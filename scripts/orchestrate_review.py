#!/usr/bin/env python3
"""Run the Claude builder -> deterministic QA -> Codex supervisor loop.

The orchestration core is deliberately adapter-driven: live adapters use provider
HTTP APIs, while tests and local rehearsals use deterministic mock transcripts.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import traceback
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Protocol

import jsonschema
import yaml

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / "model" / "base-v2.4.5.xlsx"
BASE_SHA256 = "5764fa6dc28137bf4cb2348e04400bd73a4663cb22ee14a546ae60e9082f4e15"
RUN_STATES = {
    "BUILDING", "QA_FAILED", "SUPERVISOR_REVIEW", "PATCHING",
    "ESCALATION_REQUIRED", "REVIEWED_RELEASE_CANDIDATE", "SUBSCRIPTION_LIMIT_PAUSED",
    "INFRASTRUCTURE_FAILED", "FAILED",
}

SECRET_PATTERNS = [
    re.compile(r"(?i)(api[_-]?key|authorization|bearer|token|cookie)(\s*[:=]\s*)([^\s\"']+)"),
    re.compile(r"\b(sk-[A-Za-z0-9_-]{12,})\b"),
]


class SubscriptionLimitError(RuntimeError):
    def __init__(self, service: str, detail: str) -> None:
        safe = redact(detail)
        super().__init__(safe[-2000:])
        self.service = service
        self.output = safe


def redact(text: str) -> str:
    for pattern in SECRET_PATTERNS:
        text = pattern.sub(lambda match: (match.group(1) + match.group(2) if match.lastindex and match.lastindex >= 3 else "") + "[REDACTED]", text)
    return text


def is_usage_limit(text: str) -> bool:
    lowered = text.lower()
    return any(marker in lowered for marker in (
        "usage limit", "rate limit", "rate_limit", "quota exceeded", "limit reached",
        "you've hit your limit", "you have hit your limit", "resets at", "too many requests",
    ))
ESCALATIONS = {
    "PRINCIPAL_DECISION", "NEW_UNDERWRITING_ASSUMPTION",
    "EVIDENCE_CONFLICT", "UNRESOLVED_AGENT_DISAGREEMENT",
}
BLOCKING_QA_TESTS = {
    "shared_formulas_zero", "shared_formula_records_zero", "shared_formula_groups_zero",
    "formula_pattern_regression", "formula_errors_zero", "missing_caches_zero",
    "external_links_zero", "merged_cells_zero", "wrapped_populated_cells_zero",
    "economic_invariants", "version_hash_checks",
}


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def read_yaml(path: Path) -> dict[str, Any]:
    return yaml.safe_load(path.read_text(encoding="utf-8")) or {}


def resolve_repo_path(value: str | Path) -> Path:
    path = Path(value)
    path = path if path.is_absolute() else ROOT / path
    resolved = path.resolve()
    if resolved != ROOT and ROOT not in resolved.parents:
        raise ValueError(f"Path escapes repository: {value}")
    return resolved


def extract_json(text: str) -> Any:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I | re.S)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start_candidates = [i for i in (text.find("{"), text.find("[")) if i >= 0]
        if not start_candidates:
            raise
        decoder = json.JSONDecoder()
        value, _ = decoder.raw_decode(text[min(start_candidates):])
        return value


def apply_builder_operations(result: dict[str, Any]) -> list[str]:
    """Apply Claude's explicit, auditable local operations (never shell strings)."""
    changed: list[str] = []
    for operation in result.get("operations", []):
        kind = operation.get("type")
        target = resolve_repo_path(operation.get("path", ""))
        if target == BASE:
            raise RuntimeError("Builder attempted to modify the frozen Base")
        if kind == "write_text":
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(operation.get("content", ""), encoding="utf-8")
            changed.append(str(target.relative_to(ROOT)))
        elif kind == "run_python":
            if target.suffix != ".py":
                raise ValueError("run_python is limited to repository Python files")
            subprocess.run([sys.executable, str(target), *map(str, operation.get("args", []))],
                           cwd=ROOT, check=True)
        else:
            raise ValueError(f"Unsupported builder operation: {kind}")
    return changed


def http_json(url: str, headers: dict[str, str], payload: dict[str, Any]) -> dict[str, Any]:
    request = urllib.request.Request(
        url, data=json.dumps(payload).encode(), headers={"content-type": "application/json", **headers}, method="POST"
    )
    try:
        with urllib.request.urlopen(request, timeout=600) as response:
            return json.loads(response.read().decode())
    except urllib.error.HTTPError as error:
        detail = error.read().decode(errors="replace")
        raise RuntimeError(f"Provider API error {error.code}: {detail[:2000]}") from error


class AgentAdapter(Protocol):
    def call(self, request: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]: ...


class AnthropicBuilderAdapter:
    def __init__(self, api_key: str, model: str = "claude-sonnet-4-6") -> None:
        self.api_key, self.model = api_key, model

    def call(self, request: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
        raw = http_json(
            "https://api.anthropic.com/v1/messages",
            {"x-api-key": self.api_key, "anthropic-version": "2023-06-01"},
            {"model": self.model, "max_tokens": 8192, "system": request["system"],
             "messages": [{"role": "user", "content": request["user"]}]},
        )
        text = "".join(part.get("text", "") for part in raw.get("content", []) if part.get("type") == "text")
        return extract_json(text), raw


class OpenAISupervisorAdapter:
    def __init__(self, api_key: str, model: str = "gpt-5") -> None:
        self.api_key, self.model = api_key, model

    def call(self, request: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
        raw = http_json(
            "https://api.openai.com/v1/responses", {"Authorization": f"Bearer {self.api_key}"},
            {"model": self.model, "store": False, "instructions": request["system"],
             "input": request["user"], "text": {"format": {"type": "json_schema",
             "name": "supervisor_review", "strict": True, "schema": request["schema"]}}},
        )
        pieces = []
        for output in raw.get("output", []):
            for part in output.get("content", []):
                if part.get("type") == "output_text":
                    pieces.append(part.get("text", ""))
        return extract_json("".join(pieces)), raw


class SubscriptionCLIAdapter:
    service = "subscription"

    def __init__(self, executable: str, timeout: int = 1800,
                 runner: Callable[..., subprocess.CompletedProcess[str]] = subprocess.run) -> None:
        self.executable, self.timeout, self.runner = executable, timeout, runner

    def _run(self, command: list[str], prompt: str) -> subprocess.CompletedProcess[str]:
        environment = os.environ.copy()
        # Subscription mode must not accidentally select provider API billing.
        for name in ("ANTHROPIC_API_KEY", "OPENAI_API_KEY", "CODEX_API_KEY"):
            environment.pop(name, None)
        try:
            completed = self.runner(command, input=prompt, text=True, capture_output=True,
                                    timeout=self.timeout, cwd=ROOT, env=environment)
        except subprocess.TimeoutExpired as error:
            raise RuntimeError(f"{self.service} CLI timed out after {self.timeout} seconds") from error
        combined = f"{completed.stdout}\n{completed.stderr}"
        if is_usage_limit(combined):
            raise SubscriptionLimitError(self.service, combined)
        if completed.returncode:
            raise RuntimeError(f"{self.service} CLI exited {completed.returncode}: {redact(combined)[-2000:]}")
        return completed


class ClaudeSubscriptionAdapter(SubscriptionCLIAdapter):
    service = "claude"

    def call(self, request: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
        prompt = request["system"] + "\n\n" + request["user"]
        schema = {
            "type": "object", "additionalProperties": False,
            "required": ["summary", "candidate_manifest", "files_changed", "assumptions_used",
                         "findings_addressed", "escalation", "operations"],
            "properties": {
                "summary": {"type": "string"},
                "candidate_manifest": {"type": "string"},
                "files_changed": {"type": "array", "items": {"type": "string"}},
                "assumptions_used": {"type": "array", "items": {}},
                "findings_addressed": {"type": "array", "items": {}},
                "escalation": {"type": ["object", "null"]},
                "operations": {"type": "array", "items": {"type": "object"}},
            },
        }
        completed = self._run([self.executable, "-p", "--output-format", "json",
                               "--json-schema", json.dumps(schema), "--permission-mode", "plan"], prompt)
        outer = extract_json(completed.stdout)
        structured = outer.get("structured_output") if isinstance(outer, dict) else None
        if structured is not None:
            result = structured
        else:
            text = outer.get("result", "") if isinstance(outer, dict) else str(outer)
            result = extract_json(text) if isinstance(text, str) else text
        if not isinstance(result, dict):
            raise TypeError(f"Claude builder result must be an object, got {type(result).__name__}")
        return result, {"exit_code": completed.returncode, "stdout": redact(completed.stdout),
                        "stderr": redact(completed.stderr)}


class CodexSubscriptionAdapter(SubscriptionCLIAdapter):
    service = "codex"

    def call(self, request: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
        schema_path = ROOT / "reports" / "review-schema.json"
        prompt = request["system"] + "\n\n" + request["user"]
        completed = self._run([self.executable, "exec", "-", "--ephemeral", "--sandbox", "read-only",
                               "--output-schema", str(schema_path), "--color", "never"], prompt)
        result = extract_json(completed.stdout)
        return result, {"exit_code": completed.returncode, "stdout": redact(completed.stdout),
                        "stderr": redact(completed.stderr)}


class TranscriptAdapter:
    """Deterministic adapter used by --mock and unit tests."""
    def __init__(self, responses: list[dict[str, Any]]) -> None:
        self.responses = list(responses)
        self.calls: list[dict[str, Any]] = []

    def call(self, request: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
        self.calls.append(request)
        response = self.responses.pop(0) if self.responses else {}
        return response, {"mock": True, "response": response}


@dataclass
class RunResult:
    status: str
    exit_code: int
    payload: dict[str, Any]


def command_check(command: list[str], timeout: int = 20) -> dict[str, Any]:
    executable = shutil.which(command[0])
    if not executable:
        return {"command": command, "installed": False, "ok": False, "output": "not found"}
    environment = os.environ.copy()
    for name in ("ANTHROPIC_API_KEY", "OPENAI_API_KEY", "CODEX_API_KEY"):
        environment.pop(name, None)
    completed = subprocess.run([executable, *command[1:]], text=True, capture_output=True,
                               timeout=timeout, cwd=ROOT, env=environment)
    output = redact(f"{completed.stdout}\n{completed.stderr}").strip()
    return {"command": command, "executable": executable, "installed": True,
            "ok": completed.returncode == 0, "exit_code": completed.returncode, "output": output}


def subscription_preflight() -> dict[str, Any]:
    checks = {
        "claude_version": command_check(["claude", "--version"]),
        "claude_auth": command_check(["claude", "auth", "status", "--json"]),
        "codex_version": command_check(["codex", "--version"]),
        "codex_auth": command_check(["codex", "login", "status"]),
        "github_auth": command_check(["gh", "auth", "status"]),
    }
    checks["base"] = {"ok": BASE.exists() and sha256(BASE) == BASE_SHA256,
                      "expected": BASE_SHA256, "actual": sha256(BASE) if BASE.exists() else None}
    status = subprocess.run(["git", "status", "--porcelain"], cwd=ROOT, text=True, capture_output=True)
    checks["repository"] = {"ok": status.returncode == 0 and not status.stdout.strip(),
                            "output": redact(status.stdout)}
    warnings = []
    if os.getenv("ANTHROPIC_API_KEY"):
        warnings.append("ANTHROPIC_API_KEY is set; it will be removed from Claude's subscription subprocess")
    if os.getenv("OPENAI_API_KEY") or os.getenv("CODEX_API_KEY"):
        warnings.append("OpenAI API key environment variable is set; it will be removed from Codex's subscription subprocess")
    blockers = []
    if not checks["claude_version"]["ok"]:
        blockers.append("Claude Code CLI missing; install it, then run: claude auth login")
    elif not checks["claude_auth"]["ok"]:
        blockers.append("Claude Code is not authenticated; run: claude auth login")
    elif not any(marker in checks["claude_auth"]["output"].lower() for marker in ("claude.ai", "subscription", '"pro"', '"max"')):
        blockers.append("Claude Code authentication could not be verified as a Claude subscription; run claude, choose Claude App/Pro/Max, then verify with /status")
    if not checks["codex_version"]["ok"]:
        blockers.append("Codex CLI missing; install it, then run: codex login")
    elif not checks["codex_auth"]["ok"] or "not logged in" in checks["codex_auth"]["output"].lower():
        blockers.append("Codex CLI is not authenticated with ChatGPT; run: codex login")
    elif "chatgpt" not in checks["codex_auth"]["output"].lower():
        blockers.append("Codex CLI is not verified as ChatGPT subscription auth; run: codex login (do not use --with-api-key)")
    if not checks["base"]["ok"]:
        blockers.append("Frozen Base hash verification failed")
    return {"checks": checks, "warnings": warnings, "blockers": blockers, "ok": not blockers}


class Orchestrator:
    def __init__(
        self, change_request: Path, builder: AgentAdapter, supervisor: AgentAdapter,
        *, max_cycles: int = 2, run_id: str | None = None, no_push: bool = False,
        dry_run: bool = False, qa_runner: Callable[[Path, Path], tuple[int, dict[str, Any]]] | None = None,
        git_enabled: bool = True, resume: bool = False,
        fallback_builder: AgentAdapter | None = None, fallback_supervisor: AgentAdapter | None = None,
    ) -> None:
        self.change_request_path = resolve_repo_path(change_request)
        self.request = read_yaml(self.change_request_path)
        self.request_id = str(self.request.get("change_request_id", ""))
        if not self.request_id or self.request_id == "CR-YYYY-NNN":
            raise ValueError("A unique change_request_id is required")
        self.builder, self.supervisor = builder, supervisor
        self.max_cycles, self.no_push, self.dry_run = max_cycles, no_push, dry_run
        self.qa_runner = qa_runner or self._run_qa
        self.git_enabled = git_enabled
        self.resume = resume
        self.fallback_builder, self.fallback_supervisor = fallback_builder, fallback_supervisor
        safe_id = re.sub(r"[^A-Za-z0-9_.-]+", "-", self.request_id)
        self.run_id = run_id or f"{safe_id}-{datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')}"
        self.run_dir = ROOT / "reports" / "runs" / self.run_id
        self.history: list[dict[str, Any]] = []
        self.pr_url: str | None = None
        self.generated_paths: set[Path] = set()

    def _state(self, state: str, cycle: int, **extra: Any) -> None:
        if state not in RUN_STATES:
            raise ValueError(state)
        event = {"state": state, "cycle": cycle, "at": utc_now(), **extra}
        self.history.append(event)
        write_json(self.run_dir / "state.json", {"current": state, "history": self.history})

    def _verify_base(self) -> None:
        actual = sha256(BASE)
        requested = str(self.request.get("base", {}).get("sha256", ""))
        if actual != BASE_SHA256 or (requested and requested != BASE_SHA256):
            raise RuntimeError(f"Frozen Base hash mismatch: expected {BASE_SHA256}, actual {actual}, request {requested}")

    def _git(self, *args: str, check: bool = True) -> subprocess.CompletedProcess[str]:
        return subprocess.run(["git", *args], cwd=ROOT, text=True, capture_output=True, check=check)

    def _prepare_branch(self) -> str:
        branch = f"change/{re.sub(r'[^A-Za-z0-9._/-]+', '-', self.request_id)}"
        if not self.git_enabled or self.dry_run:
            return branch
        current = self._git("branch", "--show-current").stdout.strip()
        if current != branch:
            exists = self._git("show-ref", "--verify", f"refs/heads/{branch}", check=False).returncode == 0
            self._git("switch", branch if exists else "-c", *( [] if exists else [branch] ))
        return branch

    def _publish(self, branch: str) -> None:
        """Commit declared builder artifacts and the audit trail, then open/update a PR."""
        if not self.git_enabled or self.dry_run:
            return
        # Run transcripts are intentionally local/ignored; commit only the approved
        # change request and files explicitly declared by the builder.
        paths = {self.change_request_path, *self.generated_paths}
        for path in sorted(paths, key=str):
            resolved = resolve_repo_path(path)
            self._git("add", "--", str(resolved.relative_to(ROOT)))
        if self._git("diff", "--cached", "--quiet", check=False).returncode != 0:
            self._git("commit", "-m", f"Build reviewed candidate for {self.request_id}")
        if self.no_push:
            return
        self._git("push", "--set-upstream", "origin", branch)
        view = subprocess.run(["gh", "pr", "view", branch, "--json", "url", "--jq", ".url"],
                              cwd=ROOT, text=True, capture_output=True)
        if view.returncode == 0 and view.stdout.strip():
            self.pr_url = view.stdout.strip()
        else:
            created = subprocess.run(
                ["gh", "pr", "create", "--base", "main", "--head", branch,
                 "--title", f"{self.request_id}: reviewed model candidate",
                 "--body", f"Automated builder, deterministic QA, and independent supervisor run: `{self.run_id}`.\n\nDo not merge until the desktop Excel gate is complete."],
                cwd=ROOT, text=True, capture_output=True, check=True,
            )
            self.pr_url = created.stdout.strip().splitlines()[-1]
        final_path = self.run_dir / "final.json"
        if final_path.exists():
            final = json.loads(final_path.read_text(encoding="utf-8"))
            final["pr_url"] = self.pr_url
            write_json(final_path, final)
            self._git("add", "--", str(final_path.relative_to(ROOT)))
            self._git("commit", "-m", f"Record pull request for {self.request_id}")
            self._git("push")

    def _spec_bundle(self) -> str:
        chunks = []
        for path in sorted((ROOT / "spec").glob("*.yaml")):
            chunks.append(f"<document source=\"{path.relative_to(ROOT)}\">\n{path.read_text(encoding='utf-8')}\n</document>")
        return "\n".join(chunks)

    def _builder_request(self, cycle: int, feedback: dict[str, Any] | None) -> dict[str, Any]:
        user = (
            f"<change_request>\n{self.change_request_path.read_text(encoding='utf-8')}\n</change_request>\n"
            f"{self._spec_bundle()}\n<cycle>{cycle}</cycle>\n"
            f"<automated_feedback>{json.dumps(feedback or {}, indent=2)}</automated_feedback>\n"
            "Work in the repository checkout. Return one JSON object with keys summary, candidate_manifest, "
            "files_changed, assumptions_used, findings_addressed, escalation (null or authorized category package), "
            "and operations. Operations are ordered objects of type write_text (path, content) or run_python "
            "(path, args). You may write a repository-local Python builder and then run it to create/patch binary workbooks. "
            "candidate_manifest must be a repository-relative path. Do not modify the frozen Base."
        )
        return {"system": (ROOT / "prompts" / "claude-builder.md").read_text(encoding="utf-8"), "user": user}

    def _supervisor_request(self, cycle: int, manifest: Path, qa: dict[str, Any]) -> dict[str, Any]:
        schema = json.loads((ROOT / "reports" / "review-schema.json").read_text(encoding="utf-8"))
        user = json.dumps({"change_request": self.request, "cycle": cycle,
                           "candidate_manifest": read_yaml(manifest), "qa_report": qa}, indent=2, default=str)
        return {"system": (ROOT / "prompts" / "codex-supervisor.md").read_text(encoding="utf-8"),
                "user": user, "schema": schema}

    def _run_qa(self, manifest: Path, output: Path) -> tuple[int, dict[str, Any]]:
        command = [sys.executable, str(ROOT / "scripts" / "run_regression.py"),
                   "--manifest", str(manifest), "--output", str(output)]
        try:
            result = subprocess.run(command, cwd=ROOT, text=True, capture_output=True)
        except Exception as error:
            payload = {"passed": False, "results": [], "infrastructure_error": {
                "type": type(error).__name__, "message": str(error), "command": command}}
            write_json(output, payload)
            return 2, payload
        payload = json.loads(output.read_text(encoding="utf-8")) if output.exists() else {
            "passed": False, "results": [], "infrastructure_error": {
                "type": "MissingQAReport", "message": "QA did not produce a report"}}
        if result.returncode not in (0, 1):
            payload.setdefault("infrastructure_error", {})
            payload["infrastructure_error"].update({"command": command, "exit_code": result.returncode,
                                                     "stdout": redact(result.stdout), "stderr": redact(result.stderr)})
        return result.returncode, payload

    @staticmethod
    def _qa_defects(qa: dict[str, Any]) -> list[dict[str, Any]]:
        defects = []
        for result in qa.get("results", []):
            if not result.get("passed"):
                defects.extend(result.get("findings", []))
                if not result.get("findings"):
                    defects.append({"category": "deterministic_qa", "test": result.get("test"),
                                    "issue": result.get("details"), "blocks_release": True})
        return defects

    @staticmethod
    def _escalation(findings: list[dict[str, Any]]) -> dict[str, Any] | None:
        for finding in findings:
            category = finding.get("category")
            if category in ESCALATIONS - {"UNRESOLVED_AGENT_DISAGREEMENT"}:
                return {"category": category, "decision_required": finding.get("required_fix", finding.get("issue", "Decision required")),
                        "why_agents_cannot_resolve": finding.get("issue", "Requires principal authority"),
                        "options": finding.get("options", []), "quantified_effect": finding.get("economic_effect", "Not quantified"),
                        "recommendation": finding.get("recommendation", "Review the evidence and choose an option"),
                        "affected_cells_or_assumptions": [finding.get("cell_or_range")] if finding.get("cell_or_range") else []}
        return None

    def _finish(self, status: str, cycle: int, *, manifest: Path | None = None,
                qa: dict[str, Any] | None = None, findings: list[dict[str, Any]] | None = None,
                escalation: dict[str, Any] | None = None, error: str | None = None) -> RunResult:
        self._state(status, cycle, error=error) if status in RUN_STATES else None
        if escalation:
            write_json(self.run_dir / "escalation.json", escalation)
        manifest_data = read_yaml(manifest) if manifest and manifest.exists() else {}
        workbook_value = manifest_data.get("workbook", {})
        workbook_value = workbook_value.get("path") if isinstance(workbook_value, dict) else workbook_value
        candidate = resolve_repo_path(workbook_value) if workbook_value else None
        final = {"status": status, "candidate_workbook": str(candidate.relative_to(ROOT)) if candidate else None,
                 "candidate_hash": sha256(candidate) if candidate and candidate.exists() else None,
                 "base_hash": sha256(BASE), "qa_status": "PASS" if qa and qa.get("passed") else ("FAIL" if qa else "NOT_RUN"),
                 "supervisor_status": "PASS" if findings == [] else ("BLOCKED" if findings is not None else "NOT_RUN"),
                 "number_of_cycles": cycle, "unresolved_findings": findings or [], "pr_url": self.pr_url,
                 "desktop_excel_gate_required": "YES", "error": error}
        write_json(self.run_dir / "final.json", final)
        return RunResult(status, 0 if status == "REVIEWED_RELEASE_CANDIDATE" else 2 if status == "ESCALATION_REQUIRED" else 1, final)

    def _pause(self, service: str, cycle: int, stage: str, detail: str,
               manifest: Path | None = None, qa: dict[str, Any] | None = None) -> RunResult:
        checkpoint = {"service": service, "cycle": cycle, "stage": stage, "detail": redact(detail),
                      "manifest": str(manifest.relative_to(ROOT)) if manifest else None, "qa": qa}
        write_json(self.run_dir / "checkpoint.json", checkpoint)
        write_json(self.run_dir / "subscription-limit.json", {"service": service, "cycle": cycle,
                   "stage": stage, "output": redact(detail)})
        self._state("SUBSCRIPTION_LIMIT_PAUSED", cycle, service=service, resume_stage=stage)
        payload = {"status": "SUBSCRIPTION_LIMIT_PAUSED", "service": service, "run_id": self.run_id,
                   "resume_command": f"python scripts/orchestrate_review.py --resume {self.run_id}",
                   "detail": redact(detail)}
        write_json(self.run_dir / "final.json", payload)
        return RunResult("SUBSCRIPTION_LIMIT_PAUSED", 3, payload)

    def _infrastructure_failed(self, cycle: int, stage: str, report: dict[str, Any]) -> RunResult:
        error = report.get("infrastructure_error", report)
        payload = {"status": "INFRASTRUCTURE_FAILED", "stage": stage, "cycle": cycle,
                   "qa_status": "NOT_COMPLETED", "supervisor_status": "NOT_RUN",
                   "number_of_cycles": cycle - 1, "infrastructure_error": error,
                   "run_id": self.run_id}
        write_json(self.run_dir / "infrastructure-error.json", payload)
        write_json(self.run_dir / "checkpoint.json", {"stage": stage, "cycle": cycle,
                   "infrastructure_error": error})
        self._state("INFRASTRUCTURE_FAILED", cycle, stage=stage)
        write_json(self.run_dir / "final.json", payload)
        return RunResult("INFRASTRUCTURE_FAILED", 5, payload)

    def run(self) -> RunResult:
        self.run_dir.mkdir(parents=True, exist_ok=self.resume)
        if not self.resume:
            shutil.copy2(self.change_request_path, self.run_dir / "change-request.yaml")
        try:
            self._verify_base()
            branch = self._prepare_branch()
            write_json(self.run_dir / "run.json", {"run_id": self.run_id, "branch": branch,
                       "change_request": self.request_id, "started_at": utc_now(), "base_hash": sha256(BASE)})
            if self.dry_run:
                return self._finish("REVIEWED_RELEASE_CANDIDATE", 0)
            feedback: dict[str, Any] | None = None
            last_manifest: Path | None = None
            last_qa: dict[str, Any] | None = None
            last_findings: list[dict[str, Any]] = []
            supervisor_ran = False
            checkpoint_path = self.run_dir / "checkpoint.json"
            checkpoint = json.loads(checkpoint_path.read_text(encoding="utf-8")) if self.resume and checkpoint_path.exists() else {}
            start_cycle = int(checkpoint.get("cycle", 1))
            for cycle in range(start_cycle, self.max_cycles + 1):
                cycle_dir = self.run_dir / f"cycle-{cycle:02d}"
                cycle_dir.mkdir(exist_ok=True)
                resume_supervisor = checkpoint.get("stage") == "SUPERVISOR_REVIEW" and cycle == start_cycle
                if not resume_supervisor:
                    self._state("BUILDING" if cycle == 1 else "PATCHING", cycle)
                    builder_request = self._builder_request(cycle, feedback)
                    write_json(cycle_dir / "claude-request.json", builder_request)
                    try:
                        builder, raw_builder = self.builder.call(builder_request)
                    except SubscriptionLimitError as error:
                        if self.fallback_builder:
                            builder, raw_builder = self.fallback_builder.call(builder_request)
                        else:
                            return self._pause(error.service, cycle, "BUILDING", error.output)
                    write_json(cycle_dir / "claude-response.json", raw_builder)
                    write_json(cycle_dir / "builder-result.json", builder)
                    applied = apply_builder_operations(builder)
                    builder.setdefault("files_changed", []).extend(applied)
                    for value in builder.get("files_changed", []):
                        self.generated_paths.add(resolve_repo_path(value))
                else:
                    builder = {}
                if builder.get("escalation"):
                    package = builder["escalation"]
                    if package.get("category") not in ESCALATIONS - {"UNRESOLVED_AGENT_DISAGREEMENT"}:
                        feedback = {"rejected_escalation": package, "reason": "Unauthorized escalation category"}
                        continue
                    return self._finish("ESCALATION_REQUIRED", cycle, escalation=package)
                if not resume_supervisor and not builder.get("candidate_manifest"):
                    feedback = {"builder_protocol_error": "candidate_manifest is required"}
                    continue
                last_manifest = (resolve_repo_path(checkpoint["manifest"]) if resume_supervisor
                                 else resolve_repo_path(builder["candidate_manifest"]))
                manifest_data = read_yaml(last_manifest)
                workbook_value = manifest_data.get("workbook", {})
                workbook_value = workbook_value.get("path") if isinstance(workbook_value, dict) else workbook_value
                candidate = resolve_repo_path(workbook_value) if workbook_value else None
                write_json(cycle_dir / "hashes.json", {
                    "cycle": cycle, "base_sha256": sha256(BASE),
                    "manifest_sha256": sha256(last_manifest),
                    "candidate_sha256": sha256(candidate) if candidate and candidate.exists() else None,
                })
                qa_path = cycle_dir / "qa-report.json"
                if resume_supervisor:
                    last_qa, qa_code = checkpoint["qa"], 0
                else:
                    qa_code, last_qa = self.qa_runner(last_manifest, qa_path)
                    write_json(qa_path, last_qa)
                defects = self._qa_defects(last_qa)
                if qa_code not in (0, 1) or last_qa.get("infrastructure_error"):
                    return self._infrastructure_failed(cycle, "DETERMINISTIC_QA", last_qa)
                if qa_code or not last_qa.get("passed"):
                    self._state("QA_FAILED", cycle, defect_count=len(defects))
                    feedback = {"source": "deterministic_qa", "defects": defects}
                    continue
                self._state("SUPERVISOR_REVIEW", cycle)
                supervisor_request = self._supervisor_request(cycle, last_manifest, last_qa)
                write_json(cycle_dir / "codex-request.json", supervisor_request)
                try:
                    review, raw_review = self.supervisor.call(supervisor_request)
                except SubscriptionLimitError as error:
                    if self.fallback_supervisor:
                        review, raw_review = self.fallback_supervisor.call(supervisor_request)
                    else:
                        return self._pause(error.service, cycle, "SUPERVISOR_REVIEW", error.output, last_manifest, last_qa)
                if checkpoint_path.exists():
                    checkpoint_path.unlink()
                supervisor_ran = True
                write_json(cycle_dir / "codex-response.json", raw_review)
                jsonschema.validate(review, supervisor_request["schema"])
                last_findings = review["findings"]
                write_json(cycle_dir / "findings.json", last_findings)
                escalation = self._escalation(last_findings)
                if escalation:
                    return self._finish("ESCALATION_REQUIRED", cycle, manifest=last_manifest,
                                        qa=last_qa, findings=last_findings, escalation=escalation)
                blockers = [f for f in last_findings if f.get("blocks_release") or f.get("category") == "unauthorized_change"]
                if not blockers:
                    self.generated_paths.add(last_manifest)
                    result = self._finish("REVIEWED_RELEASE_CANDIDATE", cycle, manifest=last_manifest,
                                          qa=last_qa, findings=[])
                    self._publish(branch)
                    result.payload["pr_url"] = self.pr_url
                    return result
                feedback = {"source": "codex_supervisor", "findings": blockers}
            escalation = {"category": "UNRESOLVED_AGENT_DISAGREEMENT",
                          "decision_required": "Resolve findings that remained after two automated cycles",
                          "why_agents_cannot_resolve": "The builder and supervisor did not converge within the configured cycle cap",
                          "options": ["Authorize a targeted change", "Reject the candidate", "Provide clarifying evidence"],
                          "quantified_effect": "See unresolved findings", "recommendation": "Review the unresolved findings",
                          "affected_cells_or_assumptions": [f.get("cell_or_range") for f in last_findings if f.get("cell_or_range")]}
            return self._finish("ESCALATION_REQUIRED", self.max_cycles, manifest=last_manifest,
                                qa=last_qa, findings=last_findings if supervisor_ran else None, escalation=escalation)
        except Exception as error:
            return self._infrastructure_failed(len({e["cycle"] for e in self.history}) or 1,
                                               "ORCHESTRATOR_RUNTIME", {"infrastructure_error": {
                                                   "type": type(error).__name__, "message": str(error),
                                                   "traceback": traceback.format_exc()}})


def mock_adapters(path: Path | None, change_request: Path) -> tuple[TranscriptAdapter, TranscriptAdapter]:
    if path:
        data = json.loads(path.read_text(encoding="utf-8"))
        return TranscriptAdapter(data.get("claude", [])), TranscriptAdapter(data.get("codex", []))
    manifests = list((ROOT / "scenarios").glob("*/manifest.yaml"))
    manifest = str(manifests[0].relative_to(ROOT)) if manifests else "change-requests/template.yaml"
    request_id = str(read_yaml(resolve_repo_path(change_request)).get("change_request_id", "MOCK"))
    reviews = [{"change_request_id": request_id, "cycle": cycle, "findings": []} for cycle in range(1, 11)]
    return TranscriptAdapter([{"summary": "mock", "candidate_manifest": manifest}] * 10), TranscriptAdapter(reviews)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("change_request", type=Path, nargs="?")
    parser.add_argument("--auth", choices=("subscription", "api"), default="subscription")
    parser.add_argument("--allow-api-fallback", action="store_true")
    parser.add_argument("--resume", metavar="RUN_ID")
    parser.add_argument("--cli-timeout", type=int, default=1800)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--max-cycles", type=int, default=2)
    parser.add_argument("--no-push", action="store_true")
    parser.add_argument("--mock", nargs="?", const="", metavar="TRANSCRIPT.json")
    parser.add_argument("--anthropic-model", default=os.getenv("ANTHROPIC_MODEL", "claude-sonnet-4-6"))
    parser.add_argument("--openai-model", default=os.getenv("OPENAI_MODEL", "gpt-5"))
    args = parser.parse_args()
    if args.resume:
        resume_dir = ROOT / "reports" / "runs" / args.resume
        saved_request = resume_dir / "change-request.yaml"
        if not saved_request.exists():
            parser.error(f"Run {args.resume!r} does not exist or has no saved change request")
        args.change_request = saved_request
    if not args.change_request:
        parser.error("change_request is required unless --resume is supplied")
    if args.max_cycles < 1:
        parser.error("--max-cycles must be at least 1")
    fallback_builder = fallback_supervisor = None
    if args.mock is not None:
        builder, supervisor = mock_adapters(Path(args.mock) if args.mock else None, args.change_request)
    elif args.auth == "api":
        anthropic_key, openai_key = os.getenv("ANTHROPIC_API_KEY"), os.getenv("OPENAI_API_KEY")
        if not args.dry_run and (not anthropic_key or not openai_key):
            parser.error("ANTHROPIC_API_KEY and OPENAI_API_KEY are required (or use --mock)")
        builder = AnthropicBuilderAdapter(anthropic_key or "", args.anthropic_model)
        supervisor = OpenAISupervisorAdapter(openai_key or "", args.openai_model)
    else:
        preflight = subscription_preflight()
        print(json.dumps({"subscription_preflight": preflight}, indent=2))
        if not preflight["ok"]:
            return 4
        builder = ClaudeSubscriptionAdapter(preflight["checks"]["claude_version"]["executable"], args.cli_timeout)
        supervisor = CodexSubscriptionAdapter(preflight["checks"]["codex_version"]["executable"], args.cli_timeout)
        if args.allow_api_fallback:
            anthropic_key, openai_key = os.getenv("ANTHROPIC_API_KEY"), os.getenv("OPENAI_API_KEY")
            if not anthropic_key or not openai_key:
                parser.error("--allow-api-fallback requires ANTHROPIC_API_KEY and OPENAI_API_KEY")
            fallback_builder = AnthropicBuilderAdapter(anthropic_key, args.anthropic_model)
            fallback_supervisor = OpenAISupervisorAdapter(openai_key, args.openai_model)
    result = Orchestrator(args.change_request, builder, supervisor, max_cycles=args.max_cycles,
                          no_push=args.no_push, dry_run=args.dry_run,
                          run_id=args.resume, resume=bool(args.resume),
                          fallback_builder=fallback_builder, fallback_supervisor=fallback_supervisor,
                          git_enabled=args.mock is None).run()
    print(json.dumps(result.payload, indent=2))
    return result.exit_code


if __name__ == "__main__":
    raise SystemExit(main())
