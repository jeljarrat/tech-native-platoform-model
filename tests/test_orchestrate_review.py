from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from scripts.orchestrate_review import (BASE_SHA256, ClaudeSubscriptionAdapter,
    AnthropicBuilderAdapter, CodexSubscriptionAdapter, OpenAISupervisorAdapter,
    Orchestrator, ROOT, SubscriptionLimitError, TranscriptAdapter,
    subscription_preflight)


MANIFEST = "scenarios/TASK-C-PATH-10M/manifest.yaml"


def builder_response(escalation=None):
    return {"summary": "mock build", "candidate_manifest": MANIFEST, "escalation": escalation}


def review(cycle: int, findings=None):
    return {"change_request_id": "CR-TEST-001", "cycle": cycle, "findings": findings or []}


def blocker(finding_id="FND-1", category="formula_regression"):
    return {"finding_id": finding_id, "severity": "high", "category": category, "sheet": "Returns",
            "cell_or_range": "E57", "issue": "Blocking issue", "economic_effect": "Material",
            "blocks_release": True, "requires_jack_decision": False, "required_fix": "Fix it",
            "required_test": "formula_pattern_regression"}


class OrchestrationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(dir=ROOT / "reports")
        self.request = Path(self.temp.name) / "request.yaml"
        self.request.write_text(
            f"change_request_id: CR-TEST-001\nbase:\n  sha256: {BASE_SHA256}\n", encoding="utf-8")
        self.run_dirs = []

    def tearDown(self):
        self.temp.cleanup()
        for path in self.run_dirs:
            if path.exists():
                import shutil
                shutil.rmtree(path)

    def make(self, builders, supervisors, qa_results, max_cycles=2):
        queue = list(qa_results)

        def qa(_manifest, _output):
            payload = queue.pop(0)
            return (0 if payload["passed"] else 1), payload

        run_id = "test-" + next(tempfile._get_candidate_names())
        orchestrator = Orchestrator(self.request, TranscriptAdapter(builders), TranscriptAdapter(supervisors),
                                    max_cycles=max_cycles, run_id=run_id, qa_runner=qa, git_enabled=False)
        self.run_dirs.append(orchestrator.run_dir)
        return orchestrator

    def test_qa_failure_loops_back_to_claude(self):
        qa_fail = {"passed": False, "results": [{"test": "missing_caches_zero", "passed": False,
                   "details": "one missing cache", "findings": []}]}
        orch = self.make([builder_response(), builder_response()], [review(2)], [qa_fail, {"passed": True, "results": []}])
        result = orch.run()
        self.assertEqual(result.status, "REVIEWED_RELEASE_CANDIDATE")
        self.assertIn("deterministic_qa", orch.builder.calls[1]["user"])
        self.assertTrue((orch.run_dir / "cycle-01" / "qa-report.json").exists())

    def test_codex_blocker_loops_back_to_claude(self):
        orch = self.make([builder_response(), builder_response()], [review(1, [blocker()]), review(2)],
                         [{"passed": True, "results": []}] * 2)
        result = orch.run()
        self.assertEqual(result.status, "REVIEWED_RELEASE_CANDIDATE")
        self.assertIn("FND-1", orch.builder.calls[1]["user"])

    def test_clean_result_stops_automatically(self):
        orch = self.make([builder_response()], [review(1)], [{"passed": True, "results": []}])
        result = orch.run()
        self.assertEqual(result.status, "REVIEWED_RELEASE_CANDIDATE")
        self.assertEqual(result.payload["number_of_cycles"], 1)

    def test_only_authorized_category_escalates(self):
        unauthorized = blocker(category="PLEASE_ASK_JACK")
        orch = self.make([builder_response(), builder_response()], [review(1, [unauthorized]), review(2)],
                         [{"passed": True, "results": []}] * 2)
        self.assertEqual(orch.run().status, "REVIEWED_RELEASE_CANDIDATE")
        authorized = blocker(category="NEW_UNDERWRITING_ASSUMPTION")
        orch2 = self.make([builder_response()], [review(1, [authorized])], [{"passed": True, "results": []}])
        self.assertEqual(orch2.run().status, "ESCALATION_REQUIRED")

    def test_two_cycle_disagreement_escalates(self):
        orch = self.make([builder_response(), builder_response()], [review(1, [blocker()]), review(2, [blocker()])],
                         [{"passed": True, "results": []}] * 2)
        result = orch.run()
        self.assertEqual(result.status, "ESCALATION_REQUIRED")
        escalation = json.loads((orch.run_dir / "escalation.json").read_text())
        self.assertEqual(escalation["category"], "UNRESOLVED_AGENT_DISAGREEMENT")

    def test_base_mutation_stops_immediately(self):
        orch = self.make([], [], [])
        with patch("scripts.orchestrate_review.sha256", return_value="bad"):
            result = orch.run()
        self.assertEqual(result.status, "INFRASTRUCTURE_FAILED")
        self.assertEqual(len(orch.builder.calls), 0)
        self.assertIn("Frozen Base hash mismatch", result.payload["infrastructure_error"]["message"])

    def test_qa_harness_failure_stops_without_more_agents(self):
        run_id = "test-" + next(tempfile._get_candidate_names())
        builder = TranscriptAdapter([builder_response(), builder_response()])
        supervisor = TranscriptAdapter([review(1)])
        failure = {"passed": False, "results": [], "infrastructure_error": {
            "type": "RuntimeError", "message": "harness crashed", "traceback": "trace"}}
        orch = Orchestrator(self.request, builder, supervisor, run_id=run_id,
                            qa_runner=lambda *_: (2, failure), git_enabled=False)
        self.run_dirs.append(orch.run_dir)
        result = orch.run()
        self.assertEqual(result.status, "INFRASTRUCTURE_FAILED")
        self.assertEqual(len(builder.calls), 1)
        self.assertEqual(len(supervisor.calls), 0)
        self.assertEqual(result.payload["qa_status"], "NOT_COMPLETED")
        self.assertEqual(result.payload["supervisor_status"], "NOT_RUN")

    def test_run_artifacts_are_preserved(self):
        orch = self.make([builder_response()], [review(1)], [{"passed": True, "results": []}])
        orch.run()
        required = ["change-request.yaml", "run.json", "state.json", "final.json",
                    "cycle-01/claude-request.json", "cycle-01/claude-response.json",
                    "cycle-01/hashes.json", "cycle-01/qa-report.json", "cycle-01/codex-request.json",
                    "cycle-01/codex-response.json", "cycle-01/findings.json"]
        for name in required:
            self.assertTrue((orch.run_dir / name).exists(), name)

    def test_subscription_mode_claude_mock(self):
        payload = json.dumps({"result": json.dumps(builder_response())})
        completed = __import__("subprocess").CompletedProcess([], 0, payload, "progress")
        adapter = ClaudeSubscriptionAdapter("claude", runner=lambda *a, **k: completed)
        result, raw = adapter.call({"system": "builder", "user": "request"})
        self.assertEqual(result["candidate_manifest"], MANIFEST)
        self.assertEqual(raw["exit_code"], 0)

    def test_subscription_mode_codex_mock(self):
        completed = __import__("subprocess").CompletedProcess([], 0, json.dumps(review(1)), "")
        adapter = CodexSubscriptionAdapter("codex", runner=lambda *a, **k: completed)
        result, _ = adapter.call({"system": "review", "user": "state", "schema": {}})
        self.assertEqual(result["findings"], [])

    @patch("scripts.orchestrate_review.http_json")
    def test_api_adapters_remain_available(self, http):
        http.side_effect = [
            {"content": [{"type": "text", "text": json.dumps(builder_response())}]},
            {"output": [{"content": [{"type": "output_text", "text": json.dumps(review(1))}]}]},
        ]
        built, _ = AnthropicBuilderAdapter("test-key").call({"system": "s", "user": "u"})
        checked, _ = OpenAISupervisorAdapter("test-key").call({"system": "s", "user": "u", "schema": {}})
        self.assertEqual(built["candidate_manifest"], MANIFEST)
        self.assertEqual(checked["findings"], [])

    @patch("scripts.orchestrate_review.command_check")
    def test_missing_and_unauthenticated_cli(self, check):
        def result(command, timeout=20):
            if command[0] == "claude":
                return {"ok": False, "installed": False, "output": "not found"}
            if command[:2] == ["codex", "login"]:
                return {"ok": False, "installed": True, "output": "Not logged in"}
            return {"ok": True, "installed": True, "output": "ok"}
        check.side_effect = result
        preflight = subscription_preflight()
        self.assertFalse(preflight["ok"])
        self.assertTrue(any("Claude Code CLI missing" in item for item in preflight["blockers"]))
        self.assertTrue(any("codex login" in item for item in preflight["blockers"]))

    def test_subscription_usage_limit_pauses_without_api_fallback(self):
        class Limited:
            def call(self, _request):
                raise SubscriptionLimitError("claude", "usage limit resets at 8pm token=secret")
        orch = self.make([], [], [])
        orch.builder = Limited()
        result = orch.run()
        self.assertEqual(result.status, "SUBSCRIPTION_LIMIT_PAUSED")
        self.assertTrue((orch.run_dir / "checkpoint.json").exists())
        self.assertNotIn("secret", (orch.run_dir / "final.json").read_text())

    def test_explicit_api_fallback_only_when_supplied(self):
        class Limited:
            def call(self, _request):
                raise SubscriptionLimitError("claude", "usage limit")
        orch = self.make([], [review(1)], [{"passed": True, "results": []}])
        orch.builder = Limited()
        fallback = TranscriptAdapter([builder_response()])
        orch.fallback_builder = fallback
        self.assertEqual(orch.run().status, "REVIEWED_RELEASE_CANDIDATE")
        self.assertEqual(len(fallback.calls), 1)

    def test_resume_continues_at_supervisor_stage(self):
        class LimitedSupervisor:
            def call(self, _request):
                raise SubscriptionLimitError("codex", "usage limit")
        orch = self.make([builder_response()], [], [{"passed": True, "results": []}])
        orch.supervisor = LimitedSupervisor()
        paused = orch.run()
        self.assertEqual(paused.status, "SUBSCRIPTION_LIMIT_PAUSED")
        builder = TranscriptAdapter([])
        resumed = Orchestrator(orch.run_dir / "change-request.yaml", builder, TranscriptAdapter([review(1)]),
                               run_id=orch.run_id, resume=True, qa_runner=lambda *_: self.fail("QA reran"),
                               git_enabled=False)
        self.assertEqual(resumed.run().status, "REVIEWED_RELEASE_CANDIDATE")
        self.assertEqual(builder.calls, [])


if __name__ == "__main__":
    unittest.main()
