"""Permanent Task C regression entry point.

The detailed independent checks live in scripts/run_task_c_qa.py so the same
runner can be used locally and in GitHub Actions. This module intentionally
does not approve manifest-declared booleans.
"""

from pathlib import Path
import unittest

from scripts.run_task_c_qa import run_checks


class TaskCReviewedCandidateTests(unittest.TestCase):
    def test_task_c_reviewed_candidate(self) -> None:
        repo = Path(__file__).resolve().parents[1]
        result = run_checks(repo, write_reports=False)
        failures = [item for item in result["tests"] if not item["passed"]]
        self.assertFalse(failures, failures)
