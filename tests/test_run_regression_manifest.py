from pathlib import Path
import tempfile
import unittest

from scripts.run_regression import ManifestWorkbookPathError, ROOT, resolve_workbook_path


class WorkbookPathResolutionTests(unittest.TestCase):
    def setUp(self):
        self.manifest = ROOT / "scenarios" / "example" / "manifest.yaml"

    def test_workbook_mapping_path_is_repo_relative(self):
        actual = resolve_workbook_path(self.manifest, {"workbook": {"path": "scenarios/example/candidate.xlsx"}})
        self.assertEqual(actual, (ROOT / "scenarios/example/candidate.xlsx").resolve())

    def test_top_level_workbook_path_is_manifest_relative(self):
        actual = resolve_workbook_path(self.manifest, {"workbook_path": "candidate.xlsx"})
        self.assertEqual(actual, (self.manifest.parent / "candidate.xlsx").resolve())

    def test_legacy_workbook_string_is_manifest_relative(self):
        actual = resolve_workbook_path(self.manifest, {"workbook": "candidate.xlsx"})
        self.assertEqual(actual, (self.manifest.parent / "candidate.xlsx").resolve())

    def test_malformed_workbook_object_is_rejected(self):
        with self.assertRaisesRegex(ManifestWorkbookPathError, "workbook.path"):
            resolve_workbook_path(self.manifest, {"workbook": {"sha256": "abc"}})

    def test_missing_workbook_path_is_rejected(self):
        with self.assertRaisesRegex(ManifestWorkbookPathError, "must define"):
            resolve_workbook_path(self.manifest, {})


if __name__ == "__main__":
    unittest.main()
