"""Tests for automatic workbook version progression."""

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from budget_workbook.versioning import (
    WorkbookVersion,
    next_available_patch_version,
    parse_versioned_filename,
)


class VersioningTests(unittest.TestCase):
    def test_parse_versioned_filename(self) -> None:
        parsed = parse_versioned_filename("Personal_Budget_Workbook_v0.0.3.xlsx", "Personal_Budget_Workbook")
        self.assertEqual(parsed, WorkbookVersion(0, 0, 3))
        self.assertIsNone(parse_versioned_filename("not-a-match.xlsx", "Personal_Budget_Workbook"))

    def test_first_creation_uses_start_version(self) -> None:
        with TemporaryDirectory() as temp_dir:
            version = next_available_patch_version(
                directory=Path(temp_dir),
                prefix="Personal_Budget_Workbook",
                start_version=WorkbookVersion(0, 0, 0),
            )
            self.assertEqual(version, WorkbookVersion(0, 0, 0))

    def test_repeated_runs_increment_patch(self) -> None:
        with TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "Personal_Budget_Workbook_v0.0.0.xlsx").write_text("x")
            second = next_available_patch_version(
                directory=root,
                prefix="Personal_Budget_Workbook",
                start_version=WorkbookVersion(0, 0, 0),
            )
            self.assertEqual(second, WorkbookVersion(0, 0, 1))

            (root / "Personal_Budget_Workbook_v0.0.1.xlsx").write_text("x")
            third = next_available_patch_version(
                directory=root,
                prefix="Personal_Budget_Workbook",
                start_version=WorkbookVersion(0, 0, 0),
            )
            self.assertEqual(third, WorkbookVersion(0, 0, 2))

    def test_version_gaps_choose_highest_plus_one(self) -> None:
        with TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "Personal_Budget_Workbook_v0.0.0.xlsx").write_text("x")
            (root / "Personal_Budget_Workbook_v0.0.2.xlsx").write_text("x")
            version = next_available_patch_version(
                directory=root,
                prefix="Personal_Budget_Workbook",
                start_version=WorkbookVersion(0, 0, 0),
            )
            self.assertEqual(version, WorkbookVersion(0, 0, 3))

    def test_ignores_invalid_and_unrelated_files(self) -> None:
        with TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "Personal_Budget_Workbook_v0.0.bad.xlsx").write_text("x")
            (root / "Personal_Budget_Workbook_v0.1.0.csv").write_text("x")
            (root / "Another_Report_v0.0.9.xlsx").write_text("x")
            (root / "Personal_Budget_Workbook_v0.0.4.xlsx").write_text("x")
            version = next_available_patch_version(
                directory=root,
                prefix="Personal_Budget_Workbook",
                start_version=WorkbookVersion(0, 0, 0),
            )
            self.assertEqual(version, WorkbookVersion(0, 0, 5))


if __name__ == "__main__":
    unittest.main()
