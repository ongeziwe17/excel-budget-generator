"""Versioning utilities for workbook outputs."""

from dataclasses import dataclass


@dataclass(frozen=True)
class WorkbookVersion:
    """Semantic-style version used in generated workbook filenames."""

    major: int = 0
    minor: int = 0
    patch: int = 0

    def __str__(self) -> str:
        return f"v{self.major}.{self.minor}.{self.patch}"

    def bump_major(self) -> "WorkbookVersion":
        return WorkbookVersion(self.major + 1, 0, 0)

    def bump_minor(self) -> "WorkbookVersion":
        return WorkbookVersion(self.major, self.minor + 1, 0)

    def bump_patch(self) -> "WorkbookVersion":
        return WorkbookVersion(self.major, self.minor, self.patch + 1)
