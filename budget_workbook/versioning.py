"""Versioning utilities for workbook outputs."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path


_VERSION_RE = re.compile(r"^v(?P<major>\d+)\.(?P<minor>\d+)\.(?P<patch>\d+)$")


@dataclass(frozen=True, order=True)
class WorkbookVersion:
    """Semantic-style version used in generated workbook filenames."""

    major: int = 0
    minor: int = 0
    patch: int = 0

    def __str__(self) -> str:
        return f"v{self.major}.{self.minor}.{self.patch}"

    @classmethod
    def from_string(cls, value: str) -> "WorkbookVersion | None":
        """Parse a version string like v1.2.3. Return None if invalid."""
        match = _VERSION_RE.fullmatch(value)
        if not match:
            return None
        return cls(
            major=int(match.group("major")),
            minor=int(match.group("minor")),
            patch=int(match.group("patch")),
        )

    def bump_major(self) -> "WorkbookVersion":
        return WorkbookVersion(self.major + 1, 0, 0)

    def bump_minor(self) -> "WorkbookVersion":
        return WorkbookVersion(self.major, self.minor + 1, 0)

    def bump_patch(self) -> "WorkbookVersion":
        return WorkbookVersion(self.major, self.minor, self.patch + 1)


def parse_versioned_filename(filename: str, prefix: str) -> WorkbookVersion | None:
    """Extract a workbook version from `<prefix>_vX.Y.Z.xlsx` filenames."""
    expected_prefix = f"{prefix}_"
    if not filename.startswith(expected_prefix) or not filename.endswith(".xlsx"):
        return None

    version_part = filename[len(expected_prefix) : -len(".xlsx")]
    return WorkbookVersion.from_string(version_part)


def next_available_patch_version(
    directory: Path,
    prefix: str,
    start_version: WorkbookVersion,
) -> WorkbookVersion:
    """Find the next patch version not already used in `directory`.

    Only versions sharing major/minor with `start_version` are considered.
    Invalid and unrelated filenames are ignored safely.
    """
    if not directory.exists():
        return start_version

    matching_versions: list[WorkbookVersion] = []
    for file_path in directory.iterdir():
        if not file_path.is_file():
            continue
        version = parse_versioned_filename(file_path.name, prefix)
        if version is None:
            continue
        if version.major == start_version.major and version.minor == start_version.minor:
            matching_versions.append(version)

    if not matching_versions:
        return start_version

    highest = max(matching_versions)
    if highest.patch < start_version.patch and start_version not in matching_versions:
        return start_version

    return WorkbookVersion(start_version.major, start_version.minor, highest.patch + 1)
