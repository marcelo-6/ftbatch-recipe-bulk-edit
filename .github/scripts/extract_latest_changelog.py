#!/usr/bin/env python3
"""Extract the latest released section from CHANGELOG.md."""

from __future__ import annotations

import argparse
import re
from pathlib import Path

SECTION_RE = re.compile(r"^##\s+(?P<title>.+?)\s*$")


def extract_latest_release_section(changelog_text: str) -> str:
    lines = changelog_text.splitlines()
    sections: list[tuple[str, int, int]] = []

    start_idx: int | None = None
    title: str | None = None
    for idx, line in enumerate(lines):
        match = SECTION_RE.match(line.strip())
        if not match:
            continue
        if start_idx is not None and title is not None:
            sections.append((title, start_idx, idx))
        start_idx = idx
        title = match.group("title").strip()

    if start_idx is not None and title is not None:
        sections.append((title, start_idx, len(lines)))

    if not sections:
        raise ValueError("No release sections (## ...) found in changelog")

    # Prefer first non-Unreleased section; fallback to first section if needed.
    chosen = next(
        (
            section
            for section in sections
            if not section[0].lower().startswith("unreleased")
        ),
        sections[0],
    )
    _, start, end = chosen
    snippet = "\n".join(lines[start:end]).strip()
    if not snippet:
        raise ValueError("Selected changelog section is empty")
    return snippet + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract the latest released section from CHANGELOG.md",
    )
    parser.add_argument(
        "--changelog",
        default="CHANGELOG.md",
        help="Path to changelog markdown file",
    )
    parser.add_argument(
        "--output",
        default="RELEASE_NOTES.md",
        help="Path for extracted release notes",
    )
    args = parser.parse_args()

    changelog_path = Path(args.changelog)
    output_path = Path(args.output)

    text = changelog_path.read_text(encoding="utf-8")
    notes = extract_latest_release_section(text)
    output_path.write_text(notes, encoding="utf-8")
    print(f"Wrote {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
