#!/usr/bin/env python3
"""Read project metadata from pyproject.toml."""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

import tomllib


def load_project_metadata(pyproject_path: Path) -> dict[str, Any]:
    data = tomllib.loads(pyproject_path.read_text(encoding="utf-8"))
    project = data.get("project")
    if not isinstance(project, dict):
        raise ValueError(f"Missing [project] table in {pyproject_path}")
    return project


def main() -> int:
    parser = argparse.ArgumentParser(description="Read metadata from pyproject.toml")
    parser.add_argument(
        "--pyproject",
        default="pyproject.toml",
        help="Path to pyproject.toml",
    )
    parser.add_argument(
        "--field",
        choices=["name", "version", "description"],
        help="Print only one field value",
    )
    args = parser.parse_args()

    metadata = load_project_metadata(Path(args.pyproject))

    if args.field:
        value = metadata.get(args.field, "")
        print(value)
        return 0

    print(f"Name       : {metadata.get('name', '')}")
    print(f"Version    : {metadata.get('version', '')}")
    print(f"Description: {metadata.get('description', '')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
