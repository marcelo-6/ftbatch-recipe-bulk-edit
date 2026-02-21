#!/usr/bin/env python3
"""Basic packaging smoke tests run against built wheel/sdist in isolated envs."""

from __future__ import annotations

import importlib
import subprocess
import sys  # noqa F401


def _check_imports() -> None:
    for module in ("main", "cli.cli", "core.parser", "utils.logging_cfg"):
        importlib.import_module(module)


def _check_cli() -> None:
    result = subprocess.run(
        ["ftbatch-bulk-edit", "--version"],
        check=True,
        capture_output=True,
        text=True,
    )
    output = result.stdout.strip() or result.stderr.strip()
    if not output:
        raise RuntimeError("ftbatch-bulk-edit --version returned empty output")


def main() -> int:
    _check_imports()
    _check_cli()
    print("smoke test passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
