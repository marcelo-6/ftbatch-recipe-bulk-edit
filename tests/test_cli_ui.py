from __future__ import annotations

from io import StringIO

from rich.console import Console

from cli.ui import CLIRuntimeUI


def test_no_progress_announces_phase_once() -> None:
    buffer = StringIO()
    console = Console(file=buffer, force_terminal=False, color_system=None)

    with CLIRuntimeUI(console=console, enable_progress=False) as ui:
        ui.ensure_task("parse", "Parsing XML files", total=3)
        ui.ensure_task("parse", "Parsing XML files", total=3)
        ui.on_parse_progress("loaded", {"loaded": 1, "total": 3})
        ui.complete_task("parse")

    output = buffer.getvalue()
    assert output.count("Parsing XML files") == 1


def test_progress_callbacks_run_without_exceptions() -> None:
    buffer = StringIO()
    console = Console(file=buffer, force_terminal=True, width=120)

    with CLIRuntimeUI(console=console, enable_progress=True) as ui:
        ui.on_parse_progress("discovered", {"loaded": 0, "total": 4})
        ui.on_parse_progress("loaded", {"loaded": 2, "total": 4})
        ui.on_parse_progress("finished", {"loaded": 4, "total": 4})

        ui.on_import_progress("start", {"total": 2})
        ui.on_import_progress("sheet_done", {"index": 1, "total": 2, "sheet": "A"})
        ui.on_import_progress("sheet_done", {"index": 2, "total": 2, "sheet": "B"})
        ui.on_import_progress("finished", {"total": 2})

        ui.on_write_progress("start", {"total": 2})
        ui.on_write_progress("file_written", {"index": 1, "total": 2, "filename": "x"})
        ui.on_write_progress("file_written", {"index": 2, "total": 2, "filename": "y"})
        ui.on_write_progress("finished", {"total": 2})

    output = buffer.getvalue()
    assert "Parsed XML files (4/4)" in output
    assert "Applied workbook edits (2/2)" in output
    assert "Wrote updated XML files (2/2)" in output
