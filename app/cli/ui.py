"""
Shared CLI UI runtime helpers for Rich progress and status output.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from rich.console import Console
from rich.progress import (
    BarColumn,
    MofNCompleteColumn,
    Progress,
    SpinnerColumn,
    TaskID,
    TextColumn,
    TimeElapsedColumn,
)


@dataclass(slots=True)
class CLIRuntimeUI:
    """
    Coordinate progress rendering and short status messages for CLI commands.
    """

    console: Console
    enable_progress: bool
    progress: Progress | None = None
    task_ids: dict[str, TaskID] = field(default_factory=dict)
    announced: set[str] = field(default_factory=set)

    def __enter__(self) -> "CLIRuntimeUI":
        if self.enable_progress:
            self.progress = Progress(
                SpinnerColumn(style="cyan"),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(bar_width=None),
                MofNCompleteColumn(),
                TimeElapsedColumn(),
                console=self.console,
                transient=False,
            )
            self.progress.start()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self.progress is not None:
            self.progress.stop()
            self.progress = None

    def status(self, message: str) -> None:
        self.console.print(f"[cyan]{message}[/cyan]")

    def success(self, message: str) -> None:
        self.console.print(f"[bold green]{message}[/bold green]")

    def warning(self, message: str) -> None:
        self.console.print(f"[yellow]{message}[/yellow]")

    def error(self, message: str) -> None:
        self.console.print(f"[bold red]{message}[/bold red]")

    def ensure_task(self, key: str, description: str, total: int = 1) -> None:
        if self.enable_progress and self.progress is not None:
            if key not in self.task_ids:
                safe_total = max(total, 1)
                self.task_ids[key] = self.progress.add_task(
                    description=description,
                    total=safe_total,
                )
            return

        if key not in self.announced:
            self.console.print(f"[cyan]• {description}[/cyan]")
            self.announced.add(key)

    def update_task(
        self,
        key: str,
        *,
        description: str | None = None,
        completed: int | None = None,
        total: int | None = None,
        advance: int | None = None,
    ) -> None:
        if not self.enable_progress or self.progress is None:
            return
        if key not in self.task_ids:
            self.ensure_task(key, description or key, total=total or 1)
        task_id = self.task_ids[key]
        update_kwargs: dict[str, Any] = {}
        if description is not None:
            update_kwargs["description"] = description
        if total is not None:
            update_kwargs["total"] = max(total, 1)
        if completed is not None:
            update_kwargs["completed"] = max(completed, 0)
        if advance is not None:
            update_kwargs["advance"] = max(advance, 0)
        self.progress.update(task_id, **update_kwargs)

    def complete_task(self, key: str, *, description: str | None = None) -> None:
        if self.enable_progress and self.progress is not None and key in self.task_ids:
            task = self.progress.tasks[self.task_ids[key]]
            done_description = description or task.description
            self.progress.update(
                self.task_ids[key],
                completed=task.total if task.total is not None else task.completed,
                description=done_description,
            )

    def on_parse_progress(self, event: str, payload: dict[str, Any]) -> None:
        if event in {"discovered", "loaded"}:
            total = int(payload.get("total", 1))
            loaded = int(payload.get("loaded", 0))
            self.ensure_task("parse", "Parsing XML files", total=total)
            self.update_task(
                "parse",
                description=f"Parsing XML files ({loaded}/{max(total, 1)})",
                completed=loaded,
                total=total,
            )
        elif event == "finished":
            loaded = int(payload.get("loaded", 0))
            total = int(payload.get("total", max(loaded, 1)))
            self.ensure_task("parse", "Parsing XML files", total=total)
            self.complete_task(
                "parse",
                description=f"Parsed XML files ({loaded}/{max(total, 1)})",
            )

    def on_import_progress(self, event: str, payload: dict[str, Any]) -> None:
        if event == "start":
            total = int(payload.get("total", 1))
            self.ensure_task("import", "Applying workbook edits", total=total)
            self.update_task(
                "import",
                description=f"Applying workbook edits (0/{max(total, 1)})",
            )
        elif event == "sheet_done":
            index = int(payload.get("index", 0))
            total = int(payload.get("total", 1))
            sheet = str(payload.get("sheet", "sheet"))
            self.ensure_task("import", "Applying workbook edits", total=total)
            self.update_task(
                "import",
                description=f"Applying workbook edits ({index}/{max(total, 1)}) - {sheet}",
                completed=index,
                total=total,
            )
        elif event == "finished":
            total = int(payload.get("total", 1))
            self.ensure_task("import", "Applying workbook edits", total=total)
            self.complete_task(
                "import",
                description=f"Applied workbook edits ({max(total, 1)}/{max(total, 1)})",
            )

    def on_write_progress(self, event: str, payload: dict[str, Any]) -> None:
        if event == "start":
            total = int(payload.get("total", 1))
            self.ensure_task("write", "Writing updated XML files", total=total)
        elif event == "file_written":
            index = int(payload.get("index", 0))
            total = int(payload.get("total", 1))
            name = str(payload.get("filename", "file"))
            self.ensure_task("write", "Writing updated XML files", total=total)
            self.update_task(
                "write",
                description=f"Writing updated XML files ({index}/{max(total, 1)}) - {name}",
                completed=index,
                total=total,
            )
        elif event == "finished":
            total = int(payload.get("total", 1))
            self.ensure_task("write", "Writing updated XML files", total=total)
            self.complete_task(
                "write",
                description=f"Wrote updated XML files ({max(total, 1)}/{max(total, 1)})",
            )
