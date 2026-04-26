from __future__ import annotations

import csv
import uuid
from datetime import datetime
from pathlib import Path


class AppLogger:
    def __init__(self, work_dir: Path) -> None:
        self.work_dir = work_dir
        self.log_dir = work_dir / "logs"
        self.log_dir.mkdir(exist_ok=True)
        self.run_id = ""
        self.log_file = self.log_dir / "run.log"
        self.events_file = self.log_dir / "events.csv"

    def start_run(self, mode: str) -> str:
        self.run_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:8]
        self.log_file = self.log_dir / f"run_{self.run_id}.log"
        self._write_line("INFO", None, "start", f"Run started in {mode} mode", {})
        if not self.events_file.exists():
            with self.events_file.open("w", newline="", encoding="utf-8") as handle:
                writer = csv.writer(handle)
                writer.writerow(["timestamp", "run_id", "level", "party_id", "step", "message", "context"])
        return self.run_id

    def info(self, row_id: str | None, step: str, message: str, context: dict | None = None) -> None:
        self._write_line("INFO", row_id, step, message, context or {})

    def warning(self, row_id: str | None, step: str, message: str, context: dict | None = None) -> None:
        self._write_line("WARNING", row_id, step, message, context or {})

    def error(self, row_id: str | None, step: str, message: str, context: dict | None = None) -> None:
        self._write_line("ERROR", row_id, step, message, context or {})

    def _write_line(self, level: str, row_id: str | None, step: str, message: str, context: dict) -> None:
        timestamp = datetime.now().isoformat(timespec="seconds")
        clean_context = {k: ("***" if "password" in k.lower() else v) for k, v in context.items()}
        line = f"{timestamp} [{level}] run={self.run_id} row={row_id or '-'} step={step} {message} {clean_context}\n"
        with self.log_file.open("a", encoding="utf-8") as handle:
            handle.write(line)
        with self.events_file.open("a", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            writer.writerow([timestamp, self.run_id, level, row_id or "", step, message, clean_context])
