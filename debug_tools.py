from __future__ import annotations

import json
import os
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional


def _truthy(v: str | None) -> bool:
    if v is None:
        return False
    return v.strip().lower() in {"1","true","yes","y","on"}


@dataclass
class DebugSettings:
    enabled: bool = False
    # Write a verbose JSON report (wrap decisions, packing decisions)
    write_json_report: bool = True
    # Draw visual guides on slides (rectangles + caption). Use sparingly.
    draw_guides: bool = False
    # Print to console
    print_console: bool = True
    # Also write a text log file next to output
    write_text_log: bool = True

    width_safety: float = 0.97
    height_safety: float = 0.98

    @staticmethod
    def from_env() -> "DebugSettings":
        s = DebugSettings()
        s.enabled = True  # ← FORCE DEBUG ON
        s.draw_guides = False  # change to True if you want visual boxes
        s.write_json_report = True
        s.print_console = True
        s.write_text_log = True
        return s
        # Toggle all debugging:
        #   WS_DEBUG=1
        # Optional:
        #   WS_DEBUG_GUIDES=1
        #   WS_DEBUG_JSON=0
        #   WS_DEBUG_PRINT=0
        #   WS_DEBUG_LOG=0
        s = DebugSettings()
        s.enabled = _truthy(os.getenv("WS_DEBUG"))
        s.draw_guides = _truthy(os.getenv("WS_DEBUG_GUIDES"))
        if os.getenv("WS_DEBUG_JSON") is not None:
            s.write_json_report = _truthy(os.getenv("WS_DEBUG_JSON"))
        if os.getenv("WS_DEBUG_PRINT") is not None:
            s.print_console = _truthy(os.getenv("WS_DEBUG_PRINT"))
        if os.getenv("WS_DEBUG_LOG") is not None:
            s.write_text_log = _truthy(os.getenv("WS_DEBUG_LOG"))
        return s


@dataclass
class DebugRecorder:
    settings: DebugSettings
    output_path: Optional[Path] = None
    started_ts: float = field(default_factory=time.time)
    lines: List[str] = field(default_factory=list)
    report: Dict[str, Any] = field(default_factory=lambda: {"version": 1, "runs": []})

    def _stamp(self) -> str:
        return time.strftime("%Y-%m-%d %H:%M:%S")

    def log(self, msg: str) -> None:
        if not self.settings.enabled:
            return
        line = f"[{self._stamp()}] {msg}"
        self.lines.append(line)
        if self.settings.print_console:
            print(line)

    def start_run(self, run_kind: str, template_path: str, output_path: str) -> None:
        if not self.settings.enabled:
            return
        self.output_path = Path(output_path)
        self.report.setdefault("runs", []).append({
            "kind": run_kind,
            "template_path": template_path,
            "output_path": output_path,
            "started_at": self._stamp(),
            "slides": []
        })
        self.log(f"DEBUG ENABLED ({run_kind})")
        self.log(f"Template: {template_path}")
        self.log(f"Output:   {output_path}")

    def _cur_run(self) -> Optional[Dict[str, Any]]:
        if not self.settings.enabled:
            return None
        runs = self.report.get("runs") or []
        return runs[-1] if runs else None

    def add_slide_record(self, slide_rec: Dict[str, Any]) -> None:
        run = self._cur_run()
        if not run:
            return
        run["slides"].append(slide_rec)

    def flush(self) -> None:
        if not self.settings.enabled or not self.output_path:
            return
        out_dir = self.output_path.parent
        stem = self.output_path.stem

        if self.settings.write_text_log:
            (out_dir / f"{stem}_debug.log").write_text("\n".join(self.lines) + "\n", encoding="utf-8")

        if self.settings.write_json_report:
            (out_dir / f"{stem}_debug.json").write_text(json.dumps(self.report, indent=2), encoding="utf-8")
