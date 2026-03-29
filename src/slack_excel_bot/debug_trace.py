from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any


class DebugTrace:
    def __init__(self, base_dir: Path, session_key: str, timestamp_key: str):
        logs_dir = base_dir / "logs" / self._sanitize(session_key)
        logs_dir.mkdir(parents=True, exist_ok=True)
        self.path = logs_dir / f"{self._sanitize(timestamp_key)}.log"

    def write_section(self, title: str, payload: Any) -> None:
        timestamp = datetime.now().isoformat(timespec="seconds")
        with self.path.open("a", encoding="utf-8") as handle:
            handle.write(f"[{timestamp}] {title}\n")
            handle.write(self._to_text(payload))
            handle.write("\n\n")

    @staticmethod
    def _to_text(payload: Any) -> str:
        if payload is None:
            return "null"
        if hasattr(payload, "model_dump"):
            payload = payload.model_dump()
        try:
            return json.dumps(payload, ensure_ascii=False, indent=2, default=str)
        except TypeError:
            return repr(payload)

    @staticmethod
    def _sanitize(value: str) -> str:
        sanitized = re.sub(r"[^0-9A-Za-z._-]+", "_", value)
        sanitized = sanitized.strip("._")
        return sanitized or "unknown"
