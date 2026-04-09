from __future__ import annotations

import json
import re
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal


DraftTemplate = Literal["transport", "personal_expense", "attendance"]


@dataclass(frozen=True)
class DraftUpdateResult:
    snapshot: dict[str, Any]
    replaced_previous: bool


class ThreadDraftStore:
    def __init__(self, storage_dir: Path):
        self.base_dir = storage_dir / "thread_state"
        self.base_dir.mkdir(parents=True, exist_ok=True)

    def load_snapshot(self, thread_ts: str) -> dict[str, Any]:
        snapshot_path = self._snapshot_path(thread_ts)
        if snapshot_path.exists():
            return json.loads(snapshot_path.read_text(encoding="utf-8"))
        snapshot = self._empty_snapshot(thread_ts)
        self._write_snapshot(thread_ts, snapshot)
        return snapshot

    def build_context_summary(self, thread_ts: str) -> str:
        snapshot = self.load_snapshot(thread_ts)
        lines = ["現在の thread 草稿状態です。必要なら既存草稿を更新してください。"]
        for template_type in ("transport", "personal_expense", "attendance"):
            draft = snapshot["drafts_by_template"][template_type]
            if not draft["canonical_state"]:
                continue
            lines.append(
                f"[{template_type}] status={draft['status']} pending_questions={len(draft['pending_questions'])}"
            )
            lines.append(json.dumps(draft["canonical_state"], ensure_ascii=False))
            if draft["pending_questions"]:
                lines.append("pending_questions:")
                lines.extend(f"- {question}" for question in draft["pending_questions"])
        if snapshot.get("active_template"):
            lines.append(f"active_template={snapshot['active_template']}")
        return "\n".join(lines)

    def upsert_draft(
        self,
        *,
        thread_ts: str,
        template_type: DraftTemplate,
        mode: Literal["replace", "merge"],
        draft_patch: dict[str, Any],
        pending_questions: list[str] | None = None,
    ) -> DraftUpdateResult:
        snapshot = self.load_snapshot(thread_ts)
        event = {
            "type": "draft_replaced" if mode == "replace" else "fields_updated",
            "template_type": template_type,
            "draft_patch": draft_patch,
            "pending_questions": pending_questions or [],
        }
        return self._append_and_apply_event(thread_ts=thread_ts, snapshot=snapshot, event=event)

    def record_file_generated(
        self,
        *,
        thread_ts: str,
        template_type: DraftTemplate,
        generated_file: dict[str, Any],
        canonical_state: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        snapshot = self.load_snapshot(thread_ts)
        event = {
            "type": "file_generated",
            "template_type": template_type,
            "generated_file": {
                "template_id": generated_file.get("template_id"),
                "title": generated_file.get("title"),
                "output_path": generated_file.get("output_path"),
            },
            "canonical_state": canonical_state,
        }
        return self._append_and_apply_event(thread_ts=thread_ts, snapshot=snapshot, event=event).snapshot

    def get_draft(self, thread_ts: str, template_type: DraftTemplate) -> dict[str, Any]:
        snapshot = self.load_snapshot(thread_ts)
        return deepcopy(snapshot["drafts_by_template"][template_type])

    def _append_and_apply_event(
        self,
        *,
        thread_ts: str,
        snapshot: dict[str, Any],
        event: dict[str, Any],
    ) -> DraftUpdateResult:
        log_path = self._events_path(thread_ts)
        replaced_previous = self._apply_event(snapshot, event)
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(event, ensure_ascii=False) + "\n")
        self._write_snapshot(thread_ts, snapshot)
        return DraftUpdateResult(snapshot=deepcopy(snapshot), replaced_previous=replaced_previous)

    def _apply_event(self, snapshot: dict[str, Any], event: dict[str, Any]) -> bool:
        template_type = event["template_type"]
        draft = snapshot["drafts_by_template"][template_type]
        replaced_previous = False

        if event["type"] == "draft_replaced":
            replaced_previous = bool(draft["canonical_state"])
            preserved_history = draft.get("generation_history", [])
            draft["canonical_state"] = deepcopy(event["draft_patch"])
            draft["pending_questions"] = list(event.get("pending_questions") or [])
            draft["generation_history"] = preserved_history
            draft["latest_generated_file"] = None
        elif event["type"] == "fields_updated":
            draft["canonical_state"] = self._deep_merge(draft["canonical_state"], event["draft_patch"])
            draft["pending_questions"] = list(event.get("pending_questions") or draft.get("pending_questions") or [])
        elif event["type"] == "file_generated":
            if event.get("canonical_state") is not None:
                draft["canonical_state"] = deepcopy(event["canonical_state"])
            generated_file = deepcopy(event["generated_file"])
            draft["latest_generated_file"] = generated_file
            draft.setdefault("generation_history", []).append(generated_file)
            draft["pending_questions"] = []

        draft["status"] = self._determine_status(draft)
        snapshot["active_template"] = template_type
        snapshot["event_count"] = int(snapshot.get("event_count") or 0) + 1
        return replaced_previous

    @staticmethod
    def _determine_status(draft: dict[str, Any]) -> str:
        if draft.get("pending_questions"):
            return "waiting_confirmation"
        if draft.get("latest_generated_file"):
            return "generated"
        if draft.get("canonical_state"):
            return "ready"
        return "collecting"

    @staticmethod
    def _deep_merge(existing: dict[str, Any], patch: dict[str, Any]) -> dict[str, Any]:
        merged = deepcopy(existing)
        for key, value in patch.items():
            if value is None:
                continue
            if isinstance(value, dict) and isinstance(merged.get(key), dict):
                merged[key] = ThreadDraftStore._deep_merge(merged[key], value)
            else:
                merged[key] = deepcopy(value)
        return merged

    def _empty_snapshot(self, thread_ts: str) -> dict[str, Any]:
        return {
            "thread_ts": thread_ts,
            "active_template": None,
            "event_count": 0,
            "drafts_by_template": {
                "transport": self._empty_draft("transport"),
                "personal_expense": self._empty_draft("personal_expense"),
                "attendance": self._empty_draft("attendance"),
            },
        }

    @staticmethod
    def _empty_draft(template_type: DraftTemplate) -> dict[str, Any]:
        return {
            "template_type": template_type,
            "status": "collecting",
            "canonical_state": {},
            "pending_questions": [],
            "latest_generated_file": None,
            "generation_history": [],
        }

    def _write_snapshot(self, thread_ts: str, snapshot: dict[str, Any]) -> None:
        self._snapshot_path(thread_ts).write_text(
            json.dumps(snapshot, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _snapshot_path(self, thread_ts: str) -> Path:
        return self.base_dir / f"{self._sanitize(thread_ts)}.snapshot.json"

    def _events_path(self, thread_ts: str) -> Path:
        return self.base_dir / f"{self._sanitize(thread_ts)}.jsonl"

    @staticmethod
    def _sanitize(value: str) -> str:
        return re.sub(r"[^A-Za-z0-9._-]", "_", value)
