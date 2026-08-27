from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Iterable


@dataclass(frozen=True)
class MergeResult:
    notes: list[dict[str, Any]]
    total_input: int
    unique_notes: int
    duplicates_removed: int
    changed_notes: int


def _identity(note: dict[str, Any]) -> str:
    uuid = str(note.get("uuid") or "").strip()
    if uuid:
        return f"uuid:{uuid}"
    note_id = note.get("id")
    if note_id not in (None, ""):
        return f"id:{note_id}"
    return "content:" + json.dumps(
        [note.get("title", ""), note.get("text", ""), note.get("created", "")],
        ensure_ascii=False,
        separators=(",", ":"),
    )


def _timestamp(note: dict[str, Any]) -> datetime:
    value = str(note.get("modified") or note.get("created") or "").strip()
    if not value:
        return datetime.min.replace(tzinfo=timezone.utc)
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return datetime.min.replace(tzinfo=timezone.utc)


def _version_signature(note: dict[str, Any]) -> str:
    fields = (
        "title",
        "text",
        "created",
        "modified",
        "active_state",
        "color_index",
        "tags",
        "text_status",
        "encrypted",
    )
    return json.dumps(
        [note.get(field) for field in fields],
        ensure_ascii=False,
        separators=(",", ":"),
    )


def merge_backups(
    backups: Iterable[tuple[str, list[dict[str, Any]]]],
) -> MergeResult:
    """Merge parsed backups and keep the newest version of each note."""
    groups: dict[str, dict[str, Any]] = {}
    total_input = 0

    for source_name, notes in backups:
        for original in notes:
            total_input += 1
            note = dict(original)
            key = _identity(note)
            signature = _version_signature(note)
            group = groups.get(key)

            if group is None:
                groups[key] = {
                    "selected": note,
                    "selected_source": source_name,
                    "selected_time": _timestamp(note),
                    "sources": [source_name],
                    "occurrences": 1,
                    "signatures": {signature},
                }
                continue

            group["occurrences"] += 1
            if source_name not in group["sources"]:
                group["sources"].append(source_name)
            group["signatures"].add(signature)

            candidate_time = _timestamp(note)
            if candidate_time > group["selected_time"]:
                group["selected"] = note
                group["selected_source"] = source_name
                group["selected_time"] = candidate_time

    merged: list[dict[str, Any]] = []
    changed_notes = 0
    for group in groups.values():
        note = dict(group["selected"])
        distinct_versions = len(group["signatures"])
        changed = distinct_versions > 1
        if changed:
            changed_notes += 1
        note["source_file"] = group["selected_source"]
        note["source_files"] = list(group["sources"])
        note["backup_occurrences"] = group["occurrences"]
        note["distinct_versions"] = distinct_versions
        note["changed_across_backups"] = changed
        merged.append(note)

    merged.sort(key=lambda note: _timestamp(note), reverse=True)
    unique_notes = len(merged)
    return MergeResult(
        notes=merged,
        total_input=total_input,
        unique_notes=unique_notes,
        duplicates_removed=total_input - unique_notes,
        changed_notes=changed_notes,
    )
