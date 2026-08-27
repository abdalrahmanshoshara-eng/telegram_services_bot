from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from jinja2 import Environment, FileSystemLoader, select_autoescape

TEMPLATE_DIR = Path(__file__).resolve().parent / "templates"
_env = Environment(
    loader=FileSystemLoader(TEMPLATE_DIR),
    autoescape=select_autoescape(["html", "xml"]),
)


def generate_html(notes: list[dict[str, Any]], source_name: str) -> bytes:
    template = _env.get_template("export_viewer.html")
    notes_json = json.dumps(notes, ensure_ascii=False, separators=(",", ":")).replace(
        "</", "<\\/"
    )
    html = template.render(
        notes_json=notes_json,
        source_name=source_name or "ColorNote",
    )
    return html.encode("utf-8")
