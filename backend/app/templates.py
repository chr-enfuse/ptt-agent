"""Template store backing the `list_templates` / `load_template` agent tools —
the Python port of backend/src/templates.ts.

Definitions come from the repo-root `templates.json` (the single source of truth
shared with the Node generator). The placeholder inventory goes to the model; the
.pptx base64 never does — it travels only in the relayed `insert_template_slides`
payload to the pane.
"""

from __future__ import annotations

import base64
import json
from functools import lru_cache
from pathlib import Path
from typing import Any

# backend-py/app/templates.py -> repo root is two parents up.
_REPO_ROOT = Path(__file__).resolve().parent.parent.parent
_TEMPLATES_JSON = _REPO_ROOT / "templates.json"
# The committed .pptx artifacts live alongside this package.
_TEMPLATES_DIR = Path(__file__).resolve().parent.parent / "templates"


class TemplateToolError(Exception):
    """Recoverable tool error surfaced back to the model as an `Error: ...` string."""


@lru_cache(maxsize=1)
def _load_defs() -> list[dict[str, Any]]:
    return json.loads(_TEMPLATES_JSON.read_text(encoding="utf-8"))


def _require_template(template_id: str) -> dict[str, Any]:
    defs = _load_defs()
    for t in defs:
        if t["id"] == template_id:
            return t
    ids = ", ".join(t["id"] for t in defs)
    raise TemplateToolError(f'unknown template_id "{template_id}" — available templates: {ids}')


def list_templates() -> dict[str, Any]:
    return {
        "templates": [
            {
                "id": t["id"],
                "name": t["name"],
                "description": t["description"],
                "slideCount": len(t["slides"]),
            }
            for t in _load_defs()
        ]
    }


def load_template(template_id: str) -> dict[str, Any]:
    """Full placeholder inventory the model plans `apply_slide_content` fills against."""
    t = _require_template(template_id)
    return {"id": t["id"], "name": t["name"], "description": t["description"], "slides": t["slides"]}


def get_template_base64(template_id: str) -> str:
    """The .pptx itself — relayed to the task pane for insertSlidesFromBase64."""
    t = _require_template(template_id)
    path = _TEMPLATES_DIR / t["file"]
    try:
        return base64.b64encode(path.read_bytes()).decode("ascii")
    except OSError as err:
        raise TemplateToolError(
            f'template file "{t["file"]}" is missing — run '
            "`python scripts/build_templates.py` in backend/"
        ) from err
