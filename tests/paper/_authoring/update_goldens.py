"""THE explicit golden-update command (CONVENTIONS §4: goldens update only via this).

Regenerates every golden inspection JSON under `tests/paper/goldens/` from the frozen fixture
corpus. Golden diffs are human-reviewed in the PR that lands them; tests never call this.

Usage:  python tests/paper/_authoring/update_goldens.py
"""

from __future__ import annotations

import json
from pathlib import Path

from pptx import Presentation
from pptx.inspect import inspect_text

GOLDENS_DIR = Path(__file__).resolve().parent.parent / "goldens"
FIXTURES_DIR = Path(__file__).resolve().parent.parent / "fixtures"

#: (golden filename, fixture relpath, slide index) — one golden per entry
GOLDENS = (
    ("branded_template.inspect.json", "self_generated/branded_template.pptx", 0),
    ("clrmap_remap.inspect.json", "self_generated/clrmap_remap.pptx", 0),
    ("gauntlet_slide1.inspect.json", "self_generated/gauntlet.pptx", 0),
)


def golden_json(fixture_relpath: str, slide_index: int) -> str:
    """Return the canonical golden serialization for one slide's inspection."""
    prs = Presentation(str(FIXTURES_DIR / fixture_relpath))
    payload = inspect_text(prs.slides[slide_index]).to_dict()
    return json.dumps(payload, indent=2, ensure_ascii=False) + "\n"


def main() -> None:
    GOLDENS_DIR.mkdir(parents=True, exist_ok=True)
    for golden_name, fixture_relpath, slide_index in GOLDENS:
        out = GOLDENS_DIR / golden_name
        out.write_text(golden_json(fixture_relpath, slide_index), encoding="utf-8")
        print("wrote", out)


if __name__ == "__main__":
    main()
