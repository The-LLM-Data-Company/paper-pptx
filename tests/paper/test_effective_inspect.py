"""Phase 4 contract tests: effective-style inspection (`pptx.inspect`).

Sidecar-driven against the branded-template and clrMap fixtures, determinism-goldened, and
independently cross-checked against LibreOffice's own resolution of the same source deck.
The API is read-only: proven here by part-snapshot equality around every call.
"""

from __future__ import annotations

import hashlib
import json

import pytest
from lxml import etree

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.errors import UnsupportedStructureError
from pptx.inspect import content_hash, effective_font, inspect_text
from pptx.util import Emu

from . import corpus
from .contract import snapshot_parts

BRANDED = "self_generated/branded_template.pptx"
CLRMAP = "self_generated/clrmap_remap.pptx"
GAUNTLET = "self_generated/gauntlet.pptx"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _open(relpath):
    return Presentation(str(corpus.fixture_path(relpath)))


def _ground_truth(relpath):
    return corpus.load_sidecar(relpath)["ground_truth"]


def _supplied_levels(effective_value):
    return [step.level for step in effective_value.provenance if step.supplied]


# ------------------------------------------------- sidecar-driven resolution (branded deck)


def test_title_size_resolves_through_layout_override():
    expected = _ground_truth(BRANDED)["expected_effective"]["title_run"]
    font = _open(BRANDED).slides[0].shapes.title.text_frame.paragraphs[0].runs[0].effective_font()
    assert font.size.resolved
    assert font.size.value_pt == expected["size_centipoints"] / 100.0
    assert font.size.value == Emu(int(expected["size_centipoints"] * 127))
    assert _supplied_levels(font.size) == ["layout placeholder lstStyle lvl1"]


def test_title_name_resolves_through_master_theme_reference():
    expected = _ground_truth(BRANDED)["expected_effective"]["title_run"]
    font = _open(BRANDED).slides[0].shapes.title.text_frame.paragraphs[0].runs[0].effective_font()
    assert font.name.resolved
    assert font.name.value == expected["font_name"]
    assert _supplied_levels(font.name) == ["theme fontScheme majorFont"]
    # -- the +mj-lt reference itself is recorded on the master step of the chain
    master_steps = [s for s in font.name.provenance if "txStyles titleStyle" in s.level]
    assert master_steps
    assert "+mj-lt" in master_steps[0].detail


def test_body_level_one_resolves_from_master_body_style():
    expected = _ground_truth(BRANDED)["expected_effective"]["body_paragraph_0_run"]
    font = (
        _open(BRANDED).slides[0].placeholders[1].text_frame.paragraphs[0].runs[0].effective_font()
    )
    assert font.size.value_pt == expected["size_centipoints"] / 100.0
    assert font.name.value == expected["font_name"]  # -- "Trebuchet MS", explicit at master
    assert _supplied_levels(font.size) == ["master txStyles bodyStyle lvl1"]
    assert _supplied_levels(font.name) == ["master txStyles bodyStyle lvl1"]


def test_body_level_two_resolves_size_from_master_and_name_from_theme():
    expected = _ground_truth(BRANDED)["expected_effective"]["body_paragraph_1_run"]
    font = (
        _open(BRANDED).slides[0].placeholders[1].text_frame.paragraphs[1].runs[0].effective_font()
    )
    assert font.size.value_pt == expected["size_centipoints"] / 100.0
    assert _supplied_levels(font.size) == ["master txStyles bodyStyle lvl2"]
    assert font.name.value == expected["font_name"]  # -- Calibri via +mn-lt
    assert _supplied_levels(font.name) == ["theme fontScheme minorFont"]


def test_every_consulted_level_appears_in_provenance_in_walk_order():
    font = (
        _open(BRANDED).slides[0].placeholders[1].text_frame.paragraphs[1].runs[0].effective_font()
    )
    levels = [step.level for step in font.size.provenance]
    assert levels == [
        "run rPr",
        "paragraph defRPr",
        "shape lstStyle lvl2",
        "layout placeholder lstStyle lvl2",
        "master placeholder lstStyle lvl2",
        "master txStyles bodyStyle lvl2",
    ]


# ------------------------------------------------------------- scheme-color resolution


def test_scheme_color_resolves_through_remapped_clrmap():
    expected = _ground_truth(CLRMAP)["expected_resolution"]
    prs = _open(CLRMAP)
    box = next(s for s in prs.slides[0].shapes if s.name == "tx1_text")
    font = box.text_frame.paragraphs[0].runs[0].effective_font()
    assert font.color_rgb.resolved
    assert font.color_rgb.value == expected["text_run_rgb"]
    details = [step.detail for step in font.color_rgb.provenance]
    assert 'schemeClr val="tx1"' in details[0]
    assert 'tx1="lt1"' in details  # -- the clrMap remap step is visible
    assert _supplied_levels(font.color_rgb) == ["theme clrScheme lt1"]


def test_slide_clrmapovr_override_beats_master_clrmap():
    prs = _open(CLRMAP)
    slide = prs.slides[0]
    clrMapOvr = slide._element.get_or_add_clrMapOvr()
    override = etree.SubElement(clrMapOvr, "{%s}overrideClrMapping" % _A)
    master_map = dict(prs.slide_masters[0].element.find(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}clrMap").attrib)
    master_map["tx1"] = "accent3"  # -- 9BBB59 in the default theme
    for key, value in master_map.items():
        override.set(key, value)

    box = next(s for s in slide.shapes if s.name == "tx1_text")
    font = box.text_frame.paragraphs[0].runs[0].effective_font()
    assert font.color_rgb.value == "9BBB59"
    assert any(step.level == "slide clrMapOvr" for step in font.color_rgb.provenance)


def test_direct_srgb_color_supplies_immediately():
    prs = _open(BRANDED)
    run = prs.slides[0].shapes.title.text_frame.paragraphs[0].runs[0]
    run.font.color.rgb = RGBColor(0x12, 0x34, 0xAB)
    font = run.effective_font()
    assert font.color_rgb.value == "1234AB"
    assert _supplied_levels(font.color_rgb) == ["run rPr"]


# ----------------------------------------------------- non-placeholder + honesty cases


def test_plain_textbox_resolves_via_presentation_default_text_style():
    prs = _open("self_generated/minimal_clean.pptx")
    box = prs.slides[0].shapes.add_textbox(0, 0, 914400, 914400)
    box.text_frame.paragraphs[0].add_run().text = "plain"
    font = box.text_frame.paragraphs[0].runs[0].effective_font()
    assert (font.size.value_pt, font.name.value, font.color_rgb.value) == (
        18.0,
        "Calibri",
        "000000",
    )
    assert _supplied_levels(font.size) == ["presentation defaultTextStyle lvl1"]


def test_gradient_text_fill_reports_unresolved_not_a_guess():
    prs = _open("self_generated/minimal_clean.pptx")
    box = prs.slides[0].shapes.add_textbox(0, 0, 914400, 914400)
    run = box.text_frame.paragraphs[0].add_run()
    run.text = "gradient"
    rPr = run._r.get_or_add_rPr()
    etree.SubElement(rPr, "{%s}gradFill" % _A)
    font = box.text_frame.paragraphs[0].runs[0].effective_font()
    assert not font.color_rgb.resolved
    assert font.color_rgb.value is None
    assert "gradFill" in font.color_rgb.provenance[-1].detail


def test_phclr_scheme_token_reports_unresolved():
    prs = _open("self_generated/minimal_clean.pptx")
    box = prs.slides[0].shapes.add_textbox(0, 0, 914400, 914400)
    run = box.text_frame.paragraphs[0].add_run()
    run.text = "phClr"
    rPr = run._r.get_or_add_rPr()
    solidFill = etree.SubElement(rPr, "{%s}solidFill" % _A)
    schemeClr = etree.SubElement(solidFill, "{%s}schemeClr" % _A)
    schemeClr.set("val", "phClr")
    font = run.effective_font()
    assert not font.color_rgb.resolved
    assert any("unmappable" in step.detail for step in font.color_rgb.provenance)


def test_table_cell_run_refuses_instead_of_guessing():
    prs = _open(GAUNTLET)
    table_shape = next(s for s in prs.slides[2].shapes if s.name == "gauntlet_table")
    cell_run = table_shape.table.cell(0, 0).text_frame.paragraphs[0].runs[0]
    with pytest.raises(UnsupportedStructureError):
        effective_font(cell_run)


# ---------------------------------------------------------------- read-only + determinism


def test_inspection_never_mutates_the_package():
    prs = _open(BRANDED)
    before = snapshot_parts(prs)
    inspect_text(prs.slides[0])
    for shape in prs.slides[0].shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.effective_font()
    assert snapshot_parts(prs) == before


def test_inspection_is_deterministic_across_runs_and_loads():
    first = json.dumps(inspect_text(_open(BRANDED).slides[0]).to_dict())
    second = json.dumps(inspect_text(_open(BRANDED).slides[0]).to_dict())
    assert first == second


@pytest.mark.parametrize(
    ("golden_name", "fixture_relpath", "slide_index"),
    [
        ("branded_template.inspect.json", BRANDED, 0),
        ("clrmap_remap.inspect.json", CLRMAP, 0),
        ("gauntlet_slide1.inspect.json", GAUNTLET, 0),
    ],
)
def test_inspection_matches_frozen_golden(golden_name, fixture_relpath, slide_index):
    """Byte-identical to the reviewed golden; update ONLY via update_goldens.py + PR review."""
    golden_path = corpus.FIXTURES_DIR.parent / "goldens" / golden_name
    payload = inspect_text(_open(fixture_relpath).slides[slide_index]).to_dict()
    actual = (json.dumps(payload, indent=2, ensure_ascii=False) + "\n").encode("utf-8")
    assert actual == golden_path.read_bytes()  # -- byte-exact, no newline translation


def test_libreoffice_independently_confirms_branded_effective_sizes():
    """LO resolved the same inheritance when it baked run sizes into its re-export."""
    ours = sorted(
        run.font.size.value_pt
        for block in inspect_text(_open(BRANDED).slides[0]).blocks
        for run in block.runs
    )
    lo_prs = Presentation(str(corpus.fixture_path("libreoffice_export/lo_branded_template.pptx")))
    theirs = sorted(
        r.effective_font().size.value_pt
        for s in lo_prs.slides[0].shapes
        if s.has_text_frame
        for p in s.text_frame.paragraphs
        for r in p.runs
    )
    assert ours == theirs == [22.0, 26.0, 36.0]


# --------------------------------------------------------------------------------- anchors


def test_content_hash_is_pinned_sha256_nfc_prefix():
    assert content_hash("Branded Title") == (
        hashlib.sha256("Branded Title".encode("utf-8")).hexdigest()[:8]
    )


def test_content_hash_applies_nfc_normalization():
    composed = "café"  # -- é as one code point
    decomposed = "café"  # -- e + combining acute
    assert composed != decomposed
    assert content_hash(composed) == content_hash(decomposed)


def test_out_of_schema_indent_level_refuses_instead_of_crashing():
    prs = _open(BRANDED)
    run = prs.slides[0].shapes.title.text_frame.paragraphs[0].runs[0]
    pPr = run._r.getparent().find("{%s}pPr" % _A)
    if pPr is None:
        pPr = etree.SubElement(run._r.getparent(), "{%s}pPr" % _A)
        run._r.getparent().insert(0, pPr)
    pPr.set("lvl", "9")  # -- outside ST_TextIndentLevelType's 0..8
    with pytest.raises(UnsupportedStructureError, match="lvl=9"):
        run.effective_font()


def test_effective_font_payload_carries_pinned_schema_keys():
    payload = (
        _open(BRANDED).slides[0].shapes.title.text_frame.paragraphs[0].runs[0]
        .effective_font()
        .to_dict()
    )
    assert payload["schema"] == "paper-effective-font"
    assert payload["version"] == 1


def test_content_hash_treats_whitespace_as_content():
    assert content_hash("Trailing space ") != content_hash("Trailing space")


def test_blocks_carry_stable_anchors_in_sptree_order():
    inspection = inspect_text(_open(BRANDED).slides[0])
    assert [block.anchor.block_index for block in inspection.blocks] == list(
        range(len(inspection.blocks))
    )
    title_block = inspection.blocks[0]
    assert title_block.anchor.part == "/ppt/slides/slide1.xml"
    assert title_block.anchor.content_hash == content_hash("Branded Title")
    assert title_block.placeholder_type == "TITLE"
