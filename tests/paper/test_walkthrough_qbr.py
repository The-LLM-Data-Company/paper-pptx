"""The QBR walkthrough eval (PLAN-v0.1): the canonical template job as a permanent test.

The post-v0 gap review found its gaps by simulating a real job AFTER v0 was declared done.
This module makes that simulation permanent: `test_walkthrough_end_to_end` executes the
canonical template job — build a QBR deck from the gauntlet "corporate template" — using only
shipped public API, asserting through the contract harness at every step. Capabilities the
current wave has not shipped yet are strict-xfail step tests (the PR-0 stub mechanism): the
suite FAILS the moment an organ lands without its walkthrough step flipping, and v0.1 is not
done while any Phase 0-2 step here is an xfail.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.package import patch_save
from pptx.util import Pt

from . import corpus
from .contract import zip_member_map
from .idlists import dangling_section_slide_ids, duplicate_section_slide_ids
from .lo import lo_load_smoke
from .relint import dangling_relationship_targets, missing_relationship_references

GAUNTLET = "self_generated/gauntlet.pptx"


def _build_qbr_deck(tmp_path):
    """Run the canonical template job with shipped API; return (deck_path, presentation)."""
    template_path = tmp_path / "template.pptx"
    template_path.write_bytes(corpus.fixture_path(GAUNTLET).read_bytes())
    prs = Presentation(str(template_path))

    # -- structural pass first (template-editing doctrine): 3 content sections from the
    # -- branded slide, chart slide kept, autofit playground dropped, closing slide last
    for _ in range(2):
        prs.slides.clone(0)
    prs.slides.delete(prs.slides.index(next(s for s in prs.slides if any(
        shape.name == "autofit_normal_box" for shape in s.shapes
    ))))
    assert len(prs.slides) == 5  # -- 3 branded + chart + closing

    # -- content pass (Phase 1.1): anchored retitles — each branded clone gets its own title
    # -- via a hash-checked anchor; then one deck-wide format-preserving language pass
    from pptx.edit import replace_text, replace_text_at
    from pptx.inspect import inspect_text

    section_titles = ["Q4 Business Review", "Wins", "Risks"]
    branded = [s for s in prs.slides if s.shapes.title is not None
               and s.shapes.title.text == "Gauntlet: branded"]
    assert len(branded) == 3
    for slide, title in zip(branded, section_titles):
        anchor = next(
            b.anchor for b in inspect_text(slide).blocks if b.text == "Gauntlet: branded"
        )
        result = replace_text_at(prs, anchor, "Gauntlet: branded", title)
        assert result.replacements == 1
    language_pass = replace_text(prs, "Inherited body level", "Section content line")
    assert language_pass.replacements == 6  # -- two lines on each of the three sections
    closing_title = prs.slides[-1].shapes.title
    closing_title.text_frame.paragraphs[0].runs[0].text = "Next steps"

    # -- chart data update, by name, validated
    chart_slide = next(s for s in prs.slides if any(sh.has_chart for sh in s.shapes))
    chart = chart_slide.shapes.chart_by_name("gauntlet_chart")
    chart.replace_data_safe(["Q3", "Q4"], [("Revenue", (14.2, 16.8))], number_format="0.0")

    # -- speaker notes for the talk track
    chart_slide.replace_notes_text("Walk through Q4 revenue; flag the Alpha renewal.")

    # -- image swap, geometry preserved (same-format PNG)
    from PIL import Image as PILImage

    replacement = io.BytesIO()
    PILImage.new("RGB", (64, 64), (10, 90, 160)).save(replacement, format="PNG")
    closing = prs.slides[-1]
    picture = next(s for s in closing.shapes if s.name == "gauntlet_img_2")
    picture.replace_image(io.BytesIO(replacement.getvalue()))

    # -- normalize the one autofit frame we kept prose in (freeze what the reader sees)
    body_box = next(
        (s for s in prs.slides[0].shapes if s.name == "hyperlink_box"), None
    )
    if body_box is not None:
        tf = body_box.text_frame
        for paragraph in tf.paragraphs:
            paragraph.line_spacing = 1.0
        tf.normalize_autofit(resolve=True, min_font_size=Pt(11))
        assert tf.auto_size == MSO_AUTO_SIZE.NONE

    # -- footer pass: static text today; real slide-number fields are Phase 2.5 (xfail below)
    for index, slide in enumerate(prs.slides):
        footer = slide.shapes.add_textbox(Pt(20), Pt(520), Pt(200), Pt(20))
        footer.name = "meta_footer_%d" % (index + 1)
        footer.text_frame.paragraphs[0].add_run().text = "Q4 QBR - page %d" % (index + 1)

    # -- narrow save: only genuinely-changed parts differ from the template
    out_path = tmp_path / "qbr.pptx"
    diff = patch_save(str(template_path), prs, str(out_path))
    assert not diff.is_empty
    return out_path, prs


def _assert_package_integrity(pptx_bytes):
    zip_map = zip_member_map(pptx_bytes)
    assert dangling_relationship_targets(zip_map) == []
    assert missing_relationship_references(zip_map) == []
    assert dangling_section_slide_ids(zip_map) == []
    assert duplicate_section_slide_ids(zip_map) == []


def test_walkthrough_end_to_end(tmp_path):
    out_path, _ = _build_qbr_deck(tmp_path)
    saved = out_path.read_bytes()
    _assert_package_integrity(saved)

    reopened = Presentation(io.BytesIO(saved))
    assert len(reopened.slides) == 5
    assert reopened.slides[0].shapes.title.text == "Q4 Business Review"
    chart = next(
        sh.chart for s in reopened.slides for sh in s.shapes if sh.has_chart
    )
    assert [(s.name, tuple(s.values)) for s in chart.series] == [("Revenue", (14.2, 16.8))]
    chart_slide = next(s for s in reopened.slides if any(sh.has_chart for sh in s.shapes))
    assert chart_slide.read_notes_text() == (
        "Walk through Q4 revenue; flag the Alpha renewal."
    )


@pytest.mark.lo_smoke
def test_walkthrough_output_loads_in_libreoffice(tmp_path):
    out_path, _ = _build_qbr_deck(tmp_path)
    lo_load_smoke(out_path, tmp_path)


# ---------------------------------------------------------------- unshipped steps (strict)
# Each xfail names its PLAN-v0.1 item and FAILS the suite when the organ lands without the
# walkthrough growing the real step.


@pytest.mark.xfail(strict=True, reason="PLAN-v0.1 Phase 1.2: shape surgery")
def test_step_delete_leftover_placeholder_shape():
    from pptx.shapes.shapetree import SlideShapes

    assert callable(SlideShapes.delete)


@pytest.mark.xfail(strict=True, reason="PLAN-v0.1 Phase 1.3: generic by-name addressing")
def test_step_address_picture_by_name():
    from pptx.shapes.shapetree import SlideShapes

    assert callable(SlideShapes.picture_by_name)


@pytest.mark.xfail(strict=True, reason="PLAN-v0.1 Phase 2.1: structural deck manifest")
def test_step_survey_template_with_deck_manifest():
    from pptx.inspect import inspect_deck  # noqa: F401


@pytest.mark.xfail(strict=True, reason="PLAN-v0.1 Phase 2.2: effective shape fill/line")
def test_step_check_brand_accent_via_effective_shape_format():
    from pptx.inspect import effective_shape_format  # noqa: F401


@pytest.mark.xfail(strict=True, reason="PLAN-v0.1 Phase 2.3: cross-format image replace")
def test_step_swap_logo_across_image_formats(tmp_path):
    import inspect as stdlib_inspect

    from pptx.shapes.picture import Picture

    assert "allow_format_change" in stdlib_inspect.signature(
        Picture.replace_image
    ).parameters


@pytest.mark.xfail(strict=True, reason="PLAN-v0.1 Phase 2.4: workbook-less chart update")
def test_step_update_libreoffice_authored_chart():
    prs = Presentation(str(corpus.fixture_path("libreoffice_export/lo_chart_notes.pptx")))
    chart = prs.slides[0].shapes[1].chart if prs.slides[0].shapes[1].has_chart else (
        next(sh.chart for sh in prs.slides[0].shapes if sh.has_chart)
    )
    chart.replace_data_safe(["A", "B"], [("S1", (1.0, 2.0))])  # -- refuses today


@pytest.mark.xfail(strict=True, reason="PLAN-v0.1 Phase 2.5: real slide-number fields")
def test_step_add_real_page_number_fields():
    from pptx.text.text import _Paragraph

    assert callable(_Paragraph.add_slide_number_field)
