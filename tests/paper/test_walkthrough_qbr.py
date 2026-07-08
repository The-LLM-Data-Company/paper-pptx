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

    # -- declutter the chart slide (Phase 1.2): drop its decorative picture, rel-safely
    chart_slide.shapes.delete(chart_slide.shapes.picture_by_name("gauntlet_img_1"))

    # -- image swap by name (Phase 1.3), across formats (Phase 2.3: the brand asset arrives
    # -- as a JPEG), geometry preserved
    from PIL import Image as PILImage

    replacement = io.BytesIO()
    PILImage.new("RGB", (64, 64), (10, 90, 160)).save(replacement, format="JPEG")
    closing = prs.slides[-1]
    picture = closing.shapes.picture_by_name("gauntlet_img_2")
    picture.replace_image(io.BytesIO(replacement.getvalue()), allow_format_change=True)
    # -- and pull the footer rail behind everything on the closing slide (z-order)
    closing.shapes.move(closing.shapes[-1], 0)

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

    # -- footer pass (Phase 2.5): real slide-number fields, right across reordering
    for index, slide in enumerate(prs.slides):
        footer = slide.shapes.add_textbox(Pt(20), Pt(520), Pt(200), Pt(20))
        footer.name = "meta_footer_%d" % (index + 1)
        paragraph = footer.text_frame.paragraphs[0]
        paragraph.add_run().text = "Q4 QBR - page "
        paragraph.add_slide_number_field()
    prs.slide_masters[0].header_footers.slide_number_visible = True

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


def test_step_survey_template_with_deck_manifest():
    """Phase 2.1 step (was xfail): the job starts with a structural survey, not guesswork."""
    from pptx.inspect import inspect_deck

    manifest = inspect_deck(Presentation(str(corpus.fixture_path(GAUNTLET))))
    assert manifest.slide_count == 4
    chart_slides = [
        slide.part for slide in manifest.slides
        if any(shape.chart for shape in slide.shapes)
    ]
    assert chart_slides == ["/ppt/slides/slide2.xml"]


def test_step_check_brand_accent_via_effective_shape_format():
    """Phase 2.2 step (was xfail): 'is that box actually brand-colored?' is now answerable."""
    from pptx.inspect import effective_shape_format

    prs = Presentation(str(corpus.fixture_path("self_generated/clrmap_remap.pptx")))
    rect = prs.slides[0].shapes.shape_by_name("accent1_box")
    fmt = effective_shape_format(rect)
    assert fmt.fill_rgb.value == "C0504D"  # -- accent1 through the remapped clrMap
    assert fmt.fill_rgb.resolved is True


def test_step_update_libreoffice_authored_chart():
    """Phase 2.4 step (was xfail): the externally-produced deck's chart takes new numbers."""
    prs = Presentation(str(corpus.fixture_path("libreoffice_export/lo_chart_notes.pptx")))
    chart = next(sh.chart for sh in prs.slides[0].shapes if sh.has_chart)
    chart.replace_data_safe(["A", "B"], [("S1", (1.0, 2.0))])
    reopened = Presentation(io.BytesIO(_bytes_of(prs)))
    reopened_chart = next(sh.chart for sh in reopened.slides[0].shapes if sh.has_chart)
    assert [(s.name, tuple(s.values)) for s in reopened_chart.series] == [("S1", (1.0, 2.0))]


def _bytes_of(prs):
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


