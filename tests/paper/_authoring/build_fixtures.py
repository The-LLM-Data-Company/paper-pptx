"""One-time authoring script for the bootstrap fixture corpus in `tests/paper/fixtures/`.

This script is development tooling, not package code and not test code. It exists so every
self-generated fixture's construction is reviewable and reproducible in *content* (regenerated
files are semantically identical but not byte-identical: zip entry timestamps differ). The
fixtures themselves are FROZEN: tests verify them against `MANIFEST.sha256` and never call this
script. Regenerating a fixture requires updating the manifest and sidecar in a reviewed PR.

Provenance produced here is honestly "self-generated" (python-pptx, this file). The
LibreOffice-export bucket is produced separately by round-tripping these files through headless
LibreOffice (see `tests/paper/fixtures/README.md`); real-PowerPoint and Google-export fixtures
require a human and are specified in `FIXTURE-REQUESTS.md`.

Usage:  python tests/paper/_authoring/build_fixtures.py
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

from lxml import etree
from PIL import Image

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from pptx.util import Inches

OUT_DIR = Path(__file__).resolve().parent.parent / "fixtures" / "self_generated"

# -- Distinctive, non-default style values so inheritance in ground truth is unambiguous.
# -- Default-template values are asserted before each edit so template drift fails loudly.
BRAND_TITLE_MASTER_SZ = "4000"  # default 4400
BRAND_TITLE_LAYOUT_SZ = "3600"  # layout-level override, no default counterpart
BRAND_BODY_L1_SZ = "2600"  # default 3200
BRAND_BODY_L2_SZ = "2200"  # default 2800
BRAND_BODY_L1_FONT = "Trebuchet MS"  # default "+mn-lt"


def _expect(condition: bool, message: str) -> None:
    """Fail loudly if a precondition about the default template no longer holds."""
    if not condition:
        raise AssertionError("fixture-authoring precondition failed: %s" % message)


def _save(prs, filename: str) -> Path:
    path = OUT_DIR / filename
    prs.save(str(path))
    return path


def _png_bytes() -> bytes:
    """Return a small deterministic 64x64 quadrant-pattern PNG."""
    img = Image.new("RGB", (64, 64))
    quadrant_colors = {
        (0, 0): (200, 30, 30),
        (1, 0): (30, 160, 60),
        (0, 1): (30, 60, 200),
        (1, 1): (220, 180, 40),
    }
    for x in range(64):
        for y in range(64):
            img.putpixel((x, y), quadrant_colors[(x // 32, y // 32)])
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def _brand_master(prs) -> None:
    """Give the master's txStyles distinctive sizes/fonts (see module constants)."""
    txStyles = prs.slide_masters[0].element.find(qn("p:txStyles"))
    _expect(txStyles is not None, "master has no p:txStyles")

    title_defRPr = txStyles.find(qn("p:titleStyle")).find(qn("a:lvl1pPr")).find(qn("a:defRPr"))
    _expect(title_defRPr.get("sz") == "4400", "default master titleStyle lvl1 sz is not 4400")
    title_defRPr.set("sz", BRAND_TITLE_MASTER_SZ)
    title_latin = title_defRPr.find(qn("a:latin"))
    _expect(
        title_latin is not None and title_latin.get("typeface") == "+mj-lt",
        "default master titleStyle font is not the +mj-lt theme reference",
    )  # -- left as "+mj-lt" deliberately: title font must resolve through the theme

    bodyStyle = txStyles.find(qn("p:bodyStyle"))
    l1_defRPr = bodyStyle.find(qn("a:lvl1pPr")).find(qn("a:defRPr"))
    _expect(l1_defRPr.get("sz") == "3200", "default master bodyStyle lvl1 sz is not 3200")
    l1_defRPr.set("sz", BRAND_BODY_L1_SZ)
    l1_latin = l1_defRPr.find(qn("a:latin"))
    _expect(l1_latin is not None, "master bodyStyle lvl1 has no a:latin")
    l1_latin.set("typeface", BRAND_BODY_L1_FONT)

    l2_defRPr = bodyStyle.find(qn("a:lvl2pPr")).find(qn("a:defRPr"))
    _expect(l2_defRPr.get("sz") == "2800", "default master bodyStyle lvl2 sz is not 2800")
    l2_defRPr.set("sz", BRAND_BODY_L2_SZ)
    l2_latin = l2_defRPr.find(qn("a:latin"))
    _expect(
        l2_latin is not None and l2_latin.get("typeface") == "+mn-lt",
        "default master bodyStyle lvl2 font is not the +mn-lt theme reference",
    )  # -- lvl2 font left as "+mn-lt" so ground truth exercises master->theme resolution


def _override_layout_title_size(layout, sz: str) -> None:
    """Add an `a:lvl1pPr/a:defRPr@sz` override to the layout title placeholder's lstStyle."""
    title_ph = next(p for p in layout.placeholders if p.placeholder_format.idx == 0)
    lstStyle = title_ph.text_frame._txBody.find(qn("a:lstStyle"))
    _expect(lstStyle is not None, "layout title placeholder has no a:lstStyle")
    _expect(len(lstStyle) == 0, "layout title placeholder lstStyle is not empty")
    lvl1pPr = etree.SubElement(lstStyle, qn("a:lvl1pPr"))
    defRPr = etree.SubElement(lvl1pPr, qn("a:defRPr"))
    defRPr.set("sz", sz)


def _remap_clrmap(prs) -> None:
    """Swap accent1<->accent2 and invert bg1/tx1 in the master's p:clrMap."""
    clrMap = prs.slide_masters[0].element.find(qn("p:clrMap"))
    defaults = {"bg1": "lt1", "tx1": "dk1", "accent1": "accent1", "accent2": "accent2"}
    for attr, expected in defaults.items():
        _expect(clrMap.get(attr) == expected, "default clrMap %s is not %s" % (attr, expected))
    clrMap.set("bg1", "dk1")
    clrMap.set("tx1", "lt1")
    clrMap.set("accent1", "accent2")
    clrMap.set("accent2", "accent1")


def _add_named_textbox(slide, name, left, top, width, height, text):
    box = slide.shapes.add_textbox(left, top, width, height)
    box.name = name
    tf = box.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].add_run().text = text
    return box


def _add_real_bullet(paragraph, char="•", font="Arial") -> None:
    """Write a real a:buFont/a:buChar pair into the paragraph's a:pPr.

    Raw-XML on purpose: bullets are exactly the gap the fork fills in Phase 2; the fixture must
    exist before the API does. a:buFont precedes a:buChar per the a:pPr child sequence.
    """
    pPr = paragraph._p.get_or_add_pPr()
    _expect(len(pPr) == 0, "paragraph pPr is not empty; bullet insert order not handled here")
    buFont = etree.SubElement(pPr, qn("a:buFont"))
    buFont.set("typeface", font)
    buChar = etree.SubElement(pPr, qn("a:buChar"))
    buChar.set("char", char)


def _set_norm_autofit_details(text_frame, font_scale="62500", ln_spc_reduction="20000") -> None:
    """Set fontScale/lnSpcReduction on an existing a:normAutofit (raw: oxml lacks the attrs)."""
    normAutofit = text_frame._txBody.find(qn("a:bodyPr")).find(qn("a:normAutofit"))
    _expect(normAutofit is not None, "text frame has no a:normAutofit element")
    normAutofit.set("fontScale", font_scale)
    normAutofit.set("lnSpcReduction", ln_spc_reduction)


# ------------------------------------------------------------------------- fixtures


def build_minimal_clean() -> Path:
    """One title slide from the default template. The minimal-clean corpus anchor."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text_frame.paragraphs[0].add_run().text = "Minimal clean fixture"
    slide.placeholders[1].text_frame.paragraphs[0].add_run().text = "One slide, no extras."
    return _save(prs, "minimal_clean.pptx")


def build_branded_template() -> Path:
    """Feature: placeholder text inheriting size/font through layout/master/theme.

    Title run: no local size/font; layout lstStyle sz=3600 overrides master titleStyle sz=4000;
    font resolves master "+mj-lt" -> theme majorFont (Calibri).
    Body run lvl0: master bodyStyle lvl1 sz=2600, explicit "Trebuchet MS".
    Body run lvl1: master bodyStyle lvl2 sz=2200; font "+mn-lt" -> theme minorFont (Calibri).
    """
    prs = Presentation()
    _brand_master(prs)
    layout = prs.slide_layouts[1]  # -- "Title and Content"
    _override_layout_title_size(layout, BRAND_TITLE_LAYOUT_SZ)

    slide = prs.slides.add_slide(layout)
    slide.shapes.title.text_frame.paragraphs[0].add_run().text = "Branded Title"
    body_tf = slide.placeholders[1].text_frame
    body_tf.paragraphs[0].add_run().text = "Body level one"
    level_two = body_tf.add_paragraph()
    level_two.level = 1
    level_two.add_run().text = "Body level two"

    for paragraph in [slide.shapes.title.text_frame.paragraphs[0]] + list(body_tf.paragraphs):
        for run in paragraph.runs:
            _expect(run.font.size is None, "fixture run must not carry a local font size")
            _expect(run.font.name is None, "fixture run must not carry a local font name")
    return _save(prs, "branded_template.pptx")


def build_clrmap_remap() -> Path:
    """Feature: master remaps theme colors via p:clrMap.

    accent1<->accent2 swapped and bg1/tx1 inverted. A rectangle filled with scheme color
    "accent1" must resolve to the theme's accent2 RGB (C0504D); text colored "tx1" must resolve
    to theme lt1 (FFFFFF).
    """
    prs = Presentation()
    _remap_clrmap(prs)
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # -- "Blank"

    rect = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1.5)
    )
    rect.name = "accent1_box"
    rect.fill.solid()
    rect.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1

    box = _add_named_textbox(
        slide, "tx1_text", Inches(1), Inches(3), Inches(4), Inches(1), "Mapped text color"
    )
    run = box.text_frame.paragraphs[0].runs[0]
    run.font.color.theme_color = MSO_THEME_COLOR.TEXT_1

    fill_val = rect.fill.fore_color._color._xClr.get("val")
    _expect(fill_val == "accent1", "rectangle fill schemeClr val is %r" % fill_val)
    text_val = run.font.color._color._xClr.get("val")
    _expect(text_val == "tx1", "run color schemeClr val is %r" % text_val)
    return _save(prs, "clrmap_remap.pptx")


def build_chart_notes() -> Path:
    """The clone fixture: native chart WITH embedded workbook, speaker notes on same slide."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # -- "Title Only"
    slide.shapes.title.text_frame.paragraphs[0].add_run().text = "Chart and notes"

    chart_data = CategoryChartData()
    chart_data.categories = ["East", "West", "Midwest"]
    chart_data.add_series("Q1", (19.2, 21.4, 16.7))
    chart_data.add_series("Q2", (22.3, 28.6, 15.2))
    frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.5), Inches(8), Inches(5), chart_data
    )
    frame.name = "clone_fixture_chart"

    slide.notes_slide.notes_text_frame.text = "Speaker notes for the clone fixture."
    return _save(prs, "chart_notes.pptx")


def build_shared_media() -> Path:
    """Two slides that deliberately share one image part (python-pptx dedups identical bytes)."""
    prs = Presentation()
    png = _png_bytes()
    for index in range(2):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        picture = slide.shapes.add_picture(io.BytesIO(png), Inches(1), Inches(1), Inches(2))
        picture.name = "shared_image_%d" % (index + 1)
    image_parts = set()
    for slide in prs.slides:
        rId = slide.shapes[0]._element.blip_rId
        _expect(rId is not None, "picture has no blip rId")
        image_parts.add(slide.part.related_part(rId).partname)
    _expect(len(image_parts) == 1, "slides do not share a single image part: %r" % image_parts)
    return _save(prs, "shared_media.pptx")


def _autofit_deck(mode_setter) -> "Presentation":
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = _add_named_textbox(
        slide,
        "autofit_box",
        Inches(1),
        Inches(1),
        Inches(4),
        Inches(2),
        "Autofit fixture text that is long enough to exercise fitting behavior.",
    )
    mode_setter(box.text_frame)
    return prs


def build_autofit_none() -> Path:
    """Feature: explicit a:noAutofit."""

    def setter(tf):
        tf.auto_size = MSO_AUTO_SIZE.NONE

    return _save(_autofit_deck(setter), "autofit_none.pptx")


def build_autofit_normal() -> Path:
    """Feature: a:normAutofit with explicit fontScale=62500 and lnSpcReduction=20000."""

    def setter(tf):
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        _set_norm_autofit_details(tf)

    return _save(_autofit_deck(setter), "autofit_normal.pptx")


def build_autofit_shape() -> Path:
    """Feature: a:spAutoFit (shape grows to fit text)."""

    def setter(tf):
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

    return _save(_autofit_deck(setter), "autofit_shape.pptx")


def _whitespace_deck(text) -> "Presentation":
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_named_textbox(slide, "whitespace_box", Inches(1), Inches(1), Inches(4), Inches(1), text)
    return prs


def build_whitespace_pair() -> "tuple[Path, Path]":
    """Two decks identical except for one trailing space inside an a:t text node.

    The Phase 5 kernel trap: any comparison that trims meaningful whitespace will call these
    equivalent. They are not.
    """
    path_a = _save(_whitespace_deck("Trailing space "), "whitespace_trailing_a.pptx")
    path_b = _save(_whitespace_deck("Trailing space"), "whitespace_trailing_b.pptx")
    return path_a, path_b


def build_gauntlet() -> Path:
    """Everything ugly combined: branding + clrMap remap + chart/notes + shared media +
    all three autofit modes + table + real bullet + cropped picture + hyperlink + empty
    placeholder."""
    prs = Presentation()
    _brand_master(prs)
    _remap_clrmap(prs)
    _override_layout_title_size(prs.slide_layouts[1], BRAND_TITLE_LAYOUT_SZ)
    png = _png_bytes()

    # -- slide 1: branded placeholders ------------------------------------------------
    s1 = prs.slides.add_slide(prs.slide_layouts[1])
    s1.shapes.title.text_frame.paragraphs[0].add_run().text = "Gauntlet: branded"
    body_tf = s1.placeholders[1].text_frame
    body_tf.paragraphs[0].add_run().text = "Inherited body level one"
    p2 = body_tf.add_paragraph()
    p2.level = 1
    p2.add_run().text = "Inherited body level two"

    # -- slide 2: chart + embedded workbook + notes + shared image --------------------
    s2 = prs.slides.add_slide(prs.slide_layouts[5])
    s2.shapes.title.text_frame.paragraphs[0].add_run().text = "Gauntlet: chart and notes"
    chart_data = CategoryChartData()
    chart_data.categories = ["Alpha", "Beta"]
    chart_data.add_series("S1", (3.5, 4.25))
    frame = s2.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.5), Inches(5), Inches(4), chart_data
    )
    frame.name = "gauntlet_chart"
    s2.notes_slide.notes_text_frame.text = "Gauntlet speaker notes."
    s2.shapes.add_picture(io.BytesIO(png), Inches(7), Inches(2), Inches(2)).name = "gauntlet_img_1"

    # -- slide 3: autofit trio + table + real bullet + cropped picture ----------------
    s3 = prs.slides.add_slide(prs.slide_layouts[6])
    none_box = _add_named_textbox(
        s3, "autofit_none_box", Inches(0.5), Inches(0.5), Inches(3), Inches(1.2), "No autofit."
    )
    none_box.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    norm_box = _add_named_textbox(
        s3, "autofit_normal_box", Inches(0.5), Inches(2), Inches(3), Inches(1.2), "Norm autofit."
    )
    norm_box.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    _set_norm_autofit_details(norm_box.text_frame)
    shape_box = _add_named_textbox(
        s3, "autofit_shape_box", Inches(0.5), Inches(3.5), Inches(3), Inches(1.2), "Sp autofit."
    )
    shape_box.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

    bullet_box = _add_named_textbox(
        s3, "real_bullet_box", Inches(4), Inches(0.5), Inches(3), Inches(1.5), "A real bullet"
    )
    # -- explicit autofit so only the three autofit_* boxes carry their signature elements
    # -- (add_textbox's default bodyPr contains a:spAutoFit, which would muddy ground truth)
    bullet_box.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    _add_real_bullet(bullet_box.text_frame.paragraphs[0])

    table_shape = s3.shapes.add_table(3, 3, Inches(4), Inches(2.5), Inches(4), Inches(2))
    table_shape.name = "gauntlet_table"
    for row in range(3):
        for col in range(3):
            cell_tf = table_shape.table.cell(row, col).text_frame
            cell_tf.paragraphs[0].add_run().text = "r%dc%d" % (row, col)

    cropped = s3.shapes.add_picture(io.BytesIO(png), Inches(8.5), Inches(0.5), Inches(1.5))
    cropped.name = "gauntlet_cropped"
    cropped.crop_left = 0.25
    cropped.crop_top = 0.1

    # -- slide 4: shared media again + hyperlink + empty body placeholder -------------
    s4 = prs.slides.add_slide(prs.slide_layouts[1])
    s4.shapes.title.text_frame.paragraphs[0].add_run().text = "Gauntlet: end"
    # -- body placeholder (idx 1) left deliberately empty --
    s4.shapes.add_picture(io.BytesIO(png), Inches(6), Inches(4), Inches(2)).name = "gauntlet_img_2"
    link_box = _add_named_textbox(
        s4, "hyperlink_box", Inches(1), Inches(4), Inches(4), Inches(0.8), "External link"
    )
    link_box.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    link_box.text_frame.paragraphs[0].runs[0].hyperlink.address = "https://example.com/paper"

    return _save(prs, "gauntlet.pptx")


def build_corrupt_dangling_sldid() -> Path:
    """Corrupt-by-construction: p:sldIdLst entry referencing a relationship that does not exist.

    Opens under python-pptx (parts load lazily) but any traversal touching the second sldId
    raises KeyError. Negative tests only.
    """
    buf = io.BytesIO()
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text_frame.paragraphs[0].add_run().text = "Corrupt fixture"
    prs.save(buf)

    source = zipfile.ZipFile(buf)
    presentation_xml = etree.fromstring(source.read("ppt/presentation.xml"))
    sldIdLst = presentation_xml.find(qn("p:sldIdLst"))
    _expect(sldIdLst is not None and len(sldIdLst) == 1, "expected exactly one p:sldId")
    dangling = etree.SubElement(sldIdLst, qn("p:sldId"))
    dangling.set("id", "9999")
    dangling.set(qn("r:id"), "rId99")

    path = OUT_DIR / "corrupt_dangling_sldid.pptx"
    with zipfile.ZipFile(str(path), "w", zipfile.ZIP_DEFLATED) as out:
        for info in source.infolist():
            if info.filename == "ppt/presentation.xml":
                out.writestr(
                    info,
                    etree.tostring(
                        presentation_xml, xml_declaration=True, encoding="UTF-8", standalone=True
                    ),
                )
            else:
                out.writestr(info, source.read(info.filename))
    return path


def build_large_smoke() -> Path:
    """Large deck for perf smoke: 120 text slides, a picture every 10th slide."""
    prs = Presentation()
    png = _png_bytes()
    for index in range(120):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text_frame.paragraphs[0].add_run().text = "Slide %d" % (index + 1)
        body_tf = slide.placeholders[1].text_frame
        body_tf.paragraphs[0].add_run().text = "First line of body content on slide %d." % (
            index + 1
        )
        body_tf.add_paragraph().add_run().text = "Second line of body content."
        if index % 10 == 0:
            slide.shapes.add_picture(io.BytesIO(png), Inches(7), Inches(4), Inches(2))
    return _save(prs, "large_smoke.pptx")


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    builders = [
        build_minimal_clean,
        build_branded_template,
        build_clrmap_remap,
        build_chart_notes,
        build_shared_media,
        build_autofit_none,
        build_autofit_normal,
        build_autofit_shape,
        build_whitespace_pair,
        build_gauntlet,
        build_corrupt_dangling_sldid,
        build_large_smoke,
    ]
    for builder in builders:
        result = builder()
        for path in result if isinstance(result, tuple) else (result,):
            with zipfile.ZipFile(str(path)) as z:
                names = z.namelist()
            print("%s  (%d bytes, %d parts)" % (path.name, path.stat().st_size, len(names)))


if __name__ == "__main__":
    main()
