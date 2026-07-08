"""Phase 9 contract tests: chart addressing + safe data replacement (route, don't build).

The organ is `SlideShapes.chart_by_name` plus `Chart.replace_data_safe`: full validation,
typed refusals for structures the mechanism can't honor, then routing to upstream's public
`replace_data`. The quarantined chart XML-writer module gains no code.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.errors import (
    AmbiguousTargetError,
    PaperRefusal,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from pptx.util import Inches

from . import corpus
from .contract import assert_changed_parts, assert_refusal_atomic, save_to_bytes
from .lo import lo_load_smoke

CHART_NOTES = "self_generated/chart_notes.pptx"
LO_CHART_NOTES = "libreoffice_export/lo_chart_notes.pptx"
MINIMAL = "self_generated/minimal_clean.pptx"
CHART_NAME = "clone_fixture_chart"


def _open(relpath):
    return Presentation(str(corpus.fixture_path(relpath)))


# ----------------------------------------------------------------------------- addressing


def test_chart_by_name_finds_the_frozen_fixture_chart():
    chart = _open(CHART_NOTES).slides[0].shapes.chart_by_name(CHART_NAME)
    assert chart.chart_type == XL_CHART_TYPE.COLUMN_CLUSTERED


def test_chart_by_name_reports_missing_and_wrong_kind_distinctly():
    shapes = _open(CHART_NOTES).slides[0].shapes
    with pytest.raises(TargetNotFoundError, match="no shape named"):
        shapes.chart_by_name("does_not_exist")
    with pytest.raises(TargetNotFoundError, match="holds no chart"):
        shapes.chart_by_name("Title 1")


def test_chart_by_name_refuses_to_pick_between_duplicates():
    prs = _open(CHART_NOTES)
    chart_data = CategoryChartData()
    chart_data.categories = ["a"]
    chart_data.add_series("s", (1,))
    frame = prs.slides[0].shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1), Inches(3), Inches(2), chart_data
    )
    frame.name = CHART_NAME
    with pytest.raises(AmbiguousTargetError):
        prs.slides[0].shapes.chart_by_name(CHART_NAME)


# ------------------------------------------------------------------------------ replacing


def test_replace_data_safe_round_trips_with_exact_budget():
    prs = _open(CHART_NOTES)
    chart = prs.slides[0].shapes.chart_by_name(CHART_NAME)
    before = save_to_bytes(prs)
    chart.replace_data_safe(
        ["North", "South"],
        [("FY25", (10.5, None)), ("FY26", (12, 13))],
        number_format="0.0",
    )
    after = save_to_bytes(prs)
    assert_changed_parts(
        before,
        after,
        expect_changed=[
            "ppt/charts/chart1.xml",
            "ppt/embeddings/Microsoft_Excel_Sheet1.xlsx",
        ],
    )

    reopened_chart = (
        Presentation(io.BytesIO(after)).slides[0].shapes.chart_by_name(CHART_NAME)
    )
    assert [(s.name, tuple(s.values)) for s in reopened_chart.series] == [
        ("FY25", (10.5, None)),
        ("FY26", (12.0, 13.0)),
    ]
    assert list(reopened_chart.plots[0].categories) == ["North", "South"]


@pytest.mark.parametrize(
    "bad_call",
    [
        lambda c: c.replace_data_safe([], [("a", ())]),
        lambda c: c.replace_data_safe(["x"], []),
        lambda c: c.replace_data_safe(["x"], [("a", (1, 2))]),
        lambda c: c.replace_data_safe(["x"], [("a", ("y",))]),
        lambda c: c.replace_data_safe(["x"], [("a", (True,))]),
        lambda c: c.replace_data_safe(["x"], [("a", (1,)), ("a", (2,))]),
        lambda c: c.replace_data_safe([1], [("a", (1,))]),
        lambda c: c.replace_data_safe(["x"], [("", (1,))]),
        lambda c: c.replace_data_safe(["x"], ["not-a-pair"]),
        lambda c: c.replace_data_safe(["x"], [("a", (1,))], number_format=7),
    ],
)
def test_data_shape_problems_are_valueerrors_and_leave_the_chart_untouched(bad_call):
    prs = _open(CHART_NOTES)
    chart = prs.slides[0].shapes.chart_by_name(CHART_NAME)
    before = save_to_bytes(prs)
    with pytest.raises(ValueError):
        bad_call(chart)
    assert_changed_parts(before, save_to_bytes(prs))  # -- empty budget


def test_refuses_chart_without_embedded_workbook():
    """LibreOffice charts have no c:externalData; replacing would desync XML and workbook."""
    prs = _open(LO_CHART_NOTES)
    chart_shape_name = next(s for s in prs.slides[0].shapes if s.has_chart).name

    def operation(p):
        chart = p.slides[0].shapes.chart_by_name(chart_shape_name)
        chart.replace_data_safe(["x"], [("a", (1,))])

    raised = assert_refusal_atomic(prs, operation, UnsupportedStructureError)
    assert "no embedded workbook" in str(raised)
    assert isinstance(raised, PaperRefusal)


def test_refuses_unsupported_chart_types_atomically():
    prs = _open(MINIMAL)
    xy_data = XyChartData()
    series = xy_data.add_series("s1")
    series.add_data_point(1, 2)
    series.add_data_point(3, 4)
    frame = prs.slides[0].shapes.add_chart(
        XL_CHART_TYPE.XY_SCATTER, Inches(1), Inches(1), Inches(4), Inches(3), xy_data
    )
    frame.name = "xy_chart"

    def operation(p):
        chart = p.slides[0].shapes.chart_by_name("xy_chart")
        chart.replace_data_safe(["x"], [("a", (1,))])

    raised = assert_refusal_atomic(prs, operation, UnsupportedStructureError)
    assert "not supported" in str(raised)


def test_refuses_multi_plot_combo_charts_atomically():
    """A combo chart (two plots in one plotArea) refuses rather than desyncing one plot."""
    import copy

    from pptx.oxml.ns import qn

    prs = _open(CHART_NOTES)
    chart = prs.slides[0].shapes.chart_by_name(CHART_NAME)
    plotArea = chart._chartSpace.chart.plotArea
    barChart = plotArea.find(qn("c:barChart"))
    barChart.addnext(copy.deepcopy(barChart))  # -- now a two-plot chart
    assert len(chart.plots) == 2

    def operation(p):
        p.slides[0].shapes.chart_by_name(CHART_NAME).replace_data_safe(["x"], [("a", (1,))])

    raised = assert_refusal_atomic(prs, operation, UnsupportedStructureError)
    assert "multi-plot" in str(raised)


def test_lone_surrogate_strings_are_rejected_before_any_mutation():
    """Regression: a str containing a lone surrogate passed isinstance validation, then
    exploded during serialization AFTER the chart XML had already been rewritten."""
    prs = _open(CHART_NOTES)
    chart = prs.slides[0].shapes.chart_by_name(CHART_NAME)
    before = save_to_bytes(prs)
    with pytest.raises(ValueError, match="not encodable"):
        chart.replace_data_safe(["ok"], [("bad\udc80name", (1,))])
    with pytest.raises(ValueError, match="not encodable"):
        chart.replace_data_safe(["bad\udc80cat"], [("name", (1,))])
    assert_changed_parts(before, save_to_bytes(prs))  # -- empty budget


def test_default_number_format_path_round_trips():
    prs = _open(CHART_NOTES)
    chart = prs.slides[0].shapes.chart_by_name(CHART_NAME)
    chart.replace_data_safe(["One", "Two"], [("Only", (4.5, 6.5))])
    reopened_chart = (
        Presentation(io.BytesIO(save_to_bytes(prs))).slides[0].shapes.chart_by_name(CHART_NAME)
    )
    assert [(s.name, tuple(s.values)) for s in reopened_chart.series] == [("Only", (4.5, 6.5))]


# --------------------------------------------------------------------------------- lo_smoke


@pytest.mark.lo_smoke
def test_replaced_chart_output_loads_in_libreoffice(tmp_path):
    prs = _open(CHART_NOTES)
    chart = prs.slides[0].shapes.chart_by_name(CHART_NAME)
    chart.replace_data_safe(["Alpha", "Beta", "Gamma"], [("S1", (1.5, 2.5, 3.5))])
    out = tmp_path / "chart_replaced.pptx"
    prs.save(str(out))
    lo_load_smoke(out, tmp_path)
