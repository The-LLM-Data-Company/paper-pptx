"""PR-0 stub tests (CONVENTIONS §8): one strict xfail per unimplemented organ.

Each stub imports the names API-PROPOSAL.md pins for its organ. `strict=True` means the stub
FAILS the suite the moment the organ lands (XPASS), forcing the landing phase to replace its
stub with real contract tests in the same change — and keeping this file an accurate ledger of
what remains unimplemented.
"""

from __future__ import annotations

import pytest


@pytest.mark.xfail(strict=True, reason="PR-0 stub - lands with Phase 7 (slide operations)")
def test_pr0_slide_ops_api():
    from pptx.slide import SlideClonePolicy, Slides  # noqa: F401

    for name in ("clone", "delete", "reorder", "move"):
        assert callable(getattr(Slides, name))


@pytest.mark.xfail(strict=True, reason="PR-0 stub - lands with Phase 8 (image replacement)")
def test_pr0_image_replacement_api():
    from pptx.shapes.picture import Picture

    assert callable(Picture.replace_image)


@pytest.mark.xfail(strict=True, reason="PR-0 stub - lands with Phase 9 (chart data routing)")
def test_pr0_chart_routing_api():
    from pptx.chart.chart import Chart
    from pptx.shapes.shapetree import SlideShapes

    assert callable(SlideShapes.chart_by_name)
    assert callable(Chart.replace_data_safe)
