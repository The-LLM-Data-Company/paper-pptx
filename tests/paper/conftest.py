"""Shared pytest configuration for the paper test suite (tests/paper only)."""

from __future__ import annotations

import pytest

from pptx.opc.package import PartFactory

from .clock import FrozenClock

# -- Captured at collection time, i.e. before any test has run, so this is the canonical
# -- import-time registration state established by pptx/__init__.py.
_CANONICAL_PART_TYPES = dict(PartFactory.part_type_for)


def pytest_configure(config):
    # -- registered here (not in pyproject.toml) so no upstream config file changes; the
    # -- upstream suite runs warnings-as-errors and an unregistered mark would be an error.
    config.addinivalue_line(
        "markers",
        "lo_smoke: independent-loader smoke via headless LibreOffice; skipped when soffice "
        "is unavailable",
    )


@pytest.fixture(autouse=True)
def _canonical_part_factory_registrations():
    """Repair `PartFactory.part_type_for` before every paper test.

    Upstream's own unit tests mutate this class-level registry without restoring it
    (`DescribePartFactory` in tests/opc/test_package.py leaves a mock class registered for the
    slide-part content type), so in a combined run every presentation opened after that test
    loads its slide parts as MagicMocks. Paper tests are integration tests over real files and
    must always see the canonical import-time registrations, whatever ran before them.
    """
    PartFactory.part_type_for.clear()
    PartFactory.part_type_for.update(_CANONICAL_PART_TYPES)


@pytest.fixture
def frozen_clock():
    """A FrozenClock pinned to `tests.paper.clock.PAPER_TEST_INSTANT`."""
    return FrozenClock()
