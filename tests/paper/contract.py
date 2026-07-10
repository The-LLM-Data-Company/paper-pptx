"""Contract-harness utilities implementing the five contract assertions.

Every mutating paper API must pass, on the relevant fixtures:

1. Save → reopen (never assert on the in-memory object): `save_reopen()`.
2. Intended effect present in the reopened document: caller asserts on the reopened object.
3. Changed-part budget: `diff_zip_members()` / `assert_changed_parts()` — the package diff
   between two outputs shows exactly the expected members changed and nothing else.
4. Independent-loader smoke: `tests.paper.lo` (marked `lo_smoke`, skipped without soffice).
5. Refusal atomicity: `assert_refusal_atomic()` (in-memory tree unchanged) and
   `assert_file_bytes_unchanged()` (files on disk unchanged), plus the typed refusal.

Comparison here is deliberately BYTE-level per zip member. Byte-compare can never call two
different parts "the same", so it can never mask a real change; semantic XML comparison is the
package kernel's job and is not test infrastructure. Because upstream `save()` stamps wall-clock
times into zip *entry headers* (not member bytes), budgets must always compare two save outputs
member-by-member — never whole-file bytes, and never a save output against the original input
file (re-serialization makes every XML part differ from the original bytes).
"""

from __future__ import annotations

import io
import zipfile
from contextlib import contextmanager
from pathlib import Path
from typing import TYPE_CHECKING, Callable, Dict, Sequence, Tuple, Type, Union

import pytest

from pptx import Presentation as open_presentation

if TYPE_CHECKING:
    from pptx.presentation import Presentation


def save_to_bytes(prs: Presentation) -> bytes:
    """Return the full .pptx package bytes of `prs` as saved by the API under test."""
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def save_reopen(prs: Presentation) -> Presentation:
    """Save `prs` and reopen it from the saved bytes (assertion 1: the reopen rule).

    Assert on the returned object, never on `prs`: an edit that landed in the in-memory tree
    but never reached disk is the classic silent failure this rule exists to catch.
    """
    return open_presentation(io.BytesIO(save_to_bytes(prs)))


def zip_member_map(pptx_bytes: bytes) -> Dict[str, bytes]:
    """Return {member-name: member-bytes} for every member of the zip package.

    Fails loudly on duplicate member names: OPC forbids them, zip readers disagree on which
    copy wins, and a dict would silently keep only one — masking exactly the
    duplicate-partname corruption (e.g. a clone that forgot to rename) that budget assertions
    exist to catch.
    """
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as zipf:
        names = zipf.namelist()
        duplicates = sorted({n for n in names if names.count(n) > 1})
        assert not duplicates, "package contains duplicate zip member names: %r" % duplicates
        return {name: zipf.read(name) for name in names}


class PartsDiff:
    """Byte-level member diff between two .pptx packages. Immutable value object."""

    def __init__(
        self, added: Sequence[str], removed: Sequence[str], changed: Sequence[str]
    ):
        self.added = tuple(sorted(added))
        self.removed = tuple(sorted(removed))
        self.changed = tuple(sorted(changed))

    @property
    def is_empty(self) -> bool:
        return not (self.added or self.removed or self.changed)

    def describe(self) -> str:
        return "added=%r removed=%r changed=%r" % (
            list(self.added),
            list(self.removed),
            list(self.changed),
        )

    def __repr__(self) -> str:
        return "PartsDiff(%s)" % self.describe()


def diff_zip_members(a_bytes: bytes, b_bytes: bytes) -> PartsDiff:
    """Return the byte-level member diff between package `a_bytes` and package `b_bytes`.

    Covers EVERY zip member — parts, `.rels` items, and `[Content_Types].xml` — so
    relationship or content-type churn cannot hide from a changed-part budget.
    """
    a_map, b_map = zip_member_map(a_bytes), zip_member_map(b_bytes)
    added = [n for n in b_map if n not in a_map]
    removed = [n for n in a_map if n not in b_map]
    changed = [n for n in a_map if n in b_map and a_map[n] != b_map[n]]
    return PartsDiff(added, removed, changed)


def assert_changed_parts(
    before_bytes: bytes,
    after_bytes: bytes,
    expect_changed: Sequence[str] = (),
    expect_added: Sequence[str] = (),
    expect_removed: Sequence[str] = (),
) -> None:
    """Assert the member diff is EXACTLY the expected budget (assertion 3), no more, no less."""
    diff = diff_zip_members(before_bytes, after_bytes)
    expected = PartsDiff(expect_added, expect_removed, expect_changed)
    assert (diff.added, diff.removed, diff.changed) == (
        expected.added,
        expected.removed,
        expected.changed,
    ), "changed-part budget violated:\n  actual:   %s\n  expected: %s" % (
        diff.describe(),
        expected.describe(),
    )


def snapshot_parts(prs: Presentation) -> Dict[str, bytes]:
    """Serialize every part and relationship collection of `prs`'s package to bytes.

    Keys are partnames, plus `<partname>::rels` for each part's relationships and
    `::package-rels` for the package-level relationships. Used to prove refusal atomicity of
    in-memory state at part granularity.
    """
    package = prs.part.package
    snapshot = {"::package-rels": package._rels.xml}
    for part in package.iter_parts():
        snapshot[str(part.partname)] = part.blob
        # -- _content_type read directly: the public property is a lazyproperty whose cache
        # -- would hide a mutation between two snapshots. Save-visible via [Content_Types].xml.
        snapshot[str(part.partname) + "::content-type"] = part._content_type.encode("utf-8")
        if len(part.rels):
            snapshot[str(part.partname) + "::rels"] = part.rels.xml
    return snapshot


def assert_refusal_atomic(
    prs: Presentation,
    operation: Callable[[Presentation], object],
    expected_exception: Union[Type[BaseException], Tuple[Type[BaseException], ...]],
) -> BaseException:
    """Assert `operation(prs)` raises `expected_exception` AND leaves `prs` untouched.

    "Untouched" is part-serialization equality over every part, its relationships, and the
    package relationships (assertion 5, in-memory half). Returns the raised exception so the
    caller can additionally assert on its message.
    """
    before = snapshot_parts(prs)
    with pytest.raises(expected_exception) as exc_info:
        operation(prs)
    after = snapshot_parts(prs)
    all_keys = set(before) | set(after)
    dirty = sorted(k for k in all_keys if before.get(k) != after.get(k))
    assert not dirty, (
        "refused operation left the in-memory package dirty; differing keys: %r" % dirty
    )
    return exc_info.value


@contextmanager
def assert_file_bytes_unchanged(*paths: "Union[str, Path]"):
    """Context manager asserting every file in `paths` is byte-identical on exit.

    The on-disk half of refusal atomicity (assertion 5): wrap the refused operation; any write
    that reaches any of the input files is a failure. Files must exist on entry.

    The checks run in a `finally` block so they execute even when the wrapped operation raises
    — which a refused operation always does. (An AssertionError raised here supersedes the
    in-flight refusal exception, which is the correct loud outcome for a dirty refusal.)
    """
    resolved = [Path(p) for p in paths]
    before = [p.read_bytes() for p in resolved]
    try:
        yield
    finally:
        for path, original in zip(resolved, before):
            assert path.exists(), "input file %s disappeared during a refused operation" % path
            assert path.read_bytes() == original, (
                "input file %s was modified by a refused operation" % path
            )
