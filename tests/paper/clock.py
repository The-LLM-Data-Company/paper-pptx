"""Frozen-clock utility (CONVENTIONS.md §4: any API that stamps dates takes an injectable clock).

The injection mechanism for package APIs is pinned at PR-0; this utility supports both plausible
call shapes (a callable, or an object with `.now()`) so tests freeze time the same way either
way.
"""

from __future__ import annotations

from datetime import datetime, timezone

#: The corpus-wide frozen instant. Deliberately has distinct non-zero components so a value
#: that "looks right" by accident (midnight, first-of-month) cannot pass.
PAPER_TEST_INSTANT = datetime(2026, 1, 2, 3, 4, 5, tzinfo=timezone.utc)


class FrozenClock:
    """A clock that always reports the same instant."""

    def __init__(self, instant: datetime = PAPER_TEST_INSTANT):
        self._instant = instant

    def now(self) -> datetime:
        return self._instant

    def __call__(self) -> datetime:
        return self._instant

    def __repr__(self) -> str:
        return "FrozenClock(%r)" % (self._instant,)
