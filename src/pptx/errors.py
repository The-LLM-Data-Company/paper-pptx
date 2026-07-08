"""Typed safe-refusal exceptions for paper-pptx APIs.

A `PaperRefusal` means the operation declined to run and the document — the in-memory XML tree
and any file on disk — is exactly as it was. Refusals are a success mode: the API judged the
requested change unsafe or unsupported and said so instead of guessing. Callers can therefore
catch `PaperRefusal` distinctly from bugs; programmer errors (bad argument types or values)
remain `TypeError`/`ValueError` as usual.

Every paper mutating API is structured validate-fully-then-mutate, so a refusal can never
leave a partial edit behind.
"""

from __future__ import annotations

__all__ = [
    "AmbiguousTargetError",
    "BoundaryViolationError",
    "PaperRefusal",
    "RelationshipPolicyError",
    "StaleAnchorError",
    "TargetNotFoundError",
    "UnsupportedStructureError",
]


class PaperRefusal(Exception):
    """Base class for all safe refusals raised by paper-pptx APIs."""


class AmbiguousTargetError(PaperRefusal):
    """The addressing given matches more than one target; refusing to pick one."""


class TargetNotFoundError(PaperRefusal):
    """The addressing given matches nothing in this document."""


class StaleAnchorError(TargetNotFoundError):
    """The block at an anchor's position no longer matches the anchor's content hash.

    The document changed since the anchor was produced. Refusing beats guessing: use
    `pptx.edit.refind()` to recover a fresh anchor explicitly. (Subclass of
    |TargetNotFoundError| so existing handlers keep working.)
    """


class UnsupportedStructureError(PaperRefusal):
    """The document contains structure this API cannot operate on safely."""


class BoundaryViolationError(PaperRefusal):
    """The operation would cross a boundary it promised to stay inside."""


class RelationshipPolicyError(PaperRefusal):
    """The relationship graph cannot be honored under the requested policy."""
