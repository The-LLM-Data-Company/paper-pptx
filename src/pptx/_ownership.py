"""Attachment checks for live shape proxies used by paper-pptx operations."""

from __future__ import annotations


def require_shape_attached(shape, *, argument: str = "shape") -> None:
    """Refuse a shape proxy whose XML is no longer attached to its remembered part."""
    from pptx.errors import TargetNotFoundError

    element = getattr(shape, "_element", None)
    try:
        part_root = shape.part._element
    except (AttributeError, ValueError):
        part_root = None
    root = element
    while root is not None and root.getparent() is not None:
        root = root.getparent()
    if element is None or part_root is None or root is not part_root:
        raise TargetNotFoundError(
            "%s is stale: its shape was removed from the presentation" % argument
        )
