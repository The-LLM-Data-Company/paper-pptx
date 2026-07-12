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
    require_part_reachable(shape.part, argument=argument)


def require_element_attached(element, part, *, argument: str) -> None:
    """Refuse an element detached from the active root of its remembered part."""
    from pptx.errors import TargetNotFoundError

    root = element
    while root is not None and root.getparent() is not None:
        root = root.getparent()
    if root is not getattr(part, "_element", None):
        raise TargetNotFoundError(
            "%s is stale: its content was removed from the presentation" % argument
        )
    require_part_reachable(part, argument=argument)


def require_shape_tree_attached(shapes, *, argument: str = "target shape tree") -> None:
    """Refuse a shape collection that no longer wraps its part's active tree."""
    from pptx.errors import TargetNotFoundError

    spTree = getattr(shapes, "_spTree", None)
    part = shapes.part
    require_element_attached(spTree, part, argument=argument)
    live_spTree = part._element.cSld.spTree
    if spTree is not live_spTree:
        raise TargetNotFoundError(
            "%s is stale: it is no longer the part's active shape tree" % argument
        )


def require_part_reachable(part, *, argument: str) -> None:
    """Refuse a part proxy no longer reachable from its package root."""
    from pptx.errors import TargetNotFoundError, UnsupportedStructureError

    try:
        reachable = any(candidate is part for candidate in part.package.iter_parts())
    except (AssertionError, AttributeError, KeyError, TypeError, ValueError) as exc:
        raise UnsupportedStructureError(
            "cannot verify %s ownership because the package relationship graph is broken (%s)"
            % (argument, exc)
        ) from exc
    if not reachable:
        raise TargetNotFoundError(
            "%s is stale: its package part is no longer in the presentation" % argument
        )
