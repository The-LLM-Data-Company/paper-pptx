"""Internal machinery for slide clone/delete/reorder (paper-pptx). Not public API.

The public surface is `Slides.clone/delete/reorder/move` and `SlideClonePolicy` in
`pptx.slide`. Everything here operates on the in-memory opc package — parts, relationships,
content types — never on unpacked files; `[Content_Types].xml` regenerates from live parts at
save, and a part no longer reachable through the relationship graph is simply never
serialized, so orphans structurally cannot reach disk.

Clone is validate-fully-then-mutate: the complete relationship plan for the source slide (and
every deep-copied part's own relationships) is validated against the policy BEFORE any part is
created, so a `RelationshipPolicyError` provably leaves the package untouched.
"""

from __future__ import annotations

import copy
import re
from typing import TYPE_CHECKING, Dict, List, Tuple

from pptx.errors import RelationshipPolicyError
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI

if TYPE_CHECKING:
    from pptx.parts.slide import SlidePart

_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
#: Microsoft chart-style extension parts (LibreOffice and recent PowerPoint emit these)
_CHART_COLOR_STYLE = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle"
_CHART_STYLE = "http://schemas.microsoft.com/office/2011/relationships/chartStyle"

#: media-like relationship types, shared between clone and source by default
_MEDIA_RELTYPES = frozenset([RT.IMAGE, RT.MEDIA, RT.VIDEO, RT.AUDIO])
#: relationship types allowed FROM a chart part, all deep-copied with it
_CHART_CHILD_RELTYPES = frozenset([RT.PACKAGE, _CHART_COLOR_STYLE, _CHART_STYLE])
#: relationship types allowed FROM a notes-slide part
_NOTES_CHILD_RELTYPES = frozenset([RT.NOTES_MASTER, RT.SLIDE])


def clone_slide_part(source_part: "SlidePart", policy) -> "SlidePart":
    """Return a new |SlidePart| that is a policy-governed deep copy of `source_part`.

    The new part is fully related (layout, media, charts+workbooks, notes per `policy`) but
    NOT yet added to the presentation's slide list — the caller owns `p:sldIdLst`.
    """
    from pptx.parts.slide import SlidePart

    package = source_part.package
    plan = _validated_plan(source_part, policy)

    # -- Parts created here are unreachable from the package rels graph until the caller
    # -- relates the new slide, so `package.next_partname` cannot see them. `allocated`
    # -- tracks every partname assigned during THIS clone so two deep copies sharing a
    # -- template (e.g. two charts) can never collide.
    allocated: "set" = set()
    new_part = SlidePart(
        _allocate_partname(package, "/ppt/slides/slide%d.xml", allocated),
        CT.PML_SLIDE,
        package,
        copy.deepcopy(source_part._element),
    )

    rId_mapping: "Dict[str, str]" = {}
    for old_rId, action, rel in plan:
        if action == "external":
            rId_mapping[old_rId] = new_part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        elif action == "share":
            rId_mapping[old_rId] = new_part.relate_to(rel.target_part, rel.reltype)
        elif action == "copy":
            rId_mapping[old_rId] = new_part.relate_to(
                _copy_leaf_part(rel.target_part, allocated), rel.reltype
            )
        elif action == "chart":
            rId_mapping[old_rId] = new_part.relate_to(
                _copy_chart_part(rel.target_part, allocated), rel.reltype
            )
        elif action == "notes":
            rId_mapping[old_rId] = new_part.relate_to(
                _copy_notes_part(rel.target_part, new_part, allocated), rel.reltype
            )
        # -- action == "drop": no relationship on the clone (notes policy) --

    _rewrite_r_references(new_part._element, rId_mapping)
    return new_part


def _allocate_partname(package, template: str, allocated: "set") -> PackURI:
    """Return the lowest-numbered partname from `template` unused in `package` OR `allocated`.

    Records the returned name in `allocated` so subsequent allocations within the same
    (not-yet-related) clone operation cannot reuse it.
    """
    used = {str(part.partname) for part in package.iter_parts()} | allocated
    index = 1
    while template % index in used:
        index += 1
    partname = template % index
    allocated.add(partname)
    return PackURI(partname)


def _validated_plan(source_part, policy) -> "List[Tuple[str, str, object]]":
    """Return [(rId, action, rel)] for every source relationship, or raise before mutating.

    Also pre-validates the relationship graphs of parts that will be deep-copied (charts and
    notes), so no failure can occur after part creation begins.
    """
    plan = []
    unsupported = []
    for rId in sorted(source_part.rels, key=_rId_sort_key):
        rel = source_part.rels[rId]
        if rel.is_external:
            plan.append((rId, "external", rel))
        elif rel.reltype == RT.SLIDE_LAYOUT:
            plan.append((rId, "share", rel))
        elif rel.reltype in _MEDIA_RELTYPES:
            plan.append((rId, "share" if policy.share_media else "copy", rel))
        elif rel.reltype == RT.CHART:
            if not policy.deep_copy_charts:
                raise RelationshipPolicyError(
                    "cloning a slide with a chart requires deep_copy_charts=True: sharing an"
                    " editable chart part between slides is exactly the cross-contamination"
                    " this API exists to prevent, and is not offered in v0"
                )
            _validate_chart_rels(rel.target_part)
            plan.append((rId, "chart", rel))
        elif rel.reltype == RT.NOTES_SLIDE:
            if policy.deep_copy_notes:
                _validate_notes_rels(rel.target_part)
                plan.append((rId, "notes", rel))
            else:
                plan.append((rId, "drop", rel))
        else:
            unsupported.append(rel.reltype)
    if unsupported:
        raise RelationshipPolicyError(
            "slide has relationship types clone does not support in v0: %s"
            % ", ".join(sorted(unsupported))
        )
    return plan


def _validate_chart_rels(chart_part) -> None:
    for rId in chart_part.rels:
        rel = chart_part.rels[rId]
        if rel.is_external:
            continue
        if rel.reltype not in _CHART_CHILD_RELTYPES:
            raise RelationshipPolicyError(
                "chart part %s has relationship type clone does not support in v0: %s"
                % (chart_part.partname, rel.reltype)
            )
        if any(not child.is_external for child in rel.target_part.rels.values()):
            raise RelationshipPolicyError(
                "chart child part %s has internal relationships of its own; clone supports"
                " only leaf chart children in v0" % rel.target_part.partname
            )


def _validate_notes_rels(notes_part) -> None:
    for rId in notes_part.rels:
        rel = notes_part.rels[rId]
        if not rel.is_external and rel.reltype not in _NOTES_CHILD_RELTYPES:
            raise RelationshipPolicyError(
                "notes slide %s has relationship type clone does not support in v0: %s"
                % (notes_part.partname, rel.reltype)
            )


def _copy_chart_part(chart_part, allocated):
    """Return a deep copy of `chart_part` including its embedded workbook and style parts."""
    new_chart = _copy_leaf_part(chart_part, allocated)
    rId_mapping = {}
    for rId in sorted(chart_part.rels, key=_rId_sort_key):
        rel = chart_part.rels[rId]
        if rel.is_external:
            rId_mapping[rId] = new_chart.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            rId_mapping[rId] = new_chart.relate_to(
                _copy_leaf_part(rel.target_part, allocated), rel.reltype
            )
    _rewrite_r_references(new_chart._element, rId_mapping)
    return new_chart


def _copy_notes_part(notes_part, new_slide_part, allocated):
    """Return a deep copy of `notes_part`, related to the notes master and the CLONE slide."""
    new_notes = _copy_leaf_part(notes_part, allocated)
    rId_mapping = {}
    for rId in sorted(notes_part.rels, key=_rId_sort_key):
        rel = notes_part.rels[rId]
        if rel.is_external:
            rId_mapping[rId] = new_notes.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        elif rel.reltype == RT.SLIDE:
            rId_mapping[rId] = new_notes.relate_to(new_slide_part, RT.SLIDE)
        else:  # -- RT.NOTES_MASTER: shared singleton
            rId_mapping[rId] = new_notes.relate_to(rel.target_part, rel.reltype)
    _rewrite_r_references(new_notes._element, rId_mapping)
    return new_notes


def _copy_leaf_part(part, allocated):
    """Return a new part of `part`'s class with copied content and a fresh partname."""
    package = part.package
    partname = _allocate_partname(package, _partname_template(str(part.partname)), allocated)
    if isinstance(part, XmlPart):
        return type(part)(partname, part.content_type, package, copy.deepcopy(part._element))
    return type(part)(partname, part.content_type, package, part.blob)


def _partname_template(partname: str) -> str:
    """Return a next_partname template for `partname`, e.g. "/ppt/charts/chart%d.xml"."""
    template, substitutions = re.subn(r"[0-9]+(?=\.[^.]+$)", "%d", partname, count=1)
    if substitutions == 0:
        stem, dot, ext = partname.rpartition(".")
        template = "%s%%d%s%s" % (stem, dot, ext)
    return template


def _rewrite_r_references(root, rId_mapping: "Dict[str, str]") -> None:
    """Rewrite every r-namespace attribute in `root` per `rId_mapping`, in place."""
    if not rId_mapping:
        return
    prefix = "{%s}" % _R_NS
    for element in root.iter():
        for attr_name, attr_value in element.attrib.items():
            if attr_name.startswith(prefix) and attr_value in rId_mapping:
                element.set(attr_name, rId_mapping[attr_value])


def _rId_sort_key(rId: str):
    match = re.fullmatch(r"rId([0-9]+)", rId)
    return (0, int(match.group(1))) if match else (1, rId)
