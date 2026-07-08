"""Deck scrubbing - the exit gate (paper-pptx, v0.11 Phase 3).

`Presentation.scrub(...)` removes exactly the individually-toggled targets and nothing
else. Reachability is the heart: every removal is expressed as dropping a relationship
(or an id-list entry plus its relationship), and parts leave the package only by becoming
unreachable through the relationship graph - upstream's serializer never writes an
unreachable part, so a part reachable from any live slide, layout, or master structurally
cannot be removed. The returned |ScrubReport| is typed, deterministic, and carries the
exact zip-member budget (`parts_removed`/`parts_modified`) so tests can assert the
changed-part diff matches the report member for member.

Declared behaviors (deliberate, documented):
- The notes master is RETAINED by `notes=True`: it is referenced from the presentation
  part (never orphaned by notes-slide removal) and PowerPoint decks routinely carry one
  with zero notes slides.
- `metadata=True` clears the core-properties text fields and removes the extended
  (app.xml), custom-properties, and thumbnail parts; created/modified/revision survive -
  they are pipeline-relevant, not personal.
- Comment removal matches both classic (`.../comments`, `.../commentAuthors`) and modern
  (2018/10 `.../comments`, `.../authors`) relationship types; only classic parts exist in
  the frozen corpus today (FIXTURE-REQUESTS.md R12 tracks the real-PowerPoint modern
  fixture).
"""

from __future__ import annotations

import posixpath
from dataclasses import dataclass
from dataclasses import fields as dataclass_fields
from typing import TYPE_CHECKING, Dict, List, Set, Tuple

if TYPE_CHECKING:
    from pptx.presentation import Presentation

SCHEMA_NAME = "paper-scrub-report"
SCHEMA_VERSION = 1

_MEDIA_RELTYPE_SUFFIXES = ("/image", "/media", "/video", "/audio")
_COMMENT_RELTYPE_SUFFIX = "/comments"  # -- classic and modern reltypes both end this way
_COMMENT_AUTHOR_RELTYPE_SUFFIXES = ("/commentAuthors", "2018/10/relationships/authors")
_FONT_RELTYPE_SUFFIX = "/font"
_METADATA_RELTYPE_SUFFIXES = (
    "/extended-properties",
    "/custom-properties",
    "/metadata/thumbnail",
)
_CLEARED_CORE_FIELDS = (
    "author",
    "category",
    "comments",
    "content_status",
    "identifier",
    "keywords",
    "language",
    "last_modified_by",
    "subject",
    "title",
    "version",
)


@dataclass(frozen=True)
class ScrubReport:
    """What one scrub actually did. Deterministic; `.to_dict()` is goldenable.

    Two addressing conventions, deliberately: the per-category fields name OPC parts by
    PARTNAME (leading slash, e.g. "/ppt/comments/comment1.xml") because they identify
    document parts; `parts_removed`/`parts_modified` name ZIP MEMBERS (no leading slash,
    including `.rels` members and `[Content_Types].xml`, which are not parts) because
    they are the exact save-output budget a byte-level diff is held against.
    """

    notes_slides_removed: Tuple[str, ...] = ()
    comment_parts_removed: Tuple[str, ...] = ()
    comment_author_parts_removed: Tuple[str, ...] = ()
    metadata_fields_cleared: Tuple[str, ...] = ()
    metadata_parts_removed: Tuple[str, ...] = ()
    hidden_slides_removed: Tuple[str, ...] = ()
    unused_layouts_removed: Tuple[str, ...] = ()
    unused_masters_removed: Tuple[str, ...] = ()
    unreachable_media_rels_dropped: Tuple[str, ...] = ()
    embedded_font_parts_removed: Tuple[str, ...] = ()
    notes_master_retained: bool = False
    parts_removed: Tuple[str, ...] = ()
    parts_modified: Tuple[str, ...] = ()

    def to_dict(self) -> dict:
        payload = {"schema": SCHEMA_NAME, "version": SCHEMA_VERSION}
        for f in dataclass_fields(self):
            value = getattr(self, f.name)
            payload[f.name] = list(value) if isinstance(value, tuple) else value
        return payload


def _rels_member(partname: str) -> str:
    """Zip member holding `partname`'s relationships, e.g. ppt/slides/_rels/slide1.xml.rels."""
    directory, basename = posixpath.split(partname)
    return "%s/_rels/%s.rels" % (directory.lstrip("/"), basename)


def _member(partname: str) -> str:
    return partname.lstrip("/")


def scrub_presentation(
    prs: "Presentation",
    *,
    notes: bool = False,
    comments: bool = False,
    metadata: bool = False,
    hidden_slides: bool = False,
    unused_layouts: bool = False,
    unused_masters: bool = False,
    unreachable_media: bool = False,
    embedded_fonts: bool = False,
) -> ScrubReport:
    """Perform the toggled scrub passes on `prs` and return the |ScrubReport|."""
    toggles = {
        "notes": notes,
        "comments": comments,
        "metadata": metadata,
        "hidden_slides": hidden_slides,
        "unused_layouts": unused_layouts,
        "unused_masters": unused_masters,
        "unreachable_media": unreachable_media,
        "embedded_fonts": embedded_fonts,
    }
    for name, value in toggles.items():
        if not isinstance(value, bool):
            raise ValueError("%s must be True or False, got %r" % (name, value))

    from pptx.errors import UnsupportedStructureError, materialize_slides

    slides = materialize_slides(prs, "scrub")  # -- typed refusal on a broken graph
    if unused_layouts or unused_masters:
        # -- the layout-usage pass resolves every slide's layout; validate that NOW so
        # -- a broken slide->layout relationship refuses BEFORE any pass has mutated
        # -- anything (refusal atomicity, CONVENTIONS 1.3)
        try:
            for slide in slides:
                slide.slide_layout
        except KeyError as exc:
            raise UnsupportedStructureError(
                "scrub refused: a slide's layout relationship is broken (%s); repair "
                "the package before operating on it" % exc
            ) from exc

    package = prs.part.package
    before_parts: "Dict[str, object]" = {
        str(part.partname): part for part in package.iter_parts()
    }
    modified: "Set[str]" = set()
    collected: "Dict[str, List[str]]" = {key: [] for key in (
        "notes_slides_removed",
        "comment_parts_removed",
        "comment_author_parts_removed",
        "metadata_fields_cleared",
        "metadata_parts_removed",
        "hidden_slides_removed",
        "unused_layouts_removed",
        "unused_masters_removed",
        "unreachable_media_rels_dropped",
        "embedded_font_parts_removed",
    )}

    presentation_part = prs.part
    presentation_partname = str(presentation_part.partname)

    # -- 1. hidden slides first: their notes/comments/media leave with them ---------------
    if hidden_slides:
        for slide in [s for s in prs.slides if s._element.get("show") in ("0", "false")]:
            collected["hidden_slides_removed"].append(str(slide.part.partname))
            prs.slides.delete(slide)
        if collected["hidden_slides_removed"]:
            modified.add(_member(presentation_partname))
            modified.add(_rels_member(presentation_partname))

    # -- 2. speaker notes (notes master deliberately retained - module docstring) ---------
    if notes:
        for slide in prs.slides:
            part = slide.part
            for rId, rel in list(part.rels.items()):
                if rel.is_external:
                    continue
                if rel.reltype.endswith("/notesSlide"):
                    collected["notes_slides_removed"].append(str(rel.target_part.partname))
                    part.drop_rel(rId)
                    modified.add(_rels_member(str(part.partname)))

    # -- 3. comments: per-slide comment parts + deck-level author registries --------------
    if comments:
        for slide in prs.slides:
            part = slide.part
            for rId, rel in list(part.rels.items()):
                if rel.is_external:
                    continue
                if rel.reltype.endswith(_COMMENT_RELTYPE_SUFFIX):
                    collected["comment_parts_removed"].append(str(rel.target_part.partname))
                    part.drop_rel(rId)
                    modified.add(_rels_member(str(part.partname)))
        for rId, rel in list(presentation_part.rels.items()):
            if rel.is_external:
                continue
            if any(rel.reltype.endswith(s) for s in _COMMENT_AUTHOR_RELTYPE_SUFFIXES):
                collected["comment_author_parts_removed"].append(
                    str(rel.target_part.partname)
                )
                presentation_part.drop_rel(rId)
                modified.add(_rels_member(presentation_partname))

    # -- 4. embedded fonts: the p:embeddedFontLst and every font-data relationship --------
    if embedded_fonts:
        fontLst = prs._element.find(
            "{http://schemas.openxmlformats.org/presentationml/2006/main}embeddedFontLst"
        )
        for rId, rel in list(presentation_part.rels.items()):
            if rel.is_external:
                continue
            if rel.reltype.endswith(_FONT_RELTYPE_SUFFIX):
                collected["embedded_font_parts_removed"].append(
                    str(rel.target_part.partname)
                )
                if fontLst is not None and fontLst.getparent() is not None:
                    prs._element.remove(fontLst)
                    modified.add(_member(presentation_partname))
                presentation_part.drop_rel(rId)
                modified.add(_rels_member(presentation_partname))

    # -- 5. unused layouts / masters (usage computed after hidden-slide removal; the
    # -- layout resolution below was validated up front, so it cannot fail mid-scrub)
    sldMasterIdLst = prs._element.sldMasterIdLst
    if (unused_layouts or unused_masters) and sldMasterIdLst is not None:
        used_layout_partnames = {str(s.slide_layout.part.partname) for s in prs.slides}
        for master in list(prs.slide_masters):
            master_part = master.part
            layout_rels = [
                (rId, rel)
                for rId, rel in master_part.rels.items()
                if not rel.is_external and rel.reltype.endswith("/slideLayout")
            ]
            master_serves_slides = any(
                str(rel.target_part.partname) in used_layout_partnames
                for _, rel in layout_rels
            )
            if unused_masters and not master_serves_slides:
                for rId, rel in presentation_part.rels.items():
                    if not rel.is_external and rel.target_part is master_part:
                        for entry in list(sldMasterIdLst):
                            if entry.get(
                                "{http://schemas.openxmlformats.org/officeDocument/2006/"
                                "relationships}id"
                            ) == rId:
                                sldMasterIdLst.remove(entry)
                        collected["unused_masters_removed"].append(
                            str(master_part.partname)
                        )
                        presentation_part.drop_rel(rId)
                        modified.add(_member(presentation_partname))
                        modified.add(_rels_member(presentation_partname))
                        break
                continue
            if unused_layouts:
                sldLayoutIdLst = master.element.sldLayoutIdLst
                for rId, rel in layout_rels:
                    if str(rel.target_part.partname) in used_layout_partnames:
                        continue
                    for entry in list(sldLayoutIdLst):
                        if entry.get(
                            "{http://schemas.openxmlformats.org/officeDocument/2006/"
                            "relationships}id"
                        ) == rId:
                            sldLayoutIdLst.remove(entry)
                    collected["unused_layouts_removed"].append(str(rel.target_part.partname))
                    master_part.drop_rel(rId)
                    modified.add(_member(str(master_part.partname)))
                    modified.add(_rels_member(str(master_part.partname)))

    # -- 6. unreachable media: rels no XML reference uses (never touches referenced media)
    if unreachable_media:
        from pptx.shapes.picture import _part_xml_references_rId

        for part in list(package.iter_parts()):
            element = getattr(part, "_element", None)
            if element is None:
                continue
            for rId, rel in list(part.rels.items()):
                if rel.is_external:
                    continue
                if not any(rel.reltype.endswith(s) for s in _MEDIA_RELTYPE_SUFFIXES):
                    continue
                if _part_xml_references_rId(element, rId):
                    continue
                collected["unreachable_media_rels_dropped"].append(
                    str(rel.target_part.partname)
                )
                part.drop_rel(rId)
                modified.add(_rels_member(str(part.partname)))

    # -- 7. metadata: core-props text fields + app/custom/thumbnail parts -----------------
    if metadata:
        # -- upstream's core_properties accessor CREATES a default part when absent;
        # -- scrub may never create parts, so probe the relationship first
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT

        has_core_part = any(
            not rel.is_external and rel.reltype == RT.CORE_PROPERTIES
            for rel in package._rels.values()
        )
        if has_core_part:
            core = prs.core_properties
            for field_name in _CLEARED_CORE_FIELDS:
                if getattr(core, field_name):
                    setattr(core, field_name, "")
                    collected["metadata_fields_cleared"].append(field_name)
            if collected["metadata_fields_cleared"]:
                modified.add("docProps/core.xml")
        for rId, rel in list(package._rels.items()):
            if rel.is_external:
                continue
            if any(rel.reltype.endswith(s) for s in _METADATA_RELTYPE_SUFFIXES):
                collected["metadata_parts_removed"].append(str(rel.target_part.partname))
                package.drop_rel(rId)
                modified.add("_rels/.rels")

    # -- settle the exact member budget from actual reachability ---------------------------
    after_partnames = {str(part.partname) for part in package.iter_parts()}
    removed_partnames = sorted(set(before_parts) - after_partnames)
    appeared = after_partnames - set(before_parts)
    if appeared:  # pragma: no cover - structural invariant, never expected
        raise AssertionError("scrub may never create parts, but found %r" % sorted(appeared))

    parts_removed: "List[str]" = []
    for partname in removed_partnames:
        parts_removed.append(_member(partname))
        if len(before_parts[partname].rels):  # type: ignore[attr-defined]
            parts_removed.append(_rels_member(partname))

    if removed_partnames:
        # -- [Content_Types].xml changes iff a removed part carried its own Override
        # -- (its (extension, content_type) is not in the serializer's Default table -
        # -- e.g. image/svg+xml media), or the last Default-covered part of an
        # -- extension left. Mirrors opc/serialized.py's actual emission rule.
        from pptx.opc.spec import default_content_types

        def _is_default_typed(part) -> bool:
            extension = posixpath.splitext(str(part.partname))[1][1:].lower()
            return (extension, part.content_type) in default_content_types

        surviving_default_extensions = {
            posixpath.splitext(str(part.partname))[1][1:].lower()
            for part in package.iter_parts()
            if _is_default_typed(part)
        }
        for partname in removed_partnames:
            removed_part = before_parts[partname]
            if not _is_default_typed(removed_part):
                modified.add("[Content_Types].xml")
                break
            extension = posixpath.splitext(partname)[1][1:].lower()
            if extension not in surviving_default_extensions:
                modified.add("[Content_Types].xml")
                break

    return ScrubReport(
        notes_slides_removed=tuple(sorted(collected["notes_slides_removed"])),
        comment_parts_removed=tuple(sorted(collected["comment_parts_removed"])),
        comment_author_parts_removed=tuple(
            sorted(collected["comment_author_parts_removed"])
        ),
        metadata_fields_cleared=tuple(sorted(collected["metadata_fields_cleared"])),
        metadata_parts_removed=tuple(sorted(collected["metadata_parts_removed"])),
        hidden_slides_removed=tuple(sorted(collected["hidden_slides_removed"])),
        unused_layouts_removed=tuple(sorted(collected["unused_layouts_removed"])),
        unused_masters_removed=tuple(sorted(collected["unused_masters_removed"])),
        unreachable_media_rels_dropped=tuple(
            sorted(collected["unreachable_media_rels_dropped"])
        ),
        embedded_font_parts_removed=tuple(
            sorted(collected["embedded_font_parts_removed"])
        ),
        notes_master_retained=any(
            rel.reltype.endswith("/notesMaster")
            for rel in presentation_part.rels.values()
            if not rel.is_external
        ),
        parts_removed=tuple(sorted(parts_removed)),
        parts_modified=tuple(sorted(modified)),
    )
