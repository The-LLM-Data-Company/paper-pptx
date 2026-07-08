"""Cross-presentation slide import and deck merge (paper-pptx, v0.11 Phase 5).

Composition is how decks are actually made, and composition is relationship-and-
inheritance surgery - the corruption-prone, package-level mechanics this fork exists to
own. `Presentation.import_slide` transplants one slide between packages under one of
three CONSCIOUS reconciliation modes (there is no right default):

- "adopt_theme"  - content transplants; the slide rebinds to a destination layout (the
  Phase 4 machinery), orphan placeholders bake from their SOURCE-resolved look, and every
  run whose resolved values changed is reported. The slide takes the house style.
- "keep_appearance" - the source layout + master + theme chain transplants with it.
  Support parts deduplicate by content fingerprint, so ten slides from one source share
  ONE transplanted master, which gains additional layouts on demand.
- "bake" - every resolvable run's effective values become explicit local properties
  (resolved in the SOURCE package), furniture placeholders (dt/ftr/sldNum) drop, other
  placeholders become free shapes, and the slide attaches to a destination layout.
  Visually stable without importing masters.

The source presentation is never mutated (the cross-contamination guarantee, byte-
tested). All transplant decisions - the full refusal ledger - validate BEFORE the first
destination write. Media always copies across packages (never shared); charts deep-copy
with their embedded workbooks; SmartArt carries opaquely; comments drop (reported);
OLE/ActiveX/internal-link relationships refuse. Fingerprint dedupe is guaranteed within
one destination |Presentation| session; across sessions it applies only when content is
identical after rId normalization (a transplanted master pruned to fewer layouts no
longer fingerprints like its source - declared).
"""

from __future__ import annotations

import copy
import hashlib
import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Dict, List, Optional, Tuple

from pptx.errors import RelationshipPolicyError, TargetNotFoundError, UnsupportedStructureError
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.slideops import (
    _allocate_partname,
    _partname_template,
    _rewrite_r_references,
    _rId_sort_key,
    _validate_chart_rels,
    _validate_notes_rels,
)

if TYPE_CHECKING:
    from pptx.presentation import Presentation
    from pptx.slide import Slide, SlideLayout

SCHEMA_NAME = "paper-import-report"
SCHEMA_VERSION = 1

_MODES = ("adopt_theme", "keep_appearance", "bake")
_MEDIA_RELTYPES = frozenset([RT.IMAGE, RT.MEDIA, RT.VIDEO, RT.AUDIO])
_HF_PH_TYPE_TOKENS = ("dt", "ftr", "sldNum")

#: SmartArt (DrawingML diagram) relationship types - carried opaquely, never edited
_DGM_RELTYPES = frozenset(
    [
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
        "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
    ]
)

_MC_ALTERNATE_CONTENT = (
    "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"
)
_P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"


@dataclass(frozen=True)
class ImportReport:
    """What one import did. Deterministic; `.to_dict()` is goldenable."""

    mode: str
    source_slide: str
    dest_slide: str
    dest_slide_id: int
    position: int
    layout_binding: str
    layout_binding_method: str
    parts_added: Tuple[str, ...]
    parts_reused: Tuple[str, ...]
    notes_copied: bool
    comments_dropped: int
    section: Optional[str]
    baked_shapes: Tuple[str, ...]
    dropped_placeholders: Tuple[str, ...]
    run_shifts: tuple

    def to_dict(self) -> dict:
        return {
            "schema": SCHEMA_NAME,
            "version": SCHEMA_VERSION,
            "mode": self.mode,
            "source_slide": self.source_slide,
            "dest_slide": self.dest_slide,
            "dest_slide_id": self.dest_slide_id,
            "position": self.position,
            "layout_binding": self.layout_binding,
            "layout_binding_method": self.layout_binding_method,
            "parts_added": list(self.parts_added),
            "parts_reused": list(self.parts_reused),
            "notes_copied": self.notes_copied,
            "comments_dropped": self.comments_dropped,
            "section": self.section,
            "baked_shapes": list(self.baked_shapes),
            "dropped_placeholders": list(self.dropped_placeholders),
            "run_shifts": [shift.to_dict() for shift in self.run_shifts],
        }


# ---------------------------------------------------------------------- public entrypoints


def import_slide(
    dest_prs: "Presentation",
    source_prs: "Presentation",
    slide,
    *,
    mode: str,
    position: "Optional[int]" = None,
    notes: bool = True,
    section: "Optional[str]" = None,
    target_layout: "Optional[SlideLayout]" = None,
) -> ImportReport:
    """Import one slide from `source_prs` into `dest_prs`; return the |ImportReport|."""
    source_slide = _validate_arguments(
        dest_prs, source_prs, slide, mode, position, notes, section, target_layout
    )
    plan = _validated_transplant_plan(source_slide.part, notes)
    binding = _resolve_layout_binding(dest_prs, source_slide, mode, target_layout)
    prep = _validate_mode_preparation(source_slide, mode, binding)
    return _perform_import(
        dest_prs, source_slide, plan, mode, binding, prep, position, notes, section
    )


def append_deck(
    dest_prs: "Presentation",
    source_prs: "Presentation",
    *,
    mode: str,
    notes: bool = True,
) -> "Tuple[ImportReport, ...]":
    """Import every `source_prs` slide, in order, at the end of `dest_prs`.

    The COMPLETE source deck validates before the first destination write: a refusal on
    any source slide leaves the destination untouched.
    """
    staged = []
    for source_slide in source_prs.slides:
        source_slide = _validate_arguments(
            dest_prs, source_prs, source_slide, mode, None, notes, None, None
        )
        plan = _validated_transplant_plan(source_slide.part, notes)
        binding = _resolve_layout_binding(dest_prs, source_slide, mode, None)
        prep = _validate_mode_preparation(source_slide, mode, binding)
        staged.append((source_slide, plan, binding, prep))

    reports = []
    for source_slide, plan, binding, prep in staged:
        reports.append(
            _perform_import(
                dest_prs, source_slide, plan, mode, binding, prep, None, notes, None
            )
        )
    return tuple(reports)


# ----------------------------------------------------------------------------- validation


def _validate_arguments(
    dest_prs, source_prs, slide, mode, position, notes, section, target_layout
) -> "Slide":
    from pptx.presentation import Presentation as _Presentation
    from pptx.slide import Slide as _Slide
    from pptx.slide import SlideLayout as _SlideLayout

    if not isinstance(source_prs, _Presentation):
        raise ValueError("source_prs must be a Presentation, got %r" % (source_prs,))
    if source_prs.part.package is dest_prs.part.package:
        raise ValueError(
            "source_prs is this same presentation; same-package duplication is "
            "Slides.clone's job"
        )
    from pptx.errors import materialize_slides

    materialize_slides(source_prs, "import_slide")  # -- typed refusal on broken source
    materialize_slides(dest_prs, "import_slide")
    if isinstance(slide, bool) or not isinstance(slide, (int, _Slide)):
        raise ValueError("slide must be a Slide or int index, got %r" % (slide,))
    if isinstance(slide, int):
        if not (0 <= slide < len(source_prs.slides)):
            raise ValueError(
                "slide index %d out of range 0..%d" % (slide, len(source_prs.slides) - 1)
            )
        slide = source_prs.slides[slide]
    elif slide.part.package is not source_prs.part.package:
        raise ValueError("slide does not belong to source_prs")
    if mode not in _MODES:
        raise ValueError("mode must be one of %s, got %r" % (", ".join(_MODES), mode))
    if position is not None and (
        isinstance(position, bool)
        or not isinstance(position, int)
        or not (0 <= position <= len(dest_prs.slides))
    ):
        raise ValueError(
            "position must be an int in range 0..%d or None, got %r"
            % (len(dest_prs.slides), position)
        )
    if not isinstance(notes, bool):
        raise ValueError("notes must be True or False")
    if section is not None:
        if not isinstance(section, str):
            raise ValueError("section must be a str section name or None")
        if _find_section(dest_prs._element, section) is None:
            raise TargetNotFoundError(
                "destination has no section named %r" % (section,)
            )
    if target_layout is not None:
        if mode == "keep_appearance":
            raise ValueError(
                "target_layout does not apply to keep_appearance (the source layout "
                "chain transplants)"
            )
        if not isinstance(target_layout, _SlideLayout):
            raise ValueError("target_layout must be a SlideLayout")
        if target_layout.part.package is not dest_prs.part.package:
            raise ValueError("target_layout must belong to the destination presentation")
    if mode in ("adopt_theme", "bake") and any(
        True for _ in slide._element.spTree.iterchildren(_MC_ALTERNATE_CONTENT)
    ):
        raise UnsupportedStructureError(
            "source slide contains mc:AlternateContent; shapes inside it are invisible "
            "to placeholder reconciliation and baking - import it with "
            "mode='keep_appearance' (opaque transplant) instead"
        )
    return slide


def _validated_transplant_plan(source_part, notes: bool):
    """[(rId, action, rel)] for every source-slide relationship, or refuse before any write."""
    plan = []
    unsupported = []
    for rId in sorted(source_part.rels, key=_rId_sort_key):
        rel = source_part.rels[rId]
        if rel.is_external:
            plan.append((rId, "external", rel))
        elif rel.reltype == RT.SLIDE_LAYOUT:
            plan.append((rId, "layout", rel))
        elif rel.reltype in _MEDIA_RELTYPES:
            plan.append((rId, "media", rel))  # -- cross-package media ALWAYS copies
        elif rel.reltype == RT.CHART:
            _validate_chart_rels(rel.target_part)
            plan.append((rId, "chart", rel))
        elif rel.reltype == RT.NOTES_SLIDE:
            if notes:
                _validate_notes_rels(rel.target_part)
                plan.append((rId, "notes", rel))
            else:
                plan.append((rId, "drop", rel))
        elif rel.reltype.endswith("/comments"):
            plan.append((rId, "drop-comments", rel))
        elif rel.reltype in _DGM_RELTYPES:
            _validate_dgm_rels(rel.target_part)
            plan.append((rId, "diagram", rel))
        else:
            unsupported.append(rel.reltype)
    if unsupported:
        raise RelationshipPolicyError(
            "slide has relationship types import does not support (the v0.11 refusal "
            "ledger: OLE objects, controls, internal slide links, and anything not in "
            "the ledger): %s" % ", ".join(sorted(unsupported))
        )
    return plan


def _validate_dgm_rels(dgm_part) -> None:
    """SmartArt parts carry opaquely; only media children are expected below them."""
    for rel in dgm_part.rels.values():
        if rel.is_external:
            continue
        if rel.reltype not in _MEDIA_RELTYPES:
            raise RelationshipPolicyError(
                "SmartArt part %s has relationship type import does not support: %s"
                % (dgm_part.partname, rel.reltype)
            )


def _validate_mode_preparation(source_slide, mode, binding):
    """Pre-compute everything the mode needs from the SOURCE, refusing before any write."""
    from pptx.rebind import _compute_mapping, _resolution_state

    if mode == "keep_appearance":
        _validate_support_chain(source_slide.slide_layout)
        return {"before_state": _resolution_state(source_slide)}

    slide_phs = [shape for shape in source_slide.shapes if shape.is_placeholder]
    hf_phs = [
        shape for shape in slide_phs if shape.element.ph.get("type") in _HF_PH_TYPE_TOKENS
    ]
    content_phs = [shape for shape in slide_phs if shape not in hf_phs]

    if mode == "adopt_theme":
        mapping = _compute_mapping(slide_phs, binding.layout, "auto")
        orphans = [s for s in slide_phs if mapping[s.element.ph_idx] is None]
    else:  # -- bake: every content placeholder converts; furniture drops
        mapping = None
        orphans = content_phs

    baked_values = {}
    for shape in orphans:
        if shape.element.findall(
            ".//{http://schemas.openxmlformats.org/drawingml/2006/main}fld"
        ):
            raise UnsupportedStructureError(
                "placeholder %r contains a field (a:fld); baking would freeze volatile "
                "content - import with keep_appearance or remove the field first"
                % shape.name
            )
        geometry = {
            attribute: getattr(shape, attribute)
            for attribute in ("left", "top", "width", "height")
        }
        if any(value is None for value in geometry.values()):
            raise UnsupportedStructureError(
                "placeholder %r has no resolvable geometry; baking would place it "
                "unpredictably" % shape.name
            )
        baked_values[shape.shape_id] = {
            "geometry": geometry,
            "runs": _resolved_run_values(shape),
        }

    if mode == "bake":
        # -- bake also localizes resolvable formatting on every remaining text shape
        for shape in source_slide.shapes:
            if shape.shape_id in baked_values or not shape.has_text_frame:
                continue
            if shape.is_placeholder and shape in hf_phs:
                continue
            baked_values[shape.shape_id] = {"geometry": None, "runs": _resolved_run_values(shape)}

    return {
        "before_state": _resolution_state(source_slide),
        "mapping": mapping,
        "orphan_ids": [shape.shape_id for shape in orphans],
        "hf_ph_ids": [shape.shape_id for shape in hf_phs],
        "baked_values": baked_values,
        "source_idx_of": {shape.shape_id: shape.element.ph_idx for shape in slide_phs},
    }


def _validate_support_chain(source_layout) -> None:
    """Refuse before any write if the layout/master/theme chain has unsupported children."""
    layout_part = source_layout.part
    master_part = source_layout.slide_master.part
    for part, allowed in (
        (layout_part, _MEDIA_RELTYPES | {RT.SLIDE_MASTER}),
        (master_part, _MEDIA_RELTYPES | {RT.SLIDE_LAYOUT, RT.THEME}),
    ):
        for rel in part.rels.values():
            if rel.is_external:
                continue
            if rel.reltype not in allowed:
                raise RelationshipPolicyError(
                    "support part %s has relationship type import does not support: %s"
                    % (part.partname, rel.reltype)
                )
    theme_part = master_part.part_related_by(RT.THEME)
    for rel in theme_part.rels.values():
        if not rel.is_external and rel.reltype not in _MEDIA_RELTYPES:
            raise RelationshipPolicyError(
                "theme part %s has relationship type import does not support: %s"
                % (theme_part.partname, rel.reltype)
            )


def _resolved_run_values(shape) -> list:
    """[(paragraph_idx, run_idx, {facet: value})] of RESOLVED effective values only."""
    resolved = []
    for p_idx, paragraph in enumerate(shape.text_frame.paragraphs):
        for r_idx, run in enumerate(paragraph.runs):
            effective = run.effective_font()
            values = {}
            if effective.size.resolved and effective.size.value is not None:
                values["size"] = effective.size.value
            if effective.name.resolved and effective.name.value is not None:
                values["name"] = effective.name.value
            if effective.color_rgb.resolved and effective.color_rgb.value is not None:
                values["color_rgb"] = effective.color_rgb.value
            for facet in ("bold", "italic"):
                facet_value = getattr(effective, facet)
                if facet_value is not None and facet_value.resolved and (
                    facet_value.value is not None
                ):
                    values[facet] = facet_value.value
            if effective.underline is not None and effective.underline.resolved and (
                effective.underline.value is not None
            ):
                values["underline"] = effective.underline.value
            if values:
                resolved.append((p_idx, r_idx, values))
    return resolved


# ------------------------------------------------------------------------------ mechanics


@dataclass
class _LayoutBinding:
    layout: "Optional[SlideLayout]"  # -- None only for keep_appearance (pre-transplant)
    # -- name-match | type-match | explicit | blank-fallback | first-fallback | transplant
    method: str


def _resolve_layout_binding(dest_prs, source_slide, mode, target_layout) -> _LayoutBinding:
    if mode == "keep_appearance":
        return _LayoutBinding(None, "transplant")
    if target_layout is not None:
        return _LayoutBinding(target_layout, "explicit")
    source_layout = source_slide.slide_layout
    dest_layouts = [
        layout for master in dest_prs.slide_masters for layout in master.slide_layouts
    ]
    source_name = source_layout.name
    if source_name:
        for layout in dest_layouts:
            if layout.name == source_name:
                return _LayoutBinding(layout, "name-match")
    source_type = source_layout._element.get("type")
    if source_type and source_type != "cust":
        for layout in dest_layouts:
            if layout._element.get("type") == source_type:
                return _LayoutBinding(layout, "type-match")
    if mode == "bake":
        for layout in dest_layouts:
            if layout._element.get("type") == "blank":
                return _LayoutBinding(layout, "blank-fallback")
        return _LayoutBinding(dest_layouts[0], "first-fallback")
    raise UnsupportedStructureError(
        "no destination layout matches source layout %r by name or type; pass "
        "target_layout= explicitly (or use keep_appearance to transplant the source "
        "layout chain)" % (source_name or str(source_layout.part.partname))
    )


def _perform_import(
    dest_prs, source_slide, plan, mode, binding, prep, position, notes, section
) -> ImportReport:
    from pptx.parts.slide import SlidePart
    from pptx.rebind import _resolution_state, _shifts_between

    dest_package = dest_prs.part.package
    before_partnames = {str(part.partname) for part in dest_package.iter_parts()}
    reused: "List[str]" = []
    allocated: "set" = set()

    # -- keep_appearance: the support chain first (fingerprint-deduped) ------------------
    if mode == "keep_appearance":
        dest_layout_part = _transplant_layout_chain(
            dest_prs, source_slide.slide_layout, allocated, reused
        )
    else:
        dest_layout_part = binding.layout.part

    # -- the slide part itself ------------------------------------------------------------
    new_slide_part = SlidePart(
        _allocate_partname(dest_package, "/ppt/slides/slide%d.xml", allocated),
        source_slide.part.content_type,
        dest_package,
        copy.deepcopy(source_slide.part._element),
    )
    rId_mapping: "Dict[str, str]" = {}
    comments_dropped = 0
    notes_copied = False
    for old_rId, action, rel in plan:
        if action == "external":
            rId_mapping[old_rId] = new_slide_part.rels.get_or_add_ext_rel(
                rel.reltype, rel.target_ref
            )
        elif action == "layout":
            continue  # -- rebound below, per mode
        elif action == "media":
            rId_mapping[old_rId] = new_slide_part.relate_to(
                _import_support_part(dest_package, rel.target_part, allocated, reused),
                rel.reltype,
            )
        elif action == "chart":
            rId_mapping[old_rId] = new_slide_part.relate_to(
                _import_chart_part(dest_package, rel.target_part, allocated), rel.reltype
            )
        elif action == "notes":
            rId_mapping[old_rId] = new_slide_part.relate_to(
                _import_notes_part(dest_prs, rel.target_part, new_slide_part, allocated),
                rel.reltype,
            )
            notes_copied = True
        elif action == "diagram":
            rId_mapping[old_rId] = new_slide_part.relate_to(
                _import_diagram_part(dest_package, rel.target_part, allocated, reused),
                rel.reltype,
            )
        elif action == "drop-comments":
            comments_dropped += 1
        # -- action == "drop": notes policy says leave them behind
    _rewrite_r_references(new_slide_part._element, rId_mapping)

    # -- mode-specific placeholder handling on the COPY -----------------------------------
    baked_names: "List[str]" = []
    dropped_names: "List[str]" = []
    if mode in ("adopt_theme", "bake"):
        baked_names, dropped_names = _reconcile_copied_placeholders(
            new_slide_part, mode, prep
        )
    new_slide_part.relate_to(dest_layout_part, RT.SLIDE_LAYOUT)

    # -- enroll in the destination slide sequence ------------------------------------------
    pres_part = dest_prs.part
    rId = pres_part.relate_to(new_slide_part, RT.SLIDE)
    sldIdLst = dest_prs._element.get_or_add_sldIdLst()
    sldId = sldIdLst.add_sldId(rId)
    slide_count = len(sldIdLst.sldId_lst)
    final_position = slide_count - 1 if position is None else position
    if final_position < slide_count - 1:
        sldIdLst.sldId_lst[final_position].addprevious(sldId)
    new_slide_id = int(sldId.get("id"))

    _enroll_in_section(dest_prs._element, new_slide_id, final_position, section)

    # -- report ----------------------------------------------------------------------------
    after_state = _resolution_state(new_slide_part.slide)
    shifts = _shifts_between(prep["before_state"], after_state)
    after_partnames = {str(part.partname) for part in dest_package.iter_parts()}
    parts_added = tuple(sorted(after_partnames - before_partnames))

    return ImportReport(
        mode=mode,
        source_slide=str(source_slide.part.partname),
        dest_slide=str(new_slide_part.partname),
        dest_slide_id=new_slide_id,
        position=final_position,
        layout_binding=str(dest_layout_part.partname),
        layout_binding_method=binding.method,
        parts_added=parts_added,
        parts_reused=tuple(sorted(set(reused))),
        notes_copied=notes_copied,
        comments_dropped=comments_dropped,
        section=section,
        baked_shapes=tuple(baked_names),
        dropped_placeholders=tuple(dropped_names),
        run_shifts=shifts,
    )


def _reconcile_copied_placeholders(new_slide_part, mode, prep):
    """Bake/drop/re-idx placeholders in the copied slide XML per the precomputed prep."""
    from pptx.dml.color import RGBColor
    from pptx.enum.text import MSO_TEXT_UNDERLINE_TYPE
    from pptx.util import Emu

    slide = new_slide_part.slide
    baked_names: "List[str]" = []
    dropped_names: "List[str]" = []
    for shape in list(slide.shapes):
        shape_values = prep["baked_values"].get(shape.shape_id)
        if shape.is_placeholder and mode == "bake" and shape.shape_id in prep["hf_ph_ids"]:
            dropped_names.append(shape.name)
            shape._element.getparent().remove(shape._element)
            continue
        if shape_values is not None:
            if shape_values["geometry"] is not None:
                for attribute, value in shape_values["geometry"].items():
                    setattr(shape, attribute, Emu(value))
            paragraphs = shape.text_frame.paragraphs if shape.has_text_frame else ()
            for p_idx, r_idx, values in shape_values["runs"]:
                run = paragraphs[p_idx].runs[r_idx]
                if "size" in values:
                    run.font.size = Emu(values["size"])
                if "name" in values:
                    run.font.name = values["name"]
                if "color_rgb" in values:
                    run.font.color.rgb = RGBColor.from_string(values["color_rgb"])
                if "bold" in values:
                    run.font.bold = values["bold"]
                if "italic" in values:
                    run.font.italic = values["italic"]
                if "underline" in values:
                    run.font.underline = MSO_TEXT_UNDERLINE_TYPE.from_xml(
                        values["underline"]
                    )
        if shape.is_placeholder and shape.shape_id in prep["orphan_ids"]:
            baked_names.append(shape.name)
            ph = shape.element.ph
            ph.getparent().remove(ph)
        elif shape.is_placeholder and mode == "adopt_theme":
            target = prep["mapping"][prep["source_idx_of"][shape.shape_id]]
            if target is not None:
                ph = shape.element.ph
                ph.type = target[0]
                ph.idx = target[1]
    return baked_names, dropped_names


# --------------------------------------------------------------- support-part transplant


def _dedupe_cache(package) -> dict:
    cache = getattr(package, "_paper_compose_fingerprints", None)
    if cache is None:
        cache = {}
        package._paper_compose_fingerprints = cache
    return cache


def _fingerprint(part, depth: int = 0, visiting=None) -> str:
    """Content fingerprint: SHA-256 over content type + rId-normalized blob + children."""
    if visiting is None:
        visiting = set()
    key = id(part)
    if key in visiting or depth > 3:
        return "cycle"
    visiting = visiting | {key}
    normalized = re.sub(rb'"rId[0-9]+"', b'"rId#"', part.blob)
    digest = hashlib.sha256()
    digest.update(part.content_type.encode("utf-8"))
    digest.update(normalized)
    children = []
    for rel in part.rels.values():
        if rel.is_external:
            children.append("%s>%s" % (rel.reltype, rel.target_ref))
        else:
            children.append(
                "%s>%s" % (rel.reltype, _fingerprint(rel.target_part, depth + 1, visiting))
            )
    for entry in sorted(children):
        digest.update(entry.encode("utf-8"))
    return digest.hexdigest()


def _import_support_part(dest_package, part, allocated, reused):
    """Copy a leaf-ish part (media, theme) into `dest_package` with fingerprint dedupe."""
    cache = _dedupe_cache(dest_package)
    fingerprint = _fingerprint(part)
    hit = cache.get(fingerprint)
    if hit is not None:
        reused.append(str(hit.partname))
        return hit
    partname = _allocate_partname(
        dest_package, _partname_template(str(part.partname)), allocated
    )
    if isinstance(part, XmlPart):
        new_part = type(part)(
            partname, part.content_type, dest_package, copy.deepcopy(part._element)
        )
    else:
        new_part = type(part)(partname, part.content_type, dest_package, part.blob)
    rId_mapping = {}
    for rId in sorted(part.rels, key=_rId_sort_key):
        rel = part.rels[rId]
        if rel.is_external:
            rId_mapping[rId] = new_part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            rId_mapping[rId] = new_part.relate_to(
                _import_support_part(dest_package, rel.target_part, allocated, reused),
                rel.reltype,
            )
    if rId_mapping and isinstance(new_part, XmlPart):
        _rewrite_r_references(new_part._element, rId_mapping)
    cache[fingerprint] = new_part
    return new_part


def _import_chart_part(dest_package, chart_part, allocated):
    """Deep-copy a chart (with workbook and style parts) into `dest_package` - never shared.

    Same recipe as the v0 clone machinery's `_copy_chart_part`, but the copies land in
    the DESTINATION package.
    """
    new_chart = _import_leaf_into(dest_package, chart_part, allocated)
    rId_mapping = {}
    for rId in sorted(chart_part.rels, key=_rId_sort_key):
        rel = chart_part.rels[rId]
        if rel.is_external:
            rId_mapping[rId] = new_chart.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            rId_mapping[rId] = new_chart.relate_to(
                _import_leaf_into(dest_package, rel.target_part, allocated), rel.reltype
            )
    _rewrite_r_references(new_chart._element, rId_mapping)
    return new_chart


def _import_notes_part(dest_prs, notes_part, new_slide_part, allocated):
    """Copy a notes part, re-linked to the NEW slide and the DESTINATION notes master."""
    new_notes = _import_leaf_into(dest_prs.part.package, notes_part, allocated)
    dest_notes_master_part = dest_prs.notes_master.part  # -- created if absent (documented)
    rId_mapping = {}
    for rId in sorted(notes_part.rels, key=_rId_sort_key):
        rel = notes_part.rels[rId]
        if rel.is_external:
            rId_mapping[rId] = new_notes.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        elif rel.reltype == RT.SLIDE:
            rId_mapping[rId] = new_notes.relate_to(new_slide_part, RT.SLIDE)
        else:  # -- RT.NOTES_MASTER: the destination's, never the source's
            rId_mapping[rId] = new_notes.relate_to(dest_notes_master_part, RT.NOTES_MASTER)
    _rewrite_r_references(new_notes._element, rId_mapping)
    return new_notes


def _import_diagram_part(dest_package, dgm_part, allocated, reused):
    """SmartArt part: opaque leaf copy with its media children."""
    return _import_support_part(dest_package, dgm_part, allocated, reused)


def _import_leaf_into(dest_package, part, allocated):
    partname = _allocate_partname(
        dest_package, _partname_template(str(part.partname)), allocated
    )
    if isinstance(part, XmlPart):
        return type(part)(
            partname, part.content_type, dest_package, copy.deepcopy(part._element)
        )
    return type(part)(partname, part.content_type, dest_package, part.blob)


def _transplant_layout_chain(dest_prs, source_layout, allocated, reused):
    """Transplant layout+master+theme into the destination, fingerprint-deduped."""
    dest_package = dest_prs.part.package
    cache = _dedupe_cache(dest_package)
    layout_part = source_layout.part
    layout_fingerprint = _fingerprint(layout_part)
    hit = cache.get(layout_fingerprint)
    if hit is not None:
        reused.append(str(hit.partname))
        return hit

    dest_master_part = _transplant_master(dest_prs, source_layout.slide_master, allocated, reused)

    new_layout = _import_leaf_into(dest_package, layout_part, allocated)
    rId_mapping = {}
    for rId in sorted(layout_part.rels, key=_rId_sort_key):
        rel = layout_part.rels[rId]
        if rel.is_external:
            rId_mapping[rId] = new_layout.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        elif rel.reltype == RT.SLIDE_MASTER:
            rId_mapping[rId] = new_layout.relate_to(dest_master_part, RT.SLIDE_MASTER)
        else:  # -- media (validated)
            rId_mapping[rId] = new_layout.relate_to(
                _import_support_part(dest_package, rel.target_part, allocated, reused),
                rel.reltype,
            )
    _rewrite_r_references(new_layout._element, rId_mapping)

    # -- enroll the layout in the transplanted master
    master_rId = dest_master_part.relate_to(new_layout, RT.SLIDE_LAYOUT)
    sldLayoutIdLst = dest_master_part._element.get_or_add_sldLayoutIdLst()
    entry = sldLayoutIdLst.makeelement(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}sldLayoutId", {}
    )
    entry.set("id", str(_next_layout_or_master_id(dest_prs)))
    entry.set(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", master_rId
    )
    sldLayoutIdLst.append(entry)

    cache[layout_fingerprint] = new_layout
    return new_layout


def _transplant_master(dest_prs, source_master, allocated, reused):
    dest_package = dest_prs.part.package
    cache = _dedupe_cache(dest_package)
    master_part = source_master.part
    master_fingerprint = _fingerprint(master_part)
    hit = cache.get(master_fingerprint)
    if hit is not None:
        reused.append(str(hit.partname))
        return hit

    new_master = _import_leaf_into(dest_package, master_part, allocated)
    # -- start the copy with NO layouts: they enroll on demand (dedupe requirement)
    P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    sldLayoutIdLst = new_master._element.find(P + "sldLayoutIdLst")
    if sldLayoutIdLst is not None:
        for entry in list(sldLayoutIdLst):
            sldLayoutIdLst.remove(entry)

    rId_mapping = {}
    for rId in sorted(master_part.rels, key=_rId_sort_key):
        rel = master_part.rels[rId]
        if rel.is_external:
            rId_mapping[rId] = new_master.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        elif rel.reltype == RT.SLIDE_LAYOUT:
            continue  # -- deliberately not copied; enrolled on demand
        elif rel.reltype == RT.THEME:
            rId_mapping[rId] = new_master.relate_to(
                _import_support_part(dest_package, rel.target_part, allocated, reused),
                RT.THEME,
            )
        else:  # -- media (validated)
            rId_mapping[rId] = new_master.relate_to(
                _import_support_part(dest_package, rel.target_part, allocated, reused),
                rel.reltype,
            )
    _rewrite_r_references(new_master._element, rId_mapping)

    # -- enroll the master in the destination presentation
    pres_rId = dest_prs.part.relate_to(new_master, RT.SLIDE_MASTER)
    sldMasterIdLst = dest_prs._element.get_or_add_sldMasterIdLst()
    entry = sldMasterIdLst.makeelement(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}sldMasterId", {}
    )
    entry.set("id", str(_next_layout_or_master_id(dest_prs)))
    entry.set(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", pres_rId
    )
    sldMasterIdLst.append(entry)

    cache[master_fingerprint] = new_master
    return new_master


def _next_layout_or_master_id(dest_prs) -> int:
    """Next document-unique slide-master/layout id (schema minimum 2147483648)."""
    used = [2147483647]
    P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    for entry in dest_prs._element.findall(".//%ssldMasterId" % P):
        used.append(int(entry.get("id") or 0))
    for master in dest_prs.slide_masters:
        for entry in master._element.findall(".//%ssldLayoutId" % P):
            used.append(int(entry.get("id") or 0))
    return max(used) + 1


# ---------------------------------------------------------------------------- sections


def _find_section(presentation_elm, name: str):
    for section in presentation_elm.findall(".//{%s}sectionLst/{%s}section" % (_P14_NS, _P14_NS)):
        if section.get("name") == name:
            return section
    return None


def _enroll_in_section(presentation_elm, new_slide_id, final_position, section_name) -> None:
    """Enroll the imported slide: named section, or adjacent to the insertion point."""
    sections = presentation_elm.findall(
        ".//{%s}sectionLst/{%s}section" % (_P14_NS, _P14_NS)
    )
    if not sections:
        return
    if section_name is not None:
        section = _find_section(presentation_elm, section_name)
        sldIdLst = section.find("{%s}sldIdLst" % _P14_NS)
        if sldIdLst is None:
            sldIdLst = section.makeelement("{%s}sldIdLst" % _P14_NS, {})
            section.insert(0, sldIdLst)
        entry = sldIdLst.makeelement("{%s}sldId" % _P14_NS, {"id": str(new_slide_id)})
        sldIdLst.append(entry)
        return
    # -- adjacent enrollment: after the preceding slide's entry; first section if none
    P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    deck_sldIds = presentation_elm.findall(".//%ssldIdLst/%ssldId" % (P, P))
    preceding_id = None
    if final_position > 0:
        preceding_id = deck_sldIds[final_position - 1].get("id")
    if preceding_id is not None:
        for section in sections:
            for entry in section.findall(".//{%s}sldId" % _P14_NS):
                if entry.get("id") == preceding_id:
                    new_entry = entry.makeelement(
                        "{%s}sldId" % _P14_NS, {"id": str(new_slide_id)}
                    )
                    entry.addnext(new_entry)
                    return
    first_sldIdLst = sections[0].find("{%s}sldIdLst" % _P14_NS)
    if first_sldIdLst is None:
        first_sldIdLst = sections[0].makeelement("{%s}sldIdLst" % _P14_NS, {})
        sections[0].insert(0, first_sldIdLst)
    entry = first_sldIdLst.makeelement("{%s}sldId" % _P14_NS, {"id": str(new_slide_id)})
    first_sldIdLst.insert(0, entry)
