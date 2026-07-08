"""Effective-style inspection: resolved font values with provenance (paper-pptx).

Upstream python-pptx reports `None` for any character property that is inherited, leaving
callers blind to what a deck actually renders. This module implements the documented
inheritance walk for **font size, font name, and font color** and reports, for every resolved
value, the ordered chain of sources consulted (CONVENTIONS §2: read-only inspection is
provenance-bearing).

The pinned walk, per run (levels consulted in order until one supplies the value):

1. the run's own `a:rPr`
2. the containing paragraph's `a:pPr/a:defRPr`
3. the shape's own `p:txBody/a:lstStyle` at the paragraph's indent level
4. placeholder shapes: the layout placeholder's `lstStyle` (matched by `idx`, falling back to
   `type`), then the master placeholder's `lstStyle` (title-family → master title placeholder,
   otherwise master body placeholder), then the master's `p:txStyles` family style
   (title/ctrTitle → `titleStyle`; body/subTitle/obj and vertical variants → `bodyStyle`;
   anything else → `otherStyle`)
5. non-placeholder shapes: the presentation's `p:defaultTextStyle`

Font-name theme references (`+mj-lt`/`+mn-lt`) resolve through the master's theme part
(`a:fontScheme`). Scheme colors resolve through the slide's `p:clrMapOvr` override when
present, else the master's `p:clrMap`, then the theme's `a:clrScheme` (`a:sysClr` uses its
`lastClr`). Anything genuinely outside this walk reports `resolved=False` with the consulted
chain — never a guessed default.

Everything here is strictly read-only: no method mutates any element, and raw-XML reads are
used throughout (several upstream proxy accessors are get-or-add and would dirty the tree).
"""

from __future__ import annotations

import hashlib
import unicodedata
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Iterator, Optional, Tuple

from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.errors import UnsupportedStructureError
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Centipoints, Length

if TYPE_CHECKING:
    from pptx.slide import Slide
    from pptx.text.text import _Run

SCHEMA_NAME = "paper-text-inspection"
SCHEMA_VERSION = 1

_TITLE_FAMILY = frozenset([PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE])
_BODY_FAMILY = frozenset(
    [
        PP_PLACEHOLDER.BODY,
        PP_PLACEHOLDER.SUBTITLE,
        PP_PLACEHOLDER.OBJECT,
        PP_PLACEHOLDER.VERTICAL_BODY,
        PP_PLACEHOLDER.VERTICAL_OBJECT,
    ]
)
#: schemeClr tokens addressing theme slots directly, bypassing the clrMap indirection
_DIRECT_THEME_SLOTS = frozenset(
    ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
     "accent6", "hlink", "folHlink"]
)


def content_hash(text: str) -> str:
    """First 8 hex chars of SHA-256 over the NFC-normalized text (the pinned anchor hash).

    Unicode normalization only — whitespace is content and is never trimmed.
    """
    normalized = unicodedata.normalize("NFC", text)
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()[:8]


@dataclass(frozen=True)
class ProvenanceStep:
    """One consulted level of an inheritance walk."""

    level: str  #: e.g. "layout placeholder lstStyle lvl1"
    part: Optional[str]  #: partname consulted, |None| for in-part levels
    detail: str  #: what was found (or not) at this level
    supplied: bool  #: True on the step that supplied the resolved value

    def to_dict(self) -> dict:
        return {
            "level": self.level,
            "part": self.part,
            "detail": self.detail,
            "supplied": self.supplied,
        }


@dataclass(frozen=True)
class EffectiveValue:
    """A resolved (or honestly-unresolved) effective value with its provenance chain."""

    value: object  #: EMU int for sizes, str for names, "RRGGBB" str for colors; None if unresolved
    value_pt: Optional[float]  #: convenience for sizes only, never instead of the EMU value
    resolved: bool
    provenance: Tuple[ProvenanceStep, ...]

    def to_dict(self) -> dict:
        return {
            "value": int(self.value) if isinstance(self.value, int) else self.value,
            "value_pt": self.value_pt,
            "resolved": self.resolved,
            "provenance": [step.to_dict() for step in self.provenance],
        }


@dataclass(frozen=True)
class EffectiveFont:
    """Effective size, name, and color of one run."""

    size: EffectiveValue
    name: EffectiveValue
    color_rgb: EffectiveValue

    def to_dict(self) -> dict:
        return {
            "schema": "paper-effective-font",
            "version": 1,
            "size": self.size.to_dict(),
            "name": self.name.to_dict(),
            "color_rgb": self.color_rgb.to_dict(),
        }


@dataclass(frozen=True)
class BlockAnchor:
    """The pinned anchor shape: part + block index + content hash (detects staleness)."""

    part: str
    block_index: int
    content_hash: str

    def to_dict(self) -> dict:
        return {
            "part": self.part,
            "block_index": self.block_index,
            "content_hash": self.content_hash,
        }


@dataclass(frozen=True)
class InspectedRun:
    text: str
    font: EffectiveFont

    def to_dict(self) -> dict:
        return {"text": self.text, "font": self.font.to_dict()}


@dataclass(frozen=True)
class TextBlock:
    """One paragraph of one shape, with anchor and per-run effective fonts."""

    anchor: BlockAnchor
    shape_id: int
    shape_name: str
    placeholder_type: Optional[str]
    level: int
    text: str
    runs: Tuple[InspectedRun, ...]

    def to_dict(self) -> dict:
        return {
            "anchor": self.anchor.to_dict(),
            "shape_id": self.shape_id,
            "shape_name": self.shape_name,
            "placeholder_type": self.placeholder_type,
            "level": self.level,
            "text": self.text,
            "runs": [run.to_dict() for run in self.runs],
        }


@dataclass(frozen=True)
class TextInspection:
    """Inspection payload for one slide. `.to_dict()` is deterministic (golden-tested)."""

    part: str
    blocks: Tuple[TextBlock, ...] = field(default_factory=tuple)

    def to_dict(self) -> dict:
        return {
            "schema": SCHEMA_NAME,
            "version": SCHEMA_VERSION,
            "part": self.part,
            "blocks": [block.to_dict() for block in self.blocks],
        }


def effective_font(run: "_Run") -> EffectiveFont:
    """Return the |EffectiveFont| for `run`, a `_Run` on a slide shape.

    Raises |UnsupportedStructureError| for runs outside a `p:sp` shape on a slide part
    (table-cell and chart text are out of v0 scope).
    """
    r = run._r
    sp = _ancestor_sp(r)
    if sp is None:
        raise UnsupportedStructureError(
            "effective_font resolves runs inside p:sp shapes only in v0 (got a run outside"
            " any p:sp, e.g. table-cell or chart text)"
        )
    part = run.part
    if not hasattr(part, "slide_layout"):
        raise UnsupportedStructureError(
            "effective_font resolves runs on slide parts only in v0, got %r"
            % type(part).__name__
        )
    return _FontResolver(part).effective_font(r, sp)


def inspect_text(slide: "Slide") -> TextInspection:
    """Return a |TextInspection| of every text-bearing `p:sp` shape on `slide`.

    Blocks appear in spTree (z-)order, paragraphs in document order; `block_index` numbers
    them consecutively within the slide part. Table and chart text are out of v0 scope and do
    not appear.
    """
    part = slide.part
    partname = str(part.partname)
    resolver = _FontResolver(part)
    blocks = []
    block_index = 0
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        sp = shape._element
        placeholder_type = (
            shape.placeholder_format.type.name if shape.is_placeholder else None
        )
        for paragraph in shape.text_frame.paragraphs:
            runs = tuple(
                InspectedRun(r.text, resolver.effective_font(r._r, sp))
                for r in paragraph.runs
            )
            text = "".join(inspected.text for inspected in runs)
            # -- raw read: the paragraph.level proxy getter is get-or-add and would mutate
            pPr = paragraph._p.find(qn("a:pPr"))
            level = int(pPr.get("lvl", "0")) if pPr is not None else 0
            blocks.append(
                TextBlock(
                    anchor=BlockAnchor(partname, block_index, content_hash(text)),
                    shape_id=shape.shape_id,
                    shape_name=shape.name,
                    placeholder_type=placeholder_type,
                    level=level,
                    text=text,
                    runs=runs,
                )
            )
            block_index += 1
    return TextInspection(part=partname, blocks=tuple(blocks))


def _ancestor_sp(element):
    """Return the nearest `p:sp` ancestor of `element`, or None."""
    parent = element.getparent()
    sp_tag = qn("p:sp")
    while parent is not None and parent.tag != sp_tag:
        parent = parent.getparent()
    return parent


class _FontResolver(object):
    """Executes the pinned inheritance walk for runs on one slide part. Strictly read-only."""

    def __init__(self, slide_part):
        self._slide_part = slide_part
        self._layout = slide_part.slide_layout
        self._master = self._layout.slide_master
        self._theme_root, self._theme_partname = self._load_theme()

    # ------------------------------------------------------------------------- public

    def effective_font(self, r, sp) -> EffectiveFont:
        p = r.getparent()
        pPr = p.find(qn("a:pPr"))
        level = int(pPr.get("lvl", "0")) if pPr is not None else 0
        if not 0 <= level <= 8:
            raise UnsupportedStructureError(
                "paragraph carries out-of-schema indent level lvl=%d (valid range 0..8);"
                " refusing to resolve styles for malformed structure" % level
            )
        chain = list(self._chain(r, p, sp, level))
        return EffectiveFont(
            size=self._resolve_size(chain),
            name=self._resolve_name(chain),
            color_rgb=self._resolve_color(chain, sp),
        )

    # ------------------------------------------------------------------- chain assembly

    def _chain(self, r, p, sp, level) -> Iterator[Tuple[str, str, object]]:
        """Generate (level-label, partname, rPr-like element or None) in consultation order."""
        slide_partname = str(self._slide_part.partname)
        yield "run rPr", slide_partname, r.find(qn("a:rPr"))

        pPr = p.find(qn("a:pPr"))
        yield (
            "paragraph defRPr",
            slide_partname,
            pPr.find(qn("a:defRPr")) if pPr is not None else None,
        )

        txBody = p.getparent()
        yield (
            "shape lstStyle lvl%d" % (level + 1),
            slide_partname,
            self._defRPr_from_lstStyle(txBody.find(qn("a:lstStyle")), level),
        )

        ph = sp.find(qn("p:nvSpPr") + "/" + qn("p:nvPr") + "/" + qn("p:ph"))
        if ph is not None:
            for step in self._placeholder_chain(sp, level):
                yield step
        else:
            yield (
                "presentation defaultTextStyle lvl%d" % (level + 1),
                str(self._presentation_part.partname),
                self._defRPr_from_lstStyle(
                    self._presentation_part._element.find(qn("p:defaultTextStyle")), level
                ),
            )

    def _placeholder_chain(self, sp, level):
        idx, ph_type = sp.ph_idx, sp.ph_type
        layout_ph = self._layout.placeholders.get(idx=idx)
        if layout_ph is None:
            layout_ph = next(
                (ph for ph in self._layout.placeholders if ph.element.ph_type == ph_type),
                None,
            )
        yield (
            "layout placeholder lstStyle lvl%d" % (level + 1),
            str(self._layout.part.partname),
            self._ph_lstStyle_defRPr(layout_ph, level),
        )

        master_ph_type = (
            PP_PLACEHOLDER.TITLE if ph_type in _TITLE_FAMILY else PP_PLACEHOLDER.BODY
        )
        master_ph = self._master.placeholders.get(master_ph_type)
        yield (
            "master placeholder lstStyle lvl%d" % (level + 1),
            str(self._master.part.partname),
            self._ph_lstStyle_defRPr(master_ph, level),
        )

        family, style = self._master_family_style(ph_type)
        yield (
            "master txStyles %s lvl%d" % (family, level + 1),
            str(self._master.part.partname),
            self._defRPr_from_lstStyle(style, level),
        )

    def _master_family_style(self, ph_type):
        txStyles = self._master.element.txStyles
        if ph_type in _TITLE_FAMILY:
            return "titleStyle", txStyles.titleStyle if txStyles is not None else None
        if ph_type in _BODY_FAMILY:
            return "bodyStyle", txStyles.bodyStyle if txStyles is not None else None
        return "otherStyle", txStyles.otherStyle if txStyles is not None else None

    @staticmethod
    def _ph_lstStyle_defRPr(placeholder_proxy, level):
        if placeholder_proxy is None:
            return None
        txBody = placeholder_proxy._element.find(qn("p:txBody"))
        if txBody is None:
            return None
        return _FontResolver._defRPr_from_lstStyle(txBody.find(qn("a:lstStyle")), level)

    @staticmethod
    def _defRPr_from_lstStyle(lstStyle, level):
        if lstStyle is None:
            return None
        lvl_pPr = lstStyle.pPr_for_lvl(level)
        if lvl_pPr is None:
            return None
        return lvl_pPr.find(qn("a:defRPr"))

    # ------------------------------------------------------------------------ resolvers

    def _resolve_size(self, chain) -> EffectiveValue:
        steps = []
        for label, partname, rPr in chain:
            sz = rPr.get("sz") if rPr is not None else None
            if sz is not None:
                steps.append(ProvenanceStep(label, partname, 'sz="%s"' % sz, True))
                size = Length(Centipoints(int(sz)))
                return EffectiveValue(size, size.pt, True, tuple(steps))
            steps.append(ProvenanceStep(label, partname, self._absence(rPr, "no size"), False))
        return EffectiveValue(None, None, False, tuple(steps))

    def _resolve_name(self, chain) -> EffectiveValue:
        steps = []
        for label, partname, rPr in chain:
            latin = rPr.find(qn("a:latin")) if rPr is not None else None
            typeface = latin.get("typeface") if latin is not None else None
            if typeface is None:
                steps.append(
                    ProvenanceStep(label, partname, self._absence(rPr, "no latin typeface"), False)
                )
                continue
            if typeface.startswith("+"):
                steps.append(
                    ProvenanceStep(
                        label, partname, 'latin typeface="%s" (theme reference)' % typeface, False
                    )
                )
                return self._resolve_theme_font(typeface, steps)
            steps.append(
                ProvenanceStep(label, partname, 'latin typeface="%s"' % typeface, True)
            )
            return EffectiveValue(typeface, None, True, tuple(steps))
        return EffectiveValue(None, None, False, tuple(steps))

    def _resolve_theme_font(self, token, steps) -> EffectiveValue:
        scheme_element = {"mj": "a:majorFont", "mn": "a:minorFont"}.get(token[1:3])
        if self._theme_root is None or scheme_element is None or not token.endswith("-lt"):
            steps.append(
                ProvenanceStep(
                    "theme fontScheme",
                    self._theme_partname,
                    "unresolvable theme font reference %r" % token,
                    False,
                )
            )
            return EffectiveValue(None, None, False, tuple(steps))
        latin = self._theme_root.find(
            qn("a:themeElements") + "/" + qn("a:fontScheme") + "/" + qn(scheme_element)
            + "/" + qn("a:latin")
        )
        typeface = latin.get("typeface") if latin is not None else None
        if not typeface:
            steps.append(
                ProvenanceStep(
                    "theme fontScheme",
                    self._theme_partname,
                    "%s has no latin typeface" % scheme_element,
                    False,
                )
            )
            return EffectiveValue(None, None, False, tuple(steps))
        steps.append(
            ProvenanceStep(
                "theme fontScheme %s" % scheme_element[2:],
                self._theme_partname,
                'latin typeface="%s"' % typeface,
                True,
            )
        )
        return EffectiveValue(typeface, None, True, tuple(steps))

    def _resolve_color(self, chain, sp) -> EffectiveValue:
        steps = []
        for label, partname, rPr in chain:
            solidFill = rPr.find(qn("a:solidFill")) if rPr is not None else None
            if solidFill is None:
                fill_kind = self._non_solid_fill_kind(rPr)
                if fill_kind is not None:
                    steps.append(
                        ProvenanceStep(
                            label, partname, "%s (not a single solid color)" % fill_kind, False
                        )
                    )
                    return EffectiveValue(None, None, False, tuple(steps))
                steps.append(
                    ProvenanceStep(label, partname, self._absence(rPr, "no solidFill"), False)
                )
                continue
            srgbClr = solidFill.find(qn("a:srgbClr"))
            if srgbClr is not None:
                detail = 'srgbClr val="%s"%s' % (
                    srgbClr.get("val"),
                    self._transforms_note(srgbClr),
                )
                steps.append(ProvenanceStep(label, partname, detail, True))
                return EffectiveValue(srgbClr.get("val").upper(), None, True, tuple(steps))
            schemeClr = solidFill.find(qn("a:schemeClr"))
            if schemeClr is not None:
                steps.append(
                    ProvenanceStep(
                        label,
                        partname,
                        'schemeClr val="%s"%s'
                        % (schemeClr.get("val"), self._transforms_note(schemeClr)),
                        False,
                    )
                )
                return self._resolve_scheme_color(schemeClr.get("val"), steps)
            steps.append(
                ProvenanceStep(
                    label,
                    partname,
                    "solidFill with unsupported color element (%s)"
                    % ", ".join(
                        el.tag.rsplit("}", 1)[-1] for el in solidFill
                    ),
                    False,
                )
            )
            return EffectiveValue(None, None, False, tuple(steps))
        return EffectiveValue(None, None, False, tuple(steps))

    def _resolve_scheme_color(self, token, steps) -> EffectiveValue:
        slot, map_step = self._map_scheme_token(token)
        steps.append(map_step)
        if slot is None:
            return EffectiveValue(None, None, False, tuple(steps))
        if self._theme_root is None:
            steps.append(
                ProvenanceStep(
                    "theme clrScheme", self._theme_partname, "theme part unavailable", False
                )
            )
            return EffectiveValue(None, None, False, tuple(steps))
        slot_element = self._theme_root.find(
            qn("a:themeElements") + "/" + qn("a:clrScheme") + "/" + qn("a:%s" % slot)
        )
        color_child = slot_element[0] if slot_element is not None and len(slot_element) else None
        if color_child is None:
            steps.append(
                ProvenanceStep(
                    "theme clrScheme %s" % slot,
                    self._theme_partname,
                    "slot missing from clrScheme",
                    False,
                )
            )
            return EffectiveValue(None, None, False, tuple(steps))
        local = color_child.tag.rsplit("}", 1)[-1]
        if local == "srgbClr":
            rgb = color_child.get("val").upper()
            detail = 'srgbClr val="%s"' % color_child.get("val")
        elif local == "sysClr":
            rgb = (color_child.get("lastClr") or "").upper() or None
            detail = 'sysClr val="%s" lastClr="%s"' % (
                color_child.get("val"),
                color_child.get("lastClr"),
            )
        else:
            rgb, detail = None, "unsupported clrScheme child %s" % local
        steps.append(
            ProvenanceStep(
                "theme clrScheme %s" % slot, self._theme_partname, detail, rgb is not None
            )
        )
        return EffectiveValue(rgb, None, rgb is not None, tuple(steps))

    def _map_scheme_token(self, token):
        """Return (theme-slot, ProvenanceStep) for schemeClr `token`."""
        override = self._clrMapOvr_mapping()
        if override is not None:
            mapping, source_label, source_part = override
        else:
            clrMap = self._master.element.find(qn("p:clrMap"))
            mapping = dict(clrMap.attrib) if clrMap is not None else {}
            source_label, source_part = "master clrMap", str(self._master.part.partname)
        if token in mapping:
            slot = mapping[token]
            return slot, ProvenanceStep(
                source_label, source_part, '%s="%s"' % (token, slot), False
            )
        if token in _DIRECT_THEME_SLOTS:
            return token, ProvenanceStep(
                source_label,
                source_part,
                '"%s" addresses the theme slot directly (no mapping entry)' % token,
                False,
            )
        return None, ProvenanceStep(
            source_label, source_part, "unmappable scheme color token %r" % token, False
        )

    def _clrMapOvr_mapping(self):
        """Return (mapping, label, partname) from the slide's clrMapOvr, or None to defer."""
        clrMapOvr = self._slide_part._element.find(qn("p:clrMapOvr"))
        if clrMapOvr is None:
            return None
        overrideClrMapping = clrMapOvr.find(qn("a:overrideClrMapping"))
        if overrideClrMapping is None:
            return None  # -- a:masterClrMapping (or empty): defer to the master's clrMap
        return (
            dict(overrideClrMapping.attrib),
            "slide clrMapOvr",
            str(self._slide_part.partname),
        )

    # ------------------------------------------------------------------------- helpers

    @staticmethod
    def _transforms_note(color_element) -> str:
        """Return a note naming any color-transform children (lumMod, tint, …), or ""."""
        transforms = [
            "%s=%s" % (child.tag.rsplit("}", 1)[-1], child.get("val"))
            for child in color_element
        ]
        if not transforms:
            return ""
        return " with unapplied transforms [%s]" % ", ".join(transforms)

    @staticmethod
    def _non_solid_fill_kind(rPr):
        if rPr is None:
            return None
        for tag in ("a:gradFill", "a:blipFill", "a:pattFill", "a:noFill", "a:grpFill"):
            if rPr.find(qn(tag)) is not None:
                return tag
        return None

    @staticmethod
    def _absence(rPr, what) -> str:
        return "level not present" if rPr is None else what + " here"

    @property
    def _presentation_part(self):
        return self._slide_part.package.presentation_part

    def _load_theme(self):
        try:
            theme_part = self._master.part.part_related_by(RT.THEME)
        except KeyError:
            return None, None
        try:
            root = parse_xml(theme_part.blob)
        except Exception:
            raise UnsupportedStructureError(
                "theme part %s is not parseable XML" % theme_part.partname
            )
        return root, str(theme_part.partname)
