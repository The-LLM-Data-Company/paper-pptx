"""Main presentation object."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, cast

from pptx.shared import PartElementProxy
from pptx.slide import SlideMasters, Slides
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from datetime import datetime

    from pptx.oxml.presentation import CT_Presentation, CT_SlideId
    from pptx.parts.presentation import PresentationPart
    from pptx.slide import NotesMaster, SlideLayouts
    from pptx.util import Length


class Presentation(PartElementProxy):
    """PresentationML (PML) presentation.

    Not intended to be constructed directly. Use :func:`pptx.Presentation` to open or
    create a presentation.
    """

    _element: CT_Presentation
    part: PresentationPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def apply_footers(
        self,
        *,
        footer: str | None = None,
        slide_number: bool = False,
        date_format: str | None = None,
        fixed_date: str | None = None,
        skip_title_slides: bool = False,
        now: "datetime | None" = None,
    ) -> None:
        """Apply the complete footer state to every slide ("Apply to All").

        paper-pptx addition (v0.11 Phase 2). Persists exactly what PowerPoint's
        Insert > Header & Footer dialog does: materializes minimal `dt`/`ftr`/`sldNum`
        placeholder shapes per slide (binding to the layout furniture by `idx`), writes
        slide numbers and automatic dates as real `a:fld` elements whose cached text
        consumers refresh on open, and *removes* the placeholders for unchecked elements —
        each call sets the full three-element state, like the dialog.

        `footer`: literal footer text, or None to remove the footer placeholder.
        `slide_number`: True writes a `slidenum` field cached with the current position
        (honoring `firstSlideNum`); consumers renumber live after any reorder.
        `date_format`: a `datetime`..`datetime13` token for an automatically-updating date
        field; `fixed_date`: literal date text (the dialog's "Fixed" mode); passing both
        raises |ValueError|. `now` seeds the date field's cached text (None = wall clock);
        the package never vouches for cached values — they are consumer-refreshed hints.
        `skip_title_slides`: the dialog's "Don't show on title slide" — slides on a
        `type="title"` layout get the all-removed state.

        Refuses atomically (|UnsupportedStructureError|, validated deck-wide before the
        first write) when a wanted element has no layout furniture to inherit from, or
        when explicit `p:hf` flags on a layout/master disable it (clear those via
        `header_footers` first — this API never flips them silently).
        """
        from pptx.hf import apply_presentation_footers

        apply_presentation_footers(
            self,
            footer=footer,
            slide_number=slide_number,
            date_format=date_format,
            fixed_date=fixed_date,
            skip_title_slides=skip_title_slides,
            now=now,
        )

    @property
    def core_properties(self):
        """|CoreProperties| instance for this presentation.

        Provides read/write access to the Dublin Core document properties for the presentation.
        """
        return self.part.core_properties

    @property
    def notes_master(self) -> NotesMaster:
        """Instance of |NotesMaster| for this presentation.

        If the presentation does not have a notes master, one is created from a default template
        and returned. The same single instance is returned on each call.
        """
        return self.part.notes_master

    def save(self, file: str | IO[bytes]):
        """Writes this presentation to `file`.

        `file` can be either a file-path or a file-like object open for writing bytes.
        """
        self.part.save(file)

    @property
    def slide_height(self) -> Length | None:
        """Height of slides in this presentation, in English Metric Units (EMU).

        Returns |None| if no slide width is defined. Read/write.
        """
        sldSz = self._element.sldSz
        if sldSz is None:
            return None
        return sldSz.cy

    @slide_height.setter
    def slide_height(self, height: Length):
        sldSz = self._element.get_or_add_sldSz()
        sldSz.cy = height

    @property
    def slide_layouts(self) -> SlideLayouts:
        """|SlideLayouts| collection belonging to the first |SlideMaster| of this presentation.

        A presentation can have more than one slide master and each master will have its own set
        of layouts. This property is a convenience for the common case where the presentation has
        only a single slide master.
        """
        return self.slide_masters[0].slide_layouts

    @property
    def slide_master(self):
        """
        First |SlideMaster| object belonging to this presentation. Typically,
        presentations have only a single slide master. This property provides
        simpler access in that common case.
        """
        return self.slide_masters[0]

    @lazyproperty
    def slide_masters(self) -> SlideMasters:
        """|SlideMasters| collection of slide-masters belonging to this presentation."""
        return SlideMasters(self._element.get_or_add_sldMasterIdLst(), self)

    @property
    def slide_width(self):
        """
        Width of slides in this presentation, in English Metric Units (EMU).
        Returns |None| if no slide width is defined. Read/write.
        """
        sldSz = self._element.sldSz
        if sldSz is None:
            return None
        return sldSz.cx

    @slide_width.setter
    def slide_width(self, width: Length):
        sldSz = self._element.get_or_add_sldSz()
        sldSz.cx = width

    @lazyproperty
    def slides(self):
        """|Slides| object containing the slides in this presentation."""
        sldIdLst = self._element.get_or_add_sldIdLst()
        self.part.rename_slide_parts([cast("CT_SlideId", sldId).rId for sldId in sldIdLst])
        return Slides(sldIdLst, self)
