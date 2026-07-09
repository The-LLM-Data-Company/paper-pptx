.. _hf_api:

Fields and footers (``pptx.hf``)
================================

*paper-pptx addition.* Apply footer, date, and slide-number content the way PowerPoint's
Insert → Header & Footer dialog does — materializing the placeholder shapes on each slide and
writing slide numbers and dates as real ``a:fld`` fields, never as static text.

This package **authors** fields; it never computes their values. A field's cached text is a hint
that PowerPoint and LibreOffice refresh when they open the file, so a slide number written here
stays correct after a later reorder — retiring the static-page-number anti-pattern that
paper-pptx's own slide-reordering made hazardous.

The entry points are methods on |Presentation| and |Slide|; see
:meth:`.Presentation.apply_footers` and :meth:`.Slide.apply_footers`.

The ``date_format`` argument accepts the ISO/IEC 29500 date field tokens, held in the module
constant ``pptx.hf.DATETIME_FIELD_FORMATS``: ``"datetime"`` (the rendering application's default
format) and ``"datetime1"`` through ``"datetime13"`` (the fixed format variants, e.g.
``"datetime1"`` → ``MM/DD/YYYY``, ``"datetime3"`` → ``DD Month YYYY``). ``fixed_date`` instead
writes a literal, non-updating date string.
