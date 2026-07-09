.. _rebind_api:

Layout rebind (``pptx.rebind``)
===============================

*paper-pptx addition.* The template-migration *primitive*: move a slide to another layout in the
same presentation, matching placeholders by type and index (with an explicit map and an orphan
policy for the rest). Bulk-rebrand *orchestration* is deliberately out of scope — this is the
one safe primitive; the workflow that decides what maps where stays in the harness.

The report is required, not optional: the effective-value resolver runs before and after, and
every text run whose *resolved* appearance changed appears in it. Nothing about the look shifts
silently.

The entry point is a method on |Slide|; see :meth:`.Slide.rebind_layout`. This page documents the
report it returns.

.. currentmodule:: pptx.rebind

.. autoclass:: RebindReport()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: RunShift()
   :members:
   :undoc-members:
   :member-order: bysource
