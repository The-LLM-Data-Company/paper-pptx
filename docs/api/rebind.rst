.. _rebind_api:

Layout rebind (``pptx.rebind``)
===============================

*paper-pptx addition.* Move a slide to another layout in the same presentation, matching
placeholders by type and index with an explicit map and an orphan policy for the rest.

The report is required. The effective-value resolver runs before and after, and every text run
whose *resolved* appearance changed appears in it.

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
