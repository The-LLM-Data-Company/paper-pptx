.. _diff_api:

Deck diff (``pptx.diff``)
=========================

*paper-pptx addition.* :func:`diff_decks` compares two decks and returns a typed report. It lists
slides added, removed, or **moved** (matched by permanent slide id, so a reorder is reported as a
move rather than delete-plus-add) and the shape, text, chart-data, image, and notes changes within
matched slides. At ``detail="full"`` it also reports per-run effective-value shifts via the
resolver.

Release job evaluations compare each import / rebind / refresh operation report with
``diff_decks(input, output)`` and require them to agree.

Matching uses the permanent slide id, which serves lineage-derived decks (v4 saved from v3). Decks
rebuilt from scratch get fresh ids and will not match.

.. currentmodule:: pptx.diff

.. autofunction:: diff_decks

.. autoclass:: DeckDiff()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: SlideChange()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: SlideRef()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: MovedSlide()
   :members:
   :undoc-members:
   :member-order: bysource
