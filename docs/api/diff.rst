.. _diff_api:

Deck diff (``pptx.diff``)
=========================

*paper-pptx addition.* The verification mirror — git-diff for decks. PPTX carries no revision
markup, so "what changed between v3 and v4?" has had no programmable answer anywhere in the
ecosystem. :func:`diff_decks` provides one as a typed report beside the file: slides added,
removed, or **moved** (matched by permanent slide id, so a reorder reads as a move rather than a
delete-plus-add), and within matched slides the shape, text, chart-data, image, and notes
deltas. At ``detail="full"`` it also reports per-run effective-value shifts via the resolver.

It is report-only — presenting or rendering the diff is harness territory — and it is the
independent witness the release's job evals check every operation report against: on each
import / rebind / refresh job, the operation's own report and ``diff_decks(input, output)`` must
agree.

Matching is by permanent slide id, which serves lineage-derived decks (v4 saved from v3). Decks
rebuilt from scratch get fresh ids and will not match; a content-fingerprint fallback is a
declared future option, not a promise here.

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
