*paper-pptx* is Paper Instruments' hard fork of *python-pptx* (from upstream ``v1.0.2``) — a
Python library for creating, reading, and updating PowerPoint (.pptx) files, extended into a
library for safely **inspecting, editing, and composing** real, brand-driven, multi-source
decks. It remains a 100% drop-in replacement: ``from pptx import Presentation`` and every other
existing call keep working, unchanged.

Stock python-pptx is a superb deck *generator* — reliable on files its own code created.
paper-pptx adds the other three verbs professional deck work needs: perceive an existing deck
(effective font/color resolution through the placeholder → layout → master → theme chain, with
provenance), edit it safely (anchored formatting-preserving text replacement, slide and shape
and table surgery, image and chart operations — all validate-fully-then-mutate, refusing loudly
rather than corrupting), and assemble it from other decks (cross-file slide import and merge,
layout rebind, send-safe scrub, and a deck-to-deck diff that proves what changed). It runs on
any Python-capable platform, macOS and Linux included, and does not require PowerPoint to be
installed or licensed.

The founding commitment is against *silent wrongness* — the deck that opens fine and lies. Every
mutating operation either does exactly what it claims (proven by save → reopen, exact
changed-part budgets, and an independent LibreOffice load smoke) or raises a typed refusal and
leaves the document byte-identical.

The paper additions are surveyed in :ref:`paper_additions` (start there); each added module has
its own page under `API Documentation`_. The rest of this documentation is inherited from
python-pptx and describes the shared, unchanged foundation.
