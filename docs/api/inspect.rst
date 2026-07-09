.. _inspect_api:

Inspection (``pptx.inspect``)
=============================

*paper-pptx addition.* Read-only, provenance-bearing resolution of the values a deck actually
renders — the sizes, fonts, and colors that stock python-pptx returns as ``None`` because they
are inherited through the placeholder → layout → master → theme chain. Nothing here mutates the
document. Every resolved value can explain where it came from; a value that cannot be resolved is
reported as unresolved, never guessed.

.. currentmodule:: pptx.inspect


Functions
---------

.. autofunction:: effective_font

.. autofunction:: effective_paragraph_format

.. autofunction:: effective_shape_format

.. autofunction:: inspect_text

.. autofunction:: inspect_deck


Resolved values and provenance
-------------------------------

.. autoclass:: EffectiveValue()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: ProvenanceStep()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: EffectiveFont()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: EffectiveParagraphFormat()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: EffectiveShapeFormat()
   :members:
   :undoc-members:
   :member-order: bysource


Text inspection payloads
------------------------

.. autoclass:: BlockAnchor()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: InspectedRun()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: TextBlock()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: TextInspection()
   :members:
   :undoc-members:
   :member-order: bysource


Deck manifest payloads
----------------------

.. autoclass:: DeckManifest()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: SlideManifest()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: ShapeManifest()
   :members:
   :undoc-members:
   :member-order: bysource
