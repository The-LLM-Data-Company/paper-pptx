.. _errors_api:

Refusals (``pptx.errors``)
==========================

*paper-pptx addition.* Every added operation either does exactly what it claims or refuses
atomically. Mutating operations follow *validate-fully-then-mutate*. If one cannot proceed
safely, it raises a typed :exc:`.PaperRefusal` and leaves the document byte-for-byte unchanged
in memory and on disk. Programmer errors remain plain ``ValueError`` / ``TypeError``. Callers can
handle unsupported deck operations separately by catching |PaperRefusal|.

This module is distinct from :mod:`pptx.exc`, which holds the exceptions inherited from
python-pptx.

.. currentmodule:: pptx.errors

.. autoexception:: PaperRefusal
   :show-inheritance:

.. autoexception:: AmbiguousTargetError
   :show-inheritance:

.. autoexception:: TargetNotFoundError
   :show-inheritance:

.. autoexception:: StaleAnchorError
   :show-inheritance:

.. autoexception:: UnsupportedStructureError
   :show-inheritance:

.. autoexception:: BoundaryViolationError
   :show-inheritance:

.. autoexception:: RelationshipPolicyError
   :show-inheritance:
