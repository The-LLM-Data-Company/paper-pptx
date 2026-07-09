.. _slides_api:

Slides
======

|Slides| objects
-----------------

The |Slides| object is accessed using the
:attr:`~pptx.presentation.Presentation.slides` property of |Presentation|. It
is not intended to be constructed directly.

.. autoclass:: pptx.slide.Slides()
   :members:
   :member-order: bysource
   :undoc-members:


|Slide| objects
---------------

An individual |Slide| object is accessed by index from |Slides| or as the
return value of :meth:`add_slide`.

.. autoclass:: pptx.slide.Slide()
   :members:
   :exclude-members: part
   :inherited-members:
   :undoc-members:


|SlideLayouts| objects
----------------------

The |SlideLayouts| object is accessed using the
:attr:`~pptx.slide.SlideMaster.slide_layouts` property of |SlideMaster|, typically::

    >>> from pptx import Presentation
    >>> prs = Presentation()
    >>> slide_layouts = prs.slide_master.slide_layouts

As a convenience, since most presentations have only a single slide master, the
|SlideLayouts| collection for the first master may be accessed directly from the
|Presentation| object::

    >>> slide_layouts = prs.slide_layouts

This class is not intended to be constructed directly.

.. autoclass:: pptx.slide.SlideLayouts()
   :members:
   :exclude-members: element, parent
   :inherited-members:
   :undoc-members:


|SlideLayout| objects
---------------------

.. autoclass:: pptx.slide.SlideLayout
   :members:
   :exclude-members: iter_cloneable_placeholders


|SlideMasters| objects
----------------------

The |SlideMasters| object is accessed using the
:attr:`~pptx.presentation.slide_masters` property of |Presentation|, typically::

    >>> from pptx import Presentation
    >>> prs = Presentation()
    >>> slide_masters = prs.slide_masters

As a convenience, since most presentations have only a single slide master, the
first master may be accessed directly from the |Presentation| object without indexing
the collection::

    >>> slide_master = prs.slide_master

This class is not intended to be constructed directly.

.. autoclass:: pptx.slide.SlideMasters()
   :members:
   :exclude-members: element, parent
   :inherited-members:
   :undoc-members:

|SlideMaster| objects
---------------------

.. autoclass:: pptx.slide.SlideMaster
   :members:
   :exclude-members: related_slide_layout, sldLayoutIdLst


|SlidePlaceholders| objects
---------------------------

.. autoclass:: pptx.shapes.shapetree.SlidePlaceholders
   :members:


|NotesSlide| objects
--------------------

.. autoclass:: pptx.slide.NotesSlide
   :members:
   :exclude-members: clone_master_placeholders
   :inherited-members:


|SlideClonePolicy| objects
--------------------------

*paper-pptx addition.* The policy object accepted by :meth:`.Slides.clone`, controlling how a
cloned slide's related parts are handled.

.. autoclass:: pptx.slide.SlideClonePolicy()
   :members:
   :undoc-members:
   :member-order: bysource


|HeaderFooters| objects
-----------------------

*paper-pptx addition.* Tri-state (``True`` / ``False`` / ``None`` = inherit) visibility flags for
the date, footer, and slide-number placeholders, read from a layout or master through the
:attr:`.SlideLayout.header_footers` / :attr:`.SlideMaster.header_footers` properties.

.. autoclass:: pptx.slide.HeaderFooters()
   :members:
