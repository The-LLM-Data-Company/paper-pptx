"""ID-list integrity scans: slide-id references outside the relationship graph.

Sections (`p14:sectionLst`, PowerPoint 2010+ extension) reference slides by *slide id*, not
relationship id, so `relint.py`'s r-namespace scans cannot see a dangling entry. A section
that names a deleted slide's id is exactly the "opens in python-pptx, breaks in PowerPoint"
corruption class the slide operations exist to eliminate — this scan makes the class visible to every
fixture test and every slide-op output test.

Custom shows (`p:custShowLst`) reference slides by `r:id`, which `relint.py` already covers;
`duplicate_section_slide_ids` covers the other section invariant (an id in two sections).
"""

from __future__ import annotations

from typing import Dict, List, Tuple

from lxml import etree

_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"

_PRESENTATION_MEMBER = "ppt/presentation.xml"


def _section_membership(presentation_xml) -> "List[Tuple[str, str]]":
    """Return (section-name, slide-id) for every p14:sldId entry, in document order."""
    return [
        (section.get("name"), sldId.get("id"))
        for section in presentation_xml.iter("{%s}section" % _P14)
        for sldId in section.iter("{%s}sldId" % _P14)
    ]


def dangling_section_slide_ids(zip_map: "Dict[str, bytes]") -> "List[Tuple[str, str]]":
    """Return (section-name, slide-id) for section entries naming no existing slide."""
    if _PRESENTATION_MEMBER not in zip_map:
        return []
    presentation_xml = etree.fromstring(zip_map[_PRESENTATION_MEMBER])
    sldIdLst = presentation_xml.find("{%s}sldIdLst" % _P)
    real_ids = (
        {sldId.get("id") for sldId in sldIdLst} if sldIdLst is not None else set()
    )
    return [
        (name, slide_id)
        for name, slide_id in _section_membership(presentation_xml)
        if slide_id not in real_ids
    ]


def duplicate_section_slide_ids(zip_map: "Dict[str, bytes]") -> "List[str]":
    """Return slide ids enrolled in more than one section (or twice in one)."""
    if _PRESENTATION_MEMBER not in zip_map:
        return []
    presentation_xml = etree.fromstring(zip_map[_PRESENTATION_MEMBER])
    seen: "Dict[str, int]" = {}
    for _, slide_id in _section_membership(presentation_xml):
        seen[slide_id] = seen.get(slide_id, 0) + 1
    return sorted(slide_id for slide_id, count in seen.items() if count > 1)
