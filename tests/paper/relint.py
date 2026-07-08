"""Relationship-integrity scans, mined from the reference verifier's mechanical checks.

Two failure classes turn a deck that "opens in python-pptx" into one PowerPoint rejects:

- a part XML references an `r:id` that its `.rels` item does not define
  (`missing_relationship_references`), and
- a `.rels` item targets a package member that does not exist
  (`dangling_relationship_targets`).

Both scans work on the raw zip members - never through the pptx object model - so they stay
valid as an independent oracle when the object model itself is the code under test. Phase 7
(slide clone/delete/reorder) runs these over every output; corrupt-by-construction fixtures
prove they actually detect what they claim to.
"""

from __future__ import annotations

import posixpath
from typing import Dict, List, Tuple

from lxml import etree

_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _source_member_for_rels(rels_member: str) -> str:
    """Return the member a `.rels` item belongs to ('' for the package rels)."""
    directory, _, filename = rels_member.rpartition("/")
    parent = posixpath.dirname(directory)  # -- strip the trailing "_rels"
    source_name = filename[: -len(".rels")]
    return posixpath.join(parent, source_name) if parent else source_name


def _resolve_target(rels_member: str, target: str) -> str:
    """Resolve a relationship `target` to a zip member name.

    Targets beginning with "/" are package-root-relative (legal OPC; some producers emit
    them); all others resolve against the source part's base directory.
    """
    if target.startswith("/"):
        return posixpath.normpath(target).lstrip("/")
    base_dir = posixpath.dirname(posixpath.dirname(rels_member))
    return posixpath.normpath(posixpath.join(base_dir, target))


def relationship_ids_by_member(zip_map: "Dict[str, bytes]") -> "Dict[str, Dict[str, str]]":
    """Return {source-member: {rId: target}} for every `.rels` item in the package."""
    rels_by_member = {}
    for name, blob in zip_map.items():
        if not name.endswith(".rels"):
            continue
        source = _source_member_for_rels(name)
        rels_by_member[source] = {
            rel.get("Id"): rel.get("Target")
            for rel in etree.fromstring(blob).iter("{%s}Relationship" % _RELS_NS)
        }
    return rels_by_member


def dangling_relationship_targets(
    zip_map: "Dict[str, bytes]",
) -> "List[Tuple[str, str, str]]":
    """Return (rels-member, rId, resolved-target) for internal targets that do not exist."""
    problems = []
    for name, blob in sorted(zip_map.items()):
        if not name.endswith(".rels"):
            continue
        for rel in etree.fromstring(blob).iter("{%s}Relationship" % _RELS_NS):
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target")
            if not target:
                # -- report rather than crash: a Target-less Relationship is itself corrupt,
                # -- and one malformed rel must not abort the scan before later real problems
                problems.append((name, rel.get("Id"), "<missing Target attribute>"))
                continue
            resolved = _resolve_target(name, target)
            if resolved not in zip_map:
                problems.append((name, rel.get("Id"), resolved))
    return problems


def missing_relationship_references(
    zip_map: "Dict[str, bytes]",
) -> "List[Tuple[str, str]]":
    """Return (member, rId) for every r-namespace attribute naming an undefined relationship."""
    rels_by_member = relationship_ids_by_member(zip_map)
    problems = []
    for name, blob in sorted(zip_map.items()):
        if not name.endswith(".xml") or name.endswith(".rels"):
            continue
        try:
            root = etree.fromstring(blob)
        except etree.XMLSyntaxError:
            continue  # -- not this scan's failure class; malformed XML fails other oracles
        defined = rels_by_member.get(name, {})
        for element in root.iter():
            for attr_name, attr_value in element.attrib.items():
                is_rel_ref = attr_name.startswith("{%s}" % _R_NS) and attr_value
                if is_rel_ref and attr_value not in defined:
                    problems.append((name, attr_value))
    return problems
