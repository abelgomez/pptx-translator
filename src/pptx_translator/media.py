"""Detection and removal of embedded audio from a presentation.

Audio objects in PowerPoint are internally represented as ``<p:pic>`` shapes
whose ``<p:nvPr>`` section contains an ``<a:audioFile>`` tag. This module
detects those shapes (whether or not they are inside a group) and removes
them from the slide's XML tree, together with:

* Their relationship entries (``audio``/``media``), to avoid leaving
  dangling relationship links.
* Any leftover reference to the removed shape in the slide's
  ``<p:timing>`` animation tree, which is what PowerPoint uses to trigger
  "play automatically"/"play on click" audio behavior. Leaving that
  reference in place after deleting the shape is what makes PowerPoint
  flag the file as needing repair, even though it can still open it.
"""

from __future__ import annotations

import logging

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.slide import Slide

logger = logging.getLogger(__name__)

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _is_audio_shape(shape) -> bool:
    element = shape._element  # noqa: SLF001 - intentional access to the underlying XML
    return element.find(f".//{{{_A_NS}}}audioFile") is not None


def _shape_id(shape) -> str | None:
    """Returns the shape's own numeric id (``<p:cNvPr id="...">``).

    Needed to later find and remove any dangling reference to this shape
    elsewhere in the slide (namely in the ``<p:timing>`` animation tree)
    once the shape itself has been deleted.
    """

    c_nv_pr = shape._element.find(f".//{qn('p:cNvPr')}")  # noqa: SLF001
    if c_nv_pr is None:
        return None
    return c_nv_pr.attrib.get("id")


def _remove_audio_relationship(shape) -> None:
    """Removes the relationship entries that link the audio media file.

    Removing the shape alone is not enough: the slide may retain both the
    ``audio`` relationship and the ``media`` relationship pointing to the
    same target file. These are leftover (orphaned) links and can leave the
    package in a partially inconsistent state, even if PowerPoint still opens
    the file.
    """

    audio_file = shape._element.find(f".//{{{_A_NS}}}audioFile")  # noqa: SLF001
    if audio_file is None:
        return

    rel_id = audio_file.attrib.get(f"{{{_REL_NS}}}link")
    if not rel_id:
        return

    part = getattr(shape, "part", None)
    if part is None or not hasattr(part, "rels"):
        return

    target_part_name = None
    try:
        target_part = part.rels[rel_id].target_part
        target_part_name = getattr(target_part, "partname", None)
    except KeyError:
        target_part_name = None

    for rel_key, rel in list(part.rels.items()):
        if rel_key == rel_id:
            part.rels.pop(rel_key)
            continue
        if target_part_name is not None:
            rel_target = getattr(rel.target_part, "partname", None)
            if rel_target == target_part_name and rel.reltype in {
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio",
                "http://schemas.microsoft.com/office/2007/relationships/media",
            }:
                part.rels.pop(rel_key)


def _remove_audio_recursive(shapes) -> tuple[int, list[str]]:
    """Removes audio shapes recursively.

    Returns ``(removed_count, removed_shape_ids)`` so the caller can also
    clean up any leftover reference to those shape ids elsewhere in the
    slide (see :func:`_remove_orphaned_timing_nodes`).
    """

    removed = 0
    removed_ids: list[str] = []
    for shape in list(shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            sub_removed, sub_ids = _remove_audio_recursive(shape.shapes)
            removed += sub_removed
            removed_ids.extend(sub_ids)
        if _is_audio_shape(shape):
            shape_id = _shape_id(shape)
            _remove_audio_relationship(shape)
            parent = shape._element.getparent()  # noqa: SLF001
            if parent is not None:
                parent.remove(shape._element)  # noqa: SLF001
                removed += 1
                if shape_id:
                    removed_ids.append(shape_id)
    return removed, removed_ids


def _remove_orphaned_timing_nodes(slide: Slide, removed_shape_ids: list[str]) -> None:
    """Removes ``<p:timing>`` animation entries that target a removed shape.

    PowerPoint stores audio "play automatically"/"play on click" behavior
    as an entry in the slide's ``<p:timing>`` tree (a sibling of
    ``<p:cSld>``), which references the shape by its numeric id via
    ``<p:spTgt spid="...">``. Deleting only the shape and leaving this
    reference behind is what causes PowerPoint to report (and silently
    repair) a corrupted file.
    """

    if not removed_shape_ids:
        return

    sld = slide._element  # noqa: SLF001
    timing = sld.find(qn("p:timing"))
    if timing is None:
        return

    dangling_ids = set(removed_shape_ids)

    # Each animation entry PowerPoint generates for a shape is
    # self-contained inside its own <p:par> ("parallel time node") subtree,
    # a sibling of the entries for any other, still-valid animations. Drop
    # the nearest enclosing <p:par> of every dangling <p:spTgt> reference,
    # re-scanning after each removal since the tree gets mutated.
    changed = True
    while changed:
        changed = False
        for sp_tgt in timing.iter(qn("p:spTgt")):
            if sp_tgt.attrib.get("spid") not in dangling_ids:
                continue
            node = sp_tgt
            enclosing_par = None
            while node is not None:
                node = node.getparent()
                if node is not None and node.tag == qn("p:par"):
                    enclosing_par = node
                    break
            target = enclosing_par if enclosing_par is not None else sp_tgt
            target_parent = target.getparent()
            if target_parent is not None:
                target_parent.remove(target)
                changed = True
            break  # the tree was mutated; restart the scan
        else:
            break

    # If no valid target reference remains, the whole timing tree serves
    # no purpose and can be safely dropped.
    if timing.find(qn("p:spTgt")) is None:
        parent = timing.getparent()
        if parent is not None:
            parent.remove(timing)


def remove_audio_from_slide(slide: Slide) -> int:
    """Removes all audio shapes from a slide (recursively).

    Also cleans up any leftover animation/timing reference to the removed
    shapes, to avoid leaving the package in a state PowerPoint considers
    corrupted. Returns the number of audio shapes removed.
    """

    removed, removed_ids = _remove_audio_recursive(slide.shapes)
    _remove_orphaned_timing_nodes(slide, removed_ids)
    return removed
