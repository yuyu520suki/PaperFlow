"""Append end-of-manuscript figure and table copies in a separate DOCX."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import posixpath
import re

from lxml import etree as ET

from . import word
from .citations import _owner, _scan_fields
from .floating import ANCHOR, INLINE, inline_picture_copy
from .package import part_fields, q, read_package, write_package, xml


_PICTURE_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
_WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_PIC = f"{{{_PICTURE_NS}}}pic"
_DRAWING_IDS = {f"{{{_WP_NS}}}docPr", f"{{{_PICTURE_NS}}}cNvPr"}
_CLONED_ID_ATTRIBUTES = {
    "{http://schemas.microsoft.com/office/word/2010/wordml}paraId",
    "{http://schemas.microsoft.com/office/word/2010/wordml}textId",
    "{http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing}anchorId",
    "{http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing}editId",
}
_SEQ = re.compile(r'^\s*SEQ\s+(?:"(Figure|Fig\.|Table)"|(Figure|Fig\.|Table)(?=\s|$))', re.I)
_FREEZEABLE = re.compile(r"^\s*(?:SEQ|REF|PAGEREF|STYLEREF)\s", re.I)
_TOC = re.compile(r"^\s*TOC(?=\s|$)", re.I)
_FIELD_SWITCH = re.compile(r'"[^"]*"|(\\[a-z]+)\b', re.I)
_MAIN = "word/document.xml"


def _visible(node):
    return "".join(text.text or "" for text in node.iter(q("w:t")))


def _caption_kind(code):
    match = _SEQ.match(code)
    if match is None:
        return None
    return "table" if (match[1] or match[2]).lower() == "table" else "figure"


def _caption_candidates(children, index, kind, captions):
    if kind == "figure" and captions.get(index) == kind:
        return {index}
    picture_labels = {
        element.get("descr", "").strip()
        for element in children[index].iter(f"{{{_WP_NS}}}docPr")
    } - {""}
    matches = []
    for direction in (-1, 1):
        candidate = index + direction
        while 0 <= candidate < len(children):
            sibling = children[candidate]
            if candidate in captions:
                if captions[candidate] == kind:
                    matches.append(candidate)
                break
            # Word can emit a bookmark directly in the body. Older Pandoc
            # output also emitted a plain-text caption equal to the image alt.
            if sibling.tag in {q("w:bookmarkStart"), q("w:bookmarkEnd")}:
                candidate += direction
                continue
            if sibling.tag != q("w:p") or any(
                sibling.find(f".//{q('w:' + name)}") is not None
                for name in ("drawing", "object", "pict", "fldChar", "fldSimple", "sectPr", "br")
            ):
                break
            value = _visible(sibling).strip()
            if value and (kind != "figure" or value not in picture_labels):
                break
            candidate += direction
    if not matches:
        raise ValueError(f"{kind.title()} at body index {index} has missing adjacent SEQ caption")
    return set(matches)


def _pair_captions(children, objects, captions):
    # Consecutive figures can share the same adjacent candidate. Resolve
    # forced pairs from either end before declaring an association ambiguous.
    options = {i: _caption_candidates(children, i, kind, captions) for i, kind in objects.items()}
    adjacent = {i: [] for i in captions}
    for obj, candidates in options.items():
        for caption in candidates:
            adjacent[caption].append(obj)
    pending = [obj for obj, candidates in options.items() if len(candidates) == 1]
    pairs = {}
    while pending:
        obj = pending.pop()
        candidates = options.pop(obj)
        if not candidates:
            raise ValueError(f"No distinct adjacent SEQ caption for object at body index {obj}")
        caption, = candidates
        pairs[obj] = caption
        for other in adjacent[caption]:
            if other in options:
                options[other].remove(caption)
                if len(options[other]) == 1:
                    pending.append(other)
    if options:
        raise ValueError(f"Ambiguous adjacent SEQ captions for objects at body indices {sorted(options)}")
    unused = set(captions) - set(pairs.values())
    if unused:
        raise ValueError(f"Unmatched figure or table captions at body indices {sorted(unused)}")
    return pairs


def _figure_parts(node, relationships, parts):
    used = set()
    unsupported = {q("w:" + name) for name in ("txbxContent", "sdt", "del", "ins", "moveFrom", "moveTo", "object", "pict")}
    for drawing in node.iter(q("w:drawing")):
        if any(ancestor.tag in unsupported for ancestor in drawing.iterancestors()):
            raise ValueError("Drawing in a text box, protected content or revision needs manual placement")
        if drawing.find(ANCHOR) is not None:
            inline_picture_copy(deepcopy(drawing))
            inline = drawing.find(ANCHOR)
        else:
            inline = drawing.find(INLINE)
        if inline is None or len(list(inline.iter(_PIC))) != 1:
            raise ValueError("End placement supports simple embedded pictures; found a complex drawing")
        embeds = {item.get(q("r:embed")) for item in drawing.iter() if item.get(q("r:embed"))}
        if not embeds or any(item.get(q("r:link")) for item in drawing.iter()):
            raise ValueError("End placement requires an embedded picture instead of a linked drawing")
        for rid in embeds:
            relation = relationships.get(rid)
            if relation is None or relation.get("TargetMode") == "External" or not relation.get("Type", "").endswith("/image"):
                raise ValueError(f"Invalid embedded image relationship: {rid}")
            target = relation.get("Target", "")
            part = posixpath.normpath(target.lstrip("/") if target.startswith("/") else f"word/{target}")
            if part not in parts:
                raise ValueError(f"Missing embedded image part: {part}")
            used.add(part)
    return sorted(used)


def _freeze_fields(node):
    fields = _scan_fields(node)
    for field in fields:
        if not (_FREEZEABLE.match(field["code"]) or _owner(field["code"])):
            raise ValueError(f"Unsupported field in end copy: {field['code'][:70]}")
        result = field["result"].strip()
        boundary = field.get("simple", field.get("begin"))
        if (
            not result
            or "PAPERFLOW_PENDING" in result
            or result.startswith("Error!")
            or boundary.get(q("w:dirty")) in {"1", "true", "on"}
        ):
            raise ValueError("Refresh caption and object fields in Word before making end copies")
        if "simple" in field:
            parent = boundary.getparent()
            index = parent.index(boundary)
            for child in list(boundary):
                parent.insert(index, child)
                index += 1
            parent.remove(boundary)
        else:
            for element in (boundary, *field["code_nodes"], field["separator"], field["end"]):
                if element is not None:
                    element.getparent().remove(element)


def _sanitize_clone(clone, drawing_ids, *, allow_mathtype=False):
    unsafe = {
        q("o:OLEObject"),
        *(q("w:" + name) for name in (
            "object", "pict", "sectPr", "sdt", "txbxContent", "ins", "del",
            "moveFrom", "moveTo", "footnoteReference", "endnoteReference",
            "commentRangeStart", "commentRangeEnd", "commentReference",
        )),
    }
    if allow_mathtype:
        unsafe -= {q("w:object"), q("o:OLEObject")}
    if any(child.tag in unsafe for child in clone.iter()):
        raise ValueError("End copy contains an OLE object, review markup, note or another unsupported structure")
    if any(br.get(q("w:type")) in {"page", "column"} for br in clone.iter(q("w:br"))):
        raise ValueError("End copy contains a manual page or column break")
    _freeze_fields(clone)
    for child in list(clone.iter()):
        if child.tag in {
            q("w:bookmarkStart"), q("w:bookmarkEnd"), q("w:lastRenderedPageBreak"), q("w:pageBreakBefore")
        }:
            child.getparent().remove(child)
            continue
        for name in _CLONED_ID_ATTRIBUTES:
            child.attrib.pop(name, None)
        if child.tag in _DRAWING_IDS:
            child.set("id", str(next(drawing_ids)))
    return clone


def _copy_block(node, drawing_ids):
    return _sanitize_clone(deepcopy(node), drawing_ids)


def _copy_figure(node, drawing_ids, *, caption_follows):
    drawings = list(node.iter(q("w:drawing")))
    anchors = sum(drawing.find(ANCHOR) is not None for drawing in drawings)
    if not anchors and not _visible(node).strip() and next(node.iter(q("w:object")), None) is None:
        return _copy_block(node, drawing_ids), anchors
    if anchors and len(drawings) != 1:
        raise ValueError("An anchored end figure needs a single picture, not an ambiguous floating panel layout")
    paragraph = ET.Element(q("w:p"))
    properties = ET.SubElement(paragraph, q("w:pPr"))
    if caption_follows:
        ET.SubElement(properties, q("w:keepNext"))
    ET.SubElement(properties, q("w:jc")).set(q("w:val"), "center")
    for drawing in drawings:
        # Move no prose, fields or source anchors to the end page.
        copy = deepcopy(drawing)
        if copy.find(ANCHOR) is not None:
            inline_picture_copy(copy)
        ET.SubElement(paragraph, q("w:r")).append(copy)
    return _sanitize_clone(paragraph, drawing_ids), anchors


def _copy_same_paragraph_caption(node, drawing_ids):
    caption = deepcopy(node)
    for drawing in list(caption.iter(q("w:drawing"))):
        drawing.getparent().remove(drawing)
    return _sanitize_clone(caption, drawing_ids)


def _start_end_page(paragraph):
    properties = paragraph.find(q("w:pPr"))
    if properties is None:
        properties = ET.Element(q("w:pPr"))
        paragraph.insert(0, properties)
    leading = {q("w:pStyle"), q("w:keepNext"), q("w:keepLines")}
    index = next((i for i, child in enumerate(properties) if child.tag not in leading), len(properties))
    # A separate paragraph containing a page break can spill onto its own
    # blank page when the final body page is already full.
    properties.insert(index, ET.Element(q("w:pageBreakBefore")))


def _highest_drawing_id(document, parts):
    stories = [document]
    for name, data in parts.items():
        if (
            name != _MAIN and name.startswith("word/") and name.endswith(".xml")
            and (b"docPr" in data or b"cNvPr" in data)
        ):
            stories.append(xml(data))
    return max(
        (int(node.get("id", "0")) for story in stories for node in story.iter() if node.tag in _DRAWING_IDS),
        default=0,
    )


def _scope_style_indexes(body, first_end_copy):
    """Keep unscoped, style-selected TOC fields from indexing copied styles."""
    indexes = []
    for field in _scan_fields(body):
        code = field["code"]
        if _TOC.match(code):
            switches = set(_FIELD_SWITCH.findall(code.casefold()))
            if r"\t" in switches and r"\b" not in switches:
                indexes.append(field)
    if not indexes:
        return None

    bookmarks = list(body.iter(q("w:bookmarkStart")))
    names = {bookmark.get(q("w:name"), "").casefold() for bookmark in bookmarks}
    name, suffix = "PaperFlowBodyScope", 0
    while name.casefold() in names:
        suffix += 1
        name = f"PaperFlowBodyScope{suffix}"
    identifier = str(max((int(bookmark.get(q("w:id"))) for bookmark in bookmarks), default=0) + 1)
    start = ET.Element(q("w:bookmarkStart"), {q("w:id"): identifier, q("w:name"): name})
    end = ET.Element(q("w:bookmarkEnd"), {q("w:id"): identifier})
    body.insert(0, start)
    first_end_copy.addprevious(end)
    for field in indexes:
        switch = f" \\b {name} "
        if "simple" in field:
            field["simple"].set(q("w:instr"), field["code"].rstrip() + switch)
        else:
            instruction = field["code_nodes"][-1]
            instruction.text = (instruction.text or "").rstrip() + switch
    return {"bookmark": name, "fields_scoped": len(indexes)}


def append_end_objects(source: str | Path, output: str | Path) -> dict:
    """Publish a new DOCX with each paired figure and table on end pages.

    Refresh the source in Word before calling. Cached values in the new copies
    are snapshots. Original bookmarks and Word fields retain their semantics.
    The report establishes package structure. Visual pagination needs a render.
    """
    source = word._source_path(source)
    output = word._output_path(output, (source,))
    if source.suffix.lower() != ".docx" or output.suffix.lower() != ".docx":
        raise ValueError("End placement requires DOCX source and output files")
    parts = read_package(source)
    document = xml(parts[_MAIN])
    body = document.find(q("w:body"))
    if body is None:
        raise ValueError("DOCX main document has no body")
    children = list(body)
    if any(child.tag == q("w:sectPr") for child in children[:-1]):
        raise ValueError("Unexpected section properties in the middle of the DOCX body")
    end_start = len(children) - int(bool(children and children[-1].tag == q("w:sectPr")))
    insert_at = end_start
    # Each body table owns its cell sub-tables. Copy that whole tree once
    # under its outer caption, keeping the cells, merges and formula data.
    for table in body.iter(q("w:tbl")):
        parent = table.getparent()
        if parent is not body and not (
            parent.tag == q("w:tc") and parent.getparent().tag == q("w:tr")
            and parent.getparent().getparent().tag == q("w:tbl")
        ):
            raise ValueError("End placement requires a body table or a table nested in a regular cell")
    if any(not any(parent.tag == q("w:object") for parent in picture.iterancestors())
           for picture in body.iter(q("w:pict"))):
        raise ValueError("End placement found a legacy VML drawing; convert it to an inline embedded picture first")

    rel_name = "word/_rels/document.xml.rels"
    rel_root = xml(parts[rel_name]) if rel_name in parts else None
    relationships = {rel.get("Id"): rel for rel in rel_root.iter(q("rel:Relationship"))} if rel_root is not None else {}
    original_fields, errors = part_fields(body)
    if errors:
        raise ValueError(f"Word field errors in source document: {errors}")
    original_bookmarks = [item.get(q("w:id")) for item in body.iter(q("w:bookmarkStart"))]
    captions, objects = {}, {}
    for index, child in enumerate(children):
        if child.tag == q("w:p"):
            codes, _ = part_fields(child)
            kinds = [kind for code in codes if (kind := _caption_kind(code))]
            if len(kinds) > 1:
                raise ValueError(f"More than one figure or table caption at body index {index}")
            if kinds:
                captions[index] = kinds[0]
        if child.tag == q("w:tbl"):
            objects[index] = "table"
        elif child.find(f".//{q('w:drawing')}") is not None:
            if child.tag != q("w:p") or child.find(f".//{_PIC}") is None:
                raise ValueError(f"Unsupported or nested drawing at body index {index}")
            objects[index] = "figure"

    highest_id = _highest_drawing_id(document, parts)
    drawing_ids = iter(range(highest_id + 1, 2**32))
    paired_captions = _pair_captions(children, objects, captions)
    copies = []
    part_names = set(parts)
    original_content_types = parts.get("[Content_Types].xml")
    for index, kind in objects.items():
        caption = paired_captions[index]
        image_parts = _figure_parts(children[index], relationships, parts)
        # An end table starts with its caption, allowing the new page to be
        # attached to visible content rather than to a blank break paragraph.
        caption_follows = kind == "figure" and index <= caption
        anchors = 0
        math_copies = []
        if kind == "figure":
            object_copy, anchors = _copy_figure(children[index], drawing_ids, caption_follows=caption_follows)
        else:
            object_copy = deepcopy(children[index])
            if next(object_copy.iter(ANCHOR), None) is not None:
                raise ValueError("A floating drawing inside a table needs manual placement")
            if next(object_copy.iter(q("w:object")), None) is not None:
                if rel_root is None:
                    raise ValueError("Table OLE object lacks a document relationship")
                from .endcopy_math import clone_mathtype_objects

                math_copies = clone_mathtype_objects(object_copy, document, parts, rel_root)
            object_copy = _sanitize_clone(object_copy, drawing_ids, allow_mathtype=bool(math_copies))
        caption_copy = (_copy_same_paragraph_caption(children[caption], drawing_ids)
                        if caption == index else _copy_block(children[caption], drawing_ids))
        clones = (object_copy, caption_copy) if caption_follows else (caption_copy, object_copy)
        _start_end_page(clones[0])
        for offset, clone in enumerate(clones):
            body.insert(insert_at + offset, clone)
        copies.append({
            "kind": kind,
            "caption": _visible(children[caption]),
            "source_body_index": index,
            "end_body_index": insert_at + int(not caption_follows),
            "embedded_image_parts": image_parts,
            "floating_pictures_converted_in_end_copy": anchors,
            "mathtype_end_copies": math_copies,
        })
        insert_at += 2
    if part_fields(body) != (original_fields, []) or [
        item.get(q("w:id")) for item in body.iter(q("w:bookmarkStart"))
    ] != original_bookmarks:
        raise ValueError("End placement would change the original Word field or bookmark inventory")

    scope = _scope_style_indexes(body, body[end_start]) if copies else None
    if scope:
        for item in copies:
            item["end_body_index"] += 2  # Both new bookmark boundaries precede the end copies.

    changed_parts = []
    if copies:
        parts[_MAIN] = ET.tostring(document, encoding="UTF-8", xml_declaration=True, standalone=True)
        changed_parts.append(_MAIN)
    if any(item["mathtype_end_copies"] for item in copies):
        parts[rel_name] = ET.tostring(rel_root, encoding="UTF-8", xml_declaration=True, standalone=True)
        changed_parts += [rel_name, *sorted(set(parts) - part_names)]
        if parts.get("[Content_Types].xml") != original_content_types and "[Content_Types].xml" not in changed_parts:
            changed_parts.append("[Content_Types].xml")
    stage = word._stage_path(output)
    try:
        write_package(parts, stage)
        word._publish(stage, output)
    finally:
        stage.unlink(missing_ok=True)
    return {
        "output": str(output),
        "figure_copies": sum(item["kind"] == "figure" for item in copies),
        "table_copies": sum(item["kind"] == "table" for item in copies),
        "floating_end_copies": sum(item["floating_pictures_converted_in_end_copy"] for item in copies),
        "mathtype_end_copies": sum(len(item["mathtype_end_copies"]) for item in copies),
        "copies": copies,
        "placement_verified": True,
        "pagination_verified": False,
        "preserved_original_fields_and_bookmarks": True,
        "scoped_style_indexes": scope,
        "changed_parts": changed_parts,
    }
