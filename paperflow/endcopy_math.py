"""Clone the native MathType objects in an end-page table copy."""

from __future__ import annotations

from dataclasses import dataclass
import posixpath
import re

from lxml import etree as ET

from .mathtype import MATHTYPE_PROGID, inspect_native_embedding
from .package import digest, q, xml


_MAIN = "word/document.xml"
_CONTENT_TYPES = "[Content_Types].xml"
_CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
_V = "{urn:schemas-microsoft-com:vml}"
_ANCHOR_ID = "{http://schemas.microsoft.com/office/word/2010/wordml}anchorId"
_DRAWING_ANCHOR_ID = "{http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing}anchorId"
_SPID = "{urn:schemas-microsoft-com:office:office}spid"
_SHAPE = re.compile(r"_x0000_[is]([0-9]+)$")
_INLINE = re.compile(r"_x0000_i[0-9]+$")
_OBJECT_ID = re.compile(r"_([0-9]+)$")
_DIMENSION = re.compile(r"(?:[0-9]+(?:\.[0-9]+)?|\.[0-9]+)pt$")
_RELATION = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
_OLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.oleObject"


@dataclass(frozen=True)
class _MathObject:
    container: ET._Element
    shape: ET._Element
    preview: ET._Element
    ole: ET._Element
    embedding_part: str
    preview_part: str
    mtef_sha256: str


def _validate_container(container):
    tags = [child.tag for child in container]
    shape_type = _V + "shapetype"
    base = [_V + "shape", q("o:OLEObject")]
    if tags not in (base, [shape_type, *base]):
        raise ValueError("End table supports only a simple inline MathType object")
    shape, ole = container[-2], container[-1]
    if (
        container.getparent().tag != q("w:r")
        or not any(parent.tag == q("w:tc") for parent in container.iterancestors())
        or ole.get("ProgID") != MATHTYPE_PROGID
        or ole.get("Type") != "Embed"
        or ole.get("DrawAspect") != "Content"
        or len(ole)
        or not _INLINE.fullmatch(shape.get("id", ""))
        or not _OBJECT_ID.fullmatch(ole.get("ObjectID", ""))
        or ole.get("ShapeID") != shape.get("id")
        or shape.get("type") != "#_x0000_t75"
        or shape.get(_SPID) is not None
        or any(ET.QName(name).localname in {"gfxdata", "href", "src"} for name in shape.attrib)
        or [child.tag for child in shape] != [_V + "imagedata"]
    ):
        raise ValueError("End table supports only embedded Equation.DSMT4 with a simple inline preview")
    if len(container) == 3 and container[0].get("id") != "_x0000_t75":
        raise ValueError("Unexpected VML shape type in the MathType object")
    entries = [entry.strip() for entry in shape.get("style", "").split(";") if entry.strip()]
    properties = dict(tuple(value.strip() for value in entry.split(":", 1)) for entry in entries if ":" in entry)
    if len(entries) != 2 or set(properties) != {"width", "height"} or not all(
        _DIMENSION.fullmatch(value) and float(value[:-2]) > 0 for value in properties.values()
    ):
        raise ValueError("End table cannot copy a positioned or nonstandard MathType shape")
    preview = shape[0]
    if any(ET.QName(name).localname in {"src", "href", "relid", "link"} for name in preview.attrib):
        raise ValueError("End table cannot copy a linked MathType preview")
    return shape, preview, ole


def _internal_part(node, relationships, parts, kind):
    relation = relationships.get(node.get(q("r:id")))
    if relation is None or relation.get("Type") != _RELATION + kind or relation.get("TargetMode", "Internal") != "Internal":
        raise ValueError(f"Invalid MathType {kind} relationship in the end table")
    target = relation.get("Target", "")
    part = posixpath.normpath(target.lstrip("/") if target.startswith("/") else f"word/{target}")
    if not target or part not in parts:
        raise ValueError(f"Missing MathType {kind} data in the end table")
    return part


def _content_type(part, content_types):
    override = [node.get("ContentType") for node in content_types if node.get("PartName") == "/" + part]
    if override:
        return override[0]
    extension = posixpath.splitext(part)[1][1:].lower()
    return next((node.get("ContentType") for node in content_types
                 if node.get("Extension", "").lower() == extension), None)


def _reserve_name(used, prefix, suffix=""):
    i = 1
    while f"{prefix}{i}{suffix}".lower() in used:
        i += 1
    result = f"{prefix}{i}{suffix}"
    used.add(result.lower())
    return result


def _used_identifiers(main_document, parts):
    used_shapes, used_objects, used_anchors = set(), set(), set()
    stories = [main_document]
    for name, payload in parts.items():
        if name != _MAIN and name.startswith("word/") and name.endswith(".xml") and (
            b"_x0000_" in payload or b"ObjectID" in payload or b"anchorId" in payload
        ):
            stories.append(xml(payload))
    for story in stories:
        for node in story.iter():
            for name in (_SPID, "ShapeID", "id"):
                shape_id = node.get(name)
                if shape_id and (match := _SHAPE.fullmatch(shape_id)):
                    used_shapes.add(int(match[1]))
            if node.tag == q("o:OLEObject") and (match := _OBJECT_ID.fullmatch(node.get("ObjectID", ""))):
                used_objects.add(int(match[1]))
            for name in (_ANCHOR_ID, _DRAWING_ANCHOR_ID):
                if (anchor := node.get(name)) is not None:
                    used_anchors.add(anchor.upper())
    return used_shapes, used_objects, used_anchors


def clone_mathtype_objects(table_clone, main_document, parts, relationship_root):
    """Give cloned native equations independent embeddings and previews.

    This changes only the cloned table and adds new package entries. The
    caller owns serialization of relationship_root after a successful clone.
    """
    containers = list(table_clone.iter(q("w:object")))
    if not containers and not list(table_clone.iter(q("o:OLEObject"))):
        return []
    if table_clone.tag != q("w:tbl") or relationship_root is None or _CONTENT_TYPES not in parts:
        raise ValueError("End copy needs a complete Word table with local MathType data")

    content_types = xml(parts[_CONTENT_TYPES])
    relationships = {node.get("Id"): node for node in relationship_root}
    if None in relationships or len(relationships) != len(relationship_root):
        raise ValueError("Ambiguous MathType relationship identifiers in the Word package")
    if not any(node.get("id") == "_x0000_t75" for node in main_document.iter(_V + "shapetype")):
        raise ValueError("The Word document has no VML shape type for its MathType previews")
    objects = []
    for container in containers:
        shape, preview, ole = _validate_container(container)
        embedding = _internal_part(ole, relationships, parts, "oleObject")
        preview_part = _internal_part(preview, relationships, parts, "image")
        if (
            not embedding.startswith("word/embeddings/")
            or not preview_part.startswith("word/media/")
            or _content_type(embedding, content_types) != _OLE_CONTENT_TYPE
            or not (_content_type(preview_part, content_types) or "").startswith("image/")
        ):
            raise ValueError("Unsupported MathType part name or content type in the end table")
        audit = inspect_native_embedding(parts[embedding])
        if not audit["valid_native_storage"]:
            raise ValueError("End table contains an invalid native MathType embedding")
        objects.append(_MathObject(container, shape, preview, ole, embedding, preview_part, audit["mtef_sha256"]))
    if len(objects) != sum(1 for _ in table_clone.iter(q("o:OLEObject"))):
        raise ValueError("End table contains an OLE object outside a supported inline container")

    shape_ids, object_ids, anchor_ids = _used_identifiers(main_document, parts)
    # Allocate in Word's existing first inline VML block so the clone does
    # not require editing the document's VML ID map.
    free_shapes = [i for i in range(1025, 2049) if i not in shape_ids]
    if len(free_shapes) < len(objects):
        raise ValueError("Insufficient unused inline VML shape IDs for end-page MathType copies")
    next_object = max(object_ids, default=0) + 1
    if next_object + len(objects) > 2**31:
        next_object = 1
    next_anchor = 1
    used_parts = {name.lower() for name in parts}
    used_parts.update(node.get("PartName").lstrip("/").lower()
                      for node in content_types if node.get("PartName"))
    used_rels = {key.lower() for key in relationships}
    added_overrides = False

    def clone_part(old, prefix, kind):
        nonlocal added_overrides
        new = _reserve_name(used_parts, prefix, posixpath.splitext(old)[1])
        identifier = _reserve_name(used_rels, "rId")
        content_type = _content_type(old, content_types)
        if _content_type(new, content_types) != content_type:
            ET.SubElement(content_types, _CT + "Override", PartName="/" + new, ContentType=content_type)
            added_overrides = True
        parts[new] = parts[old]
        ET.SubElement(relationship_root, q("rel:Relationship"), Id=identifier,
                      Type=_RELATION + kind, Target=new.removeprefix("word/"))
        return new, identifier

    reports = []
    for shape_number, obj in zip(free_shapes, objects):
        shape_id = f"_x0000_i{shape_number}"
        while next_object in object_ids:
            next_object += 1
        object_ids.add(next_object)
        old_shape_id = obj.ole.get("ShapeID")
        object_id = f"_{next_object}"
        if obj.container.get(_ANCHOR_ID) is not None:
            while f"{next_anchor:08X}" in anchor_ids:
                next_anchor += 1
            anchor_id = f"{next_anchor:08X}"
            anchor_ids.add(anchor_id)
            obj.container.set(_ANCHOR_ID, anchor_id)
        if obj.container[0].tag == _V + "shapetype":
            obj.container.remove(obj.container[0])
        embedding, ole_rid = clone_part(obj.embedding_part, "word/embeddings/oleObject", "oleObject")
        preview_part, preview_rid = clone_part(obj.preview_part, "word/media/image", "image")
        obj.shape.set("id", shape_id)
        obj.ole.set("ShapeID", shape_id)
        obj.ole.set("ObjectID", object_id)
        obj.ole.set(q("r:id"), ole_rid)
        obj.preview.set(q("r:id"), preview_rid)
        reports.append({"original_shape_id": old_shape_id, "shape_id": shape_id, "object_id": object_id,
                        "original_embedding_part": obj.embedding_part, "embedding_part": embedding,
                        "original_preview_part": obj.preview_part, "preview_part": preview_part,
                        "mtef_sha256": obj.mtef_sha256, "preview_sha256": digest(parts[preview_part])})
    if added_overrides:
        parts[_CONTENT_TYPES] = ET.tostring(content_types, encoding="UTF-8", xml_declaration=True, standalone=True)
    return reports
