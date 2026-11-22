"""Convert a cloned, fixed-size floating Word picture to an inline drawing."""

from __future__ import annotations

from lxml import etree as ET

from .package import q


_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
_WP14 = "{http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_PICTURE_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
_PIC = f"{{{_PICTURE_NS}}}"
_PIC14 = "{http://schemas.microsoft.com/office/drawing/2010/picture}"
_MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"

ANCHOR = _WP + "anchor"
INLINE = _WP + "inline"

# Open XML SDK's CT_Inline and CT_Anchor models describe which children carry
# the drawing, and which belong only to floating page placement.
_INLINE_ORDER = (_WP + "extent", _WP + "effectExtent", _WP + "docPr", _WP + "cNvGraphicFramePr", _A + "graphic")
_INLINE_REQUIRED = {_WP + "extent", _WP + "docPr", _A + "graphic"}
_POSITIONS = {_WP + name for name in ("simplePos", "positionH", "positionV")}
_WRAPS = {_WP + name for name in ("wrapNone", "wrapSquare", "wrapTight", "wrapThrough", "wrapTopAndBottom")}
_RELATIVE_SIZES = {_WP14 + "sizeRelH": _WP14 + "pctWidth", _WP14 + "sizeRelV": _WP14 + "pctHeight"}
_COMPATIBILITY = {_MC + name for name in ("Ignorable", "MustUnderstand", "ProcessContent", "PreserveAttributes", "PreserveElements")}
_ANCHOR_ATTRIBUTES = {
    "distT", "distB", "distL", "distR", "simplePos", "relativeHeight", "behindDoc", "locked", "layoutInCell", "allowOverlap",
    _WP14 + "anchorId", _WP14 + "editId",
}


def _validate_picture(anchor: ET._Element) -> None:
    graphic = anchor.find(_A + "graphic")
    if len(graphic) != 1 or graphic[0].tag != _A + "graphicData":
        raise ValueError("Floating drawing does not contain one picture graphic")
    data = graphic[0]
    if data.get("uri") != _PICTURE_NS or len(data) != 1 or data[0].tag != _PIC + "pic":
        raise ValueError("Floating drawing is not a simple picture; charts and grouped objects need manual placement")
    picture = data[0]
    children = [node.tag for node in picture]
    base = [_PIC + "nvPicPr", _PIC + "blipFill", _PIC + "spPr"]
    extensions = [_PIC14 + "style", _PIC14 + "extLst"]
    if children[:3] != base or children[3:] != [name for name in extensions if name in children[3:]]:
        raise ValueError("Floating picture contains unsupported picture elements")
    blips = list(picture.iter(_A + "blip"))
    if len(blips) != 1 or blips[0].getparent() is not picture[1] or not blips[0].get(q("r:embed")):
        raise ValueError("Floating picture needs one embedded image")
    if any(q("r:link") in node.attrib for node in anchor.iter()):
        raise ValueError("Floating picture uses a linked image")
    if anchor.find(_WP + "docPr").get("hidden") in {"1", "true", "on"}:
        raise ValueError("Floating picture is hidden")
    extent = anchor.find(_WP + "extent")
    try:
        valid = all(int(extent.get(axis, "0")) > 0 for axis in ("cx", "cy"))
    except ValueError:
        valid = False
    if not valid:
        raise ValueError("Floating picture needs positive fixed dimensions")


def inline_picture_copy(drawing: ET._Element) -> ET._Element:
    """Inline the passed w:drawing clone, preserving the picture and its extents.

    Callers keep the original anchor and assign unique drawing IDs on the
    clone. Relationship targets and final Word pagination are checked by the
    placement stage.
    """
    if drawing.tag != q("w:drawing") or len(drawing) != 1 or drawing[0].tag != ANCHOR:
        raise ValueError("Expected a Word drawing with one floating anchor")
    anchor = drawing[0]
    tags = [node.tag for node in anchor]
    allowed = set(_INLINE_ORDER) | _POSITIONS | _WRAPS | _RELATIVE_SIZES.keys()
    if set(tags) - allowed or set(anchor.attrib) - (_ANCHOR_ATTRIBUTES | _COMPATIBILITY):
        raise ValueError("Floating picture contains unsupported anchor markup")
    if len(tags) != len(set(tags)) or len(set(tags) & _WRAPS) > 1:
        raise ValueError("Floating picture has duplicate anchor elements")
    retained = [tag for tag in tags if tag in _INLINE_ORDER]
    if not _INLINE_REQUIRED.issubset(retained) or retained != [tag for tag in _INLINE_ORDER if tag in retained]:
        raise ValueError("Floating picture has incomplete or misordered inline-compatible elements")
    for node in anchor:
        percent_tag = _RELATIVE_SIZES.get(node.tag)
        if percent_tag is not None and (
            set(node.attrib) - {"relativeFrom"} or len(node) != 1 or node[0].tag != percent_tag
            or len(node[0]) or node[0].attrib or (node[0].text or "").strip() != "0"
        ):
            raise ValueError("Floating picture uses relative sizing and needs manual placement")
    _validate_picture(anchor)

    anchor.tag = INLINE
    for name in _ANCHOR_ATTRIBUTES:
        anchor.attrib.pop(name, None)
    for node in list(anchor):
        if node.tag not in _INLINE_ORDER:
            anchor.remove(node)
    return drawing
