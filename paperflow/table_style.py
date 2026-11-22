"""Apply a template-owned Word table style without rebuilding its cells."""
from collections import Counter

from lxml import etree as ET

from .package import NS, q, read_package, xml


_LOOK_OPTIONS = {
    "firstRow": 0x0020, "lastRow": 0x0040, "firstColumn": 0x0080,
    "lastColumn": 0x0100, "noHBand": 0x0200, "noVBand": 0x0400,
}
_ON = {"1", "true", "on"}
_OMITTED_LOOK = 0x04A0  # The default used by Microsoft Word.
_SKIP_STORIES = {q("w:" + name) for name in ("txbxContent", "del", "moveFrom")}
_ADJACENT_ANNOTATIONS = {q("w:" + name) for name in (
    "bookmarkStart", "bookmarkEnd", "proofErr", "commentRangeStart", "commentRangeEnd", "permStart", "permEnd",
)}
_TABLE_FORMAT = ("tblBorders", "shd", "tblStyleRowBandSize", "tblStyleColBandSize")
_OWNED_PROPERTIES = {
    q("w:tblPr"): ("tbl", _TABLE_FORMAT),
    q("w:tblPrEx"): ("tr", _TABLE_FORMAT),
    q("w:tcPr"): ("tc", ("tcBorders", "shd", "cnfStyle")),
    q("w:trPr"): ("tr", ("cnfStyle",)),
    q("w:pPr"): ("p", ("cnfStyle",)),
}


def body_tables(document):
    body = document.find(q("w:body"))
    if body is None:
        return []
    return [table for table in body.iter(q("w:tbl"))
            if not any(parent.tag in _SKIP_STORIES for parent in table.iterancestors())]


def table_look_mask(node):
    if node is None:
        return _OMITTED_LOOK
    if any(q("w:" + name) in node.attrib for name in _LOOK_OPTIONS):
        return sum(bit for name, bit in _LOOK_OPTIONS.items()
                   if node.get(q("w:" + name), "").casefold() in _ON)
    return int(node.get(q("w:val"), "0"), 16) & sum(_LOOK_OPTIONS.values())


def _set_look(properties, mask):
    look = properties.find(q("w:tblLook"))
    if look is None:
        look = ET.Element(q("w:tblLook"))
        after = {q("w:" + name) for name in ("tblCaption", "tblDescription", "tblPrChange")}
        index = next((i for i, node in enumerate(properties) if node.tag in after), len(properties))
        properties.insert(index, look)
    for name, bit in _LOOK_OPTIONS.items():
        look.set(q("w:" + name), str(int(bool(mask & bit))))
    look.set(q("w:val"), f"{mask:04X}")


def _style_label(style, fallback):
    name = style.find(q("w:name"))
    return name.get(q("w:val"), fallback) if name is not None else fallback


def _table_style_id(table, default=None):
    style = table.find("w:tblPr/w:tblStyle", NS)
    return style.get(q("w:val")) if style is not None else default


def _default_table_style(styles):
    return next((sid for sid, node in styles.items()
                 if node.get(q("w:type")) == "table" and node.get(q("w:default"), "").casefold() in _ON), None)


def _reject_adjacent_tables(tables):
    selected = set(tables)
    for table in tables:
        previous = table.getprevious()
        while previous is not None and previous.tag in _ADJACENT_ANNOTATIONS:
            previous = previous.getprevious()
        if previous in selected and not any(
            member.find("w:tblPr/w:tblpPr", NS) is not None for member in (table, previous)
        ):
            raise ValueError("Add a paragraph between adjacent tables before applying the same template style")


def _remove_overrides(table, use_template_look):
    counts = Counter()
    for properties in table.iter(*_OWNED_PROPERTIES):
        owner, remove = _OWNED_PROPERTIES[properties.tag]
        if properties.getparent().tag != q("w:" + owner):
            continue
        if next(properties.iterancestors(q("w:tbl")), None) is not table:
            continue
        if any(parent.tag in _SKIP_STORIES for parent in properties.iterancestors()):
            continue
        if use_template_look and properties.tag == q("w:tblPrEx"):
            remove += ("tblLook",)
        for name in remove:
            for node in properties.findall(q("w:" + name)):
                properties.remove(node)
                counts[name] += 1
    return counts


def apply_template_table_style(document, template, target_styles, style_id):
    selected = target_styles.get(style_id)
    if selected is None or selected.get(q("w:type")) != "table":
        raise ValueError(f"Template Word table style ID not found: {style_id}")
    tables = body_tables(document)
    _reject_adjacent_tables(tables)
    default_style = _default_table_style(target_styles)
    exemplars = [table for table in body_tables(template) if _table_style_id(table, default_style) == style_id]
    exemplar_looks = {table_look_mask(table.find("w:tblPr/w:tblLook", NS)) for table in exemplars}
    uniform = len(exemplar_looks) == 1 and all(
        table.find("w:tr/w:tblPrEx/w:tblLook", NS) is None for table in exemplars
    )
    mask = next(iter(exemplar_looks)) if uniform else None
    look_policy = "template_example" if uniform else "source_options"
    unresolved = []
    if tables and exemplars and not uniform:
        unresolved.append("Template tables have different table style options. The source table options were retained.")
    removed, looks = Counter(), Counter()
    for table in tables:
        properties = table.find(q("w:tblPr"))
        if properties is None:
            properties = ET.Element(q("w:tblPr"))
            table.insert(0, properties)
        style = properties.find(q("w:tblStyle"))
        if style is None:
            style = ET.Element(q("w:tblStyle"))
            properties.insert(0, style)
        style.set(q("w:val"), style_id)
        if mask is not None:
            _set_look(properties, mask)
        removed.update(_remove_overrides(table, mask is not None))
        looks[f"{table_look_mask(properties.find(q('w:tblLook'))):04X}"] += 1
    return {"style_id": style_id, "style_name": _style_label(selected, style_id), "applied_tables": len(tables),
            "look_policy": look_policy, "template_exemplars": len(exemplars), "look_counts": dict(looks),
            "removed_overrides": dict(removed), "unresolved": unresolved}


def _style_looks(path, name):
    parts = read_package(path)
    styles = {node.get(q("w:styleId")): node for node in xml(parts["word/styles.xml"]).findall(q("w:style"))}
    matches = {sid for sid, node in styles.items() if node.get(q("w:type")) == "table"
               and _style_label(node, sid).casefold() == name.casefold()}
    default_style = _default_table_style(styles)
    tables = body_tables(xml(parts["word/document.xml"]))
    return Counter(f"{table_look_mask(table.find('w:tblPr/w:tblLook', NS)):04X}"
                   for table in tables if _table_style_id(table, default_style) in matches)


def verify_template_table_style(template, output, applications):
    """Check the selected style and table-level options after Word saves."""
    name = applications[0]["style_name"]
    expected = _style_looks(template, name)
    for application in applications:
        expected.update(application["look_counts"])
    actual = _style_looks(output, name)
    if actual != expected:
        raise RuntimeError(f"Word changed the selected table style or merged styled tables: "
                           f"expected {dict(expected)}, observed {dict(actual)} for {name!r}")
    return {"style_name": name, "template_and_imported_tables": sum(actual.values()),
            "table_look_counts": dict(actual), "word_style_assignments_verified": True}
