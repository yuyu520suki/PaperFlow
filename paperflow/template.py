"""Inspect explicit insertion slots and prepare a copy of a Word template."""
from pathlib import Path
import posixpath
import re

from lxml import etree as ET

from .package import NS, digest, part_fields, q, read_package, text_run, write_package, xml


DEFAULT_SLOT = "{{PAPERFLOW_BODY}}"
_MARKER = re.compile(r"\{\{PAPERFLOW_[A-Z_]+\}\}")
_STORY_EXCLUSIONS = {q("w:txbxContent"), q("w:del"), q("w:moveFrom")}
_TEXT_NODES = {q("w:t"), q("w:tab"), q("w:br"), q("w:cr")}


def _main_story_events(root):
    walker = ET.iterwalk(root, events=("start", "end"))
    for event, element in walker:
        if element.tag in _STORY_EXCLUSIONS:
            if event == "start":
                walker.skip_subtree()
            continue
        yield event, element


def _paragraph_text(paragraph):
    chunks = []
    for event, element in _main_story_events(paragraph):
        if event != "start" or element.tag not in _TEXT_NODES:
            continue
        if element.tag == q("w:t"):
            chunks.append(element.text or "")
        else:
            chunks.append("\t" if element.tag == q("w:tab") else "\n")
    return "".join(chunks)


def _index_body(document):
    body = document.find(q("w:body"))
    if body is None:
        raise ValueError("Template is missing its Word body")
    terminal_sections = body.findall(q("w:sectPr"))
    if len(terminal_sections) > 1 or (terminal_sections and terminal_sections[0] is not body[-1]):
        raise ValueError("Template has a misplaced or repeated final section definition")

    sections, paragraphs, starts, ends = [], [], [], []
    for event, element in _main_story_events(body):
        section = len(sections) + 1
        if element.tag == q("w:p"):
            if event == "start":
                paragraphs.append((element, section, _paragraph_text(element)))
            else:
                properties = element.find("w:pPr/w:sectPr", NS)
                if properties is not None:
                    if any(parent.tag == q("w:tbl") for parent in element.iterancestors()):
                        raise ValueError("Template has a section break inside a table")
                    sections.append(properties)
        elif event == "start" and element.tag == q("w:bookmarkStart"):
            starts.append((element, section))
        elif event == "start" and element.tag == q("w:bookmarkEnd"):
            ends.append((element, section))
    sections.append(terminal_sections[0] if terminal_sections else None)
    return body, sections, paragraphs, starts, ends


def _body_level_boundary(node, body):
    parent = node.getparent()
    return parent is body or (parent.tag == q("w:p") and parent.getparent() is body)


def _resolve_slot(index, slot):
    body, sections, paragraphs, starts, ends = index
    if slot is None:
        if any(DEFAULT_SLOT in text for _, _, text in paragraphs):
            slot = DEFAULT_SLOT
        elif len(sections) == 1:
            return 1, {"slot": None, "slot_kind": "sole_section"}
        else:
            raise ValueError("Multi-section templates require an explicit section/slot mapping")
    if not isinstance(slot, str) or not slot.strip() or len(slot) > 255 or any(char in slot for char in "\r\n\t\x07"):
        raise ValueError("Provide a nonempty single-line template marker or bookmark:NAME")

    if slot.startswith("bookmark:"):
        name = slot.removeprefix("bookmark:")
        if not name:
            raise ValueError("Template bookmark name is required")
        matches = [(node, section) for node, section in starts
                   if node.get(q("w:name"), "").casefold() == name.casefold()]
        if len(matches) != 1:
            raise ValueError(f"Template bookmark must occur exactly once: {name}")
        start, section = matches[0]
        identifier = start.get(q("w:id"))
        matches = [node for node, _ in starts if node.get(q("w:id")) == identifier]
        end_matches = [(node, number) for node, number in ends if node.get(q("w:id")) == identifier]
        if identifier is None or len(matches) != 1 or len(end_matches) != 1:
            raise ValueError(f"Template bookmark must have one matching start and end: {name}")
        end, end_section = end_matches[0]
        if section != end_section:
            raise ValueError(f"Template bookmark crosses a section boundary: {name}")
        if not all(_body_level_boundary(node, body) for node in (start, end)):
            raise ValueError(f"Template bookmark endpoints must be in the main body outside containers: {name}")
        endpoints = [node for node in body.iter() if node is start or node is end]
        if endpoints != [start, end]:
            raise ValueError(f"Template bookmark end precedes its start: {name}")
        first = start if start.getparent() is body else start.getparent()
        last = end if end.getparent() is body else end.getparent()
        return section, {"slot": slot, "slot_kind": "bookmark", "bookmark_name": start.get(q("w:name")),
                         "body_child_index": body.index(first) + 1, "body_end_index": body.index(last) + 1}

    matches = [(paragraph, section, text) for paragraph, section, text in paragraphs if slot in text]
    occurrences = sum(text.count(slot) for _, _, text in matches)
    if occurrences != 1:
        raise ValueError(f"Template marker must occur exactly once; found {occurrences}: {slot}")
    paragraph, section, text = matches[0]
    if paragraph.getparent() is not body or text.strip(" \t\r\n\x07") != slot:
        raise ValueError("Template body marker must occupy its own top-level paragraph")
    if paragraph.xpath(".//w:fldChar | .//w:fldSimple | .//w:instrText | .//w:drawing | .//w:object | .//w:pict | .//m:oMath", namespaces=NS):
        raise ValueError("Template body marker paragraph contains a field, drawing or object")
    return section, {"slot": slot, "slot_kind": "marker", "body_child_index": body.index(paragraph) + 1}


def _attributes(element):
    return {ET.QName(name).localname: value for name, value in element.attrib.items()} if element is not None else None


def _section_report(section, number):
    report = {"section_index": number, "explicit_properties": section is not None}
    for name in ("pgSz", "pgMar", "cols", "type", "pgNumType", "titlePg", "vAlign", "docGrid"):
        report[name] = _attributes(section.find(q("w:" + name))) if section is not None else None
    for kind in ("header", "footer"):
        report[kind + "_references"] = [
            {"type": ref.get(q("w:type")), "relationship_id": ref.get(q("r:id"))}
            for ref in section.findall(q("w:" + kind + "Reference"))
        ] if section is not None else []
    return report


def select_template_section(document, slot=None):
    """Return the slot's section XML and its one-based location evidence."""
    index = _index_body(document)
    number, location = _resolve_slot(index, slot)
    sections = index[1]
    section = sections[number - 1]
    return section, {**location, **_section_report(section, number), "section_count": len(sections)}


def _read_template(template):
    path = Path(template).resolve(strict=True)
    sha256 = digest(path.read_bytes())
    parts = read_package(path)
    if digest(path.read_bytes()) != sha256:
        raise ValueError("Template changed while it was being inspected")
    document = xml(parts["word/document.xml"])
    return path, sha256, parts, document


def inventory_template(template):
    """List authored slots and immutable paragraph locators for a template copy."""
    path, sha256, _, document = _read_template(template)
    index = _index_body(document)
    body, sections, paragraphs, starts, _ = index
    markers = {marker for _, _, text in paragraphs for marker in _MARKER.findall(text)}
    bookmarks = {"bookmark:" + name for node, _ in starts
                 if (name := node.get(q("w:name"))) and not name.startswith("_")}
    slots = []
    for slot in sorted(markers | bookmarks):
        try:
            section, location = _resolve_slot(index, slot)
            slots.append({**location, "valid": True, "section_index": section})
        except ValueError as error:
            slots.append({"slot": slot, "valid": False, "reason": str(error)})

    anchors = [{"body_child_index": body.index(paragraph) + 1, "section_index": section,
                "insertion_section_index": section + int(paragraph.find("w:pPr/w:sectPr", NS) is not None),
                "text_preview": text[:100], "style_id": style.get(q("w:val")) if style is not None else None}
               for paragraph, section, text in paragraphs if paragraph.getparent() is body
               for style in [paragraph.find("w:pPr/w:pStyle", NS)]]
    proposal = None if slots else {"slot": DEFAULT_SLOT, "after_paragraph": None,
                                  "section_index": len(sections), "operation": "append before final section properties"}
    return {"path": str(path), "sha256": sha256,
            "sections": [_section_report(section, number) for number, section in enumerate(sections, start=1)],
            "slots": slots, "paragraph_anchors": anchors, "insertion_proposal": proposal}


def _part_relationships(parts, part):
    folder, name = posixpath.split(part)
    path = posixpath.join(folder, "_rels", name + ".rels")
    return {item.get("Id"): item for item in xml(parts[path])} if path in parts else {}


def _internal_target(part, target):
    if not isinstance(target, str) or not target:
        raise ValueError(f"Missing relationship target in {part}")
    return target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def _story_signature(parts, part):
    root = xml(parts[part])
    fields, errors = part_fields(root)
    depth = 0
    text = []
    for event, node in ET.iterwalk(root, events=("start", "end")):
        if node.tag == q("w:fldSimple"):
            depth += 1 if event == "start" else -1
        elif node.tag == q("w:fldChar") and event == "start":
            operation = node.get(q("w:fldCharType"))
            depth += (operation == "begin") - (operation == "end")
        elif not depth and event == "start" and node.tag == q("w:t"):
            text.append(node.text or "")
    ids = {node.get(attr) for node in root.iter() for attr in (q("r:id"), q("r:embed"), q("r:link"))} - {None}
    assets = []
    for key, relationship in _part_relationships(parts, part).items():
        if key not in ids:
            continue
        target = relationship.get("Target")
        kind = relationship.get("Type", "").rsplit("/", 1)[-1]
        if relationship.get("TargetMode") == "External":
            assets.append((kind, target))
        else:
            resolved = _internal_target(part, target)
            assets.append((kind, digest(parts[resolved]) if resolved in parts else "missing:" + resolved))
    return ("".join(text).strip(), tuple(" ".join(f.split()) for f in fields), tuple(sorted(assets)), tuple(errors))


def section_story_signature(path):
    """Compare section stories and page numbering independently of field caches."""
    parts = read_package(path)
    main = "word/document.xml"
    _, sections, *_ = _index_body(xml(parts[main]))
    relationships = _part_relationships(parts, main)
    result = []
    effective = {}
    for index, section in enumerate(sections):
        for story in ("header", "footer"):
            references = {} if section is None else {e.get(q("w:type")): e for e in section.findall(q("w:" + story + "Reference"))}
            for kind in ("default", "first", "even"):
                reference = references.get(kind)
                key = (story, kind)
                if reference is None:
                    previous = effective.get(key)
                    if previous:
                        effective[key] = (False, previous[1])
                    continue
                relationship = relationships.get(reference.get(q("r:id")))
                if relationship is None:
                    raise ValueError(f"Missing {story} relationship in template section")
                signature = _story_signature(parts, _internal_target(main, relationship.get("Target")))
                if any(signature) or key in effective:
                    effective[key] = (True, signature)
        page_number = section.find(q("w:pgNumType")) if section is not None else None
        numbering = _attributes(page_number) or {}
        numbering.setdefault("fmt", "decimal")
        if index == 0:
            numbering.setdefault("start", "1")
        first_page = section.find(q("w:titlePg")) if section is not None else None
        result.append((dict(effective), numbering, _enabled(first_page)))
    settings = xml(parts["word/settings.xml"]) if "word/settings.xml" in parts else None
    even_and_odd = settings.find(q("w:evenAndOddHeaders")) if settings is not None else None
    return (_enabled(even_and_odd), result)


def _enabled(element):
    return element is not None and element.get(q("w:val"), "true") in {"1", "true", "on"}


def prepare_template(template, output, *, after_paragraph, expected_sha256, slot=DEFAULT_SLOT):
    """Insert a slot after a verified body-child locator, or at the end for None."""
    path, sha256, parts, document = _read_template(template)
    if sha256 != expected_sha256:
        raise ValueError("Template SHA-256 differs from the inspected copy")
    output = Path(output).resolve()
    if path.suffix.lower() not in {".docx", ".dotx"} or output.suffix.lower() != path.suffix.lower():
        raise ValueError("Template copy must use the source .docx or .dotx extension")
    if not isinstance(slot, str) or not _MARKER.fullmatch(slot):
        raise ValueError("Prepared markers must use the form {{PAPERFLOW_NAME}}")
    body, _, paragraphs, _, _ = _index_body(document)
    if any(slot in text for _, _, text in paragraphs):
        raise ValueError(f"Template already contains this marker: {slot}")
    if after_paragraph is None:
        index = len(body) - int(len(body) > 0 and body[-1].tag == q("w:sectPr"))
    else:
        if (type(after_paragraph) is not int or not 1 <= after_paragraph <= len(body)
                or body[after_paragraph - 1].tag != q("w:p")):
            raise ValueError("Choose a top-level paragraph body_child_index from the inventory")
        index = after_paragraph
    paragraph = ET.Element(q("w:p"))
    paragraph.append(text_run(slot))
    body.insert(index, paragraph)
    _, location = select_template_section(document, slot)
    parts["word/document.xml"] = ET.tostring(document, encoding="UTF-8", xml_declaration=True, standalone=True)
    write_package(parts, output)
    return {"output": str(output), "source_sha256": sha256, "after_paragraph": after_paragraph,
            "changed_parts": ["word/document.xml"], "location": location}
