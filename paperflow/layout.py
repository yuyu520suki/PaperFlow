"""Inspect and partition one manuscript for reviewed template positions."""
from __future__ import annotations

from collections import Counter
from copy import deepcopy
from pathlib import Path

from lxml import etree as ET

from .package import MARKER, NS, digest, part_fields, q, read_package, write_package, xml
from .template import inventory_template


def _source_package(source):
    source = Path(source).resolve(strict=True)
    if source.suffix.lower() != ".docx":
        raise ValueError("Inspect or recover other inputs to DOCX before assigning them to multiple template slots")
    signature = digest(source.read_bytes())
    parts = read_package(source)
    if digest(source.read_bytes()) != signature:
        raise ValueError("The source changed while it was being inspected")
    document = xml(parts["word/document.xml"])
    body = document.find(q("w:body"))
    if body is None:
        raise ValueError("The source does not contain a Word body")
    terminals = body.findall(q("w:sectPr"))
    if len(terminals) > 1 or (terminals and terminals[0] is not body[-1]):
        raise ValueError("The source has misplaced section properties")
    blocks = list(body)[:-1] if terminals else list(body)
    if not blocks:
        raise ValueError("The source Word body is empty")
    return source, signature, parts, document, blocks, terminals[0] if terminals else None


def _source_inventory(source, signature, blocks):
    return {"path": str(source), "sha256": signature, "blocks": [
        {"body_child_index": i, "kind": ET.QName(block).localname,
         "text_preview": " ".join(text.text or "" for text in block.iter(q("w:t")))[:120],
         "section_break": block.find("w:pPr/w:sectPr", NS) is not None}
        for i, block in enumerate(blocks, 1)
    ]}


def inspect_layout(source, template):
    source, signature, _, _, blocks, _ = _source_package(source)
    destinations = inventory_template(template)
    slots = sorted((slot for slot in destinations["slots"] if slot["valid"]), key=lambda item: item["body_child_index"])
    return {"source": _source_inventory(source, signature, blocks), "template": destinations,
            "plan": {"schema_version": 1, "source_sha256": signature, "template_sha256": destinations["sha256"],
                     "assignments": [{"slot": slot["slot"], "start": None, "end": None} for slot in slots]}}


def _checked_ranges(assignments, block_count):
    if not isinstance(assignments, list) or not assignments:
        raise ValueError("A layout plan needs at least one assignment")
    next_index = 1
    for assignment in assignments:
        if not isinstance(assignment, dict):
            raise ValueError("Each layout assignment must be an object")
        name, first, last = (assignment.get(key) for key in ("slot", "start", "end"))
        if (not isinstance(name, str) or not name or type(first) is not int or type(last) is not int
                or first != next_index or not first <= last <= block_count):
            raise ValueError(f"Layout must cover every source body block in order. Expected block {next_index} for {name!r}")
        role_map = assignment.get("role_style_map", {})
        if not isinstance(role_map, dict) or any(not isinstance(k, str) or not k or not isinstance(v, str) or not v
                                                 for k, v in role_map.items()):
            raise ValueError(f"The style map for {name!r} must map roles to template style IDs")
        next_index = last + 1
    if next_index != block_count + 1:
        raise ValueError(f"The layout plan leaves source body blocks {next_index} through {block_count} unassigned")


_RANGE_TAGS = {
    q(f"w:{prefix}Range{end}"): (prefix, end == "Start")
    for prefix in ("comment", "moveFrom", "moveTo", "customXmlIns", "customXmlDel", "customXmlMoveFrom", "customXmlMoveTo")
    for end in ("Start", "End")
}
_RANGE_TAGS.update({q("w:bookmarkStart"): ("bookmark", True), q("w:bookmarkEnd"): ("bookmark", False),
                    q("w:permStart"): ("permission", True), q("w:permEnd"): ("permission", False)})


def _verify_boundaries(document, blocks, assignments):
    _, errors = part_fields(document)
    if errors:
        raise ValueError(f"Source contains malformed Word fields: {errors}")
    fields = 0
    ranges = set()
    boundaries = {part["end"]: part["slot"] for part in assignments[:-1]}
    for index, block in enumerate(blocks, 1):
        for element in block.iter():
            if element.tag == q("w:fldChar"):
                event = element.get(q("w:fldCharType"))
                fields += (event == "begin") - (event == "end")
            if element.tag in _RANGE_TAGS:
                kind, starts = _RANGE_TAGS[element.tag]
                key = (kind, element.get(q("w:id")))
                if key[1] is None or starts == (key in ranges):
                    raise ValueError(f"Malformed {kind} range around source body block {index}")
                if starts:
                    ranges.add(key)
                else:
                    ranges.remove(key)
        if index in boundaries and (fields or ranges):
            active = sorted({kind for kind, _ in ranges} | ({"Word field"} if fields else set()))
            raise ValueError(f"The boundary after {boundaries[index]!r} cuts a range at block {index}: {', '.join(active)}")
    if fields or ranges:
        raise ValueError("Source contains an unclosed field or annotation range")


def validate_layout(source, template, plan):
    source, signature, _, document, blocks, _ = _source_package(source)
    if not isinstance(plan, dict) or type(plan.get("schema_version")) is not int or plan["schema_version"] != 1:
        raise ValueError("Unknown layout plan version")
    destinations = inventory_template(template)
    if plan.get("source_sha256") != signature or plan.get("template_sha256") != destinations["sha256"]:
        raise ValueError("The source or template has changed since the layout plan was inspected")
    assignments = plan.get("assignments")
    _checked_ranges(assignments, len(blocks))
    slots = {item["slot"].casefold(): item for item in destinations["slots"]}
    assigned = set()
    last = 0
    for assignment in assignments:
        name = assignment["slot"].casefold()
        destination = slots.get(name)
        if destination is None or not destination["valid"]:
            raise ValueError(f"Invalid or missing destination slot: {assignment['slot']!r}")
        if name in assigned or destination["body_child_index"] <= last:
            raise ValueError("Template positions must be unique, nonoverlapping and assigned in document order")
        assigned.add(name)
        last = destination.get("body_end_index", destination["body_child_index"])
    unused_markers = [item["slot"] for key, item in slots.items() if key not in assigned and item["slot"].startswith("{{")]
    if unused_markers:
        raise ValueError(f"Unassigned template placeholders: {unused_markers}")
    _verify_boundaries(document, blocks, assignments)
    validate_bookmark_names(document, xml(read_package(template)["word/document.xml"]), include_markers=True)
    return [{**item, "slot": slots[item["slot"].casefold()]["slot"]} for item in assignments]


def _bookmark_names(document):
    return [name for item in document.iter(q("w:bookmarkStart"))
            if (name := item.get(q("w:name"), "").casefold()) and name != "_goback"]


def validate_bookmark_names(source_document, template_document, *, include_markers=False):
    """Prevent name collisions without requiring references to resolve before assembly."""
    source_names = _bookmark_names(source_document)
    template_names = _bookmark_names(template_document)
    if include_markers:
        for paragraph in source_document.iter(q("w:p")):
            for match in MARKER.finditer("".join(paragraph.itertext())):
                command, _, args = match[1].partition(" ")
                if command in {"EQNO", "ANCHOR"}:
                    source_names.append(args.casefold())
                elif command in {"CAPTION", "CAPTIONNO"}:
                    source_names.append(args.partition(" ")[2].partition("|")[0].casefold())
    conflicts = {name for name, count in Counter(source_names + template_names).items() if count > 1}
    if conflicts:
        raise ValueError(f"Bookmark names must be distinct across the source and template: {sorted(conflicts)}")


def split_source(source, assignments, output_dir):
    """Write chunk copies, retaining opaque media, settings and native payloads."""
    source, signature, parts, document, blocks, terminal = _source_package(source)
    _checked_ranges(assignments, len(blocks))
    _verify_boundaries(document, blocks, assignments)
    root = Path(output_dir).resolve()
    root.mkdir(parents=True, exist_ok=False)
    outputs = []
    for number, assignment in enumerate(assignments, 1):
        selected = blocks[assignment["start"] - 1:assignment["end"]]
        chunk = deepcopy(document)
        body = chunk.find(q("w:body"))
        for element in list(body):
            body.remove(element)
        for element in selected:
            body.append(deepcopy(element))
        # Source section properties at this split point are superseded by
        # the destination slot. Keep the paragraph and its text in the chunk.
        final_break = body[-1].find("w:pPr/w:sectPr", NS)
        if final_break is not None:
            final_break.getparent().remove(final_break)
        following = (part.find("w:pPr/w:sectPr", NS) for part in blocks[assignment["end"] - 1:])
        section = next((item for item in following if item is not None), terminal)
        if section is not None:
            body.append(deepcopy(section))
        parts["word/document.xml"] = ET.tostring(chunk, encoding="UTF-8", xml_declaration=True, standalone=True)
        outputs.append(write_package(parts, root / f"part-{number:03d}.docx"))
    if digest(source.read_bytes()) != signature:
        raise ValueError("The source changed while splitting the reviewed ranges")
    return outputs
