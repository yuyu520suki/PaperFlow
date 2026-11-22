"""Plan conservative upgrades for a Word draft with literal numbered references."""

from __future__ import annotations

from collections import Counter, defaultdict
from collections.abc import Mapping
from copy import deepcopy
from pathlib import Path
import re

from lxml import etree as ET

from .formatting import _default_paragraph_style, _paragraph_role, _styles
from .package import (digest, field_runs, part_fields, q, read_package, sequence_number_runs,
                      text_run, validate_chapter_caption_options, write_package, xml)
from .placement import _figure_parts
from .zotero import ITEM_KEY


_MAIN = "word/document.xml"
_RELATIONSHIPS = "word/_rels/document.xml.rels"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_LABEL = r"(?P<label>Figure|Fig\.?|Table|图|表)"
_NUMBER = r"(?P<number>[0-9]+(?:[.\-–][0-9]+)*)"
_CAPTION_START = r"^\s*" + _LABEL + r"\s*" + _NUMBER + r"(?![0-9]|[.\-–][0-9])"
_CAPTION = re.compile(_CAPTION_START + r"(?:\s*[:：.、]\s*\S|\s*$)", re.I)
_STYLED_CAPTION = re.compile(_CAPTION_START + r"(?:(?:\s*[:：.、]\s*|\s+)\S|\s*$)", re.I)
_REFERENCE = re.compile(r"(?<![A-Za-z0-9_])" + _LABEL + r"\s*" + _NUMBER, re.I)
_PLURAL_REFERENCE = re.compile(r"\b(?:Figures|Figs\.?|Tables)\s+[0-9]+", re.I)
_REFERENCE_GROUP = re.compile(r"^\s*(?:[,，、&]\s*|and\s+|和\s*|与\s*)[0-9]", re.I)
_EXTERNAL_REFERENCE = re.compile(r"^\s+(?:of|in|from)\s+", re.I)
_CHAPTER_RANGE_SEPARATOR = r"(?:[-–—]|to|through|至|到)"
_LABELED_RANGE = re.compile(r"\s*" + _CHAPTER_RANGE_SEPARATOR + r"\s*", re.I)
_BARE_CHAPTER_RANGE = re.compile(r"^\s*" + _CHAPTER_RANGE_SEPARATOR + r"\s*[0-9]", re.I)
_BARE_RANGE_START = re.compile(r"(?<![A-Za-z0-9_])[0-9]+(?:[.\-–][0-9]+)*\s*"
                               + _CHAPTER_RANGE_SEPARATOR + r"\s*$", re.I)
_CHAPTER_EXTERNAL_REFERENCE = re.compile(r"^\s*(?:来自|引自|出自)", re.I)
_POSITIVE = r"[1-9][0-9]{0,8}"
_CHAPTER_NUMBER = re.compile(rf"({_POSITIVE})[.\-–]({_POSITIVE})")
_CITATION = re.compile(r"(?<![A-Za-z0-9_\[\]])\[([0-9]+(?:\s*[,;–-]\s*[0-9]+)*)\](?![A-Za-z0-9_\[\]])")
_BIBLIOGRAPHY = re.compile(
    r"^(?:[0-9]+[.、]?\s+)?(?:references|bibliography|works cited|参考文献|引用文献)\s*[:：]?$", re.I
)
_APPENDIX = re.compile(r"^(?:appendix(?:\s+[A-Z0-9]+)?|附录\s*[A-Z0-9]*)(?:\s*[:：.]\s*.+)?$", re.I)
_EXISTING_SEQ = re.compile(r'^\s*SEQ\s+(?:"([^"]+)"|([^\s\\]+))', re.I)


def _kind(label):
    if label.lower() in {"figure", "fig", "fig.", "图"}:
        return "Figure"
    if label.lower() in {"table", "表"}:
        return "Table"
    return None


def _visible(node):
    return "".join(text.text or "" for text in node.iter(q("w:t")))


def _plain_runs(paragraph):
    """Represent simple text runs by character spans, or refuse mixed content."""
    offset = 0
    spans = []
    for child in paragraph:
        if child.tag == q("w:pPr"):
            if child.find(q("w:sectPr")) is not None or child.find(q("w:pPrChange")) is not None:
                return None
            continue
        if child.tag != q("w:r"):
            return None
        text = child.find(q("w:t"))
        properties = child.find(q("w:rPr"))
        if text is None or len(child) != 1 + (properties is not None):
            return None
        if properties is not None and properties.find(q("w:rPrChange")) is not None:
            return None
        value = text.text or ""
        if value:
            spans.append((offset, offset + len(value), child))
            offset += len(value)
    return spans


def _span_properties(spans, start, end):
    if not spans:
        raise ValueError("Paragraph contains fields, objects, annotations or other complex content")
    overlaps = [run.find(q("w:rPr")) for left, right, run in spans if start < right and end > left]
    if not overlaps:
        raise ValueError("Text span does not reach a simple run")
    signatures = {
        ET.tostring(properties, method="c14n", exclusive=True) if properties is not None else b""
        for properties in overlaps
    }
    if len(signatures) != 1:
        raise ValueError("Number token spans different character formatting")
    return overlaps[0]


def _item_map(citation_map):
    if citation_map is None:
        return {}
    if not isinstance(citation_map, Mapping):
        raise ValueError("Citation map must map reference numbers to Zotero item keys")
    result = {}
    for number, key in citation_map.items():
        number = str(number)
        if not re.fullmatch(r"[1-9][0-9]{0,8}", number) or not isinstance(key, str) or not ITEM_KEY.fullmatch(key):
            raise ValueError("Citation map requires positive decimal reference numbers and eight-character Zotero item keys")
        result[number] = key
    return result


def _citation_numbers(value):
    numbers = []
    for group in re.split(r"\s*[,;]\s*", value):
        boundaries = re.split(r"\s*[-–]\s*", group)
        if len(boundaries) > 2 or any(not re.fullmatch(r"[1-9][0-9]{0,8}", number) for number in boundaries):
            return None
        first = int(boundaries[0])
        last = int(boundaries[-1])
        if not 0 <= last - first < 50 or len(numbers) + last - first >= 50:
            return None
        numbers.extend(str(number) for number in range(first, last + 1))
    return list(dict.fromkeys(numbers))


def _caption_style(parts):
    if "word/styles.xml" not in parts:
        return None
    for style in xml(parts["word/styles.xml"]).iter(q("w:style")):
        name = style.find(q("w:name"))
        if style.get(q("w:type")) == "paragraph" and name is not None:
            if name.get(q("w:val"), "").lower() in {"caption", "题注"}:
                return style.get(q("w:styleId"))
    return None


def _load_source(source):
    source = Path(source).resolve(strict=True)
    fingerprint = digest(source.read_bytes())
    parts = read_package(source)
    if fingerprint != digest(source.read_bytes()):
        raise ValueError("Source changed while it was being inspected. Run the plan again on the saved draft.")
    return parts, fingerprint


def _chapter_policy(level, separator):
    validate_chapter_caption_options(level, separator)
    return {"level": level, "separator": separator} if level is not None else None


def _number_key(value, chapter_policy):
    match = _CHAPTER_NUMBER.fullmatch(value) if chapter_policy is not None else None
    return tuple(map(int, match.groups())) if match else value


def _chapter_owners(parts, children, paragraphs, existing_fields, level):
    """Map body paragraphs to the preceding heading role used by normalization."""
    styles = _styles(xml(parts["word/styles.xml"])) if "word/styles.xml" in parts else {}
    default_style = _default_paragraph_style(styles)
    owners = {}
    heading = None
    for index, (text, _) in paragraphs.items():
        if index not in existing_fields and text.strip():
            properties = children[index].find(q("w:pPr"))
            style = properties.find(q("w:pStyle")) if properties is not None else None
            style_id = style.get(q("w:val")) if style is not None else default_style
            if _paragraph_role(properties, style_id, styles) == f"heading{level}":
                heading = index
        owners[index] = heading
    return owners


def _make_plan(parts, fingerprint, citation_map, chapter_policy=None):
    document = xml(parts[_MAIN])
    body = document.find(q("w:body"))
    if body is None:
        raise ValueError("Word draft has no main document body")
    children = list(body)
    instructions, errors = part_fields(body)
    if errors:
        raise ValueError("Cannot plan edits around malformed Word fields: " + "; ".join(errors))
    native_kinds = set()
    for instruction in instructions:
        match = _EXISTING_SEQ.match(instruction)
        if match:
            native_kinds.add(_kind(match[1] or match[2]))

    relationships = {}
    if _RELATIONSHIPS in parts:
        relationships = {r.get("Id"): r for r in xml(parts[_RELATIONSHIPS])}
    objects, paragraphs, issues = {}, {}, []
    existing_field_paragraphs = set()
    open_fields = 0
    for index, child in enumerate(children):
        field_events = [
            event.get(q("w:fldCharType")) for event in child.iter(q("w:fldChar"))
        ]
        if child.tag == q("w:p"):
            value = _visible(child)
            if open_fields or field_events or child.find(f".//{q('w:fldSimple')}") is not None:
                existing_field_paragraphs.add(index)
            spans = None if index in existing_field_paragraphs else _plain_runs(child)
            paragraphs[index] = (value, spans)
            if child.find(f".//{q('w:drawing')}") is not None and child.find(f".//{q('w:object')}") is None:
                try:
                    embedded = _figure_parts(child, relationships, parts)
                except ValueError:
                    embedded = []
                if embedded:
                    objects[index] = "Figure"
        elif child.tag == q("w:tbl") and child.find(f"./{q('w:tr')}/{q('w:tc')}") is not None:
            objects[index] = "Table"
        open_fields += field_events.count("begin") - field_events.count("end")

    bibliography = next((index for index, (value, _) in paragraphs.items()
                         if index not in existing_field_paragraphs and _BIBLIOGRAPHY.fullmatch(value.strip())), None)
    appendix = next((index for index, (value, _) in paragraphs.items()
                     if bibliography is not None and index > bibliography and index not in existing_field_paragraphs
                     and _APPENDIX.fullmatch(value.strip())
                     and len(value.strip()) <= 80), len(children))

    def in_bibliography(index):
        return bibliography is not None and bibliography <= index < appendix

    def issue(code, index, token, reason):
        issues.append({"code": code, "body_index": index, "token": token, "reason": reason})

    caption_style = _caption_style(parts)
    candidates = {}
    for index, (value, _) in paragraphs.items():
        if in_bibliography(index) or index in existing_field_paragraphs:
            continue
        style = children[index].find(f"./{q('w:pPr')}/{q('w:pStyle')}")
        is_caption = caption_style is not None and style is not None and style.get(q("w:val")) == caption_style
        match = (_STYLED_CAPTION if is_caption else _CAPTION).match(value)
        if not match:
            continue
        kind = _kind(match["label"])
        neighbors = [
            neighbor for neighbor in (index - 1, index + 1)
            if objects.get(neighbor) == kind and not in_bibliography(neighbor)
            and children[min(index, neighbor)].find(f"./{q('w:pPr')}/{q('w:sectPr')}") is None
        ]
        if neighbors:
            candidates[index] = (match, kind, neighbors)

    by_object = Counter(neighbor for _, _, neighbors in candidates.values() for neighbor in neighbors)
    keys = {index: _number_key(match["number"], chapter_policy) for index, (match, _, _) in candidates.items()}
    by_number = Counter((kind, keys[index]) for index, (_, kind, _) in candidates.items())
    owners, chapter_prefixes, prefix_owners, numbering_modes = {}, defaultdict(set), defaultdict(set), defaultdict(set)
    if chapter_policy is not None:
        owners = _chapter_owners(parts, children, paragraphs, existing_field_paragraphs, chapter_policy["level"])
        for index, (_, kind, _) in candidates.items():
            key = keys[index]
            if isinstance(key, tuple):
                numbering_modes[kind].add("chapter")
                owner = owners[index]
                if owner is not None:
                    chapter_prefixes[owner].add(key[0])
                    prefix_owners[key[0]].add(owner)
            elif re.fullmatch(_POSITIVE, key):
                numbering_modes[kind].add("continuous")
    names = {bookmark.get(q("w:name"), "").casefold() for bookmark in document.iter(q("w:bookmarkStart"))}
    targets, edits = {}, []

    def append_edit(operation, index, match_span, **fields):
        value, spans = paragraphs[index]
        start, end = match_span
        token = value[start:end]
        try:
            _span_properties(spans, start, end)
        except ValueError as exc:
            issue("unsafe_text_span", index, token, str(exc))
            return False
        edits.append({"operation": operation, "body_index": index, "start": start, "end": end,
                      "source_text": token, **fields})
        return True

    sequence_count = Counter()
    for index, (match, kind, neighbors) in candidates.items():
        number = match["number"]
        key = keys[index]
        if not isinstance(key, tuple) and not re.fullmatch(_POSITIVE, number):
            issue("unsupported_caption_number", index, number, "Use plain positive sequence numbers or map the chapter-numbered caption manually")
            continue
        if len(neighbors) != 1 or by_object[neighbors[0]] != 1 or by_number[(kind, key)] != 1:
            issue("ambiguous_caption", index, number, "The object or number has more than one matching caption")
            continue
        if kind in native_kinds:
            issue("mixed_caption_numbering", index, number, f"Existing native {kind} sequences need a coordinated numbering review")
            continue
        if len(numbering_modes[kind]) > 1:
            issue("mixed_caption_numbering", index, number, "Continuous and chapter numbering for the same caption kind need review")
            continue
        details = {}
        counter_key = (kind, None)
        if isinstance(key, tuple):
            prefix, sequence_index = key
            owner = owners[index]
            if owner is None:
                issue("caption_without_chapter", index, number, "No preceding source heading has the selected chapter level")
                continue
            if len(chapter_prefixes[owner]) != 1 or len(prefix_owners[prefix]) != 1:
                issue("ambiguous_chapter_prefix", index, number, "The chapter heading and printed caption prefix do not match one-to-one")
                continue
            counter_key = (kind, owner)
            separator = chapter_policy["separator"]
            details = {"chapter_level": chapter_policy["level"], "chapter_heading_body_index": owner,
                       "chapter_prefix": prefix, "sequence_index": sequence_index, "separator": separator,
                       "target_number": f"{prefix}{separator}{sequence_index}"}
        else:
            sequence_index = int(number)
        if sequence_index != sequence_count[counter_key] + 1:
            issue("caption_numbering_gap", index, number, "The proposed Word sequence would change this original number on refresh")
            continue
        bookmark = f"PF_{kind}_{index + 1}"
        while bookmark.casefold() in names:
            bookmark += "_"
        if len(bookmark) > 40:
            raise ValueError("Cannot allocate an unused Word bookmark name")
        if append_edit("caption", index, match.span("number"), sequence=kind, bookmark=bookmark,
                       object_body_index=neighbors[0], caption_style=caption_style, **details):
            names.add(bookmark.casefold())
            targets[(kind, key)] = (bookmark, details.get("target_number"))
            sequence_count[counter_key] += 1

    for index, (value, _) in paragraphs.items():
        if index in candidates or in_bibliography(index) or index in existing_field_paragraphs:
            continue
        for match in _PLURAL_REFERENCE.finditer(value):
            issue("ambiguous_cross_reference", index, match[0], "Grouped references need review")
        references = list(_REFERENCE.finditer(value))
        range_starts = set()
        if chapter_policy is not None:
            for first, second in zip(references, references[1:]):
                if _LABELED_RANGE.fullmatch(value[first.end():second.start()]):
                    range_starts.update((first.start(), second.start()))
        for match in references:
            tail = value[match.end():]
            number = match["number"]
            target = targets.get((_kind(match["label"]), _number_key(number, chapter_policy)))
            if (
                _REFERENCE_GROUP.match(tail) or _EXTERNAL_REFERENCE.match(tail)
                or re.match(r"^[A-Za-z0-9_]", tail)
                or (chapter_policy is not None and (match.start() in range_starts
                    or _BARE_CHAPTER_RANGE.match(tail) or _BARE_RANGE_START.search(value, 0, match.start())
                    or _CHAPTER_EXTERNAL_REFERENCE.match(tail)))
            ):
                issue("ambiguous_cross_reference", index, match[0], "Grouped, suffixed or externally qualified references need review")
            elif target is None:
                issue("unresolved_cross_reference", index, match[0], "No uniquely matched static caption is being converted")
            else:
                bookmark, target_number = target
                details = {"target_number": target_number} if target_number is not None else {}
                append_edit("cross_reference", index, match.span("number"), bookmark=bookmark, **details)
        for match in _CITATION.finditer(value):
            if not value[:match.start()].strip():
                issue("possible_reference_list_entry", index, match[0], "A leading bracketed number may label a reference entry")
                continue
            numbers = _citation_numbers(match[1])
            if numbers is None:
                issue("unsupported_citation_numbers", index, match[0], "Citation number range needs explicit review")
            elif any(number not in citation_map for number in numbers):
                missing = [number for number in numbers if number not in citation_map]
                issue("unmapped_citation", index, match[0], "Missing Zotero item keys for " + ", ".join(missing))
            else:
                keys = list(dict.fromkeys(citation_map[number] for number in numbers))
                append_edit("zotero_marker", index, match.span(), numbers=numbers, item_keys=keys)

    plan = {
        "schema_version": 1,
        "source_sha256": fingerprint,
        "citation_map": citation_map,
        "edits": sorted(edits, key=lambda edit: (edit["body_index"], edit["start"])),
        "issues": issues,
        "scope": {
            "main_body_only": True,
            "existing_field_paragraphs_excluded": sorted(existing_field_paragraphs),
        },
        "bibliography": {
            "heading_body_index": bibliography,
            "end_body_index_exclusive": appendix if bibliography is not None else None,
            "action": "retain_existing_text",
        },
    }
    if chapter_policy is not None:
        plan["chapter_caption_policy"] = chapter_policy
    return plan


def plan_draft(source_docx, *, citation_map=None, caption_chapter_level=None, caption_chapter_separator="-"):
    """Return a JSON-ready plan. Body indices and character offsets are zero-based."""
    parts, fingerprint = _load_source(source_docx)
    policy = _chapter_policy(caption_chapter_level, caption_chapter_separator)
    return _make_plan(parts, fingerprint, _item_map(citation_map), policy)


def _clone_text(run, value):
    clone = deepcopy(run)
    element = clone.find(q("w:t"))
    element.text = value
    if value != value.strip():
        element.set(_XML_SPACE, "preserve")
    return clone


def _set_caption_style(paragraph, style):
    if not style:
        return
    properties = paragraph.find(q("w:pPr"))
    if properties is None:
        properties = ET.Element(q("w:pPr"))
        paragraph.insert(0, properties)
    assigned = properties.find(q("w:pStyle"))
    if assigned is None:
        assigned = ET.Element(q("w:pStyle"))
        properties.insert(0, assigned)
    elif assigned.get(q("w:val")) not in {"Normal", style}:
        return
    assigned.set(q("w:val"), style)


def _apply_edits(paragraph, edits, next_bookmark_id):
    spans = _plain_runs(paragraph)
    replacements = {}
    previous_end = 0
    value = _visible(paragraph)
    for edit in edits:
        start, end = edit["start"], edit["end"]
        token = edit["source_text"]
        if start < previous_end or value[start:end] != token:
            raise ValueError("Overlapping or stale draft edits")
        previous_end = end
        properties = _span_properties(spans, start, end)
        if edit["operation"] == "caption":
            next_bookmark_id += 1
            bookmark_id = str(next_bookmark_id)
            begin = ET.Element(q("w:bookmarkStart"), {q("w:id"): bookmark_id, q("w:name"): edit["bookmark"]})
            finish = ET.Element(q("w:bookmarkEnd"), {q("w:id"): bookmark_id})
            sequence = sequence_number_runs(edit["sequence"], chapter_level=edit.get("chapter_level"),
                                            separator=edit.get("separator", "-"), properties=properties,
                                            chapter_result=str(edit.get("chapter_prefix", "")),
                                            sequence_result=str(edit.get("sequence_index", token)))
            replacements[start] = [begin, *sequence, finish]
            _set_caption_style(paragraph, edit["caption_style"])
        elif edit["operation"] == "cross_reference":
            replacements[start] = field_runs(f' REF {edit["bookmark"]} \\h ', edit.get("target_number", token), properties)
        else:
            replacements[start] = [text_run("[[CITE " + "; ".join(edit["item_keys"]) + "]]", properties)]

    for left, right, run in spans:
        overlapping = [edit for edit in edits if edit["start"] < right and edit["end"] > left]
        if not overlapping:
            continue
        position = paragraph.index(run)
        paragraph.remove(run)
        cursor = left
        for edit in overlapping:
            start, end = edit["start"], edit["end"]
            if cursor < start:
                paragraph.insert(position, _clone_text(run, value[cursor:start]))
                position += 1
            if left <= start:
                for replacement in replacements[start]:
                    paragraph.insert(position, replacement)
                    position += 1
            cursor = min(end, right)
        if cursor < right:
            paragraph.insert(position, _clone_text(run, value[cursor:right]))
    return next_bookmark_id


def apply_draft_plan(source_docx, output_docx, plan):
    """Apply the exact generated plan to a separate copy of the same source."""
    parts, fingerprint = _load_source(source_docx)
    if not isinstance(plan, Mapping) or fingerprint != plan.get("source_sha256"):
        raise ValueError("Draft plan source SHA-256 does not match")
    policy = plan.get("chapter_caption_policy")
    if policy is not None:
        if not isinstance(policy, Mapping):
            raise ValueError("Chapter caption policy must give a heading level and separator")
        policy = _chapter_policy(policy.get("level"), policy.get("separator"))
    verified = _make_plan(parts, fingerprint, _item_map(plan.get("citation_map")), policy)
    if dict(plan) != verified:
        raise ValueError("Draft plan differs from the validated proposal. Regenerate it with the intended citation map.")
    edits = defaultdict(list)
    for edit in verified["edits"]:
        edits[edit["body_index"]].append(edit)
    if edits:
        document = xml(parts[_MAIN])
        body = document.find(q("w:body"))
        bookmark_id = max(
            (int(node.get(q("w:id"))) for tag in ("w:bookmarkStart", "w:bookmarkEnd") for node in document.iter(q(tag))),
            default=0,
        )
        for index, operations in edits.items():
            bookmark_id = _apply_edits(body[index], operations, bookmark_id)
        _, errors = part_fields(body)
        if errors:
            raise ValueError("Draft edits would damage Word field structure")
        parts[_MAIN] = ET.tostring(document, encoding="UTF-8", xml_declaration=True, standalone=True)
    target = write_package(parts, output_docx)
    counts = Counter(edit["operation"] for edit in verified["edits"])
    references = Counter(edit["bookmark"] for edit in verified["edits"] if edit["operation"] == "cross_reference")
    return {
        "output": str(target),
        "source_sha256": fingerprint,
        "captions": counts["caption"],
        "cross_references": counts["cross_reference"],
        "chapter_caption_policy": policy,
        "chapter_captions": [{"bookmark": edit["bookmark"], "number": edit["target_number"],
                              "references": references[edit["bookmark"]]}
                             for edit in verified["edits"] if edit["operation"] == "caption" and "target_number" in edit],
        "citation_markers": counts["zotero_marker"],
        "issues": verified["issues"],
        "bibliography": verified["bibliography"],
        "word_refresh_needed": bool(counts["caption"] or counts["cross_reference"]),
        "requires_zotero_resolution": bool(counts["zotero_marker"]),
        "changed_parts": [_MAIN] if edits else [],
    }
