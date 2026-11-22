"""Migrate an explicitly numbered bibliography to links to local Zotero records."""
from __future__ import annotations

from collections import defaultdict
from copy import deepcopy
from pathlib import Path
import re

from lxml import etree as ET

from . import semantics
from .doi import normalize_doi, resolve_dois
from .package import NS, digest, q, text_run, write_package, xml
from .zotero import ZoteroClient


_MAIN = "word/document.xml"
_ENTRY = re.compile(r"^\s*\[([1-9][0-9]{0,8})\]\s*(\S.*)$", re.S)
_DOI_TOKEN = re.compile(r"(?:https?://(?:dx\.)?doi\.org/|\b10\.[0-9]+(?:\.[0-9]+)*/)[^\s<>\"“”]+", re.I)
_UNSAFE_ENTRY = {q("w:" + local) for local in (
    "fldSimple", "fldChar", "instrText", "drawing", "pict", "object", "sdt", "ins", "del", "moveFrom", "moveTo",
    "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "commentReference",
    "footnoteReference", "endnoteReference", "sectPr", "permStart", "permEnd",
)} | {q("m:oMath")}


def _visible(paragraph):
    return "".join(node.text or "" for node in paragraph.iter(q("w:t")))


def _extract_dois(value):
    result = []
    for match in _DOI_TOKEN.finditer(value):
        literal = match[0]
        candidate = literal.rstrip(".,;:，。；：")
        for left, right in (("(", ")"), ("[", "]"), ("{", "}")):
            while candidate.endswith(right) and candidate.count(right) > candidate.count(left):
                candidate = candidate[:-1].rstrip(".,;:，。；：")
        # A punctuation character can be part of a DOI suffix. Let exact
        # library metadata distinguish a literal token from sentence punctuation.
        choices = {doi for token in (literal, candidate) if (doi := normalize_doi(token))}
        if choices:
            result.append(choices)
    return result


def _read_static_source(source):
    source = Path(source).resolve(strict=True)
    if source.suffix.lower() != ".docx":
        raise ValueError("Bibliography migration requires an editable DOCX draft")
    parts, signature = semantics._load_source(source)
    for name, data in parts.items():
        if name.startswith("word/") and name.endswith(".xml"):
            story = "".join(xml(data).itertext())
            if any(marker in story for marker in ("ZOTERO_ITEM CSL_CITATION", "ZOTERO_BIBL", "[[CITE ", "[[BIBLIOGRAPHY")):
                raise ValueError("Existing Zotero fields or markers require a mixed-citation review in Word")
    document = xml(parts[_MAIN])
    body = document.find(q("w:body"))
    if body is None:
        raise ValueError("Word draft has no body")
    headings = [i for i, child in enumerate(body) if child.tag == q("w:p")
                and semantics._BIBLIOGRAPHY.fullmatch(_visible(child).strip())]
    if len(headings) != 1:
        raise ValueError("Expected one standalone References or 参考文献 heading")
    heading = headings[0]
    end = len(body)
    entries = []
    numbers = set()
    relations = {}
    if "word/_rels/document.xml.rels" in parts:
        relations = {item.get("Id"): item for item in xml(parts["word/_rels/document.xml.rels"])}
    for index in range(heading + 1, len(body)):
        paragraph = body[index]
        if paragraph.tag == q("w:sectPr"):
            end = index
            break
        value = _visible(paragraph).strip()
        if paragraph.tag == q("w:p") and semantics._APPENDIX.fullmatch(value):
            end = index
            break
        if paragraph.tag != q("w:p") or any(node.tag in _UNSAFE_ENTRY for node in paragraph.iter()):
            raise ValueError(f"References contain a protected or non-paragraph item at body index {index}")
        if not value:
            continue
        match = _ENTRY.fullmatch(value)
        if match is None:
            raise ValueError(f"Expected one numbered reference per paragraph; check body index {index}")
        number = match[1]
        if number in numbers:
            raise ValueError(f"Reference number {number} occurs more than once")
        numbers.add(number)
        dois = _extract_dois(match[2])
        for link in paragraph.iter(q("w:hyperlink")):
            relationship = relations.get(link.get(q("r:id")))
            if relationship is not None and relationship.get("TargetMode") == "External":
                target = normalize_doi(relationship.get("Target", ""))
                if target:
                    dois.append({target})
        variants = set.intersection(*dois) if dois else set()
        if dois and not variants:
            raise ValueError(f"Reference {number} contains more than one DOI. Resolve its identity first.")
        variants = sorted(variants, key=lambda value: (len(value), value))
        entries.append({"number": number, "text": match[2], "doi": variants[0] if variants else None,
                        "doi_variants": variants, "body_index": index})
    if not entries:
        raise ValueError("No static numbered references found below the heading")
    return source, parts, signature, document, heading, end, entries


def _citation_locations(parts, document, heading, end):
    body = document.find(q("w:body"))
    result = []
    for index, child in enumerate(body):
        if heading <= index < end:
            continue
        for paragraph in child.iter(q("w:p")):
            text = _visible(paragraph)
            for match in semantics._CITATION.finditer(text):
                result.append({"part": _MAIN, "body_index": index, "top_level": paragraph is child,
                               "start": match.start(), "end": match.end(), "source_text": match[0],
                               "context": text[max(0, match.start() - 70):match.end() + 70],
                               "numbers": semantics._citation_numbers(match[1])})
    for name, data in parts.items():
        if not re.fullmatch(r"word/(?:header[0-9]+|footer[0-9]+|footnotes|endnotes)\.xml", name):
            continue
        for paragraph in xml(data).iter(q("w:p")):
            for match in semantics._CITATION.finditer(_visible(paragraph)):
                result.append({"part": name, "body_index": None, "top_level": False,
                               "source_text": match[0], "numbers": semantics._citation_numbers(match[1])})
    return result


def _edit_locator(value):
    return (value["body_index"], value.get("start"), value.get("end"), value["source_text"])


def plan_migration(source, *, client=None):
    """Suggest only unique exact DOI matches, keeping editable choices separate."""
    _, parts, signature, document, heading, end, entries = _read_static_source(source)
    dois = {doi for entry in entries for doi in entry["doi_variants"]}
    matches = resolve_dois(sorted(dois), client=client) if dois else {}
    citation_map = {}
    for entry in entries:
        candidates = [candidate for doi in entry["doi_variants"] for candidate in matches.get(doi, [])]
        entry["candidates"] = candidates
        entry["match_status"] = ("unique_doi" if len(candidates) == 1 else "multiple_doi_matches" if candidates
                                 else "missing_in_library" if entry["doi"] else "manual_key_required")
        if len(candidates) == 1:
            entry["doi"] = candidates[0]["doi"]
            citation_map[entry["number"]] = candidates[0]["key"]
    locations = _citation_locations(parts, document, heading, end)
    cited = {number for item in locations for number in item["numbers"] or []}
    available = {entry["number"] for entry in entries}
    edits = [edit for edit in semantics._make_plan(parts, signature, citation_map)["edits"] if edit["operation"] == "zotero_marker"]
    proposed = {_edit_locator(edit) for edit in edits}
    return {"schema_version": 1, "source_sha256": signature, "citation_map": citation_map, "include_uncited": False,
            "bibliography_body_range": [heading + 1, end], "entries": entries, "citation_locations": locations,
            "not_automatically_linked": [loc for loc in locations if not loc["top_level"] or _edit_locator(loc) not in proposed],
            "missing_reference_numbers": sorted(cited - available, key=int),
            "uncited_numbers": sorted(available - cited, key=int)}


def apply_migration(source, reviewed_plan, output, *, client=None):
    """Apply all reviewed links and retire only the static bibliography in a new copy."""
    source, parts, signature, document, heading, end, entries = _read_static_source(source)
    if (not isinstance(reviewed_plan, dict) or type(reviewed_plan.get("schema_version")) is not int
            or reviewed_plan["schema_version"] != 1 or reviewed_plan.get("source_sha256") != signature):
        raise ValueError("Bibliography migration plan does not match this saved source")
    include_uncited = reviewed_plan.get("include_uncited", False)
    if type(include_uncited) is not bool:
        raise ValueError("include_uncited must be true or false")
    citation_map = semantics._item_map(reviewed_plan.get("citation_map"))
    available = {entry["number"] for entry in entries}
    if set(citation_map) != available or len(set(citation_map.values())) != len(entries):
        raise ValueError("Assign one distinct reviewed Zotero key to every static bibliography entry")
    locations = _citation_locations(parts, document, heading, end)
    cited = {number for loc in locations for number in loc["numbers"] or []}
    if cited - available:
        raise ValueError(f"Cited numbers have no static bibliography entry: {sorted(cited - available, key=int)}")
    if not cited:
        raise ValueError("No numbered in-text citations were found to migrate")
    uncited = sorted(available - cited, key=int)
    if uncited and not include_uncited:
        raise ValueError(f"Static bibliography entries are uncited: {uncited}. Review and explicitly retain them or edit the source.")
    edits = [edit for edit in semantics._make_plan(parts, signature, citation_map)["edits"] if edit["operation"] == "zotero_marker"]
    proposed = {_edit_locator(edit) for edit in edits}
    unsupported = [loc for loc in locations if not loc["top_level"] or _edit_locator(loc) not in proposed]
    if unsupported or len(locations) != len(edits):
        raise ValueError(f"Some numbered citations need a Word review before replacing the static bibliography: {unsupported[:5]}")

    output = Path(output).resolve()
    if output.suffix.lower() != ".docx" or output.exists():
        raise ValueError("Bibliography migration needs a new DOCX output path")
    client = client or ZoteroClient()
    identities = []
    for entry in entries:
        key = citation_map[entry["number"]]
        item = client.get_item(key)
        data = item["item_data"]
        actual_doi = normalize_doi(data.get("DOI", ""))
        if entry["doi_variants"] and actual_doi not in entry["doi_variants"]:
            raise ValueError(f"Zotero key {key} does not have the reviewed DOI for reference {entry['number']}")
        identities.append({"number": entry["number"], "key": key, "title": data.get("title"), "doi": actual_doi,
                           "match": "exact_doi" if entry["doi"] else "reviewed_manual_key"})

    body = document.find(q("w:body"))
    by_paragraph = defaultdict(list)
    for edit in edits:
        by_paragraph[edit["body_index"]].append(edit)
    for index, operations in by_paragraph.items():
        semantics._apply_edits(body[index], operations, 0)
    paragraph = ET.Element(q("w:p"))
    formatting = body[entries[0]["body_index"]].find("w:pPr", NS)
    if formatting is not None:
        paragraph.append(deepcopy(formatting))
    keys = [citation_map[number] for number in uncited]
    paragraph.append(text_run("[[BIBLIOGRAPHY" + (" " + "; ".join(keys) if keys else "") + "]]"))
    retired = [body[entry["body_index"]] for entry in entries]
    for old in retired:
        body.remove(old)
    body.insert(entries[0]["body_index"], paragraph)
    parts[_MAIN] = ET.tostring(document, encoding="UTF-8", xml_declaration=True, standalone=True)
    if signature != digest(source.read_bytes()):
        raise ValueError("The original draft changed while applying the migration")
    target = write_package(parts, output)
    return {"output": str(target), "source_unchanged": True, "source_sha256": signature,
            "citation_markers": len(edits), "retired_entries": len(entries), "uncited_numbers": uncited,
            "identities": identities, "word_zotero_refresh_required": True}
