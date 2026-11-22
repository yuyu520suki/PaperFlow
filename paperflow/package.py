"""Read and patch DOCX without round-tripping opaque OLE or citation parts."""
from __future__ import annotations

import hashlib
import json
import re
import zipfile
from collections import Counter
from copy import deepcopy
from pathlib import Path

from lxml import etree as ET

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "o": "urn:schemas-microsoft-com:office:office",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
_BIDI_CONTROLS = dict.fromkeys(map(ord, "\u061c\u200e\u200f\u202a\u202b\u202c\u202d\u202e\u2066\u2067\u2068\u2069"))
_PENDING_MARKERS = (
    "PAPERFLOW_PENDING", "{{PAPERFLOW_", "[[CITE ", "[[REF ", "[[TOC]]", "[[LOF]]", "[[LOT]]",
    "[[PAGEREF ", "[[HEADINGREF ", "[[ANCHOR ", "[[CAPTION ", "[[CAPTIONNO ", "[[EQNO ", "[[BIBLIOGRAPHY",
)


def q(name):
    prefix, local = name.split(":", 1)
    return "{" + NS[prefix] + "}" + local


def xml(data):
    return ET.fromstring(data, ET.XMLParser(resolve_entities=False, no_network=True))


def digest(data):
    return hashlib.sha256(data).hexdigest()


def without_bidi_controls(text):
    """Compare displayed Word text with exports that omit direction controls."""
    return text.translate(_BIDI_CONTROLS)


def read_package(path):
    path = Path(path).resolve(strict=True)
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if len(names) != len(set(names)):
            raise ValueError("Duplicate ZIP part names")
        if sum(i.file_size for i in z.infolist()) > 512 * 1024 * 1024:
            raise ValueError("DOCX exceeds inspection size limit")
        if any(n.startswith(("/", "\\")) or ".." in n.replace("\\", "/").split("/") for n in names):
            raise ValueError("Unsafe DOCX part name")
        parts = {n: z.read(n) for n in names}
    if "word/document.xml" not in parts:
        raise ValueError("Not a WordprocessingML document")
    return parts


def write_package(parts, output):
    target = Path(output).resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    # Exclusive creation protects both user originals and prior results.
    with target.open("xb") as stream:
        with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as z:
            for name, value in parts.items():
                z.writestr(name, value)
    return target


def part_fields(root):
    """Read complete complex field instructions, including instructions split across runs."""
    stack, fields, errors = [], [], []
    for e in root.iter():
        if e.tag == q("w:fldSimple"):
            fields.append(e.get(q("w:instr"), ""))
        elif e.tag == q("w:fldChar"):
            kind = e.get(q("w:fldCharType"))
            if kind == "begin":
                stack.append({"instruction": [], "result": False})
            elif kind == "separate":
                if not stack:
                    errors.append("field separator without begin")
                else:
                    stack[-1]["result"] = True
            elif kind == "end":
                if not stack:
                    errors.append("field end without begin")
                else:
                    fields.append("".join(stack.pop()["instruction"]))
        elif e.tag == q("w:instrText") and stack and not stack[-1]["result"]:
            stack[-1]["instruction"].append(e.text or "")
    if stack:
        errors.append(f"{len(stack)} unclosed fields")
    return fields, errors


def inspect_docx(path):
    path = Path(path).resolve(strict=True)
    parts = read_package(path)
    document = xml(parts["word/document.xml"])
    styles = xml(parts["word/styles.xml"]) if "word/styles.xml" in parts else None
    body = document.find(q("w:body"))
    fields, field_errors = [], []
    for name, data in parts.items():
        if name.startswith("word/") and name.endswith(".xml"):
            root = xml(data)
            fs, es = part_fields(root)
            fields.extend({"part": name, "instruction": f} for f in fs)
            field_errors.extend(f"{name}: {e}" for e in es)
    style_list = []
    if styles is not None:
        for s in styles.findall(q("w:style")):
            name = s.find(q("w:name"))
            style_list.append({
                "id": s.get(q("w:styleId")), "type": s.get(q("w:type")),
                "name": name.get(q("w:val")) if name is not None else None,
                "xml": ET.tostring(s, encoding="unicode"),
            })
    sections = [ET.tostring(s, encoding="unicode") for s in document.iter(q("w:sectPr"))]
    visible = "\n".join("".join(p.itertext()) for p in body.findall(q("w:p")))
    return {
        "path": str(path), "sha256": digest(path.read_bytes()),
        "sections": sections, "section_count": len(sections), "styles": style_list,
        "fields": fields, "field_errors": field_errors,
        "body_markers": re.findall(r"\{\{PAPERFLOW_[A-Z_]+\}\}", visible),
        "bookmarks": [b.get(q("w:name")) for b in document.iter(q("w:bookmarkStart"))],
        "content_controls": [e.get(q("w:val")) for e in document.iter(q("w:tag"))],
        "omml_count": len(document.findall(".//m:oMath", NS)),
        "ole_objects": [{"prog_id": o.get("ProgID"), "relationship": o.get(q("r:id"))} for o in document.iter(q("o:OLEObject"))],
        "parts": {n: {"size": len(d), "sha256": digest(d)} for n, d in parts.items()},
        "validation": {"package_inspected": True, "word_refreshed": False,
                       "zotero_refreshed": False, "mathtype_edit_verified": False,
                       "visual_reviewed": False, "real_template_verified": False},
    }


def audit_docx(path):
    report = inspect_docx(path)
    parts = read_package(path)
    root = xml(parts["word/document.xml"])
    bookmarks = [name.casefold() for name in report["bookmarks"] if name is not None]
    errors = list(report["field_errors"])
    if len(bookmarks) != len(set(bookmarks)):
        errors.append("Duplicate bookmark names")
    for field in report["fields"]:
        match = re.match(r'\s*(?:REF|PAGEREF)\s+(?:"([^"]+)"|([^\s\\]+))', field["instruction"], re.I)
        if match and (match[1] or match[2]).casefold() not in bookmarks:
            errors.append(f"Missing cross-reference target: {match[1] or match[2]}")
        if "ZOTERO_ITEM CSL_CITATION" in field["instruction"]:
            try:
                payload = json.loads(field["instruction"].split("CSL_CITATION", 1)[1])
                if not payload.get("citationItems"):
                    errors.append("Empty Zotero citation")
                for item in payload.get("citationItems", []):
                    if not item.get("uris"):
                        errors.append("Zotero citation item missing library URI")
            except (ValueError, KeyError):
                errors.append("Invalid Zotero citation JSON")
    pending_parts = {}
    for name, data in parts.items():
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        text = "".join(xml(data).itertext())
        tokens = [token for token in _PENDING_MARKERS if token in text]
        if tokens:
            pending_parts[name] = tokens
    pending = [token for token in _PENDING_MARKERS if any(token in tokens for tokens in pending_parts.values())]
    # Ordinary and native feature checks are intentionally separate.
    report["errors"] = errors
    report["pending_markers"] = pending
    report["pending_marker_parts"] = pending_parts
    report["structural_pass"] = not errors and not pending
    report["complete"] = False
    return report


def text_run(text, properties=None):
    r = ET.Element(q("w:r"))
    if properties is not None:
        r.append(deepcopy(properties))
    t = ET.SubElement(r, q("w:t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return r


def field_runs(instruction, result="PAPERFLOW_PENDING", properties=None):
    result_nodes = []
    for kind in ["begin", "instruction", "separate", "result", "end"]:
        if kind == "result":
            result_nodes.append(text_run(result, properties))
            continue
        r = ET.Element(q("w:r"))
        if kind == "instruction":
            e = ET.SubElement(r, q("w:instrText"))
            e.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            e.text = instruction
        else:
            e = ET.SubElement(r, q("w:fldChar"))
            e.set(q("w:fldCharType"), kind)
            if kind == "begin":
                e.set(q("w:dirty"), "true")
        result_nodes.append(r)
    return result_nodes


_BOOKMARK_NAME = r"[A-Za-z][A-Za-z0-9_]{0,39}"
MARKER = re.compile(
    r"\[\[(TOC|LOF|LOT|BIBLIOGRAPHY(?: [A-Z0-9;, ]+)?|CITE [A-Z0-9;, ]+|"
    rf"(?:EQNO|REF|PAGEREF|HEADINGREF|ANCHOR) {_BOOKMARK_NAME}|CAPTIONNO (?:Figure|Table) {_BOOKMARK_NAME}|"
    rf"CAPTION (?:Figure|Table|Equation) {_BOOKMARK_NAME}\|[^\]\r\n]+)\]\]"
)
_CAPTION_LISTS = {"LOF": r' TOC \h \z \c "Figure" ', "LOT": r' TOC \h \z \c "Table" '}
CHAPTER_CAPTION_SEPARATORS = ("-", ".", ":", "–", "—")
_CAPTION_LIST_PARAGRAPH_CHILDREN = {q("w:" + name) for name in (
    "pPr", "r", "bookmarkStart", "bookmarkEnd", "proofErr",
)}
_CAPTION_LIST_RUN_CHILDREN = {q("w:rPr"), q("w:t")}


def navigation_index_kind(instruction):
    if not re.match(r"\s*TOC(?:\s|$)", instruction, re.I):
        return None
    label = re.search(r'\s\\[ca]\s+(?:"([^"]+)"|([^\s\\]+))', instruction, re.I)
    if label:
        return {"figure": "LOF", "fig.": "LOF", "table": "LOT"}.get((label[1] or label[2]).casefold())
    if re.search(r"\s\\[ou](?=\s|$)", instruction, re.I):
        return "TOC"
    if not re.search(r"\s\\[acft](?=\s|$)", instruction, re.I):
        return "TOC"
    return None


def _standalone_caption_list(paragraph, text, marker):
    return (text.strip() == marker
            and all(node.tag in _CAPTION_LIST_PARAGRAPH_CHILDREN for node in paragraph)
            and all(node.tag in _CAPTION_LIST_RUN_CHILDREN
                    for run in paragraph.findall(q("w:r")) for node in run))


def _marker_keys(command):
    return [key for key in re.split(r"[;, ]+", command.partition(" ")[2]) if key]


def marked_zotero_keys(path):
    documents = _marker_documents(read_package(path))
    return {key for _, _, paragraphs in documents for kind, paragraph in paragraphs
            for match in MARKER.finditer("".join(paragraph.itertext()))
            if match[1].startswith("CITE ") or kind is None and match[1].startswith("BIBLIOGRAPHY ")
            for key in _marker_keys(match[1])}


def _marker_documents(parts):
    from .notes import note_ids, note_parts, normal_notes

    document = xml(parts["word/document.xml"])
    documents = [("word/document.xml", document, [(None, paragraph) for paragraph in document.iter(q("w:p"))])]
    for kind, (name, root) in note_parts(parts).items():
        references = Counter(note_ids(document, kind))
        definitions = Counter(note.get(q("w:id")) for note in normal_notes(root, kind))
        paragraphs = []
        for note in root:
            identifier = note.get(q("w:id"))
            supported = (note.tag == q("w:" + kind) and note.get(q("w:type"), "normal") == "normal"
                         and references[identifier] == 1 and definitions[identifier] == 1)
            if not supported:
                if MARKER.search("".join(note.itertext())):
                    raise ValueError(f"Markers require a uniquely referenced normal {kind}; check {name} note {identifier}")
                continue
            for paragraph in note.iter(q("w:p")):
                if MARKER.search("".join(paragraph.itertext())) and any(
                    ancestor.tag in {q("w:" + tag) for tag in ("txbxContent", "drawing", "pict", "del", "moveFrom")}
                    for ancestor in paragraph.iterancestors()
                ):
                    raise ValueError(f"A marker in a nested or revised {kind} needs review")
                paragraphs.append((kind, paragraph))
        documents.append((name, root, paragraphs))
    return documents


def validate_chapter_caption_options(caption_chapter_level, caption_chapter_separator):
    if caption_chapter_level is not None and (
        type(caption_chapter_level) is not int or not 1 <= caption_chapter_level <= 9
    ):
        raise ValueError("Chapter caption level must be a Word heading level from 1 to 9")
    if caption_chapter_separator not in CHAPTER_CAPTION_SEPARATORS:
        raise ValueError("Unsupported chapter caption separator")
    if caption_chapter_level is None and caption_chapter_separator != "-":
        raise ValueError("A chapter caption separator requires a chapter heading level")


def sequence_number_runs(kind, *, chapter_level=None, separator="-", properties=None,
                         chapter_result="PAPERFLOW_PENDING", sequence_result="PAPERFLOW_PENDING"):
    """Build the full number that a caption's cross-reference bookmark covers."""
    seq = f" SEQ {kind} \\* ARABIC "
    if chapter_level is None:
        return field_runs(seq, sequence_result, properties)
    return [*field_runs(f" STYLEREF {chapter_level} \\s ", chapter_result, properties),
            text_run(separator, properties),
            *field_runs(seq + f"\\s {chapter_level} ", sequence_result, properties)]


def expand_markers(source, output, citation_items=None, caption_style="Caption", *, equation_style="Equation",
                   caption_chapter_level=None, caption_chapter_separator="-"):
    """Expand semantic markers, preserving all nonedited ZIP parts byte for byte."""
    validate_chapter_caption_options(caption_chapter_level, caption_chapter_separator)
    parts = read_package(source)
    documents = _marker_documents(parts)
    document = documents[0][1]
    known = {e.get(q("w:name"), "").casefold() for e in document.iter(q("w:bookmarkStart"))}
    next_id = max([int(e.get(q("w:id"))) for e in document.iter(q("w:bookmarkStart"))] or [0]) + 1
    citations_used = False
    chapter_bookmarks = []

    def bookmark(name):
        nonlocal next_id
        if name.casefold() in known:
            raise ValueError(f"Duplicate bookmark: {name}")
        known.add(name.casefold())
        identifier = str(next_id)
        next_id += 1
        start = ET.Element(q("w:bookmarkStart"), {q("w:name"): name, q("w:id"): identifier})
        end = ET.Element(q("w:bookmarkEnd"), {q("w:id"): identifier})
        return start, end

    def numbered_bookmark(kind, name):
        start, end = bookmark(name)
        if caption_chapter_level is not None:
            chapter_bookmarks.append(name)
        return [start, *sequence_number_runs(kind, chapter_level=caption_chapter_level,
                                             separator=caption_chapter_separator), end]

    def replacement(command, properties):
        nonlocal citations_used
        if command == "TOC":
            return field_runs(' TOC \\o "1-3" \\h \\z \\u ')
        if command in _CAPTION_LISTS:
            return field_runs(_CAPTION_LISTS[command])
        if command.startswith(("REF ", "PAGEREF ")):
            return field_runs(f" {command} \\h ", properties=properties)
        if command.startswith("HEADINGREF "):
            _, name = command.split()
            return field_runs(f" REF {name} \\h \\w \\t ", properties=properties)
        if command.startswith("ANCHOR "):
            return bookmark(command[7:])
        if command.startswith("EQNO "):
            tab = ET.Element(q("w:r"))
            ET.SubElement(tab, q("w:tab"))
            return [tab, text_run("("), *numbered_bookmark("Equation", command[5:]), text_run(")")]
        if command.startswith("CAPTIONNO "):
            _, kind, name = command.split()
            return [text_run(kind + " "), *numbered_bookmark(kind, name)]
        if command.startswith("CAPTION "):
            kind, rest = command[8:].split(" ", 1)
            name, label = rest.split("|", 1)
            return [text_run(kind + " "), *numbered_bookmark(kind, name), text_run(" " + label)]
        from . import zotero
        if command == "BIBLIOGRAPHY" or command.startswith("BIBLIOGRAPHY "):
            citations_used = True
            uris = [(citation_items or {})[key]["uri"] for key in _marker_keys(command)]
            return field_runs(zotero.make_bibliography_field(uncited_uris=uris))
        if command.startswith("CITE "):
            citations_used = True
            keys = _marker_keys(command)
            items = [(citation_items or {})[k] for k in keys]
            return field_runs(zotero.make_citation_field(items), properties=properties)
        raise ValueError(f"Unsupported marker {command}")

    for part_name, root, paragraphs in documents:
        changed = False
        for kind, paragraph in paragraphs:
            combined = "".join(paragraph.itertext())
            if "[[" not in combined:
                continue
            for run in list(paragraph.findall(q("w:r"))):
                texts = run.findall(q("w:t"))
                if len(texts) != 1:
                    continue
                text = texts[0].text or ""
                matches = list(MARKER.finditer(text))
                if not matches:
                    continue
                if len(run) != len(texts) + int(run.find(q("w:rPr")) is not None):
                    raise ValueError("Marker run also contains an object; refusing to flatten it")
                nodes, cursor = [], 0
                props = run.find(q("w:rPr"))
                for match in matches:
                    if kind is not None and not match[1].startswith(("CITE ", "REF ", "PAGEREF ", "HEADINGREF ")):
                        raise ValueError(f"{match[0]} is not supported in a {kind}; place it in the main body")
                    if match[1] in _CAPTION_LISTS and not _standalone_caption_list(paragraph, combined, match[0]):
                        raise ValueError(f"{match[0]} must be the only content in its paragraph")
                    if match.start() > cursor:
                        nodes.append(text_run(text[cursor:match.start()], props))
                    nodes.extend(replacement(match[1], props))
                    cursor = match.end()
                    if match[1].startswith(("CAPTION ", "CAPTIONNO ", "EQNO ")):
                        pp = paragraph.find(q("w:pPr"))
                        if pp is None:
                            pp = ET.Element(q("w:pPr")); paragraph.insert(0, pp)
                        style = pp.find(q("w:pStyle"))
                        if style is None:
                            style = ET.Element(q("w:pStyle"))
                            pp.insert(0, style)
                        style.set(q("w:val"), equation_style if match[1].startswith("EQNO ") else caption_style)
                if cursor < len(text):
                    nodes.append(text_run(text[cursor:], props))
                index = paragraph.index(run)
                paragraph.remove(run)
                for node in nodes:
                    paragraph.insert(index, node); index += 1
                changed = True
            remaining = "".join(paragraph.itertext())
            if MARKER.search(remaining):
                raise ValueError("Marker split across runs or nested in a hyperlink; normalize this marker before building")
        if changed:
            parts[part_name] = ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    if "word/settings.xml" in parts:
        settings = xml(parts["word/settings.xml"])
        update = settings.find(q("w:updateFields"))
        if update is None:
            update = ET.SubElement(settings, q("w:updateFields"))
        update.set(q("w:val"), "true")
        parts["word/settings.xml"] = ET.tostring(settings, encoding="UTF-8", xml_declaration=True, standalone=True)
    write_package(parts, output)
    return {"citations_used": citations_used, "chapter_bookmarks": chapter_bookmarks}
