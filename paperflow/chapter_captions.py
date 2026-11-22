"""Validate heading context and refreshed Word cross-references."""

from collections import Counter, defaultdict
from dataclasses import dataclass, field
import re

from .package import NS, q, read_package, without_bidi_controls, xml
from .template import _main_story_events


_TOC_FIELD = re.compile(r"^\s*TOC(?:\s|$)", re.I)
_REF_FIELD = re.compile(r'^\s*(REF|PAGEREF)\s+(?:"([^"]+)"|([^\s\\"]+))(?=\s|\\|$)', re.I)
_FULL_CONTEXT = re.compile(r"(?:^|\s)\\w(?=\s|$)", re.I)
_INVALID_HEADING_NUMBER = re.compile(r"PAPERFLOW_PENDING|[!！?？\ufffc]", re.I)
_INDEX_STYLE = re.compile(r"toc(?:[1-9]|heading)|tableoffigures", re.I)
_CUSTOM_STYLE = {"1", "on", "true"}
_DISPLAY_NODES = {
    q("w:tab"): "\t", q("w:br"): "\n", q("w:cr"): "\n",
    q("w:noBreakHyphen"): "\u2011", q("w:softHyphen"): "\u00ad",
    q("w:sym"): "\ufffc", q("w:drawing"): "\ufffc", q("w:pict"): "\ufffc",
    q("w:object"): "\ufffc", q("m:oMath"): "\ufffc",
}


@dataclass
class _Paragraph:
    number: int
    chapter_heading: bool
    index_style: bool
    has_text: bool = False


@dataclass
class _Field:
    instruction: str = ""
    in_result: bool = False
    reference: tuple[str, str] | None = None
    result: list[str] = field(default_factory=list)

    def begin_result(self):
        self.in_result = True
        self.reference = _reference_target(self.instruction)


def _style_value(node):
    return node.get(q("w:val"), "") if node is not None else ""


def _style_key(value):
    return "".join(value.casefold().split())


def _style_sets(parts, level):
    data = parts.get("word/styles.xml")
    if data is None:
        return set(), set()
    headings, indexes = set(), set()
    native_names = {f"{prefix}{level}" for prefix in ("heading", "标题", "標題")} if level is not None else set()
    for style in xml(data).findall(q("w:style")):
        if style.get(q("w:type"), "paragraph") != "paragraph":
            continue
        style_id = style.get(q("w:styleId"), "")
        name = _style_value(style.find(q("w:name")))
        keys = {_style_key(name), _style_key(style_id)}
        if style.get(q("w:customStyle"), "").casefold() not in _CUSTOM_STYLE and keys & native_names:
            headings.add(style_id)
        if any(_INDEX_STYLE.fullmatch(key) for key in keys):
            indexes.add(style_id)
    return headings, indexes


def _is_index_style(style, index_styles):
    return style in index_styles or bool(_INDEX_STYLE.fullmatch(_style_key(style)))


def _reference_target(instruction):
    match = _REF_FIELD.match(instruction)
    if match is None:
        return None
    kind = match[1].upper()
    if kind == "REF" and _FULL_CONTEXT.search(instruction[match.end():]):
        kind = "HEADINGREF"
    return kind, (match[2] or match[3]).casefold()


def _display_piece(node):
    return node.text or "" if node.tag == q("w:t") else _DISPLAY_NODES.get(node.tag, "")


def _story_nodes(body):
    """Yield story events, TOC context and completed cross-reference fields."""
    fields = []
    for event, node in _main_story_events(body):
        completed = None
        if node.tag == q("w:fldSimple"):
            if event == "start":
                fields.append(_Field(node.get(q("w:instr"), "")))
                fields[-1].begin_result()
            elif fields:
                completed = fields.pop()
        elif event == "start" and node.tag == q("w:fldChar"):
            kind = node.get(q("w:fldCharType"))
            if kind == "begin":
                fields.append(_Field())
            elif kind == "separate" and fields:
                fields[-1].begin_result()
            elif kind == "end" and fields:
                completed = fields.pop()
        elif event == "start" and fields:
            if node.tag == q("w:instrText") and not fields[-1].in_result:
                fields[-1].instruction += node.text or ""
            elif piece := _display_piece(node):
                for active in fields:
                    if active.reference is not None:
                        active.result.append(piece)
        in_toc = any(_TOC_FIELD.match(active.instruction) for active in fields)
        reference = None
        if completed is not None and not in_toc and completed.reference is not None:
            reference = (*completed.reference, "".join(completed.result))
        yield event, node, in_toc, reference


def validate_chapter_captions(built_path, level, bookmark_names):
    """Require a preceding built-in heading for each generated caption bookmark.

    This checks body order before Word updates STYLEREF. It does not evaluate
    the heading's list label or either caption field's result.
    """
    if type(level) is not int or not 1 <= level <= 9:
        raise ValueError("Chapter captions require a heading level from 1 to 9")
    names = {}
    for name in bookmark_names:
        if not isinstance(name, str) or not name:
            raise ValueError("Chapter caption bookmark names must be nonempty strings")
        if name.casefold() in names:
            raise ValueError(f"Duplicate chapter caption bookmark {name!r}")
        names[name.casefold()] = name
    if not names:
        return {"heading_level": level, "bookmarks_checked": 0, "builtin_heading_paragraphs": 0, "bookmarks": []}

    parts = read_package(built_path)
    headings, index_styles = _style_sets(parts, level)
    body = xml(parts["word/document.xml"]).find(q("w:body"))
    if body is None:
        raise ValueError("Chapter captions require a main Word document body")
    paragraphs = []
    matches, ends = defaultdict(list), defaultdict(list)
    paragraph_number = heading_count = 0
    last_heading = None
    for position, (event, node, in_toc, _) in enumerate(_story_nodes(body)):
        tag = node.tag
        if tag == q("w:p"):
            if event == "start":
                paragraph_number += 1
                style = _style_value(node.find("w:pPr/w:pStyle", NS))
                chapter_heading = style in headings and next(node.iterancestors(q("w:tbl")), None) is None
                paragraphs.append(_Paragraph(paragraph_number, chapter_heading, _is_index_style(style, index_styles)))
            else:
                paragraph = paragraphs.pop()
                if paragraph.chapter_heading and paragraph.has_text:
                    heading_count += 1
                    last_heading = paragraph.number
            continue
        if event != "start" or in_toc:
            continue
        paragraph = paragraphs[-1] if paragraphs else None
        if paragraph is not None and paragraph.index_style:
            continue
        if tag == q("w:t") and paragraph is not None and (node.text or "").strip():
            paragraph.has_text = True
        elif tag == q("w:bookmarkStart"):
            name = node.get(q("w:name"), "").casefold()
            if name in names:
                matches[name].append((node.get(q("w:id")), paragraph, position, last_heading))
        elif tag == q("w:bookmarkEnd"):
            ends[node.get(q("w:id"))].append((paragraph, position))

    checked = []
    for key, name in names.items():
        found = matches[key]
        if len(found) != 1:
            raise ValueError(f"Chapter caption bookmark {name!r} must occur once in the main story; found {len(found)}")
        identifier, paragraph, start, preceding = found[0]
        closing = ends[identifier]
        if (identifier is None or paragraph is None or len(closing) != 1
                or closing[0][0] is not paragraph or closing[0][1] <= start):
            raise ValueError(f"Chapter caption bookmark {name!r} must start and end in one body paragraph")
        if preceding is None:
            raise ValueError(f"Chapter caption bookmark {name!r} has no preceding built-in Heading {level} in the document body")
        checked.append({"bookmark": name, "body_paragraph": paragraph.number, "preceding_heading_paragraph": preceding})
    return {"heading_level": level, "bookmarks_checked": len(checked), "builtin_heading_paragraphs": heading_count,
            "bookmarks": checked}


def _saved_reference_values(path, names):
    parts = read_package(path)
    body = xml(parts["word/document.xml"]).find(q("w:body"))
    if body is None:
        raise ValueError("Cross-reference validation requires a main Word document body")
    _, index_styles = _style_sets(parts, None)
    bookmarks, references = defaultdict(list), defaultdict(list)
    active = {}
    paragraph = None
    in_index_style = False
    for event, node, in_toc, reference in _story_nodes(body):
        tag = node.tag
        if tag == q("w:p"):
            paragraph = node if event == "start" else None
            style = _style_value(node.find("w:pPr/w:pStyle", NS))
            in_index_style = paragraph is not None and _is_index_style(style, index_styles)
        if in_toc or in_index_style:
            continue
        if reference is not None and paragraph is not None and reference[1] in names:
            kind, name, result = reference
            references[kind, name].append(without_bidi_controls(result).strip())
        if event != "start":
            continue
        if tag == q("w:bookmarkStart"):
            name = node.get(q("w:name"), "").casefold()
            if name in names:
                identifier = node.get(q("w:id"))
                if identifier is None or identifier in active or paragraph is None:
                    raise ValueError(f"Number bookmark {name!r} has an invalid body range")
                pieces = []
                bookmarks[name].append(pieces)
                active[identifier] = (name, paragraph, pieces)
        elif tag == q("w:bookmarkEnd"):
            identifier = node.get(q("w:id"))
            if identifier in active:
                name, start_paragraph, _ = active.pop(identifier)
                if paragraph is not start_paragraph:
                    raise ValueError(f"Number bookmark {name!r} crosses body paragraphs")
        elif active and (piece := _display_piece(node)):
            for _, _, pieces in active.values():
                pieces.append(piece)
    if active:
        raise ValueError(f"Unclosed number bookmarks in Word body: {list(active)!r}")
    numbers = {name: [without_bidi_controls("".join(pieces)).strip() for pieces in bookmarks[name]] for name in names}
    return numbers, references


def verify_generated_reference_values(path, targets, references):
    """Check that Word saved the generated anchors and refreshed references."""
    planned = {target["bookmark"].casefold(): target for target in targets}
    if len(planned) != len(targets):
        raise ValueError("Generated references must point to unique planned bookmarks")
    number_kinds = {name: "HEADINGREF" if target["kind"] == "Heading" else "REF"
                    for name, target in planned.items()}
    counts = Counter()
    for ref in references:
        name = ref["bookmark"].casefold()
        if name not in planned:
            raise ValueError("Generated references must point to unique planned bookmarks")
        kind = "PAGEREF" if ref.get("command") == "pageref" else number_kinds[name]
        counts[kind, name] += 1

    numbers, results = _saved_reference_values(path, planned)
    checked = []
    for name, target in planned.items():
        found = numbers[name]
        kind = number_kinds[name]
        number_count, number_results = counts[kind, name], results[kind, name]
        if kind == "HEADINGREF":
            # A heading bookmark surrounds a position in the paragraph, not its automatic number.
            number = number_results[0] if number_results else None
            valid_number = (not results["REF", name] and len(number_results) == number_count
                            and (number_count == 0 or bool(number) and not _INVALID_HEADING_NUMBER.search(number)
                                 and number_results == [number] * number_count))
        else:
            number = found[0] if len(found) == 1 else None
            valid_number = (bool(number) and not results["HEADINGREF", name]
                            and re.fullmatch(r"(?:[^\s]+[-.:–—])?[0-9]+", number)
                            and number_results == [number] * number_count)
        if len(found) != 1 or not valid_number:
            raise ValueError(f"Word did not refresh the expected bookmark and {number_count} number references for "
                             f"{target['bookmark']!r}. Observed {found!r} and {number_results!r}")
        pages = results["PAGEREF", name]
        page_count = counts["PAGEREF", name]
        if (len(pages) != page_count or len(set(pages)) > 1 or (pages and (
            not re.fullmatch(r"[^\s!！?？]+", pages[0]) or pages[0] == "PAPERFLOW_PENDING"
        ))):
            raise ValueError(f"Word did not refresh {page_count} page references for {target['bookmark']!r}: {pages!r}")
        checked.append({**target, "number": number, "references": number_count,
                        "page_reference_count": page_count, "page_reference_result": pages[0] if pages else None})
    return {"bookmarks_checked": len(checked), "references_checked": sum(counts.values()),
            "page_references_checked": sum(count for (kind, _), count in counts.items() if kind == "PAGEREF"),
            "heading_references_checked": sum(count for (kind, _), count in counts.items() if kind == "HEADINGREF"),
            "targets": checked}


def validate_chapter_migration_values(path, expected):
    """Compare exact numbers cached in Word's caption bookmarks and body REF fields."""
    planned = {}
    for item in expected:
        name, number, count = item["bookmark"], item["number"], item["references"]
        if not isinstance(name, str) or not name or not isinstance(number, str) or not number:
            raise ValueError("Chapter migration needs nonempty bookmark names and expected numbers")
        if type(count) is not int or count < 0:
            raise ValueError(f"Chapter migration reference count for {name!r} must be a nonnegative integer")
        if name.casefold() in planned:
            raise ValueError(f"Duplicate chapter migration bookmark {name!r}")
        planned[name.casefold()] = item
    if not planned:
        return {"bookmarks_checked": 0, "references_checked": 0, "bookmarks": []}

    numbers, references = _saved_reference_values(path, planned)
    checked, problems = [], []
    for key, item in planned.items():
        name, wanted, count = item["bookmark"], item["number"], item["references"]
        found = numbers[key]
        ref_values = references["REF", key] + references["HEADINGREF", key]
        if found != [wanted]:
            problems.append(f"{name!r} expected one bookmark displaying {wanted!r}, observed {found!r}")
        if ref_values != [wanted] * count:
            problems.append(f"{name!r} expected {count} REF fields displaying {wanted!r}, observed {ref_values!r}")
        checked.append({"bookmark": name, "number": wanted, "references": count})
    if problems:
        raise ValueError("Word chapter migration numbers differ from the reviewed plan. "
                         "Check the template chapter list's starting number and restarts, "
                         "and verify each source caption and reference. " + ". ".join(problems))
    return {"bookmarks_checked": len(checked), "references_checked": sum(item["references"] for item in checked),
            "bookmarks": checked}
