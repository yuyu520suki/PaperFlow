"""Read labeled abstract and keywords from the main Word story."""
import re
from dataclasses import dataclass

from lxml import etree

from .package import NS, q, xml

_LABEL = re.compile(
    r"^(abstract|key\s*words)(?:(?:[^\S\n]*[:：.\u2013\u2014\n]| +-(?= |$))\s*(.*))?$",
    re.I | re.S,
)
_SECTION = re.compile(
    r"(?:\d+(?:\.\d+)*[.)]?\s+)?"
    r"(?:introduction|background|methods|materials and methods|results(?: and discussion)?|"
    r"discussion|conclusions?|references|acknowledg(?:e)?ments|appendix(?: [a-z])?)",
    re.I,
)
_INTRODUCTION = re.compile(r"(?:\d+[.)]?\s+)?introduction", re.I)
_TOC_STYLE = re.compile(r"toc\s*(?:\d+|heading)", re.I)
_HEADING_LEVELS = frozenset("012345678")
_WORD = re.compile(r"\b\w+(?:[-\u2010\u2011'\u2019]\w+)*\b")
_BLOCK_WRAPPERS = {q("w:" + name) for name in ("sdt", "sdtContent", "customXml", "ins", "moveTo")}
_EXCLUDED = {q("w:" + name) for name in ("del", "moveFrom", "txbxContent", "drawing", "pict", "object", "pPr", "rPr")}


@dataclass(frozen=True)
class _Paragraph:
    index: int
    text: str
    heading: bool
    section_end: bool
    label: str | None = None
    inline_value: str = ""


def _blocks(container):
    for node in container:
        if node.tag in {q("w:p"), q("w:tbl")}:
            yield node
        elif node.tag in _BLOCK_WRAPPERS:
            yield from _blocks(node)


def _enabled(node):
    return node is not None and node.get(q("w:val"), "1").lower() not in {"0", "false", "off"}


def _text(paragraph, field_stack, paragraph_hidden, style_hides, default_character_style):
    """Keep ordinary text and rendered field results, skipping TOC fields."""
    chunks = []
    events = etree.iterwalk(paragraph, events=("start",))
    hidden_run = None
    for _, node in events:
        tag = node.tag
        if tag in _EXCLUDED:
            events.skip_subtree()
        elif tag == q("w:r"):
            direct = node.find("w:rPr/w:vanish", NS)
            character_style = node.find("w:rPr/w:rStyle", NS)
            character_id = (character_style.get(q("w:val"), "")
                            if character_style is not None else default_character_style)
            hidden = _enabled(direct) if direct is not None else paragraph_hidden ^ style_hides.get(character_id, False)
            hidden_run = node if hidden else None
        elif tag == q("w:fldSimple") and re.match(r"\s*TOC\b", node.get(q("w:instr"), ""), re.I):
            events.skip_subtree()
        elif tag == q("w:fldChar"):
            kind = node.get(q("w:fldCharType"))
            if kind == "begin":
                field_stack.append(["", False])
            elif field_stack and kind == "separate":
                field_stack[-1][1] = True
            elif field_stack and kind == "end":
                field_stack.pop()
        elif tag == q("w:instrText") and field_stack and not field_stack[-1][1]:
            field_stack[-1][0] += node.text or ""
        elif any(not result or re.match(r"\s*TOC\b", instruction, re.I) for instruction, result in field_stack):
            continue
        elif hidden_run is not None and node.getparent() is hidden_run:
            continue
        elif tag == q("w:t"):
            chunks.append(node.text or "")
        elif tag == q("w:tab"):
            chunks.append(" ")
        elif tag in {q("w:br"), q("w:cr")}:
            chunks.append("\n")
        elif tag == q("w:noBreakHyphen"):
            chunks.append("\u2011")
    return "\n".join(" ".join(line.split()) for line in "".join(chunks).split("\n")).strip()


def _style_chain(style_id, styles):
    visited = set()
    while style_id in styles and style_id not in visited:
        style = styles[style_id]
        visited.add(style_id)
        yield style
        parent = style.find(q("w:basedOn"))
        style_id = parent.get(q("w:val"), "") if parent is not None else ""


def _style_details(style_id, styles):
    style = styles.get(style_id)
    name = style.find(q("w:name")) if style is not None else None
    label = name.get(q("w:val"), style_id) if name is not None else style_id
    for style in _style_chain(style_id, styles):
        outline = style.find("w:pPr/w:outlineLvl", NS)
        if outline is not None:
            return label, outline.get(q("w:val"), "9") in _HEADING_LEVELS
    return label, bool(re.fullmatch(r"heading\s*[1-9]", label, re.I))


def _paragraphs(document, parts):
    styles, defaults = {}, {}
    default_hidden = False
    if "word/styles.xml" in parts:
        root = xml(parts["word/styles.xml"])
        styles = {s.get(q("w:styleId")): s for s in root.findall(q("w:style"))}
        defaults = {style.get(q("w:type")): key for key, style in styles.items()
                    if style.get(q("w:default"), "").lower() in {"1", "true", "on"}}
        default_hidden = _enabled(root.find("w:docDefaults/w:rPrDefault/w:rPr/w:vanish", NS))
    # In a style, vanish toggles visibility. Direct run formatting sets it.
    style_hides = {key: bool(sum(_enabled(style.find("w:rPr/w:vanish", NS))
                                for style in _style_chain(key, styles)) % 2)
                   for key in styles}
    default_character_style = defaults.get("character", "")
    field_stack = []
    index = 0
    for node in _blocks(document.find(q("w:body"))):
        if node.tag == q("w:tbl"):
            yield _Paragraph(0, "", False, False)
            continue
        index += 1
        style_node = node.find("w:pPr/w:pStyle", NS)
        style_id = style_node.get(q("w:val"), "") if style_node is not None else defaults.get("paragraph", "")
        hidden = default_hidden ^ style_hides.get(style_id, False)
        text = _text(node, field_stack, hidden, style_hides, default_character_style)
        style, heading = _style_details(style_id, styles)
        if _TOC_STYLE.fullmatch(style_id) or _TOC_STYLE.fullmatch(style):
            text = ""
        outline = node.find("w:pPr/w:outlineLvl", NS)
        if outline is not None:
            heading = outline.get(q("w:val"), "9") in _HEADING_LEVELS
        label = _LABEL.fullmatch(text)
        yield _Paragraph(
            index, text, heading or bool(_SECTION.fullmatch(text)),
            node.find("w:pPr/w:sectPr", NS) is not None,
            "abstract" if label and label[1].casefold() == "abstract" else "keywords" if label else None,
            (label[2] or "").strip() if label else "",
        )


def _find_block(paragraphs, name):
    candidates = [i for i, p in enumerate(paragraphs) if p.label == name]
    if not candidates:
        return None, [], "No recognizable label in the Word body"
    if len(candidates) != 1:
        return None, [paragraphs[i].index for i in candidates], "Multiple labels in the Word body"
    start = candidates[0]
    label = paragraphs[start]
    indexes = [label.index]
    values = [label.inline_value] if label.inline_value else []
    if (name == "keywords" and values) or label.section_end:
        return " ".join(values), indexes, None
    following = paragraphs[start + 1:]
    for offset, paragraph in enumerate(following):
        if paragraph.label:
            break
        if paragraph.heading and paragraph.text:
            if name == "abstract" and not _INTRODUCTION.fullmatch(paragraph.text):
                for later in following[offset:]:
                    if later.label == "keywords":
                        return None, indexes, "A heading interrupts the abstract before Keywords"
                    if later.section_end or _INTRODUCTION.fullmatch(later.text):
                        break
            break
        if not paragraph.index:
            return None, indexes, "A table interrupts the labeled block"
        if paragraph.text:
            values.append(paragraph.text)
            indexes.append(paragraph.index)
        if paragraph.section_end:
            break
    else:
        if len(values) > 1:
            return None, indexes, "No recognizable end of the labeled block"
    if name == "keywords" and len(values) > 1:
        return None, indexes, "The keyword list spans multiple paragraphs"
    return " ".join(values), indexes, None


def manuscript_frontmatter(document, parts):
    paragraphs = list(_paragraphs(document, parts))
    return {name: _find_block(paragraphs, name) for name in ("abstract", "keywords")}


def abstract_word_count(text):
    return len(_WORD.findall(text.replace("\u00ad", "")))


def split_keywords(text):
    delimiter = r"[;；]" if re.search(r"[;；]", text) else r"[,，]"
    return [" ".join(term.split()) for term in re.split(delimiter, text)] if text else []
