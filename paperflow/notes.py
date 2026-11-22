"""Find native note stories and apply template-owned note markers."""

from .character_style import _TYPOGRAPHY
from .package import q, xml
from .template import _internal_target, _part_relationships


_MAIN = "word/document.xml"
KINDS = ("footnote", "endnote")
_STYLE_NAMES = {
    "footnote": ({"脚注文字", "脚注文本", "註腳文字", "註腳本文"}, {"脚注引用", "脚注引用标记", "註腳參照"}),
    "endnote": ({"尾注文字", "尾注文本", "章節附註文字", "章節附註本文"}, {"尾注引用", "尾注引用标记", "章節附註參照"}),
}
_MARKERS = {q("w:" + kind + suffix): kind for kind in KINDS for suffix in ("Ref", "Reference")}
_MARKER_FORMAT = _TYPOGRAPHY | {q("w:" + tag) for tag in ("b", "bCs", "i", "iCs", "u", "vertAlign")}


def note_parts(parts):
    """Resolve the native note parts via the main document relationships."""
    output = {}
    for relation in _part_relationships(parts, _MAIN).values():
        suffix = relation.get("Type", "").rsplit("/", 1)[-1]
        if suffix not in {"footnotes", "endnotes"}:
            continue
        kind = suffix[:-1]
        if kind in output or relation.get("TargetMode", "Internal") != "Internal":
            raise ValueError(f"Invalid Word relationship for {suffix}")
        path = _internal_target(_MAIN, relation.get("Target"))
        if path not in parts:
            raise ValueError(f"Missing Word {suffix} part: {path}")
        output[kind] = (path, xml(parts[path]))
    return output


def note_ids(document, kind):
    body = document.find(q("w:body"))
    return [mark.get(q("w:id")) for mark in body.iter(q("w:" + kind + "Reference"))] if body is not None else []


def normal_notes(root, kind):
    return [node for node in root.findall(q("w:" + kind)) if node.get(q("w:type"), "normal") == "normal"]


def _style_id(styles, canonical, labels, style_type):
    candidates = []
    for sid, style in styles.items():
        if not sid or style.get(q("w:type")) != style_type:
            continue
        if sid.casefold() == canonical:
            return sid
        name = style.find(q("w:name"))
        if name is not None and "".join(name.get(q("w:val"), "").casefold().split()) in labels | {canonical}:
            candidates.append(sid)
    return candidates[0] if len(candidates) == 1 else None


def note_style_ids(styles):
    text, references = {}, {}
    for kind, (text_names, reference_names) in _STYLE_NAMES.items():
        if sid := _style_id(styles, kind + "text", text_names, "paragraph"):
            text[kind] = sid
        if sid := _style_id(styles, kind + "reference", reference_names, "character"):
            references[kind] = sid
    return text, references


def normalize_note_marker(run, references):
    """Return the kind of a native marker without changing its note link."""
    kind = next((_MARKERS[node.tag] for node in run if node.tag in _MARKERS), None)
    if kind is None or kind not in references:
        return kind
    props = run.find(q("w:rPr"))
    if props is None:
        props = run.makeelement(q("w:rPr"))
        run.insert(0, props)
    for node in list(props):
        if node.tag in _MARKER_FORMAT:
            props.remove(node)
    style = props.find(q("w:rStyle"))
    if style is None:
        style = props.makeelement(q("w:rStyle"))
        props.insert(0, style)
    style.set(q("w:val"), references[kind])
    return kind
