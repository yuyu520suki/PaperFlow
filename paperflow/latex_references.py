"""Convert LaTeX reference semantics from Pandoc's raw-TeX-aware AST."""

from collections import Counter
from copy import deepcopy
from dataclasses import dataclass
from hashlib import sha256
import re


_LABEL = re.compile(r"\\label\s*\{([^{}]+)\}")
_DISPLAY_ENV = re.compile(r"\s*\\begin\{([^{}]+)\}(.*?)\\end\{\1\}\s*", re.S)
_NUMBERED_ENVS = {"equation", "align", "alignat", "flalign", "gather", "multline", "eqnarray", "subequations"}
_CUSTOM_EQUATION = re.compile(r"\\(?:tag|notag|nonumber)(?![A-Za-z])")
_REF = re.compile(r"\\([A-Za-z]+)(\*)?\s*\{([^{}]+)\}\s*")
_REF_COMMAND = re.compile(r"\s*\\([A-Za-z]+)")
_REFERENCE_COMMANDS = {"ref", "eqref", "autoref", "pageref", "cref", "Cref", "cpageref", "Cpageref",
                       "vref", "Vref", "nameref", "Nameref", "fref", "Fref", "subref", "hyperref"}
_HEADING_NAMES = {1: "Section", 2: "Subsection", 3: "Subsubsection"}
_INDEX_COMMANDS = {"tableofcontents": ("TOC", "Contents"), "listoffigures": ("LOF", "List of Figures"),
                   "listoftables": ("LOT", "List of Tables")}
_BARE_COMMAND = re.compile(r"\\[A-Za-z]+(?:\s*\{\s*\})?")
_ID_INDEX = {"Div": 0, "Figure": 0, "Table": 0, "Header": 1, "Span": 0}
_BLOCKED = {"Header": "heading", "Link": "hyperlink", "Image": "image description", "Note": "note"}
_INVISIBLE_SPACE = {"Space", "SoftBreak"}


@dataclass
class _Target:
    kind: str
    labels: tuple[str, ...]
    node: dict
    position: int
    content: list | str | None = None
    reason: str | None = None


def _walk(value):
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from _walk(child)
    elif isinstance(value, list):
        for child in value:
            yield from _walk(child)


def _identifier(node):
    index = _ID_INDEX.get(node.get("t"))
    return node["c"][index][0] if index is not None else ""


def _raw_tex(node):
    if node.get("t") in {"RawInline", "RawBlock"} and node["c"][0] in {"latex", "tex"}:
        return node["c"][1]
    return ""


def _labels(node):
    identifier = _identifier(node)
    if identifier:
        return [identifier]
    if node.get("t") == "Math":
        return _LABEL.findall(node["c"][1])
    return _LABEL.findall(_raw_tex(node))


def _heading_reference_labels(blocks):
    return {match[3].strip() for node in _walk(blocks)
            if (match := _REF.fullmatch(_raw_tex(node).strip()))
            and match[1] in {"ref", "autoref", "pageref"} and not match[2]}


def needs_heading_labels(ast):
    """Determine whether an AST with auto IDs needs a check for explicit labels."""
    blocks = ast["blocks"]
    headings = {_identifier(block) for block in blocks if block.get("t") == "Header"}
    return bool(headings & _heading_reference_labels(blocks))


def _simple_image(figure):
    blocks = [block for block in figure["c"][2] if _raw_tex(block).strip() != r"\centering"]
    if len(blocks) == 1 and blocks[0].get("t") in {"Plain", "Para"}:
        inlines = blocks[0]["c"]
        if len(inlines) == 1 and inlines[0].get("t") == "Image":
            return inlines[0]
    return None


def _float_target(node, position, wrapper=None, *, supported=True):
    label_ids = tuple(dict.fromkeys(identifier for item in (wrapper, node) if item is not None
                                    if (identifier := _identifier(item))))
    _, long_caption = node["c"][1]
    if not long_caption and not label_ids:
        return None
    first = long_caption[0] if long_caption else None
    target = _Target(node["t"], label_ids, node, position)
    if not supported:
        target.reason = "Nested float needs an explicit Word caption placement"
    elif len(label_ids) > 1:
        target.reason = "The table and its wrapper name different labels"
    elif first is None or first.get("t") not in {"Plain", "Para"} or not first["c"]:
        target.reason = "Cannot place a number in the first caption paragraph"
    elif target.kind == "Figure" and _simple_image(node) is None:
        target.reason = "A figure with more than one image or additional content needs review"
    else:
        target.content = first["c"]
    return target


def _equation_target(node, position, *, supported=True):
    math_type, text = node["c"]
    labels = tuple(_LABEL.findall(text))
    environment = _DISPLAY_ENV.fullmatch(text) if math_type.get("t") == "DisplayMath" else None
    env_name, body = environment.groups() if environment else (None, None)
    if not labels and env_name not in _NUMBERED_ENVS:
        return None
    target = _Target("Equation", labels, node, position)
    if not supported:
        target.reason = "Nested equation needs an explicit Word number placement"
    elif env_name != "equation" or len(labels) > 1:
        target.reason = "Only a simple numbered equation with at most one label can receive a Word number"
    elif _CUSTOM_EQUATION.search(body):
        target.reason = "An equation with custom numbering or suppressed numbering needs review"
    elif not (expression := _LABEL.sub("", body).strip()):
        target.reason = "Cannot number an empty equation"
    else:
        target.content = expression
    return target


def _target_candidates(blocks, heading_labels):
    for position, block in enumerate(blocks, 1):
        top_float = block if block.get("t") in {"Figure", "Table"} else None
        wrapper = None
        if block.get("t") == "Div":
            children = block["c"][1]
            if len(children) == 1 and children[0].get("t") == "Table":
                top_float = children[0]
                wrapper = block
        direct_inlines = {id(node) for node in block["c"]} if block.get("t") in {"Para", "Plain"} else set()
        for node in _walk(block):
            kind = node.get("t")
            if kind in {"Figure", "Table"}:
                target = _float_target(node, position, wrapper if node is top_float else None,
                                       supported=node is top_float)
            elif kind == "Math":
                target = _equation_target(node, position, supported=id(node) in direct_inlines)
            elif node is block and kind == "Header" and (label := _identifier(node)) in heading_labels:
                target = _Target("Heading", (label,), node, position)
            else:
                continue
            if target is not None:
                yield target


def _marker(command):
    return {"t": "Str", "c": command}


def _replace_raw(node, text):
    marker = _marker(text)
    if node["t"] == "RawBlock":
        node.update(t="Para", c=[marker])
    else:
        node.update(marker)


def _index_command(node):
    raw = _raw_tex(node).strip()
    match = _REF_COMMAND.match(raw)
    return (match[1], raw) if match and match[1] in _INDEX_COMMANDS else None


def _index_review(node, command, raw, reason, report):
    report["issues"].append({"command": command, "source": raw, "reason": reason})
    safe_display = raw.replace("[", "［").replace("]", "］")
    _replace_raw(node, f"[LaTeX index needs review: {safe_display}]")


def _index_blocks(blocks):
    # A command on its own TeX line can still be inline in Pandoc's paragraph.
    for block in blocks:
        if block["t"] not in {"Para", "Plain"}:
            yield block
            continue
        inlines = block["c"]
        start = 0
        for index, node in enumerate(inlines):
            command = _index_command(node)
            if (command and _BARE_COMMAND.fullmatch(command[1])
                    and (index == 0 or inlines[index - 1]["t"] == "SoftBreak")
                    and (index == len(inlines) - 1 or inlines[index + 1]["t"] == "SoftBreak")):
                if prefix := _trim_inlines(inlines[start:index]):
                    yield {"t": block["t"], "c": prefix}
                yield {"t": "RawBlock", "c": node["c"]}
                start = index + 1
        if start == 0:
            yield block
        elif suffix := _trim_inlines(inlines[start:]):
            yield {"t": block["t"], "c": suffix}


def _prepare_indexes(blocks, metadata, report, template_indexes):
    seen = set()
    for block in _index_blocks(blocks):
        index = _index_command(block)
        if index and block["t"] == "RawBlock" and _BARE_COMMAND.fullmatch(index[1]):
            command, raw = index
            kind, title = _INDEX_COMMANDS[command]
            if kind in seen:
                _index_review(block, command, raw, "Duplicate LaTeX index command", report)
                yield block
                continue
            seen.add(kind)
            mode = "template" if kind in template_indexes else "source"
            report["indexes"].append({"kind": kind, "command": command, "mode": mode})
            if mode == "source":
                yield {"t": "Para", "c": [{"t": "Strong", "c": [_marker(title)]}]}
                yield {"t": "Para", "c": [_marker(f"[[{kind}]]")]}
            continue
        for node in _walk(block):
            if index := _index_command(node):
                _index_review(node, *index, "Place the unmodified index command in the main body", report)
        yield block
    for node in _walk(metadata):
        if index := _index_command(node):
            _index_review(node, *index, "An index command in metadata needs review", report)


def _numbered_blocks(blocks, equations):
    for block in blocks:
        if block.get("t") not in {"Para", "Plain"} or not any(id(node) in equations for node in block["c"]):
            yield block
            continue
        pending = []
        for node in block["c"]:
            number = equations.get(id(node))
            if number is None:
                pending.append(node)
                continue
            if pending := _trim_inlines(pending):
                yield {"t": block["t"], "c": pending}
            yield {"t": "Para", "c": [node, _marker(f"[[EQNO {number}]]")]}
            pending = []
        if pending := _trim_inlines(pending):
            yield {"t": block["t"], "c": pending}


def _trim_inlines(inlines):
    start, end = 0, len(inlines)
    while start < end and inlines[start].get("t") in _INVISIBLE_SPACE:
        start += 1
    while end > start and inlines[end - 1].get("t") in _INVISIBLE_SPACE:
        end -= 1
    return inlines[start:end]


def _replace_references(value, labels, targets, report, context=None):
    if isinstance(value, list):
        for child in value:
            _replace_references(child, labels, targets, report, context)
        return
    if not isinstance(value, dict):
        return
    kind = value.get("t")
    if kind in {"RawInline", "RawBlock"} and (match := _REF_COMMAND.match(_raw_tex(value))):
        command = match[1]
        if command not in _REFERENCE_COMMANDS:
            return
        raw = _raw_tex(value).strip()
        match = _REF.fullmatch(raw)
        label = match[3].strip() if match else None
        target = targets.get(label)
        if context:
            reason = f"A reference in {context} needs review"
        elif command not in {"ref", "eqref", "autoref", "pageref"} or match is None or (command == "pageref" and match[2]):
            reason = f"LaTeX \\{command} has no verified Word equivalent"
        elif target is None:
            reason = "Target is missing" if label not in labels else "Target has no unique supported Word bookmark"
        elif command == "eqref" and target["kind"] != "Equation":
            reason = "An equation reference points to a non-equation target"
        elif target["kind"] == "Heading" and command != "pageref" and match[2]:
            reason = "A non-hyperlinked LaTeX heading reference needs review"
        elif target["kind"] == "Heading" and command != "pageref" and target["numbering_review"]:
            reason = target["numbering_review"]
        elif target["kind"] == "Heading" and command == "autoref" and not target["autoref_name"]:
            reason = "The LaTeX heading type name needs review"
        else:
            field = "PAGEREF" if command == "pageref" else "HEADINGREF" if target["kind"] == "Heading" else "REF"
            marker = f"[[{field} {target['bookmark']}]]"
            if command == "eqref":
                marker = f"({marker})"
            elif command == "autoref":
                marker = f"{target.get('autoref_name', target['kind'])} {marker}"
            report["references"].append({"label": label, "command": command, "bookmark": target["bookmark"]})
            _replace_raw(value, marker)
            return
        report["issues"].append({"label": label, "command": command, "source": raw, "reason": reason})
        safe_display = raw.replace("[", "［").replace("]", "］")
        _replace_raw(value, f"[LaTeX reference needs review: {safe_display}]")
        return
    context = context or _BLOCKED.get(kind)
    if kind in {"Figure", "Table"}:
        for index, child in enumerate(value["c"]):
            _replace_references(child, labels, targets, report, context or ("caption" if index == 1 else None))
    else:
        for child in value.values():
            _replace_references(child, labels, targets, report, context)


def prepare_latex_references(ast, *, ambiguous_float_numbering=False, explicit_heading_ids=None, template_indexes=(),
                             numbered_heading_levels=(), ambiguous_heading_numbering=False, heading_autoref=True):
    """Emit editable Word number, reference and index markers in a copy of a LaTeX AST."""
    result = deepcopy(ast)
    blocks = result["blocks"]
    report = {"targets": [], "references": [], "indexes": [], "issues": []}
    labels = Counter(label for node in _walk(blocks) for label in _labels(node))
    heading_labels = _heading_reference_labels(blocks)
    if explicit_heading_ids is not None:
        heading_labels &= explicit_heading_ids
    candidates = list(_target_candidates(blocks, heading_labels))
    for target in candidates:
        if target.kind == "Table" and len(target.labels) == 1:
            # A Div and its direct Table may repeat the same label in the AST.
            wrapper = blocks[target.position - 1]
            if (wrapper.get("t") == "Div" and wrapper["c"][1] == [target.node]
                    and _identifier(wrapper) == _identifier(target.node) == target.labels[0]):
                labels[target.labels[0]] -= 1
        if target.reason is None and any(labels[label] > 1 for label in target.labels):
            target.reason = "The LaTeX label occurs more than once"
        if target.reason is None and ambiguous_float_numbering and target.kind in {"Figure", "Table"}:
            target.reason = "Starred captions or included files make automatic float numbering ambiguous"

    by_label, equations, figures, headings = {}, {}, [], {}
    for ordinal, target in enumerate(candidates, 1):
        if target.reason is not None:
            report["issues"].append({"labels": list(target.labels), "kind": target.kind, "block": target.position,
                                     "reason": target.reason})
            continue
        label = target.labels[0] if target.labels else None
        seed = f"label\0{label}" if label else f"unnamed\0{target.kind}\0{ordinal}"
        bookmark = "Latex" + sha256(seed.encode("utf-8")).hexdigest()[:35]
        entry = {"label": label, "kind": target.kind, "bookmark": bookmark}
        report["targets"].append(entry)
        if label is not None:
            by_label[label] = entry
        if target.kind == "Equation":
            # The DOCX writer would force DisplayMath and its number into
            # separate paragraphs. The equation paragraph supplies the layout.
            target.node["c"] = [{"t": "InlineMath"}, target.content]
            equations[id(target.node)] = bookmark
        elif target.kind == "Heading":
            level, (_, classes, _), _ = target.node["c"]
            entry["level"] = level
            entry["autoref_name"] = _HEADING_NAMES.get(level) if heading_autoref else None
            entry["numbering_review"] = None
            if "unnumbered" in classes:
                entry["numbering_review"] = "Numbered heading reference points to an unnumbered LaTeX heading"
            elif ambiguous_heading_numbering:
                entry["numbering_review"] = "Numbered heading reference with custom or included LaTeX numbering needs review"
            elif level not in numbered_heading_levels:
                entry["numbering_review"] = f"Numbered heading {level} has no verified Word numbering for this conversion"
            headings[bookmark] = target.node
        else:
            target.content[:0] = [_marker(f"[[CAPTIONNO {target.kind} {bookmark}]]"), {"t": "Space"}]
            if target.kind == "Figure":
                figures.append(target.node)

    result["blocks"] = list(_numbered_blocks(blocks, equations))
    _replace_references(result["blocks"], labels, by_label, report)
    _replace_references(result.get("meta", {}), labels, by_label, report, "metadata")
    used_headings = {ref["bookmark"] for ref in report["references"]} & headings.keys()
    for bookmark, heading in headings.items():
        if bookmark in used_headings:
            inlines = heading["c"][2]
            # A bookmark at paragraph start can attach to text inserted before it.
            position = next((index + 1 for index, node in enumerate(inlines)
                             if node["t"] not in _INVISIBLE_SPACE | {"RawInline", "Note"}), len(inlines))
            inlines.insert(position, _marker(f"[[ANCHOR {bookmark}]]"))
    report["targets"] = [entry for entry in report["targets"]
                         if entry["kind"] != "Heading" or entry["bookmark"] in used_headings]
    result["blocks"] = list(_prepare_indexes(result["blocks"], result.get("meta", {}), report, template_indexes))
    for figure in figures:
        # The DOCX writer renders a Figure as an image in a layout table.
        # Emit a standalone image and caption for Word's object placement.
        image = _simple_image(figure)
        attributes, (_, caption), _ = figure["c"]
        figure.update(t="Div", c=[attributes, [{"t": "Para", "c": [image]}, *caption]])
    return result, report
