"""Inspect Pandoc citations before linking them to live Zotero records."""

from __future__ import annotations

from collections.abc import Iterator, Mapping
from copy import deepcopy
import re
from typing import Any


_LATEX_CITE = re.compile(r"^\\(?:citep|parencite|cite)(?=\s*[\[{])")
_PAGE_LOCATOR = re.compile(
    r"(?:,\s*)?(?:(?:p{1,2}\.|pages?)\s*)?"
    r"(?P<locator>[0-9]+(?:\s*[-–]\s*[0-9]+)?)(?P<suffix>,\s*.+)?",
    re.IGNORECASE,
)
_UNSUPPORTED_LOCATOR = re.compile(
    r"^[\s,]*(?:p{1,2}\.|pages?\b|ch(?:ap(?:ter)?)?\.|chapters?\b|"
    r"secs?\.|sections?\b|vols?\.|volumes?\b|pars?\.|paras?\.|paragraphs?\b|"
    r"figs?\.|figures?\b|lines?\b|verses?\b|notes?\b|folios?\b|nos?\.)",
    re.IGNORECASE,
)
_BLOCKED_CONTEXTS = {"Header": "heading", "Image": "image description", "Link": "hyperlink"}


def _document_parts(ast: Mapping[str, Any]) -> tuple[list, Mapping]:
    if not isinstance(ast, Mapping) or not isinstance(ast.get("blocks"), list):
        raise ValueError("Pandoc did not return a document AST")
    meta = ast.get("meta", {})
    if not isinstance(meta, Mapping):
        raise ValueError("Pandoc returned invalid document metadata")
    return ast["blocks"], meta


def _bibliography_div(content: list) -> bool:
    if not content or not isinstance(content[0], list) or len(content[0]) != 3:
        return False
    identifier, classes, _ = content[0]
    return identifier == "refs" or (
        isinstance(classes, list) and any(name in classes for name in ("thebibliography", "references"))
    )


def _walk_citations(
    value: Any, blocked: str | None = None, in_note: bool = False,
) -> Iterator[tuple[dict, str | None, bool]]:
    if isinstance(value, list):
        for child in value:
            yield from _walk_citations(child, blocked, in_note)
        return
    if not isinstance(value, dict):
        return
    kind = value.get("t")
    if kind == "Cite":
        yield value, blocked, in_note
        return
    if kind is None:
        for child in value.values():
            yield from _walk_citations(child, blocked, in_note)
        return

    content = value.get("c")
    blocked = blocked or _BLOCKED_CONTEXTS.get(kind)
    if kind == "Div" and isinstance(content, list) and _bibliography_div(content):
        blocked = blocked or "existing bibliography"
    if kind == "Note":
        yield from _walk_citations(content, blocked or ("nested note" if in_note else None), True)
        return
    if kind in {"Table", "Figure"} and isinstance(content, list):
        for index, child in enumerate(content):
            context = blocked or ("object caption" if index == 1 else None)
            yield from _walk_citations(child, context, in_note)
        return
    yield from _walk_citations(content, blocked, in_note)


def _metadata_issues(meta: Mapping) -> list[dict]:
    issues = []
    for key, value in meta.items():
        if key == "nocite":
            reason = "Uncited bibliography entries need a separate review"
        elif next(_walk_citations(value), None) is not None:
            reason = "A citation in document metadata has no reviewed Word placement"
        else:
            continue
        issues.append({"index": None, "context": f"metadata.{key}", "reason": reason})
    return issues


def _plain_affix(nodes: Any) -> str:
    if not isinstance(nodes, list):
        raise ValueError("Pandoc returned invalid citation affix data")
    pieces = []
    for node in nodes:
        if not isinstance(node, Mapping):
            raise ValueError("Pandoc returned invalid citation affix data")
        kind = node.get("t")
        text = node.get("c")
        if kind in {"Space", "SoftBreak"}:
            pieces.append(" ")
        elif kind == "Str" and isinstance(text, str):
            text = text.replace("\u00a0", " ").replace("\u202f", " ")
            if text and not text.isprintable():
                raise ValueError("Invalid characters in a citation affix need review")
            pieces.append(text)
        else:
            raise ValueError("Formatted citation prefixes or suffixes need review")
    result = "".join(pieces).strip()
    if any(char in result for char in "<>{}\\"):
        raise ValueError("Markup or explicit locator syntax in citation affixes needs review")
    return result


def _literal_suffix(suffix: str) -> str:
    if re.search(r"\d|§", suffix) or _UNSUPPORTED_LOCATOR.match(suffix):
        raise ValueError("Ambiguous citation locator or numeric suffix needs review")
    return f" {suffix}" if suffix[0].isalnum() else suffix


def _citation_item(value: Any, latex: bool) -> dict:
    if not isinstance(value, Mapping):
        raise ValueError("Pandoc returned invalid citation item data")
    key = value.get("citationId")
    if not isinstance(key, str) or not key or key == "*" or not key.isprintable() or any(char.isspace() for char in key):
        raise ValueError("A citation requires an explicit BibTeX key")
    mode = value.get("citationMode")
    if not isinstance(mode, Mapping):
        raise ValueError("Pandoc returned invalid citation mode")
    mode = mode.get("t")
    if mode == "AuthorInText":
        raise ValueError("Author-in-text citations need review before conversion")
    if mode not in {"NormalCitation", "SuppressAuthor"} or (latex and mode == "SuppressAuthor"):
        raise ValueError("This citation mode has no verified Zotero equivalent")

    item = {"citekey": key}
    if mode == "SuppressAuthor":
        item["suppress-author"] = True
    prefix = _plain_affix(value.get("citationPrefix"))
    suffix = _plain_affix(value.get("citationSuffix"))
    if prefix:
        item["prefix"] = prefix
    locator = _PAGE_LOCATOR.fullmatch(suffix)
    if locator:
        item.update(locator=re.sub(r"\s+", "", locator["locator"]), label="page")
        suffix = locator["suffix"] or ""
    if suffix:
        item["suffix"] = _literal_suffix(suffix)
    return item


def _citation_items(node: dict, latex: bool) -> list[dict]:
    content = node.get("c")
    if not isinstance(content, list) or len(content) != 2 or not isinstance(content[0], list) or not content[0]:
        raise ValueError("Pandoc returned an empty or invalid Cite node")
    citations, fallback = content
    if latex:
        if not isinstance(fallback, list) or len(fallback) != 1 or not isinstance(fallback[0], Mapping):
            raise ValueError("Cannot identify the original LaTeX citation command")
        raw = fallback[0].get("c")
        if (fallback[0].get("t") != "RawInline" or not isinstance(raw, list) or len(raw) != 2
                or raw[0] != "latex" or not isinstance(raw[1], str) or not _LATEX_CITE.match(raw[1].lstrip())):
            raise ValueError("Only unstarred LaTeX \\cite, \\citep and \\parencite can be linked automatically")
    return [_citation_item(item, latex) for item in citations]


def inspect_citations(ast: Mapping[str, Any], source_format: str) -> tuple[list[dict], list[dict]]:
    """Collect citation clusters in body order, and flag unsupported sites."""
    blocks, meta = _document_parts(ast)
    source_format = source_format.casefold().lstrip(".")
    latex = source_format == "tex" or source_format.startswith("latex")
    if not latex and source_format != "md" and not source_format.startswith("markdown"):
        raise ValueError("Citekey linking requires Markdown or LaTeX")
    clusters = []
    issues = _metadata_issues(meta)
    for index, (node, blocked, in_note) in enumerate(_walk_citations(blocks), 1):
        context = blocked or ("note" if in_note else "body")
        try:
            if blocked:
                raise ValueError(f"Citations in {blocked} need review before conversion")
            clusters.append({"index": index, "items": _citation_items(node, latex)})
        except ValueError as error:
            issues.append({"index": index, "context": context, "reason": str(error)})
    return clusters, issues


def replace_citations(ast: Mapping[str, Any], fields_by_index: Mapping[int, str]) -> dict:
    """Replace every reviewed body Cite with its prepared Word field XML."""
    result = deepcopy(ast)
    blocks, meta = _document_parts(result)
    nodes = list(_walk_citations(blocks))
    if _metadata_issues(meta) or any(blocked for _, blocked, _ in nodes):
        raise ValueError("Cannot replace citations in unsupported Pandoc contexts")
    expected = set(range(1, len(nodes) + 1))
    if any(type(index) is not int for index in fields_by_index) or set(fields_by_index) != expected:
        raise ValueError("Prepared fields do not cover every Pandoc citation exactly once")
    for index, (node, _, _) in enumerate(nodes, 1):
        field = fields_by_index[index]
        if not isinstance(field, str) or not field.strip() or "\0" in field:
            raise ValueError(f"Missing or invalid Word field for citation {index}")
        node.clear()
        node.update(t="RawInline", c=["openxml", field])
    return result
