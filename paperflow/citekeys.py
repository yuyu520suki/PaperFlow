"""Review Pandoc citekeys against BibTeX and create linked Word citation fields."""
from collections import Counter
import json
from pathlib import Path
import re
import subprocess

from lxml import etree as ET

from .citations import _owner, _payload
from .citekey_ast import inspect_citations, replace_citations
from .doi import normalize_doi, resolve_dois
from .ingest import LATEX_AST_FORMAT, PANDOC_FORMATS, find_pandoc
from .package import MARKER, digest, field_runs, part_fields, read_package, xml
from .zotero import ITEM_KEY, ZoteroClient, make_citation_field


_SOURCE_SUFFIXES = {".md", ".markdown", ".tex"}
_BIB_MARKER = "[[BIBLIOGRAPHY]]"
_REFERENCES = {"references", "bibliography", "works cited", "参考文献", "引用文献"}


def _run_pandoc(command, *, data=None, cwd=None):
    result = subprocess.run([find_pandoc(), *command], input=data, cwd=cwd, capture_output=True, timeout=180)
    warning = result.stderr.decode("utf-8", errors="replace").strip()
    if result.returncode:
        raise RuntimeError(f"Pandoc citation input could not be read: {warning}")
    try:
        return json.loads(result.stdout), warning
    except ValueError as exc:
        raise RuntimeError("Pandoc did not return valid citation input JSON") from exc


def _fingerprint_ast(ast):
    return digest(json.dumps(ast, sort_keys=True, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))


def _source_ast(source):
    source = Path(source).resolve(strict=True)
    suffix = source.suffix.lower()
    if suffix not in _SOURCE_SUFFIXES:
        raise ValueError("Citekey migration requires a Markdown or LaTeX source")
    source_hash = digest(source.read_bytes())
    fmt = LATEX_AST_FORMAT if suffix == ".tex" else PANDOC_FORMATS[suffix]
    ast, warning = _run_pandoc(["--from", fmt, "--to=json", str(source)], cwd=source.parent)
    if source_hash != digest(source.read_bytes()):
        raise ValueError("Source changed while Pandoc was reading its citations")
    clusters, issues = inspect_citations(ast, fmt)
    if warning:
        issues.append({"reason": "Pandoc source warning", "detail": warning})
    if not clusters and not issues:
        issues.append({"reason": "No Pandoc citation groups were found"})
    return ast, source_hash, fmt, clusters, issues


def require_plan_for_structured_citations(source):
    if Path(source).suffix.lower() not in _SOURCE_SUFFIXES:
        return
    _, _, _, clusters, issues = _source_ast(source)
    if clusters or any("index" in issue for issue in issues):
        raise ValueError("This manuscript has Pandoc citations. Review them with plan-citekeys and use --citekey-plan.")


def _bibliography_data(bibliography):
    path = Path(bibliography).resolve(strict=True)
    if path.suffix.lower() != ".bib":
        raise ValueError("Provide a .bib file for citekey migration")
    data = path.read_bytes()
    records, warning = _run_pandoc(["--sandbox", "--from=biblatex", "--to=csljson"], data=data)
    if warning:
        raise ValueError(f"Review the BibTeX file before linking citations: {warning}")
    if not isinstance(records, list):
        raise ValueError("Pandoc did not return a BibTeX record list")
    indexed = {}
    for record in records:
        key = record.get("id") if isinstance(record, dict) else None
        if not isinstance(key, str) or not key or key in indexed:
            raise ValueError(f"The BibTeX file has a missing or duplicate citekey: {key!r}")
        indexed[key] = record
    return {"path": str(path), "sha256": digest(data)}, indexed


def _inline_text(inlines):
    if not all(node.get("t") in {"Str", "Space", "SoftBreak"} for node in inlines):
        return None
    return "".join(node.get("c", "") if node["t"] == "Str" else " " for node in inlines).strip()


def _bibliography_location(ast):
    """Honor one explicit bibliography position or append one reference section."""
    targets, headings, appendices = [], [], []
    standalone_markers = 0
    for index, block in enumerate(ast["blocks"]):
        kind, content = block.get("t"), block.get("c")
        if kind in {"Para", "Plain"}:
            text = _inline_text(content)
            match = MARKER.fullmatch(text) if text is not None else None
            if match and match[1].startswith("BIBLIOGRAPHY"):
                standalone_markers += 1
                targets.append({"mode": "existing_marker", "index": index})
        elif kind == "Header":
            title = _inline_text(content[2])
            if title and title.rstrip(":：").casefold() in _REFERENCES:
                headings.append(index)
        elif kind == "Div":
            (identifier, classes, _), blocks = content
            if blocks and (identifier == "refs" or any(name in classes for name in ("thebibliography", "references"))):
                return None, "A static bibliography is already present. Review it and leave an explicit bibliography position."
            if identifier == "refs":
                targets.append({"mode": "replace", "index": index})
        elif kind == "RawBlock" and content[0] in {"latex", "tex"}:
            raw = content[1].strip()
            if raw == r"\printbibliography":
                mode = "replace" if index - 1 in headings else "replace_with_heading"
                targets.append({"mode": mode, "index": index})
            elif raw == r"\appendix":
                appendices.append(index)
            elif re.search(r"\\(?:printbibliography\b|begin\{thebibliography\}|bibitem\b)", raw):
                return None, "The LaTeX bibliography command needs an explicit, supported position"
    if json.dumps(ast["blocks"]).count("[[BIBLIOGRAPHY") != standalone_markers:
        return None, "Place the PaperFlow bibliography marker in one standalone body paragraph"
    if len(targets) > 1:
        return None, "More than one bibliography position was found"
    if targets:
        return targets[0], None
    if headings and headings != [len(ast["blocks"]) - 1]:
        return None, "The References heading already has content or later sections. Add one explicit bibliography position."
    if appendices and not headings:
        return None, "A LaTeX appendix needs an explicit bibliography position"
    return {"mode": "append" if headings else "append_with_heading", "index": None}, None


def _keys(clusters):
    return list(dict.fromkeys(item["citekey"] for cluster in clusters for item in cluster["items"]))


def _entry_doi(record):
    return normalize_doi(record.get("DOI")) or normalize_doi(record.get("URL"))


def _inspect(source, bibliography):
    ast, source_hash, fmt, clusters, issues = _source_ast(source)
    bib_info, bib_records = _bibliography_data(bibliography)
    location, reason = _bibliography_location(ast)
    if reason:
        issues.append({"reason": reason})
    cited = _keys(clusters)
    missing = [key for key in cited if key not in bib_records]
    if missing:
        issues.append({"reason": "Citekeys missing from the .bib file", "citekeys": missing})
    return ast, bib_records, {
        "schema_version": 1, "source_sha256": source_hash, "source_ast_sha256": _fingerprint_ast(ast),
        "source_format": fmt, "bibliography": bib_info, "bibliography_location": location,
        "clusters": clusters, "issues": issues,
    }


def plan_citekeys(source, bibliography, *, client=None):
    """Suggest one local Zotero key only for a unique exact DOI match."""
    _, bib_records, plan = _inspect(source, bibliography)
    cited = _keys(plan["clusters"])
    dois = {key: _entry_doi(bib_records[key]) for key in cited if key in bib_records}
    matches = resolve_dois(sorted({doi for doi in dois.values() if doi}), client=client)
    entries, citation_map = [], {}
    for key, doi in dois.items():
        candidates = matches.get(doi, [])
        status = ("unique_doi" if len(candidates) == 1 else "multiple_doi_matches" if candidates
                  else "missing_in_library" if doi else "manual_key_required")
        if len(candidates) == 1:
            citation_map[key] = candidates[0]["key"]
        entries.append({"citekey": key, "title": bib_records[key].get("title"), "doi": doi,
                        "status": status, "candidates": candidates})
    return {**plan, "entries": entries, "citation_map": citation_map}


def _place_bibliography(ast, location):
    mode, index = location["mode"], location["index"]
    if mode == "existing_marker":
        return
    marker = {"t": "Para", "c": [{"t": "Str", "c": _BIB_MARKER}]}
    heading = {"t": "Header", "c": [1, ["", [], []], [{"t": "Str", "c": "References"}]]}
    additions = [heading, marker] if mode.endswith("_with_heading") else [marker]
    if index is None:
        ast["blocks"].extend(additions)
    else:
        ast["blocks"][index:index + 1] = additions


def prepare_citekey_import(source, plan, *, client=None):
    """Revalidate the reviewed input and return an AST with native Zotero fields."""
    if not isinstance(plan, dict) or type(plan.get("schema_version")) is not int or plan["schema_version"] != 1:
        raise ValueError("Unknown citekey plan version")
    bib = plan.get("bibliography")
    if not isinstance(bib, dict) or not isinstance(bib.get("path"), str):
        raise ValueError("The citekey plan does not specify its .bib file")
    ast, bib_records, observed = _inspect(source, bib["path"])
    if any(plan.get(key) != value for key, value in observed.items()):
        raise ValueError("The manuscript, parsed citations, or .bib file differs from the reviewed citekey plan")
    if observed["issues"]:
        raise ValueError(f"Resolve citekey plan issues before converting: {observed['issues']}")
    mapping = plan.get("citation_map")
    cited = _keys(observed["clusters"])
    if (not isinstance(mapping, dict) or set(mapping) != set(cited)
            or any(not isinstance(key, str) or not ITEM_KEY.fullmatch(key) for key in mapping.values())):
        missing = [key for key in cited if not isinstance(mapping, dict) or key not in mapping]
        raise ValueError(f"Assign one reviewed Zotero item key to every citekey. Missing keys: {missing}")
    client = client if client is not None else ZoteroClient()
    records = {key: client.get_item(key) for key in dict.fromkeys(mapping.values())}
    for key, item_key in mapping.items():
        doi = _entry_doi(bib_records[key])
        actual = _entry_doi(records[item_key]["item_data"])
        if doi and doi != actual:
            raise ValueError(f"Zotero item {item_key} does not match the reviewed DOI for citekey {key!r}")
    fields, identifiers = {}, []
    for cluster in observed["clusters"]:
        items = []
        for item in cluster["items"]:
            record = records[mapping[item["citekey"]]]
            items.append({**record, **item})
        code = make_citation_field(items)
        identifiers.append(_payload(code, "ITEM")["citationID"])
        fields[cluster["index"]] = "".join(ET.tostring(run, encoding="unicode") for run in field_runs(code))
    result = replace_citations(ast, fields)
    result["meta"].pop("bibliography", None)
    _place_bibliography(result, observed["bibliography_location"])
    evidence = {"source_sha256": observed["source_sha256"], "source_ast_sha256": observed["source_ast_sha256"],
                "bibliography": bib, "cluster_count": len(fields), "citation_map": dict(mapping),
                "bibliography_location": observed["bibliography_location"], "library_modified": False}
    return result, identifiers, evidence


def verify_imported_citekey_fields(path, identifiers):
    """Verify every generated citation reached the Word package exactly once."""
    saved = []
    for part, data in read_package(path).items():
        if part.startswith("word/") and part.endswith(".xml"):
            fields, errors = part_fields(xml(data))
            if errors:
                raise ValueError(f"Pandoc produced malformed Word fields in {part}: {errors}")
            saved.extend(_payload(code, "ITEM").get("citationID") for code in fields if _owner(code) == "ITEM")
    if Counter(saved) != Counter(identifiers):
        raise ValueError("Pandoc did not preserve all generated Zotero citations in the Word package")
    return len(saved)
