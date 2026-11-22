"""Exercise static numbered citations through the installed Zotero Word integration."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from scripts.zotero_fixture import FIRST_DOI, SECOND_DOI, fixture_key
from paperflow.citations import _owner, _payload, _scan_fields, _uri_identity
from paperflow.package import audit_docx, digest, q, read_package, xml
from paperflow.pipeline import CNS_STYLE, run_pipeline
from paperflow.zotero import ZoteroClient


PROJECT = Path(__file__).resolve().parents[1]
TEMPLATE = PROJECT / "templates" / "computers-and-structures.docx"
OLD_ROW_1, OLD_ROW_2 = "Static row sentinel Alpha", "Static row sentinel Beta"
APPENDIX_TEXT = "This appendix text must survive the bibliography replacement."
APPENDIX_CELL = "APPENDIX-CELL-220"


def _make_source(path: Path, *, leave_first_uncited: bool = False) -> None:
    document = Document()
    document.add_paragraph("Reviewed Citation Migration Fixture", style="Title")
    document.add_heading("1 Introduction", 1)
    paragraph = document.add_paragraph("Earlier in the text we cite ")
    # Word can split a literal citation across identical formatting runs.
    for token in ("[", "2", "]"):
        paragraph.add_run(token)
    paragraph.add_run(" before the next discussion.")
    if not leave_first_uncited:
        document.add_paragraph("The second discussion cites [1].")
    document.add_paragraph("A later sentence repeats [2].")
    document.add_heading("References", 1)
    document.add_paragraph(f"[1] {OLD_ROW_1}. doi: {FIRST_DOI}.")
    document.add_paragraph(f"[2] {OLD_ROW_2}. https://doi.org/{SECOND_DOI}.")
    document.add_heading("Appendix A", 1)
    document.add_paragraph(APPENDIX_TEXT)
    document.add_table(1, 1).cell(0, 0).text = APPENDIX_CELL
    document.save(path)


def _main_xml(path: Path):
    return xml(read_package(path)["word/document.xml"])


def _visible(root) -> str:
    return "\n".join(node.text or "" for node in root.iter(q("w:t")))


def _field_results(master: Path) -> tuple[list, list]:
    citations, bibliographies = [], []
    for field in _scan_fields(_main_xml(master)):
        owner = _owner(field["code"])
        if owner == "ITEM":
            payload = json.loads(field["code"].split("CSL_CITATION", 1)[1])
            keys = [item["uris"][0].rsplit("/", 1)[-1] for item in payload["citationItems"]]
            citations.append({"keys": keys, "plain_citation": payload["properties"]["plainCitation"],
                              "visible_citation": field["result"]})
        elif owner == "BIBL":
            data = _payload(field["code"], "BIBL")
            bibliographies.append({"visible_result": field["result"], "uncited": data.get("uncited", [])})
    return citations, bibliographies


def _verify_master(master: Path, workflow: dict, *, uncited_uri: str | None = None) -> tuple[dict, list, list]:
    citations, bibliographies = _field_results(master)
    stages = {entry["stage"]: entry["evidence"] for entry in workflow["steps"]}
    refresh = stages["zotero"]
    expected = [([fixture_key(SECOND_DOI)], "[1]"), ([fixture_key(FIRST_DOI)], "[2]"), ([fixture_key(SECOND_DOI)], "[1]")]
    if uncited_uri:
        expected.pop(1)
    cited = [(field["keys"], field["plain_citation"]) for field in citations]
    visible = _visible(_main_xml(master))
    bibliography = bibliographies[0]["visible_result"].casefold() if len(bibliographies) == 1 else ""
    uncited = bibliographies[0]["uncited"] if len(bibliographies) == 1 else []
    rendered_both = bibliography.count(FIRST_DOI) == bibliography.count(SECOND_DOI) == 1
    checks = {
        "native_refresh_verified": refresh["zotero_native_refresh_verified"] is True,
        "two_current_library_records": refresh["current_library_records_checked"] == 2,
        "chosen_csl_style": refresh["style_preflight"]["document_style"] == CNS_STYLE,
        "citations_follow_first_appearance": cited == expected,
        "cached_citations_match_field_payload": all(
            field["plain_citation"] == field["visible_citation"].strip() for field in citations
        ),
        "exactly_one_dynamic_bibliography": len(bibliographies) == 1,
        "bibliography_renders_both_records": rendered_both,
        "original_static_entries_replaced": all(row not in visible for row in (OLD_ROW_1, OLD_ROW_2)),
        "appendix_preserved": APPENDIX_TEXT in visible and APPENDIX_CELL in visible,
        "structural_audit_passed": audit_docx(master)["structural_pass"],
    }
    if uncited_uri:
        checks["uncited_real_library_uri_retained"] = (
            len(uncited) == 1 and len(uncited[0]) >= 1
            and _uri_identity(uncited[0][0]) == _uri_identity(uncited_uri)
        )
    else:
        checks["bibliography_order_matches_citations"] = rendered_both and bibliography.index(SECOND_DOI) < bibliography.index(FIRST_DOI)
        checks["no_uncited_items_in_regular_fixture"] = not uncited
    return checks, citations, bibliographies


def _appendix_xml(path: Path) -> list[bytes]:
    from lxml import etree

    body = _main_xml(path).find(q("w:body"))
    index = next(i for i, block in enumerate(body) if _visible(block) == "Appendix A")
    return [etree.tostring(block, method="c14n", exclusive=True) for block in body[index:]]


def _render_pages(folder: Path) -> tuple[list, str]:
    import pymupdf

    pages = []
    with pymupdf.open(folder / "editable-master.pdf") as document:
        for number, page in enumerate(document, 1):
            preview = folder / f"review-page-{number}.png"
            page.get_pixmap(dpi=130, alpha=False).save(preview)
            pages.append({"image": str(preview), "text": page.get_text()})
    return pages, " ".join(" ".join(page["text"] for page in pages).split())


def _verify_uncited_gate(output: Path, uncited_uri: str) -> tuple[dict, dict]:
    from paperflow.bibliography import apply_migration, plan_migration

    source = output / "uncited-source.docx"
    _make_source(source, leave_first_uncited=True)
    fingerprint = digest(source.read_bytes())
    plan = plan_migration(source)
    refused_output = output / "unexpected-uncited-output.docx"
    reason = None
    try:
        apply_migration(source, plan, refused_output)
    except ValueError as error:
        reason = str(error)
    reviewed = {**plan, "include_uncited": True}
    retained = output / "reviewed-uncited.docx"
    report = apply_migration(source, reviewed, retained)
    checks = {
        "default_uncited_replacement_refused": (
            isinstance(reason, str) and "uncited" in reason.casefold() and not refused_output.exists()
        ),
        "explicit_uncited_retention_preserved_number": report["uncited_numbers"] == ["1"],
        "uncited_source_unchanged": fingerprint == digest(source.read_bytes()),
    }
    built = output / "built-with-uncited"
    workflow = run_pipeline(retained, TEMPLATE, built, csl_style=CNS_STYLE)
    verified, citations, bibliography = _verify_master(built / "editable-master.docx", workflow, uncited_uri=uncited_uri)
    checks.update({f"uncited_variant_{key}": value for key, value in verified.items()})
    pages, text = _render_pages(built)
    checks["uncited_variant_has_only_two_rendered_body_citations"] = (
        "Earlier in the text we cite [1] before" in text and "A later sentence repeats [1]." in text
        and "The second discussion" not in text
    )
    return checks, {"refusal_reason": reason, "reviewed_retention": report, "citations": citations,
                    "bibliography": bibliography, "pages": pages}


def main() -> int:
    from paperflow.bibliography import apply_migration, plan_migration

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    output = parser.parse_args().output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    source = output / "static-citations.docx"
    _make_source(source)
    fingerprint = digest(source.read_bytes())
    client = ZoteroClient()
    records = {key: client.get_item(key) for key in (fixture_key(FIRST_DOI), fixture_key(SECOND_DOI))}
    versions_before = {key: item["version"] for key, item in records.items()}
    plan = plan_migration(source)
    (output / "reviewed-migration-plan.json").write_text(
        json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    migrated = output / "citation-markers.docx"
    report = apply_migration(source, plan, migrated)
    expected_map = {"1": fixture_key(FIRST_DOI), "2": fixture_key(SECOND_DOI)}
    checks = {
        "exact_local_doi_matches": plan["citation_map"] == expected_map and
                                  {str(entry["number"]): entry["doi"] for entry in plan["entries"]} ==
                                  {"1": FIRST_DOI, "2": SECOND_DOI},
        "three_literal_citations_converted": report["citation_markers"] == 3,
        "only_two_numbered_rows_retired": report["retired_entries"] == 2,
        "appendix_xml_unchanged_during_migration": _appendix_xml(source) == _appendix_xml(migrated),
        "migration_source_hash_verified": report["source_unchanged"] is True and plan["source_sha256"] == fingerprint,
    }
    built = output / "built"
    workflow = run_pipeline(migrated, TEMPLATE, built, csl_style=CNS_STYLE)
    verified, citations, bibliography = _verify_master(built / "editable-master.docx", workflow)
    checks.update(verified)
    pages, text = _render_pages(built)
    gate_checks, gate_evidence = _verify_uncited_gate(output, records[fixture_key(FIRST_DOI)]["uri"])
    checks.update(gate_checks)
    checks.update(
        rendered_text_uses_new_numbers=all(snippet in text for snippet in (
            "Earlier in the text we cite [1] before", "The second discussion cites [2].", "A later sentence repeats [1]."
        )),
        rendered_appendix_preserved=APPENDIX_TEXT in text and APPENDIX_CELL in text,
        source_unchanged=fingerprint == digest(source.read_bytes()),
        cited_library_item_versions_unchanged=versions_before == {key: client.get_item(key)["version"] for key in (fixture_key(FIRST_DOI), fixture_key(SECOND_DOI))},
    )
    evidence = {"pass": all(checks.values()), "checks": checks, "migration": report,
                "uncited_gate": gate_evidence, "citations": citations, "bibliography": bibliography, "pages": pages}
    (output / "bibliography-evidence.json").write_text(
        json.dumps(evidence, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(checks, ensure_ascii=False, indent=2), flush=True)
    return 0 if evidence["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
