"""Synthetic fixture only: validate native Word without touching personal files."""
from __future__ import annotations
import argparse
from copy import deepcopy
import hashlib
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from scripts.zotero_fixture import FIRST_DOI, fixture_key
from paperflow.ingest import import_document
from paperflow.package import expand_markers, field_runs, inspect_docx, audit_docx
from paperflow.word import private_word, build_from_template, refresh_fields, export_pdf
from paperflow.zotero import ZoteroClient, document_preferences, set_docx_preferences


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("output_dir"); args = ap.parse_args()
    out = Path(args.output_dir).resolve(); out.mkdir(parents=True, exist_ok=False)
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    template = Document()
    sec = template.sections[0]
    sec.top_margin = sec.bottom_margin = Inches(1)
    sec.left_margin = sec.right_margin = Inches(1)
    style = template.styles["Normal"]
    style.font.name = "Times New Roman"; style.font.size = Pt(12)
    style.paragraph_format.line_spacing = 1.5
    template.styles["Heading 1"].font.name = "Times New Roman"
    template.styles["Heading 1"].font.size = Pt(14)
    template.styles["Heading 1"].font.bold = True
    sec.header.paragraphs[0].text = "Paperflow integration fixture"
    footer = sec.footer.paragraphs[0]
    for e in field_runs(" PAGE ", "1"): footer._p.append(e)
    footer.add_run(" / ")
    for e in field_runs(" NUMPAGES ", "1"): footer._p.append(e)
    template.add_paragraph("{{PAPERFLOW_BODY}}")
    fixture = out / "template.docx"; template.save(fixture)
    before_hash = hashlib.sha256(fixture.read_bytes()).hexdigest()
    source = out / "source.md"
    source.write_text("""# 1 Introduction

This document tests editable native fields and template preservation.

[[TOC]]

## 1.1 References and equations

The linked reference is [[CITE FIXTURE_KEY]]. Figure [[REF response]] is referenced dynamically.

$$
\\mathbf{M}\\ddot{\\mathbf{x}}+\\mathbf{C}\\dot{\\mathbf{x}}+\\mathbf{K}\\mathbf{x}=\\mathbf{f}(t)
$$

$$
\\mathbf{A}=\\begin{bmatrix}1&2\\\\3&4\\end{bmatrix}
$$

[[CAPTION Figure response|A fixture caption]]

| Case | Value |
|:-----|------:|
| A | 1.25 |
| B | 2.50 |

[[CAPTION Table results|Fixture values]]

# 2 References

[[BIBLIOGRAPHY]]
""".replace("FIXTURE_KEY", fixture_key(FIRST_DOI)), encoding="utf-8")
    print("Importing Markdown", flush=True)
    imported = out / "imported.docx"
    ingestion = import_document(source, imported, template=fixture)
    item = ZoteroClient().get_item(fixture_key(FIRST_DOI))
    expected_title = item["item_data"]["title"]
    item["item_data"]["title"] = "PAPERFLOW_STALE_TITLE_FOR_DYNAMIC_REFRESH_TEST"
    fields = out / "fields.docx"
    expand_markers(imported, fields, {fixture_key(FIRST_DOI): item})
    set_docx_preferences(fields, document_preferences("http://www.zotero.org/styles/ieee"))
    print("Starting isolated Word instance", flush=True)
    with private_word() as app:
        print("Word version " + app.Version, flush=True)
        built = out / "built.docx"
        report = build_from_template(fixture, fields, built, app=app)
        print("Template content inserted", flush=True)
        refreshed = out / "refreshed.docx"
        ordinary = refresh_fields(built, refreshed, app=app)
        print("Ordinary fields refreshed", flush=True)
        pdf = out / "refreshed.pdf"
        exported = export_pdf(refreshed, pdf, app=app)
        print("Word PDF exported", flush=True)
    result = {"fixture_only": True, "template_unchanged": before_hash == hashlib.sha256(fixture.read_bytes()).hexdigest(),
              "ingestion": ingestion, "build": report, "refresh": ordinary, "export": exported,
              "structural_audit": audit_docx(refreshed), "expected_library_title": expected_title,
              "zotero_native_refresh_verified": False, "mathtype_native_verified": False,
              "official_journal_template_verified": False}
    (out / "native-smoke.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"output_dir": str(out), "template_unchanged": result["template_unchanged"],
                      "sections": result["structural_audit"]["section_count"],
                      "fields": len(result["structural_audit"]["fields"]),
                      "pending": result["structural_audit"]["pending_markers"]}, ensure_ascii=False), flush=True)


if __name__ == "__main__": main()
