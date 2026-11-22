"""Verify editable MathType equations in manuscript and end-page table copies."""

from __future__ import annotations

import argparse
from copy import deepcopy
import hashlib
import json
from pathlib import Path
import sys

import pymupdf
from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from paperflow.mathtype import MATHTYPE_PROGID, MathTypeSDK, audit_docx_math
from paperflow.package import digest, part_fields, q, read_package, xml
from paperflow.pipeline import CNS, run_pipeline
from paperflow.recovery import _word_equations
from paperflow.word import DOCX_FORMAT, private_word


PROJECT = Path(__file__).resolve().parents[1]
ABSTRACT = "This synthetic manuscript checks native editable equations inside a table."
CAPTION = "Synthetic table with an editable frequency equation"
TABLE_TEXT = ["Case", "Expression", "Natural frequency", ""]
_V = "{urn:schemas-microsoft-com:vml}"


def _make_source(path):
    document = Document()
    document.add_paragraph("MathType Table End Page Fixture", style="Title")
    document.add_heading("Abstract", 1)
    document.add_paragraph(ABSTRACT)
    document.add_paragraph("Keywords: Document workflow; Editable equations")
    document.add_heading("1 Table equations", 1)
    document.add_paragraph("Table [[REF freqtable]] records the frequency expression for the synthetic model.")
    document.add_paragraph(f"[[CAPTION Table freqtable|{CAPTION}]]")
    table = document.add_table(rows=2, cols=2)
    for cell, label in zip((cell for row in table.rows for cell in row.cells), TABLE_TEXT):
        cell.text = label
    equation = _word_equations([r"\omega_n = \sqrt{k/m}"])[0]
    table.cell(1, 1).paragraphs[0]._p.append(deepcopy(equation))
    document.add_paragraph("The final table page provides a second independent editable copy of the equation.")
    document.save(path)


def _inspect(path):
    parts = read_package(path)
    document = xml(parts["word/document.xml"])
    body = document.find(q("w:body"))
    all_tables = body.findall(q("w:tbl"))
    tables = [["".join(text.text or "" for text in cell.iter(q("w:t"))) for cell in table.iter(q("w:tc"))]
              for table in all_tables]
    equation_counts = [sum(1 for _ in table.iter(q("o:OLEObject"))) for table in all_tables]
    refs = []
    for obj in body.iter(q("o:OLEObject")):
        container = obj.getparent()
        shape = container.find(_V + "shape")
        if shape is None or shape.get("id") != obj.get("ShapeID"):
            raise ValueError("Word-saved MathType object has no matching preview shape ID")
        refs.append({"shape_id": shape.get("id"), "object_id": obj.get("ObjectID")})
    fields, field_errors = part_fields(body)
    audit = audit_docx_math(path)
    return {"tables": tables, "table_equation_counts": equation_counts, "fields": fields, "field_errors": field_errors,
            "native_objects": audit["objects"], "omml_equations": audit["omml_equations"], "ids": refs}


def _render(path, output, label):
    pages = []
    with pymupdf.open(path) as pdf:
        for number, page in enumerate(pdf, 1):
            pixmap = page.get_pixmap(dpi=130, colorspace=pymupdf.csRGB, alpha=False)
            preview = output / f"{label}-{number}.png"
            pixmap.save(preview)
            text = " ".join(page.get_text().split())
            pages.append({"page": number, "preview": str(preview), "caption": f"Table 1 {CAPTION}" in text,
                          "table_text": "Natural frequency" in text,
                          "pixels_sha256": hashlib.sha256(pixmap.samples).hexdigest()})
    return pages


def _edit_clone(submission, destination):
    # Use a private Word session and only save to a new disposable fixture.
    with private_word() as app:
        document = app.Documents.Open(FileName=str(submission), ReadOnly=True, AddToRecentFiles=False, Visible=False)
        try:
            native = [shape for shape in document.InlineShapes if shape.Type == 1 and shape.OLEFormat.ProgID == MATHTYPE_PROGID]
            if len(native) != 2:
                raise RuntimeError("Expected two independently accessible MathType objects in Word")
            last = native[-1]
            with MathTypeSDK() as sdk:
                last.OLEFormat.DoVerb(2)
                ole = last.OLEFormat.Object
                sdk.set_ole_equation(ole, r"+\eta", "latex")
                ole = None
            last = native = None
            document.SaveAs2(FileName=str(destination), FileFormat=DOCX_FORMAT, AddToRecentFiles=False)
        finally:
            document.Close(SaveChanges=0)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path, help="New directory for the fixture and build evidence")
    parser.add_argument("--edit-copy", action="store_true", help="Prove independent editing in a disposable Word copy")
    args = parser.parse_args()
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    source = output / "source.docx"
    _make_source(source)
    before = digest(source.read_bytes())
    built = output / "built"
    report = run_pipeline(source, PROJECT / "templates" / f"{CNS}.docx", built, journal=CNS,
                          table_policy="three-line", submission_copy=True,
                          metadata={"abstract": ABSTRACT, "keywords": ["Document workflow", "Editable equations"]})
    master, submission = built / "editable-master.docx", built / "submission-copy.docx"
    master_data, submission_data = _inspect(master), _inspect(submission)
    master_objects = master_data["native_objects"]
    submission_objects = submission_data["native_objects"]
    master_hashes = [obj.get("mtef_sha256") for obj in master_objects]
    submission_hashes = [obj.get("mtef_sha256") for obj in submission_objects]
    master_pages = _render(master.with_suffix(".pdf"), output, "master")
    submission_pages = _render(submission.with_suffix(".pdf"), output, "submission")
    stages = {step["stage"]: step["evidence"] for step in report["steps"]}
    checks = {
        "source_unchanged": before == digest(source.read_bytes()),
        "master_has_one_native_equation_in_table": master_data["table_equation_counts"] == [1],
        "submission_has_one_native_equation_in_each_table": submission_data["table_equation_counts"] == [1, 1],
        "no_omml_remaining": not master_data["omml_equations"] and not submission_data["omml_equations"],
        "native_storage_valid": all(obj["valid_native_storage"] for obj in master_objects + submission_objects),
        "both_equations_match_master": len(master_hashes) == 1 and submission_hashes == master_hashes * 2,
        "distinct_saved_embedding_parts": len({obj.get("embedding_part") for obj in submission_objects}) == 2,
        "distinct_saved_shape_ids": len({item["shape_id"] for item in submission_data["ids"]}) == 2,
        "distinct_saved_object_ids": len({item["object_id"] for item in submission_data["ids"]}) == 2,
        "table_text_preserved": master_data["tables"] == [TABLE_TEXT] and submission_data["tables"] == [TABLE_TEXT, TABLE_TEXT],
        "original_fields_preserved": not submission_data["field_errors"] and submission_data["fields"] == master_data["fields"],
        "body_render_unchanged": [page["pixels_sha256"] for page in master_pages] ==
                                  [page["pixels_sha256"] for page in submission_pages[:len(master_pages)]],
        "table_on_separate_end_page": stages["journal_object_pages"]["pagination_verified"] and
                                      len(submission_pages) == len(master_pages) + 1 and submission_pages[-1]["caption"] and
                                      submission_pages[-1]["table_text"],
    }
    evidence = {"checks": checks, "master": master_data, "submission": submission_data,
                "placement": stages["journal_object_placement"],
                "master_pages": master_pages, "submission_pages": submission_pages}
    if args.edit_copy:
        submission_sha = digest(submission.read_bytes())
        edited = output / "edited-disposable-copy.docx"
        _edit_clone(submission, edited)
        edited_data = _inspect(edited)
        edited_hashes = [item.get("mtef_sha256") for item in edited_data["native_objects"]]
        checks["editing_clone_changes_only_second_equation"] = (
            len(edited_hashes) == len(submission_hashes) == 2 and edited_hashes[0] == submission_hashes[0] and
            edited_hashes[1] != submission_hashes[1] and all(item["valid_native_storage"] for item in edited_data["native_objects"])
        )
        checks["unedited_submission_unchanged"] = submission_sha == digest(submission.read_bytes())
        evidence["edited_disposable_copy"] = edited_data
    evidence["pass"] = all(checks.values())
    (output / "mathtype-table-end-evidence.json").write_text(json.dumps(evidence, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(checks, ensure_ascii=False, indent=2), flush=True)
    return 0 if evidence["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
