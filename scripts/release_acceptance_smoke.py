"""Accept one complete synthetic paper through build, native edits and revision.

Only run this script while the caller owns the private Word/MathType channel.
It reads two existing Zotero records, never changes the library, and writes
only new fixture/output files. Every exported PDF page is rendered for review.
"""
from __future__ import annotations

import argparse
import csv
from collections import Counter
from copy import deepcopy
import json
import math
from pathlib import Path
import re
import shutil
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt
from lxml import etree as ET
import pymupdf

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from paperflow.mathtype import MathTypeSDK, audit_docx_math
from paperflow.package import NS, audit_docx, digest, expand_markers, field_runs, q, read_package, xml
from paperflow.pipeline import CNS, CNS_STYLE, revise_master, run_pipeline
from paperflow.recovery import _word_equations
from paperflow.template import section_story_signature
from paperflow.word import private_word
from paperflow.zotero import ZoteroClient, document_preferences, make_citation_field, set_docx_preferences


TITLE = "Complete local manuscript workflow acceptance"
HEADER = "PaperFlow acceptance header"
FOOTER = "PaperFlow acceptance footer"
BODY_WITNESS = "This complete paper is a synthetic software acceptance fixture."
PREFACE = "Revision preface"
FIGURE_CAPTION = "Synthetic response for complete workflow acceptance"
NEW_FIGURE_CAPTION = "Revision auxiliary response"
TABLE_CAPTION = "Editable synthetic response measurements"
WIDE_ALT = "PaperFlow release original response"
NEW_ALT = "PaperFlow release revision response"
STALE_TITLE = "PAPERFLOW_RELEASE_STALE_TITLE"
ABSTRACT = (
    "This synthetic manuscript validates a local paper workflow. It combines a supplied Word template, "
    "editable mathematical equations, linked literature citations, document navigation, figures, tables, "
    "and cross-references. A subsequent working-copy revision checks that citations, numbering and page "
    "references refresh while the original files remain intact. The numerical values are software test "
    "fixtures and do not describe experimental research."
)
KEYWORDS = ["Document workflow", "Editable equations", "Cross-references"]
INDEX_LABEL = re.compile(r'\\c\s+"?(Figure|Table)"?(?=\s|$)', re.I)


def _read(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def _write(path, value):
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _stale_item(item):
    result = deepcopy(item)
    result["item_data"]["title"] = STALE_TITLE
    return result


def _template(path):
    document = Document(PROJECT / "templates/computers-and-structures.docx")
    section = document.sections[0]
    section.header.paragraphs[0].text = HEADER
    footer = section.footer.paragraphs[0]
    footer.text = FOOTER + " "
    for run in field_runs(" PAGE ", "0"):
        footer._p.append(run)
    # Keep the template's A4/25.4 mm geometry and declared project typography.
    document.save(path)


def _fixed(table, widths):
    table.autofit = False
    table.style = "Table Grid"
    for column, width in zip(table.columns, widths, strict=True):
        column.width = Inches(width)
    for row in table.rows:
        for cell, width in zip(row.cells, widths, strict=True):
            cell.width = Inches(width)
    preferred = table._tbl.tblPr.find(q("w:tblW"))
    preferred.set(q("w:type"), "dxa")
    preferred.set(q("w:w"), str(sum(round(width * 1440) for width in widths)))


def _source(path, item):
    document = Document()
    section = document.sections[0]
    section.page_width = Inches(11)
    section.left_margin = section.right_margin = Inches(0.5)
    document.styles["Normal"].font.name = "Calibri"
    document.styles["Normal"].font.size = Pt(14)
    document.add_paragraph(TITLE, style="Title")
    document.add_paragraph("Workflow Test Author — synthetic software validation")
    document.add_heading("Abstract", 1)
    document.add_paragraph(ABSTRACT)
    document.add_paragraph("Keywords: " + "; ".join(KEYWORDS))
    document.add_paragraph("Contents").paragraph_format.page_break_before = True
    for heading, marker in ((None, "TOC"), ("List of Figures", "LOF"), ("List of Tables", "LOT")):
        if heading:
            document.add_paragraph(heading)
        document.add_paragraph("[[" + marker + "]]")
    introduction = document.add_heading("1 Introduction", 1)
    introduction.paragraph_format.page_break_before = True
    document.add_paragraph(BODY_WITNESS)
    paragraph = document.add_paragraph("Initial literature-link witness ")
    for run in field_runs(make_citation_field([_stale_item(item)]), "[999]"):
        paragraph._p.append(run)
    paragraph.add_run(". This source is cited only to validate the library connection.")
    document.add_paragraph("See Figure [[REF response]] and Table [[REF results]] on page [[PAGEREF results]]. "
                           "The model is Equation ([[REF dynamics]]).")
    document.add_heading("2 Methods", 1)
    document.add_paragraph("The following two independent editable equations exercise the native equation workflow.")
    equations = _word_equations(["y=x^2+1", "A=\\begin{bmatrix}1&2\\\\3&4\\end{bmatrix}"])
    for equation, label in zip(equations, ("dynamics", "matrix"), strict=True):
        paragraph = document.add_paragraph()
        paragraph._p.append(deepcopy(equation))
        paragraph.add_run(" [[EQNO " + label + "]]")
    document.add_heading("3 Results", 1)
    document.add_paragraph("The plot and measurements contain declared synthetic data.")
    shape = document.add_picture(str(path.parent / "response.png"), width=Inches(9))
    shape._inline.docPr.set("descr", WIDE_ALT)
    document.add_paragraph(f"[[CAPTION Figure response|{FIGURE_CAPTION}]]")
    document.add_paragraph(f"[[CAPTION Table results|{TABLE_CAPTION}]]")
    table = document.add_table(rows=5, cols=3)
    _fixed(table, (3, 3.5, 2.5))
    for row, values in zip(table.rows, (
        ("Synthetic response series", "", "Observation"),
        ("Pair A-B", "7.500", "Stable"),
        ("", "7.625", "Stable"),
        ("Settings", "Nested acquisition settings", "Native equation"),
        ("Case C", "7.750", "Stable"),
    ), strict=True):
        for cell, value in zip(row.cells, values, strict=True):
            cell.text = value
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    table.cell(0, 0).merge(table.cell(0, 1))
    table.cell(1, 0).merge(table.cell(2, 0))
    table.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    nested = table.cell(3, 1).add_table(rows=2, cols=2)
    _fixed(nested, (1.8, 0.9))
    for row, values in zip(nested.rows, (("Interval", "Value"), ("Sample", "0.02 s")), strict=True):
        for cell, value in zip(row.cells, values, strict=True):
            cell.text = value
    table.cell(3, 2).paragraphs[0]._p.append(deepcopy(_word_equations(["x+1=2"])[0]))
    document.add_heading("4 Conclusion", 1)
    document.add_paragraph("This document reports software behavior only. All fixture values remain editable.")
    document.add_heading("References", 1)
    document.add_paragraph("[[BIBLIOGRAPHY]]")
    document.save(path)
    set_docx_preferences(path, document_preferences(CNS_STYLE, locale="en-US"))


def _revision_artwork(output):
    """A distinct synthetic plot makes figure-number/source remapping testable."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    points = [(index / 10, math.exp(-0.2 * index / 10) * math.cos(index / 10)) for index in range(61)]
    with (output / "revision-response.csv").open("x", encoding="utf-8", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(("time_s", "auxiliary_synthetic_response"))
        writer.writerows(points)
    with plt.rc_context({"font.family": "Times New Roman", "font.size": 10, "pdf.fonttype": 42}):
        figure, axis = plt.subplots(figsize=(5, 2.5))
        axis.plot([point[0] for point in points], [point[1] for point in points], color="black", linewidth=1)
        axis.set(xlabel="Time (s)", ylabel="Auxiliary synthetic response")
        figure.tight_layout()
        figure.savefig(output / "revision-response.png", dpi=300)
        figure.savefig(output / "revision-response.pdf")
        plt.close(figure)


def _paragraph(document, prefix, *, exact=False):
    matches = [document.Paragraphs.Item(index) for index in range(1, document.Paragraphs.Count + 1)
               if ((str(document.Paragraphs.Item(index).Range.Text).strip("\r\x07 ") == prefix) if exact else
                   str(document.Paragraphs.Item(index).Range.Text).strip("\r\x07 ").startswith(prefix))]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one native paragraph beginning {prefix!r}; found {len(matches)}")
    return matches[0]


def _native_edit(master, target, image):
    """Edit only a new Word copy and its second controlled native equation."""
    with private_word() as app:
        document = app.Documents.Open(FileName=str(master), ReadOnly=True, AddToRecentFiles=False, Visible=False)
        try:
            heading = _paragraph(document, "1 Introduction", exact=True)
            position = document.Range(heading.Range.Start, heading.Range.Start)
            position.InsertBefore(PREFACE + "\rA separate revision page checks navigation and pagination.\r")
            preface = _paragraph(document, PREFACE, exact=True)
            preface.Style = document.Styles.Item(-2)
            preface.Format.PageBreakBefore = True
            _paragraph(document, "A separate revision page").Style = document.Styles.Item(-1)
            _paragraph(document, "1 Introduction", exact=True).Format.PageBreakBefore = True
            witness = _paragraph(document, BODY_WITNESS)
            position = document.Range(witness.Range.Start, witness.Range.Start)
            position.InsertBefore("Revision literature-link witness [[CITE REVISION_ITEM]].\r")
            _paragraph(document, "Revision literature-link witness").Style = document.Styles.Item(-1)
            originals = [document.InlineShapes.Item(index) for index in range(1, document.InlineShapes.Count + 1)
                         if int(document.InlineShapes.Item(index).Type) == 3
                         and str(document.InlineShapes.Item(index).AlternativeText) == WIDE_ALT]
            if len(originals) != 1:
                raise RuntimeError("Could not locate the original native response picture")
            point = originals[0].Range.Paragraphs.Item(1).Range.Start
            document.Range(point, point).InsertBefore(
                "Revision picture placeholder\r[[CAPTION Figure revisionResponse|" + NEW_FIGURE_CAPTION + "]]\r")
            placeholder = _paragraph(document, "Revision picture placeholder")
            placeholder.Style = document.Styles.Item(-1)
            insertion = placeholder.Range.Duplicate
            insertion.End -= 1
            insertion.Text = ""
            insertion.Collapse(1)
            picture = document.InlineShapes.AddPicture(FileName=str(image), LinkToFile=False,
                                                       SaveWithDocument=True, Range=insertion)
            picture.LockAspectRatio = -1
            picture.Width = 216
            picture.AlternativeText = NEW_ALT
            conclusion = _paragraph(document, "4 Conclusion", exact=True)
            document.Range(conclusion.Range.Start, conclusion.Range.Start).InsertBefore(
                "Revision reference witness: original Figure [[REF response]]; new Figure [[REF revisionResponse]].\r")
            _paragraph(document, "Revision reference witness").Style = document.Styles.Item(-1)
            value = document.Content.Duplicate
            if not value.Find.Execute(FindText="7.500", Forward=True, Wrap=0):
                raise RuntimeError("Could not locate the table edit witness")
            value.Text = "7.875"
            equations = [document.InlineShapes.Item(index) for index in range(1, document.InlineShapes.Count + 1)
                         if int(document.InlineShapes.Item(index).Type) == 1
                         and str(document.InlineShapes.Item(index).OLEFormat.ProgID) == "Equation.DSMT4"]
            if len(equations) != 3:
                raise RuntimeError(f"Expected three independent native formulas; found {len(equations)}")
            with MathTypeSDK() as sdk:
                equations[1].OLEFormat.DoVerb(2)
                native = equations[1].OLEFormat.Object
                sdk.set_ole_equation(native, r"+\eta", "latex")
                native = None
            equations = originals = []
            document.SaveAs2(FileName=str(target), FileFormat=12, AddToRecentFiles=False)
        finally:
            document.Close(SaveChanges=0)


def _replace_revision_key(source, output, key, item):
    # Replace the one placeholder marker on a package copy, not Word/lib data.
    from paperflow.package import write_package
    parts = read_package(source)
    root = xml(parts["word/document.xml"])
    nodes = [node for node in root.iter(q("w:t")) if "[[CITE REVISION_ITEM]]" in (node.text or "")]
    if len(nodes) != 1:
        raise RuntimeError("The revision citation marker was split or duplicated")
    nodes[0].text = nodes[0].text.replace("[[CITE REVISION_ITEM]]", "[[CITE " + key + "]]")
    intermediate = output.with_name("revision-markers.docx")
    parts["word/document.xml"] = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    write_package(parts, intermediate)
    return expand_markers(intermediate, output, {key: _stale_item(item)})


def _stage(workflow, stage):
    return next(step["evidence"] for step in reversed(workflow["steps"]) if step["stage"] == stage)


def _snapshot(workflow, submission=False):
    names = {"submission_word_fields"} if submission else {"word_fields", "word_fields_after_zotero"}
    return next(step["evidence"]["snapshot"] for step in reversed(workflow["steps"]) if step["stage"] in names)


def _render(pdf_path, folder):
    folder.mkdir(exist_ok=True)
    result = []
    with pymupdf.open(pdf_path) as document:
        for index, page in enumerate(document, 1):
            preview = folder / f"page-{index:03}.png"
            page.get_pixmap(dpi=120, alpha=False).save(preview)
            result.append({"page": index, "text": " ".join(page.get_text().split()),
                           "size_pt": [page.rect.width, page.rect.height], "preview": str(preview)})
    return result


def _citation_items(path):
    result = []
    for field in audit_docx(path)["fields"]:
        if "ZOTERO_ITEM CSL_CITATION" not in field["instruction"]:
            continue
        payload = json.loads(field["instruction"].split("CSL_CITATION", 1)[1].strip())
        for item in payload["citationItems"]:
            result.append({"key": item["uris"][0].rsplit("/", 1)[1],
                           "title": item.get("itemData", {}).get("title"),
                           "citation": payload.get("properties", {}).get("plainCitation")})
    return result


def _math_hashes(path):
    audit = audit_docx_math(path)
    return audit, [obj["mtef_sha256"] for obj in audit["objects"]]


def _table_contract(path):
    root = xml(read_package(path)["word/document.xml"])
    table, = root.findall("w:body/w:tbl", NS)
    rows = table.findall(q("w:tr"))
    return {
        "cells": [[(
            cell.find("w:tcPr/w:gridSpan", NS).get(q("w:val"), "1")
            if cell.find("w:tcPr/w:gridSpan", NS) is not None else "1",
            cell.find("w:tcPr/w:vMerge", NS).get(q("w:val"), "continue")
            if cell.find("w:tcPr/w:vMerge", NS) is not None else None,
        ) for cell in row.findall(q("w:tc"))] for row in rows],
        "repeat_header": [row.find("w:trPr/w:tblHeader", NS) is not None for row in rows],
        "cant_split": [row.find("w:trPr/w:cantSplit", NS) is not None for row in rows],
        "nested_count": len(table.findall(".//w:tbl", NS)),
    }


def _checkpoint(output, paths=()):
    path = output / "acceptance-checkpoints.json"
    state = _read(path) if path.is_file() else {}
    for name, checksum in state.items():
        if digest(Path(name).read_bytes()) != checksum:
            raise RuntimeError(f"A completed acceptance artifact changed: {name}")
    for artifact in paths:
        state[str(artifact)] = digest(artifact.read_bytes())
    _write(path, state)


def _stage_paths(folder):
    return [folder / name for name in ("editable-master.docx", "editable-master.pdf", "submission-copy.docx",
                                      "submission-copy.pdf", "workflow.json")]


def _check_stage(folder, workflow, template, items, *, revised):
    docs, checks = {}, {}
    expected_figures = [f"Figure 1 {FIGURE_CAPTION}"] if not revised else [
        f"Figure 1 {NEW_FIGURE_CAPTION}", f"Figure 2 {FIGURE_CAPTION}"]
    fields = _snapshot(workflow)["fields"]
    figure_indexes = [field for field in fields if field["kind"] == "TOC" and (match := INDEX_LABEL.search(field["code"]))
                      and match[1].casefold() == "figure"]
    table_indexes = [field for field in fields if field["kind"] == "TOC" and (match := INDEX_LABEL.search(field["code"]))
                     and match[1].casefold() == "table"]
    contents = [field for field in fields if field["kind"] == "TOC" and not INDEX_LABEL.search(field["code"])]
    figure_rows = [row.strip() for field in figure_indexes for row in field["result"].splitlines() if row.strip()]
    table_rows = [row.strip() for field in table_indexes for row in field["result"].splitlines() if row.strip()]
    checks["contents_and_both_indexes_exist"] = len(contents) == len(figure_indexes) == len(table_indexes) == 1
    checks["contents_include_paper_sections_and_revision_heading"] = bool(contents) and all(
        text in contents[0]["result"] for text in ("1 Introduction", "2 Methods", "3 Results", "4 Conclusion")) \
        and (not revised or PREFACE in contents[0]["result"])
    checks["indexes_have_exactly_the_expected_captions"] = (
        len(figure_rows) == len(expected_figures) and all(row.startswith(expected + "\t")
            for row, expected in zip(figure_rows, expected_figures, strict=True))
        and len(table_rows) == 1 and table_rows[0].startswith(f"Table 1 {TABLE_CAPTION}\t"))
    expected_refs = {"response": "2" if revised else "1", "results": "1", "dynamics": "1"}
    if revised:
        expected_refs["revisionResponse"] = "1"
    for label, expected in expected_refs.items():
        observed = [field["result"].strip() for field in fields if re.match(rf"\s*REF\s+{label}\b", field["code"], re.I)]
        checks["reference_" + label] = bool(observed) and all(value == expected for value in observed)
    layout = _stage(workflow, "table_layout")
    checks["native_reopened_table_and_cell_objects_fit"] = layout["native_reopened"] and not layout["unresolved"] \
        and len(layout["tables"]) == 2 and all(table["within_width_limit"] for table in layout["tables"])
    checks["merged_cells_nested_table_and_row_controls_preserved"] = _table_contract(folder / "editable-master.docx") == \
        _table_contract(folder.parent / "manuscript.docx")
    native = _snapshot(workflow)
    target_section = Document(template).sections[0]
    expected_geometry = {"PageWidth": target_section.page_width.pt, "PageHeight": target_section.page_height.pt,
                         "LeftMargin": target_section.left_margin.pt, "RightMargin": target_section.right_margin.pt,
                         "TopMargin": target_section.top_margin.pt, "BottomMargin": target_section.bottom_margin.pt}
    checks["native_page_geometry_matches_template"] = bool(native["sections"]) and all(
        abs(section[name] - value) < 0.15 for section in native["sections"] for name, value in expected_geometry.items())
    checks["template_body_and_heading_fonts_observed_in_native_word"] = any(
        BODY_WITNESS in paragraph["text"] and paragraph["font"]["Name"] == "Times New Roman"
        and abs(paragraph["font"]["Size"] - 12) < 0.1 for paragraph in native["paragraphs"]) and any(
        paragraph["text"].strip() == "1 Introduction" and paragraph["font"]["Name"] == "Times New Roman"
        and abs(paragraph["font"]["Size"] - 14) < 0.1 for paragraph in native["paragraphs"])
    body_width = native["sections"][0]["PageWidth"] - native["sections"][0]["LeftMargin"] - native["sections"][0]["RightMargin"]
    pictures = [shape for shape in native["inline_shapes"] if shape["type"] == 3]
    checks["all_native_inline_pictures_fit_body_width"] = len(pictures) == (2 if revised else 1) \
        and all(shape["width"] <= body_width + 0.2 for shape in pictures)
    citations = _citation_items(folder / "editable-master.docx")
    expected_order = list(reversed(list(items))) if revised else [next(iter(items))]
    checks["zotero_current_titles_and_expected_numbering"] = [entry["key"] for entry in citations] == expected_order and all(
        entry["title"] == items[entry["key"]]["item_data"]["title"] and entry["citation"] == f"[{index}]"
        for index, entry in enumerate(citations, 1))
    checks["native_zotero_refresh_executed"] = _stage(workflow, "zotero")["zotero_native_refresh_verified"]
    checks["supported_journal_checks_have_no_measured_failures"] = not workflow["journal_audit"]["failed"]
    for name in ("editable-master", "submission-copy"):
        path = folder / (name + ".docx")
        pages = _render(folder / (name + ".pdf"), folder / (name + "-pages"))
        audit = audit_docx(path)
        math, hashes = _math_hashes(path)
        text = " ".join(page["text"] for page in pages)
        checks[name + "_structural_audit_passes"] = audit["structural_pass"]
        checks[name + "_template_furniture_preserved"] = section_story_signature(path) == section_story_signature(template)
        checks[name + "_headers_footers_and_expected_body_are_rendered"] = bool(pages) and all(
            HEADER in page["text"] and FOOTER in page["text"] for page in pages) and BODY_WITNESS in text
        checks[name + "_all_equations_are_valid_native_mathtype"] = math["all_mathtype_objects_valid"] \
            and not math["omml_equations"] and not math["legacy_equation_objects"] \
            and math["native_mathtype_objects"] == (4 if name == "submission-copy" else 3)
        if name == "submission-copy":
            checks["submission_citations_are_fixed_text"] = not _citation_items(path) and not any(
                "ZOTERO_BIBL" in field["instruction"] for field in audit["fields"])
            master_hashes = docs["editable-master"]["math_hashes"]
            checks["submission_preserves_body_equations_and_independent_table_copy"] = Counter(hashes) == (
                Counter(master_hashes) + Counter(master_hashes[-1:]))
        docs[name] = {"pages": pages, "math_hashes": hashes, "citation_items": _citation_items(path)}
    pages = docs["editable-master"]["pages"]
    # The known two-page title/navigation front matter precedes body captions.
    caption_pages = {caption: [page["page"] for page in pages if page["page"] > (3 if revised else 2)
                               and caption in page["text"]] for caption in expected_figures + [f"Table 1 {TABLE_CAPTION}"]}
    checks["index_page_numbers_match_actual_body_caption_pages"] = all(
        len(caption_pages[caption]) == 1 and row == caption + "\t" + str(caption_pages[caption][0])
        for row, caption in zip(figure_rows + table_rows, expected_figures + [f"Table 1 {TABLE_CAPTION}"], strict=True))
    page_refs = [field["result"].strip() for field in fields if re.match(r"\s*PAGEREF\s+results\b", field["code"], re.I)]
    checks["table_page_reference_matches_pdf"] = len(caption_pages[f"Table 1 {TABLE_CAPTION}"]) == 1 and page_refs == [
        str(caption_pages[f"Table 1 {TABLE_CAPTION}"][0])]
    checks["journal_end_object_pagination_verified"] = _stage(workflow, "journal_object_pages")["pagination_verified"]
    artwork = _stage(workflow, "journal_artwork")
    # The packaging module records all staged original figure files separately.
    artwork_paths = list((folder / "artwork").glob("*.pdf"))
    sources = {"1": folder.parent / "response.pdf"} if not revised else {
        "1": folder.parent / "revision-response.pdf", "2": folder.parent / "response.pdf"}
    checks["separate_artwork_originals_packaged"] = len(artwork_paths) == (2 if revised else 1) \
        and artwork["all_figures_staged"] and not artwork["unused_declarations"] \
        and all(record.get("output") and record["figure_number"] in sources
                and digest(Path(record["output"]).read_bytes()) == digest(sources[record["figure_number"]].read_bytes())
                for record in artwork["figures"])
    return {"checks": checks, "documents": docs, "figure_index_rows": figure_rows, "table_index_rows": table_rows,
            "caption_pages": caption_pages, "native_table_layout": layout, "artwork": artwork,
            "journal_audit": workflow["journal_audit"],
            "workflow": str(folder / "workflow.json")}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument("--zotero-item-keys", nargs=2, required=True, metavar=("INITIAL", "REVISION"))
    parser.add_argument("--resume", action="store_true", help="Reuse only completed, hash-matching stages")
    args = parser.parse_args()
    if len(set(args.zotero_item_keys)) != 2:
        raise ValueError("Supply two distinct existing Zotero parent-record keys")
    output = args.output_dir.resolve()
    if args.resume:
        manifest = _read(output / "fixtures.json")
        if manifest["zotero_keys"] != args.zotero_item_keys:
            raise RuntimeError("Resume keys differ from the original fixture")
        if any(digest(Path(path).read_bytes()) != checksum for path, checksum in manifest["input_sha256"].items()):
            raise RuntimeError("An original fixture input changed; refusing to reuse stages")
        _checkpoint(output)
    else:
        output.mkdir(parents=True, exist_ok=False)
    client = ZoteroClient()
    items = {key: client.get_item(key) for key in args.zotero_item_keys}
    source, template = output / "manuscript.docx", output / "template.docx"
    metadata = {"abstract": ABSTRACT, "keywords": KEYWORDS, "fixture_only": True, "figure_sources": {"1": "response.pdf"}}
    if not args.resume:
        for name in ("response.png", "response.pdf"):
            shutil.copyfile(PROJECT / "examples" / CNS / name, output / name)
        _revision_artwork(output)
        _template(template)
        _source(source, items[args.zotero_item_keys[0]])
        original_paths = [source, template, *[output / name for name in (
            "response.png", "response.pdf", "revision-response.png", "revision-response.pdf", "revision-response.csv")]]
        manifest = {"zotero_keys": args.zotero_item_keys,
                    "input_sha256": {str(path): digest(path.read_bytes()) for path in original_paths}}
        _write(output / "fixtures.json", manifest)
    built = output / "built"
    if args.resume and (built / "workflow.json").is_file():
        build = _read(built / "workflow.json")
        if not build.get("native_build_finished"):
            raise RuntimeError("The existing build stage is incomplete; inspect its workflow before recovery")
    else:
        build = run_pipeline(source, template, built, journal=CNS, submission_copy=True,
                             table_policy="three-line", metadata=metadata)
    _checkpoint(output, _stage_paths(built))
    initial = _check_stage(built, build, template, items, revised=False)
    _write(output / "initial-acceptance.json", initial)
    master = built / "editable-master.docx"
    master_hash = digest(master.read_bytes())
    _, initial_hashes = _math_hashes(master)
    raw, linked = output / "word-edited.docx", output / "word-edited-linked.docx"
    if not args.resume or not linked.exists():
        if raw.exists():
            raise RuntimeError("An unfinished native edit exists; inspect it before recovery")
        _native_edit(master, raw, output / "revision-response.png")
        _replace_revision_key(raw, linked, args.zotero_item_keys[1], items[args.zotero_item_keys[1]])
    _checkpoint(output, (raw, linked))
    edited_hash = digest(linked.read_bytes())
    _, edited_math = _math_hashes(linked)
    revised_metadata = deepcopy(metadata)
    revised_metadata["figure_sources"] = {"1": "revision-response.pdf", "2": "response.pdf"}
    revised = output / "revised"
    if args.resume and (revised / "workflow.json").is_file():
        revision = _read(revised / "workflow.json")
        if not revision.get("native_build_finished"):
            raise RuntimeError("The existing revision stage is incomplete; inspect its workflow before recovery")
    else:
        revision = revise_master(linked, revised, journal=CNS, submission_copy=True,
                                 metadata=revised_metadata, artwork_root=output)
    _checkpoint(output, _stage_paths(revised))
    final = _check_stage(revised, revision, template, items, revised=True)
    _, revised_math = _math_hashes(revised / "editable-master.docx")
    checks = {**{"initial_" + key: value for key, value in initial["checks"].items()},
              **{"revised_" + key: value for key, value in final["checks"].items()}}
    checks.update({
        "both_native_pipeline_stages_finished": build["native_build_finished"] and revision["native_build_finished"],
        "initial_source_has_a_deliberately_stale_linked_citation": _citation_items(source)[0]["title"] == STALE_TITLE,
        "revision_source_has_a_deliberately_stale_new_citation": any(
            entry["key"] == args.zotero_item_keys[1] and entry["title"] == STALE_TITLE for entry in _citation_items(linked)),
        "native_sdk_edit_changed_only_the_second_master_equation": len(initial_hashes) == len(edited_math) == 3
            and initial_hashes[0] == edited_math[0] and initial_hashes[1] != edited_math[1] and initial_hashes[2] == edited_math[2],
        "revision_kept_exact_edited_native_equation_payloads": edited_math == revised_math,
        "table_cell_edit_survived_revision": "7.875" in " ".join(page["text"] for page in final["documents"]["editable-master"]["pages"]),
        "original_figure_number_and_page_changed_after_insertions": bool(initial["caption_pages"][f"Figure 1 {FIGURE_CAPTION}"])
            and bool(final["caption_pages"][f"Figure 2 {FIGURE_CAPTION}"])
            and final["caption_pages"][f"Figure 2 {FIGURE_CAPTION}"][0] > initial["caption_pages"][f"Figure 1 {FIGURE_CAPTION}"][0],
        "input_template_master_and_revision_sources_unchanged": all(digest(Path(path).read_bytes()) == checksum
            for path, checksum in manifest["input_sha256"].items()) and digest(master.read_bytes()) == master_hash
            and digest(linked.read_bytes()) == edited_hash,
    })
    report = {"pass": all(checks.values()), "checks": checks, "visual_reviewed": False,
              "initial": initial, "revised": final, "input_sha256": manifest["input_sha256"],
              "math_edit": {"method": "MathType native SDK insertion into the second equation on a new private Word copy",
                            "original": initial_hashes, "edited": edited_math, "revised": revised_math},
              "boundaries": ["Synthetic complete paper and project-owned example template, not a real manuscript acceptance.",
                             "Existing Zotero metadata was read only; no item/library changes were requested.",
                             "The script renders all pages but visual review must be recorded by the caller."]}
    evidence = output / "release-acceptance-evidence.json"
    _write(evidence, report)
    print(json.dumps({"pass": report["pass"], "checks": checks, "evidence": str(evidence)}, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
