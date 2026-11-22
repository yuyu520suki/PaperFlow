"""Assemble one DOCX into three distinct template sections in real Word."""
from __future__ import annotations

import argparse
from copy import deepcopy
import json
from pathlib import Path
import sys

import pymupdf
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from scripts.zotero_fixture import FIRST_DOI, fixture_key
from paperflow.layout import inspect_layout
from paperflow.mathtype import audit_docx_math
from paperflow.package import digest, field_runs
from paperflow.pipeline import CNS_STYLE, run_pipeline
from paperflow.recovery import _word_equations


PROJECT = Path(__file__).resolve().parents[1]
TITLE = "Three Section Manuscript Fixture"
AUTHOR = "Cover Author Probe"
ABSTRACT_TEXT = "The scanned response is discussed in the article section."


def _bookmark(paragraph, name, identifier, text=None):
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(identifier))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(identifier))
    paragraph._p.append(start)
    if text:
        paragraph.add_run(text)
    paragraph._p.append(end)


def _furniture(section, label, number, fmt="decimal"):
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False
    section.header.paragraphs[0].text = f"{label} Header"
    paragraph = section.footer.paragraphs[0]
    paragraph.text = f"{label} Footer "
    for run in field_runs(" PAGE ", "0"):
        paragraph._p.append(run)
    numbering = section._sectPr.find(qn("w:pgNumType"))
    if numbering is None:
        numbering = OxmlElement("w:pgNumType")
        section._sectPr.insert_element_before(numbering, "w:cols", "w:docGrid")
    numbering.set(qn("w:start"), str(number))
    numbering.set(qn("w:fmt"), fmt)


def _make_template(path):
    document = Document()
    document.styles["Normal"].font.size = Pt(10)
    cover_line = document.styles.add_style("Cover Line", WD_STYLE_TYPE.PARAGRAPH)
    cover_line.font.name = "Times New Roman"
    cover_line.font.size = Pt(16)
    cover_line.font.bold = True
    for name in ("Title", "Heading 1"):
        properties = document.styles[name]._element.find(qn("w:pPr"))
        border = properties.find(qn("w:pBdr")) if properties is not None else None
        if border is not None:
            properties.remove(border)
    cover = document.sections[0]
    cover.left_margin = Cm(2.5)
    _furniture(cover, "Cover", 1)
    document.add_paragraph("Fixed cover content")
    document.add_paragraph("{{PAPERFLOW_COVER}}")
    abstract = document.add_section(WD_SECTION.NEW_PAGE)
    abstract.left_margin = Cm(3.4)
    _furniture(abstract, "Abstract", 1, "lowerRoman")
    _bookmark(document.add_paragraph(), "AbstractSlot", 11)
    document.add_paragraph("Fixed abstract closing text")
    main = document.add_section(WD_SECTION.NEW_PAGE)
    main.orientation = WD_ORIENT.LANDSCAPE
    main.page_width, main.page_height = main.page_height, main.page_width
    main.left_margin = Cm(2)
    main.top_margin = Cm(1.8)
    main.bottom_margin = Cm(1.6)
    _furniture(main, "Article", 7)
    document.add_paragraph("{{PAPERFLOW_BODY}}")
    document.save(path)
    return cover_line.style_id


def _block_count(document):
    return len(document._element.body) - 1


def _make_source(path):
    document = Document()
    _bookmark(document.add_paragraph(style="Title"), "coverTitle", 15, TITLE)
    document.add_paragraph(AUTHOR)
    document.add_section(WD_SECTION.NEW_PAGE)
    first = _block_count(document)
    document.add_heading("Abstract", 1)
    document.add_paragraph(ABSTRACT_TEXT)
    document.add_paragraph(f"See Eq. ([[REF model]]) and Figure [[REF chart]]. A prior study [[CITE {fixture_key(FIRST_DOI)}]] is a citation fixture.")
    document.add_section(WD_SECTION.NEW_PAGE)
    second = _block_count(document)
    document.add_paragraph("[[TOC]]")
    document.add_heading("1 Layout methods", 1)
    document.add_paragraph("Cover bookmark: [[REF coverTitle]]. The data appear in Table [[REF settings]].")
    equation = document.add_paragraph()
    equation._p.append(deepcopy(_word_equations(["y=x^2+1"])[0]))
    equation.add_run(" [[EQNO model]]")
    document.add_paragraph("[[CAPTION Table settings|Test settings]]")
    table = document.add_table(2, 2)
    for row, values in zip(table.rows, (("Case", "Force"), ("A", "220"))):
        for cell, text in zip(row.cells, values):
            cell.text = text
    document.add_picture(str(PROJECT / "examples" / "computers-and-structures" / "response.png"), width=Cm(8))
    document.add_paragraph("[[CAPTION Figure chart|Test response]]")
    document.add_heading("References", 1)
    document.add_paragraph("[[BIBLIOGRAPHY]]")
    document.save(path)
    return first, second, _block_count(document)


def _render(folder):
    pages = []
    with pymupdf.open(folder / "editable-master.pdf") as pdf:
        for n, page in enumerate(pdf, 1):
            target = folder / f"review-page-{n}.png"
            page.get_pixmap(dpi=130, alpha=False).save(target)
            pages.append({"text": page.get_text(), "size": tuple(page.rect)[2:], "image": str(target)})
    return pages


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    output = parser.parse_args().output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    source, template = output / "source.docx", output / "template.docx"
    custom_style = _make_template(template)
    first, second, last = _make_source(source)
    report = inspect_layout(source, template)
    plan = report["plan"]
    expected_slots = ["{{PAPERFLOW_COVER}}", "bookmark:AbstractSlot", "{{PAPERFLOW_BODY}}"]
    if [item["slot"] for item in plan["assignments"]] != expected_slots:
        raise RuntimeError("The layout inventory changed destination order")
    for assignment, (start, end) in zip(plan["assignments"], [(1, first), (first + 1, second), (second + 1, last)]):
        assignment.update(start=start, end=end)
    plan["assignments"][0]["role_style_map"] = {"body": custom_style}
    (output / "layout.json").write_text(json.dumps(plan, indent=2), encoding="utf-8")
    folder = output / "built"
    workflow = run_pipeline(source, template, folder, layout_plan=plan, csl_style=CNS_STYLE, table_policy="three-line")
    stages = {step["stage"]: step["evidence"] for step in workflow["steps"]}
    snapshot = stages["word_fields_after_zotero"]["snapshot"]
    pages = _render(folder)
    fields = [field for field in snapshot["fields"] if field["kind"] == "REF"]
    by_target = {field["code"].split()[1]: field["result"].strip() for field in fields}
    checks = {
        "source_unchanged": digest(source.read_bytes()) == plan["source_sha256"],
        "template_unchanged": digest(template.read_bytes()) == plan["template_sha256"],
        "geometry": stages["template_insert"]["section_geometry_matches_template"] and snapshot["section_count"] == 3,
        "cover_role_override": any(p["text"].strip() == AUTHOR and p["font"]["Size"] == 16 for p in snapshot["paragraphs"]),
        "cross_slot_fields": by_target == {"model": "1", "chart": "1", "coverTitle": TITLE, "settings": "1"},
        "toc_contains_article": any(f["kind"] == "TOC" and "1 Layout methods" in f["result"] for f in snapshot["fields"]),
        "true_zotero_refresh": stages["zotero"]["zotero_native_refresh_verified"],
        "native_mathtype": audit_docx_math(folder / "editable-master.docx")["native_mathtype_objects"] == 1,
        "three_pages": len(pages) == 3,
    }
    if len(pages) == 3:
        a, b, c = (page["text"] for page in pages)
        checks.update(cover_content=all(text in a for text in ("Cover Header", "Cover Footer 1", AUTHOR, "Fixed cover content")),
                      abstract_content=all(text in b for text in ("Abstract Header", "Abstract Footer i", ABSTRACT_TEXT,
                                                                 "Fixed abstract closing text", "See Eq. (1) and Figure 1")),
                      article_content=all(text in c for text in ("Article Header", "Article Footer 7", "Test settings",
                                                                 "Test response", "220", TITLE, "References")),
                      landscape=pages[2]["size"][0] > pages[2]["size"][1])
    result = {"checks": checks, "pages": pages, "field_results": by_target, "pass": all(checks.values())}
    (output / "layout-evidence.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(checks, ensure_ascii=False, indent=2))
    return 0 if result["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
