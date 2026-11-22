"""Verify direct Word outline levels through template styles and chapter references."""
from __future__ import annotations

import argparse
from collections import Counter
import json
from pathlib import Path
import re
import sys

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm
import pymupdf

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from paperflow.mathtype import audit_docx_math
from paperflow.package import digest, without_bidi_controls
from paperflow.pipeline import run_pipeline
from paperflow.semantics import plan_draft
from paperflow.word import snapshot
from scripts.chapter_captions_smoke import CAPTION_STYLE, INDEX_LABEL, REFERENCE, _make_template
from scripts.chapter_draft_smoke import _add_equation


CHAPTER_ONE = "Experimental program"
SUBHEADING = "Sensor arrangement"
BODY_OVERRIDE = "This paragraph stays in the body"
CHAPTER_TWO = "Numerical comparison"
CAPTIONS = (
    ("Figure", "1-1", "Reference signal", CHAPTER_ONE),
    ("Figure", "1-2", "Measured response", CHAPTER_ONE),
    ("Table", "2-1", "Calibration coefficients", CHAPTER_TWO),
)
OUTLINE_LEVELS = {CHAPTER_ONE: 1, SUBHEADING: 2, BODY_OVERRIDE: 10, CHAPTER_TWO: 1}


def _outline(paragraph, level):
    properties = paragraph._p.get_or_add_pPr()
    node = OxmlElement("w:outlineLvl")
    node.set(qn("w:val"), str(level))
    properties.append(node)


def _make_source(path):
    doc = Document()
    derived = doc.styles.add_style("Source Chapter Variant", WD_STYLE_TYPE.PARAGRAPH)
    derived.base_style = doc.styles["Heading 1"]
    for title, marker in (("Contents", "TOC"), ("List of Figures", "LOF"), ("List of Tables", "LOT")):
        doc.add_paragraph().add_run(title).bold = True
        doc.add_paragraph(f"[[{marker}]]")

    _outline(doc.add_paragraph(CHAPTER_ONE), 0)
    _outline(doc.add_paragraph(SUBHEADING), 1)
    picture = PROJECT / "examples" / "computers-and-structures" / "response.png"
    for index, (kind, number, title, _) in enumerate(CAPTIONS[:2]):
        if index == 1:
            _outline(doc.add_paragraph(BODY_OVERRIDE, style=derived), 9)
        doc.add_picture(str(picture), width=Cm(5))
        doc.add_paragraph(f"{kind} {number}. {title}", style="Caption")
        doc.add_paragraph(f"See {kind} {number}.")
    _add_equation(doc)

    # Word's built-in heading level takes precedence over this direct value.
    _outline(doc.add_heading(CHAPTER_TWO, level=1), 1)
    kind, number, title, _ = CAPTIONS[2]
    doc.add_paragraph(f"{kind} {number}. {title}", style="Caption")
    table = doc.add_table(rows=2, cols=2)
    for row, values in zip(table.rows, (("Constant", "Value"), ("A", "1.25"))):
        for cell, value in zip(row.cells, values):
            cell.text = value
    doc.add_paragraph(f"See {kind} {number}.")
    owners = {}
    for index, child in enumerate(doc.element.body):
        texts = child.xpath("./w:r/w:t")
        if texts and texts[0].text in (CHAPTER_ONE, CHAPTER_TWO):
            owners[texts[0].text] = index
    doc.save(path)
    return owners


def _review_plan(plan, owners):
    captions = [edit for edit in plan["edits"] if edit["operation"] == "caption"]
    actual = [(item["sequence"], item["target_number"], item["chapter_heading_body_index"]) for item in captions]
    expected = [(kind, number, owners[owner]) for kind, number, _, owner in CAPTIONS]
    targets = {item["bookmark"]: item["target_number"] for item in captions}
    references = Counter((edit["bookmark"], edit["target_number"]) for edit in plan["edits"]
                         if edit["operation"] == "cross_reference")
    if plan["issues"] or actual != expected or references != Counter(targets.items()):
        raise RuntimeError(f"Unexpected caption ownership or references in the draft plan: {plan}")
    return targets


def _read_pages(folder):
    with pymupdf.open(folder / "editable-master.pdf") as pdf:
        pages = []
        for index, page in enumerate(pdf, 1):
            path = folder / f"review-page-{index}.png"
            page.get_pixmap(dpi=130, alpha=False).save(path)
            pages.append({"page": index, "png": str(path), "text": " ".join(page.get_text().split())})
    return pages


def _witness_paragraphs(word):
    return {paragraph["text"].strip(): paragraph for paragraph in word["paragraphs"]
            if paragraph["text"].strip() in OUTLINE_LEVELS}


def _inspect_build(folder, workflow, targets):
    word = next(step["evidence"]["snapshot"] for step in reversed(workflow["steps"])
                if step["stage"] in {"word_fields", "word_fields_after_zotero"})
    paragraphs = _witness_paragraphs(word)
    styles = {1: {"heading1", "标题1", "標題1"}, 2: {"heading2", "标题2", "標題2"}, 10: {"normal", "正文"}}
    actual_refs = Counter()
    toc, indexes = [], {"Figure": [], "Table": []}
    for field in word["fields"]:
        if field["story"] != 1:
            continue
        result = without_bidi_controls(field["result"])
        if field["kind"] == "REF" and (match := REFERENCE.match(field["code"])) and not field["locked"]:
            actual_refs[match[1].strip('"'), result.strip()] += 1
        if field["kind"] == "TOC":
            rows = [line.strip() for line in result.splitlines() if line.strip()]
            if match := INDEX_LABEL.search(field["code"]):
                indexes[match[1].title()].extend(rows)
            else:
                toc.extend(rows)

    pages = _read_pages(folder)
    rendered = [page["text"] for page in pages]
    chapter_pages = {name: [i for i, page in enumerate(rendered, 1) if i > 1 and name in page]
                     for name in (CHAPTER_ONE, SUBHEADING, CHAPTER_TWO)}
    expected_toc = []
    for name, prefix in ((CHAPTER_ONE, "1 "), (SUBHEADING, ""), (CHAPTER_TWO, "2 ")):
        locations = chapter_pages[name]
        if len(locations) == 1:
            expected_toc.append(f"{prefix}{name}\t{locations[0]}")
    expected_indexes = {"Figure": [], "Table": []}
    for kind, number, title, owner in CAPTIONS:
        if len(chapter_pages[owner]) == 1:
            expected_indexes[kind].append(f"{kind} {number}. {title}\t{chapter_pages[owner][0]}")
    rendered_captions = {f"{kind} {number}. {title}" for kind, number, title, _ in CAPTIONS}
    caption_styles = [paragraph["style"] for paragraph in word["paragraphs"]
                      if without_bidi_controls(paragraph["text"]).strip() in rendered_captions]
    math = audit_docx_math(folder / "editable-master.docx")
    checks = {
        "native_word_build_and_structural_audit_passed": (
            workflow["native_build_finished"] and workflow["structural_audit"]["structural_pass"]
        ),
        "word_reports_template_heading_and_body_roles": (
            paragraphs.keys() == OUTLINE_LEVELS.keys() and all(
                paragraphs[name]["outline_level"] == level
                and "".join(paragraphs[name]["style"].casefold().split()) in styles[level]
                for name, level in OUTLINE_LEVELS.items()
            )
        ),
        "toc_contains_only_true_headings_and_current_pages": len(expected_toc) == 3 and toc == expected_toc,
        "toc_matches_rendered_contents_page": (
            all(name in rendered[0] for name in (CHAPTER_ONE, SUBHEADING, CHAPTER_TWO))
            and BODY_OVERRIDE not in rendered[0] and any(BODY_OVERRIDE in page for page in rendered[1:])
        ),
        "all_native_refs_resolve_to_full_chapter_numbers": actual_refs == Counter({**targets, "eqAlpha": "1-1"}.items()),
        "figure_and_table_indexes_match_target_chapters": (
            sum(map(len, expected_indexes.values())) == len(CAPTIONS) and indexes == expected_indexes
        ),
        "caption_template_style_applied": caption_styles == [CAPTION_STYLE] * len(CAPTIONS),
        "captions_and_ref_values_visible_in_word_pdf": (
            all(any(caption in page for page in rendered[1:]) for caption in rendered_captions)
            and all(any(f"See {kind} {number}." in page for page in rendered[1:]) for kind, number, _, _ in CAPTIONS)
            and any("See Equation 1-1." in page for page in rendered[1:])
        ),
        "word_verified_the_migration_values": any(step["stage"] == "chapter_migration_values" for step in workflow["steps"]),
        "equation_remains_editable_mathtype": (
            math["all_mathtype_objects_valid"] and math["native_mathtype_objects"] == 1 and not math["omml_equations"]
        ),
    }
    return checks, {"paragraphs": paragraphs, "toc": toc, "figure_table_indexes": indexes,
                    "references": [{"bookmark": bookmark, "value": value, "count": count}
                                   for (bookmark, value), count in actual_refs.items()], "pages": pages}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    output = parser.parse_args().output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    source, template = output / "manuscript.docx", output / "template.docx"
    owners = _make_source(source)
    caption_style_id = _make_template(template)
    plan = plan_draft(source, caption_chapter_level=1)
    plan_path = output / "draft-plan.json"
    plan_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    targets = _review_plan(plan, owners)
    originals = {path: digest(path.read_bytes()) for path in (source, template, plan_path)}
    source_paragraphs = _witness_paragraphs(snapshot(source))

    built = output / "built"
    workflow = run_pipeline(source, template, built, draft_plan=plan, role_style_map={"caption": caption_style_id})
    checks, evidence = _inspect_build(built, workflow, targets)
    checks["source_levels_confirmed_in_word"] = (
        {name: paragraph["outline_level"] for name, paragraph in source_paragraphs.items()} == OUTLINE_LEVELS
    )
    checks["all_inputs_unchanged"] = all(digest(path.read_bytes()) == value for path, value in originals.items())
    result = {"pass": all(checks.values()), "checks": checks, "source_chapter_indexes": owners,
              "source_paragraphs": source_paragraphs, **evidence}
    report = output / "source-outline-evidence.json"
    report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"pass": result["pass"], "checks": checks, "evidence": str(report)}, ensure_ascii=False, indent=2))
    return 0 if result["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
