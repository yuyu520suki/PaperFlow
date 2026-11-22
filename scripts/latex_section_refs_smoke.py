"""Verify native LaTeX heading references in numbered and unnumbered Word templates."""
from __future__ import annotations

import argparse
from collections import Counter
from copy import deepcopy
import json
from pathlib import Path
import re
import sys

from docx import Document
from lxml import etree as ET

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from paperflow.mathtype import audit_docx_math
from paperflow.package import digest, q, read_package, text_run, without_bidi_controls, write_package, xml
from paperflow.pipeline import revise_master, run_pipeline
from paperflow.word import WordAutomationError
from scripts.chapter_captions_smoke import NUMBERING_ID, _element, _set_numbering
from scripts.latex_navigation_smoke import _pdf_pages


MAIN_HEADING = "Measured response chapter"
SUB_HEADING = "Actuator arrangement subsection"
INSERTED = "Earlier numbered chapter"
TARGET_TITLES = {"sec:main": MAIN_HEADING, "sec:detail": SUB_HEADING}
FIELD_TARGET = re.compile(r'^\s*(REF|PAGEREF)\s+"?([^"\s]+)"?(?=\s|$)', re.I)
LATEX = r"""\documentclass{article}
\usepackage{hyperref}
\title{LaTeX section references fixture}
\begin{document}
\maketitle

\tableofcontents

\section{Measured response chapter}\label{sec:main}

Chapter number witness: \ref{sec:main}.

Subsection number witness: \ref{sec:detail}.

Named subsection witness: \autoref{sec:detail}.

The chapter begins on page~\pageref{sec:main}.

The subsection begins on page~\pageref{sec:detail}.

This simple formula remains editable in MathType: $x+1=2$.

\subsection{Actuator arrangement subsection}\label{sec:detail}

This subsection contains the second heading target.

\end{document}
"""


def _make_template(path, *, numbered):
    document = Document()
    abstract = _element("abstractNum", abstractNumId=42)
    abstract.append(_element("multiLevelType", val="multilevel"))
    for level, pattern in enumerate(("%1", "%1.%2"), 1):
        heading = document.styles[f"Heading {level}"]
        heading.paragraph_format.page_break_before = True
        if numbered:
            index = level - 1
            definition = _element("lvl", ilvl=index)
            definition.extend((_element("start", val=1), _element("numFmt", val="decimal"),
                               _element("pStyle", val=heading.style_id), _element("suff", val="space"),
                               _element("lvlText", val=pattern), _element("lvlJc", val="left")))
            abstract.append(definition)
            properties = heading._element.get_or_add_pPr()
            _set_numbering(properties)
            properties.numPr.get_or_add_ilvl().val = index
    if numbered:
        numbering = document.part.numbering_part.element
        numbering.insert_element_before(abstract, "w:num")
        instance = _element("num", numId=NUMBERING_ID)
        instance.append(_element("abstractNumId", val=42))
        numbering.append(instance)
    document.add_paragraph("{{PAPERFLOW_BODY}}")
    document.save(path)


def _insert_earlier_heading(master, revision):
    parts = read_package(master)
    document = xml(parts["word/document.xml"])
    body = document.find(q("w:body"))
    (existing,) = [paragraph for paragraph in body.findall(q("w:p"))
                   if "".join(run.text or "" for run in paragraph.iter(q("w:t"))) == MAIN_HEADING]
    inserted = ET.Element(q("w:p"))
    inserted.append(deepcopy(existing.find(q("w:pPr"))))
    inserted.append(text_run(INSERTED))
    existing.addprevious(inserted)
    parts["word/document.xml"] = ET.tostring(document, encoding="UTF-8", xml_declaration=True)
    write_package(parts, revision)


def _remove_heading_numbering(master, revision):
    parts = read_package(master)
    document = xml(parts["word/document.xml"])
    for paragraph in document.find(q("w:body")).findall(q("w:p")):
        text = "".join(node.text or "" for node in paragraph.iter(q("w:t")))
        if text not in TARGET_TITLES.values():
            continue
        properties = paragraph.find(q("w:pPr"))
        numbering = properties.find(q("w:numPr"))
        if numbering is not None:
            properties.remove(numbering)
        numbering = ET.SubElement(properties, q("w:numPr"))
        ET.SubElement(numbering, q("w:numId")).set(q("w:val"), "0")
    parts["word/document.xml"] = ET.tostring(document, encoding="UTF-8", xml_declaration=True)
    write_package(parts, revision)


def _inspect(folder, workflow, anchors, *, number=None):
    pages = _pdf_pages(folder)
    body = [page for page in pages if page["page"] > 1]
    rendered = " ".join(page["text"] for page in body)
    targets = {label: [page["page"] for page in body if title in page["text"]]
               for label, title in TARGET_TITLES.items()}
    expected_page_refs = Counter((anchors[label], str(page[0])) for label, page in targets.items() if len(page) == 1)
    page_refs, number_refs = Counter(), Counter()
    snapshot = next(step["evidence"]["snapshot"] for step in reversed(workflow["steps"])
                    if step["stage"] in {"word_fields", "word_fields_after_zotero"})
    heading_names = set(anchors.values())
    active_fields = []
    for field in snapshot["fields"]:
        target = FIELD_TARGET.match(field["code"])
        if field["story"] != 1 or target is None or target[2].casefold() not in heading_names:
            continue
        result = without_bidi_controls(field["result"]).strip()
        counters = page_refs if target[1].upper() == "PAGEREF" else number_refs
        counters[(target[2].casefold(), result)] += 1
        active_fields.append({**field, "result": result})
    math = audit_docx_math(folder / "editable-master.docx")
    checks = {
        "native_word_build_and_audit": workflow["native_build_finished"] and workflow["structural_audit"]["structural_pass"],
        "live_heading_pagerefs_match_word_pdf": len(expected_page_refs) == 2 and page_refs == expected_page_refs,
        "heading_pagerefs_visible_in_word_pdf": all(
            len(targets[label]) == 1 and f"The {role} begins on page {targets[label][0]}." in rendered
            for label, role in (("sec:main", "chapter"), ("sec:detail", "subsection"))
        ),
        "heading_reference_fields_unlocked": bool(active_fields) and not any(field["locked"] for field in active_fields),
        "native_mathtype_preserved": math["all_mathtype_objects_valid"] and math["native_mathtype_objects"] == 1
        and not math["omml_equations"],
    }
    if number is None:
        checks["unnumbered_template_requires_review_without_numeric_refs"] = (
            workflow["input_needs_semantic_review"] and not number_refs and "needs review" in rendered.casefold()
        )
    else:
        checks["live_heading_number_fields_match_full_word_numbers"] = number_refs == Counter({
            (anchors["sec:main"], str(number)): 1, (anchors["sec:detail"], f"{number}.1"): 2,
        })
        checks["numeric_reference_fields_use_full_context_switch"] = all(
            field["kind"] != "REF" or re.search(r"\\w(?=\s|$)", field["code"], re.I) for field in active_fields
        )
        checks["numeric_heading_references_visible_in_word_pdf"] = (
            f"Chapter number witness: {number}." in rendered
            and f"Subsection number witness: {number}.1." in rendered
            and re.search(rf"Named subsection witness: (?:sub)?section {number}\.1\.", rendered, re.I) is not None
        )
        checks["native_toc_shows_template_numbered_headings"] = (
            f"{number} {MAIN_HEADING}" in pages[0]["text"] and f"{number}.1 {SUB_HEADING}" in pages[0]["text"]
        )
    return {"checks": checks, "target_pages": targets, "fields": active_fields, "pages": pages,
            "math_payloads": [item["mtef_sha256"] for item in math["objects"]]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument("--unnumbered", action="store_true", help="Verify the review path for an unnumbered template")
    args = parser.parse_args()
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    source, template = output / "manuscript.tex", output / "template.docx"
    source.write_text(LATEX, encoding="utf-8")
    _make_template(template, numbered=not args.unnumbered)
    originals = {path: digest(path.read_bytes()) for path in (source, template)}
    built = output / "built"
    workflow = run_pipeline(source, template, built)
    imported = next(step["evidence"]["latex_references"] for step in workflow["steps"] if step["stage"] == "import")
    anchors = {target["label"]: target["bookmark"].casefold() for target in imported["targets"]
               if target["label"] in TARGET_TITLES}
    if anchors.keys() != TARGET_TITLES.keys():
        raise RuntimeError(f"Missing LaTeX heading bookmarks: {imported}")
    initial = _inspect(built, workflow, anchors, number=None if args.unnumbered else 1)
    checks = {f"initial_{name}": result for name, result in initial["checks"].items()}
    revised = None
    if not args.unnumbered:
        master = built / "editable-master.docx"
        revision = output / "word-edited.docx"
        originals[master] = digest(master.read_bytes())
        _insert_earlier_heading(master, revision)
        originals[revision] = digest(revision.read_bytes())
        refreshed = output / "refreshed"
        revised_workflow = revise_master(revision, refreshed)
        revised = _inspect(refreshed, revised_workflow, anchors, number=2)
        checks.update({f"revised_{name}": result for name, result in revised["checks"].items()})
        checks["earlier_heading_moves_both_targets_one_page"] = (
            any(INSERTED in page["text"] for page in revised["pages"][1:])
            and all(len(initial["target_pages"][label]) == 1
                    and revised["target_pages"][label] == [initial["target_pages"][label][0] + 1]
                    for label in anchors)
        )
        checks["formula_payload_unchanged_after_refresh"] = initial["math_payloads"] == revised["math_payloads"]
        checks["no_initial_heading_review"] = not workflow["input_needs_semantic_review"] and not imported["issues"]
        broken = output / "word-edited-without-numbering.docx"
        _remove_heading_numbering(refreshed / "editable-master.docx", broken)
        originals[broken] = digest(broken.read_bytes())
        rejected = output / "rejected-revision"
        try:
            revise_master(broken, rejected, submission_copy=True)
        except WordAutomationError as error:
            failure = json.loads((rejected / "workflow.json").read_text(encoding="utf-8"))
            checks["revision_detects_removed_automatic_heading_numbers"] = (
                "no automatic paragraph number" in str(error)
                and failure.get("failure", {}).get("type") == "WordAutomationError"
                and not failure.get("native_build_finished")
            )
        else:
            checks["revision_detects_removed_automatic_heading_numbers"] = False
        checks["invalid_revision_stops_before_final_exports"] = all(
            not (rejected / name).exists() for name in (
                "editable-master.docx", "editable-master.pdf", "submission-copy.docx", "submission-copy.pdf",
            )
        )
    checks["input_files_unchanged"] = all(digest(path.read_bytes()) == value for path, value in originals.items())
    evidence = output / "latex-section-refs-evidence.json"
    report = {"pass": all(checks.values()), "checks": checks, "import": imported, "initial": initial, "revised": revised}
    evidence.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"pass": report["pass"], "checks": checks, "evidence": str(evidence)}, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
