"""Verify LaTeX directories and page references before and after a Word revision."""
from __future__ import annotations

import argparse
from collections import Counter
import json
from pathlib import Path
import re
import sys

from docx import Document
import pymupdf

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from scripts.zotero_fixture import SECOND_DOI, fixture_key
from paperflow.citations import _package_state, _payload
from paperflow.citekeys import plan_citekeys
from paperflow.mathtype import audit_docx_math
from paperflow.package import digest, field_runs, read_package, without_bidi_controls
from paperflow.pipeline import CNS, revise_master, run_pipeline
from paperflow.word import private_word
from scripts.latex_references_smoke import BIBTEX, _make_figure, _read_stage


TITLE = "LaTeX navigation fixture"
INSERTED = "A page inserted during Word revision."
CHAPTERS = ("Navigation overview", "Response observations", "Calibration observations", "Formula observations", "References")
FIGURE = "Single oscillator response"
TABLE = "Sensor calibration values"
BOOKMARK_WITNESSES = {
    "sec:overview": ("Navigation overview", "The heading target is on page"),
    "fig:response": ("Response observations", f"Figure 1 {FIGURE}"),
    "tab:calibration": ("Calibration observations", f"Table 1 {TABLE}"),
    "eq:response": ("Formula observations", "The equation target appears here."),
}
INDEX_LABEL = re.compile(r'\\c\s+"?(Figure|Table)"?(?=\s|$)', re.I)
FIELD_TARGET = re.compile(r'^\s*(?:PAGEREF|REF)\s+"?([^"\s]+)"?(?=\s|$)', re.I)
LATEX = r"""\documentclass{article}
\usepackage{amsmath,graphicx,hyperref,natbib}
\title{LaTeX navigation fixture}
\begin{document}
\maketitle

\begin{abstract}
This sample checks editable directories and cross references after repagination.
\end{abstract}

Keywords: Word fields; Structural dynamics

\tableofcontents

\listoffigures

\listoftables

\section{Navigation overview}\label{sec:overview}
The heading target is on page~\pageref{sec:overview}.
The figure target is on page~\pageref{fig:response}.
The table target is on page~\pageref{tab:calibration}.
The equation target is on page~\pageref{eq:response}.
See Figure~\ref{fig:response} and equation~\eqref{eq:response}.
The supporting reference is \citep{pf-vibrations}.

\section{Response observations}
\begin{figure}
\centering
\includegraphics[width=2.6in]{response.png}
\caption{Single oscillator response}\label{fig:response}
\end{figure}

\section{Calibration observations}
\begin{table}
\caption{Sensor calibration values}\label{tab:calibration}
\begin{tabular}{lr}
Sensor & Value \\
A & 1.25 \\
B & 2.50 \\
\end{tabular}
\end{table}

\section{Formula observations}
The equation target appears here.

\begin{equation}\label{eq:response}
x+1=2
\end{equation}

\bibliography{references}
\end{document}
"""


def _insert_body_page(master, revision):
    with private_word() as app:
        doc = app.Documents.Open(FileName=str(master), ReadOnly=True, AddToRecentFiles=False, Visible=False)
        try:
            doc.SaveAs2(FileName=str(revision), FileFormat=12, AddToRecentFiles=False)
            for index in range(1, doc.Paragraphs.Count + 1):
                paragraph = doc.Paragraphs.Item(index)
                if without_bidi_controls(str(paragraph.Range.Text)).strip() == CHAPTERS[0]:
                    start = int(paragraph.Range.Start)
                    break
            else:
                raise RuntimeError("Expected the first body heading after the directories")
            doc.Range(Start=start, End=start).InsertBefore(INSERTED + "\r")
            inserted = doc.Paragraphs.Item(index)
            inserted.Range.Style = doc.Styles.Item(-1)  # wdStyleNormal
            inserted.Format.PageBreakBefore = True
            doc.Save()
        finally:
            doc.Close(SaveChanges=0)


def _pdf_pages(folder):
    with pymupdf.open(folder / "editable-master.pdf") as pdf:
        pages = []
        for index, page in enumerate(pdf, 1):
            png = folder / f"review-page-{index}.png"
            page.get_pixmap(dpi=110, alpha=False).save(png)
            pages.append({"page": index, "png": str(png), "text": " ".join(page.get_text().split())})
    return pages


def _inspect(folder, workflow, targets, *, revised, template_toc):
    pages = _pdf_pages(folder)
    first_body_page = 3 if revised else 2
    body = [page for page in pages if page["page"] >= first_body_page]
    directory_page = pages[0]["text"]
    target_pages = {label: [page["page"] for page in body if all(text in page["text"] for text in witnesses)]
                    for label, witnesses in BOOKMARK_WITNESSES.items()}
    chapter_pages = {title: [page["page"] for page in body if title in page["text"]] for title in CHAPTERS}
    expected_refs = Counter((targets[label].casefold(), str(locations[0]))
                            for label, locations in target_pages.items() if len(locations) == 1)
    expected_index = {
        "TOC": [f"{title}\t{chapter_pages[title][0]}" for title in CHAPTERS if len(chapter_pages[title]) == 1],
        "LOF": [f"Figure 1 {FIGURE}\t{page}" for page in target_pages["fig:response"]],
        "LOT": [f"Table 1 {TABLE}\t{page}" for page in target_pages["tab:calibration"]],
    }
    actual_refs, number_refs = Counter(), Counter()
    index_fields = {name: [] for name in expected_index}
    locked = []
    names = {name.casefold() for name in targets.values()}
    snapshot = _read_stage(workflow, "word_fields_after_zotero")["snapshot"]
    for field in snapshot["fields"]:
        if field["story"] != 1:
            continue
        result = without_bidi_controls(field["result"]).strip()
        if field["kind"] == "TOC":
            caption = INDEX_LABEL.search(field["code"])
            key = {"figure": "LOF", "table": "LOT"}[caption[1].casefold()] if caption else "TOC"
            index_fields[key].append([line.strip() for line in result.splitlines() if line.strip()])
        elif (target := FIELD_TARGET.match(field["code"])) and target[1].casefold() in names:
            collection = actual_refs if field["kind"] == "PAGEREF" else number_refs
            collection[target[1].casefold(), result] += 1
        else:
            continue
        if field["locked"]:
            locked.append(field["code"])

    rendered_body = " ".join(page["text"] for page in body)
    labels_by_object = {"Heading": "sec:overview", "Figure": "fig:response", "Table": "tab:calibration", "Equation": "eq:response"}
    master = folder / "editable-master.docx"
    math = audit_docx_math(master)
    fields = _package_state(read_package(master))["fields"]
    linked = [item["uris"][0].rsplit("/", 1)[-1] for field in fields if field["owner"] == "ITEM"
              for item in _payload(field["code"], "ITEM")["citationItems"]]
    checks = {
        "native_build_and_structural_audit": workflow["native_build_finished"] and workflow["structural_audit"]["structural_pass"],
        "one_live_directory_per_command_with_pdf_page_values": (
            len(expected_index["TOC"]) == len(CHAPTERS) and len(expected_index["LOF"]) == len(expected_index["LOT"]) == 1
            and index_fields == {key: [value] for key, value in expected_index.items()} and not locked
        ),
        "directory_labels_and_entries_visible_on_directory_page": all(text in directory_page for text in (
            TITLE, "Contents", "List of Figures", "List of Tables", CHAPTERS[0], FIGURE, TABLE,
        )),
        "all_four_pageref_fields_match_rendered_target_pages": len(expected_refs) == 4 and actual_refs == expected_refs,
        "all_four_page_references_visible_in_pdf": all(
            len(target_pages[label]) == 1 and f"The {kind.casefold()} target is on page {target_pages[label][0]}." in rendered_body
            for kind, label in labels_by_object.items()
        ),
        "ordinary_number_references_still_live": number_refs == Counter({
            (targets["fig:response"].casefold(), "1"): 1, (targets["eq:response"].casefold(), "1"): 1,
        }) and "See Figure 1 and equation (1)." in rendered_body,
        "native_mathtype_and_zotero_preserved": (
            math["all_mathtype_objects_valid"] and math["native_mathtype_objects"] == 1 and not math["omml_equations"]
            and workflow["citation_style"]["native_refresh_verified"] and linked == [fixture_key(SECOND_DOI)]
            and "The supporting reference is [1]." in rendered_body
        ),
    }
    if revised:
        checks["word_revision_inserted_separate_body_page"] = len(pages) > 1 and (
            INSERTED in pages[1]["text"] and not any(title in pages[1]["text"] for title in CHAPTERS)
        )
    if template_toc:
        checks["template_directory_title_and_position_preserved"] = (
            "Contents from template" in directory_page and directory_page.index("Contents from template") < directory_page.index(TITLE)
        )
    return {"checks": checks, "pages": pages, "target_pages": target_pages, "index_fields": index_fields,
            "page_references": [{"bookmark": name, "value": value, "count": count}
                                for (name, value), count in actual_refs.items()],
            "math_payloads": [item["mtef_sha256"] for item in math["objects"]]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument("--template-toc", action="store_true", help="Reuse a directory already placed in the template")
    args = parser.parse_args()
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    template_source = PROJECT / "templates" / f"{CNS}.docx"
    source, bibliography, image, template = (output / name for name in (
        "manuscript.tex", "references.bib", "response.png", "template.docx",
    ))
    source.write_text(LATEX, encoding="utf-8")
    bibliography.write_text(BIBTEX, encoding="utf-8")
    _make_figure(image)
    reference = Document(template_source)
    reference.styles["Heading 1"].paragraph_format.page_break_before = True
    if args.template_toc:
        slot = next(p for p in reference.paragraphs if p.text == "{{PAPERFLOW_BODY}}")
        slot.insert_paragraph_before("Contents from template")
        slot.insert_paragraph_before()._p.extend(field_runs(r' TOC \o "1-3" \h \z \u '))
    reference.save(template)
    plan = plan_citekeys(source, bibliography)
    if plan["issues"] or plan["citation_map"] != {"pf-vibrations": fixture_key(SECOND_DOI)}:
        raise RuntimeError(f"Expected the existing fixture citation to be matched uniquely: {plan}")
    (output / "citekey-plan.json").write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    originals = {path: digest(path.read_bytes()) for path in (source, bibliography, image, template, template_source)}

    built = output / "built"
    workflow = run_pipeline(source, template, built, journal=CNS, citekey_plan=plan, table_policy="three-line")
    imported = _read_stage(workflow, "import")["latex_references"]
    targets = {target["label"]: target["bookmark"] for target in imported["targets"] if target["label"]}
    if targets.keys() != BOOKMARK_WITNESSES.keys():
        raise RuntimeError(f"Unexpected LaTeX page reference targets: {imported}")
    initial = _inspect(built, workflow, targets, revised=False, template_toc=args.template_toc)

    master, revision = built / "editable-master.docx", output / "word-edited.docx"
    originals[master] = digest(master.read_bytes())
    _insert_body_page(master, revision)
    originals[revision] = digest(revision.read_bytes())
    refreshed = output / "refreshed"
    revised_workflow = revise_master(revision, refreshed, journal=CNS)
    revised = _inspect(refreshed, revised_workflow, targets, revised=True, template_toc=args.template_toc)
    toc_mode = "template" if args.template_toc else "source"
    checks = {
        **{f"initial_{key}": value for key, value in initial["checks"].items()},
        **{f"revised_{key}": value for key, value in revised["checks"].items()},
        "recognized_all_commands_without_review": (
            [(index["kind"], index["mode"]) for index in imported["indexes"]] == [("TOC", toc_mode), ("LOF", "source"), ("LOT", "source")]
            and Counter(ref["command"] for ref in imported["references"]) == {"pageref": 4, "ref": 1, "eqref": 1}
            and not imported["issues"] and not workflow["input_needs_semantic_review"]
        ),
        "each_rendered_target_moved_one_page": all(
            len(initial["target_pages"][label]) == 1
            and revised["target_pages"][label] == [initial["target_pages"][label][0] + 1] for label in targets
        ) and len(revised["pages"]) == len(initial["pages"]) + 1,
        "formula_payload_unchanged": initial["math_payloads"] == revised["math_payloads"],
        "all_inputs_unchanged": all(digest(path.read_bytes()) == value for path, value in originals.items()),
    }
    evidence = output / "latex-navigation-evidence.json"
    report = {"pass": all(checks.values()), "checks": checks, "import": imported, "initial": initial, "revised": revised}
    evidence.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"pass": report["pass"], "checks": checks, "evidence": str(evidence)}, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
