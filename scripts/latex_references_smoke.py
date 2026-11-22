"""Verify native LaTeX labels, rich captions, MathType, and Zotero in Word."""
from __future__ import annotations

import argparse
from collections import Counter
import json
from math import exp, sin
from pathlib import Path
import re
import sys

from PIL import Image, ImageDraw
import pymupdf

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from scripts.zotero_fixture import SECOND_DOI, fixture_key
from paperflow.citations import _package_state, _payload
from paperflow.citekeys import plan_citekeys
from paperflow.mathtype import audit_docx_math
from paperflow.package import q, digest, read_package, without_bidi_controls, xml
from paperflow.pipeline import CNS, CNS_STYLE, revise_master, run_pipeline
from paperflow.word import private_word


TITLE = "Real-time neural network based semiactive model predictive control of structural vibrations"
FIGURE = "Free oscillation response"
TABLE = "Sensor readings"
BIBTEX = """@article{pf-vibrations,
  title = {Real-time neural network based semiactive model predictive control of structural vibrations},
  doi = {10.1016/j.compstruc.2022.106899}
}
"""
LATEX = r"""\documentclass{article}
\usepackage{amsmath,graphicx,hyperref,natbib}
\title{LaTeX references fixture}
\begin{document}
\maketitle

\begin{abstract}
This local fixture demonstrates editable labels, references and captions.
\end{abstract}

Keywords: Word fields; Structural dynamics

\section{Results}

See Figure~\ref{fig:response} and Table~\ref{tab:readings}.
The equation reference is \eqref{eq:response}.
The named reference is \autoref{fig:response}.
The literature reference is \citep{pf-vibrations}.

\begin{figure}
\centering
\includegraphics[width=2.6in]{response.png}
\caption{Free \emph{oscillation} response}\label{fig:response}
\end{figure}

\begin{table}
\caption{Sensor \textbf{readings}}\label{tab:readings}
\begin{tabular}{lr}
Sensor & Value \\
A & 1.25 \\
B & 2.50 \\
\end{tabular}
\end{table}

\begin{equation}
u=0
\end{equation}

\begin{equation}
\label{eq:response}
x+1=2
\end{equation}

\bibliography{references}
\end{document}
"""


def _make_figure(path):
    image = Image.new("RGB", (3550, 1430), "white")
    draw = ImageDraw.Draw(image)
    draw.line(((165, 140), (165, 1230), (3390, 1230)), fill="#202730", width=8)
    draw.line(((165, 685), (3390, 685)), fill="#d4dbe0", width=6)
    curve = [
        (x + 165, 685 - 480 * exp(-x / 1400) * sin(x / 135))
        for x in range(0, 3226, 5)
    ]
    draw.line(curve, fill="#145486", width=14)
    image.save(path, dpi=(1000, 1000))


def _read_stage(workflow, name):
    return next(step["evidence"] for step in reversed(workflow["steps"]) if step["stage"] == name)


def _visible(element):
    return "".join(node.text or "" for node in element.iter(q("w:t")))


def _caption_markup(parts, caption, token, property_name, character_style):
    document = xml(parts["word/document.xml"])
    for paragraph in document.iter(q("w:p")):
        if caption not in _visible(paragraph):
            continue
        for run in paragraph.iter(q("w:r")):
            if token not in _visible(run):
                continue
            properties = run.find(q("w:rPr"))
            if properties is None:
                continue
            applied = properties.find(q("w:" + property_name))
            style = properties.find(q("w:rStyle"))
            if (
                applied is not None and applied.get(q("w:val"), "1") not in {"0", "false", "off"}
                or style is not None and style.get(q("w:val")) == character_style
            ):
                return True
    return False


def _insert_earlier_sequence(master, revised, bookmark):
    """Edit only a new copy, using the same native Word field an author would add."""
    with private_word() as app:
        doc = app.Documents.Open(FileName=str(master), ReadOnly=True, AddToRecentFiles=False, Visible=False)
        try:
            doc.SaveAs2(FileName=str(revised), FileFormat=12, AddToRecentFiles=False)
            # Add the new sequence in a regular paragraph before both existing equations.
            content = doc.Content
            content.Collapse(Direction=1)  # wdCollapseStart
            label = "Earlier equation number "
            start = int(content.Start)
            content.InsertBefore(label + "00.\r")
            position = start + len(label)
            doc.Fields.Add(Range=doc.Range(Start=position, End=position + 2),
                           Type=-1, Text=r"SEQ Equation \* ARABIC", PreserveFormatting=True)
            if not doc.Bookmarks.Exists(bookmark):
                raise RuntimeError("Word edit removed the original LaTeX equation bookmark")
            doc.Save()
        finally:
            doc.Close(SaveChanges=0)


def _check_build(folder, workflow, targets, *, revised):
    master = folder / "editable-master.docx"
    parts = read_package(master)
    snapshot = _read_stage(workflow, "word_fields_after_zotero")["snapshot"]
    fields = [field for field in snapshot["fields"] if field["story"] == 1]
    seq = {kind: [] for kind in ("Figure", "Table", "Equation")}
    references = Counter()
    locked = []
    for field in fields:
        result = without_bidi_controls(field["result"]).strip()
        if match := re.match(r'^\s*SEQ\s+"?(Figure|Table|Equation)"?(?=\s|$)', field["code"], re.I):
            seq[match[1].title()].append(result)
        elif match := re.match(r'^\s*REF\s+"?([^"\s]+)"?(?=\s|$)', field["code"], re.I):
            references[match[1].casefold(), result] += 1
        else:
            continue
        if field["locked"]:
            locked.append(field["code"])

    number = "3" if revised else "2"
    expected_refs = Counter({
        (targets["fig:response"].casefold(), "1"): 2,
        (targets["tab:readings"].casefold(), "1"): 1,
        (targets["eq:response"].casefold(), number): 1,
    })
    pages = []
    with pymupdf.open(folder / "editable-master.pdf") as pdf:
        for index, page in enumerate(pdf, 1):
            png = folder / f"review-page-{index}.png"
            page.get_pixmap(dpi=130, alpha=False).save(png)
            pages.append({"number": index, "text": page.get_text(), "png": str(png)})
    rendered = " ".join(" ".join(page["text"] for page in pages).split())

    citation_state = _package_state(parts)["fields"]
    citations = [entry for entry in citation_state if entry["owner"] == "ITEM"]
    bibliography = [entry for entry in citation_state if entry["owner"] == "BIBL"]
    item_keys = [
        item["uris"][0].rsplit("/", 1)[-1]
        for entry in citations for item in _payload(entry["code"], "ITEM")["citationItems"]
    ]
    math = audit_docx_math(master)
    hashes = [obj["mtef_sha256"] for obj in math["objects"]]
    checks = {
        "native_word_build_finished": workflow["native_build_finished"],
        "figure_table_and_equation_sequence_results": seq == {
            "Figure": ["1"], "Table": ["1"], "Equation": ["1", "2", "3"] if revised else ["1", "2"],
        },
        "four_native_references_resolve_to_their_targets": references == expected_refs and not locked,
        "source_caption_emphasis_retained": (
            _caption_markup(parts, FIGURE, "oscillation", "i", "Emphasis")
            and _caption_markup(parts, TABLE, "readings", "b", "Strong")
        ),
        "captions_and_reference_text_rendered": all(text in rendered for text in (
            f"Figure 1 {FIGURE}", f"Table 1 {TABLE}",
            "See Figure 1 and Table 1.",
            f"The equation reference is ({number}).",
            "The named reference is Figure 1.",
        )),
        "two_equations_remain_native_mathtype": (
            math["all_mathtype_objects_valid"] and math["native_mathtype_objects"] == len(hashes) == 2
            and all(hashes) and math["omml_equations"] == 0
        ),
        "zotero_refreshed_with_real_link_and_bibliography": (
            workflow["citation_style"]["native_refresh_verified"] and item_keys == [fixture_key(SECOND_DOI)]
            and workflow["citation_style"]["saved_style_id"] == CNS_STYLE and len(bibliography) == 1
            and TITLE.casefold() in " ".join(bibliography[0]["result"].casefold().split())
            and "The literature reference is [1]." in rendered
        ),
        "no_unresolved_fields_or_markers": workflow["structural_audit"]["structural_pass"],
    }
    if revised:
        checks["earlier_sequence_rendered"] = "Earlier equation number 1." in rendered
    return {"checks": checks, "sequences": seq, "references": [
        {"bookmark": bookmark, "number": result, "count": count}
        for (bookmark, result), count in references.items()
    ], "mathtype_payloads": hashes, "pages": pages}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    output = parser.parse_args().output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    template = PROJECT / "templates" / f"{CNS}.docx"
    source, bibliography, figure = (output / name for name in ("manuscript.tex", "references.bib", "response.png"))
    source.write_text(LATEX, encoding="utf-8")
    bibliography.write_text(BIBTEX, encoding="utf-8")
    _make_figure(figure)
    plan = plan_citekeys(source, bibliography)
    if plan["issues"] or plan["citation_map"] != {"pf-vibrations": fixture_key(SECOND_DOI)}:
        raise RuntimeError(f"Expected the existing fixture citation to be matched uniquely: {plan}")
    (output / "citekey-plan.json").write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    originals = {path: digest(path.read_bytes()) for path in (source, bibliography, template, figure)}

    built = output / "built"
    metadata = {"figure_sources": {"1": {"path": figure.name, "kind": "line", "layout": "single_column"}}}
    initial_workflow = run_pipeline(source, template, built, journal=CNS, citekey_plan=plan,
                                    table_policy="three-line", submission_copy=True, metadata=metadata)
    reference_import = _read_stage(initial_workflow, "import")["latex_references"]
    targets = {entry["label"]: entry["bookmark"] for entry in reference_import["targets"] if entry["label"]}
    if set(targets) != {"fig:response", "tab:readings", "eq:response"}:
        raise RuntimeError(f"Not all labeled source targets were migrated: {reference_import}")
    initial = _check_build(built, initial_workflow, targets, revised=False)
    placement = _read_stage(initial_workflow, "journal_object_pages")
    frozen = _package_state(read_package(built / "submission-copy.docx"))

    master = built / "editable-master.docx"
    originals[master] = digest(master.read_bytes())
    revision = output / "word-edited.docx"
    _insert_earlier_sequence(master, revision, targets["eq:response"])
    originals[revision] = digest(revision.read_bytes())
    refreshed = output / "refreshed"
    revised_workflow = revise_master(revision, refreshed, journal=CNS)
    final = _check_build(refreshed, revised_workflow, targets, revised=True)
    checks = {
        **{f"initial_{key}": value for key, value in initial["checks"].items()},
        **{f"revised_{key}": value for key, value in final["checks"].items()},
        "recognized_three_labels_and_four_references_without_review": (
            len(reference_import["references"]) == 4 and not reference_import["issues"]
            and not initial_workflow["input_needs_semantic_review"]
        ),
        "journal_submission_copies_figure_and_table_on_distinct_end_pages": (
            placement["pagination_verified"] and len(placement["objects"]) == 2
        ),
        "submission_copy_freezes_zotero_without_losing_word_fields": (
            not frozen["fields"] and initial_workflow["submission_audit"]["structural_pass"]
            and any(field["instruction"].lstrip().startswith("REF ")
                    for field in initial_workflow["submission_audit"]["fields"])
        ),
        "same_math_payloads_after_renumbering": initial["mathtype_payloads"] == final["mathtype_payloads"],
        "all_inputs_unchanged": all(digest(path.read_bytes()) == value for path, value in originals.items()),
    }
    report = {"pass": all(checks.values()), "checks": checks, "reference_import": reference_import,
              "journal_placement": placement, "initial": initial, "revised": final}
    evidence = output / "latex-references-evidence.json"
    evidence.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"pass": report["pass"], "checks": checks, "evidence": str(evidence)}, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
