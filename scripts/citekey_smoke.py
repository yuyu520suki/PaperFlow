"""Verify native Word output from Markdown and LaTeX citekeys and one BibTeX file."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
import subprocess
import sys

import pymupdf

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from scripts.zotero_fixture import FIRST_DOI, SECOND_DOI, fixture_key
from paperflow.citations import _package_state, _payload
from paperflow.mathtype import audit_docx_math
from paperflow.package import digest, read_package, without_bidi_controls
from paperflow.pipeline import CNS, CNS_STYLE
from paperflow.zotero import preference_style_id, read_document_preferences


SECOND_TITLE = "Real-time neural network based semiactive model predictive control of structural vibrations"
BIBTEX = """@article{pf-alpha,
  title = {基于自适应伸缩因子的变论域模糊PID振动控制方法},
  doi = {10.6052/j.issn.1000-4750.2020.11.0786}
}
@article{pf-beta,
  title = {Real-time neural network based semiactive model predictive control of structural vibrations},
  doi = {10.1016/j.compstruc.2022.106899}
}
"""
MARKDOWN = """---
title: Citekey migration example
bibliography: references.bib
---

## Abstract

This is a local conversion fixture using synthetic manuscript text.

Keywords: Citation migration; Structural dynamics

# Methods

A grouped reference cites both records [@pf-alpha; @pf-beta].

A page-specific follow-up cites the first record [@pf-alpha, p. 21].

$$x+1=2$$
"""
LATEX = r"""\documentclass{article}
\usepackage{natbib}
\title{Citekey migration example}
\begin{document}
\maketitle
\begin{abstract}
This is a local conversion fixture using synthetic manuscript text.
\end{abstract}

Keywords: Citation migration; Structural dynamics

\section{Methods}
A grouped reference cites both records \citep{pf-alpha,pf-beta}.

A page-specific follow-up cites the first record \citep[21]{pf-alpha}.

\[
x+1=2
\]
\bibliography{references}
\end{document}
"""


def _cli(*args):
    result = subprocess.run(
        [sys.executable, "-X", "utf8", "-m", "paperflow", *map(str, args)],
        cwd=PROJECT, capture_output=True, encoding="utf-8",
    )
    if result.returncode:
        raise RuntimeError(f"PaperFlow {args[0]} exited {result.returncode}\n{result.stdout}{result.stderr}")


def _read_json(path):
    return json.loads(path.read_text(encoding="utf-8"))


def _review_plan(plan, source_hash):
    if plan["citation_map"] != {"pf-alpha": fixture_key(FIRST_DOI), "pf-beta": fixture_key(SECOND_DOI)} or plan["issues"] or len(plan["clusters"]) != 2:
        raise RuntimeError(f"Citekey plan did not match the fixture: {plan}")
    first, second = plan["clusters"]
    if (
        plan["source_sha256"] != source_hash
        or [cluster["index"] for cluster in (first, second)] != [1, 2]
        or [[item["citekey"] for item in cluster["items"]] for cluster in (first, second)]
        != [["pf-alpha", "pf-beta"], ["pf-alpha"]]
        or second["items"][0].get("locator") != "21"
    ):
        raise RuntimeError(f"Citekey plan did not retain source citation details: {plan}")


def _compact(text):
    return "".join(without_bidi_controls(text).split())


def _check_build(output, *, with_submission):
    master = output / "editable-master.docx"
    workflow = _read_json(output / "workflow.json")
    state = _package_state(read_package(master))
    citations = [field for field in state["fields"] if field["owner"] == "ITEM"]
    bibliographies = [field for field in state["fields"] if field["owner"] == "BIBL"]
    items = [_payload(field["code"], "ITEM")["citationItems"] for field in citations]
    item_keys = [[item["uris"][0].rsplit("/", 1)[-1] for item in cluster] for cluster in items]
    citation_results = [_compact(field["result"]) for field in citations]
    math = audit_docx_math(master)
    saved_style = preference_style_id(read_document_preferences(master))
    with pymupdf.open(output / "editable-master.pdf") as pdf:
        text = _compact(" ".join(page.get_text() for page in pdf))
    checks = {
        "native_word_build_finished": workflow["native_build_finished"],
        "native_zotero_refresh_verified": workflow["citation_style"]["native_refresh_verified"],
        "selected_journal_style_saved": saved_style == CNS_STYLE,
        "citation_clusters_link_real_items": item_keys == [
            [fixture_key(FIRST_DOI), fixture_key(SECOND_DOI)], [fixture_key(FIRST_DOI)],
        ],
        "page_locator_saved_in_dynamic_citation": (
            len(items) == 2 and len(items[1]) == 1
            and items[1][0].get("locator") == "21" and items[1][0].get("label", "page") == "page"
        ),
        "visible_grouped_and_repeated_citations": (
            len(citation_results) == 2
            and re.fullmatch(r"\[1[,–-]2\]", citation_results[0]) is not None
            and citation_results[1].startswith("[1")
            and f"bothrecords{citation_results[0]}." in text
            and f"firstrecord{citation_results[1]}." in text
        ),
        "one_visible_live_bibliography": (
            len(bibliographies) == 1 and _compact(SECOND_TITLE).casefold() in _compact(
                bibliographies[0]["result"]
            ).casefold() and text.casefold().count(_compact(SECOND_TITLE).casefold()) == 1
        ),
        "editable_native_mathtype": math["native_mathtype_objects"] == 1 and math["omml_equations"] == 0,
    }
    if with_submission:
        submission = output / "submission-copy.docx"
        frozen = _package_state(read_package(submission))
        with pymupdf.open(output / "submission-copy.pdf") as pdf:
            submission_text = _compact(" ".join(page.get_text() for page in pdf))
        checks["separate_static_submission_keeps_citation_text"] = (
            not frozen["fields"] and len(citation_results) == 2
            and f"bothrecords{citation_results[0]}." in submission_text
            and f"firstrecord{citation_results[1]}." in submission_text
        )
    return {
        "checks": checks, "saved_style": saved_style, "citation_item_keys": item_keys,
        "citation_results": citation_results,
        "bibliography_result": bibliographies[0]["result"] if len(bibliographies) == 1 else None,
        "mathtype_mtef_sha256": [obj["mtef_sha256"] for obj in math["objects"]],
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    output = parser.parse_args().output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    template = PROJECT / "templates" / f"{CNS}.docx"
    bibliography = output / "references.bib"
    bibliography.write_text(BIBTEX, encoding="utf-8")
    originals = {path: digest(path.read_bytes()) for path in (template, bibliography)}
    stages = {}
    for kind, suffix, text in (("markdown", ".md", MARKDOWN), ("latex", ".tex", LATEX)):
        source = output / f"manuscript{suffix}"
        source.write_text(text, encoding="utf-8")
        originals[source] = digest(source.read_bytes())
        plan = output / f"{kind}-plan.json"
        _cli("plan-citekeys", source, "--bibliography", bibliography, "--report", plan)
        _review_plan(_read_json(plan), originals[source])
        originals[plan] = digest(plan.read_bytes())
        built = output / kind
        with_submission = kind == "markdown"
        copy_options = ["--submission-copy"] if with_submission else []
        _cli("run", source, "--template", template, "--journal", CNS, "--citekey-plan", plan,
             "--output-dir", built, *copy_options)
        stages[kind] = _check_build(built, with_submission=with_submission)
    checks = {f"{kind}_{name}": passed
              for kind, stage in stages.items() for name, passed in stage["checks"].items()}
    checks["all_inputs_and_reviewed_plans_unchanged"] = all(
        digest(path.read_bytes()) == value for path, value in originals.items()
    )
    evidence = output / "citekey-evidence.json"
    report = {"pass": all(checks.values()), "checks": checks, "stages": stages}
    evidence.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"pass": report["pass"], "checks": checks, "evidence": str(evidence)}, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
