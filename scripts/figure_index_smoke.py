"""Verify Word figure and table indexes through build, revision and C&S end pages."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
import shutil
import sys

from lxml import etree
import pymupdf

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from paperflow.mathtype import audit_docx_math
from paperflow.package import NS, digest, part_fields, q, read_package, text_run, write_package, xml
from paperflow.pipeline import CNS, revise_master, run_pipeline


HEADING = "3 Fixture output"
PAGE_INSERTION = "This fixture page shifts the objects to test index page numbers."
CAPTIONS = {
    "Figure": ("response", "Synthetic response used for software validation", "Revised synthetic response"),
    "Table": ("results", "Synthetic fixture values", "Revised synthetic fixture values"),
}
INDEX_LABEL = re.compile(r'\\c\s+"?(Figure|Table)"?(?=\s|$)', re.I)
FRONT_MATTER = """**Contents**

[[TOC]]

**List of Figures**

[[LOF]]

**List of Tables**

[[LOT]]

```{=openxml}
<w:p><w:r><w:br w:type="page"/></w:r></w:p>
```

"""


def _make_source(examples, output):
    source = output / "manuscript.md"
    draft = (examples / source.name).read_text(encoding="utf-8")
    heading = "# 1 Introduction"
    if draft.count(heading) != 1:
        raise RuntimeError("Expected one introductory heading in the C&S fixture")
    source.write_text(draft.replace(heading, FRONT_MATTER + heading), encoding="utf-8")
    for name in ("response.png", "response.pdf"):
        shutil.copyfile(examples / name, output / name)
    return source


def _page_break_before(paragraph):
    properties = paragraph.find(q("w:pPr"))
    if properties is None:
        properties = etree.Element(q("w:pPr"))
        paragraph.insert(0, properties)
    marker = properties.find(q("w:pageBreakBefore"))
    if marker is None:
        # The fixture only has a paragraph style before pageBreakBefore.
        style = properties.find(q("w:pStyle"))
        properties.insert(0 if style is None else 1, etree.Element(q("w:pageBreakBefore")))
    else:
        marker.set(q("w:val"), "1")


def _revise_copy(master, revised):
    parts = read_package(master)
    document = xml(parts["word/document.xml"])
    body = document.find(q("w:body"))
    direct_paragraphs = body.findall(q("w:p"))
    heading, = [p for p in direct_paragraphs if "".join(t.text or "" for t in p.iter(q("w:t"))) == HEADING]
    insertion = etree.Element(q("w:p"))
    insertion.append(text_run(PAGE_INSERTION))
    _page_break_before(insertion)
    heading.addprevious(insertion)
    _page_break_before(heading)
    for kind, (bookmark, old, new) in CAPTIONS.items():
        seq = re.compile(rf"^\s*SEQ\s+{kind}\b", re.I)
        ref = re.compile(rf"^\s*REF\s+{bookmark}\b", re.I)
        caption, = [p for p in direct_paragraphs if any(seq.match(code) for code in part_fields(p)[0])]
        text, = [t for t in caption.iter(q("w:t")) if old in (t.text or "")]
        text.text = text.text.replace(old, new)
        reference, = [p for p in direct_paragraphs if any(ref.match(code) for code in part_fields(p)[0])]
        stale_result, = reference.xpath('./w:r[w:fldChar[@w:fldCharType="separate"]]'
                                        '/following-sibling::w:r[1]/w:t', namespaces=NS)
        if stale_result.text != "1":
            raise RuntimeError(f"Expected a numeric REF cache for {bookmark!r}")
        stale_result.text = "99"
    parts["word/document.xml"] = etree.tostring(document, encoding="UTF-8", xml_declaration=True)
    write_package(parts, revised)


def _fields(workflow, submission):
    phases = {"submission_word_fields"} if submission else {"word_fields", "word_fields_after_zotero"}
    return next(step["evidence"]["snapshot"]["fields"] for step in reversed(workflow["steps"])
                if step["stage"] in phases)


def _check_outputs(workflow, output, revised):
    docs, checks = {}, {}
    selected_names = {kind: values[2 if revised else 1] for kind, values in CAPTIONS.items()}
    for name in ("editable-master", "submission-copy"):
        submission = name == "submission-copy"
        fields = _fields(workflow, submission)
        toc_fields = [field for field in fields if field["kind"] == "TOC" and field["story"] == 1]
        indexes = {kind: [field for field in toc_fields if (match := INDEX_LABEL.search(field["code"]))
                          and match[1].casefold() == kind.casefold()] for kind in CAPTIONS}
        toc = [field for field in toc_fields if not INDEX_LABEL.search(field["code"])]
        with pymupdf.open(output / (name + ".pdf")) as pdf:
            pages = [" ".join(page.get_text().split()) for page in pdf]
            if not submission:
                pdf[0].get_pixmap(dpi=150, alpha=False).save(output / "navigation-page.png")
        observed = {}
        for kind, label in selected_names.items():
            full_caption = f"{kind} 1 {label}"
            rows = [line.strip() for field in indexes[kind] for line in field["result"].splitlines() if line.strip()]
            body_pages = [number for number, text in enumerate(pages, 1) if number > 1 and full_caption in text]
            page = body_pages[0] if body_pages else 0
            rendered_row = re.search(rf"{re.escape(full_caption)}\s*\.*\s+{page}\b", pages[0])
            correct_result = rows == [f"{full_caption}\t{page}"]
            bookmark = CAPTIONS[kind][0]
            references = [field["result"].strip() for field in fields
                          if re.match(rf"^\s*REF\s+{bookmark}\b", field["code"], re.I)]
            observed[kind] = {
                "index_field": indexes[kind],
                "index_rows": rows,
                "caption_pages": body_pages,
                "reference_results": references,
                "rendered_index_row": bool(rendered_row),
            }
            checks[f"{name}_{kind.lower()}_index_matches_rendered_page_once"] = (
                len(indexes[kind]) == 1 and not indexes[kind][0]["locked"]
                and correct_result and bool(rendered_row) and len(body_pages) == (2 if submission else 1)
            )
            checks[f"{name}_{kind.lower()}_ref_updated"] = references == ["1"]
        checks[f"{name}_contents_kept_separate"] = (len(toc) == 1 and HEADING in toc[0]["result"]
                                                     and all(label not in toc[0]["result"] for label in selected_names.values()))
        math = audit_docx_math(output / (name + ".docx"))
        docs[name] = {
            "indexes": observed, "contents": toc, "pdf_pages": len(pages),
            "math_hashes": [entry["mtef_sha256"] for entry in math["objects"]],
        }
    checks["end_pages_verified"] = next(step["evidence"]["pagination_verified"]
                                         for step in workflow["steps"] if step["stage"] == "journal_object_pages")
    checks["zotero_native_refresh"] = any(step["stage"] == "zotero"
                                           and step["evidence"]["zotero_native_refresh_verified"]
                                           for step in workflow["steps"])
    return {"checks": checks, "documents": docs, "workflow": str(output / "workflow.json")}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    args = parser.parse_args()
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    examples = PROJECT / "examples" / CNS
    template = PROJECT / "templates" / f"{CNS}.docx"
    source = _make_source(examples, output)
    metadata = json.loads((examples / "metadata.json").read_text(encoding="utf-8"))
    before = digest(source.read_bytes())
    built = output / "built"
    original_report = run_pipeline(source, template, built, journal=CNS, submission_copy=True,
                                   table_policy="three-line", metadata=metadata)
    original = _check_outputs(original_report, built, revised=False)
    master = built / "editable-master.docx"
    master_before = digest(master.read_bytes())
    draft = output / "word-revision.docx"
    _revise_copy(master, draft)
    draft_before = digest(draft.read_bytes())
    refreshed = output / "refreshed"
    revised_report = revise_master(draft, refreshed, journal=CNS, submission_copy=True,
                                   metadata=metadata, artwork_root=output)
    revised = _check_outputs(revised_report, refreshed, revised=True)
    first_master = original["documents"]["editable-master"]
    final_master = revised["documents"]["editable-master"]
    checks = {
        **{f"original_{name}": passed for name, passed in original["checks"].items()},
        **{f"revised_{name}": passed for name, passed in revised["checks"].items()},
        "both_index_page_numbers_changed": all(
            first_master["indexes"][kind]["caption_pages"] and final_master["indexes"][kind]["caption_pages"]
            and final_master["indexes"][kind]["caption_pages"][0] > first_master["indexes"][kind]["caption_pages"][0]
            for kind in CAPTIONS
        ),
        "mathtype_objects_unchanged": len(first_master["math_hashes"]) == 2 and all(
            details["math_hashes"] == first_master["math_hashes"]
            for result in (original, revised) for details in result["documents"].values()
        ),
        "input_template_master_and_revision_unchanged": (
            original_report["template_unchanged"] and revised_report["source_unchanged"]
            and before == digest(source.read_bytes()) and master_before == digest(master.read_bytes())
            and draft_before == digest(draft.read_bytes())
        ),
    }
    evidence = output / "figure-index-evidence.json"
    report = {"pass": all(checks.values()), "checks": checks, "original": original, "revised": revised}
    evidence.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"pass": report["pass"], "checks": checks, "evidence": str(evidence)}, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
