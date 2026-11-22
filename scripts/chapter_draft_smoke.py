"""Prove native Word migration of typed chapter-numbered captions and callouts."""
from __future__ import annotations

import argparse
from collections import Counter
import json
from pathlib import Path
import re
import subprocess
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Cm
import pymupdf

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from paperflow.mathtype import audit_docx_math
from paperflow.package import digest
from paperflow.pipeline import CNS_STYLE
from scripts.chapter_captions_smoke import (
    BIDI_FORMATTING,
    CAPTION_STYLE,
    CHAPTERS,
    INDEX_LABEL,
    REFERENCE,
    _insert_chapter,
    _make_numbering,
    _make_template,
    _set_numbering,
)


def _numbered_objects(offset=0):
    for chapter_number, (_, objects) in enumerate(CHAPTERS, 1):
        counts = Counter()
        for kind, _, title in objects:
            counts[kind] += 1
            separator, punctuation = ("-", ".") if kind == "Figure" else (".", ":")
            yield {
                "kind": kind,
                "original_number": f"{chapter_number}{separator}{counts[kind]}",
                "expected_number": f"{chapter_number + offset}-{counts[kind]}",
                "punctuation": punctuation,
                "title": title,
            }


def _add_equation(document):
    paragraph = document.add_paragraph()
    math = OxmlElement("m:oMath")
    run = OxmlElement("m:r")
    text = OxmlElement("m:t")
    text.text = "x+1=2"
    run.append(text)
    math.append(run)
    paragraph._p.append(math)
    paragraph.add_run(" [[EQNO eqAlpha]]")
    document.add_paragraph("See Equation [[REF eqAlpha]].")


def _make_source(path, with_zotero):
    document = Document()
    heading = document.styles["Heading 1"]
    _make_numbering(document, first=1, paragraph_style=heading.style_id, multilevel=True)
    _set_numbering(heading._element.get_or_add_pPr())
    for title, marker in (("Contents", "[[TOC]]"), ("List of Figures", "[[LOF]]"), ("List of Tables", "[[LOT]]")):
        document.add_paragraph().add_run(title).bold = True
        document.add_paragraph(marker)
    picture = PROJECT / "examples" / "computers-and-structures" / "response.png"
    records = iter(_numbered_objects())
    for chapter_index, (chapter, objects) in enumerate(CHAPTERS):
        document.add_heading(chapter, 1)
        for _ in objects:
            record = next(records)
            kind, number, punctuation, title = (record[name] for name in (
                "kind", "original_number", "punctuation", "title"
            ))
            if kind == "Figure":
                document.add_picture(str(picture), width=Cm(6))
            document.add_paragraph(f"{kind} {number}{punctuation} {title}", style="Caption")
            if kind == "Table":
                table = document.add_table(1, 2)
                table.cell(0, 0).text = title
                table.cell(0, 1).text = "1.25"
            document.add_paragraph(f"See {kind} {number}.")
        if chapter_index == 0:
            _add_equation(document)
            if with_zotero:
                document.add_paragraph("Prior work [1] accompanies the validation.")
    if with_zotero:
        document.add_heading("References", level=2)
        document.add_paragraph("[[BIBLIOGRAPHY]]")
    document.save(path)


def _cli(*args):
    result = subprocess.run(
        [sys.executable, "-X", "utf8", "-m", "paperflow", *map(str, args)],
        cwd=PROJECT,
        encoding="utf-8",
        capture_output=True,
    )
    if result.returncode:
        raise RuntimeError(f"PaperFlow {args[0]} exited {result.returncode}\n{result.stdout}{result.stderr}")
    return json.loads(result.stdout) if result.stdout.strip() else None


def _read_json(path):
    return json.loads(path.read_text(encoding="utf-8"))


def _review_plan(plan, source_hash):
    expected = list(_numbered_objects())
    captions = [edit for edit in plan["edits"] if edit["operation"] == "caption"]
    references = [edit for edit in plan["edits"] if edit["operation"] == "cross_reference"]
    if (plan["source_sha256"] != source_hash or plan["issues"] or len(captions) != len(expected)
            or plan["chapter_caption_policy"] != {"level": 1, "separator": "-"}):
        raise RuntimeError(f"Plan did not match the reviewed fixture: {plan}")
    keys = {}
    for record, caption in zip(expected, captions):
        key = (record["kind"], record["original_number"])
        if key != (caption["sequence"], caption["source_text"]) or caption["target_number"] != record["expected_number"]:
            raise RuntimeError(f"Unexpected caption plan edit: {caption}")
        keys[key] = caption["bookmark"]
    planned_references = Counter((edit["bookmark"], edit["source_text"]) for edit in references)
    if planned_references != Counter((bookmark, number) for (_, number), bookmark in keys.items()):
        raise RuntimeError(f"Unexpected cross-reference plan edits: {references}")
    return keys


def _read_pages(folder):
    with pymupdf.open(folder / "editable-master.pdf") as pdf:
        pages = []
        for page_index, page in enumerate(pdf, 1):
            page.get_pixmap(dpi=130, alpha=False).save(folder / f"review-page-{page_index}.png")
            pages.append(" ".join(page.get_text().split()))
    return pages


def _check_stage(folder, keys, *, offset, with_zotero):
    workflow = _read_json(folder / "workflow.json")
    snapshot = next(step["evidence"]["snapshot"] for step in reversed(workflow["steps"])
                    if step["stage"] in {"word_fields", "word_fields_after_zotero"})
    fields = snapshot["fields"]
    expected_refs = {"eqAlpha": f"{1 + offset}-1"}
    expected_indexes = {"Figure": [], "Table": []}
    for item in _numbered_objects(offset):
        kind = item["kind"]
        expected_refs[keys[(kind, item["original_number"])]] = item["expected_number"]
        expected_indexes[kind].append(
            f"{kind} {item['expected_number']}{item['punctuation']} {item['title']}"
        )
    refs, indexes = {}, {"Figure": [], "Table": []}
    for field in fields:
        value = field["result"].translate(BIDI_FORMATTING)
        if field["kind"] == "REF" and (match := REFERENCE.match(field["code"])):
            refs[match[1]] = value.strip()
        if field["kind"] == "TOC" and (match := INDEX_LABEL.search(field["code"])):
            indexes[match[1].title()].extend(row.strip() for row in value.splitlines() if row.strip())

    pages = _read_pages(folder)
    expected_rows = {"Figure": [], "Table": []}
    for kind, entries in expected_indexes.items():
        for entry in entries:
            matches = [index for index, text in enumerate(pages, 1) if index > 1 and entry in text]
            if len(matches) == 1:
                expected_rows[kind].append(f"{entry}\t{matches[0]}")
    all_captions = {caption for entries in expected_indexes.values() for caption in entries}
    styles = [paragraph["style"] for paragraph in snapshot["paragraphs"]
              if paragraph["text"].translate(BIDI_FORMATTING).strip() in all_captions]
    math = audit_docx_math(folder / "editable-master.docx")
    math_hashes = [obj.get("mtef_sha256") for obj in math["objects"]]
    chapter_migration = [step["evidence"] for step in workflow["steps"]
                         if step["stage"] == "chapter_migration_values"]
    checks = {
        "word_native_build_complete": workflow["native_build_finished"],
        "chapter_migration_verified_on_build": offset > 0 or len(chapter_migration) == 1,
        "word_refs_match_full_caption_numbers": refs == expected_refs,
        "word_lists_match_captions_and_pdf_page_numbers": (
            indexes == expected_rows and all(len(expected_rows[kind]) == len(expected_indexes[kind]) for kind in indexes)
        ),
        "rendered_references_match": all(any(
            f"See {item['kind']} {item['expected_number']}." in page for page in pages[1:]
        ) for item in _numbered_objects(offset)),
        "rendered_equation_ref_matches": any(f"See Equation {expected_refs['eqAlpha']}." in page for page in pages[1:]),
        "template_caption_style_kept": styles == [CAPTION_STYLE] * len(all_captions),
        "editable_mathtype_equation": len(math_hashes) == math["native_mathtype_objects"] == 1
                                     and not math["omml_equations"] and bool(math_hashes[0]),
    }
    if with_zotero:
        checks["native_zotero_refresh"] = any(step["stage"] == "zotero"
                                              and step["evidence"]["zotero_native_refresh_verified"]
                                              for step in workflow["steps"])
    return {"checks": checks, "refs": refs, "index_rows": indexes, "caption_styles": styles,
            "mtef_sha256": math_hashes, "chapter_migration": chapter_migration, "pages": pages}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument("--zotero-item-key", help="Link one old numbered citation to this local Zotero item")
    args = parser.parse_args()
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    source = output / "typed-manuscript.docx"
    template = output / "template.docx"
    plan_path = output / "reviewed-draft-plan.json"
    roles = output / "roles.json"
    _make_source(source, bool(args.zotero_item_key))
    style_id = _make_template(template)
    roles.write_text(json.dumps({"caption": style_id}, indent=2) + "\n", encoding="utf-8")
    originals = {path: digest(path.read_bytes()) for path in (source, template, roles)}
    citation_options, style_options = [], []
    if args.zotero_item_key:
        citation_map = output / "citation-map.json"
        citation_map.write_text(json.dumps({"1": args.zotero_item_key}), encoding="utf-8")
        originals[citation_map] = digest(citation_map.read_bytes())
        citation_options = ["--citation-map", citation_map]
        style_options = ["--csl-style", CNS_STYLE]
    _cli("plan-draft", source, "--caption-chapter-level", "1", "--report", plan_path, *citation_options)
    keys = _review_plan(_read_json(plan_path), originals[source])
    originals[plan_path] = digest(plan_path.read_bytes())

    built = output / "built"
    _cli("run", source, "--template", template, "--role-style-map", roles, "--draft-plan", plan_path,
         "--output-dir", built, *style_options)
    initial = _check_stage(built, keys, offset=0, with_zotero=bool(args.zotero_item_key))

    master = built / "editable-master.docx"
    originals[master] = digest(master.read_bytes())
    revision = output / "revised-input.docx"
    _insert_chapter(master, revision)
    originals[revision] = digest(revision.read_bytes())
    refreshed = output / "refreshed"
    _cli("revise-master", revision, "--output-dir", refreshed)
    final = _check_stage(refreshed, keys, offset=1, with_zotero=bool(args.zotero_item_key))

    checks = {
        **{f"initial_{key}": passed for key, passed in initial["checks"].items()},
        **{f"revised_{key}": passed for key, passed in final["checks"].items()},
        "all_inputs_unchanged": all(digest(path.read_bytes()) == value for path, value in originals.items()),
        "same_mathtype_payload_after_renumbering": initial["mtef_sha256"] == final["mtef_sha256"],
    }
    evidence = output / "chapter-draft-evidence.json"
    report = {"pass": all(checks.values()), "checks": checks, "initial": initial, "revised": final}
    evidence.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"pass": report["pass"], "checks": checks, "evidence": str(evidence)}, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
