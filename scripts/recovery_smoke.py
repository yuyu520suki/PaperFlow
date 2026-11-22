"""Verify reviewed scan recovery all the way to native Word and MathType."""
from __future__ import annotations

import argparse
from io import BytesIO
import json
from pathlib import Path
import shutil
import sys

import pymupdf
from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from scripts.zotero_fixture import FIRST_DOI, fixture_key
from paperflow.mathtype import audit_docx_math
from paperflow.package import digest, q, read_package, xml
from paperflow.pipeline import CNS, run_pipeline
from paperflow.recovery import apply_recovery, prepare_recovery


PROJECT = Path(__file__).resolve().parents[1]
EXAMPLE = PROJECT / "examples" / CNS
TITLE = "Scanned Mechanics Workflow Fixture"
ABSTRACT = "This synthetic manuscript checks the recovery of manuscript structure."
TABLE = [["Case", "Force", "Current"], ["Steel A", "220", "1.5"], ["Steel B", "215", "1.4"], ["Steel C", "230", "1.6"]]
FORMULA = "y = x^2 + 1"
FIGURE_CAPTION = "Figure 1. Synthetic response fixture"
CROP = (85, 545, 1085, 1045)
SCAN_SIZE = (1250, 1200)


def _make_fixture(output):
    scan = Image.new("RGB", SCAN_SIZE, "white")
    draw = ImageDraw.Draw(scan)
    regular = ImageFont.truetype(r"C:\Windows\Fonts\arial.ttf", 36)
    heading = ImageFont.truetype(r"C:\Windows\Fonts\arialbd.ttf", 38)
    draw.text((90, 45), "2 Results", font=heading, fill="black")
    draw.text((90, 105), "Table 1. Bench settings", font=regular, fill="black")
    for row, y in zip(TABLE, (175, 245, 315, 385)):
        for text, x in zip(row, (105, 470, 835)):
            draw.text((x, y), text, font=regular, fill="black")
    draw.text((90, 465), FORMULA, font=regular, fill="black")
    with Image.open(EXAMPLE / "response.png") as graph:
        scan.paste(graph.convert("RGB").resize((CROP[2] - CROP[0], CROP[3] - CROP[1])), CROP[:2])
    draw.text((90, 1080), FIGURE_CAPTION, font=regular, fill="black")
    image = BytesIO()
    scan.save(image, format="PNG")
    path = output / "mixed-source.pdf"
    with pymupdf.open() as pdf:
        first = pdf.new_page(width=625, height=600)
        for y, line in enumerate((TITLE, "Abstract", ABSTRACT, "Contents", "1 Methods",
                                  "The sample refers to Table 1 and Eq. (1). Prior work [1] is a citation fixture.",
                                  "References", "[1] Citation placeholder for the local Zotero record."), start=1):
            first.insert_text((45, 48 + 29 * y), line, fontsize=13)
        second = pdf.new_page(width=625, height=600)
        second.insert_image(second.rect, stream=image.getvalue())
        pdf.save(path)
    shutil.copyfile(EXAMPLE / "response.pdf", output / "fixture-graph.pdf")
    return path


def _find_line(lines, match):
    hits = [i for i, line in enumerate(lines) if match(line)]
    if len(hits) != 1:
        raise RuntimeError(f"Expected one fixture line, found {hits}")
    return hits[0]


def _review(intake, plan, output):
    """Apply explicit decisions from this known synthetic source, as a reviewer would."""
    lines = intake["source_lines"]
    by_text = lambda text: _find_line(lines, lambda line: line["text"].strip() == text)
    regions = {"title": by_text(TITLE), "abstract_heading": by_text("Abstract"), "abstract": by_text(ABSTRACT),
               "contents": by_text("Contents"), "methods": by_text("1 Methods"), "references": by_text("References")}
    regions["intro"] = _find_line(lines, lambda line: "Prior work [1]" in line["text"])
    regions["bibliography"] = _find_line(lines, lambda line: "Citation placeholder" in line["text"])
    regions["results"] = _find_line(lines, lambda line: line["page"] == 2 and "2 Results" in line["text"])
    regions["table_caption"] = _find_line(lines, lambda line: line["page"] == 2 and "Bench settings" in line["text"])
    recognized_math = [i for i, line in enumerate(lines) if line["page"] == 2 and "=" in line["text"] and "+ 1" in line["text"]]
    if len(recognized_math) > 1:
        raise RuntimeError("Fixture has more than one candidate OCR formula")
    if recognized_math:
        regions["formula"] = recognized_math[0]
    regions["figure_caption"] = _find_line(lines, lambda line: line["page"] == 2 and "Synthetic response fixture" in line["text"])
    table = next((block for block in plan["blocks"] if block["kind"] == "table" and block["rows"] == TABLE), None)
    if table is None:
        raise RuntimeError("The scan did not produce the expected geometry-based editable table proposal")
    table_lines = table.get("line_indices", list(range(table["start_line"], table["end_line"] + 1)))
    table_lines = sorted(set(table_lines + [regions["table_caption"]]))
    figure_box = [CROP[0] / SCAN_SIZE[0], CROP[1] / SCAN_SIZE[1], CROP[2] / SCAN_SIZE[0], CROP[3] / SCAN_SIZE[1]]

    def block(key, kind, **fields):
        line = regions[key]
        return {"start_line": line, "end_line": line, "kind": kind, **fields}

    formula = {"start_line": None, "end_line": None, "kind": "equation", "page": 2,
               "bbox": [90 / SCAN_SIZE[0], 465 / SCAN_SIZE[1], 425 / SCAN_SIZE[0], 513 / SCAN_SIZE[1]],
               "latex": "y=x^2+1", "label": "model"}
    if recognized_math:
        formula.update(start_line=recognized_math[0], end_line=recognized_math[0])
    blocks = [block("title", "title"), block("abstract_heading", "heading", level=2), block("abstract", "paragraph"),
              block("contents", "paragraph", text="[[TOC]]"), block("methods", "heading", level=1),
              block("intro", "paragraph", text="The sample refers to Table [[REF settings]], Eq. ([[REF model]]) and "
                    f"Figure [[REF response]]. Prior work [[CITE {fixture_key(FIRST_DOI)}]] is a citation fixture."),
              block("results", "heading", level=1),
              {"start_line": table_lines[0], "end_line": table_lines[-1], "line_indices": table_lines,
               "kind": "table", "label": "settings", "caption": "Bench settings", "rows": TABLE, "header_rows": 1},
              formula,
              block("figure_caption", "figure", page=2, bbox=figure_box, caption="Synthetic response fixture", label="response"),
              block("references", "heading", level=1), block("bibliography", "paragraph", text="[[BIBLIOGRAPHY]]")]
    assigned = set(regions.values()).union(table_lines)
    for i, line in enumerate(lines):
        if i in assigned:
            continue
        bbox = line.get("bbox")
        if line["page"] != 2 or not bbox or not figure_box[1] <= (bbox[1] + bbox[3]) / 2 <= figure_box[3]:
            raise RuntimeError(f"Unhandled fixture line {i}: {line['text']!r}")
        blocks.append({"start_line": i, "end_line": i, "kind": "omit", "reason": "Graph labels are retained in the figure crop"})
    reviewed = {**plan, "blocks": blocks}
    output.write_text(json.dumps(reviewed, ensure_ascii=False, indent=2), encoding="utf-8")
    return len(table_lines) - 1, len(lines)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    source = _make_fixture(output)
    before = digest(source.read_bytes())
    prepared = prepare_recovery(source, output / "intake", ocr_language="en-US")
    print(json.dumps(prepared, ensure_ascii=False), flush=True)
    intake = json.loads(Path(prepared["intake"]).read_text(encoding="utf-8"))
    plan = json.loads(Path(prepared["plan"]).read_text(encoding="utf-8"))
    reviewed = output / "reviewed-plan.json"
    proposed_table_lines, lines = _review(intake, plan, reviewed)
    recovered = output / "recovered.docx"
    applied = apply_recovery(prepared["intake"], reviewed, recovered)
    print(json.dumps(applied, ensure_ascii=False), flush=True)
    final = output / "built"
    built = run_pipeline(recovered, PROJECT / "templates" / f"{CNS}.docx", final, journal=CNS,
                         table_policy="three-line", submission_copy=True,
                         metadata={"abstract": ABSTRACT, "keywords": ["Workflow"], "figure_sources": {"1": "fixture-graph.pdf"}})
    master = final / "editable-master.docx"
    stage = {step["stage"]: step["evidence"] for step in built["steps"]}
    body = xml(read_package(master)["word/document.xml"])
    texts = ["".join(cell.itertext()) for cell in body.iter(q("w:tc"))]
    checks = {
        "original_unchanged": before == digest(source.read_bytes()),
        "all_source_lines_reviewed": applied["source_lines_accounted_for"] == lines,
        "ocr_table_source_lines": proposed_table_lines,
        "text_and_scan_adapters": {line["adapter"] for line in intake["source_lines"]} == {"pdf-text-layer", "windows-ocr"},
        "editable_cells": texts == [cell for row in TABLE for cell in row],
        "one_native_mathtype": audit_docx_math(master)["native_mathtype_objects"] == 1,
        "real_zotero_refresh": stage["zotero"]["zotero_native_refresh_verified"],
        "expected_fields": {item["kind"] for item in stage["word_fields_after_zotero"]["snapshot"]["fields"]}
                           >= {"TOC", "SEQ", "REF", "PAGE", "ADDIN"},
        "figure_and_table_at_end": stage["journal_object_pages"]["pagination_verified"],
        "original_vector_staged": stage["journal_artwork"]["all_figures_ready"],
    }
    report = {"checks": checks, "review": applied, "plan": prepared, "master": str(master), "pass": all(checks.values())}
    (output / "recovery-evidence.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(checks, ensure_ascii=False, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
