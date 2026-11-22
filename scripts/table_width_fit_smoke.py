"""Native fixed-table width fixture, including merges, nesting and pagination.

Run with a new output directory. Word and MathType are required. Every pipeline
and native inspection runs serially; the input documents are never edited.
"""
from __future__ import annotations

import argparse
from collections import Counter
import json
from pathlib import Path
import re
import sys

from docx import Document
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt
from lxml import etree as ET
import pymupdf

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from paperflow.mathtype import audit_docx_math
from paperflow.package import NS, digest, q, read_package, xml
from paperflow.pipeline import revise_master, run_pipeline
from paperflow.word import _opened, private_word


BODY = "Table width fixture body typography."
HEADERS = {"wide": "Measured response series", "nested": "Inner control", "small": "Fixed setting"}
COLORS = {"wide": "2D5B81", "nested": "7B5B9A", "small": "2E6B46"}
CAPTION = "Measured displacement response"
SMALL_CAPTION = "Fixed acquisition reference"
DATA_ROWS = 36
NESTED_ROW = 6  # zero-based row in the outer table
WIDTH_TOLERANCE_TWIPS = 3


def _element(name, **attributes):
    element = OxmlElement("w:" + name)
    for key, value in attributes.items():
        element.set(q("w:" + key), str(value))
    return element


def _fixed_table(table, widths, color):
    table.autofit = False
    lengths = [Inches(width) for width in widths]
    for column, width in zip(table.columns, lengths, strict=True):
        column.width = width
    for row in table.rows:
        for cell, width in zip(row.cells, lengths, strict=True):
            cell.width = width
    properties = table._tbl.tblPr
    preferred = properties.find(q("w:tblW"))
    preferred.set(q("w:type"), "dxa")
    preferred.set(q("w:w"), str(sum(width.twips for width in lengths)))
    borders = _element("tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(_element(edge, val="single", sz=8, color=color))
    properties.insert_element_before(borders, "w:tblLayout", "w:tblLook")
    properties.insert_element_before(_element("jc", val="left"), "w:tblBorders", "w:tblLayout", "w:tblLook")
    # Word aligns cell text at the margin by default. Offset by the explicit
    # left cell padding so the outside border starts at the page body margin.
    properties.insert_element_before(_element("tblInd", w=72, type="dxa"), "w:tblBorders", "w:tblLayout", "w:tblLook")
    margins = _element("tblCellMar")
    for edge in ("top", "bottom", "left", "right"):
        margins.append(_element(edge, w=72 if edge in {"left", "right"} else 24, type="dxa"))
    properties.insert_element_before(margins, "w:tblLook")


def _row_properties(row, *, repeat=False, cant_split=True):
    properties = row._tr.get_or_add_trPr()
    if repeat:
        properties.append(_element("tblHeader"))
    if cant_split:
        properties.append(_element("cantSplit"))
    row.height = Pt(20)
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST


def _make_template(path):
    document = Document()
    section = document.sections[0]
    section.page_width, section.page_height = Inches(7), Inches(11)
    section.left_margin = section.right_margin = Inches(1)
    section.top_margin = section.bottom_margin = Inches(1)
    for name, font, size in (("Normal", "Calibri", 10), ("Heading 1", "Arial", 13), ("Caption", "Arial", 9)):
        document.styles[name].font.name = font
        document.styles[name].font.size = Pt(size)
    document.styles["Normal"].paragraph_format.space_after = Pt(2)
    document.add_paragraph("{{PAPERFLOW_BODY}}")
    document.save(path)


def _make_source(path):
    document = Document()
    section = document.sections[0]
    section.page_width, section.page_height = Inches(11), Inches(11)
    section.left_margin = section.right_margin = Inches(0.5)
    section.top_margin = section.bottom_margin = Inches(0.5)
    document.styles["Normal"].font.name = "Times New Roman"
    document.styles["Normal"].font.size = Pt(13)
    document.add_heading("Displacement response measurements", level=1)
    document.add_paragraph(BODY)
    document.add_paragraph("See Table [[REF WideResults]] on page [[PAGEREF WideResults]].")
    document.add_paragraph(f"[[CAPTION Table WideResults|{CAPTION}]]")
    table = document.add_table(rows=DATA_ROWS + 3, cols=3)
    _fixed_table(table, (3, 3.75, 2.25), COLORS["wide"])
    table.cell(0, 0).merge(table.cell(0, 2)).text = HEADERS["wide"]
    for cell, text in zip(table.rows[1].cells, ("Case ID", "Peak (mm)", "Observation"), strict=True):
        cell.text = text
    for index in range(DATA_ROWS):
        row = table.rows[index + 2]
        row.cells[0].text = f"Case {index + 1:03d}"
        row.cells[1].text = f"{7.5 + index * 0.125:.3f}"
        row.cells[2].text = f"Record {index + 1:03d}: stable"
    table.cell(2, 0).merge(table.cell(3, 0)).text = "Pair 001-002"
    table.cell(DATA_ROWS + 2, 0).text = "Summary"
    table.cell(DATA_ROWS + 2, 1).merge(table.cell(DATA_ROWS + 2, 2)).text = "All 36 response records retained."
    host = table.cell(NESTED_ROW, 1)
    host.text = "Acquisition settings"
    nested = host.add_table(rows=2, cols=2)
    _fixed_table(nested, (1.8, 1), COLORS["nested"])
    for row, values in zip(nested.rows, ((HEADERS["nested"], "Value"), ("Step", "0.02 s")), strict=True):
        for cell, value in zip(row.cells, values, strict=True):
            cell.text = value
        _row_properties(row)
    equation = table.cell(8, 2).paragraphs[0]
    equation.text = "Record 007; equation "
    math = ET.SubElement(equation._p, q("m:oMath"))
    ET.SubElement(ET.SubElement(math, q("m:r")), q("m:t")).text = "x+1=2"
    for index, row in enumerate(table.rows):
        _row_properties(row, repeat=index < 2, cant_split=index != 10)
    document.add_paragraph("Small reference witness: Table [[REF SmallSettings]].")
    document.add_paragraph(f"[[CAPTION Table SmallSettings|{SMALL_CAPTION}]]")
    small = document.add_table(rows=2, cols=2)
    _fixed_table(small, (1.2, 0.8), COLORS["small"])
    for row, values in zip(small.rows, ((HEADERS["small"], "Value"), ("Gain", "1.00")), strict=True):
        for cell, value in zip(row.cells, values, strict=True):
            cell.text = value
        _row_properties(row)
    document.save(path)


def _visible(element):
    return "".join(node.text or "" for node in element.iter(q("w:t")))


def _flag(parent, name):
    node = parent.find("w:trPr/w:" + name, NS)
    return node is not None and node.get(q("w:val"), "true").lower() not in {"0", "false", "off"}


def _horizontal_cell_margins(cell):
    """Resolve the fixture's explicit cell/table margins without fit helpers."""
    owner = next(cell.iterancestors(q("w:tbl")))
    margins = {"left": 108, "right": 108}
    for group in (owner.find("w:tblPr/w:tblCellMar", NS), cell.find("w:tcPr/w:tcMar", NS)):
        if group is not None:
            for side in margins:
                node = group.find(q("w:" + side))
                if node is not None:
                    margins[side] = int(node.get(q("w:w"), "0"))
    return margins


def _body_width_twips(template):
    document = xml(read_package(template)["word/document.xml"])
    section = document.find("w:body/w:sectPr", NS)
    page, margins = section.find(q("w:pgSz")), section.find(q("w:pgMar"))
    return int(page.get(q("w:w"))) - sum(int(margins.get(q("w:" + name), "0"))
                                         for name in ("left", "right", "gutter"))


def _state(path):
    document = xml(read_package(path)["word/document.xml"])
    found = {}
    for table in document.iter(q("w:tbl")):
        rows = table.findall(q("w:tr"))
        first = _visible(rows[0].find(q("w:tc")))
        kind = next((key for key, header in HEADERS.items() if first == header), None)
        if kind is None:
            continue
        cells = [row.findall(q("w:tc")) for row in rows]
        width = table.find("w:tblPr/w:tblW", NS)
        parent = table.getparent()
        host_width = parent.find("w:tcPr/w:tcW", NS) if parent.tag == q("w:tc") else None
        indent = table.find("w:tblPr/w:tblInd", NS)
        found[kind] = {
            "grid": [int(item.get(q("w:w"))) for item in table.findall("w:tblGrid/w:gridCol", NS)],
            "preferred_width": {"type": width.get(q("w:type")), "twips": int(width.get(q("w:w")))},
            "cell_widths": [[int(cell.find("w:tcPr/w:tcW", NS).get(q("w:w"))) for cell in row] for row in cells],
            "cell_text": [[_visible(cell) for cell in row] for row in cells],
            "grid_spans": [[cell.find("w:tcPr/w:gridSpan", NS).get(q("w:val"), "1")
                            if cell.find("w:tcPr/w:gridSpan", NS) is not None else "1" for cell in row] for row in cells],
            "vertical_merges": [[cell.find("w:tcPr/w:vMerge", NS).get(q("w:val"), "continue")
                                 if cell.find("w:tcPr/w:vMerge", NS) is not None else None for cell in row] for row in cells],
            "repeat_header": [_flag(row, "tblHeader") for row in rows],
            "cant_split": [_flag(row, "cantSplit") for row in rows],
            "nesting_level": 1 + sum(ancestor.tag == q("w:tbl") for ancestor in table.iterancestors()),
            "indent_twips": int(indent.get(q("w:w"))) if indent is not None else 0,
            "host_cell_width_twips": int(host_width.get(q("w:w"))) if host_width is not None else None,
            "host_cell_margins_twips": _horizontal_cell_margins(parent) if host_width is not None else None,
            "native_equations": len(list(table.iter(q("o:OLEObject")))),
        }
    if set(found) != set(HEADERS):
        raise RuntimeError(f"Expected the wide, nested and small tables, found {sorted(found)}")
    return found


def _snapshot(workflow):
    return next(step["evidence"]["snapshot"] for step in reversed(workflow["steps"])
                if step["stage"] in {"word_fields", "word_fields_after_zotero"})


def _native_geometry(path):
    """Read actual laid-out cell widths, including the nested table, serially."""
    result, seen = {}, set()
    with private_word() as word, _opened(word, path) as document:
        document.Repaginate()

        def visit(collection):
            for index in range(1, collection.Count + 1):
                table = collection.Item(index)
                identity = (int(table.Range.Start), int(table.Range.End))
                if identity in seen:
                    continue
                seen.add(identity)
                first = str(table.Cell(1, 1).Range.Text).strip("\r\x07 ")
                kind = next((key for key, header in HEADERS.items() if first == header), None)
                if kind is not None:
                    columns = 3 if kind == "wide" else 2
                    widths = [float(table.Cell(2, column).Width) for column in range(1, columns + 1)]
                    start, end = table.Range.Duplicate, table.Range.Duplicate
                    start.Collapse(1)
                    end.Collapse(0)
                    entry = {"column_widths_pt": widths, "width_pt": sum(widths),
                             "nesting_level": int(table.NestingLevel),
                             "page_start": int(start.Information(3)), "page_end": int(end.Information(3))}
                    if kind == "wide":
                        host = table.Cell(NESTED_ROW + 1, 2)
                        entry["nested_container_width_pt"] = float(host.Width)
                        entry["nested_container_padding_pt"] = [float(host.LeftPadding), float(host.RightPadding)]
                    result[kind] = entry
                # Table.Tables is Word's nested-table collection. Range.Tables
                # selects the enclosing table again and can omit its children.
                visit(table.Tables)

        visit(document.Tables)
    if set(result) != set(HEADERS):
        raise RuntimeError(f"Native geometry did not find all fixture tables: {sorted(result)}")
    return result


def _color_matches(observed, expected):
    return observed is not None and all(abs(a - b) <= 2 / 255 for a, b in zip(observed, expected, strict=True))


def _render(pdf_path, output):
    output.mkdir(exist_ok=True)
    pages = []
    colors = {key: tuple(int(value[index:index + 2], 16) / 255 for index in (0, 2, 4)) for key, value in COLORS.items()}
    with pymupdf.open(pdf_path) as pdf:
        for index, page in enumerate(pdf, 1):
            preview = output / f"page-{index:03d}.png"
            page.get_pixmap(dpi=130, colorspace=pymupdf.csRGB, alpha=False).save(preview)
            borders = {key: [list(path["rect"]) for path in page.get_drawings()
                             if _color_matches(path.get("color"), color) or _color_matches(path.get("fill"), color)]
                       for key, color in colors.items()}
            words = page.get_text("words")
            pages.append({"page": index, "size_pt": [page.rect.width, page.rect.height], "preview": str(preview),
                          "text": " ".join(page.get_text().split()), "border_rects_pt": borders,
                          "text_rects_pt": [list(word[:4]) for word in words],
                          "wide_header_occurrences": len(page.search_for(HEADERS["wide"])),
                          "caption_occurrences": len(page.search_for(CAPTION))})
    return pages


def _scaled(before, after):
    if len(before["grid"]) != len(after["grid"]) or [len(row) for row in before["cell_widths"]] != [len(row) for row in after["cell_widths"]]:
        return False
    factor = sum(after["grid"]) / sum(before["grid"])
    return all(abs(new - old * factor) <= WIDTH_TOLERANCE_TWIPS
               for old, new in zip(before["grid"] + sum(before["cell_widths"], []),
                                   after["grid"] + sum(after["cell_widths"], []), strict=True))


def _same_widths(before, after):
    return (before["preferred_width"] == after["preferred_width"] and before["grid"] == after["grid"]
            and before["cell_widths"] == after["cell_widths"])


def _preserved(before, after):
    return all(before[key] == after[key] for key in ("cell_text", "grid_spans", "vertical_merges", "repeat_header", "cant_split", "nesting_level", "indent_twips", "host_cell_margins_twips"))


def _fields_ok(snapshot, pages):
    fields = [(item["code"], item["result"].strip()) for item in snapshot["fields"]]
    caption_pages = [str(page["page"]) for page in pages if page["caption_occurrences"]]
    return (sorted(value for code, value in fields if re.match(r"\s*SEQ\s+Table\b", code, re.I)) == ["1", "2"]
            and all(any(re.match(rf"\s*{kind}\s+{label}\b", code, re.I) and value in expected for code, value in fields)
                    for kind, label, expected in (("REF", "WideResults", ["1"]), ("REF", "SmallSettings", ["2"]),
                                                  ("PAGEREF", "WideResults", caption_pages))))


def _pdf_ok(pages):
    main_pages = [page for page in pages if page["border_rects_pt"]["wide"]]
    rectangles = [rect for page in pages for values in page["border_rects_pt"].values() for rect in values]
    text = " ".join(page["text"] for page in pages)
    return (len(main_pages) >= 2 and all(page["wide_header_occurrences"] == 1 for page in main_pages)
            and bool(rectangles) and all(rect[0] >= 71 and rect[2] <= 433 for rect in rectangles)
            and all(rect[0] >= 71 and rect[2] <= 433 for page in pages for rect in page["text_rects_pt"])
            and all(abs(page["size_pt"][0] - 504) < 0.1 for page in pages)
            and f"Table 1 {CAPTION}" in text and f"Table 2 {SMALL_CAPTION}" in text
            and "See Table 1 on page 1." in text and "Small reference witness: Table 2." in text
            and "Case 036" in text and "All 36 response records retained." in text)


def _native_ok(geometry, state, expected_main_width_twips):
    main, nested, small = (geometry[key] for key in ("wide", "nested", "small"))
    padding = sum(main["nested_container_padding_pt"])
    return (abs(main["width_pt"] - expected_main_width_twips / 20) <= 0.2 and abs(small["width_pt"] - 144) <= 0.2
            and main["page_end"] > main["page_start"] and nested["nesting_level"] == 2
            and nested["width_pt"] <= main["nested_container_width_pt"] - padding - max(0, state["nested"]["indent_twips"]) / 20 + 0.2
            and all(abs(geometry[key]["width_pt"] - sum(value["grid"]) / 20) <= 0.2 for key, value in state.items()))


def _math_ok(audit):
    return audit["native_mathtype_objects"] == 1 and audit["all_mathtype_objects_valid"] and not audit["omml_equations"]


def _existing_build(output):
    built = output / "built"
    workflow = json.loads((built / "workflow.json").read_text(encoding="utf-8"))
    source, template = output / "source.docx", output / "template.docx"
    if not workflow.get("native_build_finished"):
        raise RuntimeError("Cannot resume an unfinished native build")
    for key, path in (("source", source), ("template", template)):
        if Path(workflow[key]).resolve() != path or digest(path.read_bytes()) != workflow[key + "_sha256"]:
            raise RuntimeError(f"Cannot resume: {key} differs from the completed build's original input")
    for name in ("editable-master.docx", "editable-master.pdf"):
        if not (built / name).is_file():
            raise RuntimeError(f"Cannot resume: missing completed build artifact {name}")
    return workflow


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument("--resume-existing", action="store_true",
                        help="Verify and reuse a finished build, then complete native inspection and revision")
    args = parser.parse_args()
    output = args.output_dir.resolve()
    source, template = output / "source.docx", output / "template.docx"
    if args.resume_existing:
        build = _existing_build(output)
    else:
        output.mkdir(parents=True, exist_ok=False)
        _make_source(source)
        _make_template(template)
    originals = {path: digest(path.read_bytes()) for path in (source, template)}
    before = _state(source)
    body_width_twips = _body_width_twips(template)
    expected_main_width_twips = body_width_twips - max(0, before["wide"]["indent_twips"])
    built = output / "built"
    if not args.resume_existing:
        build = run_pipeline(source, template, built)
    master = built / "editable-master.docx"
    master_hash = digest(master.read_bytes())
    after, math_after = _state(master), audit_docx_math(master)
    native_after, snapshot_after = _native_geometry(master), _snapshot(build)
    render_after = _render(built / "editable-master.pdf", output / "built-pages")
    revised = output / "revised"
    if args.resume_existing and revised.exists():
        revision = json.loads((revised / "workflow.json").read_text(encoding="utf-8"))
        if (not revision.get("native_build_finished") or Path(revision["source"]).resolve() != master
                or revision["source_sha256"] != master_hash):
            raise RuntimeError("Existing revision does not describe the unchanged completed build")
    else:
        revision = revise_master(master, revised)
    final = revised / "editable-master.docx"
    stable, math_stable = _state(final), audit_docx_math(final)
    native_stable, snapshot_stable = _native_geometry(final), _snapshot(revision)
    render_stable = _render(revised / "editable-master.pdf", output / "revised-pages")
    checks = {
        "native_build_and_revision_finished": build["native_build_finished"] and revision["native_build_finished"],
        "source_is_9in_asymmetric_fixed_table": before["wide"]["grid"] == [4320, 5400, 3240]
                                                and before["wide"]["preferred_width"] == {"type": "dxa", "twips": 12960},
        "wide_table_fitted_to_body_minus_indent_with_grid_and_cell_ratios": abs(sum(after["wide"]["grid"]) - expected_main_width_twips) <= 3
                                                            and _scaled(before["wide"], after["wide"]),
        "nested_table_fitted_proportionally_to_parent_cell": _scaled(before["nested"], after["nested"])
            and sum(after["nested"]["grid"]) <= after["nested"]["host_cell_width_twips"]
                - sum(after["nested"]["host_cell_margins_twips"].values()) - max(0, after["nested"]["indent_twips"]) + 3,
        "small_2in_table_widths_unchanged": _same_widths(before["small"], after["small"]),
        "editable_text_merges_headers_and_split_controls_preserved": all(_preserved(before[key], after[key]) for key in HEADERS),
        "revision_widths_and_structure_stable": all(_same_widths(after[key], stable[key]) and _preserved(after[key], stable[key]) for key in HEADERS),
        "actual_word_geometry_matches_fitted_package": _native_ok(native_after, after, expected_main_width_twips)
                                                       and _native_ok(native_stable, stable, expected_main_width_twips),
        "native_caption_ref_and_pageref_valid": _fields_ok(snapshot_after, render_after) and _fields_ok(snapshot_stable, render_stable),
        "native_mathtype_cell_equation_preserved": _math_ok(math_after) and _math_ok(math_stable)
            and after["wide"]["native_equations"] == stable["wide"]["native_equations"] == 1
            and Counter(item["mtef_sha256"] for item in math_after["objects"]) == Counter(item["mtef_sha256"] for item in math_stable["objects"]),
        "word_pdf_borders_text_and_repeated_headers_within_page_body": _pdf_ok(render_after) and _pdf_ok(render_stable),
        "all_pdf_pages_rendered": len(render_after) >= 2 and len(render_stable) >= 2
                                  and all(Path(page["preview"]).is_file() for page in render_after + render_stable),
        "source_template_and_revision_input_unchanged": all(digest(path.read_bytes()) == sha for path, sha in originals.items())
                                                       and digest(master.read_bytes()) == master_hash,
    }
    fit_steps = [{"stage": step["stage"], "table_fit": step["evidence"]["table_fit"]}
                 for step in build["steps"] if "table_fit" in step["evidence"]]
    baseline = any(step["table_fit"].get("policy") == "baseline_before_table_fit" for step in fit_steps)
    report = {"pass": all(checks.values()), "checks": checks, "visual_reviewed": False,
              "execution": {"mode": "baseline_no_table_fit" if baseline else "default_table_fit",
                            "existing_build_reused": args.resume_existing, "table_fit_steps": fit_steps},
              "expected_geometry": {"template_body_width_twips": body_width_twips,
                                    "main_positive_indent_twips": max(0, before["wide"]["indent_twips"]),
                                    "main_width_limit_twips": expected_main_width_twips,
                                    "derivation": "template page width minus margins and gutter, minus positive table indent"},
              "reproduce": {"script": str(Path(__file__).resolve()), "source": str(source), "template": str(template),
                            "input_sha256": {str(path): sha for path, sha in originals.items()}},
              "source_tables": before, "built_tables": after, "revised_tables": stable,
              "built_native_geometry": native_after, "revised_native_geometry": native_stable,
              "built_native_snapshot": snapshot_after, "revised_native_snapshot": snapshot_stable,
              "built_mathtype": math_after, "revised_mathtype": math_stable,
              "built_render": render_after, "revised_render": render_stable,
              "workflows": [str(built / "workflow.json"), str(revised / "workflow.json")]}
    evidence = output / "table-width-fit-evidence.json"
    if evidence.exists():
        previous = evidence.read_bytes()
        backup = evidence.with_name(f"table-width-fit-evidence.before-resume-{digest(previous)[:12]}.json")
        if not backup.exists():
            with backup.open("xb") as stream:
                stream.write(previous)
    evidence.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"pass": report["pass"], "checks": checks, "evidence": str(evidence)}, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
