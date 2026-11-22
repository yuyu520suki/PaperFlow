"""Exercise auto, percentage, inferred-grid and indented tables through Word.

Run serially with a fresh output directory. Word is required; this fixture has
no equations, citations or inline objects and does not require MathType. The
script checks the normal pipeline's reopened native audit and every PDF page.
PNG generation and geometric checks do not constitute human visual review.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt, Twips
import pymupdf

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from paperflow.package import NS, digest, q, read_package, xml
from paperflow.pipeline import run_pipeline


CASES = ("AUTO9", "PCT150", "PCT50", "SMALL_FIXED2", "MISSING_GRID",
         "ROW_INDENT_EXCEPTION", "NEGATIVE_INDENT")
COLORS = dict(zip(CASES, ("2D5B81", "7B5B9A", "2E6B46", "9D5B22",
                          "8A3948", "367A82", "6A6F2A"), strict=True))
BODY_WIDTH_TWIPS = 7200  # 7-inch page minus two 1-inch margins.
EXPECTED_WIDTHS = {case: BODY_WIDTH_TWIPS for case in CASES}
EXPECTED_WIDTHS.update(PCT50=3600, SMALL_FIXED2=2880, ROW_INDENT_EXCEPTION=6480)
WIDTH_TOLERANCE_TWIPS = 3
NATIVE_TOLERANCE_PT = 0.6
PDF_TOLERANCE_PT = 1.0


def _element(name, **attributes):
    node = OxmlElement("w:" + name)
    for key, value in attributes.items():
        node.set(q("w:" + key), str(value))
    return node


def _make_template(path):
    document = Document()
    section = document.sections[0]
    section.page_width, section.page_height = Inches(7), Inches(11)
    section.left_margin = section.right_margin = Inches(1)
    section.top_margin = section.bottom_margin = Inches(1)
    document.styles["Normal"].font.name = "Calibri"
    document.styles["Normal"].font.size = Pt(10)
    document.styles["Normal"].paragraph_format.space_after = Pt(5)
    document.add_paragraph("{{PAPERFLOW_BODY}}")
    document.save(path)


def _make_source(path):
    document = Document()
    section = document.sections[0]
    section.page_width, section.page_height = Inches(11), Inches(11)
    section.left_margin = section.right_margin = Inches(0.5)
    document.styles["Normal"].font.name = "Calibri"
    document.styles["Normal"].font.size = Pt(10)
    document.add_paragraph("Table width variants: native layout acceptance fixture.")
    for case in CASES:
        document.add_paragraph("Variant " + case)
        table = document.add_table(rows=2, cols=2)
        widths = (1728, 1152) if case == "SMALL_FIXED2" else (7776, 5184)
        table.autofit = case in {"AUTO9", "PCT150", "PCT50", "MISSING_GRID"}
        for column, width in zip(table.columns, widths, strict=True):
            column.width = Twips(width)
        for row in table.rows:
            for cell, width in zip(row.cells, widths, strict=True):
                cell.width = Twips(width)
        table.cell(0, 0).text = case
        table.cell(0, 1).text = "Value"
        table.cell(1, 0).text = "Sample A"
        table.cell(1, 1).text = "0.25"
        properties = table._tbl.tblPr
        width = properties.find(q("w:tblW"))
        kind, value = ("auto", 0) if case in {"AUTO9", "MISSING_GRID"} else (
            ("pct", 7500) if case == "PCT150" else ("pct", 2500) if case == "PCT50" else ("dxa", sum(widths)))
        width.set(q("w:type"), kind)
        width.set(q("w:w"), str(value))
        properties.insert_element_before(_element("jc", val="left"), "w:tblLayout", "w:tblLook")
        properties.insert_element_before(_element("tblInd", type="dxa", w=-720 if case == "NEGATIVE_INDENT" else 0),
                                         "w:tblBorders", "w:tblLayout", "w:tblLook")
        borders = _element("tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            borders.append(_element(edge, val="single", sz=8, color=COLORS[case]))
        properties.insert_element_before(borders, "w:tblLayout", "w:tblLook")
        # Zero horizontal cell padding makes the outer-border position directly
        # comparable with the page body and the explicit row indentation.
        margins = _element("tblCellMar")
        for edge in ("top", "bottom", "left", "right"):
            margins.append(_element(edge, type="dxa", w=24 if edge in {"top", "bottom"} else 0))
        properties.insert_element_before(margins, "w:tblLook")
        for row in table.rows:
            row._tr.get_or_add_trPr().append(_element("cantSplit"))
            if case == "ROW_INDENT_EXCEPTION":
                exception = _element("tblPrEx")
                exception.append(_element("jc", val="left"))
                exception.append(_element("tblInd", type="dxa", w=720))
                row._tr.insert(0, exception)
        if case == "MISSING_GRID":
            table._tbl.remove(table._tbl.tblGrid)
        # A real paragraph separates tables so native insertion cannot join
        # neighboring tables that happen to use the same default style.
        document.add_paragraph(" ")
    document.save(path)


def _text(node):
    return "".join(item.text or "" for item in node.iter(q("w:t")))


def _state(path):
    parts = read_package(path)
    root = xml(parts["word/document.xml"])
    found = {}
    for table in root.findall("w:body/w:tbl", NS):
        rows = table.findall(q("w:tr"))
        label = _text(rows[0].find(q("w:tc")))
        if label not in CASES or label in found:
            raise RuntimeError(f"Unexpected or duplicate fixture table label: {label!r}")
        properties = table.find(q("w:tblPr"))
        preferred = properties.find(q("w:tblW"))
        layout = properties.find(q("w:tblLayout"))
        indent = properties.find(q("w:tblInd"))
        found[label] = {
            "grid": [int(column.get(q("w:w"))) for column in table.findall("w:tblGrid/w:gridCol", NS)],
            "preferred_width": {"type": preferred.get(q("w:type")), "value": preferred.get(q("w:w"))},
            "layout": layout.get(q("w:type")) if layout is not None else None,
            "indent_twips": int(indent.get(q("w:w"), "0")) if indent is not None else 0,
            "row_indents_twips": [int(node.get(q("w:w"), "0")) if node is not None else None
                                   for node in (row.find("w:tblPrEx/w:tblInd", NS) for row in rows)],
            "cell_widths": [[int(cell.find("w:tcPr/w:tcW", NS).get(q("w:w")))
                             for cell in row.findall(q("w:tc"))] for row in rows],
            "cell_width_types": [[cell.find("w:tcPr/w:tcW", NS).get(q("w:type"))
                                  for cell in row.findall(q("w:tc"))] for row in rows],
            "cell_text": [[_text(cell) for cell in row.findall(q("w:tc"))] for row in rows],
            "cant_split": [row.find("w:trPr/w:cantSplit", NS) is not None for row in rows],
        }
    if tuple(found) != CASES:
        raise RuntimeError(f"Fixture tables changed order or disappeared: {tuple(found)!r}")
    fields = sum(1 for node in root.iter() if node.tag in {q("w:fldSimple"), q("w:fldChar")})
    return {"tables": found, "field_markup_count": fields}


def _stage(workflow, name):
    return next(step["evidence"] for step in reversed(workflow["steps"]) if step["stage"] == name)


def _color_matches(observed, expected):
    return observed is not None and len(observed) == 3 and all(
        abs(left - right) <= 2 / 255 for left, right in zip(observed, expected, strict=True))


def _render(pdf_path, output):
    output.mkdir(exist_ok=False)
    colors = {case: tuple(int(color[index:index + 2], 16) / 255 for index in (0, 2, 4))
              for case, color in COLORS.items()}
    pages = []
    with pymupdf.open(pdf_path) as pdf:
        for index, page in enumerate(pdf, 1):
            preview = output / f"page-{index:03d}.png"
            page.get_pixmap(dpi=140, colorspace=pymupdf.csRGB, alpha=False).save(preview)
            drawings = page.get_drawings()
            borders = {case: [list(item["rect"]) for item in drawings
                              if _color_matches(item.get("color"), color) or _color_matches(item.get("fill"), color)]
                       for case, color in colors.items()}
            pages.append({"page": index, "preview": str(preview),
                          "size_pt": [page.rect.width, page.rect.height],
                          "text": " ".join(page.get_text().split()),
                          "text_rects_pt": [list(word[:4]) for word in page.get_text("words")],
                          "border_rects_pt": borders})
    return pages


def _geometry_ok(state):
    for case, table in state["tables"].items():
        expected = EXPECTED_WIDTHS[case]
        if (abs(sum(table["grid"]) - expected) > WIDTH_TOLERANCE_TWIPS
                or table["layout"] != "fixed" or table["preferred_width"]["type"] != "dxa"
                or abs(int(table["preferred_width"]["value"]) - expected) > WIDTH_TOLERANCE_TWIPS
                or any(types != ["dxa", "dxa"] for types in table["cell_width_types"])
                or any(len(row) != 2 or any(abs(a - b) > WIDTH_TOLERANCE_TWIPS for a, b in zip(row, table["grid"], strict=True))
                       for row in table["cell_widths"])):
            return False
    return True


def _native_ok(native, state):
    if (not native.get("native_reopened") or not native.get("source_unchanged")
            or native["unresolved"] or len(native["tables"]) != len(CASES)
            or native["content"]):
        return False
    for case, record in zip(CASES, native["tables"], strict=True):
        if (not record["within_width_limit"] or not record["cells_complete"] or record["floating"]
                or record["nesting_level"] != 1 or len(record["cells"]) != 4
                or abs(record["measured_width_pt"] - sum(state["tables"][case]["grid"]) / 20) > NATIVE_TOLERANCE_PT):
            return False
    return True


def _pdf_ok(pages):
    if not pages or any(abs(page["size_pt"][0] - 504) > 0.1 for page in pages):
        return False
    for case in CASES:
        rectangles = [rect for page in pages for rect in page["border_rects_pt"][case]]
        if not rectangles:
            return False
        if any(rect[0] < 72 - PDF_TOLERANCE_PT or rect[2] > 432 + PDF_TOLERANCE_PT for rect in rectangles):
            return False
        left, right = min(rect[0] for rect in rectangles), max(rect[2] for rect in rectangles)
        if abs((right - left) - EXPECTED_WIDTHS[case] / 20) > PDF_TOLERANCE_PT:
            return False
        if case == "ROW_INDENT_EXCEPTION" and abs(left - 108) > PDF_TOLERANCE_PT:
            return False
    text = " ".join(page["text"] for page in pages)
    return all(case in text for case in CASES) and all(
        rect[0] >= 72 - PDF_TOLERANCE_PT and rect[2] <= 432 + PDF_TOLERANCE_PT
        for page in pages for rect in page["text_rects_pt"])


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    output = parser.parse_args().output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    source, template = output / "source.docx", output / "template.docx"
    _make_source(source)
    _make_template(template)
    original_hashes = {path: digest(path.read_bytes()) for path in (source, template)}
    before = _state(source)
    workflow = run_pipeline(source, template, output / "built", require_mathtype=False)
    normalized = _state(output / "built/02-template-styles.docx")
    saved = _state(output / "built/editable-master.docx")
    fit = _stage(workflow, "normalize")["table_fit"]
    native = _stage(workflow, "table_layout")
    snapshot = next(step["evidence"]["snapshot"] for step in reversed(workflow["steps"])
                    if step["stage"] in {"word_fields", "word_fields_after_zotero"})
    pages = _render(output / "built/editable-master.pdf", output / "pages")
    checks = {
        "native_build_finished": workflow["native_build_finished"],
        "source_contains_the_declared_width_variants": all(before["tables"][case]["grid"] == (
                [] if case == "MISSING_GRID" else [1728, 1152] if case == "SMALL_FIXED2" else [7776, 5184]) for case in CASES)
            and all(before["tables"][case]["preferred_width"] == {"type": kind, "value": value}
                    for case, kind, value in (("AUTO9", "auto", "0"), ("PCT150", "pct", "7500"),
                                               ("PCT50", "pct", "2500"), ("SMALL_FIXED2", "dxa", "2880"),
                                               ("MISSING_GRID", "auto", "0")))
            and before["tables"]["ROW_INDENT_EXCEPTION"]["row_indents_twips"] == [720, 720],
        "normal_pipeline_fitted_six_variants_and_preserved_small_fixed": len(fit["tables"]) == len(CASES)
            and fit["resized"] == 6 and [record["status"] for record in fit["tables"]]
            == ["resized", "resized", "resized", "unchanged", "resized", "resized", "resized"],
        "missing_grid_inference_is_explicitly_reported": len(fit["unresolved"]) == 1
            and "missing grid" in fit["unresolved"][0].lower()
            and fit["tables"][CASES.index("MISSING_GRID")].get("grid_inferred") is True,
        "normalized_and_saved_fixed_geometry_matches_each_variant": _geometry_ok(normalized) and _geometry_ok(saved),
        "small_fixed_2in_grid_and_cell_widths_unchanged": all(before["tables"]["SMALL_FIXED2"][key] == saved["tables"]["SMALL_FIXED2"][key]
            for key in ("grid", "preferred_width", "cell_widths")),
        "negative_indent_was_clamped_to_body_margin": before["tables"]["NEGATIVE_INDENT"]["indent_twips"] == -720
            and normalized["tables"]["NEGATIVE_INDENT"]["indent_twips"] == saved["tables"]["NEGATIVE_INDENT"]["indent_twips"] == 0,
        "table_text_and_row_controls_preserved": all(before["tables"][case][key] == saved["tables"][case][key]
            for case in CASES for key in ("cell_text", "cant_split")),
        "reopened_native_geometry_matches_saved_package_and_limits": _native_ok(native, saved),
        "no_fields_or_inline_objects_were_introduced": before["field_markup_count"] == saved["field_markup_count"] == 0
            and snapshot["fields"] == [] and snapshot["inline_shapes"] == [] and native["content"] == [],
        "all_colored_table_borders_and_text_stay_in_pdf_body": _pdf_ok(pages),
        "every_exported_pdf_page_rendered": bool(pages) and all(Path(page["preview"]).is_file() for page in pages),
        "source_and_template_unchanged": all(digest(path.read_bytes()) == expected for path, expected in original_hashes.items()),
    }
    report = {"pass": all(checks.values()), "checks": checks, "visual_reviewed": False,
              "expected_geometry": {"body_width_twips": BODY_WIDTH_TWIPS, "table_widths_twips": EXPECTED_WIDTHS,
                                    "row_exception_indent_twips": 720, "border_colors": COLORS},
              "reproduce": {"script": str(Path(__file__).resolve()),
                            "input_sha256": {str(path): value for path, value in original_hashes.items()}},
              "source": before, "normalized": normalized, "saved": saved,
              "table_fit": fit, "native_table_layout": native, "native_snapshot": snapshot,
              "render": pages, "workflow": str(output / "built/workflow.json")}
    evidence = output / "table-width-variants-evidence.json"
    evidence.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"pass": report["pass"], "checks": checks, "evidence": str(evidence)}, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
