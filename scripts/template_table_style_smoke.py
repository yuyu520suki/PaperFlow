"""Verify a template-owned Word table style against a conflicting editable table."""
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
from docx.shared import Inches
import pymupdf

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from paperflow.citations import _scan_fields
from paperflow.mathtype import audit_docx_math
from paperflow.package import NS, digest, q, read_package, xml
from paperflow.pipeline import run_pipeline


STYLE = "Fixture Table Contract"
HEADER = "Imported measurement headers"
LAST_CELL = "Equation witness"
CAPTION = "Template styled results"
BORDER = "246197"
HEADER_FILL = "D5F2EC"
OLD_BORDER = "F02044"
OLD_FILL = "FFB347"


def _element(name, **attributes):
    element = OxmlElement(f"w:{name}")
    for key, value in attributes.items():
        element.set(qn(f"w:{key}"), str(value))
    return element


def _borders(name, color):
    borders = _element(name)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(_element(edge, val="single", sz=16, color=color))
    return borders


def _shading(color):
    return _element("shd", val="clear", color="auto", fill=color)


def _style(document, border, header_fill):
    style = document.styles.add_style(STYLE, WD_STYLE_TYPE.TABLE)
    style.base_style = document.styles["Normal Table"]
    properties = _element("tblPr")
    properties.append(_borders("tblBorders", border))
    style._element.append(properties)
    first_row = _element("tblStylePr", type="firstRow")
    cell_properties = _element("tcPr")
    cell_properties.append(_shading(header_fill))
    first_row.append(cell_properties)
    style._element.append(first_row)
    return style


def _look(table, first_row):
    flags = {"firstRow": first_row, "lastRow": 0, "firstColumn": 0, "lastColumn": 0,
             "noHBand": 1, "noVBand": 1, "val": "0620" if first_row else "0600"}
    look = table._tbl.tblPr.find(qn("w:tblLook"))
    for key, value in flags.items():
        look.set(qn(f"w:{key}"), str(value))


def _make_template(path):
    document = Document()
    style = _style(document, BORDER, HEADER_FILL)
    exemplar = document.add_table(rows=2, cols=2)
    exemplar.style = style
    exemplar.cell(0, 0).text = "Template exemplar"
    exemplar.cell(0, 1).text = "Header design"
    exemplar.cell(1, 0).text = "Reference row"
    _look(exemplar, 1)
    document.add_paragraph("{{PAPERFLOW_BODY}}")
    document.save(path)
    return style.style_id


def _make_source(path):
    document = Document()
    style = _style(document, OLD_BORDER, OLD_FILL)
    document.add_heading("Imported measurements", level=1)
    document.add_paragraph("See Table [[REF styledtable]] for the values.")
    document.add_paragraph(f"[[CAPTION Table styledtable|{CAPTION}]]")
    table = document.add_table(rows=4, cols=3)
    table.style = style
    table.autofit = False
    widths = [Inches(1.30), Inches(1.35), Inches(1.40)]
    for column, width in zip(table.columns, widths, strict=True):
        column.width = width
    for row in table.rows:
        for cell, width in zip(row.cells, widths, strict=True):
            cell.width = width
    preferred = table._tbl.tblPr.find(qn("w:tblW"))
    preferred.set(qn("w:type"), "dxa")
    preferred.set(qn("w:w"), str(sum(width.twips for width in widths)))
    _look(table, 0)
    table.cell(0, 0).merge(table.cell(0, 1)).text = HEADER
    table.cell(0, 2).text = "Value"
    table.cell(1, 0).merge(table.cell(2, 0)).text = "Series A"
    table.cell(1, 1).text = "First reading"
    table.cell(1, 2).text = "7.25"
    table.cell(2, 1).text = "Second reading"
    table.cell(2, 2).text = "9.5"
    table.cell(3, 0).text = "Series B"
    table.cell(3, 1).text = LAST_CELL
    equation = OxmlElement("m:oMath")
    run = OxmlElement("m:r")
    value = OxmlElement("m:t")
    value.text = "x+1=2"
    run.append(value)
    equation.append(run)
    table.cell(3, 2).paragraphs[0]._p.append(equation)
    properties = table._tbl.tblPr
    properties.insert_element_before(_borders("tblBorders", OLD_BORDER), "w:tblLayout", "w:tblLook")
    properties.insert_element_before(_shading(OLD_FILL), "w:tblLayout", "w:tblLook")
    for cell in table._tbl.iter(qn("w:tc")):
        properties = cell.get_or_add_tcPr()
        properties.append(_borders("tcBorders", OLD_BORDER))
        properties.append(_shading(OLD_FILL))
    document.save(path)


def _visible(element):
    return "".join(node.text or "" for node in element.iter(q("w:t")))


def _state(path, token):
    parts = read_package(path)
    document = xml(parts["word/document.xml"])
    table, = [element for element in document.iter(q("w:tbl")) if token in _visible(element)]
    style_id = table.find("w:tblPr/w:tblStyle", NS).get(q("w:val"))
    style, = [element for element in xml(parts["word/styles.xml"]).findall(q("w:style"))
              if element.get(q("w:styleId")) == style_id]
    style_name = style.find(q("w:name")).get(q("w:val"))
    look = table.find("w:tblPr/w:tblLook", NS)
    width = table.find("w:tblPr/w:tblW", NS)
    rows = table.findall(q("w:tr"))
    cells = [row.findall(q("w:tc")) for row in rows]
    style_fill = style.find("w:tblStylePr[@w:type='firstRow']/w:tcPr/w:shd", NS)
    return {
        "style_name": style_name, "style_id": style_id,
        "style_first_row_fill": style_fill.get(q("w:fill")) if style_fill is not None else None,
        "first_row_flag": look.get(q("w:firstRow")) if look is not None else None,
        "width": {"type": width.get(q("w:type")), "twips": int(width.get(q("w:w")))},
        "grid": [int(col.get(q("w:w"))) for col in table.findall("w:tblGrid/w:gridCol", NS)],
        "cell_text": [[_visible(cell) for cell in row] for row in cells],
        "horizontal_merges": [node.get(q("w:val")) for node in table.iter(q("w:gridSpan"))],
        "vertical_merges": [node.get(q("w:val"), "continue") for node in table.iter(q("w:vMerge"))],
        "has_old_color_overrides": any(
            node.get(q("w:fill")) == OLD_FILL or node.get(q("w:color")) == OLD_BORDER for node in table.iter()
        ),
        "native_equations_in_table": sum(1 for _ in table.iter(q("o:OLEObject"))),
        "fields": [(field["code"].strip(), field["result"].strip()) for field in _scan_fields(document)],
    }


def _rgb(value):
    return tuple(bytes.fromhex(value))


def _render(pdf_path, preview, width_twips):
    with pymupdf.open(pdf_path) as pdf:
        matches = [(index, page.search_for(HEADER), page.search_for(LAST_CELL))
                   for index, page in enumerate(pdf, 1) if page.search_for(HEADER)]
        if len(matches) != 1 or len(matches[0][1]) != 1 or len(matches[0][2]) != 1:
            raise RuntimeError("Expected both imported table witnesses on one Word-rendered page")
        index, (header,), (last,) = matches[0]
        page = pdf[index - 1]
        page.get_pixmap(dpi=140, colorspace=pymupdf.csRGB, alpha=False).save(preview)
        clip = pymupdf.Rect(min(header.x0, last.x0) - 15, header.y0 - 12,
                            header.x0 + width_twips / 20 + 8, last.y1 + 18) & page.rect
        samples = page.get_pixmap(dpi=160, clip=clip, colorspace=pymupdf.csRGB, alpha=False).samples
        colors = Counter(zip(samples[0::3], samples[1::3], samples[2::3]))
        # Word's PDF color components and the rasterizer round at different
        # steps. Count a pixel when each channel is within two RGB levels.
        def near_count(color):
            expected = _rgb(color)
            return sum(count for observed, count in colors.items()
                       if all(abs(a - b) <= 2 for a, b in zip(observed, expected, strict=True)))

        return {
            "page": index, "clip_pt": list(clip), "preview": str(preview),
            "pixel_counts": {color: near_count(color) for color in (BORDER, HEADER_FILL, OLD_BORDER, OLD_FILL)},
            "rgb_tolerance_per_channel": 2,
            "text": " ".join(page.get_text().split()),
        }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    output = parser.parse_args().output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    source, template = output / "source.docx", output / "template.docx"
    style_id = _make_template(template)
    _make_source(source)
    originals = {path: digest(path.read_bytes()) for path in (source, template)}
    before, exemplar_before = _state(source, HEADER), _state(template, "Template exemplar")
    built = output / "built"
    workflow = run_pipeline(source, template, built, table_style=style_id)
    master = built / "editable-master.docx"
    after, exemplar_after = _state(master, HEADER), _state(master, "Template exemplar")
    render = _render(built / "editable-master.pdf", output / "review-page.png", after["width"]["twips"])
    math = audit_docx_math(master)
    checks = {
        "native_build_finished": workflow["native_build_finished"],
        "template_table_style_saved": after["style_name"].casefold() == STYLE.casefold()
                                      and after["style_first_row_fill"] == HEADER_FILL,
        "template_first_row_flag_adopted": before["first_row_flag"] == "0" and after["first_row_flag"] == "1",
        "conflicting_source_overrides_removed": before["has_old_color_overrides"] and not after["has_old_color_overrides"],
        "table_grid_and_width_preserved": before["grid"] == after["grid"] and before["width"] == after["width"],
        "merged_cells_and_editable_text_preserved": (
            before["horizontal_merges"] == after["horizontal_merges"] == ["2"]
            and before["vertical_merges"] == after["vertical_merges"] == ["restart", "continue"]
            and before["cell_text"] == after["cell_text"]
        ),
        "native_mathtype_in_table": after["native_equations_in_table"] == math["native_mathtype_objects"] == 1
                                    and not math["omml_equations"] and math["all_mathtype_objects_valid"],
        "native_caption_and_cross_reference": (
            any(re.match(r"SEQ\s+Table\b", code, re.I) and value == "1" for code, value in after["fields"])
            and any(re.match(r"REF\s+styledtable\b", code, re.I) and value == "1" for code, value in after["fields"])
            and f"Table 1 {CAPTION}" in render["text"] and "See Table 1 for the values." in render["text"]
        ),
        "template_exemplar_unchanged_in_output": all(exemplar_before[key] == exemplar_after[key] for key in (
            "style_name", "first_row_flag", "grid", "cell_text",
        )),
        "word_pdf_renders_selected_style_on_imported_table": (
            render["pixel_counts"][HEADER_FILL] > 500 and render["pixel_counts"][BORDER] > 40
            and render["pixel_counts"][OLD_BORDER] == render["pixel_counts"][OLD_FILL] == 0
        ),
        "source_and_template_unchanged": all(digest(path.read_bytes()) == sha for path, sha in originals.items()),
    }
    evidence = output / "template-table-style-evidence.json"
    report = {"pass": all(checks.values()), "checks": checks, "source": before, "saved_table": after,
              "render": render, "workflow": str(built / "workflow.json")}
    evidence.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"pass": report["pass"], "checks": checks, "evidence": str(evidence)}, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
