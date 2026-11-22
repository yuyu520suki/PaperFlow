"""Read saved Word table and cell geometry without changing the document."""
from __future__ import annotations

from collections import defaultdict
from math import isfinite
from pathlib import Path

from .package import NS, digest, q, read_package, xml
from .word import _application, _opened
from .table_style import body_tables


_TOLERANCE_PT = 0.6
_INLINE_KINDS = {1: "embedded_ole", 2: "linked_ole", 3: "picture"}
_BROKEN_SERVER = {0x80010108, 0x800706BA, 0x800706BE, 0x80010007, 0x80010012}


def _number(value, *, positive=False):
    value = float(value)
    if not isfinite(value) or abs(value) >= 1000000:
        raise ValueError("Word returned missing or undefined geometry")
    if positive and value <= 0:
        raise ValueError("Word returned no positive width")
    return value


def _review(issues, context, error):
    hresult = getattr(error, "hresult", None)
    if hresult is not None and int(hresult) & 0xFFFFFFFF in _BROKEN_SERVER:
        raise error
    issues.append(f"{context}: {error}")


def _section_width(table, top_gutter):
    section = table.Range.Sections.Item(1)
    page = section.PageSetup
    width = _number(page.PageWidth, positive=True)
    left, right = _number(page.LeftMargin), _number(page.RightMargin)
    if min(left, right) < 0:
        raise ValueError("Word returned negative horizontal page margins")
    width -= left + right
    if not top_gutter:
        gutter = _number(page.Gutter)
        if gutter < 0:
            raise ValueError("Word returned a negative page gutter")
        width -= gutter
    columns = page.TextColumns
    if columns.Count:
        width = min(width, *(_number(columns.Item(i).Width, positive=True)
                            for i in range(1, columns.Count + 1)))
    if width <= 0:
        raise ValueError("section margins or columns leave no positive table width")
    return width


def _native_cells(table, level, issues, label):
    result = []
    complete = True
    cells = table.Range.Cells
    for index in range(1, cells.Count + 1):
        try:
            cell = cells.Item(index)
            if int(cell.NestingLevel) != level:
                continue
            location = cell.Range
            record = {
                "row": int(cell.RowIndex), "column": int(cell.ColumnIndex),
                "start": int(location.Start), "end": int(location.End),
                "width_pt": _number(cell.Width, positive=True),
                "left_padding_pt": None, "right_padding_pt": None,
            }
            result.append(record)
            try:
                left, right = _number(cell.LeftPadding), _number(cell.RightPadding)
                if min(left, right) < 0:
                    raise ValueError("Word returned negative cell padding")
                record["left_padding_pt"], record["right_padding_pt"] = left, right
            except Exception as error:
                _review(issues, f"{label} cell {record['row']},{record['column']} padding", error)
        except Exception as error:
            complete = False
            _review(issues, f"{label} cell {index} geometry", error)
    return result, complete


def audit_table_layout(path, *, app=None):
    """Reopen a saved document and inspect native table and inline-object widths.

    The table metric is the largest sum of native cell widths sharing a row
    index. Word can omit vertical-merge continuations from that collection, so
    this is a useful width check rather than proof of every rendered border.
    Floating placement and unavailable geometry are reported for review.
    """
    source = Path(path).resolve(strict=True)
    signature = digest(source.read_bytes())
    parts = read_package(source)
    settings = xml(parts["word/settings.xml"]) if "word/settings.xml" in parts else None
    gutter = settings.find(q("w:gutterAtTop")) if settings is not None else None
    top_gutter = gutter is not None and gutter.get(q("w:val"), "1").casefold() not in {"0", "false", "off"}
    report = {"tables": [], "content": [], "unresolved": [], "native_reopened": False,
              "tolerance_pt": _TOLERANCE_PT,
              "width_metric": "maximum_enumerated_row_cell_width_sum",
              "width_metric_note": "Vertical-merge continuation cells can be omitted by Word; this metric does not prove every border position."}
    issues = report["unresolved"]
    xml_tables = body_tables(xml(parts["word/document.xml"]))
    if not xml_tables:
        report.update(applicable=False, source_unchanged=True)
        return report
    report["applicable"] = True
    # These concrete content constraints have no width metric in this audit.
    # Keep them visible even when all measured borders/inline objects fit.
    for index, table in enumerate(xml_tables, 1):
        owned = [node for node in table.iter() if next(node.iterancestors(q("w:tbl")), None) is table]
        if any(node.tag == q("w:noWrap") and node.get(q("w:val"), "1").casefold()
               not in {"0", "false", "off"} for node in owned):
            issues.append(f"Table {index} contains nonwrapping cell text; review text bounds in the Word export")
        if any(node.tag == q("m:oMath") for node in owned):
            issues.append(f"Table {index} contains Word equations whose rendered width was not measured")
        if any(node.tag == "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor"
               for node in owned):
            issues.append(f"Table {index} contains a floating drawing; review its cell placement in the Word export")
    with _application(app) as word, _opened(word, source) as document:
        report["native_reopened"] = True
        pending = [document.Tables.Item(index) for index in range(1, document.Tables.Count + 1)]
        seen, tables = set(), []
        while pending:
            table = pending.pop()
            location = table.Range
            if int(location.StoryType) != 1:
                continue
            start, end, level = int(location.Start), int(location.End), int(table.NestingLevel)
            key = start, end, level
            if key in seen:
                continue
            seen.add(key)
            tables.append((key, table))
            try:
                nested = table.Tables
                pending.extend(nested.Item(index) for index in range(1, nested.Count + 1))
            except Exception as error:
                _review(issues, f"Table at {start} nested-table enumeration", error)
        tables.sort(key=lambda item: (item[0][0], item[0][2], item[0][1]))
        parent_cells = []
        for (start, end, level), table in tables:
            index = len(report["tables"]) + 1
            label = f"Table {index}"
            record = {"index": index, "start": start, "end": end, "nesting_level": level,
                      "measured_width_pt": None, "width_limit_pt": None, "limit_policy": "section_column",
                      "left_indent_pt": None, "floating": None, "cells": [], "cells_complete": False,
                      "row_widths_pt": {}, "within_width_limit": None}
            report["tables"].append(record)
            if index <= len(xml_tables):
                columns = xml_tables[index - 1].findall("w:tblGrid/w:gridCol", NS)
                if columns:
                    try:
                        widths = [int(column.get(q("w:w"), "0")) for column in columns]
                        if min(widths) <= 0:
                            raise ValueError("saved grid contains nonpositive column widths")
                        record["saved_grid_width_pt"] = sum(widths) / 20
                    except ValueError as error:
                        _review(issues, f"{label} saved grid", error)
            try:
                cells, complete = _native_cells(table, level, issues, label)
                record["cells"], record["cells_complete"] = cells, complete
                widths = defaultdict(float)
                for cell in cells:
                    widths[cell["row"]] += cell["width_pt"]
                record["row_widths_pt"] = dict(sorted(widths.items()))
                if not widths:
                    raise ValueError("Word returned no measurable cells")
                record["measured_width_pt"] = max(widths.values())
            except Exception as error:
                _review(issues, f"{label} width measurement", error)
            try:
                limit = _section_width(table, top_gutter)
                if level > 1:
                    parents = [(parent_index, cell) for parent_level, parent_index, cell in parent_cells
                               if parent_level == level - 1 and cell["start"] <= start and cell["end"] >= end]
                    if not parents:
                        raise ValueError("nested table has no measurable containing parent cell")
                    parent_index, parent = min(parents, key=lambda item: item[1]["end"] - item[1]["start"])
                    if parent["left_padding_pt"] is None or parent["right_padding_pt"] is None:
                        raise ValueError("nested table parent-cell padding is unavailable")
                    limit = parent["width_pt"] - parent["left_padding_pt"] - parent["right_padding_pt"]
                    if limit <= 0:
                        raise ValueError("parent cell padding leaves no positive nested-table width")
                    record["parent_table_index"] = parent_index
                    record["parent_cell"] = {"row": parent["row"], "column": parent["column"]}
                    record["limit_policy"] = "native_parent_cell_content_width"
                record["width_limit_pt"] = limit
            except Exception as error:
                _review(issues, f"{label} available width", error)
            try:
                indent = _number(table.Rows.LeftIndent)
                record["left_indent_pt"] = indent
                if int(table.Rows.Alignment) == 0 and record["width_limit_pt"] is not None:
                    record["width_limit_pt"] -= max(0, indent)
                    if indent < -_TOLERANCE_PT:
                        record["placement_needs_review"] = True
                        issues.append(f"{label} has a negative leading indent; review its left boundary in the Word export")
            except Exception as error:
                hresult = getattr(error, "hresult", None)
                if hresult is not None and int(hresult) & 0xFFFFFFFF in _BROKEN_SERVER:
                    raise
                # Mixed/merged row collections can have no single indent.
                record["indent_note"] = "Word did not expose a single row indent."
            try:
                wrapping = int(table.Rows.WrapAroundText)
                if wrapping not in {0, -1, 1}:
                    raise ValueError("Word returned mixed floating-table positioning")
                record["floating"] = bool(wrapping)
                if wrapping:
                    record["horizontal_position"] = float(table.Rows.HorizontalPosition)
                    record["relative_horizontal_position"] = int(table.Rows.RelativeHorizontalPosition)
                    issues.append(f"{label} is floating; review its horizontal placement in the Word export")
            except Exception as error:
                _review(issues, f"{label} floating placement", error)
            measured, limit = record["measured_width_pt"], record["width_limit_pt"]
            if measured is not None and limit is not None:
                saved_grid = record.get("saved_grid_width_pt")
                record["within_width_limit"] = max(measured, saved_grid if saved_grid is not None else measured) <= limit + _TOLERANCE_PT
                if measured > limit + _TOLERANCE_PT:
                    issues.append(f"{label} native row-cell width {measured:.2f} pt exceeds its {limit:.2f} pt limit")
                if saved_grid is not None and saved_grid > max(limit, measured) + _TOLERANCE_PT:
                    issues.append(f"{label} saved logical grid width {saved_grid:.2f} pt exceeds its {limit:.2f} pt limit; review omitted columns and cell boundaries")
            parent_cells.extend((level, index, cell) for cell in record["cells"])
        shapes = document.InlineShapes
        for index in range(1, shapes.Count + 1):
            try:
                shape = shapes.Item(index)
                shape_type = int(shape.Type)
                location = shape.Range
                if shape_type not in _INLINE_KINDS or int(location.StoryType) != 1 or not location.Information(12):
                    continue
                cell = location.Cells.Item(1)
                paragraph = location.ParagraphFormat
                record = {"index": index, "kind": _INLINE_KINDS[shape_type], "start": int(location.Start),
                          "cell_row": int(cell.RowIndex), "cell_column": int(cell.ColumnIndex),
                          "width_pt": _number(shape.Width, positive=True), "width_limit_pt": None,
                          "within_width_limit": None}
                report["content"].append(record)
                containing = [table for table in report["tables"]
                              if table["start"] <= record["start"] < table["end"]]
                if containing:
                    record["table_index"] = max(containing, key=lambda table: table["nesting_level"])["index"]
                left, right = _number(cell.LeftPadding), _number(cell.RightPadding)
                if min(left, right) < 0:
                    raise ValueError("Word returned negative cell padding")
                padding = left + right
                indents = max(0, _number(paragraph.LeftIndent)) + max(0, _number(paragraph.RightIndent))
                limit = _number(cell.Width, positive=True) - padding - indents
                if limit <= 0:
                    raise ValueError("cell padding and paragraph indents leave no positive inline-object width")
                record["width_limit_pt"] = limit
                record["within_width_limit"] = record["width_pt"] <= limit + _TOLERANCE_PT
                if not record["within_width_limit"]:
                    issues.append(f"Inline {record['kind']} {index} width {record['width_pt']:.2f} pt exceeds its cell paragraph width {limit:.2f} pt")
            except Exception as error:
                _review(issues, f"Inline object {index} cell-width measurement", error)
    if digest(source.read_bytes()) != signature:
        raise ValueError("The saved document changed while its native table layout was inspected")
    report["source_unchanged"] = True
    return report
