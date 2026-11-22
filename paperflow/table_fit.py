"""Bound imported table geometry without reconstructing cells or their contents."""
from __future__ import annotations

from copy import deepcopy
from fractions import Fraction

from lxml import etree as ET

from .package import NS, q
from .picture_fit import EMU_PER_TWIP, _limits
from .table_style import body_tables


_SKIP = {q("w:" + name) for name in ("del", "moveFrom", "txbxContent")}
_ORDER = {q("w:" + name): n for n, name in enumerate((
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
    "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar",
    "tblLook", "tblCaption", "tblDescription", "tblPrChange",
))}
_ON = {"1", "true", "on"}


def _own(root, tag, owner_tag):
    return [node for node in root.iter(q("w:" + tag))
            if next(node.iterancestors(q("w:" + owner_tag)), None) is root
            and not any(parent.tag in _SKIP for parent in node.iterancestors())]


def _number(node, name="val", default=None):
    raw = node.get(q("w:" + name)) if node is not None else None
    if raw is None and default is not None:
        return default
    try:
        value = int(raw)
    except (ValueError, TypeError):
        raise ValueError("missing or invalid table grid dimension") from None
    if value < 0:
        raise ValueError("negative table grid dimension")
    return value


def _width(node, base, *, default=None):
    if node is None:
        return default
    kind, raw = node.get(q("w:type"), "dxa"), node.get(q("w:w"), "0")
    if kind in {"auto", "nil"}:
        return default
    try:
        if kind == "pct":
            ratio = Fraction(raw[:-1]) / 100 if raw.endswith("%") else Fraction(raw) / 5000
            value = int(base * ratio)
        elif kind == "dxa":
            value = int(raw)
        else:
            raise ValueError
    except (ValueError, ZeroDivisionError):
        raise ValueError("invalid table preferred width or unit") from None
    if value < 0:
        raise ValueError("negative table preferred width")
    return value or default


def _put(properties, name):
    tag = q("w:" + name)
    node = properties.find(tag)
    if node is None:
        node = ET.Element(tag)
        position = next((i for i, child in enumerate(properties)
                         if _ORDER.get(child.tag, -1) > _ORDER.get(tag, -1)), len(properties))
        properties.insert(position, node)
    return node


def _set_width(node, width):
    node.set(q("w:type"), "dxa")
    node.set(q("w:w"), str(width))


def _indent(node):
    # Unlike tblW, percentage/auto indentation is ignored by Word.
    if node is None or node.get(q("w:type"), "dxa") != "dxa":
        return 0
    try:
        return int(node.get(q("w:w"), "0"))
    except ValueError:
        raise ValueError("invalid table indentation") from None


def _properties(table, styles):
    direct = table.find(q("w:tblPr"))
    definitions = {style.get(q("w:styleId")): style for style in styles.findall(q("w:style"))}
    selected = direct.find(q("w:tblStyle")) if direct is not None else None
    style_id = selected.get(q("w:val")) if selected is not None else next((
        key for key, style in definitions.items() if style.get(q("w:type")) == "table"
        and style.get(q("w:default"), "").casefold() in _ON
    ), None)
    chain, visited = [], set()
    while style_id:
        if style_id in visited or style_id not in definitions:
            raise ValueError("unresolved table style inheritance")
        visited.add(style_id)
        style = definitions[style_id]
        chain.append(style)
        parent = style.find(q("w:basedOn"))
        style_id = parent.get(q("w:val")) if parent is not None else None
    chain.reverse()
    layers = [style.find(q("w:tblPr")) for style in chain] + [direct]
    effective = {}
    for layer in layers:
        if layer is not None:
            effective.update({ET.QName(child).localname: child for child in layer})
    return layers, effective, chain


def _row_map(table):
    result, count = [], 0
    for row in _own(table, "tr", "tbl"):
        start = _number(row.find("w:trPr/w:gridBefore", NS), default=0)
        cells = []
        for cell in _own(row, "tc", "tr"):
            span = _number(cell.find("w:tcPr/w:gridSpan", NS), default=1)
            if span < 1:
                raise ValueError("table cell has an empty grid span")
            cells.append((cell, start, span))
            start += span
        after = _number(row.find("w:trPr/w:gridAfter", NS), default=0)
        count = max(count, start + after)
        result.append((row, cells, after))
    if not result or count < 1:
        raise ValueError("table has no usable logical grid")
    return result, count


def _grid(table, rows, count, fallback):
    columns = table.findall("w:tblGrid/w:gridCol", NS)
    weights = [_number(col, "w") for col in columns]
    if any(value <= 0 for value in weights):
        raise ValueError("table grid contains nonpositive column widths")
    count = max(count, len(weights))
    inferred = len(weights) != count
    weights += [None] * (count - len(weights))
    # Saved grid values win over conflicting cell preferences. Preferences are
    # useful constraints only for omitted columns, including old Word inputs.
    for _, cells, _ in rows:
        for cell, start, span in cells:
            missing = [i for i in range(start, start + span) if weights[i] is None]
            preferred = _width(cell.find("w:tcPr/w:tcW", NS), fallback)
            if missing and preferred:
                residual = preferred - sum(weights[i] or 0 for i in range(start, start + span))
                if residual > 0:
                    for i in missing:
                        weights[i] = max(1, residual // len(missing))
    known = sum(value or 0 for value in weights)
    missing = sum(value is None for value in weights)
    residual = max(missing, fallback - known)
    weights = [value if value is not None else max(1, residual // max(1, missing)) for value in weights]
    return weights, inferred


def _allocate(weights, total):
    if total < len(weights):
        raise ValueError("target width cannot retain positive table columns")
    exact = [Fraction(value * total, sum(weights)) for value in weights]
    result = [max(1, int(value)) for value in exact]
    remainder = total - sum(result)
    if remainder < 0:
        for i in sorted(range(len(result)), key=lambda i: result[i], reverse=True):
            take = min(result[i] - 1, -remainder)
            result[i] -= take
            remainder += take
            if not remainder:
                break
    else:
        order = sorted(range(len(result)), key=lambda i: exact[i] - int(exact[i]), reverse=True)
        for i in order[:remainder]:
            result[i] += 1
    if sum(result) != total:
        raise ValueError("could not allocate bounded table columns")
    return result


def _margins(layers, row, cell, styles):
    # Word's omitted horizontal cell-margin default is 108 twips. Resolve
    # inherited table defaults per side rather than replacing the whole group.
    values = {"left": 108, "right": 108}
    groups = [layer.find(q("w:tblCellMar")) for layer in layers if layer is not None]
    groups += [row.find("w:tblPrEx/w:tblCellMar", NS), cell.find("w:tcPr/w:tcMar", NS)]
    for group in groups:
        if group is None:
            continue
        for side, logical in (("left", "start"), ("right", "end")):
            node = group.find(q("w:" + logical))
            if node is None:
                node = group.find(q("w:" + side))
            if node is not None:
                values[side] = _width(node, 0, default=0)
    return sum(values.values())


def fit_tables(document, target_section, styles_root, *, settings_root=None):
    """Fit imported body tables to template columns and enclosing cells.

    Modified tables receive coherent fixed dxa geometry. Text, fields, row
    controls, merge markup and embedded objects are never rebuilt or scaled.
    Native layout and cell-content checks remain necessary after Word saves.
    """
    report = {"policy": "fit_table_grid_to_template", "tables": [], "resized": 0, "unresolved": []}
    tables = body_tables(document)
    if not tables:
        return report
    try:
        column_emu, _, column_policy = _limits(target_section, settings_root)
        limit = column_emu // EMU_PER_TWIP
        page_section = deepcopy(target_section)
        for columns in page_section.findall(q("w:cols")):
            page_section.remove(columns)
        page_width = _limits(page_section, settings_root)[0] // EMU_PER_TWIP
    except ValueError as exc:
        report["unresolved"].append(f"Table layout: {exc}")
        report["tables"] = [{"index": i, "status": "unresolved", "reason": str(exc)}
                            for i in range(1, len(tables) + 1)]
        return report
    cell_limits = {}
    failed = set()
    for table in tables:
        record = {"index": len(report["tables"]) + 1, "status": "unresolved", "column_policy": column_policy}
        report["tables"].append(record)
        try:
            parent_cell = next(table.iterancestors(q("w:tc")), None)
            parent_table = next(table.iterancestors(q("w:tbl")), None)
            if parent_table in failed:
                raise ValueError("containing table geometry is unresolved")
            containing = limit if parent_cell is None else cell_limits.get(parent_cell)
            if containing is None or containing <= 0:
                raise ValueError("containing cell has no positive inner width")
            layers, properties, style_chain = _properties(table, styles_root)
            rows, count = _row_map(table)
            first_exception = rows[0][0].find("w:tblPrEx/w:tblW", NS)
            preference = first_exception if first_exception is not None else properties.get("tblW")
            percentage_base = page_width if parent_cell is None else containing
            preferred = _width(preference, percentage_base)
            weights, inferred = _grid(table, rows, count, preferred or containing)
            requested = preferred or sum(weights)
            # Validate even an alignment-ignored dxa value before touching XML.
            # Invalid input must never leave a partially adapted table behind.
            base_indent = _indent(properties.get("tblInd"))
            row_limits, negative_indents, has_spacing = [], [], False
            for row, _, _ in rows:
                effective = dict(properties)
                exception = row.find(q("w:tblPrEx"))
                if exception is not None:
                    effective.update({ET.QName(child).localname: child for child in exception})
                validated_indent = _indent(effective.get("tblInd"))
                alignment = effective.get("jc")
                alignment = alignment.get(q("w:val"), "left") if alignment is not None else "left"
                indent = validated_indent if alignment in {"left", "start"} else 0
                spacing = _width(effective.get("tblCellSpacing"), containing, default=0)
                internal_spacing = _width(row.find("w:trPr/w:tblCellSpacing", NS), containing, default=0)
                row_limits.append(containing - max(0, indent) - 2 * spacing)
                has_spacing |= bool(spacing or internal_spacing)
                if indent < 0:
                    negative_indents.append(row)
            width_limit = min(row_limits)
            if width_limit <= 0:
                raise ValueError("table indentation or spacing consumes its containing width")
            target = min(requested, width_limit)
            adapt = inferred or bool(negative_indents) or requested > width_limit or preferred is not None and sum(weights) != target
            grid = _allocate(weights, target) if adapt else weights
            planned_cells = {}
            for row, cells, _ in rows:
                for cell, start, span in cells:
                    cell_width = sum(grid[start:start + span])
                    planned_cells[cell] = (cell_width, cell_width - _margins(layers, row, cell, style_chain))
                    if planned_cells[cell][1] <= 0:
                        raise ValueError("cell margins leave no positive inner width")
            record.update(before_grid_twips=[_number(node, "w") for node in table.findall("w:tblGrid/w:gridCol", NS)],
                          after_grid_twips=grid, grid_inferred=inferred, width_limit_twips=width_limit,
                          requested_width_twips=requested, containing_width_twips=containing,
                          status="resized" if adapt else "unchanged")
            if adapt:
                direct = table.find(q("w:tblPr"))
                if direct is None:
                    direct = ET.Element(q("w:tblPr"))
                    table.insert(0, direct)
                _set_width(_put(direct, "tblW"), sum(grid))
                _put(direct, "tblLayout").set(q("w:type"), "fixed")
                if base_indent < 0:
                    _set_width(_put(direct, "tblInd"), 0)
                for row in negative_indents:
                    exception = row.find(q("w:tblPrEx"))
                    if exception is None:
                        exception = ET.Element(q("w:tblPrEx"))
                        row.insert(0, exception)
                    _set_width(_put(exception, "tblInd"), 0)
                grid_node = table.find(q("w:tblGrid"))
                if grid_node is None:
                    grid_node = ET.Element(q("w:tblGrid"))
                    table.insert(table.index(direct) + 1, grid_node)
                for old in grid_node.findall(q("w:gridCol")):
                    grid_node.remove(old)
                for position, width in enumerate(grid):
                    grid_node.insert(position, ET.Element(q("w:gridCol"), {q("w:w"): str(width)}))
                # Word applies a first-row width exception to the whole table.
                if first_exception is not None:
                    _set_width(first_exception, sum(grid))
                for row, _, _ in rows:
                    exception_layout = row.find("w:tblPrEx/w:tblLayout", NS)
                    if exception_layout is not None:
                        exception_layout.set(q("w:type"), "fixed")
                report["resized"] += 1
            for row, cells, after in rows:
                if adapt:
                    tr_properties = row.find(q("w:trPr"))
                    for name, value in (("wBefore", sum(grid[:cells[0][1]]) if cells else 0),
                                        ("wAfter", sum(grid[len(grid) - after:]) if after else 0)):
                        node = tr_properties.find(q("w:" + name)) if tr_properties is not None else None
                        if node is not None:
                            _set_width(node, value)
                for cell, start, span in cells:
                    cell_width, inner_width = planned_cells[cell]
                    if adapt:
                        tc_properties = cell.find(q("w:tcPr"))
                        if tc_properties is None:
                            tc_properties = ET.Element(q("w:tcPr"))
                            cell.insert(0, tc_properties)
                        width = tc_properties.find(q("w:tcW"))
                        if width is None:
                            width = ET.Element(q("w:tcW"))
                            tc_properties.insert(0, width)
                        _set_width(width, cell_width)
                    cell_limits[cell] = inner_width
            if inferred:
                report["unresolved"].append(f"Table {record['index']}: missing grid columns were inferred; verify column proportions")
            if "tblpPr" in properties:
                report["unresolved"].append(f"Table {record['index']}: floating table position requires native placement review")
            if has_spacing:
                report["unresolved"].append(f"Table {record['index']}: nonzero cell spacing requires native placement review")
            if any(style.find("w:tblStylePr/w:tcPr/w:tcMar", NS) is not None
                   or style.find("w:tblStylePr/w:tblPr/w:tblCellMar", NS) is not None for style in style_chain):
                report["unresolved"].append(f"Table {record['index']}: conditional cell margins require native nested-content review")
        except ValueError as exc:
            record.update(status="unresolved", reason=str(exc))
            failed.add(table)
            report["unresolved"].append(f"Table {record['index']}: {exc}")
    return report
