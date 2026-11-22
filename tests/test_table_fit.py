"""Fit table geometry while preserving rows, scientific objects and fields."""
import base64
import io
import unittest
from copy import deepcopy

from docx import Document
from docx.shared import Inches
from lxml import etree as ET

from paperflow.package import NS, q
from paperflow.table_fit import fit_tables


NS = {**NS, "v": "urn:schemas-microsoft-com:vml"}
BODY_WIDTH = 9360


def section():
    node = ET.Element(q("w:sectPr"), nsmap={"w": NS["w"]})
    ET.SubElement(node, q("w:pgSz"), {q("w:w"): "12240", q("w:h"): "15840"})
    ET.SubElement(node, q("w:pgMar"), {
        q("w:top"): "1440", q("w:bottom"): "1440",
        q("w:left"): "1440", q("w:right"): "1440", q("w:gutter"): "0",
    })
    return node


def styles():
    node = ET.Element(q("w:styles"), nsmap={"w": NS["w"]})
    normal = ET.SubElement(node, q("w:style"), {
        q("w:type"): "paragraph", q("w:default"): "1", q("w:styleId"): "Normal",
    })
    ET.SubElement(normal, q("w:name"), {q("w:val"): "Normal"})
    return node


def width_node(parent, name, width, kind="dxa"):
    return ET.SubElement(parent, q("w:" + name), {q("w:type"): kind, q("w:w"): str(width)})


def table(grid=(12000, 6000), *, width=None, kind="dxa", layout="fixed"):
    node = ET.Element(q("w:tbl"))
    properties = ET.SubElement(node, q("w:tblPr"))
    width_node(properties, "tblW", (0 if kind == "auto" else sum(grid)) if width is None else width, kind)
    if layout is not None:
        ET.SubElement(properties, q("w:tblLayout"), {q("w:type"): layout})
    columns = ET.SubElement(node, q("w:tblGrid"))
    for value in grid:
        ET.SubElement(columns, q("w:gridCol"), {q("w:w"): str(value)})
    add_row(node, [1] * len(grid))
    return node


def grid_widths(node):
    return [int(column.get(q("w:w"))) for column in node.findall("w:tblGrid/w:gridCol", NS)]


def add_row(node, spans, *, before=0, after=0):
    grid = grid_widths(node)
    row = ET.SubElement(node, q("w:tr"))
    properties = ET.SubElement(row, q("w:trPr"))
    if before:
        ET.SubElement(properties, q("w:gridBefore"), {q("w:val"): str(before)})
        width_node(properties, "wBefore", sum(grid[:before]))
    if after:
        ET.SubElement(properties, q("w:gridAfter"), {q("w:val"): str(after)})
        width_node(properties, "wAfter", sum(grid[-after:]))
    offset = before
    for index, span in enumerate(spans):
        cell = ET.SubElement(row, q("w:tc"))
        cell_properties = ET.SubElement(cell, q("w:tcPr"))
        width_node(cell_properties, "tcW", sum(grid[offset:offset + span]))
        if span != 1:
            ET.SubElement(cell_properties, q("w:gridSpan"), {q("w:val"): str(span)})
        paragraph = ET.SubElement(cell, q("w:p"))
        run = ET.SubElement(paragraph, q("w:r"))
        ET.SubElement(run, q("w:t")).text = "Scientific measurement " + str(index + 1)
        offset += span
    return row


def document(*tables):
    root = ET.Element(q("w:document"), nsmap=NS)
    body = ET.SubElement(root, q("w:body"))
    for node in tables:
        body.append(node)
        ET.SubElement(body, q("w:p"))
    body.append(section())
    return root


def margins(parent, name="tblCellMar", *, left=100, right=100):
    node = ET.SubElement(parent, q("w:" + name))
    for edge, value in (("left", left), ("right", right)):
        width_node(node, edge, value)
    return node


def without_width_geometry(root):
    clone = deepcopy(root)
    allowed = {q("w:" + name) for name in ("tblW", "tcW", "wBefore", "wAfter", "tblLayout")}
    for node in list(clone.iter()):
        if node.tag in allowed:
            node.getparent().remove(node)
        elif node.tag == q("w:gridCol"):
            node.attrib.pop(q("w:w"), None)
    return ET.tostring(clone)


class TableFitTests(unittest.TestCase):
    def assert_consistent_cell_widths(self, node):
        grid = grid_widths(node)
        for row in node.findall(q("w:tr")):
            before = row.find("w:trPr/w:gridBefore", NS)
            after = row.find("w:trPr/w:gridAfter", NS)
            start = int(before.get(q("w:val"))) if before is not None else 0
            for cell in row.findall(q("w:tc")):
                span = cell.find("w:tcPr/w:gridSpan", NS)
                count = int(span.get(q("w:val"))) if span is not None else 1
                width = cell.find("w:tcPr/w:tcW", NS)
                self.assertEqual(width.get(q("w:type")), "dxa")
                self.assertEqual(int(width.get(q("w:w"))), sum(grid[start:start + count]))
                start += count
            for name, count in (("wBefore", int(before.get(q("w:val"))) if before is not None else 0),
                                ("wAfter", int(after.get(q("w:val"))) if after is not None else 0)):
                width = row.find("w:trPr/w:" + name, NS)
                if width is not None:
                    expected = sum(grid[:count]) if name == "wBefore" else sum(grid[-count:]) if count else 0
                    self.assertEqual(width.get(q("w:type")), "dxa")
                    self.assertEqual(int(width.get(q("w:w"))), expected)

    def assert_fixed_width(self, node, expected):
        width = node.find("w:tblPr/w:tblW", NS)
        self.assertEqual(width.get(q("w:type")), "dxa")
        self.assertEqual(int(width.get(q("w:w"))), expected)
        self.assertEqual(sum(grid_widths(node)), expected)
        self.assertEqual(node.find("w:tblPr/w:tblLayout", NS).get(q("w:type")), "fixed")
        self.assert_consistent_cell_widths(node)

    def test_shrinks_oversized_fixed_grid_proportionally_and_reports_widths(self):
        source_table = table()
        root, target, target_styles = document(source_table), section(), styles()
        before_target, before_styles = ET.tostring(target), ET.tostring(target_styles)
        report = fit_tables(root, target, target_styles)
        self.assertEqual(grid_widths(source_table), [6240, 3120])
        self.assert_fixed_width(source_table, BODY_WIDTH)
        self.assertEqual(report["resized"], 1)
        self.assertEqual(report["unresolved"], [])
        self.assertEqual(len(report["tables"]), 1)
        record = report["tables"][0]
        self.assertEqual(record["index"], 1)
        self.assertEqual(record["status"], "resized")
        self.assertEqual(record["before_grid_twips"], [12000, 6000])
        self.assertEqual(record["after_grid_twips"], [6240, 3120])
        self.assertEqual(record["width_limit_twips"], BODY_WIDTH)
        self.assertEqual(ET.tostring(target), before_target)
        self.assertEqual(ET.tostring(target_styles), before_styles)

    def test_real_python_docx_table_keeps_fonts_text_and_original_document(self):
        source = Document()
        source_table = source.add_table(rows=2, cols=2)
        source_table.autofit = False
        for column, width in zip(source_table.columns, (Inches(8), Inches(4))):
            column.width = width
        for row in source_table.rows:
            for cell, width in zip(row.cells, (Inches(8), Inches(4))):
                cell.width = width
                run = cell.paragraphs[0].add_run("Measured response")
                run.bold = True
                run.italic = True
        source_table._tbl.tblPr.find(q("w:tblW")).set(q("w:type"), "dxa")
        source_table._tbl.tblPr.find(q("w:tblW")).set(q("w:w"), "17280")
        original = ET.tostring(source._element)
        root = ET.fromstring(original)
        report = fit_tables(root, section(), styles())
        fitted = root.find("w:body/w:tbl", NS)
        self.assert_fixed_width(fitted, BODY_WIDTH)
        self.assertEqual(without_width_geometry(root), without_width_geometry(ET.fromstring(original)))
        self.assertEqual(ET.tostring(source._element), original)
        self.assertEqual(report["resized"], 1)

    def test_auto_and_percentage_tables_with_large_grids_are_fitted(self):
        for kind, width, layout in (("auto", 0, None), ("auto", 0, "autofit"),
                                    ("pct", "7500", "autofit"), ("pct", "150%", "fixed")):
            with self.subTest(kind=kind, width=width, layout=layout):
                source_table = table(kind=kind, width=width, layout=layout)
                report = fit_tables(document(source_table), section(), styles())
                self.assertEqual(grid_widths(source_table), [6240, 3120])
                self.assert_fixed_width(source_table, BODY_WIDTH)
                self.assertEqual(report["resized"], 1)

    def test_percentage_width_uses_target_and_can_expand_a_small_saved_grid(self):
        for percentage, expected in ((2500, BODY_WIDTH // 2), (5000, BODY_WIDTH)):
            with self.subTest(percentage=percentage):
                source_table = table((1000, 500), kind="pct", width=percentage, layout="autofit")
                fit_tables(document(source_table), section(), styles())
                self.assert_fixed_width(source_table, expected)
                self.assertEqual(grid_widths(source_table), [expected * 2 // 3, expected // 3])

    def test_does_not_enlarge_small_fixed_or_auto_tables_and_is_idempotent(self):
        for kind, layout in (("dxa", "fixed"), ("auto", None)):
            with self.subTest(kind=kind):
                root = document(table((1500, 2500), kind=kind, layout=layout))
                before = ET.tostring(root)
                report = fit_tables(root, section(), styles())
                self.assertEqual(grid_widths(root.find("w:body/w:tbl", NS)), [1500, 2500])
                self.assertEqual(report["resized"], 0)
                if kind == "dxa":
                    self.assertEqual(ET.tostring(root), before)
        root = document(table((10001, 10003, 10007)))
        first = fit_tables(root, section(), styles())
        after_first = ET.tostring(root)
        second = fit_tables(root, section(), styles())
        self.assertEqual(ET.tostring(root), after_first)
        self.assertEqual(first["resized"], 1)
        self.assertEqual(second["resized"], 0)
        self.assertEqual(second["tables"][0]["status"], "unchanged")
        self.assertEqual(second["unresolved"], [])

    def test_rounding_keeps_exact_total_positive_columns_and_proportions(self):
        original = [10001, 10003, 10007]
        source_table = table(original)
        fit_tables(document(source_table), section(), styles())
        fitted = grid_widths(source_table)
        self.assertEqual(sum(fitted), BODY_WIDTH)
        for before, after in zip(original, fitted):
            self.assertGreater(after, 0)
            self.assertLessEqual(abs(after - before * BODY_WIDTH / sum(original)), 1)
        self.assert_consistent_cell_widths(source_table)

    def test_fitted_grid_columns_precede_unchanged_grid_revision_history(self):
        source_table = table()
        grid = source_table.find(q("w:tblGrid"))
        history = ET.SubElement(grid, q("w:tblGridChange"), {
            q("w:id"): "19", q("w:author"): "Original author", q("w:date"): "2026-01-02T03:04:05Z",
        })
        previous_grid = ET.SubElement(history, q("w:tblGrid"))
        for width in (16000, 8000, 4000):
            ET.SubElement(previous_grid, q("w:gridCol"), {q("w:w"): str(width)})
        row = source_table.find(q("w:tr"))
        row_properties = row.find(q("w:trPr"))
        ET.SubElement(row_properties, q("w:tblHeader"))
        ET.SubElement(row_properties, q("w:cantSplit"))
        root = document(source_table)
        original_history = ET.tostring(history)
        original_row_properties = ET.tostring(row_properties)
        original_cell_data = [ET.tostring(cell.find(q("w:p"))) for cell in row.findall(q("w:tc"))]

        report = fit_tables(root, section(), styles())

        self.assert_fixed_width(source_table, BODY_WIDTH)
        self.assertEqual(grid_widths(source_table), [6240, 3120])
        self.assertEqual([child.tag for child in grid], [q("w:gridCol"), q("w:gridCol"), q("w:tblGridChange")])
        self.assertIs(grid[-1], history)
        self.assertEqual(ET.tostring(history), original_history)
        self.assertEqual(ET.tostring(row_properties), original_row_properties)
        self.assertEqual([ET.tostring(cell.find(q("w:p"))) for cell in row.findall(q("w:tc"))], original_cell_data)
        self.assertEqual(report["resized"], 1)
        self.assertEqual(report["unresolved"], [])

    def test_percentage_and_auto_cell_preferences_are_reconciled_to_the_fitted_grid(self):
        source_table = table()
        preferences = source_table.findall("w:tr/w:tc/w:tcPr/w:tcW", NS)
        preferences[0].set(q("w:type"), "pct")
        preferences[0].set(q("w:w"), "3300")
        preferences[1].set(q("w:type"), "auto")
        preferences[1].set(q("w:w"), "0")
        fit_tables(document(source_table), section(), styles())
        self.assertEqual(grid_widths(source_table), [6240, 3120])
        self.assert_fixed_width(source_table, BODY_WIDTH)

    def test_updates_merged_cell_widths_row_omissions_and_first_row_exception(self):
        source_table = table((6000, 3000, 9000))
        first_row = source_table.find(q("w:tr"))
        source_table.remove(first_row)
        row = add_row(source_table, [1], before=1, after=1)
        exception = ET.Element(q("w:tblPrEx"))
        width_node(exception, "tblW", 24000)
        row.insert(0, exception)
        add_row(source_table, [2, 1])
        fit_tables(document(source_table), section(), styles())
        self.assertEqual(grid_widths(source_table), [3120, 1560, 4680])
        self.assert_fixed_width(source_table, BODY_WIDTH)
        self.assertEqual(int(exception.find(q("w:tblW")).get(q("w:w"))), BODY_WIDTH)
        self.assertEqual(exception.find(q("w:tblW")).get(q("w:type")), "dxa")
        self.assertEqual(row.find("w:trPr/w:gridBefore", NS).get(q("w:val")), "1")
        self.assertEqual(row.find("w:trPr/w:gridAfter", NS).get(q("w:val")), "1")

    def test_preserves_scientific_payload_merges_fields_media_metadata_and_row_flags(self):
        source_table = table((9000, 4500, 4500))
        row = source_table.find(q("w:tr"))
        source_table.remove(row)
        row = add_row(source_table, [2, 1])
        second = add_row(source_table, [2, 1])
        ET.SubElement(row.find(q("w:trPr")), q("w:tblHeader"))
        ET.SubElement(row.find(q("w:trPr")), q("w:cantSplit"))
        ET.SubElement(row.find("w:tc/w:tcPr", NS), q("w:vMerge"), {q("w:val"): "restart"})
        ET.SubElement(second.find("w:tc/w:tcPr", NS), q("w:vMerge"))
        paragraph = row.find("w:tc/w:p", NS)
        payload = ET.fromstring(f"""<w:p xmlns:w="{NS['w']}" xmlns:r="{NS['r']}"
          xmlns:m="{NS['m']}" xmlns:o="{NS['o']}" xmlns:v="{NS['v']}">
          <w:r><w:rPr><w:rFonts w:ascii="Cambria"/><w:sz w:val="23"/><w:b/></w:rPr><w:t>7.25 MPa</w:t></w:r>
          <w:fldSimple w:instr="REF response \\h"><w:r><w:t>1</w:t></w:r></w:fldSimple>
          <m:oMath><m:r><m:t>x+1=2</m:t></m:r></m:oMath>
          <w:r><w:object w:dxaOrig="2800" w:dyaOrig="900"><v:shape id="_x0000_i1025" style="width:140pt;height:45pt"><v:imagedata r:id="rIdPreview"/></v:shape>
            <o:OLEObject ProgID="Equation.DSMT4" Type="Embed" ShapeID="_x0000_i1025" ObjectID="_1" r:id="rIdMathType"/>
          </w:object></w:r>
        </w:p>""".encode())
        for child in list(payload):
            paragraph.append(child)
        png = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jY9kAAAAASUVORK5CYII=")
        picture_doc = Document()
        picture_doc.add_paragraph().add_run().add_picture(io.BytesIO(png), width=Inches(1))
        picture_run = ET.fromstring(ET.tostring(picture_doc.paragraphs[0]._p[0]))
        paragraph.append(picture_run)
        root = document(source_table)
        before = deepcopy(root)
        image_before = ET.tostring(picture_run)
        report = fit_tables(root, section(), styles())
        self.assertEqual(without_width_geometry(root), without_width_geometry(before))
        self.assertEqual(ET.tostring(picture_run), image_before)
        self.assertEqual(report["resized"], 1)
        self.assert_consistent_cell_widths(source_table)

    def test_nested_tables_use_new_parent_cell_width_and_direct_margins(self):
        outer = table()
        margins(outer.find(q("w:tblPr")), left=100, right=100)
        cells = outer.findall("w:tr/w:tc", NS)
        margins(cells[0].find(q("w:tcPr")), "tcMar", left=200, right=300)
        first_inner, second_inner = table((9000, 9000)), table((6000, 6000))
        cells[0].insert(1, first_inner)
        cells[1].insert(1, second_inner)
        report = fit_tables(document(outer), section(), styles())
        self.assert_fixed_width(outer, BODY_WIDTH)
        self.assert_fixed_width(first_inner, 6240 - 200 - 300)
        self.assert_fixed_width(second_inner, 3120 - 100 - 100)
        self.assertEqual(len(report["tables"]), 3)
        self.assertEqual([record["width_limit_twips"] for record in report["tables"]], [BODY_WIDTH, 5740, 2920])
        self.assertEqual(report["resized"], 3)

    def test_nested_tables_use_inherited_logical_cell_margins_and_direct_override(self):
        target_styles = styles()
        base = ET.SubElement(target_styles, q("w:style"), {q("w:type"): "table", q("w:styleId"): "BaseTable"})
        inherited = ET.SubElement(ET.SubElement(base, q("w:tblPr")), q("w:tblCellMar"))
        width_node(inherited, "start", 300)
        width_node(inherited, "end", 200)
        child = ET.SubElement(target_styles, q("w:style"), {q("w:type"): "table", q("w:styleId"): "TargetTable"})
        ET.SubElement(child, q("w:basedOn"), {q("w:val"): "BaseTable"})
        outer = table()
        ET.SubElement(outer.find(q("w:tblPr")), q("w:tblStyle"), {q("w:val"): "TargetTable"})
        cells = outer.findall("w:tr/w:tc", NS)
        direct = ET.SubElement(cells[0].find(q("w:tcPr")), q("w:tcMar"))
        width_node(direct, "start", 400)
        first_inner, second_inner = table(), table()
        cells[0].insert(1, first_inner)
        cells[1].insert(1, second_inner)
        fit_tables(document(outer), section(), target_styles)
        self.assert_fixed_width(first_inner, 6240 - 400 - 200)
        self.assert_fixed_width(second_inner, 3120 - 300 - 200)

    def test_positive_indent_is_deducted_only_for_leading_alignment(self):
        for alignment, expected in (("left", BODY_WIDTH - 360), ("center", BODY_WIDTH), ("right", BODY_WIDTH)):
            with self.subTest(alignment=alignment):
                source_table = table()
                properties = source_table.find(q("w:tblPr"))
                ET.SubElement(properties, q("w:jc"), {q("w:val"): alignment})
                width_node(properties, "tblInd", 360)
                report = fit_tables(document(source_table), section(), styles())
                self.assert_fixed_width(source_table, expected)
                self.assertEqual(report["tables"][0]["width_limit_twips"], expected)

    def test_percentage_and_auto_indent_units_are_ignored(self):
        for kind in ("pct", "auto"):
            with self.subTest(kind=kind):
                source_table = table()
                width_node(source_table.find(q("w:tblPr")), "tblInd", 2500, kind)
                report = fit_tables(document(source_table), section(), styles())
                self.assert_fixed_width(source_table, BODY_WIDTH)
                self.assertEqual(report["tables"][0]["width_limit_twips"], BODY_WIDTH)

    def test_row_exceptions_constrain_outer_indent_spacing_and_justification(self):
        for alignment, indent, spacing, expected in (("left", 720, 0, BODY_WIDTH - 720),
                                                      ("right", 720, 0, BODY_WIDTH),
                                                      ("left", 0, 120, BODY_WIDTH - 240)):
            with self.subTest(alignment=alignment, indent=indent, spacing=spacing):
                source_table = table()
                row = source_table.find(q("w:tr"))
                exception = ET.Element(q("w:tblPrEx"))
                ET.SubElement(exception, q("w:jc"), {q("w:val"): alignment})
                width_node(exception, "tblInd", indent)
                width_node(exception, "tblCellSpacing", spacing)
                row.insert(0, exception)
                report = fit_tables(document(source_table), section(), styles())
                self.assert_fixed_width(source_table, expected)
                self.assertEqual(report["tables"][0]["width_limit_twips"], expected)

    def test_table_spacing_reserves_outer_edges_and_row_spacing_stays_internal(self):
        for row_level, expected in ((False, BODY_WIDTH - 240), (True, BODY_WIDTH)):
            with self.subTest(row_level=row_level):
                source_table = table()
                owner = source_table.find("w:tr/w:trPr", NS) if row_level else source_table.find(q("w:tblPr"))
                spacing = width_node(owner, "tblCellSpacing", 120)
                root = document(source_table)
                before_attributes = dict(spacing.attrib)
                report = fit_tables(root, section(), styles())
                self.assert_fixed_width(source_table, expected)
                self.assertEqual(report["tables"][0]["width_limit_twips"], expected)
                self.assertEqual(dict(spacing.attrib), before_attributes)
                if row_level:
                    self.assertTrue(report["unresolved"])

    def test_explicit_zero_cell_spacing_does_not_create_a_nonzero_spacing_warning(self):
        for row_level in (False, True):
            with self.subTest(row_level=row_level):
                source_table = table()
                owner = source_table.find("w:tr/w:trPr", NS) if row_level else source_table.find(q("w:tblPr"))
                width_node(owner, "tblCellSpacing", 0)
                report = fit_tables(document(source_table), section(), styles())
                self.assert_fixed_width(source_table, BODY_WIDTH)
                self.assertEqual(report["unresolved"], [])

    def test_equal_unequal_columns_and_percentage_page_basis(self):
        equal = section()
        ET.SubElement(equal, q("w:cols"), {q("w:num"): "2", q("w:space"): "720", q("w:equalWidth"): "1"})
        unequal = section()
        columns = ET.SubElement(unequal, q("w:cols"), {q("w:equalWidth"): "0"})
        ET.SubElement(columns, q("w:col"), {q("w:w"): "3000", q("w:space"): "720"})
        ET.SubElement(columns, q("w:col"), {q("w:w"): "5640"})
        for target, expected in ((equal, 4320), (unequal, 3000)):
            with self.subTest(expected=expected):
                source_table = table()
                report = fit_tables(document(source_table), target, styles())
                self.assert_fixed_width(source_table, expected)
                self.assertEqual(report["tables"][0]["width_limit_twips"], expected)
        source_table = table((1000, 500), kind="pct", width=2500)
        fit_tables(document(source_table), equal, styles())
        self.assert_fixed_width(source_table, 4320)  # Half of full-page text width, capped to the column.

    def test_missing_grid_is_inferred_from_cell_preferences_and_flagged(self):
        source_table = table((8000, 4000), kind="auto", layout=None)
        source_table.remove(source_table.find(q("w:tblGrid")))
        report = fit_tables(document(source_table), section(), styles())
        self.assert_fixed_width(source_table, BODY_WIDTH)
        self.assertEqual(grid_widths(source_table), [6240, 3120])
        self.assertTrue(report["unresolved"])

    def test_missing_grid_and_auto_cell_widths_get_equal_residual_with_review(self):
        source_table = table((1, 1, 1), kind="auto", layout=None)
        source_table.remove(source_table.find(q("w:tblGrid")))
        for node in source_table.findall("w:tr/w:tc/w:tcPr/w:tcW", NS):
            node.set(q("w:type"), "auto")
            node.set(q("w:w"), "0")
        report = fit_tables(document(source_table), section(), styles())
        self.assertEqual(grid_widths(source_table), [3120, 3120, 3120])
        self.assert_fixed_width(source_table, BODY_WIDTH)
        self.assertTrue(report["unresolved"])

    def test_legacy_horizontal_merge_markers_are_preserved(self):
        source_table = table((6000, 6000, 6000))
        cells = source_table.findall("w:tr/w:tc", NS)
        ET.SubElement(cells[0].find(q("w:tcPr")), q("w:hMerge"), {q("w:val"): "restart"})
        ET.SubElement(cells[1].find(q("w:tcPr")), q("w:hMerge"))
        root = document(source_table)
        before = deepcopy(root)
        fit_tables(root, section(), styles())
        self.assertEqual(grid_widths(source_table), [3120, 3120, 3120])
        self.assertEqual(without_width_geometry(root), without_width_geometry(before))

    def test_floating_table_keeps_position_but_width_is_fitted_and_reviewed(self):
        source_table = table()
        position = ET.SubElement(source_table.find(q("w:tblPr")), q("w:tblpPr"), {
            q("w:horzAnchor"): "margin", q("w:vertAnchor"): "paragraph",
            q("w:tblpX"): "1440", q("w:tblpY"): "0",
        })
        root = document(source_table)
        before = ET.tostring(position)
        report = fit_tables(root, section(), styles())
        self.assert_fixed_width(source_table, BODY_WIDTH)
        self.assertEqual(ET.tostring(position), before)
        self.assertEqual(report["resized"], 1)
        self.assertTrue(report["unresolved"])

    def test_invalid_grid_span_or_target_geometry_leaves_table_untouched(self):
        def invalid_grid(node):
            node.find("w:tblGrid/w:gridCol", NS).set(q("w:w"), "bad")

        def invalid_span(node):
            properties = node.find("w:tr/w:tc/w:tcPr", NS)
            ET.SubElement(properties, q("w:gridSpan"), {q("w:val"): "0"})

        for mutate, target in ((invalid_grid, section()), (invalid_span, section()),
                               (lambda node: None, None), (lambda node: None, ET.Element(q("w:sectPr")))):
            with self.subTest(mutate=mutate.__name__, target=target):
                source_table = table()
                mutate(source_table)
                root = document(source_table)
                before = ET.tostring(root)
                report = fit_tables(root, target, styles())
                self.assertEqual(ET.tostring(root), before)
                self.assertEqual(report["resized"], 0)
                self.assertEqual(report["tables"][0]["status"], "unresolved")
                self.assertTrue(report["unresolved"])

    def test_malformed_alignment_ignored_indents_never_partially_change_table(self):
        for alignment, row_override in (("center", False), ("right", False), ("center", True)):
            with self.subTest(alignment=alignment, row_override=row_override):
                source_table = table()
                properties = source_table.find(q("w:tblPr"))
                ET.SubElement(properties, q("w:jc"), {q("w:val"): alignment})
                if row_override:
                    properties = ET.Element(q("w:tblPrEx"))
                    source_table.find(q("w:tr")).insert(0, properties)
                width_node(properties, "tblInd", "invalid")
                root = document(source_table)
                before = ET.tostring(root)
                report = fit_tables(root, section(), styles())
                self.assertEqual(ET.tostring(root), before)
                self.assertEqual(report["resized"], 0)
                self.assertEqual(report["tables"][0]["status"], "unresolved")
                self.assertTrue(report["unresolved"])

    def test_row_leading_alignment_clamps_inherited_negative_indent(self):
        source_table = table()
        properties = source_table.find(q("w:tblPr"))
        ET.SubElement(properties, q("w:jc"), {q("w:val"): "center"})
        width_node(properties, "tblInd", -720)
        row = source_table.find(q("w:tr"))
        exception = ET.Element(q("w:tblPrEx"))
        ET.SubElement(exception, q("w:jc"), {q("w:val"): "left"})
        row.insert(0, exception)
        report = fit_tables(document(source_table), section(), styles())
        self.assertEqual(report["unresolved"], [])
        self.assertEqual(report["resized"], 1)
        self.assert_fixed_width(source_table, BODY_WIDTH)
        self.assertEqual(exception.find(q("w:tblInd")).get(q("w:w")), "0")

    def test_no_tables_means_no_missing_geometry_warning(self):
        report = fit_tables(document(), None, styles())
        self.assertEqual(report["tables"], [])
        self.assertEqual(report["resized"], 0)
        self.assertEqual(report["unresolved"], [])

    def test_deleted_and_textbox_story_tables_are_preserved_without_geometry_warnings(self):
        for wrapper_name in ("del", "moveFrom", "txbxContent"):
            with self.subTest(wrapper=wrapper_name):
                root = document()
                wrapper = ET.Element(q("w:" + wrapper_name))
                wrapper.append(table())
                root.find(q("w:body")).insert(0, wrapper)
                before = ET.tostring(root)
                report = fit_tables(root, None, styles())
                self.assertEqual(ET.tostring(root), before)
                self.assertEqual(report["tables"], [])
                self.assertEqual(report["resized"], 0)
                self.assertEqual(report["unresolved"], [])


if __name__ == "__main__":
    unittest.main()
