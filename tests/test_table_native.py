"""Saved-table layout checks exercise native geometry without opening Word."""
from contextlib import nullcontext
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from paperflow.package import NS
from paperflow.table_native import audit_table_layout


class Collection:
    def __init__(self, values):
        self.values = list(values)

    @property
    def Count(self):
        return len(self.values)

    def Item(self, index):
        return self.values[index - 1]


class NativeTable:
    def __init__(self, start, end, cells, *, level=1, nested=(), rows=None):
        page = SimpleNamespace(PageWidth=500, LeftMargin=50, RightMargin=50, Gutter=0,
                               TextColumns=Collection([SimpleNamespace(Width=400)]))
        self.Range = SimpleNamespace(Start=start, End=end, StoryType=1, Cells=Collection(cells),
                                     Sections=Collection([SimpleNamespace(PageSetup=page)]))
        self.NestingLevel = level
        self.Tables = Collection(nested)
        self.rows = rows or SimpleNamespace(LeftIndent=0, Alignment=0, WrapAroundText=0)

    @property
    def Rows(self):
        if isinstance(self.rows, Exception):
            raise self.rows
        return self.rows


def cell(start, end, width, *, row=1, column=1, level=1, padding=10):
    return SimpleNamespace(Range=SimpleNamespace(Start=start, End=end), Width=width,
                           RowIndex=row, ColumnIndex=column, NestingLevel=level,
                           LeftPadding=padding, RightPadding=padding)


def inline_shape(start, width, containing_cell, *, kind=1, left_indent=0, right_indent=0):
    location = SimpleNamespace(Start=start, End=start + 1, StoryType=1,
                               Cells=Collection([containing_cell]), Information=lambda kind: True,
                               ParagraphFormat=SimpleNamespace(LeftIndent=left_indent, RightIndent=right_indent))
    return SimpleNamespace(Type=kind, Width=width, Range=location)


def source_document(folder, *, table=True, body_xml=None):
    path = Path(folder) / "saved.docx"
    body = "<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>" if table else "<w:p/>"
    if body_xml is not None:
        body = body_xml
    with ZipFile(path, "w") as package:
        package.writestr("word/document.xml", f'<w:document xmlns:w="{NS["w"]}" xmlns:m="{NS["m"]}" '
                         'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
                         f'<w:body>{body}</w:body></w:document>')
        package.writestr("word/header1.xml", f'<w:hdr xmlns:w="{NS["w"]}"><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:hdr>')
        package.writestr("word/embeddings/oleObject1.bin", b"native equation payload")
    return path


def inspect_mock(source, tables, shapes=()):
    document = SimpleNamespace(Tables=Collection(tables), InlineShapes=Collection(shapes))
    with patch("paperflow.table_native._opened", return_value=nullcontext(document)) as opened:
        report = audit_table_layout(source, app=object())
    opened.assert_called_once()
    return report


class NativeTableLayoutTests(unittest.TestCase):
    def test_negative_leading_indent_requires_placement_review_even_when_width_fits(self):
        rows = SimpleNamespace(LeftIndent=-100, Alignment=0, WrapAroundText=0)
        table = NativeTable(0, 101, [cell(1, 100, 300)], rows=rows)
        with tempfile.TemporaryDirectory() as temporary:
            report = inspect_mock(source_document(temporary), [table])
        record, = report["tables"]
        self.assertTrue(record["within_width_limit"])
        self.assertTrue(record["placement_needs_review"])
        self.assertTrue(any("negative leading indent" in issue for issue in report["unresolved"]))

    def test_omitted_grid_columns_cannot_hide_an_overwide_saved_table(self):
        body = """<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="8000"/>
          <w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr><w:gridBefore w:val="1"/>
          <w:gridAfter w:val="1"/></w:trPr><w:tc><w:p/></w:tc></w:tr></w:tbl>"""
        table = NativeTable(0, 101, [cell(1, 100, 400, column=2)])
        with tempfile.TemporaryDirectory() as temporary:
            report = inspect_mock(source_document(temporary, body_xml=body), [table])
        record, = report["tables"]
        self.assertEqual(record["measured_width_pt"], 400)
        self.assertEqual(record["saved_grid_width_pt"], 600)
        self.assertFalse(record["within_width_limit"])
        self.assertTrue(any("saved logical grid width" in issue for issue in report["unresolved"]))

    def test_fitting_native_cells_keep_nested_content_constraints_visible_for_review(self):
        body = """
        <w:tbl><w:tr><w:tc><w:tcPr><w:noWrap w:val="false"/></w:tcPr><w:p/>
          <w:tbl><w:tr><w:tc><w:tcPr><w:noWrap/></w:tcPr>
            <w:p><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath>
              <w:r><w:drawing><wp:anchor/></w:drawing></w:r></w:p>
          </w:tc></w:tr></w:tbl><w:p/>
        </w:tc></w:tr></w:tbl><w:p/>
        <w:tbl><w:tr><w:tc><w:tcPr><w:noWrap w:val="0"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>
        """
        nested = NativeTable(20, 80, [cell(21, 79, 200, level=2)], level=2)
        outer = NativeTable(0, 101, [cell(1, 100, 350)], nested=[nested])
        separate = NativeTable(110, 150, [cell(111, 149, 100)])
        with tempfile.TemporaryDirectory() as temporary:
            report = inspect_mock(source_document(temporary, body_xml=body), [separate, outer])
        self.assertEqual(len(report["tables"]), 3)
        self.assertTrue(all(table["within_width_limit"] for table in report["tables"]))
        self.assertEqual(len(report["unresolved"]), 3)
        self.assertTrue(all(issue.startswith("Table 2 ") for issue in report["unresolved"]))
        self.assertTrue(any("nonwrapping" in issue for issue in report["unresolved"]))
        self.assertTrue(any("Word equations" in issue for issue in report["unresolved"]))
        self.assertTrue(any("floating drawing" in issue for issue in report["unresolved"]))

    def test_nested_table_and_inline_objects_use_actual_cell_content_widths(self):
        left = cell(1, 100, 200)
        right = cell(101, 200, 200, column=2)
        inner_cell = cell(21, 79, 190, level=2, padding=5)
        nested = NativeTable(20, 80, [inner_cell], level=2)
        outer = NativeTable(0, 201, [left, inner_cell, right], nested=[nested])
        shapes = [inline_shape(22, 185, inner_cell, left_indent=10, right_indent=5),
                  inline_shape(23, 160, inner_cell, kind=3, left_indent=10, right_indent=5)]
        with tempfile.TemporaryDirectory() as temporary:
            source = source_document(temporary)
            before = source.read_bytes()
            # Document and nested collections can both expose the same table.
            report = inspect_mock(source, [nested, outer], shapes)
            self.assertEqual(source.read_bytes(), before)
        first, second = report["tables"]
        self.assertEqual(len(first["cells"]), 2)
        self.assertEqual(first["measured_width_pt"], 400)
        self.assertTrue(first["within_width_limit"])
        self.assertEqual(second["parent_table_index"], first["index"])
        self.assertEqual(second["width_limit_pt"], 180)
        self.assertFalse(second["within_width_limit"])
        equation, picture = report["content"]
        self.assertEqual(equation["width_limit_pt"], 165)
        self.assertFalse(equation["within_width_limit"])
        self.assertTrue(picture["within_width_limit"])
        self.assertEqual(equation["table_index"], second["index"])
        self.assertTrue(report["native_reopened"])
        self.assertTrue(report["source_unchanged"])
        self.assertEqual(len(report["unresolved"]), 2)

    def test_merged_row_collection_failure_keeps_cell_geometry_and_metric_limitations(self):
        cells = [cell(1, 150, 150), cell(151, 160, 200, column=2),
                 cell(161, 170, 200, row=2, column=2)]
        table = NativeTable(0, 171, cells, rows=RuntimeError("merged rows are unavailable"))
        with tempfile.TemporaryDirectory() as temporary:
            report = inspect_mock(source_document(temporary), [table])
        (record,) = report["tables"]
        self.assertEqual(record["row_widths_pt"], {1: 350, 2: 200})
        self.assertEqual(record["measured_width_pt"], 350)
        self.assertEqual(record["width_limit_pt"], 400)
        self.assertTrue(record["cells_complete"])
        self.assertTrue(record["within_width_limit"])
        self.assertIn("indent_note", record)
        self.assertIn("continuation", report["width_metric_note"])
        self.assertTrue(any("floating placement" in issue for issue in report["unresolved"]))

    def test_undefined_cell_width_or_padding_requires_review(self):
        for undefined in ("width", "padding"):
            with self.subTest(undefined=undefined), tempfile.TemporaryDirectory() as temporary:
                parent_cell = cell(1, 100, 9999999 if undefined == "width" else 200,
                                   padding=9999999 if undefined == "padding" else 10)
                inner_cell = cell(21, 79, 100, level=2)
                nested = NativeTable(20, 80, [inner_cell], level=2)
                outer = NativeTable(0, 101, [parent_cell], nested=[nested])
                report = inspect_mock(source_document(temporary), [outer])
                first, second = report["tables"]
                self.assertTrue(report["unresolved"])
                self.assertIsNone(second["width_limit_pt"])
                self.assertIsNone(second["within_width_limit"])
                if undefined == "width":
                    self.assertFalse(first["cells_complete"])
                    self.assertIsNone(first["measured_width_pt"])
                else:
                    self.assertIsNone(first["cells"][0]["left_padding_pt"])
                    self.assertTrue(any("padding" in issue for issue in report["unresolved"]))

    def test_disconnected_word_server_propagates_instead_of_becoming_review(self):
        class DisconnectedServer(RuntimeError):
            hresult = -2147417848

        table = NativeTable(0, 101, [cell(1, 100, 200)], rows=DisconnectedServer("Word exited"))
        with tempfile.TemporaryDirectory() as temporary:
            source = source_document(temporary)
            before = source.read_bytes()
            with self.assertRaisesRegex(DisconnectedServer, "Word exited"):
                inspect_mock(source, [table])
            self.assertEqual(source.read_bytes(), before)

    def test_no_body_table_does_not_open_word_even_when_header_has_table(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = source_document(temporary, table=False)
            with patch("paperflow.table_native._application", side_effect=AssertionError("Word must not start")), \
                    patch("paperflow.table_native._opened") as opened:
                report = audit_table_layout(source)
            opened.assert_not_called()
        self.assertFalse(report["applicable"])
        self.assertFalse(report["native_reopened"])
        self.assertEqual(report["tables"], [])
        self.assertEqual(report["content"], [])
        self.assertEqual(report["unresolved"], [])


if __name__ == "__main__":
    unittest.main()
