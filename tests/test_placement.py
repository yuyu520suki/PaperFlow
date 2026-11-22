import tempfile
import unittest
from pathlib import Path

from lxml import etree as ET

from paperflow.package import field_runs, part_fields, q, read_package, text_run, write_package, xml
from paperflow.placement import _pair_captions, append_end_objects


_MAIN = "word/document.xml"
_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_MASTER = Path(__file__).resolve().parents[1] / "output/computers-and-structures-001/editable-master.docx"


def _body(parts):
    return xml(parts[_MAIN]).find(q("w:body"))


def _table_source(path, *, caption=True, pending=False, ole=False):
    document = ET.Element(q("w:document"))
    body = ET.SubElement(document, q("w:body"))
    if caption:
        paragraph = ET.SubElement(body, q("w:p"))
        paragraph.append(text_run("Table "))
        fields = field_runs(" SEQ Table ", "PAPERFLOW_PENDING" if pending else "1")
        if not pending:
            fields[0][0].attrib.pop(q("w:dirty"), None)
        paragraph.extend(fields)
    table = ET.SubElement(body, q("w:tbl"))
    cell = ET.SubElement(ET.SubElement(table, q("w:tr")), q("w:tc"))
    paragraph = ET.SubElement(cell, q("w:p"))
    if ole:
        embedded = ET.SubElement(ET.SubElement(paragraph, q("w:r")), q("w:object"))
        ET.SubElement(embedded, q("o:OLEObject")).set("ProgID", "Equation.DSMT4")
    paragraph.append(text_run("Editable table value"))
    ET.SubElement(body, q("w:sectPr"))
    return write_package({_MAIN: ET.tostring(document)}, path)


class PlacementTests(unittest.TestCase):
    def test_consecutive_figure_caption_assignment_is_unique_in_both_directions(self):
        for order, expected in (("FCFCFC", {0: 1, 2: 3, 4: 5}), ("CFCFCF", {1: 0, 3: 2, 5: 4})):
            children = [ET.Element(q("w:p")) for _ in order]
            objects = {i: "figure" for i, tag in enumerate(order) if tag == "F"}
            captions = {i: "figure" for i, tag in enumerate(order) if tag == "C"}
            for i in objects:
                ET.SubElement(ET.SubElement(children[i], q("w:r")), q("w:drawing"))
            self.assertEqual(_pair_captions(children, objects, captions), expected)

    @unittest.skipUnless(_MASTER.is_file(), "Native journal integration fixture is unavailable")
    def test_native_fixture_preserves_fields_equations_and_picture_assets(self):
        before_bytes = _MASTER.read_bytes()
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / "submission.docx"
            report = append_end_objects(_MASTER, output)
            self.assertEqual((report["figure_copies"], report["table_copies"]), (1, 1))
            self.assertTrue(report["placement_verified"])
            self.assertFalse(report["pagination_verified"])
            before = read_package(_MASTER)
            after = read_package(output)
            self.assertEqual(before.keys(), after.keys())
            self.assertEqual([name for name in before if before[name] != after[name]], [_MAIN])
            original, result = _body(before), _body(after)
            self.assertEqual(
                [ET.tostring(node) for node in original[:-1]],
                [ET.tostring(node) for node in result[: len(original) - 1]],
            )
            self.assertEqual(part_fields(original), part_fields(result))
            appended = result[len(original) - 1 : -1]
            self.assertEqual(
                sum(1 for node in appended for _ in node.iter(q("w:pageBreakBefore"))),
                2,
            )
            self.assertEqual(len(list(result.iter(q("w:tbl")))), 2)
            drawing_ids = [node.get("id") for node in result.iter(f"{{{_WP}}}docPr")]
            self.assertEqual(len(set(drawing_ids)), 2)
            self.assertEqual(len(list(result.iter(q("o:OLEObject")))), 2)
            self.assertEqual(result[-1].tag, q("w:sectPr"))
        self.assertEqual(_MASTER.read_bytes(), before_bytes)

    def test_nested_table_is_copied_with_its_parent_under_one_caption(self):
        with tempfile.TemporaryDirectory() as temp:
            source, nested, output = [Path(temp) / name for name in ("source.docx", "nested.docx", "out.docx")]
            _table_source(source)
            parts = read_package(source)
            body = _body(parts)
            cell = next(body.iter(q("w:tc")))
            inner = ET.SubElement(cell, q("w:tbl"))
            inner_cell = ET.SubElement(ET.SubElement(inner, q("w:tr")), q("w:tc"))
            ET.SubElement(inner_cell, q("w:p")).append(text_run("Nested measurement"))
            ET.SubElement(cell, q("w:p"))
            parts[_MAIN] = ET.tostring(body.getparent())
            write_package(parts, nested)
            report = append_end_objects(nested, output)
            self.assertEqual(report["table_copies"], 1)
            tables = _body(read_package(output)).findall(q("w:tbl"))
            self.assertEqual(len(tables), 2)
            self.assertEqual(ET.tostring(tables[0]), ET.tostring(tables[1]))
            self.assertIn(b"Nested measurement", ET.tostring(tables[1]))

    def test_missing_caption_does_not_publish(self):
        with tempfile.TemporaryDirectory() as temp:
            source, output = [Path(temp) / name for name in ("source.docx", "out.docx")]
            _table_source(source, caption=False)
            with self.assertRaisesRegex(ValueError, "missing adjacent SEQ caption"):
                append_end_objects(source, output)
            self.assertFalse(output.exists())

    def test_pending_caption_does_not_publish(self):
        with tempfile.TemporaryDirectory() as temp:
            source, output = [Path(temp) / name for name in ("source.docx", "out.docx")]
            _table_source(source, pending=True)
            with self.assertRaisesRegex(ValueError, "Refresh caption"):
                append_end_objects(source, output)
            self.assertFalse(output.exists())

    def test_end_copies_scope_unrestricted_style_indexes_only(self):
        with tempfile.TemporaryDirectory() as temp:
            source, indexed, output = [Path(temp) / name for name in ("source.docx", "indexed.docx", "out.docx")]
            _table_source(source)
            parts = read_package(source)
            body = _body(parts)
            codes = [r' TOC \t "CustomCaption,1" ', r' TOC \t "CustomCaption,1" \b Existing ', r' TOC \c "Table" ']
            for index, code in enumerate(codes):
                paragraph = ET.Element(q("w:p"))
                paragraph.extend(field_runs(code, "Table 1"))
                body.insert(index, paragraph)
            parts[_MAIN] = ET.tostring(body.getparent())
            write_package(parts, indexed)
            report = append_end_objects(indexed, output)
            result = _body(read_package(output))
            scope = report["scoped_style_indexes"]
            self.assertEqual(scope["fields_scoped"], 1)
            expected = [codes[0].rstrip() + f' \\b {scope["bookmark"]} ', *codes[1:], " SEQ Table "]
            self.assertEqual(part_fields(result), (expected, []))
            self.assertEqual(result[0].get(q("w:name")), scope["bookmark"])
            self.assertEqual(result[-4].tag, q("w:bookmarkEnd"))
            self.assertEqual(result[0].get(q("w:id")), result[-4].get(q("w:id")))
            self.assertEqual(result[report["copies"][0]["end_body_index"]].tag, q("w:tbl"))

    def test_table_with_ole_requires_independent_embedding(self):
        with tempfile.TemporaryDirectory() as temp:
            source, output = [Path(temp) / name for name in ("source.docx", "out.docx")]
            _table_source(source, ole=True)
            with self.assertRaisesRegex(ValueError, "OLE object"):
                append_end_objects(source, output)
            self.assertFalse(output.exists())

    def test_existing_target_is_preserved(self):
        with tempfile.TemporaryDirectory() as temp:
            source, output = [Path(temp) / name for name in ("source.docx", "out.docx")]
            _table_source(source)
            output.write_bytes(b"existing submission")
            with self.assertRaises(FileExistsError):
                append_end_objects(source, output)
            self.assertEqual(output.read_bytes(), b"existing submission")


if __name__ == "__main__":
    unittest.main()
