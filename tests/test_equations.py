"""Equation mapping and original/output protection, without Word automation."""
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from paperflow.equations import (EquationConversionError, _check_equation_alignment,
                                 _check_native_result, _math_layout,
                                 convert_to_mathtype, extract_math_from_ast)
from paperflow.word import OutputExistsError


def math_node(source, mode="InlineMath"):
    return {"t": "Math", "c": [{"t": mode}, source]}


class EquationMappingTests(unittest.TestCase):
    def test_nested_body_structures_preserve_document_order(self):
        # A citation contains ordinary strings and must not be mistaken for an
        # equation. Table-cell and list equations remain in their reading order.
        ast = {"pandoc-api-version": [1, 23, 1], "meta": {}, "blocks": [
            {"t": "Header", "c": [1, ["intro", [], []], [math_node("a^2")]]},
            {"t": "Para", "c": [{"t": "Cite", "c": [[{"citationId": "not_math"}], [{"t": "Str", "c": "[1]"}]]},
                                    math_node("b_1")]},
            {"t": "Table", "c": [["", [], []], [None, []], [],
                                    ["head", [{"t": "Plain", "c": [math_node("c+d")]}]],
                                    [["row", [{"t": "Plain", "c": [math_node(r"\frac{x}{y}", "DisplayMath")]}]]]]},
            {"t": "BulletList", "c": [[{"t": "Plain", "c": [math_node("z")] }]]},
        ]}
        self.assertEqual(extract_math_from_ast(ast), ["a^2", "b_1", "c+d", r"\frac{x}{y}", "z"])

    def test_note_or_metadata_equation_is_not_silently_reordered(self):
        for location in ("note", "metadata"):
            with self.subTest(location=location):
                ast = {"meta": {}, "blocks": []}
                if location == "note":
                    ast["blocks"] = [{"t": "Para", "c": [{"t": "Note", "c": [{"t": "Para", "c": [math_node("x")]}]}]}]
                else:
                    ast["meta"] = {"title": {"t": "MetaInlines", "c": [math_node("x")]}}
                with self.assertRaisesRegex(EquationConversionError, "Unsupported equation location"):
                    extract_math_from_ast(ast)

    def test_raw_or_empty_math_cannot_be_used_as_native_source(self):
        cases = [math_node(""), math_node("x", "UnknownMath"),
                 {"t": "RawInline", "c": ["openxml", "<m:oMath><m:r/></m:oMath>"]}]
        for item in cases:
            with self.subTest(item=item):
                with self.assertRaises(EquationConversionError):
                    extract_math_from_ast({"meta": {}, "blocks": [{"t": "Para", "c": [item]}]})

    def test_count_alignment_checks_both_extractor_and_word(self):
        with self.assertRaisesRegex(EquationConversionError, "source/XML mismatch"):
            _check_equation_alignment(["x"], 2, 2)
        with self.assertRaisesRegex(EquationConversionError, "Word/XML equation mismatch"):
            _check_equation_alignment(["x", "y"], 2, 1)
        with self.assertRaisesRegex(EquationConversionError, "ordered sequence"):
            _check_equation_alignment("xy", 2)
        _check_equation_alignment(["x", "y"], 2, 2)

    def test_native_audit_requires_exact_total_and_zero_omml(self):
        before = {"native_mathtype_objects": 2}
        good = {"native_mathtype_objects": 5, "omml_equations": 0, "all_mathtype_objects_valid": True}
        self.assertEqual(_check_native_result(before, good, 3), 5)
        for broken in ({**good, "native_mathtype_objects": 4},
                       {**good, "omml_equations": 1},
                       {**good, "all_mathtype_objects_valid": False}):
            with self.subTest(broken=broken):
                with self.assertRaisesRegex(EquationConversionError, "native audit"):
                    _check_native_result(before, broken, 3)


class InputProtectionTests(unittest.TestCase):
    def setUp(self):
        work = tempfile.TemporaryDirectory()
        self.addCleanup(work.cleanup)
        self.root = Path(work.name)

    def package(self, name="source.docx", body="", other_parts=None):
        source = self.root / name
        with ZipFile(source, "w") as z:
            z.writestr("word/document.xml", self.document(body))
            for part, data in (other_parts or {}).items():
                z.writestr(part, data)
        return source

    @staticmethod
    def document(body):
        return ('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
                '<w:body>' + body + '</w:body></w:document>')

    def test_existing_output_is_preserved_before_any_word_or_sdk_call(self):
        source = self.package()
        output = self.root / "existing.docx"
        output.write_bytes(b"User result")
        with patch("paperflow.equations._application") as word:
            with self.assertRaises(OutputExistsError):
                convert_to_mathtype(source, output)
            word.assert_not_called()
        self.assertEqual(output.read_bytes(), b"User result")

    def test_in_place_conversion_is_rejected(self):
        source = self.package()
        original = source.read_bytes()
        with self.assertRaisesRegex(ValueError, "Output must differ"):
            convert_to_mathtype(source, source)
        self.assertEqual(source.read_bytes(), original)

    def test_formula_free_copy_preserves_opaque_parts_without_word(self):
        source = self.package(body='<w:p><w:r><w:t>本文没有公式</w:t></w:r></w:p>',
                              other_parts={"customXml/item1.xml": b"<keep>opaque payload</keep>"})
        original = source.read_bytes()
        output = self.root / "copied.docx"
        with patch("paperflow.equations._application") as word, patch("paperflow.equations.find_pandoc") as pandoc:
            result = convert_to_mathtype(source, output)
            word.assert_not_called()
            pandoc.assert_not_called()
        self.assertEqual(source.read_bytes(), original)
        self.assertEqual(output.read_bytes(), original)
        self.assertEqual(result["expected_native_count"], 0)
        self.assertTrue(result["native_count_matches"])

    def test_header_textbox_and_nested_omml_fail_before_output_is_created(self):
        equation = '<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>'
        cases = [
            (equation, {"word/header1.xml": self.document(equation)}),
            ('<w:p><w:r><w:pict><w:txbxContent><w:p>' + equation + '</w:p></w:txbxContent></w:pict></w:r></w:p>', {}),
            ('<m:oMath>' + equation + '</m:oMath>', {}),
        ]
        for index, (body, extras) in enumerate(cases):
            source = self.package(f"source-{index}.docx", body, extras)
            output = self.root / f"failed-{index}.docx"
            with patch("paperflow.equations._application") as word:
                with self.assertRaisesRegex(EquationConversionError, "Unsupported OMML locations"):
                    convert_to_mathtype(source, output, equations=["x"])
                word.assert_not_called()
            self.assertFalse(output.exists())

    def test_explicit_source_count_mismatch_does_not_touch_word_or_source(self):
        source = self.package(body='<w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>')
        original = source.read_bytes()
        with patch("paperflow.equations._application") as word:
            with self.assertRaisesRegex(EquationConversionError, "source/XML mismatch"):
                convert_to_mathtype(source, self.root / "output.docx", equations=[])
            word.assert_not_called()
        self.assertEqual(source.read_bytes(), original)
        self.assertFalse((self.root / "output.docx").exists())

    def test_invalid_legacy_equation_fails_before_storage_or_word_operations(self):
        source = self.package(body='<w:p><o:OLEObject xmlns:o="urn:schemas-microsoft-com:office:office" '
                                   'ProgID="Equation.3"/></w:p>')
        original = source.read_bytes()
        output = self.root / "output.docx"
        with patch("paperflow.legacy_equations._vendor_blank") as storage, patch("paperflow.equations._application") as word:
            with self.assertRaisesRegex(EquationConversionError, "Unusable Microsoft Equation 3 object"):
                convert_to_mathtype(source, output)
            storage.assert_not_called()
            word.assert_not_called()
        self.assertEqual(source.read_bytes(), original)
        self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
