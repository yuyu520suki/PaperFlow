"""MathType requirements must not pass for OMML, links, or pictures."""
import os
import struct
import tempfile
import unittest
import zipfile
from pathlib import Path

from paperflow.mathtype import (MathTypeError, MathTypeSDK, _CompoundFile,
                                audit_docx_math, discover_mathtype,
                                inspect_native_embedding)


class NativeMathAuditTests(unittest.TestCase):
    def package(self, body, rels="", embedding=None):
        work = tempfile.TemporaryDirectory()
        self.addCleanup(work.cleanup)
        target = Path(work.name) / "document.docx"
        document = ('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
                    'xmlns:o="urn:schemas-microsoft-com:office:office" '
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                    '<w:body>' + body + '</w:body></w:document>')
        with zipfile.ZipFile(target, "w") as z:
            z.writestr("word/document.xml", document)
            if rels:
                z.writestr("word/_rels/document.xml.rels",
                           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                           + rels + '</Relationships>')
            if embedding is not None:
                z.writestr("word/embeddings/oleObject1.bin", embedding)
        return target

    def test_omml_cannot_satisfy_native_mathtype(self):
        path = self.package('<w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>')
        report = audit_docx_math(path)
        self.assertEqual(report["omml_equations"], 1)
        self.assertEqual(report["native_mathtype_objects"], 0)
        self.assertFalse(report["native_equations_present"])
        self.assertFalse(report["edit_roundtrip_verified"])

    def test_progid_without_real_native_storage_is_rejected(self):
        path = self.package('<w:p><o:OLEObject Type="Embed" ProgID="Equation.DSMT4" r:id="rId1"/></w:p>',
                            '<Relationship Id="rId1" Target="embeddings/oleObject1.bin" '
                            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"/>',
                            b"Not a native equation")
        report = audit_docx_math(path)
        self.assertEqual(report["mathtype_objects"], 1)
        self.assertFalse(report["all_mathtype_objects_valid"])
        self.assertIn("Compound File", report["objects"][0]["errors"][0])

    def test_linked_equation_is_rejected(self):
        path = self.package('<w:p><o:OLEObject Type="Link" ProgID="Equation.DSMT4" r:id="rId1"/></w:p>',
                            '<Relationship Id="rId1" TargetMode="External" Target="file:///test.eqn" '
                            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"/>')
        report = audit_docx_math(path)
        self.assertFalse(report["native_equations_present"])
        self.assertIn("linked", report["objects"][0]["errors"][0])

    def test_missing_native_part_is_rejected(self):
        path = self.package('<w:p><o:OLEObject ProgID="Equation.DSMT4" r:id="rId1"/></w:p>')
        report = audit_docx_math(path)
        self.assertFalse(report["all_mathtype_objects_valid"])
        self.assertIn("Missing", report["objects"][0]["errors"][0])

    def test_equation_editor_object_is_not_mathtype(self):
        path = self.package('<w:p><o:OLEObject ProgID="Equation.3" r:id="rId1"/></w:p>')
        report = audit_docx_math(path)
        self.assertEqual(report["mathtype_objects"], 0)
        self.assertFalse(report["native_equations_present"])
        self.assertEqual(report["legacy_equation_objects"], 1)
        self.assertFalse(report["legacy_objects"][0]["valid_legacy_storage"])
        self.assertIn("Missing", report["legacy_objects"][0]["errors"][0])

    def test_bare_source_transform_fails_explicitly(self):
        with self.assertRaisesRegex(MathTypeError, "OLE host"):
            MathTypeSDK().to_mtef('<math><mi>x</mi></math>')

    def test_discovery_does_not_connect(self):
        sdk = MathTypeSDK()
        self.assertIsNone(sdk._dll)
        self.assertFalse(sdk._connected)

    @unittest.skipUnless(os.name == "nt", "Windows installation fixture")
    def test_installed_vendor_blank_equation_native_stream(self):
        installation = discover_mathtype()
        if not installation.directory:
            self.skipTest("MathType not installed")
        source = Path(installation.directory) / "Office Support" / "BlankEqn.doc"
        if not source.exists():
            self.skipTest("Vendor BlankEqn.doc is missing")
        compound = _CompoundFile(source.read_bytes())
        native = compound.stream("Equation Native")
        header = struct.unpack_from("<I", native, 0)[0]
        self.assertEqual(header, 28)
        self.assertEqual(native[header], 5)
        # A Word .doc hosting a MathType object is not itself a standalone
        # native embedding, and must not pass the DOCX-embedding audit.
        self.assertFalse(inspect_native_embedding(source.read_bytes())["valid_native_storage"])


if __name__ == "__main__":
    unittest.main()
