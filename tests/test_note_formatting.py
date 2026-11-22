from pathlib import Path
import tempfile
import unittest

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from lxml import etree as ET

from paperflow.formatting import normalize_to_template
from paperflow.package import NS, q, read_package, xml


def _note_doc(path):
    doc = Document()
    body_mark = ET.SubElement(doc.add_paragraph("Main text").add_run()._r, q("w:footnoteReference"))
    body_mark.set(q("w:id"), "7")
    body_mark.set(q("w:customMarkFollows"), "1")
    note_xml = f'''<w:footnotes xmlns:w="{NS['w']}" xmlns:m="{NS['m']}">
      <w:footnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
      <w:footnote w:id="7"><w:p><w:pPr><w:spacing w:after="200"/></w:pPr>
        <w:r><w:rPr><w:vertAlign w:val="baseline"/></w:rPr><w:footnoteRef/></w:r>
        <w:r><w:rPr><w:sz w:val="42"/><w:i/></w:rPr><w:t>Note text</w:t></w:r>
        <m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>
      </w:p></w:footnote>
      <w:footnote w:id="9"><w:p><w:r><w:rPr><w:sz w:val="52"/></w:rPr><w:t>Other slot</w:t></w:r></w:p></w:footnote>
    </w:footnotes>'''
    part = Part(PackURI("/word/other-note-part.xml"), CT.WML_FOOTNOTES, note_xml.encode(), doc.part.package)
    doc.part.relate_to(part, RT.FOOTNOTES)
    doc.save(path)


class NoteFormattingTests(unittest.TestCase):
    def test_formats_referenced_note_and_preserves_separator_and_other_slot(self):
        with tempfile.TemporaryDirectory() as temp:
            source, template, output = [Path(temp) / name for name in ("src.docx", "template.docx", "out.docx")]
            _note_doc(source)
            destination = Document()
            body = destination.styles.add_style("Thesis Note", WD_STYLE_TYPE.PARAGRAPH)
            marker = destination.styles.add_style("Footnote Reference", WD_STYLE_TYPE.CHARACTER, builtin=True)
            marker.font.superscript = True
            destination.save(template)
            report = normalize_to_template(source, template, output, role_style_map={"footnote": body.style_id})
            self.assertEqual(report["notes"]["footnote"], {
                "referenced_notes": 1, "text_style": body.style_id, "reference_style": marker.style_id,
                "styled_markers": 2,
            })
            parts = read_package(output)
            root = xml(parts["word/other-note-part.xml"])
            note = root.find('w:footnote[@w:id="7"]', NS)
            self.assertEqual(note.find("w:p/w:pPr/w:pStyle", NS).get(q("w:val")), body.style_id)
            self.assertIsNone(note.find(".//w:spacing", NS))
            self.assertIsNone(note.find(".//w:sz", NS))
            self.assertIsNotNone(note.find(".//w:i", NS))
            self.assertEqual(note.find(".//m:t", NS).text, "x")
            self.assertIsNone(note.find(".//w:vertAlign", NS))
            document = xml(parts["word/document.xml"])
            reference = document.find(".//w:footnoteReference", NS)
            self.assertEqual(reference.get(q("w:customMarkFollows")), "1")
            self.assertEqual(reference.getparent().find("w:rPr/w:rStyle", NS).get(q("w:val")), marker.style_id)
            original = xml(read_package(source)["word/other-note-part.xml"])
            for identifier in ("-1", "9"):
                selector = f'w:footnote[@w:id="{identifier}"]'
                self.assertEqual(ET.tostring(root.find(selector, NS)), ET.tostring(original.find(selector, NS)))

    def test_template_without_note_styles_preserves_note_part(self):
        with tempfile.TemporaryDirectory() as temp:
            source, template, output = [Path(temp) / name for name in ("src.docx", "template.docx", "out.docx")]
            _note_doc(source)
            Document().save(template)
            report = normalize_to_template(source, template, output)
            self.assertIsNone(report["notes"]["footnote"]["text_style"])
            part = "word/other-note-part.xml"
            self.assertEqual(read_package(output)[part], read_package(source)[part])


if __name__ == "__main__":
    unittest.main()
