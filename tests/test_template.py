from pathlib import Path
import tempfile
import unittest

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches

from paperflow.package import read_package, xml
from paperflow.template import DEFAULT_SLOT, inventory_template, prepare_template, select_template_section


class TemplateTests(unittest.TestCase):
    def test_prepare_anchors_after_section_break_and_keeps_template_parts(self):
        with tempfile.TemporaryDirectory() as temp:
            original = Path(temp) / "original.docx"
            prepared = Path(temp) / "prepared.docx"
            document = Document()
            document.add_paragraph("Cover")
            document.sections[0].header.paragraphs[0].text = "Cover header"
            main = document.add_section()
            main.page_width = Inches(11)
            main.header.is_linked_to_previous = False
            main.header.paragraphs[0].text = "Article header"
            document.add_paragraph("Article headings")
            document.save(original)
            before = original.read_bytes()

            report = inventory_template(original)
            self.assertEqual(report["insertion_proposal"]["section_index"], 2)
            boundary = next(anchor for anchor in report["paragraph_anchors"]
                            if anchor["insertion_section_index"] != anchor["section_index"])
            with self.assertRaisesRegex(ValueError, "SHA-256"):
                prepare_template(original, prepared, after_paragraph=boundary["body_child_index"], expected_sha256="stale")
            self.assertFalse(prepared.exists())
            result = prepare_template(original, prepared, after_paragraph=boundary["body_child_index"],
                                      expected_sha256=report["sha256"])

            self.assertEqual(original.read_bytes(), before)
            parts_before, parts_after = read_package(original), read_package(prepared)
            self.assertEqual(parts_before.keys(), parts_after.keys())
            self.assertEqual([part for part in parts_before if parts_before[part] != parts_after[part]], ["word/document.xml"])
            self.assertEqual(result["location"]["section_index"], 2)
            self.assertEqual(result["location"]["pgSz"]["w"], "15840")
            self.assertEqual([p.text for p in Document(prepared).paragraphs], ["Cover", "", DEFAULT_SLOT, "Article headings"])

    def test_split_run_marker_uses_section_before_break_and_rejects_duplicates(self):
        document = Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("{{PAPERFLOW_")
        paragraph.add_run("BODY}}")
        document.add_section()
        document.add_paragraph("Article")
        _, location = select_template_section(xml(document._element.xml.encode()), DEFAULT_SLOT)
        self.assertEqual(location["section_index"], 1)

        document.add_table(rows=1, cols=1).cell(0, 0).text = DEFAULT_SLOT
        with self.assertRaisesRegex(ValueError, "exactly once"):
            select_template_section(xml(document._element.xml.encode()), DEFAULT_SLOT)

    def test_rejects_cross_section_bookmark(self):
        document = Document()
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:name"), "body")
        start.set(qn("w:id"), "2")
        document.add_paragraph("Cover")._p.append(start)
        document.add_section()
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), "2")
        document.add_paragraph("Article")._p.append(end)
        with self.assertRaisesRegex(ValueError, "crosses a section boundary"):
            select_template_section(xml(document._element.xml.encode()), "bookmark:body")
