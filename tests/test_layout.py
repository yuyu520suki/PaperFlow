from pathlib import Path
import tempfile
import unittest

from docx import Document
from docx.enum.section import WD_SECTION

from paperflow.layout import inspect_layout, split_source, validate_layout
from paperflow.package import field_runs, read_package
from paperflow.template import section_story_signature


class LayoutTests(unittest.TestCase):
    def test_plan_uses_template_order_and_rejects_an_open_word_field_at_a_cut(self):
        with tempfile.TemporaryDirectory() as folder:
            base = Path(folder)
            source = base / "source.docx"
            template = base / "template.docx"
            doc = Document()
            doc.add_paragraph("{{PAPERFLOW_Z}}")
            doc.add_paragraph("{{PAPERFLOW_A}}")
            doc.save(template)
            doc = Document()
            runs = field_runs(" REF name ", "Cached result")
            doc.add_paragraph()._p.extend(runs[:2])
            doc.add_paragraph()._p.extend(runs[2:])
            doc.add_paragraph("Separate content")
            doc.save(source)

            plan = inspect_layout(source, template)["plan"]
            first, second = plan["assignments"]
            self.assertEqual([first["slot"], second["slot"]], ["{{PAPERFLOW_Z}}", "{{PAPERFLOW_A}}"])
            first.update(start=1, end=1)
            second.update(start=2, end=3)
            with self.assertRaisesRegex(ValueError, "cuts a range.*Word field"):
                validate_layout(source, template, plan)
            first["end"] = 2
            second["start"] = 3
            parts = split_source(source, validate_layout(source, template, plan), base / "parts")
            for part in parts:
                before, after = read_package(source), read_package(part)
                self.assertEqual({k: v for k, v in before.items() if k != "word/document.xml"},
                                 {k: v for k, v in after.items() if k != "word/document.xml"})

    def test_plan_allows_forward_references_but_rejects_caption_bookmark_collisions(self):
        with tempfile.TemporaryDirectory() as folder:
            base = Path(folder)
            source = base / "source.docx"
            template = base / "template.docx"
            destination = Document()
            destination.add_paragraph()._p.extend(field_runs(" REF results ", "Cached"))
            destination.add_paragraph("{{PAPERFLOW_COVER}}")
            destination.add_paragraph("{{PAPERFLOW_BODY}}")
            destination.save(template)

            def plan_for(first_marker):
                manuscript = Document()
                manuscript.add_paragraph(first_marker)
                manuscript.add_paragraph("[[CAPTIONNO Table results]] Table caption")
                manuscript.save(source)
                plan = inspect_layout(source, template)["plan"]
                for number, assignment in enumerate(plan["assignments"], 1):
                    assignment.update(start=number, end=number)
                return plan

            self.assertEqual(len(validate_layout(source, template, plan_for("See [[REF results]]."))), 2)
            with self.assertRaisesRegex(ValueError, "Bookmark names must be distinct.*results"):
                validate_layout(source, template, plan_for("Equation [[EQNO Results]]"))

    def test_story_signature_ignores_caches_and_detects_inherited_footer(self):
        with tempfile.TemporaryDirectory() as folder:
            first, cached, changed = (Path(folder) / name for name in ("before.docx", "cached.docx", "changed.docx"))
            doc = Document()
            doc.sections[0].footer.paragraphs[0].text = "Cover"
            body = doc.add_section(WD_SECTION.NEW_PAGE)
            body.footer.is_linked_to_previous = False
            footer = body.footer.paragraphs[0]
            footer.text = "Article "
            footer._p.extend(field_runs(" PAGE ", "1"))
            doc.save(first)
            footer.runs[-2].text = "7"
            doc.save(cached)
            self.assertEqual(section_story_signature(first), section_story_signature(cached))
            body.footer.is_linked_to_previous = True
            doc.save(changed)
            self.assertNotEqual(section_story_signature(first), section_story_signature(changed))


if __name__ == "__main__":
    unittest.main()
