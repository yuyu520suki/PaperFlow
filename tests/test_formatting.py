import tempfile
import unittest
from pathlib import Path
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from paperflow.formatting import normalize_to_template
from paperflow.package import read_package, xml, NS


class FormattingTests(unittest.TestCase):
    def test_inline_picture_uses_selected_template_section_and_destination_defaults(self):
        with tempfile.TemporaryDirectory() as temp:
            source, template, output = [Path(temp) / name for name in ("source.docx", "template.docx", "out.docx")]
            draft = Document()
            asset = Path(__file__).resolve().parents[1] / "examples/computers-and-structures/response.png"
            draft.add_picture(str(asset), width=Inches(9))
            # Source document defaults must not override the destination layout.
            default = draft.styles.element.find(qn("w:docDefaults")).find(qn("w:pPrDefault")).find(qn("w:pPr"))
            source_indent = OxmlElement("w:ind")
            source_indent.set(qn("w:left"), "2880")
            default.append(source_indent)
            draft.save(source)
            reference = Document()
            reference.sections[0].page_width = Inches(12)
            reference.add_paragraph("Cover")
            body_section = reference.add_section(WD_SECTION.NEW_PAGE)
            body_section.page_width = Inches(7)
            body_section.left_margin = body_section.right_margin = Inches(1)
            reference.styles["Normal"].paragraph_format.left_indent = Inches(0.5)
            reference.add_paragraph("{{PAPERFLOW_BODY}}")
            reference.save(template)
            original_source, original_template = source.read_bytes(), template.read_bytes()

            result = normalize_to_template(source, template, output)
            self.assertEqual(result["template_location"]["section_index"], 2)
            self.assertEqual(result["picture_fit"]["resized"], 1)
            self.assertEqual(result["picture_fit"]["unresolved"], [])
            shape = Document(output).inline_shapes[0]
            self.assertEqual(shape.width, Inches(4.5))
            self.assertEqual(shape.height, Inches(2.25))
            original_parts, output_parts = read_package(source), read_package(output)
            self.assertEqual({k: v for k, v in original_parts.items() if k.startswith("word/media/")},
                             {k: v for k, v in output_parts.items() if k.startswith("word/media/")})
            self.assertEqual(source.read_bytes(), original_source)
            self.assertEqual(template.read_bytes(), original_template)

    def test_outline_role_uses_direct_level_and_preserves_builtin_heading_priority(self):
        with tempfile.TemporaryDirectory() as temp:
            source, template, output = [Path(temp) / name for name in ("source.docx", "template.docx", "out.docx")]
            draft = Document()
            heading = draft.styles.add_style("Source Chapter", WD_STYLE_TYPE.PARAGRAPH)
            heading.base_style = draft.styles["Heading 1"]
            body = draft.styles.add_style("Source Body", WD_STYLE_TYPE.PARAGRAPH)
            body.base_style = heading
            body._element.get_or_add_pPr().get_or_add_outlineLvl().val = 9
            for text, style, outline in (
                ("Direct subheading", "Normal", 1),
                ("Custom style with a lower level", heading, 2),
                ("Custom style as body", heading, 9),
                ("Body from inherited style override", body, None),
                ("Native chapter", "Heading 1", 9),
                ("Caption with body outline level", "Caption", 9),
            ):
                paragraph = draft.add_paragraph(text, style=style)
                if outline is not None:
                    paragraph._p.get_or_add_pPr().get_or_add_outlineLvl().val = outline
            draft.save(source)
            reference = Document()
            subheading = reference.styles.add_style("Target Section", WD_STYLE_TYPE.PARAGRAPH)
            subheading.base_style = reference.styles["Heading 2"]
            reference.save(template)

            report = normalize_to_template(source, template, output, role_style_map={"heading2": subheading.style_id})
            paragraphs = Document(output).paragraphs
            self.assertEqual([p.style.name for p in paragraphs], [
                "Target Section", "Heading 3", "Normal", "Normal", "Heading 1", "Caption",
            ])
            self.assertEqual(report["roles"], {"heading2": 1, "heading3": 1, "body": 2, "heading1": 1, "caption": 1})
            self.assertTrue(all(p._p.pPr.find(qn("w:outlineLvl")) is None for p in paragraphs))

    def test_template_table_style_preserves_direct_paragraph_and_run_shading(self):
        with tempfile.TemporaryDirectory() as temp:
            src, ref, out = [Path(temp) / name for name in ("source.docx", "template.docx", "out.docx")]
            source = Document()
            table = source.add_table(rows=1, cols=1)
            cell = table.cell(0, 0)
            paragraph = cell.paragraphs[0]
            run = paragraph.add_run("Scientific highlight")
            for owner, color in ((cell._tc.get_or_add_tcPr(), "FF0000"),
                                 (paragraph._p.get_or_add_pPr(), "FFFF00"),
                                 (run._r.get_or_add_rPr(), "00FFFF")):
                shade = OxmlElement("w:shd")
                shade.set(qn("w:fill"), color)
                owner.append(shade)
            source.save(src)
            template = Document()
            style = template.styles.add_style("Target Table", WD_STYLE_TYPE.TABLE)
            template.save(ref)
            report = normalize_to_template(src, ref, out, table_style=style.style_id)
            root = xml(read_package(out)["word/document.xml"])
            self.assertEqual(report["table_style"]["look_policy"], "source_options")
            self.assertEqual(report["table_style"]["removed_overrides"], {"shd": 1})
            self.assertFalse(root.findall(".//w:tcPr/w:shd", NS))
            self.assertEqual([node.get(qn("w:fill")) for node in root.findall(".//w:shd", NS)], ["FFFF00", "00FFFF"])

    def test_rejects_adjacent_tables_that_word_would_merge(self):
        with tempfile.TemporaryDirectory() as temp:
            src, ref, out = [Path(temp) / name for name in ("source.docx", "template.docx", "out.docx")]
            source = Document()
            source.add_table(rows=1, cols=1)
            source.add_table(rows=1, cols=1)
            source.save(src)
            Document().save(ref)
            with self.assertRaisesRegex(ValueError, "Add a paragraph between adjacent tables"):
                normalize_to_template(src, ref, out, table_style="TableGrid")
            self.assertFalse(out.exists())

    def test_resets_typography_but_retains_scientific_emphasis_and_source(self):
        with tempfile.TemporaryDirectory() as temp:
            src,ref,out=[Path(temp)/x for x in ["source.docx","template.docx","out.docx"]]
            source=Document(); p=source.add_paragraph(style="Heading 1")
            r=p.add_run("Vector x"); r.bold=True; r.italic=True; r.font.size=Pt(40)
            source.save(src)
            template=Document(); template.styles["Heading 1"].font.size=Pt(14); template.save(ref)
            before=src.read_bytes()
            result=normalize_to_template(src,ref,out)
            self.assertEqual(src.read_bytes(),before)
            parts=read_package(out); root=xml(parts["word/document.xml"])
            self.assertTrue(root.findall(".//w:rPr/w:b",NS))
            self.assertTrue(root.findall(".//w:rPr/w:i",NS))
            self.assertFalse(root.findall(".//w:rPr/w:sz",NS))
            self.assertEqual(result["roles"]["heading1"],1)

    def test_multisection_template_requires_explicit_mapping(self):
        with tempfile.TemporaryDirectory() as temp:
            src,ref,out=[Path(temp)/x for x in ["source.docx","template.docx","out.docx"]]
            source=Document(); source.add_paragraph("body"); source.save(src)
            template=Document(); template.add_section(); template.save(ref)
            with self.assertRaisesRegex(ValueError,"section/slot mapping"):
                normalize_to_template(src,ref,out)
            self.assertFalse(out.exists())

    def test_explicit_second_section_bookmark_controls_imported_geometry(self):
        with tempfile.TemporaryDirectory() as temp:
            src, ref, out = [Path(temp) / name for name in ("source.docx", "template.docx", "out.docx")]
            source = Document()
            source.add_paragraph("Manuscript")
            source.save(src)
            template = Document()
            template.sections[0].page_width = Inches(8.5)
            template.add_paragraph("Cover")
            main = template.add_section()
            main.page_width = Inches(11)
            main.left_margin = Inches(1.75)
            main.header.is_linked_to_previous = False
            main.header.paragraphs[0].text = "Main header"
            paragraph = template.add_paragraph()
            start = OxmlElement("w:bookmarkStart")
            start.set(qn("w:name"), "manuscript")
            start.set(qn("w:id"), "17")
            end = OxmlElement("w:bookmarkEnd")
            end.set(qn("w:id"), "17")
            paragraph._p.extend((start, end))
            template.save(ref)

            report = normalize_to_template(src, ref, out, slot="bookmark:manuscript")
            self.assertEqual(report["template_location"]["section_index"], 2)
            self.assertEqual(report["template_location"]["section_count"], 2)
            self.assertTrue(report["template_location"]["header_references"])
            result = Document(out)
            self.assertEqual(result.sections[0].page_width, Inches(11))
            self.assertEqual(result.sections[0].left_margin, Inches(1.75))
            self.assertIsNone(result.sections[0]._sectPr.find(qn("w:headerReference")))

    def test_guide_table_policy_is_opt_in(self):
        with tempfile.TemporaryDirectory() as temp:
            src, ref, normal, guide = [Path(temp) / name for name in (
                "source.docx", "template.docx", "normal.docx", "guide.docx",
            )]
            source = Document()
            table = source.add_table(rows=2, cols=2)
            table.style = "Table Grid"
            properties = table.cell(1, 0)._tc.get_or_add_tcPr()
            shade = OxmlElement("w:shd")
            shade.set(qn("w:fill"), "AA0000")
            properties.append(shade)
            source.save(src)
            Document().save(ref)

            normalize_to_template(src, ref, normal)
            normalize_to_template(src, ref, guide, table_policy="three-line")
            normal_root = xml(read_package(normal)["word/document.xml"])
            self.assertEqual(normal_root.find(".//w:tblStyle", NS).get(qn("w:val")), "TableGrid")
            self.assertEqual(normal_root.find(".//w:shd", NS).get(qn("w:fill")), "AA0000")
            self.assertIsNone(normal_root.find(".//w:tblBorders", NS))
            guide_root = xml(read_package(guide)["word/document.xml"])
            self.assertIsNone(guide_root.find(".//w:shd", NS))
            self.assertEqual(guide_root.find(".//w:tblBorders/w:insideV", NS).get(qn("w:val")), "nil")

    def test_source_section_furniture_is_replaced_by_the_slot_but_page_intent_survives(self):
        with tempfile.TemporaryDirectory() as temp:
            src, ref, out = [Path(temp) / name for name in ("source.docx", "template.docx", "out.docx")]
            source = Document()
            source.add_paragraph("First source section")
            source.sections[0].header.paragraphs[0].text = "Source header"
            source.add_section(WD_SECTION.CONTINUOUS)
            source.paragraphs[-1].add_run("Text in the boundary paragraph")
            source.add_paragraph("Continuous source section")
            final_section = source.add_section(WD_SECTION.NEW_PAGE)
            final_section.different_first_page_header_footer = True
            restart = OxmlElement("w:pgNumType")
            restart.set(qn("w:start"), "19")
            final_section._sectPr.insert_element_before(restart, "w:cols", "w:docGrid")
            source.add_paragraph("Next-page source section")
            source.save(src)
            original = src.read_bytes()

            template = Document()
            template.sections[0].page_width = Inches(11)
            template.add_paragraph("{{PAPERFLOW_BODY}}")
            template.save(ref)
            report = normalize_to_template(src, ref, out, slot="{{PAPERFLOW_BODY}}")

            self.assertEqual(src.read_bytes(), original)
            normalized = Document(out)
            self.assertEqual(len(normalized.sections), 1)
            self.assertEqual(normalized.sections[0].page_width, Inches(11))
            self.assertEqual([paragraph.text for paragraph in normalized.paragraphs if paragraph.text], [
                "First source section", "Text in the boundary paragraph", "Continuous source section", "Next-page source section",
            ])
            self.assertIsNone(normalized.paragraphs[2].paragraph_format.page_break_before)
            self.assertTrue(normalized.paragraphs[-1].paragraph_format.page_break_before)
            terminal = normalized.sections[0]._sectPr
            for name in ("w:headerReference", "w:footerReference", "w:pgNumType", "w:titlePg", "w:type"):
                self.assertIsNone(terminal.find(qn(name)))
            self.assertEqual(report["source_section_breaks"], [
                {"after_source_section": 1, "source_start_type": "continuous", "conversion": "no_added_break"},
                {"after_source_section": 2, "source_start_type": "nextPage", "conversion": "page_break_before"},
            ])

    def test_odd_even_and_column_breaks_require_review_before_single_slot_import(self):
        with tempfile.TemporaryDirectory() as temp:
            src, ref, out = [Path(temp) / name for name in ("source.docx", "template.docx", "out.docx")]
            template = Document()
            template.add_paragraph("{{PAPERFLOW_BODY}}")
            template.save(ref)
            for start, name in ((WD_SECTION.ODD_PAGE, "oddPage"), (WD_SECTION.EVEN_PAGE, "evenPage"),
                                (WD_SECTION.NEW_COLUMN, "nextColumn")):
                with self.subTest(mode=name):
                    source = Document()
                    source.add_paragraph("First source section")
                    source.add_section(start)
                    source.add_paragraph("Following source section")
                    source.save(src)
                    with self.assertRaisesRegex(ValueError, name):
                        normalize_to_template(src, ref, out, slot="{{PAPERFLOW_BODY}}")
                    self.assertFalse(out.exists())


if __name__=="__main__": unittest.main()
