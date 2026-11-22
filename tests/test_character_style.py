import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.shared import Pt

from paperflow.formatting import normalize_to_template
from paperflow.package import NS, q, read_package, xml


def _style_chain(properties, styles):
    reference = properties.find(q("w:rStyle"))
    chain = []
    while reference is not None:
        style_id = reference.get(q("w:val"))
        if style_id not in styles or styles[style_id] in chain:
            raise AssertionError(f"Unresolved character style chain at {style_id}")
        style = styles[style_id]
        chain.append(style)
        reference = style.find(q("w:basedOn"))
    return chain


class CharacterStyleTests(unittest.TestCase):
    def test_normalizes_imported_character_typography_without_losing_style_semantics(self):
        with tempfile.TemporaryDirectory() as folder:
            source, template, output = (Path(folder) / name for name in ("source.docx", "template.docx", "out.docx"))
            manuscript = Document()
            typography = manuscript.styles.add_style("Only typography", WD_STYLE_TYPE.CHARACTER)
            typography.font.name = "Courier New"
            typography.font.size = Pt(36)

            base = manuscript.styles.add_style("Variable base", WD_STYLE_TYPE.CHARACTER)
            base.font.name = "Courier New"
            base.font.size = Pt(34)
            base.font.bold = True
            notation = manuscript.styles.add_style("Shared notation", WD_STYLE_TYPE.CHARACTER)
            notation.base_style = base
            notation.font.italic = True
            notation.font.subscript = True
            double_toggle = manuscript.styles.add_style("Double bold toggle", WD_STYLE_TYPE.CHARACTER)
            double_toggle.base_style = base
            double_toggle.font.bold = True

            paragraph = manuscript.add_paragraph()
            paragraph.add_run("Typography alone", style=typography)
            paragraph.add_run("Scientific notation", style=notation)
            override = paragraph.add_run("Explicitly unbold", style=notation)
            override.font.bold = False
            paragraph.add_run("Two style toggles", style=double_toggle)
            cell = manuscript.add_table(1, 1).cell(0, 0)
            cell.paragraphs[0].add_run("Preserved table text", style=typography)

            math = OxmlElement("m:oMath")
            math_run = OxmlElement("m:r")
            math_properties = OxmlElement("w:rPr")
            math_size = OxmlElement("w:sz")
            math_size.set(q("w:val"), "42")
            math_properties.append(math_size)
            math_text = OxmlElement("m:t")
            math_text.text = "x+1"
            math_run.extend((math_properties, math_text))
            math.append(math_run)
            manuscript.add_paragraph()._p.append(math)
            manuscript.save(source)

            reference = Document()
            reference.styles["Normal"].font.name = "Times New Roman"
            reference.styles["Normal"].font.size = Pt(11)
            collision = reference.styles.add_style("Shared notation", WD_STYLE_TYPE.CHARACTER)
            collision._element.set(q("w:styleId"), "TemplateNotation")
            collision.font.size = Pt(27)
            collision.font.superscript = True
            reference.save(template)
            source_bytes, template_bytes = source.read_bytes(), template.read_bytes()

            normalize_to_template(source, template, output)
            parts = read_package(output)
            styles = {style.get(q("w:styleId")): style for style in xml(parts["word/styles.xml"]).findall(q("w:style"))}
            document = xml(parts["word/document.xml"])
            runs = {run.find(q("w:t")).text: run.find(q("w:rPr")) for run in document.iter(q("w:r"))
                    if run.find(q("w:t")) is not None}

            typography_run = runs["Typography alone"]
            self.assertTrue(typography_run is None or typography_run.find(q("w:rStyle")) is None)
            notation_chain = _style_chain(runs["Scientific notation"], styles)
            self.assertTrue(notation_chain)
            self.assertNotIn("TemplateNotation", [style.get(q("w:styleId")) for style in notation_chain])
            self.assertNotEqual(notation_chain[0].find(q("w:name")).get(q("w:val")).casefold(), "shared notation")
            self.assertFalse(any(style.findall("w:rPr/w:rFonts", NS) or style.findall("w:rPr/w:sz", NS)
                                 for style in notation_chain))
            self.assertTrue(any(style.find("w:rPr/w:b", NS) is not None for style in notation_chain))
            self.assertTrue(any(style.find("w:rPr/w:i", NS) is not None for style in notation_chain))
            self.assertEqual([align.get(q("w:val")) for style in notation_chain
                              for align in style.findall("w:rPr/w:vertAlign", NS)], ["subscript"])
            self.assertEqual(_style_chain(runs["Explicitly unbold"], styles), notation_chain)
            self.assertEqual(runs["Explicitly unbold"].find(q("w:b")).get(q("w:val")), "0")

            double_properties = runs["Two style toggles"]
            self.assertTrue(double_properties is None or double_properties.find(q("w:b")) is None)
            if double_properties is not None:
                toggles = [bold for style in _style_chain(double_properties, styles)
                           for bold in style.findall("w:rPr/w:b", NS) if bold.get(q("w:val")) not in {"0", "off", "false"}]
                self.assertEqual(len(toggles) % 2, 0)

            table_chain = _style_chain(runs["Preserved table text"], styles)
            self.assertEqual(table_chain[0].get(q("w:styleId")), typography.style_id)
            self.assertEqual(table_chain[0].find("w:rPr/w:sz", NS).get(q("w:val")), "72")
            self.assertEqual(document.find(".//m:oMath/m:r/m:t", NS).text, "x+1")
            self.assertEqual(document.find(".//m:oMath/m:r/w:rPr/w:sz", NS).get(q("w:val")), "42")
            self.assertEqual(source.read_bytes(), source_bytes)
            self.assertEqual(template.read_bytes(), template_bytes)


if __name__ == "__main__":
    unittest.main()
