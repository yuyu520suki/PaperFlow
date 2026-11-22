import tempfile
import unittest
from pathlib import Path

from docx import Document
from lxml import etree as ET

from paperflow.formatting import normalize_to_template
from paperflow.package import NS, q, read_package, write_package, xml


def _direct_number(paragraph, number_id):
    properties = paragraph._p.get_or_add_pPr().get_or_add_numPr()
    properties.get_or_add_ilvl().val = 0
    properties.get_or_add_numId().val = number_id


def _numbered_template(path, *, picture=False):
    template = Document()
    numbering = template.part.numbering_part.element
    concrete = next(node for node in numbering.findall(q("w:num")) if node.get(q("w:numId")) == "1")
    abstract_id = concrete.find(q("w:abstractNumId")).get(q("w:val"))
    abstract = next(node for node in numbering.findall(q("w:abstractNum"))
                    if node.get(q("w:abstractNumId")) == abstract_id)
    # A Word-saved document can use numeric style IDs even for built-in styles.
    heading = template.styles["Heading 1"].element
    heading.set(q("w:styleId"), "104")
    heading.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = 1
    for level in abstract.findall(q("w:lvl")):
        abstract.remove(level)
    for index, style_id in enumerate(("104", "Heading2")):
        level = ET.SubElement(abstract, q("w:lvl"), {q("w:ilvl"): str(index)})
        for property_name, value in (("start", "1"), ("numFmt", "decimal"), ("pStyle", style_id),
                                     ("lvlText", "%1" if index == 0 else "%1.%2")):
            ET.SubElement(level, q("w:" + property_name), {q("w:val"): value})
        if picture:
            ET.SubElement(level, q("w:lvlPicBulletId"), {q("w:val"): "1"})
    template.styles["Heading 2"].element.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = 1
    template.save(path)


def _num_id(paragraph):
    node = paragraph.find("w:pPr/w:numPr/w:numId", NS)
    return node.get(q("w:val")) if node is not None else None


def _style(part, style_id):
    return next(node for node in xml(part).findall(q("w:style")) if node.get(q("w:styleId")) == style_id)


class HeadingNumberingTests(unittest.TestCase):
    def test_target_heading_list_avoids_collision_and_keeps_source_lists(self):
        with tempfile.TemporaryDirectory() as directory:
            source, template, output = (Path(directory) / name for name in ("source.docx", "template.docx", "out.docx"))
            draft = Document()
            _direct_number(draft.add_heading("Chapter", 1), 1)
            draft.add_heading("Section", 2)
            _direct_number(draft.add_paragraph("Direct list"), 1)
            draft.add_paragraph("Style list", style="List Number")
            draft.save(source)
            _numbered_template(template)
            original_source, original_template = source.read_bytes(), template.read_bytes()
            original_parts = read_package(source)

            report = normalize_to_template(source, template, output)

            self.assertEqual(report["numbered_heading_levels"], [1, 2])
            self.assertEqual(report["role_style_ids"]["heading1"], "104")
            self.assertEqual(report["unresolved"], [])
            parts = read_package(output)
            body = xml(parts["word/document.xml"]).findall("w:body/w:p", NS)
            self.assertIsNone(_num_id(body[0]))
            self.assertIsNone(_num_id(body[1]))
            self.assertEqual(_num_id(body[2]), "1")
            self.assertEqual(_num_id(body[3]), _num_id(_style(original_parts["word/styles.xml"], "ListNumber")))

            style1 = _style(parts["word/styles.xml"], "104")
            style2 = _style(parts["word/styles.xml"], "Heading2")
            new_id = _num_id(style1)
            self.assertNotEqual(new_id, "1")
            self.assertEqual(_num_id(style2), new_id)
            numbering = xml(parts["word/numbering.xml"])
            concrete = numbering.find(f"w:num[@w:numId='{new_id}']", NS)
            abstract_id = concrete.find(q("w:abstractNumId")).get(q("w:val"))
            abstract = numbering.find(f"w:abstractNum[@w:abstractNumId='{abstract_id}']", NS)
            self.assertEqual([node.get(q("w:val")) for node in abstract.findall("w:lvl/w:lvlText", NS)], ["%1", "%1.%2"])
            for tag, identifier in (("num", "numId"), ("abstractNum", "abstractNumId")):
                for node in xml(original_parts["word/numbering.xml"]).findall(q(f"w:{tag}")):
                    destination = numbering.find(f"w:{tag}[@w:{identifier}='{node.get(q(f'w:{identifier}'))}']", NS)
                    self.assertEqual(ET.tostring(destination, method="c14n", exclusive=True),
                                     ET.tostring(node, method="c14n", exclusive=True))
            self.assertEqual(source.read_bytes(), original_source)
            self.assertEqual(template.read_bytes(), original_template)

    def test_missing_source_numbering_part_is_created_with_a_relationship(self):
        with tempfile.TemporaryDirectory() as directory:
            base, source, template, output = (Path(directory) / name for name in
                                              ("base.docx", "source.docx", "template.docx", "out.docx"))
            draft = Document()
            draft.add_heading("Chapter", 1)
            draft.save(base)
            parts = read_package(base)
            del parts["word/numbering.xml"]
            for name, attribute, value in (
                ("word/_rels/document.xml.rels", "Target", "numbering.xml"),
                ("[Content_Types].xml", "PartName", "/word/numbering.xml"),
            ):
                root = xml(parts[name])
                root.remove(next(node for node in root if node.get(attribute) == value))
                parts[name] = ET.tostring(root)
            write_package(parts, source)
            _numbered_template(template)

            report = normalize_to_template(source, template, output)

            self.assertEqual(report["numbered_heading_levels"], [1, 2])
            result = Document(output)
            self.assertEqual(result.paragraphs[0].style.style_id, "104")
            self.assertEqual(result.part.numbering_part.element.find(q("w:num")).get(q("w:numId")), "1")

    def test_picture_linked_numbering_requires_review_and_keeps_direct_source_number(self):
        with tempfile.TemporaryDirectory() as directory:
            source, template, output = (Path(directory) / name for name in ("source.docx", "template.docx", "out.docx"))
            draft = Document()
            _direct_number(draft.add_heading("Chapter", 1), 1)
            draft.save(source)
            _numbered_template(template, picture=True)

            report = normalize_to_template(source, template, output)

            self.assertEqual(report["numbered_heading_levels"], [])
            self.assertTrue(all("picture numbering" in issue for issue in report["unresolved"]))
            self.assertEqual(len(report["unresolved"]), 2)
            parts = read_package(output)
            self.assertEqual(parts["word/numbering.xml"], read_package(source)["word/numbering.xml"])
            self.assertEqual(_num_id(_style(parts["word/styles.xml"], "104")), "0")
            self.assertEqual(_num_id(xml(parts["word/document.xml"]).find("w:body/w:p", NS)), "1")


if __name__ == "__main__":
    unittest.main()
