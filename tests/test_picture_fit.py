"""Picture fitting keeps the image payload while adapting to destination width."""
import base64
import io
import unittest
from copy import deepcopy

from docx import Document
from docx.shared import Inches
from lxml import etree as ET

from paperflow.package import NS, q
from paperflow.picture_fit import fit_inline_pictures


NS = {
    **NS,
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "v": "urn:schemas-microsoft-com:vml",
}
EMU_PER_TWIP = 635
BODY_WIDTH = (12240 - 1440 - 1440) * EMU_PER_TWIP
BODY_HEIGHT = (15840 - 1440 - 1440) * EMU_PER_TWIP


def tag(name):
    prefix, local = name.split(":")
    return "{" + NS[prefix] + "}" + local


def section(*, gutter=0):
    root = ET.Element(q("w:sectPr"), nsmap={"w": NS["w"]})
    ET.SubElement(root, q("w:pgSz"), {q("w:w"): "12240", q("w:h"): "15840"})
    ET.SubElement(root, q("w:pgMar"), {
        q("w:top"): "1440", q("w:bottom"): "1440",
        q("w:left"): "1440", q("w:right"): "1440", q("w:gutter"): str(gutter),
    })
    return root


def styles():
    root = ET.Element(q("w:styles"), nsmap={"w": NS["w"]})
    style = ET.SubElement(root, q("w:style"), {
        q("w:type"): "paragraph", q("w:default"): "1", q("w:styleId"): "Normal",
    })
    ET.SubElement(style, q("w:name"), {q("w:val"): "Normal"})
    return root


def picture(width=BODY_WIDTH * 2, height=BODY_WIDTH // 2, *, inner=None):
    inner_width, inner_height = inner or (width, height)
    return ET.fromstring(f"""
    <w:drawing xmlns:w="{NS['w']}" xmlns:wp="{NS['wp']}" xmlns:a="{NS['a']}"
      xmlns:pic="{NS['pic']}" xmlns:r="{NS['r']}">
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="{width}" cy="{height}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:docPr id="42" name="Picture 42" descr="Original scientific plot"/>
        <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
        <a:graphic><a:graphicData uri="{NS['pic']}"><pic:pic>
          <pic:nvPicPr><pic:cNvPr id="0" name="response.png"/><pic:cNvPicPr/></pic:nvPicPr>
          <pic:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{inner_width}" cy="{inner_height}"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </pic:spPr>
        </pic:pic></a:graphicData></a:graphic>
      </wp:inline>
    </w:drawing>
    """.encode())


def document(drawing=None):
    root = ET.Element(q("w:document"), nsmap=NS)
    body = ET.SubElement(root, q("w:body"))
    paragraph = ET.SubElement(body, q("w:p"))
    if drawing is not None:
        ET.SubElement(paragraph, q("w:r")).append(drawing)
    body.append(section())
    return root


def dimensions(root, path=".//wp:inline/wp:extent"):
    node = root.find(path, NS)
    return [int(node.get(axis)) for axis in ("cx", "cy")]


def add_indent(parent, **values):
    properties = parent.find(q("w:pPr"))
    if properties is None:
        properties = ET.SubElement(parent, q("w:pPr"))
    return ET.SubElement(properties, q("w:ind"), {q("w:" + key): str(value) for key, value in values.items()})


class InlinePictureFitTests(unittest.TestCase):
    def test_fits_real_python_docx_inline_shape_and_reports_original_dimensions(self):
        # A generated DOCX picture exercises the same namespaces/shape structure
        # as a normal Word insert, without opening Word or modifying any file.
        png = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jY9kAAAAASUVORK5CYII="
        )
        source = Document()
        source.add_paragraph().add_run().add_picture(io.BytesIO(png), width=Inches(13), height=Inches(2))
        root = ET.fromstring(ET.tostring(source._element))
        target, target_styles = section(), styles()
        before_geometry, before_styles = ET.tostring(target), ET.tostring(target_styles)

        report = fit_inline_pictures(root, target, target_styles)

        self.assertEqual(dimensions(root), [int(Inches(6.5)), int(Inches(1))])
        self.assertEqual(report["resized"], 1)
        self.assertEqual(report["unresolved"], [])
        self.assertEqual(len(report["pictures"]), 1)
        record = report["pictures"][0]
        self.assertEqual(record["status"], "resized")
        self.assertEqual(record["before_emu"], [int(Inches(13)), int(Inches(2))])
        self.assertEqual(record["after_emu"], dimensions(root))
        self.assertEqual(ET.tostring(target), before_geometry)
        self.assertEqual(ET.tostring(target_styles), before_styles)
        self.assertEqual(source.inline_shapes[0].width, Inches(13))

    def test_scales_both_extents_by_same_ratio_and_preserves_picture_payload(self):
        drawing = picture(inner=(BODY_WIDTH * 3 // 2, BODY_WIDTH // 4))
        fill = drawing.find(".//pic:blipFill", NS)
        fill.insert(1, ET.Element(tag("a:srcRect"), l="2500", r="7500", t="1000", b="0"))
        transform = drawing.find(".//a:xfrm", NS)
        transform.set("rot", "600000")
        transform.set("flipH", "1")
        transform.set("flipV", "0")
        root = document(drawing)
        before = deepcopy(root)

        report = fit_inline_pictures(root, section(), styles())

        self.assertEqual(dimensions(root), [BODY_WIDTH, BODY_WIDTH // 4])
        self.assertEqual(dimensions(root, ".//pic:spPr/a:xfrm/a:ext"), [BODY_WIDTH * 3 // 4, BODY_WIDTH // 8])
        restored = deepcopy(root)
        for path in (".//wp:inline/wp:extent", ".//pic:spPr/a:xfrm/a:ext"):
            original = before.find(path, NS)
            resized = restored.find(path, NS)
            for axis in ("cx", "cy"):
                resized.set(axis, original.get(axis))
        self.assertEqual(ET.tostring(restored), ET.tostring(before))
        self.assertEqual(report["resized"], 1)

    def test_deducts_gutter_and_effect_margins_without_changing_effects(self):
        root = document(picture())
        effect = root.find(".//wp:effectExtent", NS)
        effect.attrib.update({"l": "19050", "r": "31750", "t": "12700", "b": "25400"})
        original_effect = dict(effect.attrib)
        available = BODY_WIDTH - 240 * EMU_PER_TWIP - 19050 - 31750

        report = fit_inline_pictures(root, section(gutter=240), styles())

        self.assertLessEqual(dimensions(root)[0], available)
        self.assertLessEqual(abs(dimensions(root)[0] - available), 1)
        self.assertEqual(dict(effect.attrib), original_effect)
        self.assertEqual(report["unresolved"], [])

    def test_never_enlarges_and_repeated_fitting_is_idempotent(self):
        for width in (BODY_WIDTH // 3, BODY_WIDTH, BODY_WIDTH * 2 + 1):
            with self.subTest(width=width):
                root = document(picture(width=width, height=BODY_WIDTH // 4))
                first = fit_inline_pictures(root, section(), styles())
                after_first = ET.tostring(root)
                second = fit_inline_pictures(root, section(), styles())
                self.assertLessEqual(dimensions(root)[0], min(width, BODY_WIDTH))
                self.assertEqual(ET.tostring(root), after_first)
                self.assertEqual(second["resized"], 0)
                self.assertEqual(second["pictures"][0]["status"], "unchanged")
                self.assertEqual(second["unresolved"], [])
                self.assertEqual(first["resized"], int(width > BODY_WIDTH))

    def test_equal_columns_deduct_spacing_and_unequal_columns_use_smallest_width(self):
        equal = section()
        ET.SubElement(equal, q("w:cols"), {q("w:num"): "2", q("w:space"): "720", q("w:equalWidth"): "1"})
        unequal = section()
        columns = ET.SubElement(unequal, q("w:cols"), {q("w:equalWidth"): "0"})
        ET.SubElement(columns, q("w:col"), {q("w:w"): "3000", q("w:space"): "720"})
        ET.SubElement(columns, q("w:col"), {q("w:w"): "5640"})
        for target, expected in ((equal, 4320 * EMU_PER_TWIP), (unequal, 3000 * EMU_PER_TWIP)):
            with self.subTest(expected=expected):
                root = document(picture())
                report = fit_inline_pictures(root, target, styles())
                self.assertEqual(dimensions(root)[0], expected)
                self.assertEqual(report["resized"], 1)
                self.assertEqual(report["unresolved"], [])

    def test_resolves_inherited_paragraph_indents_with_direct_start_end_overrides(self):
        target_styles = styles()
        normal = target_styles.find(q("w:style"))
        add_indent(normal, left=300, right=200)
        parent = ET.SubElement(target_styles, q("w:style"), {q("w:type"): "paragraph", q("w:styleId"): "Parent"})
        ET.SubElement(parent, q("w:basedOn"), {q("w:val"): "Normal"})
        add_indent(parent, left=500)
        child = ET.SubElement(target_styles, q("w:style"), {q("w:type"): "paragraph", q("w:styleId"): "TargetFigure"})
        ET.SubElement(child, q("w:basedOn"), {q("w:val"): "Parent"})
        add_indent(child, right=400)
        for direct, deduction in ((None, 900), ({"left": 900, "right": 800, "start": 700, "end": 600}, 1300)):
            with self.subTest(direct=direct):
                root = document(picture())
                paragraph = root.find("w:body/w:p", NS)
                properties = ET.Element(q("w:pPr"))
                ET.SubElement(properties, q("w:pStyle"), {q("w:val"): "TargetFigure"})
                paragraph.insert(0, properties)
                if direct:
                    add_indent(paragraph, **direct)
                report = fit_inline_pictures(root, section(), target_styles)
                self.assertEqual(dimensions(root)[0], BODY_WIDTH - deduction * EMU_PER_TWIP)
                self.assertEqual(report["unresolved"], [])

    def test_negative_indents_do_not_enlarge_the_printable_width(self):
        root = document(picture())
        add_indent(root.find("w:body/w:p", NS), left=-300, right=-400)
        fit_inline_pictures(root, section(), styles())
        self.assertEqual(dimensions(root)[0], BODY_WIDTH)

    def test_default_style_and_first_line_indent_constrain_an_unstyled_paragraph(self):
        target_styles = styles()
        add_indent(target_styles.find(q("w:style")), left=400, right=100, firstLine=200)
        root = document(picture())
        report = fit_inline_pictures(root, section(), target_styles)
        self.assertEqual(dimensions(root)[0], BODY_WIDTH - 700 * EMU_PER_TWIP)
        self.assertEqual(report["unresolved"], [])

    def test_top_gutter_changes_height_review_without_reducing_picture_width(self):
        settings = ET.Element(q("w:settings"))
        ET.SubElement(settings, q("w:gutterAtTop"))
        root = document(picture(width=BODY_WIDTH, height=BODY_HEIGHT - 100 * EMU_PER_TWIP))
        before = ET.tostring(root)
        report = fit_inline_pictures(root, section(gutter=200), styles(), settings_root=settings)
        self.assertEqual(ET.tostring(root), before)
        self.assertEqual(report["resized"], 0)
        self.assertTrue(report["unresolved"])

    def test_numbered_paragraph_is_retained_for_native_indent_review(self):
        for inherited in (False, True):
            with self.subTest(inherited=inherited):
                root, target_styles = document(picture()), styles()
                owner = target_styles.find(q("w:style")) if inherited else root.find("w:body/w:p", NS)
                properties = ET.SubElement(owner, q("w:pPr"))
                numbering = ET.SubElement(properties, q("w:numPr"))
                ET.SubElement(numbering, q("w:ilvl"), {q("w:val"): "0"})
                ET.SubElement(numbering, q("w:numId"), {q("w:val"): "7"})
                before = ET.tostring(root)
                report = fit_inline_pictures(root, section(), target_styles)
                self.assertEqual(ET.tostring(root), before)
                self.assertEqual(report["resized"], 0)
                self.assertEqual(report["pictures"][0]["status"], "unresolved")
                self.assertTrue(report["unresolved"])

    def test_direct_numid_zero_cancels_inherited_numbering_and_allows_width_fit(self):
        root, target_styles = document(picture()), styles()
        for owner, number in ((target_styles.find(q("w:style")), "7"),
                              (root.find("w:body/w:p", NS), "0")):
            properties = ET.SubElement(owner, q("w:pPr"))
            numbering = ET.SubElement(properties, q("w:numPr"))
            ET.SubElement(numbering, q("w:numId"), {q("w:val"): number})
        report = fit_inline_pictures(root, section(), target_styles)
        self.assertEqual(dimensions(root)[0], BODY_WIDTH)
        self.assertEqual(report["resized"], 1)
        self.assertEqual(report["unresolved"], [])

    def test_inherited_or_document_default_paragraph_frame_is_retained_for_review(self):
        for location in ("style", "docDefaults"):
            with self.subTest(location=location):
                root, target_styles = document(picture()), styles()
                if location == "style":
                    owner = target_styles.find(q("w:style"))
                else:
                    defaults = ET.SubElement(target_styles, q("w:docDefaults"))
                    owner = ET.SubElement(defaults, q("w:pPrDefault"))
                properties = ET.SubElement(owner, q("w:pPr"))
                ET.SubElement(properties, q("w:framePr"), {q("w:w"): "3000", q("w:h"): "2000"})
                before = ET.tostring(root)
                report = fit_inline_pictures(root, section(), target_styles)
                self.assertEqual(ET.tostring(root), before)
                self.assertEqual(report["resized"], 0)
                self.assertEqual(report["pictures"][0]["status"], "unresolved")
                self.assertTrue(report["unresolved"])

    def test_character_based_indents_are_retained_and_reported_for_review(self):
        for attribute in ("startChars", "endChars", "leftChars", "rightChars"):
            with self.subTest(attribute=attribute):
                root = document(picture())
                add_indent(root.find("w:body/w:p", NS), **{attribute: 100})
                before = ET.tostring(root)
                report = fit_inline_pictures(root, section(), styles())
                self.assertEqual(ET.tostring(root), before)
                self.assertEqual(report["pictures"][0]["status"], "unresolved")
                self.assertEqual(report["resized"], 0)
                self.assertTrue(report["unresolved"])

    def test_table_cell_picture_is_reported_without_changing_image_or_table_widths(self):
        root = document()
        body = root.find(q("w:body"))
        body.remove(body[0])
        table = ET.Element(q("w:tbl"))
        properties = ET.SubElement(table, q("w:tblPr"))
        ET.SubElement(properties, q("w:tblW"), {q("w:type"): "dxa", q("w:w"): "6000"})
        ET.SubElement(properties, q("w:tblLayout"), {q("w:type"): "fixed"})
        grid = ET.SubElement(table, q("w:tblGrid"))
        for _ in range(2):
            ET.SubElement(grid, q("w:gridCol"), {q("w:w"): "3000"})
        cell = ET.SubElement(ET.SubElement(table, q("w:tr")), q("w:tc"))
        cell_properties = ET.SubElement(cell, q("w:tcPr"))
        ET.SubElement(cell_properties, q("w:tcW"), {q("w:type"): "dxa", q("w:w"): "6000"})
        ET.SubElement(cell_properties, q("w:gridSpan"), {q("w:val"): "2"})
        ET.SubElement(ET.SubElement(cell, q("w:p")), q("w:r")).append(picture())
        body.insert(0, table)
        before = ET.tostring(root)

        report = fit_inline_pictures(root, section(), styles())

        self.assertEqual(ET.tostring(root), before)
        self.assertEqual(report["resized"], 0)
        self.assertEqual(report["pictures"][0]["status"], "unresolved")
        self.assertTrue(report["unresolved"])

    def test_tall_picture_is_reported_without_automatic_height_shrink(self):
        root = document(picture(width=BODY_WIDTH // 2, height=BODY_HEIGHT + 1))
        before = ET.tostring(root)
        report = fit_inline_pictures(root, section(), styles())
        self.assertEqual(ET.tostring(root), before)
        self.assertEqual(report["resized"], 0)
        self.assertEqual(report["pictures"][0]["status"], "unchanged")
        self.assertTrue(report["unresolved"])

    def test_reports_height_still_oversize_after_width_fit_and_counts_effect_height(self):
        for width, height, effect_bottom in (
            (BODY_WIDTH * 2, BODY_HEIGHT * 3, 0),
            (BODY_WIDTH // 2, BODY_HEIGHT - 10, 11),
        ):
            with self.subTest(width=width):
                root = document(picture(width=width, height=height))
                root.find(".//wp:effectExtent", NS).set("b", str(effect_bottom))
                report = fit_inline_pictures(root, section(), styles())
                if width > BODY_WIDTH:
                    self.assertEqual(dimensions(root), [BODY_WIDTH, height // 2])
                    self.assertEqual(report["resized"], 1)
                else:
                    self.assertEqual(dimensions(root), [width, height])
                    self.assertEqual(report["resized"], 0)
                self.assertTrue(report["unresolved"])

    def test_malformed_external_or_inner_extents_are_not_partially_changed(self):
        def invalid_outer(root):
            root.find(".//wp:extent", NS).set("cx", "not-an-integer")

        def zero_outer(root):
            root.find(".//wp:extent", NS).set("cy", "0")

        def missing_inner(root):
            node = root.find(".//pic:spPr/a:xfrm/a:ext", NS)
            node.getparent().remove(node)

        def invalid_inner(root):
            root.find(".//pic:spPr/a:xfrm/a:ext", NS).set("cx", "-1")

        for mutate in (invalid_outer, zero_outer, missing_inner, invalid_inner):
            with self.subTest(case=mutate.__name__):
                root = document(picture())
                mutate(root)
                before = ET.tostring(root)
                report = fit_inline_pictures(root, section(), styles())
                self.assertEqual(ET.tostring(root), before)
                self.assertEqual(report["resized"], 0)
                self.assertEqual(report["pictures"][0]["status"], "unresolved")
                self.assertTrue(report["unresolved"])

    def test_missing_destination_geometry_is_unresolved_only_when_needed(self):
        for target in (None, ET.Element(q("w:sectPr"))):
            with self.subTest(target=target):
                root = document(picture())
                before = ET.tostring(root)
                report = fit_inline_pictures(root, target, styles())
                self.assertEqual(ET.tostring(root), before)
                self.assertEqual(report["resized"], 0)
                self.assertEqual(report["pictures"][0]["status"], "unresolved")
                self.assertTrue(report["unresolved"])
                empty_report = fit_inline_pictures(document(), target, styles())
                self.assertEqual(empty_report["pictures"], [])
                self.assertEqual(empty_report["unresolved"], [])

    def test_skips_floating_linked_ole_deleted_textbox_and_alternate_content(self):
        def floating(root):
            root.find(".//wp:inline", NS).tag = tag("wp:anchor")

        def linked(root):
            root.find(".//a:blip", NS).set(q("r:link"), "rIdExternal")

        def wrap_run(root, name):
            paragraph = root.find("w:body/w:p", NS)
            run = paragraph[0]
            paragraph.remove(run)
            wrapper = ET.SubElement(paragraph, tag(name))
            wrapper.append(run)
            return wrapper

        def ole(root):
            drawing = root.find(".//w:drawing", NS)
            run = drawing.getparent()
            run.remove(drawing)
            obj = ET.SubElement(run, q("w:object"))
            obj.append(drawing)
            ET.SubElement(obj, q("o:OLEObject"), {"ProgID": "Equation.DSMT4", q("r:id"): "rIdEquation"})

        def deleted(root):
            wrap_run(root, "w:del").set(q("w:id"), "7")

        def moved_from(root):
            wrap_run(root, "w:moveFrom").set(q("w:id"), "8")

        def textbox(root):
            paragraph = root.find("w:body/w:p", NS)
            run = paragraph[0]
            paragraph.remove(run)
            outer = ET.SubElement(paragraph, q("w:r"))
            pict = ET.SubElement(outer, q("w:pict"))
            shape = ET.SubElement(pict, tag("v:shape"), id="TextBox1")
            content = ET.SubElement(ET.SubElement(shape, tag("v:textbox")), q("w:txbxContent"))
            ET.SubElement(content, q("w:p")).append(run)

        def alternate(root):
            drawing = root.find(".//w:drawing", NS)
            run = drawing.getparent()
            run.remove(drawing)
            wrapper = ET.SubElement(run, tag("mc:AlternateContent"))
            ET.SubElement(wrapper, tag("mc:Choice"), Requires="wps").append(drawing)
            ET.SubElement(wrapper, tag("mc:Fallback"))

        for mutate in (floating, linked, ole, deleted, moved_from, textbox, alternate):
            with self.subTest(case=mutate.__name__):
                root = document(picture())
                mutate(root)
                before = ET.tostring(root)
                report = fit_inline_pictures(root, section(), styles())
                self.assertEqual(ET.tostring(root), before)
                self.assertEqual(report["resized"], 0)
                self.assertEqual(report["pictures"][0]["status"], "skipped")
                self.assertTrue(report["pictures"][0].get("reason"))

    def test_skips_simple_and_active_complex_field_but_fits_picture_after_field_end(self):
        root = document()
        paragraph = root.find("w:body/w:p", NS)
        simple = ET.SubElement(paragraph, q("w:fldSimple"), {q("w:instr"): "INCLUDEPICTURE \"plot.png\""})
        ET.SubElement(simple, q("w:r")).append(picture())
        simple_before = ET.tostring(simple)
        begin = ET.SubElement(paragraph, q("w:r"))
        ET.SubElement(begin, q("w:fldChar"), {q("w:fldCharType"): "begin"})
        instruction = ET.SubElement(ET.SubElement(paragraph, q("w:r")), q("w:instrText"))
        instruction.text = ' INCLUDEPICTURE "response.png" '
        ET.SubElement(ET.SubElement(paragraph, q("w:r")), q("w:fldChar"), {q("w:fldCharType"): "separate"})
        field_picture = picture()
        ET.SubElement(paragraph, q("w:r")).append(field_picture)
        field_before = ET.tostring(field_picture)
        ET.SubElement(ET.SubElement(paragraph, q("w:r")), q("w:fldChar"), {q("w:fldCharType"): "end"})
        ordinary = picture()
        ET.SubElement(paragraph, q("w:r")).append(ordinary)

        report = fit_inline_pictures(root, section(), styles())

        self.assertEqual(ET.tostring(simple), simple_before)
        self.assertEqual(ET.tostring(field_picture), field_before)
        self.assertEqual([int(ordinary.find("wp:inline/wp:extent", NS).get(axis)) for axis in ("cx", "cy")],
                         [BODY_WIDTH, BODY_WIDTH // 4])
        self.assertEqual([record["status"] for record in report["pictures"]], ["skipped", "skipped", "resized"])
        self.assertEqual(report["resized"], 1)


if __name__ == "__main__":
    unittest.main()
