import unittest
from copy import deepcopy

from lxml import etree as ET

from paperflow.floating import ANCHOR, INLINE, inline_picture_copy
from paperflow.package import q


_DRAWING = b"""
<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <wp:anchor distT="0" distB="0" distL="114300" distR="114300"
    simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0"
    layoutInCell="1" allowOverlap="1" wp14:anchorId="23456789" wp14:editId="01234567">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="column"><wp:posOffset>50000</wp:posOffset></wp:positionH>
    <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
    <wp:extent cx="1828800" cy="914400"/>
    <wp:effectExtent l="0" t="0" r="127" b="0"/>
    <wp:wrapSquare wrapText="bothSides"/>
    <wp:docPr id="42" name="Picture 42" descr="Fixture figure"/>
    <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr><pic:cNvPr id="0" name="Plot.png"/><pic:cNvPicPr/></pic:nvPicPr>
          <pic:blipFill>
            <a:blip r:embed="rId5"/>
            <a:srcRect l="2500" r="7500"/>
            <a:stretch><a:fillRect/></a:stretch>
          </pic:blipFill>
          <pic:spPr>
            <a:xfrm rot="600000" flipH="1"><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
    <wp14:sizeRelH relativeFrom="margin"><wp14:pctWidth>0</wp14:pctWidth></wp14:sizeRelH>
    <wp14:sizeRelV relativeFrom="margin"><wp14:pctHeight>0</wp14:pctHeight></wp14:sizeRelV>
  </wp:anchor>
</w:drawing>
"""


class FloatingPictureTests(unittest.TestCase):
    def test_inlining_preserves_picture_payload_and_original_anchor(self):
        original = ET.fromstring(_DRAWING)
        original_bytes = ET.tostring(original)
        clone = deepcopy(original)
        keep = [child for child in clone[0] if ET.QName(child).localname in {
            "extent", "effectExtent", "docPr", "cNvGraphicFramePr", "graphic",
        }]
        self.assertIs(inline_picture_copy(clone), clone)
        self.assertEqual(clone[0].tag, INLINE)
        self.assertEqual(list(clone[0]), keep)
        self.assertEqual(clone[0].attrib, {})
        self.assertEqual(ET.tostring(original), original_bytes)
        self.assertEqual(original[0].tag, ANCHOR)
        picture = clone.find(".//pic:pic", clone.nsmap)
        original_picture = original.find(".//pic:pic", original.nsmap)
        self.assertEqual(ET.tostring(picture), ET.tostring(original_picture))

    def test_unsupported_drawings_are_rejected_before_mutation(self):
        def linked(drawing):
            drawing.find(".//a:blip", drawing.nsmap).set(q("r:link"), "rId6")

        def chart(drawing):
            drawing.find(".//a:graphicData", drawing.nsmap).set("uri", "http://schemas.openxmlformats.org/drawingml/2006/chart")

        def relative(drawing):
            drawing.find(".//wp14:pctWidth", drawing.nsmap).text = "80000"

        def unexpected(drawing):
            ET.SubElement(drawing[0], f"{{{drawing.nsmap['wp']}}}unsupported")

        for mutate, message in (
            (linked, "linked image"),
            (chart, "not a simple picture"),
            (relative, "relative sizing"),
            (unexpected, "unsupported anchor markup"),
        ):
            with self.subTest(case=mutate.__name__):
                drawing = ET.fromstring(_DRAWING)
                mutate(drawing)
                before = ET.tostring(drawing)
                with self.assertRaisesRegex(ValueError, message):
                    inline_picture_copy(drawing)
                self.assertEqual(ET.tostring(drawing), before)


if __name__ == "__main__":
    unittest.main()
