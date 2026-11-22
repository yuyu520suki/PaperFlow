"""Check note equation ordering against native Word note relationships."""

import unittest

from paperflow.equations import EquationConversionError
from paperflow.note_equations import collect_note_equations
from paperflow.package import NS, q


_NAMESPACES = f'xmlns:w="{NS["w"]}" xmlns:m="{NS["m"]}"'
_PARTS = {"footnote": "word/notes/foot-content.xml", "endnote": "word/notes/end-content.xml"}


def _equation(text):
    return f"<m:oMath><m:r><m:t>{text}</m:t></m:r></m:oMath>"


def _note(kind, identifier, content, note_type=None):
    type_attribute = f' w:type="{note_type}"' if note_type else ""
    return f'<w:{kind} w:id="{identifier}"{type_attribute}><w:p>{content}</w:p></w:{kind}>'


def _parts(anchors, definitions):
    references = "".join(f'<w:r><w:{kind}Reference w:id="{identifier}"/></w:r>' for kind, identifier in anchors)
    parts = {"word/document.xml": (
        f"<w:document {_NAMESPACES}><w:body><w:p>{references}</w:p></w:body></w:document>"
    ).encode()}
    relationships = []
    for index, (kind, notes) in enumerate(definitions.items(), start=1):
        part = _PARTS[kind]
        target = part.removeprefix("word/") if kind == "footnote" else f"/{part}"
        relationships.append(f'<Relationship Id="rId{index}" Type="{NS["r"]}/{kind}s" Target="{target}"/>')
        parts[part] = f'<w:{kind}s {_NAMESPACES}>{"".join(notes)}</w:{kind}s>'.encode()
    parts["word/_rels/document.xml.rels"] = f'<Relationships xmlns="{NS["rel"]}">{"".join(relationships)}</Relationships>'.encode()
    return parts


class NoteEquationTests(unittest.TestCase):
    def test_groups_by_body_reference_order_in_each_note_kind(self):
        parts = _parts(
            [("footnote", 7), ("endnote", 7), ("footnote", 9), ("footnote", 4)],
            {"footnote": [_note("footnote", -1, "<w:r><w:separator/></w:r>", "separator"),
                          _note("footnote", 4, _equation("last")),
                          _note("footnote", 7, _equation("first") + _equation("second"), "normal"),
                          _note("footnote", 9, "<w:r><w:t>No math here</w:t></w:r>")],
             "endnote": [_note("endnote", 7, _equation("end"))]},
        )
        observed = collect_note_equations(parts)
        self.assertEqual(set(observed), {"footnote", "endnote"})
        expected = {"footnote": (("first", "second"), (), ("last",)), "endnote": (("end",),)}
        for kind, expected_groups in expected.items():
            with self.subTest(kind=kind):
                part, groups = observed[kind]
                self.assertEqual(part, _PARTS[kind])
                self.assertIsInstance(groups, tuple)
                for group in groups:
                    self.assertIsInstance(group, tuple)
                    for equation in group:
                        self.assertEqual(equation.tag, q("m:oMath"))
                        self.assertEqual(equation.getparent().tag, q("w:p"))
                self.assertEqual(tuple(tuple(node.find(".//m:t", NS).text for node in group)
                                       for group in groups), expected_groups)

    def test_rejects_ambiguous_or_unreferenced_note_equations(self):
        note = _note("footnote", 1, _equation("x"))
        for label, anchors, notes in (
            ("unreferenced equation", [("footnote", 2)], [note, _note("footnote", 2, "")]),
            ("duplicate reference", [("footnote", 1), ("footnote", 1)], [note]),
            ("duplicate definition", [("footnote", 1)], [note, note]),
        ):
            with self.subTest(label=label):
                with self.assertRaises(EquationConversionError):
                    collect_note_equations(_parts(anchors, {"footnote": notes}))

    def test_rejects_unsupported_nested_or_special_note_equations(self):
        for label, notes in (
            ("nested equation", [_note("footnote", 1, f"<m:oMath>{_equation('x')}</m:oMath>")]),
            ("floating text box", [_note("footnote", 1, f"<w:r><w:pict><w:txbxContent><w:p>{_equation('x')}"
                                                        "</w:p></w:txbxContent></w:pict></w:r>")]),
            ("separator", [_note("footnote", 1, ""), _note("footnote", -1, _equation("x"), "separator")]),
        ):
            with self.subTest(label=label):
                with self.assertRaises(EquationConversionError):
                    collect_note_equations(_parts([("footnote", 1)], {"footnote": notes}))


if __name__ == "__main__":
    unittest.main()
