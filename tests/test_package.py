import json
import tempfile
import unittest
import zipfile
from pathlib import Path

from paperflow.package import (NS, audit_docx, expand_markers, marked_zotero_keys, navigation_index_kind,
                              part_fields, q, read_package, xml)


def fixture(path, text, extra=None):
    from xml.sax.saxutils import escape
    w = NS["w"]
    parts = {"word/document.xml": f'<w:document xmlns:w="{w}"><w:body><w:p><w:r><w:t>{escape(text)}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'.encode()}
    parts.update(extra or {})
    with zipfile.ZipFile(path, "w") as z:
        for k, v in parts.items(): z.writestr(k, v)


def note_fixture(path, text, *, reference_count=1, note_type="normal"):
    from xml.sax.saxutils import escape

    w = NS["w"]
    relationships = []
    references = []
    extra = {"word/embeddings/oleObject1.bin": b"opaque native equation"}
    for kind in ("footnote", "endnote"):
        target = f"notes/fixture-{kind}s.xml"
        relationships.append(
            f'<Relationship Id="{kind}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}s" Target="{target}"/>'
        )
        references.extend(f'<w:r><w:{kind}Reference w:id="1"/></w:r>' for _ in range(reference_count))
        extra["word/" + target] = (
            f'<w:{kind}s xmlns:w="{w}"><w:{kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:{kind}>'
            f'<w:{kind} w:type="{note_type}" w:id="1"><w:p><w:pPr><w:pStyle w:val="FixtureNote"/></w:pPr>'
            f'<w:r><w:{kind}Ref/></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>{escape(text)}</w:t></w:r></w:p></w:{kind}></w:{kind}s>'
        ).encode()
    extra["word/_rels/document.xml.rels"] = (
        f'<Relationships xmlns="{NS["rel"]}">{"".join(relationships)}</Relationships>'
    ).encode()
    extra["word/document.xml"] = (
        f'<w:document xmlns:w="{w}"><w:body><w:p><w:r><w:t>[[ANCHOR Chapter]]Methods</w:t></w:r>'
        f'{"".join(references)}</w:p><w:sectPr/></w:body></w:document>'
    ).encode()
    fixture(path, "", extra)


class PackageTests(unittest.TestCase):
    def test_markers_make_dynamic_fields_and_preserve_native_bytes(self):
        with tempfile.TemporaryDirectory() as temp:
            src, dst = Path(temp)/"src.docx", Path(temp)/"dst.docx"
            native = b"opaque existing native equation content"
            fixture(src, "[[CAPTION Figure response|Response curve]] See [[REF response]]", {"word/embeddings/oleObject1.bin": native})
            expand_markers(src, dst)
            result = audit_docx(dst)
            self.assertEqual(result["errors"], [])
            self.assertEqual(read_package(dst)["word/embeddings/oleObject1.bin"], native)
            self.assertEqual(len(result["fields"]), 2)
            self.assertFalse(result["structural_pass"], "Unrefreshed visible fields must remain pending")
            self.assertFalse(result["complete"])

    def test_missing_reference_is_rejected_by_audit(self):
        with tempfile.TemporaryDirectory() as temp:
            src, dst = Path(temp)/"src.docx", Path(temp)/"dst.docx"
            fixture(src, "[[REF unknown]]")
            expand_markers(src, dst)
            self.assertIn("Missing cross-reference target: unknown", audit_docx(dst)["errors"])

    def test_page_reference_uses_a_real_word_bookmark_and_field(self):
        with tempfile.TemporaryDirectory() as temp:
            src, dst = Path(temp) / "source.docx", Path(temp) / "expanded.docx"
            fixture(src, "[[ANCHOR Chapter]]Heading text and page [[PAGEREF Chapter]].")
            self.assertIn("[[PAGEREF ", audit_docx(src)["pending_markers"])
            expand_markers(src, dst)
            report = audit_docx(dst)
            self.assertEqual(report["bookmarks"], ["Chapter"])
            self.assertEqual(report["fields"][0]["instruction"].strip(), r"PAGEREF Chapter \h")
            self.assertEqual(report["errors"], [])
            self.assertEqual(report["pending_markers"], ["PAPERFLOW_PENDING"])

    def test_only_recognizable_template_indexes_override_source_requests(self):
        for instruction, expected in (
            (r'TOC \o "1-3" \h \z', "TOC"), (r'TOC \h \z', "TOC"),
            (r'TOC \c "Figure" \h', "LOF"), (r'toc \c Table \h', "LOT"),
            (r'TOC \a "Scheme"', None), (r'TOC \t "Custom Style,1"', None),
            (r'TOC \f Q', None), (r'PAGEREF Heading \h', None),
        ):
            with self.subTest(instruction=instruction):
                self.assertEqual(navigation_index_kind(instruction), expected)

    def test_output_never_overwrites_original(self):
        with tempfile.TemporaryDirectory() as temp:
            src = Path(temp)/"src.docx"
            fixture(src, "[[TOC]]")
            previous = src.read_bytes()
            with self.assertRaises(FileExistsError): expand_markers(src, src)
            self.assertEqual(src.read_bytes(), previous)

    def test_caption_list_markers_create_distinct_native_fields(self):
        with tempfile.TemporaryDirectory() as temp:
            for marker, label in (("LOF", "Figure"), ("LOT", "Table")):
                with self.subTest(marker=marker):
                    src, dst = Path(temp)/f"{marker}-src.docx", Path(temp)/f"{marker}-dst.docx"
                    fixture(src, f"[[{marker}]]")
                    self.assertIn(f"[[{marker}]]", audit_docx(src)["pending_markers"])
                    expand_markers(src, dst)
                    report = audit_docx(dst)
                    self.assertEqual(report["fields"][0]["instruction"].strip(), f'TOC \\h \\z \\c "{label}"')
                    self.assertEqual(report["pending_markers"], ["PAPERFLOW_PENDING"])

    def test_caption_list_must_be_its_own_paragraph(self):
        with tempfile.TemporaryDirectory() as temp:
            src, dst = Path(temp)/"src.docx", Path(temp)/"dst.docx"
            fixture(src, "List of figures [[LOF]]")
            with self.assertRaisesRegex(ValueError, "only content"):
                expand_markers(src, dst)
            self.assertFalse(dst.exists())

    def test_chapter_numbers_place_both_fields_and_separator_in_ref_bookmark(self):
        with tempfile.TemporaryDirectory() as temp:
            for marker, kind, expected_style in (
                ("[[CAPTION Table label|Title]]", "Table", "ThesisCaption"),
                ("[[CAPTIONNO Figure label]] Figure text", "Figure", "ThesisCaption"),
                ("x [[EQNO label]]", "Equation", "ThesisEquation"),
            ):
                with self.subTest(kind=kind):
                    src, dst = Path(temp)/f"{kind}-src.docx", Path(temp)/f"{kind}-dst.docx"
                    fixture(src, marker)
                    report = expand_markers(src, dst, caption_style="ThesisCaption", equation_style="ThesisEquation",
                                            caption_chapter_level=2, caption_chapter_separator=".")
                    paragraph = xml(read_package(dst)["word/document.xml"]).find(".//w:p", NS)
                    self.assertEqual(paragraph.find("w:pPr/w:pStyle", NS).get(q("w:val")), expected_style)
                    start = paragraph.index(paragraph.find(q("w:bookmarkStart")))
                    end = paragraph.index(paragraph.find(q("w:bookmarkEnd")))
                    self.assertEqual(part_fields(paragraph), ([r" STYLEREF 2 \s ", f" SEQ {kind} \\* ARABIC \\s 2 "], []))
                    self.assertEqual("".join(
                        node.text for child in paragraph[start + 1:end] for node in child.iter(q("w:t"))
                    ), "PAPERFLOW_PENDING.PAPERFLOW_PENDING")
                    self.assertEqual(report["chapter_bookmarks"], ["label"])

    def test_duplicate_bookmark_rejected_before_output(self):
        with tempfile.TemporaryDirectory() as temp:
            src, dst = Path(temp)/"src.docx", Path(temp)/"dst.docx"
            fixture(src, "[[CAPTION Table example|A]] [[CAPTION Table example|B]]")
            with self.assertRaisesRegex(ValueError, "Duplicate bookmark"):
                expand_markers(src, dst)
            self.assertFalse(dst.exists())

    def test_note_markers_create_native_fields_and_preserve_note_identity(self):
        with tempfile.TemporaryDirectory() as temporary:
            source, output = Path(temporary) / "source.docx", Path(temporary) / "expanded.docx"
            note_fixture(source, "Evidence [[CITE ABCD1234]] target [[REF Chapter]] page [[PAGEREF Chapter]] section [[HEADINGREF Chapter]].")
            original = source.read_bytes()
            before = read_package(source)
            self.assertEqual(marked_zotero_keys(source), {"ABCD1234"})
            item = {"uri": "http://zotero.org/users/123456/items/ABCD1234",
                    "item_data": {"type": "article-journal", "title": "Fixture metadata"}}
            expansion = expand_markers(source, output, citation_items={"ABCD1234": item})
            self.assertTrue(expansion["citations_used"])
            after = read_package(output)
            citation_ids = []
            for kind in ("footnote", "endnote"):
                name = f"word/notes/fixture-{kind}s.xml"
                root = xml(after[name])
                (note,) = [node for node in root if node.get(q("w:id")) == "1"]
                self.assertEqual(note.get(q("w:type")), "normal")
                self.assertEqual(len(list(note.iter(q(f"w:{kind}Ref")))), 1)
                self.assertEqual(note.find(".//w:pStyle", NS).get(q("w:val")), "FixtureNote")
                self.assertEqual(len(root.findall(f'w:{kind}[@w:type="separator"]', NS)), 1)
                fields, errors = part_fields(root)
                self.assertEqual(errors, [])
                self.assertEqual([code.strip() for code in fields[1:]],
                                 [r"REF Chapter \h", r"PAGEREF Chapter \h", r"REF Chapter \h \w \t"])
                payload = json.loads(fields[0].split("CSL_CITATION", 1)[1])
                self.assertEqual(payload["citationItems"][0]["uris"], [item["uri"]])
                citation_ids.append(payload["citationID"])
                for run in note.findall(".//w:r", NS):
                    if run.find(q("w:t")) is not None:
                        self.assertIsNotNone(run.find("w:rPr/w:i", NS))
            self.assertEqual(len(set(citation_ids)), 2)
            for name in ("word/embeddings/oleObject1.bin", "word/_rels/document.xml.rels"):
                self.assertEqual(after[name], before[name])
            main = xml(after["word/document.xml"])
            for kind in ("footnote", "endnote"):
                self.assertEqual([ref.get(q("w:id")) for ref in main.iter(q(f"w:{kind}Reference"))], ["1"])
            audit = audit_docx(output)
            self.assertEqual(audit["errors"], [])
            self.assertEqual(audit["pending_markers"], ["PAPERFLOW_PENDING"])
            self.assertEqual(set(audit["pending_marker_parts"]),
                             {"word/notes/fixture-footnotes.xml", "word/notes/fixture-endnotes.xml"})
            self.assertEqual(source.read_bytes(), original)

    def test_unsupported_note_markers_require_body_placement(self):
        with tempfile.TemporaryDirectory() as temporary:
            for index, marker in enumerate(("[[BIBLIOGRAPHY]]", "[[TOC]]", "[[ANCHOR nested]]")):
                with self.subTest(marker=marker):
                    source, output = Path(temporary) / f"source-{index}.docx", Path(temporary) / f"output-{index}.docx"
                    note_fixture(source, marker)
                    with self.assertRaisesRegex(ValueError, "not supported in a footnote"):
                        expand_markers(source, output)
                    self.assertFalse(output.exists())

    def test_note_markers_require_one_normal_note_definition_and_body_reference(self):
        with tempfile.TemporaryDirectory() as temporary:
            for index, options in enumerate((
                {"reference_count": 0}, {"reference_count": 2}, {"note_type": "separator"},
            )):
                with self.subTest(options=options):
                    source, output = Path(temporary) / f"source-{index}.docx", Path(temporary) / f"output-{index}.docx"
                    note_fixture(source, "[[REF Chapter]]", **options)
                    with self.assertRaisesRegex(ValueError, "uniquely referenced normal footnote"):
                        expand_markers(source, output)
                    self.assertFalse(output.exists())

    def test_audit_finds_pending_markers_in_notes_and_headers(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "source.docx"
            note_fixture(source, "[[CITE ABCD1234]]")
            parts = read_package(source)
            parts["word/header1.xml"] = (
                f'<w:hdr xmlns:w="{NS["w"]}"><w:p><w:r><w:t>[[PAGEREF Chapter]]</w:t></w:r></w:p></w:hdr>'
            ).encode()
            with zipfile.ZipFile(source, "w") as package:
                for name, data in parts.items():
                    package.writestr(name, data)
            audit = audit_docx(source)
            self.assertFalse(audit["structural_pass"])
            self.assertEqual(audit["pending_marker_parts"]["word/header1.xml"], ["[[PAGEREF "])
            for kind in ("footnote", "endnote"):
                self.assertEqual(audit["pending_marker_parts"][f"word/notes/fixture-{kind}s.xml"], ["[[CITE "])


if __name__ == "__main__": unittest.main()
