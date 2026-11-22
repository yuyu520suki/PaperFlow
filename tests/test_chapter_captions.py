import tempfile
import unittest
from pathlib import Path

from paperflow.chapter_captions import (validate_chapter_captions, validate_chapter_migration_values,
                                       verify_generated_reference_values)
from paperflow.package import NS, write_package


NAMESPACE = f'xmlns:w="{NS["w"]}"'


def heading(text, style="1"):
    return f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'


def caption(name, identifier, text="Caption"):
    return (f'<w:p><w:bookmarkStart w:id="{identifier}" w:name="{name}"/>'
            f'<w:r><w:t>{text}</w:t></w:r><w:bookmarkEnd w:id="{identifier}"/></w:p>')


def simple_ref(name, result):
    return (f'<w:fldSimple w:instr=" REF &quot;{name}&quot; \\h ">'
            f'<w:r><w:t>{result}</w:t></w:r></w:fldSimple>')


def table(paragraph):
    return f"<w:tbl><w:tr><w:tc>{paragraph}</w:tc></w:tr></w:tbl>"


class ChapterCaptionTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.path = Path(directory.name) / "built.docx"

    def prepare(self, *body):
        styles = ('<w:style w:type="paragraph" w:styleId="1"><w:name w:val="标题 1"/></w:style>'
                  '<w:style w:type="paragraph" w:styleId="Heading1" w:customStyle="1">'
                  '<w:name w:val="Custom chapter"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>')
        parts = {
            "word/document.xml": f'<w:document {NAMESPACE}><w:body>{"".join(body)}</w:body></w:document>'.encode(),
            "word/styles.xml": f"<w:styles {NAMESPACE}>{styles}</w:styles>".encode(),
        }
        write_package(parts, self.path)

    def test_uses_preceding_localized_heading_and_allows_caption_in_table(self):
        self.prepare(heading("First chapter"), caption("FigureOne", 1),
                     heading("Second chapter"), table(caption("EquationInTable", 2)))
        report = validate_chapter_captions(self.path, 1, ["figureone", "EquationInTable"])
        self.assertEqual(report["bookmarks_checked"], 2)
        self.assertEqual(report["builtin_heading_paragraphs"], 2)
        self.assertEqual([entry["preceding_heading_paragraph"] for entry in report["bookmarks"]], [1, 3])

    def test_toc_deleted_text_box_and_table_headings_do_not_supply_context(self):
        begin_toc = ('<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText> TO</w:instrText></w:r>'
                     '<w:r><w:instrText>C </w:instrText><w:fldChar w:fldCharType="separate"/></w:r></w:p>')
        end_toc = '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
        self.prepare(begin_toc, heading("A TOC entry"), end_toc,
                     '<w:del>' + heading("A deleted heading") + '</w:del>',
                     '<w:p><w:r><w:pict><w:txbxContent>' + heading("A text box heading")
                     + '</w:txbxContent></w:pict></w:r></w:p>',
                     table(heading("A heading in a table")), heading("A custom heading", "Heading1"),
                     caption("EarlyFigure", 1), heading("First real heading"))
        with self.assertRaisesRegex(ValueError, "EarlyFigure.*no preceding built-in Heading 1"):
            validate_chapter_captions(self.path, 1, ["EarlyFigure"])

    def test_duplicate_bookmarks_cannot_identify_a_caption(self):
        self.prepare(heading("First chapter"), caption("Duplicate", 1), caption("duplicate", 2))
        with self.assertRaisesRegex(ValueError, "must occur once.*found 2"):
            validate_chapter_captions(self.path, 1, ["Duplicate"])

    def test_compares_saved_bookmark_and_split_ref_fields_in_main_story(self):
        complex_ref = ('<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText> RE</w:instrText></w:r>'
                       '<w:r><w:instrText>F FigureOne \\h </w:instrText><w:fldChar w:fldCharType="separate"/></w:r>'
                       '<w:r><w:t>\u200e2.</w:t></w:r><w:r><w:t>1</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p>')
        outside = ('<w:p><w:r><w:pict><w:txbxContent>' + caption("FigureOne", 4, "5.1")
                   + f'<w:p>{simple_ref("FigureOne", "5.1")}</w:p></w:txbxContent></w:pict></w:r></w:p>')
        index_ref = f'<w:p><w:fldSimple w:instr=" TOC ">{simple_ref("FigureOne", "5.1")}</w:fldSimple></w:p>'
        self.prepare(heading("Chapter"), caption("FigureOne", 1, " \u200e2.1 "), complex_ref,
                     f'<w:p>{simple_ref("FIGUREONE", "2.1")}</w:p>', index_ref, outside)
        expected = [{"bookmark": "figureone", "number": "2.1", "references": 2}]
        evidence = validate_chapter_migration_values(self.path, expected)
        self.assertEqual(evidence["bookmarks"], expected)
        self.assertEqual(evidence["references_checked"], 2)

    def test_reports_changed_number_and_reference_value_without_normalizing_separators(self):
        self.prepare(heading("Chapter"), caption("FigureOne", 1, "3.1"),
                     f'<w:p>{simple_ref("FigureOne", "2-1")}</w:p>')
        expected = [{"bookmark": "FigureOne", "number": "2.1", "references": 2}]
        with self.assertRaisesRegex(ValueError, r"starting number.*observed \['3\.1'\].*2 REF fields.*\['2-1'\]"):
            validate_chapter_migration_values(self.path, expected)

    def test_generated_references_reject_a_stale_word_result(self):
        self.prepare(caption("EquationOne", 1, "2"), f'<w:p>{simple_ref("EquationOne", "1")}</w:p>')
        targets = [{"bookmark": "EquationOne", "kind": "Equation", "label": "eq:mass"}]
        with self.assertRaisesRegex(ValueError, "did not refresh.*EquationOne"):
            verify_generated_reference_values(self.path, targets, [{"bookmark": "EquationOne"}])


if __name__ == "__main__":
    unittest.main()
