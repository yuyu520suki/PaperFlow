import tempfile
import unittest
from pathlib import Path
from xml.sax.saxutils import escape

from paperflow.journal import CNS_STYLE, audit_journal
from paperflow.package import NS, write_package


def paragraph(text):
    return f"<w:p><w:r><w:t>{escape(text)}</w:t></w:r></w:p>"


def table(style=None, properties="", cell_properties=""):
    reference = f'<w:tblStyle w:val="{style}"/>' if style else ""
    return (f"<w:tbl><w:tblPr>{reference}{properties}</w:tblPr><w:tr><w:tc>"
            f"<w:tcPr>{cell_properties}</w:tcPr>{paragraph('Cell')}</w:tc></w:tr></w:tbl>")


class JournalTests(unittest.TestCase):
    def audit(self, *body, metadata=None, styles=None, citation_style=None):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "manuscript.docx"
            namespace = f'xmlns:w="{NS["w"]}"'
            content = "".join(body)
            parts = {
                "word/document.xml":
                    f"<w:document {namespace}><w:body>{content}<w:sectPr/></w:body></w:document>".encode(),
            }
            if styles:
                parts["word/styles.xml"] = f"<w:styles {namespace}>{styles}</w:styles>".encode()
            write_package(parts, source)
            report = audit_journal(source, metadata, citation_style=citation_style)
        self.assertFalse(report["full_journal_compliance_proven"])
        return report, {check["name"]: check for check in report["checks"]}

    def test_checks_actual_word_content_even_with_valid_metadata(self):
        report, checks = self.audit(
            paragraph("Abstract"), paragraph(" ".join(["word"] * 251)),
            paragraph("Keywords: a; b; c; d; e; f; g; h"),
            metadata={"abstract": "Short abstract.", "keywords": ["one"]},
        )
        self.assertEqual(checks["abstract_words"]["detail"]["observed_count"], 251)
        self.assertEqual(checks["abstract_words"]["detail"]["body_paragraphs"], [1, 2])
        self.assertEqual(checks["keyword_count"]["detail"]["observed_count"], 8)
        self.assertEqual(set(report["failed"]), {
            "abstract_words", "keyword_count", "abstract_metadata_consistency", "keywords_metadata_consistency",
        })

    def test_explicit_different_csl_style_is_an_advisory(self):
        style = {"citation_fields": 2, "bibliography_fields": 1, "native_refresh_verified": True}
        _, checks = self.audit(paragraph("Paper"), citation_style={**style, "saved_style_id": CNS_STYLE})
        self.assertTrue(checks["journal_reference_style_selection"]["pass"])
        report, checks = self.audit(paragraph("Paper"), citation_style={
            **style, "saved_style_id": "http://www.zotero.org/styles/apa",
        })
        self.assertIsNone(checks["journal_reference_style_selection"]["pass"])
        self.assertEqual(checks["journal_reference_style_selection"]["detail"]["status"], "different_style")
        self.assertIn("journal_reference_style_selection", report["unresolved"])

    def test_labeled_empty_content_fails_and_unrecognized_content_needs_review(self):
        empty, _ = self.audit(paragraph("Abstract"), paragraph("Key words:"), paragraph("1 Introduction"))
        self.assertEqual(set(empty["failed"]), {"abstract_words", "keyword_count"})
        missing, checks = self.audit(paragraph("Unknown manuscript"),
                                     metadata={"abstract": "Valid on its own.", "keywords": ["one"]})
        self.assertEqual(missing["failed"], [])
        self.assertIn("abstract_words", missing["unresolved"])
        self.assertIn("keyword_count", missing["unresolved"])
        self.assertIsNone(checks["abstract_metadata_consistency"]["pass"])
        self.assertIn("competing_interests_attachment", missing["unresolved"])
        self.assertIn("research_data_option_c", missing["unresolved"])

    def test_inline_labels_keyword_separators_and_heading_boundary(self):
        report, checks = self.audit(
            paragraph("Abstract - Two well-known methods."), paragraph("Keywords"),
            paragraph("elasticity, anisotropic； plates; mode shapes"),
            '<w:p><w:pPr><w:pStyle w:val="21"/></w:pPr><w:r><w:t>1. Model setup</w:t></w:r></w:p>',
            paragraph("A long discussion outside the keywords."),
            metadata={"abstract": "Two  well-known methods.",
                      "keywords": ["Elasticity, anisotropic", "Plates", "Mode shapes"]},
            styles='<w:style w:styleId="21"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>',
        )
        self.assertEqual(report["failed"], [])
        self.assertEqual(checks["abstract_words"]["detail"]["observed_count"], 3)
        self.assertEqual(checks["keyword_count"]["detail"]["keywords"], [
            "elasticity, anisotropic", "plates", "mode shapes",
        ])
        self.assertTrue(checks["abstract_metadata_consistency"]["pass"])
        self.assertTrue(checks["keywords_metadata_consistency"]["pass"])

    def test_toc_revisions_hidden_runs_and_table_cells_do_not_supply_labels(self):
        _, checks = self.audit(
            '<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText> TOC </w:instrText>'
            '<w:fldChar w:fldCharType="separate"/></w:r></w:p>',
            paragraph("Abstract"),
            '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
            '<w:p><w:pPr><w:pStyle w:val="10"/></w:pPr><w:r><w:t>Abstract</w:t></w:r></w:p>',
            '<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>Abstract</w:t></w:r></w:p>',
            '<w:tbl><w:tr><w:tc>' + paragraph("Keywords: wrong") + '</w:tc></w:tr></w:tbl>',
            paragraph("Abstract"),
            '<w:p><w:del><w:r><w:delText>old draft</w:delText></w:r></w:del>'
            '<w:ins><w:r><w:t>Visible draft.</w:t></w:r></w:ins></w:p>',
            paragraph("Keywords: plates, frames"),
            styles='<w:style w:styleId="10"><w:name w:val="toc 1"/></w:style>',
        )
        self.assertTrue(checks["abstract_words"]["pass"])
        self.assertEqual(checks["abstract_words"]["detail"]["observed_count"], 2)
        self.assertEqual(checks["keyword_count"]["detail"]["keywords"], ["plates", "frames"])

    def test_section_break_closes_abstract_before_template_furniture(self):
        _, checks = self.audit(
            paragraph("Abstract"),
            '<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>Three short words.</w:t></w:r></w:p>',
            paragraph("This text belongs to a different template section."),
        )
        self.assertTrue(checks["abstract_words"]["pass"])
        self.assertEqual(checks["abstract_words"]["detail"]["observed_count"], 3)

    def test_hidden_styles_do_not_supply_labels_and_direct_visibility_wins(self):
        _, checks = self.audit(
            '<w:p><w:pPr><w:pStyle w:val="HiddenParagraph"/></w:pPr><w:r><w:t>Abstract</w:t></w:r></w:p>',
            '<w:p><w:r><w:rPr><w:rStyle w:val="HiddenCharacter"/></w:rPr><w:t>Keywords: wrong</w:t></w:r></w:p>',
            paragraph("Abstract"),
            '<w:p><w:r><w:rPr><w:rStyle w:val="HiddenCharacter"/><w:vanish w:val="0"/></w:rPr>'
            '<w:t>Visible abstract.</w:t></w:r></w:p>',
            paragraph("Keywords - correct"),
            styles=(
                '<w:style w:styleId="HiddenParagraph"><w:rPr><w:vanish/></w:rPr></w:style>'
                '<w:style w:styleId="HiddenCharacter"><w:basedOn w:val="HiddenBase"/>'
                '<w:rPr><w:vanish w:val="0"/></w:rPr></w:style>'
                '<w:style w:styleId="HiddenBase"><w:rPr><w:vanish/></w:rPr></w:style>'
            ),
        )
        self.assertTrue(checks["abstract_words"]["pass"])
        self.assertEqual(checks["abstract_words"]["detail"]["observed_count"], 2)
        self.assertEqual(checks["keyword_count"]["detail"]["keywords"], ["correct"])

    def test_ambiguous_blocks_do_not_pass(self):
        _, checks = self.audit(paragraph("Abstract"), paragraph("One version."),
                               paragraph("Abstract"), paragraph("Another version."))
        self.assertIsNone(checks["abstract_words"]["pass"])
        _, checks = self.audit(paragraph("Abstract"), paragraph("First part."),
                               paragraph("Methods"), paragraph("Second part."),
                               paragraph("Keywords: elasticity"), paragraph("Introduction"))
        self.assertIsNone(checks["abstract_words"]["pass"])

    def test_highlights_metadata_is_checked_separately_from_attachment(self):
        _, checks = self.audit(paragraph("Paper"), metadata={"highlights": ["one", "two", "three"]})
        self.assertTrue(checks["highlights_limits"]["pass"])
        self.assertIsNone(checks["highlights_attachment"]["pass"])
        _, checks = self.audit(paragraph("Paper"), metadata={"highlights": ["one", "two", ""]})
        self.assertFalse(checks["highlights_limits"]["pass"])

    def test_inherited_table_format_is_reported_for_review(self):
        styles = (
            '<w:style w:type="table" w:styleId="Base" w:default="1"><w:tblPr><w:tblBorders>'
            '<w:left w:val="single"/></w:tblBorders></w:tblPr></w:style>'
            '<w:style w:type="table" w:styleId="Header"><w:basedOn w:val="Base"/>'
            '<w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:val="clear" w:themeFill="accent1"/>'
            '</w:tcPr></w:tblStylePr></w:style>'
        )
        report, checks = self.audit(table("Header"), table(), table("Header", '<w:tblLook w:firstRow="0"/>'),
                                    styles=styles)
        vertical = checks["tables_no_vertical_borders"]
        shading = checks["tables_no_shading"]
        self.assertIsNone(vertical["pass"])
        self.assertIsNone(shading["pass"])
        self.assertEqual(vertical["detail"]["style_review_tables"], [1, 2, 3])
        self.assertEqual(shading["detail"]["style_review_tables"], [1])
        self.assertIn("tables_no_shading", report["unresolved"])

    def test_direct_table_violations_still_fail_with_other_style_risks(self):
        styles = ('<w:style w:type="table" w:styleId="Colored"><w:tcPr>'
                  '<w:shd w:fill="DDDDDD"/></w:tcPr></w:style>')
        _, checks = self.audit(
            table("Colored"),
            table(cell_properties='<w:tcBorders><w:right w:val="single"/></w:tcBorders><w:shd w:fill="AAAAAA"/>'),
            styles=styles,
        )
        self.assertFalse(checks["tables_no_vertical_borders"]["pass"])
        self.assertEqual(checks["tables_no_vertical_borders"]["detail"], [2])
        self.assertFalse(checks["tables_no_shading"]["pass"])
        self.assertEqual(checks["tables_no_shading"]["detail"], {"direct_tables": [2], "style_review_tables": [1]})

    def test_three_line_suppresses_table_grid_but_not_inherited_cell_edges(self):
        grid = ('<w:style w:type="table" w:styleId="TableGrid"><w:tblPr><w:tblBorders>'
                '<w:left w:val="single"/><w:right w:val="single"/><w:insideV w:val="single"/>'
                '</w:tblBorders></w:tblPr></w:style>')
        cell_style = ('<w:style w:type="table" w:styleId="CellEdges"><w:basedOn w:val="TableGrid"/>'
                      '<w:tblStylePr w:type="firstRow"><w:tcPr><w:tcBorders><w:right w:val="single"/>'
                      '</w:tcBorders></w:tcPr></w:tblStylePr></w:style>')
        three_line = ('<w:tblBorders><w:left w:val="nil"/><w:right w:val="nil"/>'
                      '<w:insideV w:val="nil"/></w:tblBorders>')
        _, checks = self.audit(table("TableGrid", three_line), styles=grid)
        self.assertTrue(checks["tables_no_vertical_borders"]["pass"])
        _, checks = self.audit(table("CellEdges", three_line), styles=grid + cell_style)
        self.assertIsNone(checks["tables_no_vertical_borders"]["pass"])
        self.assertEqual(checks["tables_no_vertical_borders"]["detail"]["style_review_tables"], [1])


if __name__ == "__main__":
    unittest.main()
