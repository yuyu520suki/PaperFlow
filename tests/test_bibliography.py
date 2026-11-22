from pathlib import Path
import tempfile
import unittest

from docx import Document

from paperflow.bibliography import apply_migration, plan_migration
from paperflow.package import digest, q, read_package, xml


class _Zotero:
    dois = {"AAAA0001": "10.5555/fixture", "AAAA0002": "10.5555/fixture.", "AAAA0003": "10.5555/different"}

    def search(self, query, **kwargs):
        return [{"key": key, "data": {"DOI": doi, "title": key}} for key, doi in self.dois.items() if doi == query]

    def get_item(self, key):
        return {"item_data": {"DOI": self.dois[key], "title": key}}


class BibliographyMigrationTests(unittest.TestCase):
    def test_doi_punctuation_ambiguity_and_appendix_page_break(self):
        with tempfile.TemporaryDirectory() as folder:
            source, target = Path(folder) / "source.docx", Path(folder) / "target.docx"
            doc = Document()
            doc.add_paragraph("The cited work [1] describes the model.")
            doc.add_heading("References", 1)
            doc.add_paragraph("[1] Example DOI: 10.5555/fixture.")
            doc.add_page_break()
            doc.add_heading("Appendix A", 1)
            doc.add_paragraph("Appendix evidence")
            doc.save(source)
            before = digest(source.read_bytes())
            client = _Zotero()
            plan = plan_migration(source, client=client)
            self.assertEqual(plan["citation_map"], {})
            self.assertEqual({candidate["key"] for candidate in plan["entries"][0]["candidates"]}, {"AAAA0001", "AAAA0002"})
            with self.assertRaisesRegex(ValueError, "does not have the reviewed DOI"):
                apply_migration(source, {**plan, "citation_map": {"1": "AAAA0003"}}, target, client=client)
            self.assertFalse(target.exists())
            report = apply_migration(source, {**plan, "citation_map": {"1": "AAAA0002"}}, target, client=client)
            self.assertEqual(report["identities"][0]["doi"], "10.5555/fixture.")
            self.assertEqual(digest(source.read_bytes()), before)
            root = xml(read_package(target)["word/document.xml"])
            self.assertTrue(any(node.get(q("w:type")) == "page" for node in root.iter(q("w:br"))))
            self.assertIn("Appendix evidence", "".join(root.itertext()))


if __name__ == "__main__":
    unittest.main()
