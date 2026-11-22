"""Citation-owner boundaries and submission preservation without Word or API."""
import json
from contextlib import contextmanager
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from lxml import etree as ET

from paperflow import citations
from paperflow.package import NS, q, read_package, xml
from paperflow.zotero import document_preferences, preference_style_id, read_document_preferences, set_docx_preferences


URI = "http://zotero.org/users/12345/items/ABCDEFGH"


def payload(*, title="Current library title", uri=URI, rendered=True):
    properties = {"noteIndex": 0}
    if rendered:
        properties.update(formattedCitation="[1]", plainCitation="[1]")
    return {"citationID": "unchangedID", "properties": properties,
            "citationItems": [{"id": 7, "uris": [uri], "itemData": {"id": 7, "title": title}}]}


def citation_code(**kwargs):
    return " ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(payload(**kwargs)) + " "


def add_field(paragraph, code, result):
    for kind in ("begin", "instruction", "separate", "result", "end"):
        run = ET.SubElement(paragraph, q("w:r"))
        if kind in {"begin", "separate", "end"}:
            ET.SubElement(run, q("w:fldChar"), {q("w:fldCharType"): kind})
        elif kind == "instruction":
            # Intentionally split a code across runs, as Word commonly does.
            for chunk in (code[:12], code[12:]):
                ET.SubElement(run, q("w:instrText")).text = chunk
        else:
            ET.SubElement(run, q("w:t")).text = result


def make_docx(path, *, zotero=True, pending=False, bibliography=True, citation_text="[1]"):
    root = ET.Element(q("w:document"), nsmap={key: value for key, value in NS.items() if key in {"w", "r", "o", "m"}})
    body = ET.SubElement(root, q("w:body"))
    if zotero:
        add_field(ET.SubElement(body, q("w:p")), citation_code(), "PAPERFLOW_PENDING" if pending else citation_text)
        if bibliography:
            add_field(ET.SubElement(body, q("w:p")), ' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY ', "[1] Current library title")
    for code, value in (("SEQ Figure", "1"), ("REF fig1", "1"), ("PAGE", "1"), ('TOC \\o "1-3"', "Contents")):
        add_field(ET.SubElement(body, q("w:p")), code, value)
    other_addin = ET.SubElement(body, q("w:p"))
    add_field(other_addin, "ADDIN OTHER_VENDOR native data", "other add-in result")
    obj = ET.SubElement(ET.SubElement(ET.SubElement(body, q("w:p")), q("w:r")), q("w:object"))
    ET.SubElement(obj, q("o:OLEObject"), {"ProgID": "Equation.DSMT4", q("r:id"): "rIdMath"})
    with ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        package.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
        package.writestr("word/document.xml", ET.tostring(root))
        package.writestr("word/embeddings/oleObject1.bin", b"fixture-native-payload\x00\x01")
        package.writestr("word/_rels/document.xml.rels", b"fixture-native-relationship")
    return path


class CitationCopyTests(unittest.TestCase):
    def test_nested_ordinary_field_in_citation_result_remains_dynamic(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = make_docx(Path(temporary) / "master.docx")
            parts = read_package(source)
            root = xml(parts["word/document.xml"])
            outer = next(f for f in citations._scan_fields(root) if citations._owner(f["code"]) == "ITEM")
            paragraph = outer["end"].getparent().getparent()
            index = paragraph.index(outer["end"].getparent())
            nested = ET.Element(q("w:p"))
            add_field(nested, "REF preservedNestedReference", "2")
            for run in list(nested):
                paragraph.insert(index, run)
                index += 1
            with ZipFile(source, "w") as package:
                for name, value in parts.items():
                    package.writestr(name, ET.tostring(root) if name == "word/document.xml" else value)
            output = Path(temporary) / "submission.docx"
            report = citations.create_submission_copy(source, output)
            self.assertEqual(report["before"]["ordinary_fields"], 6)
            self.assertEqual(report["after"]["ordinary_fields"], 6)
            saved = citations._package_state(read_package(output))
            self.assertIn("REF preservedNestedReference", [code for _, code in saved["ordinary_fields"]])

    def test_submission_unlinks_only_zotero_and_retains_every_opaque_part(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = make_docx(Path(temporary) / "master.docx")
            output = Path(temporary) / "submission.docx"
            original = source.read_bytes()
            before = read_package(source)
            with patch.object(citations.word, "private_word", side_effect=AssertionError("Word must not start")):
                report = citations.create_submission_copy(source, output)
            after = read_package(output)
            self.assertEqual(source.read_bytes(), original)
            self.assertEqual(report["before"]["citation_fields"], 1)
            self.assertEqual(report["after"]["citation_fields"], 0)
            self.assertEqual(report["after"]["ordinary_fields"], 5)
            self.assertTrue(report["ole_object_types_and_payloads_preserved"])
            self.assertEqual(report["changed_parts"], ["word/document.xml"])
            for name in before:
                if name != "word/document.xml":
                    self.assertEqual(before[name], after[name])
            result = xml(after["word/document.xml"])
            self.assertIn("Current library title", "".join(result.itertext()))
            self.assertNotIn("ZOTERO_ITEM", "".join(result.itertext()))

    def test_pending_submission_fails_without_output(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = make_docx(Path(temporary) / "master.docx", pending=True)
            output = Path(temporary) / "submission.docx"
            with self.assertRaisesRegex(citations.CitationError, "pending/empty"):
                citations.create_submission_copy(source, output)
            self.assertFalse(output.exists())

    def test_no_zotero_refresh_is_explicit_and_never_starts_word(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = make_docx(Path(temporary) / "master.docx", zotero=False)
            output = Path(temporary) / "copy.docx"
            with patch.object(citations.word, "private_word", side_effect=AssertionError("Word must not start")):
                result = citations.refresh_zotero(source, output)
            self.assertEqual(result["status"], "not_needed")
            self.assertEqual(result["reason"], "no_zotero_fields")
            self.assertFalse(result["macro_invoked"])
            self.assertFalse(result["zotero_native_refresh_verified"])
            self.assertEqual(source.read_bytes(), output.read_bytes())

    def test_no_citation_items_in_bibliography_returns_without_word_or_api(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = make_docx(Path(temporary) / "master.docx", zotero=False)
            parts = read_package(source)
            root = xml(parts["word/document.xml"])
            add_field(ET.SubElement(root.find(q("w:body")), q("w:p")), 'ADDIN ZOTERO_BIBL {"uncited":[]} CSL_BIBLIOGRAPHY', "")
            # Rewrite only a test fixture, never a user's source.
            with ZipFile(source, "w") as package:
                for name, value in parts.items():
                    package.writestr(name, ET.tostring(root) if name == "word/document.xml" else value)
            with patch.object(citations, "ZoteroClient", side_effect=AssertionError("API must not start")), patch.object(citations.word, "private_word", side_effect=AssertionError("Word must not start")):
                report = citations.refresh_zotero(source, Path(temporary) / "copy.docx")
            self.assertEqual(report["reason"], "bibliography_has_no_citation_or_uncited_items")


class ZoteroVerificationTests(unittest.TestCase):
    def test_timeout_never_publishes_pending_results_or_restarts_macro(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = make_docx(Path(temporary) / "master.docx", pending=True)
            output = Path(temporary) / "refreshed.docx"

            class App:
                def __init__(self):
                    self.calls = []

                def Run(self, name):
                    self.calls.append(name)

            class Doc:
                def Activate(self):
                    pass

            @contextmanager
            def opened(*_):
                yield Doc()

            app = App()
            fields = self.fields(rendered=False)
            fields[0]["result"] = "PAPERFLOW_PENDING"
            with patch.object(citations, "_expectations", return_value={URI: "Current library title"}), patch.object(citations, "_style_preflight", return_value={}), patch.object(citations, "_pump_messages"), patch.object(citations, "_live_fields", return_value=fields), patch.object(citations.word, "_opened", opened), patch.object(citations.word, "_save_docx", side_effect=lambda doc, stage: stage.write_bytes(source.read_bytes())):
                with self.assertRaises(citations.ZoteroRefreshTimeout):
                    citations.refresh_zotero(source, output, app=app, timeout=0.001)
            self.assertEqual(app.calls, ["Zotero.ZoteroRefresh"])
            self.assertFalse(output.exists())
            self.assertEqual(sorted(path.name for path in Path(temporary).iterdir()), ["master.docx"])

    def fields(self, **kwargs):
        return [{"owner": "ITEM", "code": citation_code(**kwargs), "result": "[1]"},
                {"owner": "BIBL", "code": "ADDIN ZOTERO_BIBL {} CSL_BIBLIOGRAPHY", "result": "Reference entry"}]

    def errors(self, fields):
        return citations._verification_errors(fields, {URI: "Current library title"}, {"citation_fields": 1, "bibliography_fields": 1})

    def test_current_library_metadata_and_rendered_fields_are_required(self):
        self.assertFalse(self.errors(self.fields()))
        self.assertTrue(any("current library title" in e for e in self.errors(self.fields(title="Outdated metadata"))))
        self.assertTrue(any("not rendered" in e for e in self.errors(self.fields(rendered=False))))
        self.assertTrue(any("real library identity" in e for e in self.errors(self.fields(uri="http://zotero.org/users/0/items/ABCDEFGH"))))

    def test_pending_bibliography_and_wrong_visible_citation_are_rejected(self):
        fields = self.fields()
        fields[0]["result"] = "[99]"
        fields[1]["result"] = "PAPERFLOW_PENDING_BIBLIOGRAPHY"
        errors = self.errors(fields)
        self.assertTrue(any("plainCitation" in e for e in errors))
        self.assertTrue(any("pending" in e for e in errors))


class ZoteroStylePreflightTests(unittest.TestCase):
    def test_style_switch_preserves_document_preferences_and_source(self):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            source = make_docx(folder / "master.docx")
            old_id = "http://www.zotero.org/styles/apa"
            requested = "https://www.zotero.org/styles/ieee"
            initial = document_preferences(old_id, locale="en-GB", session_id="originalSession")
            initial = initial.replace('bibliographyStyleHasBeenSet="0"', 'bibliographyStyleHasBeenSet="1"')
            initial = initial.replace("</prefs>", '<pref name="delayCitationUpdates" value="true"/></prefs>')
            set_docx_preferences(source, initial)
            before = source.read_bytes()
            same, unchanged = citations.prepare_citation_style(source, folder / "same.docx", old_id)
            self.assertEqual(same, source)
            self.assertFalse(unchanged["preferences_changed"])
            self.assertFalse((folder / "same.docx").exists())
            changed, report = citations.prepare_citation_style(source, folder / "changed.docx", requested)
            actual = read_document_preferences(changed)
            self.assertEqual(preference_style_id(actual), "http://www.zotero.org/styles/ieee")
            self.assertEqual(actual.find("style").get("locale"), "en-GB")
            self.assertEqual(actual.find("style").get("bibliographyStyleHasBeenSet"), "0")
            self.assertNotEqual(actual.find("session").get("id"), "originalSession")
            self.assertEqual(actual.find("./prefs/pref[@name='delayCitationUpdates']").get("value"), "true")
            self.assertEqual(report["previous_style_id"], old_id)
            self.assertEqual(source.read_bytes(), before)
            with self.assertRaisesRegex(citations.CitationError, "changed the document's Zotero citation style"):
                citations.verify_zotero_fields_preserved(source, changed)

    def test_manually_edited_citation_blocks_automatic_style_change(self):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            source = make_docx(folder / "master.docx", citation_text="Manually edited display")
            set_docx_preferences(source, document_preferences("http://www.zotero.org/styles/apa"))
            output = folder / "changed.docx"
            with self.assertRaisesRegex(citations.CitationError, "manually edited"):
                citations.prepare_citation_style(source, output, "http://www.zotero.org/styles/ieee")
            self.assertFalse(output.exists())

    def test_custom_bibliography_blocks_automatic_style_change(self):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            source = make_docx(folder / "master.docx")
            set_docx_preferences(source, document_preferences("http://www.zotero.org/styles/apa"))
            fields = [{"owner": "BIBL", "code": 'ADDIN ZOTERO_BIBL {"custom":[["key","edited entry"]]}',
                       "result": "edited entry"}]
            output = folder / "changed.docx"
            with patch.object(citations, "_package_state", return_value={"fields": fields}):
                with self.assertRaisesRegex(citations.CitationError, "manually edited entries"):
                    citations.prepare_citation_style(source, output, "http://www.zotero.org/styles/ieee")
            self.assertFalse(output.exists())

    def test_missing_parent_refuses_refresh_without_installing(self):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            source = make_docx(folder / "master.docx")
            set_docx_preferences(source, document_preferences("http://www.zotero.org/styles/custom-dependent"))
            styles = folder / "styles"; styles.mkdir()
            (styles / "dependent.csl").write_text('<style xmlns="http://purl.org/net/xbiblio/csl"><info><id>http://www.zotero.org/styles/custom-dependent</id><link rel="independent-parent" href="http://www.zotero.org/styles/custom-parent"/></info></style>', encoding="utf-8")
            with self.assertRaisesRegex(citations.CitationError, "not installed"):
                citations._style_preflight(source, styles)
            (styles / "parent.csl").write_text('<style xmlns="http://purl.org/net/xbiblio/csl"><info><id>http://www.zotero.org/styles/custom-parent</id></info></style>', encoding="utf-8")
            report = citations._style_preflight(source, styles)
            self.assertEqual(report["installed_style_chain"], ["http://www.zotero.org/styles/custom-parent", "http://www.zotero.org/styles/custom-dependent"])
            self.assertFalse(report["styles_installed_by_operation"])


if __name__ == "__main__":
    unittest.main()
