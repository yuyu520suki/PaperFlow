"""Offline safety/field-order tests. These never initialize Word or COM."""
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from docx import Document
from docx.oxml import OxmlElement

from paperflow import word
from paperflow.package import field_runs, q


def _bookmark(paragraph, name, identifier):
    start, end = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
    start.set(q("w:name"), name)
    start.set(q("w:id"), identifier)
    end.set(q("w:id"), identifier)
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


class Collection:
    def __init__(self, values):
        self.values = list(values)

    @property
    def Count(self):
        return len(self.values)

    def Item(self, index):
        return self.values[index - 1]


class FakeField:
    def __init__(self, code, log, *, locked=False, result=True):
        self.Code = type("Range", (), {"Text": code})()
        self.Locked = locked
        self.log = log
        self.result = result

    def Update(self):
        self.log.append(self.Code.Text.strip())
        return self.result


class FakeStory:
    StoryType = 1
    NextStoryRange = None

    def __init__(self, fields):
        self.Fields = Collection(fields)


class FakeStories:
    def __init__(self, story):
        self.story = story

    def Item(self, kind):
        if kind != 1:
            raise KeyError(kind)
        return self.story


class FakeDocument:
    def __init__(self, story):
        self.StoryRanges = FakeStories(story)
        self.repagination_count = 0

    def Repaginate(self):
        self.repagination_count += 1


class FakeBookmarks:
    def __init__(self, numbering):
        self.numbering = numbering

    def Exists(self, name):
        return name in self.numbering

    def Item(self, name):
        list_type, label = self.numbering[name]
        paragraph_range = SimpleNamespace(ListFormat=SimpleNamespace(ListType=list_type, ListString=label))
        paragraphs = Collection([SimpleNamespace(Range=paragraph_range)])
        return SimpleNamespace(Range=SimpleNamespace(Paragraphs=paragraphs))


class WordSafetyTests(unittest.TestCase):
    def test_zotero_preferences_read_only_owner_chunks(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "source.docx"
            with ZipFile(source, "w") as package:
                package.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>')
                package.writestr("docProps/custom.xml", '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property name="ZOTERO_PREF_1"><vt:lpwstr>&lt;data&gt;用户&lt;/data&gt;</vt:lpwstr></property><property name="Project"><vt:lpwstr>keep me</vt:lpwstr></property></Properties>')
            self.assertEqual(word.read_zotero_preferences(source), {"ZOTERO_PREF_1": "<data>用户</data>"})

    def test_zotero_chunk_copy_replaces_stale_tail_and_keeps_template_properties(self):
        class Property:
            def __init__(self, owner, name, value):
                self.owner, self.Name, self.Value = owner, name, value

            def Delete(self):
                del self.owner.values[self.Name]

        class Properties:
            def __init__(self):
                self.values = {"ZOTERO_PREF_1": "old1", "ZOTERO_PREF_2": "old2", "Project": "retained"}

            @property
            def Count(self):
                return len(self.values)

            def Item(self, name):
                if isinstance(name, int):
                    name = list(self.values)[name - 1]
                return Property(self, name, self.values[name])

            def Add(self, name, linked, property_type, value, /):
                if linked is not False or property_type != 4:
                    raise AssertionError("Zotero preference must be a non-linked string")
                self.values[name] = value

        props = Properties()
        doc = type("Doc", (), {"CustomDocumentProperties": props})()
        with patch.object(word, "read_zotero_preferences", return_value={"ZOTERO_PREF_1": "new"}):
            report = word._copy_zotero_preferences(doc, Path("source.docx"))
        self.assertEqual(props.values, {"Project": "retained", "ZOTERO_PREF_1": "new"})
        self.assertEqual(report["source_chunks"], 1)
        self.assertTrue(report["changed"])
        self.assertFalse(report["citation_ids_rewritten"])

    def test_existing_outputs_and_sources_are_never_overwritten(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "source.docx"
            output = Path(temporary) / "output.docx"
            source.write_bytes(b"original source")
            output.write_bytes(b"original output")
            with patch.object(word, "private_word", side_effect=AssertionError("COM must not start")):
                with self.assertRaises(word.OutputExistsError):
                    word.convert_to_docx(source, output)
                with self.assertRaises(ValueError):
                    word.convert_to_docx(source, source)
            self.assertEqual(source.read_bytes(), b"original source")
            self.assertEqual(output.read_bytes(), b"original output")

    def test_build_rejects_case_insensitive_bookmark_collisions_before_word(self):
        with tempfile.TemporaryDirectory() as temporary:
            source, template, output = (Path(temporary) / name for name in ("source.docx", "template.dotx", "output.docx"))
            manuscript = Document()
            _bookmark(manuscript.add_paragraph("Source target"), "Target", "1")
            _bookmark(manuscript.add_paragraph("Source heading"), "_Toc123", "2")
            manuscript.save(source)
            destination = Document()
            _bookmark(destination.add_paragraph("Template target"), "target", "1")
            _bookmark(destination.add_paragraph("Template heading"), "_TOC123", "2")
            destination.add_paragraph("{{PAPERFLOW_BODY}}")
            destination.save(template)
            with patch.object(word, "_application") as application:
                with self.assertRaisesRegex(ValueError, "Bookmark names must be distinct.*_toc123.*target"):
                    word.build_from_template(template, source, output)
                application.assert_not_called()
            self.assertFalse(output.exists())

    def test_build_bookmark_check_allows_cross_document_refs_and_goback(self):
        with tempfile.TemporaryDirectory() as temporary:
            source, template, output = (Path(temporary) / name for name in ("source.docx", "template.docx", "output.docx"))
            manuscript = Document()
            _bookmark(manuscript.add_paragraph("Source target"), "ManuscriptTarget", "1")
            _bookmark(manuscript.add_paragraph(), "_GoBack", "2")
            manuscript.add_paragraph()._p.extend(field_runs(" REF TemplateTarget ", "Cached"))
            manuscript.save(source)
            destination = Document()
            _bookmark(destination.add_paragraph("Template target"), "TemplateTarget", "1")
            _bookmark(destination.add_paragraph(), "_GoBack", "2")
            destination.add_paragraph()._p.extend(field_runs(" REF ManuscriptTarget ", "Cached"))
            destination.add_paragraph("{{PAPERFLOW_BODY}}")
            destination.save(template)
            with patch.object(word, "_application", side_effect=RuntimeError("Passed the preflight")):
                with self.assertRaisesRegex(RuntimeError, "Passed the preflight"):
                    word.build_from_template(template, source, output)

    def test_publish_does_not_replace_a_racing_writer(self):
        with tempfile.TemporaryDirectory() as temporary:
            stage = Path(temporary) / "stage.docx"
            destination = Path(temporary) / "paper.docx"
            stage.write_bytes(b"new paper")
            destination.write_bytes(b"racing writer")
            with self.assertRaises(word.OutputExistsError):
                word._publish(stage, destination)
            self.assertEqual(destination.read_bytes(), b"racing writer")
            self.assertEqual(stage.read_bytes(), b"new paper")

    def test_publish_fallback_preserves_existing_destination(self):
        with tempfile.TemporaryDirectory() as temporary:
            stage = Path(temporary) / "stage.docx"
            destination = Path(temporary) / "paper.docx"
            stage.write_bytes(b"new paper")
            destination.write_bytes(b"racing writer")
            with patch.object(word.os, "link", side_effect=OSError("unsupported")):
                with self.assertRaises(FileExistsError):
                    word._publish(stage, destination)
            self.assertEqual(destination.read_bytes(), b"racing writer")

    def test_native_word_fails_explicitly_on_other_platforms(self):
        with patch.object(word.os, "name", "posix"):
            with self.assertRaisesRegex(word.WordAutomationError, "requires Windows"):
                with word.private_word():
                    self.fail("must not yield")


class NativeFieldPolicyTests(unittest.TestCase):
    def test_numbered_reference_target_ignores_quoted_format_literals(self):
        cases = [
            (r' REF "SecOne" \h \w \t ', "SecOne"),
            (r" ReF SecOne\h\n", "SecOne"),
            (r" REF SecOne \R \* MERGEFORMAT", "SecOne"),
            (r''' REF SecOne \# "0 'literal \w'" ''', None),
            (r' REF SecOne \# "\" \w \""', None),
            (r" REF SecOne \wrong", None),
            (r" REF SecOne \\w", None),
            (r' REF SecOne \# "0" \w', "SecOne"),
            (r" PAGEREF SecOne \w", None),
            (r" REF SecOne \h", None),
        ]
        for code, expected in cases:
            with self.subTest(code=code):
                self.assertEqual(word._numbered_reference_target(code), expected)

    def test_numbered_references_without_automatic_numbering_fail_before_update(self):
        for list_type, label in ((0, ""), (2, "•"), (6, "Picture"), (4, "")):
            with self.subTest(list_type=list_type, label=label):
                calls = []
                fields = [FakeField(r"REF SecOne " + switch, calls) for switch in (r"\w", r"\n", r"\r")]
                doc = FakeDocument(FakeStory(fields))
                doc.Bookmarks = FakeBookmarks({"SecOne": (list_type, label)})
                report = word.refresh_document_fields(doc)
                self.assertEqual(calls, [])
                self.assertEqual(report["updated"], [])
                self.assertEqual(len(report["errors"]), 3)
                self.assertTrue(all("SecOne" in item["error"] and "no automatic paragraph number" in item["error"]
                                    for item in report["errors"]))

    def test_automatic_zero_start_number_is_valid(self):
        calls = []
        doc = FakeDocument(FakeStory([FakeField(r"REF SecOne \w", calls)]))
        doc.Bookmarks = FakeBookmarks({"SecOne": (3, "\u200e0")})
        report = word.refresh_document_fields(doc)
        self.assertFalse(report["errors"])
        self.assertEqual(calls, [r"REF SecOne \w"])
        self.assertEqual(report["numbered_references"][0]["list_label"], "0")
        self.assertEqual(report["numbered_references"][0]["bookmark"], "SecOne")

    def test_locked_number_reference_skips_numbering_validation(self):
        calls = []
        doc = FakeDocument(FakeStory([FakeField(r"REF MissingTarget \w", calls, locked=True)]))
        report = word.refresh_document_fields(doc)
        self.assertFalse(report["errors"])
        self.assertEqual(calls, [])
        self.assertEqual(report["numbered_references"], [])
        self.assertEqual(len(report["skipped"]), 1)
        self.assertEqual(report["skipped"][0]["reason"], "field is locked")

    def test_number_then_reference_then_toc_then_page(self):
        calls = []
        fields = [FakeField(code, calls) for code in [
            "PAGEREF figA", 'TOC \\o "1-3"', "REF figA", "SEQ Figure",
        ]]
        doc = FakeDocument(FakeStory(fields))
        result = word.refresh_document_fields(doc)
        self.assertEqual(calls, ["SEQ Figure", "REF figA", 'TOC \\o "1-3"', "PAGEREF figA"])
        self.assertFalse(result["errors"])
        self.assertFalse(result["zotero_refreshed"])
        self.assertGreaterEqual(doc.repagination_count, 4)

    def test_toc_regenerates_entries_then_refreshes_new_page_references(self):
        calls = []
        toc = FakeField('TOC \\o "1-3"', calls)
        toc.Code.Start = 11
        toc.Result = type("Range", (), {"End": 38})()
        story = FakeStory([toc, FakeField("PAGEREF old_entry", calls)])
        doc = FakeDocument(story)

        class NativeToc:
            Range = type("Range", (), {"Start": 10, "End": 39})()

            def Update(self):
                calls.append("regenerate TOC entries")
                story.Fields.values[1:] = [FakeField("PAGEREF new_entry", calls)]

        doc.TablesOfContents = Collection([NativeToc()])
        result = word.refresh_document_fields(doc)
        self.assertEqual(calls, ["regenerate TOC entries", "PAGEREF new_entry"])
        self.assertFalse(result["errors"])

    def test_zotero_external_data_and_locked_fields_remain_untouched(self):
        calls = []
        fields = [FakeField(code, calls) for code in [
            'ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"abc"}',
            "ADDIN ZOTERO_BIBL {} CSL_BIBLIOGRAPHY", "INCLUDETEXT external.docx",
            "DDEAUTO server topic", "MACROBUTTON dangerous action", "ASK name",
        ]]
        fields += [FakeField("PAGE", calls, locked=True), FakeField("NUMPAGES", calls)]
        doc = FakeDocument(FakeStory(fields))
        result = word.refresh_document_fields(doc)
        self.assertEqual(calls, ["NUMPAGES"])
        self.assertEqual(len(result["skipped"]), 7)
        self.assertFalse(result["zotero_refreshed"])

    def test_word_false_update_is_an_error(self):
        calls = []
        doc = FakeDocument(FakeStory([FakeField("REF missing", calls, result=False)]))
        result = word.refresh_document_fields(doc)
        self.assertEqual(len(result["errors"]), 1)
        self.assertEqual(result["errors"][0]["kind"], "REF")


if __name__ == "__main__":
    unittest.main()
