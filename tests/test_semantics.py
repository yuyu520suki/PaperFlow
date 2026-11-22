from io import BytesIO
import json
from pathlib import Path
import tempfile
import unittest

from docx import Document
from lxml import etree as ET
from PIL import Image

from paperflow.package import field_runs, part_fields, q, read_package, write_package, xml
from paperflow.semantics import apply_draft_plan, plan_draft


def _picture(document):
    image = BytesIO()
    Image.new("RGB", (16, 12), (74, 92, 117)).save(image, format="PNG")
    document.add_picture(image)


def _body(parts):
    return xml(parts["word/document.xml"]).find(q("w:body"))


def _visible(element):
    return "".join(text.text or "" for text in element.iter(q("w:t")))


class DraftSemanticsTests(unittest.TestCase):
    def test_prose_starting_with_a_reference_is_not_a_second_caption(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "draft.docx"
            document = Document()
            document.add_paragraph("Figure 1 shows the results. Prior work [1] covers the method.")
            _picture(document)
            document.add_paragraph("Figure 1. Actual caption")
            document.save(source)
            plan = plan_draft(source, citation_map={"1": "ABCD1234"})
            self.assertEqual([edit["operation"] for edit in plan["edits"]],
                             ["cross_reference", "zotero_marker", "caption"])
            self.assertEqual(plan["issues"], [])

    def test_builds_local_refs_and_retains_source_native_fields_and_parts(self):
        with tempfile.TemporaryDirectory() as tmp:
            source, output, intermediate = (Path(tmp) / name for name in ("source.docx", "output.docx", "draft.docx"))
            document = Document()
            body = document.add_paragraph("如图1所示，Table ")
            body.add_run("1").bold = True
            body.add_run(" describes the settings. Prior work [")
            body.add_run("1, 2–3] was also reviewed.")
            _picture(document)
            caption = document.add_paragraph("图 ")
            caption.add_run("1").bold = True
            caption.add_run("：位移时程").italic = True
            document.add_paragraph("Table 1. Test settings")
            table = document.add_table(rows=1, cols=1)
            table.cell(0, 0).text = "Sensor"
            existing = document.add_paragraph("Existing field and OLE: ")
            for element in field_runs(" PAGE ", "3"):
                existing._p.append(element)
            object_run = ET.SubElement(existing._p, q("w:r"))
            ole = ET.SubElement(ET.SubElement(object_run, q("w:object")), q("o:OLEObject"))
            ole.set("ProgID", "Equation.DSMT4")
            document.add_heading("References", level=1)
            document.add_paragraph("[1] Existing bibliography entry")
            document.save(intermediate)
            original = read_package(intermediate)
            original["word/embeddings/oleObject1.bin"] = b"opaque embedded equation bytes"
            write_package(original, source)
            before = source.read_bytes()

            plan = plan_draft(source, citation_map={"1": "ABCD1234", "2": "EFGH5678", "3": "JKLM9012"})
            result = apply_draft_plan(source, output, json.loads(json.dumps(plan)))
            self.assertEqual(source.read_bytes(), before)
            self.assertEqual(result["captions"], 2)
            self.assertEqual(result["cross_references"], 2)
            self.assertEqual(result["citation_markers"], 1)
            self.assertEqual(result["issues"], [])
            parts = read_package(output)
            self.assertEqual({name for name in parts if parts[name] != original[name]}, {"word/document.xml"})
            content = _body(parts)
            fields, errors = part_fields(content)
            self.assertEqual(errors, [])
            self.assertIn(" SEQ Figure \\* ARABIC ", fields)
            self.assertIn(" SEQ Table \\* ARABIC ", fields)
            self.assertIn(" REF PF_Figure_3 \\h ", fields)
            self.assertIn(" REF PF_Table_4 \\h ", fields)
            self.assertIn(" PAGE ", fields)
            self.assertIn("[[CITE ABCD1234; EFGH5678; JKLM9012]]", _visible(content[0]))
            self.assertEqual(_visible(content[2]), "图 1：位移时程")
            self.assertIsNotNone(content[2].find(f".//{q('w:i')}"))
            self.assertEqual(ET.tostring(content[5], method="c14n"), ET.tostring(_body(original)[5], method="c14n"))
            self.assertEqual(_visible(content[7]), "[1] Existing bibliography entry")

    def test_duplicate_or_missing_caption_numbers_leave_references_for_review(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "draft.docx"
            document = Document()
            document.add_paragraph("Figures 1 and 2 are discussed. Figure 1 appears first. Table 3 is later.")
            _picture(document)
            document.add_paragraph("Figure 1. First caption")
            document.add_paragraph("Figure 1. Second caption")
            _picture(document)
            document.add_paragraph("Table 3. Missing earlier table")
            document.add_table(rows=1, cols=1)
            document.save(source)
            plan = plan_draft(source)
            self.assertEqual(plan["edits"], [])
            self.assertEqual(
                {issue["code"] for issue in plan["issues"]},
                {"ambiguous_caption", "caption_numbering_gap", "ambiguous_cross_reference", "unresolved_cross_reference"},
            )

    def test_chapter_captions_bind_to_headings_and_convert_equivalent_separators(self):
        with tempfile.TemporaryDirectory() as tmp:
            source, output = (Path(tmp) / name for name in ("draft.docx", "output.docx"))
            document = Document()
            document.add_heading("First chapter", level=1)
            _picture(document)
            document.add_paragraph("图 1–1：时程")
            document.add_paragraph("结果见图1.1。")
            document.add_heading("Second chapter", level=1)
            document.add_paragraph("Table 2.1: Settings")
            document.add_table(rows=1, cols=1)
            document.add_paragraph("See Table 2-1.")
            document.save(source)
            self.assertEqual(plan_draft(source)["edits"], [])
            plan = plan_draft(source, caption_chapter_level=1, caption_chapter_separator="-")
            self.assertEqual(plan["issues"], [])
            self.assertEqual(plan["chapter_caption_policy"], {"level": 1, "separator": "-"})
            captions = [edit for edit in plan["edits"] if edit["operation"] == "caption"]
            self.assertEqual([edit["chapter_heading_body_index"] for edit in captions], [0, 4])
            self.assertEqual([edit["target_number"] for edit in captions], ["1-1", "2-1"])
            report = apply_draft_plan(source, output, json.loads(json.dumps(plan)))
            self.assertEqual(report["captions"], 2)
            self.assertEqual(report["cross_references"], 2)
            self.assertEqual([entry["references"] for entry in report["chapter_captions"]], [1, 1])
            body = _body(read_package(output))
            self.assertEqual(_visible(body[2]), "图 1-1：时程")
            self.assertEqual(_visible(body[3]), "结果见图1-1。")
            self.assertEqual(part_fields(body[2]), ([r" STYLEREF 1 \s ", r" SEQ Figure \* ARABIC \s 1 "], []))
            self.assertEqual(part_fields(body[5]), ([r" STYLEREF 1 \s ", r" SEQ Table \* ARABIC \s 1 "], []))

    def test_chapter_plan_keeps_ranges_gaps_and_conflicting_prefixes_for_review(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "draft.docx"
            document = Document()
            _picture(document)
            document.add_paragraph("Figure 9-1. No chapter")
            document.add_heading("First chapter", level=1)
            for number in ("1-1", "1-2", "1-4"):
                _picture(document)
                document.add_paragraph(f"Figure {number}. Example")
                document.add_paragraph("Separation between objects.")
            document.add_paragraph("See Figure 1-1 through Figure 1-2.")
            document.add_paragraph("See Figure 1-1 – 1-2.")
            document.add_paragraph("见1-1至图1-2。")
            document.add_heading("Second chapter", level=1)
            document.add_paragraph("Table 2-1: First")
            document.add_table(rows=1, cols=1)
            document.add_paragraph("Separation between objects.")
            document.add_paragraph("Table 3-1: Conflicting chapter")
            document.add_table(rows=1, cols=1)
            document.save(source)
            plan = plan_draft(source, caption_chapter_level=1)
            self.assertEqual([(edit["operation"], edit["source_text"]) for edit in plan["edits"]],
                             [("caption", "1-1"), ("caption", "1-2")])
            codes = [issue["code"] for issue in plan["issues"]]
            self.assertIn("caption_without_chapter", codes)
            self.assertIn("caption_numbering_gap", codes)
            self.assertEqual(codes.count("ambiguous_chapter_prefix"), 2)
            self.assertEqual(codes.count("ambiguous_cross_reference"), 4)

    def test_missing_citation_map_and_stale_or_edited_plan_do_not_publish_edits(self):
        with tempfile.TemporaryDirectory() as tmp:
            source, output = Path(tmp) / "draft.docx", Path(tmp) / "out.docx"
            document = Document()
            document.add_paragraph("Prior work [1] established the approach. Invalid range [1-2-3].")
            document.add_heading("References", level=1)
            document.add_paragraph("[1] Original reference")
            document.save(source)
            plan = plan_draft(source)
            self.assertEqual(plan["edits"], [])
            self.assertEqual({issue["code"] for issue in plan["issues"]}, {"unmapped_citation", "unsupported_citation_numbers"})
            edited = plan_draft(source, citation_map={"1": "ABCD1234"})
            edited["edits"][0]["item_keys"] = ["EFGH5678"]
            with self.assertRaisesRegex(ValueError, "differs from the validated proposal"):
                apply_draft_plan(source, output, edited)
            self.assertFalse(output.exists())
            document.add_paragraph("A saved revision")
            document.save(source)
            with self.assertRaisesRegex(ValueError, "SHA-256"):
                apply_draft_plan(source, output, plan)
            self.assertFalse(output.exists())

    def test_external_qualification_and_complex_runs_stay_literal(self):
        with tempfile.TemporaryDirectory() as tmp:
            source, output = Path(tmp) / "draft.docx", Path(tmp) / "out.docx"
            document = Document()
            document.add_paragraph("Figure 1 from another study needs a reviewer.")
            document.add_paragraph("Local Figure 1(a) is shown.")
            complex_text = document.add_paragraph("A footnote and Figure 1 appear together.")
            ET.SubElement(complex_text.add_run()._r, q("w:footnoteReference"), {q("w:id"): "2"})
            _picture(document)
            document.add_paragraph("Figure 1. Local")
            document.save(source)
            plan = plan_draft(source)
            result = apply_draft_plan(source, output, plan)
            self.assertEqual(result["captions"], 1)
            self.assertEqual(result["cross_references"], 1)
            self.assertEqual({issue["code"] for issue in result["issues"]}, {"ambiguous_cross_reference", "unsafe_text_span"})
            before, after = _body(read_package(source)), _body(read_package(output))
            for index in (0, 2):
                self.assertEqual(ET.tostring(before[index], method="c14n"), ET.tostring(after[index], method="c14n"))
            self.assertEqual(_visible(after[1]), "Local Figure 1(a) is shown.")


if __name__ == "__main__":
    unittest.main()
