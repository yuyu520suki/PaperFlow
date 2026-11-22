"""Polishing edits prose while preserving scientific and Word object identity."""
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import urllib.error

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from lxml import etree as ET

from paperflow import polish
from paperflow.package import NS, q, read_package, write_package, xml


def write_plan(plan, path):
    Path(path).write_text(json.dumps(plan, ensure_ascii=False), encoding="utf-8")


def accept(plan, old, new):
    changed = 0
    for paragraph in plan["paragraphs"]:
        if old in paragraph["text"] and paragraph["editable"]:
            paragraph["proposed_text"] = paragraph["text"].replace(old, new)
            paragraph["accepted"] = True
            changed += 1
    return changed


def fixture(path):
    document = Document()
    custom = document.styles.add_style("ChapterNative", WD_STYLE_TYPE.PARAGRAPH)
    custom.base_style = document.styles["Heading 1"]
    document.add_paragraph("Protected heading title", style=custom)
    paragraph = document.add_paragraph()
    paragraph.add_run("It is worth noting that ")
    paragraph.add_run("the measured response").bold = True
    paragraph.add_run(" was 12.5 mm; see ")
    for kind, value in (("begin", None), (None, ' ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"stable"} '),
                        ("separate", None), (None, "[1]"), ("end", None)):
        run = ET.SubElement(paragraph._p, q("w:r"))
        if kind:
            ET.SubElement(run, q("w:fldChar"), {q("w:fldCharType"): kind})
        else:
            ET.SubElement(run, q("w:instrText") if value.startswith(" ADDIN") else q("w:t")).text = value
    paragraph.add_run(" for the observed result.")
    math = ET.SubElement(paragraph._p, q("m:oMath"))
    ET.SubElement(ET.SubElement(math, q("m:r")), q("m:t")).text = "x+1=2"
    run = ET.SubElement(paragraph._p, q("w:r"))
    obj = ET.SubElement(run, q("w:object"))
    ET.SubElement(obj, q("o:OLEObject"), ProgID="Equation.DSMT4")
    table = document.add_table(rows=1, cols=2)
    table.cell(0, 0).text = "The table measurement is reliable."
    table.cell(0, 1).text = "42"
    document.save(path)
    parts = read_package(path)
    parts["word/embeddings/oleObject1.bin"] = b"opaque native MathType identity"
    parts["word/footnotes.xml"] = b"opaque unchanged footnote story"
    replacement = Path(path).with_name("source-with-opaque.docx")
    write_package(parts, replacement)
    return replacement


class PolishingTests(unittest.TestCase):
    def test_docx_edits_adjacent_prose_and_preserves_fields_equations_runs_and_other_parts(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = fixture(root / "source.docx")
            before = source.read_bytes()
            plan = polish.prepare_polish(source, root / "plan")
            self.assertEqual(accept(plan, "It is worth noting that ", ""), 1)
            self.assertEqual(accept(plan, "The table measurement is reliable.", "The table measurement is consistent."), 1)
            write_plan(plan, plan["plan_path"])
            result = polish.apply_polish(source, plan["plan_path"], root / "polished.docx")
            original, saved = read_package(source), read_package(result["output"])
            self.assertEqual(source.read_bytes(), before)
            self.assertEqual(result["accepted_changes"], 2)
            self.assertEqual({key: value for key, value in original.items() if key != "word/document.xml"},
                             {key: value for key, value in saved.items() if key != "word/document.xml"})
            first, second = (xml(parts["word/document.xml"]) for parts in (original, saved))
            for path in (".//w:fldChar", ".//w:instrText", ".//m:oMath", ".//w:object", ".//w:rPr"):
                self.assertEqual([ET.tostring(node) for node in first.findall(path, NS)],
                                 [ET.tostring(node) for node in second.findall(path, NS)])
            self.assertIn("[1]", "".join(second.itertext()))
            self.assertIn("12.5 mm", "".join(second.itertext()))
            self.assertIn("Protected heading title", "".join(second.itertext()))
            self.assertNotIn("It is worth noting that", "".join(second.itertext()))
            self.assertFalse(plan["paragraphs"][0]["editable"])
            self.assertFalse(plan["paragraphs"][-1]["editable"])

    def test_protected_capsule_number_and_citation_changes_abort_without_output(self):
        for change in ("remove_capsule", "new_number", "new_citation"):
            with self.subTest(change=change), tempfile.TemporaryDirectory() as temporary:
                root = Path(temporary)
                source = fixture(root / "source.docx")
                before = source.read_bytes()
                plan = polish.prepare_polish(source, root / "plan")
                paragraph = next(item for item in plan["paragraphs"] if item["editable"] and item["capsules"])
                paragraph["accepted"] = True
                if change == "remove_capsule":
                    paragraph["proposed_text"] = paragraph["text"].replace(paragraph["capsules"][0]["token"], "")
                elif change == "new_number":
                    paragraph["proposed_text"] += " We observed 100 trials."
                else:
                    paragraph["proposed_text"] += " See [9]."
                write_plan(plan, plan["plan_path"])
                with self.assertRaises(polish.PolishError):
                    polish.apply_polish(source, plan["plan_path"], root / "out.docx")
                self.assertFalse((root / "out.docx").exists())
                self.assertEqual(source.read_bytes(), before)

    def test_source_or_immutable_plan_changes_are_rejected(self):
        for changed in ("source", "plan"):
            with self.subTest(changed=changed), tempfile.TemporaryDirectory() as temporary:
                root = Path(temporary)
                source = root / "source.txt"
                source.write_text("This method is clear.", encoding="utf-8")
                plan = polish.prepare_polish(source, root / "plan")
                accept(plan, "clear", "concise")
                if changed == "source":
                    source.write_text("The author changed this text.", encoding="utf-8")
                else:
                    plan["paragraphs"][0]["text"] = "Tampered source evidence."
                write_plan(plan, plan["plan_path"])
                with self.assertRaises(polish.PolishError):
                    polish.apply_polish(source, plan["plan_path"], root / "out.txt")
                self.assertFalse((root / "out.txt").exists())

    def test_markdown_metadata_code_tables_long_links_and_math_remain_exact(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "paper.md"
            ticks = chr(96)
            link = "https://example.org/" + "a" * 300
            text = ('---\ntitle: "Unchanged title"\n---\n\n# Heading\n\n'
                    f'{ticks * 4}python\nvalue = "This is code"\n{ticks * 3}\ncode_still_here = 42\n{ticks * 4}\n\n'
                    'Case | Result\n--- | ---\nA | 12.5\n\n'
                    f'It is worth noting that the measured value is 12.5 mm [@source] with $x+1=2$ at {link}.\n')
            source.write_text(text, encoding="utf-8")
            plan = polish.prepare_polish(source, root / "plan")
            self.assertEqual(accept(plan, "It is worth noting that ", ""), 1)
            write_plan(plan, plan["plan_path"])
            result = polish.apply_polish(source, plan["plan_path"], root / "out.md")
            self.assertEqual(Path(result["output"]).read_text(encoding="utf-8"), text.replace("It is worth noting that ", ""))
            self.assertEqual(source.read_text(encoding="utf-8"), text)

    def test_text_bom_and_windows_line_endings_are_retained(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "source.txt"
            text = "It is worth noting that the response\r\nwas measured.\r\n\r\nThe value is 12.5 mm.\r\n"
            source.write_bytes(b"\xef\xbb\xbf" + text.encode())
            plan = polish.prepare_polish(source, root / "plan")
            accept(plan, "It is worth noting that ", "")
            write_plan(plan, plan["plan_path"])
            polish.apply_polish(source, plan["plan_path"], root / "out.txt")
            self.assertEqual((root / "out.txt").read_bytes(), b"\xef\xbb\xbf" + text.replace("It is worth noting that ", "").encode())

    def test_manual_negation_change_is_explicitly_flagged_for_review(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "source.txt"
            source.write_text("The method does not converge.", encoding="utf-8")
            plan = polish.prepare_polish(source, root / "plan")
            accept(plan, "does not converge", "fails to converge")
            write_plan(plan, plan["plan_path"])
            result = polish.apply_polish(source, plan["plan_path"], root / "out.txt")
            self.assertTrue(result["review_required"])
            self.assertIn("negation", result["warnings"][0])

    def test_existing_output_and_source_overwrite_are_refused(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "source.txt"
            source.write_text("The source text.", encoding="utf-8")
            plan = polish.prepare_polish(source, root / "plan")
            for target in (source, root / "existing.txt"):
                target.write_text("Existing text.", encoding="utf-8")
                with self.assertRaises(polish.PolishError):
                    polish.apply_polish(source, plan["plan_path"], target)
                self.assertEqual(target.read_text(encoding="utf-8"), "Existing text.")

    def test_offline_preparation_and_application_need_no_provider_configuration(self):
        with tempfile.TemporaryDirectory() as temporary, patch.dict("os.environ", {}, clear=True), \
                patch("paperflow.polish._chat", side_effect=AssertionError("No request is authorized")):
            root = Path(temporary)
            source = root / "source.txt"
            source.write_text("The source prose.", encoding="utf-8")
            plan = polish.prepare_polish(source, root / "plan", voice_sample="Use concise formal prose.")
            accept(plan, "source prose", "original prose")
            write_plan(plan, plan["plan_path"])
            result = polish.apply_polish(source, plan["plan_path"], root / "out.txt")
            self.assertFalse(result["remote_requests_performed"])
            with self.assertRaisesRegex(polish.PolishError, "Configure"):
                polish.run_polish(source, root / "online")

    def test_online_workflow_uses_independent_reviews_and_requires_both_approvals(self):
        for reviewer_approves in (True, False):
            with self.subTest(approved=reviewer_approves), tempfile.TemporaryDirectory() as temporary:
                root = Path(temporary)
                source = root / "source.txt"
                source.write_text("It is worth noting that the method is stable.", encoding="utf-8")
                calls = []

                def model(role, prompt, payload, configuration):
                    calls.append(role)
                    paragraph_id = payload["paragraphs"][0]["id"]
                    text = payload["paragraphs"][0]["original"].replace("It is worth noting that ", "")
                    if role == "editor":
                        return {"edits": [{"id": paragraph_id, "text": text, "reason": "Remove filler"}]}
                    if role in {"factual_reviewer", "academic_reviewer"}:
                        approved = reviewer_approves or role != "factual_reviewer"
                        return {"reviews": [{"id": paragraph_id, "approved": approved, "issues": [] if approved else ["Unsupported change"]}]}
                    return {"edits": [{"id": paragraph_id, "text": text, "accepted": True}]}

                with patch("paperflow.polish._configuration", return_value=("http://127.0.0.1/v1/chat/completions", "mock-model", "not-a-real-key")), \
                        patch("paperflow.polish._chat", side_effect=model):
                    result = polish.run_polish(source, root / "run", max_rounds=1)
                self.assertEqual(set(calls), set(polish._PROMPTS))
                self.assertEqual(result["accepted_changes"], int(reviewer_approves))
                expected = "the method is stable." if reviewer_approves else source.read_text(encoding="utf-8")
                self.assertEqual(Path(result["output"]).read_text(encoding="utf-8"), expected)
                self.assertTrue(Path(result["review_log"]).is_file())
                self.assertEqual(result["rounds"], 1)

    def test_editor_revises_rejected_candidate_using_independent_feedback(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "source.txt"
            source.write_text("The method may be stable.", encoding="utf-8")
            rounds_seen = []

            def model(role, prompt, payload, configuration):
                para = payload["paragraphs"][0]
                key = para["id"]
                if role == "editor":
                    previous = payload.get("previous_feedback", {}).get(key)
                    rounds_seen.append(previous)
                    if previous:
                        self.assertEqual(previous["reviews"]["factual_reviewer"]["issues"], ["Restore uncertainty"])
                    text = "The method appears stable." if previous else "The method is stable."
                    return {"edits": [{"id": key, "text": text}]}
                self.assertNotIn("previous_feedback", payload)
                proposal = payload["proposals"][0]["text"]
                if role.endswith("reviewer"):
                    rejected = role == "factual_reviewer" and proposal == "The method is stable."
                    return {"reviews": [{"id": key, "approved": not rejected,
                                          "issues": ["Restore uncertainty"] if rejected else []}]}
                return {"edits": [{"id": key, "text": proposal, "accepted": True}]}

            with patch.object(polish, "_configuration", return_value=("endpoint", "fixture", "fake")), \
                    patch.object(polish, "_chat", side_effect=model):
                result = polish.run_polish(source, root / "run")
            self.assertEqual(result["rounds"], 2)
            self.assertEqual(result["accepted_changes"], 1)
            self.assertIsNone(rounds_seen[0])
            self.assertEqual(Path(result["output"]).read_text(encoding="utf-8"), "The method appears stable.")

    def test_protected_only_source_does_not_claim_any_model_requests(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "source.md"
            source.write_text("# Locked heading\n", encoding="utf-8")
            with patch.object(polish, "_configuration", return_value=("endpoint", "fixture", "fake")), \
                    patch.object(polish, "_chat", side_effect=AssertionError("Unexpected model request")):
                result = polish.run_polish(source, root / "run")
            self.assertFalse(result["remote_requests_performed"])
            self.assertEqual(result["rounds"], 0)
            self.assertEqual(source.read_bytes(), Path(result["output"]).read_bytes())

    def test_online_numeric_corruption_stops_before_review_or_output(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "source.txt"
            source.write_text("The method is stable.", encoding="utf-8")
            with patch("paperflow.polish._configuration", return_value=("http://127.0.0.1/v1/chat/completions", "mock-model", "secret")), \
                    patch("paperflow.polish._chat", return_value={"edits": [{"id": "p000001", "text": "The method is stable in 100 trials."}]}):
                with self.assertRaises(polish.PolishError):
                    polish.run_polish(source, root / "run", max_rounds=1)
            self.assertFalse((root / "run/polished.txt").exists())
            log = json.loads((root / "run/review-log.json").read_text(encoding="utf-8"))
            self.assertFalse(log["completed"])
            self.assertTrue(log["rounds"][0]["batches"][0]["validation_errors"])

    def test_provider_errors_and_credential_echo_do_not_log_response_secrets(self):
        configuration = ("https://example.invalid/v1/chat/completions", "mock", "secret-never-log-this")
        error = urllib.error.HTTPError(configuration[0], 401, "Unauthorized", {}, io.BytesIO(configuration[2].encode()))
        with patch("urllib.request.build_opener") as opener:
            opener.return_value.open.side_effect = error
            with self.assertRaises(polish.PolishError) as caught:
                polish._chat("editor", "prompt", {}, configuration)
            self.assertNotIn(configuration[2], str(caught.exception))
        response = io.BytesIO(json.dumps({"choices": [{"message": {"content": '{"key":"secret-never-log-this"}'}}]}).encode())
        with patch("urllib.request.build_opener") as opener:
            opener.return_value.open.return_value.__enter__.return_value = response
            with self.assertRaisesRegex(polish.PolishError, "credential"):
                polish._chat("editor", "prompt", {}, configuration)

    def test_round_limit_and_malformed_review_contract(self):
        with self.assertRaises(polish.PolishError):
            polish.run_polish("unused.txt", "unused", max_rounds=3)
        for response in ({"reviews": [{"id": "p1", "approved": True}]},
                         {"reviews": [{"id": "p1", "approved": "true", "issues": []}]},
                         {"reviews": [{"id": "p1", "approved": True, "issues": "none"}]}):
            with self.subTest(response=response), self.assertRaises(polish.PolishError):
                polish._indexed(response, "reviews", {"p1"})


if __name__ == "__main__":
    unittest.main()
