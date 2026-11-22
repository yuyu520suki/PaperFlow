"""Validate cross-chapter claims and preserve document objects during apply."""

import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from docx import Document

from paperflow import content_review as review
from paperflow import polish
from paperflow.package import read_package


SOURCE = """# 摘要

控制器让三种受测结构的峰值位移统一降至 2.0 mm，证实了台架减振效果。

# 仿真方法

在同一数值模型中仅比较低、中、高三种阻尼工况，未开展物理台架试验。

# 仿真结果

| 工况 | 峰值位移 | 位移均方根 |
|---|---|---|
| 低阻尼 | 9.0 mm | 3.0 mm |
| 中阻尼 | 6.0 mm | 2.0 mm |
| 高阻尼 | 4.0 mm | 1.0 mm |

# 讨论

该方法优于所有已发表控制器。
"""

EVIDENCE = [
    {"id": "run-conditions", "kind": "experiment", "locator": "run-log.md:14",
     "text": "This was a numerical simulation of the same structural model under three damping settings."},
    {"id": "known-paper", "kind": "literature", "locator": "Zotero example record",
     "title": "Literature on semi-active control", "doi": "10.9999/example-only", "zotero_key": "ABCD1234"},
    {"id": "excerpt-paper", "kind": "literature", "locator": "Supplied paper, section 4",
     "title": "Comparing damping configurations", "doi": "10.9999/example-excerpt",
     "text": "The particular setup compared the baseline and two configurations at the same load."},
]

REVIEWED = {"consistency": {"approved": True, "reason": "The corrected scope matches the experiment."},
            "literature": {"approved": True, "reason": "The correction makes no literature comparison."}}


def write_json(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def record(plan, text):
    return next(row for row in plan["paragraphs"] if text in row["text"])


def link(source, quote=None):
    return {"id": source["id"], "locator": source["locator"], "quote": source["text"] if quote is None else quote}


def finding(paragraph, evidence, *, key="claim", kind="contradicted", proposed=None, quote=None):
    return {"id": key, "classification": kind, "paragraph_id": paragraph["id"],
            "quote": paragraph["text"] if quote is None else quote, "issue": "Scope, metric or provenance needs checking",
            "evidence": evidence, "proposed_text": proposed, "reviews": REVIEWED, "accepted": False}


def correction(paragraph):
    token = next(item["token"] for item in paragraph["capsules"] if item["display"] == "2.0 mm")
    return f"在同一数值模型中，中阻尼工况的位移均方根为 {token}。"


class ContentReviewTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.source = self.root / "合成稿.md"
        self.source.write_text(SOURCE, encoding="utf-8", newline="\n")
        self.evidence = self.root / "evidence.json"
        write_json(self.evidence, {"evidence": EVIDENCE})

    def prepared(self):
        with patch.object(polish, "_chat", side_effect=AssertionError("Unexpected provider request")):
            result = review.prepare_content_review(self.source, self.root / "offline", evidence=self.evidence)
        self.assertFalse(result["remote_requests_performed"])
        self.assertEqual(result["finding_count"], 0)
        plan_path = Path(result["plan_path"])
        self.assertEqual(plan_path.name, "review-plan.json")
        self.assertEqual(Path(result["report_path"]).name, "report.json")
        return plan_path, read_json(plan_path)

    def test_portable_host_plan_anchors_visible_values_and_applies_only_selected_edit(self):
        path, plan = self.prepared()
        claim = record(plan, "峰值位移统一降至")
        method, results = record(plan, "未开展物理台架试验"), record(plan, "| 中阻尼 |")
        self.assertEqual((claim["chapter"], method["chapter"], results["chapter"]), ("摘要", "仿真方法", "仿真结果"))
        self.assertIn("2.0 mm", claim["text"])
        self.assertNotIn("PFLOCK", claim["text"])
        self.assertIn("PFLOCK", claim["protected_text"])
        self.assertEqual(results["role"], "table")
        self.assertFalse(results["editable"])
        self.assertEqual({row["id"] for row in plan["evidence"] if row["metadata_only"]}, {"known-paper"})
        plan["findings"] = [finding(claim, [link(method), link(results), link(plan["evidence"][0])], proposed=correction(claim))]
        write_json(path, plan)
        relocated = self.root / "relocated.md"
        relocated.write_bytes(self.source.read_bytes())
        check = review.validate_content_review(relocated, path)
        self.assertTrue(check["valid"], check["errors"])
        self.assertEqual((check["adoptable_count"], check["accepted_ids"]), (1, []))
        with self.assertRaisesRegex(review.ContentReviewError, "Select at least one"):
            review.apply_content_review(relocated, path, self.root / "unselected.md")
        destination = self.root / "reviewed.md"
        applied = review.apply_content_review(relocated, path, destination, accepted_ids=["claim"])
        self.assertEqual((applied["accepted_ids"], applied["accepted_changes"]), (["claim"], 1))
        self.assertTrue(applied["source_unchanged"])
        expected = SOURCE.replace(claim["text"], "在同一数值模型中，中阻尼工况的位移均方根为 2.0 mm。")
        self.assertEqual(destination.read_text(encoding="utf-8"), expected)
        self.assertEqual(self.source.read_text(encoding="utf-8"), SOURCE)

    def test_metadata_and_even_title_of_a_record_with_excerpt_are_report_only(self):
        path, plan = self.prepared()
        claim = record(plan, "优于所有")
        metadata, rich = plan["evidence"][1:]
        plan["findings"] = [
            finding(claim, [link(metadata, "")], key="meta", kind="supported"),
            finding(claim, [], key="none", kind="supported"),
            finding(claim, [link(rich, rich["title"])], key="title", proposed="现有资料不足以比较全部控制器。"),
            finding(claim, [link(rich)], key="excerpt", kind="citation_gap"),
        ]
        plan["findings"][2]["adoptable"] = True
        write_json(path, plan)
        check = review.validate_content_review(self.source, path)
        self.assertTrue(check["valid"], check["errors"])
        rows = {row["id"]: row for row in check["findings"]}
        self.assertEqual(rows["meta"]["classification"], "citation_gap")
        self.assertEqual(rows["none"]["classification"], "missing_evidence")
        self.assertTrue(rows["title"]["requires_source_text"])
        self.assertTrue(rows["title"]["evidence"][0]["metadata_only"])
        self.assertFalse(rows["excerpt"]["evidence"][0]["metadata_only"])
        self.assertEqual(check["adoptable_count"], 0)
        with self.assertRaisesRegex(review.ContentReviewError, "ineligible finding"):
            review.apply_content_review(self.source, path, self.root / "rejected.md", accepted_ids=["title"])

    def test_false_quotes_changed_sources_and_evidence_changes_are_detected(self):
        path, plan = self.prepared()
        claim = record(plan, "峰值位移统一降至")
        method = record(plan, "未开展物理台架试验")
        candidate = finding(claim, [link(method, "实际开展了物理台架试验。")], proposed=correction(claim))
        plan["findings"] = [candidate]
        write_json(path, plan)
        self.assertRegex(review.validate_content_review(self.source, path)["errors"][0], "evidence quote")
        candidate["evidence"] = [link(method)]
        candidate["quote"] = "从未出现在原稿中的表达"
        write_json(path, plan)
        self.assertRegex(review.validate_content_review(self.source, path)["errors"][0], "source quote")
        candidate["quote"] = claim["text"]
        plan["evidence"][0]["text"] += " Extra unverified evidence."
        write_json(path, plan)
        self.assertRegex(review.validate_content_review(self.source, path)["errors"][0], "evidence excerpts")
        self.source.write_text(SOURCE + "\n新增句子。", encoding="utf-8")
        self.assertRegex(review.validate_content_review(self.source, path)["errors"][0], "Source no longer matches")

    def test_numeric_changes_or_unapproved_candidate_cannot_be_selected(self):
        path, plan = self.prepared()
        claim, results = record(plan, "峰值位移统一降至"), record(plan, "| 中阻尼 |")
        broken = finding(claim, [link(results)], key="numeric", proposed="数值仿真的中阻尼位移为 6.0 mm。")
        no_review = finding(claim, [link(results)], key="only-one", proposed=correction(claim))
        no_review["reviews"] = {"consistency": REVIEWED["consistency"]}
        plan["findings"] = [broken, no_review]
        write_json(path, plan)
        check = review.validate_content_review(self.source, path)
        self.assertTrue(check["valid"])
        self.assertEqual(check["adoptable_count"], 0)
        self.assertIn("protected capsules", " ".join(check["findings"][0]["blocked_reasons"]))
        self.assertIn("literature review", " ".join(check["findings"][1]["blocked_reasons"]))
        no_review["accepted"] = True
        write_json(path, plan)
        self.assertFalse(review.validate_content_review(self.source, path)["valid"])

    def test_utf8_text_log_is_anchored_to_actual_line_ranges(self):
        evidence = self.root / "试验记录.txt"
        evidence.write_bytes("标题\r\n\r\n低阻尼峰值 9.0 mm。\r\n中阻尼峰值 6.0 mm。\r\n".encode("utf-8-sig"))
        result = review.prepare_content_review(self.source, self.root / "plain", evidence=evidence)
        plan = read_json(result["plan_path"])
        self.assertEqual([row["locator"] for row in plan["evidence"]], ["试验记录.txt:1", "试验记录.txt:3-4"])
        self.assertEqual(plan["evidence"][1]["text"], "低阻尼峰值 9.0 mm。\n中阻尼峰值 6.0 mm。")
        self.assertTrue(all(row["kind"] == "note" for row in plan["evidence"]))

    def test_word_heading_table_context_and_copy_preserve_package_parts(self):
        manuscript = self.root / "structure.docx"
        document = Document()
        document.add_heading("Results", level=1)
        document.add_paragraph().add_run("The measured RMS was 2.0 mm.").bold = True
        document.add_heading("Evidence", level=1)
        table = document.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "Simulation only"
        table.cell(0, 1).text = "RMS calculated from the model output"
        document.save(manuscript)
        result = review.prepare_content_review(manuscript, self.root / "word-plan")
        path = Path(result["plan_path"])
        plan = read_json(path)
        claim, cell = record(plan, "measured RMS"), record(plan, "Simulation only")
        self.assertEqual((claim["chapter"], cell["chapter"], cell["role"]), ("Results", "Evidence", "table"))
        self.assertIn("table 1, row 1, cell 1", cell["locator"])
        candidate = finding(claim, [link(cell)], proposed=claim["protected_text"].replace("measured", "simulated"))
        candidate["accepted"] = True
        plan["findings"] = [candidate]
        write_json(path, plan)
        final = self.root / "copy.docx"
        applied = review.apply_content_review(manuscript, path, final)
        self.assertEqual(applied["accepted_ids"], ["claim"])
        first, second = read_package(manuscript), read_package(final)
        self.assertEqual(set(first), set(second))
        self.assertEqual([key for key in first if first[key] != second[key]], ["word/document.xml"])
        self.assertTrue(Document(final).paragraphs[1].runs[0].bold)

    def test_missing_index_excerpt_preserves_source_with_explicit_fallback(self):
        records = [{"id": "p000001", "text": "The result was simulated."},
                   {"id": "e000001", "text": "There was no physical test."}]
        with patch.object(review, "CONTEXT_BUDGET", 10), patch.object(review, "_chat", side_effect=[
            {"anchors": [{"id": "p000001", "quote": "result was simulated"}]},
            {"anchors": [{"id": "e000001", "quote": "This quote does not exist"}]},
        ]):
            trace = []
            context = review._cross_context(records, ("endpoint", "fixture", "fake"), trace)
        self.assertEqual(context[0]["excerpts"], ["result was simulated"])
        self.assertEqual(context[1]["text"], records[1]["text"])
        self.assertEqual(trace[1]["full_text_fallback_ids"], ["e000001"])

    def test_online_analysis_uses_cross_chapter_data_and_retries_only_rejected_edit(self):
        for correct_second in (True, False):
            with self.subTest(correct_second=correct_second):
                calls = []
                unwanted = "控制器消除了三种工况下的全部振动"
                active_candidate = None

                def provider(role, _prompt, payload, _configuration):
                    nonlocal active_candidate
                    role = role.removeprefix("content_")
                    calls.append(role)
                    if role == "analyst":
                        self.assertEqual(len(payload["chapters"]), 4)
                        claim = next((item for item in payload["focus"] if "峰值位移统一降至" in item["text"]), None)
                        if claim is None:
                            return {"findings": []}
                        active_candidate = correction(claim)
                        self.assertIn("2.0 mm", claim["text"])
                        method = next(item for item in payload["document_context"] if "未开展物理台架试验" in item["text"])
                        results = next(item for item in payload["document_context"] if "| 中阻尼 |" in item["text"])
                        invalid = claim["protected_text"].replace("控制器让三种受测结构的峰值位移统一降至", unwanted)
                        return {"findings": [finding(claim, [link(method), link(results)], proposed=invalid)]}
                    if role == "refiner":
                        [rejected] = payload["rejected"]
                        self.assertFalse(rejected["reviews"]["consistency"]["approved"])
                        return {"revisions": [{"id": rejected["id"], "proposed_text": active_candidate if correct_second else "错误替代 8.0 mm。"}]}
                    reviews = []
                    for row in payload["findings"]:
                        self.assertTrue(all("text" in item for item in payload["cited_records"]))
                        accepted = role == "literature" or unwanted not in row["proposed_text"]
                        reviews.append({"id": row["id"], "approved": accepted,
                                        "reason": "Measured scope checked" if accepted else "No physical test"})
                    return {"reviews": reviews}

                with patch.object(polish, "_configuration", return_value=("endpoint", "fixture", "fake")), \
                        patch.object(polish, "_chat", side_effect=provider):
                    outcome = review.run_content_review(self.source, self.root / f"online-{correct_second}", evidence=self.evidence)
                self.assertEqual((outcome["rounds"], outcome["finding_count"], outcome["adoptable_count"]), (2, 1, int(correct_second)))
                self.assertTrue(outcome["remote_requests_performed"])
                self.assertEqual(calls.count("analyst"), 4)
                self.assertEqual(calls.count("refiner"), 1)
                self.assertEqual(calls.count("consistency"), 5)
                self.assertEqual(calls.count("literature"), 5)
                plan = read_json(outcome["plan_path"])
                self.assertTrue(all(not row["accepted"] for row in plan["findings"]))
                proposal = plan["findings"][0]["proposed_text"]
                self.assertEqual(proposal if correct_second else unwanted in proposal,
                                 active_candidate if correct_second else True)
                report = read_json(outcome["report_path"])
                first = next(item for item in report["review_trace"] if item["round"] == 1 and item["reviewed_findings"])
                self.assertFalse(first["reviewed_findings"][0]["reviews"]["consistency"]["approved"])
                self.assertFalse(any(path.suffix == ".md" for path in Path(outcome["report_path"]).parent.iterdir()))


if __name__ == "__main__":
    unittest.main()
