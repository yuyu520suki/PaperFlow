import json
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import sys
import tempfile
import threading
import time
import unittest

from paperflow import gui, polish
from paperflow.service import Service, TERMINAL_STATUSES, validate_cli_args


def finish(service, job_id):
    deadline = time.monotonic() + 30
    while time.monotonic() < deadline:
        status = service.job_status(job_id)
        if status["status"] in TERMINAL_STATUSES and status["worker_alive"] is False:
            return status
        time.sleep(0.05)
    raise AssertionError(f"Job did not finish: {job_id}")


class ServiceTests(unittest.TestCase):
    def test_content_review_forms_use_shared_cli_and_explicit_adoption(self):
        values = {"source": "待审稿.md", "evidence": "实验记录.txt", "output": "审阅输出"}
        plan = validate_cli_args(gui.build_content_args(values))
        self.assertEqual((plan.command, plan.input, plan.evidence, plan.output_dir),
                         ("content-plan", "待审稿.md", "实验记录.txt", "审阅输出"))
        online = validate_cli_args(gui.build_content_args(values, execute=True))
        self.assertEqual((online.command, online.max_rounds), ("content-review", 2))
        apply = validate_cli_args(gui.build_content_apply_args("待审稿.md", "审阅计划.json", "修订稿.md", ["f01", "f03"]))
        self.assertEqual((apply.command, apply.output, apply.accept), ("content-apply", "修订稿.md", ["f01", "f03"]))
        with self.assertRaises(ValueError):
            gui.build_content_apply_args("待审稿.md", "审阅计划.json", "修订稿.md", [])

    def test_gui_forms_use_cli_and_readable_review_tokens_round_trip(self):
        forms = {"source": "原稿.docx", "template": "样例.docx", "output": "输出目录",
                 "metadata": "metadata.json", "journal": "Computers & Structures",
                 "submission_copy": True, "artwork_root": "图片", "table_policy": "three-line"}
        self.assertEqual(validate_cli_args(gui.build_layout_args(forms)).journal, "computers-and-structures")
        revised = validate_cli_args(gui.build_layout_args(forms, revise=True))
        self.assertEqual((revised.command, revised.artwork_root), ("revise-master", "图片"))
        self.assertEqual(validate_cli_args(gui.build_polish_args({**forms, "language": "中文"})).language, "zh")
        advanced = gui.parse_cli_args(json.dumps(["polish-run", "原稿.docx", "--output-dir", "新目录",
                                                  "--max-rounds", "1", "--voice-sample", "写作样例.txt"]))
        self.assertEqual((validate_cli_args(advanced).max_rounds, validate_cli_args(advanced).voice_sample),
                         (1, "写作样例.txt"))
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "原稿.txt"
            source.write_text("结果表明位移为 12.5 mm [1]。", encoding="utf-8")
            plan = polish.prepare_polish(source, root / "plan")
            paragraph = plan["paragraphs"][0]
            markers = gui._review_markers(paragraph)
            displayed = gui._review_text(paragraph["text"], markers)
            self.assertIn("12.5 mm", displayed)
            with self.assertRaises(polish.PolishError):
                gui._save_review_paragraph(paragraph, displayed.replace("12.5", "11.0"), True)
            gui._save_review_paragraph(paragraph, displayed.replace("结果表明", ""), True)
            report = root / "reviewed.json"
            report.write_text(json.dumps(plan, ensure_ascii=False), encoding="utf-8")
            result = polish.apply_polish(source, report, root / "润色.txt")
            self.assertEqual(Path(result["output"]).read_text(encoding="utf-8"), "位移为 12.5 mm [1]。")

    def test_job_argument_limits_and_cooperative_cancellation(self):
        with tempfile.TemporaryDirectory() as tmp:
            service = Service(Path(tmp))
            with self.assertRaises(ValueError):
                service.start_cli(["doctor"], {"PATH": "not-allowed"})
            self.assertEqual(service.list_jobs(), [])
            (Path(tmp) / "source.txt").write_text("An academic sentence.", encoding="utf-8")
            started = service.start_cli(["polish-plan", "source.txt", "--output-dir", "plan"])
            service.cancel_job(started["id"])
            result = finish(service, started["id"])
            self.assertEqual(result["status"], "cancelled", result)
            self.assertEqual(result["exit_code"], 130)
            self.assertTrue(result["cancel_requested"])
            self.assertEqual(service.cancel_job(started["id"])["status"], "cancelled")

    def test_real_worker_openai_compatible_loopback_and_independent_review(self):
        requests = []
        barrier = threading.Barrier(2)
        failures = []
        key = "synthetic-loopback-credential"

        class Provider(BaseHTTPRequestHandler):
            def log_message(self, *args):
                pass

            def do_POST(self):
                try:
                    request = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
                    prompt, user = request["messages"]
                    role = next(role for role, full_prompt in polish._PROMPTS.items() if full_prompt == prompt["content"])
                    payload = json.loads(user["content"])
                    requests.append((self.path, request["model"], self.headers["Authorization"] == "Bearer " + key, role))
                    paragraph = payload["paragraphs"][0]
                    proposal = paragraph["original"].replace("It is worth noting that the", "The")
                    if role.endswith("reviewer"):
                        barrier.wait(8)
                        body = {"reviews": [{"id": paragraph["id"], "approved": True, "issues": []}]}
                    elif role == "integrator":
                        if set(payload["independent_reviews"]) != {"factual_reviewer", "academic_reviewer"}:
                            raise ValueError("Missing independent reviews")
                        body = {"edits": [{"id": paragraph["id"], "text": proposal, "accepted": True}]}
                    else:
                        body = {"edits": [{"id": paragraph["id"], "text": proposal}]}
                    data = json.dumps({"choices": [{"message": {"content": json.dumps(body)}}]}).encode()
                    self.send_response(200)
                    self.send_header("Content-Length", str(len(data)))
                    self.end_headers()
                    self.wfile.write(data)
                except Exception as error:
                    failures.append(repr(error))
                    self.send_error(500)

        with tempfile.TemporaryDirectory() as tmp, ThreadingHTTPServer(("127.0.0.1", 0), Provider) as server:
            worker = threading.Thread(target=server.serve_forever, daemon=True)
            worker.start()
            try:
                root = Path(tmp)
                service = Service(root)
                source = root / "paper.txt"
                source.write_text("It is worth noting that the deformation was 12.5 mm [1].", encoding="utf-8")
                env = gui.provider_environment(f"http://127.0.0.1:{server.server_port}/v1", "fixture", key)
                submitted = service.start_cli(["polish-run", "paper.txt", "--output-dir", "edits", "--max-rounds", "1"], env)
                result = finish(service, submitted["id"])
                self.assertFalse(failures)
                self.assertEqual(result["status"], "succeeded", result["stderr"])
                self.assertEqual({record[3] for record in requests}, set(polish._PROMPTS))
                self.assertEqual(len(requests), 4)
                self.assertTrue(all(record[:3] == ("/v1/chat/completions", "fixture", True) for record in requests))
                output = result["result"]
                self.assertEqual(output["accepted_changes"], 1)
                self.assertEqual(Path(output["output"]).read_text(encoding="utf-8"), "The deformation was 12.5 mm [1].")
                self.assertTrue(json.loads(Path(output["review_log"]).read_text(encoding="utf-8"))["completed"])
                for file in root.rglob("*"):
                    if file.is_file():
                        self.assertNotIn(key.encode(), file.read_bytes(), file.name)
            finally:
                server.shutdown()
                worker.join(2)


if __name__ == "__main__":
    unittest.main()
