from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path

from . import __version__
from .ingest import import_document, find_pandoc
from .package import CHAPTER_CAPTION_SEPARATORS, inspect_docx, audit_docx, expand_markers


def emit(data, output=None):
    text = json.dumps(data, ensure_ascii=False, indent=2)
    if output:
        target = Path(output).resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open("x", encoding="utf-8") as f:
            f.write(text + "\n")
    else:
        print(text)


def _read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig")) if path else None


def build_parser():
    parser = argparse.ArgumentParser(prog="paperflow")
    parser.add_argument("--version", action="version", version=__version__)
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("doctor", help="Read-only local dependency report")
    sub.add_parser("gui", help="Open the local manual desktop interface")
    sub.add_parser("mcp", help="Serve all workflow capabilities through MCP stdio")
    for name, help_text in (
        ("content-plan", "Prepare a manuscript and evidence review plan offline"),
        ("content-review", "Run evidence-grounded multi-agent review of a manuscript"),
    ):
        content = sub.add_parser(name, help=help_text)
        content.add_argument("input", help="Manuscript to review")
        content.add_argument("--output-dir", required=True, help="Write the review artifacts to a new directory")
        content.add_argument("--evidence", help="Evidence file in TXT, Markdown, or JSON")
        if name == "content-review":
            content.add_argument("--max-rounds", type=int, choices=[1, 2], default=2)
    content_check = sub.add_parser("content-check", help="Validate a content review plan against its source manuscript")
    content_check.add_argument("input")
    content_check.add_argument("plan")
    content_apply = sub.add_parser("content-apply", help="Apply adopted evidence-grounded changes to a new manuscript")
    content_apply.add_argument("input")
    content_apply.add_argument("plan")
    content_apply.add_argument("output")
    content_apply.add_argument("--accept", action="append", metavar="ID",
                               help="Adopt this suggestion ID (repeatable); when omitted, use accepted choices recorded in the plan")
    for name in ("polish-plan", "polish-run"):
        polish = sub.add_parser(name, help="Prepare or run reviewable multi-agent academic polishing")
        polish.add_argument("input")
        polish.add_argument("--output-dir", required=True)
        polish.add_argument("--language", choices=["auto", "zh", "en"], default="auto")
        polish.add_argument("--voice-sample", help="UTF-8 text file containing the author's preferred prose style")
        if name == "polish-run":
            polish.add_argument("--max-rounds", type=int, choices=[1, 2], default=2)
    polish_apply = sub.add_parser("polish-apply", help="Apply accepted edits while preserving protected document objects")
    polish_apply.add_argument("input"); polish_apply.add_argument("plan"); polish_apply.add_argument("output")
    zotero_find = sub.add_parser("zotero-find", help="Find item keys in the local Zotero library by title or author")
    zotero_find.add_argument("query")
    zotero_find.add_argument("--limit", type=int, default=10)
    citation_plan = sub.add_parser("plan-citations", help="Review numbered references against exact local Zotero DOI matches")
    citation_plan.add_argument("input")
    citation_plan.add_argument("--report", required=True)
    citation_apply = sub.add_parser("apply-citations", help="Apply reviewed Zotero links and one bibliography to a new DOCX")
    citation_apply.add_argument("input"); citation_apply.add_argument("plan"); citation_apply.add_argument("output")
    citation_apply.add_argument("--report")
    citekey_plan = sub.add_parser("plan-citekeys", help="Review Markdown or LaTeX citekeys against BibTeX and local Zotero")
    citekey_plan.add_argument("input")
    citekey_plan.add_argument("--bibliography", required=True)
    citekey_plan.add_argument("--report", required=True)
    for name in ["inspect-template", "audit"]:
        p = sub.add_parser(name)
        p.add_argument("input")
        p.add_argument("--report")
    imp = sub.add_parser("import")
    imp.add_argument("input"); imp.add_argument("output")
    imp.add_argument("--template"); imp.add_argument("--report")
    imp.add_argument("--ocr-mode",choices=["auto","always","never"],default="auto")
    imp.add_argument("--ocr-language")
    recovery=sub.add_parser("recovery-plan", help="Create reviewable text, table and heading proposals from TXT, PDF or a scan")
    recovery.add_argument("input")
    recovery.add_argument("--output-dir", required=True)
    recovery.add_argument("--ocr-mode", choices=["auto", "always", "never"], default="auto")
    recovery.add_argument("--ocr-language")
    recover=sub.add_parser("apply-recovery", help="Build a structured DOCX from a reviewed recovery plan")
    recover.add_argument("intake"); recover.add_argument("plan"); recover.add_argument("output")
    recover.add_argument("--report")
    expand = sub.add_parser("expand")
    expand.add_argument("input"); expand.add_argument("output")
    expand.add_argument("--csl-style", default="http://www.zotero.org/styles/ieee")
    expand.add_argument("--report")
    build = sub.add_parser("build")
    build.add_argument("input"); build.add_argument("--template", required=True)
    build.add_argument("--output", required=True)
    build.add_argument("--slot", default="{{PAPERFLOW_BODY}}")
    build.add_argument("--style-map", help="JSON map of source style names to template style names")
    build.add_argument("--report")
    refresh = sub.add_parser("refresh")
    refresh.add_argument("input"); refresh.add_argument("output")
    export = sub.add_parser("export-pdf")
    export.add_argument("input"); export.add_argument("output")
    run = sub.add_parser("run", help="Run the native workflow into a new output folder")
    run.add_argument("input"); run.add_argument("--template", required=True)
    run.add_argument("--output-dir", required=True)
    run.add_argument("--csl-style")
    run.add_argument("--metadata", help="JSON containing abstract, keywords, optional highlights and figure sources")
    run.add_argument("--journal",choices=["computers-and-structures"])
    slots=run.add_mutually_exclusive_group()
    slots.add_argument("--slot", default="{{PAPERFLOW_BODY}}")
    slots.add_argument("--layout-plan", help="Assign reviewed DOCX ranges to several template positions")
    run.add_argument("--table-policy",choices=["preserve","three-line"],default="preserve")
    run.add_argument("--table-style", help="Apply the specified Word table style ID from the target template")
    run.add_argument("--role-style-map",help="JSON mapping roles (e.g. heading1) to destination Word style IDs")
    run.add_argument("--caption-chapter-level", type=int, choices=range(1, 10),
                     help="Prefix generated caption numbers with the selected numbered Word Heading level")
    run.add_argument("--caption-chapter-separator", choices=CHAPTER_CAPTION_SEPARATORS,
                     help="Separator between the chapter and sequence numbers (default: -)")
    run.add_argument("--draft-plan",help="Apply a reviewed plan-draft report to its original DOCX")
    run.add_argument("--citekey-plan", help="Apply a reviewed plan-citekeys report to its original Markdown or LaTeX")
    run.add_argument("--submission-copy",action="store_true")
    run.add_argument("--ocr-mode",choices=["auto","always","never"],default="auto")
    run.add_argument("--ocr-language")
    revise=sub.add_parser("revise-master",help="Refresh a Word-edited DOCX without rebuilding it")
    revise.add_argument("input"); revise.add_argument("--output-dir",required=True)
    revise.add_argument("--csl-style"); revise.add_argument("--journal",choices=["computers-and-structures"])
    revise.add_argument("--submission-copy",action="store_true")
    revise.add_argument("--metadata",help="Reuse the submission metadata and separate figure declarations")
    revise.add_argument("--artwork-root",help="Directory for paths in figure_sources, usually the original manuscript folder")
    plan=sub.add_parser("plan-draft",help="Propose dynamic captions, cross-references and explicitly mapped citations")
    plan.add_argument("input"); plan.add_argument("--citation-map",help="JSON mapping numeric citations to Zotero keys")
    plan.add_argument("--caption-chapter-level", type=int, choices=range(1, 10),
                      help="Review legacy chapter captions against the selected source heading level")
    plan.add_argument("--caption-chapter-separator", choices=CHAPTER_CAPTION_SEPARATORS, default="-",
                      help="Separator for converted chapter caption numbers (default: -)")
    plan.add_argument("--report",required=True)
    apply_plan=sub.add_parser("apply-draft-plan",help="Apply a reviewed plan to a new DOCX copy")
    apply_plan.add_argument("input"); apply_plan.add_argument("plan"); apply_plan.add_argument("output")
    apply_plan.add_argument("--report")
    inventory=sub.add_parser("template-slots",help="Inspect sections and candidate insertion points")
    inventory.add_argument("input"); inventory.add_argument("--report")
    layout=sub.add_parser("inspect-layout",help="List manuscript body blocks and matching template positions")
    layout.add_argument("input"); layout.add_argument("--template", required=True)
    layout.add_argument("--report"); layout.add_argument("--plan-output")
    prepare=sub.add_parser("prepare-template",help="Add a marker to a separate copy at a reviewed paragraph anchor")
    prepare.add_argument("input"); prepare.add_argument("output")
    position=prepare.add_mutually_exclusive_group(required=True)
    position.add_argument("--after-paragraph",type=int)
    position.add_argument("--at-end",action="store_true")
    prepare.add_argument("--expected-sha256",required=True)
    prepare.add_argument("--slot",default="{{PAPERFLOW_BODY}}")
    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "gui":
            from .gui import main as gui_main
            return gui_main()
        elif args.command == "mcp":
            from .mcp_server import main as mcp_main
            return mcp_main([])
        elif args.command in {"content-plan", "content-review"}:
            from .content_review import prepare_content_review, run_content_review
            if args.command == "content-plan":
                emit(prepare_content_review(args.input, args.output_dir, evidence=args.evidence))
            else:
                emit(run_content_review(args.input, args.output_dir, evidence=args.evidence, max_rounds=args.max_rounds))
        elif args.command == "content-check":
            from .content_review import validate_content_review
            report = validate_content_review(args.input, args.plan)
            emit(report)
            if report.get("valid") is False:
                return 2
        elif args.command == "content-apply":
            from .content_review import apply_content_review
            emit(apply_content_review(args.input, args.plan, args.output, accepted_ids=args.accept))
        elif args.command in {"polish-plan", "polish-run"}:
            from .polish import prepare_polish, run_polish
            options = {"language": args.language, "voice_sample": Path(args.voice_sample) if args.voice_sample else None}
            if args.command == "polish-run":
                options["max_rounds"] = args.max_rounds
                emit(run_polish(args.input, args.output_dir, **options))
            else:
                emit(prepare_polish(args.input, args.output_dir, **options))
        elif args.command == "polish-apply":
            from .polish import apply_polish
            emit(apply_polish(args.input, args.plan, args.output))
        elif args.command == "doctor":
            try:
                pandoc = find_pandoc()
            except RuntimeError:
                pandoc = None
            from .mathtype import discover_mathtype
            emit({"version": __version__, "python": sys.executable, "pandoc": pandoc,
                  "modules": {n: importlib.util.find_spec(n) is not None for n in ["win32com", "docx", "lxml", "pymupdf", "PIL"]},
                  "mathtype": discover_mathtype().to_dict(), "word_session_started": False})
        elif args.command == "zotero-find":
            from .zotero import ZoteroClient
            query = args.query.strip()
            if not query:
                raise ValueError("Enter a title or author to find")
            records = ZoteroClient().search(query, limit=args.limit)
            emit([{"key": record["key"], "title": record["data"].get("title"),
                   "date": record["data"].get("date"), "doi": record["data"].get("DOI")}
                  for record in records if record.get("data", {}).get("title")])
        elif args.command == "plan-citations":
            from .bibliography import plan_migration
            emit(plan_migration(args.input), args.report)
        elif args.command == "apply-citations":
            from .bibliography import apply_migration
            emit(apply_migration(args.input, _read_json(args.plan), args.output), args.report)
        elif args.command == "plan-citekeys":
            from .citekeys import plan_citekeys
            emit(plan_citekeys(args.input, args.bibliography), args.report)
        elif args.command in {"inspect-template", "audit"}:
            report = inspect_docx(args.input) if args.command == "inspect-template" else audit_docx(args.input)
            emit(report, args.report)
            if args.command == "audit" and not report["structural_pass"]:
                return 2
        elif args.command == "import":
            emit(import_document(args.input,args.output,template=args.template,ocr_mode=args.ocr_mode,ocr_language=args.ocr_language),args.report)
        elif args.command == "recovery-plan":
            from .recovery import prepare_recovery
            emit(prepare_recovery(args.input, args.output_dir, ocr_mode=args.ocr_mode, ocr_language=args.ocr_language))
        elif args.command == "apply-recovery":
            from .recovery import apply_recovery
            emit(apply_recovery(args.intake, args.plan, args.output), args.report)
        elif args.command == "expand":
            from .package import marked_zotero_keys
            from .zotero import ZoteroClient, set_docx_preferences, document_preferences
            keys = marked_zotero_keys(args.input)
            client = ZoteroClient() if keys else None
            items = {k: client.get_item(k) for k in keys}
            report = expand_markers(args.input, args.output, items)
            if report["citations_used"]:
                set_docx_preferences(args.output, document_preferences(args.csl_style))
            emit(report, args.report)
        elif args.command == "build":
            from .word import build_from_template
            styles = _read_json(args.style_map)
            emit(build_from_template(args.template, args.input, args.output, slot=args.slot, style_map=styles), args.report)
        elif args.command == "refresh":
            from .word import refresh_fields
            emit(refresh_fields(args.input, args.output))
        elif args.command == "export-pdf":
            from .word import export_pdf
            emit(export_pdf(args.input, args.output))
        elif args.command == "run":
            from .pipeline import run_pipeline
            metadata=_read_json(args.metadata)
            result=run_pipeline(args.input,args.template,args.output_dir,csl_style=args.csl_style,metadata=metadata,
                                journal=args.journal,slot=args.slot,table_policy=args.table_policy,submission_copy=args.submission_copy,
                                ocr_mode=args.ocr_mode,ocr_language=args.ocr_language,
                                role_style_map=_read_json(args.role_style_map),draft_plan=_read_json(args.draft_plan),
                                layout_plan=_read_json(args.layout_plan), caption_chapter_level=args.caption_chapter_level,
                                caption_chapter_separator=args.caption_chapter_separator,
                                citekey_plan=_read_json(args.citekey_plan), table_style=args.table_style)
            emit({"native_build_finished":result.get("native_build_finished",False),"output_dir":args.output_dir,
                  "input_needs_semantic_review":result.get("input_needs_semantic_review",False),
                  "layout_needs_review":result.get("layout_needs_review",False),
                  "saved_csl_style":result["citation_style"]["saved_style_id"],
                  "journal_failed":result.get("journal_audit",{}).get("failed",[]),
                  "journal_unresolved":result.get("journal_audit",{}).get("unresolved",[])})
        elif args.command == "revise-master":
            from .pipeline import revise_master
            result=revise_master(args.input,args.output_dir,csl_style=args.csl_style,journal=args.journal,
                                 submission_copy=args.submission_copy,metadata=_read_json(args.metadata),artwork_root=args.artwork_root)
            emit({"native_build_finished":result.get("native_build_finished",False),"output_dir":args.output_dir,
                  "layout_needs_review":result.get("layout_needs_review",False),
                  "source_unchanged":result.get("source_unchanged"),
                  "saved_csl_style":result["citation_style"]["saved_style_id"],
                  "journal_failed":result.get("journal_audit",{}).get("failed",[]),
                  "journal_unresolved":result.get("journal_audit",{}).get("unresolved",[])})
        elif args.command == "template-slots":
            from .template import inventory_template
            emit(inventory_template(args.input),args.report)
        elif args.command == "inspect-layout":
            from .layout import inspect_layout
            report = inspect_layout(args.input, args.template)
            if args.plan_output:
                emit(report["plan"], args.plan_output)
            emit(report, args.report)
        elif args.command == "prepare-template":
            from .template import prepare_template
            emit(prepare_template(args.input,args.output,after_paragraph=args.after_paragraph,
                                  expected_sha256=args.expected_sha256,slot=args.slot))
        elif args.command == "plan-draft":
            from .semantics import plan_draft
            emit(plan_draft(args.input, citation_map=_read_json(args.citation_map),
                            caption_chapter_level=args.caption_chapter_level,
                            caption_chapter_separator=args.caption_chapter_separator), args.report)
        elif args.command == "apply-draft-plan":
            from .semantics import apply_draft_plan
            emit(apply_draft_plan(args.input,args.output,_read_json(args.plan)),args.report)
    except (ValueError, RuntimeError, OSError, KeyError) as exc:
        print(json.dumps({"error": type(exc).__name__, "message": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1
    return 0
