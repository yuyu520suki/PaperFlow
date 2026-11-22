"""Local manual interface for the same persistent jobs exposed through MCP."""
from __future__ import annotations

from datetime import datetime
import json
import os
from pathlib import Path
import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .service import console_python


PROJECT = Path(__file__).resolve().parents[1]
JOURNALS = {"不指定期刊": None, "Computers & Structures": "computers-and-structures"}
LANGUAGES = {"自动识别": "auto", "中文": "zh", "English": "en"}
CONTENT_CLASSIFICATIONS = {
    "supported": "已有证据支持", "overstated": "论断偏强", "contradicted": "与证据矛盾",
    "missing_evidence": "证据不足", "citation_gap": "引文待核实",
}
STATUS_NAMES = {
    "starting": "启动中", "cancel_requested": "等待任务取消",
    "queued": "排队中", "pending": "等待中", "running": "运行中", "completed": "已完成",
    "succeeded": "已完成", "success": "已完成", "done": "已完成", "failed": "失败",
    "cancelled": "已取消", "canceled": "已取消", "interrupted": "已中断",
}
SOURCE_TYPES = [("支持的原稿", "*.docx *.doc *.rtf *.odt *.md *.txt *.tex *.html *.htm *.rst *.pdf *.png *.jpg *.jpeg *.tif *.tiff"),
                ("所有文件", "*.*")]


def _required(value, label):
    value = str(value or "").strip()
    if not value:
        raise ValueError("请填写" + label)
    return value


def build_layout_args(values, *, revise=False):
    """Pure form-to-argv mapping, also usable by offline checks."""
    args = ["revise-master" if revise else "run", _required(values.get("source"), "原稿路径")]
    if not revise:
        args += ["--template", _required(values.get("template"), "模板路径")]
    args += ["--output-dir", _required(values.get("output"), "新输出目录")]
    journal = values.get("journal")
    journal = JOURNALS.get(journal, journal)
    if journal:
        if journal not in set(JOURNALS.values()):
            raise ValueError("当前仅支持 Computers & Structures 期刊配置")
        args += ["--journal", journal]
    if values.get("submission_copy"):
        args.append("--submission-copy")
    if str(values.get("metadata") or "").strip():
        args += ["--metadata", str(values["metadata"]).strip()]
    if revise:
        if str(values.get("artwork_root") or "").strip():
            args += ["--artwork-root", str(values["artwork_root"]).strip()]
    else:
        policy = values.get("table_policy", "preserve")
        if policy not in {"preserve", "three-line"}:
            raise ValueError("表格边框策略应为 preserve 或 three-line")
        args += ["--table-policy", policy]
    return args


def build_polish_args(values, *, execute=False):
    args = ["polish-run" if execute else "polish-plan", _required(values.get("source"), "待润色稿件")]
    args += ["--output-dir", _required(values.get("output"), "新输出目录")]
    language = values.get("language", "auto")
    language = LANGUAGES.get(language, language)
    if language not in {"auto", "zh", "en"}:
        raise ValueError("请选择自动识别、中文或 English")
    args += ["--language", language]
    if str(values.get("voice_sample") or "").strip():
        args += ["--voice-sample", str(values["voice_sample"]).strip()]
    return args


def build_polish_apply_args(source, plan, output):
    return ["polish-apply", _required(source, "待润色稿件"), _required(plan, "已审阅计划"),
            _required(output, "润色稿输出路径")]


def build_content_args(values, *, execute=False):
    args = ["content-review" if execute else "content-plan", _required(values.get("source"), "待审稿件"),
            "--output-dir", _required(values.get("output"), "新输出目录")]
    if str(values.get("evidence") or "").strip():
        args += ["--evidence", str(values["evidence"]).strip()]
    return args


def build_content_apply_args(source, plan, output, accepted_ids):
    args = ["content-apply", _required(source, "待审稿件"), _required(plan, "内容审阅计划"),
            _required(output, "修订稿输出路径")]
    if not accepted_ids:
        raise ValueError("请先采用至少一条有证据的建议。")
    for suggestion_id in accepted_ids:
        args += ["--accept", suggestion_id]
    return args


def build_utility_args(action, *, source="", template="", output=""):
    if action == "doctor":
        return ["doctor"]
    if action == "inspect-template":
        return [action, _required(template, "模板路径")]
    if action == "audit":
        return [action, _required(source, "Word 稿件路径")]
    if action == "export-pdf":
        return [action, _required(source, "Word 稿件路径"), _required(output, "PDF 输出路径")]
    raise ValueError("未知辅助操作")


def parse_cli_args(text):
    try:
        values = json.loads(text)
    except json.JSONDecodeError as error:
        raise ValueError("请输入 JSON 参数数组，例如 [\"doctor\"]") from error
    if not isinstance(values, list) or not values or any(not isinstance(value, str) for value in values):
        raise ValueError("参数必须是非空 JSON 字符串数组")
    if not values[0].strip() or any("\x00" in value for value in values):
        raise ValueError("命令不能为空，参数不能包含空字符")
    return values


def provider_environment(base_url, model, api_key):
    return {name: value.strip() for name, value in (
        ("PAPERFLOW_AI_BASE_URL", base_url), ("PAPERFLOW_AI_MODEL", model),
        ("PAPERFLOW_AI_API_KEY", api_key)) if value.strip()}


def mcp_configuration():
    return {"mcpServers": {"paperflow": {"command": console_python(), "args": ["-s", "-X", "utf8", "-m", "paperflow", "mcp"],
                                        "env": {"PYTHONNOUSERSITE": "1", "PYTHONPATH": str(PROJECT)}}}}


def _review_markers(paragraph):
    markers = {}
    for index, capsule in enumerate(paragraph["capsules"], start=1):
        description = " ".join(capsule["display"].split())[:35] or "文档对象"
        markers[capsule["token"]] = f"⟦{index}·{description}⟧"
    return markers


def _review_text(text, markers, *, restore=False):
    mapping = {value: key for key, value in markers.items()} if restore else markers
    if not mapping:
        return text
    return re.sub("|".join(re.escape(key) for key in mapping), lambda match: mapping[match.group()], text)


def _save_review_paragraph(paragraph, text, accepted):
    from .polish import _validate_proposal
    proposed = _review_text(text, _review_markers(paragraph), restore=True)
    if accepted:
        _validate_proposal(paragraph, proposed)
    paragraph.update(proposed_text=proposed, accepted=bool(accepted))


def _content_review_items(plan, report):
    paragraphs = {item["id"]: item for item in plan["paragraphs"]}
    evidence = {item["id"]: item for item in plan["evidence"]}
    items = []
    for finding in report["findings"]:
        paragraph = paragraphs[finding["paragraph_id"]]
        adoptable = finding["adoptable"] is True
        classification = CONTENT_CLASSIFICATIONS[finding["classification"]]
        status = classification + (" · 可采用" if adoptable else " · 仅供审阅")
        source_location = " · ".join([paragraph["chapter"], paragraph["locator"]])
        details = ["原句 · " + source_location, finding["quote"], "", "问题", finding["issue"], "", "证据与引文"]
        for link in finding["evidence"]:
            entry = evidence.get(link["id"], {})
            citation = [entry.get("title"), entry.get("doi") and "DOI " + entry["doi"],
                        entry.get("zotero_key") and "Zotero " + entry["zotero_key"],
                        entry.get("metadata_only") and "仅文献元数据，需核对全文"]
            location = " · ".join([link["locator"], *filter(None, citation)])
            details.extend([location, link["quote"], ""])
        if not finding["evidence"]:
            details.extend(["当前材料未提供对应的证据摘录。", ""])
        if finding.get("proposed_text"):
            proposed = _review_text(finding["proposed_text"], _review_markers(paragraph))
            details.extend(["建议稿", proposed, ""])
            if finding.get("blocked_reasons"):
                details.extend(["校验结果", *finding["blocked_reasons"], ""])
        else:
            details.extend(["处理建议", "请作者核查出处及所需证据。", ""])
        for role, label in (("consistency", "一致性评议"), ("literature", "文献评议")):
            review = finding.get("reviews", {}).get(role)
            if review:
                decision = "通过" if review["approved"] else "待核查"
                details.extend([label + " · " + decision, review["reason"], ""])
        items.append({"id": finding["id"], "paragraph_id": finding["paragraph_id"],
                      "adoptable": adoptable, "status": status, "details": "\n".join(details)})
    return items


class PaperFlowApp:
    def __init__(self, root, service=None):
        if service is None:
            from .service import Service
            service = Service()
        self.root, self.service = root, service
        self.jobs = {}
        self.job_outputs = {}
        self.selected_job = None
        self.last_log = None
        self.refresh_error = None
        self.layout = {key: tk.StringVar() for key in ("source", "template", "output", "metadata", "artwork_root")}
        self.layout["output"].set(self._new_output("排版"))
        self.layout_journal = tk.StringVar(value="不指定期刊")
        self.layout_table = tk.StringVar(value="保留边框")
        self.layout_submission = tk.BooleanVar(value=True)
        self.polish = {key: tk.StringVar() for key in ("source", "output", "voice_sample", "plan")}
        self.polish["output"].set(self._new_output("润色"))
        self.polish_language = tk.StringVar(value="自动识别")
        self.content = {key: tk.StringVar() for key in ("source", "evidence", "output", "plan")}
        self.content["output"].set(self._new_output("内容审阅"))
        self.provider_url = tk.StringVar(value=os.environ.get("PAPERFLOW_AI_BASE_URL", ""))
        self.provider_model = tk.StringVar(value=os.environ.get("PAPERFLOW_AI_MODEL", ""))
        self.provider_key = tk.StringVar(value=os.environ.get("PAPERFLOW_AI_API_KEY", ""))
        self._secrets = {self.provider_key.get().strip()} - {""}
        self.summary = tk.StringVar(value="选择原稿和模板，或先检查本机环境。")
        self.job_summary = tk.StringVar(value="尚未选择任务")
        self.autoscroll = tk.BooleanVar(value=True)
        self._style()
        self._build()
        self.root.after(300, self._poll)

    @staticmethod
    def _new_output(prefix, parent=None):
        return str((Path(parent) if parent else PROJECT / "output") / f"{prefix}-{datetime.now():%Y%m%d-%H%M%S}")

    def _style(self):
        self.root.title("PaperFlow · 论文工作流")
        self.root.geometry(f"1080x{max(680, min(830, self.root.winfo_screenheight() - 90))}")
        self.root.minsize(920, 680)
        self.root.configure(bg="#F3F6FA")
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure(".", font=("Microsoft YaHei UI", 10))
        style.configure("TFrame", background="#F3F6FA")
        style.configure("TLabel", background="#F3F6FA", foreground="#23364B")
        style.configure("TLabelframe", background="#F3F6FA", bordercolor="#D8E1EC")
        style.configure("TLabelframe.Label", background="#F3F6FA", foreground="#23435C", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("TButton", padding=(12, 7))
        style.configure("Primary.TButton", background="#126F77", foreground="white", padding=(16, 9))
        style.map("Primary.TButton", background=[("active", "#0D5960")])
        style.configure("TNotebook", background="#F3F6FA", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(19, 10))
        style.configure("Treeview", rowheight=29, font=("Microsoft YaHei UI", 9))
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 9, "bold"))

    def _build(self):
        header = tk.Frame(self.root, bg="#17314B", padx=24, pady=16)
        header.pack(fill="x")
        tk.Label(header, text="PaperFlow", font=("Segoe UI", 24, "bold"), fg="white", bg="#17314B").pack(anchor="w")
        tk.Label(header, text="模板排版  ·  动态文献  ·  可编辑公式  ·  审阅与润色", font=("Microsoft YaHei UI", 10),
                 fg="#CEE4EF", bg="#17314B").pack(anchor="w", pady=(3, 0))
        self.tabs = ttk.Notebook(self.root)
        self.tabs.pack(fill="both", expand=True, padx=18, pady=(14, 8))
        self.layout_tab = ttk.Frame(self.tabs, padding=16)
        self.content_tab = ttk.Frame(self.tabs, padding=16)
        self.polish_tab = ttk.Frame(self.tabs, padding=16)
        self.ai_tab = ttk.Frame(self.tabs, padding=16)
        self.jobs_tab = ttk.Frame(self.tabs, padding=16)
        for frame, name in ((self.layout_tab, "排版 / 修订"), (self.content_tab, "内容审阅"), (self.polish_tab, "润色"),
                            (self.ai_tab, "AI 接入"), (self.jobs_tab, "任务与日志")):
            self.tabs.add(frame, text=name)
        self._layout_tab()
        self._content_tab()
        self._polish_tab()
        self._ai_tab()
        self._jobs_tab()
        ttk.Label(self.root, textvariable=self.summary, padding=(22, 7)).pack(fill="x")

    def _path_row(self, parent, row, label, variable, *, kind="file", filetypes=None):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=7, padx=(0, 12))
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=7)
        if kind == "new-directory":
            command = lambda: self._choose_new_output(variable)
            text = "选择父目录"
        elif kind == "directory":
            command = lambda: self._choose_directory(variable)
            text = "选择目录"
        else:
            command = lambda: self._choose_file(variable, filetypes)
            text = "选择文件"
        ttk.Button(parent, text=text, command=command).grid(row=row, column=2, sticky="e", padx=(10, 0))
        parent.columnconfigure(1, weight=1)

    def _choose_file(self, variable, filetypes):
        name = filedialog.askopenfilename(parent=self.root, filetypes=filetypes or SOURCE_TYPES)
        if name:
            variable.set(name)

    def _choose_directory(self, variable):
        name = filedialog.askdirectory(parent=self.root)
        if name:
            variable.set(name)

    def _choose_new_output(self, variable):
        parent = filedialog.askdirectory(parent=self.root, title="选择父目录，工具会填入一个新的输出子目录")
        if parent:
            variable.set(self._new_output("paperflow", parent))

    def _layout_tab(self):
        paths = ttk.LabelFrame(self.layout_tab, text="稿件与模板", padding=13)
        paths.pack(fill="x")
        self._path_row(paths, 0, "原稿 / 已编辑主稿", self.layout["source"])
        self._path_row(paths, 1, "模板（首次套版）", self.layout["template"], filetypes=[("Word 模板", "*.docx *.dotx")])
        self._path_row(paths, 2, "新输出目录", self.layout["output"], kind="new-directory")
        self._path_row(paths, 3, "元数据（可选）", self.layout["metadata"], filetypes=[("JSON", "*.json")])
        self._path_row(paths, 4, "图件根目录（修订）", self.layout["artwork_root"], kind="directory")
        options = ttk.Frame(self.layout_tab)
        options.pack(fill="x", pady=14)
        ttk.Label(options, text="期刊配置").pack(side="left")
        ttk.Combobox(options, textvariable=self.layout_journal, values=list(JOURNALS), state="readonly", width=25).pack(side="left", padx=(10, 22))
        ttk.Label(options, text="表格边框").pack(side="left")
        ttk.Combobox(options, textvariable=self.layout_table, values=["保留边框", "三线表"], state="readonly", width=12).pack(side="left", padx=10)
        ttk.Checkbutton(options, text="生成投稿副本", variable=self.layout_submission).pack(side="left", padx=10)
        ttk.Button(options, text="载入样稿", command=self._load_example).pack(side="right")
        actions = ttk.Frame(self.layout_tab)
        actions.pack(fill="x", pady=(0, 12))
        ttk.Button(actions, text="首次套版", style="Primary.TButton", command=lambda: self._layout_job(False)).pack(side="left")
        ttk.Button(actions, text="刷新 Word 改稿", command=lambda: self._layout_job(True)).pack(side="left", padx=10)
        for label, action in (("检查环境", "doctor"), ("查看模板", "inspect-template"), ("检查主稿", "audit"), ("导出 PDF", "export-pdf")):
            ttk.Button(actions, text=label, command=lambda action=action: self._utility_job(action)).pack(side="left", padx=(0, 8))
        ttk.Label(self.layout_tab, text="输出目录需填写新目录。修改 editable-master.docx 后，使用“刷新 Word 改稿”继续。\n"
                  "完成后请核对 PDF 和报告中的版面、期刊待核查项；其他参数可在“任务与日志”页提交。", wraplength=940,
                  foreground="#5D7186", justify="left").pack(anchor="w", pady=7)

    def _polish_tab(self):
        paths = ttk.LabelFrame(self.polish_tab, text="润色稿件与审阅计划", padding=13)
        paths.pack(fill="x")
        self._path_row(paths, 0, "待润色稿件", self.polish["source"], filetypes=[("支持润色的稿件", "*.docx *.md *.txt")])
        self._path_row(paths, 1, "新输出目录", self.polish["output"], kind="new-directory")
        self._path_row(paths, 2, "写作样例（可选）", self.polish["voice_sample"], filetypes=[("文字样例", "*.txt *.md")])
        self._path_row(paths, 3, "已审阅计划 JSON", self.polish["plan"], filetypes=[("JSON", "*.json")])
        provider = ttk.LabelFrame(self.polish_tab, text="当前任务的模型接口", padding=13)
        provider.pack(fill="x", pady=12)
        for row, label, variable in ((0, "Base URL", self.provider_url), (1, "模型名称", self.provider_model), (2, "API 密钥", self.provider_key)):
            ttk.Label(provider, text=label).grid(row=row, column=0, sticky="w", padx=(0, 15), pady=5)
            ttk.Entry(provider, textvariable=variable, show="●" if row == 2 else "").grid(row=row, column=1, sticky="ew", pady=5)
        provider.columnconfigure(1, weight=1)
        actions = ttk.Frame(self.polish_tab)
        actions.pack(fill="x")
        ttk.Label(actions, text="语言").pack(side="left")
        ttk.Combobox(actions, textvariable=self.polish_language, values=list(LANGUAGES), state="readonly", width=12).pack(side="left", padx=10)
        ttk.Button(actions, text="生成审阅计划", command=lambda: self._polish_job(False)).pack(side="left", padx=(0, 10))
        ttk.Button(actions, text="多代理润色", style="Primary.TButton", command=lambda: self._polish_job(True)).pack(side="left", padx=(0, 10))
        ttk.Button(actions, text="逐段审阅", command=self._review_polish).pack(side="left", padx=(0, 10))
        ttk.Button(actions, text="应用已审阅计划", command=self._apply_polish).pack(side="left")
        ttk.Label(self.polish_tab, text="模型接口只用于当前任务，API 密钥以密码框输入并保留在内存中。\n"
                  "先查看修改计划，再应用到新文件；正文事实、数据、引用和公式仍需作者核对。", wraplength=940,
                  foreground="#5D7186", justify="left").pack(anchor="w", pady=12)

    def _content_tab(self):
        paths = ttk.LabelFrame(self.content_tab, text="稿件、证据和审阅计划", padding=13)
        paths.pack(fill="x")
        self._path_row(paths, 0, "待审稿件", self.content["source"], filetypes=[("支持的稿件", "*.docx *.md *.txt")])
        self._path_row(paths, 1, "证据文件（可选）", self.content["evidence"],
                       filetypes=[("实验记录、证据摘录", "*.txt *.md *.json")])
        self._path_row(paths, 2, "新输出目录", self.content["output"], kind="new-directory")
        self._path_row(paths, 3, "审阅计划", self.content["plan"], filetypes=[("审阅计划 JSON", "*.json")])
        actions = ttk.Frame(self.content_tab)
        actions.pack(fill="x", pady=(14, 6))
        ttk.Button(actions, text="载入审阅样例", command=self._load_content_example).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="离线准备", style="Primary.TButton", command=lambda: self._content_job(False)).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="运行模型评议", command=lambda: self._content_job(True)).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="配置模型", command=lambda: self.tabs.select(self.polish_tab)).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="校验计划", command=self._check_content).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="逐条审阅并应用", command=self._review_content).pack(side="left")
        ttk.Label(self.content_tab, text="可直接使用 TXT 或 Markdown 实验记录。离线审阅包可交给已接入 PaperFlow 的 MCP 客户端填写。\n"
                  "模型评议使用“润色”页的接口，会向该接口发送稿件和所选证据。每次准备或评议请选择新的输出目录。",
                  wraplength=850, foreground="#5D7186", justify="left").pack(anchor="w", pady=8)

    def _ai_tab(self):
        ttk.Label(self.ai_tab, text="将下面的配置加入支持 MCP 的 AI 客户端，即可调用同一套论文任务。", wraplength=930).pack(anchor="w", pady=(0, 12))
        self.mcp_text = tk.Text(self.ai_tab, height=14, wrap="none", font=("Consolas", 10), bg="#FDFEFF", relief="solid", borderwidth=1)
        self.mcp_text.pack(fill="x")
        self.mcp_text.insert("1.0", json.dumps(mcp_configuration(), ensure_ascii=False, indent=2))
        self.mcp_text.configure(state="disabled")
        controls = ttk.Frame(self.ai_tab)
        controls.pack(fill="x", pady=12)
        ttk.Button(controls, text="复制 MCP 配置", style="Primary.TButton", command=lambda: self._copy(self.mcp_text.get("1.0", "end-1c"))).pack(side="left")
        ttk.Button(controls, text="检查当前环境", command=lambda: self._utility_job("doctor")).pack(side="left", padx=10)
        ttk.Button(controls, text="打开使用说明", command=lambda: self._open_path(PROJECT / "开始使用.md")).pack(side="left")
        ttk.Label(self.ai_tab, text="当前 Python", foreground="#5D7186").pack(anchor="w", pady=(10, 4))
        runtime = ttk.Entry(self.ai_tab)
        runtime.pack(fill="x")
        runtime.insert(0, sys.executable)
        runtime.configure(state="readonly")
        ttk.Label(self.ai_tab, text="配置使用当前环境的 Python 和项目目录。Word 排版任务由共享任务服务排队。\n"
                  "内容审阅和润色共用“润色”页的接口地址、模型和密钥。", foreground="#5D7186", justify="left",
                  wraplength=940).pack(anchor="w", pady=14)

    def _jobs_tab(self):
        top = ttk.Frame(self.jobs_tab)
        top.pack(fill="x")
        self.job_tree = ttk.Treeview(top, columns=("command", "status", "time"), show="headings", height=5, selectmode="browse")
        for name, label, width in (("command", "任务", 480), ("status", "状态", 110), ("time", "提交时间", 170)):
            self.job_tree.heading(name, text=label)
            self.job_tree.column(name, width=width, minwidth=85, stretch=name == "command")
        scroll = ttk.Scrollbar(top, orient="vertical", command=self.job_tree.yview)
        self.job_tree.configure(yscrollcommand=scroll.set)
        self.job_tree.pack(side="left", fill="x", expand=True)
        scroll.pack(side="right", fill="y")
        self.job_tree.bind("<<TreeviewSelect>>", self._select_job)
        controls = ttk.Frame(self.jobs_tab)
        controls.pack(fill="x", pady=8)
        for label, command in (("刷新任务", self._manual_refresh), ("取消任务", self._cancel_job), ("打开输出目录", self._open_output),
                               ("打开主稿", lambda: self._open_output_file("editable-master.docx")),
                               ("打开 PDF", lambda: self._open_output_file("editable-master.pdf")),
                               ("复制日志", lambda: self._copy(self.log_text.get("1.0", "end-1c")))):
            ttk.Button(controls, text=label, command=command).pack(side="left", padx=(0, 8))
        ttk.Checkbutton(controls, text="自动滚动", variable=self.autoscroll).pack(side="right")
        ttk.Label(self.jobs_tab, textvariable=self.job_summary).pack(anchor="w", pady=(0, 5))
        log_frame = ttk.Frame(self.jobs_tab)
        log_frame.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_frame, height=10, wrap="word", font=("Consolas", 10), bg="#10283B", fg="#E3EDF4", relief="flat", padx=10, pady=9)
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set, state="disabled")
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")
        advanced = ttk.LabelFrame(self.jobs_tab, text="高级操作 · JSON 参数数组", padding=9)
        advanced.pack(fill="x", pady=(10, 0))
        self.advanced_text = tk.Text(advanced, height=2, wrap="word", font=("Consolas", 10))
        self.advanced_text.insert("1.0", '["doctor"]')
        self.advanced_text.pack(side="left", fill="x", expand=True)
        ttk.Button(advanced, text="提交任务", command=self._advanced_job).pack(side="right", padx=(10, 0))

    @staticmethod
    def _existing_file(value, label):
        path = Path(_required(value, label)).expanduser()
        if not path.is_file():
            raise ValueError(f"{label}不存在：{path}")
        return str(path.resolve())

    def _load_example(self):
        example = PROJECT / "examples/computers-and-structures"
        self.layout["source"].set(str(example / "manuscript.md"))
        self.layout["template"].set(str(PROJECT / "templates/computers-and-structures.docx"))
        self.layout["metadata"].set(str(example / "metadata.json"))
        self.layout["artwork_root"].set(str(example))
        self.layout["output"].set(self._new_output("样稿"))
        self.layout_journal.set("Computers & Structures")
        self.layout_table.set("三线表")
        self.layout_submission.set(True)
        self.summary.set("已载入合成样稿；可先检查环境，再点击首次套版。")

    @staticmethod
    def _new_directory(value):
        path = Path(_required(value, "新输出目录")).expanduser()
        if path.exists():
            raise ValueError("输出目录已经存在，请填写一个新的目录名称。")
        return str(path.resolve())

    def _layout_job(self, revise):
        try:
            values = {key: value.get() for key, value in self.layout.items()}
            values["source"] = self._existing_file(values["source"], "原稿")
            if revise and Path(values["source"]).suffix.casefold() != ".docx":
                raise ValueError("刷新改稿需要已保存的 DOCX 主稿。")
            if not revise:
                values["template"] = self._existing_file(values["template"], "模板")
            values["output"] = self._new_directory(values["output"])
            if values["metadata"].strip():
                values["metadata"] = self._existing_file(values["metadata"], "元数据文件")
            values.update(journal=self.layout_journal.get(), submission_copy=self.layout_submission.get(),
                          table_policy="three-line" if self.layout_table.get() == "三线表" else "preserve")
            self._start(build_layout_args(values, revise=revise), output=values["output"])
        except Exception as error:
            self._error(error)

    def _utility_job(self, action):
        try:
            source, template, output = self.layout["source"].get(), self.layout["template"].get(), ""
            if action == "inspect-template":
                template = self._existing_file(template, "模板")
            elif action in {"audit", "export-pdf"}:
                source = self._existing_file(source, "Word 稿件")
            if action == "export-pdf":
                output = filedialog.asksaveasfilename(parent=self.root, title="导出到新的 PDF 文件", defaultextension=".pdf",
                    initialfile=Path(source).stem + "-预览.pdf", filetypes=[("PDF", "*.pdf")], confirmoverwrite=False)
                if not output:
                    return
                if Path(output).exists():
                    raise ValueError("PDF 输出文件已经存在，请选择一个新文件名。")
            self._start(build_utility_args(action, source=source, template=template, output=output),
                        output=str(Path(output).parent) if output else None)
        except Exception as error:
            self._error(error)

    def _polish_job(self, execute):
        try:
            values = {key: value.get() for key, value in self.polish.items()}
            values["source"] = self._existing_file(values["source"], "待润色稿件")
            values["output"] = self._new_directory(values["output"])
            if values["voice_sample"].strip():
                values["voice_sample"] = self._existing_file(values["voice_sample"], "写作样例")
            values["language"] = self.polish_language.get()
            environment = self._provider_environment() if execute else None
            self._start(build_polish_args(values, execute=execute), output=values["output"], environment=environment)
            self.polish["plan"].set(str(Path(values["output"]) / ("reviewed-plan.json" if execute else "plan.json")))
        except Exception as error:
            self._error(error)

    def _provider_environment(self):
        environment = provider_environment(self.provider_url.get(), self.provider_model.get(), self.provider_key.get())
        if len(environment) != 3:
            raise ValueError("模型评议需要在“润色”页填写 Base URL、模型名称和 API 密钥。")
        return environment

    def _load_content_example(self):
        folder = PROJECT / "examples/content-review"
        try:
            source = self._existing_file(folder / "manuscript.md", "审阅示例稿件")
            evidence = self._existing_file(folder / "evidence.json", "审阅示例证据")
            plan = self._existing_file(folder / "review-plan.json", "已审阅示例计划")
            self.content["source"].set(source)
            self.content["evidence"].set(evidence)
            self.content["output"].set(self._new_output("内容审阅样例"))
            self.content["plan"].set(plan)
            self.summary.set("已载入审阅样例。可逐条审阅现成建议，也可离线另建审阅包。")
        except Exception as error:
            self._error(error)

    def _content_job(self, execute):
        try:
            values = {key: value.get() for key, value in self.content.items()}
            values["source"] = self._existing_file(values["source"], "待审稿件")
            values["output"] = self._new_directory(values["output"])
            if values["evidence"].strip():
                values["evidence"] = self._existing_file(values["evidence"], "证据文件")
            environment = self._provider_environment() if execute else None
            self._start(build_content_args(values, execute=execute), output=values["output"], environment=environment)
            self.content["plan"].set(str(Path(values["output"]) / "review-plan.json"))
        except Exception as error:
            self._error(error)

    def _content_paths(self):
        return (self._existing_file(self.content["source"].get(), "待审稿件"),
                self._existing_file(self.content["plan"].get(), "内容审阅计划"))

    def _check_content(self):
        try:
            source, plan = self._content_paths()
            self._start(["content-check", source, plan])
        except Exception as error:
            self._error(error)

    def _review_content(self):
        try:
            from .content_review import validate_content_review

            source, path = self._content_paths()
            plan = json.loads(Path(path).read_text(encoding="utf-8-sig"))
            report = validate_content_review(source, path)
            if not report["valid"]:
                raise ValueError("审阅计划未通过校验，请查看校验任务的结果。")
            items = _content_review_items(plan, report)
            if not items:
                messagebox.showinfo("等待审阅建议", "可将离线审阅包交给已接入 PaperFlow 的 MCP 客户端填写建议，或运行模型评议。",
                                    parent=self.root)
                return
            self._content_review_window(source, path, items)
        except Exception as error:
            self._error(error)

    def _content_review_window(self, source, plan, items):
        window = tk.Toplevel(self.root)
        window.title("内容审阅 · " + Path(plan).name)
        window.geometry("860x560")
        window.minsize(700, 450)
        window.transient(self.root)
        window.grab_set()
        panel = ttk.Frame(window, padding=14)
        panel.pack(fill="both", expand=True)
        title = tk.StringVar(master=window)
        choice = tk.BooleanVar(master=window)
        count = tk.StringVar(master=window)
        accepted = {}
        index = 0
        ttk.Label(panel, textvariable=title).pack(anchor="w", pady=(0, 6))
        reader = ttk.Frame(panel)
        reader.pack(fill="both", expand=True)
        details = tk.Text(reader, wrap="word", font=("Microsoft YaHei UI", 11), padx=12, pady=8)
        scroll = ttk.Scrollbar(reader, orient="vertical", command=details.yview)
        details.configure(yscrollcommand=scroll.set, state="disabled")
        details.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        controls = ttk.Frame(panel)
        controls.pack(fill="x", pady=(10, 0))

        def display():
            item = items[index]
            title.set(f"第 {index + 1} / {len(items)} 条  ·  {item['id']}  ·  {item['status']}")
            details.configure(state="normal")
            details.delete("1.0", "end")
            details.insert("1.0", item["details"])
            details.configure(state="disabled")
            choice.set(accepted.get(item["paragraph_id"]) == item["id"])
            adopt.configure(state="normal" if item["adoptable"] else "disabled")
            previous.configure(state="normal" if index else "disabled")
            following.configure(state="normal" if index + 1 < len(items) else "disabled")

        def move(step):
            nonlocal index
            index += step
            display()

        def update_choice():
            item = items[index]
            if choice.get() and item["adoptable"]:
                accepted[item["paragraph_id"]] = item["id"]
            elif accepted.get(item["paragraph_id"]) == item["id"]:
                del accepted[item["paragraph_id"]]
            count.set(f"已采用 {len(accepted)} 条")
            apply.configure(state="normal" if accepted else "disabled")

        def apply_selected():
            try:
                suffix = Path(source).suffix
                target = filedialog.asksaveasfilename(parent=window, title="应用建议并保存为新稿件", defaultextension=suffix,
                    initialdir=str(Path(plan).parent), initialfile=Path(source).stem + "-内容修订" + suffix,
                    filetypes=[("稿件", "*" + suffix)], confirmoverwrite=False)
                if not target:
                    return
                output = Path(target).resolve()
                if output.exists():
                    raise ValueError("输出文件已经存在，请选择一个新文件名。")
                self._start(build_content_apply_args(source, plan, str(output), accepted.values()), output=str(output.parent))
                window.destroy()
            except Exception as error:
                messagebox.showerror("应用内容审阅建议", self._redact(error), parent=window)

        adopt = ttk.Radiobutton(controls, text="采用本条", variable=choice, value=True, command=update_choice)
        adopt.pack(side="left")
        ttk.Radiobutton(controls, text="忽略本条", variable=choice, value=False, command=update_choice).pack(side="left", padx=8)
        previous = ttk.Button(controls, text="上一条", command=lambda: move(-1))
        following = ttk.Button(controls, text="下一条", command=lambda: move(1))
        previous.pack(side="left", padx=(10, 4))
        following.pack(side="left")
        apply = ttk.Button(controls, text="应用到新文件", style="Primary.TButton", command=apply_selected, state="disabled")
        apply.pack(side="right")
        ttk.Label(controls, textvariable=count).pack(side="right", padx=12)
        count.set("已采用 0 条")
        display()

    def _apply_polish(self):
        try:
            source = self._existing_file(self.polish["source"].get(), "待润色稿件")
            plan = self._existing_file(self.polish["plan"].get(), "已审阅计划")
            suffix = Path(source).suffix or ".docx"
            output = filedialog.asksaveasfilename(parent=self.root, title="保存润色稿到新文件", defaultextension=suffix,
                initialfile=Path(source).stem + "-润色稿" + suffix, filetypes=[("稿件", "*" + suffix), ("所有文件", "*.*")], confirmoverwrite=False)
            if not output:
                return
            if Path(output).exists():
                raise ValueError("输出文件已经存在，请选择一个新文件名。")
            self._start(build_polish_apply_args(source, plan, output), output=str(Path(output).parent))
        except Exception as error:
            self._error(error)

    def _review_polish(self):
        try:
            path = self.polish["plan"].get()
            if not path:
                self._choose_file(self.polish["plan"], [("审阅计划", "*.json")])
                path = self.polish["plan"].get()
            if not path:
                return
            plan = json.loads(Path(path).read_text(encoding="utf-8-sig"))
            if plan.get("schema_version") != 1 or not isinstance(plan.get("paragraphs"), list):
                raise ValueError("请选择 PaperFlow 生成的润色计划。")
            paragraphs = [item for item in plan["paragraphs"] if item["editable"]]
            if not paragraphs:
                raise ValueError("此计划没有可编辑的正文段落。")
            self._review_window(Path(path), plan, paragraphs)
        except Exception as error:
            self._error(error)

    def _review_window(self, path, plan, paragraphs):
        window = tk.Toplevel(self.root)
        window.title("逐段审阅 · " + path.name)
        window.geometry("830x580")
        window.transient(self.root)
        window.grab_set()
        panel = ttk.Frame(window, padding=14)
        panel.pack(fill="both", expand=True)
        title = tk.StringVar(master=window)
        accepted = tk.BooleanVar(master=window)
        index = 0
        ttk.Label(panel, textvariable=title).pack(anchor="w")
        source = tk.Text(panel, wrap="word", height=5, font=("Microsoft YaHei UI", 11))
        source.pack(fill="both", expand=True, pady=(5, 12))
        ttk.Label(panel, text="建议稿（保留 ⟦ ⟧ 内的数字、引文和公式）").pack(anchor="w")
        proposal = tk.Text(panel, wrap="word", height=7, font=("Microsoft YaHei UI", 11), undo=True)
        proposal.pack(fill="both", expand=True, pady=(5, 12))
        controls = ttk.Frame(panel)
        controls.pack(fill="x")
        ttk.Checkbutton(controls, text="采用本段建议", variable=accepted).pack(side="left")

        def display():
            paragraph = paragraphs[index]
            markers = _review_markers(paragraph)
            title.set(f"原稿 · 第 {index + 1} / {len(paragraphs)} 段")
            source.configure(state="normal")
            source.delete("1.0", "end")
            source.insert("1.0", _review_text(paragraph["text"], markers))
            source.configure(state="disabled")
            proposal.delete("1.0", "end")
            proposal.insert("1.0", _review_text(paragraph["proposed_text"], markers))
            proposal.edit_reset()
            accepted.set(paragraph["accepted"])
            previous.configure(state="normal" if index else "disabled")
            following.configure(state="normal" if index + 1 < len(paragraphs) else "disabled")

        def capture():
            try:
                _save_review_paragraph(paragraphs[index], proposal.get("1.0", "end-1c"), accepted.get())
                return True
            except ValueError as error:
                messagebox.showerror("请核对本段", str(error), parent=window)
                return False

        def move(step):
            nonlocal index
            if capture():
                index += step
                display()

        def save():
            if not capture():
                return
            target = filedialog.asksaveasfilename(parent=window, title="另存已审阅计划", initialdir=str(path.parent),
                       initialfile=path.stem + "-人工审阅.json", defaultextension=".json", filetypes=[("JSON", "*.json")],
                       confirmoverwrite=False)
            if not target:
                return
            try:
                self.service.write_text(target, json.dumps({**plan, "plan_path": target}, ensure_ascii=False, indent=2))
            except OSError as error:
                messagebox.showerror("请选择新文件名", str(error), parent=window)
                return
            self.polish["plan"].set(target)
            self.polish["source"].set(plan["source"])
            self.summary.set("已保存人工审阅计划，可应用到稿件副本。")
            window.destroy()

        previous = ttk.Button(controls, text="上一段", command=lambda: move(-1))
        following = ttk.Button(controls, text="下一段", command=lambda: move(1))
        previous.pack(side="left", padx=(12, 4))
        following.pack(side="left")
        ttk.Button(controls, text="另存审阅计划", command=save).pack(side="right")
        display()

    def _advanced_job(self):
        try:
            args = parse_cli_args(self.advanced_text.get("1.0", "end-1c"))
            environment = self._provider_environment() if args[0] in {"polish-run", "content-review"} else None
            self._start(args, environment=environment)
        except Exception as error:
            self._error(error)

    def _start(self, args, *, output=None, environment=None):
        if environment and environment.get("PAPERFLOW_AI_API_KEY"):
            self._secrets.add(environment["PAPERFLOW_AI_API_KEY"])
        job = self.service.start_cli(args, env_overrides=environment)
        job_id = str(job.get("job_id") or job.get("id") or "")
        if not job_id:
            raise RuntimeError("任务服务没有返回任务编号")
        if output:
            self.job_outputs[job_id] = output
        self.selected_job = job_id
        self.summary.set("任务已提交：" + args[0] + " · " + job_id[:12])
        self.tabs.select(self.jobs_tab)
        self._refresh_jobs()
        if self.job_tree.exists(job_id):
            self.job_tree.selection_set(job_id)
            self.job_tree.see(job_id)
        self._refresh_selected()

    def _refresh_jobs(self):
        records = self.service.list_jobs(limit=50)
        self.jobs = {str(job.get("job_id") or job.get("id")): job for job in records if job.get("job_id") or job.get("id")}
        for item in self.job_tree.get_children():
            if item not in self.jobs:
                self.job_tree.delete(item)
        for job_id, job in self.jobs.items():
            args = job.get("args") or job.get("argv") or []
            command = " ".join(str(value) for value in args[:2]) if isinstance(args, list) else str(args)
            command = command or str(job.get("command") or job.get("kind") or job_id)
            status = str(job.get("status", "unknown"))
            created = str(job.get("created_at") or job.get("created_utc") or job.get("started_at") or "")[:19].replace("T", " ")
            values = (command, STATUS_NAMES.get(status, status), created)
            if self.job_tree.exists(job_id):
                self.job_tree.item(job_id, values=values)
            else:
                self.job_tree.insert("", "end", iid=job_id, values=values)

    def _select_job(self, _event=None):
        selection = self.job_tree.selection()
        if selection:
            self.selected_job = selection[0]
            self.last_log = None
            try:
                self._refresh_selected()
            except Exception as error:
                self.summary.set("读取任务失败：" + str(error))

    def _redact(self, value):
        text = str(value or "")
        secrets = self._secrets | ({self.provider_key.get().strip()} - {""})
        for secret in sorted(secrets, key=len, reverse=True):
            text = text.replace(secret, "[已隐藏密钥]")
        return text

    def _manual_refresh(self):
        try:
            self._refresh_jobs()
            self._refresh_selected()
        except Exception as error:
            self._error(error)

    def _cancel_job(self):
        try:
            if not self.selected_job:
                raise ValueError("请先选择任务。")
            self.service.cancel_job(self.selected_job)
            self._manual_refresh()
        except Exception as error:
            self._error(error)

    def _refresh_selected(self):
        if not self.selected_job:
            return
        job = self.service.job_status(self.selected_job, tail_chars=12000)
        self.jobs[self.selected_job] = job
        status = str(job.get("status", "unknown"))
        self.job_summary.set("任务 " + self.selected_job + "  ·  " + STATUS_NAMES.get(status, status))
        segments = ["标准输出\n" + self._redact(job.get("stdout", "")), "诊断日志\n" + self._redact(job.get("stderr", ""))]
        if job.get("result") is not None:
            segments.append("结果\n" + self._redact(json.dumps(job["result"], ensure_ascii=False, indent=2, default=str)))
        if job.get("exit_code") is not None:
            segments.append("退出码：" + str(job["exit_code"]))
        text = "\n\n".join(segments)
        if text != self.last_log:
            position = self.log_text.yview()
            self.log_text.configure(state="normal")
            self.log_text.delete("1.0", "end")
            self.log_text.insert("1.0", text)
            self.log_text.configure(state="disabled")
            if self.autoscroll.get():
                self.log_text.see("end")
            elif position:
                self.log_text.yview_moveto(position[0])
            self.last_log = text

    def _poll(self):
        try:
            self._refresh_jobs()
            self._refresh_selected()
            self.refresh_error = None
        except Exception as error:
            detail = self._redact(error)
            if detail != self.refresh_error:
                self.summary.set("任务状态暂未更新：" + detail)
                self.refresh_error = detail
        self.root.after(1200, self._poll)

    def _output_path(self):
        if not self.selected_job:
            raise ValueError("请先选择一个任务")
        if self.selected_job in self.job_outputs:
            return Path(self.job_outputs[self.selected_job])
        job = self.jobs.get(self.selected_job, {})
        result = job.get("result")
        if isinstance(result, dict) and result.get("output_dir"):
            return Path(result["output_dir"])
        args = job.get("args") or job.get("argv") or []
        if isinstance(args, list) and "--output-dir" in args:
            index = args.index("--output-dir")
            if index + 1 < len(args):
                return Path(args[index + 1])
        raise ValueError("此任务未提供输出目录，请在日志中查看文件路径。")

    def _open_output(self):
        try:
            self._open_path(self._output_path())
        except Exception as error:
            self._error(error)

    def _open_output_file(self, name):
        try:
            self._open_path(self._output_path() / name)
        except Exception as error:
            self._error(error)

    def _open_path(self, path):
        try:
            path = Path(path).expanduser().resolve(strict=True)
            os.startfile(str(path))
        except Exception as error:
            self._error(error)

    def _copy(self, text):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.summary.set("已复制到剪贴板")

    def _error(self, error):
        detail = self._redact(error)
        self.summary.set(detail)
        messagebox.showerror("PaperFlow", detail, parent=self.root)


def main():
    root = tk.Tk()
    PaperFlowApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
