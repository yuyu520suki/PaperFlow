# PaperFlow 论文排版工作流

PaperFlow 将原稿套入指定的 Word 模板，输出可继续修改的主稿、分页预览和可选的投稿副本。主稿沿用模板的分节、页眉页脚与样式，保留 Word 目录、图表和公式编号、交叉引用、Zotero 动态引文及 MathType 公式。手动界面、命令行和 MCP 客户端共用排版、内容审阅和润色任务。

## 安装

在 Windows x64 的 Conda 命令行中进入本项目目录，创建独立环境并安装项目。

```powershell
conda env create -f environment-win-64.lock.yml
conda activate paperflow
python -s -m pip install --no-user --no-build-isolation --no-deps -e .
.\PaperFlow.cmd doctor
```

锁定文件记录本版所用的 Python 3.12、Pandoc 3.8.3、Tk 和 Python 包版本。如需让 Conda 重新求解兼容依赖，创建环境时使用 `environment.yml`。请保留项目目录供可编辑安装、模板和示例图件使用。启动器会识别已激活的 `paperflow` 环境和默认用户环境目录。在其他位置安装环境时，可在 PowerShell 设置 `$env:PAPERFLOW_PYTHON = "$env:CONDA_PREFIX\python.exe"` 后运行启动器。

| 用途 | 本机需要的组件 |
| --- | --- |
| 原生排版、更新域、导出 PDF | Microsoft Word 桌面版 |
| 创建和编辑 MathType 公式 | 已注册 `Equation.DSMT4` 的 MathType，含可用的本机 SDK |
| 引文、书目和样式切换 | 运行中的 Zotero Desktop、Zotero 的 Word 插件、已打开的本地 API 和选用的 CSL 样式 |
| 扫描页识别 | Windows OCR 及原稿对应的识别语言 |

Computers & Structures 引用样式依赖 Elsevier Vancouver。可通过 Zotero 的引用样式管理功能依次添加项目所附的父样式和期刊样式。样式文件保留了原作者和 CC BY-SA 3.0 许可。 [CSL 文件](references/computers-and-structures/)

## 手动运行

运行 `PaperFlow-界面.cmd`，或在已激活环境中运行 `.\PaperFlow.cmd gui`。选择“载入样稿”可以填入自带的 Markdown、Word 模板和绘图文件。选择“首次套版”后，在“任务与日志”查看进度与输出路径。用自己的论文时，把原稿和模板换成相应文件，选择新的输出目录。操作步骤和改稿流程见[开始使用](开始使用.md)。

| 输出 | 用途 |
| --- | --- |
| `editable-master.docx` | 在 Word 中继续修改的版本，保留动态域和可编辑公式 |
| `editable-master.pdf` | 检查分页、字体、目录和图表位置 |
| `submission-copy.docx`、`submission-copy.pdf` | 选择生成投稿副本后得到。Zotero 引文固定为当前显示值，期刊配置可追加文末图表 |
| `workflow.json`、`artwork/` | 步骤记录、期刊检查结果和已配置的独立图件 |

`docs/排版与修订.md` 说明如何使用自定义模板插入点、按章编号、BibTeX 引用映射、多节布局和旧稿迁移。对于扫描 PDF 和图片，先核对识别文本、表格与公式，再将审阅后的结构化稿件用于排版。 [排版与修订](docs/排版与修订.md) · [OCR 说明](docs/ocr.md)

## AI 操作

在“AI 接入”页复制当前机器的 MCP 配置，加入支持 stdio 的客户端。客户端按当前 Windows 用户权限读取原稿、运行命令并写输出。可先调用 `capabilities` 取得命令清单，再用 `command_help` 查询参数。

例如，调用 `start_cli`，传入 `{"args":["inspect-template","templates/computers-and-structures.docx"]}`。用返回的 `job_id` 调用 `job_status` 查看结果。长任务请保持宿主连接到任务结束。GUI 与 MCP 都可在默认工作区查看持久化记录。用 `read_text`、`write_text` 可编辑审阅计划，用 `cancel_job` 请求取消。任务记录保存在本地 `.paperflow/jobs/`。

## 学术润色

“润色”支持 DOCX、Markdown 和 TXT。生成审阅计划后，可在“逐段审阅”修改建议、选择采用的段落并另存计划，再应用到新文件。此流程可离线使用，也可由已接入的 AI 通过 MCP 填写候选修改。

“多代理润色”会将可编辑段落及可选写作样例提交至所配置的模型接口。语言编辑提出改写，事实评审与学术表达评审各自检查，整合者汇总结果。最多两轮，生成 `reviewed-plan.json`、`review-log.json` 和新稿。支持本机回环 HTTP 或远程 HTTPS 的 Chat Completions 接口，设置 Base URL、模型名和密钥即可。请核对论断强度、术语和引用，按目标期刊要求披露 AI 辅助。程序检查数字和受保护的文档对象，检测器的 AI 百分比须由相应检测器另行测量。 [编辑方法和参考项目](docs/polishing-sources.md)

## 内容审阅

在“内容审阅”页载入合成样例，或选择自己的 DOCX、Markdown 或 TXT 稿件。证据可选 TXT 或 Markdown 实验记录，也可导入带出处的 JSON 摘录。可以离线生成审阅包并交给接入 PaperFlow 的 MCP 宿主填写，也可使用“润色”页的当前接口配置运行模型评议。逐条核对原文、问题、证据摘录和出处后，采用有证据支持的建议，另存为修订稿。 [内容审阅操作步骤](开始使用.md#审阅稿件内容)

这一流程参考了 PaperOrchestra 的分工和审阅思路。PaperFlow 从已有稿件和用户选定的证据开始，审阅者负责确认采用的修改。 [PaperOrchestra 项目](https://github.com/google-research/paper-orchestra)

## Computers & Structures 示例的适用范围

本仓库根据作者指南制作了单栏示例模板。A4、25.4 mm 页边距、12 磅正文等是项目采用的排版参数。正式稿可直接指定自己的 DOCX 或 DOTX 模板。示例内容与图表为合成数据，默认示例无需绑定某位用户的 Zotero 条目。 [期刊格式依据](docs/computers-and-structures-requirements.md) · [示例模板出处](templates/computers-and-structures.provenance.json)
