# 学术润色的来源与边界

核查日期为 2026-11-16。以下为 GitHub 当日页面星数快照，后续会变化。PaperFlow 使用独立编写的流程和短提示词，参考这些项目的编辑思路，没有复制其完整 Skill。

| 来源 | 星数快照及许可 | 参考内容 |
|---|---|---|
| [blader/humanizer](https://github.com/blader/humanizer) | 约 69.6k，MIT | 结构化套话、重复收尾、空泛强调与作者声音校准，采用 [v3.1.0 的保真原则](https://github.com/blader/humanizer/blob/v3.1.0/SKILL.md) |
| [op7418/Humanizer-zh](https://github.com/op7418/Humanizer-zh) | 约 22.9k，MIT | 中文长定语、弱动词、万能背景及模板化结尾，保留文体、归因和不确定程度。[指导文件](https://github.com/op7418/Humanizer-zh/blob/main/SKILL.md) |
| [AIScientists-Dev/academic-humanizer](https://github.com/AIScientists-Dev/academic-humanizer) | 约 5.2k，MIT | 学术语体、主张与证据强度匹配、术语及数字公式引用保护。[指导文件](https://github.com/AIScientists-Dev/academic-humanizer/blob/main/SKILL.md) |
| [madaan/self-refine](https://github.com/madaan/self-refine) | 851，Apache-2.0；作为机制参考 | 有上限的生成、反馈、修订循环。[实际实现](https://github.com/madaan/self-refine/blob/main/src/responsegen/run.py) |

Humanizer 的 [v4.0.0 发布说明](https://github.com/blader/humanizer/releases/tag/v4.0.0) 允许为概括文字补入模型认为真实的具体细节。PaperFlow 不采用这条规则。缺少数据、来源或证据时，润色不能补造事实。

许可来源分别为 [Humanizer MIT](https://github.com/blader/humanizer/blob/main/LICENSE)、[Humanizer-zh MIT](https://github.com/op7418/Humanizer-zh/blob/main/LICENSE)、[Academic Humanizer MIT](https://github.com/AIScientists-Dev/academic-humanizer/blob/main/LICENSE) 和 [Self-Refine Apache-2.0](https://github.com/madaan/self-refine/blob/main/LICENSE)。如果后续直接引入其中的代码或大段提示词，应随引入内容保留相应许可和作者声明。

## 工作流程

离线准备创建带源文件 SHA-256 的 plan.json。DOCX 中的动态域、Zotero 引文、公式、OLE、图片、链接和书签范围作为不可更改的保护片段；相邻的普通文字仍可润色。数字、单位、引文及图表引用也形成保护片段。Markdown 的元数据、代码、标题和表格结构保留。

在线流程使用四个独立的对话上下文。语言编辑提出修改，事实评审和学术表达评审分别检查，整合者只接受两位评审都通过的原候选文字。每轮结果、被拒建议和最终差异留在 review-log.json。最多两轮。四个角色默认可使用同一个已配置模型，也可通过角色模型环境变量指定不同模型；角色不同不等于模型提供商不同。

prepare_polish 和 apply_polish 不发网络请求。run_polish 只使用明确设置的 PAPERFLOW_AI_BASE_URL、PAPERFLOW_AI_MODEL、PAPERFLOW_AI_API_KEY，调用兼容的 /chat/completions 接口。远程服务使用 HTTPS，本机回环服务可使用 HTTP。API Key 仅在请求头使用，不进入计划和日志。没有模型配置时，可由接入 MCP 的外部 AI 完成候选编辑和评议，再提交审阅计划供离线应用。

模型输出必须保留保护片段及顺序。源文件变化、数字增加、引文修改或保护片段丢失会终止应用。所有输出都是新文件，原稿保留。DOCX 修改只触及正文中允许编辑的文字节点，保留运行格式和其他 ZIP 部件。润色后的完整 Word 稿仍可用 revise-master 刷新目录和交叉引用。

确定性检查无法证明所有语义完全等价。否定、条件、范围、因果和归因仍需事实评审及作者核对。自然表达与清晰度是润色目标，不以模型自评分、词语命中或文本长度声称作者身份，也不承诺固定 AI 检测率。使用时遵守目标期刊的 AI 辅助披露要求。
