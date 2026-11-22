# 本地 OCR 输入适配器

`paperflow.ocr.recognize_image(source, output=None, language=None, timeout=60)` 使用已安装的 Windows.Media.Ocr。只读图片，在本机完成识别；不上传图像、不安装 OCR 语言或修改系统配置。可选 JSON 输出拒绝覆盖已有文件。

2026-11-04 本机实测可用语言为 `en-US`、`ja`、`zh-Hans-CN`。帮助程序显式调用 Windows PowerShell 5.1，因为本机默认 PowerShell 7 的运行时不能直接投影 WinRT 类型。脚本为 UTF-8 BOM，并已由 PowerShell 5.1 解析器检查。

返回内容包括引擎与语言、图像尺寸、逐行和逐词的像素坐标边界框。`raw_text` 保留原生合并文本；`text` 恢复引擎识别出的行分隔；`normalized_text` 进一步去除相邻汉字之间的空格，原始文本和边界框仍可用于复核。Windows OCR 未提供逐词置信度，本适配器不虚构置信度。

支持单帧 PNG、JPEG、BMP、TIFF。多帧图像或超过本机引擎最大尺寸 10000 像素的图像会明确拒绝，避免默默漏页或缩放。扫描 PDF 应由导入层逐页渲染为工作图片，再逐页调用本接口，保留页号与侧车 JSON。

OCR 只恢复可疑文本。公式符号、上下标、表格结构和阅读顺序都需复核；结果始终包含 `requires_review=true`、`equations_reconstructed=false`、`table_structure_reconstructed=false`。识别出的公式字符不等于已经构造 OMML 或 MathType 原生对象。

复现合成样本测试：

```powershell
conda run -n paperflow python -s scripts/ocr_smoke.py output/ocr-smoke-new
```

中英文各三行的合成样本已实测识别，文本在忽略空白后与输入匹配，行和词边界框存在。该测试没有覆盖手写、复杂表格、多栏阅读顺序或复杂数学公式。测试输入、OCR JSON 与证据报告位于 `output/ocr-smoke-v2`。

官方接口依据：[OcrEngine 结果结构](https://github.com/MicrosoftDocs/winrt-api/blob/docs/windows.media.ocr/ocrengine.md)、[RecognizeAsync](https://learn.microsoft.com/en-us/uwp/api/windows.media.ocr.ocrengine.recognizeasync)、[已安装识别语言](https://github.com/MicrosoftDocs/winrt-api/blob/docs/windows.media.ocr/ocrengine_availablerecognizerlanguages.md)、[微软 OCR 示例](https://github.com/Microsoft/Windows-universal-samples/blob/main/Samples/OCR/cs/OcrFileImage.xaml.cs)。
