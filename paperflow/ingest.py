"""Input adapters. Original DOCX objects are never passed through Pandoc."""
from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
from collections import Counter, defaultdict
from math import isfinite
from pathlib import Path
from .ocr import SUPPORTED_IMAGES, _normalize_ocr_text, recognize_image


PANDOC_FORMATS = {".md": "markdown-implicit_figures", ".markdown": "markdown-implicit_figures", ".html": "html", ".htm": "html", ".tex": "latex", ".rst": "rst"}
LATEX_AST_FORMAT = "latex+raw_tex"
_LATEX_FLOAT_AMBIGUITY = re.compile(rb"\\caption(?:of)?\s*\*")
_LATEX_INCLUDED_FILE = re.compile(rb"\\(?:input|include|subfile|import)\b")
_LATEX_HEADING_NUMBERING_AMBIGUITY = re.compile(
    rb"\\(?:(?:setcounter|addtocounter)\s*\{\s*secnumdepth\s*\}|(?:appendix|frontmatter|backmatter)\b)"
)
_LATEX_HEADING_AUTOREF_AMBIGUITY = re.compile(rb"\\(?:chapter|part|appendix|[A-Za-z]*autorefname)\b")
_LATEX_COMMENT = re.compile(rb"(?<!\\)(?:\\\\)*%[^\r\n]*")


def find_pandoc():
    # MCP clients may start the environment's Python without activating Conda.
    prefix = Path(sys.prefix)
    for relative in ("Library/bin/pandoc.exe", "Scripts/pandoc.exe", "bin/pandoc"):
        candidate = prefix / relative
        if candidate.is_file():
            return str(candidate)
    executable = shutil.which("pandoc")
    if not executable:
        candidate = Path.home() / "AppData/Local/Pandoc/pandoc.exe"
        if candidate.is_file():
            executable = str(candidate)
    if not executable:
        raise RuntimeError("Pandoc is not installed; no software will be installed automatically")
    return executable


def _read_latex_ast(pandoc, source, report):
    # Disabling generated heading identifiers distinguishes real LaTeX labels.
    parsed = subprocess.run([pandoc, "--from", LATEX_AST_FORMAT + "-auto_identifiers", "--to=json", str(source)],
                            cwd=source.parent, capture_output=True, check=False, timeout=180)
    warning = parsed.stderr.decode("utf-8", errors="replace").strip()
    if parsed.returncode:
        raise RuntimeError(warning)
    if warning:
        report["warnings"].append(warning)
        report["semantic_review_required"] = True
    return json.loads(parsed.stdout)


def _paragraph_lines(text):
    return [line for line in text.splitlines() if line.strip()]


def _write_text_docx(text, output):
    from docx import Document
    doc = Document()
    for line in _paragraph_lines(text):
        doc.add_paragraph(line)
    with output.open("xb") as stream:
        doc.save(stream)


def _source_line(text, page, adapter, bbox=None):
    return {"text": text, "page": page, "bbox": bbox, "adapter": adapter}


def _normalized_box(box, width, height):
    if width <= 0 or height <= 0 or not all(isfinite(value) for value in (*box, width, height)):
        return None
    x0, y0, x1, y1 = box
    values = [max(0.0, min(1.0, value)) for value in (x0 / width, y0 / height, x1 / width, y1 / height)]
    return values if values[0] < values[2] and values[1] < values[3] else None


def _ocr_box(box, width, height):
    try:
        x, y, box_width, box_height = (float(box[key]) for key in ("x", "y", "width", "height"))
    except (KeyError, TypeError, ValueError):
        return None
    return _normalized_box((x, y, x + box_width, y + box_height), width, height)


def _ocr_source_lines(text, recognized, page_number, adapter):
    source_lines = [_source_line(line, page_number, adapter) for line in _paragraph_lines(text)]
    image = recognized.get("image", {})
    try:
        width, height = float(image["width"]), float(image["height"])
    except (KeyError, TypeError, ValueError):
        return source_lines

    matched_lines = []
    for recognized_line in recognized["lines"]:
        parts = _paragraph_lines(_normalize_ocr_text(recognized_line["text"]))
        for part in parts:
            matched_lines.append((part, recognized_line if len(parts) == 1 else None))
    if len(source_lines) != len(matched_lines):
        return source_lines

    for source_line, (part, recognized_line) in zip(source_lines, matched_lines):
        if recognized_line is None or source_line["text"].strip() != part.strip():
            continue
        source_line["bbox"] = _ocr_box(recognized_line.get("bbox"), width, height)
        words = []
        for word in recognized_line.get("words", []):
            box = _ocr_box(word.get("bbox"), width, height)
            token = word.get("text")
            if box is None or not isinstance(token, str) or not token.strip():
                words = []
                break
            words.append({"text": token, "bbox": box})
        if words:
            source_line["words"] = words
    return source_lines


def _pdf_source_lines(text, page, page_number):
    import pymupdf

    # Sorted plain-text extraction can merge or wrap rows. Keep that text and
    # attach a box only when a native line is an unambiguous match.
    flags = pymupdf.TEXTFLAGS_DICT & ~pymupdf.TEXT_PRESERVE_IMAGES
    native = defaultdict(list)
    for block in page.get_text("dict", flags=flags)["blocks"]:
        for line in block.get("lines", []):
            key = " ".join("".join(span["text"] for span in line["spans"]).split())
            if key:
                native[key].append(line["bbox"])

    lines = _paragraph_lines(text)
    keys = [" ".join(line.split()) for line in lines]
    counts = Counter(keys)
    source_lines = []
    display = page.rect
    rotation = page.rotation_matrix
    for line, key in zip(lines, keys):
        bounds = native.get(key, [])
        box = None
        if counts[key] == 1 and len(bounds) == 1:
            rect = pymupdf.Rect(bounds[0]) * rotation
            box = _normalized_box(
                (rect.x0 - display.x0, rect.y0 - display.y0, rect.x1 - display.x0, rect.y1 - display.y0),
                display.width,
                display.height,
            )
        source_lines.append(_source_line(line, page_number, "pdf-text-layer", box))
    return source_lines


def _decode_text(source):
    raw = source.read_bytes()
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16"), "UTF-16"
    try:
        return raw.decode("utf-8-sig"), "UTF-8"
    except UnicodeDecodeError:
        return raw.decode("gb18030"), "GB18030"


def import_document(source, output, template=None, app=None, *, ocr_mode="auto", ocr_language=None,
                    line_provenance=False, allow_empty=False, pandoc_ast=None, reuse_template_indexes=False,
                    role_style_map=None):
    """Import a source document and optionally locate each literal draft paragraph.

    TXT, PDF, and OCR images can report source_lines. Boxes are normalized
    [x0, y0, x1, y1] in the displayed page or decoded image, or None if no
    unambiguous box is available. Page numbers are one-based, with None for TXT.
    Recovery callers may allow empty PDF/image text for page-region review.
    """
    source, output = Path(source).resolve(strict=True), Path(output).resolve()
    if output.exists():
        raise FileExistsError(output)
    if ocr_mode not in {"auto", "always", "never"}:
        raise ValueError("OCR mode must be auto, always, or never")
    output.parent.mkdir(parents=True, exist_ok=True)
    suffix = source.suffix.lower()
    if pandoc_ast is not None and suffix not in PANDOC_FORMATS:
        raise ValueError("A prepared Pandoc document requires a structured input adapter")
    report = {"source": str(source), "output": str(output), "warnings": []}
    if suffix == ".docx":
        from .package import read_package
        read_package(source)
        with source.open("rb") as src, output.open("xb") as dst:
            shutil.copyfileobj(src, dst)
        report["adapter"] = "docx-preserve-package"
    elif suffix in {".doc", ".rtf", ".odt"}:
        from .word import convert_to_docx
        convert_to_docx(source, output, app=app)
        report["adapter"] = "word-native-import"
    elif suffix in SUPPORTED_IMAGES:
        if ocr_mode == "never":
            raise ValueError("An image needs OCR; OCR was explicitly disabled")
        sidecar=output.with_suffix(".ocr.json")
        result=recognize_image(source,sidecar,language=ocr_language)
        if not result["normalized_text"].strip():
            if not allow_empty:
                raise ValueError("OCR did not detect any text")
            report["warnings"].append("No text was recognized. Inspect the image for omitted equations and figures.")
        _write_text_docx(result["normalized_text"],output)
        report.update(adapter="windows-ocr-image",ocr_sidecars=[str(sidecar)],semantic_review_required=True)
        if line_provenance:
            report["source_lines"] = _ocr_source_lines(result["normalized_text"], result, 1, "windows-ocr-image")
        report["warnings"].append("Review OCR text against the source; formulas, table cells and heading roles are not reconstructed")
    elif suffix in {".pdf", ".txt"}:
        if suffix == ".txt":
            text, encoding = _decode_text(source)
            report.update(adapter="literal-text", encoding=encoding, semantic_review_required=True)
            if line_provenance:
                report["source_lines"] = [_source_line(line, None, "literal-text") for line in _paragraph_lines(text)]
            report["warnings"].append("Text is imported literally; supply heading, equation and table structure separately")
        else:
            import pymupdf
            chunks, pages, sidecars = [], [], []
            source_lines = []
            with pymupdf.open(source) as document:
                for page_number, page in enumerate(document, 1):
                    text = page.get_text(sort=True).strip() if ocr_mode!="always" else ""
                    adapter="pdf-text-layer"
                    if not text:
                        if ocr_mode=="never": raise ValueError(f"PDF page {page_number} has no text layer; OCR is disabled")
                        image=output.with_name(f"{output.stem}.page-{page_number:03d}.png")
                        with image.open("xb") as f: f.write(page.get_pixmap(dpi=240,alpha=False).tobytes("png"))
                        sidecar=image.with_suffix(".ocr.json")
                        recognized=recognize_image(image,sidecar,language=ocr_language)
                        text=recognized["normalized_text"].strip()
                        sidecars.append(str(sidecar)); adapter="windows-ocr"
                        if line_provenance:
                            source_lines.extend(_ocr_source_lines(text, recognized, page_number, adapter))
                    elif line_provenance:
                        source_lines.extend(_pdf_source_lines(text, page, page_number))
                    pages.append({"page":page_number,"adapter":adapter,"characters":len(text)})
                    chunks.append(text)
            text="\n\n".join(chunks)
            report.update(adapter="pdf-with-ocr" if sidecars else "pdf-text-layer",pages=pages,
                          ocr_sidecars=sidecars,semantic_review_required=True)
            if line_provenance:
                report["source_lines"] = source_lines
            report["warnings"].append("Review PDF reading order, headings, formula symbols and table structure against the original")
        if not text.strip():
            if not allow_empty or suffix != ".pdf":
                raise ValueError("Input does not contain recoverable text")
            report["warnings"].append("No text was recognized. Inspect the original pages for omitted equations and figures.")
        _write_text_docx(text,output)
    else:
        if suffix not in PANDOC_FORMATS:
            raise ValueError(f"No input adapter for {suffix}; original file retained")
        fmt = PANDOC_FORMATS[suffix]
        pandoc = find_pandoc()
        if suffix == ".tex":
            from .latex_references import needs_heading_labels, prepare_latex_references
            from .package import navigation_index_kind, part_fields, read_package, xml

            original = source.read_bytes()
            explicit_heading_ids = None if pandoc_ast is None else set()
            if pandoc_ast is None:
                pandoc_ast = _read_latex_ast(pandoc, source, report)
            elif needs_heading_labels(pandoc_ast):
                # Keep the AST used to verify existing citekey plans and their prepared Word fields.
                explicit = _read_latex_ast(pandoc, source, report)
                explicit_heading_ids = {block["c"][1][0] for block in explicit["blocks"] if block["t"] == "Header"}
            if source.read_bytes() != original:
                raise ValueError("The LaTeX source changed while Pandoc was reading it")
            clean_source = _LATEX_COMMENT.sub(b"", original)
            has_included_files = bool(_LATEX_INCLUDED_FILE.search(clean_source))
            template_indexes = set()
            numbered_heading_levels = set()
            if template is not None and reuse_template_indexes:
                template_parts = read_package(template)
                template_fields, _ = part_fields(xml(template_parts["word/document.xml"]))
                template_indexes = {kind for field in template_fields if (kind := navigation_index_kind(field))}
                if needs_heading_labels(pandoc_ast):
                    from .formatting import template_numbered_heading_styles

                    styles = template_numbered_heading_styles(template_parts, role_style_map)
                    if app is None:
                        numbered_heading_levels = set(styles)
                    else:
                        from .word import numbered_template_headings

                        numbered_heading_levels = numbered_template_headings(template, styles, app=app)
            pandoc_ast, refs = prepare_latex_references(
                pandoc_ast, explicit_heading_ids=explicit_heading_ids, template_indexes=template_indexes,
                ambiguous_float_numbering=has_included_files or bool(_LATEX_FLOAT_AMBIGUITY.search(clean_source)),
                numbered_heading_levels=numbered_heading_levels,
                ambiguous_heading_numbering=has_included_files or bool(_LATEX_HEADING_NUMBERING_AMBIGUITY.search(clean_source)),
                heading_autoref=not _LATEX_HEADING_AUTOREF_AMBIGUITY.search(clean_source),
            )
            report["latex_references"] = refs
            if refs["issues"]:
                report["semantic_review_required"] = True
        command = [pandoc, "--to", "docx", "--output", "-", "--resource-path", str(source.parent)]
        data = None
        if pandoc_ast is None:
            command.extend(["--from", fmt, str(source)])
        else:
            command.extend(["--from=json"])
            data = json.dumps(pandoc_ast, ensure_ascii=False).encode("utf-8")
        if template:
            command.extend(["--reference-doc", str(Path(template).resolve(strict=True))])
        result = subprocess.run(command, input=data, cwd=source.parent if pandoc_ast is not None else None,
                                capture_output=True, check=False, timeout=180)
        if result.returncode:
            raise RuntimeError(result.stderr.decode("utf-8", errors="replace"))
        with output.open("xb") as stream: stream.write(result.stdout)
        report["adapter"] = "pandoc-" + fmt if pandoc_ast is None else "pandoc-ast-docx"
        if result.stderr.strip():
            report["warnings"].append(result.stderr.decode("utf-8", errors="replace").strip())
            report["semantic_review_required"] = True
    return report
