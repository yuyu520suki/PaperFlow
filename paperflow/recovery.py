"""Build editable manuscript structure from a reviewed text/OCR intake."""
from __future__ import annotations

from collections import Counter
from copy import deepcopy
from io import BytesIO
import json
from pathlib import Path
import re
import subprocess

from .ingest import find_pandoc, import_document
from .ocr import SUPPORTED_IMAGES
from .package import digest, q


_VERSION = 1
_LABEL = re.compile(r"[A-Za-z][A-Za-z0-9_]{0,39}\Z")


def _read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def _write_json(path, value):
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, ensure_ascii=False, indent=2)
        stream.write("\n")


def _indices(block, line_count):
    first, last = block.get("start_line"), block.get("end_line")
    if ("start_line" in block and "end_line" in block and first is None and last is None
            and block.get("kind") in {"equation", "figure"} and "line_indices" not in block):
        # An OCR engine can omit a symbol-only line or an entire drawing.
        # An explicitly reviewed page region supplies the missing anchor.
        return []
    if type(first) is not int or type(last) is not int or not 0 <= first <= last < line_count:
        raise ValueError("A recovery block has an invalid source line range")
    sparse = block.get("line_indices")
    if sparse is None:
        return list(range(first, last + 1))
    if (not isinstance(sparse, list) or not sparse or any(type(i) is not int for i in sparse)
            or sparse != sorted(set(sparse)) or sparse[0] != first or sparse[-1] != last):
        raise ValueError("line_indices must be sorted, unique and match the block's range endpoints")
    return sparse


def _proposals(lines):
    from .recovery_proposals import propose_blocks
    from .recovery_geometry import detect_ocr_tables

    geometry = detect_ocr_tables(lines)
    claimed = {i for block in geometry for i in _indices(block, len(lines))}
    blocks = list(geometry)
    for block in propose_blocks([line["text"] for line in lines]):
        indices = _indices(block, len(lines))
        if not claimed.intersection(indices):
            blocks.append(block)
        else:
            blocks.extend({"start_line": i, "end_line": i, "kind": "paragraph"}
                          for i in indices if i not in claimed)
    blocks.sort(key=lambda block: block["start_line"])
    for block in blocks:
        block["source_preview"] = [lines[i]["text"] for i in _indices(block, len(lines))]
    return blocks


def prepare_recovery(source, output_dir, *, ocr_mode="auto", ocr_language=None):
    """Import an unstructured original and publish line evidence plus editable proposals."""
    source = Path(source).resolve(strict=True)
    if source.suffix.lower() not in {".txt", ".pdf", *SUPPORTED_IMAGES}:
        raise ValueError("Recovery accepts TXT, PDF or a supported single image")
    original_hash = digest(source.read_bytes())
    output = Path(output_dir).resolve()
    output.mkdir(parents=True, exist_ok=False)
    imported = output / "literal-import.docx"
    report = import_document(source, imported, ocr_mode=ocr_mode, ocr_language=ocr_language,
                             line_provenance=True, allow_empty=True)
    lines = report.pop("source_lines")
    if digest(source.read_bytes()) != original_hash:
        raise RuntimeError("The original file changed while importing source lines")
    intake = output / "source-lines.json"
    _write_json(intake, {"schema_version": _VERSION, "source": str(source), "source_sha256": original_hash,
                         "import": report, "source_lines": lines})
    plan = output / "recovery-plan.json"
    blocks = _proposals(lines)
    _write_json(plan, {"schema_version": _VERSION, "intake_sha256": digest(intake.read_bytes()), "blocks": blocks})
    empty_pages = [page["page"] for page in report.get("pages", []) if not page["characters"]]
    if not lines and source.suffix.lower() in SUPPORTED_IMAGES:
        empty_pages = [1]
    return {"intake": str(intake), "plan": str(plan), "literal_import": str(imported), "source_lines": len(lines),
            "proposals": dict(Counter(block["kind"] for block in blocks)), "pages_without_recognized_text": empty_pages,
            "review_required": True}


def _text(value, description):
    if not isinstance(value, str) or not value.strip() or any(ord(char) < 32 and char not in "\t\r\n" for char in value):
        raise ValueError(f"{description} must contain valid text")
    return value


def _caption(block, kind):
    caption, label = block.get("caption"), block.get("label")
    if caption is None and label is None:
        return None
    if not isinstance(label, str) or not _LABEL.fullmatch(label):
        raise ValueError(f"{kind} caption needs an ASCII label of up to 40 characters")
    caption = _text(caption, f"{kind} caption")
    if any(char in caption for char in "]\r\n"):
        raise ValueError(f"{kind} caption must be a single line without closing brackets")
    return f"[[CAPTION {kind} {label}|{caption}]]"


def _table_rows(block):
    rows = block.get("rows")
    if (not isinstance(rows, list) or not rows or not isinstance(rows[0], list) or not rows[0]
            or any(not isinstance(row, list) or len(row) != len(rows[0]) for row in rows)):
        raise ValueError("A table needs a nonempty, rectangular rows array")
    for row in rows:
        for cell in row:
            if cell != "":
                _text(cell, "Table cell")
    header_count = block.get("header_rows", 0)
    if type(header_count) is not int or not 0 <= header_count < len(rows):
        raise ValueError("header_rows must be a count smaller than the table's row count")
    return rows, header_count


def _word_equations(formulas):
    if not formulas:
        return []
    from docx import Document

    # Keep the converter local and only pass display math. A parsing warning
    # can indicate that part of the reviewed formula would become plain text.
    markdown = "\n\n".join(f"$$\n{formula}\n$$" for formula in formulas).encode("utf-8")
    result = subprocess.run([find_pandoc(), "--sandbox", "--from=markdown", "--to=docx", "--output=-"],
                            input=markdown, capture_output=True, timeout=60)
    if result.returncode or result.stderr.strip():
        raise ValueError("Reviewed LaTeX could not be converted to editable Word equations: " +
                         result.stderr.decode("utf-8", errors="replace").strip())
    document = Document(BytesIO(result.stdout))
    math = []
    for paragraph in document.paragraphs:
        objects = paragraph._p.findall(q("m:oMathPara"))
        if paragraph.text.strip() or len(objects) != 1 or len(list(objects[0].iter(q("m:oMath")))) != 1:
            raise ValueError("Each reviewed equation must produce exactly one display formula")
        math.append(objects[0])
    if len(math) != len(formulas):
        raise ValueError("One or more reviewed formulas were lost or split during conversion")
    return math


def _source_region(source, block):
    import math

    page_number = block.get("page")
    bbox = block.get("bbox")
    if (type(page_number) is not int or page_number < 1 or not isinstance(bbox, list) or len(bbox) != 4
            or any(type(value) not in {int, float} or not math.isfinite(value) or not 0 <= value <= 1 for value in bbox)
            or bbox[0] >= bbox[2] or bbox[1] >= bbox[3]):
        raise ValueError("A source region needs a page number and an ordered, normalized bbox [x0, y0, x1, y1]")
    if source.suffix.lower() != ".pdf" and (source.suffix.lower() not in SUPPORTED_IMAGES or page_number != 1):
        raise ValueError("A source region must come from a PDF or single-image original")
    return page_number, bbox


def _figure_crop(source, block):
    page_number, bbox = _source_region(source, block)
    if source.suffix.lower() == ".pdf":
        import pymupdf

        with pymupdf.open(source) as pdf:
            if page_number > len(pdf):
                raise ValueError("Figure page is outside the original PDF")
            page = pdf[page_number - 1]
            bounds = page.rect
            clip = pymupdf.Rect(bounds.x0 + bbox[0] * bounds.width, bounds.y0 + bbox[1] * bounds.height,
                               bounds.x0 + bbox[2] * bounds.width, bounds.y0 + bbox[3] * bounds.height)
            if clip.width < 1 or clip.height < 1:
                raise ValueError("The selected figure region is too small")
            return page.get_pixmap(clip=clip, dpi=240, alpha=False).tobytes("png")
    from PIL import Image

    with Image.open(source) as image:
        width, height = image.size
        clip = (round(bbox[0] * width), round(bbox[1] * height), round(bbox[2] * width), round(bbox[3] * height))
        if clip[0] >= clip[2] or clip[1] >= clip[3]:
            raise ValueError("The selected figure region is too small")
        result = BytesIO()
        image.crop(clip).convert("RGB").save(result, format="PNG")
        return result.getvalue()


def apply_recovery(intake_path, plan_path, output_path):
    """Apply reviewed roles, cells, formulas and figure regions to a new DOCX."""
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.shared import Cm

    intake_bytes = Path(intake_path).read_bytes()
    intake = json.loads(intake_bytes.decode("utf-8-sig"))
    plan = _read_json(plan_path)
    if (not isinstance(intake, dict) or not isinstance(plan, dict) or intake.get("schema_version") != _VERSION
            or plan.get("schema_version") != _VERSION or plan.get("intake_sha256") != digest(intake_bytes)):
        raise ValueError("The review plan does not match this source-lines report")
    source = Path(intake["source"]).resolve(strict=True)
    if digest(source.read_bytes()) != intake["source_sha256"]:
        raise ValueError("The original file changed after intake; prepare it again")
    output = Path(output_path).resolve()
    if output.suffix.lower() != ".docx" or output.exists():
        raise ValueError("Recovery output must be a new .docx file")
    lines, blocks = intake["source_lines"], plan.get("blocks")
    if not isinstance(blocks, list) or not blocks or any(not isinstance(block, dict) for block in blocks):
        raise ValueError("Recovery plan needs a nonempty blocks array")
    covered = Counter(i for block in blocks for i in _indices(block, len(lines)))
    if len(covered) != len(lines) or any(count != 1 for count in covered.values()):
        raise ValueError("Each original line must be assigned exactly once. Use omit with a reason for removed material.")
    unresolved = [block["start_line"] for block in blocks if block.get("kind") == "review"]
    if unresolved:
        raise ValueError(f"Classify or correct the unresolved blocks beginning at source lines {unresolved}")

    formulas = [_text(block.get("latex"), "Reviewed LaTeX formula") for block in blocks if block.get("kind") == "equation"]
    if any("$$" in formula for formula in formulas):
        raise ValueError("Provide LaTeX formula content without $$ display delimiters")
    equations = iter(_word_equations(formulas))
    document = Document()
    labels = set()
    notes = []
    rendered = Counter()
    for block in blocks:
        kind = block.get("kind")
        selected = _indices(block, len(lines))
        label = block.get("label")
        if label is not None:
            if not isinstance(label, str) or not _LABEL.fullmatch(label) or label.casefold() in labels:
                raise ValueError("Recovery labels must be valid and unique within the document")
            labels.add(label.casefold())
        if kind in {"title", "heading", "paragraph"}:
            content = block.get("text", " ".join(lines[i]["text"] for i in selected))
            content = _text(content, f"Recovery {kind}")
            if kind == "heading":
                level = block.get("level")
                if type(level) is not int or level not in (1, 2, 3):
                    raise ValueError("Heading level must be 1, 2 or 3")
                document.add_heading(content, level)
            else:
                document.add_paragraph(content, style="Title" if kind == "title" else None)
        elif kind == "table":
            rows, header_count = _table_rows(block)
            caption = _caption(block, "Table")
            if caption:
                document.add_paragraph(caption)
            table = document.add_table(rows=len(rows), cols=len(rows[0]))
            for row_index, row in enumerate(rows):
                if row_index < header_count:
                    table.rows[row_index]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
                for column_index, content in enumerate(row):
                    table.cell(row_index, column_index).text = content
        elif kind == "equation":
            if not selected:
                page, bbox = _source_region(source, block)
                if source.suffix.lower() == ".pdf":
                    import pymupdf

                    with pymupdf.open(source) as pdf:
                        if page > len(pdf):
                            raise ValueError("Equation region is outside the original PDF")
                notes.append({"kind": "reviewed_equation_region", "page": page, "bbox": bbox, "latex": block["latex"]})
            paragraph = document.add_paragraph()
            paragraph._p.append(deepcopy(next(equations)))
            if label:
                paragraph.add_run(f" [[EQNO {label}]]")
        elif kind == "figure":
            caption = _caption(block, "Figure")
            document.add_picture(BytesIO(_figure_crop(source, block)), width=Cm(12))
            if caption:
                document.add_paragraph(caption)
            notes.append({"kind": "raster_figure_crop", "page": block["page"], "bbox": block["bbox"], "lines": selected})
        elif kind == "omit":
            notes.append({"kind": "omitted_source_lines", "lines": selected, "reason": _text(block.get("reason"), "Omission reason")})
        else:
            raise ValueError(f"Unknown recovery block kind: {kind}")
        rendered[kind] += 1
    if not any(kind != "omit" for kind in rendered):
        raise ValueError("The recovery plan omits all source lines")
    if digest(source.read_bytes()) != intake["source_sha256"]:
        raise ValueError("The original file changed while applying the plan")
    output.parent.mkdir(parents=True, exist_ok=True)
    data = BytesIO()
    document.save(data)
    with output.open("xb") as stream:
        stream.write(data.getbuffer())
    return {"output": str(output), "source_sha256": intake["source_sha256"], "source_unchanged": True,
            "source_lines_accounted_for": len(covered), "applied_blocks": dict(rendered), "review_actions": notes}
