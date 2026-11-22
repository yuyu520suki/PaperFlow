"""Stage original Computers & Structures artwork and check its file properties."""

from __future__ import annotations

from collections import Counter
from hashlib import sha256
from io import BytesIO
from pathlib import Path
import math
import re

from .package import without_bidi_controls


_FIGURE_NUMBER = re.compile(
    r"^\s*(?:Figure|Fig\.)\s*((?:[A-Z]\.?)?[1-9]\d*)(?![A-Za-z0-9]|[.\-:–—]\d)",
    re.IGNORECASE | re.ASCII,
)
_RASTER_FORMATS = {".tif": "TIFF", ".tiff": "TIFF", ".jpg": "JPEG", ".jpeg": "JPEG", ".png": "PNG"}
_RASTER_RULES = {
    "photo": {"minimum_dpi": 300, "single_column": 1063, "full_page": 2244},
    "line": {"minimum_dpi": 1000, "single_column": 3543, "full_page": 7480},
}
_VECTOR_SUFFIXES = {".pdf", ".eps"}


def _source_path(source_root, value):
    if not isinstance(value, str) or not value.strip():
        raise ValueError("Figure source needs an input-relative file path")
    relative = Path(value)
    if relative.anchor or relative.drive or ".." in relative.parts or ":" in value:
        raise ValueError("Figure source path must stay under the manuscript input directory")
    candidate = (source_root / relative).resolve(strict=True)
    if not candidate.is_relative_to(source_root):
        raise ValueError("Resolved figure source points outside the manuscript input directory")
    if not candidate.is_file():
        raise ValueError("Figure source is not a file")
    return candidate


def _inspect_pdf(data):
    import pymupdf

    with pymupdf.open(stream=data, filetype="pdf") as document:
        if not document.is_pdf:
            raise ValueError("Artwork extension is PDF but the file is another format")
        if document.needs_pass or document.page_count != 1:
            return {"pages": document.page_count}, ["An artwork PDF needs one accessible page"], []
        page = document[0]
        paths = len(page.get_drawings())
        raster_images = len(page.get_image_info())
        fonts = []
        for xref, _, subtype, name, *_ in page.get_fonts(full=True):
            if subtype == "Type3":
                charprocs, _ = document.xref_get_key(xref, "CharProcs")
                embedded = charprocs in {"dict", "xref"}
            else:
                embedded = bool(document.extract_font(xref)[3])
            fonts.append({"name": name, "subtype": subtype, "embedded": embedded})

    facts = {"pages": 1, "vector_paths": paths, "painted_raster_images": raster_images, "fonts": fonts}
    failures, reviews = [], []
    if any(not font["embedded"] for font in fonts):
        failures.append("PDF references fonts without embedded glyph data")
    if raster_images:
        reviews.append("PDF paints raster artwork; verify the embedded resolution and artwork classification")
    if not paths:
        reviews.append("No vector drawing paths were detected")
    return facts, failures, reviews


def _inspect_eps(data):
    header = data.split(b"\n", 1)[0][:256]
    is_postscript = header.startswith(b"%!PS-Adobe-") and b"EPSF-" in header
    is_binary_eps = data.startswith(b"\xc5\xd0\xd3\xc6") and len(data) >= 30
    if not (is_postscript or is_binary_eps):
        raise ValueError("EPS extension does not match an encapsulated PostScript header")
    return {"binary_preview_header": is_binary_eps}, [], ["Verify EPS fonts and the drawing in an EPS-capable renderer"]


def _inspect_raster(data, suffix, kind, layout):
    from PIL import Image

    try:
        with Image.open(BytesIO(data)) as image:
            if image.format != _RASTER_FORMATS[suffix]:
                raise ValueError(f"Image data has format {image.format}, which disagrees with its extension")
            if getattr(image, "n_frames", 1) != 1:
                raise ValueError("Supply a separate single-frame raster file for each figure")
            width, height = image.size
            raw_dpi = image.info.get("dpi")
            image.verify()
        # Reading EXIF can load PNG pixels, after which Pillow no longer
        # permits verify(). Use a separate header read for orientation.
        with Image.open(BytesIO(data)) as image:
            orientation = image.getexif().get(274, 1)
    except Image.DecompressionBombError as exc:
        raise ValueError("Image dimensions exceed the local inspector's safety limit") from exc
    dpi = (
        [float(axis) for axis in raw_dpi]
        if isinstance(raw_dpi, (tuple, list)) and len(raw_dpi) == 2
        and all(math.isfinite(axis) and axis > 0 for axis in raw_dpi)
        else None
    )
    if orientation in {5, 6, 7, 8}:
        width, height = height, width
        if dpi is not None:
            dpi.reverse()
    facts = {"pixels": [width, height], "declared_dpi": dpi, "print_layout": layout}
    if kind not in _RASTER_RULES:
        return facts, [], ["Declare this raster artwork as photo or line to choose its resolution rule"]

    rule = _RASTER_RULES[kind]
    failures, reviews = [], []
    minimum_dpi = rule["minimum_dpi"]
    facts["minimum_dpi"] = minimum_dpi
    if dpi is None:
        reviews.append("Image has no usable DPI metadata; verify its effective resolution at print size")
    elif min(dpi) + 0.02 < minimum_dpi:
        # PNG stores whole pixels per metre and can round an exact DPI down.
        failures.append(f"Declared resolution falls below {minimum_dpi} dpi")
    if layout:
        minimum_width = rule[layout]
        facts["minimum_width_pixels"] = minimum_width
        if width < minimum_width:
            failures.append(f"Image width falls below {minimum_width} pixels for {layout}")
    elif width < rule["single_column"]:
        failures.append(f"Image width falls below the smallest permitted width of {rule['single_column']} pixels")
    elif width < rule["full_page"]:
        reviews.append("Declare the intended single_column or full_page print layout to check minimum width")
    return facts, failures, reviews


def _stage_figure(record, declaration, source_root, output_dir):
    if isinstance(declaration, str):
        declaration = {"path": declaration}
    if not isinstance(declaration, dict):
        record.update(status="missing", issues=["Declare the separate source file for this figure in figure_sources"])
        return
    try:
        source = _source_path(source_root, declaration.get("path"))
    except FileNotFoundError:
        record.update(status="missing", issues=["The declared figure source file could not be found"])
        return
    except ValueError as exc:
        record.update(status="failed", issues=[str(exc)])
        return

    suffix = source.suffix.lower()
    kind = declaration.get("kind", "vector" if suffix in _VECTOR_SUFFIXES else None)
    layout = declaration.get("layout")
    record.update(source=str(source), kind=kind)
    if (
        (suffix in _VECTOR_SUFFIXES and kind != "vector")
        or (suffix in _RASTER_FORMATS and kind not in (None, "photo", "line"))
        or layout not in (None, "single_column", "full_page")
    ):
        record.update(status="failed", issues=["Artwork type or print layout is inconsistent with its declared source"])
        return
    if suffix not in _VECTOR_SUFFIXES and suffix not in _RASTER_FORMATS:
        record.update(status="manual_review", issues=["Artwork staging supports PDF, EPS, TIFF, JPEG and PNG"])
        return
    try:
        data = source.read_bytes()
        if suffix == ".pdf":
            facts, failures, reviews = _inspect_pdf(data)
        elif suffix == ".eps":
            facts, failures, reviews = _inspect_eps(data)
        else:
            facts, failures, reviews = _inspect_raster(data, suffix, kind, layout)
    except ImportError:
        record.update(status="manual_review", issues=["An artwork inspection dependency is unavailable"])
        return
    except (ValueError, OSError, RuntimeError, SyntaxError) as exc:
        record.update(status="failed", issues=[f"Could not read the artwork file: {exc}"])
        return

    target = output_dir / f"Figure_{record['figure_number']}{suffix}"
    with target.open("xb") as stream:
        stream.write(data)
    status = "failed" if failures else "manual_review" if reviews else "verified"
    record.update(
        output=str(target), sha256=sha256(data).hexdigest(), checks=facts,
        status=status, issues=failures + reviews,
    )


def stage_artwork(placement: dict, metadata: dict | None, source_root: str | Path, output_dir: str | Path) -> dict:
    """Stage the declared artwork for figure numbers in an end-placement report.

    In metadata, set figure_sources to a mapping such as {"1": "response.pdf"}
    or {"2": {"path": "photo.png", "kind": "photo", "layout": "single_column"}}.
    Files are read relative to source_root and copied into a fresh output_dir.
    Checks cover artwork structure and declared raster resolution. Matching the
    source visually to the manuscript figure remains an editorial check.
    """
    if not placement.get("placement_verified") or not isinstance(placement.get("copies"), list):
        raise ValueError("Artwork staging requires a structural figure placement report")
    declarations = (metadata or {}).get("figure_sources", {})
    if not isinstance(declarations, dict) or any(not isinstance(key, str) for key in declarations):
        raise ValueError("figure_sources must be a mapping keyed by figure number strings")
    root = Path(source_root).resolve(strict=True)
    if not root.is_dir():
        raise ValueError("Artwork source root must be a directory")
    output = Path(output_dir).resolve()
    output.mkdir(parents=True, exist_ok=False)

    records = []
    for entry in placement["copies"]:
        if entry["kind"] != "figure":
            continue
        caption = entry["caption"]
        match = _FIGURE_NUMBER.match(without_bidi_controls(caption))
        records.append({"figure_number": match[1].upper() if match else None, "caption": caption})
    counts = Counter(record["figure_number"] for record in records)
    for record in records:
        number = record["figure_number"]
        if number is None or counts[number] != 1:
            record.update(status="manual_review", issues=["Figure caption does not give a unique supported figure number"])
        else:
            _stage_figure(record, declarations.get(number), root, output)
    unused = sorted(set(declarations) - {record["figure_number"] for record in records})
    return {
        "directory": str(output),
        "figures": records,
        "unused_declarations": unused,
        "all_figures_staged": all(record.get("output") for record in records),
        "all_figures_ready": not unused and all(record["status"] == "verified" for record in records),
        "visual_match_verified": False,
        "guide_pdf_page": 11,
    }
