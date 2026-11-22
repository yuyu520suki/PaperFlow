"""Verify ordinary inline-picture fitting through the normal build and revision paths.

Run with the project's Python and a new output directory. This requires Word and
MathType; it uses no Zotero items. PNG previews cover every exported PDF page.
The evidence proves geometric fitting and native preservation, not human visual QA.
"""
from __future__ import annotations

import argparse
from collections import Counter
import io
import json
from pathlib import Path
import posixpath
import re
import sys

from docx import Document
from docx.shared import Inches, Pt
from lxml import etree as ET
from PIL import Image, ImageChops, ImageStat
import pymupdf

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from paperflow.mathtype import audit_docx_math
from paperflow.package import digest, q, read_package, xml
from paperflow.pipeline import revise_master, run_pipeline


ASSET = PROJECT / "examples/computers-and-structures/response.png"
BODY = "Body typography witness for the inline picture fixture."
WIDE_ALT = "PaperFlow wide 9in fixture"
SMALL_ALT = "PaperFlow small 2in fixture"
WIDE_CAPTION = "Wide response witness"
SMALL_CAPTION = "Small response witness"
DRAWING_NS = {
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
}
EMU_PER_PT = 12700
TOLERANCE_PT = 0.1


def _make_template(path):
    document = Document()
    section = document.sections[0]
    section.page_width, section.page_height = Inches(7), Inches(11)
    section.left_margin = section.right_margin = Inches(1)
    section.top_margin = section.bottom_margin = Inches(1)
    for name, font, size in (("Normal", "Calibri", 11), ("Heading 1", "Arial", 14),
                             ("Caption", "Arial", 9)):
        document.styles[name].font.name = font
        document.styles[name].font.size = Pt(size)
    document.styles["Normal"].paragraph_format.space_after = Pt(5)
    document.add_paragraph("{{PAPERFLOW_BODY}}")
    document.save(path)


def _make_source(path):
    document = Document()
    section = document.sections[0]
    section.page_width, section.page_height = Inches(11), Inches(11)
    section.left_margin = section.right_margin = Inches(0.5)
    document.styles["Normal"].font.name = "Times New Roman"
    document.styles["Normal"].font.size = Pt(15)
    document.add_heading("Inline picture fitting fixture", level=1)
    document.add_paragraph(BODY)
    document.add_paragraph("See Figure [[REF WidePicture]] and Figure [[REF SmallPicture]].")
    for width, label, alt, caption in ((9, "WidePicture", WIDE_ALT, WIDE_CAPTION),
                                       (2, "SmallPicture", SMALL_ALT, SMALL_CAPTION)):
        shape = document.add_picture(str(ASSET), width=Inches(width))
        shape._inline.docPr.set("descr", alt)
        document.add_paragraph(f"[[CAPTION Figure {label}|{caption}]]")
    paragraph = document.add_paragraph("Editable formula witness: ")
    math = ET.SubElement(paragraph._p, q("m:oMath"))
    ET.SubElement(ET.SubElement(math, q("m:r")), q("m:t")).text = "x+1=2"
    document.save(path)


def _decoded_identity(data):
    with Image.open(io.BytesIO(data)) as image:
        canonical = image.convert("RGBA")
        return {"pixels": list(canonical.size), "rgba_sha256": digest(canonical.tobytes())}


def _pictures(path):
    parts = read_package(path)
    document = xml(parts["word/document.xml"])
    relationships = {item.get("Id"): item for item in xml(parts["word/_rels/document.xml.rels"])}
    pictures = []
    for inline in document.findall(".//wp:inline", DRAWING_NS):
        picture = inline.find("a:graphic/a:graphicData/pic:pic", DRAWING_NS)
        if picture is None:
            continue
        extent = inline.find("wp:extent", DRAWING_NS)
        blip = picture.find(".//a:blip", DRAWING_NS)
        relation = relationships[blip.get(q("r:embed"))]
        target = relation.get("Target")
        part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("word", target))
        properties = inline.find("wp:docPr", DRAWING_NS)
        pictures.append({
            "alt": properties.get("descr", ""), "part": part,
            "width_pt": int(extent.get("cx")) / EMU_PER_PT,
            "height_pt": int(extent.get("cy")) / EMU_PER_PT,
            "encoded_sha256": digest(parts[part]), **_decoded_identity(parts[part]),
        })
    return pictures


def _snapshot(workflow):
    for step in reversed(workflow["steps"]):
        if step["stage"] in {"word_fields", "word_fields_after_zotero"}:
            return step["evidence"]["snapshot"]
    raise RuntimeError("No native Word field-refresh snapshot was recorded")


def _pdf_picture_match(pdf, image, reference):
    """Match decoded visual content after Word's PDF image downsampling.

    The exact decoded identity assertion lives at the DOCX boundary. PDF export
    resamples images, so require both the source aspect ratio and a strict RGB
    thumbnail match here. A formula raster or an empty rectangle cannot satisfy
    the source-image comparison merely by occupying the expected rectangle.
    """
    expected_ratio = reference.width / reference.height
    ratio = image["width"] / image["height"]
    result = {"xref": image["xref"], "pixels": [image["width"], image["height"]],
              "bbox_pt": list(image["bbox"]), "matched": False,
              "relative_aspect_error": abs(ratio / expected_ratio - 1)}
    if result["relative_aspect_error"] > 0.01:
        result["reason"] = "different source aspect ratio"
        return result
    if not image["xref"]:
        result["reason"] = "no extractable PDF image object"
        return result
    pixmap = pymupdf.Pixmap(pdf, image["xref"])
    if pixmap.colorspace != pymupdf.csRGB:
        pixmap = pymupdf.Pixmap(pymupdf.csRGB, pixmap)
    if pixmap.alpha:
        pixmap = pymupdf.Pixmap(pixmap, 0)
    decoded = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
    thumbnail = decoded.resize(reference.size, Image.Resampling.LANCZOS)
    difference = ImageStat.Stat(ImageChops.difference(reference, thumbnail))
    result.update(rgb_sha256=digest(decoded.tobytes()), thumbnail_size=list(reference.size),
                  mean_absolute_rgb_error=difference.mean, root_mean_square_rgb_error=difference.rms,
                  max_mean_absolute_error_allowed=1.0, max_root_mean_square_error_allowed=2.0)
    result["matched"] = max(difference.mean) <= 1.0 and max(difference.rms) <= 2.0
    return result


def _render(pdf_path, output, asset=ASSET, *, render_pages=True):
    if render_pages:
        output.mkdir(exist_ok=True)
    with Image.open(asset) as image:
        rgba = image.convert("RGBA")
        reference = Image.alpha_composite(Image.new("RGBA", rgba.size, "white"), rgba).convert("RGB")
        reference = reference.resize((128, 64), Image.Resampling.LANCZOS)
    pages = []
    with pymupdf.open(pdf_path) as pdf:
        for index, page in enumerate(pdf, 1):
            preview = output / f"page-{index:03d}.png"
            if render_pages:
                page.get_pixmap(dpi=130, colorspace=pymupdf.csRGB, alpha=False).save(preview)
            matches = [_pdf_picture_match(pdf, image, reference)
                       for image in page.get_image_info(hashes=True, xrefs=True)]
            pages.append({
                "page": index, "size_pt": [page.rect.width, page.rect.height],
                "preview": str(preview), "text": " ".join(page.get_text().split()),
                "picture_rects_pt": [image["bbox_pt"] for image in matches if image["matched"]],
                "pdf_image_content_matches": matches,
            })
    return pages


def _near(left, right):
    return abs(left - right) <= TOLERANCE_PT


def _same_geometry(left, right):
    return len(left) == len(right) and all(
        _near(a[dimension], b[dimension]) for a, b in zip(left, right, strict=True)
        for dimension in ("width_pt", "height_pt")
    )


def _identities(pictures):
    # Word can rename, deduplicate, or re-encode media. Compare decoded pixels
    # for every drawing occurrence, not ZIP names or compressed media bytes.
    return Counter((tuple(item["pixels"]), item["rgba_sha256"]) for item in pictures)


def _math_ok(audit):
    return audit["native_mathtype_objects"] == 1 and audit["all_mathtype_objects_valid"] and not audit["omml_equations"]


def _native_fields_ok(snapshot):
    fields = [(item["code"], item["result"].strip()) for item in snapshot["fields"]]
    return (sorted(value for code, value in fields if re.match(r"\s*SEQ\s+Figure\b", code, re.I)) == ["1", "2"]
            and all(any(re.match(rf"\s*REF\s+{label}\b", code, re.I) and value == number
                        for code, value in fields) for label, number in (("WidePicture", "1"), ("SmallPicture", "2"))))


def _native_pictures_ok(snapshot, pictures):
    shapes = [item for item in snapshot["inline_shapes"] if item["type"] == 3]
    return len(shapes) == len(pictures) == 2 and all(
        _near(shape["width"], picture["width_pt"]) and _near(shape["height"], picture["height_pt"])
        for shape, picture in zip(shapes, pictures, strict=True)
    )


def _native_template_ok(snapshot):
    return (bool(snapshot["sections"]) and all(
        _near(section["PageWidth"], 504) and _near(section["LeftMargin"], 72)
        and _near(section["RightMargin"], 72) for section in snapshot["sections"])
        and any(BODY in paragraph["text"] and paragraph["font"]["Name"] == "Calibri"
                and _near(paragraph["font"]["Size"], 11) for paragraph in snapshot["paragraphs"]))


def _pdf_ok(pages, pictures):
    rectangles = [rect for page in pages for rect in page["picture_rects_pt"]]
    text = " ".join(page["text"] for page in pages)
    return (len(rectangles) == len(pictures) == 2
            and all(_near(page["size_pt"][0], 504) for page in pages)
            and all(rect[0] >= 72 - TOLERANCE_PT and rect[2] <= 432 + TOLERANCE_PT for rect in rectangles)
            and all(_near(rect[2] - rect[0], picture["width_pt"])
                    and _near(rect[3] - rect[1], picture["height_pt"])
                    for rect, picture in zip(rectangles, pictures, strict=True))
            and "See Figure 1 and Figure 2." in text
            and f"Figure 1 {WIDE_CAPTION}" in text and f"Figure 2 {SMALL_CAPTION}" in text)


def _recheck(output):
    """Reanalyze existing PDF artifacts without invoking Word or editing them."""
    evidence = output / "inline-picture-fit-evidence.json"
    previous = evidence.read_bytes()
    report = json.loads(previous)
    for path, expected in report["reproduce"]["input_sha256"].items():
        if digest(Path(path).read_bytes()) != expected:
            raise RuntimeError(f"Cannot reuse evidence: original input changed: {path}")
    source = Path(report["reproduce"]["source"])
    if _pictures(source) != report["source_pictures"]:
        raise RuntimeError("Source picture state differs from the saved native-run evidence")
    artifacts = []
    for directory, key in (("built", "built"), ("revised", "revised")):
        folder = output / directory
        master, pdf, workflow = (folder / name for name in ("editable-master.docx", "editable-master.pdf", "workflow.json"))
        artifacts.extend((master, pdf, workflow))
        saved = json.loads(workflow.read_text(encoding="utf-8"))
        if (not saved["native_build_finished"] or _snapshot(saved) != report[f"{key}_native_snapshot"]
                or _pictures(master) != report[f"{key}_pictures"]
                or audit_docx_math(master) != report[f"{key}_mathtype"]):
            raise RuntimeError(f"Saved {directory} native evidence does not describe the current artifacts")
    hashes = {str(path): digest(path.read_bytes()) for path in artifacts}
    report["built_render"] = _render(output / "built/editable-master.pdf", output / "built-pages", render_pages=False)
    report["revised_render"] = _render(output / "revised/editable-master.pdf", output / "revised-pages", render_pages=False)
    report["checks"]["word_pdf_picture_bounds_dimensions_and_references_valid"] = (
        _pdf_ok(report["built_render"], report["built_pictures"])
        and _pdf_ok(report["revised_render"], report["revised_pictures"]))
    report["checks"]["all_pdf_pages_rendered"] = bool(report["built_render"]) and bool(report["revised_render"]) and all(
        Path(page["preview"]).is_file() for page in report["built_render"] + report["revised_render"])
    if any(digest(Path(path).read_bytes()) != sha for path, sha in hashes.items()):
        raise RuntimeError("A native document, PDF, or workflow changed during read-only reanalysis")
    report["pdf_reanalysis"] = {"native_word_rerun": False, "artifact_sha256": hashes,
                                "method": "decoded RGB source-image thumbnail match after PDF downsampling"}
    report["pass"] = all(report["checks"].values())
    backup = evidence.with_name("inline-picture-fit-evidence.before-pdf-recheck.json")
    if not backup.exists():
        with backup.open("xb") as stream:
            stream.write(previous)
    temporary = evidence.with_suffix(".recheck.tmp")
    temporary.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(evidence)
    print(json.dumps({"pass": report["pass"], "checks": report["checks"], "evidence": str(evidence),
                      "previous_evidence": str(backup), "native_word_rerun": False}, indent=2))
    return 0 if report["pass"] else 1


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument("--recheck", action="store_true", help="Reanalyze existing PDFs without running Word")
    args = parser.parse_args()
    output = args.output_dir.resolve()
    if args.recheck:
        return _recheck(output)
    output.mkdir(parents=True, exist_ok=False)
    source, template = output / "source.docx", output / "template.docx"
    _make_template(template)
    _make_source(source)
    originals = {path: digest(path.read_bytes()) for path in (source, template, ASSET)}
    before = _pictures(source)
    built = output / "built"
    build = run_pipeline(source, template, built)
    master = built / "editable-master.docx"
    master_hash = digest(master.read_bytes())
    after, math_after = _pictures(master), audit_docx_math(master)
    native_after = _snapshot(build)
    render_after = _render(built / "editable-master.pdf", output / "built-pages")
    revised = output / "revised"
    revision = revise_master(master, revised)
    final = revised / "editable-master.docx"
    stable, math_stable = _pictures(final), audit_docx_math(final)
    native_stable = _snapshot(revision)
    render_stable = _render(revised / "editable-master.pdf", output / "revised-pages")
    checks = {
        "normal_native_build_and_revision_finished": build["native_build_finished"] and revision["native_build_finished"],
        "source_has_large_9in_and_small_2in_picture": len(before) == 2 and _near(before[0]["width_pt"], 648)
                                                       and _near(before[1]["width_pt"], 144),
        "large_picture_fitted_to_5in_template_width": len(after) == 2 and _near(after[0]["width_pt"], 360),
        "small_picture_not_enlarged_or_resized": len(after) == 2 and _same_geometry(before[1:], after[1:]),
        "picture_aspect_ratios_preserved": len(before) == len(after) == 2 and all(
            abs(a["width_pt"] / a["height_pt"] - b["width_pt"] / b["height_pt"]) < 0.001
            for a, b in zip(before, after, strict=True)),
        "decoded_media_identity_preserved": _identities(before) == _identities(after) == _identities(stable),
        "revision_picture_geometry_stable": _same_geometry(after, stable),
        "native_word_picture_sizes_match_saved_package": _native_pictures_ok(native_after, after)
                                                        and _native_pictures_ok(native_stable, stable),
        "native_template_page_margins_and_body_font_adopted": _native_template_ok(native_after)
                                                              and _native_template_ok(native_stable),
        "dynamic_figure_captions_and_refs_valid": _native_fields_ok(native_after) and _native_fields_ok(native_stable),
        "native_mathtype_equation_valid_and_revision_payload_preserved": _math_ok(math_after) and _math_ok(math_stable)
            and Counter(item["mtef_sha256"] for item in math_after["objects"])
            == Counter(item["mtef_sha256"] for item in math_stable["objects"]),
        "word_pdf_picture_bounds_dimensions_and_references_valid": _pdf_ok(render_after, after) and _pdf_ok(render_stable, stable),
        "all_pdf_pages_rendered": bool(render_after) and bool(render_stable)
                                  and all(Path(page["preview"]).is_file() for page in render_after + render_stable),
        "source_template_asset_and_revision_input_unchanged": all(digest(path.read_bytes()) == sha for path, sha in originals.items())
                                                               and digest(master.read_bytes()) == master_hash,
    }
    report = {
        "pass": all(checks.values()), "checks": checks, "visual_reviewed": False,
        "reproduce": {"script": str(Path(__file__).resolve()), "source": str(source), "template": str(template),
                      "asset": str(ASSET), "input_sha256": {str(path): sha for path, sha in originals.items()}},
        "source_pictures": before, "built_pictures": after, "revised_pictures": stable,
        "built_native_snapshot": native_after, "revised_native_snapshot": native_stable,
        "built_mathtype": math_after, "revised_mathtype": math_stable,
        "built_render": render_after, "revised_render": render_stable,
        "workflows": [str(built / "workflow.json"), str(revised / "workflow.json")],
    }
    evidence = output / "inline-picture-fit-evidence.json"
    evidence.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"pass": report["pass"], "checks": checks, "evidence": str(evidence)}, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
