"""Exercise real Word insertion and rendering with two dissimilar sections."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

import pymupdf
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from paperflow.formatting import normalize_to_template
from paperflow.package import digest, field_runs
from paperflow.template import DEFAULT_SLOT, inventory_template, prepare_template
from paperflow.word import build_from_template, export_pdf, private_word, refresh_fields, snapshot


HEADING = "Imported Mechanics Heading"
BODY = "Imported article text for the landscape section."
SPLIT_BODY = "Imported second source section after an ordinary page boundary."
CLOSING = "Landscape Closing Text"


def _furniture(section, label):
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False
    section.header.paragraphs[0].text = f"{label} Header"
    footer = section.footer.paragraphs[0]
    footer.text = f"{label} Footer "
    for run in field_runs(" PAGE ", "0"):
        footer._p.append(run)


def _restart_page_number(section, number):
    page_numbers = OxmlElement("w:pgNumType")
    page_numbers.set(qn("w:start"), str(number))
    section._sectPr.insert_element_before(page_numbers, "w:cols", "w:docGrid")


def _make_template(path, *, bookmark=False):
    document = Document()
    for name, points in (("Normal", 12), ("Title", 22), ("Heading 1", 16)):
        style = document.styles[name]
        font = style.font
        font.name = "Times New Roman"
        font.size = Pt(points)
        font.color.rgb = RGBColor(0, 0, 0)
        fonts = style._element.find("w:rPr/w:rFonts", namespaces=style._element.nsmap)
        for attribute in ("asciiTheme", "hAnsiTheme"):
            fonts.attrib.pop(qn("w:" + attribute), None)
        border = style._element.find("w:pPr/w:pBdr", namespaces=style._element.nsmap)
        if border is not None:
            border.getparent().remove(border)
    cover = document.sections[0]
    cover.page_width, cover.page_height = Inches(8.5), Inches(11)
    cover.left_margin = Inches(1)
    _furniture(cover, "Cover")
    document.add_paragraph("Mechanics Template Cover", style="Title")
    document.add_paragraph("Portrait Cover Text")

    article = document.add_section(WD_SECTION.NEW_PAGE)
    article.orientation = WD_ORIENT.LANDSCAPE
    article.page_width, article.page_height = Inches(11), Inches(8.5)
    article.left_margin, article.right_margin = Inches(1.25), Inches(0.75)
    article.top_margin, article.bottom_margin = Inches(0.8), Inches(0.7)
    _furniture(article, "Landscape")
    _restart_page_number(article, 7)
    document.add_paragraph("Landscape Context")
    closing = document.add_paragraph()
    if bookmark:
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:name"), "manuscript")
        start.set(qn("w:id"), "19")
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), "19")
        closing._p.extend((start, end))
    closing.add_run(CLOSING)
    document.save(path)


def _make_source(path, *, split_source=False):
    document = Document()
    document.styles["Normal"].font.name = "Calibri"
    document.styles["Normal"].font.size = Pt(10)
    document.styles["Heading 1"].font.name = "Calibri"
    document.styles["Heading 1"].font.size = Pt(30)
    document.add_heading(HEADING, level=1)
    run = document.add_paragraph().add_run(BODY)
    run.font.name = "Calibri"
    run.font.size = Pt(25)
    run.italic = True
    if split_source:
        _furniture(document.sections[0], "Source First")
        following = document.add_section(WD_SECTION.NEW_PAGE)
        following.left_margin = Inches(2.5)
        _furniture(following, "Source Second")
        _restart_page_number(following, 19)
        document.add_paragraph(SPLIT_BODY)
    document.save(path)


def _rendered_pages(path, output):
    pages = []
    with pymupdf.open(path) as document:
        for number, page in enumerate(document, start=1):
            image_path = output / f"page-{number}.png"
            page.get_pixmap(dpi=144, alpha=False).save(image_path)
            pages.append({"page": number, "width_points": page.rect.width, "height_points": page.rect.height,
                          "text": page.get_text(), "png": str(image_path)})
    return pages


def _check_native(before, after, pages, *, split_source):
    heading = next((p for p in after["paragraphs"] if p["text"].strip() == HEADING), {})
    body = next((p for p in after["paragraphs"] if p["text"].strip() == BODY), {})
    checks = {
        "section_geometry_matches_template": before["sections"] == after["sections"],
        "portrait_then_landscape": [s["Orientation"] for s in after["sections"]] == [0, 1],
        "heading_style_and_font": heading.get("outline_level") == 1 and heading.get("font", {}).get("Name") == "Times New Roman"
            and heading.get("font", {}).get("Size") == 16,
        "body_font": body.get("font", {}).get("Name") == "Times New Roman" and body.get("font", {}).get("Size") == 12,
    }
    expected_pages = 3 if split_source else 2
    checks["expected_rendered_page_count"] = len(pages) == expected_pages
    if len(pages) == expected_pages:
        cover, landscape, *continued = pages
        checks.update({
            "rendered_page_geometry": cover["width_points"] < cover["height_points"]
                and all(page["width_points"] > page["height_points"] for page in (landscape, *continued)),
            "cover_furniture_and_text": all(token in cover["text"] for token in (
                "Cover Header", "Cover Footer 1", "Mechanics Template Cover", "Portrait Cover Text",
            )) and all(token not in cover["text"] for token in ("Landscape Header", HEADING, BODY)),
            "landscape_furniture_and_text": all(token in landscape["text"] for token in (
                "Landscape Header", "Landscape Footer 7", "Landscape Context", HEADING, BODY,
            )) and all(token not in landscape["text"] for token in ("Cover Header", "Portrait Cover Text", DEFAULT_SLOT)),
        })
        if split_source:
            checks["second_source_section_renders_on_continued_template_page"] = all(
                token in continued[0]["text"] for token in ("Landscape Header", "Landscape Footer 8", SPLIT_BODY, CLOSING)
            ) and SPLIT_BODY not in landscape["text"] and BODY not in continued[0]["text"]
            checks["source_headers_and_footers_absent"] = all(
                label not in page["text"] for page in pages for label in ("Source First", "Source Second")
            )
        else:
            checks["closing_text_in_landscape_section"] = CLOSING in landscape["text"]
    return checks


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument("--split-source", action="store_true", help="Import two source sections into the same template slot")
    args = parser.parse_args()
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    source = output / "source.docx"
    _make_source(source, split_source=args.split_source)
    source_hash = digest(source.read_bytes())
    cases = {}
    with private_word() as app:
        for name, slot in (("marker", DEFAULT_SLOT), ("bookmark", "bookmark:manuscript")):
            print(f"Checking {name} slot in Word {app.Version}", flush=True)
            folder = output / name
            folder.mkdir()
            original = folder / "original-template.docx"
            _make_template(original, bookmark=name == "bookmark")
            original_hash = digest(original.read_bytes())
            template = original
            preparation = None
            if name == "marker":
                template = folder / "prepared-template.docx"
                inventory = inventory_template(original)
                context = next(anchor for anchor in inventory["paragraph_anchors"] if anchor["text_preview"] == "Landscape Context")
                preparation = prepare_template(original, template, after_paragraph=context["body_child_index"],
                                               expected_sha256=inventory["sha256"])
            template_hash = digest(template.read_bytes())
            normalized = folder / "normalized.docx"
            normalization = normalize_to_template(source, template, normalized, slot=slot)
            built = folder / "inserted.docx"
            insertion = build_from_template(template, normalized, built, slot=slot, app=app)
            final = folder / "word-refreshed.docx"
            refresh_fields(built, final, app=app)
            after = snapshot(final, app=app)
            pdf = folder / "word-render.pdf"
            export = export_pdf(final, pdf, app=app)
            pages = _rendered_pages(pdf, folder)
            checks = _check_native(insertion["before"], after, pages, split_source=args.split_source)
            checks.update({
                "source_unchanged": digest(source.read_bytes()) == source_hash,
                "original_template_unchanged": digest(original.read_bytes()) == original_hash,
                "build_template_unchanged": digest(template.read_bytes()) == template_hash,
                "slot_resolved_to_second_section": normalization["template_location"]["section_index"] == 2,
            })
            if args.split_source:
                checks["source_break_converted_to_ordinary_page_boundary"] = normalization["source_section_breaks"] == [
                    {"after_source_section": 1, "source_start_type": "nextPage", "conversion": "page_break_before"},
                ]
            cases[name] = {"pass": all(checks.values()), "checks": checks, "prepared": preparation,
                           "normalization": normalization, "native_before": insertion["before"], "native_after": after,
                           "word_export": export, "pages": pages}
    report = {"pass": all(case["pass"] for case in cases.values()), "split_source": args.split_source, "cases": cases}
    report_path = output / "multisection-evidence.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"evidence": str(report_path), "checks": {name: case["checks"] for name, case in cases.items()}}, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
