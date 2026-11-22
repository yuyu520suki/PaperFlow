"""Verify template-owned chapter numbering and editable Word caption fields."""
from __future__ import annotations

import argparse
from copy import deepcopy
import json
from pathlib import Path
import re
import sys

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from lxml import etree
import pymupdf

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from paperflow.package import digest, q, read_package, text_run, write_package, xml
from paperflow.mathtype import audit_docx_math
from paperflow.pipeline import revise_master, run_pipeline


NUMBERING_ID = 43
CAPTION_STYLE = "Thesis Caption Probe"
CHAPTERS = (
    ("Alpha chapter", (
        ("Figure", "figAlphaA", "Alpha response"),
        ("Table", "tabAlphaA", "Alpha constants"),
        ("Figure", "figAlphaB", "Alpha comparison"),
    )),
    ("Beta chapter", (
        ("Table", "tabBetaA", "Beta constants"),
        ("Figure", "figBetaA", "Beta response"),
        ("Table", "tabBetaB", "Beta comparison"),
    )),
)
REFERENCE = re.compile(r"^\s*REF\s+(\S+)", re.I)
INDEX_LABEL = re.compile(r'\\c\s+"?(Figure|Table)"?(?=\s|$)', re.I)
CAPTION_SEQUENCE = re.compile(r'^\s*SEQ\s+"?(Figure|Table|Equation)"?(?=\s|$)', re.I)
RESET_AT_CHAPTER = re.compile(r"\\s\s+1(?=\s|$)", re.I)
# Word may wrap a computed chapter number in Unicode bidi controls.
BIDI_FORMATTING = dict.fromkeys((0x061C, 0x200E, 0x200F, *range(0x202A, 0x202F), *range(0x2066, 0x206A)))


def _element(tag, **attributes):
    element = OxmlElement(f"w:{tag}")
    for name, value in attributes.items():
        element.set(qn(f"w:{name}"), str(value))
    return element


def _make_numbering(document, *, first, paragraph_style, multilevel=False):
    numbering = document.part.numbering_part.element
    abstract = _element("abstractNum", abstractNumId=42)
    abstract.append(_element("multiLevelType", val="multilevel" if multilevel else "singleLevel"))
    level = _element("lvl", ilvl=0)
    level.extend((_element("start", val=first), _element("numFmt", val="decimal")))
    level.append(_element("pStyle", val=paragraph_style))
    level.extend((_element("suff", val="space"), _element("lvlText", val="%1"),
                  _element("lvlJc", val="left")))
    abstract.append(level)
    numbering.insert_element_before(abstract, "w:num")
    number = _element("num", numId=NUMBERING_ID)
    number.append(_element("abstractNumId", val=42))
    numbering.append(number)


def _set_numbering(properties, *, include_level=False):
    numbering = properties.get_or_add_numPr()
    if include_level:
        numbering.get_or_add_ilvl().val = 0
    numbering.get_or_add_numId().val = NUMBERING_ID


def _make_template(path):
    document = Document()
    document.styles["Normal"].font.name = "Times New Roman"
    document.styles["Normal"].font.size = Pt(11)
    heading = document.styles["Heading 1"]
    heading.paragraph_format.page_break_before = True
    _make_numbering(document, first=1, paragraph_style=heading.style_id, multilevel=True)
    _set_numbering(heading._element.get_or_add_pPr())
    caption = document.styles.add_style(CAPTION_STYLE, WD_STYLE_TYPE.PARAGRAPH)
    caption.base_style = document.styles["Normal"]
    caption.font.name = "Times New Roman"
    caption.font.size = Pt(9)
    caption.font.italic = True
    caption.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption.paragraph_format.space_after = Pt(4)
    document.add_paragraph("{{PAPERFLOW_BODY}}")
    document.save(path)
    return caption.style_id


def _make_source(path):
    document = Document()
    # These IDs deliberately collide with template numbering IDs. Source list
    # items start at seven, while headings should use the template list.
    list_base = document.styles.add_style("Source List Base Probe", WD_STYLE_TYPE.PARAGRAPH)
    list_base.base_style = document.styles["Normal"]
    _set_numbering(list_base._element.get_or_add_pPr())
    list_style = document.styles.add_style("Source List Derived Probe", WD_STYLE_TYPE.PARAGRAPH)
    list_style.base_style = list_base
    _make_numbering(document, first=7, paragraph_style=list_base.style_id)
    for title, token in (("Contents", "[[TOC]]"), ("List of Figures", "[[LOF]]"),
                         ("List of Tables", "[[LOT]]")):
        document.add_paragraph().add_run(title).bold = True
        document.add_paragraph(token)
    picture = PROJECT / "examples" / "computers-and-structures" / "response.png"
    for chapter, objects in CHAPTERS:
        heading = document.add_heading(chapter, level=1)
        _set_numbering(heading._p.get_or_add_pPr(), include_level=True)
        for kind, target, caption in objects:
            if kind == "Figure":
                document.add_picture(str(picture), width=Cm(6))
            document.add_paragraph(f"[[CAPTION {kind} {target}|{caption}]]")
            if kind == "Table":
                table = document.add_table(1, 2)
                table.cell(0, 0).text = caption
                table.cell(0, 1).text = "1.25"
            document.add_paragraph(f"See {kind} [[REF {target}]].")
        if chapter == CHAPTERS[0][0]:
            first_item = document.add_paragraph("Source list first item")
            _set_numbering(first_item._p.get_or_add_pPr(), include_level=True)
            document.add_paragraph("Source list second item", style=list_style)
            equation = document.add_paragraph()
            omml = OxmlElement("m:oMath")
            math_run = OxmlElement("m:r")
            math_text = OxmlElement("m:t")
            math_text.text = "x+1=2"
            math_run.append(math_text)
            omml.append(math_run)
            equation._p.append(omml)
            equation.add_run(" [[EQNO eqAlpha]]")
            document.add_paragraph("See Equation [[REF eqAlpha]].")
    document.save(path)


def _insert_chapter(master, revision):
    parts = read_package(master)
    document = xml(parts["word/document.xml"])
    body = document.find(q("w:body"))
    chapter_name = CHAPTERS[0][0]
    (existing,) = [p for p in body.findall(q("w:p"))
                   if "".join(t.text or "" for t in p.iter(q("w:t"))) == chapter_name]
    new = etree.Element(q("w:p"))
    new.append(deepcopy(existing.find(q("w:pPr"))))
    new.append(text_run("Prefatory chapter"))
    text = etree.Element(q("w:p"))
    text.append(text_run("This added chapter exercises native caption renumbering."))
    existing.addprevious(new)
    existing.addprevious(text)
    parts["word/document.xml"] = etree.tostring(document, encoding="UTF-8", xml_declaration=True)
    write_package(parts, revision)


def _check_stage(workflow, folder, *, offset):
    snapshot = next(step["evidence"]["snapshot"] for step in reversed(workflow["steps"])
                    if step["stage"] in {"word_fields", "word_fields_after_zotero"})
    fields = snapshot["fields"]
    chapters = []
    expected_refs = {}
    expected_indexes = {"Figure": [], "Table": []}
    for chapter_number, (chapter, objects) in enumerate(CHAPTERS, 1 + offset):
        chapters.append((chapter_number, chapter))
        counts = {"Figure": 0, "Table": 0}
        for kind, target, caption in objects:
            counts[kind] += 1
            number = f"{chapter_number}-{counts[kind]}"
            expected_refs[target] = number
            expected_indexes[kind].append(f"{kind} {number} {caption}")
    if offset:
        chapters.insert(0, (1, "Prefatory chapter"))
    equation_number = f"{1 + offset}-1"
    expected_refs["eqAlpha"] = equation_number

    references, indexes, toc = {}, {"Figure": [], "Table": []}, []
    for field in fields:
        result = field["result"].translate(BIDI_FORMATTING)
        if field["kind"] == "REF" and (match := REFERENCE.match(field["code"])):
            references[match[1]] = result.strip()
        if field["kind"] == "TOC":
            if match := INDEX_LABEL.search(field["code"]):
                indexes[match[1].title()].extend(result.splitlines())
            else:
                toc.extend(result.splitlines())
    # Each native list entry includes a tab followed by its resolved page.
    index_text = {kind: [row.rsplit("\t", 1)[0].strip() for row in rows if row.strip()]
                  for kind, rows in indexes.items()}
    expected_caption_text = {caption for rows in expected_indexes.values() for caption in rows}
    styles = [paragraph["style"] for paragraph in snapshot["paragraphs"]
              if paragraph["text"].translate(BIDI_FORMATTING).strip() in expected_caption_text]
    pages = []
    with pymupdf.open(folder / "editable-master.pdf") as pdf:
        for index, page in enumerate(pdf, 1):
            page.get_pixmap(dpi=130, alpha=False).save(folder / f"review-page-{index}.png")
            pages.append(" ".join(page.get_text().split()))
    rendered = "\n".join(pages)
    expected_page_rows = {"Figure": [], "Table": []}
    for kind, captions in expected_indexes.items():
        for caption in captions:
            locations = [index for index, page in enumerate(pages, 1) if index > 1 and caption in page]
            if len(locations) == 1:
                expected_page_rows[kind].append(f"{caption}\t{locations[0]}")
    cached_page_rows = {kind: [row.strip() for row in rows if row.strip()] for kind, rows in indexes.items()}
    prefixes = [field for field in fields if field["kind"] == "STYLEREF"]
    sequences = [field for field in fields if field["kind"] == "SEQ" and CAPTION_SEQUENCE.match(field["code"])]
    math = audit_docx_math(folder / "editable-master.docx")
    math_hashes = [item.get("mtef_sha256") for item in math["objects"]]
    checks = {
        "whole_number_refs": references == expected_refs,
        "chapter_local_figure_and_table_indexes": index_text == expected_indexes,
        "index_pages_match_rendered_captions": cached_page_rows == expected_page_rows,
        "editable_native_chapter_fields": (
            len(prefixes) == len(sequences) == len(expected_refs)
            and all(not field["locked"] for field in prefixes + sequences)
            and all(RESET_AT_CHAPTER.search(field["code"]) for field in sequences)
        ),
        "toc_uses_numbered_headings": all(any(re.match(rf"^\s*{number}\s+{re.escape(chapter)}\b", row)
                                                   for row in toc) for number, chapter in chapters),
        "custom_caption_style": styles == [CAPTION_STYLE] * len(expected_caption_text),
        "rendered_chapter_numbers": all(any(re.search(rf"(?:^|\s){number}\s+{re.escape(chapter)}\b", page)
                                              for page in pages[1:]) for number, chapter in chapters),
        "rendered_chapter_captions": all(any(caption in page for page in pages[1:])
                                         for rows in expected_indexes.values() for caption in rows),
        "source_list_numbers_preserved": all(re.search(rf"\b{number}\s+Source list {label} item\b", rendered)
                                              for number, label in ((7, "first"), (8, "second"))),
        "native_mathtype_equation": math["native_mathtype_objects"] == math["mathtype_objects"] == 1
                                    and not math["omml_equations"] and bool(math_hashes[0]),
        "rendered_equation_number_and_ref": any(
            re.search(rf"\(\s*{equation_number}\s*\)", page) and f"See Equation {equation_number}." in page
            for page in pages[1:]
        ),
    }
    return {"checks": checks, "references": references, "index_rows": indexes, "contents": toc,
            "caption_styles": styles, "mathtype_mtef_sha256": math_hashes, "pages": pages}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    output = parser.parse_args().output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    source, template = output / "manuscript.docx", output / "template.docx"
    style_id = _make_template(template)
    _make_source(source)
    originals = {path: digest(path.read_bytes()) for path in (source, template)}
    built = output / "built"
    initial_workflow = run_pipeline(source, template, built, caption_chapter_level=1,
                                    role_style_map={"caption": style_id})
    initial = _check_stage(initial_workflow, built, offset=0)
    master = built / "editable-master.docx"
    revision = output / "revised-input.docx"
    _insert_chapter(master, revision)
    originals.update({path: digest(path.read_bytes()) for path in (master, revision)})
    revised = output / "refreshed"
    revised_workflow = revise_master(revision, revised)
    final = _check_stage(revised_workflow, revised, offset=1)
    checks = {
        **{f"initial_{key}": value for key, value in initial["checks"].items()},
        **{f"revised_{key}": value for key, value in final["checks"].items()},
        "same_mathtype_payload_after_renumbering": initial["mathtype_mtef_sha256"] == final["mathtype_mtef_sha256"],
        "all_inputs_unchanged": all(digest(path.read_bytes()) == value for path, value in originals.items()),
    }
    evidence = output / "chapter-captions-evidence.json"
    report = {"pass": all(checks.values()), "checks": checks, "initial": initial, "revised": final}
    evidence.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"pass": report["pass"], "checks": checks, "evidence": str(evidence)}, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
