"""Exercise template formatting on real Word footnotes and endnotes."""
from __future__ import annotations

import argparse
from collections import Counter
import json
from pathlib import Path
import re
import sys

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.shared import Pt, RGBColor
from lxml import etree as ET
import pymupdf

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from scripts.zotero_fixture import SECOND_DOI, fixture_key
from paperflow.citations import _package_state, _payload
from paperflow.journal import CNS_STYLE
from paperflow.mathtype import MATHTYPE_PROGID, MathTypeSDK, audit_docx_math
from paperflow.package import NS, digest, field_runs, q, read_package, text_run, without_bidi_controls, xml
from paperflow.pipeline import run_pipeline
from paperflow.word import DOCX_FORMAT, _opened, export_pdf, private_word
from paperflow.zotero import ZoteroClient, document_preferences, make_bibliography_field, make_citation_field, set_docx_preferences


BODY = "Manuscript body typography"
NOTE_KINDS = {
    "footnote": ("Footnotes", CT.WML_FOOTNOTES, RT.FOOTNOTES, "Footnote Text", "Footnote Reference",
                 "Arial", 10, 8, "Primary footnote witness", 3),
    "endnote": ("Endnotes", CT.WML_ENDNOTES, RT.ENDNOTES, "Endnote Text", "Endnote Reference",
                "Calibri", 11, 9, "Primary endnote witness", 5),
}
MARKER_COLOR = "165380"
FONT_PROPERTIES = ("NameAscii", "Size", "Italic", "Bold", "Superscript", "Color")
NOTE_EQUATIONS = {"footnote": "z-3=1", "endnote": "w+2=7"}


def _style(document, name, style_type):
    try:
        return document.styles[name]
    except KeyError:
        return document.styles.add_style(name, style_type, builtin=not name.startswith("Fixture "))


def _typography(style, font, size, *, color=None):
    style.font.name = font
    style.font.size = Pt(size)
    if color:
        style.font.color.rgb = RGBColor.from_string(color)


def _child(parent, name, value=None, **attributes):
    if value is not None:
        attributes["val"] = value
    return ET.SubElement(parent, q("w:" + name), {q("w:" + key): str(data) for key, data in attributes.items()})


def _reference(paragraph, tag, style_id, identifier=None):
    run = _child(paragraph, "r")
    _child(_child(run, "rPr"), "rStyle", style_id)
    _child(run, tag, **({"id": identifier} if identifier is not None else {}))


def _equation(paragraph, value):
    math = ET.SubElement(paragraph, q("m:oMath"))
    ET.SubElement(ET.SubElement(math, q("m:r")), q("m:t")).text = value


def _settings(document, kind, start):
    settings = document.settings.element
    properties = ET.Element(q("w:" + kind + "Pr"))
    _child(properties, "numFmt", "decimal")
    _child(properties, "numStart", start)
    for identifier in (-1, 0):
        _child(properties, kind, id=identifier)
    settings.insert_element_before(properties, "w:compat")
    section = document.sections[0]._sectPr
    numbering = ET.Element(q("w:" + kind + "Pr"))
    _child(numbering, "numFmt", "decimal")
    _child(numbering, "numStart", start)
    section.insert_element_before(numbering, "w:type", "w:pgSz", "w:pgMar")


def _make_note(document, kind, *, template, item=None, note_equation=False, markers=False):
    _, content_type, relationship_type, text_style, reference_style, font, size, marker_size, witness, start = NOTE_KINDS[kind]
    note_style = _style(document, text_style, WD_STYLE_TYPE.PARAGRAPH)
    note_style.base_style = document.styles["Normal"]
    note_style.paragraph_format.space_after = Pt(3 if template else 13)
    _typography(note_style, font if template else "Georgia", size if template else 17)
    if not template:
        custom_note = _style(document, f"Fixture Imported {kind.capitalize()}", WD_STYLE_TYPE.PARAGRAPH)
        custom_note.base_style = note_style
        note_style = custom_note
    marker_style = _style(document, reference_style, WD_STYLE_TYPE.CHARACTER)
    _typography(marker_style, font if template else "Georgia", marker_size if template else 16,
                color=MARKER_COLOR if template else "AA3333")
    marker_style.font.superscript = True
    marker_style.font.bold = template

    settings_start = start if template else start + 8
    _settings(document, kind, settings_start)
    body_anchor = document.add_paragraph(f"{'Template' if template else 'Imported'} {kind} anchor")
    _reference(body_anchor._p, kind + "Reference", marker_style.style_id, 1)

    root = ET.Element(q("w:" + kind + "s"), nsmap={key: NS[key] for key in ("w", "r", "m")})
    for identifier, role in ((-1, "separator"), (0, "continuationSeparator")):
        special = _child(root, kind, id=identifier, type=role)
        _child(_child(_child(special, "p"), "r"), role)
    note = _child(root, kind, id=1)
    paragraph = _child(note, "p")
    properties = _child(paragraph, "pPr")
    _child(properties, "pStyle", note_style.style_id)
    _reference(paragraph, kind + "Ref", marker_style.style_id)

    path = PackURI(f"/word/{kind}s.xml")
    part = Part(path, content_type, b"", document.part.package)
    if template:
        paragraph.append(text_run(f" Template {kind} exemplar"))
    else:
        _child(properties, "spacing", after=300, line=320, lineRule="auto")
        explicit = ET.Element(q("w:rPr"))
        _child(explicit, "rFonts", ascii="Georgia", hAnsi="Georgia")
        _child(explicit, "i")
        _child(explicit, "sz", 40)
        paragraph.append(text_run(" " + witness, explicit))
        paragraph.append(text_run(". Cited item "))
        if markers:
            paragraph.append(text_run(f"[[CITE {fixture_key(SECOND_DOI)}]]. Table [[REF NoteTable]], page [[PAGEREF NoteTable]]"))
        else:
            paragraph.extend(field_runs(make_citation_field([item], citation_id=f"paperflow{kind}", note_index=1)))
        paragraph.append(text_run(". "))
        url = f"https://example.org/paperflow/{kind}"
        link = ET.SubElement(paragraph, q("w:hyperlink"), {q("r:id"): part.relate_to(url, RT.HYPERLINK, is_external=True)})
        link_style = ET.Element(q("w:rPr"))
        _child(link_style, "rStyle", _style(document, "Hyperlink", WD_STYLE_TYPE.CHARACTER).style_id)
        link.append(text_run(f"{kind.capitalize()} external link", link_style))
        if note_equation:
            paragraph.append(text_run(". Note equation: "))
            _equation(paragraph, NOTE_EQUATIONS[kind])
    part._blob = ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    document.part.relate_to(part, relationship_type)


def _create(path, *, template, item=None, note_equation=False, markers=False):
    document = Document()
    _typography(document.styles["Normal"], "Arial" if template else "Georgia", 13 if template else 18)
    link = _style(document, "Hyperlink", WD_STYLE_TYPE.CHARACTER)
    link.font.underline = True
    link.font.color.rgb = RGBColor.from_string("16709B" if template else "AA3333")
    for kind in NOTE_KINDS:
        _make_note(document, kind, template=template, item=item, note_equation=note_equation, markers=markers)
    if template:
        document.add_paragraph("{{PAPERFLOW_BODY}}")
    else:
        paragraph = document.add_paragraph(BODY + ". Body equation: ")
        _equation(paragraph._p, "x+1=2")
        if markers:
            document.add_paragraph("[[CAPTION Table NoteTable|Note reference witness]]")
            table = document.add_table(rows=1, cols=2)
            table.cell(0, 0).text = "Measured quantity"
            table.cell(0, 1).text = "Value"
        document.add_heading("References", 1)
        if markers:
            document.add_paragraph("[[BIBLIOGRAPHY]]")
        else:
            document.add_paragraph()._p.extend(field_runs(make_bibliography_field()))
    document.save(path)
    if not template:
        set_docx_preferences(path, document_preferences(CNS_STYLE, locale="en-US"))


def _font(word_range):
    font = word_range.Font
    return {name: getattr(font, name) for name in FONT_PROPERTIES}


def _find_in_range(word_range, witness):
    text = str(word_range.Text)
    offset = text.index(witness)
    start = int(word_range.Start) + len(text[:offset].encode("utf-16-le")) // 2
    target = word_range.Duplicate
    target.SetRange(start, start + len(witness.encode("utf-16-le")) // 2)
    if str(target.Text) != witness:
        raise RuntimeError(f"Unexpected Word range for {witness!r}: {target.Text!r}")
    return target


def _word_state(app, path, *, template):
    with _opened(app, path) as document:
        report = {"notes": {}}
        for kind, (collection, _, _, _, _, _, _, _, witness, _) in NOTE_KINDS.items():
            notes = getattr(document, collection)
            rows = []
            for index in range(1, notes.Count + 1):
                note = notes.Item(index)
                note_range = note.Range
                content = without_bidi_controls(str(note_range.Text))
                token = witness if witness in content else f"Template {kind} exemplar"
                reference = note.Reference
                rows.append({"text": content, "font": _font(_find_in_range(note_range, token)),
                             "reference_font": _font(reference), "body_anchor_text": str(reference.Paragraphs.Item(1).Range.Text)})
            options = getattr(document.Sections.Item(1).Range, kind.capitalize() + "Options")
            report["notes"][kind] = {"starting_number": int(options.StartingNumber), "count": int(notes.Count), "rows": rows}
        if not template:
            report["body_font"] = _font(_find_in_range(document.Content, BODY))
    return report


def _package_notes(path):
    parts = read_package(path)
    document = xml(parts["word/document.xml"])
    result = {}
    for kind in NOTE_KINDS:
        root = xml(parts[f"word/{kind}s.xml"])
        relationship_path = f"word/_rels/{kind}s.xml.rels"
        relationships = xml(parts[relationship_path]) if relationship_path in parts else ()
        targets = {rel.get("Id"): rel.get("Target") for rel in relationships if rel.get("TargetMode") == "External"}
        result[kind] = {
            "body_references": [node.get(q("w:id")) for node in document.iter(q("w:" + kind + "Reference"))],
            "native_notes": [node.get(q("w:id")) for node in root if node.get(q("w:type"), "normal") == "normal"],
            "separators": [node.get(q("w:type")) for node in root if node.get(q("w:type")) in {"separator", "continuationSeparator"}],
            "links": [targets.get(node.get(q("r:id"))) for node in root.iter(q("w:hyperlink"))],
        }
    fields = _package_state(parts)["fields"]
    citations = []
    for field in fields:
        if field["owner"] == "ITEM":
            data = _payload(field["code"], "ITEM")
            citations.append({"part": field["part"], "result": without_bidi_controls(field["result"]).strip(),
                              "items": [entry["uris"][0].rsplit("/", 1)[-1] for entry in data["citationItems"]],
                              "note_index": data.get("properties", {}).get("noteIndex")})
    return result, citations, sum(field["owner"] == "BIBL" for field in fields)


def _render(pdf_path, label="review-page"):
    result = []
    with pymupdf.open(pdf_path) as pdf:
        for page in pdf:
            image = pdf_path.parent / f"{label}-{page.number + 1}.png"
            page.get_pixmap(dpi=130, alpha=False).save(image)
            result.append({"page": page.number + 1, "text": " ".join(page.get_text().split()), "image": str(image)})
    return result


def _edit_note_equations(master, output):
    copy = output / "edited-note-equations.docx"
    pdf = output / "edited-note-equations.pdf"
    edited = []
    with private_word() as app:
        with _opened(app, master) as document, MathTypeSDK() as sdk:
            for kind, (collection, _, _, _, _, _, _, _, witness, _) in NOTE_KINDS.items():
                notes = getattr(document, collection)
                matches = []
                for index in range(1, notes.Count + 1):
                    note_range = notes.Item(index).Range
                    if witness not in str(note_range.Text):
                        continue
                    shapes = note_range.InlineShapes
                    for shape_index in range(1, shapes.Count + 1):
                        shape = shapes.Item(shape_index)
                        if shape.Type == 1 and str(shape.OLEFormat.ProgID) == MATHTYPE_PROGID:
                            matches.append(shape)
                if len(matches) != 1:
                    raise RuntimeError(f"Expected one native MathType object in the imported {kind}")
                shape = matches[0]
                shape.OLEFormat.DoVerb(2)
                ole = shape.OLEFormat.Object
                sdk.set_ole_equation(ole, r"+\eta", "latex")
                ole = shape = matches = shapes = note_range = notes = None
                edited.append(kind)
            document.SaveAs2(FileName=str(copy), FileFormat=DOCX_FORMAT, AddToRecentFiles=False)
        export_pdf(copy, pdf, app=app)
    return {"path": str(copy), "edited_notes": edited, "math": audit_docx_math(copy),
            "pages": _render(pdf, "edited-note-equations-page")}


def _math_by_part(report):
    return {item["part"]: item["mtef_sha256"] for item in report["objects"] if item["valid_native_storage"]}


def _matches(state, **expected):
    return all(state.get(name) == value for name, value in expected.items())


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument("--note-equation", action="store_true", help="Convert and re-edit native MathType in both note stories")
    parser.add_argument("--markers", action="store_true", help="Expand note citation and cross-reference markers")
    args = parser.parse_args()
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    source, template = output / "source.docx", output / "template.docx"
    _create(source, template=False, item=ZoteroClient().get_item(fixture_key(SECOND_DOI)), note_equation=args.note_equation,
            markers=args.markers)
    _create(template, template=True)
    hashes = {path: digest(path.read_bytes()) for path in (source, template)}
    with private_word() as app:
        original = _word_state(app, source, template=False)
        target = _word_state(app, template, template=True)
    built = output / "built"
    workflow = run_pipeline(source, template, built, csl_style=CNS_STYLE if args.markers else None)
    master = built / "editable-master.docx"
    with private_word() as app:
        saved = _word_state(app, master, template=False)
    package_notes, citations, bibliography_count = _package_notes(master)
    math = audit_docx_math(master)
    pages = _render(built / "editable-master.pdf")
    visible = " ".join(page["text"] for page in pages)
    expected_math = {"word/document.xml": 1}
    if args.note_equation:
        expected_math.update({f"word/{kind}s.xml": 1 for kind in NOTE_EQUATIONS})
    checks = {
        "body_uses_template_font": _matches(saved["body_font"], NameAscii="Arial", Size=13),
        "both_note_citations_refreshed_in_original_stories": (
            workflow["citation_style"]["native_refresh_verified"] and bibliography_count == 1
            and Counter(citation["part"] for citation in citations) == {"word/footnotes.xml": 1, "word/endnotes.xml": 1}
            and all(citation["result"] == "[1]" and citation["items"] == [fixture_key(SECOND_DOI)] for citation in citations)
        ),
        "native_math_type": math["all_mathtype_objects_valid"] and not math["omml_equations"]
        and Counter(item["part"] for item in math["objects"]) == expected_math,
        "source_and_template_unchanged": all(digest(path.read_bytes()) == value for path, value in hashes.items()),
    }
    if args.markers:
        from paperflow.citations import _scan_fields
        from paperflow.package import audit_docx

        checks["all_note_markers_expanded_and_audited"] = audit_docx(master)["structural_pass"]
        checks["note_marker_indices_recomputed_by_zotero"] = all(
            citation["note_index"] == saved["notes"][kind]["count"]
            for kind in NOTE_KINDS for citation in citations if citation["part"] == f"word/{kind}s.xml"
        )
        for kind in NOTE_KINDS:
            root = xml(read_package(master)[f"word/{kind}s.xml"])
            fields = _scan_fields(root)
            numbers = [without_bidi_controls(field["result"]).strip() for field in fields
                       if re.match(r"^\s*REF\s+NoteTable\b", field["code"], re.I)]
            page_refs = [without_bidi_controls(field["result"]).strip() for field in fields
                         if re.match(r"^\s*PAGEREF\s+NoteTable\b", field["code"], re.I)]
            caption_pages = [str(page["page"]) for page in pages if "Table 1 Note reference witness" in page["text"]]
            checks[f"{kind}_live_table_number_and_page_refs"] = numbers == ["1"] and page_refs == caption_pages
    marker_rgb = int(MARKER_COLOR[:2], 16) + (int(MARKER_COLOR[2:4], 16) << 8) + (int(MARKER_COLOR[4:], 16) << 16)
    for kind, (_, _, _, _, _, font, size, marker_size, witness, start) in NOTE_KINDS.items():
        old, before, after = (record["notes"][kind] for record in (original, target, saved))
        state = package_notes[kind]
        checks[f"{kind}_source_format_distinct"] = old["count"] == 1 and _matches(old["rows"][0]["font"], NameAscii="Georgia", Size=20, Italic=-1)
        checks[f"{kind}_native_identity_order_and_numbering"] = (
            before["count"] == 1 and after["count"] == 2 and after["starting_number"] == before["starting_number"] == start
            and [row["body_anchor_text"].startswith(prefix) for row, prefix in zip(after["rows"], ("Template", "Imported"))] == [True, True]
            and state["body_references"] == state["native_notes"] and len(set(state["native_notes"])) == 2
            and state["separators"] == ["separator", "continuationSeparator"]
        )
        checks[f"{kind}_template_text_and_anchor_styles"] = (
            after["count"] == 2 and after["rows"][0]["font"] == before["rows"][0]["font"]
            and _matches(after["rows"][1]["font"], NameAscii=font, Size=size, Italic=-1)
            and _matches(after["rows"][1]["reference_font"], NameAscii=font, Size=marker_size,
                         Bold=-1, Superscript=-1, Color=marker_rgb)
        )
        checks[f"{kind}_external_link_and_rendered_text"] = (
            state["links"] == [f"https://example.org/paperflow/{kind}"]
            and witness in visible and f"{kind.capitalize()} external link" in visible
            and f"Template {kind} exemplar" in visible
        )
        checks[f"{kind}_rendered_note_numbers_follow_template"] = all(
            re.search(rf"(?<!\d){number}\s*{re.escape(token)}", visible)
            for number, token in ((start, f"Template {kind} exemplar"), (start + 1, witness))
        )
    edit = None
    if args.note_equation:
        checks["native_note_math_can_be_edited_independently"] = False
        if checks["native_math_type"]:
            master_hash = digest(master.read_bytes())
            edit = _edit_note_equations(master, output)
            before, after = _math_by_part(math), _math_by_part(edit["math"])
            checks["native_note_math_can_be_edited_independently"] = (
                edit["math"]["all_mathtype_objects_valid"] and not edit["math"]["omml_equations"]
                and Counter(item["part"] for item in edit["math"]["objects"]) == expected_math
                and before["word/document.xml"] == after["word/document.xml"]
                and all(before[f"word/{kind}s.xml"] != after[f"word/{kind}s.xml"] for kind in NOTE_EQUATIONS)
                and digest(master.read_bytes()) == master_hash
            )
    result = {"pass": all(checks.values()), "checks": checks, "source_word": original, "template_word": target,
              "saved_word": saved, "native_notes": package_notes, "zotero_citations": citations, "math": math, "pages": pages}
    if edit is not None:
        result["note_math_edit_witness"] = edit
    evidence = output / "note-format-evidence.json"
    evidence.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"pass": result["pass"], "checks": checks, "evidence": str(evidence)}, ensure_ascii=False, indent=2))
    return 0 if result["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
