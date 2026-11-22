"""Verify Word template fonts without losing imported character semantics."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
import pymupdf

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from paperflow.mathtype import audit_docx_math
from paperflow.package import digest, without_bidi_controls
from paperflow.pipeline import run_pipeline
from paperflow.word import _opened, private_word


WESTERN, EAST_ASIAN = "Arial", "SimSun"
TEMPLATE_COLOR, SOURCE_COLOR, EXEMPLAR_COLOR, LINK_COLOR = "253C5A", "A82731", "287346", "11639B"
SHARED_STYLE = "Fixture Shared Annotation"
EXEMPLAR = "Template exemplar"
WESTERN_WITNESS = "inheritedVariable"
EAST_ASIAN_WITNESS = "汉字校验"
OVERRIDE_WITNESS = "overriddenVariable"
SHARED_WITNESS = "sharedAnnotation"
HYPERLINK_WITNESS = "Native source link"
URL = "https://example.org/paperflow-character-fixture"
WITNESSES = (WESTERN_WITNESS, EAST_ASIAN_WITNESS, OVERRIDE_WITNESS, SHARED_WITNESS, HYPERLINK_WITNESS)
FONT_PROPERTIES = ("NameAscii", "NameFarEast", "Size", "Color", "Bold", "Italic", "Subscript", "Superscript", "Underline")


def _font(style, western, east_asian, size, color):
    style.font.name = western
    style.font.size = Pt(size)
    style.font.color.rgb = RGBColor.from_string(color)
    fonts = style._element.get_or_add_rPr().get_or_add_rFonts()
    fonts.set(qn("w:eastAsia"), east_asian)
    for kind in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme"):
        fonts.attrib.pop(qn("w:" + kind), None)


def _add_style(document, name, base=None, *, builtin=False):
    style = document.styles.add_style(name, WD_STYLE_TYPE.CHARACTER, builtin=builtin)
    if base is not None:
        style.base_style = base
    return style


def _make_template(path):
    template = Document()
    _font(template.styles["Normal"], WESTERN, EAST_ASIAN, 12, TEMPLATE_COLOR)
    shared = _add_style(template, SHARED_STYLE)
    _font(shared, "Calibri", "Microsoft YaHei", 16, EXEMPLAR_COLOR)
    shared.font.italic = True
    link = _add_style(template, "Hyperlink", builtin=True)
    link.font.color.rgb = RGBColor.from_string(LINK_COLOR)
    link.font.underline = True
    template.add_paragraph().add_run(EXEMPLAR, style=shared)
    template.add_paragraph("{{PAPERFLOW_BODY}}")
    template.save(path)


def _add_hyperlink(paragraph, style):
    relationship = paragraph.part.relate_to(URL, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship)
    run = paragraph.add_run(HYPERLINK_WITNESS, style=style)
    hyperlink.append(run._r)
    paragraph._p.append(hyperlink)


def _make_source(path):
    draft = Document()
    _font(draft.styles["Normal"], "Georgia", "Microsoft YaHei", 11, SOURCE_COLOR)
    parent = _add_style(draft, "Fixture Legacy Variables")
    _font(parent, "Georgia", "Microsoft YaHei", 22, SOURCE_COLOR)
    parent.font.italic = True
    parent.font.subscript = True
    child = _add_style(draft, "Fixture Inherited Variable", base=parent)
    shared = _add_style(draft, SHARED_STYLE)
    _font(shared, "Georgia", "Microsoft YaHei", 18, SOURCE_COLOR)
    shared.font.bold = True
    link = _add_style(draft, "Hyperlink", builtin=True)
    _font(link, "Georgia", "Microsoft YaHei", 20, SOURCE_COLOR)
    link.font.underline = True

    draft.add_paragraph("Inherited character formatting: ").add_run(WESTERN_WITNESS, style=child)
    draft.add_paragraph("East Asian font: ").add_run(EAST_ASIAN_WITNESS, style=child)
    direct = draft.add_paragraph("Explicit run overrides: ").add_run(OVERRIDE_WITNESS, style=child)
    direct.font.italic = False
    direct.font.superscript = True
    draft.add_paragraph("Source character style with shared name: ").add_run(SHARED_WITNESS, style=shared)
    _add_hyperlink(draft.add_paragraph("External hyperlink: "), link)

    equation = OxmlElement("m:oMath")
    run = OxmlElement("m:r")
    text = OxmlElement("m:t")
    text.text = "x+1=2"
    run.append(text)
    equation.append(run)
    draft.add_paragraph("Editable equation: ")._p.append(equation)
    draft.save(path)


def _inspect_word(app, path, witnesses):
    with _opened(app, path) as doc:
        results = {}
        for witness in witnesses:
            target = doc.Content.Duplicate
            target.Find.ClearFormatting()
            found = target.Find.Execute(FindText=witness, MatchCase=True, Forward=True, Wrap=0, Format=False)
            if not found or without_bidi_controls(str(target.Text)) != witness:
                raise RuntimeError(f"Could not locate the unique Word witness {witness!r} in {path}")
            font = target.Font
            results[witness] = {name: getattr(font, name) for name in FONT_PROPERTIES}
        links = [{"text": without_bidi_controls(str(link.Range.Text)), "address": str(link.Address)}
                 for link in (doc.Hyperlinks.Item(index) for index in range(1, doc.Hyperlinks.Count + 1))]
    return {"runs": results, "hyperlinks": links}


def _rgb(value):
    red, green, blue = bytes.fromhex(value)
    return red + green * 256 + blue * 65536


def _matches(run, **expected):
    return all(run[name] == value for name, value in expected.items())


def _font_aliases(run):
    east_asian = run["NameFarEast"]
    if east_asian.casefold() in {"微软雅黑", "microsoft yahei"}:
        east_asian = "Microsoft YaHei"
    return {**run, "NameFarEast": east_asian}


def _render(built):
    rendered = []
    with pymupdf.open(built / "editable-master.pdf") as pdf:
        for index, page in enumerate(pdf, 1):
            preview = built / f"review-page-{index}.png"
            page.get_pixmap(dpi=140, alpha=False).save(preview)
            rendered.append({"page": index, "text": " ".join(page.get_text().split()), "preview": str(preview)})
    return rendered


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    output = parser.parse_args().output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    source, template = output / "source.docx", output / "template.docx"
    _make_source(source)
    _make_template(template)
    originals = {path: digest(path.read_bytes()) for path in (source, template)}
    with private_word() as app:
        source_word = _inspect_word(app, source, WITNESSES)
        template_word = _inspect_word(app, template, (EXEMPLAR,))

    built = output / "built"
    workflow = run_pipeline(source, template, built)
    master = built / "editable-master.docx"
    with private_word() as app:
        saved_word = _inspect_word(app, master, (EXEMPLAR, *WITNESSES))
    old, saved = source_word["runs"], saved_word["runs"]
    rendered = _render(built)
    text = " ".join(page["text"] for page in rendered)
    math = audit_docx_math(master)
    normalized = next(step["evidence"] for step in workflow["steps"] if step["stage"] == "normalize")
    body = {"NameAscii": WESTERN, "Size": 12, "Color": _rgb(TEMPLATE_COLOR)}
    checks = {
        "native_build_finished": workflow["native_build_finished"] and workflow["structural_audit"]["structural_pass"],
        "source_style_chain_has_inherited_typography_and_semantics": _matches(
            old[WESTERN_WITNESS], NameAscii="Georgia", Size=22, Color=_rgb(SOURCE_COLOR), Italic=-1, Subscript=-1,
        ),
        "western_template_typography_and_inherited_emphasis": _matches(saved[WESTERN_WITNESS], **body, Italic=-1, Subscript=-1),
        "east_asian_template_font": saved[EAST_ASIAN_WITNESS]["NameFarEast"].casefold() in {"simsun", "宋体"}
                                  and saved[EAST_ASIAN_WITNESS]["Size"] == 12,
        "direct_italic_off_and_superscript_preserved": _matches(
            saved[OVERRIDE_WITNESS], **body, Italic=0, Subscript=0, Superscript=-1,
        ),
        "same_name_style_does_not_replace_source_emphasis": _matches(saved[SHARED_WITNESS], **body, Bold=-1, Italic=0),
        "template_exemplar_retains_its_character_style": _font_aliases(saved[EXEMPLAR]) == _font_aliases(template_word["runs"][EXEMPLAR]),
        "native_link_and_target_preserved": saved_word["hyperlinks"] == source_word["hyperlinks"] == [
            {"text": HYPERLINK_WITNESS, "address": URL},
        ],
        "native_link_uses_template_font_color_and_underline": _matches(
            saved[HYPERLINK_WITNESS], NameAscii=WESTERN, Size=12, Color=_rgb(LINK_COLOR), Underline=1,
        ),
        "editable_mathtype": math["all_mathtype_objects_valid"] and math["native_mathtype_objects"] == 1 and not math["omml_equations"],
        "word_render_contains_all_witnesses": all(token in text for token in (EXEMPLAR, *WITNESSES)),
        "source_and_template_unchanged": all(digest(path.read_bytes()) == value for path, value in originals.items()),
    }
    report = {"pass": all(checks.values()), "checks": checks, "source": source_word, "template": template_word,
              "saved": saved_word, "normalization": normalized, "rendered": rendered, "workflow": str(built / "workflow.json")}
    evidence = output / "character-style-evidence.json"
    evidence.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"pass": report["pass"], "checks": checks, "evidence": str(evidence)}, ensure_ascii=False, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
