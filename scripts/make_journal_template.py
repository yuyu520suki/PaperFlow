"""Create a project-owned submission template; this is not a publisher DOCX."""
import hashlib
import json
from pathlib import Path
import sys

from docx import Document
from docx.shared import Mm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from paperflow.package import field_runs, inspect_docx


def main():
    root = Path(__file__).resolve().parents[1]
    destination = root / "templates/computers-and-structures.docx"
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        prior=destination.with_suffix(".provenance.json")
        evidence=json.loads(prior.read_text(encoding="utf-8")) if prior.exists() else {}
        if evidence.get("template_kind")!="project_authored_submission_template" or evidence.get("template_sha256")!=hashlib.sha256(destination.read_bytes()).hexdigest():
            raise FileExistsError("Existing template has no matching generator provenance; it will not be overwritten")
    profile = json.loads((root / "profiles/computers-and-structures.json").read_text(encoding="utf-8"))
    doc = Document()
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.text import WD_TAB_ALIGNMENT
    doc.styles.add_style("Equation",WD_STYLE_TYPE.PARAGRAPH)
    doc.styles["Equation"].base_style=doc.styles["Normal"]
    doc.styles["Equation"].paragraph_format.tab_stops.add_tab_stop(Mm(159.2),WD_TAB_ALIGNMENT.RIGHT)
    doc.styles.add_style("TableText",WD_STYLE_TYPE.PARAGRAPH)
    doc.styles["TableText"].base_style=doc.styles["Normal"]
    doc.styles["TableText"].font.size=Pt(10)
    doc.styles["TableText"].paragraph_format.line_spacing=1
    doc.styles["TableText"].paragraph_format.space_after=Pt(3)
    section = doc.sections[0]
    section.page_width, section.page_height = Mm(210), Mm(297)
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Mm(25.4)
    # Only single-column is a publisher requirement. The following typography
    # and page dimensions are explicit project defaults, never official rules.
    for name in ["Normal", "Body Text", "Title", "Subtitle", "Caption", "TOC Heading", *[f"Heading {i}" for i in range(1, 10)]]:
        if name not in doc.styles: continue
        style = doc.styles[name]
        style.font.name = "Times New Roman"
        style.font.size = Pt(12)
        style.font.color.rgb = RGBColor(0, 0, 0)
        rpr = style.element.get_or_add_rPr()
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is not None:
            for a in ["asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"]: fonts.attrib.pop(qn("w:"+a), None)
            fonts.set(qn("w:eastAsia"), "宋体")
        style.paragraph_format.line_spacing = 1.5
        style.paragraph_format.space_after = Pt(6)
        if name.startswith("Heading "):
            style.font.bold = True
            style.font.size = Pt(14 if name == "Heading 1" else 12)
            style.paragraph_format.keep_with_next = True
        if name in {"Title","Subtitle"}:
            for border in list(style.element.findall(".//"+qn("w:pBdr"))):
                border.getparent().remove(border)
    doc.styles["Title"].font.size = Pt(16)
    doc.styles["Title"].font.bold = True
    doc.styles["Caption"].font.bold = False
    doc.styles["Caption"].font.size = Pt(10)
    doc.styles["Caption"].paragraph_format.line_spacing = 1
    # Page footer is a declared working convention; the guide does not mandate it.
    footer = section.footer.paragraphs[0]
    for run in field_runs(" PAGE ", "1"): footer._p.append(run)
    doc.add_paragraph("{{PAPERFLOW_BODY}}")
    doc.save(destination)
    provenance = {
        "template_kind": "project_authored_submission_template",
        "not_an_official_publisher_docx": True,
        "guide_sha256": profile["source"]["pdf_sha256"],
        "official_rule_applied": {"word_columns": 1},
        "author_defaults": {"paper": "A4", "margins_mm": 25.4, "western_font": "Times New Roman", "east_asian_font": "宋体", "body_size_pt": 12, "line_spacing": 1.5, "heading_color": "000000", "footer": "PAGE field", "running_header": None},
        "slot": "{{PAPERFLOW_BODY}}", "template_sha256": hashlib.sha256(destination.read_bytes()).hexdigest(),
        "validation": {"visual_reviewed": False, "journal_rules_fully_validated": False},
    }
    destination.with_suffix(".provenance.json").write_text(json.dumps(provenance, ensure_ascii=False, indent=2), encoding="utf-8")
    destination.with_suffix(".inventory.json").write_text(json.dumps(inspect_docx(destination), ensure_ascii=False, indent=2), encoding="utf-8")
    print(destination)


if __name__ == "__main__": main()
