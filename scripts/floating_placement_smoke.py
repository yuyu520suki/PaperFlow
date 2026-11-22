"""Verify C&S end pages with a cropped, text-wrapped picture made in Word."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import posixpath
import re
import shutil
import sys

import pymupdf
from docx import Document
from docx.shared import Cm

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from paperflow.package import digest, part_fields, q, read_package, xml
from paperflow.pipeline import CNS, run_pipeline
from paperflow.semantics import plan_draft
from paperflow.word import DOCX_FORMAT, private_word


PROJECT = Path(__file__).resolve().parents[1]
EXAMPLE = PROJECT / "examples" / CNS
TITLE = "Floating Figure Placement Fixture"
ABSTRACT = "This synthetic manuscript checks a cropped floating figure and an editable table."
KEYWORDS = ["Document workflow", "Floating figure"]
ANCHOR_TEXT = "The floating anchor belongs to this sentence."
FIGURE_CAPTION = "Synthetic wrapped response"
TABLE_CAPTION = "Synthetic table settings"
TABLE = [["Case", "Value"], ["A", "220"], ["B", "215"]]

_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}pic"


def _make_inline_seed(path: Path, *, static_captions=False) -> None:
    document = Document()
    document.add_paragraph(TITLE, style="Title")
    document.add_heading("Abstract", 1)
    document.add_paragraph(ABSTRACT)
    document.add_paragraph("Keywords: " + "; ".join(KEYWORDS))
    document.add_heading("1 Floating figure", 1)
    figure_ref = "1" if static_captions else "[[REF response]]"
    table_ref = "1" if static_captions else "[[REF settings]]"
    document.add_paragraph(f"The response is shown in Figure {figure_ref}.")
    paragraph = document.add_paragraph(ANCHOR_TEXT + " ")
    paragraph.add_run().add_picture(str(EXAMPLE / "response.png"), width=Cm(6.3))
    paragraph.add_run(
        "The adjoining lines provide an observable test of Word text wrapping. The curve "
        "starts near zero and changes sign later in the recorded interval. The time scale "
        "and the axis labels are part of the image. The illustrative values have no role "
        "in a research claim. This additional discussion keeps the source paragraph long "
        "enough to run alongside the entire picture in the main manuscript. The original "
        "paragraph remains in the main body when an independent copy of the figure is "
        "prepared for the end of the submission."
    )
    document.add_paragraph(f"Figure 1. {FIGURE_CAPTION}" if static_captions else f"[[CAPTION Figure response|{FIGURE_CAPTION}]]")
    document.add_heading("2 Table fixture", 1)
    document.add_paragraph(f"The fixture values are listed in Table {table_ref}.")
    document.add_paragraph(f"Table 1. {TABLE_CAPTION}" if static_captions else f"[[CAPTION Table settings|{TABLE_CAPTION}]]")
    table = document.add_table(rows=len(TABLE), cols=len(TABLE[0]))
    for row, values in zip(table.rows, TABLE):
        for cell, value in zip(row.cells, values):
            cell.text = value
    document.save(path)


def _float_in_word(inline: Path, output: Path) -> None:
    # Native source fixture, using Microsoft's InlineShape.ConvertToShape and
    # WdRelativeVerticalPosition/WdRelativeHorizontalPosition APIs.
    with private_word() as app:
        document = app.Documents.Open(
            FileName=str(inline), ReadOnly=True, AddToRecentFiles=False, Visible=False,
        )
        try:
            if int(document.InlineShapes.Count) != 1 or int(document.Shapes.Count):
                raise RuntimeError("Expected exactly one inline seed picture")
            picture = document.InlineShapes.Item(1).ConvertToShape()
            picture.WrapFormat.Type = 0  # wdWrapSquare
            picture.WrapFormat.Side = 1  # wdWrapLeft
            picture.WrapFormat.DistanceLeft = 9.0
            picture.WrapFormat.DistanceRight = 9.0
            picture.RelativeHorizontalPosition = 0  # wdRelativeHorizontalPositionMargin
            picture.Left = 238.0
            picture.RelativeVerticalPosition = 2  # wdRelativeVerticalPositionParagraph
            picture.Top = 0.0
            picture.PictureFormat.CropLeft = 2.0
            picture.PictureFormat.CropRight = 3.0
            if int(document.Shapes.Count) != 1 or int(document.InlineShapes.Count):
                raise RuntimeError("Word did not convert the inline seed into a floating picture")
            document.SaveAs2(FileName=str(output), FileFormat=DOCX_FORMAT, AddToRecentFiles=False)
        finally:
            document.Close(SaveChanges=0)


def _inspect(path: Path) -> dict:
    parts = read_package(path)
    body = xml(parts["word/document.xml"]).find(q("w:body"))
    relationships = {
        rel.get("Id"): rel for rel in xml(parts["word/_rels/document.xml.rels"]).iter(q("rel:Relationship"))
    }
    pictures = []
    for index, child in enumerate(body):
        for drawing in child.iter(q("w:drawing")):
            container = drawing.find(_WP + "anchor")
            if container is None:
                container = drawing.find(_WP + "inline")
            if container is None or container.find(".//" + _PIC) is None:
                continue
            embeds = []
            for node in container.iter():
                rid = node.get(q("r:embed"))
                if rid:
                    relation = relationships[rid]
                    target = relation.get("Target", "")
                    part = posixpath.normpath(target.lstrip("/") if target.startswith("/") else "word/" + target)
                    embeds.append(digest(parts[part]))
            crop = container.find(".//" + _A + "srcRect")
            extent = container.find(_WP + "extent")
            picture_transform = container.find(".//" + _A + "xfrm")
            pictures.append({
                "body_index": index,
                "mode": container.tag.rsplit("}", 1)[-1],
                "host_text": "".join(t.text or "" for t in child.iter(q("w:t"))),
                "image_sha256": sorted(embeds),
                "crop": dict(crop.attrib) if crop is not None else {},
                "extent": dict(extent.attrib) if extent is not None else {},
                "picture_transform": dict(picture_transform.attrib) if picture_transform is not None else {},
                "wrap_square": container.find(_WP + "wrapSquare") is not None,
            })
    tables = [
        [["".join(t.text or "" for t in cell.iter(q("w:t"))) for cell in row.findall(q("w:tc"))]
         for row in table.findall(q("w:tr"))]
        for table in body.findall(q("w:tbl"))
    ]
    fields, errors = part_fields(body)
    return {
        "pictures": pictures, "tables": tables, "fields": fields, "field_errors": errors,
        "anchor_text_count": "".join(t.text or "" for t in body.iter(q("w:t"))).count(ANCHOR_TEXT),
    }


def _render(path: Path, directory: Path, label: str) -> list[dict]:
    result = []
    with pymupdf.open(path) as document:
        for number, page in enumerate(document, 1):
            pixmap = page.get_pixmap(dpi=130, colorspace=pymupdf.csRGB, alpha=False)
            preview = directory / f"{label}-{number}.png"
            pixmap.save(preview)
            text = " ".join(page.get_text().split())
            result.append({
                "page": number, "preview": str(preview), "width": pixmap.width, "height": pixmap.height,
                "pixels_sha256": hashlib.sha256(pixmap.samples).hexdigest(),
                "figure_caption": re.search(r"Figure 1[.:]? " + re.escape(FIGURE_CAPTION), text) is not None,
                "table_caption": re.search(r"Table 1[.:]? " + re.escape(TABLE_CAPTION), text) is not None,
                "anchor_text": text.count(ANCHOR_TEXT),
                "images": len(page.get_image_info()),
            })
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path, help="New directory for the fixture and native build evidence")
    parser.add_argument("--static-captions", action="store_true", help="Convert literal legacy figure and table numbers first")
    args = parser.parse_args()
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    inline = output / "source-inline.docx"
    source = output / "source-floating.docx"
    _make_inline_seed(inline, static_captions=args.static_captions)
    inline_hash = digest(inline.read_bytes())
    print("Creating a native floating picture with Word", flush=True)
    _float_in_word(inline, source)
    source_hash = digest(source.read_bytes())
    source_package = _inspect(source)
    if [p["mode"] for p in source_package["pictures"]] != ["anchor"]:
        raise RuntimeError("The native fixture did not serialize a single DrawingML anchor")
    anchor = source_package["pictures"][0]
    if not (anchor["wrap_square"] and ANCHOR_TEXT in anchor["host_text"] and
            any(int(value) for value in anchor["crop"].values())):
        raise RuntimeError("The native fixture did not preserve the expected Square-wrapped cropped anchor in prose")

    figure_source = output / "response.pdf"
    shutil.copyfile(EXAMPLE / figure_source.name, figure_source)
    build = output / "built"
    draft_plan = plan_draft(source) if args.static_captions else None
    if draft_plan is not None:
        (output / "caption-conversion-plan.json").write_text(json.dumps(draft_plan, ensure_ascii=False, indent=2), encoding="utf-8")
        if len(draft_plan["edits"]) != 4:
            raise RuntimeError(f"Expected two caption and two cross-reference edits: {draft_plan['issues']}")
    print("Building the master and C&S submission copy", flush=True)
    manifest = run_pipeline(
        source, PROJECT / "templates" / f"{CNS}.docx", build, journal=CNS,
        table_policy="three-line", submission_copy=True, draft_plan=draft_plan,
        metadata={"abstract": ABSTRACT, "keywords": KEYWORDS, "figure_sources": {"1": figure_source.name}},
    )
    master = _inspect(build / "editable-master.docx")
    submission = _inspect(build / "submission-copy.docx")
    master_picture, = master["pictures"]
    body_picture, end_picture = submission["pictures"]
    stages = {step["stage"]: step["evidence"] for step in manifest["steps"]}
    pages = stages["journal_object_pages"]
    figure_page = next(item for item in pages["objects"] if item["kind"] == "figure")["submission_pages"][-1]
    table_page = next(item for item in pages["objects"] if item["kind"] == "table")["submission_pages"][-1]

    previews = output / "previews"
    previews.mkdir()
    master_render = _render(build / "editable-master.pdf", previews, "master")
    submission_render = _render(build / "submission-copy.pdf", previews, "submission")
    master_prefix = [{k: p[k] for k in ("width", "height", "pixels_sha256")} for p in master_render]
    submission_prefix = [{k: p[k] for k in ("width", "height", "pixels_sha256")}
                         for p in submission_render[:len(master_render)]]
    figure_preview = submission_render[figure_page - 1]
    table_preview = submission_render[table_page - 1]
    master_figure = [page for page in master_render if page["figure_caption"]]
    artwork = stages["journal_artwork"]
    staged_figure = Path(artwork["figures"][0]["output"])
    checks = {
        "seed_and_source_unchanged": inline_hash == digest(inline.read_bytes()) and source_hash == digest(source.read_bytes()),
        "master_anchor_preserves_square_wrap": master_picture["mode"] == "anchor" and master_picture["wrap_square"],
        "submission_body_anchor_unchanged": body_picture == master_picture,
        "end_copy_is_inline": end_picture["mode"] == "inline",
        "asset_crop_extent_transform_preserved": all(
            master_picture[key] == end_picture[key] for key in ("image_sha256", "crop", "extent", "picture_transform")
        ) and anchor["crop"] == master_picture["crop"] and anchor["image_sha256"] == master_picture["image_sha256"],
        "body_prose_not_duplicated": master["anchor_text_count"] == submission["anchor_text_count"] == 1
                                    and not end_picture["host_text"].strip()
                                    and sum(page["anchor_text"] for page in submission_render) == 1,
        "editable_table_cells_preserved": master["tables"] == [TABLE] and submission["tables"] == [TABLE, TABLE],
        "original_body_fields_preserved": not (master["field_errors"] or submission["field_errors"])
                                          and master["fields"] == submission["fields"]
                                          and sum(field.lstrip().startswith("SEQ ") for field in master["fields"]) == 2
                                          and sum(field.lstrip().startswith("REF ") for field in master["fields"]) == 2,
        "master_figure_rendered": len(master_figure) == 1 and master_figure[0]["images"] == 1,
        "original_pdf_pages_pixel_identical": master_prefix == submission_prefix,
        "figure_and_table_on_distinct_end_pages": pages["pagination_verified"]
                                                  and len(submission_render) == len(master_render) + 2
                                                  and figure_preview["figure_caption"] and figure_preview["images"] == 1
                                                  and table_preview["table_caption"] and table_preview["images"] == 0,
        "separate_vector_source_preserved": artwork["all_figures_ready"] and digest(staged_figure.read_bytes()) == digest(figure_source.read_bytes()),
        "native_build_completed": manifest["native_build_finished"] and manifest["source_unchanged"],
    }
    if draft_plan is not None:
        checks["literal_captions_and_refs_migrated"] = (stages["draft_plan"]["captions"] == 2
                                                        and stages["draft_plan"]["cross_references"] == 2)
    report = {
        "pass": all(checks.values()), "checks": checks, "source": source_package, "master": master,
        "submission": submission, "master_render": master_render, "submission_render": submission_render,
        "pagination": pages, "workflow": str(build / "workflow.json"),
    }
    (output / "floating-placement-evidence.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(checks, ensure_ascii=False, indent=2), flush=True)
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
