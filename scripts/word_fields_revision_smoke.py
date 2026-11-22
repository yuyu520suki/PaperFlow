"""Prove native TOC, figure references and page fields after a Word revision."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
import sys

import pymupdf

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from paperflow.package import audit_docx, digest
from paperflow.pipeline import revise_master
from paperflow.word import private_word


HEADING = "2 Field Refresh Checkpoint"
REFERENCES = "3 References"
BODY = "This new section starts on a separate page after the Word revision."
BOOKMARK = "field_refresh_checkpoint"
PAGE_SENTENCE = "The new section starts on page "
ADDED_CAPTION = "Additional fixture caption"


def _field(doc, start, stop, instruction):
    return doc.Fields.Add(Range=doc.Range(Start=start, End=stop), Type=-1,
                          Text=instruction, PreserveFormatting=True)


def _edit_in_word(seed, draft):
    with private_word() as app:
        doc = app.Documents.Open(FileName=str(seed), ReadOnly=True, AddToRecentFiles=False, Visible=False)
        try:
            doc.SaveAs2(FileName=str(draft), FileFormat=12, AddToRecentFiles=False)
            before_caption = int(doc.Bookmarks.Item("response").Range.Paragraphs.Item(1).Range.Start)
            page_line = PAGE_SENTENCE + "00.\r"
            caption_line = f"Figure 00 {ADDED_CAPTION}\r"
            doc.Range(Start=before_caption, End=before_caption).InsertBefore(page_line + caption_line)
            caption_number = before_caption + len(page_line) + len("Figure ")
            _field(doc, caption_number, caption_number + 2, r"SEQ Figure \* ARABIC")

            references = next(p.Range.Duplicate for p in doc.Paragraphs if str(p.Range.Text).strip() == "2 References")
            references.End -= 1
            start = int(references.Start)
            references.Text = REFERENCES
            doc.Range(Start=start, End=start).InsertBefore(HEADING + "\r" + BODY + "\r")
            for paragraph in doc.Paragraphs:
                line = str(paragraph.Range.Text).strip()
                if line == HEADING:
                    paragraph.Style = doc.Styles.Item(-2)  # wdStyleHeading1
                    paragraph.Format.PageBreakBefore = True
                    heading_range = paragraph.Range.Duplicate
                    heading_range.End -= 1
                    doc.Bookmarks.Add(Name=BOOKMARK, Range=heading_range)
                elif line == BODY:
                    paragraph.Style = doc.Styles.Item(-1)  # wdStyleNormal
            if not doc.Bookmarks.Exists(BOOKMARK):
                raise RuntimeError("New Word heading was not placed")
            page_number = before_caption + len(PAGE_SENTENCE)
            _field(doc, page_number, page_number + 2, f"PAGEREF {BOOKMARK}")
            doc.Save()
        finally:
            doc.Close(SaveChanges=0)


def _verify(report, output):
    refreshed = output / "editable-master.docx"
    fields = next(step["evidence"]["snapshot"]["fields"] for step in reversed(report["steps"])
                  if step["stage"] in {"word_fields", "word_fields_after_zotero"})
    toc = [field["result"] for field in fields if field["kind"] == "TOC"]
    figure_ref = [field["result"].strip() for field in fields if re.match(r"^\s*REF\s+response\b", field["code"], re.I)]
    figure_sequence = [field["result"].strip() for field in fields if re.match(r"^\s*SEQ\s+Figure\b", field["code"], re.I)]
    page_ref = [field["result"].strip() for field in fields if re.match(rf"^\s*PAGEREF\s+{BOOKMARK}\b", field["code"], re.I)]

    pages = []
    with pymupdf.open(output / "editable-master.pdf") as pdf:
        for index, page in enumerate(pdf, start=1):
            raster = output / f"field-roundtrip-page-{index}.png"
            page.get_pixmap(dpi=144, alpha=False).save(raster)
            pages.append({"number": index, "text": page.get_text(), "png": str(raster)})
    heading_pages = [page["number"] for page in pages if HEADING in page["text"] and BODY in page["text"]]
    heading_page = str(heading_pages[0]) if len(heading_pages) == 1 else ""
    toc_entry = [line.strip() for result in toc for line in result.splitlines() if HEADING in line]
    rendered = "\n".join(page["text"] for page in pages)
    audit = audit_docx(refreshed)
    checks = {
        "two_rendered_pages": len(pages) == 2,
        "section_appears_after_page_one": bool(heading_page) and int(heading_page) > 1,
        "toc_includes_new_heading_and_page": len(toc) == len(toc_entry) == 1 and toc_entry[0] == f"{HEADING}\t{heading_page}",
        "toc_replaces_renamed_reference_heading": any(f"{REFERENCES}\t{heading_page}" in result for result in toc)
            and all("2 References\t" not in result for result in toc),
        "rendered_toc_entries_match_heading_page": bool(pages) and all(
            re.search(rf"(?m)^{re.escape(title)}\s*\.+\s*{heading_page}\s*$", pages[0]["text"]) is not None
            for title in (HEADING, REFERENCES)
        ),
        "original_bookmark_ref_updates_to_two": figure_ref == ["2"],
        "figure_sequence_in_document_order": figure_sequence == ["1", "2"],
        "pageref_points_to_rendered_heading_page": bool(heading_page) and page_ref == [heading_page],
        "updated_fields_rendered": all(text in rendered for text in (
            "Figure 2 is referenced dynamically.", f"Figure 1 {ADDED_CAPTION}", "Figure 2 A fixture caption",
            f"{PAGE_SENTENCE}{heading_page}.",
        )),
        "page_and_numpages_rendered": all(f"{page['number']} / {len(pages)}" in page["text"] for page in pages),
        "no_unresolved_word_fields": not (audit["errors"] or audit["pending_markers"]),
    }
    return {"checks": checks, "toc": toc, "toc_entry": toc_entry, "figure_ref": figure_ref,
            "figure_sequence": figure_sequence, "pageref": page_ref, "heading_page": heading_page, "pages": pages}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("seed", type=Path)
    parser.add_argument("output_dir", type=Path)
    args = parser.parse_args()
    source = args.seed.resolve(strict=True)
    before = digest(source.read_bytes())
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    draft = output / "word-edited.docx"
    _edit_in_word(source, draft)
    draft_hash = digest(draft.read_bytes())
    refreshed = output / "refreshed"
    report = revise_master(draft, refreshed)
    evidence = _verify(report, refreshed)
    evidence["checks"].update({"seed_unchanged": digest(source.read_bytes()) == before,
                               "edited_source_unchanged": digest(draft.read_bytes()) == draft_hash == report["source_sha256"]})
    evidence["pass"] = all(evidence["checks"].values())
    (output / "field-revision-evidence.json").write_text(json.dumps(evidence, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"pass": evidence["pass"], "checks": evidence["checks"]}, indent=2))
    return 0 if evidence["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
