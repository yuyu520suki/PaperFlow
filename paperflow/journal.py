"""Checks supported by explicit author-guide evidence, with unresolved gates."""
from .package import read_package, without_bidi_controls, xml, q, NS
from .journal_content import abstract_word_count, manuscript_frontmatter, split_keywords
from .table_style import table_look_mask

CNS_STYLE = "http://www.zotero.org/styles/computers-and-structures"
_VERTICAL_EDGES = {q("w:" + edge) for edge in ("left", "right", "insideV", "start", "end")}
_CONDITIONAL_LOOK = {
    "firstRow": 0x20, "lastRow": 0x40, "firstCol": 0x80, "lastCol": 0x100,
    "nwCell": 0xA0, "neCell": 0x120, "swCell": 0xC0, "seCell": 0x140,
}
_BAND_LOOK = {"band1Horz": 0x200, "band2Horz": 0x200, "band1Vert": 0x400, "band2Vert": 0x400}


def _vertical_edges(nodes):
    return {node.tag for node in nodes
            if node.tag in _VERTICAL_EDGES and node.get(q("w:val"), "single") not in {"nil", "none"}}


def _colored_fill(node):
    return node.get(q("w:fill"), "auto") not in {"auto", "FFFFFF", "ffffff"}


def _possible_conditional_format(kind, masks):
    if kind == "wholeTable":
        return False  # Word uses the style's ordinary properties for whole-table formatting.
    if masks is None:
        return True
    if kind in _BAND_LOOK:
        return any(not mask & _BAND_LOOK[kind] for mask in masks)
    required = _CONDITIONAL_LOOK.get(kind, 0)
    return any(mask & required == required for mask in masks)


def _table_style_risks(table, styles, default_style):
    reference = table.find("w:tblPr/w:tblStyle", NS)
    style_id = reference.get(q("w:val")) if reference is not None else None
    style_id = style_id if style_id in styles else default_style
    if style_id is None:
        return False, False

    looks = [table.find("w:tblPr/w:tblLook", NS)]
    conditions = table.xpath(".//w:tblPrEx/w:tblLook | .//w:cnfStyle", namespaces=NS)
    conditions = [node for node in conditions if next(node.iterancestors(q("w:tbl"))) is table]
    looks.extend(node for node in conditions if node.tag == q("w:tblLook"))
    try:
        masks = None if any(node.tag == q("w:cnfStyle") for node in conditions) else [
            table_look_mask(look) for look in looks
        ]
    except ValueError:
        masks = None

    # Three-line formatting directly suppresses the inherited table edges. This
    # says nothing about conditional styles or separate cell-level borders.
    suppressed = {node.tag for node in table.findall("w:tblPr/w:tblBorders/*", NS)
                  if node.get(q("w:val")) == "nil"}
    borders, shading, visited = False, False, set()
    while style_id in styles and style_id not in visited:
        visited.add(style_id)
        style = styles[style_id]
        blocks = [style, *(node for node in style.findall(q("w:tblStylePr"))
                           if _possible_conditional_format(node.get(q("w:type")), masks))]
        for block in blocks:
            edges = _vertical_edges(block.findall("w:tblPr/w:tblBorders/*", NS))
            if block is style:
                edges -= suppressed
            borders |= bool(edges or _vertical_edges(block.findall("w:tcPr/w:tcBorders/*", NS)))
            shading |= any(
                node.get(q("w:val")) != "nil" and (
                    _colored_fill(node) or node.get(q("w:themeFill")) is not None
                    or node.get(q("w:val"), "clear") not in {"clear", "pct0"}
                ) for node in block.findall("*/w:shd", NS)
            )
        parent = style.find(q("w:basedOn"))
        style_id = parent.get(q("w:val")) if parent is not None else None
    return borders, shading


def _table_format_findings(document, parts):
    style_root = xml(parts["word/styles.xml"]) if "word/styles.xml" in parts else None
    styles = {node.get(q("w:styleId")): node for node in style_root.findall('w:style[@w:type="table"]', NS)
              if node.get(q("w:styleId"))} if style_root is not None else {}
    default_style = next((key for key, node in styles.items()
                          if node.get(q("w:default"), "").lower() in {"1", "true", "on"}), None)
    vertical, shading, styled_vertical, styled_shading = [], [], [], []
    for number, table in enumerate(document.iter(q("w:tbl")), 1):
        direct = table.xpath(".//w:tblBorders/* | .//w:tcBorders/* | .//w:shd", namespaces=NS)
        direct = [node for node in direct if next(node.iterancestors(q("w:tbl"))) is table]
        if _vertical_edges(direct):
            vertical.append(number)
        if any(node.tag == q("w:shd") and _colored_fill(node) for node in direct):
            shading.append(number)
        style_borders, style_shading = _table_style_risks(table, styles, default_style)
        if style_borders:
            styled_vertical.append(number)
        if style_shading:
            styled_shading.append(number)
    return ((vertical, styled_vertical), (shading, styled_shading))


def audit_rendered_object_pages(master_pdf, submission_pdf, placement):
    """Check that each matched caption gains one occurrence on its own end page.

    Existing occurrences include both the body caption and any navigation entry.
    The comparison does not establish object appearance or visual placement.
    """
    import pymupdf

    def compact(text):
        return " ".join(without_bidi_controls(text).split())

    def read_pages(pdf):
        with pymupdf.open(pdf) as document:
            return [compact(page.get_text()) for page in document]

    master_pages, submission_pages = map(read_pages, (master_pdf, submission_pdf))
    checks = []
    for copy in placement["copies"]:
        caption = compact(copy["caption"])
        before = [i for i, page in enumerate(master_pages, 1) for _ in range(page.count(caption))] if caption else []
        after = [i for i, page in enumerate(submission_pages, 1) for _ in range(page.count(caption))] if caption else []
        observed = bool(before and after) and after[:-1] == before and after[-1] > len(master_pages)
        checks.append({"kind": copy["kind"], "caption": caption, "master_pages": before,
                       "submission_pages": after, "separate_end_page_observed": observed})
    end_pages = [item["submission_pages"][-1] for item in checks if item["separate_end_page_observed"]]
    return {"master_page_count": len(master_pages), "submission_page_count": len(submission_pages),
            "objects": checks, "pagination_verified": len(end_pages) == len(checks) == len(set(end_pages)),
            "visual_layout_reviewed": False, "relevant_text_proximity_verified": False}


def audit_journal(path, metadata=None, *, placement=None, artwork=None, citation_style=None):
    metadata = metadata or {}
    parts = read_package(path)
    document = xml(parts["word/document.xml"])
    checks = []
    def record(name, passed, evidence_page, detail):
        checks.append({"name": name, "pass": passed, "guide_pdf_page": evidence_page, "detail": detail})
    columns = []
    for section in document.iter(q("w:sectPr")):
        cols = section.find(q("w:cols"))
        columns.append(int(cols.get(q("w:num"), "1")) if cols is not None else 1)
    record("word_single_column", bool(columns) and all(n == 1 for n in columns), 8, columns)
    frontmatter = manuscript_frontmatter(document, parts)
    abstract, indexes, reason = frontmatter["abstract"]
    count = abstract_word_count(abstract) if abstract is not None else None
    record("abstract_words", None if count is None else 0 < count <= 250, 9,
           {"source": "word/document.xml", "body_paragraphs": indexes, "observed_count": count,
            "max_words": 250, "reason": reason})
    keyword_text, indexes, reason = frontmatter["keywords"]
    keywords = split_keywords(keyword_text) if keyword_text is not None else None
    count = len(keywords) if keywords is not None else None
    record("keyword_count", None if count is None else 1 <= count <= 7 and all(keywords), 9,
           {"source": "word/document.xml", "body_paragraphs": indexes, "observed_count": count,
            "keywords": keywords, "min_count": 1, "max_count": 7, "reason": reason})
    if "abstract" in metadata:
        declared = metadata["abstract"]
        passed = _normalized_text(abstract) == _normalized_text(declared) if abstract is not None else None
        record("abstract_metadata_consistency", passed, None,
               "Compare the supplied metadata with the abstract in the final Word body")
    if "keywords" in metadata:
        declared = metadata["keywords"]
        passed = None
        if keywords is not None:
            passed = isinstance(declared, list) and [term.casefold() for term in keywords] == [
                " ".join(term.casefold().split()) if isinstance(term, str) else None for term in declared
            ]
        record("keywords_metadata_consistency", passed, None,
               "Compare the supplied metadata with the keywords in the final Word body")
    if metadata.get("highlights") is not None:
        highlights = metadata["highlights"]
        lengths = [len(text) for text in highlights] if isinstance(highlights, list) and all(
            isinstance(text, str) and text.strip() for text in highlights
        ) else None
        record("highlights_limits", lengths is not None and 3 <= len(lengths) <= 5 and max(lengths) <= 85, 10,
               {"source": "metadata", "count": len(lengths) if lengths is not None else None, "lengths": lengths})
        record("highlights_attachment", None, 9, "Check that the highlights are in a separate editable submission file")
    else:
        record("highlights_limits", True, 9, "Highlights are encouraged. None declared in metadata.")
    for name, (direct, inherited) in zip(
        ("tables_no_vertical_borders", "tables_no_shading"), _table_format_findings(document, parts)
    ):
        inherited = sorted(set(inherited) - set(direct))
        detail = {"direct_tables": direct, "style_review_tables": inherited} if inherited else direct
        record(name, False if direct else None if inherited else True, 10, detail)
    if placement and placement.get("pagination_verified"):
        detail = "Each matched caption gains one occurrence on a distinct end page; review appearance and proximity to the discussion."
        if not placement["objects"]:
            detail = "No figures or tables in the checked submission copy."
        record("figure_table_dual_placement", None if placement["objects"] else True, 14, detail)
    else:
        record("figure_table_dual_placement", None, 14, "Needs object inventory and separate end-page verification")
    record("figures_separate_source_files", artwork.get("all_figures_staged") if artwork else None, 10,
           "See staged artwork report" if artwork else "Needs figure delivery manifest")
    if artwork:
        figures = artwork["figures"]
        failed = any(figure["status"] == "failed" for figure in figures)
        record("figure_file_properties", False if failed else (True if artwork["all_figures_ready"] else None), 11,
               [{"figure": f["figure_number"], "status": f["status"], "issues": f["issues"]} for f in figures])
    record("title_page_details", None, 9, "Review title, authors, full affiliations, and corresponding author contact")
    record("abstract_and_keywords_language_review", None, 9,
           "Review the abstract as standalone text, any essential abstract citations, and English keywords")
    record("competing_interests_attachment", None, 5,
           "Complete the publisher declarations tool and check the separate Word attachment")
    record("funding_disclosure", None, 5,
           "Confirm funding sources and funder roles, or confirm the no-funding statement")
    record("generative_ai_disclosure", None, 6,
           "Confirm applicability and, when required, check the declaration before the references")
    record("research_data_option_c", None, 13,
           "Confirm the data deposit and article citation/link, or confirm the reason data cannot be shared")
    if citation_style and (citation_style["citation_fields"] or citation_style["bibliography_fields"]):
        saved = citation_style["saved_style_id"]
        if saved == CNS_STYLE and citation_style["native_refresh_verified"]:
            status = "matches"
            note = "The saved editable master used the journal CSL style for the native Zotero refresh."
        elif saved and saved != CNS_STYLE:
            status = "different_style"
            note = "The draft uses a different CSL style. The guide permits consistent reference formatting at initial submission."
        else:
            status = "unverified"
            note = "The saved style or its native refresh could not be verified."
        record("journal_reference_style_selection", True if status == "matches" else None, 17,
               {"status": status, "expected_style_id": CNS_STYLE, "saved_style_id": saved, "note": note})
    record("citations_support_claims", None, 15, "Requires review against cited sources")
    return {"journal": "Computers & Structures", "checks": checks,
            "failed": [c["name"] for c in checks if c["pass"] is False],
            "unresolved": [c["name"] for c in checks if c["pass"] is None],
            "full_journal_compliance_proven": False}


def _normalized_text(value):
    return " ".join(value.split()) if isinstance(value, str) else None
