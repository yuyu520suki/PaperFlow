"""Normalize a working source copy to explicit template styles before insertion.

Only typography/layout properties are reset. Emphasis, superscripts, dynamic
fields, equations, drawings and embedded objects retain their structure.
"""
from copy import deepcopy
from .character_style import CharacterStyles
from .package import read_package, write_package, xml, q, NS
from .template import select_template_section
from .numbering import HeadingNumbering, numbered_heading_styles
from .notes import KINDS, normal_notes, normalize_note_marker, note_ids, note_parts, note_style_ids
from .table_style import apply_template_table_style
from .picture_fit import fit_inline_pictures
from .table_fit import fit_tables
from lxml import etree as ET


_ON = {"1", "true", "on"}


def _styles(root):
    return {s.get(q("w:styleId")): s for s in root.findall(q("w:style"))}


def _default_paragraph_style(definitions):
    return next((style_id for style_id, style in definitions.items()
                 if style.get(q("w:type")) == "paragraph" and style.get(q("w:default"), "").casefold() in _ON), None)


def _style_name(definition):
    name = definition.find(q("w:name"))
    return "".join(name.get(q("w:val"), "").casefold().split()) if name is not None else ""


def _builtin_heading(definition):
    if definition is None or definition.get(q("w:customStyle"), "").casefold() in _ON:
        return None
    label = _style_name(definition)
    style_id = definition.get(q("w:styleId"), "").casefold()
    return next((level for level in range(1, 10) if style_id == f"heading{level}"
                 or label in {f"heading{level}", f"标题{level}", f"標題{level}"}), None)


def _outline_level(properties):
    outline = properties.find(q("w:outlineLvl")) if properties is not None else None
    if outline is None:
        return None
    try:
        level = int(outline.get(q("w:val"), ""))
    except ValueError:
        return None
    return level if 0 <= level <= 9 else None


def _role(style_id, definitions, visited=None):
    visited = set() if visited is None else visited
    if not style_id or style_id in visited: return "body"
    visited.add(style_id)
    definition = definitions.get(style_id)
    if definition is None: return "body"
    if (heading := _builtin_heading(definition)) is not None: return f"heading{heading}"
    label = _style_name(definition)
    level = _outline_level(definition.find(q("w:pPr")))
    if level is not None and level < 9: return f"heading{level + 1}"
    if label == "title" or style_id.lower() == "title": return "title"
    if label == "subtitle" or style_id.lower() == "subtitle": return "subtitle"
    if "caption" in label or "caption" in style_id.lower() or label == "题注": return "caption"
    if label in {"equation", "公式"}: return "equation"
    if "bibliograph" in label: return "bibliography"
    parent = definition.find(q("w:basedOn"))
    inherited = _role(parent.get(q("w:val")), definitions, visited) if parent is not None else "body"
    return "body" if level == 9 and inherited.startswith("heading") else inherited


def _paragraph_role(properties, style_id, definitions):
    role = _role(style_id, definitions)
    level = _outline_level(properties)
    # Word keeps the outline level of a native Heading 1 through Heading 9.
    if level is None or _builtin_heading(definitions.get(style_id)) is not None:
        return role
    if level < 9:
        return f"heading{level + 1}"
    return "body" if role.startswith("heading") else role


_PARAGRAPH_LAYOUT = {q("w:" + name) for name in ("spacing", "ind", "jc", "tabs", "keepNext", "keepLines", "outlineLvl")}
_BEFORE_NUM_PR = {q("w:" + name) for name in ("pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl")}
_SECTION_LAYOUT = {q("w:" + name) for name in ("pgSz", "pgMar", "cols", "docGrid", "textDirection", "pgBorders", "vAlign")}
_SECTION_ORDER = {q("w:" + name): position for position, name in enumerate((
    "headerReference", "footerReference", "footnotePr", "endnotePr", "type",
    "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols",
    "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi",
    "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
))}


def _insert_before(parent, node, successors):
    for index, child in enumerate(parent):
        if child.tag in successors:
            parent.insert(index, node)
            return
    parent.append(node)


def _three_line_tables(document):
    for table in document.findall(".//w:tbl", NS):
        properties = table.find(q("w:tblPr"))
        if properties is None:
            properties = ET.Element(q("w:tblPr"))
            table.insert(0, properties)
        borders = properties.find(q("w:tblBorders"))
        if borders is None:
            borders = ET.Element(q("w:tblBorders"))
            _insert_before(properties, borders, {q("w:" + name) for name in (
                "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange",
            )})
        else:
            borders[:] = []
        for name in ("top", "left", "bottom", "right", "insideH", "insideV"):
            border = ET.SubElement(borders, q("w:" + name))
            border.set(q("w:val"), "single" if name in {"top", "bottom"} else "nil")
            border.set(q("w:sz"), "4")
            border.set(q("w:color"), "000000")
        for override in table.xpath(".//w:shd | ./w:tr/w:tblPrEx/w:tblBorders | ./w:tr/w:tc/w:tcPr/w:tcBorders", namespaces=NS):
            override.getparent().remove(override)
        first_row = table.find(q("w:tr"))
        if first_row is not None:
            for cell in first_row.findall(q("w:tc")):
                cell_properties = cell.find(q("w:tcPr"))
                if cell_properties is None:
                    cell_properties = ET.Element(q("w:tcPr"))
                    cell.insert(0, cell_properties)
                cell_borders = ET.Element(q("w:tcBorders"))
                _insert_before(cell_properties, cell_borders, {q("w:" + name) for name in (
                    "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark",
                    "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange",
                )})
                bottom = ET.SubElement(cell_borders, q("w:bottom"))
                bottom.set(q("w:val"), "single")
                bottom.set(q("w:sz"), "4")
                bottom.set(q("w:color"), "000000")


def _apply_section_layout(section, target_section):
    references = {q("w:headerReference"), q("w:footerReference")}
    for node in list(section):
        if node.tag in _SECTION_LAYOUT or node.tag in references:
            section.remove(node)
    if target_section is None:
        return
    for node in target_section:
        if node.tag in _SECTION_LAYOUT:
            order = _SECTION_ORDER[node.tag]
            position = next((i for i, child in enumerate(section)
                             if _SECTION_ORDER.get(child.tag, -1) > order), len(section))
            section.insert(position, deepcopy(node))


def _keep_page_boundary(paragraph):
    following = paragraph.getnext()
    if following is not None and following.tag == q("w:p"):
        properties = following.find(q("w:pPr"))
        if properties is None:
            properties = ET.Element(q("w:pPr"))
            following.insert(0, properties)
        page_break = properties.find(q("w:pageBreakBefore"))
        if page_break is None:
            page_break = ET.Element(q("w:pageBreakBefore"))
            prefix = {q("w:pStyle"), q("w:keepNext"), q("w:keepLines")}
            index = next((i for i, node in enumerate(properties) if node.tag not in prefix), len(properties))
            properties.insert(index, page_break)
        page_break.set(q("w:val"), "1")
        return "page_break_before"

    # An ordinary paragraph break property cannot be attached to a following
    # table or content control. Keep the break in the existing boundary paragraph.
    run = ET.SubElement(paragraph, q("w:r"))
    ET.SubElement(run, q("w:br")).set(q("w:type"), "page")
    return "inline_page_break"


def _collapse_source_sections(document, source_sections, target_section):
    """Fit the source body into one template slot without importing its furniture."""
    body = document.find(q("w:body"))
    terminal_sections = body.findall(q("w:sectPr"))
    if len(terminal_sections) > 1 or (terminal_sections and terminal_sections[0] is not body[-1]):
        raise ValueError("Source has a misplaced or repeated final section definition")
    terminal = terminal_sections[0] if terminal_sections else None
    boundaries = [section for section in source_sections if section is not terminal]

    changes = []
    for number, (boundary, following) in enumerate(zip(boundaries, [*boundaries[1:], terminal]), start=1):
        # The following section's type determines how it starts relative to
        # this boundary. A missing type is the standard next-page break.
        start_type = following.find(q("w:type")) if following is not None else None
        mode = "nextPage" if start_type is None else start_type.get(q("w:val"))
        if mode not in {"nextPage", "continuous"}:
            raise ValueError(f"Source section {number + 1} uses {mode!r}, which cannot be preserved in one template slot")
        paragraph = boundary.getparent().getparent()
        if any(ancestor.tag == q("w:tbl") for ancestor in paragraph.iterancestors()):
            raise ValueError("Source has an internal section break inside a table")
        boundary.getparent().remove(boundary)
        conversion = _keep_page_boundary(paragraph) if mode == "nextPage" else "no_added_break"
        changes.append({"after_source_section": number, "source_start_type": mode, "conversion": conversion})

    if terminal is None:
        terminal = ET.SubElement(body, q("w:sectPr"))
    # Header/footer links, restarts, first-page flags and other section options
    # belong to the destination. Match only its layout in the staging copy.
    terminal[:] = []
    _apply_section_layout(terminal, target_section)
    return changes


def _template_roles(target_styles, role_style_map):
    roles = {"body": _default_paragraph_style(target_styles) or "Normal"}
    for sid, style in target_styles.items():
        if style.get(q("w:type")) != "paragraph":
            continue
        role = _role(sid, target_styles)
        if role != "body" and role not in roles: roles[role] = sid
    note_text_styles, _ = note_style_ids(target_styles)
    roles.update(note_text_styles)
    roles.update(role_style_map or {})
    if "TableText" in target_styles: roles.setdefault("table","TableText")
    for sid in roles.values():
        if sid not in target_styles or target_styles[sid].get(q("w:type")) != "paragraph":
            raise ValueError(f"Template paragraph style not found: {sid}")
    return roles


def template_numbered_heading_styles(template_parts, role_style_map=None):
    if "word/styles.xml" not in template_parts:
        raise ValueError("The template must contain explicit styles")
    target_styles = _styles(xml(template_parts["word/styles.xml"]))
    roles = _template_roles(target_styles, role_style_map)
    candidates = {level: roles[f"heading{level}"] for level in range(1, 10) if f"heading{level}" in roles}
    eligible = numbered_heading_styles(template_parts, target_styles, candidates.values())
    styles = {}
    for level, style_id in candidates.items():
        if style_id not in eligible:
            continue
        definition = target_styles[style_id]
        builtin = _builtin_heading(definition)
        name = definition.find(q("w:name"))
        if builtin is not None:
            styles[level] = f"Heading {builtin}"
        elif name is not None and name.get(q("w:val")):
            styles[level] = name.get(q("w:val"))
    return styles


def normalize_to_template(source, template, output, *, role_style_map=None, slot=None, table_policy="preserve",
                          table_style=None):
    if table_policy not in {"preserve", "three-line"}:
        raise ValueError("Table policy must be preserve or three-line")
    if table_style is not None and table_policy != "preserve":
        raise ValueError("Select either a template table style or the three-line policy")
    source_parts, template_parts = read_package(source), read_package(template)
    if "word/styles.xml" not in source_parts or "word/styles.xml" not in template_parts:
        raise ValueError("Both documents must contain explicit styles")
    document = xml(source_parts["word/document.xml"])
    source_styles_root, target_styles_root = xml(source_parts["word/styles.xml"]), xml(template_parts["word/styles.xml"])
    source_styles, target_styles = _styles(source_styles_root), _styles(target_styles_root)
    template_document = xml(template_parts["word/document.xml"])
    target_section, template_location = select_template_section(template_document, slot)
    roles = _template_roles(target_styles, role_style_map)
    _, note_reference_styles = note_style_ids(target_styles)
    numbering = HeadingNumbering(source_parts, template_parts, target_styles,
                                 [sid for role, sid in roles.items() if role.startswith("heading")])
    numbered_levels = [level for level in range(1, 10) if (sid := roles.get(f"heading{level}")) in numbering.adopted_styles
                       and _builtin_heading(target_styles[sid]) == level]
    counts, unresolved = {}, list(numbering.unresolved)
    source_default_style = _default_paragraph_style(source_styles)
    characters = CharacterStyles(source_styles, target_styles)
    paragraphs = [(None, paragraph) for paragraph in document.findall(".//w:body//w:p", NS)]
    note_documents = note_parts(source_parts)
    notes_report = {}
    for kind in KINDS:
        ids = set(note_ids(document, kind))
        if not ids:
            continue
        notes = []
        if kind in note_documents:
            _, root = note_documents[kind]
            notes = [node for node in normal_notes(root, kind) if node.get(q("w:id")) in ids]
        missing = ids - {node.get(q("w:id")) for node in notes}
        if missing:
            unresolved.append(f"Missing {kind} definitions for references: {sorted(missing, key=str)}")
        notes_report[kind] = {"referenced_notes": len(notes), "text_style": roles.get(kind),
                              "reference_style": note_reference_styles.get(kind), "styled_markers": 0}
        if kind in roles or kind in note_reference_styles:
            paragraphs.extend((kind, paragraph) for note in notes for paragraph in note.iter(q("w:p")))
    for story, paragraph in paragraphs:
        if paragraph.xpath("ancestor::w:txbxContent", namespaces=NS):
            unresolved.append("text box formatting requires a separate role mapping")
            continue
        in_table = bool(paragraph.xpath("ancestor::w:tc", namespaces=NS))
        preserve_table = in_table and table_policy == "preserve" and "table" not in roles
        pp = paragraph.find(q("w:pPr"))
        sid = pp.find(q("w:pStyle")) if pp is not None else None
        source_id = sid.get(q("w:val")) if sid is not None else source_default_style
        role = story or _paragraph_role(pp, source_id, source_styles)
        if in_table and "table" in roles: role="table"
        target_id = None if preserve_table else roles.get(role)
        if target_id is None and story is None and not preserve_table:
            unresolved.append(f"No target style for role {role}")
        if target_id is not None:
            if pp is None:
                pp = ET.Element(q("w:pPr"))
                paragraph.insert(0, pp)
            if sid is None:
                sid = ET.Element(q("w:pStyle"))
                pp.insert(0, sid)
            sid.set(q("w:val"), target_id)
            for node in list(pp):
                if node.tag in _PARAGRAPH_LAYOUT:
                    pp.remove(node)
            direct_number = pp.find(q("w:numPr"))
            if role.startswith("heading") and target_id in numbering.adopted_styles:
                if direct_number is not None:
                    pp.remove(direct_number)
            elif role in {"body", "table", *KINDS} and source_id != target_id and direct_number is None:
                list_properties = numbering.source_list_properties(source_id, source_styles)
                if list_properties is not None:
                    index = next((index for index, node in enumerate(pp) if node.tag not in _BEFORE_NUM_PR), len(pp))
                    pp.insert(index, list_properties)
            counts[role] = counts.get(role, 0) + 1
        for run in paragraph.xpath(".//w:r[not(ancestor::w:txbxContent)]", namespaces=NS):
            if any(run.find(q("w:" + kind)) is not None for kind in ("object", "drawing", "pict", "sym")):
                continue
            marker_kind = normalize_note_marker(run, note_reference_styles)
            if marker_kind:
                if marker_kind in notes_report and marker_kind in note_reference_styles:
                    notes_report[marker_kind]["styled_markers"] += 1
            elif target_id is not None:
                characters.normalize(run)
    unresolved.extend(characters.unresolved)
    if table_policy == "three-line":
        _three_line_tables(document)
    style_application = None
    if table_style is not None:
        style_application = apply_template_table_style(document, template_document, target_styles, table_style)
        unresolved.extend(style_application["unresolved"])
    # Modern Word may open the effects copy of the style table. Use the same
    # neutral definitions in both parts, and keep the template's own effects.
    styles_part = "word/styles.xml"
    effects_part = "word/stylesWithEffects.xml"
    for part_name in (styles_part, effects_part):
        if part_name not in source_parts:
            continue
        if part_name == styles_part:
            root, existing, definitions = source_styles_root, source_styles, target_styles
        else:
            root = xml(source_parts[part_name])
            existing = _styles(root)
            template_effects = _styles(xml(template_parts[part_name])) if part_name in template_parts else {}
            definitions = {**target_styles, **template_effects}
        for sid, definition in {**definitions, **characters.styles}.items():
            if sid in existing:
                root.remove(existing[sid])
            root.append(deepcopy(definition))
        source_parts[part_name] = ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    source_sections = document.xpath("w:body/w:sectPr | w:body//w:p/w:pPr/w:sectPr[not(ancestor::w:txbxContent or ancestor::w:del or ancestor::w:moveFrom)]", namespaces=NS)
    if template_location["slot"] is not None:
        section_changes = _collapse_source_sections(document, source_sections, target_section)
        section_policy = "single_template_slot"
    else:
        for section in source_sections:
            _apply_section_layout(section, target_section)
        section_changes = []
        section_policy = "preserve_source_sections"
    table_fit = fit_tables(
        document, target_section, source_styles_root,
        settings_root=xml(template_parts["word/settings.xml"]) if "word/settings.xml" in template_parts else None,
    )
    unresolved.extend(table_fit["unresolved"])
    picture_fit = fit_inline_pictures(
        document, target_section, target_styles_root,
        settings_root=xml(template_parts["word/settings.xml"]) if "word/settings.xml" in template_parts else None,
    )
    unresolved.extend(picture_fit["unresolved"])
    source_parts["word/document.xml"]=ET.tostring(document,encoding="UTF-8",xml_declaration=True,standalone=True)
    for kind, (part, root) in note_documents.items():
        if kind in notes_report and (kind in roles or kind in note_reference_styles):
            source_parts[part] = ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    write_package(source_parts,output)
    return {"roles":counts,"role_style_ids":roles,"numbered_heading_levels":numbered_levels,
            "character_styles":characters.report(), "notes":notes_report, "picture_fit":picture_fit, "table_fit":table_fit,
            "sections_normalized":len(source_sections),"unresolved":unresolved,
            "template_location":template_location,"table_policy":table_policy,"table_style":style_application,
            "section_policy":section_policy,"source_section_breaks":section_changes,
            "header_footer_policy":"use destination template furniture",
            "retained_semantics":["bold","italic","superscript","fields","OMML","OLE","drawings"],
            "source_original_modified":False}
