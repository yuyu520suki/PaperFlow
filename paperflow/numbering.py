"""Move the list definitions used by target heading styles into a staging DOCX."""
from copy import deepcopy

from lxml import etree as ET

from .package import NS, q, xml
from .template import _internal_target, _part_relationships

_DOCUMENT = "word/document.xml"
_RELATIONSHIPS = "word/_rels/document.xml.rels"
_NUMBERING = "word/numbering.xml"
_NUMBERING_REL = NS["r"] + "/numbering"
_TYPES = "[Content_Types].xml"
_OVERRIDE = "{http://schemas.openxmlformats.org/package/2006/content-types}Override"
_NUMBERING_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
_DEPENDENCIES = {q("w:" + tag) for tag in ("numStyleLink", "styleLink", "lvlPicBulletId")}


def _serialize(element):
    return ET.tostring(element, encoding="UTF-8", xml_declaration=True, standalone=True)


def _numbering(parts):
    matches = [rel for rel in _part_relationships(parts, _DOCUMENT).values()
               if rel.get("Type") == _NUMBERING_REL]
    if not matches:
        return None, None
    if len(matches) != 1 or matches[0].get("TargetMode") == "External":
        raise ValueError("Word numbering must have one internal document relationship")
    path = _internal_target(_DOCUMENT, matches[0].get("Target"))
    if path not in parts:
        raise ValueError("Word document relationship points to a missing numbering part")
    return path, xml(parts[path])


def _add_numbering(parts):
    if _NUMBERING in parts:
        raise ValueError("Word numbering part exists without a document relationship")
    relationships = (xml(parts[_RELATIONSHIPS]) if _RELATIONSHIPS in parts else
                     ET.Element(q("rel:Relationships"), nsmap={None: NS["rel"]}))
    used = {relationship.get("Id") for relationship in relationships}
    index = 1
    while f"rId{index}" in used:
        index += 1
    ET.SubElement(relationships, q("rel:Relationship"), Id=f"rId{index}", Type=_NUMBERING_REL,
                  Target="numbering.xml")
    parts[_RELATIONSHIPS] = _serialize(relationships)
    types = xml(parts[_TYPES])
    override = next((item for item in types.findall(_OVERRIDE)
                     if item.get("PartName") == "/" + _NUMBERING), None)
    if override is None:
        override = ET.SubElement(types, _OVERRIDE, PartName="/" + _NUMBERING)
    override.set("ContentType", _NUMBERING_TYPE)
    parts[_TYPES] = _serialize(types)
    return _NUMBERING, ET.Element(q("w:numbering"), nsmap={"w": NS["w"]})


def style_numbering(style_id, styles):
    """Find the nearest style-owned numbering reference, including an explicit zero."""
    seen = set()
    while style_id in styles and style_id not in seen:
        seen.add(style_id)
        style = styles[style_id]
        props = style.find("w:pPr/w:numPr", NS)
        if props is not None:
            num_id = props.find(q("w:numId"))
            if num_id is not None:
                return style_id, props, num_id
        parent = style.find(q("w:basedOn"))
        style_id = parent.get(q("w:val")) if parent is not None else None
    return None


def _definitions(root, element, attr):
    return {node.get(q(attr)): node for node in root.findall(q(element))} if root is not None else {}


def _levels(abstract, concrete):
    levels = _definitions(abstract, "w:lvl", "w:ilvl")
    for override in concrete.findall(q("w:lvlOverride")):
        level = override.find(q("w:lvl"))
        if level is not None:
            levels[override.get(q("w:ilvl"))] = level
    return levels


def _style_level(style_ids, abstract, concrete):
    matches = [(index, level) for index, level in _levels(abstract, concrete).items()
               if any(node.get(q("w:val")) in style_ids for node in level.findall(q("w:pStyle")))]
    return matches[0] if len(matches) == 1 else (None, None)


def _concrete_abstract(concrete, abstracts):
    link = concrete.find(q("w:abstractNumId"))
    return abstracts.get(link.get(q("w:val"))) if link is not None else None


def _same(first, second):
    return first is not None and ET.tostring(first, method="c14n", exclusive=True) == ET.tostring(
        second, method="c14n", exclusive=True)


def _allocate(used, preferred):
    return preferred if preferred not in used else str(max(map(int, used)) + 1)


def _heading_numbering_sources(template_parts, target_styles, heading_style_ids):
    refs = [(style_id, reference[0], reference[2].get(q("w:val"))) for style_id in dict.fromkeys(heading_style_ids)
            if (reference := style_numbering(style_id, target_styles)) is not None
            and reference[2].get(q("w:val")) != "0"]
    if not refs:
        return
    _, target_root = _numbering(template_parts)
    nums = _definitions(target_root, "w:num", "w:numId")
    abstracts = _definitions(target_root, "w:abstractNum", "w:abstractNumId")
    for style_id, owner, number_id in refs:
        concrete = nums.get(number_id)
        abstract = _concrete_abstract(concrete, abstracts) if concrete is not None else None
        reason = "missing numbering definition"
        if abstract is not None:
            _, level = _style_level({style_id}, abstract, concrete)
            fmt = level.find(q("w:numFmt")) if level is not None else None
            format_value = fmt.get(q("w:val")) if fmt is not None else None
            has_dependencies = any(node.tag in _DEPENDENCIES for root in (abstract, concrete) for node in root.iter())
            reason = ("linked or picture numbering dependency" if has_dependencies else
                      "no unique numbered level linked to the heading style"
                      if not format_value or format_value in {"none", "bullet"} else None)
        yield style_id, owner, number_id, concrete, abstract, reason


def numbered_heading_styles(template_parts, target_styles, heading_style_ids):
    return {style_id for style_id, _, _, _, _, reason in
            _heading_numbering_sources(template_parts, target_styles, heading_style_ids) if reason is None}


class HeadingNumbering:
    def __init__(self, source_parts, template_parts, target_styles, heading_style_ids):
        self.source_parts = source_parts
        self.source_path, self.source_root = _numbering(source_parts)
        self.source_nums = _definitions(self.source_root, "w:num", "w:numId")
        self.source_abstracts = _definitions(self.source_root, "w:abstractNum", "w:abstractNumId")
        self.original_nums = self.source_nums.copy()
        self.original_abstracts = self.source_abstracts.copy()
        self.adopted_styles = set()
        self.unresolved = []
        remapped_nums, remapped_abstracts, refs_to_replace = {}, {}, {}
        for style_id, owner, number_id, concrete, abstract, reason in _heading_numbering_sources(
            template_parts, target_styles, heading_style_ids,
        ):
            if reason is not None:
                self.unresolved.append(f"Template heading style {style_id} requires numbering review: {reason}")
                refs_to_replace.setdefault(owner, "0")
                continue
            self.adopted_styles.add(style_id)
            if number_id not in remapped_nums:
                remapped_nums[number_id] = self._clone(concrete, abstract, remapped_abstracts)
            refs_to_replace[owner] = remapped_nums[number_id]
        # The nodes belong to the parsed template copy, never the original.
        for style_id, style in target_styles.items():
            ref = style.find("w:pPr/w:numPr/w:numId", NS)
            if style_id in refs_to_replace:
                ref.set(q("w:val"), refs_to_replace[style_id])
            elif ref is not None and ref.get(q("w:val")) in remapped_nums:
                ref.set(q("w:val"), remapped_nums[ref.get(q("w:val"))])
        if self.source_path is not None and remapped_nums:
            self.source_parts[self.source_path] = _serialize(self.source_root)

    def _clone(self, concrete, abstract, abstract_ids):
        number_id = concrete.get(q("w:numId"))
        abstract_id = abstract.get(q("w:abstractNumId"))
        if _same(self.source_nums.get(number_id), concrete) and _same(self.source_abstracts.get(abstract_id), abstract):
            return number_id
        if self.source_root is None:
            self.source_path, self.source_root = _add_numbering(self.source_parts)
        if abstract_id not in abstract_ids:
            destination = abstract_id
            if not _same(self.source_abstracts.get(abstract_id), abstract):
                destination = _allocate(self.source_abstracts, abstract_id)
                imported = deepcopy(abstract)
                imported.set(q("w:abstractNumId"), destination)
                for index, child in enumerate(self.source_root):
                    if child.tag in {q("w:num"), q("w:numIdMacAtCleanup")}:
                        self.source_root.insert(index, imported)
                        break
                else:
                    self.source_root.append(imported)
                self.source_abstracts[destination] = imported
            abstract_ids[abstract_id] = destination
        destination = _allocate(self.source_nums, number_id)
        imported = deepcopy(concrete)
        imported.set(q("w:numId"), destination)
        imported.find(q("w:abstractNumId")).set(q("w:val"), abstract_ids[abstract_id])
        cleanup = self.source_root.find(q("w:numIdMacAtCleanup"))
        if cleanup is not None:
            cleanup.addprevious(imported)
        else:
            self.source_root.append(imported)
        self.source_nums[destination] = imported
        return destination

    def source_list_properties(self, style_id, styles):
        reference = style_numbering(style_id, styles)
        if reference is None:
            return None
        owner, props, num_id = reference
        concrete = self.original_nums.get(num_id.get(q("w:val")))
        abstract = _concrete_abstract(concrete, self.original_abstracts) if concrete is not None else None
        if abstract is None:
            return None
        direct = deepcopy(props)
        if direct.find(q("w:ilvl")) is None:
            level, _ = _style_level({owner, style_id}, abstract, concrete)
            if level is None and len(_levels(abstract, concrete)) != 1:
                return None
            level = level if level is not None else next(iter(_levels(abstract, concrete)))
            direct.insert(0, ET.Element(q("w:ilvl"), {q("w:val"): level}))
        return direct
