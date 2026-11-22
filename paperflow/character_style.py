"""Keep inline semantics while letting the destination set typography."""

from collections import Counter
from copy import deepcopy
from hashlib import sha256

from lxml import etree as ET

from .package import q


_TYPOGRAPHY = {q("w:" + name) for name in (
    "rFonts", "sz", "szCs", "color", "spacing", "kern", "position", "w", "fitText",
)}
_LINK_ROLES = {"hyperlink": "link", "超链接": "link", "超連結": "link",
               "followedhyperlink": "followed", "已访问的超链接": "followed", "已瀏覽的超連結": "followed"}


def _value(style, property_name, fallback=""):
    element = style.find(q("w:" + property_name)) if style is not None else None
    return element.get(q("w:val"), fallback) if element is not None else fallback


def _link_role(style_id, style):
    label = _value(style, "name", style_id)
    return _LINK_ROLES.get("".join(label.casefold().split())) or _LINK_ROLES.get(style_id.casefold())


class CharacterStyles:
    def __init__(self, source, target):
        self.source = source
        self.links = {role: sid for sid, style in target.items() if style.get(q("w:type")) == "character"
                      if (role := _link_role(sid, style)) is not None}
        self.occupied_ids = {sid.casefold() for sid in source.keys() | target.keys() if sid}
        self.occupied_names = {name.strip().casefold() for style in (*source.values(), *target.values())
                               for name in (_value(style, "name"), *_value(style, "aliases").split(",")) if name.strip()}
        self.styles = {}
        self.resolved = {}
        self.unresolved = []
        self.counts = Counter()

    def _neutral_style(self, sid, ancestors=frozenset()):
        if sid in self.resolved:
            return self.resolved[sid]
        if sid in ancestors:
            raise ValueError(f"Character style inheritance is circular at {sid!r}")
        definition = self.source.get(sid)
        if definition is None or definition.get(q("w:type")) != "character":
            self.unresolved.append(f"No source character style definition for {sid!r}")
            self.resolved[sid] = None
            return None
        parent_id = _value(definition, "basedOn")
        parent = self._neutral_style(parent_id, ancestors | {sid}) if parent_id else None
        source_properties = definition.find(q("w:rPr"))
        kept = [] if source_properties is None else [
            child for child in source_properties if child.tag not in _TYPOGRAPHY and child.tag != q("w:rStyle")
        ]
        if not kept:
            self.resolved[sid] = parent
            return parent

        # Retain Word's toggle and inheritance layers. Direct formatting still wins.
        payload = "\0".join((sid, parent or "", *(ET.tostring(node, encoding="unicode") for node in kept)))
        label = _value(definition, "name", sid).replace(",", " ")[:44].strip()
        salt = 0
        while True:
            token = sha256(f"{salt}\0{payload}".encode("utf-8")).hexdigest()[:24]
            neutral_id = "PFInline" + token
            name = f"PaperFlow {label} {token}"
            if neutral_id.casefold() not in self.occupied_ids and name.casefold() not in self.occupied_names:
                break
            salt += 1
        self.occupied_ids.add(neutral_id.casefold())
        self.occupied_names.add(name.casefold())
        style = ET.Element(q("w:style"), {q("w:type"): "character", q("w:styleId"): neutral_id,
                                          q("w:customStyle"): "1"})
        ET.SubElement(style, q("w:name"), {q("w:val"): name})
        if parent:
            ET.SubElement(style, q("w:basedOn"), {q("w:val"): parent})
        properties = ET.SubElement(style, q("w:rPr"))
        properties.extend(deepcopy(node) for node in kept)
        self.styles[neutral_id] = style
        self.resolved[sid] = neutral_id
        return neutral_id

    def normalize(self, run):
        properties = run.find(q("w:rPr"))
        if properties is None:
            return
        for node in list(properties):
            if node.tag in _TYPOGRAPHY:
                properties.remove(node)
        reference = properties.find(q("w:rStyle"))
        if reference is None:
            return
        sid = reference.get(q("w:val"), "")
        link = self.links.get(_link_role(sid, self.source.get(sid)))
        if link and next(run.iterancestors(q("w:hyperlink")), None) is not None:
            reference.set(q("w:val"), link)
            self.counts["template_hyperlink_runs"] += 1
        elif neutral := self._neutral_style(sid):
            reference.set(q("w:val"), neutral)
            self.counts["remapped_runs"] += 1
        else:
            properties.remove(reference)
            self.counts["paragraph_typography_runs"] += 1

    def report(self):
        return {**self.counts, "source_style_map": self.resolved, "neutral_styles": list(self.styles)}
