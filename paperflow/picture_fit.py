"""Fit ordinary inline pictures to the selected template's text width.

Only drawing dimensions change. Media, relationships, picture transforms and
equation objects remain intact. Word pagination is still authoritative for
vertical placement and table-cell layout.
"""
from __future__ import annotations

from fractions import Fraction

from .package import NS, q


WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PIC_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture"
PIC = "{" + PIC_URI + "}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
EMU_PER_TWIP = 635
_OFF = {"0", "false", "off"}


def _integer(node, name, *, default=None, positive=False):
    raw = node.get(name) if node is not None else None
    if raw is None and default is not None:
        return default
    try:
        value = int(raw)
    except (ValueError, TypeError):
        raise ValueError("missing or invalid drawing/layout dimension") from None
    if value < (1 if positive else 0):
        raise ValueError("drawing/layout dimensions must be positive or zero")
    return value


def _limits(section, settings):
    if section is None:
        raise ValueError("target section has no explicit page geometry")
    direction = section.find(q("w:textDirection"))
    if direction is not None and direction.get(q("w:val"), "lrTb") != "lrTb":
        raise ValueError("vertical text layout needs native picture placement review")
    size, margins = section.find(q("w:pgSz")), section.find(q("w:pgMar"))
    width = _integer(size, q("w:w"), positive=True)
    height = _integer(size, q("w:h"), positive=True)
    width -= sum(_integer(margins, q("w:" + side)) for side in ("left", "right"))
    height -= sum(_integer(margins, q("w:" + side)) for side in ("top", "bottom"))
    gutter = _integer(margins, q("w:gutter"), default=0)
    top_gutter = settings.find(q("w:gutterAtTop")) if settings is not None else None
    if top_gutter is not None and top_gutter.get(q("w:val"), "1").casefold() not in _OFF:
        height -= gutter
    else:
        width -= gutter
    if min(width, height) <= 0:
        raise ValueError("target margins leave no positive printable page area")
    columns = section.find(q("w:cols"))
    policy = "single_column"
    if columns is not None:
        unequal = columns.get(q("w:equalWidth"), "1").casefold() in _OFF
        explicit = columns.findall(q("w:col"))
        if unequal:
            count = _integer(columns, q("w:num"), default=len(explicit), positive=True)
            if not explicit or len(explicit) != count:
                raise ValueError("unequal target columns have incomplete widths")
            widths = [_integer(col, q("w:w"), positive=True) for col in explicit]
            gaps = [_integer(col, q("w:space"), default=0) for col in explicit[:-1]]
            if sum(widths) + sum(gaps) > width:
                raise ValueError("target column widths exceed printable page width")
            width = min(widths)
            policy = "smallest_explicit_column"
        else:
            count = _integer(columns, q("w:num"), default=1, positive=True)
            gap = _integer(columns, q("w:space"), default=720)
            width = (width - gap * (count - 1)) // count
            policy = "equal_columns" if count > 1 else policy
    if width <= 0:
        raise ValueError("target columns leave no positive picture width")
    return width * EMU_PER_TWIP, height * EMU_PER_TWIP, policy


def _paragraph_indent(paragraph, styles):
    properties = paragraph.find(q("w:pPr"))
    if properties is not None and properties.find(q("w:framePr")) is not None:
        raise ValueError("picture in a paragraph frame needs native layout review")
    definitions = {node.get(q("w:styleId")): node for node in styles.findall(q("w:style"))}
    reference = properties.find(q("w:pStyle")) if properties is not None else None
    style_id = reference.get(q("w:val")) if reference is not None else next((
        key for key, node in definitions.items() if node.get(q("w:type")) == "paragraph"
        and node.get(q("w:default"), "").casefold() in {"1", "true", "on"}
    ), None)
    chain, visited = [], set()
    while style_id:
        if style_id in visited or style_id not in definitions:
            raise ValueError("picture paragraph has an unresolved style inheritance chain")
        visited.add(style_id)
        definition = definitions[style_id]
        chain.append(definition.find(q("w:pPr")))
        based = definition.find(q("w:basedOn"))
        style_id = based.get(q("w:val")) if based is not None else None
    layers = [styles.find("w:docDefaults/w:pPrDefault/w:pPr", NS), *reversed(chain), properties]
    values = {"left": 0, "right": 0, "firstLine": 0, "hanging": 0}
    numbered = False
    for layer in layers:
        if layer is not None and layer.find(q("w:framePr")) is not None:
            raise ValueError("picture in a paragraph frame needs native layout review")
        numbering = layer.find(q("w:numPr")) if layer is not None else None
        if numbering is not None:
            number = numbering.find(q("w:numId"))
            numbered = number is None or number.get(q("w:val")) != "0"
        indent = layer.find(q("w:ind")) if layer is not None else None
        if indent is None:
            continue
        for key in ("leftChars", "rightChars", "startChars", "endChars", "firstLineChars", "hangingChars"):
            if indent.get(q("w:" + key)) not in {None, "0"}:
                raise ValueError("character-based paragraph indents need native layout review")
        for key, logical in (("left", "start"), ("right", "end"), ("firstLine", "firstLine"), ("hanging", "hanging")):
            raw = indent.get(q("w:" + logical), indent.get(q("w:" + key)))
            if raw is not None:
                try:
                    values[key] = int(raw)
                except ValueError:
                    raise ValueError("picture paragraph has an invalid indent") from None
    if numbered:
        raise ValueError("picture in a numbered paragraph needs native list-indent review")
    # Negative indentation never grants permission to draw outside page margins.
    left = max(0, values["left"], values["left"] + values["firstLine"] - values["hanging"])
    return (left + max(0, values["right"])) * EMU_PER_TWIP


def fit_inline_pictures(document, target_section, styles_root, *, settings_root=None):
    """Shrink embedded inline pictures uniformly to the template column width.

    Pictures in tables, fields, compatibility choices and unsupported layouts
    are retained and reported. Tall pictures receive a review message instead
    of a guessed page-height reduction. Existing small pictures are not enlarged.
    """
    report = {"policy": "shrink_to_template_column_width", "pictures": [], "resized": 0, "unresolved": []}
    body = document.find(q("w:body"))
    if body is None:
        return report
    # A field result can be regenerated by its field producer, so do not resize
    # INCLUDEPICTURE or another generated picture, including multi-paragraph fields.
    depth, field_drawings = 0, set()
    for node in body.iter():
        if node.tag == q("w:fldChar"):
            kind = node.get(q("w:fldCharType"))
            if kind == "begin":
                depth += 1
            elif kind == "end":
                depth = max(0, depth - 1)
        elif node.tag == q("w:drawing") and depth:
            field_drawings.add(node)
    for drawing in body.iter(q("w:drawing")):
        inline = drawing.find(WP + "inline")
        record = {"index": len(report["pictures"]) + 1, "status": "skipped",
                  "before_emu": None, "after_emu": None}
        report["pictures"].append(record)
        ancestors = {parent.tag for parent in drawing.iterancestors()}
        if ancestors & {q("w:del"), q("w:moveFrom"), q("w:txbxContent"), q("w:object")}:
            record["reason"] = "picture belongs to a deleted, text-box or OLE story"
            continue
        if inline is None:
            record["reason"] = "floating or non-inline drawing retains its own placement"
            continue
        picture = inline.find(A + "graphic/" + A + "graphicData/" + PIC + "pic")
        if picture is None:
            record["reason"] = "drawing is not an ordinary picture"
            continue
        extent = inline.find(WP + "extent")
        properties = inline.find(WP + "docPr")
        record["alt"] = properties.get("descr", "") if properties is not None else ""
        try:
            before = [_integer(extent, axis, positive=True) for axis in ("cx", "cy")]
            record["before_emu"] = record["after_emu"] = before
            if len(drawing) != 1 or len(inline.findall(A + "graphic/" + A + "graphicData")) != 1:
                raise ValueError("picture drawing has unsupported extra payloads")
            data = picture.getparent()
            if len(data) != 1 or data.get("uri") != PIC_URI:
                raise ValueError("picture graphic data is not a standalone picture")
            blips = list(picture.iter(A + "blip"))
            if len(blips) != 1 or not blips[0].get(q("r:embed")) or any(q("r:link") in node.attrib for node in inline.iter()):
                record["reason"] = "linked or nonembedded picture retains its original dimensions"
                report["unresolved"].append(f"Inline picture {record['index']}: {record['reason']}")
                continue
            if drawing in field_drawings or ancestors & {q("w:fldSimple"), MC + "AlternateContent"}:
                record["reason"] = "generated or compatibility picture needs native layout review"
                report["unresolved"].append(f"Inline picture {record['index']}: {record['reason']}")
                continue
            if q("w:tc") in ancestors:
                raise ValueError("table-cell picture width requires native cell layout review")
            paragraph = next((parent for parent in drawing.iterancestors() if parent.tag == q("w:p")), None)
            if paragraph is None:
                raise ValueError("picture has no paragraph layout context")
            inner = picture.find(PIC + "spPr/" + A + "xfrm/" + A + "ext")
            inner_before = [_integer(inner, axis, positive=True) for axis in ("cx", "cy")]
            width, height, policy = _limits(target_section, settings_root)
            width -= _paragraph_indent(paragraph, styles_root)
            effects = inline.find(WP + "effectExtent")
            width -= sum(_integer(effects, side, default=0) for side in ("l", "r"))
            height -= sum(_integer(effects, side, default=0) for side in ("t", "b"))
            if min(width, height) <= 0:
                raise ValueError("paragraph indents or picture effects leave no positive text area")
            record.update(width_limit_emu=width, height_limit_emu=height, column_policy=policy)
            factor = min(Fraction(1), Fraction(width, before[0]))
            after = [int(value * factor) for value in before]
            inner_after = [int(value * factor) for value in inner_before]
            if min(*after, *inner_after) <= 0:
                raise ValueError("picture cannot fit without losing its dimensions")
            record["status"] = "resized" if factor < 1 else "unchanged"
            record["scale"] = float(factor)
            record["after_emu"] = after
            if factor < 1:
                for node, values in ((extent, after), (inner, inner_after)):
                    for axis, value in zip(("cx", "cy"), values):
                        node.set(axis, str(value))
                report["resized"] += 1
            if after[1] > height:
                report["unresolved"].append(f"Inline picture {record['index']}: height exceeds the target printable page; review vertical placement")
        except ValueError as exc:
            record.update(status="unresolved", reason=str(exc))
            report["unresolved"].append(f"Inline picture {record['index']}: {exc}")
    return report
