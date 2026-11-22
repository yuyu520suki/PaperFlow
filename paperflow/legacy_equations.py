"""Rewrap Equation Editor MTEF in trusted MathType storage without activating it."""

from collections import Counter
from pathlib import Path, PurePosixPath
import struct
from uuid import uuid4

from lxml import etree as ET

from .mathtype import (LEGACY_EQUATION_PROGID, MATHTYPE_CLSID, MATHTYPE_PROGID,
                       audit_docx_math, discover_mathtype, inspect_native_embedding, legacy_mtef)
from .package import digest, q, read_package, write_package, xml
from .template import _internal_target


def _vendor_blank():
    import pythoncom
    from win32com.storagecon import STGM_READ, STGM_SHARE_DENY_WRITE, STGM_SHARE_EXCLUSIVE

    from .equations import EquationConversionError

    installation = discover_mathtype()
    if not installation.directory or not installation.progid_registered:
        raise EquationConversionError("Installed MathType OLE registration is required for Equation 3 conversion")
    path = Path(installation.directory) / "Office Support" / "BlankEqn.doc"
    if not path.is_file():
        raise EquationConversionError("Installed vendor BlankEqn.doc is missing")
    mode = STGM_READ | STGM_SHARE_EXCLUSIVE
    document = pythoncom.StgOpenStorage(str(path), None, STGM_READ | STGM_SHARE_DENY_WRITE)
    pool = document.OpenStorage("ObjectPool", None, mode)
    matches = [item for item in pool.EnumElements()
               if item[1] == 1 and str(item[8]).strip("{}").casefold() == MATHTYPE_CLSID]
    if len(matches) != 1:
        raise EquationConversionError("Vendor blank does not contain exactly one MathType storage")
    entry = matches[0]
    return document, pool, pool.OpenStorage(entry[0], None, mode), entry[8]


def _rewrap(blank, clsid, mtef):
    import pythoncom
    from win32com.storagecon import STGM_CREATE, STGM_READ, STGM_READWRITE, STGM_SHARE_EXCLUSIVE

    from .equations import EquationConversionError

    header = bytearray(blank.OpenStream("Equation Native", None, STGM_READ | STGM_SHARE_EXCLUSIVE).Read(28))
    if len(header) != 28 or struct.unpack_from("<HI", header) != (28, 0x00020000):
        raise EquationConversionError("Vendor blank has an unsupported Equation Native header")
    struct.pack_into("<I", header, 8, len(mtef))
    native = bytes(header) + mtef
    memory = pythoncom.CreateILockBytesOnHGlobal()
    mode = STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE
    storage = pythoncom.StgCreateDocfileOnILockBytes(memory, mode)
    blank.CopyTo(None, None, storage)
    storage.SetClass(clsid)
    stream = storage.CreateStream("Equation Native", mode)
    if stream.Write(native) != len(native):
        raise EquationConversionError("Incomplete write of MathType equation storage")
    stream.Commit(0)
    stream = None
    storage.Commit(0)
    storage = None
    data = memory.ReadAt(0, memory.Stat(1)[2])
    audit = inspect_native_embedding(data)
    if not audit["valid_native_storage"] or audit.get("mtef_sha256") != digest(mtef):
        raise EquationConversionError("MathType storage did not preserve the legacy MTEF payload")
    return data


def _embedding_targets(parts, objects, progid):
    """Ensure rewriting an embedding cannot change another relationship's object."""
    from .equations import EquationConversionError

    targets = {item["embedding_part"] for item in objects}
    owners = {(item["part"], item["relationship_id"]) for item in objects}
    for part, data in parts.items():
        path = PurePosixPath(part)
        if path.suffix != ".rels" and path.name != ".rels":
            continue
        owner = str(path.parent.parent / path.name.removesuffix(".rels"))
        for relation in xml(data):
            target = relation.get("Target")
            if (target and relation.get("TargetMode", "Internal") == "Internal"
                    and _internal_target(owner, target) in targets and (owner, relation.get("Id")) not in owners):
                raise EquationConversionError(f"Equation 3 embedding has a shared relationship in {part}")
    roots = {part: xml(parts[part]) for part in {part for part, _ in owners}}
    seen = Counter()
    for part, root in roots.items():
        for node in root.iter():
            key = (part, node.get(q("r:id")))
            if key not in owners:
                if any((part, node.get(q("r:" + attribute))) in owners for attribute in ("embed", "link")):
                    raise EquationConversionError(f"Equation 3 relationship is shared with other content in {part}")
                continue
            if node.tag != q("o:OLEObject") or node.get("ProgID", "").casefold() != progid.casefold():
                raise EquationConversionError(f"Equation 3 relationship is shared with other content in {part}")
            seen[key] += 1
    if seen != Counter((item["part"], item["relationship_id"]) for item in objects):
        raise EquationConversionError("Equation 3 XML objects do not match their audited relationships")
    return targets, roots


def rewrap_legacy_equations(parts, legacy, output, *, anchor_for_insertion=False):
    """Prepare a DOCX copy before Word can change the original MTEF bytes."""
    from .equations import EquationConversionError

    invalid = [item for item in legacy if not item["valid_legacy_storage"]]
    if invalid:
        detail = [(item["part"], item["relationship_id"], item["errors"]) for item in invalid]
        raise EquationConversionError(f"Unusable Microsoft Equation 3 object: {detail}")
    targets, roots = _embedding_targets(parts, legacy, LEGACY_EQUATION_PROGID)
    payloads = {target: legacy_mtef(parts[target]) for target in targets}
    target_by_relationship = {(item["part"], item["relationship_id"]): item["embedding_part"] for item in legacy}
    anchors = {}
    for part, root in roots.items():
        next_id = 1 + max((int(node.get(q("w:id"))) for node in root.iter(q("w:bookmarkStart"))), default=0)
        for node in list(root.iter(q("o:OLEObject"))):
            target = target_by_relationship.get((part, node.get(q("r:id"))))
            if target is None:
                continue
            node.set("ProgID", MATHTYPE_PROGID)
            if anchor_for_insertion:
                run = next(node.iterancestors(q("w:r")), None)
                if (run is None or run.getparent().tag != q("w:p")
                        or len(list(run.iter(q("o:OLEObject")))) != 1):
                    raise EquationConversionError("A legacy equation needs its own inline Word run for template insertion")
                name = "PFEquation3" + uuid4().hex[:28]
                start = ET.Element(q("w:bookmarkStart"), {q("w:name"): name, q("w:id"): str(next_id)})
                end = ET.Element(q("w:bookmarkEnd"), {q("w:id"): str(next_id)})
                run.addprevious(start)
                run.addnext(end)
                anchors[name.casefold()] = payloads[target]
                next_id += 1
    document, pool, blank, clsid = _vendor_blank()
    for target, payload in payloads.items():
        parts[target] = _rewrap(blank, clsid, payload)
    blank = pool = document = None
    for part, root in roots.items():
        parts[part] = ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    write_package(parts, output)
    return {"converted_legacy_equation_objects": len(legacy),
            "mtef_sha256": [item["mtef_sha256"] for item in legacy]}, anchors


def restore_legacy_payloads(source, output, anchors):
    """Word InsertFile can upgrade MTEF. Restore each payload at its tracked object."""
    from .equations import EquationConversionError

    parts = read_package(source)
    audit = audit_docx_math(source)
    by_relationship = {(item["part"], item["relationship_id"]): item for item in audit["objects"]}
    seen, mapped, payloads, roots = set(), [], {}, {}
    for part in {item["part"] for item in audit["objects"]}:
        root = xml(parts[part])
        active, remove = {}, []
        for _, node in ET.iterwalk(root, events=("start",)):
            if node.tag == q("w:bookmarkStart"):
                name = node.get(q("w:name"), "").casefold()
                if name in anchors:
                    identifier = node.get(q("w:id"))
                    if name in seen or active or identifier is None:
                        raise EquationConversionError("Word duplicated or overlapped a legacy equation bookmark")
                    seen.add(name)
                    active[identifier] = (name, [])
                    remove.append(node)
            elif node.tag == q("o:OLEObject") and active:
                next(iter(active.values()))[1].append(by_relationship.get((part, node.get(q("r:id")))))
            elif node.tag == q("w:bookmarkEnd") and (identifier := node.get(q("w:id"))) in active:
                name, found = active.pop(identifier)
                if len(found) != 1 or found[0] is None or not found[0]["valid_native_storage"]:
                    raise EquationConversionError("Word changed the object tracked by a legacy equation bookmark")
                item = found[0]
                mapped.append(item)
                target = item["embedding_part"]
                if payloads.setdefault(target, anchors[name]) != anchors[name]:
                    raise EquationConversionError("Word merged distinct legacy equation payloads")
                remove.append(node)
        if active:
            raise EquationConversionError("Word lost a closing legacy equation bookmark")
        if remove:
            for node in remove:
                node.getparent().remove(node)
            roots[part] = root
    if seen != anchors.keys():
        raise EquationConversionError("Word lost a legacy equation bookmark during template insertion")
    _embedding_targets(parts, mapped, MATHTYPE_PROGID)
    changed = {item["embedding_part"] for item in mapped
               if item["mtef_sha256"] != digest(payloads[item["embedding_part"]])}
    if changed:
        document, pool, blank, clsid = _vendor_blank()
        for target in changed:
            parts[target] = _rewrap(blank, clsid, payloads[target])
        blank = pool = document = None
    for part, root in roots.items():
        parts[part] = ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    write_package(parts, output)
    return {"matched_equations": len(mapped), "restored_embeddings": len(changed)}


def convert_with_legacy_equations(source, output, parts, before, *, app, equations, source_digest):
    from .equations import EquationConversionError, _check_native_result, _native_fingerprints, convert_to_mathtype
    from .word import _publish, _stage_path

    converted, ready = _stage_path(output), _stage_path(output)
    try:
        legacy, _ = rewrap_legacy_equations(parts, before["legacy_objects"], converted)
        count = legacy["converted_legacy_equation_objects"]
        report = convert_to_mathtype(converted, ready, app=app, equations=equations)
        after = report["math_audit"]
        expected = _check_native_result(before, after, before["omml_equations"] + count)
        fingerprints = _native_fingerprints(source, before)
        fingerprints.update(legacy["mtef_sha256"])
        if fingerprints - _native_fingerprints(ready, after):
            raise EquationConversionError("A legacy or existing MathType payload changed during conversion")
        if digest(source.read_bytes()) != source_digest:
            raise EquationConversionError("Source changed during conversion; output was not published")
        _publish(ready, output)
        return {**report, "output": str(output), "source": str(source), "source_sha256": source_digest,
                "converted_legacy_equation_objects": count, "legacy_mtef_preserved": True,
                "preserved_native_count": before["native_mathtype_objects"], "expected_native_count": expected}
    finally:
        converted.unlink(missing_ok=True)
        ready.unlink(missing_ok=True)
