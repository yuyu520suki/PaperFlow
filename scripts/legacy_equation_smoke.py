"""Exercise synthetic Equation Editor 3.0 conversion without loading its server."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import struct
import sys
import uuid

from docx import Document
from lxml import etree as ET
import pymupdf


PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from paperflow.equations import convert_to_mathtype
from paperflow.mathtype import MATHTYPE_CLSID, MATHTYPE_PROGID, MathTypeSDK, _CompoundFile, audit_docx_math
from paperflow.package import digest, q, read_package, write_package, xml
from paperflow.pipeline import run_pipeline
from paperflow.template import _internal_target, _part_relationships
from paperflow.word import _opened, _save_docx, export_pdf, private_word


LEGACY_PROGID = "Equation.3"
LEGACY_CLSID = "0002ce02-0000-0000-c000-000000000046"
LEGACY_TEXT = "Legacy Equation Editor formula"
NATIVE_TEXT = "Existing native MathType formula"
OMML_TEXT = "New Word equation"
TEMPLATE_TEXT = "Legacy formula migration fixture"
_MAIN = "word/document.xml"
# Synthetic x+1=2, assembled from the recovered vendor MTEF 3 format spec.
# https://rtf2latex2e.sourceforge.net/MTEF3.html
V3_EQUATION = bytes.fromhex("03 01 01 03 00 0a 01 02 83 78 00 02 86 2b 00 02 88 31 00 02 86 3d 00 02 88 32 00 00 00")


def _add_math(paragraph, expression):
    equation = ET.SubElement(paragraph, q("m:oMath"))
    ET.SubElement(ET.SubElement(equation, q("m:r")), q("m:t")).text = expression


def _make_donor(path):
    doc = Document()
    _add_math(doc.add_paragraph(LEGACY_TEXT + ": ")._p, "x+1=2")
    _add_math(doc.add_paragraph(NATIVE_TEXT + ": ")._p, "u=4")
    doc.save(path)


def _make_template(path):
    doc = Document()
    doc.add_paragraph(TEMPLATE_TEXT, style="Title")
    doc.add_paragraph("{{PAPERFLOW_BODY}}")
    doc.sections[0].header.paragraphs[0].text = "Template header witness"
    doc.sections[0].footer.paragraphs[0].text = "Template footer witness"
    doc.save(path)


def _prefixed_ansi(text):
    raw = text.encode("ascii") + b"\0"
    return struct.pack("<I", len(raw)) + raw


def _synthetic_legacy_storage(native_bytes):
    """Build an Equation.3 container from data, with no OLE object activation."""
    import pythoncom
    import pywintypes

    donor = _CompoundFile(native_bytes)
    native = donor.stream("Equation Native")
    clipboard_format = struct.unpack_from("<H", native, 6)[0]
    # EQNOLEFILEHDR documented by Design Science, followed by the exact MTEF.
    old_stream = struct.pack("<HIHI4I", 28, 0x00020000, clipboard_format, len(V3_EQUATION), 0, 0, 0, 0) + V3_EQUATION
    # MS-OLEDS 2.3.8. Rebuild the type/progid fields using legacy identity.
    old_compobj = (
        donor.stream("\x01CompObj")[:12] + uuid.UUID(LEGACY_CLSID).bytes_le
        + _prefixed_ansi("Microsoft Equation 3.0") + _prefixed_ansi("MathType EF") + _prefixed_ansi(LEGACY_PROGID)
        + struct.pack("<4I", 0x71B239F4, 0, 0, 0)
    )
    memory = pythoncom.CreateILockBytesOnHGlobal()
    storage = pythoncom.StgCreateDocfileOnILockBytes(memory, 0x1012)
    storage.SetClass(pywintypes.IID("{" + LEGACY_CLSID + "}"))
    for name, payload in {
        "\x01Ole": donor.stream("\x01Ole"), "\x01CompObj": old_compobj,
        "\x03ObjInfo": b"\x00\x00\x03\x00\x04\x00", "Equation Native": old_stream,
    }.items():
        stream = storage.CreateStream(name, 0x1012, 0, 0)
        stream.Write(payload)
        stream.Commit(0)
        stream = None
    storage.Commit(0)
    storage = None
    return memory.ReadAt(0, memory.Stat(1)[2])


def _object_paths(parts, document, progid):
    rels = _part_relationships(parts, _MAIN)
    return [(item, _internal_target(_MAIN, rels[item.get(q("r:id"))].get("Target")))
            for item in document.iter(q("o:OLEObject")) if item.get("ProgID") == progid]


def _make_legacy_carrier(donor, output, *, include_omml):
    parts = read_package(donor)
    document = xml(parts[_MAIN])
    first, other = _object_paths(parts, document, MATHTYPE_PROGID)
    if first[1] == other[1]:
        raise RuntimeError("Expected separate donor embeddings")
    first[0].set("ProgID", LEGACY_PROGID)
    parts[first[1]] = _synthetic_legacy_storage(parts[first[1]])
    if include_omml:
        body = document.find(q("w:body"))
        paragraph = ET.Element(q("w:p"))
        text = ET.SubElement(ET.SubElement(paragraph, q("w:r")), q("w:t"))
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text.text = OMML_TEXT + ": "
        _add_math(paragraph, "t=7")
        section = body.find(q("w:sectPr"))
        body.insert(body.index(section), paragraph)
    parts[_MAIN] = ET.tostring(document, encoding="UTF-8", xml_declaration=True, standalone=True)
    write_package(parts, output)


def _inspect_legacy(path):
    parts = read_package(path)
    records = []
    for _, target in _object_paths(parts, xml(parts[_MAIN]), LEGACY_PROGID):
        storage = _CompoundFile(parts[target])
        raw = storage.stream("Equation Native")
        header, _, _, length = struct.unpack_from("<HIHI", raw)
        payload = raw[header:header + length]
        records.append({"clsid": storage.root["clsid"], "mtef_version": payload[0], "mtef_sha256": digest(payload),
                        "legacy_compobj": b"Equation.3\0" in storage.stream("\x01CompObj")})
    return records


def _edit_first(app, master, destination, count):
    with _opened(app, master) as document:
        shapes = document.InlineShapes
        first = shapes.Item(1)
        if int(shapes.Count) != count or first.OLEFormat.ProgID != MATHTYPE_PROGID:
            raise RuntimeError(f"The converted source must have {count} native inline equations")
        with MathTypeSDK() as sdk:
            first.OLEFormat.DoVerb(2)
            ole = first.OLEFormat.Object
            sdk.set_ole_equation(ole, r"+\eta", "latex")
            ole = None
        shapes = first = None
        _save_docx(document, destination)


def _visible_formula(page, witness):
    rects = page.search_for(witness + ":")
    if len(rects) != 1:
        return ""
    marker = rects[0]
    # Word's vector picture text can be stored out of reading order.
    symbols = [item for block in page.get_text("rawdict")["blocks"] for line in block.get("lines", [])
               for span in line["spans"] for item in span.get("chars", [])
               if item["bbox"][0] >= marker.x1 - 0.25 and marker.y0 - 4 <= item["origin"][1] <= marker.y1 + 4]
    return "".join(item["c"] for item in sorted(symbols, key=lambda item: item["origin"][0]) if not item["c"].isspace())


def _render(path, directory, label, witnesses):
    pages = []
    with pymupdf.open(path) as document:
        for index, page in enumerate(document, 1):
            image = directory / f"{label}-page-{index}.png"
            pixmap = page.get_pixmap(dpi=130, alpha=False)
            pixmap.save(image)
            pages.append({"preview": str(image), "text": page.get_text(),
                          "equations": {witness: _visible_formula(page, witness) for witness in witnesses},
                          "pixels_sha256": hashlib.sha256(pixmap.samples).hexdigest()})
    return pages


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path, help="New directory for the synthetic fixture and evidence")
    parser.add_argument("--binary-doc", action="store_true", help="Test Word 97–2003 import in the same workflow")
    args = parser.parse_args()
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    donor_source, donor, carrier, template = (output / name for name in (
        "donor-source.docx", "donor-mathtype.docx", "synthetic-equation3.docx", "template.docx",
    ))
    expected_formulas = {LEGACY_TEXT: "x+1=2", NATIVE_TEXT: "u=4"}
    if not args.binary_doc:
        expected_formulas[OMML_TEXT] = "t=7"
    count = len(expected_formulas)
    _make_donor(donor_source)
    _make_template(template)
    with private_word() as app:
        convert_to_mathtype(donor_source, donor, app=app)
        # Word 97–2003 cannot retain modern OMML when saving a binary DOC.
        _make_legacy_carrier(donor, carrier, include_omml=not args.binary_doc)
        source = carrier
        if args.binary_doc:
            source = output / "synthetic-equation3.doc"
            with _opened(app, carrier) as document:
                document.SaveAs2(FileName=str(source), FileFormat=0, AddToRecentFiles=False)
    original = {path: digest(path.read_bytes()) for path in (source, carrier, template)}
    built = output / "built"
    workflow = run_pipeline(source, template, built)
    master = built / "editable-master.docx"
    master_hash = digest(master.read_bytes())
    edited_path, edited_pdf = output / "math-edited.docx", output / "math-edited.pdf"
    with private_word() as app:
        _edit_first(app, master, edited_path, count)
    export_pdf(edited_path, edited_pdf)
    return _verify_output(output, source, carrier, original, workflow, master_hash, expected_formulas)


def _verify_output(output, source, carrier, original, workflow, master_hash, expected_formulas):
    count = len(expected_formulas)
    binary_doc = source.suffix.lower() == ".doc"
    source_legacy = _inspect_legacy(carrier)
    source_native = audit_docx_math(carrier)["objects"][0]["mtef_sha256"]
    provenance = {"kind": "synthetic_equation_editor_3_container", "carrier": str(carrier)}
    if binary_doc:
        storage = _CompoundFile(source.read_bytes())
        provenance.update(binary_document_identifier=storage.stream("WordDocument")[:2].hex(),
                          object_pool_classes=sorted(entry["clsid"] for entry in storage.entries.values()
                                                     if entry["type"] == 1 and entry["clsid"] in {LEGACY_CLSID, MATHTYPE_CLSID}))
    built = output / "built"
    imported_legacy = _inspect_legacy(built / "01-imported.docx")
    steps = {step["stage"]: step["evidence"] for step in workflow["steps"]}
    provenance["import_adapter"] = steps["import"]["adapter"]
    math_step = steps["mathtype"]
    master = built / "editable-master.docx"
    saved = audit_docx_math(master)
    edited_path, edited_pdf = output / "math-edited.docx", output / "math-edited.pdf"
    edited = audit_docx_math(edited_path)
    renders = _render(built / "editable-master.pdf", output, "original", expected_formulas)
    edited_renders = _render(edited_pdf, output, "edited", expected_formulas)
    saved_hashes = [item["mtef_sha256"] for item in saved["objects"]]
    edited_hashes = [item["mtef_sha256"] for item in edited["objects"]]
    checks = {
        "source_has_equation3_class_v3_payload_and_legacy_metadata": source_legacy == [{
            "clsid": LEGACY_CLSID, "mtef_version": 3, "mtef_sha256": digest(V3_EQUATION), "legacy_compobj": True,
        }],
        "word_import_keeps_legacy_mtef_until_conversion": len(imported_legacy) == 1
        and imported_legacy[0]["mtef_sha256"] == digest(V3_EQUATION),
        "pipeline_counts_legacy_conversion": math_step["converted_legacy_equation_objects"] == 1,
        "output_valid_mathtype_objects_without_omml_or_legacy": saved["all_mathtype_objects_valid"]
        and saved["native_mathtype_objects"] == count and saved["omml_equations"] == 0 and not _inspect_legacy(master),
        "legacy_mtef_and_existing_native_mtef_exact": len(saved_hashes) == count
        and saved_hashes[:2] == [digest(V3_EQUATION), source_native],
        "reopened_native_editor_changes_only_migrated_equation": len(edited_hashes) == len(saved_hashes) == count
        and edited_hashes[0] != saved_hashes[0] and edited_hashes[1:] == saved_hashes[1:],
        "word_render_matches_source_formulas": len(renders) == 1 and renders[0]["equations"] == expected_formulas,
        "edited_word_render_keeps_other_formulas": len(edited_renders) == 1
        and edited_renders[0]["equations"][LEGACY_TEXT].startswith(expected_formulas[LEGACY_TEXT] + "+")
        and all(edited_renders[0]["equations"][name] == value for name, value in expected_formulas.items() if name != LEGACY_TEXT),
        "render_preserves_template_and_surrounding_text": len(renders) == 1 and all(
            token in renders[0]["text"] for token in (
                TEMPLATE_TEXT, "Template header witness", "Template footer witness", *expected_formulas,
            )
        ),
        "input_and_master_unchanged_after_verification": digest(master.read_bytes()) == master_hash
        and all(digest(path.read_bytes()) == before for path, before in original.items()),
    }
    if binary_doc:
        checks["genuine_word_binary_uses_native_import_adapter"] = (
            provenance["binary_document_identifier"] == "eca5"
            and provenance["object_pool_classes"] == [LEGACY_CLSID, MATHTYPE_CLSID]
            and provenance["import_adapter"] == "word-native-import"
        )
    report = {"pass": all(checks.values()), "source_format": source.suffix, "checks": checks, "legacy_before": source_legacy,
              "provenance": provenance, "legacy_after_import": imported_legacy,
              "saved_native": saved["objects"], "edited_native": edited["objects"],
              "original_render": renders, "edited_render": edited_renders}
    evidence = output / "legacy-equation-evidence.json"
    evidence.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"pass": report["pass"], "checks": checks, "evidence": str(evidence)}, ensure_ascii=False, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
