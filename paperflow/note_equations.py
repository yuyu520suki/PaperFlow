"""Transfer native MathType objects to the exact original Word note ranges."""

from collections import Counter
from copy import deepcopy
from pathlib import Path
from tempfile import TemporaryDirectory

from .notes import KINDS, normal_notes, note_ids, note_parts
from .package import digest, q, read_package, xml


_STORY = {"footnote": ("Footnotes", 2), "endnote": ("Endnotes", 3)}
_UNSUPPORTED = {q("w:" + tag) for tag in ("txbxContent", "drawing", "pict", "del", "moveFrom")}


def collect_note_equations(parts):
    """Order each note's equations by the native references in the body."""
    from .equations import EquationConversionError

    result = {}
    body = xml(parts["word/document.xml"])
    stories = note_parts(parts)
    for kind in KINDS:
        if kind not in stories:
            continue
        part, root = stories[kind]
        equations = set(root.iter(q("m:oMath")))
        if not equations:
            continue
        notes = normal_notes(root, kind)
        by_id = {note.get(q("w:id")): note for note in notes}
        references = note_ids(body, kind)
        if (len(by_id) != len(notes) or len(set(references)) != len(references)
                or any(identifier is None or identifier not in by_id for identifier in references)):
            raise EquationConversionError(f"Ambiguous {kind} references cannot be mapped to MathType equations")
        groups = tuple(tuple(by_id[identifier].iter(q("m:oMath"))) for identifier in references)
        if {equation for group in groups for equation in group} != equations:
            raise EquationConversionError(f"An unreferenced or special {kind} contains an equation")
        for equation in equations:
            ancestors = {parent.tag for parent in equation.iterancestors()}
            if q("m:oMath") in ancestors or q("w:p") not in ancestors or ancestors & _UNSUPPORTED:
                raise EquationConversionError(f"An unsupported {kind} equation is nested or in a noneditable range")
        result[kind] = (part, groups)
    return result


def _donor_document(path, layout):
    from docx import Document

    donor = Document()
    for _, groups in layout.values():
        for group in groups:
            for equation in group:
                donor.add_paragraph()._p.append(deepcopy(equation))
    donor.save(path)


def _transfer_notes(word, source, donor, output, layout):
    from .equations import EquationConversionError
    from .mathtype import MATHTYPE_PROGID
    from .word import _opened, _save_docx

    with _opened(word, donor) as library, _opened(word, source) as document:
        plan = []
        for kind, (_, groups) in layout.items():
            collection, story_type = _STORY[kind]
            notes = getattr(document, collection)
            if int(notes.Count) != len(groups):
                raise EquationConversionError(f"Word and source have different {kind} counts")
            previous = -1
            for index, group in enumerate(groups, 1):
                note = notes.Item(index)
                span = note.Range
                position = int(note.Reference.Start)
                if position <= previous or int(span.StoryType) != story_type or int(span.OMaths.Count) != len(group):
                    raise EquationConversionError(f"Cannot align source and Word equations in {kind} {index}")
                previous = position
                end = int(span.Start)
                for number in range(1, len(group) + 1):
                    formula = span.OMaths.Item(number).Range
                    if not end <= int(formula.Start) < int(formula.End) <= int(span.End):
                        raise EquationConversionError(f"Word equations overlap in {kind} {index}")
                    end = int(formula.End)
                    plan.append((collection, index, number, len(plan) + 1))
        if int(library.InlineShapes.Count) != len(plan):
            raise EquationConversionError("MathType donor objects do not cover every note equation")
        for collection, index, number, donor_index in reversed(plan):
            shape = library.InlineShapes.Item(donor_index)
            if int(shape.Type) != 1 or str(shape.OLEFormat.ProgID) != MATHTYPE_PROGID:
                raise EquationConversionError("A note donor shape is not a native MathType object")
            location = getattr(document, collection).Item(index).Range.OMaths.Item(number).Range.Duplicate
            location.FormattedText = shape.Range.FormattedText
        location = shape = formula = span = note = notes = None
        _save_docx(document, output)


def _part_fingerprints(audit, part):
    return Counter(item["mtef_sha256"] for item in audit["objects"] if item["part"] == part)


def convert_with_note_equations(source, output, layout, *, app, equations, source_digest):
    """Build native note objects first, then use the ordinary body converter."""
    from .equations import EquationConversionError, _check_native_result, _native_fingerprints, convert_to_mathtype
    from .mathtype import audit_docx_math
    from .word import _application, _publish

    before = audit_docx_math(source)
    if not before["all_mathtype_objects_valid"]:
        raise EquationConversionError("Source contains an invalid existing MathType embedding")
    counts = {kind: sum(map(len, groups)) for kind, (_, groups) in layout.items()}
    with TemporaryDirectory(prefix=".paperflow-note-math-", dir=output.parent) as folder, _application(app) as word:
        scratch = Path(folder)
        donor_source = scratch / "donor-source.docx"
        donor = scratch / "donor-native.docx"
        notes_native = scratch / "notes-native.docx"
        ready = scratch / "ready.docx"
        _donor_document(donor_source, layout)
        donor_audit = convert_to_mathtype(donor_source, donor, app=word)["math_audit"]
        _transfer_notes(word, source, donor, notes_native, layout)
        transferred = audit_docx_math(notes_native)
        converted_notes = sum(counts.values())
        if (transferred["omml_equations"] != before["omml_equations"] - converted_notes
                or transferred["native_mathtype_objects"] != before["native_mathtype_objects"] + converted_notes
                or not transferred["all_mathtype_objects_valid"]):
            raise EquationConversionError("Word changed the equation inventory while transferring note equations")
        destinations = note_parts(read_package(notes_native))
        expected_by_kind = {}
        start = 0
        for kind, count in counts.items():
            target_part = destinations.get(kind, (None,))[0]
            expected = _part_fingerprints(before, layout[kind][0])
            expected.update(item["mtef_sha256"] for item in donor_audit["objects"][start:start + count])
            if _part_fingerprints(transferred, target_part) != expected:
                raise EquationConversionError(f"The native MathType payloads changed while transferring to {kind}s")
            expected_by_kind[kind] = expected
            start += count
        report = convert_to_mathtype(notes_native, ready, app=word, equations=equations)
        after = report["math_audit"]
        total = report["converted_omml_count"] + converted_notes
        if total != before["omml_equations"]:
            raise EquationConversionError("MathType conversion did not cover every body and note equation")
        expected_native = _check_native_result(before, after, total)
        final_destinations = note_parts(read_package(ready))
        for kind, expected in expected_by_kind.items():
            if _part_fingerprints(after, final_destinations.get(kind, (None,))[0]) != expected:
                raise EquationConversionError(f"Body conversion changed the native MathType payloads in {kind}s")
        if _native_fingerprints(source, before) - _native_fingerprints(ready, after):
            raise EquationConversionError("An existing MathType equation changed during note conversion")
        if digest(source.read_bytes()) != source_digest:
            raise EquationConversionError("Source changed during conversion; output was not published")
        _publish(ready, output)
    return {**report, "output": str(output), "source_sha256": source_digest,
            "converted_omml_count": total, "converted_note_equations": counts,
            "preserved_native_count": before["native_mathtype_objects"], "expected_native_count": expected_native}
