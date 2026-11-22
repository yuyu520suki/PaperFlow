"""Replace DOCX OMML with editable native MathType equations in a new copy.

Pandoc is a read-only source extractor here. It never regenerates the DOCX,
which would risk flattening Zotero fields and existing native OLE objects.
Word/MathType operations belong to the caller's serialized COM channel.
"""
from __future__ import annotations

from collections import Counter
import hashlib
import json
from pathlib import Path
import re
import shutil
import struct
import subprocess
from typing import Any, Mapping, Sequence
from zipfile import ZipFile

from lxml import etree as ET

from .ingest import find_pandoc
from .mathtype import (MATHTYPE_PROGID, MathTypeError, MathTypeSDK,
                       _CompoundFile, audit_docx_math, discover_mathtype)
from .package import q, read_package, xml
from .word import (_application, _output_path, _publish, _save_docx,
                   _source_path, _stage_path, _stories)


class EquationConversionError(MathTypeError):
    """A source equation cannot be converted without losing its identity."""


_CASES_MATRIX = re.compile(
    r"\\left\\\{\s*\\begin\{matrix\}(.*?)\\end\{matrix\}\s*\\right\.(?:\\ )?", re.S
)


def extract_math_from_ast(ast: Mapping[str, Any]) -> list[str]:
    """Read Math nodes in Pandoc's body order, including tables and lists.

    Math moved into document metadata, notes, or raw Office fragments cannot be
    mapped unambiguously to Word's main-story OMaths collection and is rejected.
    """
    if not isinstance(ast, Mapping) or not isinstance(ast.get("blocks"), list):
        raise EquationConversionError("Pandoc did not return a document AST")
    equations: list[str] = []

    def walk(value: Any, context: str, collect: bool = True) -> None:
        if isinstance(value, Mapping):
            kind = value.get("t")
            if kind == "Math":
                if not collect:
                    raise EquationConversionError(f"Unsupported equation location in Pandoc AST: {context}")
                content = value.get("c")
                if (not isinstance(content, list) or len(content) != 2
                        or not isinstance(content[0], Mapping)
                        or content[0].get("t") not in {"InlineMath", "DisplayMath"}
                        or not isinstance(content[1], str) or not content[1].strip()
                        or "\0" in content[1]):
                    raise EquationConversionError("Pandoc returned an empty or unsupported Math node")
                equations.append(content[1])
                return
            if kind == "Note":
                walk(value.get("c"), "note", collect=False)
                return
            if kind in {"RawInline", "RawBlock"}:
                content = value.get("c")
                if (isinstance(content, list) and len(content) == 2
                        and isinstance(content[1], str)
                        and any(token in content[1] for token in ("<m:oMath", "<mml:math", "<math"))):
                    raise EquationConversionError("Pandoc left an equation in unsupported raw markup")
            for key, child in value.items():
                if key not in {"t", "pandoc-api-version"}:
                    walk(child, context, collect=collect)
        elif isinstance(value, list):
            for child in value:
                walk(child, context, collect=collect)

    walk(ast.get("meta", {}), "metadata", collect=False)
    walk(ast["blocks"], "body")
    return equations


def _math_layout(source: Path, *, parts=None, validated_note_parts=()) -> dict[str, Any]:
    """Find every OMML object before Word is permitted to edit a copy."""
    parts = read_package(source) if parts is None else parts
    count = 0
    alignment_arrays = []
    noncentered_matrices = []
    unsupported = []
    for name, data in parts.items():
        if name in validated_note_parts or not name.startswith("word/") or not name.endswith(".xml"):
            continue
        root = xml(data)
        for equation in root.iter(q("m:oMath")):
            ancestors = list(equation.iterancestors())
            reason = None
            if name != "word/document.xml":
                reason = "non-main document part"
            elif any(a.tag == q("m:oMath") for a in ancestors):
                reason = "nested OMath object"
            elif any(a.tag in {q("w:txbxContent"), q("w:drawing"), q("w:pict")} for a in ancestors):
                reason = "textbox, drawing or legacy floating object"
            elif not any(a.tag == q("w:body") for a in ancestors):
                reason = "outside the main document body"
            if reason:
                unsupported.append({"part": name, "reason": reason})
            else:
                count += 1
                if equation.find(f".//{q('m:eqArr')}") is not None:
                    alignment_arrays.append(count)
                if any(element.get(q("m:val")) in {"left", "right"}
                       for element in equation.iter(q("m:mcJc"))):
                    noncentered_matrices.append(count)
    if unsupported:
        raise EquationConversionError(f"Unsupported OMML locations require explicit handling: {unsupported}")
    return {"main_omml_count": count, "alignment_arrays": alignment_arrays,
            "noncentered_matrices": noncentered_matrices}


def extract_docx_equations(source_docx: str | Path) -> list[str]:
    """Ask installed Pandoc for JSON only; never create an intermediate DOCX."""
    source = _source_path(source_docx)
    expected = _math_layout(source)["main_omml_count"]
    if not expected:
        return []
    result = subprocess.run(
        [find_pandoc(), str(source), "--from=docx", "--to=json"],
        capture_output=True, check=False, timeout=60,
    )
    if result.returncode:
        raise EquationConversionError("Pandoc could not extract equations: "
                                      + result.stderr.decode("utf-8", errors="replace"))
    # A warning can signal dropped or unsupported OMML. Do not report a native
    # conversion as successful after the extractor admitted a lossy read.
    if result.stderr.strip():
        raise EquationConversionError("Pandoc source extraction requires review: "
                                      + result.stderr.decode("utf-8", errors="replace"))
    try:
        equations = extract_math_from_ast(json.loads(result.stdout))
    except (ValueError, UnicodeError) as exc:
        raise EquationConversionError("Pandoc returned invalid equation JSON") from exc
    _check_equation_alignment(equations, expected)
    return equations


def _office_mathml_sources(source: Path, office_directory: str) -> list[str]:
    stylesheet = Path(office_directory) / "OMML2MML.XSL"
    if not stylesheet.is_file():
        raise EquationConversionError("Installed Word OMML-to-MathML stylesheet is missing")
    root = xml(read_package(source)["word/document.xml"])
    try:
        transform = ET.XSLT(ET.parse(str(stylesheet)), access_control=ET.XSLTAccessControl.DENY_ALL)
        return [ET.tostring(transform(equation).getroot(), encoding="unicode")
                for equation in root.iter(q("m:oMath"))]
    except (ET.Error, TypeError) as exc:
        raise EquationConversionError(f"Word's stylesheet could not convert the original OMML: {exc}") from exc


def _check_equation_alignment(equations: Sequence[str], xml_count: int,
                               word_count: int | None = None) -> None:
    if isinstance(equations, (str, bytes)):
        raise EquationConversionError("Equations must be an ordered sequence of LaTeX strings")
    if any(not isinstance(e, str) or not e.strip() or "\0" in e for e in equations):
        raise EquationConversionError("Each OMML object needs one non-empty LaTeX source")
    if len(equations) != xml_count:
        raise EquationConversionError(
            f"Equation source/XML mismatch: {len(equations)} LaTeX sources for {xml_count} OMML objects")
    if word_count is not None and word_count != xml_count:
        raise EquationConversionError(
            f"Word/XML equation mismatch: Word reports {word_count}, XML reports {xml_count}")


def _native_fingerprints(path: Path, audit: Mapping[str, Any]) -> Counter[str]:
    fingerprints: Counter[str] = Counter()
    with ZipFile(path) as package:
        for item in audit["objects"]:
            native = _CompoundFile(package.read(item["embedding_part"])).stream("Equation Native")
            header_size = struct.unpack_from("<I", native, 0)[0]
            fingerprints[hashlib.sha256(native[header_size:]).hexdigest()] += 1
    return fingerprints


def _check_native_result(before: Mapping[str, Any], after: Mapping[str, Any],
                           converted_count: int) -> int:
    expected = before["native_mathtype_objects"] + converted_count
    if (after["omml_equations"] != 0 or after["native_mathtype_objects"] != expected
            or after.get("legacy_equation_objects", 0)
            or not after["all_mathtype_objects_valid"]):
        raise EquationConversionError(
            "MathType output failed native audit: "
            f"expected {expected}, actual {after['native_mathtype_objects']}, "
            f"residual OMML {after['omml_equations']}, legacy {after.get('legacy_equation_objects', 0)}")
    return expected


def convert_to_mathtype(
    source_docx: str | Path, output_docx: str | Path, *,
    app: Any = None, equations: Sequence[str] | None = None,
) -> dict[str, Any]:
    """Convert OMML and legacy Equation 3, preserving prior MathType objects.

    Explicit equations correspond one-for-one to the source's Word OMaths in
    document order. With no explicit list, installed Pandoc extracts LaTeX.
    Unsupported stories and nested equations fail before mutation. The caller owns any
    supplied Word application and serializes MathType operations on that thread.
    A returned success proves structural native storage, not visual or editing
    fidelity; those remain separate render and reopen acceptance steps.
    """
    source = _source_path(source_docx)
    if source.suffix.lower() != ".docx":
        raise ValueError("MathType conversion source must be DOCX")
    if Path(output_docx).suffix.lower() != ".docx":
        raise ValueError("MathType conversion output must be DOCX")
    output = _output_path(output_docx, (source,))
    source_digest = hashlib.sha256(source.read_bytes()).hexdigest()
    from .note_equations import collect_note_equations, convert_with_note_equations

    parts = read_package(source)
    note_layout = collect_note_equations(parts)
    layout = _math_layout(source, parts=parts, validated_note_parts={part for part, _ in note_layout.values()})
    xml_count = layout["main_omml_count"]
    before = audit_docx_math(source)
    if not before["all_mathtype_objects_valid"]:
        raise EquationConversionError("Source contains an invalid existing MathType embedding")
    if equations is not None:
        _check_equation_alignment(equations, xml_count)
    if before["legacy_equation_objects"]:
        from .legacy_equations import convert_with_legacy_equations

        return convert_with_legacy_equations(source, output, parts, before, app=app, equations=equations,
                                             source_digest=source_digest)
    if note_layout:
        return convert_with_note_equations(source, output, note_layout, app=app, equations=equations,
                                           source_digest=source_digest)
    preserved = _native_fingerprints(source, before)
    sources = extract_docx_equations(source) if equations is None else list(equations) if not isinstance(equations, (str, bytes)) else equations
    _check_equation_alignment(sources, xml_count)
    stage = _stage_path(output)
    mathml_fallbacks = []
    mathml_alignments = []
    texvc_alignments = []
    texvc_cases = []
    try:
        if not xml_count:
            with source.open("rb") as src, stage.open("xb") as dst:
                shutil.copyfileobj(src, dst)
        else:
            installation = discover_mathtype()
            if not installation.native_sdk_available:
                raise EquationConversionError("Installed MathType native SDK is unavailable")
            blank = Path(installation.directory) / "Office Support" / "BlankEqn.doc"
            if not blank.is_file():
                raise EquationConversionError("Installed vendor BlankEqn.doc is missing")
            with _application(app) as word:
                doc = None
                shape = location = insertion = ole = story = equation = None
                try:
                    doc = word.Documents.Open(
                        FileName=str(source), ConfirmConversions=False, ReadOnly=True,
                        AddToRecentFiles=False, Visible=False, NoEncodingDialog=True,
                        OpenAndRepair=False,
                    )
                    _check_equation_alignment(sources, xml_count, int(doc.OMaths.Count))
                    # XML is authoritative for package-wide coverage; live Word
                    # adds a second guard for headers, notes and linked stories.
                    for story in _stories(doc):
                        if int(story.StoryType) != 1 and int(story.OMaths.Count):
                            raise EquationConversionError(
                                f"Non-main Word story {story.StoryType} contains OMML equations")
                    story = None
                    positions = []
                    for index in range(1, xml_count + 1):
                        equation = doc.OMaths.Item(index)
                        if int(equation.Range.StoryType) != 1:
                            raise EquationConversionError("Word OMath collection includes a non-main equation")
                        positions.append((int(equation.Range.Start), int(equation.Range.End)))
                    equation = None
                    if any(a[1] > b[0] for a, b in zip(positions, positions[1:])):
                        raise EquationConversionError("Word OMath ranges overlap or are out of document order")
                    mathml_sources = None

                    def mathml(index):
                        nonlocal mathml_sources
                        if mathml_sources is None:
                            mathml_sources = _office_mathml_sources(source, str(word.Path))
                            _check_equation_alignment(mathml_sources, xml_count)
                        return mathml_sources[index - 1]

                    with MathTypeSDK(installation) as sdk:
                        for index in range(xml_count, 0, -1):
                            location = doc.OMaths.Item(index).Range.Duplicate
                            start = int(location.Start)
                            # Word's smart deletion also removes the adjacent
                            # space in inline prose. Set the exact range instead.
                            location.Text = ""
                            insertion = doc.Range(Start=start, End=start)
                            insertion.InsertFile(
                                FileName=str(blank), Range="MTBlankEqn",
                                ConfirmConversions=False, Link=False, Attachment=False,
                            )
                            matches = []
                            for number in range(1, int(doc.InlineShapes.Count) + 1):
                                candidate = doc.InlineShapes.Item(number)
                                if int(candidate.Range.Start) == start:
                                    matches.append(candidate)
                                candidate = None
                            if len(matches) != 1:
                                raise EquationConversionError("Vendor blank insertion did not produce exactly one inline object")
                            shape = matches[0]
                            matches = None
                            if int(shape.Type) != 1 or str(shape.OLEFormat.ProgID) != MATHTYPE_PROGID:
                                raise EquationConversionError("Vendor blank insertion is not a native MathType object")
                            shape.OLEFormat.DoVerb(2)
                            ole = shape.OLEFormat.Object
                            written = False
                            try:
                                text = sources[index - 1]
                                texvc_align = False
                                texvc_case = False
                                use_mathml = False
                                if equations is None:
                                    if index in layout["noncentered_matrices"]:
                                        # Pandoc expresses a Word cases block as a fenced
                                        # matrix. Texvc's cases form preserves the full brace
                                        # and left alignment, which generic MathML can lose.
                                        text, case_count = _CASES_MATRIX.subn(
                                            lambda match: r"\begin{cases}" + match[1] + r"\end{cases}", text
                                        )
                                        texvc_case = case_count > 0
                                        use_mathml = not texvc_case
                                    if index in layout["alignment_arrays"]:
                                        # The installed Texvc importer accepts align but
                                        # renders the aligned environment as literal text.
                                        texvc_align = r"\begin{aligned}" in text and r"\end{aligned}" in text
                                        if texvc_align:
                                            text = text.replace(r"\begin{aligned}", r"\begin{align}")
                                            text = text.replace(r"\end{aligned}", r"\end{align}")
                                        else:
                                            use_mathml = True
                                if use_mathml:
                                    try:
                                        sdk.set_ole_equation(ole, mathml(index), "mathml", save_close=False)
                                    except (MathTypeError, ValueError) as exc:
                                        raise EquationConversionError(f"MathType rejected formula {index} as MathML: {exc}") from exc
                                    mathml_alignments.append(index)
                                else:
                                    try:
                                        sdk.set_ole_equation(ole, text, "latex", save_close=False)
                                        if texvc_align:
                                            texvc_alignments.append(index)
                                        if texvc_case:
                                            texvc_cases.append(index)
                                    except MathTypeError as tex_error:
                                        if equations is not None:
                                            raise EquationConversionError(f"MathType rejected explicit formula {index}: {tex_error}") from tex_error
                                        sdk.close_ole_equation(ole, save_changes=False)
                                        ole = None
                                        shape.OLEFormat.DoVerb(2)
                                        ole = shape.OLEFormat.Object
                                        try:
                                            sdk.set_ole_equation(ole, mathml(index), "mathml", save_close=False)
                                        except (MathTypeError, ValueError) as mathml_error:
                                            raise EquationConversionError(
                                                f"MathType rejected formula {index} as TeX and as MathML: {tex_error}; {mathml_error}"
                                            ) from mathml_error
                                        mathml_fallbacks.append(index)
                                written = True
                                shape.Range.Fields.Update()
                                import pythoncom
                                pythoncom.PumpWaitingMessages()
                            finally:
                                if ole is not None:
                                    sdk.close_ole_equation(ole, save_changes=written)
                            ole = shape = location = insertion = None
                    if int(doc.OMaths.Count):
                        raise EquationConversionError("Word still contains unconverted OMML equations")
                    _save_docx(doc, stage)
                finally:
                    ole = shape = location = insertion = equation = story = None
                    if doc is not None:
                        doc.Close(SaveChanges=0)
                        doc = None
        after = audit_docx_math(stage)
        expected = _check_native_result(before, after, xml_count)
        remaining = _native_fingerprints(stage, after)
        if preserved - remaining:
            raise EquationConversionError("An existing MathType equation's native data changed or disappeared")
        if hashlib.sha256(source.read_bytes()).hexdigest() != source_digest:
            raise EquationConversionError("Source changed during conversion; output was not published")
        _publish(stage, output)
        return {
            "output": str(output), "source": str(source),
            "converted_omml_count": xml_count,
            "converted_legacy_equation_objects": 0,
            "office_mathml_fallback_formula_numbers": sorted(mathml_fallbacks),
            "office_mathml_alignment_formula_numbers": sorted(mathml_alignments),
            "texvc_alignment_formula_numbers": sorted(texvc_alignments),
            "texvc_cases_formula_numbers": sorted(texvc_cases),
            "preserved_native_count": before["native_mathtype_objects"],
            "expected_native_count": expected,
            "native_count_matches": after["native_mathtype_objects"] == expected,
            "residual_omml_count": after["omml_equations"],
            "source_sha256": source_digest, "source_unchanged": True,
            "native_data_preserved": True,
            "math_audit": after,
            "visual_fidelity_verified": False, "edit_roundtrip_verified": False,
        }
    finally:
        stage.unlink(missing_ok=True)
