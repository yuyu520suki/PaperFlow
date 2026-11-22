"""Windows Word automation that never attaches to a user's Word instance.

All writing operations produce a new file and refuse to overwrite an existing
file.  This module does not load add-ins, execute macros, change Trust Center
settings, or refresh Zotero fields.  InsertFile preserves Word objects rather
than extracting their display text; preservation still requires package audits.

Official API references:
https://learn.microsoft.com/en-us/office/vba/api/word.range.insertfile
https://learn.microsoft.com/en-us/office/vba/api/word.application.automationsecurity
https://learn.microsoft.com/en-us/office/vba/api/word.range.nextstoryrange
https://learn.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat
https://learn.microsoft.com/en-us/office/vba/api/word.tableofcontents.update
https://learn.microsoft.com/en-us/office/vba/api/word.tableoffigures.update
"""

from __future__ import annotations

from contextlib import contextmanager
import os
from pathlib import Path
import re
import shutil
from typing import Any, Iterator, Mapping
import uuid
import xml.etree.ElementTree as ET
from zipfile import ZipFile


class WordAutomationError(RuntimeError):
    """Word is unavailable or a requested native operation failed."""


class OutputExistsError(FileExistsError):
    """A writing operation would replace a pre-existing file."""


DOCX_FORMAT = 12
PDF_FORMAT = 17
_SAFE_FIELDS = {
    "SEQ", "REF", "PAGEREF", "PAGE", "NUMPAGES", "SECTION",
    "SECTIONPAGES", "STYLEREF", "TOC", "INDEX", "TOA",
}
_PAGE_FIELDS = {"PAGEREF", "PAGE", "NUMPAGES", "SECTIONPAGES"}
_TABLE_FIELDS = {"TOC", "INDEX", "TOA"}
_REF_TARGET = re.compile(r'^\s*REF\s+(?:"([^"]+)"|([^\s\\"]+))(?=\s|\\|$)', re.I)
_FIELD_ARGUMENTS = re.compile(r'"(?:\\.|[^"\\])*"|\\\\|\\[^\s\\"]+|[^\s\\"]+')
_BUILTIN_STYLES = {"Normal": -1, "Caption": -35}
_BUILTIN_STYLES.update({f"Heading {n}": -1 - n for n in range(1, 10)})
_BUILTIN_STYLES.update({f"Heading{n}": -1 - n for n in range(1, 10)})


def _source_path(value: str | Path) -> Path:
    path = Path(value).expanduser().resolve(strict=True)
    if not path.is_file():
        raise ValueError(f"Source is not a file: {path}")
    return path


def _output_path(value: str | Path, sources: tuple[Path, ...] = ()) -> Path:
    path = Path(value).expanduser().resolve()
    if any(os.path.normcase(str(path)) == os.path.normcase(str(s)) for s in sources):
        raise ValueError("Output must differ from every source file")
    if path.exists():
        raise OutputExistsError(f"Refusing to overwrite existing output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def _stage_path(output: Path) -> Path:
    return output.with_name(f".{output.stem}-{uuid.uuid4().hex}.paperflow{output.suffix}")


def _publish(stage: Path, output: Path) -> None:
    """Publish without a check/replace race; an existing destination is sacred."""
    if not stage.is_file() or not stage.stat().st_size:
        raise WordAutomationError(f"Word did not produce a non-empty file: {stage}")
    # NTFS hard-link creation is atomic and fails when the destination exists.
    # A different file system may not support links; exclusive creation is the
    # fallback. Never use os.replace, which could overwrite a racing writer.
    try:
        os.link(stage, output)
    except FileExistsError:
        raise OutputExistsError(f"Output appeared during operation: {output}") from None
    except OSError:
        created = False
        try:
            with output.open("xb") as target:
                created = True
                with stage.open("rb") as source:
                    shutil.copyfileobj(source, target)
        except BaseException:
            if created:
                output.unlink(missing_ok=True)
            raise
    stage.unlink()


@contextmanager
def _private_word_instance() -> Iterator[Any]:
    """Own a hidden DispatchEx instance; only that instance is closed on exit.

    COM automation is serialized by the caller. Keep this context and every
    object obtained from it on one thread. No GetActiveObject call is used.
    """
    if os.name != "nt":
        raise WordAutomationError("Native Word automation requires Windows")
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise WordAutomationError("The active Python runtime lacks pywin32") from exc
    pythoncom.CoInitialize()
    app = None
    previous_security = None
    try:
        try:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            app.ScreenUpdating = False
            previous_security = app.AutomationSecurity
            app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        except Exception as exc:
            raise WordAutomationError(f"Cannot start private Microsoft Word: {exc}") from exc
        yield app
    finally:
        try:
            if app is not None:
                try:
                    # These are only documents belonging to this private instance.
                    for n in range(app.Documents.Count, 0, -1):
                        app.Documents.Item(n).Close(SaveChanges=0)
                finally:
                    try:
                        if previous_security is not None:
                            app.AutomationSecurity = previous_security
                        # Do not persist incidental Normal.dotm changes on shutdown.
                        app.NormalTemplate.Saved = True
                    finally:
                        try:
                            if isinstance(app, win32com.client.dynamic.CDispatch):
                                # Some late-bound Word instances advertise Quit as a property.
                                app._FlagAsMethod("Quit")
                            app.Quit(0)
                        finally:
                            app = None
        finally:
            pythoncom.CoUninitialize()


@contextmanager
def private_word() -> Iterator[Any]:
    """Serialize a hidden owned instance across CLI, GUI and MCP processes."""
    from .native_lock import native_word_lock

    with native_word_lock(), _private_word_instance() as app:
        yield app


@contextmanager
def _application(app: Any | None) -> Iterator[Any]:
    if app is None:
        with private_word() as owned:
            yield owned
    else:
        yield app


@contextmanager
def _opened(app: Any, path: Path) -> Iterator[Any]:
    doc = app.Documents.Open(
        FileName=str(path), ConfirmConversions=False, ReadOnly=True,
        AddToRecentFiles=False, Visible=False, NoEncodingDialog=True,
        OpenAndRepair=False,
    )
    try:
        yield doc
    finally:
        doc.Close(SaveChanges=0)


def _save_docx(doc: Any, stage: Path) -> None:
    doc.SaveAs2(FileName=str(stage), FileFormat=DOCX_FORMAT, AddToRecentFiles=False)


def _trim_generated_terminal_paragraph(path: Path) -> bool:
    from lxml import etree as lxml

    from .package import q, read_package, write_package, xml

    parts = read_package(path)
    main = "word/document.xml"
    document = xml(parts[main])
    body = document.find(q("w:body"))
    end = len(body) - int(len(body) > 0 and body[-1].tag == q("w:sectPr"))
    if end < 2 or body[end - 2].tag != q("w:p") or body[end - 1].tag != q("w:p"):
        return False
    paragraph = body[end - 1]
    for child in paragraph:
        if child.tag == q("w:pPr"):
            if any(prop.tag not in {q("w:pStyle"), q("w:rPr")} for prop in child):
                return False
        elif child.tag != q("w:r") or any(prop.tag not in {q("w:rPr"), q("w:t")} or (prop.tag == q("w:t") and prop.text)
                                              for prop in child):
            return False
    body.remove(paragraph)
    parts[main] = lxml.tostring(document, encoding="UTF-8", xml_declaration=True, standalone=True)
    replacement = _stage_path(path)
    try:
        write_package(parts, replacement)
        replacement.replace(path)
    finally:
        replacement.unlink(missing_ok=True)
    return True


def convert_to_docx(input_path: str | Path, output_path: str | Path, *, app: Any = None) -> dict:
    """Use Word's import converter, preserving native fields/OLE when supported."""
    source = _source_path(input_path)
    output = _output_path(output_path, (source,))
    if output.suffix.lower() != ".docx":
        raise ValueError("Native conversion output must have .docx extension")
    stage = _stage_path(output)
    try:
        with _application(app) as word, _opened(word, source) as doc:
            _save_docx(doc, stage)
            result = snapshot(doc)
        _publish(stage, output)
        return {"output": str(output), "snapshot": result}
    finally:
        stage.unlink(missing_ok=True)


def _find_body_text(doc: Any, slot: str) -> Any:
    text = str(doc.Content.Text)
    occurrences = text.count(slot)
    if occurrences != 1:
        raise ValueError(f"Template marker must occur exactly once; found {occurrences}: {slot}")
    found = doc.Content.Duplicate
    found.Find.ClearFormatting()
    ok = found.Find.Execute(
        FindText=slot, Forward=True, Wrap=0, MatchCase=True,
        MatchWholeWord=False, MatchWildcards=False, Format=False,
    )
    if not ok:
        raise ValueError(f"Template marker could not be located: {slot}")
    return found


def _slot_range(doc: Any, slot: str) -> Any:
    if not slot:
        raise ValueError("An explicit template insertion slot is required")
    if slot.startswith("bookmark:"):
        name = slot.removeprefix("bookmark:")
        if not name or not doc.Bookmarks.Exists(name):
            raise ValueError(f"Template bookmark not found: {name}")
        return doc.Bookmarks.Item(name).Range.Duplicate
    found = _find_body_text(doc, slot)
    paragraph_text = str(found.Paragraphs.Item(1).Range.Text).strip(" \t\r\n\x07")
    if paragraph_text != slot:
        raise ValueError("Template body marker must be the only content in its paragraph")
    return found


def _style(doc: Any, name: str | int) -> Any:
    selector = _BUILTIN_STYLES.get(name, name) if isinstance(name, str) else name
    try:
        return doc.Styles.Item(selector)
    except Exception as exc:
        raise ValueError(f"Word style not found in template: {name}") from exc


def _style_name(style: Any) -> str:
    try:
        return str(style.NameLocal)
    except Exception:
        return str(style)


def _apply_style_map(doc: Any, start: int, end: int, mapping: Mapping[str | int, Any]) -> dict:
    """Only remap explicitly selected styles inside the inserted range.

    Keys are Word style names or built-in WdBuiltinStyle integer IDs, not
    OOXML styleId values. Direct run formatting is deliberately preserved;
    a separate audited normalization policy is needed to remove overrides.
    """
    localized = {}
    for key, destination in mapping.items():
        try:
            source_name = _style_name(_style(doc, key))
        except ValueError:
            source_name = str(key)
        localized[source_name] = destination
    counts: dict[str, int] = {}
    inserted = doc.Range(Start=start, End=end)
    for number in range(1, inserted.Paragraphs.Count + 1):
        para = inserted.Paragraphs.Item(number)
        source_name = _style_name(para.Style)
        if source_name in localized:
            para.Style = localized[source_name]
            counts[source_name] = counts.get(source_name, 0) + 1
    return counts


def read_zotero_preferences(source_docx: str | Path) -> dict[str, str]:
    """Read only Zotero's custom-property chunks from an OPC package."""
    source = _source_path(source_docx)
    content_types = "http://schemas.openxmlformats.org/package/2006/content-types"
    custom_type = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
    with ZipFile(source) as package:
        types = ET.fromstring(package.read("[Content_Types].xml"))
        parts = [node.attrib["PartName"].lstrip("/") for node in types.findall(f"{{{content_types}}}Override")
                 if node.attrib.get("ContentType") == custom_type]
        if not parts:
            return {}
        if len(parts) != 1:
            raise ValueError("Source contains multiple custom-property parts")
        root = ET.fromstring(package.read(parts[0]))
    preferences = {}
    for prop in root:
        name = prop.attrib.get("name", "")
        if name.upper().startswith("ZOTERO_PREF_"):
            if name.upper() in preferences:
                raise ValueError(f"Duplicate Zotero custom property: {name}")
            children = list(prop)
            if len(children) != 1 or children[0].tag.rsplit("}", 1)[-1] not in ("lpwstr", "lpstr", "bstr"):
                raise ValueError(f"Zotero preference is not a single string property: {name}")
            preferences[name.upper()] = children[0].text or ""
    return preferences


def _copy_zotero_preferences(doc: Any, source_docx: Path) -> dict:
    source = read_zotero_preferences(source_docx)
    properties = doc.CustomDocumentProperties
    existing = {}
    for number in range(1, properties.Count + 1):
        prop = properties.Item(number)
        name = str(prop.Name)
        if name.upper().startswith("ZOTERO_PREF_"):
            existing[name] = str(prop.Value)
    if not source:
        return {"policy": "source_has_no_preferences_preserve_template", "source_chunks": 0,
                "template_chunks": len(existing), "changed": False}
    # Replace the entire source-owned chunk sequence, including removal of stale
    # tail chunks from a longer template preference value. Other properties stay.
    for name in existing:
        properties.Item(name).Delete()
    for name, value in source.items():
        # Office DocumentProperties is often exposed as a late-bound IDispatch
        # collection without named-argument metadata, unlike Word's own methods.
        properties.Add(name, False, 4, value)
    return {"policy": "source_zotero_preferences_override_template_only", "source_chunks": len(source),
            "template_chunks": len(existing), "changed": source != existing,
            "citation_ids_rewritten": False}


def _main_document(path: Path):
    from .package import xml

    with ZipFile(path) as package:
        return xml(package.read("word/document.xml"))


def build_from_template(
    template_path: str | Path, source_docx: str | Path, output_path: str | Path,
    *, slot: str = "{{PAPERFLOW_BODY}}", style_map: Mapping[str | int, str | int] | None = None,
    app: Any = None, preserve_section_stories: bool = False,
) -> dict:
    """Insert a source document into one explicit slot in a template copy.

    This does not guess which template paragraphs are examples. Source section
    breaks and styles may import; return before/after snapshots so an upstream
    template contract can reject changes in page geometry or recurring content.
    """
    template = _source_path(template_path)
    source = _source_path(source_docx)
    output = _output_path(output_path, (template, source))
    if source.suffix.lower() != ".docx" or output.suffix.lower() != ".docx":
        raise ValueError("Source and output must be DOCX files")
    from .layout import validate_bookmark_names

    validate_bookmark_names(_main_document(source), _main_document(template))
    stage = _stage_path(output)
    trim_placeholder = False
    try:
        with _application(app) as word, _opened(word, template) as doc:
            before = snapshot(doc)
            target_styles = {key: _style(doc, value) for key, value in (style_map or {}).items()}
            location = _slot_range(doc, slot)
            insertion_start = int(location.Start)
            terminal = int(doc.Content.End) - 1
            deletion_end = min(int(location.End), terminal)
            tail = None
            if deletion_end == terminal:
                # InsertFile at the very end can replace the last section's
                # header/footer links with those of the imported document.
                # Retain trailing text during insertion, then remove it.
                tail = "PaperFlowTemporaryTail" + uuid.uuid4().hex
                doc.Range(Start=terminal, End=terminal).Text = tail
            if deletion_end > insertion_start:
                doc.Range(Start=insertion_start, End=deletion_end).Text = ""
            before_end = int(doc.Content.End)
            insertion = doc.Range(Start=insertion_start, End=insertion_start)
            insertion.InsertFile(
                FileName=str(source), ConfirmConversions=False, Link=False, Attachment=False,
            )
            insertion_end = insertion_start + int(doc.Content.End) - before_end
            if tail:
                reserved = _find_body_text(doc, tail)
                if int(reserved.Start) < insertion_end:
                    raise WordAutomationError("Word moved the temporary end anchor into the imported content")
                trim_placeholder = (not slot.startswith("bookmark:") and
                                    str(reserved.Paragraphs.Item(1).Range.Text).strip(" \t\r\n") == tail)
                reserved.Text = ""
            remapped = _apply_style_map(doc, insertion_start, insertion_end, target_styles)
            zotero_preferences = _copy_zotero_preferences(doc, source)
            _save_docx(doc, stage)
            after = snapshot(doc)
        trimmed = trim_placeholder and _trim_generated_terminal_paragraph(stage)
        if trimmed:
            after["paragraphs"].pop()
            after["paragraph_count"] -= 1
        from .template import section_story_signature
        stories_preserved = section_story_signature(template) == section_story_signature(stage)
        if preserve_section_stories and not stories_preserved:
            raise WordAutomationError(f"Word changed the template headers, footers, or page numbering while filling {slot!r}")
        _publish(stage, output)
        return {
            "output": str(output), "template": str(template), "source": str(source),
            "slot": slot, "inserted_range": [insertion_start, insertion_end],
            "styles_remapped": remapped, "before": before, "after": after,
            "section_stories_match_template": stories_preserved,
            "terminal_template_placeholder_removed": trimmed,
            "zotero_preferences": zotero_preferences,
            "requires_template_fidelity_audit": True,
            "zotero_refreshed": False,
        }
    finally:
        stage.unlink(missing_ok=True)


def _stories(doc: Any) -> Iterator[Any]:
    # Numeric StoryRanges access avoids locale-dependent constants. A missing
    # story is normal. NextStoryRange reaches later-section headers/footers.
    for kind in range(1, 18):
        try:
            story = doc.StoryRanges.Item(kind)
        except Exception:
            continue
        for _ in range(1000):
            if story is None:
                break
            yield story
            story = story.NextStoryRange
        else:
            raise WordAutomationError("Unexpected cycle or more than 1000 linked Word stories")


def _field_kind(code: str) -> str:
    match = re.match(r"\s*([A-Za-z]+)", code)
    return match.group(1).upper() if match else ""


def _numbered_reference_target(code: str) -> str | None:
    match = _REF_TARGET.match(code)
    if match is None:
        return None
    # Quoted formatting arguments can contain literal backslashes.
    arguments = _FIELD_ARGUMENTS.findall(code[match.end():])
    if any(argument.casefold() in {r"\n", r"\r", r"\w"} for argument in arguments):
        return match[1] or match[2]
    return None


def _paragraph_list_label(doc: Any, name: str) -> str:
    from .package import without_bidi_controls

    if not doc.Bookmarks.Exists(name):
        raise WordAutomationError(f"Word cannot find the numbered reference target {name!r}")
    paragraph = doc.Bookmarks.Item(name).Range.Paragraphs.Item(1).Range
    numbering = paragraph.ListFormat
    label = without_bidi_controls(str(numbering.ListString or "")).strip()
    # WdListType values for simple numbering and outline numbering.
    if numbering.ListType not in {3, 4} or not label:
        raise WordAutomationError(
            f"Numbered reference target {name!r} has no automatic paragraph number. "
            "Restore the target's numbering or change the reference to text or a page number."
        )
    return label


def _update_field(doc: Any, field: Any, kind: str, story_type: int) -> Any:
    if kind == "TOC" and story_type == 1:
        # A generic field update can refresh page numbers without rebuilding
        # the entries. Use the owning Word table's explicit update method.
        for collection_name in ("TablesOfContents", "TablesOfFigures"):
            collection = getattr(doc, collection_name, None)
            if collection is None:
                continue
            for index in range(1, collection.Count + 1):
                table = collection.Item(index)
                location = table.Range
                if location.Start <= field.Code.Start and field.Result.End <= location.End:
                    return table.Update()
    return field.Update()


def refresh_document_fields(doc: Any) -> dict:
    """Refresh deterministic Word fields, deliberately skipping every ADDIN.

    External-data, macro, ASK/FILLIN, and unknown fields are also skipped.
    Zotero refresh belongs to Zotero's integration, never generic F9/Update.
    """
    updated: list[dict] = []
    skipped: list[dict] = []
    errors: list[dict] = []
    numbered_references: list[dict] = []
    # Re-enumerate after TOC regeneration because its fields can be replaced.
    # Numbering must finish before REF fields read its cached result. Regenerate
    # navigation tables afterwards, then resolve all page-dependent fields.
    phases = [
        _SAFE_FIELDS - _PAGE_FIELDS - _TABLE_FIELDS - {"REF"},
        {"REF"}, _TABLE_FIELDS, _PAGE_FIELDS,
    ]
    doc.Repaginate()
    for phase_no, selected in enumerate(phases):
        if phase_no:
            doc.Repaginate()
        for story in _stories(doc):
            indices = range(1, story.Fields.Count + 1)
            # A regenerated table can add or remove nested page fields.
            # Working from the end keeps earlier indices valid in this pass.
            if selected == _TABLE_FIELDS:
                indices = reversed(indices)
            for n in indices:
                field = story.Fields.Item(n)
                code = str(field.Code.Text)
                kind = _field_kind(code)
                record = {"kind": kind, "code": code, "story": int(story.StoryType)}
                if phase_no == 0 and kind not in _SAFE_FIELDS:
                    record["reason"] = "add-in requires its owning application" if kind == "ADDIN" else "not in safe refresh allowlist"
                    skipped.append(record)
                if kind not in selected:
                    continue
                if bool(field.Locked):
                    if phase_no == 0 or kind in selected:
                        skipped.append({**record, "reason": "field is locked"})
                    continue
                try:
                    target = _numbered_reference_target(code) if kind == "REF" else None
                    if target is not None:
                        label = _paragraph_list_label(doc, target)
                        numbered_references.append({**record, "bookmark": target, "list_label": label})
                    success = _update_field(doc, field, kind, record["story"])
                    if success is False or success == 0:
                        errors.append({**record, "error": "Word returned False from field refresh"})
                    else:
                        updated.append(record)
                except Exception as exc:
                    errors.append({**record, "error": str(exc)})
    doc.Repaginate()
    return {"updated": updated, "skipped": skipped, "errors": errors,
            "numbered_references": numbered_references, "zotero_refreshed": False}


def refresh_fields(
    input_path: str | Path, output_path: str | Path, *, app: Any = None,
) -> dict:
    """Create a DOCX copy with ordinary fields refreshed; input stays unchanged."""
    source = _source_path(input_path)
    output = _output_path(output_path, (source,))
    if output.suffix.lower() != ".docx":
        raise ValueError("Field refresh output must have .docx extension")
    stage = _stage_path(output)
    try:
        with _application(app) as word, _opened(word, source) as doc:
            report = refresh_document_fields(doc)
            if report["errors"]:
                raise WordAutomationError(f"Word field refresh failed: {report['errors']}")
            _save_docx(doc, stage)
            report["snapshot"] = snapshot(doc)
        _publish(stage, output)
        return {"output": str(output), **report}
    finally:
        stage.unlink(missing_ok=True)


def numbered_template_headings(template_path: str | Path, styles: Mapping[int, str], *, app: Any = None) -> set[int]:
    """Keep only template heading styles Word recognizes as linked to an automatic list."""
    if not styles:
        return set()
    source = _source_path(template_path)
    levels = set()
    with _application(app) as word, _opened(word, source) as doc:
        for level, name in styles.items():
            try:
                style = _style(doc, name)
            except ValueError:
                continue
            if style.ListLevelNumber > 0 and style.ListTemplate is not None:
                levels.add(level)
    return levels


def heading_list_labels(input_path: str | Path, names, *, app: Any = None) -> dict[str, str]:
    """Check that Word resolves each heading bookmark to an automatic list number."""
    if not names:
        return {}
    source = _source_path(input_path)
    labels = {}
    with _application(app) as word, _opened(word, source) as doc:
        for name in names:
            labels[name] = _paragraph_list_label(doc, name)
    return labels


def export_pdf(input_path: str | Path, output_path: str | Path, *, app: Any = None) -> dict:
    """Export using Word's renderer without saving changes to the input DOCX."""
    source = _source_path(input_path)
    output = _output_path(output_path, (source,))
    if output.suffix.lower() != ".pdf":
        raise ValueError("PDF export output must have .pdf extension")
    stage = _stage_path(output)
    try:
        with _application(app) as word, _opened(word, source) as doc:
            doc.Repaginate()
            doc.ExportAsFixedFormat(
                OutputFileName=str(stage), ExportFormat=PDF_FORMAT,
                OpenAfterExport=False, OptimizeFor=0, CreateBookmarks=1,
                DocStructureTags=True, BitmapMissingFonts=True,
            )
            pages = int(doc.ComputeStatistics(2))
        _publish(stage, output)
        return {"output": str(output), "renderer": "Microsoft Word", "pages": pages}
    finally:
        stage.unlink(missing_ok=True)


def _read(obj: Any, name: str, default: Any = None) -> Any:
    try:
        value = getattr(obj, name)
        return value if isinstance(value, (str, int, float, bool, type(None))) else str(value)
    except Exception:
        return default


def _snapshot_document(doc: Any) -> dict:
    sections = []
    for n in range(1, doc.Sections.Count + 1):
        section = doc.Sections.Item(n)
        setup = section.PageSetup
        sections.append({"index": n, **{name: _read(setup, name) for name in (
            "PageWidth", "PageHeight", "TopMargin", "BottomMargin", "LeftMargin", "RightMargin",
            "HeaderDistance", "FooterDistance", "Orientation", "SectionStart", "Gutter",
            "DifferentFirstPageHeaderFooter", "OddAndEvenPagesHeaderFooter",
        )}})
    paragraphs = []
    for n in range(1, doc.Paragraphs.Count + 1):
        para = doc.Paragraphs.Item(n)
        paragraphs.append({
            "index": n, "start": int(para.Range.Start), "end": int(para.Range.End),
            "text": str(para.Range.Text)[:1000], "style": _style_name(para.Style),
            "outline_level": _read(para, "OutlineLevel"),
            "font": {name: _read(para.Range.Font, name) for name in ("Name", "NameFarEast", "Size", "Bold", "Italic")},
        })
    fields = []
    inline_shapes = []
    for story in _stories(doc):
        for n in range(1, story.Fields.Count + 1):
            field = story.Fields.Item(n)
            code = str(field.Code.Text)
            fields.append({
                "story": int(story.StoryType), "code": code, "kind": _field_kind(code),
                "result": str(field.Result.Text), "locked": bool(field.Locked),
            })
        for n in range(1, story.InlineShapes.Count + 1):
            shape = story.InlineShapes.Item(n)
            item = {"story": int(story.StoryType), "type": int(shape.Type),
                    "width": float(shape.Width), "height": float(shape.Height),
                    "alt": _read(shape, "AlternativeText", "")}
            if int(shape.Type) in (1, 2):
                item["progid"] = _read(shape.OLEFormat, "ProgID")
            inline_shapes.append(item)
    shapes = []
    for n in range(1, doc.Shapes.Count + 1):
        shape = doc.Shapes.Item(n)
        item = {"name": _read(shape, "Name"), "type": int(shape.Type)}
        if int(shape.Type) in (7, 10):
            item["progid"] = _read(shape.OLEFormat, "ProgID")
        shapes.append(item)
    return {"sections": sections, "paragraphs": paragraphs, "fields": fields,
            "inline_shapes": inline_shapes, "floating_shapes": shapes,
            "section_count": len(sections), "paragraph_count": len(paragraphs),
            "field_count": len(fields), "inline_shape_count": len(inline_shapes),
            "word_version": _read(doc.Application, "Version")}


def snapshot(doc_or_path: Any, *, app: Any = None) -> dict:
    """Read a live owned Document or open a path read-only in a private instance."""
    if isinstance(doc_or_path, (str, Path)):
        source = _source_path(doc_or_path)
        with _application(app) as word, _opened(word, source) as doc:
            return _snapshot_document(doc)
    return _snapshot_document(doc_or_path)
