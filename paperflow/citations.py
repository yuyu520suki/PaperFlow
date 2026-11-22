"""Native Zotero refresh and a separate static submission copy.

The refresh operation invokes only the already installed Zotero Word add-in.
It does not install styles, change macro security, or write to the library.
The submission operation unlinks only Zotero-owned fields in OOXML, preserving
their cached results and all unrelated fields and embedded object parts.
"""
from __future__ import annotations

from collections import Counter
import configparser
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import time
from typing import Any

from lxml import etree as ET

from . import word
from .package import q, read_package, without_bidi_controls, write_package, xml
from .zotero import (ITEM_URI, ZoteroClient, normalize_style_id, preference_style_id, preferences_for_style,
                     read_document_preferences, set_docx_preferences)


class CitationError(RuntimeError):
    """The citation state cannot be safely refreshed or frozen."""


class ZoteroRefreshTimeout(CitationError):
    """The add-in's document results did not satisfy verification in time."""


_OWNER = re.compile(r"^\s*ADDIN\s+ZOTERO_(ITEM|BIBL)\b", re.I)
_PENDING = "PAPERFLOW_PENDING"


def _owner(code: str) -> str | None:
    match = _OWNER.match(code)
    return match.group(1).upper() if match else None


def _payload(code: str, owner: str) -> dict:
    tail = code.split("CSL_CITATION", 1)[1] if owner == "ITEM" and "CSL_CITATION" in code else _OWNER.sub("", code, count=1)
    try:
        result, _ = json.JSONDecoder().raw_decode(tail.lstrip())
    except (ValueError, IndexError) as exc:
        raise CitationError(f"Malformed Zotero {owner} field JSON") from exc
    if not isinstance(result, dict):
        raise CitationError(f"Zotero {owner} field JSON is not an object")
    return result


def _uri_identity(uri: str) -> str:
    return uri.replace("https://www.zotero.org/", "http://zotero.org/").replace("https://zotero.org/", "http://zotero.org/")


def _valid_uris(item: dict) -> list[str]:
    uris = item.get("uris", [])
    if not isinstance(uris, list):
        return []
    return [uri for uri in uris if isinstance(uri, str) and ITEM_URI.fullmatch(uri)]


def _scan_fields(root: Any) -> list[dict]:
    """Collect field boundaries and cached text, including split instructions."""
    fields, stack = [], []
    for element in root.iter():
        if element.tag == q("w:fldSimple"):
            fields.append({"code": element.get(q("w:instr"), ""),
                           "result": "".join(node.text or "" for node in element.iter(q("w:t"))), "simple": element})
        elif element.tag == q("w:fldChar"):
            kind = element.get(q("w:fldCharType"))
            if kind == "begin":
                stack.append({"instruction": [], "code_nodes": [], "result_nodes": [],
                              "separated": False, "begin": element, "separator": None})
            elif kind == "separate":
                if not stack:
                    raise CitationError("Field separator without a beginning")
                stack[-1]["separated"] = True
                stack[-1]["separator"] = element
            elif kind == "end":
                if not stack:
                    raise CitationError("Field end without a beginning")
                field = stack.pop()
                field["code"] = "".join(field.pop("instruction"))
                field["result"] = "".join(field.pop("result_nodes"))
                field["end"] = element
                fields.append(field)
        elif element.tag == q("w:instrText") and stack and not stack[-1]["separated"]:
            stack[-1]["instruction"].append(element.text or "")
            stack[-1]["code_nodes"].append(element)
        elif element.tag in {q("w:t"), q("w:tab"), q("w:br")}:
            text = element.text or "" if element.tag == q("w:t") else ("\t" if element.tag == q("w:tab") else "\n")
            for active in stack:
                if active["separated"]:
                    active["result_nodes"].append(text)
    if stack:
        raise CitationError("Unclosed Word field")
    return fields


def _package_state(parts: dict[str, bytes]) -> dict:
    fields, ordinary, ole = [], [], []
    for name, data in parts.items():
        if name.startswith("word/") and name.endswith(".xml"):
            root = xml(data)
            for field in _scan_fields(root):
                code = field["code"]
                owner = _owner(code)
                if owner:
                    fields.append({"part": name, "code": code, "result": field["result"], "owner": owner})
                else:
                    ordinary.append((name, code))
            ole.extend((name, element.get("ProgID", ""), element.get(q("r:id"), ""))
                       for element in root.iter(q("o:OLEObject")))
    embeddings = {name: hashlib.sha256(data).hexdigest() for name, data in parts.items()
                  if name.startswith("word/embeddings/")}
    return {"fields": fields, "ordinary_fields": ordinary, "ole_objects": ole, "embeddings": embeddings}


def _summary(state: dict) -> dict:
    return {"citation_fields": sum(f["owner"] == "ITEM" for f in state["fields"]),
            "bibliography_fields": sum(f["owner"] == "BIBL" for f in state["fields"]),
            "ordinary_fields": len(state["ordinary_fields"]), "ole_objects": len(state["ole_objects"]),
            "embedding_parts": len(state["embeddings"])}


def verify_zotero_fields_preserved(before: str | Path, after: str | Path) -> dict:
    """Check that a Word layout update preserved the refreshed Zotero fields."""
    fields_before = _package_state(read_package(before))["fields"]
    fields_after = _package_state(read_package(after))["fields"]
    if fields_before != fields_after:
        raise CitationError("Word layout refresh changed the Zotero citations or bibliography")
    before_style = preference_style_id(read_document_preferences(before))
    after_style = preference_style_id(read_document_preferences(after))
    if before_style != after_style:
        raise CitationError("Word layout refresh changed the document's Zotero citation style")
    return {"citation_fields": sum(field["owner"] == "ITEM" for field in fields_after),
            "bibliography_fields": sum(field["owner"] == "BIBL" for field in fields_after),
            "saved_style_id": after_style,
            "zotero_fields_preserved": True}


def _guards(before: dict, after: dict) -> dict:
    ordinary_before = Counter(re.sub(r"\s+", " ", code.strip()) for _, code in before["ordinary_fields"])
    ordinary_after = Counter(re.sub(r"\s+", " ", code.strip()) for _, code in after["ordinary_fields"])
    ordinary_ok = ordinary_before == ordinary_after
    native_ok = (Counter(progid for _, progid, _ in before["ole_objects"]) ==
                 Counter(progid for _, progid, _ in after["ole_objects"]) and
                 Counter(before["embeddings"].values()) == Counter(after["embeddings"].values()))
    if not ordinary_ok or not native_ok:
        raise CitationError("Citation operation changed ordinary field instructions or embedded OLE payloads")
    return {"ordinary_field_instructions_preserved": ordinary_ok,
            "ole_object_types_and_payloads_preserved": native_ok}


def _paths(source: str | Path, output: str | Path) -> tuple[Path, Path]:
    source = word._source_path(source)
    output = word._output_path(output, (source,))
    if source.suffix.lower() != ".docx" or output.suffix.lower() != ".docx":
        raise ValueError("Citation operations require DOCX source and output files")
    return source, output


def prepare_citation_style(source: str | Path, output: str | Path, style_id: str):
    """Prepare an independent document copy for an actual CSL style change."""
    source, output = _paths(source, output)
    if not isinstance(style_id, str):
        raise ValueError("style_id must be a full Zotero style identifier")
    requested = normalize_style_id(style_id)
    preferences = read_document_preferences(source)
    previous = preference_style_id(preferences)
    report = {"id": requested, "previous_style_id": previous, "preferences_changed": previous != requested,
              "library_modified": False}
    if requested == previous:
        return source, report
    for field in _package_state(read_package(source))["fields"]:
        owner = field["owner"]
        payload = _payload(field["code"], owner)
        if owner == "BIBL":
            if payload.get("custom"):
                raise CitationError("The Zotero bibliography contains manually edited entries. "
                                    "Review them in Word before automatically changing the CSL style.")
            continue
        properties = payload.get("properties", {})
        plain = properties.get("plainCitation")
        if (properties.get("dontUpdate") or properties.get("custom")
                or (isinstance(plain, str) and plain
                    and _normalized_text(plain) != _normalized_text(field["result"]))):
            raise CitationError("A Zotero citation has custom or manually edited text. "
                                "Review it in Word before automatically changing the CSL style.")
    stage = word._stage_path(output)
    try:
        shutil.copyfile(source, stage)
        set_docx_preferences(stage, preferences_for_style(preferences, requested))
        word._publish(stage, output)
    finally:
        stage.unlink(missing_ok=True)
    return output, report


def _copy_without_refresh(source: Path, output: Path, state: dict, reason: str) -> dict:
    stage = word._stage_path(output)
    try:
        shutil.copyfile(source, stage)
        word._publish(stage, output)
    finally:
        stage.unlink(missing_ok=True)
    return {"output": str(output), "status": "not_needed", "reason": reason,
            "macro_invoked": False, "zotero_native_refresh_verified": False,
            "before": _summary(state), "after": _summary(state),
            "source_preserved": True}


def _expectations(fields: list[dict]) -> dict[str, str]:
    """Read current metadata for exactly the citation/uncited records in this DOCX."""
    expected: dict[str, str] = {}
    clients = {}
    for field in fields:
        payload = _payload(field["code"], field["owner"])
        if field["owner"] == "ITEM":
            items = payload.get("citationItems", [])
            if not items:
                raise CitationError("A citation contains no items")
            uris = []
            for item in items:
                valid = _valid_uris(item)
                if not valid:
                    raise CitationError("Citation item lacks a genuine Zotero library URI")
                uris.append(valid[0])
        else:
            uris = [entry[0] for entry in payload.get("uncited", []) if isinstance(entry, list) and entry]
        for uri in uris:
            if not isinstance(uri, str) or not ITEM_URI.fullmatch(uri):
                raise CitationError("Bibliography uncited item lacks a genuine library URI")
            identity = _uri_identity(uri)
            if identity in expected:
                continue
            key = uri.rsplit("/", 1)[1]
            group = re.search(r"/groups/([0-9]+)/", uri)
            scope = ("groups", int(group.group(1))) if group else ("users", 0)
            if scope not in clients:
                clients[scope] = ZoteroClient(library_type=scope[0], library_id=scope[1])
            record = clients[scope].get_item(key)
            if _uri_identity(record["uri"]) != identity:
                raise CitationError(f"Requested citation does not resolve to the same local library record: {uri}")
            expected[identity] = str(record["item_data"]["title"])
    return expected


def _discover_styles_dir() -> Path:
    """Resolve a configured/default Zotero data directory without changing it."""
    if override := os.environ.get("PAPERFLOW_ZOTERO_STYLES_DIR"):
        return Path(override).expanduser().resolve(strict=True)
    appdata = os.environ.get("APPDATA")
    if appdata:
        zotero_root = Path(appdata) / "Zotero" / "Zotero"
        ini = zotero_root / "profiles.ini"
        if ini.is_file():
            config = configparser.ConfigParser(interpolation=None)
            config.read(ini, encoding="utf-8")
            candidates = [config[section] for section in config.sections()
                          if section.startswith("Profile") and config[section].get("Default") == "1"]
            if not candidates:
                candidates = [config[section] for section in config.sections() if section.startswith("Profile")]
            if len(candidates) != 1:
                raise CitationError("Zotero profile is ambiguous; pass styles_dir explicitly")
            selected = candidates[0]
            profile = Path(selected["Path"])
            if selected.get("IsRelative", "1") == "1":
                profile = zotero_root / profile
            prefs = profile / "prefs.js"
            if prefs.is_file():
                text = prefs.read_text(encoding="utf-8")
                data_match = re.search(r'user_pref\("extensions\.zotero\.dataDir",\s*("(?:\\.|[^"\\])*")\s*\);', text)
                legacy = re.search(r'user_pref\("extensions\.zotero\.useDataDir",\s*false\s*\);', text)
                if legacy:
                    return (profile / "zotero" / "styles").resolve(strict=True)
                if data_match:
                    return (Path(json.loads(data_match.group(1))) / "styles").resolve(strict=True)
    standard = Path.home() / "Zotero"
    if (standard / "zotero.sqlite").is_file() and (standard / "styles").is_dir():
        return (standard / "styles").resolve()
    raise CitationError("Cannot resolve installed Zotero styles; pass styles_dir explicitly")


def _style_preflight(source: Path, styles_dir: str | Path | None) -> dict:
    """Refuse a missing CSL or independent parent before invoking Zotero."""
    preferences = read_document_preferences(source)
    if preferences is None:
        raise CitationError("Document has no Zotero preferences; set its citation style before refresh")
    wanted = preference_style_id(preferences)
    if wanted is None:
        raise CitationError("Document Zotero preferences do not select a CSL style")
    directory = Path(styles_dir).expanduser().resolve(strict=True) if styles_dir else _discover_styles_dir()
    if not directory.is_dir():
        raise CitationError("Zotero styles_dir is not a directory")
    csl_ns = {"c": "http://purl.org/net/xbiblio/csl"}
    index = {}
    for folder in (directory, directory / "hidden"):
        if not folder.is_dir():
            continue
        for path in folder.glob("*.csl"):
            if path.stat().st_size > 2 * 1024 * 1024:
                continue
            try:
                root = xml(path.read_bytes())
            except ET.XMLSyntaxError:
                continue
            identity = root.findtext("c:info/c:id", namespaces=csl_ns)
            if identity:
                parents = [link.get("href") for link in root.findall("c:info/c:link", csl_ns)
                           if link.get("rel") == "independent-parent" and link.get("href")]
                index[normalize_style_id(identity)] = {"parents": [normalize_style_id(p) for p in parents],
                                                       "style_class": root.get("class")}
    checked, active = [], set()

    def require(identity: str) -> None:
        if identity in active:
            raise CitationError("Installed CSL independent-parent cycle")
        if identity not in index:
            raise CitationError(f"CSL style is not installed; refresh will not download it: {identity}")
        if identity in checked:
            return
        active.add(identity)
        for parent in index[identity]["parents"]:
            require(parent)
        active.remove(identity)
        checked.append(identity)

    require(wanted)
    style_class = next((index[identity]["style_class"] for identity in reversed(checked)
                        if index[identity]["style_class"]), None)
    note_type = preferences.find("./prefs/pref[@name='noteType']")
    uses_notes = note_type is not None and note_type.get("value") in {"1", "2"}
    if style_class in {"note", "in-text"} and (style_class == "note") != uses_notes:
        raise CitationError("This CSL style requires changing where citations are stored. "
                            "Switch between note and in-text styles in Word's Zotero Document Preferences.")
    return {"styles_dir": str(directory), "document_style": wanted, "installed_style_chain": checked,
            "style_class": style_class,
            "styles_installed_by_operation": False}


def _live_fields(doc: Any) -> list[dict]:
    fields = []
    for story in word._stories(doc):
        for number in range(1, story.Fields.Count + 1):
            field = story.Fields.Item(number)
            code = str(field.Code.Text)
            if owner := _owner(code):
                fields.append({"owner": owner, "code": code, "result": str(field.Result.Text)})
    return fields


def _normalized_text(value: str) -> str:
    return re.sub(r"[\s\u200b\u200c\u200d\ufeff]+", "", without_bidi_controls(value))


def _verification_errors(fields: list[dict], expected: dict[str, str], counts: dict) -> list[str]:
    errors = []
    if sum(f["owner"] == "ITEM" for f in fields) != counts["citation_fields"]:
        errors.append("Citation field count changed during refresh")
    if sum(f["owner"] == "BIBL" for f in fields) != counts["bibliography_fields"]:
        errors.append("Bibliography field count changed during refresh")
    seen = set()
    for field in fields:
        result = field["result"]
        if not _normalized_text(result) or _PENDING in result:
            errors.append(f"{field['owner']} field still has an empty or pending visible result")
        if field["owner"] == "BIBL":
            continue
        try:
            payload = _payload(field["code"], "ITEM")
            properties = payload.get("properties", {})
            for key in ("formattedCitation", "plainCitation"):
                if not isinstance(properties.get(key), str) or not properties[key].strip() or _PENDING in properties[key]:
                    errors.append(f"Citation properties.{key} is not rendered")
            plain = properties.get("plainCitation", "")
            if isinstance(plain, str) and _normalized_text(plain) != _normalized_text(result):
                errors.append("Visible citation differs from properties.plainCitation")
            items = payload.get("citationItems", [])
            if not items:
                errors.append("Citation has no items after refresh")
            for item in items:
                uris = [_uri_identity(uri) for uri in _valid_uris(item)]
                matched = next((uri for uri in uris if uri in expected), None)
                if matched is None:
                    errors.append("Refreshed citation lost its real library identity")
                    continue
                seen.add(matched)
                if item.get("itemData", {}).get("title") != expected[matched]:
                    errors.append(f"Citation metadata has not reached current library title: {matched}")
        except (CitationError, TypeError, AttributeError) as exc:
            errors.append(str(exc))
    # Bibliography-only uncited records have no citation-item metadata to inspect.
    if counts["citation_fields"] and not seen:
        errors.append("No current library records could be verified")
    return errors


def _pump_messages() -> None:
    import pythoncom
    pythoncom.PumpWaitingMessages()


def refresh_zotero(
    source: str | Path, output: str | Path, *, app: Any = None, timeout: float = 90,
    styles_dir: str | Path | None = None,
) -> dict:
    """Refresh a new master copy through the preinstalled Zotero Word add-in.

    ``timeout`` bounds result verification after macro dispatch. A synchronous
    COM call itself cannot be safely interrupted in the owning apartment; a CLI
    supervisor may impose a separate process deadline if that is required.
    """
    if not isinstance(timeout, (int, float)) or isinstance(timeout, bool) or timeout <= 0:
        raise ValueError("timeout must be a positive number of seconds")
    source, output = _paths(source, output)
    before = _package_state(read_package(source))
    if not before["fields"]:
        return _copy_without_refresh(source, output, before, "no_zotero_fields")
    expected = _expectations(before["fields"])
    if not expected:
        return _copy_without_refresh(source, output, before, "bibliography_has_no_citation_or_uncited_items")
    style_preflight = _style_preflight(source, styles_dir)
    counts = _summary(before)
    stage = word._stage_path(output)
    started = time.monotonic()
    last_errors = ["No refreshed state observed"]
    try:
        with word._application(app) as active, word._opened(active, source) as doc:
            word._save_docx(doc, stage)
            doc.Activate()
            try:
                active.Run("Zotero.ZoteroRefresh")
            except Exception as exc:
                raise CitationError(f"Installed Zotero Word refresh macro could not be invoked: {exc}") from exc
            deadline = time.monotonic() + float(timeout)
            previous_signature = None
            while time.monotonic() < deadline:
                _pump_messages()
                try:
                    fields = _live_fields(doc)
                    last_errors = _verification_errors(fields, expected, counts)
                    signature = json.dumps(fields, sort_keys=True, ensure_ascii=False)
                    if not last_errors and signature == previous_signature:
                        break
                    previous_signature = signature if not last_errors else None
                except Exception as exc:
                    last_errors = [f"Word results not yet readable: {exc}"]
                    previous_signature = None
                time.sleep(min(0.25, max(0, deadline - time.monotonic())))
            else:
                raise ZoteroRefreshTimeout(f"Zotero refresh verification timed out: {last_errors}")
            doc.Save()
        after = _package_state(read_package(stage))
        guard = _guards(before, after)
        saved_errors = _verification_errors(after["fields"], expected, counts)
        if saved_errors:
            raise CitationError(f"Saved Zotero fields failed verification: {saved_errors}")
        saved_style = preference_style_id(read_document_preferences(stage))
        if saved_style != style_preflight["document_style"]:
            raise CitationError("Zotero saved a different CSL style than the one selected for refresh")
        word._publish(stage, output)
        return {"output": str(output), "status": "refreshed", "macro_invoked": True,
                "zotero_native_refresh_verified": True, "current_library_records_checked": len(expected),
                "bibliography_present": bool(counts["bibliography_fields"]),
                "before": counts, "after": _summary(after), **guard,
                "style_preflight": style_preflight, "saved_style_id": saved_style,
                "elapsed_seconds": time.monotonic() - started, "source_preserved": True}
    finally:
        stage.unlink(missing_ok=True)


def create_submission_copy(source: str | Path, output: str | Path) -> dict:
    """Freeze Zotero display text in a new copy, retaining native Word features.

    No Word process is needed. Unlike flattening the document, this removes only
    Zotero-owned field instructions and delimiters. Nested ordinary fields and
    the entire cached result subtree remain editable in their existing form.
    """
    source, output = _paths(source, output)
    parts = read_package(source)
    before = _package_state(parts)
    for field in before["fields"]:
        if _PENDING in field["result"] or (field["owner"] == "ITEM" and not field["result"].strip()):
            raise CitationError("Refresh pending/empty Zotero citations before making a submission copy")
    changed = []
    for name, data in list(parts.items()):
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        root = xml(data)
        owned = [field for field in _scan_fields(root) if _owner(field["code"])]
        if not owned:
            continue
        for field in owned:
            if "simple" in field:
                element = field["simple"]
                parent = element.getparent()
                index = parent.index(element)
                for child in list(element):
                    element.remove(child)
                    parent.insert(index, child)
                    index += 1
                parent.remove(element)
            else:
                for element in [field["begin"], *field["code_nodes"], field["separator"], field["end"]]:
                    if element is not None and element.getparent() is not None:
                        element.getparent().remove(element)
        parts[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True, standalone=True)
        changed.append(name)
    after = _package_state(parts)
    if after["fields"]:
        raise CitationError("Zotero-owned fields remain after unlinking")
    guard = _guards(before, after)
    stage = word._stage_path(output)
    try:
        write_package(parts, stage)
        word._publish(stage, output)
    finally:
        stage.unlink(missing_ok=True)
    return {"output": str(output), "status": "submission_copy" if changed else "not_needed",
            "before": _summary(before), "after": _summary(after), "changed_parts": changed,
            "only_zotero_fields_unlinked": True, **guard, "source_preserved": True,
            "dynamic_master": str(source)}
