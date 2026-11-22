"""Read Zotero metadata and construct Word integration fields.

This module never writes to a Zotero library. A field passing structural checks
is only a candidate dynamic citation until the real Word add-in has refreshed
it and saved the result. Generic Word Fields.Update() is not that operation.
"""

from __future__ import annotations

import copy
import json
import os
import re
import secrets
import string
import tempfile
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any, Iterable

CSL_SCHEMA = "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
PREF_PREFIX = "ZOTERO_PREF_"
CUSTOM_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
CUSTOM_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
CUSTOM_TYPE = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
CUSTOM_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
ITEM_KEY = re.compile(r"^[A-Z0-9]{8}$")
ITEM_URI = re.compile(
    r"^https?://(?:www\.)?zotero\.org/(?:users/(?:[1-9][0-9]*|local/[A-Za-z0-9]+)|groups/[1-9][0-9]*)/items/([A-Z0-9]{8})$"
)


class ZoteroError(RuntimeError):
    """A local API or metadata contract failed."""


def _random_id(length: int = 12) -> str:
    return "".join(secrets.choice(string.ascii_letters + string.digits) for _ in range(length))


class ZoteroClient:
    """A read-only client for Zotero Desktop's loopback API.

    ``users/0`` denotes the current desktop user for API requests only. It is
    deliberately forbidden as an identity inside generated citation URIs.
    """

    def __init__(
        self,
        base_url: str = "http://127.0.0.1:23119",
        *,
        library_type: str = "users",
        library_id: int = 0,
        timeout: float = 15.0,
    ) -> None:
        parsed = urllib.parse.urlsplit(base_url)
        if parsed.scheme != "http" or parsed.hostname not in {"127.0.0.1", "localhost", "::1"}:
            raise ValueError("Zotero Desktop access requires a loopback HTTP URL")
        if parsed.username or parsed.password or parsed.query or parsed.fragment or parsed.path not in {"", "/"}:
            raise ValueError("base_url must contain only the loopback origin")
        if library_type not in {"users", "groups"} or not isinstance(library_id, int) or library_id < 0:
            raise ValueError("Expected users/groups and a nonnegative numeric library ID")
        if library_type == "groups" and library_id == 0:
            raise ValueError("Group library ID must be positive")
        self.base_url = base_url.rstrip("/")
        self.scope = f"/api/{library_type}/{library_id}"
        self.timeout = timeout
        # Loopback requests must not be sent through a configured web proxy.
        self._opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))

    def _json(self, path: str, **query: Any) -> Any:
        url = self.base_url + path
        if query:
            url += "?" + urllib.parse.urlencode(query)
        request = urllib.request.Request(url, headers={"Zotero-API-Version": "3", "Accept": "application/json"})
        try:
            with self._opener.open(request, timeout=self.timeout) as response:
                return json.loads(response.read().decode("utf-8"))
        except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, UnicodeDecodeError) as exc:
            raise ZoteroError(f"Zotero local API read failed: {path}: {exc}") from exc

    def status(self) -> dict[str, Any]:
        """Return read-only desktop API information."""
        return self._json("/api/")

    def search(self, query: str, *, limit: int = 20, mode: str = "titleCreatorYear") -> list[dict[str, Any]]:
        """Search parent-record metadata, without reading attachments/full text."""
        if not 1 <= limit <= 100:
            raise ValueError("limit must be between 1 and 100")
        if mode not in {"titleCreatorYear", "everything"}:
            raise ValueError("Unknown Zotero search mode")
        return self._json(self.scope + "/items/top", q=query, limit=limit, qmode=mode)

    def get_item(self, item_key: str) -> dict[str, Any]:
        """Get CSL JSON and a resolvable library identity for one explicit item."""
        if not ITEM_KEY.fullmatch(item_key):
            raise ValueError("Expected an eight-character Zotero item key, not a BibTeX citekey")
        path = self.scope + "/items/" + item_key
        record = self._json(path)
        if record.get("data", {}).get("itemType") in {"attachment", "note", "annotation"}:
            raise ZoteroError("Citations require a parent bibliographic record")
        csl = self._json(path, format="csljson")
        if isinstance(csl, list):
            if len(csl) != 1:
                raise ZoteroError("Expected one CSL record for the requested item")
            csl = csl[0]
        if not isinstance(csl, dict):
            raise ZoteroError("Invalid CSL JSON returned by Zotero")
        uri = str(csl.get("id", ""))
        match = ITEM_URI.fullmatch(uri)
        if not match or match.group(1) != item_key:
            # Alternate URLs are human-facing and use www/https. Canonical
            # CSL item IDs normally already contain Zotero's http identity.
            candidate = record.get("links", {}).get("alternate", {}).get("href", "")
            match = ITEM_URI.fullmatch(candidate)
            if not match or match.group(1) != item_key:
                raise ZoteroError("No real Zotero library URI was returned; refusing users/0 identity")
            uri = candidate.replace("https://www.zotero.org/", "http://zotero.org/").replace("https://zotero.org/", "http://zotero.org/")
        item_data = copy.deepcopy(csl)
        item_data["id"] = uri
        return {"key": item_key, "uri": uri, "item_data": item_data, "version": record.get("version"), "library": record.get("library", {})}


def make_citation_field(
    items: Iterable[dict[str, Any]],
    *,
    citation_id: str | None = None,
    note_index: int = 0,
) -> str:
    """Create a complete Word ADDIN code containing Zotero CSL_CITATION data.

    Items are records returned by ``get_item``. Optional per-citation keys:
    locator, label, prefix, suffix, suppress-author, author-only. Unknown keys
    are not copied. There is no fabricated bibliography or formatting fallback.
    """
    if not isinstance(note_index, int) or note_index < 0:
        raise ValueError("note_index must be a nonnegative integer")
    citation_items: list[dict[str, Any]] = []
    for item in items:
        uri = item.get("uri", "")
        if not ITEM_URI.fullmatch(uri):
            raise ValueError("Each citation needs a genuine library URI; users/0 is not valid")
        data = item.get("item_data")
        if not isinstance(data, dict) or not data.get("type") or not data.get("title"):
            raise ValueError("Each citation needs CSL bibliographic data with type and title")
        entry: dict[str, Any] = {"id": uri, "uris": [uri], "itemData": copy.deepcopy(data)}
        entry["itemData"]["id"] = uri
        for key in ("locator", "label", "prefix", "suffix", "suppress-author", "author-only"):
            if key in item:
                entry[key] = item[key]
        citation_items.append(entry)
    if not citation_items:
        raise ValueError("At least one citation item is required")
    payload = {
        "citationID": citation_id or _random_id(),
        "properties": {"noteIndex": note_index},
        "citationItems": citation_items,
        "schema": CSL_SCHEMA,
    }
    return " ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + " "


def make_bibliography_field(*, uncited_uris: Iterable[str] = ()) -> str:
    """Create a Word Zotero bibliography code, to be rendered by the add-in."""
    uris = list(uncited_uris)
    if any(not ITEM_URI.fullmatch(uri) for uri in uris):
        raise ValueError("Uncited items need valid Zotero library URIs")
    payload = {"uncited": [[uri] for uri in uris], "omitted": [], "custom": []}
    return " ADDIN ZOTERO_BIBL " + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + " CSL_BIBLIOGRAPHY "


def document_preferences(
    style_id: str = "http://www.zotero.org/styles/china-national-standard-gb-t-7714-2015-numeric",
    *,
    locale: str = "zh-CN",
    zotero_version: str = "10.0.7",
    session_id: str | None = None,
    note_type: int = 0,
) -> str:
    """Create Zotero integration data v3; the add-in must validate the style."""
    if not style_id.startswith(("http://www.zotero.org/styles/", "https://www.zotero.org/styles/")):
        raise ValueError("style_id must be a full Zotero style identifier")
    if note_type not in {0, 1, 2}:
        raise ValueError("note_type must be 0 (text), 1 (footnote), or 2 (endnote)")
    root = ET.Element("data", {"data-version": "3", "zotero-version": zotero_version})
    ET.SubElement(root, "session", {"id": session_id or _random_id()})
    ET.SubElement(root, "style", {"id": style_id, "locale": locale, "hasBibliography": "1", "bibliographyStyleHasBeenSet": "0"})
    prefs = ET.SubElement(root, "prefs")
    for name, value in (("fieldType", "Field"), ("noteType", str(note_type)), ("automaticJournalAbbreviations", "true")):
        ET.SubElement(prefs, "pref", {"name": name, "value": value})
    return ET.tostring(root, encoding="unicode", short_empty_elements=True)


def normalize_style_id(value: str) -> str:
    return re.sub(r"^https?://(?:www\.)?zotero\.org/styles/", "http://www.zotero.org/styles/", value)


def read_document_preferences(path: str | Path):
    """Read the Word integration preference document, in custom-property order."""
    from .word import read_zotero_preferences

    chunks = read_zotero_preferences(path)
    if not chunks:
        return None
    ordered = {}
    try:
        for name, value in chunks.items():
            match = re.fullmatch(r"ZOTERO_PREF_([1-9][0-9]*)", name, re.I)
            if not match or int(match[1]) in ordered:
                raise ValueError("Invalid Zotero preference chunk name")
            ordered[int(match[1])] = value
        if set(ordered) != set(range(1, len(ordered) + 1)):
            raise ValueError("Missing Zotero preference chunk")
        root = ET.fromstring("".join(ordered[index] for index in range(1, len(ordered) + 1)))
        if root.tag != "data" or len(root.findall("style")) > 1:
            raise ValueError("Invalid Zotero preference document")
    except (ValueError, ET.ParseError) as exc:
        raise ZoteroError("Cannot read Word's Zotero XML preferences") from exc
    return root


def preference_style_id(preferences) -> str | None:
    style = preferences.find("style") if preferences is not None else None
    identity = style.get("id") if style is not None else None
    return normalize_style_id(identity) if identity else None


def preferences_for_style(preferences, style_id: str) -> str:
    """Select a style for a new working copy, retaining document-level choices."""
    if not style_id.startswith("http://www.zotero.org/styles/") or not style_id.rsplit("/", 1)[-1]:
        raise ValueError("style_id must be a full Zotero style identifier")
    if preferences is None:
        return document_preferences(style_id, locale="en-US")
    root = copy.deepcopy(preferences)
    style = root.find("style")
    if style is None:
        style = ET.SubElement(root, "style")
    style.set("id", style_id)
    style.set("bibliographyStyleHasBeenSet", "0")
    session = root.find("session")
    if session is None:
        session = ET.Element("session")
        root.insert(0, session)
    # A separate copy must not reuse a live editor's Zotero citation session.
    session.set("id", _random_id())
    return ET.tostring(root, encoding="unicode")


def preference_properties(preferences_xml: str) -> dict[str, str]:
    """Split preferences into native Word custom properties of <=255 UTF16 units."""
    ET.fromstring(preferences_xml)  # Fail before touching a package for invalid XML.
    chunks: list[str] = []
    current: list[str] = []
    units = 0
    for char in preferences_xml:
        char_units = len(char.encode("utf-16-le")) // 2
        if units + char_units > 255:
            chunks.append("".join(current))
            current, units = [], 0
        current.append(char)
        units += char_units
    if current:
        chunks.append("".join(current))
    return {f"{PREF_PREFIX}{index}": value for index, value in enumerate(chunks, start=1)}


def set_docx_preferences(path: str | Path, preferences_xml: str) -> dict[str, str]:
    """Update only Zotero preferences in an existing output DOCX atomically.

    Other custom properties and every unrelated package part are retained.
    The caller must pass its output copy, never an original user template.
    """
    target = Path(path).resolve()
    properties = preference_properties(preferences_xml)
    with zipfile.ZipFile(target) as package:
        infos = package.infolist()
        parts = {info.filename: package.read(info.filename) for info in infos}
    types = ET.fromstring(parts["[Content_Types].xml"])
    declared = [part.get("PartName", "").lstrip("/") for part in types.findall(f"{{{CT_NS}}}Override")
                if part.get("ContentType") == CUSTOM_TYPE]
    if len(declared) > 1:
        raise ValueError("Document contains multiple custom-property parts")
    custom_path = declared[0] if declared else "docProps/custom.xml"
    custom = ET.fromstring(parts[custom_path]) if custom_path in parts else ET.Element(f"{{{CUSTOM_NS}}}Properties")
    for prop in list(custom):
        if re.fullmatch(r"ZOTERO_PREF_[0-9]+", prop.get("name", ""), re.I):
            custom.remove(prop)
    next_pid = max([1] + [int(prop.get("pid", "1")) for prop in custom]) + 1
    for name, value in properties.items():
        prop = ET.SubElement(custom, f"{{{CUSTOM_NS}}}property", {"fmtid": CUSTOM_FMTID, "pid": str(next_pid), "name": name})
        ET.SubElement(prop, f"{{{VT_NS}}}lpwstr").text = value
        next_pid += 1
    parts[custom_path] = ET.tostring(custom, encoding="utf-8", xml_declaration=True)
    if not declared:
        ET.SubElement(types, f"{{{CT_NS}}}Override", {"PartName": "/" + custom_path, "ContentType": CUSTOM_TYPE})
    parts["[Content_Types].xml"] = ET.tostring(types, encoding="utf-8", xml_declaration=True)
    rels = ET.fromstring(parts["_rels/.rels"])
    if not any(rel.get("Type") == CUSTOM_REL for rel in rels):
        existing_ids = {rel.get("Id") for rel in rels}
        index = 1
        while f"rId{index}" in existing_ids:
            index += 1
        ET.SubElement(rels, f"{{{REL_NS}}}Relationship", {"Id": f"rId{index}", "Type": CUSTOM_REL, "Target": custom_path})
    parts["_rels/.rels"] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)
    handle, temporary = tempfile.mkstemp(prefix=target.stem + ".zotero-", suffix=".docx", dir=target.parent)
    os.close(handle)
    try:
        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as package:
            for info in infos:
                package.writestr(info, parts.pop(info.filename))
            for name, value in parts.items():
                package.writestr(name, value)
        os.replace(temporary, target)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
    return properties
