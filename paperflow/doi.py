"""Resolve bibliographic DOI tokens against the local Zotero library."""

from __future__ import annotations

import re
import string
from collections.abc import Iterable
from urllib.parse import unquote, urlsplit

from .zotero import ITEM_KEY, ZoteroClient, ZoteroError


_DOI = re.compile(r"10\.[0-9]+(?:\.[0-9]+)*/\S+")
_ASCII_LOWER = str.maketrans(string.ascii_uppercase, string.ascii_lowercase)
_SEARCH_LIMIT = 100


def normalize_doi(raw: str) -> str | None:
    """Normalize a DOI token, DOI-prefixed value, or doi.org resolver URL.

    Keep suffix punctuation, including a final period, as part of the token.
    The bibliography parser is responsible for separating sentence punctuation.
    DOI name comparison folds only Basic Latin uppercase letters.
    """
    if not isinstance(raw, str):
        return None
    value = raw.strip()
    if value[:4].lower() == "doi:":
        value = value[4:].strip()
    if not value.isprintable():
        return None
    if value.lower().startswith(("https://", "http://")):
        try:
            resolver = urlsplit(value)
            if resolver.netloc.lower() not in {"doi.org", "dx.doi.org"}:
                return None
            if re.search(r"%(?![0-9a-fA-F]{2})", resolver.path):
                return None
            value = unquote(resolver.path.removeprefix("/"), errors="strict")
        except ValueError:
            return None
    return value.translate(_ASCII_LOWER) if value.isprintable() and _DOI.fullmatch(value) else None


def resolve_dois(dois: Iterable[str], *, client: ZoteroClient | None = None) -> dict[str, list[dict[str, str]]]:
    """Return exact local parent-record matches, including any duplicate keys.

    Zotero's broader quick search supplies candidates. Only a matching DOI in
    a record's own metadata qualifies. If the result could be truncated, it
    cannot establish that the exact match is unique.
    """
    requested: dict[str, list[dict[str, str]]] = {}
    for raw in dois:
        doi = normalize_doi(raw)
        if doi:
            requested.setdefault(doi, [])
    if not requested:
        return requested

    if client is None:
        client = ZoteroClient()
    for doi in requested:
        records = client.search(doi, limit=_SEARCH_LIMIT, mode="everything")
        if not isinstance(records, list):
            raise ZoteroError("Expected a list of records from the local Zotero DOI search")
        if len(records) >= _SEARCH_LIMIT:
            raise ZoteroError(f"The local search for {doi} reached its result limit. Review the DOI match in Zotero.")
        matches = {}
        for record in records:
            if not isinstance(record, dict) or not isinstance(record.get("data"), dict):
                raise ZoteroError("Invalid metadata record returned by the local Zotero DOI search")
            data = record["data"]
            if data.get("itemType") in {"attachment", "note", "annotation"} or data.get("deleted"):
                continue
            if normalize_doi(data.get("DOI")) != doi:
                continue
            key = record.get("key")
            if not isinstance(key, str) or not ITEM_KEY.fullmatch(key) or data.get("key", key) != key:
                raise ZoteroError(f"Exact DOI {doi} matched a record without a valid Zotero item key")
            title = data.get("title", "")
            match = {"key": key, "title": title if isinstance(title, str) else "", "doi": doi}
            if key in matches and matches[key] != match:
                raise ZoteroError(f"The local search returned inconsistent metadata for Zotero item {key}")
            matches[key] = match
        requested[doi] = [matches[key] for key in sorted(matches)]
    return requested
