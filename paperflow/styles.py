"""Install explicitly approved CSL bytes through Zotero's native connector.

Dry-run is the default. Installation never restarts Zotero and never uses a
library-write endpoint. Dependencies must be supplied or already installed;
this code does not download additional styles.
"""

from __future__ import annotations

import hashlib
import json
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Iterable, Mapping

CSL_NS = {"c": "http://purl.org/net/xbiblio/csl"}


class StyleInstallationError(RuntimeError):
    pass


def describe_style(path: str | Path) -> dict[str, Any]:
    source = Path(path).resolve(strict=True)
    data = source.read_bytes()
    root = ET.fromstring(data)
    if root.tag != "{http://purl.org/net/xbiblio/csl}style":
        raise ValueError(f"Not a CSL style: {source}")
    ident = root.findtext("c:info/c:id", namespaces=CSL_NS)
    title = root.findtext("c:info/c:title", namespaces=CSL_NS)
    if not ident or not title:
        raise ValueError(f"Missing CSL identity/title: {source}")
    parents = [link.get("href") for link in root.findall("c:info/c:link", CSL_NS) if link.get("rel") == "independent-parent"]
    if len(parents) > 1:
        raise ValueError("A dependent style must not have multiple independent parents")
    return {"path": str(source), "filename": source.name, "style_id": ident, "title": title, "independent_parent": parents[0] if parents else None, "sha256": hashlib.sha256(data).hexdigest(), "size_bytes": len(data)}


def style_directory_snapshot(styles_dir: str | Path) -> dict[str, dict[str, Any]]:
    root = Path(styles_dir).resolve(strict=True)
    result = {}
    for path in sorted(root.rglob("*.csl")):
        style = describe_style(path)
        result[path.relative_to(root).as_posix()] = style
    return result


def install_styles(
    paths: Iterable[str | Path],
    *,
    expected_sha256: Mapping[str, str],
    styles_dir: str | Path,
    install: bool = False,
    base_url: str = "http://127.0.0.1:23119",
) -> dict[str, Any]:
    """Plan, or explicitly install, hash-approved styles in dependency order.

    ``expected_sha256`` maps exact source filenames to user-approved digests.
    Existing differing styles are never overwritten by this helper.
    """
    origin = urllib.parse.urlsplit(base_url)
    if origin.scheme != "http" or origin.hostname not in {"127.0.0.1", "localhost", "::1"} or origin.path not in {"", "/"} or origin.query or origin.fragment or origin.username or origin.password:
        raise ValueError("Connector origin must be a loopback HTTP origin")
    styles = [describe_style(path) for path in paths]
    if len({s["style_id"] for s in styles}) != len(styles):
        raise ValueError("Duplicate style IDs in requested scope")
    for style in styles:
        approved = expected_sha256.get(style["filename"], "").lower()
        if len(approved) != 64 or approved != style["sha256"]:
            raise ValueError(f"Missing approval digest or checksum mismatch: {style['filename']}")
    before = style_directory_snapshot(styles_dir)
    existing_ids = {s["style_id"]: s for s in before.values()}
    planned_ids = {s["style_id"] for s in styles}
    for style in styles:
        parent = style["independent_parent"]
        if parent and parent not in planned_ids and parent not in existing_ids:
            raise ValueError(f"Dependency not supplied/installed; refusing implicit download: {parent}")
        existing = existing_ids.get(style["style_id"])
        if existing and existing["sha256"] != style["sha256"]:
            raise ValueError(f"Different style already installed; refusing overwrite: {style['style_id']}")
    ordered = []
    pending = styles[:]
    available = set(existing_ids)
    while pending:
        candidates = [s for s in pending if not s["independent_parent"] or s["independent_parent"] in available]
        if not candidates:
            raise ValueError("Cyclic or unresolved style dependencies")
        for style in candidates:
            ordered.append(style)
            available.add(style["style_id"])
            pending.remove(style)
    report = {"mode": "install" if install else "dry_run", "install_requested_explicitly": install, "endpoint": base_url.rstrip("/") + "/connector/installStyle", "plan": ordered, "before": before, "results": []}
    if not install:
        report["writes_performed"] = False
        return report
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    for style in ordered:
        data = Path(style["path"]).read_bytes()
        if hashlib.sha256(data).hexdigest() != style["sha256"]:
            raise StyleInstallationError(f"Source bytes changed after approval check: {style['filename']}")
        if style["style_id"] in existing_ids:
            report["results"].append({"style_id": style["style_id"], "status": "already_present_same_bytes"})
            continue
        url = report["endpoint"] + "?" + urllib.parse.urlencode({"origin": style["path"]})
        request = urllib.request.Request(url, data=data, method="POST", headers={"Content-Type": "application/vnd.citationstyles.style+xml", "X-Zotero-Connector-API-Version": "3"})
        try:
            with opener.open(request, timeout=30) as response:
                body = response.read().decode("utf-8")
                result = {"style_id": style["style_id"], "http_status": response.status, "response": json.loads(body)}
        except urllib.error.HTTPError as exc:
            raise StyleInstallationError(f"Zotero style installation returned {exc.code}: {exc.read().decode('utf-8', errors='replace')}") from exc
        if result["http_status"] != 201 or result["response"].get("name") != style["title"]:
            raise StyleInstallationError(f"Unexpected Zotero style install result: {result}")
        report["results"].append(result)
    after = style_directory_snapshot(styles_dir)
    report["after"] = after
    report["added_paths"] = sorted(set(after) - set(before))
    report["removed_paths"] = sorted(set(before) - set(after))
    report["changed_paths"] = sorted(name for name in set(before) & set(after) if before[name]["sha256"] != after[name]["sha256"])
    report["writes_performed"] = any(r.get("http_status") == 201 for r in report["results"])
    report["verification"] = []
    for style in styles:
        installed = [s for s in after.values() if s["style_id"] == style["style_id"]]
        if len(installed) != 1 or installed[0]["sha256"] != style["sha256"] or installed[0]["independent_parent"] != style["independent_parent"]:
            raise StyleInstallationError(f"Installed style identity/bytes mismatch: {style['style_id']}")
        report["verification"].append(installed[0])
    return report
