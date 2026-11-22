"""Plan or create an immutable local PaperFlow source release; never install software.

The default is read-only. Root must complete acceptance before using --build.
Only the source allowlist and selected artifacts are copied, never the output tree.
"""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import importlib.metadata
import json
import os
from pathlib import Path, PurePosixPath
import re
import sys
import tomllib
from zipfile import ZIP_DEFLATED, ZipFile


PROJECT = Path(__file__).resolve().parents[1]
FIXED_FILES = (
    "PaperFlow.cmd", "PaperFlow-界面.cmd", "pyproject.toml", "environment.yml", "environment-win-64.lock.yml",
    "README.md", "开始使用.md",
    "templates/computers-and-structures.docx",
    "templates/computers-and-structures.provenance.json",
    "profiles/computers-and-structures.json",
    "references/computers-and-structures/computers-and-structures.csl",
    "references/computers-and-structures/elsevier-vancouver.csl",
    "docs/computers-and-structures-requirements.md", "docs/排版与修订.md", "docs/ocr.md", "docs/polishing-sources.md",
)
SOURCE_PATTERNS = {
    "paperflow": {".py"}, "scripts": {".py", ".ps1"}, "tests": {".py"},
    "examples/computers-and-structures": {".md", ".json", ".csv", ".png", ".pdf"},
    "examples/content-review": {".md", ".json", ".txt"},
}
ACCEPTANCE_ROOT_FILES = {
    "release-acceptance-evidence.json", "acceptance-evidence.json", "coverage.md",
    "source.docx", "manuscript.docx", "template.docx", "metadata.json", "response.png", "response.pdf",
    "revision-response.png", "revision-response.pdf", "revision-response.csv", "initial-acceptance.json",
}
ACCEPTANCE_STAGES = {"built", "revised", "refreshed", "initial", "final"}
ACCEPTANCE_STAGE_FILES = {
    "editable-master.docx", "editable-master.pdf", "submission-copy.docx", "submission-copy.pdf", "workflow.json",
}
ARTWORK_SUFFIXES = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".pdf", ".eps", ".svg"}
RUNTIME_DISTRIBUTIONS = ("python-docx", "lxml", "pywin32", "PyMuPDF", "Pillow", "mcp", "matplotlib")


def sha256(data):
    return hashlib.sha256(data).hexdigest()


def source_path(relative):
    candidate = (PROJECT / relative).resolve(strict=True)
    if not candidate.is_relative_to(PROJECT) or not candidate.is_file():
        raise ValueError(f"Source must be a workspace file: {relative}")
    return candidate


def relative_name(value):
    name = PurePosixPath(str(value).replace("\\", "/"))
    if not name.parts or name.is_absolute() or any(part in {"", ".", ".."} for part in name.parts) or ":" in str(name):
        raise ValueError(f"Unsafe package-relative path: {value}")
    return name.as_posix()


def acceptance_files(folder, extra):
    """Use explicit root/stage/page/artwork rules, not an arbitrary directory copy."""
    folder = folder.resolve(strict=True)
    if not folder.is_dir() or not folder.is_relative_to(PROJECT):
        raise ValueError("Acceptance directory must be inside this workspace")
    found = {}
    for candidate in folder.rglob("*"):
        if not candidate.is_file():
            continue
        relative = candidate.relative_to(folder)
        parts = relative.parts
        allowed = len(parts) == 1 and relative.name in ACCEPTANCE_ROOT_FILES
        allowed |= len(parts) == 2 and parts[0] in ACCEPTANCE_STAGES and relative.name in ACCEPTANCE_STAGE_FILES
        allowed |= (len(parts) >= 2 and parts[0] in ACCEPTANCE_STAGES and parts[1] == "artwork"
                    and candidate.suffix.casefold() in ARTWORK_SUFFIXES)
        allowed |= (len(parts) == 2 and (parts[0] == "pages" or parts[0].endswith("-pages"))
                    and re.fullmatch(r"page-\d+\.png", relative.name) is not None)
        allowed |= (len(parts) == 3 and parts[0] in ACCEPTANCE_STAGES
                    and parts[1] in {"editable-master-pages", "submission-copy-pages"}
                    and re.fullmatch(r"page-\d+\.png", relative.name) is not None)
        if allowed:
            resolved = candidate.resolve(strict=True)
            if not resolved.is_relative_to(folder):
                raise ValueError(f"Acceptance path leaves its directory: {candidate}")
            found[relative.as_posix()] = resolved
    for value in extra:
        name = relative_name(value)
        path = (folder / name).resolve(strict=True)
        if not path.is_file() or not path.is_relative_to(folder):
            raise ValueError(f"Acceptance file leaves its directory: {value}")
        found[name] = path
    return found


def collect(acceptance=None, extra=()):
    files = {name: source_path(name) for name in FIXED_FILES}
    for directory, suffixes in SOURCE_PATTERNS.items():
        for path in sorted((PROJECT / directory).rglob("*")):
            if path.is_file() and path.suffix.casefold() in suffixes and "__pycache__" not in path.parts:
                name = path.relative_to(PROJECT).as_posix()
                files[name] = source_path(name)
    if acceptance is not None:
        prefix = "acceptance/" + acceptance.resolve(strict=True).name
        for name, path in acceptance_files(acceptance, extra).items():
            files[prefix + "/" + name] = path
    elif extra:
        raise ValueError("--acceptance-file requires --acceptance")
    return dict(sorted(files.items()))


def runtime_metadata():
    versions = {}
    for name in RUNTIME_DISTRIBUTIONS:
        try:
            versions[name] = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            versions[name] = None
    return {
        "python": {"version": sys.version},
        "python_dependencies": versions,
        "launcher": "PaperFlow.cmd uses the paperflow Conda environment or an explicit PAPERFLOW_PYTHON.",
        "external_software_not_bundled": ["Microsoft Word for Windows", "MathType with native SDK/OLE support",
                                          "Zotero Desktop and its Word add-in", "Windows OCR and chosen language pack"],
        "additional_fixture_dependency": "matplotlib (only to regenerate example plot data/assets)",
        "installs_performed": False,
    }


def verify_archive(path, root_name, expected):
    with ZipFile(path) as archive:
        names = {name.removeprefix(root_name + "/"): name for name in archive.namelist()}
        if set(names) != set(expected):
            raise RuntimeError("ZIP entry inventory differs from the release manifest")
        for relative, digest in expected.items():
            if sha256(archive.read(names[relative])) != digest:
                raise RuntimeError(f"ZIP verification failed: {relative}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--version", default="0.1.0-rc1", help="Package label; does not edit the application version")
    parser.add_argument("--output-root", type=Path, default=PROJECT / "dist")
    parser.add_argument("--acceptance", type=Path, help="Final inspected acceptance directory within the workspace")
    parser.add_argument("--acceptance-file", action="append", default=[], metavar="RELATIVE_PATH",
                        help="Explicit additional acceptance file, relative to --acceptance")
    parser.add_argument("--build", action="store_true", help="Create a new release after acceptance; default only prints a plan")
    args = parser.parse_args()
    if re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]*", args.version) is None:
        parser.error("Version must contain only letters, digits, dots, underscores and hyphens")
    files = collect(args.acceptance, args.acceptance_file)
    before = {name: {"sha256": sha256(path.read_bytes()), "bytes": path.stat().st_size} for name, path in files.items()}
    root = args.output_root.resolve()
    if not root.is_relative_to(PROJECT):
        raise ValueError("Release output must stay inside this workspace")
    name = "PaperFlow-v" + args.version
    folder, archive_path = root / name, root / (name + ".zip")
    evidence_path = root / (name + ".build-evidence.json")
    if folder.exists() or archive_path.exists() or evidence_path.exists():
        raise FileExistsError("A release folder, ZIP or build record already exists; choose a new package version")
    report = {"mode": "build" if args.build else "plan", "package": name, "source_file_count": len(files),
              "source_bytes": sum(record["bytes"] for record in before.values()), "folder": str(folder),
              "zip": str(archive_path), "acceptance_included": args.acceptance is not None,
              "file_allowlist": list(files), "writes_performed": False}
    if not args.build:
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0
    if os.name != "nt":
        raise RuntimeError("This release targets Windows and must be built on the verified Windows workspace")
    runtime = runtime_metadata()
    missing = [name for name, version in runtime["python_dependencies"].items() if version is None]
    if missing:
        raise RuntimeError("Required local Python dependencies are missing: " + ", ".join(missing))
    root.mkdir(parents=True, exist_ok=True)
    folder.mkdir(exist_ok=False)
    payload = {}
    for relative, source in files.items():
        data = source.read_bytes()
        if sha256(data) != before[relative]["sha256"]:
            raise RuntimeError(f"Source changed after planning: {relative}")
        target = folder / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open("xb") as stream:
            stream.write(data)
        payload[relative] = before[relative]["sha256"]
    generated = {
        "runtime-requirements.json": json.dumps(runtime, ensure_ascii=False, indent=2) + "\n",
        "requirements-runtime.txt": "# Versions present when this package was built; no installation is performed.\n"
            + "\n".join(f"{name}=={version}" for name, version in runtime["python_dependencies"].items()) + "\n",
    }
    for relative, value in generated.items():
        data = value.encode("utf-8")
        with (folder / relative).open("xb") as stream:
            stream.write(data)
        payload[relative] = sha256(data)
    manifest = {
        "release_label": args.version, "source_application_version": tomllib.loads((PROJECT / "pyproject.toml").read_text(encoding="utf-8"))["project"]["version"],
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "scope": "Local source handoff; external licensed applications and the Python runtime are not included.",
        "acceptance_directory": args.acceptance.name if args.acceptance else None,
        "excluded": ["prior output tree", "caches", "Zotero API/library/install snapshots", "credentials", "runtimes", "installed applications"],
        "files": [{"path": relative, "sha256": digest, "bytes": (folder / relative).stat().st_size}
                  for relative, digest in sorted(payload.items())],
        "manifest_self_hash_excluded": True,
    }
    data = (json.dumps(manifest, ensure_ascii=False, indent=2) + "\n").encode("utf-8")
    with (folder / "RELEASE-MANIFEST.json").open("xb") as stream:
        stream.write(data)
    payload["RELEASE-MANIFEST.json"] = sha256(data)
    sums = "".join(f"{digest}  {relative}\n" for relative, digest in sorted(payload.items())).encode("utf-8")
    with (folder / "SHA256SUMS.txt").open("xb") as stream:
        stream.write(sums)
    payload["SHA256SUMS.txt"] = sha256(sums)
    for relative, source in files.items():
        if sha256(source.read_bytes()) != before[relative]["sha256"]:
            raise RuntimeError(f"Source changed during packaging: {relative}")
    with ZipFile(archive_path, "x", compression=ZIP_DEFLATED, compresslevel=9) as archive:
        for relative in sorted(payload):
            archive.write(folder / relative, arcname=name + "/" + relative)
    verify_archive(archive_path, name, payload)
    report.update(writes_performed=True, source_unchanged=True, archive_verified=True,
                  packaged_file_count=len(payload), zip_bytes=archive_path.stat().st_size,
                  zip_sha256=sha256(archive_path.read_bytes()))
    with evidence_path.open("x", encoding="utf-8") as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2)
        stream.write("\n")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
