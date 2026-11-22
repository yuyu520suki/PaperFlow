"""Build a C&S sample, then recheck a changed Word master with its old metadata."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

from lxml import etree

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from paperflow.package import digest, q, read_package, write_package, xml
from paperflow.pipeline import CNS, revise_master, run_pipeline


def _hashes(paths):
    return {str(path): digest(path.read_bytes()) for path in paths}


def _revise_copy(master, revised, changes):
    """Change the two plain text paragraphs in the fixture's copied package."""
    parts = read_package(master)
    document = xml(parts["word/document.xml"])
    paragraphs = document.find(q("w:body")).findall(q("w:p"))
    for original, replacement in changes.items():
        matches = []
        for paragraph in paragraphs:
            runs = list(paragraph.iter(q("w:t")))
            if "".join(run.text or "" for run in runs) == original:
                matches.append(runs)
        if len(matches) != 1:
            raise RuntimeError(f"Expected one paragraph for {original!r}, found {len(matches)}")
        first, *remaining = matches[0]
        first.text = replacement
        for run in remaining:
            run.text = ""
    parts["word/document.xml"] = etree.tostring(document, encoding="UTF-8", xml_declaration=True)
    write_package(parts, revised)


def _preflight(report):
    return {check["name"]: check for check in report["journal_audit"]["checks"]}


def _observed(check, expected_count, expected_pass):
    detail = check["detail"]
    return (check["pass"] is expected_pass
            and detail["source"] == "word/document.xml"
            and detail["observed_count"] == expected_count
            and bool(detail["body_paragraphs"]))


def _zotero_refreshed(report):
    return any(step["stage"] == "zotero"
               and step["evidence"]["macro_invoked"]
               and step["evidence"]["zotero_native_refresh_verified"]
               for step in report["steps"])


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    args = parser.parse_args()
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    examples = PROJECT / "examples" / CNS
    source = examples / "manuscript.md"
    template = PROJECT / "templates" / f"{CNS}.docx"
    metadata_path = examples / "metadata.json"
    metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
    original_inputs = [source, template, metadata_path, examples / "response.pdf", examples / "response.png"]
    inputs_before = _hashes(original_inputs)

    build = output / "built"
    original = run_pipeline(source, template, build, journal=CNS, submission_copy=True,
                            table_policy="three-line", metadata=metadata)
    first_outputs = sorted(path for path in build.rglob("*") if path.is_file())
    first_outputs_before = _hashes(first_outputs)
    original_checks = _preflight(original)

    revised = output / "word-revision.docx"
    abstract_count = 260
    new_abstract = " ".join(["response"] * abstract_count) + "."
    new_keywords = ["Response", "Damping", "Vibration", "Structures", "Dynamics",
                    "Simulation", "Verification", "Workflow"]
    _revise_copy(build / "editable-master.docx", revised, {
        metadata["abstract"]: new_abstract,
        "Keywords: " + "; ".join(metadata["keywords"]): "Keywords: " + "; ".join(new_keywords),
    })
    revised_before = _hashes([revised])
    refreshed = output / "refreshed"
    revision = revise_master(revised, refreshed, journal=CNS, submission_copy=True,
                             metadata=metadata, artwork_root=examples)
    revised_checks = _preflight(revision)
    checks = {
        "original_abstract_count": _observed(original_checks["abstract_words"], len(metadata["abstract"].split()), True),
        "original_keyword_count": _observed(original_checks["keyword_count"], len(metadata["keywords"]), True),
        "revised_abstract_exceeds_limit": _observed(revised_checks["abstract_words"], abstract_count, False),
        "revised_keywords_exceed_limit": _observed(revised_checks["keyword_count"], len(new_keywords), False),
        "stale_metadata_detected": all(
            revised_checks[name]["pass"] is False
            for name in ("abstract_metadata_consistency", "keywords_metadata_consistency")
        ),
        "zotero_refreshed_in_both_builds": _zotero_refreshed(original) and _zotero_refreshed(revision),
        "inputs_and_metadata_unchanged": _hashes(original_inputs) == inputs_before,
        "original_build_unchanged": _hashes(first_outputs) == first_outputs_before,
        "revised_source_unchanged": revision["source_unchanged"] and _hashes([revised]) == revised_before,
    }
    report = {
        "pass": all(checks.values()), "checks": checks,
        "original": original["journal_audit"], "revised": revision["journal_audit"],
        "workflows": [str(build / "workflow.json"), str(refreshed / "workflow.json")],
    }
    evidence = output / "journal-preflight-evidence.json"
    evidence.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"pass": report["pass"], "checks": checks, "evidence": str(evidence)}, ensure_ascii=False))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
