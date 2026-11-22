"""Exercise saved-table audit failures and recheck existing native build artifacts."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Pt

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

from paperflow.package import digest, q
from paperflow.table_native import audit_table_layout
from paperflow.word import private_word


def _width(parent, name, value):
    node = OxmlElement("w:" + name)
    node.set(q("w:type"), "dxa")
    node.set(q("w:w"), str(value))
    if parent.tag == q("w:tblPr"):
        parent.insert_element_before(node, "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook")
    else:
        parent.append(node)
    return node


def _make_source(path, *, omitted):
    document = Document()
    section = document.sections[0]
    section.page_width, section.page_height = Pt(544), Pt(792)
    section.left_margin = section.right_margin = Pt(72)
    widths = [100, 400, 100] if omitted else [300]
    table = document.add_table(rows=1, cols=len(widths))
    table.autofit = False
    for column, width in zip(table.columns, widths, strict=True):
        column.width = Pt(width)
    for cell, width in zip(table.rows[0].cells, widths, strict=True):
        cell.width = Pt(width)
        cell.text = "Saved table width witness"
    preferred = table._tbl.tblPr.find(q("w:tblW"))
    preferred.set(q("w:type"), "dxa")
    preferred.set(q("w:w"), str(sum(widths) * 20))
    if omitted:
        row = table.rows[0]._tr
        cells = list(row.findall(q("w:tc")))
        row.remove(cells[0])
        row.remove(cells[-1])
        properties = row.get_or_add_trPr()
        for name in ("gridBefore", "gridAfter"):
            node = OxmlElement("w:" + name)
            node.set(q("w:val"), "1")
            properties.append(node)
        _width(properties, "wBefore", 2000)
        _width(properties, "wAfter", 2000)
    else:
        _width(table._tbl.tblPr, "tblInd", -2000)
    document.save(path)


def _checks(results, hashes):
    negative, omitted = (results[name] for name in ("negative-indent", "omitted-columns"))
    return {
        "negative_indent_requires_native_placement_review": any("negative leading indent" in issue for issue in negative["unresolved"]),
        "overwide_omitted_grid_document_requires_native_review": any("exceeds" in issue for issue in omitted["unresolved"])
            and omitted["tables"][0]["saved_grid_width_pt"] == 600
            and not omitted["tables"][0]["within_width_limit"],
        "existing_native_build_revision_and_variants_still_fit": all(
            results[name]["native_reopened"] and not results[name]["unresolved"]
            and all(table["within_width_limit"] for table in results[name]["tables"])
            for name in ("built", "revised", "variants")),
        "all_inputs_unchanged": all(digest(Path(path).read_bytes()) == before for path, before in hashes.items()),
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument("--recheck", action="store_true", help="Recheck recorded native results without reopening Word")
    args = parser.parse_args()
    output = args.output_dir.resolve()
    path = output / "table-native-audit-evidence.json"
    if args.recheck:
        previous = path.read_bytes()
        report = json.loads(previous)
        if not all(digest(Path(name).read_bytes()) == before for name, before in report["input_sha256"].items()):
            raise RuntimeError("A native input changed; saved audit evidence cannot be reused")
        report["checks"] = _checks(report["native"], report["input_sha256"])
        report["pass"] = all(report["checks"].values())
        report["assertion_recheck"] = {
            "native_word_rerun": False,
            "reason": "Word exposes the omitted-grid example as a 600 pt native cell. Either its native-cell or saved-grid overflow message proves the required review; exact message text is not the behavior under test.",
        }
        backup = path.with_name("table-native-audit-evidence.before-assertion-recheck.json")
        if not backup.exists():
            backup.write_bytes(previous)
        path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(json.dumps({"pass": report["pass"], "checks": report["checks"], "evidence": str(path)}, indent=2))
        return 0 if report["pass"] else 1
    output.mkdir(parents=True, exist_ok=False)
    sources = {name: output / (name + ".docx") for name in ("negative-indent", "omitted-columns")}
    for name, path in sources.items():
        _make_source(path, omitted=name == "omitted-columns")
    artifacts = {
        "built": PROJECT / "output/table-width-fit-smoke-001/built/editable-master.docx",
        "revised": PROJECT / "output/table-width-fit-smoke-001/revised/editable-master.docx",
        "variants": PROJECT / "output/table-width-variants-smoke-001/built/editable-master.docx",
    }
    hashes = {str(path): digest(path.read_bytes()) for path in (*sources.values(), *artifacts.values())}
    results = {}
    with private_word() as app:
        for name, path in {**sources, **artifacts}.items():
            results[name] = audit_table_layout(path, app=app)
    checks = _checks(results, hashes)
    report = {"pass": all(checks.values()), "checks": checks, "input_sha256": hashes, "native": results,
              "boundary": "Synthetic audit cases and existing synthetic native fixtures; no manuscript compliance claim."}
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"pass": report["pass"], "checks": checks, "evidence": str(path)}, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
