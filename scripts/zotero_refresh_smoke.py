"""Refresh only a test document using the installed trusted Zotero Word add-in."""
import argparse
import json
from pathlib import Path
import sys
import time

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from scripts.zotero_fixture import FIRST_DOI, fixture_key
from paperflow.word import private_word, export_pdf, snapshot
from paperflow.zotero import ZoteroClient
from paperflow.mathtype import audit_docx_math
from paperflow.package import audit_docx


def main():
    parser = argparse.ArgumentParser(); parser.add_argument("source"); parser.add_argument("output"); args = parser.parse_args()
    source, output = Path(args.source).resolve(strict=True), Path(args.output).resolve()
    if output.exists(): raise FileExistsError(output)
    expected = ZoteroClient().get_item(fixture_key(FIRST_DOI))["item_data"]["title"]
    print("Starting isolated Word with installed Zotero add-in", flush=True)
    with private_word() as app:
        doc = app.Documents.Open(FileName=str(source), ReadOnly=True, AddToRecentFiles=False, Visible=False)
        try:
            doc.SaveAs2(FileName=str(output), FileFormat=12, AddToRecentFiles=False)
            doc.Activate()
            addins = [{"name": a.Name, "installed": bool(a.Installed)} for a in app.AddIns]
            print(json.dumps({"loaded_addins": addins}, ensure_ascii=False), flush=True)
            # This invokes a preinstalled trusted global template. No templates,
            # macro policies or Trust Center settings are installed or changed.
            app.Run("Zotero.ZoteroRefresh")
            print("Zotero native refresh macro returned; checking document", flush=True)
            verified = False
            started = time.monotonic()
            while time.monotonic() - started < 90:
                import pythoncom
                pythoncom.PumpWaitingMessages()
                codes = "\n".join(str(doc.Fields.Item(n).Code.Text) for n in range(1, doc.Fields.Count + 1))
                visible = str(doc.Content.Text)
                if "PAPERFLOW_PENDING" not in visible and "PAPERFLOW_STALE_TITLE" not in codes and expected in visible:
                    verified = True
                    break
                time.sleep(0.5)
            doc.Save()
            report = {"zotero_native_refresh_verified": verified, "elapsed": time.monotonic()-started,
                      "expected_title": expected, "snapshot": snapshot(doc), "addins": addins}
            if not verified:
                report["visible_text"] = str(doc.Content.Text)
        finally:
            doc.Close(SaveChanges=0); doc = None
        if verified:
            report["pdf"] = export_pdf(output, output.with_suffix(".pdf"), app=app)
    report["structural_audit"] = audit_docx(output)
    report["mathtype_audit"] = audit_docx_math(output)
    output.with_suffix(".audit.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"verified": verified, "output": str(output)}, ensure_ascii=False), flush=True)
    if not verified: raise RuntimeError("Native Zotero refresh did not meet the verification gate")


if __name__ == "__main__": main()
