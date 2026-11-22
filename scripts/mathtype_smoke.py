"""Replace the fixture's OMML equations with installed MathType OLE objects."""
import argparse
import hashlib
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from paperflow.word import private_word, export_pdf
from paperflow.mathtype import MathTypeSDK, discover_mathtype, audit_docx_math


def main():
    p = argparse.ArgumentParser(); p.add_argument("source"); p.add_argument("output"); args = p.parse_args()
    source, output = Path(args.source).resolve(strict=True), Path(args.output).resolve()
    if output.exists(): raise FileExistsError(output)
    installation = discover_mathtype()
    blank = Path(installation.directory) / "Office Support/BlankEqn.doc"
    equations = [r"\mathbf{M}\ddot{\mathbf{x}}+\mathbf{C}\dot{\mathbf{x}}+\mathbf{K}\mathbf{x}=\mathbf{f}(t)",
                 r"\mathbf{A}=\begin{bmatrix}1&2\\3&4\end{bmatrix}"]
    print("Starting private Word / native MathType", flush=True)
    with private_word() as app:
        document = app.Documents.Open(FileName=str(source), ReadOnly=True, AddToRecentFiles=False, Visible=False)
        try:
            if document.OMaths.Count != len(equations):
                raise RuntimeError(f"Expected {len(equations)} OMML equations, got {document.OMaths.Count}")
            with MathTypeSDK() as sdk:
                for index in range(len(equations), 0, -1):
                    print(f"Replacing equation {index}", flush=True)
                    location = document.OMaths.Item(index).Range.Duplicate
                    start = location.Start
                    location.Delete()
                    insertion = document.Range(start, start)
                    insertion.InsertFile(FileName=str(blank), Range="MTBlankEqn", ConfirmConversions=False, Link=False, Attachment=False)
                    shapes = [document.InlineShapes.Item(n) for n in range(1, document.InlineShapes.Count + 1)]
                    shape = next(s for s in shapes if s.Range.Start == start)
                    shape.OLEFormat.DoVerb(2)
                    ole = shape.OLEFormat.Object
                    sdk.set_ole_equation(ole, equations[index - 1], "latex")
                    ole = shape = shapes = location = insertion = None
            document.SaveAs2(FileName=str(output), FileFormat=12, AddToRecentFiles=False)
        finally:
            document.Close(SaveChanges=0); document = None
        print("Saved native OLE equations; reopening", flush=True)
        doc = app.Documents.Open(FileName=str(output), ReadOnly=True, AddToRecentFiles=False, Visible=False)
        try:
            # Reopening and reading native equation bytes is a real OLE round trip.
            roundtrip = []
            with MathTypeSDK() as sdk:
                for n in range(1, doc.InlineShapes.Count + 1):
                    shape = doc.InlineShapes.Item(n)
                    if shape.Type == 1 and shape.OLEFormat.ProgID == "Equation.DSMT4":
                        shape.OLEFormat.DoVerb(2)
                        ole = shape.OLEFormat.Object
                        native = sdk.get_ole_mtef(ole)
                        roundtrip.append({"mtef_bytes": len(native), "sha256": hashlib.sha256(native).hexdigest()})
                        sdk.close_ole_equation(ole, save_changes=False)
                        ole = None
                    shape = None
        finally:
            doc.Close(SaveChanges=0); doc = None
        exported = export_pdf(output, output.with_suffix(".pdf"), app=app)
    audit = audit_docx_math(output)
    output.with_suffix(".audit.json").write_text(json.dumps({"math_audit": audit, "roundtrip": roundtrip,
        "pdf": exported, "fixture_only": True}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(audit, ensure_ascii=False), flush=True)


if __name__ == "__main__": main()
