"""Edit a real generated Word master, then prove citation and MathType round trips."""
import argparse
import hashlib
import json
from pathlib import Path
import sys

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from scripts.zotero_fixture import FIRST_DOI, SECOND_DOI, fixture_key
from paperflow.package import expand_markers, audit_docx, read_package, xml, q
from paperflow.zotero import ZoteroClient
from paperflow.mathtype import MathTypeSDK, audit_docx_math
from paperflow.word import private_word
from paperflow.pipeline import revise_master, CNS


def mtef_hashes(path):
    return [o["mtef_sha256"]
            for o in audit_docx_math(path)["objects"] if o["valid_native_storage"]]


def main():
    p=argparse.ArgumentParser(); p.add_argument("master"); p.add_argument("output_dir")
    p.add_argument("--native-edit-docx",help="A separate copy of the master already edited in MathType")
    args=p.parse_args()
    src=Path(args.master).resolve(strict=True); out=Path(args.output_dir).resolve(); out.mkdir(parents=True,exist_ok=False)
    original_hash=hashlib.sha256(src.read_bytes()).hexdigest(); old_ole=mtef_hashes(src)
    edit_src=Path(args.native_edit_docx).resolve(strict=True) if args.native_edit_docx else src
    edit_source_hash=hashlib.sha256(edit_src.read_bytes()).hexdigest()
    edit_source_mtef=mtef_hashes(edit_src)
    if args.native_edit_docx and (len(old_ole)!=2 or len(edit_source_mtef)!=2 or
                                  edit_source_mtef[0]!=old_ole[0] or edit_source_mtef[1]==old_ole[1]):
        raise RuntimeError("The manually edited source must change just the second native equation")
    new_key=fixture_key(SECOND_DOI)
    item=ZoteroClient().get_item(new_key)
    draft=out/"word-revision.docx"
    with private_word() as app:
        doc=app.Documents.Open(FileName=str(edit_src),ReadOnly=True,AddToRecentFiles=False,Visible=False)
        try:
            doc.SaveAs2(FileName=str(draft),FileFormat=12,AddToRecentFiles=False)
            intro=next(para for para in doc.Paragraphs if str(para.Range.Text).startswith("This file is an integration test."))
            point=doc.Range(intro.Range.Start,intro.Range.Start)
            point.InsertBefore("An additional bibliography fixture is [[CITE "+new_key+"]]. ")
            cell=doc.Tables.Item(1).Cell(2,2).Range.Duplicate; cell.End-=1; cell.Text="7.50"
            shapes=[s for s in doc.InlineShapes if s.Type==1 and s.OLEFormat.ProgID=="Equation.DSMT4"]
            if len(shapes)!=2: raise RuntimeError("Fixture requires two existing native formulas")
            if not args.native_edit_docx:
                equation=shapes[1]
                with MathTypeSDK() as sdk:
                    equation.OLEFormat.DoVerb(2)
                    target=equation.OLEFormat.Object
                    # The SDK inserts at the active selection.
                    sdk.set_ole_equation(target,r"+\mathbf{I}","latex")
                    target=None
                equation=None
            shapes=[]; intro=point=cell=None
            doc.Save()
        finally: doc.Close(SaveChanges=0)
    linked=out/"word-revision-linked.docx"
    expand_markers(draft,linked,{new_key:item})
    final=out/"refreshed"
    artwork_root=Path(__file__).resolve().parents[1]/"examples"/CNS
    metadata=json.loads((artwork_root/"metadata.json").read_text(encoding="utf-8"))
    report=revise_master(linked,final,journal=CNS,submission_copy=True,metadata=metadata,artwork_root=artwork_root)
    refreshed=final/"editable-master.docx"
    audit=audit_docx(refreshed)
    citations=[]
    for field in audit["fields"]:
        code=field["instruction"]
        if "ZOTERO_ITEM CSL_CITATION" in code:
            payload=json.loads(code.split("CSL_CITATION",1)[1].strip())
            citations.append((payload["citationItems"][0]["uris"][0].rsplit("/",1)[-1],payload["properties"].get("plainCitation")))
    revised_ole=mtef_hashes(refreshed)
    contents="".join(xml(read_package(refreshed)["word/document.xml"]).itertext())
    result={"original_master_unchanged":original_hash==hashlib.sha256(src.read_bytes()).hexdigest(),
            "math_edit_method":"MathType editor" if args.native_edit_docx else "MathType SDK insertion",
            "math_edit_input_unchanged":edit_source_hash==hashlib.sha256(edit_src.read_bytes()).hexdigest(),
            "dynamic_citations_in_order":citations,"citations_renumbered":citations==[(new_key,"[1]"),(fixture_key(FIRST_DOI),"[2]")],
            "table_revision_retained":"7.50" in contents,
            "first_equation_unchanged":len(revised_ole)==len(old_ole)==2 and revised_ole[0]==old_ole[0],
            "second_native_equation_edit_saved":len(revised_ole)==len(old_ole)==2 and revised_ole[1]!=old_ole[1]
                and (not args.native_edit_docx or revised_ole[1]==edit_source_mtef[1]),
            "math_original":old_ole,"math_revised":revised_ole,
            "word_revised_source_unchanged":report["source_unchanged"],"unresolved_docx_fields":audit["errors"]+audit["pending_markers"]}
    (out/"revision-evidence.json").write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding="utf-8")
    for k in ["original_master_unchanged","math_edit_input_unchanged","citations_renumbered","table_revision_retained","first_equation_unchanged","second_native_equation_edit_saved","word_revised_source_unchanged"]:
        if not result[k]: raise RuntimeError("Revision verification failed: "+k)
    if result["unresolved_docx_fields"]: raise RuntimeError("Unresolved Word fields")
    print(json.dumps(result,ensure_ascii=False,indent=2),flush=True)


if __name__=="__main__": main()
