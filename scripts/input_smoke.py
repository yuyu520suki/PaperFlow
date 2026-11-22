"""Exercise local import adapters with labeled synthetic fixtures."""
import argparse
import hashlib
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from paperflow.ingest import import_document
from paperflow.package import read_package, xml, q
from paperflow.word import private_word


def main():
    p=argparse.ArgumentParser(); p.add_argument("output_dir"); args=p.parse_args()
    out=Path(args.output_dir).resolve(); out.mkdir(parents=True,exist_ok=False)
    samples=out/"source"; samples.mkdir()
    phrase="Paperflow local input fixture"
    from docx import Document
    import pymupdf
    from PIL import Image, ImageDraw, ImageFont
    source_doc=Document(); source_doc.add_heading(phrase,1); source_doc.add_paragraph("Editable table and text.");
    source_doc.add_table(rows=2,cols=2).cell(0,0).text="Case"
    source_doc.save(samples/"fixture.docx")
    (samples/"fixture.md").write_text(f"# {phrase}\n\n$x^2$\n",encoding="utf-8")
    (samples/"fixture.txt").write_text(f"{phrase}\nUse _underscores_ and [square brackets] literally.",encoding="utf-8")
    (samples/"fixture.html").write_text(f"<h1>{phrase}</h1><p>Math: <math xmlns='http://www.w3.org/1998/Math/MathML'><msup><mi>x</mi><mn>2</mn></msup></math></p>",encoding="utf-8")
    (samples/"fixture.tex").write_text("\\documentclass{article}\n\\begin{document}\n\\section{"+phrase+"}\nThe equation is $x^2$.\n\\end{document}",encoding="utf-8")
    (samples/"fixture.rst").write_text(phrase+"\n"+"="*len(phrase)+"\n\nPlain paragraph.",encoding="utf-8")
    # Mixed text-layer and raster PDF exercises per-page routing, not just an image alias.
    font=ImageFont.truetype(r"C:\Windows\Fonts\arial.ttf",48)
    image=Image.new("RGB",(1200,360),"white"); draw=ImageDraw.Draw(image)
    draw.text((45,65),phrase,font=font,fill="black"); draw.text((45,150),"Scanned page sample",font=font,fill="black")
    image.save(samples/"fixture.png")
    pdf=pymupdf.open(); page=pdf.new_page(); page.insert_text((60,80),phrase,fontsize=20)
    page=pdf.new_page(); page.insert_image(pymupdf.Rect(25,50,575,215),filename=str(samples/"fixture.png"))
    pdf.save(samples/"fixture.pdf"); pdf.close()
    results={}
    with private_word() as app:
        for ext, word_format in [("doc",0),("rtf",6),("odt",23)]:
            doc=app.Documents.Open(FileName=str(samples/"fixture.docx"),ReadOnly=True,AddToRecentFiles=False,Visible=False)
            try: doc.SaveAs2(FileName=str(samples/f"fixture.{ext}"),FileFormat=word_format,AddToRecentFiles=False)
            finally: doc.Close(SaveChanges=0)
        for source in sorted(samples.glob("fixture.*")):
            extension=source.suffix[1:]
            target=out/(extension+".docx")
            before=hashlib.sha256(source.read_bytes()).hexdigest()
            report=import_document(source,target,app=app,ocr_language="en-US")
            root=xml(read_package(target)["word/document.xml"])
            text=" ".join(t.text or "" for t in root.iter(q("w:t")))
            report.update(fixture_text_present=phrase in text,source_unchanged=before==hashlib.sha256(source.read_bytes()).hexdigest(),
                          omml_objects=len(list(root.iter(q("m:oMath")))))
            if extension=="pdf": report["ocr_second_page_verified"]="Scanned page sample" in text
            if extension=="txt": report["punctuation_literal_verified"]="_underscores_ and [square brackets]" in text
            results[extension]=report
            print(extension+": "+report["adapter"],flush=True)
    if not all(v["fixture_text_present"] and v["source_unchanged"] for v in results.values()):
        raise RuntimeError("An adapter lost text or changed a source")
    if not results["pdf"]["ocr_second_page_verified"] or not results["txt"]["punctuation_literal_verified"]:
        raise RuntimeError("OCR or literal text gate failed")
    if any(results[fmt]["omml_objects"] != 1 for fmt in ("md", "html", "tex")):
        raise RuntimeError("Structured input did not produce an editable equation")
    (out/"input-validation.json").write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding="utf-8")


if __name__=="__main__": main()
