"""Create local synthetic OCR fixtures and validate recognized text and boxes."""
import argparse
import hashlib
import json
from pathlib import Path
import sys

from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from paperflow.ocr import ocr_status, recognize_image


def main():
    parser=argparse.ArgumentParser(); parser.add_argument("output_dir"); args=parser.parse_args()
    output=Path(args.output_dir).resolve(); output.mkdir(parents=True,exist_ok=False)
    fixtures=[
        ("english","en-US",r"C:\Windows\Fonts\arial.ttf",["Editable manuscript workflow","Equation x = 1 needs review.","Table Case A 1.25"]),
        ("chinese","zh-Hans-CN",r"C:\Windows\Fonts\simsun.ttc",["论文格式转换测试","公式和表格需要人工复核","本地识别 不上传图片"]),
    ]
    report={"fixture_only":True,"status":ocr_status(),"cases":[],"scientific_formula_or_table_recovery_proven":False}
    for name,language,font_file,lines in fixtures:
        image=Image.new("RGB",(1600,420),"white")
        draw=ImageDraw.Draw(image); font=ImageFont.truetype(font_file,50)
        for n,line in enumerate(lines): draw.text((70,60+n*100),line,font=font,fill="black")
        source=output/(name+".png"); image.save(source)
        result=recognize_image(source,output/(name+".json"),language=language)
        expected="\n".join(lines)
        passed="".join(expected.split())=="".join(result["normalized_text"].split())
        boxes=all(line["bbox"] and all(word["bbox"] for word in line["words"]) for line in result["lines"])
        report["cases"].append({"name":name,"language":result["language"],"expected":expected,
            "actual":result["normalized_text"],"text_match_ignoring_whitespace":passed,
            "line_count":len(result["lines"]),"boxes_present":boxes,
            "source_sha256":hashlib.sha256(source.read_bytes()).hexdigest(),"requires_review":result["requires_review"]})
        if not passed or not boxes: raise RuntimeError("OCR fixture verification failed: "+name)
    (output/"summary.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
    print(json.dumps({"output_dir":str(output),"cases_passed":len(report["cases"])}))


if __name__=="__main__": main()
