import sys
from pptx import Presentation

sys.stdout.reconfigure(encoding='utf-8')

TEMPLATE = r'D:\BETEKK\opencode-config-template\opencode_app\scripts\templates\template.pptx'

prs = Presentation(TEMPLATE)
print(f"Initial slides: {len(prs.slides)}")

def delete_slide(prs, index):
    rId = prs.slides._sldIdLst[index].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    prs.part.drop_rel(rId)
    sldId = prs.slides._sldIdLst[index]
    prs.slides._sldIdLst.remove(sldId)

# Delete slides 10, 8, 5 (from highest to lowest)
delete_slide(prs, 10)
print(f"After delete 10: {len(prs.slides)}")

delete_slide(prs, 8)
print(f"After delete 8: {len(prs.slides)}")

delete_slide(prs, 5)
print(f"After delete 5: {len(prs.slides)}")

for i, slide in enumerate(prs.slides):
    title = ""
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                txt = "".join(r.text for r in para.runs)
                if txt.strip():
                    title = txt.strip()[:60]
                    break
        if title:
            break
    print(f"  Slide {i}: {title}")
