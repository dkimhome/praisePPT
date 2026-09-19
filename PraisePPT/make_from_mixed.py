from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import re

def make_ppt_from_mixed(mixed_file, template_file, output_file):
    prs = Presentation(template_file)
    with open(mixed_file, 'r', encoding='utf-8') as f:
        content = f.read()
        
    slides_text = content.split('\n\n')
    
    for slide_content in slides_text:
        if not slide_content.strip(): 
            continue
            
        slide = prs.slides.add_slide(prs.slide_layouts[0]) 
        for shape in slide.shapes:
            sp = shape.element
            sp.getparent().remove(sp)
            
        left = Inches(0)         
        # 1. 위쪽 여백 조정: 0.8 -> 0.5 (더 위로 바짝 붙습니다)
        top = Inches(0.5)        
        width = prs.slide_width  
        height = prs.slide_height - Inches(1)
        
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP 
        
        lines = slide_content.strip().split('\n')
        
        for i, line in enumerate(lines):
            if not line.strip():
                continue
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
                
            p.text = line.strip()
            p.alignment = PP_ALIGN.CENTER
            
            is_korean = bool(re.search(r'[가-힣]', line))
            font_size = 45 if is_korean else 40
            font_name = 'Malgun Gothic' if is_korean else 'Arial'
            
            # 2. 색상 조정: 영어를 더 화사하고 밝은 오렌지색(255, 220, 60)으로 변경
            font_color = RGBColor(255, 255, 255) if is_korean else RGBColor(255, 220, 60)
            
            for run in p.runs:
                run.font.size = Pt(font_size)
                run.font.name = font_name
                run.font.color.rgb = font_color
                
    prs.save(output_file)
    print("🎉 완성된 텍스트 파일로 PPT 구성을 완료했습니다! 폴더를 확인해 보세요!")

# 파일 이름은 already_mixed_lyrics.txt 로 유지했습니다.
make_ppt_from_mixed('already_mixed_lyrics.txt', 'template.pptx', 'final_praise_direct.pptx')