
import json
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE

def create_ppt(json_file_path, output_file):
    # 1. Load JSON Data
    with open(json_file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 2. Render Logic
    prs = Presentation()
    prs.slide_width = Inches(data['slide']['width_inches'])
    prs.slide_height = Inches(data['slide']['height_inches'])
    slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank

    # Background Color
    if data['slide'].get('background_color'):
        bg_rgb = RGBColor.from_string(data['slide']['background_color'])
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = bg_rgb

    for el in data['elements']:
        if el['type'] == 'SHAPE':
            shape_type_enum = getattr(MSO_SHAPE, el['shape_type'], MSO_SHAPE.RECTANGLE)
            
            x = Inches(el['position']['x'])
            y = Inches(el['position']['y'])
            w = Inches(el['position']['w'])
            h = Inches(el['position']['h'])
            
            shape = slide.shapes.add_shape(shape_type_enum, x, y, w, h)
            
            # Style Defaults
            style = el.get('style', {})
            
            # Text
            if el.get('text'):
                shape.text = el['text']
                if shape.has_text_frame:
                    text_frame = shape.text_frame
                    # Handle word wrap based on style, default to True if not specified
                    text_frame.word_wrap = style.get('word_wrap', True)
                    
                    if style.get('alignment') == 'TOP':
                         text_frame.vertical_anchor = MSO_ANCHOR.TOP
                    else:
                         text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                    for p in text_frame.paragraphs:
                        p.alignment = PP_ALIGN.CENTER
                        if style.get('font_size'):
                            p.font.size = Pt(style['font_size'])
                        if style.get('bold'):
                            p.font.bold = True
                        
                        # Fallback for font color if not provided
                        font_color_hex = style.get('font_color', '000000') 
                        p.font.color.rgb = RGBColor.from_string(font_color_hex)
                        
                        p.font.name = 'Microsoft YaHei'
            
            # Fill
            if style.get('fill_color'):
                shape.fill.solid()
                shape.fill.fore_color.rgb = RGBColor.from_string(style['fill_color'])
            else:
                shape.fill.background()

            # Border
            line = shape.line
            
            # Default to a subtle gray border if fill is white and no border specified
            default_border = 'CCCCCC' if style.get('fill_color') == 'FFFFFF' else None
            border_color_hex = style.get('border_color', default_border)
            
            if border_color_hex:
                line.color.rgb = RGBColor.from_string(border_color_hex)
                line.width = Pt(style.get('border_width', 1.0))
            else:
                line.fill.background()
            
            if style.get('is_dashed'):
                line.dash_style = MSO_LINE_DASH_STYLE.DASH
                if not border_color_hex: # Ensure dashed lines are visible even if no border color specified
                    line.color.rgb = RGBColor(0, 0, 0)
                line.width = Pt(style.get('border_width', 1.5))
                
    prs.save(output_file)
    print(f"Generated {output_file}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python ppt_renderer.py <input.json> <output.pptx>")
        sys.exit(1)
    create_ppt(sys.argv[1], sys.argv[2])

