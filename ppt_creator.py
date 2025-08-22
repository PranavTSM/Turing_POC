# ppt_creator.py

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

def create_ppt_from_template(output_file, template, user_inputs):
    prs = Presentation()
    slide_layout = prs.slide_layouts[6]  # Blank slide
    slide = prs.slides.add_slide(slide_layout)

    for placeholder in template["placeholders"]:
        name = placeholder["name"]
        x, y, cx, cy = placeholder["x"], placeholder["y"], placeholder["cx"], placeholder["cy"]
        alignment = placeholder.get("alignment", "left")

        txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(cx), Inches(cy))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = user_inputs.get(name, f"<{name}>")
        p.font.size = Pt(24)

        if alignment == "center":
            p.alignment = PP_ALIGN.CENTER
        elif alignment == "right":
            p.alignment = PP_ALIGN.RIGHT
        else:
            p.alignment = PP_ALIGN.LEFT

    prs.save(output_file)
    print(f"✅ PPT created successfully: {output_file}")
