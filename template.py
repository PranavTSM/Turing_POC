import toml
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

class TemplateManager:
    def __init__(self, toml_file="config/templates.toml"):
        self.templates = toml.load(toml_file)

    def get_template_names(self):
        return list(self.templates.keys())

    def get_placeholders(self, template_name):
        return self.templates.get(template_name, {}).get("placeholders", [])

    def apply_content(self, slide, template_name, contents):
        placeholders = self.get_placeholders(template_name)

        for ph in placeholders:
            name = ph["name"]
            ph_type = ph["type"]
            x, y, w, h = Inches(ph["x"]), Inches(ph["y"]), Inches(ph["w"]), Inches(ph["h"])

            if ph_type == "text":
                textbox = slide.shapes.add_textbox(x, y, w, h)
                tf = textbox.text_frame
                tf.text = contents.get(name, "")
                p = tf.paragraphs[0]

                # Alignment
                align = ph.get("align", "left").lower()
                if align == "center":
                    p.alignment = PP_ALIGN.CENTER
                elif align == "right":
                    p.alignment = PP_ALIGN.RIGHT
                else:
                    p.alignment = PP_ALIGN.LEFT

                # Font size
                if p.runs:
                    run = p.runs[0]
                    run.font.size = Pt(18)

            elif ph_type == "image":
                img_path = contents.get(name)
                if img_path:
                    try:
                        slide.shapes.add_picture(img_path, x, y, w, h)
                    except FileNotFoundError:
                        print(f"⚠️ Image not found: {img_path}")
