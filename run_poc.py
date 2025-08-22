from pptx import Presentation
from template import TemplateManager

def main():
    tm = TemplateManager("config/templates.toml")
    prs = Presentation()

    # Show available templates
    print("Available templates:")
    for name in tm.get_template_names():
        print("-", name)

    template_name = input("Enter template name: ").strip()

    if template_name not in tm.get_template_names():
        print("❌ Invalid template name.")
        return

    # Collect dynamic content from user
    sample_contents = {}
    placeholders = tm.get_placeholders(template_name)

    print("\n👉 Please provide content:")
    for ph in placeholders:
        if ph["type"] == "text":
            value = input(f"Enter text for {ph['name']}: ").strip()
            sample_contents[ph["name"]] = value
        elif ph["type"] == "image":
            value = input(f"Enter path for image {ph['name']}: ").strip()
            sample_contents[ph["name"]] = value

    # Add slide
    slide_layout = prs.slide_layouts[6]  # blank slide
    slide = prs.slides.add_slide(slide_layout)

    # Apply template
    tm.apply_content(slide, template_name, sample_contents)

    # Save
    prs.save("output.pptx")
    print("✅ Presentation saved as output.pptx")

if __name__ == "__main__":
    main()
