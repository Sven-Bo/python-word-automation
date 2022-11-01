from pathlib import Path  # core library

from docx.shared import Cm, Emu, Inches, Mm  # pip install python-docx
from docxtpl import DocxTemplate, InlineImage  # pip install docxtpl

# Path settings
current_dir = Path(__file__).parent
template_path = current_dir / "Template.docx"
output_path = current_dir / "Template_Rendered.docx"
img_placeholder1_path = current_dir / "Placeholders" / "Placeholder_1.png"
img_placeholder2_path = current_dir / "Placeholders" / "Placeholder_2.png"


doc = DocxTemplate(template_path)
placeholder_1 = InlineImage(doc, str(img_placeholder1_path), Cm(5))
placeholder_2 = InlineImage(doc, str(img_placeholder2_path), Cm(5))
context = {
    "name": "Sven",
    "placeholder_1": placeholder_1,
    "placeholder_2": placeholder_2,
}

doc.render(context)
doc.save(output_path)
