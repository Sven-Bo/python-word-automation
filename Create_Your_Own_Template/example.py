import os, sys  # Standard Python Libraries
from docxtpl import DocxTemplate, InlineImage  # pip install docxtpl
from docx.shared import Cm, Inches, Mm, Emu  # pip install python-docx

# Change path to current working directory
os.chdir(sys.path[0])

doc = DocxTemplate("Template.docx")
placeholder_1 = InlineImage(doc, "Placeholders/Placeholder_1.png", Cm(5))
placeholder_2 = InlineImage(doc, "Placeholders/Placeholder_2.png", Cm(5))
context = {
    "name": "Sven",
    "placeholder_1": placeholder_1,
    "placeholder_2": placeholder_2,
}

doc.render(context)
doc.save("Template_Rendered.docx")
