from pathlib import Path  # core library

import matplotlib.pyplot as plt  # pip install matplotlib
import pandas as pd  # pip install pandas
import win32com.client as win32  # pip install pywin32
import xlwings as xw  # pip install xlwings
from docxtpl import DocxTemplate  # pip install docxtpl

# -- Documentation:
# python-docx-template: https://docxtpl.readthedocs.io/en/latest/


def create_barchart(df, barchart_output):
    """Group DataFrame by sub-category, plot barchart, save plot as PNG"""
    top_products = df.groupby(by=df["Sub-Category"]).sum()[["Sales"]]
    top_products = top_products.sort_values(by="Sales")
    plt.rcParams["figure.dpi"] = 300
    plot = top_products.plot(kind="barh")
    fig = plot.get_figure()
    fig.savefig(barchart_output, bbox_inches="tight")
    return None


def convert_to_pdf(doc):
    """Convert given word document to pdf"""
    word = win32.DispatchEx("Word.Application")
    new_name = doc.replace(".docx", ".pdf")
    worddoc = word.Documents.Open(doc)
    worddoc.SaveAs(new_name, FileFormat=17)
    worddoc.Close()
    return None


def main():
    # Path settings
    current_dir = Path(__file__).parent
    template_path = current_dir / "Sales_Report_TEMPLATE.docx"

    # Conection to Excel
    wb = xw.Book.caller()
    sht_panel = wb.sheets["PANEL"]
    sht_sales = wb.sheets["Sales"]
    context = sht_panel.range("A2").options(dict, expand="table", numbers=int).value
    df = sht_sales.range("A1").options(pd.DataFrame, index=False, expand="table").value

    # Initialize template
    doc = DocxTemplate(str(template_path))

    # -- Create Barchart & Replace Placeholder
    barchart_name = "sales_by_subcategory"
    barchart_output = current_dir / f"{barchart_name}.png"
    create_barchart(df, barchart_output)
    doc.replace_pic("Placeholder_1.png", barchart_output)

    # -- Render & Save Word Document
    output_name = current_dir / f'Sales_Report_{context["month"]}.docx'
    doc.render(context)
    doc.save(output_name)

    # -- Convert to PDF [OPTIONAL]
    convert_to_pdf(str(output_name))

    # -- Show Message Box [OPTIONAL]
    show_msgbox = wb.macro("Module1.ShowMsgBox")
    show_msgbox("DONE!")


if __name__ == "__main__":
    xw.Book("word_automation.xlsm").set_mock_caller()
    main()
