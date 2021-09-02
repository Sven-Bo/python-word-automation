import os, sys  # Standard Python Libraries
import xlwings as xw  # pip install xlwings
from docxtpl import DocxTemplate  # pip install docxtpl
import pandas as pd  # pip install pandas
import matplotlib.pyplot as plt  # pip install matplotlib
import win32com.client as win32  # pip install pywin32

# -- Documentation:
# python-docx-template: https://docxtpl.readthedocs.io/en/latest/

# Change path to current working directory
os.chdir(sys.path[0])


def create_barchart(df, barchart_name):
    """Group DataFrame by sub-category, plot barchart, save plot as PNG"""
    top_products = df.groupby(by=df["Sub-Category"]).sum()[["Sales"]]
    top_products = top_products.sort_values(by="Sales")
    plt.rcParams["figure.dpi"] = 300
    plot = top_products.plot(kind="barh")
    fig = plot.get_figure()
    fig.savefig(f"{barchart_name}.png", bbox_inches="tight")
    return None


def convert_to_pdf(doc):
    """Convert given word document to pdf"""
    word = win32.DispatchEx("Word.Application")
    new_name = doc.replace(".docx", r".pdf")
    worddoc = word.Documents.Open(doc)
    worddoc.SaveAs(new_name, FileFormat=17)
    worddoc.Close()
    return None


def main():
    wb = xw.Book.caller()
    sht_panel = wb.sheets["PANEL"]
    sht_sales = wb.sheets["Sales"]
    doc = DocxTemplate("Sales_Report_TEMPLATE.docx")
    # -- Get values from Excel
    context = sht_panel.range("A2").options(dict, expand="table", numbers=int).value
    df = sht_sales.range("A1").options(pd.DataFrame, index=False, expand="table").value

    # -- Create Barchart & Replace Placeholder
    barchart_name = "sales_by_subcategory"
    create_barchart(df, barchart_name)
    doc.replace_pic("Placeholder_1.png", f"{barchart_name}.png")

    # -- Render & Save Word Document
    output_name = f'Sales_Report_{context["month"]}.docx'
    doc.render(context)
    doc.save(output_name)

    # -- Convert to PDF [OPTIONAL]
    path_to_word_document = os.path.join(os.getcwd(), output_name)
    convert_to_pdf(path_to_word_document)

    # -- Show Message Box [OPTIONAL]
    show_msgbox = wb.macro("Module1.ShowMsgBox")
    show_msgbox("DONE!")


if __name__ == "__main__":
    xw.Book("word_automation.xlsm").set_mock_caller()
    main()
