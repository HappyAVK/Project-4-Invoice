import pandas
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoice/*xlsx")

for fpath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    pdf.set_font(family="Helvetica", style="B", size=24)
    filename = Path(fpath).stem

    invoice = filename.split("-")
    pdf.cell(w=50, h=24, txt=f"Invoice No. {invoice[0]}", ln=1)

    pdf.set_font(family="Helvetica", style="B", size=24)
    pdf.cell(w=50, h=24, txt=f"Date: {invoice[1]}", ln=1)

    df = pandas.read_excel(fpath, sheet_name="Sheet 1")
    columns = list(df.columns)
    col = [item.replace("_", " ") for item in columns]
    for c in col:
        pdf.set_font(family="Helvetica", size=10, style="B")
        pdf.set_text_color(90, 90, 90)
        pdf.cell(w=40, h=8, txt=f"{c}", border=1)
    pdf.cell(w=1, h=8, ln=1)

    for index, row in df.iterrows():

        pdf.set_font(family="Helvetica", size=10)
        pdf.set_text_color(90, 90, 90)
        pdf.cell(w=40, h=8, txt=f"{row['product_id']}", border=1)
        pdf.set_font(family="Helvetica", size=7)
        pdf.cell(w=40, h=8, txt=f"{row['product_name']}", border=1)
        pdf.set_font(family="Helvetica", size=10)
        pdf.cell(w=40, h=8, txt=f"{row['amount_purchased']}", border=1)
        pdf.cell(w=40, h=8, txt=f"{row['price_per_unit']}", border=1)
        pdf.cell(w=40, h=8, txt=f"{row['total_price']}", border=1, ln=1)

    pdf.output(f"pdf/{invoice[0]}.pdf")
