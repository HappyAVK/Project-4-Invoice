import pandas
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoice/*xlsx")

for fpath in filepaths:
    df = pandas.read_excel(fpath, sheet_name="Sheet 1")

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    pdf.set_font(family="Helvetica", style="B", size=24)
    filename = Path(fpath).stem

    invoice = filename.split("-")
    pdf.cell(w=50, h=24, txt=f"Invoice No. {invoice[0]}", ln=1)

    pdf.set_font(family="Helvetica", style="B", size=24)
    pdf.cell(w=50, h=24, txt=f"Date: {invoice[1]}")

    pdf.output(f"pdf/{invoice[0]}.pdf")
