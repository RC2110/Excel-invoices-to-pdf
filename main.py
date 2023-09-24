import pandas as pd
import glob
from pathlib import Path
from fpdf import FPDF

files = glob.glob("Invoices_excel/*xlsx")

for filepaths in files:
    data = pd.read_excel(filepaths, sheet_name="Sheet 1")
    afile = Path(filepaths).stem
    invoice_details = afile.split('-')
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_text_color(100, 100, 100)
    pdf.set_font(family='Times', style='B', size=20)
    pdf.cell(w=25, h=2, txt=f"invoice number:{invoice_details[0]}", align='L', ln=1)
    pdf.output(f"pdfs/{afile}.pdf")


# data = pd.read_excel()