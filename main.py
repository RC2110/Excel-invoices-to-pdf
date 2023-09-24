import pandas as pd
import glob
from pathlib import Path
from fpdf import FPDF

files = glob.glob("Invoices_excel/*xlsx")

for filepaths in files:
    data = pd.read_excel(filepaths, sheet_name="Sheet 1")
    afile = Path(filepaths).stem
    invoice_details = afile.split('-')
    raw_date=invoice_details[1].split('.')
    date=f"{raw_date[2]}-{raw_date[1]}-{raw_date[0]}"

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_text_color(100, 100, 100)
    pdf.set_font(family='Times', style='B', size=15)
    pdf.cell(w=110, h=8, txt=f"Invoice number:{invoice_details[0]}", align='L')
    pdf.cell(w=0, h=8, txt=f"Date:{date}", align='L',
             ln=1)

    headings= data.columns
    headings = [i.replace('_', ' ') for i in headings]

    pdf.set_font(family='Times', style= 'B', size=9)
    pdf.cell(w=25, h=7, txt=headings[0], align='M', border=1)
    pdf.cell(w=35, h=7, txt=headings[1], align='M', border=1)
    pdf.cell(w=28, h=7, txt=headings[2], align='M',
                 border=1)
    pdf.cell(w=30, h=7, txt=headings[3], align='M',
                 border=1)
    pdf.cell(w=30, h=7, txt=headings[4], align='M',
                 border=1, ln=1)

    for index, rows in data.iterrows():
        pdf.set_font(family='Times', size=9)
        pdf.cell(w=25, h=7, txt=str(rows['product_id']), align='M', border=1)
        pdf.cell(w=35, h=7, txt=str(rows['product_name']), align='M', border=1)
        pdf.cell(w=28, h=7, txt=str(rows['amount_purchased']), align='M',
                 border=1)
        pdf.cell(w=30, h=7, txt=str(rows['price_per_unit']), align='M',
                 border=1)
        pdf.cell(w=30, h=7, txt=str(rows['total_price']), align='M',
                 border=1, ln=1)

    total_price = data['total_price'].sum()

    pdf.set_font(family='Times', size=9)
    pdf.cell(w=25, h=7, txt='', align='M', border=1)
    pdf.cell(w=35, h=7, txt='', align='M', border=1)
    pdf.cell(w=28, h=7, txt='', align='M',
               border=1)
    pdf.cell(w=30, h=7, txt='', align='M',
                 border=1)
    pdf.cell(w=30, h=7, txt=f"{total_price}", align='M',
                 border=1, ln=1)

    pdf.set_font(family='Times', style='I', size=12)
    pdf.cell(w=0,h=9, txt=f"The total price is {total_price}", ln=1)

    pdf.set_font(family='Times', style='I', size=9)
    pdf.cell(w=20, h=9, txt="The Ai Place")
    pdf.image("AI.png")



    pdf.output(f"pdfs/{afile}.pdf")


# data = pd.read_excel()