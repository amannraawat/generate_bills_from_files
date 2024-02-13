import glob
import pandas as pd
# PyFPDF is a library for PDF document generation under Python
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('data/*.xlsx')

for filepath in filepaths:
    data = pd.read_excel(filepath, sheet_name='Sheet1')
    # print(data)
    
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    # print(filename)
    
    invoice_no = filename.split("-")[0]
    pdf.set_font(family="Arial", size=16, style="I")
    pdf.cell(w=50, h=8, txt=f"Invoice No.{invoice_no}", ln=1)
    
    date = filename.split("-")[1]
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date:{date}", ln=2)
    
    #for adding rows to the pdf
    data = pd.read_excel(filepath, sheet_name='Sheet1')
    # print(data)
    for index, row in data.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=50, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased(in kg/quantity)"]), border=1)
        pdf.cell(w=20, h=8, txt=str(row["price(per unit)"]), border=1)
        pdf.cell(w=20, h=8, txt=str(row["total_price"]), border=1, ln=1)

    pdf.output(f"Invoices/{filename}.pdf")
    