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
    
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice No.{invoice_no}")
    pdf.output(f"Invoices/{filename}.pdf")
    
