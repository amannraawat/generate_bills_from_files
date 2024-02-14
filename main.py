import glob
import pandas as pd
# PyFPDF is a library for PDF document generation under Python
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('data/*.xlsx')

for filepath in filepaths:
    
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    
    #separating filenames
    filename = Path(filepath).stem
    # print(filename)
    
    #adding invoice name and number on the pdf
    month = filename.split("-")[0]
    month = month.title()
    pdf.set_font(family="Arial", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Bill For {month}", ln=1)
    
    #adding date on the invoice pdf
    date = filename.split("-")[1]
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date:{date}", ln=2)
    
    #reading data from the excel file
    data = pd.read_excel(filepath, sheet_name='Sheet1')
    # print(data)
    
    
    #adding headers(product_id, name..) from the excel files
    columns_name = list(data.columns)
    columns_name = [item.replace("_", " ").title() for item in columns_name]
    # print(columns_name)
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80,80,80)
    pdf.cell(w=30, h=8, txt=columns_name[0], border=1)
    pdf.cell(w=25, h=8, txt=columns_name[1], border=1)
    pdf.cell(w=60, h=8, txt=columns_name[2], border=1)
    pdf.cell(w=30, h=8, txt=columns_name[3], border=1)
    pdf.cell(w=20, h=8, txt=columns_name[4], border=1, ln=1)
    
    
    #for adding rows to the pdf
    for index, row in data.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=25, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["amount_purchased(in kg/quantity)"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price(per unit)"]), border=1)
        pdf.cell(w=20, h=8, txt=str(row["total_price"]), border=1, ln=1)
        
    #calculating total price for each month
    total_price = data['total_price'].sum()
    # print(total_price)
    
    pdf.set_font(family='Times', size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=25, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=20, h=8, txt=str(total_price), border=1, ln=1)
    
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt="", ln=1)
    
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price for {month} is {total_price}.")

    pdf.output(f"Invoices/{filename}.pdf")
    