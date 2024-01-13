import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')

# for loop to read .xlsx files and create PDFs
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')

    # creates the PDF
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    # splits the invoice # and date onto 2 different lines
    filename = Path(filepath).stem
    invoice_number, date = filename.split('-')

    # styling for the invoice number
    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Invoice #{invoice_number}", ln=1)

    # styling for the date
    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Date: {date}")

    #  creates the PDF file and stores in the PDFs directory
    pdf.output(f"PDFs/{filename}.pdf")





