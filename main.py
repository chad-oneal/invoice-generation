import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')

# for loop to read .xlsx files and create PDFs
for filepath in filepaths:

    # creates the PDF
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    # splits the invoice # and date onto 2 different lines
    filename = Path(filepath).stem
    invoice_number, date = filename.split('-')

    # styling for the invoice number
    pdf.set_font(family='Arial', size=14, style='B')
    pdf.cell(w=50, h=8, txt=f"Invoice # {invoice_number}", ln=1)

    # styling for the date
    pdf.set_font(family='Arial', size=14, style='B')
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)
    # creates space between date and table
    pdf.cell(w=0, h=10, ln=1)

    # creates a dataframe from the excel files
    df = pd.read_excel(filepath, sheet_name='Sheet 1')

    # removes underscore from the headers
    columns = [col.replace('_', ' ').title() for col in df.columns]
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=60, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # creates structure for the cells of the PDF
    for index, row in df.iterrows():
        pdf.set_font(family='Arial', size=10)
        pdf.set_text_color(80, 80 ,80)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=60, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['amount_purchased']), border=1)

        # Format 'price_per_unit' and 'total_price' with dollar sign
        price_per_unit = f"${row['price_per_unit']:.2f}"
        total_price = f"${row['total_price']:.2f}"
        pdf.cell(w=30, h=8, txt=price_per_unit, border=1)
        pdf.cell(w=30, h=8, txt=total_price, border=1, ln=1)

    # creates spacing in the table to generate the sum and output sum message and company image
    pdf.set_font(family='Arial', size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=60, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt='', border=1)

    # Format 'price_per_unit' and 'total_price' with dollar sign
    price_per_unit = f"${row['price_per_unit']:.2f}"
    total_price = f"${row['total_price']:.2f}"
    pdf.cell(w=30, h=8, txt='', border=1)

    # Format Total Price (Sum) to have a dollar sign
    total_sum = f"${df['total_price'].sum():.2f}"
    pdf.set_text_color(42, 195, 69)

    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)
    # creates space between table and the sum message / company logo
    pdf.cell(w=0, h=8, ln=1)
    pdf.set_text_color(255, 0, 69)

    # Add total sum message
    pdf.set_font(family='Arial', style='B', size=10)
    pdf.set_text_color(0, 0, 0)  # Black color for the text
    pdf.cell(w=0, h=8, txt='The total price is: ', border=0, ln=0)
    x = pdf.get_x()
    y = pdf.get_y()

    # Set green color for only the total_sum
    pdf.set_text_color(42, 195, 69)  # Green color for the total sum
    pdf.set_xy(40, y)  # Position the total sum right after the preceding text
    pdf.cell(w=0, h=8, txt=total_sum, border=0, ln=1)
    pdf.cell(w=0, h=2, ln=1)

    # Reset color to black after total sum
    pdf.set_text_color(0, 0, 0)

    # Add company logo
    pdf.image('wicked-robot.png', x=10, w=8)  # Position and size of the image

    # Position for the text after the image
    x_after_image = 6.5 + 13 + 0  # Moves Image x, then y, and creates spacing
    pdf.set_xy(x_after_image, pdf.get_y() - 7.5)  # Set Y a bit higher to align with the image

    # Add company name next to the Logo
    pdf.set_font(family='Arial', style='BI', size=12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=8, txt='- Wicked Robot Distributing', ln=1)

    #  creates the PDF file and stores in the PDFs directory
    pdf.output(f"PDFs/{filename}.pdf")





