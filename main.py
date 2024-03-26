import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")


def create_header(pdf_form, columns):
    columns = [column.replace('_', ' ').title() for column in columns]

    pdf_form.set_font(family="Times", size=10, style="B")
    pdf_form.set_text_color(80, 80, 80)
    pdf_form.cell(w=30, h=8, txt=f"{columns[0]}", border=1)
    pdf_form.cell(w=70, h=8, txt=f"{columns[1]}", border=1, align="C")
    pdf_form.cell(w=32, h=8, txt=f"{columns[2]}", border=1)
    pdf_form.cell(w=30, h=8, txt=f"{columns[3]}", border=1)
    pdf_form.cell(w=30, h=8, txt=f"{columns[4]}", border=1, ln=1)


for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_num, date = filename.split('-')

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice #: {invoice_num}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    pdf.ln(10)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    create_header(pdf_form=pdf, columns=df.columns)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=f"{row.product_id}", border=1)
        pdf.cell(w=70, h=8, txt=f"{row.product_name}", border=1, align="C")
        pdf.cell(w=32, h=8, txt=f"{row.amount_purchased}", border=1)
        pdf.cell(w=30, h=8, txt=f"{row.price_per_unit}", border=1)
        pdf.cell(w=30, h=8, txt=f"{row.total_price}", border=1, ln=1)

    pdf.output(f"pdfs/{filename}.pdf")

