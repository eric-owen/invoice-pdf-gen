import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


def create_invoice_pdf(filepath):
    pdf = FPDF(orientation="P", unit="mm", format="A4")

    pdf.add_page()

    filename = Path(filepath).stem
    inv_num, inv_date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice# {inv_num}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Invoice Date: {inv_date} ", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Convert column names to title case
    columns = df.columns.str.replace('_', ' ').str.title()

    total_sum = df['total_price'].sum()
    pdf.cell(w=30, h=8, txt=f'Invoice Total: {total_sum:.2f}', ln=1)

    # Print column headers
    pdf.set_font(family="Times", size=10, style='B')
    for col in columns:
        pdf.cell(w=30, h=8, txt=col, border=1)
    pdf.ln()

    # Print rows
    pdf.set_font(family="Times", size=10)
    for _, row in df.iterrows():
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=f'{row["price_per_unit"]:.2f}', border=1)
        pdf.cell(w=30, h=8, txt=f'{row["total_price"]:.2f}', border=1, ln=1)

    pdf.output(f"PDFs/{inv_num}.pdf")


if __name__ == "__main__":
    filepaths = glob.glob("invoices/*.xlsx")

    for filepath in filepaths:
        create_invoice_pdf(filepath)
