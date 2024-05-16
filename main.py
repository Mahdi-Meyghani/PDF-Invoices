import glob
import pandas as pd
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob("invoices/*.xlsx")
table_width = [30, 70, 30, 30, 30]

for filepath in filepaths:
    pdf = FPDF(orientation="L", unit="mm", format="A4")

    filename = Path(filepath).stem
    invoice_nr, date_nr = filename.split("-")

    pdf.add_page()
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date {date_nr}")
    pdf.ln(20)

    df = pd.read_excel(filepath)
    columns = df.columns
    columns = [title.replace("_", " ").title() for title in columns]

    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(80, 80, 80)

    pdf.cell(w=30, h=16, txt=columns[0], border=1)
    pdf.cell(w=70, h=16, txt=columns[1], border=1)
    pdf.cell(w=40, h=16, txt=columns[2], border=1)
    pdf.cell(w=30, h=16, txt=columns[3], border=1)
    pdf.cell(w=30, h=16, txt=columns[4], border=1, ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=16)
        pdf.set_text_color(80, 80, 80)

        pdf.cell(w=30, h=16, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=16, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=16, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=16, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=16, txt=str(row["total_price"]), border=1, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")
