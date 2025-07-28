import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoces/*.xlsx");

for filepath in filepaths:
    # Extract invoces data
    df = pd.read_excel(filepath, sheet_name="Sheet 1");
    pdf = FPDF(orientation="P", unit="mm", format="A4");
    pdf.add_page();

    filename = Path(filepath).stem;
    invoice_nm = filename.split("-")[0];

    #Creating pdf
    pdf.set_font(family="Times", size=16, style="B");
    pdf.cell(w=50, h=8, text=f"invoice nr.{invoice_nm}");
    pdf.output(f"PDFS/{filename}.pdf");