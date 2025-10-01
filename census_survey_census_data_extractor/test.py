import camelot
import pandas as pd

def pdf_to_excel(pdf_path, excel_path, pages="all"):
    """
    Extract tables from a PDF and save them into an Excel file.
    
    :param pdf_path: Path to input PDF
    :param excel_path: Path to output Excel
    :param pages: Pages to parse (e.g. "1,2,3" or "all")
    """
    # Extract tables
    tables = camelot.read_pdf(pdf_path, pages=pages)

    # Save each table into a sheet
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        for i, table in enumerate(tables):
            df = table.df
            sheet_name = f"Table_{i+1}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"✅ Extracted {len(tables)} tables into {excel_path}")

# Example usage
pdf_to_excel("kundgolHescomdata.pdf", "output.xlsx", pages="all")
