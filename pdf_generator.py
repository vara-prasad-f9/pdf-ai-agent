from fpdf import FPDF

def generate_pdf_table(df):
    pdf = FPDF(orientation='L')  # Landscape to better fit wide tables
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=10)
    
    pdf.set_font("Arial", 'B', 12)

    # Set column width (divide page width equally)
    page_width = pdf.w - 20  # account for margins
    col_width = page_width / len(df.columns)

    # Header
    for col_name in df.columns:
        pdf.cell(col_width, 10, str(col_name), border=1, align='C')
    pdf.ln()

    # Rows
    pdf.set_font("Arial", size=10)
    for _, row in df.iterrows():
        for item in row:
            pdf.cell(col_width, 10, str(item), border=1)
        pdf.ln()

    output_path = "generated_report.pdf"
    pdf.output(output_path)
    return output_path
