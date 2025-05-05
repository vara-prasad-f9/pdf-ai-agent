import streamlit as st
import pandas as pd
from pdf_generator import generate_pdf_table
from word_to_pdf import convert_word_to_pdf
from excel_agent import summarize_excel_with_ollama

# Streamlit page config
st.set_page_config(page_title="Document to PDF Generator", layout="wide")
st.title("📄 Document to PDF Generator 📝")

# Create tabs for different document types
tab1, tab2 = st.tabs(["Excel to PDF", "Word to PDF"])

# Excel to PDF Tab
with tab1:
    st.subheader("📊 Excel to PDF")
    uploaded_excel = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"], key="excel_uploader")

    if uploaded_excel is not None:
        # Read the Excel file into a DataFrame
        excel_data = pd.read_excel(uploaded_excel)

        # Display the Excel data
        st.subheader("📄 Excel Data Preview")
        st.dataframe(excel_data.head())

        # Button to process Excel data and generate the PDF
        if st.button("Generate PDF from Excel"):
            with st.spinner("Processing Excel..."):
                # Generate the PDF with the Excel data
                pdf_path = generate_pdf_table(excel_data)

                # Provide the PDF download link
                st.success("✅ PDF Generated Successfully!")
                st.download_button(
                    label="Download the PDF",
                    data=open(pdf_path, "rb").read(),
                    file_name="excel_report.pdf",
                    mime="application/pdf"
                )

# Word to PDF Tab
with tab2:
    st.subheader("📝 Word to PDF")
    uploaded_word = st.file_uploader("Upload your Word document", type=["docx"], key="word_uploader")

    if uploaded_word is not None:
        # Button to process Word document and generate the PDF
        if st.button("Generate PDF from Word"):
            with st.spinner("Processing Word document..."):
                # Generate the PDF from Word document
                pdf_path = convert_word_to_pdf(uploaded_word)

                # Provide the PDF download link
                st.success("✅ PDF Generated Successfully!")
                st.download_button(
                    label="Download the PDF",
                    data=open(pdf_path, "rb").read(),
                    file_name="word_report.pdf",
                    mime="application/pdf"
                )
