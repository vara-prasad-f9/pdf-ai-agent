import streamlit as st
import pandas as pd
from pdf_generator import generate_pdf_table
from word_to_pdf import convert_word_to_pdf
from excel_agent import summarize_excel_with_ollama
import PyPDF2
import io
import os
from PIL import Image
import tempfile

# Streamlit page config
st.set_page_config(page_title="Document to PDF Generator", layout="wide")
st.title("📄 Document to PDF Generator 📝")

# Create tabs for different document types
tab1, tab2, tab3 = st.tabs(["Excel to PDF", "Word to PDF", "PDF Compressor"])

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

# PDF Compressor Tab
with tab3:
    st.subheader("📦 PDF Compressor")
    st.write("Upload any size PDF file to compress it to a smaller size")
    
    # Advanced compression settings
    st.write("### Advanced Compression Settings")
    col1, col2 = st.columns(2)
    
    with col1:
        image_quality = st.slider("Image Quality (%)", 10, 100, 50, 
                                help="Lower quality means smaller file size but lower image quality")
        dpi = st.slider("DPI Resolution", 72, 300, 150,
                       help="Lower DPI means smaller file size but lower resolution")
    
    with col2:
        compress_text = st.checkbox("Compress Text", value=True,
                                  help="Compress text content (recommended)")
        grayscale = st.checkbox("Convert to Grayscale", value=False,
                              help="Convert images to grayscale for smaller file size")
    
    uploaded_pdf = st.file_uploader("Upload your PDF file", type=["pdf"], key="pdf_uploader")
    
    if uploaded_pdf is not None:
        # Display original file size
        original_size = len(uploaded_pdf.getvalue())
        st.write(f"Original file size: {original_size / (1024*1024):.2f} MB")
        
        if st.button("Compress PDF"):
            with st.spinner("Compressing PDF... This may take a while for large files."):
                try:
                    # Create a temporary file to store the uploaded PDF
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
                        temp_file.write(uploaded_pdf.getvalue())
                        temp_path = temp_file.name
                    
                    # Read the PDF
                    pdf_reader = PyPDF2.PdfReader(temp_path)
                    pdf_writer = PyPDF2.PdfWriter()
                    
                    # Process each page
                    total_pages = len(pdf_reader.pages)
                    progress_bar = st.progress(0)
                    
                    for i, page in enumerate(pdf_reader.pages):
                        # Update progress
                        progress = (i + 1) / total_pages
                        progress_bar.progress(progress)
                        
                        # Process images in the page if any
                        if '/XObject' in page['/Resources']:
                            x_objects = page['/Resources']['/XObject'].get_object()
                            for obj in x_objects:
                                if x_objects[obj]['/Subtype'] == '/Image':
                                    try:
                                        image = x_objects[obj]
                                        if grayscale:
                                            # Convert to grayscale
                                            image['/ColorSpace'] = '/DeviceGray'
                                        
                                        # Reduce image quality
                                        if '/Width' in image and '/Height' in image:
                                            # Scale down large images
                                            if image['/Width'] > 1000 or image['/Height'] > 1000:
                                                scale = min(1000/image['/Width'], 1000/image['/Height'])
                                                image['/Width'] = int(image['/Width'] * scale)
                                                image['/Height'] = int(image['/Height'] * scale)
                                    except:
                                        continue
                        
                        # Add page with compression
                        pdf_writer.add_page(page)
                    
                    # Set compression parameters
                    pdf_writer._compress = True  # Enable compression
                    
                    # Save to a bytes buffer
                    output_buffer = io.BytesIO()
                    pdf_writer.write(output_buffer)
                    output_buffer.seek(0)
                    
                    # Calculate compressed size
                    compressed_size = len(output_buffer.getvalue())
                    compression_ratio = (1 - (compressed_size / original_size)) * 100
                    
                    # Display results
                    st.success(f"✅ PDF Compressed Successfully!")
                    st.write(f"Original size: {original_size / (1024*1024):.2f} MB")
                    st.write(f"Compressed size: {compressed_size / (1024*1024):.2f} MB")
                    st.write(f"Compression ratio: {compression_ratio:.1f}%")
                    
                    # Provide download button
                    st.download_button(
                        label="Download Compressed PDF",
                        data=output_buffer.getvalue(),
                        file_name="compressed_" + uploaded_pdf.name,
                        mime="application/pdf"
                    )
                    
                    # Clean up temporary file
                    os.unlink(temp_path)
                    
                except Exception as e:
                    st.error(f"An error occurred during compression: {str(e)}")
                    if 'temp_path' in locals():
                        os.unlink(temp_path)
