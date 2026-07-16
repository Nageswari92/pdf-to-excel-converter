import streamlit as st
import pdfplumber
import pandas as pd
import io
import zipfile

st.set_page_config(page_title="PDF to Excel Converter", page_icon="📦", layout="centered")

st.title("📦 PDF Tables to Excel Converter")
st.write("Upload a single PDF file or a ZIP archive containing multiple PDFs. The app will extract tables and convert them to Excel format.")

# Allows uploading either a single PDF file or a single ZIP file containing PDFs
uploaded_file = st.file_uploader(
    "Choose a PDF file or a ZIP file containing PDFs...", 
    type=["pdf", "zip"], 
    accept_multiple_files=False
)

# Helper function to extract tables from a PDF stream
def extract_tables_from_pdf(pdf_stream):
    all_tables_df = []
    with pdfplumber.open(pdf_stream) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                df = pd.DataFrame(table)
                if not df.empty:
                    df.columns = df.iloc[0]
                    df = df[1:]
                    all_tables_df.append(df)
    if all_tables_df:
        return pd.concat(all_tables_df, ignore_index=True)
    return None

if uploaded_file:
    st.info("File uploaded successfully. Processing...")
    file_name = uploaded_file.name.lower()
    
    if st.button("Extract & Convert"):
        
        # Scenario 1: User uploaded a single PDF file
        if file_name.endswith('.pdf'):
            try:
                pdf_stream = io.BytesIO(uploaded_file.read())
                combined_df = extract_tables_from_pdf(pdf_stream)
                
                if combined_df is not None:
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        combined_df.to_excel(writer, sheet_name='Data', index=False)
                    
                    st.success("Successfully converted PDF into Excel!")
                    st.download_button(
                        label="📥 Download Excel File",
                        data=excel_buffer.getvalue(),
                        file_name=uploaded_file.name.rsplit('.', 1)[0] + ".xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("No tables found in the uploaded PDF file!")
            except Exception as e:
                st.error(f"An error occurred while processing the PDF: {e}")
                
        # Scenario 2: User uploaded a ZIP file containing PDFs
        elif file_name.endswith('.zip'):
            output_zip_buffer = io.BytesIO()
            successful_conversions = 0
            
            try:
                with zipfile.ZipFile(uploaded_file, "r") as input_zip:
                    file_list = input_zip.namelist()
                    pdf_files = [f for f in file_list if f.lower().endswith('.pdf')]
                    
                    if not pdf_files:
                        st.error("No PDF files found inside the uploaded ZIP archive!")
                    else:
                        with zipfile.ZipFile(output_zip_buffer, "w", zipfile.ZIP_DEFLATED) as output_zip:
                            for pdf_name in pdf_files:
                                pdf_data = input_zip.read(pdf_name)
                                pdf_stream = io.BytesIO(pdf_data)
                                combined_df = extract_tables_from_pdf(pdf_stream)
                                
                                if combined_df is not None:
                                    excel_buffer = io.BytesIO()
                                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                        combined_df.to_excel(writer, sheet_name='Data', index=False)
                                    
                                    excel_data = excel_buffer.getvalue()
                                    excel_filename = pdf_name.rsplit('.', 1)[0] + ".xlsx"
                                    output_zip.writestr(excel_filename, excel_data)
                                    successful_conversions += 1
                                    
                if successful_conversions > 0:
                    st.success(f"Successfully converted {successful_conversions} PDF(s) into Excel inside ZIP!")
                    st.download_button(
                        label="📥 Download Converted Excel Files (ZIP)",
                        data=output_zip_buffer.getvalue(),
                        file_name="converted_excel_files.zip",
                        mime="application/zip"
                    )
                else:
                    st.error("No tables found in any of the PDFs inside the ZIP archive!")
                    
            except zipfile.BadZipFile:
                st.error("The uploaded file is not a valid ZIP file.")
