import streamlit as st
import pdfplumber
import pandas as pd
import io
import zipfile

st.set_page_config(page_title="ZIP to Excel ZIP Converter", page_icon="📦", layout="centered")

st.title("📦 ZIP of PDFs to Excel ZIP Converter")
st.write("PDF பைல்கள் அடங்கிய ஒரு ZIP பைலை அப்லோட் செய்யவும். அவை அனைத்தும் எக்செல் பைல்களாக மாற்றப்பட்டு, ஒரே ZIP பைலாக டவுன்லோட் செய்யக் கிடைக்கும்!")

# ZIP பைலை அப்லோட் செய்ய அனுமதிக்கிறது
uploaded_file = st.file_uploader("Choose a ZIP file containing PDFs...", type=["zip"], accept_multiple_files=False)

if uploaded_file:
    st.info("ZIP file uploaded successfully. Processing...")
    
    if st.button("Extract, Convert & Pack into ZIP"):
        # அவுட்புட் ZIP பைலை மெமரியில் உருவாக்குவதற்கான பப்பர்
        output_zip_buffer = io.BytesIO()
        successful_conversions = 0
        
        try:
            # அப்லோட் செய்யப்பட்ட input ZIP பைலை ஓபன் செய்கிறோம்
            with zipfile.ZipFile(uploaded_file, "r") as input_zip:
                # ZIP-க்குள் இருக்கும் அனைத்து பைல்களின் பெயர்கள்
                file_list = input_zip.namelist()
                
                # .pdf என்று முடியும் பைல்களை மட்டும் வடிகட்டுகிறோம்
                pdf_files = [f for f in file_list if f.lower().endswith('.pdf')]
                
                if not pdf_files:
                    st.error("No PDF files found inside the uploaded ZIP!")
                else:
                    # புதிய அவுட்புட் ZIP பைலை எழுதத் தொடங்குகிறோம்
                    with zipfile.ZipFile(output_zip_buffer, "w", zipfile.ZIP_DEFLATED) as output_zip:
                        
                        for pdf_name in pdf_files:
                            all_tables_df = []
                            
                            # ZIP-க்குள் இருக்கும் PDF பைலை மெமரியிலேயே படிக்கிறோம்
                            pdf_data = input_zip.read(pdf_name)
                            pdf_stream = io.BytesIO(pdf_data)
                            
                            # PDF-ல் இருந்து டேபிள்களை எடுக்கிறது
                            with pdfplumber.open(pdf_stream) as pdf:
                                for page in pdf.pages:
                                    tables = page.extract_tables()
                                    for table in tables:
                                        df = pd.DataFrame(table)
                                        if not df.empty:
                                            df.columns = df.iloc[0]
                                            df = df[1:]
                                            all_tables_df.append(df)
                            
                            # டேபிள்கள் இருந்தால் மட்டுமே எக்செல் பைலாக மாற்றும்
                            if all_tables_df:
                                combined_df = pd.concat(all_tables_df, ignore_index=True)
                                
                                # எக்செல் பைலை மெமரியில் உருவாக்குகிறது
                                excel_buffer = io.BytesIO()
                                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                    combined_df.to_excel(writer, sheet_name='Data', index=False)
                                
                                excel_data = excel_buffer.getvalue()
                                
                                # பழைய PDF பெயரின் நீட்சியை மாற்றி எக்செல் பெயராக மாற்றுகிறது 
                                # (e.g., folder/sample.pdf -> folder/sample.xlsx)
                                excel_filename = pdf_name.rsplit('.', 1)[0] + ".xlsx"
                                
                                # எக்செல் பைலை புதிய ZIP-க்குள் சேர்க்கிறது
                                output_zip.writestr(excel_filename, excel_data)
                                successful_conversions += 1
            
            # ஏதேனும் பைல்கள் கன்வெர்ட் ஆகியிருந்தால் டவுன்லோட் பட்டனை காட்டும்
            if successful_conversions > 0:
                st.success(f"Successfully converted {successful_conversions} PDF(s) into Excel!")
                
                final_zip_data = output_zip_buffer.getvalue()
                
                st.download_button(
                    label="📥 Download Converted Excel Files (ZIP)",
                    data=final_zip_data,
                    file_name="converted_excel_files.zip",
                    mime="application/zip"
                )
            else:
                st.error("No tables found in any of the PDFs inside the ZIP!")
                
        except zipfile.BadZipFile:
            st.error("The uploaded file is not a valid ZIP file.")