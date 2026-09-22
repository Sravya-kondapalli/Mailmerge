import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile

st.title("📧 Automated Mail Merge & Certificate Generator")
st.write("Upload your recipient data and generate personalized certificates right from your browser!")

# File uploaders
uploaded_file = st.file_uploader("Upload recipients file (CSV or Excel)", type=["csv", "xlsx"])
template_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"])

if uploaded_file is not None:
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    
    st.write("### Preview of Recipient Data:")
    st.dataframe(df)
    
    if template_file is not None and st.button("Run Mail Merge & Generate Certificates"):
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for index, row in df.iterrows():
                # Load the template for each row
                doc = Document(template_file)
                
                # Replace placeholders in paragraphs
                for p in doc.paragraphs:
                    for key, val in row.items():
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder in p.text:
                            p.text = p.text.replace(placeholder, str(val))
                
                # Replace placeholders in tables if any exist
                for table in doc.tables:
                    for row_t in table.rows:
                        for cell in row_t.cells:
                            for key, val in row.items():
                                placeholder = f"{{{{{key}}}}}"
                                if placeholder in cell.text:
                                    cell.text = cell.text.replace(placeholder, str(val))
                
                # Save generated certificate to a buffer
                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                
                # Add each certificate to the zip archive
                file_name = f"Certificate_{row.get('Name', index)}.docx"
                zip_file.writestr(file_name, doc_io.read())
        
        zip_buffer.seek(0)
        st.success("Mail merge completed successfully! Download your certificates below:")
        
        st.download_button(
            label="Download All Certificates (ZIP)",
            data=zip_buffer,
            file_name="Generated_Certificates.zip",
            mime="application/zip"
        )
