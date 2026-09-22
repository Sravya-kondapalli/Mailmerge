import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile
import os
import glob

st.title("📧 Automated Certificate Generator")
st.write("Upload your recipient data file to generate and download your certificates instantly.")

# Only upload the recipient file
uploaded_file = st.file_uploader("Upload recipients file (CSV or Excel)", type=["csv", "xlsx"])

# Check for template automatically in the repository
template_files = glob.glob("*.docx")
template_path = "certificate_template.docx"

if os.path.exists(template_path):
    active_template = template_path
elif template_files:
    active_template = template_files[0]
else:
    active_template = None

# If template is missing from GitHub, give a quick fallback uploader so you aren't blocked
if active_template is None:
    st.warning("⚠️ 'certificate_template.docx' was not detected in your repository yet. Please upload your template below:")
    template_upload = st.file_uploader("Upload Word Template (.docx)", type=["docx"])
    if template_upload is not None:
        active_template = template_upload

if uploaded_file is not None and active_template is not None:
    # Read recipient file
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    
    # Generate certificates in-memory
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for index, row in df.iterrows():
            if hasattr(active_template, "seek"):
                active_template.seek(0)
                doc = Document(active_template)
            else:
                doc = Document(active_template)
            
            # Replace placeholders in paragraphs
            for p in doc.paragraphs:
                for key, val in row.items():
                    placeholder = f"{{{{{key}}}}}"
                    if placeholder in p.text:
                        p.text = p.text.replace(placeholder, str(val))
            
            # Replace placeholders in tables
            for table in doc.tables:
                for row_t in table.rows:
                    for cell in row_t.cells:
                        for key, val in row.items():
                            placeholder = f"{{{{{key}}}}}"
                            if placeholder in cell.text:
                                cell.text = cell.text.replace(placeholder, str(val))
            
            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)
            
            file_name = f"Certificate_{row.get('Name', index)}.docx"
            zip_file.writestr(file_name, doc_io.read())
    
    zip_buffer.seek(0)
    
    st.success("Certificates generated successfully!")
    
    # Direct download button
    st.download_button(
        label="Download All Certificates (ZIP)",
        data=zip_buffer,
        file_name="Generated_Certificates.zip",
        mime="application/zip"
    )
