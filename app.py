import streamlit as st
import pandas as pd

st.title("📧 Automated Mail Merge & Certificate Generator")

st.write("Upload your recipient data and trigger the mail merge right from your browser!")

# File uploader for recipients
uploaded_file = st.file_uploader("Upload recipients file (CSV or Excel)", type=["csv", "xlsx"])

if uploaded_file is not None:
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
        
    st.write("Preview of Recipient Data:", df.head())

    if st.button("Run Mail Merge"):
        with st.spinner("Processing emails and generating files..."):
            # Add your processing code here or call your mail_merge functions
            st.success("Mail merge completed successfully!")
