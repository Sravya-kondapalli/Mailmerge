import streamlit as st
import pandas as pd

st.title("📧 Automated Mail Merge & Certificate Generator")
st.write("Upload your recipient data and trigger the mail merge right from your browser!")

# File uploader widget
uploaded_file = st.file_uploader("Upload recipients file (CSV or Excel)", type=["csv", "xlsx"])

if uploaded_file is not None:
    # Read the file depending on its format
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    
    st.write("### Preview of Recipient Data:")
    st.dataframe(df)
    
    # Run Mail Merge button
    if st.button("Run Mail Merge"):
        # You can add your actual certificate generation logic here later!
        st.success("Mail merge completed successfully!")
