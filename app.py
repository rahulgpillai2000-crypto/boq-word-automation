import streamlit as st
from generate_business_cases_ui import generate_document

st.title("🏗️ A.D.M Business Case Automation Tool")

st.info("Upload your project Excel and Word template to generate a business case automatically.")

excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])
word_template = st.file_uploader("Upload Word Template", type=["docx"])

if st.button("Generate Document"):

    if excel_file and word_template:

        output_file = generate_document(excel_file, word_template)

        st.success("✅ Document Generated Successfully!")

        with open(output_file, "rb") as f:
            st.download_button(
                label="Download Output File",
                data=f,
                file_name="Business_Case.docx"
            )

    else:
        st.error("Please upload both files")