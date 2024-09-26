import time
from io import BytesIO

import pandas as pd
import streamlit as st

# Set page configuration
st.set_page_config(page_title="Workbook Converter", page_icon="🗂", layout="centered")

# Inject custom CSS for the title color (SAP Blue)
st.markdown(
    """
    <style>
    .title {
        color: #0a6ed1;
        font-size: 36px;
        font-weight: bold;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# App title with custom class for styling
st.markdown('<h1 class="title">Workbook Converter</h1>', unsafe_allow_html=True)

# 1. Dropdown selector with one option "Route Map"
object_type = st.selectbox("Select the Object Type:", options=["Route Map"], index=0)

# 2. File uploader for Excel files
uploaded_file = st.file_uploader("Upload Your Excel File:", type=["xlsx"])

# Placeholder for the converted file
converted_file = None

# 3. Convert button
if st.button("Convert"):
    if uploaded_file is not None:
        with st.spinner("Converting..."):
            # Simulate file conversion
            time.sleep(2)  # Replace with actual processing logic

            # For demonstration, we read and write the same file
            df = pd.read_excel(uploaded_file)

            # Save to a BytesIO object
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False)
            processed_data = output.getvalue()

            # Indicate that conversion is complete
            st.success("Conversion complete!")
            converted_file = True
    else:
        st.error("Please upload an Excel file before conversion.")

# 4. Download button for the converted file
if converted_file:
    st.download_button(
        label="Download Converted File",
        data=processed_data,
        file_name="converted_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
