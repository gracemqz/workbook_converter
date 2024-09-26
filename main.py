from io import BytesIO

import pandas as pd
import streamlit as st

# Constants
INITIAL_ROWS = ["Route Map Name", "Route Map Description"]
STEP_FIELDS = [
    "Step Name",
    "Step Stage",
    "Step Name",
    "Step Description",
    "Step Type",
    "Type (Role)",
    "Step Introduction & Mouseover text",
    "Step Name After Completion",
    "Step Mode",
    "Exit Button Text",
    "Step Exit Text (next to Exit button)",
    "Previous Step Exit Button Text (Modify Stage Only)",
    "Previous Step Exit Text (Modify Stage Only)",
    "Entry User",
    "Exit User",
    "Start Date",
    "Exit Date",
    "Enforce Start Date",
    "Automatic send on due date",
    "Iterative Button Text (Iterative Step Type Only)",
    "Reject Button Mouseover Text (Signature Stage Only)",
    "Step Exit Reminder (Modify and Signature Only)",
    "Step Exit Reminder Text (Modify and Signature Only)",
]
FIELD_MAPPING = {
    "Route Map Name": "Route Map Name",
    "Route Map Description": "Description",
    "Step Name": "Step Name",
    "Step Stage": "Modify Stage",
    "Step Description": "Step Description",
    "Step Type": "Step Type",
    "Type (Role)": "Roles",
    "Step Introduction & Mouseover text": "Step Introduction and Mouseover",
    "Step Name After Completion": "Step Name After Completion",
    "Step Mode": "Step Mode",
    "Exit Button Text": "Exit Button Text",
    "Step Exit Text (next to Exit button)": "Step Exit Text",
    "Previous Step Exit Button Text (Modify Stage Only)": "Previous Step Exit Button Text",
    "Previous Step Exit Text (Modify Stage Only)": "Previous Step Exit Text",
    "Entry User": "Entry User",
    "Exit User": "Exit User",
    "Start Date": "Start Date",
    "Exit Date": "Exit Date",
    "Enforce Start Date": "Enforce Start Date",
    "Automatic send on due date": "Automatic send on due date",
    "Iterative Button Text (Iterative Step Type Only)": "Iterative Button Text",
    "Reject Button Mouseover Text (Signature Stage Only)": "Reject Button Mouseover Text",
    "Step Exit Reminder (Modify and Signature Only)": "Step Exit Reminder",
    "Step Exit Reminder Text (Modify and Signature Only)": "Step Exit Reminder Text",
}


def process_csv_file(file):
    # Read the CSV file into a DataFrame
    df_pcm_generated = pd.read_csv(file)

    # Filter out "Language pack" and get unique languages
    df_pcm_generated = df_pcm_generated[df_pcm_generated["Language"] != "Language pack"]
    unique_languages = df_pcm_generated["Language"].unique()

    # Determine the maximum number of steps across all languages
    max_steps = df_pcm_generated.groupby("Language").size().max()

    # Build the "Row Names" column
    row_names = INITIAL_ROWS + STEP_FIELDS * max_steps

    # Create the "Ideal Worksheet" DataFrame with "Route Map" as the header
    df_ideal_worksheet = pd.DataFrame()
    df_ideal_worksheet["Route Map"] = row_names  # First column header is "Route Map"

    # Add language columns and populate data
    for lang in unique_languages:
        # Filter data by the selected language
        lang_df = df_pcm_generated[df_pcm_generated["Language"] == lang].reset_index(
            drop=True
        )
        num_steps = len(lang_df)

        # Initialize the column with empty strings
        df_ideal_worksheet[lang] = ""

        # Populate the initial rows
        df_ideal_worksheet.loc[0, lang] = lang_df.at[
            0, "Route Map Name"
        ]  # Route Map Name
        df_ideal_worksheet.loc[1, lang] = lang_df.at[
            0, "Description"
        ]  # Route Map Description

        # Populate the step fields
        for step_index in range(num_steps):
            start_row = len(INITIAL_ROWS) + step_index * len(STEP_FIELDS)
            row_data = lang_df.loc[step_index]
            for i, row_name in enumerate(STEP_FIELDS):
                mapped_field = FIELD_MAPPING.get(row_name, "")
                value = row_data.get(mapped_field, "")
                df_ideal_worksheet.loc[start_row + i, lang] = value

    # Save the modified DataFrame to a CSV file in memory
    buffer = BytesIO()
    df_ideal_worksheet.to_csv(buffer, index=False)
    buffer.seek(0)
    return buffer, df_ideal_worksheet


def main():
    # Set the page configuration
    st.set_page_config(page_title="Workbook Converter", page_icon="🗂")

    # Display the title in SAP blue
    st.markdown(
        '<h1 style="color:#008BBF;">Workbook Converter</h1>',
        unsafe_allow_html=True,
    )

    # Add the dropdown selector for object type
    object_type = st.selectbox("Select the object type:", ["Route map"])

    # File uploader
    uploaded_file = st.file_uploader("Upload the workbook:", type="csv")
    if st.button("Process Workbook", disabled=not uploaded_file):
        buffer, df_ideal_worksheet = process_csv_file(uploaded_file)
        st.success("Workbook processed successfully!")

        st.download_button(
            label="Download Processed Workbook",
            data=buffer,
            file_name="processed_file.csv",
            mime="text/csv",
        )
        st.dataframe(df_ideal_worksheet)


if __name__ == "__main__":
    main()
