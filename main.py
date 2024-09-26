from io import BytesIO

import pandas as pd
import streamlit as st


def process_csv_file(file):
    # Read the CSV file into a DataFrame
    df_pcm_generated = pd.read_csv(file)

    # Create a new DataFrame for the "Ideal Worksheet"
    df_ideal_worksheet = pd.DataFrame()

    # Add language columns to the "Ideal Worksheet" DataFrame
    filtered_languages = df_pcm_generated[
        df_pcm_generated["Language"] != "Language pack"
    ]
    unique_languages = filtered_languages["Language"].unique()
    for lang in unique_languages:
        df_ideal_worksheet[lang] = ""

    filtered_row_data = df_pcm_generated.iloc[1:]
    for row_index, row_data in filtered_row_data.iterrows():
        index_offset = 23 * (row_index % 4)
        df_ideal_worksheet.loc[0, row_data["Language"]] = row_data["Route Map Name"]
        df_ideal_worksheet.loc[1, row_data["Language"]] = row_data["Description"]
        df_ideal_worksheet.loc[2 + index_offset, row_data["Language"]] = row_data[
            "Step Name"
        ]
        df_ideal_worksheet.loc[3 + index_offset, row_data["Language"]] = row_data[
            "Modify Stage"
        ]
        df_ideal_worksheet.loc[4 + index_offset, row_data["Language"]] = row_data[
            "Step Name"
        ]
        df_ideal_worksheet.loc[5 + index_offset, row_data["Language"]] = row_data[
            "Step Description"
        ]
        df_ideal_worksheet.loc[6 + index_offset, row_data["Language"]] = row_data[
            "Step Type"
        ]
        df_ideal_worksheet.loc[7 + index_offset, row_data["Language"]] = row_data[
            "Roles"
        ]
        df_ideal_worksheet.loc[8 + index_offset, row_data["Language"]] = row_data[
            "Step Introduction and Mouseover"
        ]
        df_ideal_worksheet.loc[9 + index_offset, row_data["Language"]] = row_data[
            "Step Name After Completion"
        ]
        df_ideal_worksheet.loc[10 + index_offset, row_data["Language"]] = row_data[
            "Step Mode"
        ]
        df_ideal_worksheet.loc[11 + index_offset, row_data["Language"]] = row_data[
            "Exit Button Text"
        ]
        df_ideal_worksheet.loc[12 + index_offset, row_data["Language"]] = row_data[
            "Step Exit Text"
        ]
        df_ideal_worksheet.loc[13 + index_offset, row_data["Language"]] = row_data[
            "Previous Step Exit Button Text"
        ]
        df_ideal_worksheet.loc[14 + index_offset, row_data["Language"]] = row_data[
            "Previous Step Exit Text"
        ]
        df_ideal_worksheet.loc[15 + index_offset, row_data["Language"]] = row_data[
            "Entry User"
        ]
        df_ideal_worksheet.loc[16 + index_offset, row_data["Language"]] = row_data[
            "Exit User"
        ]
        df_ideal_worksheet.loc[17 + index_offset, row_data["Language"]] = row_data[
            "Start Date"
        ]
        df_ideal_worksheet.loc[18 + index_offset, row_data["Language"]] = row_data[
            "Exit Date"
        ]
        df_ideal_worksheet.loc[19 + index_offset, row_data["Language"]] = row_data[
            "Enforce Start Date"
        ]
        df_ideal_worksheet.loc[20 + index_offset, row_data["Language"]] = row_data[
            "Automatic send on due date"
        ]
        df_ideal_worksheet.loc[21 + index_offset, row_data["Language"]] = row_data[
            "Iterative Button Text"
        ]
        df_ideal_worksheet.loc[22 + index_offset, row_data["Language"]] = row_data[
            "Reject Button Mouseover Text"
        ]
        df_ideal_worksheet.loc[23 + index_offset, row_data["Language"]] = row_data[
            "Step Exit Reminder"
        ]
        df_ideal_worksheet.loc[24 + index_offset, row_data["Language"]] = row_data[
            "Step Exit Reminder Text"
        ]

    # Save the modified DataFrame to a CSV file in memory
    buffer = BytesIO()
    df_ideal_worksheet.to_csv(buffer, index=False)
    buffer.seek(0)
    return buffer


def main():
    # Set the page configuration
    st.set_page_config(
        page_title="Workbook Converter", page_icon="🗂", layout="centered"
    )
    # Display the title in SAP blue
    st.markdown(
        '<h1 style="color:#008BBF;">Workbook Converter</h1>',
        unsafe_allow_html=True,
    )

    # Add the dropdown selector for Object Type
    object_type = st.selectbox("Select the object type:", ["Route map"])

    # File uploader
    uploaded_file = st.file_uploader("Upload the workbook:", type="csv")
    if uploaded_file is not None:
        if st.button("Process Workbook"):
            buffer = process_csv_file(uploaded_file)
            st.download_button(
                label="Download Processed Workbook",
                data=buffer,
                file_name="processed_file.csv",
                mime="text/csv",
            )
            st.success("Workbook processed successfully!")


if __name__ == "__main__":
    main()
