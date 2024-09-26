from io import BytesIO

import pandas as pd
import streamlit as st


def process_excel_file(file_path):
    # Read the Excel file with all sheets
    xls = pd.ExcelFile(file_path)

    # Check if "Ideal Worksheet" exists, if not create it
    if "Ideal Worksheet" in xls.sheet_names:
        sheet_ideal_worksheet = pd.read_excel(file_path, sheet_name="Ideal Worksheet")
    else:
        sheet_ideal_worksheet = pd.DataFrame()

    # Read the "PCM Generated" sheet
    sheet_pcm_generated = pd.read_excel(file_path, sheet_name="PCM Generated")

    # Add language columns to the "Ideal Worksheet" sheet
    filtered_languages = sheet_pcm_generated[
        sheet_pcm_generated["Language"] != "Language pack"
    ]
    unique_languages = filtered_languages["Language"].unique()
    largest_column_index = len(sheet_ideal_worksheet.columns) - 1

    for lang in unique_languages:
        if lang not in sheet_ideal_worksheet.columns:
            largest_column_index += 1
            sheet_ideal_worksheet.insert(largest_column_index, lang, "")

    filtered_row_data = sheet_pcm_generated.iloc[1:]

    for row_index, row_data in filtered_row_data.iterrows():
        sheet_ideal_worksheet.loc[0, row_data["Language"]] = row_data["Route Map Name"]
        sheet_ideal_worksheet.loc[1, row_data["Language"]] = row_data["Description"]
        index_offset = 23 * (row_index % 4)
        sheet_ideal_worksheet.loc[2 + index_offset, row_data["Language"]] = row_data[
            "Step Name"
        ]
        sheet_ideal_worksheet.loc[4 + index_offset, row_data["Language"]] = row_data[
            "Step Name"
        ]
        sheet_ideal_worksheet.loc[3 + index_offset, row_data["Language"]] = row_data[
            "Modify Stage"
        ]
        sheet_ideal_worksheet.loc[5 + index_offset, row_data["Language"]] = row_data[
            "Step Description"
        ]
        sheet_ideal_worksheet.loc[6 + index_offset, row_data["Language"]] = row_data[
            "Step Type"
        ]
        sheet_ideal_worksheet.loc[7 + index_offset, row_data["Language"]] = row_data[
            "Roles"
        ]
        sheet_ideal_worksheet.loc[8 + index_offset, row_data["Language"]] = row_data[
            "Step Introduction and Mouseover"
        ]
        sheet_ideal_worksheet.loc[9 + index_offset, row_data["Language"]] = row_data[
            "Step Name After Completion"
        ]
        sheet_ideal_worksheet.loc[10 + index_offset, row_data["Language"]] = row_data[
            "Step Mode"
        ]
        sheet_ideal_worksheet.loc[11 + index_offset, row_data["Language"]] = row_data[
            "Exit Button Text"
        ]
        sheet_ideal_worksheet.loc[12 + index_offset, row_data["Language"]] = row_data[
            "Step Exit Text"
        ]
        sheet_ideal_worksheet.loc[13 + index_offset, row_data["Language"]] = row_data[
            "Previous Step Exit Button Text"
        ]
        sheet_ideal_worksheet.loc[14 + index_offset, row_data["Language"]] = row_data[
            "Previous Step Exit Text"
        ]
        sheet_ideal_worksheet.loc[15 + index_offset, row_data["Language"]] = row_data[
            "Entry User"
        ]
        sheet_ideal_worksheet.loc[16 + index_offset, row_data["Language"]] = row_data[
            "Exit User"
        ]
        sheet_ideal_worksheet.loc[17 + index_offset, row_data["Language"]] = row_data[
            "Start Date"
        ]
        sheet_ideal_worksheet.loc[18 + index_offset, row_data["Language"]] = row_data[
            "Exit Date"
        ]
        sheet_ideal_worksheet.loc[19 + index_offset, row_data["Language"]] = row_data[
            "Enforce Start Date"
        ]
        sheet_ideal_worksheet.loc[20 + index_offset, row_data["Language"]] = row_data[
            "Automatic send on due date"
        ]
        sheet_ideal_worksheet.loc[21 + index_offset, row_data["Language"]] = row_data[
            "Iterative Button Text"
        ]
        sheet_ideal_worksheet.loc[22 + index_offset, row_data["Language"]] = row_data[
            "Reject Button Mouseover Text"
        ]
        sheet_ideal_worksheet.loc[23 + index_offset, row_data["Language"]] = row_data[
            "Step Exit Reminder"
        ]
        sheet_ideal_worksheet.loc[24 + index_offset, row_data["Language"]] = row_data[
            "Step Exit Reminder Text"
        ]

    # Save the modified DataFrame to a BytesIO buffer with all sheets
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name in xls.sheet_names:
            pd.read_excel(file_path, sheet_name=sheet_name).to_excel(
                writer, sheet_name=sheet_name, index=False
            )
        # Ensure "Ideal Worksheet" is always included
        sheet_ideal_worksheet.to_excel(
            writer, sheet_name="Ideal Worksheet", index=False
        )
    buffer.seek(0)
    return buffer


def main():
    st.title("Excel Processing App")

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file is not None:
        file_path = uploaded_file.name
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        if st.button("Process File"):
            buffer = process_excel_file(file_path)
            st.download_button(
                label="Download Processed File",
                data=buffer,
                file_name="processed_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.success("File processed successfully!")


if __name__ == "__main__":
    main()
