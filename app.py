import pandas as pd
import streamlit as st
from MIDP_Formatter_Functions import *
from datetime import datetime
from io import BytesIO
import io
from xlsxwriter import Workbook


st.set_page_config(page_title="MIDP Validator & Formatter", layout="centered")
st.title("MIDP Validator & Formatter")

# --- Mapping provided (keys must exist in last month; values must exist in this month) ---
columns_required = {
    "Information Container": "Information Container",
    "Information Container Title / Description": "Information Container Title",
    "ID number": "P6 Activity ID (Please Review)",
    "Project Milestone": "LWR - Phase",
    "Comment": "Comments",
    "ID": "",
    "LOA": "",
    "Status Code": "⚡ Document Status",
    "Revision": "⚡ Last Published Revision",
    "Issue Date (Planned)": "Planned Issue Date",
    "Issue Date (Actual)": "⚡ Published Date",
    "Authorised by TW Service manager": "",
    "Created by": "Document Workstream (Owning team)",
    "Y/N": "DCO Submission Document",
    "Y/N.": "",
    "File Extension": "Source File Extension (Please Review)",
    "Security Reference": "Security Reference",
    "Project": "Project",
    "Functional Breakdown": "Functional Breakdown",
    "Spatial Breakdown": "Spatial Breakdown",
    "Document Type": "Document Type",
    "Discipline": "Discipline"
}

# --- UI: Upload two CSVs ---
st.subheader("1) Upload CSV files")
last_file = st.file_uploader("Upload LAST month's MIDP (CSV)", type=["csv"], key="last")
current_file = st.file_uploader("Upload THIS month's MIDP export (CSV)", type=["csv"], key="current")

# --- Helpers ---
def load_csv_as_str(file) -> pd.DataFrame:
    """Load CSV with all columns as string, preserving headers exactly."""
    return pd.read_csv(file, dtype=str, keep_default_na=False, na_values=[])

def validate_columns(df: pd.DataFrame, required_cols: list[str]) -> list[str]:
    """Return a list of missing column names from df."""
    df_cols = set(df.columns)
    return [c for c in required_cols if c not in df_cols]

# --- 2) Helper to export Styler to Excel in memory ---
@st.cache_data
def styler_to_excel_bytes(styler: pd.io.formats.style.Styler, file_name: str = "styled.xlsx") -> BytesIO:
    """
    Render a pandas Styler to an in-memory Excel file.
    - Preserves styles (colors, formats).
    - Uses openpyxl engine.
    """
    buffer = BytesIO()
    # You can pass `ext='.xlsx'` if needed; default is fine.
    styler.to_excel(buffer, engine="openpyxl")
    buffer.seek(0)
    return buffer


# Function to apply styles based on conditions
def highlight_changes(row):
    styles = []
    
    for col in row.index:
        
        if row.name in df1.index and row.name in df2.index:
            val1 = df1.at[row.name, col]
            val2 = df2.at[row.name, col]
            #If both cells have contents in it

            if val1 != "" and val2 != "": 
                # Check if both values are the same type
                if type(val1) == type(val2):
                    #In the case where both cells have things in them BUT are different, turn yellow
                    if val1 != val2:
                        styles.append('background-color: yellow')
                    else:
                        styles.append('')
                else:
                    styles.append('')
            #If either cell has nothing in it
            else:
                # Highlight empty cells
                #If df1 has nothing and df2 has something in it, its a new cell and turn it green
                if val1 == "" and val2 != "":
                    styles.append('background-color: lime') #new cells
                #If df1 has something and df2 has nothing in it, a cell is deleted and turn it red
                elif val1 != "" and val2 == "":
                    styles.append('background-color: red') #deleted cell
                else:
                    styles.append('')
        elif row.name in df2.index and row.name not in df1.index:
            styles.append('background-color: lime')
        elif row.name in df1.index and row.name not in df2.index:
            styles.append('background-color: red')
        else:
            styles.append('')
    return styles


def red_background(s):
    return ['background-color: red' for _ in s]

# --- Submit ---
st.subheader("2) Validate & Generate")
submitted = st.button("Submit")


if submitted:
    # Basic presence check
    if not last_file or not current_file:
        st.error("Please upload both CSV files before submitting.")
        st.stop()

    # Load CSVs safely
    try:
        last_df = load_csv_as_str(last_file)
        current_df = load_csv_as_str(current_file)
    except Exception as e:
        st.error(f"Failed to read one or both CSV files. Error: {e}")
        st.stop()

    # Prepare required lists
    required_last = list(columns_required.keys())
    required_current = [v for v in columns_required.values() if v != ""]
    # Validate
    missing_last = validate_columns(last_df, required_last)
    missing_current = validate_columns(current_df, required_current)

    # Report errors if any
    if missing_last:
        st.error(
            "Last month's MIDP is missing the following required columns (keys):\n"
            + "\n".join(f"- {c}" for c in missing_last)
        )
        st.stop()

    if missing_current:
        st.error(
            "This month's MIDP export is missing the following required columns (values):\n"
            + "\n".join(f"- {c}" for c in missing_current)
        )
        st.stop()

   
# Call your formatter
    try:
        df1 = last_df
        df2 = format_midp_df(current_df)

        # Set 'Information Container' as the index
        df1.set_index('Information Container', inplace=True)
        df2.set_index('Information Container', inplace=True)

        # Combine the DataFrames to identify changes
        combined_df = df1.combine_first(df2).combine_first(df1)

        # Apply the styles to the DataFrame
        styled_df = df2.style.apply(highlight_changes, axis=1)

        # Create a DataFrame for deleted rows
        deleted_rows_df = df1[~df1.index.isin(df2.index)]


        # Apply the style to the DataFrame
        deleted_rows_styled_df = deleted_rows_df.style.apply(red_background, axis=None)

        # Concatenate the styled DataFrames
        final_styled_df = pd.concat([styled_df.data, deleted_rows_styled_df.data])

        final_styled_df_styled = final_styled_df.style.apply(highlight_changes, axis=1)

        #Export the styled DataFrame to an Excel file
        today = datetime.today().strftime('%y%m%d')
        
        final_styled_df_styled.to_excel(f'{today}.xlsx', index=True)
        
        st.write("Preview")
        st.dataframe(final_styled_df_styled) 

        # Convert styled DataFrame to Excel in memory
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            final_styled_df_styled.to_excel(writer, sheet_name='Sheet1')
        buffer.seek(0)

        # Download button
        st.download_button(
            label="Download Excel",
            data=buffer.getvalue(),
            file_name=f'{today}.xlsx',
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


    except Exception as e:
        st.error("An error occurred while running the formatter.")
        st.exception(e)

