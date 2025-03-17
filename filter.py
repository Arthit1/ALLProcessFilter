import streamlit as st
import pandas as pd
import re
from io import BytesIO
import time
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Constants
ASSET_COLUMN = 'รหัสทรัพย์สิน'
CENTRAL_ASSET_COLUMN = 'หน่วยงานกลางรับทราบและตรวจสอบ'
FILTER_VALUES = ['อภิสรา สีดาคุณ']
INCORRECT_TERMS = {
    "รหัสทรัพย์สิน", "ไม่มี", "Computer", "Notebook", "ไม่มีรหัสทรัพย์สิน",
    "แทบเล็ต", "รายละเอียด", "nan", "-", "์Notebook", "ทบ. 5589338", "ทบ. 5589339 (รูปเครื่อง)"
}

def is_invalid_asset(asset_value):
    """Check if the asset value is invalid."""
    if pd.isna(asset_value) or not isinstance(asset_value, str) or asset_value.strip() == '':
        return True
    if asset_value.isdigit() and int(asset_value) == 0:  # Detect all-zero codes
        return True
    if not asset_value.isdigit():  # Detect non-numeric values
        return True
    if any(term in asset_value for term in INCORRECT_TERMS):
        return True
    return False

def cleanse_asset_code(asset_value):
    """Remove non-numeric characters from asset codes."""
    return re.sub(r'\D', '', asset_value) if isinstance(asset_value, str) else asset_value

def ensure_utf8_encoding(df):
    """Ensure all string columns are UTF-8 encoded."""
    return df.apply(lambda x: x.str.encode('utf-8').str.decode('utf-8') if x.dtype == "object" else x)

def highlight_duplicates(workbook, sheet_name, column_name):
    """Apply highlighting to duplicate values in a specific column."""
    ws = workbook[sheet_name]
    duplicate_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Find the column letter
    column_letter = None
    for cell in ws[1]:  
        if cell.value == column_name:
            column_letter = cell.column_letter
            break

    if not column_letter:
        return  

    # Identify duplicate values
    column_values = [ws[f"{column_letter}{row}"].value for row in range(2, ws.max_row + 1)]
    duplicate_values = {val for val in column_values if column_values.count(val) > 1}

    # Apply yellow highlight to duplicates
    for row in range(2, ws.max_row + 1):
        cell = ws[f"{column_letter}{row}"]
        if cell.value in duplicate_values:
            cell.fill = duplicate_fill

def process_excel(uploaded_file, sheet_name):
    """Process the uploaded Excel file and return cleaned data."""
    progress_bar = st.progress(0)
    
    # Read file
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    progress_bar.progress(20)

    # Identify incorrect, duplicate, and correct rows
    incorrect_mask = df[ASSET_COLUMN].apply(is_invalid_asset)
    correct_data = df[~incorrect_mask].copy()
    incorrect_data = df[incorrect_mask].copy()
    
    # Clean correct asset codes
    correct_data[ASSET_COLUMN] = correct_data[ASSET_COLUMN].apply(cleanse_asset_code)
    progress_bar.progress(50)

    # Split incorrect asset values
    split_rows = []
    for _, row in incorrect_data.iterrows():
        asset_values = re.split(r'[ ,]+', str(row[ASSET_COLUMN]))  
        for asset in filter(None, map(str.strip, asset_values)):  
            new_row = row.copy()
            new_row[ASSET_COLUMN] = asset
            split_rows.append(new_row)

    split_result = pd.DataFrame(split_rows)
    progress_bar.progress(70)

    # Merge correct and split results
    merged_result = pd.concat([correct_data, split_result], ignore_index=True)

    # Detect Duplicates in 'รหัสทรัพย์สิน'
    merged_result["Duplicate"] = merged_result.duplicated(subset=[ASSET_COLUMN], keep=False).map({True: "Yes", False: "No"})

    # Filter Specific Data from the "Correct Data" sheet
    tara_silom_data = correct_data[correct_data[CENTRAL_ASSET_COLUMN] == 'อภิสรา สีดาคุณ']

    # Separate "Duplicate & Wrong Data"
    duplicate_wrong_data = merged_result[
        (merged_result["Duplicate"] == "Yes") | (merged_result[ASSET_COLUMN].apply(is_invalid_asset))
    ]

    # Create a sheet for "Correct Data (No Duplicates)"
    correct_no_duplicates = merged_result[merged_result["Duplicate"] == "No"].copy()

    # Ensure encoding
    for df in [correct_data, incorrect_data, split_result, merged_result, tara_silom_data, duplicate_wrong_data, correct_no_duplicates]:
        df[:] = ensure_utf8_encoding(df)

    progress_bar.progress(90)

    # Save to Excel (Main File)
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        sheets = {
            'Correct Data': correct_no_duplicates,  # Correct Data without duplicates
            'Tara-Silom': tara_silom_data,  # Data filtered by "อภิสรา สีดาคุณ"
            'Duplicate & Wrong Data': duplicate_wrong_data,  # Data with issues
        }
        for sheet_name, data in sheets.items():
            data.to_excel(writer, sheet_name=sheet_name, index=False)

    output_buffer.seek(0)

    # Load workbook for formatting
    workbook = load_workbook(output_buffer)
    highlight_duplicates(workbook, "Correct Data", ASSET_COLUMN)  

    output_buffer = BytesIO()
    workbook.save(output_buffer)
    output_buffer.seek(0)

    progress_bar.progress(100)
    time.sleep(0.5)
    progress_bar.empty()
    
    return output_buffer

# Streamlit UI
st.title("ALLProcess Data Cleaner V2.0")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

# Extract sheet names dynamically
sheet_names = []
if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
    except Exception as e:
        st.error(f"Error reading sheets: {e}")

# Allow selection only if sheets exist
if sheet_names:
    sheet_name = st.selectbox("Select the sheet to filter", sheet_names)
else:
    sheet_name = None

if uploaded_file and sheet_name:
    st.success(f"Processing sheet: {sheet_name}")
    output_file = process_excel(uploaded_file, sheet_name)

    st.download_button(
        label="Download Processed File",
        data=output_file,
        file_name="cleaned_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
