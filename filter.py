import streamlit as st
import pandas as pd
import re
from io import BytesIO
import time

def is_incorrect(asset_value, incorrect_terms):
    if isinstance(asset_value, str):
        for term in incorrect_terms:
            if term in asset_value:
                return True
        if len(asset_value.split(',')) > 1 or len(asset_value.split()) > 1:
            return True
    if pd.isna(asset_value) or asset_value.strip() == '':
        return True
    return False

def cleanse_asset_code(asset_value):
    if isinstance(asset_value, str):
        return re.sub(r'\D', '', asset_value)
    return asset_value

def ensure_utf8_encoding(df):
    return df.apply(lambda x: x.str.encode('utf-8').str.decode('utf-8') if x.dtype == "object" else x)

def process_excel(uploaded_file, sheet_name):
    progress_bar = st.progress(0)
    
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    progress_bar.progress(20)
    asset_column = 'รหัสทรัพย์สิน'
    central_asset_column = 'หน่วยงานกลางดูแลทรัพย์สิน'
    
    incorrect_terms = [
        "รหัสทรัพย์สิน", "ไม่มี", "Computer", "Notebook", 
        "ไม่มีรหัสทรัพย์สิน", "แทบเล็ต", "รายละเอียด", "nan", "-", "์Notebook",
        "ทบ. 5589338", "ทบ. 5589339 (รูปเครื่อง)"
    ]
    
    filter_values = ['ทรัพย์สินองค์กร พื้นที่ธาราพาร์ค', 'ทรัพย์สินองค์กร พื้นที่สีลม/สาทร']
    
    incorrect_data = df[df[asset_column].apply(lambda x: is_incorrect(x, incorrect_terms))]
    correct_data = df[~df[asset_column].apply(lambda x: is_incorrect(x, incorrect_terms))]
    correct_data[asset_column] = correct_data[asset_column].apply(cleanse_asset_code)
    progress_bar.progress(50)
    
    split_rows = []
    for _, row in incorrect_data.iterrows():
        asset_values = str(row[asset_column]).replace(',', ' ').split()
        for asset in asset_values:
            if asset.strip():
                new_row = row.copy()
                new_row[asset_column] = asset.strip()
                split_rows.append(new_row)
    split_result = pd.DataFrame(split_rows)
    progress_bar.progress(70)
    
    merged_result = pd.concat([correct_data, split_result], ignore_index=True)
    filtered_data = merged_result[merged_result[central_asset_column].isin(filter_values)]
    
    incorrect_data = ensure_utf8_encoding(incorrect_data)
    correct_data = ensure_utf8_encoding(correct_data)
    split_result = ensure_utf8_encoding(split_result)
    merged_result = ensure_utf8_encoding(merged_result)
    filtered_data = ensure_utf8_encoding(filtered_data)
    progress_bar.progress(90)
    
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        correct_data.to_excel(writer, sheet_name='Correct Data', index=False)
        incorrect_data.to_excel(writer, sheet_name='Incorrect Data', index=False)
        split_result.to_excel(writer, sheet_name='Result of Split', index=False)
        merged_result.to_excel(writer, sheet_name='Merged Data', index=False)
        filtered_data.to_excel(writer, sheet_name='Tara-Silom', index=False)
    
    output_buffer.seek(0)
    progress_bar.progress(100)
    time.sleep(0.5)
    progress_bar.empty()
    return output_buffer

st.title("Excel Data Cleaner")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
sheet_name = st.text_input("Enter the sheet name to filter", value="File ทำงาน")

if uploaded_file:
    st.write("File uploaded successfully!")
    output_file = process_excel(uploaded_file, sheet_name)
    st.download_button(
        label="Download Processed File",
        data=output_file,
        file_name="cleaned_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
