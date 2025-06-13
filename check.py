import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.title("🔍 Asset Code Comparator (Integer-Matched)")

# Clean and convert asset code to integer (removes symbols, leading/trailing junk)
def cleanse(asset):
    if pd.isna(asset):
        return None
    digits = re.sub(r'\D', '', str(asset))
    return int(digits) if digits.isdigit() else None

# Extract cleaned asset codes from all sheets in cleaned file
def extract_cleaned_codes_from_all_sheets(cleaned_excel_file, relevant_sheets=None):
    all_sheets = pd.read_excel(cleaned_excel_file, sheet_name=None)
    cleaned_codes = set()

    for sheet_name, df in all_sheets.items():
        if relevant_sheets and sheet_name not in relevant_sheets:
            continue
        if 'รหัสทรัพย์สิน' in df.columns:
            cleaned_vals = df['รหัสทรัพย์สิน'].dropna().apply(cleanse).dropna().tolist()
            cleaned_codes.update(cleaned_vals)

    return cleaned_codes

# Compare original against cleaned set
def process_comparison(original_df, cleaned_code_set):
    results = []
    for _, row in original_df.iterrows():
        raw = str(row.get('รหัสทรัพย์สิน', ''))
        assets = re.split(r'[ ,/\\*]+', raw)
        for asset in filter(None, map(str.strip, assets)):
            cleaned = cleanse(asset)
            status = "Found" if cleaned in cleaned_code_set else "Missing"
            results.append({
                "OriginalEntry": raw,
                "ExtractedAsset": asset,
                "CleanedAsset (int)": cleaned,
                "MatchStatus": status
            })
    return pd.DataFrame(results)

# --- UI ---
st.markdown("Upload your files:")

col1, col2 = st.columns(2)
with col1:
    original_file = st.file_uploader("📄 Original Excel File", type=["xlsx"], key="original")
with col2:
    cleaned_file = st.file_uploader("✅ Cleaned Excel File", type=["xlsx"], key="cleaned")

if original_file and cleaned_file:
    try:
        original_df = pd.read_excel(original_file)

        use_all_sheets = st.checkbox("Use all sheets in cleaned file", value=True)
        if not use_all_sheets:
            cleaned_preview = pd.read_excel(cleaned_file, sheet_name=None)
            sheet_list = list(cleaned_preview.keys())
            selected_sheets = st.multiselect("Select sheets", sheet_list, default=["Correct Data"])
        else:
            selected_sheets = None

        # Extract cleaned asset codes
        cleaned_code_set = extract_cleaned_codes_from_all_sheets(cleaned_file, relevant_sheets=selected_sheets)

        # Compare
        result_df = process_comparison(original_df, cleaned_code_set)

        st.markdown("### ✅ Comparison Result")
        st.dataframe(result_df, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='Comparison')
        output.seek(0)

        st.download_button(
            label="⬇️ Download Match Report",
            data=output,
            file_name="asset_comparison_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
