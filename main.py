import streamlit as st
import pandas as pd
import json
import io
import base64

st.set_page_config(page_title="Excel to JSON Converter", page_icon="📊", layout="wide")

st.title("📊 Excel to JSON Converter")
st.write("Upload an Excel file and convert it to JSON format.")

def process_excel(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names
    all_dataframes = []

    for sheet in sheets:
        df = pd.read_excel(xls, sheet_name=sheet)
        df['Sheet Name'] = sheet
        all_dataframes.append(df)

    df_combined = pd.concat(all_dataframes, ignore_index=True)

    relevant_columns = ['Data Domain', 'Products', 'Business Term', 'Business Term Abbreviations',
                        'Business Definition', 'UOM', 'Related Terms', 'Data Owner', 'Data Steward', 'Sheet Name']
    relevant_columns_combined = df_combined[relevant_columns]

    df_combined_cleaned = relevant_columns_combined.where(pd.notnull(relevant_columns_combined), None)
    return df_combined_cleaned.to_dict(orient='records')

def get_download_link(json_data, filename):
    json_string = json.dumps(json_data, indent=4)
    b64 = base64.b64encode(json_string.encode()).decode()
    href = f'<a href="data:application/json;base64,{b64}" download="{filename}">Download JSON file</a>'
    return href

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    if st.button("Convert to JSON"):
        try:
            with st.spinner("Processing Excel file..."):
                combined_glossary_json = process_excel(uploaded_file)

            st.success("Excel file processed successfully!")

            st.subheader("Generated JSON Preview")
            st.json(combined_glossary_json[:5])  # Show first 5 items as preview

            st.subheader("Download JSON")
            download_link = get_download_link(combined_glossary_json, "business_glossary_combined.json")
            st.markdown(download_link, unsafe_allow_html=True)

            st.info("Click the link above to download the full JSON file.")

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            st.error("Please make sure your Excel file has the correct structure and try again.")
else:
    st.info("Please upload an Excel file to begin.")

st.markdown("---")
st.markdown("Created with ❤️ using Streamlit")
