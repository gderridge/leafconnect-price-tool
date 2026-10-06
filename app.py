import streamlit as st
import pandas as pd
import io
from datetime import datetime

# 1. Page Configuration
st.set_page_config(page_title="LeafConnect Price Comparison", layout="wide")

# 2. Add Logo and Header
st.image("LeafConnect_Logo-09.png", width=250)
st.title("Price List Comparison Tool")

with st.expander("How to use this tool"):
    st.write("1. Upload your old and new price lists using the sidebar.")
    st.write("2. Toggle and set your desired percentage threshold for orange highlights.")
    st.write("3. Select the appropriate SKU and Price columns for each file.")
    st.write("4. Click 'Compare Prices' to view the summary and download the formatted Excel file.")

# 3. Sidebar Uploads & Settings
st.sidebar.header("Data Upload")
file_old = st.sidebar.file_uploader("Upload Old Price List", type=['csv', 'xlsx', 'xls'])
file_new = st.sidebar.file_uploader("Upload New Price List", type=['csv', 'xlsx', 'xls'])

st.sidebar.divider()
st.sidebar.header("Threshold Settings")

# Checkbox acts as an On/Off switch
use_threshold = st.sidebar.toggle("Enable Orange Highlights")

if use_threshold:
    threshold_pct = st.sidebar.number_input(
        "Threshold (%)", 
        min_value=0.1, 
        value=3.0, 
        step=0.5,
        help="Differences greater than this percentage will be highlighted orange."
    )
else:
    threshold_pct = 0  # Forces the orange formatting to be skipped

if file_old and file_new:
    def load_data(file):
        if file.name.endswith('.csv'):
            return pd.read_csv(file)
        return pd.read_excel(file)

    try:
        df_old = load_data(file_old)
        df_new = load_data(file_new)
        
        st.divider()
        st.subheader("Map Your Columns")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Old Price List**")
            old_sku_col = st.selectbox("Select SKU column", df_old.columns, key='old_sku')
            old_price_col = st.selectbox("Select Price column", df_old.columns, key='old_price')
            
        with col2:
            st.write("**New Price List**")
            new_sku_col = st.selectbox("Select SKU column", df_new.columns, key='new_sku')
            new_price_col = st.selectbox("Select Price column", df_new.columns, key='new_price')

        # Use a primary button for the main action
        if st.button("Compare Prices", type="primary"):
            
            # Show a loading spinner during calculation
            with st.spinner('Merging lists and calculating differences...'):
                df_old_clean = df_old[[old_sku_col, old_price_col]].copy()
                df_old_clean = df_old_clean.rename(columns={old_sku_col: 'SKU', old_price_col: 'Old Price'})
                
                df_new_clean = df_new[[new_sku_col, new_price_col]].copy()
                df_new_clean = df_new_clean.rename(columns={new_sku_col: 'SKU', new_price_col: 'New Price'})
                
                df_old_clean['Old Price'] = pd.to_numeric(df_old_clean['Old Price'].astype(str).str.replace(r'[\$,]', '', regex=True), errors='coerce')
                df_new_clean['New Price'] = pd.to_numeric(df_new_clean['New Price'].astype(str).str.replace(r'[\$,]', '', regex=True), errors='coerce')

                merged_df = pd.merge(df_old_clean, df_new_clean, on='SKU', how='inner')
                merged_df['% Difference'] = (merged_df['New Price'] - merged_df['Old Price']) / merged_df['Old Price']

            # Display a success metric
            st.success("Comparison Complete!")
            st.metric(label="Total SKUs Matched", value=len(merged_df))

            st.write("### Preview of Merged Data")
            st.dataframe(merged_df.head(), width='stretch')

            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            merged_df.to_excel(writer, index=False, sheet_name='Price Comparison')

            workbook = writer.book
            worksheet = writer.sheets['Price Comparison']

            # Define formatting styles
            pct_format = workbook.add_format({'num_format': '0.00%'})
            red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
            green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
            orange_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})

            worksheet.set_column('D:D', 15, pct_format)
            row_max = len(merged_df) + 1
            
            # Apply threshold formatting first so it takes priority
            if use_threshold and threshold_pct > 0:
                thresh_val = threshold_pct / 100.0
                worksheet.conditional_format(f'D2:D{row_max}', {'type': 'cell', 'criteria': '>', 'value': thresh_val, 'format': orange_format})
                worksheet.conditional_format(f'D2:D{row_max}', {'type': 'cell', 'criteria': '<', 'value': -thresh_val, 'format': orange_format})

            # Apply standard red/green formatting
            worksheet.conditional_format(f'D2:D{row_max}', {'type': 'cell', 'criteria': '<', 'value': 0, 'format': red_format})
            worksheet.conditional_format(f'D2:D{row_max}', {'type': 'cell', 'criteria': '>', 'value': 0, 'format': green_format})

            writer.close()
            output.seek(0)

            current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            dynamic_filename = f"price_comparison_{current_time}.xlsx"

            st.download_button(
                label="Download Compiled Excel File",
                data=output,
                file_name=dynamic_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")