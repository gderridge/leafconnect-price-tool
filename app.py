import streamlit as st
import pandas as pd
import io
import os
import json
import requests
from datetime import datetime
import streamlit_authenticator as stauth

# 1. Page Configuration
st.set_page_config(page_title="LeafConnect Tools", layout="wide")

# --- AUTHENTICATION SETUP ---
# Pull secure credentials from Railway environment variables
admin_username = os.environ.get("APP_USERNAME", "admin")
admin_password = os.environ.get("APP_PASSWORD", "default_password_please_change")

# Hash the password securely in memory
hashed_passwords = stauth.Hasher(passwords=[admin_password]).generate()

credentials = {
    "usernames": {
        admin_username: {
            "email": "admin@leafconnect.com",
            "name": "Admin",
            "password": hashed_passwords[0]
        }
    }
}

authenticator = stauth.Authenticate(
    credentials,
    "leafconnect_cookie",
    "secure_signature_key",
    cookie_expiry_days=30
)

# Render the login widget
authenticator.login()

# --- APP LOGIC (ONLY SHOWS IF LOGGED IN) ---
if st.session_state["authentication_status"] is False:
    st.error("Username/password is incorrect")
    
elif st.session_state["authentication_status"] is None:
    st.warning("Please enter your username and password to access the tools.")
    
elif st.session_state["authentication_status"]:
    
    # --- INITIALIZE SESSION STATE ---
    if 'msrp_input' not in st.session_state:
        st.session_state.msrp_input = pd.DataFrame({
            "SKU": ["", "", ""], 
            "Cost Price": [0.0, 0.0, 0.0]
        })

    # --- CONNECT TO SECURE ZOHO SHEET ---
    @st.cache_data(ttl=5)
    def fetch_zoho_markups():
        try:
            zoho_env = os.environ.get("ZOHO_CREDENTIALS", "{}")
            creds = json.loads(zoho_env)
            
            client_id = creds.get("client_id")
            client_secret = creds.get("client_secret")
            refresh_token = creds.get("refresh_token")
            resource_id = creds.get("resource_id")
            worksheet_name = creds.get("worksheet_name", "Sheet1")
            
            if not all([client_id, client_secret, refresh_token, resource_id]):
                return pd.DataFrame({'Brand': ['Setup Needed'], 'Subcategory': ['Setup Needed'], 'Multiplier': [1.0]})
            
            token_url = "https://accounts.zoho.com/oauth/v2/token"
            token_params = {
                "refresh_token": refresh_token,
                "client_id": client_id,
                "client_secret": client_secret,
                "grant_type": "refresh_token"
            }
            token_res = requests.post(token_url, params=token_params)
            token_res.raise_for_status()
            access_token = token_res.json().get("access_token")
            
            data_url = f"https://sheet.zoho.com/api/v2/{resource_id}"
            headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
            params = {
                "method": "worksheet.records.fetch",
                "worksheet_name": worksheet_name
            }
            
            data_res = requests.get(data_url, headers=headers, params=params)
            data_res.raise_for_status()
            
            records = data_res.json().get("records", [])
            df = pd.DataFrame(records)
            
            if df.empty or 'Brand' not in df.columns:
                return pd.DataFrame({'Brand': ['Setup Needed'], 'Subcategory': ['Setup Needed'], 'Multiplier': [1.0]})
                
            return df.dropna(subset=['Brand'])
            
        except Exception as e:
            st.error(f"⚠️ Zoho Sheets Connection Failed: {e}")
            return pd.DataFrame({'Brand': ['Setup Needed'], 'Subcategory': ['Setup Needed'], 'Multiplier': [1.0]})

    df_markups = fetch_zoho_markups()

    # --- SIDEBAR NAVIGATION ---
    st.sidebar.image("LeafConnect_Logo-09.png", width=200)
    
    # Place the logout button right under the logo in the sidebar
    authenticator.logout("Logout", "sidebar")
    st.sidebar.divider()
    
    st.sidebar.title("Navigation")
    page = st.sidebar.radio("Select a Tool:", ["Price Comparison", "MSRP Calculator"])
    st.sidebar.divider()

    # ==========================================
    # TOOL 1: PRICE COMPARISON
    # ==========================================
    if page == "Price Comparison":
        st.title("Price List Comparison Tool")

        with st.expander("How to use this tool"):
            st.write("1. Upload your old and new price lists using the sidebar.")
            st.write("2. Toggle and set your desired percentage threshold for orange highlights.")
            st.write("3. Select the appropriate SKU and Price columns for each file.")
            st.write("4. Click 'Compare Prices' to view the summary and download the formatted Excel file.")

        st.sidebar.header("Data Upload")
        file_old = st.sidebar.file_uploader("Upload Old Price List", type=['csv', 'xlsx', 'xls'])
        file_new = st.sidebar.file_uploader("Upload New Price List", type=['csv', 'xlsx', 'xls'])

        st.sidebar.divider()
        st.sidebar.header("Highlight Settings")
        use_threshold = st.sidebar.toggle("Enable Orange Highlights")

        if use_threshold:
            threshold_pct = st.sidebar.number_input("Threshold (%)", min_value=0.1, value=3.0, step=0.5)
        else:
            threshold_pct = 0

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

                if st.button("Compare Prices", type="primary"):
                    with st.spinner('Merging lists and calculating differences...'):
                        df_old_clean = df_old[[old_sku_col, old_price_col]].copy()
                        df_old_clean = df_old_clean.rename(columns={old_sku_col: 'SKU', old_price_col: 'Old Price'})
                        
                        df_new_clean = df_new[[new_sku_col, new_price_col]].copy()
                        df_new_clean = df_new_clean.rename(columns={new_sku_col: 'SKU', new_price_col: 'New Price'})
                        
                        df_old_clean['Old Price'] = pd.to_numeric(df_old_clean['Old Price'].astype(str).str.replace(r'[\$,]', '', regex=True), errors='coerce')
                        df_new_clean['New Price'] = pd.to_numeric(df_new_clean['New Price'].astype(str).str.replace(r'[\$,]', '', regex=True), errors='coerce')

                        merged_df = pd.merge(df_old_clean, df_new_clean, on='SKU', how='inner')
                        merged_df['% Difference'] = (merged_df['New Price'] - merged_df['Old Price']) / merged_df['Old Price']

                    st.success("Comparison Complete!")
                    st.metric(label="Total SKUs Matched", value=len(merged_df))
                    st.dataframe(merged_df.head(), width='stretch')

                    output = io.BytesIO()
                    writer = pd.ExcelWriter(output, engine='xlsxwriter')
                    merged_df.to_excel(writer, index=False, sheet_name='Price Comparison')

                    workbook = writer.book
                    worksheet = writer.sheets['Price Comparison']

                    pct_format = workbook.add_format({'num_format': '0.00%'})
                    red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                    green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
                    orange_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})

                    worksheet.set_column('D:D', 15, pct_format)
                    row_max = len(merged_df) + 1
                    
                    if use_threshold and threshold_pct > 0:
                        thresh_val = threshold_pct / 100.0
                        worksheet.conditional_format(f'D2:D{row_max}', {'type': 'cell', 'criteria': '>', 'value': thresh_val, 'format': orange_format})
                        worksheet.conditional_format(f'D2:D{row_max}', {'type': 'cell', 'criteria': '<', 'value': -thresh_val, 'format': orange_format})

                    worksheet.conditional_format(f'D2:D{row_max}', {'type': 'cell', 'criteria': '<', 'value': 0, 'format': red_format})
                    worksheet.conditional_format(f'D2:D{row_max}', {'type': 'cell', 'criteria': '>', 'value': 0, 'format': green_format})

                    writer.close()
                    output.seek(0)
                    dynamic_filename = f"price_comparison_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"

                    st.download_button(label="Download Compiled Excel File", data=output, file_name=dynamic_filename, type="primary")
                    
            except Exception as e:
                st.error(f"An error occurred: {e}")

    # ==========================================
    # TOOL 2: MSRP CALCULATOR
    # ==========================================
    elif page == "MSRP Calculator":
        st.title("MSRP Calculator")
        st.write("Markups are synced securely from your Zoho Sheet.")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            selected_brand = st.selectbox("Select Brand", df_markups['Brand'].unique())
        
        with col2:
            available_subcats = df_markups[df_markups['Brand'] == selected_brand]['Subcategory'].unique()
            selected_subcat = st.selectbox("Select Subcategory", available_subcats)
            
        with col3:
            try:
                active_multiplier = df_markups[
                    (df_markups['Brand'] == selected_brand) & 
                    (df_markups['Subcategory'] == selected_subcat)
                ]['Multiplier'].values[0]
                st.metric(label="Active Multiplier", value=f"x {active_multiplier}")
            except IndexError:
                st.error("No multiplier found.")
                active_multiplier = 1.0

        st.divider()
        st.subheader("Input SKU & Cost Data")
        st.write("Type directly into the table below. Click the bottom row to add more items.")
        
        edited_data = st.data_editor(st.session_state.msrp_input, num_rows="dynamic", use_container_width=True)
        
        if st.button("Calculate MSRP", type="primary"):
            result_df = edited_data.copy()
            result_df = result_df[result_df['SKU'].str.strip() != ""]
            
            if not result_df.empty:
                result_df['Calculated MSRP'] = result_df['Cost Price'] * float(active_multiplier)
                
                st.success("Calculation Complete!")
                st.dataframe(result_df, use_container_width=True)
                
                csv = result_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download MSRP Data (CSV)",
                    data=csv,
                    file_name=f"msrp_calc_{datetime.now().strftime('%Y-%m-%d')}.csv",
                    mime="text/csv",
                    type="primary"
                )
            else:
                st.warning("Please enter at least one valid SKU and Cost Price.")