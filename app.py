import streamlit as st
import pandas as pd
import re
import io
from streamlit_modal import Modal
from email_utils import send_analysis_email # Import our logic
from openpyxl.styles import Alignment, PatternFill, Border, Side

# ... [Keep your extract_area_logic, determine_config, and apply_excel_formatting functions here] ...

st.set_page_config(page_title="Real Estate Dashboard", layout="wide")

# Initialize Modal
modal = Modal(key="email_modal", title="Send Report via Email")

# Sidebar settings
st.sidebar.header("Calculation Settings")
loading_factor = st.sidebar.number_input("Loading Factor", min_value=1.0, value=1.35, step=0.001, format="%.3f")
t1 = st.sidebar.number_input("1 BHK Threshold (<)", value=600)
t2 = st.sidebar.number_input("2 BHK Threshold (<)", value=850)
t3 = st.sidebar.number_input("3 BHK Threshold (<)", value=1100)

uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type="xlsx")

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    clean_cols = {c.lower().strip(): c for c in df.columns}
    desc_col = clean_cols.get('property description')
    cons_col = clean_cols.get('consideration value')
    prop_col = clean_cols.get('property')
    
    if desc_col and cons_col and prop_col:
        # Processing Logic
        df['Carpet Area (SQ.MT)'] = df[desc_col].apply(extract_area_logic)
        df['Carpet Area (SQ.FT)'] = (df['Carpet Area (SQ.MT)'] * 10.764).round(3)
        df['Saleable Area'] = (df['Carpet Area (SQ.FT)'] * loading_factor).round(3)
        df['APR'] = df.apply(lambda r: round(r[cons_col]/r['Saleable Area'], 3) if r['Saleable Area'] > 0 else 0, axis=1)
        df['Configuration'] = df['Carpet Area (SQ.FT)'].apply(lambda x: determine_config(x, t1, t2, t3))
        
        valid_df = df[df['Carpet Area (SQ.FT)'] > 0].sort_values([prop_col, 'Configuration', 'Carpet Area (SQ.FT)'])
        summary = valid_df.groupby([prop_col, 'Configuration', 'Carpet Area (SQ.FT)']).agg(
            Min_APR=('APR', 'min'), Max_APR=('APR', 'max'), Avg_APR=('APR', 'mean'),
            Median_APR=('APR', 'median'),
            Mode_APR=('APR', lambda x: x.mode().iloc[0] if not x.mode().empty else 0),
            Property_Count=(prop_col, 'count')
        ).reset_index()
        
        # Prepare the Excel file in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            apply_excel_formatting(df, writer, 'Raw Data', is_summary=False)
            apply_excel_formatting(summary, writer, 'Summary', is_summary=True)
        
        excel_data = output.getvalue()

        st.success("Analysis Complete!")
        
        # Trigger Modal
        if st.button("📧 Send Report to Email"):
            modal.open()

        if modal.is_open():
            with modal.container():
                st.write("Enter the email address where you'd like to receive the report.")
                email_input = st.text_input("Recipient Email")
                
                if st.button("Send Now"):
                    if email_input:
                        with st.spinner("Sending..."):
                            success, message = send_analysis_email(
                                email_input, 
                                excel_data, 
                                "Property_Analysis.xlsx"
                            )
                            if success:
                                st.toast(message, icon="✅")
                                modal.close()
                            else:
                                st.error(message)
                    else:
                        st.warning("Please enter a valid email.")
