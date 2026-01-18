import streamlit as st
import pandas as pd
import re
import io
from streamlit_modal import Modal
from email_utils import send_analysis_email
from openpyxl.styles import Alignment, PatternFill, Border, Side

# --- [KEEP ALL YOUR LOGIC FUNCTIONS: extract_area_logic, determine_config, apply_excel_formatting HERE] ---

st.set_page_config(page_title="Real Estate Dashboard", layout="wide")

# Initialize Session State for the modal to prevent blinking
if "modal_open" not in st.session_state:
    st.session_state.modal_open = False

# Sidebar
st.sidebar.header("⚙️ Calculation Settings")
loading_factor = st.sidebar.number_input("Loading Factor", min_value=1.0, value=1.35, step=0.001, format="%.3f")
t1 = st.sidebar.number_input("1 BHK Threshold", value=600)
t2 = st.sidebar.number_input("2 BHK Threshold", value=850)
t3 = st.sidebar.number_input("3 BHK Threshold", value=1100)

st.title("🏙️ Real Estate Property Analysis")
uploaded_file = st.file_uploader("Upload your Excel file to begin", type="xlsx")

if uploaded_file:
    # ... [Your Processing Logic: df, valid_df, summary, excel_bin generation] ...
    # (Assuming excel_bin is generated here as in previous steps)

    st.divider()
    
    # Clean Action Bar
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("✉️ Send via Email", use_container_width=True):
            st.session_state.modal_open = True

    # Modal UI
    modal = Modal(key="email_popup", title="📫 Delivery Options")
    
    if st.session_state.modal_open:
        with modal.container():
            st.markdown("### Send Formatted Report")
            st.write("The report will be sent as a professional Excel attachment.")
            
            email_receiver = st.text_input("Recipient Email Address", placeholder="example@mail.com")
            
            # Using columns inside modal for better button alignment
            btn_col1, btn_col2 = st.columns(2)
            
            with btn_col1:
                if st.button("🚀 Send Email", use_container_width=True):
                    if email_receiver and "@" in email_receiver:
                        with st.spinner("Sending..."):
                            success, msg = send_analysis_email(email_receiver, excel_bin, "Property_Analysis.xlsx")
                            if success:
                                st.success("Email dispatched successfully!")
                                # Auto-close after a delay or keep open for confirmation
                            else:
                                st.error(msg)
                    else:
                        st.warning("Please enter a valid email address.")
            
            with btn_col2:
                if st.button("❌ Close", use_container_width=True):
                    st.session_state.modal_open = False
                    st.rerun()

# --- CSS to fix "Dirty" UI elements ---
st.markdown("""
    <style>
    /* Make the modal look cleaner */
    div[data-testid="stExpander"] { border: none !important; }
    .stButton button { border-radius: 8px; }
    </style>
    """, unsafe_allow_html=True)
