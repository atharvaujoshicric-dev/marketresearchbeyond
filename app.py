import streamlit as st
import pandas as pd
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from processor import extract_area_logic, determine_config, apply_excel_formatting

# --- SMTP SETTINGS (Atharva's Credentials) ---
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "atharvaujoshi@gmail.com"
SENDER_PASSWORD = "nybl zsnx zvdw edqr" 

def send_email(recipient_email, excel_data):
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = recipient_email
        msg['Subject'] = "Property Analysis Report - Professional"
        
        # Adding a simple body text
        from email.mime.text import MIMEText
        msg.attach(MIMEText("Please find the requested Real Estate Property Analysis Report attached."))

        part = MIMEBase('application', "octet-stream")
        part.set_payload(excel_data)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="Property_Report.xlsx"')
        msg.attach(part)

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Mail Error: {e}")
        return False

@st.dialog("📧 Send Report via Email")
def email_popup(excel_bytes):
    st.write("Enter the email address where you'd like to receive the analysis.")
    user_email = st.text_input("Recipient Email Address", placeholder="example@domain.com")
    
    if st.button("Send Report Now"):
        if user_email and "@" in user_email:
            with st.spinner("Processing email dispatch..."):
                if send_email(user_email, excel_bytes):
                    st.success(f"Report sent successfully to {user_email}!")
                    # Dialog stays open to show success, user can close it manually
        else:
            st.error("Please provide a valid email address.")

# --- STREAMLIT UI ---
st.set_page_config(page_title="Real Estate Dashboard", layout="wide")
st.title("🏙️ Property Area & APR Analyzer")

st.sidebar.header("Calculation Settings")
loading_factor = st.sidebar.number_input("Loading Factor", value=1.350, step=0.005, format="%.3f")
t1 = st.sidebar.number_input("1 BHK Threshold (sqft)", value=600)
t2 = st.sidebar.number_input("2 BHK Threshold (sqft)", value=850)
t3 = st.sidebar.number_input("3 BHK Threshold (sqft)", value=1100)

uploaded_file = st.file_uploader("Upload Property Excel (.xlsx)", type="xlsx")

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    clean_cols = {c.lower().strip(): c for c in df.columns}
    
    desc_col = clean_cols.get('property description')
    cons_col = clean_cols.get('consideration value')
    prop_col = clean_cols.get('property')
    
    if desc_col and cons_col and prop_col:
        with st.spinner('Calculating Areas...'):
            df['Carpet Area (SQ.MT)'] = df[desc_col].apply(extract_area_logic)
            df['Carpet Area (SQ.FT)'] = (df['Carpet Area (SQ.MT)'] * 10.764).round(2)
            df['Saleable Area'] = (df['Carpet Area (SQ.FT)'] * loading_factor).round(2)
            df['APR'] = df.apply(lambda r: round(r[cons_col]/r['Saleable Area'], 2) if r['Saleable Area'] > 0 else 0, axis=1)
            df['Configuration'] = df['Carpet Area (SQ.FT)'].apply(lambda x: determine_config(x, t1, t2, t3))
            
            # Filter and Sort
            valid_df = df[df['Carpet Area (SQ.MT)'] > 0].sort_values([prop_col, 'Configuration'])
            summary = valid_df.groupby([prop_col, 'Configuration', 'Carpet Area (SQ.FT)']).agg({
                'APR': ['min', 'max', 'mean', 'count']
            }).reset_index()

            # Create Excel in memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                apply_excel_formatting(df, writer, 'Raw Data', is_summary=False)
                apply_excel_formatting(summary, writer, 'Summary', is_summary=True)
            excel_bytes = output.getvalue()

            st.success("Analysis Complete!")
            
            if st.button("📧 Get Formatted Report via Email"):
                email_popup(excel_bytes)
                
            st.divider()
            st.subheader("Preview (Summary Data)")
            st.dataframe(summary.head(20), use_container_width=True)
    else:
        st.error("Column mapping failed. Please ensure Excel has 'Property', 'Property Description', and 'Consideration Value'.")
