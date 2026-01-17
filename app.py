import streamlit as st
import pandas as pd
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
from processor import extract_area_logic, determine_config, apply_excel_formatting

# --- SMTP SETTINGS ---
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "atharvaujoshi@gmail.com"
SENDER_PASSWORD = "nybl zsnx zvdw edqr" 

def send_email(recipient_email, excel_data):
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = recipient_email
        msg['Subject'] = "Property Analysis Report"
        msg.attach(MIMEText("Please find your Real Estate Analysis Report attached."))

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
        st.error(f"Mail Dispatch Error: {e}")
        return False

@st.dialog("📧 Send Report to Email")
def email_popup(excel_bytes):
    st.write("Enter the email address where you'd like to receive the report.")
    user_email = st.text_input("Recipient Email")
    if st.button("Submit"):
        if user_email and "@" in user_email:
            with st.spinner("Sending email..."):
                if send_email(user_email, excel_bytes):
                    st.success(f"Report sent successfully to {user_email}")
        else:
            st.error("Please enter a valid email.")

# --- UI ---
st.set_page_config(page_title="Property Analyzer", layout="wide")
st.title("🏙️ Property Analyzer")

st.sidebar.header("Calculation Settings")
loading_factor = st.sidebar.number_input("Loading Factor", value=1.350, format="%.3f")
t1 = st.sidebar.number_input("1 BHK Threshold", value=600)
t2 = st.sidebar.number_input("2 BHK Threshold", value=850)
t3 = st.sidebar.number_input("3 BHK Threshold", value=1100)

uploaded_file = st.file_uploader("Upload Excel File", type="xlsx")

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    clean_cols = {c.lower().strip(): c for c in df.columns}
    desc_col = clean_cols.get('property description')
    cons_col = clean_cols.get('consideration value')
    prop_col = clean_cols.get('property')
    
    if desc_col and cons_col and prop_col:
        with st.spinner('Analyzing...'):
            df['Carpet Area (SQ.MT)'] = df[desc_col].apply(extract_area_logic)
            df['Carpet Area (SQ.FT)'] = (df['Carpet Area (SQ.MT)'] * 10.764).round(2)
            df['Saleable Area'] = (df['Carpet Area (SQ.FT)'] * loading_factor).round(2)
            df['APR'] = df.apply(lambda r: round(r[cons_col]/r['Saleable Area'], 2) if r['Saleable Area'] > 0 else 0, axis=1)
            df['Configuration'] = df['Carpet Area (SQ.FT)'].apply(lambda x: determine_config(x, t1, t2, t3))
            
            valid_df = df[df['Carpet Area (SQ.MT)'] > 0].sort_values([prop_col, 'Configuration'])
            
            # 1. Perform aggregation (Creates Multi-index)
            summary = valid_df.groupby([prop_col, 'Configuration', 'Carpet Area (SQ.FT)']).agg({
                'APR': ['min', 'max', 'mean']
            }).reset_index()

            # 2. FLATTEN THE COLUMNS TO FIX NotImplementedError
            summary.columns = ['Property', 'Configuration', 'Carpet Area (SQ.FT)', 'Min APR', 'Max APR', 'Avg APR']

            # Generate File in Memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                apply_excel_formatting(df, writer, 'Raw Data', is_summary=False)
                apply_excel_formatting(summary, writer, 'Summary', is_summary=True)
            excel_bytes = output.getvalue()

            st.success("Analysis Complete!")
            if st.button("📧 Get Report via Email"):
                email_popup(excel_bytes)
