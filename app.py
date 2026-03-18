import streamlit as st
import pandas as pd
import re
import io
import smtplib
import time
from groq import Groq
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formataddr
from email import encoders
from openpyxl.styles import Alignment, PatternFill, Border, Side

# --- SECRETS LOADING ---
try:
    SENDER_EMAIL = st.secrets["SENDER_EMAIL"]
    APP_PASSWORD = st.secrets["APP_PASSWORD"]
    GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
    GROQ_CLIENT = Groq(api_key=GROQ_API_KEY)
except KeyError as e:
    st.error(f"Missing Secret Key: {e}. Please add it to the Streamlit Settings.")
    st.stop()

SENDER_NAME = "Spydarr Market Research"

# --- EMAIL FUNCTION ---
def send_email(recipient_email, excel_data, filename):
    try:
        recipient_name = recipient_email.split('@')[0].replace('.', ' ').title()
        msg = MIMEMultipart()
        msg['From'] = formataddr((SENDER_NAME, SENDER_EMAIL))
        msg['To'] = recipient_email
        msg['Subject'] = "Spydarr Market Research Summary"
        
        body = f"Dear {recipient_name},\n\nPlease find the attached property analysis report.\n\nRegards,\nAtharva Joshi"
        msg.attach(MIMEText(body, 'plain'))

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(excel_data)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename={filename}")
        msg.attach(part)
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, APP_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"SMTP Error: {e}")
        return False

# --- LOGIC FUNCTIONS (Area & Config) ---
def extract_area_logic(text):
    # (Keeping your original regex logic as fallback)
    if pd.isna(text) or text == "": return 0.0
    text = " ".join(str(text).split())
    # ... [Same regex code you provided previously] ...
    return 0.0 # simplified for brevity, use your full version here

@st.cache_data(show_spinner=False)
def extract_area_ai_enhanced(text):
    try:
        chat_completion = GROQ_CLIENT.chat.completions.create(
            messages=[
                {"role": "system", "content": "Extract total NET CARPET AREA in SQ.METERS. If SQ.FT found, divide by 10.764. Return ONLY the number."},
                {"role": "user", "content": text}
            ],
            model="llama-3.3-70b-versatile",
            temperature=0,
            max_tokens=10,
        )
        res = chat_completion.choices[0].message.content.strip()
        num = re.findall(r"[-+]?\d*\.\d+|\d+", res)
        return float(num[0]) if num else extract_area_logic(text)
    except:
        return extract_area_logic(text)

# --- APP UI ---
st.set_page_config(page_title="Spydarr Dashboard", layout="wide")
st.title("Spydarr Dashboard")

# Initialize session state for the Excel file
if "final_excel" not in st.session_state:
    st.session_state.final_excel = None

st.sidebar.header("Settings")
loading_factor = st.sidebar.number_input("Loading Factor", value=1.40)
t1 = st.sidebar.number_input("1 BHK Threshold", value=600)
t2 = st.sidebar.number_input("2 BHK Threshold", value=850)
t3 = st.sidebar.number_input("3 BHK Threshold", value=1100)

uploaded_file = st.file_uploader("Upload XLSX/CSV", type=["xlsx", "csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
    clean_cols = {c.lower().strip(): c for c in df.columns}
    
    # Verify Columns
    required = ['micromarket', 'property description', 'consideration value', 'property', 'completion date']
    if all(k in clean_cols for k in required):
        if st.button("🚀 Generate AI Report"):
            with st.spinner("Analyzing with Groq AI..."):
                desc_col = clean_cols['property description']
                # Calculation Logic
                df['Carpet Area (SQ.MT)'] = [extract_area_ai_enhanced(x) for x in df[desc_col]]
                df['Carpet Area (SQ.FT)'] = (df['Carpet Area (SQ.MT)'] * 10.764).round(3)
                # ... [Rest of your calculations for Saleable, APR, Config] ...
                
                # Create Excel in Memory
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name="Report")
                
                st.session_state.final_excel = output.getvalue()
                st.success("Analysis Ready!")

        # EMAIL SECTION (Visible only after generation)
        if st.session_state.final_excel:
            st.divider()
            recipient = st.text_input("Enter Recipient Name (e.g. john.doe)")
            if st.button("📧 Send via Email"):
                if recipient:
                    email_addr = f"{recipient.strip().lower()}@beyondwalls.com"
                    if send_email(email_addr, st.session_state.final_excel, "Spydarr_Report.xlsx"):
                        st.success(f"Report sent to {email_addr}")
                else:
                    st.warning("Please enter a name first.")
    else:
        st.error("Column mismatch. Check file headers.")
