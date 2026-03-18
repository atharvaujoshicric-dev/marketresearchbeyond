import streamlit as st
import pandas as pd
import re
import io
import smtplib
import json
from openai import OpenAI
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formataddr
from email import encoders
from openpyxl.styles import Alignment, PatternFill, Border, Side

# --- CONFIGURATION FROM SECRETS ---
SENDER_EMAIL = "atharvaujoshi@gmail.com"
SENDER_NAME = "Spydarr Market Research" 
APP_PASSWORD = st.secrets["EMAIL_PASSWORD"] # Store in Streamlit Secrets
OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"] # Store in Streamlit Secrets

client = OpenAI(api_key=OPENAI_API_KEY)

# --- AI EXTRACTION LOGIC ---
def extract_areas_with_ai(descriptions, batch_size=15):
    """Processes descriptions in batches using GPT-4o-mini for 100% accuracy."""
    all_extracted_values = []
    progress_bar = st.progress(0)
    total_batches = (len(descriptions) // batch_size) + 1

    for i in range(0, len(descriptions), batch_size):
        batch = descriptions[i:i + batch_size]
        
        prompt = f"""
        Extract the TOTAL Carpet Area in Square METERS from these {len(batch)} property descriptions.
        
        Rules:
        1. Look for 'Carpet Area', 'Chatai Kshetra', or 'RERA Area'.
        2. Sum 'Carpet' + 'Balcony' + 'Terrace' + 'Utility' if listed separately to get the total.
        3. If a value is only in Sq.Ft, convert it to Sq.Mt (Value / 10.764).
        4. Return ONLY a JSON list of floats in the same order as provided. 
        5. If no area is found, use 0.0.
        
        Descriptions:
        {json.dumps(batch)}
        """

        try:
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You are a real estate data specialist. You output only valid JSON lists of numbers."},
                    {"role": "user", "content": prompt}
                ],
                response_format={"type": "json_object"},
                temperature=0
            )
            res_content = json.loads(response.choices[0].message.content)
            # Handle different possible JSON key names from AI
            values = list(res_content.values())[0] if isinstance(res_content, dict) else res_content
            all_extracted_values.extend(values)
        except Exception as e:
            st.error(f"AI Batch Error: {e}")
            all_extracted_values.extend([0.0] * len(batch))
        
        progress_bar.progress(min((i + batch_size) / len(descriptions), 1.0))
    
    return all_extracted_values

# --- EMAIL LOGIC ---
def send_email(recipient_email, excel_data, filename):
    try:
        recipient_name = recipient_email.split('@')[0].replace('.', ' ').title()
        msg = MIMEMultipart()
        msg['From'] = formataddr((SENDER_NAME, SENDER_EMAIL))
        msg['To'] = recipient_email
        msg['Subject'] = "Spydarr Market Research Summary"
        
        body = f"Dear {recipient_name},\n\nPlease find the attached professional property analysis report.\n\nRegards,\nAtharva Joshi"
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
        st.error(f"Error sending email: {e}")
        return False

# --- UI & EXCEL FORMATTING ---
def determine_config(area, t1, t2, t3):
    if area <= 0: return "N/A"
    if area < t1: return "1 BHK"
    elif area < t2: return "2 BHK"
    elif area < t3: return "3 BHK"
    else: return "4 BHK"

def apply_excel_formatting(df, writer, sheet_name, is_summary=True):
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    colors = ["A2D2FF", "FFD6A5", "CAFFBF", "FDFFB6", "FFADAD", "BDB2FF", "9BF6FF"]
    
    for i in range(1, worksheet.max_row + 1):
        for j in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=i, column=j)
            cell.alignment = center_align
            if is_summary: cell.border = thin_border

    if is_summary:
        color_idx, start_row_prop, start_row_loc = 0, 2, 2
        last_col = len(df.columns)
        for i in range(2, len(df) + 3):
            curr_loc = df.iloc[i-2, 0] if i-2 < len(df) else None
            prev_loc = df.iloc[i-3, 0] if i-3 >= 0 else None
            curr_prop = df.iloc[i-2, 1] if i-2 < len(df) else None
            prev_prop = df.iloc[i-3, 1] if i-3 >= 0 else None
            
            if curr_prop != prev_prop and i > 2:
                fill = PatternFill(start_color=colors[color_idx % len(colors)], end_color=colors[color_idx % len(colors)], fill_type="solid")
                for r in range(start_row_prop, i):
                    for c in range(2, last_col + 1):
                        worksheet.cell(row=r, column=c).fill = fill
                if i-1 > start_row_prop:
                    worksheet.merge_cells(start_row=start_row_prop, start_column=2, end_row=i-1, end_column=2)
                start_row_prop, color_idx = i, color_idx + 1

            if curr_loc != prev_loc and i > 2:
                if i-1 > start_row_loc:
                    worksheet.merge_cells(start_row=start_row_loc, start_column=1, end_row=i-1, end_column=1)
                start_row_loc = i

# --- MAIN APP ---
st.set_page_config(page_title="Spydarr AI Dashboard", layout="wide")
st.title("Spydarr AI Dashboard 🚀")
st.info("Now powered by OpenAI for 100% extraction accuracy.")

st.sidebar.header("Calculation Settings")
loading_factor = st.sidebar.number_input("Loading Factor", value=1.35, format="%.3f")
t1, t2, t3 = st.sidebar.number_input("1BHK Thresh.", value=600), st.sidebar.number_input("2BHK Thresh.", value=850), st.sidebar.number_input("3BHK Thresh.", value=1100)

uploaded_file = st.file_uploader("Upload Data File", type=["xlsx", "csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
    clean_cols = {c.lower().strip(): c for c in df.columns}
    
    # Map columns
    desc_col = clean_cols.get('property description')
    cons_col = clean_cols.get('consideration value')
    prop_col = clean_cols.get('property')
    date_col = clean_cols.get('completion date')
    loc_col = clean_cols.get('micromarket')

    if all([desc_col, cons_col, prop_col, date_col, loc_col]):
        if st.button("Run AI Analysis"):
            with st.spinner('AI analyzing descriptions...'):
                descriptions = df[desc_col].astype(str).tolist()
                df['Carpet Area (SQ.MT)'] = extract_areas_with_ai(descriptions)
                
                df['Carpet Area (SQ.FT)'] = (df['Carpet Area (SQ.MT)'] * 10.764).round(3)
                df['Saleable Area'] = (df['Carpet Area (SQ.FT)'] * loading_factor).round(3)
                df['APR'] = df.apply(lambda r: round(r[cons_col]/r['Saleable Area'], 3) if r['Saleable Area'] > 0 else 0, axis=1)
                df['Configuration'] = df['Carpet Area (SQ.FT)'].apply(lambda x: determine_config(x, t1, t2, t3))
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                
                valid_df = df[df['Carpet Area (SQ.FT)'] > 0].sort_values([loc_col, prop_col, 'Configuration'])
                
                # Summary Logic
                summary = valid_df.groupby([loc_col, prop_col, 'Configuration', 'Carpet Area(SQ.FT)']).agg(
                    Last_Date=(date_col, 'max'), Min_APR=('APR', 'min'), Max_APR=('APR', 'max'),
                    Avg_APR=('APR', 'mean'), Median_APR=('APR', 'median'), Count=(prop_col, 'count')
                ).reset_index()
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    apply_excel_formatting(df, writer, 'Raw Data', is_summary=False)
                    apply_excel_formatting(summary, writer, 'Summary', is_summary=True)
                
                st.session_state['report'] = output.getvalue()
                st.success("Analysis Complete!")

        if 'report' in st.session_state:
            recipient = st.text_input("Recipient Username", placeholder="e.g. atharva.joshi")
            if st.button("Email Report") and recipient:
                full_email = f"{recipient.strip().lower()}@beyondwalls.com"
                if send_email(full_email, st.session_state['report'], "Spydarr_Market_Report.xlsx"):
                    st.success(f"Report sent to {full_email}")
    else:
        st.error("Missing required columns in uploaded file.")
