import streamlit as st
import pandas as pd
import re
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formataddr
from email import encoders
from openpyxl.styles import Alignment, PatternFill, Border, Side
from groq import Groq
import os
import json

# --- EMAIL CONFIGURATION ---
SENDER_EMAIL = "atharvaujoshi@gmail.com"
SENDER_NAME = "Spydarr Market Research"
APP_PASSWORD = "nybl zsnx zvdw edqr"

# --- GROQ CLIENT ---
# Reads GROQ_API_KEY from environment variable / Streamlit secrets
@st.cache_resource
def get_groq_client():
    api_key = os.environ.get("GROQ_API_KEY") or st.secrets.get("GROQ_API_KEY", None)
    if not api_key:
        st.error("GROQ_API_KEY not found. Add it to your .env or Streamlit secrets.")
        st.stop()
    return Groq(api_key=api_key)

GROQ_MODEL = "llama-3.3-70b-versatile"

SYSTEM_PROMPT = """You are an expert at reading Indian property registration documents written in Marathi and English.

Your ONLY job is to extract the FLAT/APARTMENT carpet area components and return their SUM in square meters (चौरस मी / sq.mt).

INCLUDE these area components (they belong to the flat itself):
- Carpet area: कारपेट, कार्पेट, कारपेट क्षेत्र, कार्पेट क्षेत्र, carpet area
- Balcony (any type attached to the flat): बाल्कनी, बालकनी, ओपन बाल्कनी, ओपन बालकनी, बाल्कनी एरिया, बालकनी एरिया, अटॅच बाल्कनी, एन्क्लोज बाल्कनी, enclosed balcony, open balcony, attached balcony
- Dry balcony: ड्राय बाल्कनी, ड्राय बालकनी
- Utility / utility balcony: युटिलिटी, युटिलिटी बालकनी, utility, utility balcony
- Attached terrace to the flat: लगतचे टेरेस, टेरेस (only when mentioned alongside flat components, NOT open-to-sky)

EXCLUDE these — they are NOT part of the flat's carpet area:
- Survey number land areas: स.नं., स. नं., सर्व्हे नं followed by hectare/are/चौ.मी. land extents
- Total land / plot area: एकूण क्षेत्र, प्लॉट क्षेत्र, यापैकी क्षेत्र (land subdivision)
- Open terrace / open to sky: ओपन टेरेस, ओपन टू स्काय
- Parking / car parking: पार्किंग, कार पार्किंग, पार्कींग, covered parking, कव्हर्ड कार पार्किंग — these always have a parking number or stall reference
- Road / reserved areas

CONVERSION RULE:
- Values given in चौ.फूट / चौ.फुट / sq.ft must be converted to sq.mt by dividing by 10.764
- Values already in चौ.मी. / चौ. मी. / sq.mt must be used directly
- Sometimes both units are given for the same component (e.g. "103.26 चौ.मी. व लगतेच बाल्कनी 12.59 चौ. मी.") — use the sq.mt value, do NOT double count

DOUBLE COUNTING RULE:
- If a value appears to equal the sum of other components already found, it is a stated total — do NOT add it again

IMPORTANT: Return ONLY a valid JSON object. No explanation, no markdown backticks, no extra text whatsoever.

JSON format:
{
  "components": [
    {"label": "carpet", "value_sqmt": 64.61},
    {"label": "open balcony", "value_sqmt": 5.99}
  ],
  "total_sqmt": 70.60
}

If no valid flat area is found, return exactly:
{"components": [], "total_sqmt": 0.0}
"""

def extract_area_groq(text: str, client: Groq, debug_log: list = None) -> float:
    """Use Groq LLM to extract and sum flat carpet area components from property description."""
    if pd.isna(text) or str(text).strip() == "":
        return 0.0

    raw = ""
    for attempt in range(2):  # retry once on failure
        try:
            response = client.chat.completions.create(
                model=GROQ_MODEL,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": str(text)[:4000]}
                ],
                temperature=0,
                max_tokens=512,
            )
            raw = response.choices[0].message.content.strip()

            # Strip markdown fences if model adds them
            raw_clean = re.sub(r"```json|```", "", raw).strip()

            # Extract only the JSON object (in case model adds preamble text)
            json_match = re.search(r'\{.*\}', raw_clean, re.DOTALL)
            if json_match:
                raw_clean = json_match.group(0)

            parsed = json.loads(raw_clean)
            total = float(parsed.get("total_sqmt", 0.0))

            if debug_log is not None:
                debug_log.append({
                    "description": str(text)[:120] + "...",
                    "raw_response": raw,
                    "components": parsed.get("components", []),
                    "total_sqmt": total
                })

            return round(total, 3)

        except json.JSONDecodeError:
            if attempt == 1:
                # Last resort: grab the last float from the raw output
                nums = re.findall(r'\d+\.?\d+', raw)
                fallback = round(float(nums[-1]), 3) if nums else 0.0
                if debug_log is not None:
                    debug_log.append({
                        "description": str(text)[:120] + "...",
                        "raw_response": raw,
                        "components": "JSON PARSE FAILED",
                        "total_sqmt": fallback
                    })
                return fallback
        except Exception as e:
            if attempt == 1:
                if debug_log is not None:
                    debug_log.append({
                        "description": str(text)[:120] + "...",
                        "raw_response": str(e),
                        "components": "ERROR",
                        "total_sqmt": 0.0
                    })
                return 0.0
    return 0.0


def send_email(recipient_email, excel_data, filename):
    try:
        recipient_name = recipient_email.split('@')[0].replace('.', ' ').title()
        msg = MIMEMultipart()
        msg['From'] = formataddr((SENDER_NAME, SENDER_EMAIL))
        msg['To'] = recipient_email
        msg['Subject'] = "Spydarr Market Research Summary"

        body = f"""Dear {recipient_name},

Please find the attached professional property analysis report generated by the dashboard.

The report includes:
1. Raw Data with calculated APR and Configurations.
2. A summarized view of APR statistics across properties.

Regards,
Atharva Joshi"""

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


def determine_config(area, t1, t2, t3):
    if area == 0:
        return "N/A"
    if area < t1:
        return "1 BHK"
    elif area < t2:
        return "2 BHK"
    elif area < t3:
        return "3 BHK"
    else:
        return "4 BHK"


def apply_excel_formatting(df, writer, sheet_name, is_summary=True):
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes = "A2"
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    colors = ["A2D2FF", "FFD6A5", "CAFFBF", "FDFFB6", "FFADAD", "BDB2FF", "9BF6FF"]

    for i in range(1, worksheet.max_row + 1):
        for j in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=i, column=j)
            cell.alignment = center_align
            if is_summary:
                cell.border = thin_border

    if is_summary:
        color_idx, start_row_prop = 0, 2
        start_row_loc = 2
        last_col = len(df.columns)
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        for i in range(2, len(df) + 3):
            curr_loc = df.iloc[i - 2, 0] if i - 2 < len(df) else None
            prev_loc = df.iloc[i - 3, 0] if i - 3 >= 0 else None
            curr_prop = df.iloc[i - 2, 1] if i - 2 < len(df) else None
            prev_prop = df.iloc[i - 3, 1] if i - 3 >= 0 else None

            if curr_prop != prev_prop and i > 2:
                fill = PatternFill(
                    start_color=colors[color_idx % len(colors)],
                    end_color=colors[color_idx % len(colors)],
                    fill_type="solid"
                )
                for r in range(start_row_prop, i):
                    for c in range(2, last_col + 1):
                        worksheet.cell(row=r, column=c).fill = fill
                if i - 1 > start_row_prop:
                    worksheet.merge_cells(start_row=start_row_prop, start_column=2, end_row=i - 1, end_column=2)
                    worksheet.merge_cells(start_row=start_row_prop, start_column=last_col, end_row=i - 1, end_column=last_col)
                start_row_prop = i
                color_idx += 1

            if curr_loc != prev_loc and i > 2:
                for r in range(start_row_loc, i):
                    worksheet.cell(row=r, column=1).fill = white_fill
                if i - 1 > start_row_loc:
                    worksheet.merge_cells(start_row=start_row_loc, start_column=1, end_row=i - 1, end_column=1)
                start_row_loc = i


# --- STREAMLIT UI ---
st.set_page_config(page_title="Spydarr Dashboard", layout="wide")
st.title("Spydarr Dashboard")
st.markdown(
    "<div style='margin-top: -15px; margin-bottom: 5px;'>"
    "<span style='background-color: #FFFF00; padding: 2px 8px; border-radius: 4px; "
    "border: 1px solid #E6E600; font-size: 0.9em; color: black;'>"
    "<u><strong>NOTE :-</strong> Please cross-check the report manually.</u></span></div>",
    unsafe_allow_html=True
)
st.markdown("[Property Report Tool · Streamlit](https://summarybeyondwalls.streamlit.app/)")
st.divider()

st.sidebar.header("Calculation Settings")
loading_factor = st.sidebar.number_input("Loading Factor", min_value=1.0, value=1.35, step=0.001, format="%.3f")
t1 = st.sidebar.number_input("1 BHK Threshold (sq.ft)", value=600)
t2 = st.sidebar.number_input("2 BHK Threshold (sq.ft)", value=850)
t3 = st.sidebar.number_input("3 BHK Threshold (sq.ft)", value=1100)

uploaded_file = st.file_uploader("Upload Data File", type=["xlsx", "csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
    clean_cols = {c.lower().strip(): c for c in df.columns}
    desc_col = clean_cols.get('property description')
    cons_col = clean_cols.get('consideration value')
    prop_col = clean_cols.get('property')
    date_col = clean_cols.get('completion date')
    loc_col = clean_cols.get('micromarket')

    if desc_col and cons_col and prop_col and date_col and loc_col:
        client = get_groq_client()

        with st.spinner('Extracting carpet areas via Groq LLM — this may take a moment...'):
            areas = []
            debug_log = []
            progress = st.progress(0, text="Processing rows...")
            total_rows = len(df)

            for idx, row in df.iterrows():
                area = extract_area_groq(row[desc_col], client, debug_log=debug_log)
                areas.append(area)
                progress.progress((idx + 1) / total_rows, text=f"Row {idx + 1} of {total_rows}")

            progress.empty()
            df['Carpet Area (SQ.MT)'] = areas

        with st.spinner('Calculating APR and generating report...'):
            df['Carpet Area (SQ.FT)'] = (df['Carpet Area (SQ.MT)'] * 10.764).round(3)
            df['Saleable Area'] = (df['Carpet Area (SQ.FT)'] * loading_factor).round(3)
            df['APR'] = df.apply(
                lambda r: round(r[cons_col] / r['Saleable Area'], 3) if r['Saleable Area'] > 0 else 0, axis=1
            )
            df['Configuration'] = df['Carpet Area (SQ.FT)'].apply(lambda x: determine_config(x, t1, t2, t3))
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')

            valid_df = df[df['Carpet Area (SQ.FT)'] > 0].sort_values(
                [loc_col, prop_col, 'Configuration', 'Carpet Area (SQ.FT)']
            )

            project_counts = valid_df.groupby(prop_col).size().reset_index(name='Total Count')
            summary = valid_df.groupby(
                [loc_col, prop_col, 'Configuration', 'Carpet Area (SQ.FT)']
            ).agg(
                Last_Date=(date_col, 'max'),
                Min_APR=('APR', 'min'),
                Max_APR=('APR', 'max'),
                Avg_APR=('APR', 'mean'),
                Median_APR=('APR', 'median'),
                Property_Count=(prop_col, 'count')
            ).reset_index()

            summary = summary.merge(project_counts, on=prop_col, how='left')
            summary['Last_Date'] = pd.to_datetime(summary['Last_Date'], errors='coerce')
            summary['Last_Date'] = summary['Last_Date'].apply(
                lambda x: x.strftime('%b-%Y') if pd.notnull(x) else "N/A"
            )

            summary.columns = [
                'Location', 'Property', 'Configuration', 'Carpet Area(SQ.FT)',
                'Last Completion Date', 'Min. APR', 'Max APR',
                'Average of APR', 'Median of APR', 'Count of Property', 'Total Count'
            ]
            summary = summary[[
                'Location', 'Property', 'Last Completion Date', 'Configuration',
                'Carpet Area(SQ.FT)', 'Min. APR', 'Max APR',
                'Average of APR', 'Median of APR', 'Count of Property', 'Total Count'
            ]]

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                apply_excel_formatting(df, writer, 'Raw Data', is_summary=False)
                apply_excel_formatting(summary, writer, 'Summary', is_summary=True)

        st.success("✅ Analysis Complete!")

        # Preview extracted areas
        with st.expander("🔍 Preview: Carpet Area Extraction (first 10 rows)"):
            preview_cols = [desc_col, 'Carpet Area (SQ.MT)', 'Carpet Area (SQ.FT)']
            st.dataframe(df[preview_cols].head(10), use_container_width=True)

        # Debug view — shows LLM raw response for each row, useful for spotting wrong extractions
        with st.expander("🐛 Debug: LLM Responses (check rows with 0 area)"):
            zero_rows = [d for d in debug_log if d['total_sqmt'] == 0.0]
            st.markdown(f"**Total rows processed:** {len(debug_log)} | **Rows with 0 area:** {len(zero_rows)}")
            for entry in debug_log:
                color = "🔴" if entry['total_sqmt'] == 0.0 else "🟢"
                with st.container():
                    st.markdown(f"{color} **Total:** {entry['total_sqmt']} sq.mt")
                    st.caption(entry['description'])
                    st.code(entry['raw_response'], language='json')
                    st.divider()

        recipient = st.text_input("Recipient Name", placeholder="firstname.lastname")
        if st.button("Send to Email") and recipient:
            full_email = f"{recipient.strip().lower()}@beyondwalls.com"
            if send_email(full_email, output.getvalue(), "Spydarr_Market_Report.xlsx"):
                st.success(f"✅ Report sent to {full_email}")

        st.download_button(
            label="⬇️ Download Report",
            data=output.getvalue(),
            file_name="Spydarr_Market_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error(
            "Missing required columns. Ensure file has: "
            "'Micromarket', 'Property Description', 'Consideration Value', 'Property', 'Completion Date'."
        )
