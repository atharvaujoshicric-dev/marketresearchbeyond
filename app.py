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
import time
import math

# --- EMAIL CONFIGURATION ---
SENDER_EMAIL = "atharvaujoshi@gmail.com"
SENDER_NAME = "Spydarr Market Research"
APP_PASSWORD = "nybl zsnx zvdw edqr"

# --- GROQ CLIENT ---
@st.cache_resource
def get_groq_client():
    api_key = os.environ.get("GROQ_API_KEY") or st.secrets.get("GROQ_API_KEY", None)
    if not api_key:
        st.error("GROQ_API_KEY not found. Add it to your .env or Streamlit secrets.")
        st.stop()
    return Groq(api_key=api_key)

GROQ_MODEL = "llama-3.3-70b-versatile"

# How many descriptions to send in one Groq call.
# Each description ~300-500 tokens. Batch of 20 = ~10K tokens per call.
# Free tier = 100K TPD, so ~10 batches/day max. Adjust down if still hitting limits.
BATCH_SIZE = 20

SYSTEM_PROMPT = """You are an expert at reading Indian property registration documents in Marathi and English.

You will receive a JSON array of property descriptions, each with an "id" and "text".
For EACH description, extract ONLY the flat/apartment carpet area components and return their sum in square meters.

INCLUDE (parts of the flat):
- Carpet area: कारपेट, कार्पेट, carpet area
- Any balcony: बाल्कनी, बालकनी, ओपन बाल्कनी, ओपन बालकनी, अटॅच बाल्कनी, एन्क्लोज बाल्कनी, लगतेच बाल्कनी, बाल्कनी एरिया, बालकनी एरिया
- Dry balcony: ड्राय बाल्कनी, ड्राय बालकनी
- Utility: युटिलिटी, युटिलिटी बालकनी
- Attached terrace: लगतचे टेरेस, टेरेस (ONLY if listed with carpet/balcony, NOT open-to-sky)

EXCLUDE (not flat area):
- Survey land: anything after स.नं./सर्व्हे नं. with हे/आर/hectare or large areas >500 चौ.मी.
- Land totals: एकूण क्षेत्र, यापैकी क्षेत्र for land
- Open terrace/sky: ओपन टेरेस, ओपन टू स्काय
- Parking: पार्किंग, कार पार्किंग, पार्कींग, कव्हर्ड कार पार्किंग (has stall number)

UNITS:
- चौ.मी./sq.mt → use directly
- चौ.फूट/चौ.फुट/चौ.फु/sq.ft → divide by 10.764
- If both units given for same item, use चौ.मी. only — do NOT count twice

ANTI-DOUBLE-COUNT: If one value = sum of others already listed, skip it.

Return ONLY a raw JSON array — no markdown, no explanation:
[
  {"id": 0, "components": [{"label": "carpet", "value_sqmt": 64.61}, {"label": "balcony", "value_sqmt": 5.99}], "total_sqmt": 70.60},
  {"id": 1, "components": [], "total_sqmt": 0.0}
]"""


def extract_areas_batch(descriptions: list, client: Groq) -> dict:
    """
    Send a batch of descriptions to Groq in one call.
    descriptions: list of (index, text) tuples
    Returns: dict of {index: total_sqmt}
    """
    payload = [{"id": idx, "text": str(text)[:1500]} for idx, text in descriptions]
    user_msg = json.dumps(payload, ensure_ascii=False)

    for attempt in range(3):
        try:
            response = client.chat.completions.create(
                model=GROQ_MODEL,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_msg}
                ],
                temperature=0,
                max_tokens=1024,
            )
            raw = response.choices[0].message.content.strip()
            clean = re.sub(r"```json|```", "", raw).strip()

            # Extract JSON array from response
            arr_match = re.search(r'\[.*\]', clean, re.DOTALL)
            if arr_match:
                clean = arr_match.group(0)

            parsed = json.loads(clean)
            return {item["id"]: round(float(item.get("total_sqmt", 0.0)), 3) for item in parsed}

        except json.JSONDecodeError:
            if attempt < 2:
                time.sleep(2)
            else:
                # Return 0 for all in this batch on total failure
                return {idx: 0.0 for idx, _ in descriptions}

        except Exception as e:
            err_str = str(e)
            if "429" in err_str or "rate_limit" in err_str.lower():
                # Parse wait time from error if possible
                wait = 60
                m = re.search(r'try again in (\d+)m(\d+)', err_str)
                if m:
                    wait = int(m.group(1)) * 60 + int(m.group(2)) + 5
                else:
                    m2 = re.search(r'try again in (\d+\.?\d*)s', err_str)
                    if m2:
                        wait = math.ceil(float(m2.group(1))) + 2

                st.warning(f"⏳ Rate limit hit. Waiting {wait}s before retrying...")
                time.sleep(wait)
            else:
                if attempt < 2:
                    time.sleep(2)
                else:
                    return {idx: 0.0 for idx, _ in descriptions}

    return {idx: 0.0 for idx, _ in descriptions}


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
    if area == 0: return "N/A"
    if area < t1: return "1 BHK"
    elif area < t2: return "2 BHK"
    elif area < t3: return "3 BHK"
    else: return "4 BHK"


def apply_excel_formatting(df, writer, sheet_name, is_summary=True):
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes = "A2"
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
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
            curr_loc = df.iloc[i-2, 0] if i-2 < len(df) else None
            prev_loc = df.iloc[i-3, 0] if i-3 >= 0 else None
            curr_prop = df.iloc[i-2, 1] if i-2 < len(df) else None
            prev_prop = df.iloc[i-3, 1] if i-3 >= 0 else None

            if curr_prop != prev_prop and i > 2:
                fill = PatternFill(start_color=colors[color_idx % len(colors)],
                                   end_color=colors[color_idx % len(colors)], fill_type="solid")
                for r in range(start_row_prop, i):
                    for c in range(2, last_col + 1):
                        worksheet.cell(row=r, column=c).fill = fill
                if i-1 > start_row_prop:
                    worksheet.merge_cells(start_row=start_row_prop, start_column=2,
                                          end_row=i-1, end_column=2)
                    worksheet.merge_cells(start_row=start_row_prop, start_column=last_col,
                                          end_row=i-1, end_column=last_col)
                start_row_prop = i
                color_idx += 1

            if curr_loc != prev_loc and i > 2:
                for r in range(start_row_loc, i):
                    worksheet.cell(row=r, column=1).fill = white_fill
                if i-1 > start_row_loc:
                    worksheet.merge_cells(start_row=start_row_loc, start_column=1,
                                          end_row=i-1, end_column=1)
                start_row_loc = i


# ─────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────
st.set_page_config(page_title="Spydarr Dashboard", layout="wide")
st.title("Spydarr Dashboard")
st.markdown(
    "<div style='margin-top:-15px;margin-bottom:5px;'>"
    "<span style='background-color:#FFFF00;padding:2px 8px;border-radius:4px;"
    "border:1px solid #E6E600;font-size:0.9em;color:black;'>"
    "<u><strong>NOTE :-</strong> Please cross-check the report manually.</u></span></div>",
    unsafe_allow_html=True
)
st.markdown("[Property Report Tool · Streamlit](https://summarybeyondwalls.streamlit.app/)")
st.divider()

# ── Sidebar ──────────────────────────────
st.sidebar.header("Calculation Settings")
loading_factor = st.sidebar.number_input("Loading Factor", min_value=1.0, value=1.35, step=0.001, format="%.3f")
t1 = st.sidebar.number_input("1 BHK Threshold (sq.ft)", value=600)
t2 = st.sidebar.number_input("2 BHK Threshold (sq.ft)", value=850)
t3 = st.sidebar.number_input("3 BHK Threshold (sq.ft)", value=1100)

st.sidebar.divider()
st.sidebar.header("🔧 Groq API Test")
st.sidebar.caption("Run this first to confirm your API key and batching works.")
if st.sidebar.button("▶ Test Groq API (2 descriptions)"):
    test_batch = [
        (0, 'फ्लॅट नं. 402, कारपेट क्षेत्र 64.61 चौ. मी., ओपन बाल्कनी क्षेत्र 5.99 चौ.मी., कव्हर्ड कार पार्किंग नं.(जी एल - 60), क्षेत्र 12.50 चौ. मी.'),
        (1, 'फ्लॅट नं. 602, कारपेट एरिया 58.10 चौ मी, बालकनी एरिया 7.73 चौ मी, युटिलिटी बालकनी 2.53 चौ मी, कव्हर्ड कार पार्किंग सह'),
    ]
    try:
        test_client = get_groq_client()
        results = extract_areas_batch(test_batch, test_client)
        st.sidebar.success(f"✅ API OK!\nRow 0: {results.get(0)} sq.mt (expected 70.60)\nRow 1: {results.get(1)} sq.mt (expected 68.36)")
    except Exception as e:
        st.sidebar.error(f"❌ Error: {e}")

st.sidebar.divider()
batch_size = st.sidebar.number_input(
    "Batch Size (rows per Groq call)",
    min_value=5, max_value=50, value=BATCH_SIZE, step=5,
    help="Higher = fewer API calls but more tokens per call. Lower = safer if hitting limits."
)

# ── File Upload ───────────────────────────
uploaded_file = st.file_uploader("Upload Data File", type=["xlsx", "csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
    clean_cols = {c.lower().strip(): c for c in df.columns}
    desc_col = clean_cols.get('property description')
    cons_col = clean_cols.get('consideration value')
    prop_col = clean_cols.get('property')
    date_col = clean_cols.get('completion date')
    loc_col  = clean_cols.get('micromarket')

    if desc_col and cons_col and prop_col and date_col and loc_col:
        client = get_groq_client()
        total_rows = len(df)
        descriptions = list(df[desc_col].items())  # list of (index, text)

        # Split into batches
        batches = [descriptions[i:i + batch_size] for i in range(0, total_rows, batch_size)]
        total_batches = len(batches)

        st.info(f"📦 {total_rows} rows → {total_batches} batches of ~{batch_size} (1 Groq call per batch)")

        with st.spinner(f'Sending {total_batches} batch(es) to Groq...'):
            all_results = {}
            progress = st.progress(0, text="Starting...")

            for b_idx, batch in enumerate(batches):
                progress.progress(
                    (b_idx) / total_batches,
                    text=f"Batch {b_idx + 1} of {total_batches} ({len(batch)} rows)..."
                )
                batch_results = extract_areas_batch(batch, client)
                all_results.update(batch_results)
                if b_idx < total_batches - 1:
                    time.sleep(0.5)  # small pause between batches

            progress.progress(1.0, text="Done!")
            time.sleep(0.3)
            progress.empty()

        # Map results back to df
        df['Carpet Area (SQ.MT)'] = df.index.map(lambda i: all_results.get(i, 0.0))

        with st.spinner('Calculating APR and generating report...'):
            df['Carpet Area (SQ.FT)'] = (df['Carpet Area (SQ.MT)'] * 10.764).round(3)
            df['Saleable Area'] = (df['Carpet Area (SQ.FT)'] * loading_factor).round(3)
            df['APR'] = df.apply(
                lambda r: round(r[cons_col] / r['Saleable Area'], 3) if r['Saleable Area'] > 0 else 0,
                axis=1
            )
            df['Configuration'] = df['Carpet Area (SQ.FT)'].apply(
                lambda x: determine_config(x, t1, t2, t3)
            )
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

        zero_count = sum(1 for v in all_results.values() if v == 0.0)
        st.success(f"✅ Analysis Complete! {total_rows - zero_count}/{total_rows} rows extracted successfully.")

        if zero_count > 0:
            st.warning(f"⚠️ {zero_count} rows returned 0 — expand preview below to inspect.")

        with st.expander("🔍 Preview: Carpet Area Extraction (first 20 rows)"):
            preview_cols = [desc_col, 'Carpet Area (SQ.MT)', 'Carpet Area (SQ.FT)']
            st.dataframe(df[preview_cols].head(20), use_container_width=True)

        with st.expander(f"🔍 Rows with 0 area ({zero_count} rows)"):
            zero_df = df[df['Carpet Area (SQ.MT)'] == 0.0][[desc_col, 'Carpet Area (SQ.MT)']]
            if len(zero_df) > 0:
                st.dataframe(zero_df, use_container_width=True)
            else:
                st.success("No zero rows!")

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
