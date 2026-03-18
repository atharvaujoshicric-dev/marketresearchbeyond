import streamlit as st
import pandas as pd
import re
import io
import smtplib
import json
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formataddr
from email import encoders
from openpyxl.styles import Alignment, PatternFill, Border, Side
from groq import Groq

# --- EMAIL CONFIGURATION ---
SENDER_EMAIL = "atharvaujoshi@gmail.com"
SENDER_NAME = "Spydarr Market Research"
APP_PASSWORD = "nybl zsnx zvdw edqr"

# --- GROQ CLIENT ---
@st.cache_resource
def get_groq_client():
    return Groq(api_key=st.secrets["GROQ_API_KEY"])

# ─────────────────────────────────────────────
# GROQ CROSS-VALIDATION
# ─────────────────────────────────────────────
def groq_validate(description: str, regex_area: float, regex_config: str, t1: int, t2: int, t3: int) -> dict:
    """
    Sends the raw property description + regex results to Groq.
    Returns {"area": float, "config": str, "changed": bool}
    """
    client = get_groq_client()

    system_prompt = """You are an expert Indian real estate data extraction assistant. 
You specialize in reading property registration documents written in Marathi and English.

Your job:
1. Extract the CARPET AREA of the residential flat/apartment described.
2. Classify it as 1 BHK, 2 BHK, 3 BHK, or 4 BHK based on the thresholds given.

Rules for area extraction:
- Return area in SQUARE METERS (sq.mt).
- Look for carpet area / चटई क्षेत्र. Prefer it over built-up or super built-up area.
- If multiple component areas are listed (e.g. hall + bedroom + kitchen), SUM them.
- If a total area is explicitly mentioned and matches the sum of components, use the total.
- Ignore parking dimensions, plot area, road widths, and land area.
- Ignore values with keywords: पार्किंग, parking, प्लॉट, plot, राखीव, reserve.
- Valid carpet area range: 20 sq.mt to 900 sq.mt. Return 0 if no valid area found.
- Convert sq.ft to sq.mt by dividing by 10.764 if needed.

Rules for configuration:
- Use the area in SQ.FT (area_sqmt * 10.764) and the thresholds provided.
- Below t1 sq.ft → 1 BHK
- t1 to t2 sq.ft → 2 BHK  
- t2 to t3 sq.ft → 3 BHK
- t3 and above → 4 BHK
- If area is 0 → N/A

You will be given:
- The raw property description text
- The regex-extracted area (sq.mt) and config — cross-check these
- The BHK thresholds

Respond ONLY with a valid JSON object, no explanation, no markdown:
{"area_sqmt": <float>, "config": "<string>", "reasoning": "<one line>"}"""

    user_prompt = f"""Property Description:
{description}

Regex extracted area: {regex_area} sq.mt
Regex extracted config: {regex_config}

BHK Thresholds (in sq.ft):
- 1 BHK: below {t1}
- 2 BHK: {t1} to {t2}
- 3 BHK: {t2} to {t3}
- 4 BHK: {t3} and above

Cross-validate the regex results. If they are correct, return the same values.
If you find a more accurate area or config, return the corrected values.
Return JSON only."""

    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.0,
            max_tokens=200,
        )
        raw = response.choices[0].message.content.strip()
        # Strip markdown fences if present
        raw = re.sub(r"^```(?:json)?|```$", "", raw, flags=re.MULTILINE).strip()
        result = json.loads(raw)

        groq_area = float(result.get("area_sqmt", regex_area))
        groq_config = str(result.get("config", regex_config)).strip()

        # Sanity check: only accept if in valid range
        if not (0.0 <= groq_area < 900):
            groq_area = regex_area
        if groq_config not in ["1 BHK", "2 BHK", "3 BHK", "4 BHK", "N/A"]:
            groq_config = regex_config

        changed = (
            abs(groq_area - regex_area) > 0.5 or
            groq_config != regex_config
        )
        return {"area": round(groq_area, 3), "config": groq_config, "changed": changed,
                "reasoning": result.get("reasoning", "")}

    except Exception as e:
        # On any failure, fall back to regex result silently
        return {"area": regex_area, "config": regex_config, "changed": False, "reasoning": f"Groq error: {e}"}


# ─────────────────────────────────────────────
# REGEX LOGIC
# ─────────────────────────────────────────────

# Keywords that label apartment-level components (carpet, balcony, terrace attached to a flat)
APARTMENT_COMPONENT_KEYWORDS = [
    "कारपेट", "कार्पेट", "carpet",
    "बाल्कनी", "balcony",
    "टेरेस", "terrace",
    "ड्राय बाल्कनी", "dry balcony",
    "सदनिका",          # flat/unit
    "युटिलिटी",        # utility area
]

def _is_survey_land_value(context_before: str) -> bool:
    """
    Returns True if the value is a survey-number land parcel area.
    Patterns like: स.नं.20/9 क्षेत्र 3000  or  स.नं. 20/14 क्षेत्र 5000
    """
    # Match स.नं followed (within ~30 chars) by क्षेत्र right before the number
    survey_pattern = re.search(
        r'(?:स\.?\s*नं\.?|survey\s*no\.?|s\.no\.?)\s*[\d/]+\s*(?:क्षेत्र|area)?\s*$',
        context_before.strip(),
        re.IGNORECASE
    )
    return survey_pattern is not None


def _is_yatun_pakhi_total(context_before: str, full_text: str, val: float) -> bool:
    """
    Returns True if this value follows 'यापैकी' or 'यांसी एकूण क्षेत्र' in a land-parcel
    context (i.e. a sub-division total, not the flat area).
    We detect this when the value is ≥ 1000 sq.mt AND 'यापैकी' appears nearby.
    """
    if val < 200:
        return False
    combined = context_before[-120:].lower()
    if "यापैकी" in combined or "यापैकी" in full_text[:20].lower():
        return True
    return False


def extract_area_logic(text):
    if pd.isna(text) or text == "": return 0.0

    text = " ".join(str(text).split())
    text = re.sub(r'म्हणज[च]े', 'म्हणजे', text)
    text = re.sub(r'(\d+)\.\.(\d+)', r'\1.\2', text)
    text = re.sub(r'(\d+)\s*\.\s*(\d+)', r'\1.\2', text)
    text = re.sub(r'(\d+\.\d+)\.', r'\1', text)
    text = re.sub(r'(\d+\.?)\s+(\d+)', r'\1\2', text)
    text = re.sub(r'(\d),(\d)', r'\1\2', text)
    text = re.sub(r'\d+\.?\d*\s*[\*x]\s*\d+\.?\d*', 'PARKING_DIM', text)

    m_unit = r'(?:चौरस\s*मी(?:[टत]र)?|चौ[\.\s]*मी[\.\s]*|चाै[\.\s]*मी[\.\s]*|sq\.?\s*m(?:tr)?\.?|square\s*meter(?:s)?)(?:\s*(?:कारपेट|कार्पेट|चटई क्षेत्र|एकूण क्षेत्र))?(?:\s*(?:एरिया|area|क्षेत्र))?'
    f_unit = r'(?:चौरस\s*फु[टत]|चौरस\s*फू[टत]|चौ[\.\s]*फु[टत]?|चौ[\.\s]*फू[टत]?|sq\.?\s*f(?:t)?\.?|square\s*f(?:ee|oo)t)(?:\s*(?:area|क्षेत्र))?'

    boundary_keywords = r'(?:येथील|मिळकतीवर|मिळकतीवरील|बांधण्यात|बांधत|प्रकल्पातील|गृहप्रकल्पातील|इमारतप्रकल्पातील|योजनेतील|नियोजित|इमारतीमधील|बिल्डींग|बिल्डिंग|प्रकल्प|टावर|टॉवर|प्रिस्टीन|सेक्टर|क्लस्टर)'
    parts = re.split(boundary_keywords, text, flags=re.IGNORECASE)
    relevant_text = " ".join(parts[1:]) if len(parts) > 1 else text

    exclude_keywords = [
        "पार्किंग", "पार्कींग", "parking", "road", "reserve", "राखीव",
        "प्लॉट", "plot", "वाढीव", "अविभक्त", "साईज", "size",
        "बिल्डअप", "मुल्यांकन", "दर", "rate", "७/१२", "नाकाश",
        # NOTE: "पैकी" removed from here — handled separately via _is_yatun_pakhi_total
    ]

    # ── PASS 1: collect apartment-component values (high confidence) ──────────
    # These are values immediately followed by apartment-level keywords in context_after.
    apartment_vals = []
    for match in re.finditer(rf'(\d+\.?\d*)\s?{m_unit}', relevant_text, re.IGNORECASE):
        val = float(match.group(1))
        if not (2.0 <= val < 500):   # apartment components are never > 500 sq.mt
            continue
        start_idx = match.start()
        end_idx   = match.end()
        context_before = relevant_text[max(0, start_idx - 80):start_idx]
        context_after  = relevant_text[end_idx:end_idx + 60]
        combined_ctx   = (context_before + context_after).lower()

        # Must have an apartment-component keyword nearby
        if any(kw.lower() in combined_ctx for kw in APARTMENT_COMPONENT_KEYWORDS):
            # Must NOT be a survey-land parcel
            if not _is_survey_land_value(context_before):
                if not apartment_vals or val != apartment_vals[-1]:
                    apartment_vals.append(val)

    if apartment_vals:
        return round(sum(apartment_vals), 3)

    # ── PASS 2: general metric scan (original logic, hardened) ────────────────
    m_vals = []
    total_area_found = None

    for match in re.finditer(rf'(\d+\.?\d*)\s?{m_unit}', relevant_text, re.IGNORECASE):
        val = float(match.group(1))
        full_match_text = match.group(0).lower()
        start_idx = match.start()
        context_before = relevant_text[max(0, start_idx - 80):start_idx]
        context_before_low = context_before.lower()
        bracket_context = relevant_text[max(0, start_idx - 150):start_idx]

        is_rera_duplicate = (
            "(" in bracket_context and "रेरा" in bracket_context
            and ")" not in bracket_context
        )

        # Skip survey-number land parcels
        if _is_survey_land_value(context_before_low):
            continue

        # Skip yatun/pakhi sub-totals that are land totals
        if _is_yatun_pakhi_total(context_before_low, full_match_text, val):
            continue

        if any(word in context_before_low for word in exclude_keywords):
            continue

        if 2.0 <= val < 900 and not is_rera_duplicate:
            # Only treat एकूण क्षेत्र as a total when it's NOT in a land/survey context
            is_land_context = (
                "स.नं" in bracket_context or
                "survey" in bracket_context.lower() or
                val > 500  # land totals are almost always > 500 sq.mt
            )
            if ("एकूण क्षेत्र" in context_before_low or "एकूण क्षेत्र" in full_match_text):
                if not is_land_context:
                    total_area_found = val

            if not m_vals or val != m_vals[-1]:
                m_vals.append(val)

    if total_area_found:
        return round(total_area_found, 3)

    if m_vals:
        if len(m_vals) > 1:
            m_vals.sort()
            for i in range(1, len(m_vals)):
                if abs(m_vals[i] - sum(m_vals[:i])) < 1.0:
                    return round(m_vals[i], 3)
        return round(sum(m_vals), 3)

    # ── PASS 3: imperial fallback ─────────────────────────────────────────────
    f_vals = []
    for match in re.finditer(rf'(\d+\.?\d*)\s?{f_unit}', relevant_text, re.IGNORECASE):
        val = float(match.group(1))
        start_idx = match.start()
        context_before = relevant_text[max(0, match.start() - 80):start_idx].lower()
        if not any(word in context_before for word in exclude_keywords):
            if 20.0 <= val < 9000:
                if not f_vals or val != f_vals[-1]:
                    f_vals.append(val)

    if f_vals:
        if len(f_vals) > 1:
            f_vals.sort()
            for i in range(1, len(f_vals)):
                if abs(f_vals[i] - sum(f_vals[:i])) < 5.0:
                    return round(f_vals[i] / 10.764, 3)
        return round(sum(f_vals) / 10.764, 3)

    return 0.0


def determine_config(area, t1, t2, t3):
    if area == 0: return "N/A"
    if area < t1: return "1 BHK"
    elif area < t2: return "2 BHK"
    elif area < t3: return "3 BHK"
    else: return "4 BHK"


# ─────────────────────────────────────────────
# EXCEL FORMATTING
# ─────────────────────────────────────────────
def apply_excel_formatting(df, writer, sheet_name, is_summary=True):
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes = "A2"
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    colors = ["A2D2FF", "FFD6A5", "CAFFBF", "FDFFB6", "FFADAD", "BDB2FF", "9BF6FF"]

    for i in range(1, worksheet.max_row + 1):
        for j in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=i, column=j)
            cell.alignment = center_align
            if is_summary: cell.border = thin_border

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
                fill = PatternFill(start_color=colors[color_idx % len(colors)], end_color=colors[color_idx % len(colors)], fill_type="solid")
                for r in range(start_row_prop, i):
                    for c in range(2, last_col + 1):
                        worksheet.cell(row=r, column=c).fill = fill
                if i-1 > start_row_prop:
                    worksheet.merge_cells(start_row=start_row_prop, start_column=2, end_row=i-1, end_column=2)
                    worksheet.merge_cells(start_row=start_row_prop, start_column=last_col, end_row=i-1, end_column=last_col)
                start_row_prop = i
                color_idx += 1

            if curr_loc != prev_loc and i > 2:
                for r in range(start_row_loc, i):
                    worksheet.cell(row=r, column=1).fill = white_fill
                if i-1 > start_row_loc:
                    worksheet.merge_cells(start_row=start_row_loc, start_column=1, end_row=i-1, end_column=1)
                start_row_loc = i


# ─────────────────────────────────────────────
# EMAIL
# ─────────────────────────────────────────────
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


# ─────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────
st.set_page_config(page_title="Spydarr Dashboard", layout="wide")
st.title("Spydarr Dashboard")
st.markdown(
    "<div style='margin-top: -15px; margin-bottom: 10px;'>"
    "<span style='background-color: #FFFF00; padding: 2px 8px; border-radius: 4px; "
    "border: 1px solid #E6E600; font-size: 0.9em; color: black;'>"
    "<u><strong>NOTE :-</strong> Please cross-check the report manually.</u></span></div>",
    unsafe_allow_html=True
)

st.sidebar.header("Calculation Settings")
loading_factor = st.sidebar.number_input("Loading Factor", min_value=1.0, value=1.40, step=0.001, format="%.3f")
t1 = st.sidebar.number_input("1 BHK Threshold (sq.ft)", value=600)
t2 = st.sidebar.number_input("2 BHK Threshold (sq.ft)", value=850)
t3 = st.sidebar.number_input("3 BHK Threshold (sq.ft)", value=1100)

# Groq toggle
use_groq = st.sidebar.toggle("🤖 Enable Groq AI Cross-Validation", value=True)
if use_groq:
    st.sidebar.info("Groq will cross-validate every regex result using **llama-3.3-70b-versatile**.")

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

        with st.spinner('Running regex extraction...'):
            df['_regex_area_sqmt'] = df[desc_col].apply(extract_area_logic)
            df['_regex_area_sqft'] = (df['_regex_area_sqmt'] * 10.764).round(3)
            df['_regex_config']    = df['_regex_area_sqft'].apply(lambda x: determine_config(x, t1, t2, t3))

        # ── Groq cross-validation ──────────────────────
        if use_groq:
            groq_areas   = []
            groq_configs = []
            groq_changed = []
            groq_reasons = []

            progress_bar  = st.progress(0, text="Groq is cross-validating results…")
            total_rows    = len(df)

            for idx, row in df.iterrows():
                result = groq_validate(
                    description  = str(row[desc_col]),
                    regex_area   = row['_regex_area_sqmt'],
                    regex_config = row['_regex_config'],
                    t1=t1, t2=t2, t3=t3
                )
                groq_areas.append(result["area"])
                groq_configs.append(result["config"])
                groq_changed.append(result["changed"])
                groq_reasons.append(result["reasoning"])

                pct = int(((df.index.get_loc(idx) + 1) / total_rows) * 100)
                progress_bar.progress(pct, text=f"Groq validating row {df.index.get_loc(idx)+1} / {total_rows}…")

            progress_bar.empty()

            df['Carpet Area (SQ.MT)'] = groq_areas
            df['Configuration']        = groq_configs
            df['_groq_changed']        = groq_changed
            df['_groq_reasoning']      = groq_reasons

            corrections = sum(groq_changed)
            if corrections:
                st.warning(f"🤖 Groq corrected **{corrections}** out of {total_rows} rows. See 'Raw Data' tab for details.")
            else:
                st.success(f"✅ Groq validated all {total_rows} rows — regex results were accurate!")

        else:
            df['Carpet Area (SQ.MT)'] = df['_regex_area_sqmt']
            df['Configuration']        = df['_regex_config']
            df['_groq_changed']        = False
            df['_groq_reasoning']      = ""

        # ── Final calculated columns ───────────────────
        df['Carpet Area (SQ.FT)'] = (df['Carpet Area (SQ.MT)'] * 10.764).round(3)
        df['Saleable Area']        = (df['Carpet Area (SQ.FT)'] * loading_factor).round(3)
        df['APR']                  = df.apply(
            lambda r: round(r[cons_col] / r['Saleable Area'], 3) if r['Saleable Area'] > 0 else 0, axis=1
        )
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')

        # Drop internal helper columns before export
        export_df = df.drop(columns=['_regex_area_sqmt', '_regex_area_sqft', '_regex_config'], errors='ignore')

        # ── Summary ───────────────────────────────────
        valid_df = export_df[export_df['Carpet Area (SQ.FT)'] > 0].sort_values(
            [loc_col, prop_col, 'Configuration', 'Carpet Area (SQ.FT)']
        )

        project_counts = valid_df.groupby(prop_col).size().reset_index(name='Total Count')
        summary = valid_df.groupby([loc_col, prop_col, 'Configuration', 'Carpet Area (SQ.FT)']).agg(
            Last_Date      =(date_col, 'max'),
            Min_APR        =('APR', 'min'),
            Max_APR        =('APR', 'max'),
            Avg_APR        =('APR', 'mean'),
            Median_APR     =('APR', 'median'),
            Property_Count =(prop_col, 'count')
        ).reset_index()

        summary = summary.merge(project_counts, on=prop_col, how='left')
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

        # ── Preview tabs ──────────────────────────────
        st.success("✅ Analysis Complete!")
        tab1, tab2 = st.tabs(["📋 Summary", "🗂 Raw Data"])
        with tab1:
            st.dataframe(summary, use_container_width=True)
        with tab2:
            highlight_col = '_groq_changed'
            if use_groq and highlight_col in export_df.columns:
                changed_mask = export_df[highlight_col]
                st.caption(f"🟡 Highlighted rows = Groq corrected the regex result ({changed_mask.sum()} rows)")
                st.dataframe(
                    export_df.style.apply(
                        lambda row: ['background-color: #FFF9C4' if row['_groq_changed'] else '' for _ in row],
                        axis=1
                    ),
                    use_container_width=True
                )
            else:
                st.dataframe(export_df, use_container_width=True)

        # ── Excel export ──────────────────────────────
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            apply_excel_formatting(export_df, writer, 'Raw Data', is_summary=False)
            apply_excel_formatting(summary,   writer, 'Summary',  is_summary=True)

        st.download_button(
            label     = "⬇️ Download Excel Report",
            data      = output.getvalue(),
            file_name = "Spydarr_Market_Summary.xlsx",
            mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ── Email ─────────────────────────────────────
        st.divider()
        recipient = st.text_input("Recipient Name", placeholder="firstname.lastname")
        if st.button("📧 Send to Email") and recipient:
            full_email = f"{recipient.strip().lower()}@beyondwalls.com"
            if send_email(full_email, output.getvalue(), "Spydarr_Market_Summary.xlsx"):
                st.success(f"Report sent to {full_email}")

    else:
        st.error(
            "Missing required columns. Ensure file has: "
            "'Micromarket', 'Property Description', 'Consideration Value', "
            "'Property', and 'Completion Date'."
        )
