
"""
🏭 Cement Quality Analytics Dashboard v4.0
=============================================
Beautiful modern UI with glassmorphism & gradients
Multi-input: Excel, CSV, Word, Copy-Paste, IMAGE OCR, Camera
Auto-detects material type | 6 Analysis Modules

Author: Quality Control Analytics Team
For: FAUJI CEMENT COMPANY LIMITED (NIZAMPUR)
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import base64
import io
import warnings
import cv2
from PIL import Image
import pytesseract
from datetime import datetime

warnings.filterwarnings('ignore')

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

st.set_page_config(page_title="Cement Quality Analytics", page_icon="🏭", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%); }
    .glass-card {
        background: rgba(30, 41, 59, 0.7);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 20px;
        padding: 1.5rem;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
        transition: all 0.3s ease;
    }
    .glass-card:hover {
        border-color: rgba(59, 130, 246, 0.3);
        box-shadow: 0 12px 40px rgba(59, 130, 246, 0.15);
        transform: translateY(-2px);
    }
    .kpi-card {
        background: linear-gradient(135deg, rgba(59, 130, 246, 0.15) 0%, rgba(139, 92, 246, 0.15) 100%);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 16px;
        padding: 1.5rem;
        text-align: center;
        position: relative;
        overflow: hidden;
    }
    .kpi-card::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0; height: 3px;
        background: linear-gradient(90deg, #3b82f6, #8b5cf6, #ec4899);
        border-radius: 16px 16px 0 0;
    }
    .kpi-value {
        font-size: 2.2rem;
        font-weight: 800;
        background: linear-gradient(135deg, #60a5fa, #a78bfa);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }
    .kpi-label { font-size: 0.85rem; color: #94a3b8; margin-top: 0.5rem; font-weight: 500; }
    .kpi-target { font-size: 0.75rem; color: #64748b; margin-top: 0.3rem; }
    .section-title {
        font-size: 1.5rem;
        font-weight: 700;
        color: #e2e8f0;
        margin: 2rem 0 1rem 0;
        padding-bottom: 0.75rem;
        border-bottom: 2px solid rgba(59, 130, 246, 0.3);
        text-shadow: 0 0 20px rgba(59, 130, 246, 0.3);
    }
    .status-badge {
        display: inline-block;
        padding: 0.25rem 0.75rem;
        border-radius: 9999px;
        font-size: 0.75rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .status-excellent { background: rgba(16, 185, 129, 0.2); color: #34d399; border: 1px solid rgba(16, 185, 129, 0.3); }
    .status-good { background: rgba(59, 130, 246, 0.2); color: #60a5fa; border: 1px solid rgba(59, 130, 246, 0.3); }
    .status-warning { background: rgba(245, 158, 11, 0.2); color: #fbbf24; border: 1px solid rgba(245, 158, 11, 0.3); }
    .status-danger { background: rgba(239, 68, 68, 0.2); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.3); }
    .recommendation-box {
        background: linear-gradient(135deg, rgba(59, 130, 246, 0.08) 0%, rgba(139, 92, 246, 0.08) 100%);
        border-left: 4px solid #3b82f6;
        border-radius: 0 12px 12px 0;
        padding: 1rem 1.25rem;
        margin: 0.75rem 0;
        backdrop-filter: blur(5px);
    }
    .grade-display {
        background: linear-gradient(135deg, rgba(16, 185, 129, 0.1) 0%, rgba(59, 130, 246, 0.1) 100%);
        border: 2px solid rgba(16, 185, 129, 0.3);
        border-radius: 20px;
        padding: 2rem;
        text-align: center;
        position: relative;
        overflow: hidden;
    }
    .stButton>button {
        background: linear-gradient(135deg, #3b82f6, #8b5cf6);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 25px rgba(59, 130, 246, 0.4);
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: rgba(30, 41, 59, 0.5);
        border-radius: 12px;
        padding: 0.5rem;
    }
    .stTabs [data-baseweb="tab"] {
        background: transparent;
        color: #94a3b8;
        border-radius: 8px;
        padding: 0.75rem 1.25rem;
        font-weight: 500;
    }
    .stTabs [data-baseweb="tab-highlight"] {
        background: linear-gradient(135deg, #3b82f6, #8b5cf6);
        border-radius: 8px;
    }
    .stTabs [aria-selected="true"] { color: white !important; font-weight: 600; }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ===================== HELPER FUNCTIONS =====================

def detect_material_type(df):
    cols = [c.upper() for c in df.columns]
    col_str = ' '.join(cols)
    clinker_indicators = ['C3S', 'C2S', 'C3A', 'C4AF', 'LIQ_PHASE', 'LIQUID', 'F_CAO', 'FREE CAO', 'LSF', 'CLINKER']
    rawmeal_indicators = ['LOI', 'RAW', 'MOISTURE', 'COMBINED', 'HOMOGENEITY']
    cement_indicators = ['BLAINE', 'FINENESS', 'SOUNDNESS', 'SETTING', 'INITIAL', 'FINAL', 'COMPRESSIVE', '3DAY', '7DAY', '28DAY']
    clinker_score = sum(1 for ind in clinker_indicators if ind in col_str)
    rawmeal_score = sum(1 for ind in rawmeal_indicators if ind in col_str)
    cement_score = sum(1 for ind in cement_indicators if ind in col_str)
    scores = {'Clinker': clinker_score, 'Raw Meal': rawmeal_score, 'Cement': cement_score}
    detected = max(scores, key=scores.get)
    if scores[detected] == 0:
        if 'LSF' in col_str and 'C3S' not in col_str:
            detected = 'Raw Meal'
        elif 'LSF' in col_str:
            detected = 'Clinker'
        else:
            detected = 'Unknown'
    return detected


def parse_pasted_data(text):
    if not text or not text.strip():
        return None
    separators = ['\t', ',', ';', r'\s+']
    best_df = None
    best_cols = 0
    for sep in separators:
        try:
            if sep == r'\s+':
                df = pd.read_csv(io.StringIO(text), sep=sep, engine='python')
            else:
                df = pd.read_csv(io.StringIO(text), sep=sep)
            if len(df.columns) > 1 and len(df) > 0:
                if len(df.columns) > best_cols:
                    best_df = df
                    best_cols = len(df.columns)
        except Exception:
            continue
    return best_df


def read_word_tables(file_bytes):
    if not HAS_DOCX:
        st.error("python-docx not installed. Run: pip install python-docx")
        return None
    try:
        doc = Document(io.BytesIO(file_bytes))
        tables_data = []
        for table in doc.tables:
            data = []
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                data.append(row_data)
            if len(data) > 1:
                df = pd.DataFrame(data[1:], columns=data[0])
                for col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                tables_data.append(df)
        if not tables_data:
            st.warning("No tables found in Word document.")
            return None
        largest = max(tables_data, key=lambda x: len(x.columns) * len(x))
        return largest
    except Exception as e:
        st.error(f"Error reading Word document: {e}")
        return None


def read_excel_all_sheets(file_bytes):
    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        sheets = {}
        for sheet_name in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet_name)
            if not df.empty:
                sheets[sheet_name] = df
        return sheets
    except Exception as e:
        st.error(f"Error reading Excel: {e}")
        return {}


def extract_data_from_image(image):
    """Extract tabular data from image using OCR"""
    try:
        img_array = np.array(image)
        if len(img_array.shape) == 3:
            gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
        else:
            gray = img_array
        gray = cv2.fastNlMeansDenoising(gray, None, 10, 7, 21)
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        custom_config = r'--oem 3 --psm 6'
        text = pytesseract.image_to_string(thresh, config=custom_config)
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        if not lines:
            return None
        data_rows = []
        for line in lines:
            parts = line.split('\t')
            if len(parts) < 2:
                parts = [p for p in line.split('  ') if p.strip()]
            if len(parts) < 2:
                parts = line.split()
            if len(parts) >= 2:
                data_rows.append(parts)
        if len(data_rows) < 2:
            return None
        headers = data_rows[0]
        data = data_rows[1:]
        max_cols = len(headers)
        cleaned_data = []
        for row in data:
            if len(row) >= max_cols:
                cleaned_data.append(row[:max_cols])
            elif len(row) > 1:
                row.extend([''] * (max_cols - len(row)))
                cleaned_data.append(row)
        if not cleaned_data:
            return None
        df = pd.DataFrame(cleaned_data, columns=headers)
        for col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='ignore')
        return df
    except Exception as e:
        st.error(f"OCR Error: {e}")
        return None


def calc_cpk(data, target, usl, lsl):
    mean = data.mean()
    std = data.std()
    if std == 0 or pd.isna(std):
        return np.inf, np.inf, mean, std
    cpu = (usl - mean) / (3 * std)
    cpl = (mean - lsl) / (3 * std)
    cpk = min(cpu, cpl)
    cp = (usl - lsl) / (6 * std)
    return cp, cpk, mean, std


def predict_strength(c3s, c2s, c3a, f_cao, lsf):
    base_strength = 35.0
    c3s_contrib = (c3s - 45) * 0.85
    c2s_contrib = (c2s - 25) * 0.25
    c3a_contrib = (c3a - 7) * 0.5
    f_cao_penalty = max(0, f_cao - 1.0) * (-3.0)
    lsf_bonus = (lsf - 90) * 0.3
    predicted = base_strength + c3s_contrib + c2s_contrib + c3a_contrib + f_cao_penalty + lsf_bonus
    return max(30, min(65, predicted))


def calc_coating_index(liquid_phase, c3a, c4af):
    if c4af == 0 or pd.isna(c4af):
        return 0
    ratio = c3a / c4af
    return (liquid_phase * ratio) / 10


def calc_burnability_index(f_cao, lsf):
    return f_cao * lsf / 100


def calc_quality_index(df):
    scores = []
    if 'LSF' in df.columns:
        lsf_score = 100 - abs(df['LSF'] - 92.0) * 5
        scores.append(lsf_score * 0.30)
    if 'SM' in df.columns:
        sm_score = 100 - abs(df['SM'] - 2.25) * 100
        scores.append(sm_score * 0.25)
    if 'AM' in df.columns:
        am_score = 100 - abs(df['AM'] - 1.30) * 100
        scores.append(am_score * 0.20)
    if 'F_CaO' in df.columns or 'F-CaO' in df.columns:
        f_col = 'F_CaO' if 'F_CaO' in df.columns else 'F-CaO'
        f_score = 100 - (df[f_col] - 1.2).clip(lower=0) * 50
        scores.append(f_score * 0.25)
    if not scores:
        return pd.Series([75] * len(df))
    total = sum(scores)
    return total.clip(0, 100)


def generate_report(df, material_type, specs):
    report = []
    report.append("=" * 80)
    report.append("CEMENT QUALITY ANALYTICS REPORT - " + material_type.upper())
    report.append("=" * 80)
    report.append("Samples Analyzed: " + str(len(df)))
    report.append("Generated: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    report.append("")
    report.append("PARAMETER SUMMARY")
    report.append("-" * 80)
    for col in df.select_dtypes(include=[np.number]).columns:
        if col in ['Quality_Index', 'Coating_Index', 'Burnability_Index']:
            continue
        mean = df[col].mean()
        std = df[col].std()
        cv = (std/mean*100 if mean != 0 else 0)
        report.append(f"{col:<20} Mean: {mean:>10.3f}   Std: {std:>10.3f}   CV: {cv:>6.2f}%")
    report.append("")
    report.append("PROCESS CAPABILITY")
    report.append("-" * 80)
    for param, spec in specs.items():
        if param in df.columns:
            cp, cpk, mean, std = calc_cpk(df[param], spec['target'], spec['usl'], spec['lsl'])
            status = 'Excellent' if cpk >= 1.33 else 'Capable' if cpk >= 1.0 else 'Marginal' if cpk >= 0.67 else 'Poor'
            report.append(f"{param:<15} Cp: {cp:>6.2f}   Cpk: {cpk:>6.2f}   Status: {status}")
    if 'Quality_Index' in df.columns:
        report.append("")
        report.append(f"QUALITY INDEX: {df['Quality_Index'].mean():.2f}/100")
    if 'Coating_Index' in df.columns:
        report.append("")
        report.append(f"COATING INDEX: {df['Coating_Index'].mean():.3f}")
    report.append("")
    report.append("=" * 80)
    return "\n".join(report)


def get_download_link(df, filename="analysis_report.csv"):
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}" style="text-decoration:none;"><button style="background:linear-gradient(135deg,#3b82f6,#8b5cf6);color:white;padding:12px 24px;border:none;border-radius:12px;cursor:pointer;font-weight:600;box-shadow:0 4px 15px rgba(59,130,246,0.3);">📥 Download CSV Report</button></a>'
    return href


# ===================== MAIN APP =====================

def main():
    # Animated header
    st.markdown("""
    <div style="text-align:center; padding:2rem 1rem 1rem;">
        <div style="font-size:3.5rem; margin-bottom:0.5rem; animation: bounce 2s infinite;">🏭</div>
        <h1 style="font-size:2.8rem; font-weight:800; background: linear-gradient(135deg, #60a5fa, #a78bfa, #f472b6); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; margin:0;">Cement Quality Analytics</h1>
        <p style="color:#94a3b8; font-size:1.1rem; margin-top:0.5rem;">Intelligent QC Analysis • OCR • Strength Prediction • Cost Optimization</p>
    </div>
    <style>@keyframes bounce { 0%,100%{transform:translateY(0);} 50%{transform:translateY(-10px);} }</style>
    """, unsafe_allow_html=True)

    # Session state
    if 'df' not in st.session_state:
        st.session_state.df = None
    if 'material_type' not in st.session_state:
        st.session_state.material_type = None
    if 'specs' not in st.session_state:
        st.session_state.specs = None

    # ===================== SIDEBAR =====================
    with st.sidebar:
        st.markdown("""
        <div style="text-align:center; padding:1rem 0;">
            <div style="font-size:2.5rem;">📊</div>
            <h3 style="color:#e2e8f0; margin:0.5rem 0;">Data Input</h3>
            <p style="color:#64748b; font-size:0.85rem;">Choose your input method</p>
        </div>
        """, unsafe_allow_html=True)

        input_methods = {
            "📂 Upload File": "Excel, CSV files",
            "📋 Copy & Paste": "Direct from Excel",
            "📄 Word Document": ".docx tables",
            "📸 Image / Screenshot": "OCR data extraction",
            "📷 Camera Capture": "Live photo capture"
        }

        input_method = st.radio(
            "Select method:",
            list(input_methods.keys()),
            format_func=lambda x: f"{x}\n  └ {input_methods[x]}",
            help="Choose how you want to input your QC data"
        )

        st.markdown("<hr style='border-color:rgba(255,255,255,0.1); margin:1.5rem 0;'>", unsafe_allow_html=True)

        df = None
        sheet_name = None
        ocr_used = False

        # ===== METHOD 1: FILE UPLOAD =====
        if input_method == "📂 Upload File":
            uploaded_file = st.file_uploader(
                "Drop Excel or CSV",
                type=['csv', 'xlsx', 'xls'],
                help="Supports multi-sheet Excel files"
            )

            if uploaded_file is not None:
                with st.spinner("📖 Reading file..."):
                    file_bytes = uploaded_file.read()

                    if uploaded_file.name.endswith('.csv'):
                        df = pd.read_csv(io.BytesIO(file_bytes))
                        st.success(f"✅ CSV loaded: {len(df)} rows x {len(df.columns)} cols")
                    else:
                        sheets = read_excel_all_sheets(file_bytes)
                        if len(sheets) == 0:
                            st.error("❌ No readable data found")
                        elif len(sheets) == 1:
                            df = list(sheets.values())[0]
                            sheet_name = list(sheets.keys())[0]
                            st.success(f"✅ Loaded: {len(df)} rows from '{sheet_name}'")
                        else:
                            st.markdown(f"<div style='background:rgba(59,130,246,0.15); border:1px solid rgba(59,130,246,0.3); border-radius:10px; padding:0.75rem; margin:0.5rem 0;'><strong>📑 {len(sheets)} sheets detected</strong></div>", unsafe_allow_html=True)
                            selected_sheet = st.selectbox("Choose sheet:", list(sheets.keys()))
                            df = sheets[selected_sheet]
                            sheet_name = selected_sheet
                            st.success(f"✅ '{selected_sheet}': {len(df)} rows x {len(df.columns)} cols")

                            with st.expander("📂 Other sheets"):
                                for sname, sdf in sheets.items():
                                    if sname != selected_sheet:
                                        st.write(f"• **{sname}**: {len(sdf)} rows, {len(sdf.columns)} cols")

        # ===== METHOD 2: COPY-PASTE =====
        elif input_method == "📋 Copy & Paste":
            st.markdown("""
            <div style="background:rgba(245,158,11,0.1); border:1px solid rgba(245,158,11,0.3); border-radius:12px; padding:1rem; margin:0.5rem 0;">
                <p style="color:#fbbf24; font-size:0.9rem; margin:0;">💡 Copy directly from Excel and paste below</p>
            </div>
            """, unsafe_allow_html=True)

            paste_text = st.text_area(
                "Paste data:",
                height=180,
                placeholder="Time\tSiO2\tAl2O3\t...\n07:00\t21.56\t5.38\t...\n08:00\t21.60\t5.36\t...",
                help="Include header row. Supports tab, comma, space, semicolon separators."
            )

            if st.button("🔄 Parse Data", type="primary", use_container_width=True):
                if paste_text.strip():
                    with st.spinner("🔍 Parsing..."):
                        df = parse_pasted_data(paste_text)
                        if df is not None:
                            st.success(f"✅ Parsed: {len(df)} rows x {len(df.columns)} columns")
                            st.balloons()
                        else:
                            st.error("❌ Could not parse. Check format and include header row.")
                else:
                    st.warning("⚠️ Paste some data first")

        # ===== METHOD 3: WORD =====
        elif input_method == "📄 Word Document":
            if not HAS_DOCX:
                st.error("⚠️ python-docx not installed")
                st.code("pip install python-docx", language="bash")
            else:
                uploaded_doc = st.file_uploader(
                    "Upload .docx with tables",
                    type=['docx'],
                    help="Document must contain data tables"
                )
                if uploaded_doc is not None:
                    with st.spinner("📄 Extracting tables..."):
                        file_bytes = uploaded_doc.read()
                        df = read_word_tables(file_bytes)
                        if df is not None:
                            st.success(f"✅ Table extracted: {len(df)} rows x {len(df.columns)} cols")

        # ===== METHOD 4: IMAGE / SCREENSHOT =====
        elif input_method == "📸 Image / Screenshot":
            st.markdown("""
            <div style="background:rgba(139,92,246,0.1); border:1px solid rgba(139,92,246,0.3); border-radius:12px; padding:1rem; margin:0.5rem 0;">
                <p style="color:#a78bfa; font-size:0.9rem; margin:0;">📸 Upload a screenshot of your control register</p>
                <p style="color:#64748b; font-size:0.75rem; margin:0.3rem 0 0;">The app will use OCR to extract data automatically</p>
            </div>
            """, unsafe_allow_html=True)

            uploaded_image = st.file_uploader(
                "Upload image (PNG, JPG, JPEG)",
                type=['png', 'jpg', 'jpeg'],
                help="Screenshot of control register, Excel sheet, or data table"
            )

            if uploaded_image is not None:
                image = Image.open(uploaded_image)
                st.session_state.ocr_image = image

                st.markdown("<p style='color:#94a3b8; font-size:0.85rem; margin-bottom:0.5rem;'>📷 Image Preview:</p>", unsafe_allow_html=True)
                st.image(image, use_column_width=True, caption="Original image")

                if st.button("🔍 Extract Data with OCR", type="primary", use_container_width=True):
                    with st.spinner("🧠 Processing image with OCR..."):
                        progress_bar = st.progress(0)
                        import time
                        time.sleep(0.5)
                        progress_bar.progress(30)

                        df = extract_data_from_image(image)

                        time.sleep(0.3)
                        progress_bar.progress(70)
                        time.sleep(0.3)
                        progress_bar.progress(100)

                        if df is not None and not df.empty:
                            st.success(f"✅ OCR Success! Extracted {len(df)} rows x {len(df.columns)} columns")
                            st.balloons()
                            ocr_used = True

                            with st.expander("🔍 View extracted data preview"):
                                st.dataframe(df.head(10), use_container_width=True)
                        else:
                            st.error("❌ OCR failed to extract structured data. Try:")
                            st.markdown("""
                            - Ensure the image is clear and well-lit
                            - Text should be horizontal and readable
                            - Try cropping to just the data table
                            - Use higher resolution image
                            """)
                            st.info("💡 Tip: You can also use the Copy-Paste method as fallback")

        # ===== METHOD 5: CAMERA =====
        elif input_method == "📷 Camera Capture":
            st.markdown("""
            <div style="background:rgba(16,185,129,0.1); border:1px solid rgba(16,185,129,0.3); border-radius:12px; padding:1rem; margin:0.5rem 0;">
                <p style="color:#34d399; font-size:0.9rem; margin:0;">📷 Take a live photo of your control register</p>
                <p style="color:#64748b; font-size:0.75rem; margin:0.3rem 0 0;">Best for: Plant floor, QC lab, control room screens</p>
            </div>
            """, unsafe_allow_html=True)

            camera_image = st.camera_input(
                "Take a photo",
                help="Point camera at control register screen or printed sheet"
            )

            if camera_image is not None:
                image = Image.open(camera_image)
                st.session_state.ocr_image = image

                st.image(image, use_column_width=True, caption="Captured photo")

                if st.button("🔍 Extract Data from Photo", type="primary", use_container_width=True):
                    with st.spinner("🧠 Processing photo with OCR..."):
                        progress_bar = st.progress(0)
                        import time
                        time.sleep(0.3)
                        progress_bar.progress(30)

                        df = extract_data_from_image(image)

                        time.sleep(0.3)
                        progress_bar.progress(70)
                        time.sleep(0.3)
                        progress_bar.progress(100)

                        if df is not None and not df.empty:
                            st.success(f"✅ Photo processed! Extracted {len(df)} rows x {len(df.columns)} columns")
                            st.balloons()
                            ocr_used = True
                        else:
                            st.error("❌ Could not read data from photo. Tips:")
                            st.markdown("""
                            - Hold camera steady and ensure good lighting
                            - Avoid glare on screens
                            - Keep text horizontal and in focus
                            - Get closer to fill the frame with the table
                            """)

        st.markdown("<hr style='border-color:rgba(255,255,255,0.1); margin:1.5rem 0;'>", unsafe_allow_html=True)
        st.markdown("""
        <div style="text-align:center; padding:1rem 0;">
            <p style="color:#64748b; font-size:0.8rem;">🏭 FAUJI CEMENT CO.</p>
            <p style="color:#475569; font-size:0.75rem;">QC-R-18 | Line 3</p>
        </div>
        """, unsafe_allow_html=True)


    # ===================== PROCESS DATA =====================
    if df is not None and not df.empty:
        try:
            df.columns = df.columns.str.strip()
            material_type = detect_material_type(df)

            time_cols = [c for c in df.columns if any(x in c.upper() for x in ['TIME', 'DATE', 'HOUR', 'SHIFT'])]
            if time_cols:
                df = df.rename(columns={time_cols[0]: 'Time'})

            ocr_badge = " | 🔍 OCR Extracted" if ocr_used else ""
            sheet_info = f" | 📑 {sheet_name}" if sheet_name else ""

            st.markdown(f"""
            <div style="background: linear-gradient(135deg, rgba(59,130,246,0.15), rgba(139,92,246,0.15)); 
                        border: 1px solid rgba(59,130,246,0.3); border-radius: 16px; padding: 1rem 1.5rem; 
                        margin: 1rem 0; display: flex; align-items: center; gap: 1rem;">
                <div style="font-size: 2rem;">✅</div>
                <div>
                    <div style="color: #e2e8f0; font-weight: 600; font-size: 1.1rem;">
                        Data Loaded: <span style="color: #60a5fa;">{material_type}</span>
                    </div>
                    <div style="color: #94a3b8; font-size: 0.9rem;">
                        {len(df)} samples x {len(df.columns)} parameters{sheet_info}{ocr_badge}
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # Standardize columns
            col_mapping = {
                'F-CaO': 'F_CaO', 'Free CaO': 'F_CaO', 'FreeCaO': 'F_CaO',
                'Liq.Phase': 'Liq_phase', 'Liquid Phase': 'Liq_phase', 'LIQ PHASE': 'Liq_phase',
                'Liq. Phase': 'Liq_phase', 'LIQ.PHASE': 'Liq_phase',
                'C3S': 'C3S', 'C2S': 'C2S', 'C3A': 'C3A', 'C4AF': 'C4AF',
                'Temp': 'Temp', 'Temperature': 'Temp', 'TEMP': 'Temp'
            }
            df = df.rename(columns={k: v for k, v in col_mapping.items() if k in df.columns})

            # Derived metrics
            if all(c in df.columns for c in ['Liq_phase', 'C3A', 'C4AF']):
                df['Coating_Index'] = df.apply(lambda row: calc_coating_index(row['Liq_phase'], row['C3A'], row['C4AF']), axis=1)

            if all(c in df.columns for c in ['F_CaO', 'LSF']):
                df['Burnability_Index'] = df.apply(lambda row: calc_burnability_index(row['F_CaO'], row['LSF']), axis=1)

            df['Quality_Index'] = calc_quality_index(df)

            if material_type == 'Clinker' and all(c in df.columns for c in ['C3S', 'C2S', 'C3A', 'F_CaO', 'LSF']):
                df['Predicted_Strength_28D'] = df.apply(
                    lambda row: predict_strength(row['C3S'], row['C2S'], row['C3A'], row['F_CaO'], row['LSF']), axis=1
                )

            # Specs
            if material_type == 'Clinker':
                specs = {
                    'LSF': {'target': 92.0, 'usl': 96.0, 'lsl': 88.0},
                    'SM': {'target': 2.25, 'usl': 2.50, 'lsl': 2.00},
                    'AM': {'target': 1.30, 'usl': 1.50, 'lsl': 1.10},
                    'F_CaO': {'target': 1.0, 'usl': 1.5, 'lsl': 0.0},
                    'C3S': {'target': 50.0, 'usl': 55.0, 'lsl': 45.0},
                    'Liq_phase': {'target': 28.0, 'usl': 30.0, 'lsl': 26.0}
                }
            elif material_type == 'Raw Meal':
                specs = {
                    'LSF': {'target': 92.0, 'usl': 96.0, 'lsl': 88.0},
                    'SM': {'target': 2.25, 'usl': 2.50, 'lsl': 2.00},
                    'AM': {'target': 1.30, 'usl': 1.50, 'lsl': 1.10},
                    'LOI': {'target': 35.0, 'usl': 38.0, 'lsl': 32.0}
                }
            else:
                specs = {
                    'Blaine': {'target': 350, 'usl': 400, 'lsl': 300},
                    'SO3': {'target': 2.5, 'usl': 3.0, 'lsl': 2.0}
                }

            st.session_state.df = df
            st.session_state.material_type = material_type
            st.session_state.specs = specs

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
            st.session_state.df = None

    # ===================== DISPLAY ANALYSIS =====================
    if st.session_state.df is not None:
        df = st.session_state.df
        material_type = st.session_state.material_type
        specs = st.session_state.specs

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "📊 Overview", "📈 Trends", "🔬 Capability", 
            "🎯 Coating & Strength", "💰 Operations", "📄 Export"
        ])

        # ========== TAB 1: OVERVIEW ==========
        with tab1:
            st.markdown('<div class="section-title">📊 Key Performance Indicators</div>', unsafe_allow_html=True)

            kpi_data = []
            if 'LSF' in df.columns:
                kpi_data.append(("Avg LSF", f"{df['LSF'].mean():.2f}", "Target: 92.0"))
            if 'SM' in df.columns:
                kpi_data.append(("Avg SM", f"{df['SM'].mean():.2f}", "Target: 2.25"))
            if 'AM' in df.columns:
                kpi_data.append(("Avg AM", f"{df['AM'].mean():.2f}", "Target: 1.30"))
            if 'C3S' in df.columns:
                kpi_data.append(("Avg C3S", f"{df['C3S'].mean():.1f}%", "Target: 50%"))
            if 'F_CaO' in df.columns:
                kpi_data.append(("Avg F-CaO", f"{df['F_CaO'].mean():.2f}%", "Target: <1.2%"))
            if 'Quality_Index' in df.columns:
                kpi_data.append(("Quality Index", f"{df['Quality_Index'].mean():.1f}", "Score: 0-100"))

            kpi_cols = st.columns(min(6, len(kpi_data)))
            for i, (label, value, subtext) in enumerate(kpi_data):
                with kpi_cols[i]:
                    st.markdown(f'<div class="kpi-card"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div><div class="kpi-target">{subtext}</div></div>', unsafe_allow_html=True)

            st.markdown("---")
            col1, col2 = st.columns([2, 1])
            with col1:
                st.markdown('<div class="section-title">📋 Data Preview</div>', unsafe_allow_html=True)
                st.dataframe(df.head(20), use_container_width=True, height=400)
            with col2:
                st.markdown('<div class="section-title">📊 Statistics</div>', unsafe_allow_html=True)
                st.dataframe(df.describe().round(3), use_container_width=True, height=400)

            st.markdown('<div class="section-title">🎯 Overall Assessment</div>', unsafe_allow_html=True)
            if 'Quality_Index' in df.columns:
                q_mean = df['Quality_Index'].mean()
                if q_mean >= 97:
                    grade, color = "EXCELLENT", "#10b981"
                elif q_mean >= 95:
                    grade, color = "GOOD", "#3b82f6"
                elif q_mean >= 93:
                    grade, color = "ACCEPTABLE", "#f59e0b"
                else:
                    grade, color = "NEEDS IMPROVEMENT", "#ef4444"

                st.markdown(f'<div class="grade-display"><div style="font-size: 1.3rem; color: {color}; font-weight: 700;">OVERALL GRADE</div><div style="font-size: 3rem; color: {color}; font-weight: 800; margin: 0.5rem 0;">{grade}</div><div style="font-size: 2.5rem; font-weight: 700; color: #e2e8f0;">{q_mean:.1f}<span style="font-size: 1.2rem; color: #64748b;">/100</span></div><div style="color: #64748b; font-size: 0.9rem;">Quality Index Score</div></div>', unsafe_allow_html=True)

        # ========== TAB 2: TRENDS ==========
        with tab2:
            st.markdown('<div class="section-title">📈 Parameter Trends with Control Limits</div>', unsafe_allow_html=True)
            numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            exclude = ['Quality_Index', 'Coating_Index', 'Burnability_Index', 'Predicted_Strength_28D']
            plot_cols = [c for c in numeric_cols if c not in exclude]
            selected_params = st.multiselect("Select parameters:", plot_cols, default=plot_cols[:4] if len(plot_cols) >= 4 else plot_cols)

            if selected_params:
                x_col = df['Time'] if 'Time' in df.columns else df.index.astype(str)
                n_params = len(selected_params)
                n_rows = max(1, (n_params + 1) // 2)
                n_cols = 2 if n_params > 1 else 1

                fig = make_subplots(rows=n_rows, cols=n_cols, subplot_titles=selected_params, vertical_spacing=0.12)
                for i, param in enumerate(selected_params):
                    row = i // 2 + 1
                    col = (i % 2) + 1 if n_cols > 1 else 1
                    fig.add_trace(go.Scatter(x=x_col, y=df[param], mode='lines+markers', name=param, line=dict(width=2, color='#3b82f6')), row=row, col=col)
                    if param in specs:
                        fig.add_hline(y=specs[param]['target'], line_dash="dash", line_color="#fbbf24", line_width=2, row=row, col=col)
                        fig.add_hline(y=specs[param]['usl'], line_dash="dot", line_color="#ef4444", line_width=1.5, row=row, col=col)
                        fig.add_hline(y=specs[param]['lsl'], line_dash="dot", line_color="#ef4444", line_width=1.5, row=row, col=col)

                fig.update_layout(height=300 * n_rows, showlegend=False, template="plotly_dark", paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(30,41,59,0.5)', title_text="Trend Analysis with Control Limits", font=dict(color='#e2e8f0'))
                st.plotly_chart(fig, use_container_width=True)

            if material_type == 'Clinker' and all(c in df.columns for c in ['C3S', 'C2S', 'C3A', 'C4AF']):
                st.markdown('<div class="section-title">🧪 Mineral Phase Composition</div>', unsafe_allow_html=True)
                x_col = df['Time'] if 'Time' in df.columns else df.index.astype(str)
                fig = go.Figure()
                colors_min = ['#60a5fa', '#34d399', '#fbbf24', '#f87171']
                minerals = ['C3S', 'C2S', 'C3A', 'C4AF']
                for mineral, color in zip(minerals, colors_min):
                    fig.add_trace(go.Scatter(x=x_col, y=df[mineral], mode='lines', stackgroup='one', name=mineral, line=dict(width=0.5, color=color), fillcolor=color))
                fig.update_layout(title="Mineral Phase Stacked Area", yaxis_title="Weight %", template="plotly_dark", paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(30,41,59,0.5)', font=dict(color='#e2e8f0'), height=500)
                st.plotly_chart(fig, use_container_width=True)


        # ========== TAB 3: CAPABILITY ==========
        with tab3:
            st.markdown('<div class="section-title">🔬 Process Capability Analysis</div>', unsafe_allow_html=True)
            capability_data = []
            for param, spec in specs.items():
                if param in df.columns:
                    cp, cpk, mean, std = calc_cpk(df[param], spec['target'], spec['usl'], spec['lsl'])
                    status = 'Excellent' if cpk >= 1.33 else 'Capable' if cpk >= 1.0 else 'Marginal' if cpk >= 0.67 else 'Poor'
                    capability_data.append({'Parameter': param, 'Mean': f"{mean:.3f}", 'Std': f"{std:.3f}", 'Cp': f"{cp:.2f}", 'Cpk': f"{cpk:.2f}", 'Status': status})

            if capability_data:
                cap_df = pd.DataFrame(capability_data)

                # Styled HTML table
                st.markdown("<div style='background:rgba(30,41,59,0.7); border-radius:16px; padding:1rem; border:1px solid rgba(255,255,255,0.1);'><table style='width:100%; color:#e2e8f0; border-collapse:collapse;'>" +
                           "<tr style='border-bottom:2px solid rgba(59,130,246,0.3);'>" +
                           "<th style='padding:0.75rem; text-align:left;'>Parameter</th>" +
                           "<th style='padding:0.75rem; text-align:right;'>Mean</th>" +
                           "<th style='padding:0.75rem; text-align:right;'>Std</th>" +
                           "<th style='padding:0.75rem; text-align:right;'>Cp</th>" +
                           "<th style='padding:0.75rem; text-align:right;'>Cpk</th>" +
                           "<th style='padding:0.75rem; text-align:center;'>Status</th></tr>", unsafe_allow_html=True)

                for _, row in cap_df.iterrows():
                    status_colors = {'Excellent': '#10b981', 'Capable': '#3b82f6', 'Marginal': '#f59e0b', 'Poor': '#ef4444'}
                    sc = status_colors.get(row['Status'], '#94a3b8')
                    st.markdown(f"<tr style='border-bottom:1px solid rgba(255,255,255,0.05);'>" +
                               f"<td style='padding:0.75rem;'>{row['Parameter']}</td>" +
                               f"<td style='padding:0.75rem; text-align:right;'>{row['Mean']}</td>" +
                               f"<td style='padding:0.75rem; text-align:right;'>{row['Std']}</td>" +
                               f"<td style='padding:0.75rem; text-align:right;'>{row['Cp']}</td>" +
                               f"<td style='padding:0.75rem; text-align:right; font-weight:600;'>{row['Cpk']}</td>" +
                               f"<td style='padding:0.75rem; text-align:center;'><span style='display:inline-block; padding:0.25rem 0.75rem; border-radius:999px; font-size:0.75rem; font-weight:600; background:{sc}22; color:{sc}; border:1px solid {sc}44;'>{row['Status']}</span></td></tr>", unsafe_allow_html=True)

                st.markdown("</table></div>", unsafe_allow_html=True)

                fig = px.bar(cap_df, x='Parameter', y='Cpk', color='Status', 
                            color_discrete_map={'Excellent': '#10b981', 'Capable': '#3b82f6', 'Marginal': '#f59e0b', 'Poor': '#ef4444'},
                            title="Process Capability Index (Cpk)", template="plotly_dark")
                fig.add_hline(y=1.33, line_dash="dash", line_color="#10b981", annotation_text="Excellent")
                fig.add_hline(y=1.0, line_dash="dash", line_color="#3b82f6", annotation_text="Capable")
                fig.add_hline(y=0.67, line_dash="dash", line_color="#f59e0b", annotation_text="Marginal")
                fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(30,41,59,0.5)', font=dict(color='#e2e8f0'))
                st.plotly_chart(fig, use_container_width=True)

            st.markdown('<div class="section-title">📊 Parameter Variability (CV %)</div>', unsafe_allow_html=True)
            cv_data = []
            for col in plot_cols:
                mean = df[col].mean()
                if mean != 0 and not pd.isna(mean):
                    cv = (df[col].std() / mean) * 100
                    cv_data.append({'Parameter': col, 'CV_%': cv})

            if cv_data:
                cv_df = pd.DataFrame(cv_data).sort_values('CV_%', ascending=True)
                fig = px.bar(cv_df, x='CV_%', y='Parameter', orientation='h', color='CV_%', 
                            color_continuous_scale='RdYlGn_r', title="Coefficient of Variation", template="plotly_dark")
                fig.add_vline(x=1.0, line_dash="dash", line_color="#ef4444", annotation_text="Target <1%")
                fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(30,41,59,0.5)', font=dict(color='#e2e8f0'))
                st.plotly_chart(fig, use_container_width=True)

        # ========== TAB 4: COATING & STRENGTH ==========
        with tab4:
            st.markdown('<div class="section-title">🎯 Coating Index & Burnability Analysis</div>', unsafe_allow_html=True)

            if 'Coating_Index' in df.columns:
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(f'<div class="glass-card" style="text-align:center;"><div style="font-size:0.85rem;color:#94a3b8;">Average</div><div style="font-size:2rem;font-weight:800;color:#fbbf24;">{df["Coating_Index"].mean():.3f}</div></div>', unsafe_allow_html=True)
                with col2:
                    st.markdown(f'<div class="glass-card" style="text-align:center;"><div style="font-size:0.85rem;color:#94a3b8;">Maximum</div><div style="font-size:2rem;font-weight:800;color:#f87171;">{df["Coating_Index"].max():.3f}</div></div>', unsafe_allow_html=True)
                with col3:
                    risk = "LOW" if df['Coating_Index'].mean() < 1.55 else "MODERATE" if df['Coating_Index'].mean() < 1.65 else "HIGH"
                    risk_color = "#10b981" if risk == "LOW" else "#fbbf24" if risk == "MODERATE" else "#ef4444"
                    st.markdown(f'<div class="glass-card" style="text-align:center;"><div style="font-size:0.85rem;color:#94a3b8;">Risk Level</div><div style="font-size:2rem;font-weight:800;color:{risk_color};">{risk}</div></div>', unsafe_allow_html=True)

                x_col = df['Time'] if 'Time' in df.columns else df.index.astype(str)
                fig = px.bar(df, x=x_col, y='Coating_Index', color='Coating_Index', 
                            color_continuous_scale='YlOrRd', title="Coating Index by Time", template="plotly_dark")
                fig.add_hline(y=1.55, line_dash="dash", line_color="#fbbf24", line_width=2)
                fig.add_hline(y=1.65, line_dash="dash", line_color="#ef4444", line_width=2)
                fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(30,41,59,0.5)', font=dict(color='#e2e8f0'))
                st.plotly_chart(fig, use_container_width=True)

                st.markdown("<div style='display:flex; gap:1rem; margin:1rem 0;'>" +
                           "<div style='flex:1; background:rgba(16,185,129,0.1); border:1px solid rgba(16,185,129,0.3); border-radius:10px; padding:0.75rem; text-align:center;'><div style='color:#34d399; font-weight:700;'>&lt; 1.55</div><div style='color:#94a3b8; font-size:0.8rem;'>Low Risk</div></div>" +
                           "<div style='flex:1; background:rgba(245,158,11,0.1); border:1px solid rgba(245,158,11,0.3); border-radius:10px; padding:0.75rem; text-align:center;'><div style='color:#fbbf24; font-weight:700;'>1.55 – 1.65</div><div style='color:#94a3b8; font-size:0.8rem;'>Moderate</div></div>" +
                           "<div style='flex:1; background:rgba(239,68,68,0.1); border:1px solid rgba(239,68,68,0.3); border-radius:10px; padding:0.75rem; text-align:center;'><div style='color:#f87171; font-weight:700;'>&gt; 1.65</div><div style='color:#94a3b8; font-size:0.8rem;'>High Risk</div></div>" +
                           "</div>", unsafe_allow_html=True)
            else:
                st.info("Coating Index requires Liq_phase, C3A, and C4AF columns.")

            if 'Burnability_Index' in df.columns:
                st.markdown('<div class="section-title">🔥 Burnability Index</div>', unsafe_allow_html=True)
                x_col = df['Time'] if 'Time' in df.columns else df.index.astype(str)
                fig = px.line(df, x=x_col, y='Burnability_Index', markers=True, title="Burnability Index Trend", template="plotly_dark")
                fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(30,41,59,0.5)', font=dict(color='#e2e8f0'))
                st.plotly_chart(fig, use_container_width=True)

            if 'Predicted_Strength_28D' in df.columns:
                st.markdown('<div class="section-title">💪 Predicted 28-Day Compressive Strength</div>', unsafe_allow_html=True)
                x_col = df['Time'] if 'Time' in df.columns else df.index.astype(str)
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=x_col, y=df['Predicted_Strength_28D'], mode='lines+markers', name='Predicted Strength', line=dict(color='#60a5fa', width=3)))
                fig.add_hline(y=42.5, line_dash="dash", line_color="#10b981", line_width=2, annotation_text="OPC 42.5")
                fig.add_hline(y=52.5, line_dash="dash", line_color="#a78bfa", line_width=2, annotation_text="OPC 52.5")
                fig.update_layout(title="Expected 28-Day Strength (MPa)", yaxis_title="Strength (MPa)", template="plotly_dark", paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(30,41,59,0.5)', font=dict(color='#e2e8f0'), height=500)
                st.plotly_chart(fig, use_container_width=True)

                st.markdown(f'<div class="recommendation-box"><strong style="color:#60a5fa;">💡 Strength Prediction Model</strong><br><span style="color:#cbd5e1;">Predicted range: <strong>{df["Predicted_Strength_28D"].min():.1f} – {df["Predicted_Strength_28D"].max():.1f} MPa</strong></span><br><span style="color:#cbd5e1;">Average: <strong style="color:#34d399;">{df["Predicted_Strength_28D"].mean():.1f} MPa</strong></span><br><span style="color:#94a3b8; font-size:0.85rem;">To increase: boost LSF toward 92.0, reduce F-CaO below 1.2%</span></div>', unsafe_allow_html=True)

        # ========== TAB 5: OPERATIONS ==========
        with tab5:
            st.markdown('<div class="section-title">💰 Cost & Operational Metrics</div>', unsafe_allow_html=True)

            col1, col2 = st.columns(2)
            with col1:
                st.markdown('<div class="section-title" style="font-size:1.1rem;">🔥 Fuel Efficiency</div>', unsafe_allow_html=True)
                if 'F_CaO' in df.columns:
                    x_col = df['Time'] if 'Time' in df.columns else df.index.astype(str)
                    fuel_eff = 100 - (df['F_CaO'] - 1.0).clip(lower=0) * 20
                    fig = px.line(x=x_col, y=fuel_eff, markers=True, title="Fuel Efficiency Index (%)", labels={'y': 'Efficiency %', 'x': 'Time'}, template="plotly_dark")
                    fig.add_hline(y=95, line_dash="dash", line_color="#10b981")
                    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(30,41,59,0.5)', font=dict(color='#e2e8f0'))
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("F-CaO data required")
            with col2:
                st.markdown('<div class="section-title" style="font-size:1.1rem;">📊 Raw Material Utilization</div>', unsafe_allow_html=True)
                if 'LSF' in df.columns:
                    x_col = df['Time'] if 'Time' in df.columns else df.index.astype(str)
                    rm_eff = 100 - abs(df['LSF'] - 92.0) * 2
                    fig = px.line(x=x_col, y=rm_eff, markers=True, title="Raw Material Efficiency (%)", labels={'y': 'Efficiency %', 'x': 'Time'}, template="plotly_dark")
                    fig.add_hline(y=95, line_dash="dash", line_color="#10b981")
                    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(30,41,59,0.5)', font=dict(color='#e2e8f0'))
                    st.plotly_chart(fig, use_container_width=True)

            st.markdown('<div class="section-title" style="font-size:1.1rem;">💵 Cost Impact Calculator</div>', unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)
            with col1: coal_price = st.number_input("Coal Price ($/ton)", value=120.0, step=5.0)
            with col2: limestone_price = st.number_input("Limestone Price ($/ton)", value=15.0, step=1.0)
            with col3: daily_production = st.number_input("Daily Production (tons)", value=3000.0, step=100.0)

            if 'F_CaO' in df.columns and 'LSF' in df.columns:
                avg_f_cao = df['F_CaO'].mean()
                avg_lsf = df['LSF'].mean()
                excess_fuel = max(0, avg_f_cao - 1.0) * 0.05 * daily_production
                fuel_cost_impact = excess_fuel * coal_price / 1000
                lsf_deviation = abs(avg_lsf - 92.0)
                raw_material_waste = lsf_deviation * 0.02 * daily_production * limestone_price / 1000
                total_daily_cost = fuel_cost_impact + raw_material_waste

                st.markdown(f"<div style='background: linear-gradient(135deg, rgba(245,158,11,0.1), rgba(239,68,68,0.1)); border: 1px solid rgba(245,158,11,0.3); border-radius: 16px; padding: 1.5rem; margin-top: 1rem;'>" +
                           f"<h4 style='color: #fbbf24; margin-top: 0; font-size: 1.2rem;'>💡 Estimated Daily Cost Impact</h4>" +
                           f"<p><strong>Excess Fuel Cost:</strong> ${fuel_cost_impact:.2f}/day (F-CaO = {avg_f_cao:.2f}%)</p>" +
                           f"<p><strong>Raw Material Waste:</strong> ${raw_material_waste:.2f}/day (LSF deviation = {lsf_deviation:.2f})</p>" +
                           f"<p style='font-size: 1.2rem; color: #fbbf24; font-weight: bold;'>Total Estimated Impact: ${total_daily_cost:.2f}/day</p>" +
                           f"<p style='font-size: 1rem; color: #fbbf24;'>Annual Impact: ~${total_daily_cost * 365:,.0f}/year</p></div>", unsafe_allow_html=True)

            if 'Temp' in df.columns and 'F_CaO' in df.columns:
                st.markdown('<div class="section-title" style="font-size:1.1rem;">🌡️ Temperature vs Free CaO Correlation</div>', unsafe_allow_html=True)
                fig = px.scatter(df, x='Temp', y='F_CaO', trendline="ols", title="Kiln Temperature vs Free CaO Relationship", template="plotly_dark")
                fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(30,41,59,0.5)', font=dict(color='#e2e8f0'))
                st.plotly_chart(fig, use_container_width=True)
                corr = df['Temp'].corr(df['F_CaO'])
                st.markdown(f"**Correlation Coefficient:** {corr:.3f}")
                if corr < -0.3:
                    st.success("Strong negative correlation: Higher temperatures reduce Free CaO.")
                elif corr > 0.3:
                    st.warning("Positive correlation detected — investigate kiln operation.")
                else:
                    st.info("Weak correlation — other factors dominate Free CaO variation.")

        # ========== TAB 6: EXPORT ==========
        with tab6:
            st.markdown('<div class="section-title">📄 Comprehensive Analysis Report</div>', unsafe_allow_html=True)
            report_text = generate_report(df, material_type, specs)
            st.text_area("Generated Report", report_text, height=500)
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(get_download_link(df, f"{material_type.lower()}_analysis_data.csv"), unsafe_allow_html=True)
            with col2:
                b64 = base64.b64encode(report_text.encode()).decode()
                href = f'<a href="data:file/txt;base64,{b64}" download="{material_type.lower()}_quality_report.txt" style="text-decoration:none;"><button style="background:linear-gradient(135deg,#10b981,#3b82f6);color:white;padding:12px 24px;border:none;border-radius:12px;cursor:pointer;font-weight:600;box-shadow:0 4px 15px rgba(16,185,129,0.3);">📥 Download Text Report</button></a>'
                st.markdown(href, unsafe_allow_html=True)

            st.markdown('<div class="section-title">🎯 Actionable Recommendations</div>', unsafe_allow_html=True)
            recommendations = []
            if 'LSF' in df.columns:
                lsf_mean = df['LSF'].mean()
                if lsf_mean < 91.8:
                    recommendations.append(f"🔴 **Increase LSF:** Current {lsf_mean:.2f} is below target 92.0. Add 0.5-1.0% limestone to raw mix.")
                elif lsf_mean > 92.5:
                    recommendations.append(f"🟡 **Decrease LSF:** Current {lsf_mean:.2f} is above optimal. Reduce limestone slightly.")
                else:
                    recommendations.append(f"🟢 **LSF Optimal:** Current {lsf_mean:.2f} is well-controlled.")
            if 'F_CaO' in df.columns:
                f_mean = df['F_CaO'].mean()
                if f_mean > 1.3:
                    recommendations.append(f"🔴 **Reduce Free CaO:** Current {f_mean:.2f}% is high. Optimize burn zone temp or raw meal fineness.")
                elif f_mean > 1.1:
                    recommendations.append(f"🟡 **Monitor Free CaO:** Current {f_mean:.2f}% is acceptable but can be improved.")
                else:
                    recommendations.append(f"🟢 **Free CaO Excellent:** Current {f_mean:.2f}% indicates good burnability.")
            if 'Coating_Index' in df.columns:
                ci_mean = df['Coating_Index'].mean()
                if ci_mean > 1.65:
                    recommendations.append(f"🔴 **High Coating Risk:** Index {ci_mean:.2f}. Inspect kiln for ring formation.")
                elif ci_mean > 1.55:
                    recommendations.append(f"🟡 **Moderate Coating:** Index {ci_mean:.2f}. Maintain regular kiln inspection schedule.")
                else:
                    recommendations.append(f"🟢 **Low Coating Risk:** Index {ci_mean:.2f}. Kiln operation is safe.")
            if 'C3S' in df.columns:
                c3s_mean = df['C3S'].mean()
                if c3s_mean < 48:
                    recommendations.append(f"🔴 **Low C3S:** {c3s_mean:.1f}% may result in lower early strength. Increase LSF and burning intensity.")
                elif c3s_mean < 50:
                    recommendations.append(f"🟡 **C3S Acceptable:** {c3s_mean:.1f}%. Slight increase in LSF will optimize strength.")
                else:
                    recommendations.append(f"🟢 **C3S Excellent:** {c3s_mean:.1f}% supports high early strength development.")
            for rec in recommendations:
                st.markdown(f'<div class="recommendation-box">{rec}</div>', unsafe_allow_html=True)

    else:
        # Welcome screen
        st.markdown("""
        <div style="text-align: center; padding: 3rem 1rem;">
            <div style="font-size: 4rem; margin-bottom: 1rem;">📤</div>
            <h2 style="color: #e2e8f0;">Upload Your Data to Begin Analysis</h2>
            <p style="color: #64748b; font-size: 1.1rem; max-width: 700px; margin: 0 auto;">
                Choose your input method from the sidebar. Supports Excel (multi-sheet), CSV, 
                Word documents, copy-paste, image OCR, and live camera capture.
            </p>
        </div>
        """, unsafe_allow_html=True)

        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.markdown("**📂 File Upload**<br>Excel, CSV<br>Multi-sheet support", unsafe_allow_html=True)
        with col2:
            st.markdown("**📋 Copy & Paste**<br>Tab, comma, space<br>Direct from Excel", unsafe_allow_html=True)
        with col3:
            st.markdown("**📄 Word**<br>.docx tables<br>Auto-extract", unsafe_allow_html=True)
        with col4:
            st.markdown("**📸 Image OCR**<br>Screenshots<br>Photo upload", unsafe_allow_html=True)
        with col5:
            st.markdown("**📷 Camera**<br>Live capture<br>Plant floor ready", unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### 📋 Expected Data Formats")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("**Clinker:**<br>Time, SiO2, Al2O3, Fe2O3, CaO, MgO, K2O, Na2O, SO3, LSF, SM, AM, C3S, C2S, C3A, C4AF, Liq_phase, F_CaO, Temp", unsafe_allow_html=True)
        with col2:
            st.markdown("**Raw Meal:**<br>Time, SiO2, Al2O3, Fe2O3, CaO, MgO, LOI, LSF, SM, AM, Moisture", unsafe_allow_html=True)
        with col3:
            st.markdown("**Cement:**<br>Time, Blaine, SO3, Setting_Time, Soundness, 3Day_Strength, 7Day_Strength, 28Day_Strength, Fineness", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
