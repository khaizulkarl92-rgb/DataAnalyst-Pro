from __future__ import annotations

import io
import re
import json
import base64
from datetime import datetime, timezone
from typing import Any

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- PAGE CONFIGURATION & THEME SETUP ---
APP_VERSION = "6.0.0-PRO"
MAX_ROWS_FOR_PROFILE = 250_000

st.set_page_config(
    page_title="Data Analyst Pro — AI Suite",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- MODERN DESIGN SYSTEM & GLASSMORPHISM STYLING ---
st.markdown("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">

<style>
    /* Global Typography & Reset */
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    .block-container {
        padding-top: 1.8rem;
        padding-bottom: 3rem;
        max-width: 95%;
    }

    /* Glassmorphism Hero Section */
    .hero-card {
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.95), rgba(15, 23, 42, 0.98));
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 20px;
        padding: 2.2rem;
        color: #ffffff;
        box-shadow: 0 20px 30px -10px rgba(0, 0, 0, 0.25);
        margin-bottom: 2rem;
    }

    .hero-title {
        font-size: 2.3rem;
        font-weight: 800;
        background: linear-gradient(120deg, #38bdf8, #818cf8, #c084fc);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.4rem;
    }

    .hero-subtitle {
        font-size: 1.05rem;
        color: #94a3b8;
        font-weight: 400;
    }

    /* Glass Cards & Metric Boxes */
    .metric-container {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 1.25rem;
        box-shadow: 0 4px 15px -2px rgba(0, 0, 0, 0.04);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .metric-container:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.08);
    }

    .badge-tag {
        display: inline-block;
        padding: 0.25rem 0.75rem;
        border-radius: 9999px;
        font-size: 0.75rem;
        font-weight: 700;
        letter-spacing: 0.05em;
        text-transform: uppercase;
    }
    .badge-primary { background: #e0f2fe; color: #0369a1; }
    .badge-success { background: #dcfce7; color: #15803d; }
    .badge-purple { background: #f3e8ff; color: #7e22ce; }

    /* File Drop Zone Customization */
    [data-testid="stFileUploader"] {
        border: 2px dashed #cbd5e1;
        border-radius: 16px;
        padding: 1.5rem;
        background: #f8fafc;
        transition: all 0.3s ease;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #6366f1;
        background: #f1f5f9;
    }

    /* Tabs Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f1f5f9;
        padding: 6px;
        border-radius: 14px;
    }

    .stTabs [data-baseweb="tab"] {
        height: 44px;
        white-space: pre;
        border-radius: 10px;
        font-weight: 600;
        font-size: 0.9rem;
        color: #475569;
    }

    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #4f46e5 !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.06);
    }
</style>
""", unsafe_allow_html=True)


# --- UTILITY & CORE DATA CLEANING FUNCTIONS ---

def safe_name(value: Any) -> str:
    value = str(value).strip()
    value = re.sub(r"\s+", "_", value)
    value = re.sub(r"[^A-Za-z0-9_]", "", value)
    return value or "column"


def unique_columns(columns) -> list[str]:
    output, seen = [], {}
    for col in columns:
        name = safe_name(col)
        key = name.lower()
        seen[key] = seen.get(key, 0) + 1
        if seen[key] > 1:
            name = f"{name}_{seen[key]}"
        output.append(name)
    return output


@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes, file_name: str, sheet_name=None) -> pd.DataFrame:
    name = file_name.lower()
    bio = io.BytesIO(file_bytes)
    if name.endswith(".csv"):
        try:
            return pd.read_csv(bio)
        except UnicodeDecodeError:
            bio.seek(0)
            return pd.read_csv(bio, encoding="latin-1")
    if name.endswith((".xlsx", ".xlsm")):
        return pd.read_excel(bio, sheet_name=sheet_name or 0, engine="openpyxl")
    if name.endswith(".json"):
        return pd.read_json(bio)
    if name.endswith(".parquet"):
        return pd.read_parquet(bio)
    raise ValueError("Format fail tidak disokong. Gunakan CSV, XLSX, XLSM, JSON, atau Parquet.")


@st.cache_data(show_spinner=False)
def clean_dataframe(source: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    df = source.copy()
    before_rows, before_cols = df.shape
    original_headers = list(df.columns)
    df.columns = unique_columns(df.columns)

    # Clean text values
    for col in df.columns:
        if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
            s = df[col].astype("string")
            s = s.str.replace(r"\s+", " ", regex=True).str.strip()
            s = s.replace("", pd.NA)
            df[col] = s

    # Handle standard missing tokens
    null_tokens = {"na", "n/a", "null", "none", "nan", "-", "--"}
    for col in df.columns:
        if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
            s = df[col].astype("string")
            mask = s.str.lower().isin(null_tokens)
            df.loc[mask.fillna(False), col] = pd.NA

    empty_rows = int(df.isna().all(axis=1).sum()) if len(df.columns) else len(df)
    empty_cols = [c for c in df.columns if df[c].isna().all()]
    df = df.dropna(axis=0, how="all")
    df = df.dropna(axis=1, how="all")
    rows_after_empty = len(df)

    duplicate_count = int(df.duplicated().sum())
    df = df.drop_duplicates().reset_index(drop=True)

    audit = {
        "original_rows": int(before_rows),
        "original_columns": int(before_cols),
        "empty_rows_removed": int(empty_rows),
        "empty_columns_removed": int(len(empty_cols)),
        "duplicates_removed": int(duplicate_count),
        "rows_after_empty_removal": int(rows_after_empty),
        "final_rows": int(len(df)),
        "final_columns": int(len(df.columns)),
        "original_headers": [str(x) for x in original_headers],
        "final_headers": list(df.columns),
    }
    return df, audit


def quality_profile(df: pd.DataFrame) -> pd.DataFrame:
    total = max(len(df), 1)
    rows = []
    for col in df.columns:
        missing = int(df[col].isna().sum())
        unique = int(df[col].nunique(dropna=True))
        dtype = str(df[col].dtype)
        rows.append({
            "Column": col,
            "Data Type": dtype,
            "Missing Count": missing,
            "Missing %": round(missing / total * 100, 2),
            "Unique Values": unique,
            "Duplicate Values": int(df[col].notna().sum() - unique),
        })
    return pd.DataFrame(rows)


def numeric_profile(df: pd.DataFrame) -> pd.DataFrame:
    numeric = df.select_dtypes(include=np.number)
    if numeric.empty:
        return pd.DataFrame()
    desc = numeric.describe().T.reset_index().rename(columns={"index": "Column"})
    return desc.round(3)


def outlier_profile(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for col in df.select_dtypes(include=np.number).columns:
        s = df[col].dropna()
        if len(s) < 4:
            continue
        q1, q3 = s.quantile(.25), s.quantile(.75)
        iqr = q3 - q1
        lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
        count = int(((s < lower) | (s > upper)).sum())
        rows.append({
            "Column": col, "Q1": q1, "Q3": q3, "IQR": iqr,
            "Lower Bound": lower, "Upper Bound": upper,
            "Potential Outliers": count,
            "Outlier %": round(count / max(len(s), 1) * 100, 2)
        })
    return pd.DataFrame(rows).round(3)


def quality_score(df: pd.DataFrame) -> tuple[float, str]:
    if df.empty or len(df.columns) == 0:
        return 0.0, "No usable data"
    missing_ratio = float(df.isna().sum().sum() / (len(df) * len(df.columns)))
    completeness = max(0.0, 1.0 - missing_ratio)
    score = round(completeness * 100, 1)
    label = "Excellent" if score >= 98 else "Good" if score >= 90 else "Needs review" if score >= 70 else "Poor"
    return score, label


# --- FEATURE 1: UNIVERSAL FILE CONVERTER ENGINE ---

def convert_dataframe(df: pd.DataFrame, target_format: str) -> tuple[bytes, str, str]:
    target_format = target_format.lower()
    
    if target_format == "csv":
        data = df.to_csv(index=False).encode("utf-8-sig")
        return data, "text/csv", "converted_data.csv"
        
    elif target_format == "excel (xlsx)":
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Data")
        output.seek(0)
        return output.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "converted_data.xlsx"
        
    elif target_format == "json":
        data = df.to_json(orient="records", indent=2).encode("utf-8")
        return data, "application/json", "converted_data.json"
        
    elif target_format == "parquet":
        output = io.BytesIO()
        df.to_parquet(output, index=False)
        output.seek(0)
        return output.getvalue(), "application/octet-stream", "converted_data.parquet"
        
    elif target_format == "html table":
        html_str = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>Converted Data Table</title>
            <script src="https://cdn.tailwindcss.com"></script>
        </head>
        <body class="bg-slate-50 p-6">
            <div class="max-w-7xl mx-auto bg-white rounded-xl shadow-md p-6 overflow-x-auto">
                <h2 class="text-xl font-bold mb-4 text-slate-800">Data Table Export</h2>
                {df.to_html(classes="min-w-full divide-y divide-gray-200 text-sm", index=False)}
            </div>
        </body>
        </html>
        """
        return html_str.encode("utf-8"), "text/html", "converted_data.html"

    elif target_format == "markdown":
        md_str = df.to_markdown(index=False)
        return md_str.encode("utf-8"), "text/markdown", "converted_data.md"
        
    raise ValueError("Format sasaran tidak sah.")


# --- FEATURE 2: SCAN FILE TO EXTRACT (OCR VIA GEMINI API) ---

async def scan_file_with_gemini(image_bytes: bytes, mime_type: str, custom_prompt: str = "") -> str:
    """Uses Gemini 3 Flash to perform OCR / Data Extraction from documents and images."""
    api_key = "" # Uses auto-injected runtime key
    api_url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key={api_key}"
    
    base64_image = base64.b64encode(image_bytes).decode("utf-8")
    
    prompt = custom_prompt or (
        "Ekstrak semua jadual atau data berstruktur daripada imej/dokumen ini. "
        "Kembalikan jawapan HANYA dalam format JSON Array of Objects supaya boleh ditukar ke Pandas Dataframe. "
        "Jangan sertakan teks penerangan lain, hanya JSON asli."
    )
    
    payload = {
        "contents": [{
            "role": "user",
            "parts": [
                {"text": prompt},
                {
                    "inlineData": {
                        "mimeType": mime_type,
                        "data": base64_image
                    }
                }
            ]
        }]
    }

    import urllib.request
    req = urllib.request.Request(
        api_url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"}
    )
    
    try:
        with urllib.request.urlopen(req) as response:
            res_data = json.loads(response.read().decode("utf-8"))
            text = res_data["candidates"][0]["content"]["parts"][0]["text"]
            return text
    except Exception as e:
        return f"Error executing OCR Extraction: {str(e)}"


# --- EXPORT REPORT BUILDERS (EXCEL & PDF) ---

def build_excel(df: pd.DataFrame, audit: dict, quality: pd.DataFrame,
                numeric: pd.DataFrame, outliers: pd.DataFrame,
                source_name: str, score: float, score_label: str) -> bytes:
    output = io.BytesIO()
    summary = pd.DataFrame([
        ["Product", "Data Analyst Pro — AI Edition"],
        ["Version", APP_VERSION],
        ["Source file", source_name],
        ["Generated at (UTC)", datetime.now(timezone.utc).isoformat(timespec="seconds")],
        ["Rows after cleaning", len(df)],
        ["Columns after cleaning", len(df.columns)],
        ["Completeness score", score],
        ["Score interpretation", score_label],
        ["Important", "Score measures completeness only; it does not certify accuracy."],
    ], columns=["Metric", "Value"])

    # Safe audit dict conversion to prevent serialization errors
    audit_clean = {k: (int(v) if isinstance(v, (np.integer, int)) else v) for k, v in audit.items()}
    audit_df = pd.DataFrame([{"Action": k, "Result": json.dumps(v, ensure_ascii=False) if isinstance(v, (list, dict)) else str(v)} for k, v in audit_clean.items()])

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Cleaned Data", index=False)
        summary.to_excel(writer, sheet_name="Executive Summary", index=False)
        quality.to_excel(writer, sheet_name="Data Quality", index=False)
        if not numeric.empty:
            numeric.to_excel(writer, sheet_name="Statistics", index=False)
        if not outliers.empty:
            outliers.to_excel(writer, sheet_name="Outlier Review", index=False)
        audit_df.to_excel(writer, sheet_name="Audit Log", index=False)

        wb = writer.book
        navy = "1E293B"
        white = "FFFFFF"
        thin = Side(style="thin", color="E2E8F0")
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color=white)
                cell.fill = PatternFill("solid", fgColor=navy)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = Border(bottom=thin)
            ws.row_dimensions[1].height = 26
            for col_cells in ws.columns:
                letter = get_column_letter(col_cells[0].column)
                max_len = 0
                for cell in col_cells[:300]:
                    val = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, min(len(val), 50))
                ws.column_dimensions[letter].width = min(max(max_len + 3, 12), 45)
    output.seek(0)
    return output.getvalue()


def make_pdf_summary(df: pd.DataFrame, quality: pd.DataFrame, audit: dict,
                     score: float, score_label: str, source_name: str) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.units import mm

    buff = io.BytesIO()
    doc = SimpleDocTemplate(buff, pagesize=A4, rightMargin=16*mm, leftMargin=16*mm,
                            topMargin=16*mm, bottomMargin=16*mm)
    styles = getSampleStyleSheet()
    story = [
        Paragraph("DATA ANALYST PRO — EXECUTIVE REPORT", styles["Title"]),
        Spacer(1, 8),
        Paragraph(f"<b>Source:</b> {source_name}", styles["Normal"]),
        Paragraph(f"<b>Generated:</b> {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}", styles["Normal"]),
        Spacer(1, 12),
    ]
    summary_data = [
        ["Metric", "Result"],
        ["Rows after cleaning", f"{len(df):,}"],
        ["Columns after cleaning", f"{len(df.columns):,}"],
        ["Completeness score", f"{score:.1f}% ({score_label})"],
        ["Duplicate rows removed", str(audit["duplicates_removed"])],
        ["Empty rows removed", str(audit["empty_rows_removed"])],
    ]
    t = Table(summary_data, colWidths=[75*mm, 85*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1E293B")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("GRID", (0,0), (-1,-1), .4, colors.HexColor("#CBD5E1")),
        ("PADDING", (0,0), (-1,-1), 7),
    ]))
    story += [t, Spacer(1, 14), Paragraph("Data Quality Overview", styles["Heading2"])]
    qdata = [["Column", "Type", "Missing", "Missing %", "Unique"]]
    for _, row in quality.head(20).iterrows():
        qdata.append([str(row["Column"])[:22], str(row["Data Type"])[:12],
                      str(row["Missing Count"]), f'{row["Missing %"]:.1f}%',
                      str(row["Unique Values"])])
    qt = Table(qdata, repeatRows=1, colWidths=[48*mm, 32*mm, 25*mm, 27*mm, 25*mm])
    qt.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#334155")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#E2E8F0")),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("PADDING", (0,0), (-1,-1), 5),
    ]))
    story += [qt, Spacer(1, 12),
              Paragraph("<i>Disclaimer: Completeness score measures non-missing ratio only, not factual business correctness.</i>", styles["Italic"])]
    doc.build(story)
    buff.seek(0)
    return buff.getvalue()


# --- APP HEADER UI ---
st.markdown(f"""
<div class="hero-card">
    <div style="display: flex; justify-content: space-between; align-items: center;">
        <div>
            <div class="hero-title">⚡ Data Analyst Pro</div>
            <div class="hero-subtitle">Automated Data Cleaning, AI Document OCR Scan, Universal Converter & Executive Analytics</div>
        </div>
        <div>
            <span class="badge-tag badge-purple">Version {APP_VERSION}</span>
            <span class="badge-tag badge-success" style="margin-left: 6px;">AI Enabled</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)


# --- SIDEBAR CONTROL PANEL ---
with st.sidebar:
    st.markdown("### ⚙️ Mode & Toolkit")
    app_mode = st.radio(
        "Pilih Modul Utama:",
        ["📊 Main Data Analytics Engine", "🔄 Universal File Converter", "🔍 Scan File to Extract (AI OCR)"]
    )
    st.divider()
    st.markdown("##### 🚀 Fast Features")
    st.caption("✓ 100% Client-Session Privacy")
    st.caption("✓ High-Speed In-Memory Processing")
    st.caption("✓ Interactive Plotly Charting")


# ==============================================================================
# MODULE 1: MAIN DATA ANALYTICS ENGINE
# ==============================================================================
if app_mode == "📊 Main Data Analytics Engine":
    uploaded = st.file_uploader("Muat naik Dataset anda (XLSX, XLSM, CSV, JSON, PARQUET)", type=["xlsx", "xlsm", "csv", "json", "parquet"])
    
    if uploaded is None:
        st.info("👋 **Selamat datang!** Sila muat naik fail data anda di atas untuk memulakan analisis automatik.")
        cols = st.columns(4)
        cards_info = [
            ("✨ Cleaning V1", "Pembersihan header automatik & penyingkiran duplikasi."),
            ("🛡️ Quality Profile", "Laporan kelengkapan data & pengesan outlier IQR."),
            ("📈 Interactive Charts", "Carta Plotly interaktif & matriks korelasi."),
            ("📑 Professional PDF/Excel", "Jana laporan bertaraf korporat dengan 1-klik.")
        ]
        for col, (title, desc) in zip(cols, cards_info):
            col.markdown(f"""
            <div class="metric-container">
                <h4 style="margin:0 0 6px 0; color:#1e293b;">{title}</h4>
                <p style="font-size:0.85rem; color:#64748b; margin:0;">{desc}</p>
            </div>
            """, unsafe_allow_html=True)
        st.stop()

    file_bytes = uploaded.getvalue()
    try:
        if uploaded.name.lower().endswith((".xlsx", ".xlsm")):
            xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
            sheet = st.selectbox("Pilih Lembaran Kerja (Worksheet)", xls.sheet_names)
            raw_df = load_data(file_bytes, uploaded.name, sheet)
        else:
            raw_df = load_data(file_bytes, uploaded.name)
    except Exception as exc:
        st.error(f"Gagal membaca fail: {exc}")
        st.stop()

    if raw_df.empty:
        st.warning("Fail yang dipilih tidak mempunyai baris data.")
        st.stop()

    # Executive Overview Cards
    cleaned_df, audit = clean_dataframe(raw_df)
    quality = quality_profile(cleaned_df)
    numeric = numeric_profile(cleaned_df)
    outliers = outlier_profile(cleaned_df)
    score, score_label = quality_score(cleaned_df)

    st.markdown("### 📌 Executive Dashboard Overview")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Clean Rows", f"{len(cleaned_df):,}", delta=f"{len(cleaned_df)-len(raw_df):,} vs source")
    m2.metric("Clean Columns", f"{len(cleaned_df.columns):,}")
    m3.metric("Completeness Score", f"{score:.1f}%")
    m4.metric("Duplicates Removed", f'{audit["duplicates_removed"]:,}')

    tab1, tab2, tab3, tab4 = st.tabs(["📋 Data Cleaned", "🛡️ Quality & Outliers", "📊 Statistics & Interactive Charts", "📥 Export Reports"])

    with tab1:
        st.markdown("#### Preview Data Bersih (First 1,000 Rows)")
        st.dataframe(cleaned_df.head(1000), use_container_width=True)
        st.download_button("📥 Muat Naik CSV Data Bersih", cleaned_df.to_csv(index=False).encode("utf-8-sig"),
                           file_name="cleaned_data.csv", mime="text/csv")

    with tab2:
        st.markdown("#### Profil Kelengkapan Data Lajur")
        st.dataframe(quality, use_container_width=True)
        if not outliers.empty:
            st.markdown("#### Potensi Outlier Nombor (Kriteria IQR)")
            st.dataframe(outliers, use_container_width=True)
        else:
            st.caption("Tiada lajur nombor dengan data yang mencukupi untuk analisis outlier.")

    with tab3:
        if not numeric.empty:
            st.markdown("#### Statistik Deskriptif Nombor")
            st.dataframe(numeric, use_container_width=True)
            
            numeric_cols = list(cleaned_df.select_dtypes(include=np.number).columns)
            if numeric_cols:
                col_sel1, col_sel2 = st.columns(2)
                with col_sel1:
                    chart_col = st.selectbox("Pilih Lajur Nombor", numeric_cols)
                with col_sel2:
                    chart_type = st.selectbox("Jenis Carta Interaktif", ["Histogram Dynamic", "Box Plot (Outliers)", "Line Chart"])
                
                if chart_type == "Histogram Dynamic":
                    fig = px.histogram(cleaned_df, x=chart_col, title=f"Taburan Data: {chart_col}", template="plotly_white", color_discrete_sequence=['#6366f1'])
                elif chart_type == "Box Plot (Outliers)":
                    fig = px.box(cleaned_df, y=chart_col, title=f"Box Plot & Outliers: {chart_col}", template="plotly_white", color_discrete_sequence=['#ec4899'])
                else:
                    fig = px.line(cleaned_df, y=chart_col, title=f"Trend mengikut Baris: {chart_col}", template="plotly_white", color_discrete_sequence=['#0ea5e9'])
                
                st.plotly_chart(fig, use_container_width=True)

            if len(numeric_cols) >= 2:
                st.markdown("#### Matriks Korelasi Interaktif")
                corr = cleaned_df[numeric_cols].corr(numeric_only=True).round(3)
                fig_corr = px.imshow(corr, text_auto=True, aspect="auto", color_continuous_scale="Blues", title="Correlation Heatmap")
                st.plotly_chart(fig_corr, use_container_width=True)

        categorical_cols = [c for c in cleaned_df.columns if cleaned_df[c].dtype == "object" or str(cleaned_df[c].dtype).startswith("string")]
        if categorical_cols:
            st.markdown("#### Analisis Frekuensi Kategori")
            cat_col = st.selectbox("Pilih Lajur Kategori", categorical_cols)
            top_counts = cleaned_df[cat_col].value_counts().head(15).reset_index()
            top_counts.columns = [cat_col, "Count"]
            fig_bar = px.bar(top_counts, x="Count", y=cat_col, orientation="h", title=f"Top 15 Kategori: {cat_col}", color="Count", color_continuous_scale="Viridis", template="plotly_white")
            fig_bar.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_bar, use_container_width=True)

    with tab4:
        st.markdown("#### Muat Turun Laporan Analisis")
        c1, c2 = st.columns(2)
        excel_bytes = build_excel(cleaned_df, audit, quality, numeric, outliers, uploaded.name, score, score_label)
        c1.download_button("📊 Download Excel Executive Report (.xlsx)", excel_bytes, file_name="Data_Analyst_Pro_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        
        try:
            pdf_bytes = make_pdf_summary(cleaned_df, quality, audit, score, score_label, uploaded.name)
            c2.download_button("📄 Download Executive PDF Summary (.pdf)", pdf_bytes, file_name="Executive_Summary.pdf", mime="application/pdf", use_container_width=True)
        except Exception as exc:
            c2.warning(f"Modul PDF tidak dapat diakses: {exc}")


# ==============================================================================
# MODULE 2: UNIVERSAL FILE CONVERTER
# ==============================================================================
elif app_mode == "🔄 Universal File Converter":
    st.markdown("### 🔄 Universal Data File Converter")
    st.caption("Tukar format fail dataset anda serta-merta ke pelbagai format tanpa gangguan data.")
    
    conv_file = st.file_uploader("Muat Naik Fail Untuk Ditukar (CSV, Excel, JSON, Parquet)", type=["csv", "xlsx", "json", "parquet"], key="conv_file")
    
    if conv_file:
        file_bytes = conv_file.getvalue()
        df_conv = load_data(file_bytes, conv_file.name)
        
        st.write(f"📁 **Pratonton Fail:** {conv_file.name} ({len(df_conv):,} baris × {len(df_conv.columns):,} lajur)")
        st.dataframe(df_conv.head(10), use_container_width=True)
        
        st.divider()
        c_format, c_action = st.columns([2, 1])
        
        with c_format:
            target_fmt = st.selectbox(
                "Pilih Format Sasaran (Target Format):",
                ["CSV", "Excel (XLSX)", "JSON", "Parquet", "HTML Table", "Markdown"]
            )
            
        with c_action:
            st.write(" ")
            st.write(" ")
            converted_data, mime_type, file_out_name = convert_dataframe(df_conv, target_fmt)
            st.download_button(
                label=f"🚀 Tukar & Muat Turun ({target_fmt})",
                data=converted_data,
                file_name=file_out_name,
                mime=mime_type,
                use_container_width=True
            )


# ==============================================================================
# MODULE 3: SCAN FILE TO EXTRACT (AI OCR)
# ==============================================================================
elif app_mode == "🔍 Scan File to Extract (AI OCR)":
    st.markdown("### 🔍 Scan Document to Extract Data (AI OCR)")
    st.caption("Imbas imej dokumen, resit, atau jadual untuk mengekstrak data secara automatik menggunakan teknologi AI Gemini 3 Flash.")
    
    ocr_file = st.file_uploader("Muat naik Dokumen / Imej Imbasan (PNG, JPG, WEBP)", type=["png", "jpg", "jpeg", "webp"])
    
    if ocr_file:
        img_bytes = ocr_file.getvalue()
        st.image(img_bytes, caption="Dokumen Diimbas", use_container_width=True)
        
        user_prompt = st.text_area(
            "Arahan Ekstraksi Khusus (Optional):",
            placeholder="Contoh: Ekstrak semua item dan harga daripada resit ini ke dalam bentuk jadual."
        )
        
        if st.button("✨ Ekstrak Data Dokumen Sekarang", type="primary"):
            with st.spinner("AI sedang membaca dan mengekstrak data daripada imej..."):
                import asyncio
                extracted_text = asyncio.run(scan_file_with_gemini(img_bytes, ocr_file.type, user_prompt))
                
                st.markdown("#### 📄 Hasil Ekstraksi Data:")
                st.code(extracted_text, language="json")
                
                # Try parsing as JSON to dataframe
                try:
                    # Clean markdown code block if present
                    clean_json = re.sub(r"^```json\s*", "", extracted_text, flags=re.MULTILINE)
                    clean_json = re.sub(r"^```\s*", "", clean_json, flags=re.MULTILINE).strip()
                    parsed_json = json.loads(clean_json)
                    
                    df_ocr = pd.DataFrame(parsed_json)
                    st.success("✅ AI berjaya menukar imej kepada Jadual Data Streamed!")
                    st.dataframe(df_ocr, use_container_width=True)
                    
                    st.download_button(
                        "📥 Muat Naik Hasil Ekstraksi CSV",
                        df_ocr.to_csv(index=False).encode("utf-8-sig"),
                        file_name="ocr_extracted_data.csv",
                        mime="text/csv"
                    )
                except Exception:
                    st.info("Nota: Hasil ekstrak teks dipaparkan di atas. Jika anda inginkan format CSV automatik, pastikan output Gemini memberikan JSON Array valid.")




