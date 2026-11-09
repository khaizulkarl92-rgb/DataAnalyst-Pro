⚡ Data Analyst Pro — V6 Enterprise AI Edition

Data Analyst Pro V6 ialah platform analisis data automatik, pembersihan data (ETL mini), pengimbas dokumen berkuasa AI (OCR), dan penukar format fail serbaguna yang dibina khas menggunakan **Streamlit**, **Pandas**, **Plotly**, dan **Gemini 3 Flash AI**.

---

## ✨ Ciri-Ciri Utama (Features)

1. **📊 Main Data Analytics & Cleaning Engine**
   - Pembersihan automatik: normalisasi header, penyingkiran ruang kosong, dan deduplikasi data.
   - Matriks korelasi interaktif dan carta visualisasi **Plotly** (Histogram, Box Plot Outlier, Line Chart).
   - Pengesan potensi *outliers* menggunakan kaedah **IQR**.
   - Skor Kelengkapan Data (*Completeness Score*) automatik.

2. **🔄 Universal Data File Converter**
   - Penukaran format fail automatik antara: **CSV, Excel (XLSX), JSON, Parquet, HTML Table, dan Markdown (.md)**.

3. **🔍 Scan Document to Extract (AI Smart OCR)**
   - Integrasi **Gemini 3 Flash AI** untuk mengekstrak jadual/teks daripada imej imbasan (`.png`, `.jpg`, `.webp`) atau resit terus ke dalam format JSON/CSV.

4. **📑 Laporan Ekssekutif Korporat**
   - Muat turun laporan bertajuk kemas dalam bentuk **Excel Workbook (.xlsx)** (berserta tab Audit Trail) dan **Laporan Ringkasan PDF**.

---

## 🚀 Jalankan Secara Tempatan (Local Setup)

Disyorkan menggunakan Python 3.10+.

```bash
# 1. Cipta virtual environment
python -m venv .venv

# 2. Aktifkan persekitaran (Windows)
.venv\Scripts\activate
# (macOS/Linux)
# source .venv/bin/activate

# 3. Install pustaka yang diperlukan
pip install -r requirements.txt

# 4. Jalankan aplikasi Streamlit
streamlit run app.py
