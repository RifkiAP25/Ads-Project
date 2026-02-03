import streamlit as st
import pandas as pd
import io
from datetime import datetime

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(
    page_title="CSV Iklan → Excel Berwarna",
    layout="centered"
)

st.title("📊 CSV Iklan → Excel Berwarna")
st.caption("Upload CSV iklan Shopee → otomatis rapi → download Excel berwarna")

# =========================
# UPLOAD CSV
# =========================
uploaded_file = st.file_uploader(
    "Upload file CSV iklan Shopee",
    type=["csv"]
)

# =========================
# LOAD CSV (PALING KEBAL)
# =========================
@st.cache_data
def load_uploaded_csv(file):
    # baca sebagai text mentah
    file.seek(0)
    raw = file.read().decode("utf-8", errors="ignore")
    lines = raw.splitlines()

    # cari baris header
    HEADER_KEYS = ["Nama Iklan", "Nama Iklan/Produk"]
    header_idx = None

    for i, line in enumerate(lines[:30]):
        if any(key in line for key in HEADER_KEYS):
            header_idx = i
            break

    if header_idx is None:
        raise ValueError("Header Nama Iklan tidak ditemukan.")

    # tentukan delimiter dari header
    header_line = lines[header_idx]
    delimiter = ";" if header_line.count(";") > header_line.count(",") else ","

    # potong file mulai header
    clean_csv = "\n".join(lines[header_idx:])

    # baca ulang sebagai CSV bersih
    df = pd.read_csv(
        io.StringIO(clean_csv),
        sep=delimiter,
        engine="python",
        on_bad_lines="skip"
    )

    # bersihin kolom
    df.columns = df.columns.str.strip()

    return df

# =========================
# NORMALISASI KOLOM NAMA IKLAN
# =========================
def normalize_nama_iklan_column(df):
    kandidat = ["Nama Iklan", "Nama Iklan/Produk"]

    for col in kandidat:
        if col in df.columns:
            return df.rename(columns={col: "Nama Iklan"})

    raise ValueError("Kolom Nama Iklan / Nama Iklan/Produk tidak ditemukan.")

# =========================
# STYLING LOGIC
# =========================
def highlight_row(row):
    styles = [''] * len(row)

    roas = row.get('Efektifitas Iklan')
    sales = row.get('Penjualan Langsung (GMV Langsung)')
    cost = row.get('Biaya')

    if pd.isna(roas) or pd.isna(sales) or pd.isna(cost):
        return styles

    nama_idx = row.index.get_loc('Nama Iklan')
    sales_idx = row.index.get_loc('Produk Terjual')

    # PRIORITAS 1 — rugi keras
    if sales == 0 and cost >= 10000:
        return ['color: red'] * len(row)

    # PRIORITAS 2 — aman
    if sales == 0 and cost < 10000:
        return styles

    # PRIORITAS 3 — ROAS
    if roas < 8:
        styles = ['background-color: red'] * len(row)
    elif roas < 10:
        styles = ['background-color: yellow'] * len(row)
    else:
        styles = ['background-color: lightgreen'] * len(row)

    # OVERRIDE BIRU
    if sales == 0:
        styles[nama_idx] = 'background-color: lightblue'
        styles[sales_idx] = 'background-color: lightblue'

    return styles

# =========================
# PROCESS & DOWNLOAD
# =========================
if uploaded_file:
    if st.button("🚀 Proses & Download Excel"):
        with st.spinner("Memproses file..."):
            df = load_uploaded_csv(uploaded_file)
            df = normalize_nama_iklan_column(df)

            # convert numerik aman
            numeric_cols = [
                'Efektifitas Iklan',
                'Penjualan Langsung (GMV Langsung)',
                'Biaya'
            ]

            for col in numeric_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"DataIklan_Styled_{timestamp}.xlsx"

            buffer = io.BytesIO()
            df.style.apply(highlight_row, axis=1).to_excel(
                buffer,
                engine="openpyxl",
                index=False
            )
            buffer.seek(0)

        st.success("Selesai! File siap di-download 👇")

        st.download_button(
            label="⬇️ Download Excel Berwarna",
            data=buffer,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
