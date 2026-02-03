import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(
    page_title="CSV Iklan → Excel Berwarna",
    layout="centered"
)

st.title("📊 CSV Iklan → Excel Berwarna")
st.caption("Upload CSV iklan Shopee → otomatis rapi → download Excel laporan")

# =========================
# UPLOAD CSV
# =========================
uploaded_file = st.file_uploader(
    "Upload file CSV iklan Shopee",
    type=["csv"]
)

csv_mode = st.radio(
    "Jenis CSV yang di-upload",
    options=[
        "CSV Keseluruhan (Normal)",
        "CSV Grup Iklan (hanya iklan produk)"
    ],
    horizontal=True
)

# =========================
# LOAD CSV (PALING KEBAL)
# =========================
@st.cache_data
def load_uploaded_csv(file):
    file.seek(0)
    raw = file.read().decode("utf-8", errors="ignore")
    lines = raw.splitlines()

    HEADER_KEYS = ["Nama Iklan", "Nama Iklan/Produk"]
    header_idx = None

    for i, line in enumerate(lines[:30]):
        if any(k in line for k in HEADER_KEYS):
            header_idx = i
            break

    if header_idx is None:
        raise ValueError("Header Nama Iklan tidak ditemukan")

    delimiter = ";" if lines[header_idx].count(";") > lines[header_idx].count(",") else ","
    clean_csv = "\n".join(lines[header_idx:])

    df = pd.read_csv(
        io.StringIO(clean_csv),
        sep=delimiter,
        engine="python",
        on_bad_lines="skip"
    )

    df.columns = df.columns.str.strip()
    return df

# =========================
# NORMALISASI NAMA IKLAN
# =========================
def normalize_nama_iklan_column(df):
    for col in ["Nama Iklan", "Nama Iklan/Produk"]:
        if col in df.columns:
            return df.rename(columns={col: "Nama Iklan"})
    raise ValueError("Kolom Nama Iklan tidak ditemukan")

# =========================
# PEMENDEK NAMA IKLAN (RINGKASAN SAJA)
# =========================
def short_nama_iklan(nama):
    if pd.isna(nama):
        return nama

    text = str(nama).strip()

    # 1️⃣ KUNCI GRUP IKLAN (STOP DI SINI)
    if text.lower().startswith("grup iklan"):
        return text.split(" - ")[0]

    # 2️⃣ buang tag [SB], [TEST], dll
    text = re.sub(r"\[.*?\]", "", text).strip()
    
    feature_blacklist = {
        "busui","friendly","bahan","soft","ultimate","ultimates",
        "motif","size","ukuran","promo","diskon","broad","testing",
        "rayon","katun","cotton","silk","sutra","viscose",
        "linen","polyester","jersey","crepe","chiffon",
        "woolpeach","baloteli","babyterry",
        "pink","hitam","black","putih","white","navy","biru","blue",
        "merah","red","hijau","green","coklat","brown",
        "abu","abu-abu","grey","gray","cream","krem","beige",
        "maroon","ungu","purple","tosca","olive","sage"
    }

    store_blacklist = {
        "official","shop","store","boutique","fashion",
        "my","zahir","myzahir","by","original","premium"
    }

    category_keywords = {
        "gamis","dress","tunik","abaya","set",
        "blouse","khimar","rok","pashmina","hijab","outer",
    }

    context_blacklist = {
        "terbaru","new","update","launch","launching",
        "viral","hits","best","seller","bestseller",
        "kondangan","lebaran","ramadhan","ramadan",
        "harian","pesta","formal","casual",
        "trend","trending","populer",
        "2024","2025","2026","2027", "2028", "2029", "2030"
    }

    parts = re.split(r"\s*[-|]\s*", text)

    product_keywords = {"dress", "gamis", "set"}
    product_candidates = []

    for part in parts:
        words = part.split()
        words_lower = [w.lower() for w in words]

        if not any(w in product_keywords for w in words_lower):
            continue

        while words_lower and words_lower[0] in store_blacklist:
            words_lower.pop(0)
            words.pop(0)

        unique_words = [
            w for w in words_lower
            if w not in store_blacklist
            and w not in feature_blacklist
            and w not in context_blacklist
            and w not in category_keywords
        ]

        if unique_words:
            product_candidates.append(words)

    if product_candidates:
        best_words = product_candidates[-1]
        return " ".join(best_words[:3])

    def score(part):
        s = 0
        for w in part.lower().split():
            if w in store_blacklist:
                s -= 3
            elif w in feature_blacklist:
                s -= 1
            elif w in context_blacklist:
                s -= 2
            elif w in category_keywords:
                s += 1
            else:
                s += 3
        return s

    best = max(parts, key=score)
    return " ".join(best.split()[:3])


# =========================
# STYLING DATA_IKLAN
# =========================
def highlight_row(row):
    styles = [''] * len(row)

    roas = row.get('Efektifitas Iklan')
    sales = row.get('Produk Terjual')
    gmv = row.get('Penjualan Langsung (GMV Langsung)')
    cost = row.get('Biaya')

    # jika kolom belum ada / NaN -> jangan crash
    if pd.isna(roas) or pd.isna(sales) or pd.isna(cost):
        return styles

    try:
        nama_idx = row.index.get_loc('Nama Iklan')
    except Exception:
        nama_idx = None
    try:
        gmv_idx = row.index.get_loc('Penjualan Langsung (GMV Langsung)')
    except Exception:
        gmv_idx = None

    # 🔴 MERAH TIPE A — rugi keras
    if sales == 0 and cost >= 10000:
        return ['color: red'] * len(row)

    # ⚪ NETRAL — pemanasan
    if sales == 0 and cost < 10000:
        return styles   # important: keep neutral

    # 🟥🟨🟩 WARNA ROAS
    if roas < 8:
        styles = ['background-color: red'] * len(row)
    elif roas < 10:
        styles = ['background-color: yellow'] * len(row)
    else:
        styles = ['background-color: lightgreen'] * len(row)

    # 🔵 OVERLAY BIRU — assist only (hanya highlight Nama Iklan & GMV col)
    if sales > 0 and (gmv == 0 or pd.isna(gmv)):
        if nama_idx is not None:
            styles[nama_idx] = 'background-color: lightblue'
        if gmv_idx is not None:
            styles[gmv_idx] = 'background-color: lightblue'

    return styles

# =========================
# KATEGORI IKLAN
# =========================
def get_iklan_color(row, csv_mode):
    roas = row.get('Efektifitas Iklan')
    sales = row.get('Produk Terjual')
    cost = row.get('Biaya')

    # safety
    if pd.isna(sales) or pd.isna(cost):
        return None

    # TANPA KONVERSI BESAR → keluar dari ringkasan
    if sales == 0 and cost >= 10000:
        return None

    if sales == 0 and cost < 10000:
        return None

    # CSV Grup: jika ROAS missing, fallback ke sales>0 -> HIJAU
    if csv_mode == "CSV Grup Iklan (hanya iklan produk)":
        if pd.isna(roas):
            return "HIJAU" if sales > 0 else None
        # else fall through to roas logic

    # ROAS-based
    if pd.isna(roas) or roas < 8:
        return "MERAH"
    elif roas < 10:
        return "KUNING"
    else:
        return "HIJAU"

# =========================
# PROCESS & EXPORT
# =========================
if uploaded_file:
    if st.button("🚀 Proses & Download Excel"):
        with st.spinner("Memproses data..."):
            # load + normalize
            df = load_uploaded_csv(uploaded_file)
            df = normalize_nama_iklan_column(df)

            # mark aggregate/group rows (baris yang mulai dengan "grup")
            df["IS_AGGREGATE"] = df["Nama Iklan"].astype(str).str.lower().str.match(r'^\s*grup\b')

            # convert numerik aman
            for col in [
                "Efektifitas Iklan",
                "Produk Terjual",
                "Penjualan Langsung (GMV Langsung)",
                "Biaya"
            ]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce")

            # IS_BIRU
            df["IS_BIRU"] = (
                (df.get("Produk Terjual", 0) > 0) &
                (df.get("Penjualan Langsung (GMV Langsung)", 0) == 0)
            )

            # Nama Ringkasan: jangan rubah Nama Iklan asli for aggregates
            df["Nama Ringkasan"] = df["Nama Iklan"].where(
                df["IS_AGGREGATE"],
                df["Nama Iklan"].apply(short_nama_iklan)
            )

            # Kategori (apply with csv_mode)
            df["Kategori"] = df.apply(lambda row: get_iklan_color(row, csv_mode), axis=1)

            # >>> CHANGE HERE <<<:
            # Hanya exclude aggregate rows dari ringkasan/tanpa_konversi jika user memilih CSV Grup Iklan
            if csv_mode == "CSV Grup Iklan (hanya iklan produk)":
                df_nonagg = df[~df["IS_AGGREGATE"]].copy()
            else:
                # CSV Keseluruhan: include aggregate rows in summaries (user requested)
                df_nonagg = df.copy()

            # --- build ordered flat list (MERAH -> KUNING -> HIJAU -> BIRU) ONLY FOR numbering mode ---
            ordered_for_numbering = []
            for kat in ["MERAH", "KUNING", "HIJAU"]:
                for name in df_nonagg[df_nonagg["Kategori"] == kat]["Nama Ringkasan"]:
                    ordered_for_numbering.append({"nama": name, "kategori": kat})
            for name in df_nonagg[df_nonagg["IS_BIRU"]]["Nama Ringkasan"]:
                ordered_for_numbering.append({"nama": name, "kategori": "BIRU"})

            # Build per-col content depending on csv_mode
            per_col = {"MERAH": [], "KUNING": [], "HIJAU": [], "BIRU": []}

            if csv_mode == "CSV Keseluruhan (Normal)":
                # numbered vertical lists (global numbering across all categories)
                for idx, item in enumerate(ordered_for_numbering, start=1):
                    numbered = f"{idx}. {item['nama']}"
                    per_col[item["kategori"]].append(numbered)
            else:
                # CSV Grup Iklan (hanya iklan produk): no numbering, comma-separated (trailing comma per item)
                for kat in ["MERAH", "KUNING", "HIJAU"]:
                    names = df_nonagg[df_nonagg["Kategori"] == kat]["Nama Ringkasan"].tolist()
                    # ensure each has trailing comma
                    per_col[kat] = [f"{n}," for n in names]
                # BIRU (IS_BIRU)
                names_biru = df_nonagg[df_nonagg["IS_BIRU"]]["Nama Ringkasan"].tolist()
                per_col["BIRU"] = [f"{n}," for n in names_biru]

            # >10K tanpa konversi (use df_nonagg as per-mode)
            tanpa_konversi_df = (
                df_nonagg[(df_nonagg.get("Produk Terjual", 0) == 0) & (df_nonagg.get("Biaya", 0) >= 10000)]
                [["Nama Ringkasan", "Biaya"]]
                .rename(columns={"Nama Ringkasan": "Nama Iklan"})
                .sort_values("Biaya", ascending=False)
            )

            # =========================
            # EXPORT EXCEL
            # =========================
            buffer = io.BytesIO()
            original_name = uploaded_file.name
            base_name = original_name.rsplit(".", 1)[0]
            filename = f"{base_name}.xlsx"

            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                # Sheet 1 — DATA (keep original Nama Iklan column in sheet; includes aggregate rows)
                df.style.apply(highlight_row, axis=1).to_excel(
                    writer, sheet_name="DATA_IKLAN", index=False
                )

                # create RINGKASAN_IKLAN sheet manually
                wb = writer.book
                from openpyxl.styles import Font, Alignment

                # remove if exists
                if "RINGKASAN_IKLAN" in wb.sheetnames:
                    wb.remove(wb["RINGKASAN_IKLAN"])
                ws_ring = wb.create_sheet("RINGKASAN_IKLAN")

                headers = ["MERAH", "KUNING", "HIJAU", "BIRU"]
                color_map = {
                    "MERAH": "FF0000",
                    "KUNING": "000000",
                    "HIJAU": "00AA00",
                    "BIRU": "0066CC"
                }

                # write headers
                for c_idx, h in enumerate(headers, start=1):
                    cell = ws_ring.cell(row=1, column=c_idx, value=h)
                    cell.font = Font(bold=True)

                # write content depending on mode
                if csv_mode == "CSV Keseluruhan (Normal)":
                    # put vertical numbered list in each color's single cell using newline
                    for c_idx, key in enumerate(headers, start=1):
                        items = per_col.get(key, [])
                        if items:
                            text = "\n".join(items)
                            cell = ws_ring.cell(row=2, column=c_idx, value=text)
                            cell.font = Font(color=color_map[key])
                            cell.alignment = Alignment(wrap_text=True, vertical="top")
                        else:
                            ws_ring.cell(row=2, column=c_idx, value="")
                else:
                    # CSV Grup Iklan: comma-separated in single cell (row 2), each item with trailing comma
                    for c_idx, key in enumerate(headers, start=1):
                        items = per_col.get(key, [])
                        if items:
                            # items already have trailing commas; join with space
                            joined = " ".join(items)
                            # ensure trailing comma at end
                            if not joined.strip().endswith(","):
                                joined = joined + ","
                            cell = ws_ring.cell(row=2, column=c_idx, value=joined)
                            cell.font = Font(color=color_map[key])
                            cell.alignment = Alignment(wrap_text=True, vertical="top")
                        else:
                            ws_ring.cell(row=2, column=c_idx, value="")

                # adjust column widths a bit (optional)
                for i in range(1, 5):
                    ws_ring.column_dimensions[chr(64 + i)].width = 40

                # Sheet 3 — >10K_TANPA_KONVERSI (from df_nonagg)
                # For CSV Keseluruhan this will be vertical table including aggregates;
                # For CSV Grup Iklan this will be the comma/single-cell representation if you prefer — current behavior keeps it vertical.
                tanpa_konversi_df.to_excel(
                    writer,
                    sheet_name=">10K_TANPA_KONVERSI",
                    index=False
                )

                # color >10k text red
                ws_tc = writer.book[">10K_TANPA_KONVERSI"]
                for r in range(2, ws_tc.max_row + 1):
                    for c in range(1, ws_tc.max_column + 1):
                        cell = ws_tc.cell(row=r, column=c)
                        cell.font = Font(color="FF0000")

            buffer.seek(0)

        st.success("Excel laporan siap di-download 👇")

        st.download_button(
            "Download Excel Laporan",
            buffer,
            filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
