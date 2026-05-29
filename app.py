import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO

st.set_page_config(page_title="Cleaning Data Sales", page_icon="🧹", layout="wide")

# ============================================================
# KONFIGURASI TEMPLATE
# Tambah template baru? Cukup tambah satu blok dict di sini.
# ============================================================
TEMPLATES = {

    "Data Pelanggan 2026 | Sales Online": {
        "output_file": "CLEANING_SALES_ONLINE.xlsx",
        "sheet_name": "Sales Online",
        "use_master_kota": True,
        "sku_filter": None,
        "tanggal_format": None,
        "clean_hp": True,             # aktifkan cleaning nomor HP
        "replace_nan_string": True,   # ganti string 'nan' dan '' jadi NaN
        "kolom_map": {
            # nama kolom output     : nama kolom di file original
            "Tanggal":                  "Tanggal",
            "Pelanggan":                "Pelanggan",
            "Kode #":                   "Kode #",
            "Nama Barang":              "Nama Barang",
            "Kuantitas":                "Kuantitas",
            "Total Harga":              "Total Harga",
            "Divisi":                   "Divisi",
            "Nama Merek Barang":        "Nama Merek Barang Barang & Jasa",
            "Nama Kategori Barang":     "Nama Kategori Barang Barang & Jasa",
            "Nama Tenaga Penjual":      "Nama Tenaga Penjual",
            "Nama Kategori Pelanggan":  "Nama Kategori Pelanggan Pesanan Penjualan",
            "Handphone":                "Handphone Kontak Utama Pelanggan Pesanan Penjuala",
        },
        "kolom_alamat":  "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        "kolom_hp":      "Handphone",   # nama kolom OUTPUT yang perlu di-clean
        "kolom_numeric": ["Kuantitas", "Total Harga"],
        "kolom_tanggal": ["Tanggal"],
    },

    "Monitoring Pelanggan | Sales Offline (PerBulan)": {
        "output_file": "CLEANING_MONITORING_PELANGGAN.xlsx",
        "sheet_name": "Monitoring Pelanggan",
        "use_master_kota": True,
        "sku_filter": None,
        "tanggal_format": None,
        "clean_hp": True,
        "replace_nan_string": True,
        "kolom_map": {
            "Divisi":                   "Divisi",
            "Pelanggan":                "Pelanggan",
            "Tanggal":                  "Tanggal",
            "Kode #":                   "Kode #",
            "Nama Barang":              "Nama Barang",
            "Kuantitas":                "Kuantitas",
            "Total Harga":              "Total Harga",
            "Nama Tenaga Penjual":      "Nama Tenaga Penjual",
            "Nama Kategori Pelanggan":  "Nama Kategori Pelanggan Pesanan Penjualan",
            "Handphone":                "Handphone Kontak Utama Pelanggan Pesanan Penjuala",
            "Nama Merek Barang":        "Nama Merek Barang Barang & Jasa",
            "Nama Kategori Barang":     "Nama Kategori Barang Barang & Jasa",
        },
        "kolom_alamat":  "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        "kolom_hp":      "Handphone",
        "kolom_numeric": ["Kuantitas", "Total Harga"],
        "kolom_tanggal": ["Tanggal"],
    },

    "Cleaning Monitoring Pelanggan | Sales Offline (Data2026)": {
        "output_file": "CLEANING_MONITORING_PELANGGAN_DATA2026.xlsx",
        "sheet_name": "Monitoring Pelanggan 2026",
        "use_master_kota": True,
        "sku_filter": None,
        "tanggal_format": None,
        "clean_hp": True,
        "replace_nan_string": True,
        "kolom_map": {
            "Divisi":                                            "Divisi",
            "Pelanggan":                                         "Pelanggan",
            "Tanggal":                                           "Tanggal",
            "Kode #":                                            "Kode #",
            "Nama Barang":                                       "Nama Barang",
            "Kuantitas":                                         "Kuantitas",
            "Total Harga":                                       "Total Harga",
            "Nama Tenaga Penjual":                               "Nama Tenaga Penjual",
            "Nama Kategori Pelanggan":                           "Nama Kategori Pelanggan Pesanan Penjualan",
            "Handphone":                                         "Handphone Kontak Utama Pelanggan Pesanan Penjuala",
            "Alamat Pengiriman Pesanan":                         "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
            "Alamat Pengiriman Pesanan Detail Pengiriman Pesan": "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
            "Nama Merek Barang":                                 "Nama Merek Barang Barang & Jasa",
            "Nama Kategori Barang":                              "Nama Kategori Barang Barang & Jasa",
        },
        "kolom_alamat":  "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        "kolom_hp":      "Handphone",
        "kolom_numeric": ["Kuantitas", "Total Harga"],
        "kolom_tanggal": ["Tanggal"],
    },

    "Cleaning Data Pelanggan 2026 | Sales Offline": {
        "output_file": "CLEANING_DATA_PELANGGAN_2026_OFFLINE.xlsx",
        "sheet_name": "Data Pelanggan 2026",
        "use_master_kota": True,
        "sku_filter": None,
        "tanggal_format": None,
        "clean_hp": False,          # kolom Handphone tidak ada di template ini
        "replace_nan_string": True,
        "kolom_map": {
            "Divisi":                   "Divisi",
            "Pelanggan":                "Pelanggan",
            "Tanggal":                  "Tanggal",
            "Kode #":                   "Kode #",
            "Nama Barang":              "Nama Barang",
            "Kuantitas":                "Kuantitas",
            "Total Harga":              "Total Harga",
            "Nama Tenaga Penjual":      "Nama Tenaga Penjual",
            "Nama Kategori Pelanggan":  "Nama Kategori Pelanggan Pesanan Penjualan",
            "Nama Merek Barang":        "Nama Merek Barang Barang & Jasa",
            "Nama Kategori Barang":     "Nama Kategori Barang Barang & Jasa",
        },
        "kolom_alamat":  "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        "kolom_hp":      None,
        "kolom_numeric": ["Kuantitas", "Total Harga"],
        "kolom_tanggal": ["Tanggal"],
    },

    "Master Dashboard All Sales": {
        "output_file": "MASTER_DASHBOARD_ALL_SALES.xlsx",
        "sheet_name": "Master Dashboard",
        "use_master_kota": True,
        "sku_filter": None,
        "tanggal_format": None,
        "clean_hp": False,
        "replace_nan_string": False,
        "kolom_map": {
            "Tanggal":               "Tanggal",
            "Nomor #":               "Nomor #",
            "Kode #":                "Kode #",
            "Nama Barang":           "Nama Barang",
            "Divisi":                "Divisi",
            "Brand":                 "Nama Merek Barang Barang & Jasa",
            "Nama Kategori Barang":  "Nama Kategori Barang Barang & Jasa",
            "QTY":                   "Kuantitas",
            "Total Harga":           "Total Harga",
            "Sales":                 "Nama Tenaga Penjual",
            "Pelanggan":             "Pelanggan",
            "Kategori Pelanggan":    "Nama Kategori Pelanggan Pesanan Penjualan",
            "Handphone":             "Handphone Kontak Utama Pelanggan Pesanan Penjuala",
            "Alamat Pengiriman":     "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        },
        "kolom_alamat":  "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        "kolom_hp":      None,
        "kolom_numeric": ["QTY", "Total Harga"],
        "kolom_tanggal": ["Tanggal"],
    },

    "Filter SKU Diskon (Bu Dhany)": {
        "type": "merge_sku",          # tipe khusus: bukan cleaning biasa
        "output_file": "hasil_riwayat_filtered.xlsx",
        "sheet_name": "Riwayat Filtered",
        "merge_key": "SKU",           # kolom join di kedua file
        "sku_cols": ["SKU", "SEGMENT"], # kolom yang diambil dari SKU List
    },

    # ── Slot untuk template berikutnya ──
    # "Template 8 — ...": { ... },
}

# ============================================================
# STYLING
# ============================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; }
    .header-box {
        background: linear-gradient(135deg, #1e3a5f 0%, #2563eb 100%);
        padding: 1.8rem 2.5rem; border-radius: 16px;
        margin-bottom: 1.5rem; color: white;
    }
    .header-box h1 { font-size: 1.6rem; font-weight: 700; margin: 0; }
    .header-box p  { font-size: 0.9rem; opacity: 0.85; margin: 0.3rem 0 0; }
    .stat-card {
        background: white; border-radius: 12px;
        padding: 1rem 1.2rem; border-left: 4px solid #2563eb;
        box-shadow: 0 1px 6px rgba(0,0,0,0.07);
    }
    .stat-card .label { font-size: 0.75rem; color: #64748b; font-weight: 600;
                        text-transform: uppercase; letter-spacing: 0.05em; }
    .stat-card .value { font-size: 1.6rem; font-weight: 700; color: #1e3a5f; }
    .template-badge {
        display: inline-block; background: #eff6ff; color: #1d4ed8;
        border: 1px solid #bfdbfe; border-radius: 20px;
        padding: 4px 14px; font-size: 0.8rem; font-weight: 600; margin-bottom: 1rem;
    }
    .step-label {
        font-size: 0.75rem; font-weight: 700; letter-spacing: 0.08em;
        text-transform: uppercase; color: #2563eb; margin-bottom: 0.3rem;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# HEADER
# ============================================================
st.markdown("""
<div class="header-box">
    <h1>🧹 Cleaning Data Sales 2026</h1>
    <p>Pilih template sesuai jenis data, upload file, download hasil cleaning</p>
</div>
""", unsafe_allow_html=True)

# ============================================================
# SIDEBAR — PILIH TEMPLATE
# ============================================================
with st.sidebar:
    st.markdown("### ⚙️ Pengaturan")
    selected = st.selectbox("Pilih Template", list(TEMPLATES.keys()))
    cfg = TEMPLATES[selected]

    st.divider()
    st.markdown("**Info Template**")
    if cfg.get("type") == "merge_sku":
        st.write(f"📄 Output: `{cfg['output_file']}`")
        st.write(f"📋 Sheet: `{cfg['sheet_name']}`")
        st.write("🔀 Tipe: Merge SKU List + Riwayat")
        st.write(f"🔑 Join key: `{cfg['merge_key']}`")
    else:
        st.write(f"📄 Output: `{cfg['output_file']}`")
        st.write(f"📋 Sheet: `{cfg['sheet_name']}`")
        st.write(f"🗺️ Master Kota: {'✅' if cfg['use_master_kota'] else '❌'}")
        st.write(f"🔍 Filter SKU: {'✅' if cfg['sku_filter'] else '❌'}")
        st.write(f"📱 Clean HP: {'✅' if cfg['clean_hp'] else '❌'}")
        st.write(f"🧹 Replace NaN string: {'✅' if cfg['replace_nan_string'] else '❌'}")

    st.divider()

# ============================================================
# MAIN AREA
# ============================================================
st.markdown(f'<div class="template-badge">📌 {selected}</div>', unsafe_allow_html=True)

# ============================================================
# CABANG: MERGE SKU (Filter SKU Diskon)
# ============================================================
if cfg.get("type") == "merge_sku":

    st.markdown('<div class="step-label">📋 Step 1 — SKU List (Acuan)</div>', unsafe_allow_html=True)
    file_sku = st.file_uploader("Upload file SKU List (.xlsx)", type=["xlsx"], key="sku_list",
                                help="Harus punya kolom: SKU dan SEGMENT")

    st.markdown('<div class="step-label">📂 Step 2 — Riwayat Pembelian</div>', unsafe_allow_html=True)
    file_trx = st.file_uploader("Upload file Riwayat Pembelian (.xlsx)", type=["xlsx"], key="trx")

    if file_sku and file_trx:
        try:
            sku_df = pd.read_excel(file_sku)
            trx_df = pd.read_excel(file_trx)

            # Strip kolom
            sku_df.columns = sku_df.columns.str.strip()
            trx_df.columns = trx_df.columns.str.strip()

            # Validasi kolom SKU List
            for col in cfg["sku_cols"]:
                if col not in sku_df.columns:
                    st.error(f"❌ Kolom **'{col}'** tidak ditemukan di SKU List!")
                    st.stop()

            merge_key = cfg["merge_key"]
            if merge_key not in trx_df.columns:
                st.error(f"❌ Kolom **'{merge_key}'** tidak ditemukan di Riwayat Pembelian!")
                st.stop()

            with st.expander("👁️ Preview SKU List", expanded=False):
                st.dataframe(sku_df.head(10), use_container_width=True)
            with st.expander("👁️ Preview Riwayat Pembelian", expanded=False):
                st.dataframe(trx_df.head(10), use_container_width=True)

            with st.spinner("⚙️ Sedang merge data..."):
                sku_df[merge_key] = sku_df[merge_key].astype(str)
                trx_df[merge_key] = trx_df[merge_key].astype(str)

                result = trx_df.merge(
                    sku_df[cfg["sku_cols"]],
                    on=merge_key,
                    how="inner"
                )

            # Statistik
            st.markdown("---")
            st.markdown("### 📊 Hasil Merge")
            c1, c2, c3 = st.columns(3)
            c1.markdown(f'<div class="stat-card" style="border-left-color:#2563eb"><div class="label">Total Riwayat</div><div class="value" style="color:#2563eb">{len(trx_df):,}</div></div>', unsafe_allow_html=True)
            c2.markdown(f'<div class="stat-card" style="border-left-color:#16a34a"><div class="label">Data Matched</div><div class="value" style="color:#16a34a">{len(result):,}</div></div>', unsafe_allow_html=True)
            c3.markdown(f'<div class="stat-card" style="border-left-color:#dc2626"><div class="label">Tidak Match</div><div class="value" style="color:#dc2626">{len(trx_df) - len(result):,}</div></div>', unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.dataframe(result.head(20), use_container_width=True)

            def to_excel(df, sheet):
                buf = BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as w:
                    df.to_excel(w, index=False, sheet_name=sheet)
                return buf.getvalue()

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="step-label">📥 Step 3 — Download Hasil</div>', unsafe_allow_html=True)
            st.download_button(
                label=f"⬇️ Download {cfg['output_file']}",
                data=to_excel(result, cfg["sheet_name"]),
                file_name=cfg["output_file"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"❌ Error: {e}")
            st.exception(e)
    else:
        st.info("👆 Upload kedua file di atas untuk memulai proses merge.")

# ============================================================
# CABANG: CLEANING BIASA (semua template selain merge_sku)
# ============================================================
else:

    # ── Upload Master Kota (kondisional) ──
    kw_map = []
    if cfg["use_master_kota"]:
        st.markdown('<div class="step-label">📋 Step 1 — Master Kota</div>', unsafe_allow_html=True)
        file_master = st.file_uploader("Upload Master Kota (.xlsx)", type=["xlsx"], key="master")

        if file_master:
            df_master = pd.read_excel(file_master)
            for col in ['Keyword', 'Hasil']:
                if col not in df_master.columns:
                    st.error(f"❌ Kolom **'{col}'** tidak ditemukan di Master Kota!")
                    st.stop()
            df_master = df_master.dropna(subset=['Keyword', 'Hasil'])
            df_master['Keyword'] = df_master['Keyword'].astype(str).str.strip()
            df_master['Hasil']   = df_master['Hasil'].astype(str).str.strip()
            df_master = df_master.sort_values(
                by='Keyword', key=lambda x: x.str.len(), ascending=False
            )
            kw_map = list(zip(df_master['Keyword'], df_master['Hasil']))
            st.success(f"✅ Master kota loaded — {len(kw_map)} keyword")
    else:
        file_master = True
        st.info("ℹ️ Template ini tidak menggunakan Master Kota.")

    # ── Upload Data Original ──
    st.markdown('<div class="step-label">📂 Step 2 — Data Original</div>', unsafe_allow_html=True)
    file_ori = st.file_uploader("Upload Data Original (.xlsx)", type=["xlsx"], key="original")

    if file_master and file_ori:
        try:
            df = pd.read_excel(file_ori, header=0)
            df = df.loc[:, ~df.columns.str.startswith('Unnamed')]

            with st.expander("👁️ Preview Data Original", expanded=False):
                st.dataframe(df.head(10), use_container_width=True)

            def extract_kota(alamat):
                if pd.isna(alamat) or not kw_map:
                    return np.nan
                alamat = str(alamat)
                for keyword, hasil in kw_map:
                    pattern = r'\b' + re.escape(keyword) + r'\b'
                    if re.search(pattern, alamat, flags=re.IGNORECASE):
                        return hasil
                return np.nan

            with st.spinner("⚙️ Sedang memproses..."):

                df_clean = pd.DataFrame()
                missing = []
                for out_col, in_col in cfg["kolom_map"].items():
                    if in_col in df.columns:
                        df_clean[out_col] = df[in_col]
                    else:
                        df_clean[out_col] = np.nan
                        missing.append(in_col)

                if cfg["kolom_alamat"] and cfg["use_master_kota"]:
                    if cfg["kolom_alamat"] in df.columns:
                        df_clean["Kota"] = df[cfg["kolom_alamat"]].apply(extract_kota)
                    else:
                        df_clean["Kota"] = np.nan

                for col in df_clean.select_dtypes(include='object').columns:
                    df_clean[col] = df_clean[col].str.strip()

                for col in cfg["kolom_tanggal"]:
                    if col in df_clean.columns:
                        df_clean[col] = pd.to_datetime(
                            df_clean[col], format=cfg["tanggal_format"], errors='coerce'
                        )

                for col in cfg["kolom_numeric"]:
                    if col in df_clean.columns:
                        df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce')

                if cfg["clean_hp"] and cfg["kolom_hp"] and cfg["kolom_hp"] in df_clean.columns:
                    df_clean[cfg["kolom_hp"]] = df_clean[cfg["kolom_hp"]].apply(
                        lambda x: re.sub(r'[^0-9]', '', str(x)) if pd.notna(x) else np.nan
                    )

                if cfg["replace_nan_string"]:
                    df_clean = df_clean.replace('nan', np.nan)
                    df_clean = df_clean.replace('', np.nan)

                n_before = len(df_clean)
                if cfg["sku_filter"]:
                    sku_col      = cfg["sku_filter"]["kolom"]
                    sku_keywords = cfg["sku_filter"]["keyword_aktif"]
                    if sku_col in df_clean.columns:
                        pattern_sku = '|'.join(re.escape(k) for k in sku_keywords)
                        mask = df_clean[sku_col].astype(str).str.contains(
                            pattern_sku, flags=re.IGNORECASE, na=False
                        )
                        df_clean = df_clean[mask].reset_index(drop=True)

            if missing:
                st.warning(f"⚠️ Kolom tidak ditemukan di file original (diisi NaN): `{'`, `'.join(missing)}`")

            st.markdown("---")
            st.markdown("### 📊 Hasil Cleaning")

            stats = [("Total Data", len(df_clean), "#2563eb")]
            if cfg["sku_filter"]:
                stats.append(("Setelah Filter SKU", len(df_clean),            "#7c3aed"))
                stats.append(("Dihapus (SKU)",      n_before - len(df_clean), "#dc2626"))
            if cfg["use_master_kota"] and cfg["kolom_alamat"]:
                stats.append(("Kota Terisi", int(df_clean['Kota'].notna().sum()), "#16a34a"))
                stats.append(("Kota Kosong", int(df_clean['Kota'].isna().sum()),  "#dc2626"))

            cols = st.columns(len(stats))
            for col_st, (label, value, color) in zip(cols, stats):
                col_st.markdown(f"""
                <div class="stat-card" style="border-left-color:{color}">
                    <div class="label">{label}</div>
                    <div class="value" style="color:{color}">{value:,}</div>
                </div>""", unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.dataframe(df_clean.head(20), use_container_width=True)

            def to_excel(df, sheet):
                buf = BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as w:
                    df.to_excel(w, index=False, sheet_name=sheet)
                return buf.getvalue()

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="step-label">📥 Step 3 — Download Hasil</div>', unsafe_allow_html=True)
            st.download_button(
                label=f"⬇️ Download {cfg['output_file']}",
                data=to_excel(df_clean, cfg["sheet_name"]),
                file_name=cfg["output_file"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"❌ Error saat memproses: {e}")
            st.exception(e)

    else:
        st.info("👆 Upload file di atas untuk memulai proses cleaning.")