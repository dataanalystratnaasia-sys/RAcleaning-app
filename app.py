import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO

st.set_page_config(page_title="Cleaning Data Sales", page_icon="🧹", layout="wide")

# ============================================================
# KONFIGURASI TEMPLATE
# ============================================================
TEMPLATES = {

    "Data Pelanggan 2026 | Sales Online": {
        "output_file": "CLEANING_SALES_ONLINE.xlsx",
        "sheet_name": "Sales Online",
        "use_master_kota": True,
        "sku_filter": None,
        "tanggal_format": None,
        "clean_hp": True,
        "replace_nan_string": True,
        "sales_filter":  ["ARUM", "LUTHFIAH WARDAH"],
        "sales_exclude": None,
        "kolom_map": {
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
            "Alamat Pengiriman Pesanan Detail Pengiriman Pesan": "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        },
        "kolom_alamat":     "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        "kolom_kota_after": "Nama Kategori Pelanggan",
        "kota_upper":       True,
        "kolom_hp":         "Handphone",
        "kolom_numeric":    ["Kuantitas", "Total Harga"],
        "kolom_tanggal":    ["Tanggal"],
    },
    
    "Monitoring Pelanggan | Sales Offline (PerBulan)": {
        "output_file": "CLEANING_MONITORING_PELANGGAN_PerBulan.xlsx",
        "sheet_name": "Monitoring Pelanggan",
        "use_master_kota": True,
        "sku_filter": None,
        "tanggal_format": None,
        "clean_hp": True,
        "replace_nan_string": True,
        "sales_filter":  None,
        "sales_exclude": ["ARUM", "LUTHFIAH WARDAH"],
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
            "Alamat Pengiriman Pesanan Detail Pengiriman Pesan": "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        },
        "kolom_alamat":     "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        "kolom_kota_after": "Nama Kategori Pelanggan",
        "kota_upper":       True,
        "kolom_hp":         "Handphone",
        "kolom_numeric":    ["Kuantitas", "Total Harga"],
        "kolom_tanggal":    ["Tanggal"],
    },
    
    "Monitoring Pelanggan | Sales Offline (Data2026)": {
        "output_file": "CLEANING_MONITORING_PELANGGAN_Data2026.xlsx",
        "sheet_name": "Monitoring Pelanggan 2026",
        "use_master_kota": True,
        "sku_filter": None,
        "tanggal_format": None,
        "clean_hp": True,
        "replace_nan_string": True,
        "sales_filter":  None,
        "sales_exclude": ["ARUM", "LUTHFIAH WARDAH"],
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
            "Alamat Pengiriman Pesanan Detail Pengiriman Pesan": "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        },
        "kolom_alamat":     "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        "kolom_kota_after": "Nama Kategori Pelanggan",
        "kota_upper":       True,
        "kolom_hp":         "Handphone",
        "kolom_numeric":    ["Kuantitas", "Total Harga"],
        "kolom_tanggal":    ["Tanggal"],
    },
    
    "Data Pelanggan 2026 | Sales Offline": {
        "output_file": "CLEANING_DATA_PELANGGAN_2026_OFFLINE.xlsx",
        "sheet_name": "Data Pelanggan 2026",
        "use_master_kota": True,
        "sku_filter": None,
        "tanggal_format": None,
        "clean_hp": False,
        "replace_nan_string": True,
        "sales_filter":  None,
        "sales_exclude": ["ARUM", "LUTHFIAH WARDAH"],
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
            # Kota disisipkan setelah Nama Kategori Pelanggan
            "Nama Merek Barang":        "Nama Merek Barang Barang & Jasa",
            "Nama Kategori Barang":     "Nama Kategori Barang Barang & Jasa",
            "Alamat Pengiriman Pesanan Detail Pengiriman Pesan": "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        },
        "kolom_alamat":     "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        "kolom_kota_after": "Nama Kategori Pelanggan",
        "kota_upper":       True,
        "kolom_hp":         None,
        "kolom_numeric":    ["Kuantitas", "Total Harga"],
        "kolom_tanggal":    ["Tanggal"],
    },
    
    "Master Dashboard All Sales": {
        "output_file": "MASTER_DASHBOARD_ALL_SALES.xlsx",
        "sheet_name": "Master Dashboard",
        "use_master_kota": True,
        "sku_filter": None,
        "tanggal_format": None,
        "clean_hp": False,
        "replace_nan_string": False,
        "sales_filter":  None,
        "sales_exclude": None,
        "kolom_map": {
            "Tanggal":                                       "Tanggal",
            "Nomor #":                                       "Nomor #",
            "Nama Barang":                                   "Nama Barang",
            "Kode #":                                        "Kode #",
            "Divisi":                                        "Divisi",
            "Brand":                                         "Nama Merek Barang Barang & Jasa",
            "Nama Kategori Barang Barang & Jasa":            "Nama Kategori Barang Barang & Jasa",
            "QTY":                                           "Kuantitas",
            "Total Harga":                                   "Total Harga",
            "Pelanggan":                                     "Pelanggan",
            "Nama Kategori Pelanggan Pesanan Penjualan":     "Nama Kategori Pelanggan Pesanan Penjualan",
            "Alamat Pengiriman Pesanan Detail Pengiriman Pesan": "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        },
        "kolom_alamat":     "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        "kolom_kota_after": "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",  # Kota di akhir
        "kolom_hp":         None,
        "kolom_numeric":    ["QTY", "Total Harga"],
        "kolom_tanggal":    ["Tanggal"],
        "kota_upper":       True,
        "kota_strip_prefix": False,
    },

    "Demand & Sales Velocity": {
        "output_file": "DEMAND_SALES_VELOCITY.xlsx",
        "sheet_name": "Demand & Sales Velocity",
        "use_master_kota": True,
        "sku_filter": None,
        "tanggal_format": None,
        "clean_hp": False,
        "replace_nan_string": True,
        "sales_filter":  None,
        "sales_exclude": None,
        "kolom_map": {
            "Divisi":                   "Divisi",
            "Pelanggan":                "Pelanggan",
            "Tanggal":                  "Tanggal",
            "SKU":                      "Kode #",
            "Nama Barang":              "Nama Barang",
            "Kuantitas":                "Kuantitas",
            "Total Harga":              "Total Harga",
            "Nama Tenaga Penjual":      "Nama Tenaga Penjual",
            "Nama Kategori Pelanggan":  "Nama Kategori Pelanggan Pesanan Penjualan",
            # Kota disisipkan setelah Nama Kategori Pelanggan
            "Alamat":                   "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
            "Nama Merek Barang":        "Nama Merek Barang Barang & Jasa",
            "Nama Kategori Barang":     "Nama Kategori Barang Barang & Jasa",
        },
        "kolom_alamat":      "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        "kolom_kota_after":  "Nama Kategori Pelanggan",
        "kota_upper":        True,
        "kota_strip_prefix": True,   # buang KOTA dan KAB.
        "kolom_hp":          None,
        "kolom_numeric":     ["Kuantitas", "Total Harga"],
        "kolom_tanggal":     ["Tanggal"],
    },

    "Filter SKU Diskon (Bu Dhany)": {
        "type": "merge_sku",
        "output_file": "hasil_riwayat_filtered_BuDhany.xlsx",
        "sheet_name": "Riwayat Filtered",
        "merge_key": "SKU",
        "sku_cols": ["SKU", "SEGMENT"],
    },

    "Laporan Sell Out Bosch (Pak Sonny)": {
        "type": "bosch_sellout",
        "output_file": "HASIL_CLEANING_DATA.xlsx",
        "sheet_name": "DATA",
        "use_master_kota": False,
    },

    "Penjualan Bosch PTRA (Bu Hasna)": {
        "type": "bosch_ptra",
        "output_file": "PENJUALAN_BOSCH_PTRA.xlsx",
        "use_master_kota": False,
    },

    "Kuadran Marketplace": {
        "type": "bosch_marketplace",
        "output_file": "DATA_Marketplace.xlsx",
        "sheet_name": "DATA",
        "use_master_kota": False,
    },
}

MARKETPLACE_KEYWORDS = ['shopee', 'tiktok', 'lazada', 'blibli', 'tokopedia']

def is_marketplace(nilai):
    if pd.isna(nilai):
        return False
    return any(kw in str(nilai).lower() for kw in MARKETPLACE_KEYWORDS)

def filter_toko(nilai):
    if pd.isna(nilai):
        return np.nan
    nilai_str = str(nilai).strip()
    if any(kw in nilai_str.lower() for kw in MARKETPLACE_KEYWORDS):
        return nilai_str
    return np.nan

def map_bu(divisi_raw, nama_barang):
    if divisi_raw is None or (isinstance(divisi_raw, float) and np.isnan(divisi_raw)):
        return nama_barang
    divisi_str = str(divisi_raw).strip()
    try:
        divisi_str = str(int(float(divisi_str)))
    except (ValueError, TypeError):
        pass
    if divisi_str == '71':
        return 'BE GT'
    elif divisi_str == '82':
        return 'MT'
    elif divisi_str in ('72', '73'):
        return 'HG OG'
    elif divisi_str == '74':
        return 'ACC'
    elif divisi_str == '75':
        return 'SP'
    else:
        return nama_barang

def build_ptra(df):
    df_clean = pd.DataFrame()
    df_clean['Tanggal']             = pd.to_datetime(df['Tanggal'], errors='coerce') if 'Tanggal' in df.columns else np.nan
    df_clean['Nomor #']             = df['Nomor #'] if 'Nomor #' in df.columns else np.nan
    df_clean['Pelanggan']           = df['Pelanggan'] if 'Pelanggan' in df.columns else np.nan
    df_clean['Kode #']              = df['Kode #'] if 'Kode #' in df.columns else np.nan
    df_clean['Nama Barang']         = df['Nama Barang'] if 'Nama Barang' in df.columns else np.nan
    df_clean['Kuantitas']           = pd.to_numeric(df['Kuantitas'], errors='coerce') if 'Kuantitas' in df.columns else np.nan
    df_clean['@Harga']              = pd.to_numeric(df['@Harga'], errors='coerce') if '@Harga' in df.columns else np.nan
    df_clean['Total Harga']         = pd.to_numeric(df['Total Harga'], errors='coerce') if 'Total Harga' in df.columns else np.nan
    df_clean['Nama Tenaga Penjual'] = df['Nama Tenaga Penjual'] if 'Nama Tenaga Penjual' in df.columns else np.nan
    df_clean['Kategori']            = df.apply(
        lambda row: map_bu(
            row['Divisi'] if 'Divisi' in df.columns else None,
            row['Nama Barang'] if 'Nama Barang' in df.columns else None
        ), axis=1
    )
    for col in df_clean.select_dtypes(include='object').columns:
        df_clean[col] = df_clean[col].str.strip()
    mask_mp  = df_clean['Pelanggan'].apply(is_marketplace)
    df_mp    = df_clean[mask_mp].reset_index(drop=True)
    df_sales = df_clean[~mask_mp].reset_index(drop=True)
    return df_clean, df_mp, df_sales

def build_bosch_marketplace(df):
    marketplace_keywords = ['shopee', 'tiktok', 'tik tok', 'lazada', 'blibli', 'tokopedia']
    pattern = '|'.join(marketplace_keywords)
    if 'Pelanggan' in df.columns:
        df = df[df['Pelanggan'].astype(str).str.contains(pattern, case=False, na=False)].copy()
    df_clean = pd.DataFrame()
    df_clean['Tanggal']     = pd.to_datetime(df['Tanggal'], errors='coerce') if 'Tanggal' in df.columns else np.nan
    df_clean['Nomor #']     = df['Nomor #'] if 'Nomor #' in df.columns else np.nan
    df_clean['Pelanggan']   = df['Pelanggan'] if 'Pelanggan' in df.columns else np.nan
    df_clean['Kode #']      = df['Kode #'] if 'Kode #' in df.columns else np.nan
    df_clean['Nama Barang'] = df['Nama Barang'] if 'Nama Barang' in df.columns else np.nan
    df_clean['QTY']         = pd.to_numeric(df['Kuantitas'], errors='coerce') if 'Kuantitas' in df.columns else np.nan
    df_clean['@Harga']      = pd.to_numeric(df['@Harga'], errors='coerce') if '@Harga' in df.columns else np.nan
    df_clean['Total Harga'] = pd.to_numeric(df['Total Harga'], errors='coerce') if 'Total Harga' in df.columns else np.nan
    df_clean['Laba']        = pd.to_numeric(df['Laba'], errors='coerce') if 'Laba' in df.columns else np.nan
    df_clean['Gross Profit/Item'] = df_clean['Laba'] / df_clean['QTY'].replace(0, np.nan)
    for col in df_clean.select_dtypes(include='object').columns:
        df_clean[col] = df_clean[col].astype(str).str.strip()
    mandatory_cols = [c for c in ['Tanggal', 'Nomor #', 'Kode #', 'Nama Barang'] if c in df_clean.columns]
    df_clean = df_clean.dropna(subset=mandatory_cols).reset_index(drop=True)
    return df_clean

def insert_kota_after(df, after_col, kota_series):
    """Sisipkan kolom Kota tepat setelah kolom after_col."""
    if after_col not in df.columns:
        df['Kota'] = kota_series
        return df
    cols = list(df.columns)
    idx  = cols.index(after_col) + 1
    df.insert(idx, 'Kota', kota_series)
    return df

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
# SIDEBAR
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
    elif cfg.get("type") == "bosch_sellout":
        st.write(f"📄 Output: `{cfg['output_file']}`")
        st.write(f"📋 Sheet: `{cfg['sheet_name']}`")
        st.write("🔧 Tipe: Sell Out Bosch")
        st.write("🗺️ Master Kota: ❌")
        st.write("🏪 Filter Marketplace: ✅")
        st.write("📦 Mapping BU: ✅")
    elif cfg.get("type") == "bosch_ptra":
        st.write(f"📄 Output: `{cfg['output_file']}`")
        st.write("📋 Sheet: `MP` + `SALES`")
        st.write("🔧 Tipe: Penjualan Bosch PTRA")
        st.write("🗺️ Master Kota: ❌")
        st.write("🏪 Pisah Marketplace: ✅")
        st.write("📦 Mapping Kategori: ✅")
    elif cfg.get("type") == "bosch_marketplace":
        st.write(f"📄 Output: `{cfg['output_file']}`")
        st.write(f"📋 Sheet: `{cfg['sheet_name']}`")
        st.write("🏪 Hanya Marketplace")
        st.write("🗺️ Master Kota: ❌")
        st.write("📈 Gross Profit/Item: ✅")
    else:
        st.write(f"📄 Output: `{cfg['output_file']}`")
        st.write(f"📋 Sheet: `{cfg['sheet_name']}`")
        st.write(f"🗺️ Master Kota: {'✅' if cfg['use_master_kota'] else '❌'}")
        st.write(f"🔍 Filter SKU: {'✅' if cfg['sku_filter'] else '❌'}")
        st.write(f"📱 Clean HP: {'✅' if cfg['clean_hp'] else '❌'}")
        st.write(f"🧹 Replace NaN string: {'✅' if cfg['replace_nan_string'] else '❌'}")
        if cfg.get("sales_filter"):
            st.write(f"👤 Filter Sales (hanya): `{'`, `'.join(cfg['sales_filter'])}`")
        if cfg.get("sales_exclude"):
            st.write(f"🚫 Exclude Sales: `{'`, `'.join(cfg['sales_exclude'])}` + kosong")

    st.divider()

# ============================================================
# MAIN AREA
# ============================================================
st.markdown(f'<div class="template-badge">📌 {selected}</div>', unsafe_allow_html=True)

# ============================================================
# CABANG: MERGE SKU
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
            sku_df.columns = sku_df.columns.str.strip()
            trx_df.columns = trx_df.columns.str.strip()

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
                result = trx_df.merge(sku_df[cfg["sku_cols"]], on=merge_key, how="inner")

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
                    df.to_excel(w, index=False, sheet_name=sheet, merge_cells=False)
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
# CABANG: LAPORAN SELL OUT BOSCH
# ============================================================
elif cfg.get("type") == "bosch_sellout":

    st.markdown('<div class="step-label">📂 Step 1 — Data Original Bosch</div>', unsafe_allow_html=True)
    file_ori = st.file_uploader("Upload Data Original (.xlsx)", type=["xlsx"], key="bosch_ori")

    if file_ori:
        try:
            df = pd.read_excel(file_ori, header=0)
            df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
            df.columns = df.columns.str.strip()

            with st.expander("👁️ Preview Data Original", expanded=False):
                st.dataframe(df.head(10), use_container_width=True)

            with st.expander("🔍 Debug — Cek Kolom Divisi", expanded=True):
                if 'Divisi' in df.columns:
                    sample = df['Divisi'].dropna().unique()[:10]
                    st.write("**Kolom Divisi ditemukan ✅**")
                    st.write("Nilai unik (tampilan):", list(sample))
                    st.write("Nilai unik (repr):", [repr(x) for x in sample])
                    st.write("Dtype:", df['Divisi'].dtype)
                else:
                    st.error("❌ Kolom 'Divisi' TIDAK ditemukan!")
                    st.write("Kolom tersedia:", df.columns.tolist())

            with st.spinner("⚙️ Sedang memproses Sell Out Bosch..."):
                df_clean = pd.DataFrame()
                df_clean['Invoice']                = np.nan
                df_clean['Toko']                   = df['Pelanggan'].apply(filter_toko) if 'Pelanggan' in df.columns else np.nan
                df_clean['Tgl Nota']               = pd.to_datetime(df['Tanggal'], errors='coerce') if 'Tanggal' in df.columns else np.nan
                df_clean['Bulan']                  = df_clean['Tgl Nota'].dt.month
                df_clean['Tahun']                  = df_clean['Tgl Nota'].dt.year
                df_clean['Qty']                    = pd.to_numeric(df['Kuantitas'], errors='coerce') if 'Kuantitas' in df.columns else np.nan
                df_clean['Harga Satuan (Include)'] = pd.to_numeric(df['@Harga'], errors='coerce') if '@Harga' in df.columns else np.nan
                df_clean['Nominal (Include)']      = pd.to_numeric(df['Total Harga'], errors='coerce') if 'Total Harga' in df.columns else np.nan
                df_clean['Part']                   = df['Kode #'] if 'Kode #' in df.columns else np.nan
                df_clean['Nama Barang']            = df['Nama Barang'] if 'Nama Barang' in df.columns else np.nan
                df_clean['BU']                     = df.apply(
                    lambda row: map_bu(
                        row['Divisi'] if 'Divisi' in df.columns else None,
                        row['Nama Barang'] if 'Nama Barang' in df.columns else None
                    ), axis=1
                )
                for col in df_clean.select_dtypes(include='object').columns:
                    df_clean[col] = df_clean[col].str.strip()

            st.markdown("---")
            st.markdown("### 📊 Hasil Cleaning")
            c1, c2, c3 = st.columns(3)
            c1.markdown(f'<div class="stat-card" style="border-left-color:#2563eb"><div class="label">Total Data</div><div class="value" style="color:#2563eb">{len(df_clean):,}</div></div>', unsafe_allow_html=True)
            c2.markdown(f'<div class="stat-card" style="border-left-color:#16a34a"><div class="label">Toko Terisi (Marketplace)</div><div class="value" style="color:#16a34a">{int(df_clean["Toko"].notna().sum()):,}</div></div>', unsafe_allow_html=True)
            c3.markdown(f'<div class="stat-card" style="border-left-color:#dc2626"><div class="label">Toko Kosong</div><div class="value" style="color:#dc2626">{int(df_clean["Toko"].isna().sum()):,}</div></div>', unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.dataframe(df_clean.head(20), use_container_width=True)

            def to_excel_single(df, sheet):
                buf = BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as w:
                    df.to_excel(w, index=False, sheet_name=sheet, merge_cells=False)
                return buf.getvalue()

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="step-label">📥 Step 2 — Download Hasil</div>', unsafe_allow_html=True)
            st.download_button(
                label=f"⬇️ Download {cfg['output_file']}",
                data=to_excel_single(df_clean, cfg["sheet_name"]),
                file_name=cfg["output_file"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"❌ Error saat memproses: {e}")
            st.exception(e)
    else:
        st.info("👆 Upload file data original Bosch untuk memulai proses cleaning.")

# ============================================================
# CABANG: PENJUALAN BOSCH PTRA
# ============================================================
elif cfg.get("type") == "bosch_ptra":

    st.markdown('<div class="step-label">📂 Step 1 — Data Original</div>', unsafe_allow_html=True)
    file_ori = st.file_uploader("Upload Data Original (.xlsx)", type=["xlsx"], key="ptra_ori")

    if file_ori:
        try:
            df = pd.read_excel(file_ori, header=0)
            df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
            df.columns = df.columns.str.strip()

            with st.expander("👁️ Preview Data Original", expanded=False):
                st.dataframe(df.head(10), use_container_width=True)

            with st.spinner("⚙️ Sedang memproses Penjualan Bosch PTRA..."):
                df_clean, df_mp, df_sales = build_ptra(df)

            st.markdown("---")
            st.markdown("### 📊 Hasil Cleaning")
            c1, c2, c3 = st.columns(3)
            c1.markdown(f'<div class="stat-card" style="border-left-color:#2563eb"><div class="label">Total Data</div><div class="value" style="color:#2563eb">{len(df_clean):,}</div></div>', unsafe_allow_html=True)
            c2.markdown(f'<div class="stat-card" style="border-left-color:#7c3aed"><div class="label">Sheet MP</div><div class="value" style="color:#7c3aed">{len(df_mp):,}</div></div>', unsafe_allow_html=True)
            c3.markdown(f'<div class="stat-card" style="border-left-color:#16a34a"><div class="label">Sheet SALES</div><div class="value" style="color:#16a34a">{len(df_sales):,}</div></div>', unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            tab1, tab2 = st.tabs(["🏪 Sheet MP (Marketplace)", "🧑‍💼 Sheet SALES"])
            with tab1:
                st.dataframe(df_mp.head(20), use_container_width=True)
            with tab2:
                st.dataframe(df_sales.head(20), use_container_width=True)

            def to_excel_ptra(df_mp, df_sales):
                buf = BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as w:
                    df_mp.to_excel(w, index=False, sheet_name='MP')
                    df_sales.to_excel(w, index=False, sheet_name='SALES', merge_cells=False)
                return buf.getvalue()

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="step-label">📥 Step 2 — Download Hasil</div>', unsafe_allow_html=True)
            st.download_button(
                label=f"⬇️ Download {cfg['output_file']} (2 sheet: MP + SALES)",
                data=to_excel_ptra(df_mp, df_sales),
                file_name=cfg["output_file"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"❌ Error saat memproses: {e}")
            st.exception(e)
    else:
        st.info("👆 Upload Data Original untuk memulai.")

# ============================================================
# CABANG: KUADRAN MARKETPLACE
# ============================================================
elif cfg.get("type") == "bosch_marketplace":

    st.markdown('<div class="step-label">📂 Step 1 — Data Original</div>', unsafe_allow_html=True)
    file_ori = st.file_uploader("Upload Data Original (.xlsx)", type=["xlsx"], key="bosch_marketplace")

    if file_ori:
        try:
            df = pd.read_excel(file_ori)
            df = df.loc[:, ~df.columns.astype(str).str.startswith('Unnamed')]
            df.columns = df.columns.str.strip()

            with st.expander("👁️ Preview Data Original", expanded=False):
                st.dataframe(df.head(10), use_container_width=True)

            with st.spinner("⚙️ Sedang memproses data marketplace..."):
                df_clean = build_bosch_marketplace(df)

            st.markdown("---")
            st.markdown("### 📊 Hasil Cleaning")
            c1, c2, c3, c4 = st.columns(4)
            c1.markdown(f'<div class="stat-card" style="border-left-color:#2563eb"><div class="label">Total Data</div><div class="value" style="color:#2563eb">{len(df_clean):,}</div></div>', unsafe_allow_html=True)
            c2.markdown(f'<div class="stat-card" style="border-left-color:#7c3aed"><div class="label">Nota Unik</div><div class="value" style="color:#7c3aed">{df_clean["Nomor #"].nunique():,}</div></div>', unsafe_allow_html=True)
            c3.markdown(f'<div class="stat-card" style="border-left-color:#16a34a"><div class="label">Pelanggan Unik</div><div class="value" style="color:#16a34a">{df_clean["Pelanggan"].nunique():,}</div></div>', unsafe_allow_html=True)
            c4.markdown(f'<div class="stat-card" style="border-left-color:#ea580c"><div class="label">SKU Unik</div><div class="value" style="color:#ea580c">{df_clean["Kode #"].nunique():,}</div></div>', unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.dataframe(df_clean.head(20), use_container_width=True)

            def to_excel_marketplace(df):
                buf = BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as w:
                    df.to_excel(w, index=False, sheet_name=cfg["sheet_name"], merge_cells=False)
                return buf.getvalue()

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="step-label">📥 Step 2 — Download Hasil</div>', unsafe_allow_html=True)
            st.download_button(
                label=f"⬇️ Download {cfg['output_file']}",
                data=to_excel_marketplace(df_clean),
                file_name=cfg["output_file"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"❌ Error saat memproses: {e}")
            st.exception(e)
    else:
        st.info("👆 Upload Data Original untuk memulai proses cleaning.")

# ============================================================
# CABANG: CLEANING BIASA
# ============================================================
else:

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
            df_master = df_master.sort_values(by='Keyword', key=lambda x: x.str.len(), ascending=False)
            kw_map = list(zip(df_master['Keyword'], df_master['Hasil']))
            st.success(f"✅ Master kota loaded — {len(kw_map)} keyword")
    else:
        file_master = True
        st.info("ℹ️ Template ini tidak menggunakan Master Kota.")

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

                # ── Kota: sisipkan di posisi yang tepat ───────────────
                if cfg["kolom_alamat"] and cfg["use_master_kota"]:
                    kota_series = (
                        df[cfg["kolom_alamat"]].apply(extract_kota)
                        if cfg["kolom_alamat"] in df.columns
                        else pd.Series(np.nan, index=df_clean.index)
                    )
                    after_col = cfg.get("kolom_kota_after")
                    if after_col and after_col in df_clean.columns:
                        df_clean = insert_kota_after(df_clean, after_col, kota_series)
                    else:
                        df_clean["Kota"] = kota_series

                # ── Kota kapital jika diminta ──────────────────────────
                if cfg.get("kota_upper") and "Kota" in df_clean.columns:
                    def clean_kota(val):
                        if pd.isna(val):
                            return val
                        s = str(val).upper().strip()
                        if cfg.get("kota_strip_prefix", True): # kapital aja
                            s = re.sub(r'^KOTA\s+', '', s)
                            s = re.sub(r'^KAB\.\s*', '', s)
                            s = re.sub(r'^KABUPATEN\s+', '', s)
                            s = re.sub(r'^ADMINISTRASI\s+', '', s)
                        return s
                    df_clean["Kota"] = df_clean["Kota"].apply(clean_kota)
                # ── End Kota ──────────────────────────────────────────

                for col in df_clean.select_dtypes(include='object').columns:
                    df_clean[col] = df_clean[col].str.strip()

                for col in cfg["kolom_tanggal"]:
                    if col in df_clean.columns:
                        df_clean[col] = pd.to_datetime(df_clean[col], format=cfg["tanggal_format"], errors='coerce')

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
                        mask = df_clean[sku_col].astype(str).str.contains(pattern_sku, flags=re.IGNORECASE, na=False)
                        df_clean = df_clean[mask].reset_index(drop=True)

                # ── Filter Sales (hanya tampilkan nama tertentu) ───────
                n_before_sales = len(df_clean)
                if cfg.get("sales_filter") and "Nama Tenaga Penjual" in df_clean.columns:
                    allowed = [s.upper() for s in cfg["sales_filter"]]
                    mask_sales = df_clean["Nama Tenaga Penjual"].astype(str).str.strip().str.upper().isin(allowed)
                    df_clean = df_clean[mask_sales].reset_index(drop=True)

                # ── Exclude Sales (buang nama tertentu + cell kosong) ──
                n_before_exclude = len(df_clean)
                if cfg.get("sales_exclude") and "Nama Tenaga Penjual" in df_clean.columns:
                    excluded  = [s.upper() for s in cfg["sales_exclude"]]
                    sales_col = df_clean["Nama Tenaga Penjual"].astype(str).str.strip()
                    mask_exclude = (
                        ~sales_col.str.upper().isin(excluded)
                        & (sales_col != '')
                        & (sales_col.str.upper() != 'NAN')
                        & df_clean["Nama Tenaga Penjual"].notna()
                    )
                    df_clean = df_clean[mask_exclude].reset_index(drop=True)
                # ── End Filter/Exclude Sales ───────────────────────────

            if missing:
                st.warning(f"⚠️ Kolom tidak ditemukan di file original (diisi NaN): `{'`, `'.join(missing)}`")

            st.markdown("---")
            st.markdown("### 📊 Hasil Cleaning")

            stats = [("Total Data", len(df_clean), "#2563eb")]
            if cfg.get("sales_filter"):
                stats.append(("Dihapus (Sales Filter)", n_before_sales - len(df_clean), "#dc2626"))
            if cfg.get("sales_exclude"):
                stats.append(("Dihapus (Sales Exclude)", n_before_exclude - len(df_clean), "#dc2626"))
            if cfg["sku_filter"]:
                stats.append(("Setelah Filter SKU", len(df_clean), "#7c3aed"))
                stats.append(("Dihapus (SKU)", n_before - len(df_clean), "#dc2626"))
            if cfg["use_master_kota"] and cfg["kolom_alamat"]:
                stats.append(("Kota Terisi", int(df_clean['Kota'].notna().sum()), "#16a34a"))
                stats.append(("Kota Kosong", int(df_clean['Kota'].isna().sum()), "#dc2626"))

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
                    df.to_excel(w, index=False, sheet_name=sheet, merge_cells=False)
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
