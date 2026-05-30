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
        },
        "kolom_alamat":  "Alamat Pengiriman Pesanan Detail Pengiriman Pesan",
        "kolom_hp":      "Handphone",
        "kolom_numeric": ["Kuantitas", "Total Harga"],
        "kolom_tanggal": ["Tanggal"],
    },

    "Monitoring Pelanggan | Sales Offline (PerBulan)": {
        "output_file": "CLEANING_MONITORING_PELANGGAN_PerBulan.xlsx",
        "sheet_name": "Monitoring Pelanggan",
        "use_master_kota": True,
        "sku_filter": None,
        "tanggal_format": None,
        "clean_hp": True,
        "replace_nan_string": True,
        "kolom_map": {
            "Divisi":                   "Divisi",
