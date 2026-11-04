import io
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Report Tangkap Gurame", layout="wide")

BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]

KOLOM = ["REGIONAL", "KODE_CABANG", "NAMA_CABANG", "NAMA_PRODUK", "NILAI_PRODUK",
         "TANGGAL_BAYAR", "ID_SELLER", "NAMA_SELLER", "SATELIT", "KODE SATELIT",
         "JOB SELLER", "KATEGORI CABANG"]

JOB_LIST = ("CREDIT RETENTION OFFICER", "CREDIT MARKETING OFFICER",
            "CREDIT CHANNELING OFFICER", "RELATIONSHIP OFFICER",
            "SALES OFFICER", "SPG")


def proses(file) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name="MIKRO")
    df = df.iloc[2:, :].reset_index(drop=True)
    df.columns = df.iloc[0]
    df = df.iloc[1:, :].reset_index(drop=True)

    df = df[KOLOM]
    df = df[df["NAMA_PRODUK"].isin(["Amandira Mikro", "Asira Mikro"])].copy()
    df["NILAI_PRODUK"] = pd.to_numeric(df["NILAI_PRODUK"], errors="coerce")
    df = df[df["NILAI_PRODUK"] > 75000]
    df = df[df["JOB SELLER"].fillna("").str.strip().str.startswith(JOB_LIST)]
    df = df.reset_index(drop=True)

    df["KODE_CABANG"] = df["KODE_CABANG"].fillna("").astype(str).str.zfill(4)
    df["KODE SATELIT"] = df["KODE SATELIT"].fillna("").astype(str).str.zfill(4)
    df["ID_SELLER"] = df["ID_SELLER"].fillna("").astype(str).str.zfill(8)
    return df


def buat_pivot(df: pd.DataFrame) -> pd.DataFrame:
    pivot = pd.pivot_table(
        df,
        index=["REGIONAL", "KODE_CABANG", "NAMA_CABANG", "SATELIT",
               "ID_SELLER", "NAMA_SELLER", "JOB SELLER"],
        values="NILAI_PRODUK",
        aggfunc={"NILAI_PRODUK": ["count", "sum"]},
    ).reset_index()
    pivot.columns = ["REGIONAL", "KODE_CABANG", "NAMA_CABANG", "SATELIT",
                     "ID_SELLER", "NAMA_SELLER", "JOB SELLER",
                     "TOTAL_POLIS", "TOTAL_PREMI"]
    pivot["REWARD"] = pivot["TOTAL_POLIS"] * 10000
    return pivot


def filter_tabel(df: pd.DataFrame, key: str) -> pd.DataFrame:
    with st.expander("🔎 Filter"):
        c1, c2 = st.columns(2)
        regional = c1.multiselect("REGIONAL", sorted(df["REGIONAL"].dropna().astype(str).unique()), key=key + "r")
        cabang = c2.multiselect("NAMA_CABANG", sorted(df["NAMA_CABANG"].dropna().astype(str).unique()), key=key + "c")
        cari = st.text_input("Cari ID / Nama Seller", key=key + "s").strip()

    if regional:
        df = df[df["REGIONAL"].astype(str).isin(regional)]
    if cabang:
        df = df[df["NAMA_CABANG"].astype(str).isin(cabang)]
    if cari:
        cocok = (df["ID_SELLER"].astype(str).str.contains(cari, case=False, na=False, regex=False)
                 | df["NAMA_SELLER"].astype(str).str.contains(cari, case=False, na=False, regex=False))
        df = df[cocok]
    return df


st.title("Report Tangkap Gurame 01")

files = st.file_uploader("Upload file Excel (.xlsx, sheet MIKRO)", type=["xlsx"],
                         accept_multiple_files=True)
if not files:
    st.stop()

df = pd.concat([proses(f) for f in files], ignore_index=True)
pivot = buat_pivot(df)

st.subheader("1. Data Amandira & Asira")
df_f = filter_tabel(df, "df")
st.caption(f"{len(df_f):,} baris")
st.dataframe(df_f.head(5000), use_container_width=True, hide_index=True)

st.subheader("2. Pivot Final")
pivot_f = filter_tabel(pivot, "pv")
st.caption(f"{len(pivot_f):,} seller")
st.dataframe(pivot_f.head(5000), use_container_width=True, hide_index=True)

kemarin = datetime.now() - timedelta(days=1)
nama = f"Report Tangkap Gurame 01 - {kemarin.day:02d} {BULAN[kemarin.month - 1]} {kemarin.year}.xlsx"

buf = io.BytesIO()
pivot.to_excel(buf, index=False)
st.download_button("⬇️ Download pivot final", buf.getvalue(), file_name=nama,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
