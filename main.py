import io
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Report Tangkap Gurame", layout="wide")

BULAN = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
    7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}

KOLOM_DIPAKAI = [
    "REGIONAL", "KODE_CABANG", "NAMA_CABANG", "NAMA_PRODUK", "NILAI_PRODUK",
    "TANGGAL_BAYAR", "ID_SELLER", "NAMA_SELLER", "SATELIT", "KODE SATELIT",
    "JOB SELLER", "KATEGORI CABANG",
]
PRODUK = ["Amandira Mikro", "Asira Mikro"]
JOB_LIST = (
    "CREDIT RETENTION OFFICER",
    "CREDIT MARKETING OFFICER",
    "CREDIT CHANNELING OFFICER",
    "RELATIONSHIP OFFICER",
    "SALES OFFICER",
    "SPG",
)
REWARD_PER_POLIS = 10000
MAX_BARIS_TAMPIL = 5000
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ---------------------------------------------------------------- Proses data
@st.cache_data(show_spinner=False)
def proses_file(isi: bytes) -> pd.DataFrame:
    """Baca satu file Excel (sheet MIKRO) lalu bersihkan sesuai alur di notebook."""
    df = pd.read_excel(io.BytesIO(isi), sheet_name="MIKRO", engine="calamine")

    df = df.iloc[2:, :].reset_index(drop=True)
    df.columns = df.iloc[0]
    df = df.iloc[1:, :].reset_index(drop=True)

    df = df[KOLOM_DIPAKAI]
    df = df[df["NAMA_PRODUK"].isin(PRODUK)].copy()

    df["NILAI_PRODUK"] = pd.to_numeric(df["NILAI_PRODUK"], errors="coerce")
    df = df[df["NILAI_PRODUK"] > 75000]

    df = df[df["JOB SELLER"].fillna("").str.strip().str.startswith(JOB_LIST)].reset_index(drop=True)

    df["KODE_CABANG"] = df["KODE_CABANG"].fillna("").astype(str).str.zfill(4)
    df["KODE SATELIT"] = df["KODE SATELIT"].fillna("").astype(str).str.zfill(4)
    df["ID_SELLER"] = df["ID_SELLER"].fillna("").astype(str).str.zfill(8)
    return df


@st.cache_data(show_spinner=False)
def buat_pivot(df: pd.DataFrame) -> pd.DataFrame:
    pivot = pd.pivot_table(
        df,
        index=["REGIONAL", "KODE_CABANG", "NAMA_CABANG", "SATELIT",
               "ID_SELLER", "NAMA_SELLER", "JOB SELLER"],
        values="NILAI_PRODUK",
        aggfunc={"NILAI_PRODUK": ["count", "sum"]},
    ).reset_index()

    pivot.columns = [
        "REGIONAL", "KODE_CABANG", "NAMA_CABANG", "SATELIT", "ID_SELLER",
        "NAMA_SELLER", "JOB SELLER", "TOTAL_POLIS", "TOTAL_PREMI",
    ]
    pivot["KODE_CABANG"] = pivot["KODE_CABANG"].fillna("").astype(str).str.zfill(4)
    pivot["ID_SELLER"] = pivot["ID_SELLER"].fillna("").astype(str).str.zfill(8)
    pivot["REWARD"] = pivot["TOTAL_POLIS"] * REWARD_PER_POLIS
    return pivot


@st.cache_data(show_spinner=False)
def ke_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


# ---------------------------------------------------------------- Filter
def panel_filter(df: pd.DataFrame, kolom_multi: list, kolom_cari: list, key: str,
                 kolom_tanggal: str | None = None, kolom_angka: str | None = None) -> pd.DataFrame:
    """Panel filter di dalam expander. Mengembalikan df yang sudah terfilter."""
    mask = pd.Series(True, index=df.index)

    with st.expander("🔎 Filter tabel", expanded=False):
        cols = st.columns(3)
        for i, kolom in enumerate(kolom_multi):
            opsi = sorted(df[kolom].dropna().astype(str).unique())
            pilih = cols[i % 3].multiselect(kolom, opsi, key=f"{key}_{kolom}")
            if pilih:
                mask &= df[kolom].astype(str).isin(pilih)

        cari = st.text_input(
            f"Cari teks di: {', '.join(kolom_cari)}", key=f"{key}_cari"
        ).strip()
        if cari:
            m = pd.Series(False, index=df.index)
            for kolom in kolom_cari:
                m |= df[kolom].astype(str).str.contains(cari, case=False, na=False, regex=False)
            mask &= m

        if kolom_tanggal:
            tgl = pd.to_datetime(df[kolom_tanggal], errors="coerce")
            if tgl.notna().any():
                tmin, tmax = tgl.min().date(), tgl.max().date()
                rentang = st.date_input(
                    f"Rentang {kolom_tanggal}", value=(tmin, tmax),
                    min_value=tmin, max_value=tmax, key=f"{key}_tgl",
                )
                if isinstance(rentang, tuple) and len(rentang) == 2:
                    mask &= tgl.between(
                        pd.Timestamp(rentang[0]),
                        pd.Timestamp(rentang[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1),
                    )

        if kolom_angka:
            minimal = st.number_input(
                f"Minimal {kolom_angka}", min_value=0, value=0, step=1, key=f"{key}_min"
            )
            if minimal > 0:
                mask &= df[kolom_angka] >= minimal

    return df[mask]


def tampilkan_tabel(df: pd.DataFrame):
    st.dataframe(df.head(MAX_BARIS_TAMPIL), use_container_width=True, hide_index=True)
    if len(df) > MAX_BARIS_TAMPIL:
        st.caption(f"Menampilkan {MAX_BARIS_TAMPIL:,} dari {len(df):,} baris.")


# ---------------------------------------------------------------- UI
st.title("📊 Report Tangkap Gurame")

files = st.file_uploader(
    "Upload file Excel (.xlsx, sheet MIKRO), boleh lebih dari satu",
    type=["xlsx"],
    accept_multiple_files=True,
)

if not files:
    st.info("Silakan upload file terlebih dahulu.")
    st.stop()

try:
    with st.spinner("Membaca dan memproses file..."):
        df = pd.concat([proses_file(f.getvalue()) for f in files], ignore_index=True)
        pivot = buat_pivot(df)
except Exception as e:
    st.error(f"Gagal memproses file: {e}")
    st.stop()

st.success(f"{len(files)} file berhasil dibaca.")

# --- Tabel 1: data Amandira & Asira
st.header("1️⃣ Data Amandira Mikro & Asira Mikro")
df_f = panel_filter(
    df,
    kolom_multi=["REGIONAL", "NAMA_CABANG", "NAMA_PRODUK", "SATELIT", "JOB SELLER", "KATEGORI CABANG"],
    kolom_cari=["ID_SELLER", "NAMA_SELLER"],
    key="df",
    kolom_tanggal="TANGGAL_BAYAR",
)
c1, c2 = st.columns(2)
c1.metric("Jumlah baris", f"{len(df_f):,}")
c2.metric("Total NILAI_PRODUK", f"{df_f['NILAI_PRODUK'].sum():,.0f}")
tampilkan_tabel(df_f)

st.divider()

# --- Tabel 2: pivot final
st.header("2️⃣ Pivot Final (untuk di-download)")
pivot_f = panel_filter(
    pivot,
    kolom_multi=["REGIONAL", "NAMA_CABANG", "SATELIT", "JOB SELLER"],
    kolom_cari=["ID_SELLER", "NAMA_SELLER"],
    key="pivot",
    kolom_angka="TOTAL_POLIS",
)
m1, m2, m3, m4 = st.columns(4)
m1.metric("Jumlah seller", f"{len(pivot_f):,}")
m2.metric("Total polis", f"{pivot_f['TOTAL_POLIS'].sum():,}")
m3.metric("Total premi", f"{pivot_f['TOTAL_PREMI'].sum():,.0f}")
m4.metric("Total reward", f"{pivot_f['REWARD'].sum():,.0f}")
tampilkan_tabel(pivot_f)

# --- Download
kemarin = datetime.now(ZoneInfo("Asia/Jakarta")).date() - timedelta(days=1)
nama_file = f"Report Tangkap Gurame 01 - {kemarin.day:02d} {BULAN[kemarin.month]} {kemarin.year}.xlsx"

d1, d2 = st.columns(2)

# Pivot lengkap: file dibuat sekali lalu di-cache
d1.download_button(
    "⬇️ Download pivot lengkap (.xlsx)", ke_excel(pivot),
    file_name=nama_file, mime=MIME_XLSX, use_container_width=True,
)

# Pivot hasil filter: file hanya dibuat saat tombol "Siapkan" ditekan
sig = int(pd.util.hash_pandas_object(pivot_f, index=False).sum())
if d2.button("Siapkan file hasil filter", use_container_width=True):
    st.session_state["xlsx_filter"] = (sig, ke_excel(pivot_f))

tersimpan = st.session_state.get("xlsx_filter")
if tersimpan and tersimpan[0] == sig:
    d2.download_button(
        "⬇️ Download pivot hasil filter (.xlsx)", tersimpan[1],
        file_name=nama_file.replace(".xlsx", " (filter).xlsx"), mime=MIME_XLSX,
        use_container_width=True,
    )
elif tersimpan:
    d2.caption("Filter berubah, tekan \"Siapkan file hasil filter\" lagi.")
