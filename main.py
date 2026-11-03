
import streamlit as st
import pandas as pd
import io
from datetime import datetime

# =====================================================
# KONFIGURASI HALAMAN
# =====================================================

st.set_page_config(
    page_title="Dashboard Amandira & Asira Mikro",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Dashboard Amandira Mikro & Asira Mikro")
st.caption("Dashboard monitoring polis, premi, dan reward seller")


# =====================================================
# KONSTANTA
# =====================================================

KOLOM = [
    'REGIONAL',
    'KODE_CABANG',
    'NAMA_CABANG',
    'NAMA_PRODUK',
    'NILAI_PRODUK',
    'TANGGAL_BAYAR',
    'ID_SELLER',
    'NAMA_SELLER',
    'SATELIT',
    'KODE SATELIT',
    'JOB SELLER',
    'KATEGORI CABANG'
]

JOB_LIST = (
    "CREDIT RETENTION OFFICER",
    "CREDIT MARKETING OFFICER",
    "CREDIT CHANNELING OFFICER",
    "RELATIONSHIP OFFICER",
    "SALES OFFICER",
    "SPG"
)


# =====================================================
# FUNGSI MEMBACA FILE
# =====================================================


def read_files(uploaded_files):

    semua_data = []

    for file in uploaded_files:

        nama_file = file.name.lower()

        if nama_file.endswith(('.xlsx', '.xls')):
            df_temp = pd.read_excel(
                file,
                sheet_name='MIKRO',
                header=2
            )

        elif nama_file.endswith('.csv'):
            df_temp = pd.read_csv(
                file,
                header=2
            )

        else:
            continue

        # Rapikan nama kolom
        df_temp.columns = (
            df_temp.columns
            .astype(str)
            .str.replace('\n', ' ', regex=False)
            .str.strip()
        )

        semua_data.append(df_temp)

    if semua_data:
        df = pd.concat(
            semua_data,
            ignore_index=True
        )

        return df

    return pd.DataFrame()



# =====================================================
# FUNGSI PENGOLAHAN DATA
# =====================================================

def proses_data(df):

    # Ambil kolom yang digunakan
    df = df[KOLOM].copy()

    # Filter produk
    df = df[
        df['NAMA_PRODUK'].isin([
            "Amandira Mikro",
            "Asira Mikro"
        ])
    ]

    # Pastikan nilai produk numerik
    df['NILAI_PRODUK'] = pd.to_numeric(
        df['NILAI_PRODUK'],
        errors='coerce'
    )

    # Filter nilai produk lebih dari 75.000
    df = df[df['NILAI_PRODUK'] > 75000]

    # Filter jabatan seller
    df = df[
        df['JOB SELLER']
        .fillna('')
        .astype(str)
        .str.strip()
        .str.startswith(JOB_LIST)
    ].reset_index(drop=True)

    # Normalisasi kode
    df['KODE_CABANG'] = (
        df['KODE_CABANG']
        .fillna('')
        .astype(str)
        .str.replace(r'\.0$', '', regex=True)
        .str.zfill(4)
    )

    df['KODE SATELIT'] = (
        df['KODE SATELIT']
        .fillna('')
        .astype(str)
        .str.replace(r'\.0$', '', regex=True)
        .str.zfill(4)
    )

    df['ID_SELLER'] = (
        df['ID_SELLER']
        .fillna('')
        .astype(str)
        .str.replace(r'\.0$', '', regex=True)
        .str.zfill(8)
    )

    # =================================================
    # PIVOT / REKAPITULASI
    # =================================================

    pivot = pd.pivot_table(
        df,
        index=[
            'REGIONAL',
            'KODE_CABANG',
            'NAMA_CABANG',
            'SATELIT',
            'ID_SELLER',
            'NAMA_SELLER',
            'JOB SELLER'
        ],
        values='NILAI_PRODUK',
        aggfunc=['count', 'sum']
    ).reset_index()

    pivot.columns = [
        'REGIONAL',
        'KODE_CABANG',
        'NAMA_CABANG',
        'SATELIT',
        'ID_SELLER',
        'NAMA_SELLER',
        'JOB SELLER',
        'TOTAL_POLIS',
        'TOTAL_PREMI'
    ]

    # Normalisasi kembali kode pada pivot
    pivot['KODE_CABANG'] = (
        pivot['KODE_CABANG']
        .fillna('')
        .astype(str)
        .str.replace(r'\.0$', '', regex=True)
        .str.zfill(4)
    )

    pivot['ID_SELLER'] = (
        pivot['ID_SELLER']
        .fillna('')
        .astype(str)
        .str.replace(r'\.0$', '', regex=True)
        .str.zfill(8)
    )

    # Perhitungan reward
    pivot['REWARD'] = pivot['TOTAL_POLIS'] * 10000

    return df, pivot


# =====================================================
# SIDEBAR - UPLOAD FILE
# =====================================================

st.sidebar.header("📁 Upload Data")

uploaded_files = st.sidebar.file_uploader(
    "Upload file Excel atau CSV",
    type=['xlsx', 'xls', 'csv'],
    accept_multiple_files=True
)


# =====================================================
# PROSES DASHBOARD
# =====================================================

if uploaded_files:

    try:

        with st.spinner("Sedang membaca dan mengolah data..."):

            df_raw = read_files(uploaded_files)

            if df_raw.empty:
                st.warning("Tidak ada data yang berhasil dibaca.")
                st.stop()

            df, pivot = proses_data(df_raw)

        st.sidebar.success(
            f"{len(uploaded_files)} file berhasil dibaca"
        )

        # =================================================
        # FILTER DASHBOARD
        # =================================================

        st.sidebar.header("🔎 Filter Dashboard")

        regional_list = sorted(
            df['REGIONAL'].dropna().unique().tolist()
        )

        regional_pilih = st.sidebar.multiselect(
            "Pilih Regional",
            options=regional_list,
            default=regional_list
        )

        df_filter = df[
            df['REGIONAL'].isin(regional_pilih)
        ]

        cabang_list = sorted(
            df_filter['NAMA_CABANG'].dropna().unique().tolist()
        )

        cabang_pilih = st.sidebar.multiselect(
            "Pilih Cabang",
            options=cabang_list,
            default=cabang_list
        )

        df_filter = df_filter[
            df_filter['NAMA_CABANG'].isin(cabang_pilih)
        ]

        produk_list = sorted(
            df_filter['NAMA_PRODUK'].dropna().unique().tolist()
        )

        produk_pilih = st.sidebar.multiselect(
            "Pilih Produk",
            options=produk_list,
            default=produk_list
        )

        df_filter = df_filter[
            df_filter['NAMA_PRODUK'].isin(produk_pilih)
        ]

        # Buat pivot berdasarkan hasil filter
        _, pivot_filter = proses_data(df_filter)

        # =================================================
        # KPI / METRIK
        # =================================================

        st.subheader("Ringkasan Kinerja")

        total_polis = len(df_filter)

        total_premi = df_filter['NILAI_PRODUK'].sum()

        total_reward = total_polis * 10000

        total_seller = df_filter['ID_SELLER'].nunique()

        col1, col2, col3, col4 = st.columns(4)

        col1.metric(
            "Total Polis",
            f"{total_polis:,}"
        )

        col2.metric(
            "Total Premi",
            f"Rp {total_premi:,.0f}"
        )

        col3.metric(
            "Total Reward",
            f"Rp {total_reward:,.0f}"
        )

        col4.metric(
            "Total Seller",
            f"{total_seller:,}"
        )

        st.divider()

        # =================================================
        # TAB DASHBOARD
        # =================================================

        tab1, tab2 = st.tabs([
            "📋 Data Detail",
            "📊 Rekapitulasi Seller"
        ])

        # =================================================
        # TAB 1 - DATA DETAIL
        # =================================================

        with tab1:

            st.subheader("Data Detail Polis")

            st.write(
                f"Jumlah data: **{len(df_filter):,} baris**"
            )

            st.dataframe(
                df_filter,
                use_container_width=True,
                hide_index=True
            )

            # Download data detail
            output_detail = io.BytesIO()

            with pd.ExcelWriter(
                output_detail,
                engine='openpyxl'
            ) as writer:

                df_filter.to_excel(
                    writer,
                    index=False,
                    sheet_name='Data Detail'
                )

            st.download_button(
                label="⬇️ Download Data Detail Excel",
                data=output_detail.getvalue(),
                file_name=f"Data_Detail_Mikro_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # =================================================
        # TAB 2 - REKAP SELLER
        # =================================================

        with tab2:

            st.subheader("Rekapitulasi Seller")

            st.write(
                f"Jumlah seller: **{len(pivot_filter):,} baris**"
            )

            st.dataframe(
                pivot_filter,
                use_container_width=True,
                hide_index=True
            )

            # Download rekap
            output_pivot = io.BytesIO()

            with pd.ExcelWriter(
                output_pivot,
                engine='openpyxl'
            ) as writer:

                pivot_filter.to_excel(
                    writer,
                    index=False,
                    sheet_name='Rekap Seller'
                )

            st.download_button(
                label="⬇️ Download Rekap Seller Excel",
                data=output_pivot.getvalue(),
                file_name=f"Rekap_Seller_Mikro_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:

        st.error(f"Terjadi kesalahan saat memproses data: {e}")

else:

    st.info(
        "Silakan upload file Excel atau CSV melalui menu di sebelah kiri."
    )

    st.markdown("""
    ### Panduan Upload
    1. Upload satu atau beberapa file sekaligus.
    2. File Excel harus memiliki sheet bernama `MIKRO`.
    3. Baris ke-3 digunakan sebagai header.
    4. Data dimulai dari baris ke-4.
    5. Dashboard akan menggabungkan seluruh file yang di-upload.
    """)
