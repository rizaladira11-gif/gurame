
import streamlit as st
import pandas as pd
import io

# ==========================================
# KONFIGURASI HALAMAN
# ==========================================

st.set_page_config(
    page_title="Dashboard Mikro",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Dashboard Monitoring Produk Mikro")
st.caption("Data Amandira Mikro dan Asira Mikro")

# ==========================================
# SIDEBAR
# ==========================================

st.sidebar.header("Upload Data")

uploaded_files = st.sidebar.file_uploader(
    "Upload file Excel / CSV",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)

# ==========================================
# FUNGSI MEMBACA FILE
# ==========================================

def read_files(uploaded_files):

    data = []

    for file in uploaded_files:

        try:

            if file.name.endswith((".xlsx", ".xls")):

                df_temp = pd.read_excel(
                    io.BytesIO(file.getvalue()),
                    sheet_name="MIKRO",
                    header=None
                )

            elif file.name.endswith(".csv"):

                df_temp = pd.read_csv(
                    io.BytesIO(file.getvalue()),
                    header=None
                )

            else:
                continue

            # Mengambil header dari baris ketiga
            df_temp = df_temp.iloc[2:].reset_index(drop=True)

            # Baris pertama menjadi header
            df_temp.columns = df_temp.iloc[0]

            # Menghapus baris header dari isi data
            df_temp = df_temp.iloc[1:].reset_index(drop=True)

            # Menghapus kolom dengan nama kosong
            df_temp = df_temp.loc[
                :, df_temp.columns.notna()
            ]

            data.append(df_temp)

        except Exception as e:

            st.error(
                f"Gagal membaca {file.name}: {e}"
            )

    if data:
        return pd.concat(data, ignore_index=True)

    return pd.DataFrame()


# ==========================================
# PROSES DATA
# ==========================================

if uploaded_files:

    with st.spinner("Sedang memproses data..."):

        df = read_files(uploaded_files)

        if not df.empty:

            kolom = [
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

            # Validasi kolom
            missing = [
                col for col in kolom
                if col not in df.columns
            ]

            if missing:

                st.error(
                    f"Kolom tidak ditemukan: {missing}"
                )
                st.stop()

            df = df[kolom].copy()

            # ==================================
            # KONVERSI TIPE DATA
            # ==================================

            df['NILAI_PRODUK'] = pd.to_numeric(
                df['NILAI_PRODUK'],
                errors='coerce'
            )

            # ==================================
            # FILTER PRODUK
            # ==================================

            df = df[
                df['NAMA_PRODUK'].isin([
                    "Amandira Mikro",
                    "Asira Mikro"
                ])
            ]

            # Filter nilai produk
            df = df[
                df['NILAI_PRODUK'] > 75000
            ]

            # ==================================
            # FILTER JOB SELLER
            # ==================================

            job_list = (
                "CREDIT RETENTION OFFICER",
                "CREDIT MARKETING OFFICER",
                "CREDIT CHANNELING OFFICER",
                "RELATIONSHIP OFFICER",
                "SALES OFFICER",
                "SPG"
            )

            df = df[
                df['JOB SELLER']
                .fillna('')
                .astype(str)
                .str.strip()
                .str.startswith(job_list)
            ].reset_index(drop=True)

            # ==================================
            # FORMAT KODE
            # ==================================

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

            # ==================================
            # PIVOT TABLE
            # ==================================

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
                aggfunc={
                    'NILAI_PRODUK': ['count', 'sum']
                },
                fill_value=0
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

            pivot['KODE_CABANG'] = (
                pivot['KODE_CABANG']
                .fillna('')
                .astype(str)
                .str.zfill(4)
            )

            pivot['ID_SELLER'] = (
                pivot['ID_SELLER']
                .fillna('')
                .astype(str)
                .str.zfill(8)
            )

            pivot['REWARD'] = (
                pivot['TOTAL_POLIS'] * 10000
            )

            # ==================================
            # METRIC DASHBOARD
            # ==================================

            st.success("Data berhasil diproses!")

            col1, col2, col3, col4 = st.columns(4)

            with col1:
                st.metric(
                    "Total Data",
                    f"{len(df):,}"
                )

            with col2:
                st.metric(
                    "Total Polis",
                    f"{pivot['TOTAL_POLIS'].sum():,}"
                )

            with col3:
                st.metric(
                    "Total Premi",
                    f"Rp {pivot['TOTAL_PREMI'].sum():,.0f}"
                )

            with col4:
                st.metric(
                    "Total Reward",
                    f"Rp {pivot['REWARD'].sum():,.0f}"
                )

            st.divider()

            # ==================================
            # FUNGSI FILTER TABEL
            # ==================================

            def filter_table(data, key):

                filtered = data.copy()

                st.markdown("#### 🔎 Filter Kolom")

                cols = st.columns(3)

                for i, column in enumerate(data.columns):

                    with cols[i % 3]:

                        if (
                            pd.api.types.is_numeric_dtype(
                                data[column]
                            )
                        ):

                            min_val = float(
                                data[column].min()
                            ) if not data[column].empty else 0

                            max_val = float(
                                data[column].max()
                            ) if not data[column].empty else 0

                            if min_val < max_val:

                                selected = st.slider(
                                    column,
                                    min_val,
                                    max_val,
                                    (min_val, max_val),
                                    key=f"{key}_{column}"
                                )

                                filtered = filtered[
                                    filtered[column].between(
                                        selected[0],
                                        selected[1]
                                    )
                                ]

                            else:
                                st.caption(
                                    f"{column}: {min_val:,.0f}"
                                )

                        else:

                            options = sorted(
                                data[column]
                                .dropna()
                                .astype(str)
                                .unique()
                                .tolist()
                            )

                            selected = st.multiselect(
                                column,
                                options,
                                default=options,
                                key=f"{key}_{column}"
                            )

                            filtered = filtered[
                                filtered[column]
                                .astype(str)
                                .isin(selected)
                            ]

                return filtered

            # ==================================
            # TAB DASHBOARD
            # ==================================

            tab1, tab2 = st.tabs([
                "📋 Data Detail",
                "📊 Hasil Pivot"
            ])

            # ==================================
            # TAB 1 DATA DETAIL
            # ==================================

            with tab1:

                st.subheader(
                    "Data Detail Amandira & Asira Mikro"
                )

                st.write(
                    f"Total baris: {len(df):,}"
                )

                filtered_df = filter_table(
                    df,
                    "detail"
                )

                st.dataframe(
                    filtered_df,
                    use_container_width=True,
                    height=550,
                    hide_index=True
                )

                st.caption(
                    f"Menampilkan {len(filtered_df):,} dari "
                    f"{len(df):,} baris"
                )

                # Download data detail
                output_detail = io.BytesIO()

                with pd.ExcelWriter(
                    output_detail,
                    engine='openpyxl'
                ) as writer:

                    filtered_df.to_excel(
                        writer,
                        index=False,
                        sheet_name='DATA_DETAIL'
                    )

                st.download_button(
                    "⬇️ Download Data Detail",
                    data=output_detail.getvalue(),
                    file_name="data_detail.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # ==================================
            # TAB 2 PIVOT
            # ==================================

            with tab2:

                st.subheader(
                    "Hasil Pivot per Seller"
                )

                st.write(
                    f"Total seller: {len(pivot):,}"
                )

                filtered_pivot = filter_table(
                    pivot,
                    "pivot"
                )

                st.dataframe(
                    filtered_pivot,
                    use_container_width=True,
                    height=550,
                    hide_index=True
                )

                st.caption(
                    f"Menampilkan {len(filtered_pivot):,} dari "
                    f"{len(pivot):,} baris"
                )

                # Download pivot
                output_pivot = io.BytesIO()

                with pd.ExcelWriter(
                    output_pivot,
                    engine='openpyxl'
                ) as writer:

                    filtered_pivot.to_excel(
                        writer,
                        index=False,
                        sheet_name='PIVOT'
                    )

                st.download_button(
                    "⬇️ Download Hasil Pivot",
                    data=output_pivot.getvalue(),
                    file_name="hasil_pivot.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Download gabungan
                output_all = io.BytesIO()

                with pd.ExcelWriter(
                    output_all,
                    engine='openpyxl'
                ) as writer:

                    df.to_excel(
                        writer,
                        index=False,
                        sheet_name='DATA_DETAIL'
                    )

                    pivot.to_excel(
                        writer,
                        index=False,
                        sheet_name='PIVOT'
                    )

                st.divider()

                st.download_button(
                    "⬇️ Download Semua Data (Excel)",
                    data=output_all.getvalue(),
                    file_name=f"Report Tangkap Gurame 01 - {(today - pd.Timedelta(days=1)).strftime('%d %B %Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

        else:

            st.warning(
                "Tidak ada data yang berhasil dibaca."
            )

else:

    st.info(
        "Silakan upload file Excel atau CSV melalui sidebar untuk memulai."
    )
