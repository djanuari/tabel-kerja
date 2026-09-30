import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Pencocokan & Pembersihan Data BPJS", page_icon="📊", layout="wide")

st.title("📊 Aplikasi Pencocokan & Pembersihan Data (WLKP vs BPJS)")
st.write("Unggah file data WLKP (File 1) dan data BPJS (File 2). Sistem akan otomatis membersihkan duplikat dan menyusun laporan lengkap dalam **1 file Excel**.")
st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    file1 = st.file_uploader("Unggah File 1 (Data Perusahaan WLKP)", type=["xlsx", "csv"], key="f1")
with col2:
    file2 = st.file_uploader("Unggah File 2 (Data Perusahaan BPJS)", type=["xlsx", "csv"], key="f2")

if file1 is not None and file2 is not None:
    try:
        df1 = pd.read_csv(file1) if file1.name.endswith(".csv") else pd.read_excel(file1)
        df2 = pd.read_csv(file2) if file2.name.endswith(".csv") else pd.read_excel(file2)

        st.success("Kedua file berhasil dimuat!")

        col_opt1, col_opt2 = st.columns(2)
        with col_opt1:
            kolom1 = st.selectbox("Pilih Kolom Nama Perusahaan di File 1 (WLKP):", df1.columns.tolist())
        with col_opt2:
            kolom2 = st.selectbox("Pilih Kolom Nama Perusahaan di File 2 (BPJS):", df2.columns.tolist())

        if st.button("Proses & Buat Laporan Lengkap", type="primary"):
            # 1. Membersihkan spasi dan menyamakan format teks menjadi huruf kecil untuk kunci pencocokan
            df1['clean_key'] = df1[kolom1].dropna().astype(str).str.strip().str.lower()
            df2['clean_key'] = df2[kolom2].dropna().astype(str).str.strip().str.lower()

            # 2. Menangkap SELURUH baris duplikat (menggunakan keep=False) untuk lembar khusus duplikat
            df1_duplikat = df1[df1.duplicated(subset=['clean_key'], keep=False)].sort_values(by='clean_key')
            df2_duplikat = df2[df2.duplicated(subset=['clean_key'], keep=False)].sort_values(by='clean_key')

            # 3. Data utama WLKP dan BPJS secara utuh
            df_wlkp_utuh = df1.drop(columns=['clean_key'])
            df_bpjs_utuh = df2.drop(columns=['clean_key'])

            # Versi bersih (unique key) untuk perbandingan silang
            df1_bersih_key = df1.drop_duplicates(subset=['clean_key'], keep='first')
            df2_bersih_key = df2.drop_duplicates(subset=['clean_key'], keep='first')

            set_1 = set(df1_bersih_key['clean_key'])
            set_2 = set(df2_bersih_key['clean_key'])

            kunci_sama = set_1.intersection(set_2)
            kunci_hanya_1 = set_1.difference(set_2) # Ada di WLKP, tidak ada di BPJS
            kunci_hanya_2 = set_2.difference(set_1) # Ada di BPJS, tidak ada di WLKP

            # Menyusun DataFrame untuk perbandingan silang
            df_terdaftar_bpjs = df1_bersih_key[df1_bersih_key['clean_key'].isin(kunci_sama)].drop(columns=['clean_key'])
            df_wlkp_tidak_bpjs = df1_bersih_key[df1_bersih_key['clean_key'].isin(kunci_hanya_1)].drop(columns=['clean_key'])
            df_bpjs_tidak_wlkp = df2_bersih_key[df2_bersih_key['clean_key'].isin(kunci_hanya_2)].drop(columns=['clean_key'])

            # Hitung jumlah baris duplikat murni untuk metrik
            jumlah_dup_1 = len(df1) - len(df1_bersih_key)
            jumlah_dup_2 = len(df2) - len(df2_bersih_key)

            st.markdown("---")
            st.subheader("📊 Ringkasan Hasil Analisis")

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Baris Duplikat File 1", f"{jumlah_dup_1} Baris")
            m2.metric("Baris Duplikat File 2", f"{jumlah_dup_2} Baris")
            m3.metric("Terdaftar di Keduanya", f"{len(df_terdaftar_bpjs)} Data")
            m4.metric("Perbandingan Silang", f"{len(df_wlkp_tidak_bpjs) + len(df_bpjs_tidak_wlkp)} Data")

            # Membuat 1 File Excel dengan lembar kerja (sheet) lengkap
            output_excel = BytesIO()
            with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                df1_duplikat.drop(columns=['clean_key']).to_excel(writer, index=False, sheet_name='Dup_File1_Dihapus')
                df2_duplikat.drop(columns=['clean_key']).to_excel(writer, index=False, sheet_name='Dup_File2_Dihapus')
                df_wlkp_utuh.to_excel(writer, index=False, sheet_name='WLKP_Utuh')
                df_bpjs_utuh.to_excel(writer, index=False, sheet_name='BPJS_Utuh')
                df_terdaftar_bpjs.to_excel(writer, index=False, sheet_name='Terdaftar_BPJS')
                df_wlkp_tidak_bpjs.to_excel(writer, index=False, sheet_name='WLKP_Tanpa_BPJS')
                df_bpjs_tidak_wlkp.to_excel(writer, index=False, sheet_name='BPJS_Tanpa_WLKP')

            st.markdown("---")
            st.download_button(
                label="📥 Download 1 File Laporan Lengkap (Semua Kategori)",
                data=output_excel.getvalue(),
                file_name="Laporan_Lengkap_Pencocokan_BPJS_WLKP.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")
else:
    st.info("Silakan unggah kedua file Excel atau CSV di atas untuk memulai analisis.")
