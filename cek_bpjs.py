from io import BytesIO
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Laporan Analisis Data & BPJS", page_icon="📊", layout="wide"
)

st.title("📊 Laporan Komprehensif Pencocokan Data Perusahaan & BPJS")
st.write(
    "Unggah **File 1** dan **File 2** untuk menghasilkan **5 kategori laporan"
    " perbandingan** dan mengunduhnya dalam **1 file Excel**."
)
st.markdown("---")

# Area Unggah File
col1, col2 = st.columns(2)
with col1:
  file1 = st.file_uploader(
      "Unggah File 1 (Master / Rujukan Utama)", type=["xlsx", "csv"], key="f1"
  )
with col2:
  file2 = st.file_uploader(
      "Unggah File 2 (Peserta BPJS)", type=["xlsx", "csv"], key="f2"
  )

if file1 is not None and file2 is not None:
  try:
    # Membaca file otomatis
    df1 = (
        pd.read_csv(file1)
        if file1.name.endswith(".csv")
        else pd.read_excel(file1)
    )
    df2 = (
        pd.read_csv(file2)
        if file2.name.endswith(".csv")
        else pd.read_excel(file2)
    )

    st.success("Kedua file berhasil dimuat!")

    # Pilih Kolom Nama dari masing-masing file
    col_opt1, col_opt2 = st.columns(2)
    with col_opt1:
      kolom1 = st.selectbox(
          "Pilih Kolom Nama Perusahaan di File 1:", df1.columns.tolist()
      )
    with col_opt2:
      kolom2 = st.selectbox(
          "Pilih Kolom Nama Perusahaan di File 2:", df2.columns.tolist()
      )

    if st.button("Proses Analisis 5 Kategori Laporan", type="primary"):
      # Pembersihan data (string, buang spasi, huruf kecil)
      s1_raw = df1[kolom1].dropna().astype(str).str.strip()
      s2_raw = df2[kolom2].dropna().astype(str).str.strip()

      s1_clean = s1_raw.str.lower()
      s2_clean = s2_raw.str.lower()

      # 1. Data Terinput Dobel File 1 (Duplikat internal di File 1)
      df_dobel1 = (
          df1[s1_clean.duplicated(keep=False)]
          .sort_values(by=kolom1)
          .reset_index(drop=True)
      )

      # 2. Data Terinput Dobel File 2 (Duplikat internal di File 2)
      df_dobel2 = (
          df2[s2_clean.duplicated(keep=False)]
          .sort_values(by=kolom2)
          .reset_index(drop=True)
      )

      # Set untuk perbandingan antar file
      set_1 = set(s1_clean)
      set_2 = set(s2_clean)

      # 3. Sudah Jadi Peserta (Ada di File 1 DAN ada di File 2)
      irisan = sorted(list(set_1.intersection(set_2)))
      df_sudah = pd.DataFrame(
          {
              "No": range(1, len(irisan) + 1),
              "Nama Perusahaan (Sudah Jadi Peserta)": irisan,
          }
      )

      # 4. Belum Jadi Peserta (Ada di File 1 TAPI TIDAK ADA di File 2)
      belum = sorted(list(set_1.difference(set_2)))
      df_belum = pd.DataFrame(
          {
              "No": range(1, len(belum) + 1),
              "Nama Perusahaan (Belum Jadi Peserta)": belum,
          }
      )

      # 5. Anomali (Ada di File 2 TAPI TIDAK ADA di File 1)
      anomali = sorted(list(set_2.difference(set_1)))
      df_anomali = pd.DataFrame(
          {
              "No": range(1, len(anomali) + 1),
              "Nama Perusahaan (Anomali)": anomali,
          }
      )

      st.markdown("---")
      st.subheader("📋 Ringkasan Hasil Laporan")

      # Metrik Angka Sesuai 5 Kategori
      m1, m2, m3, m4, m5 = st.columns(5)
      m1.metric("1. Dobel File 1", f"{len(df_dobel1)}")
      m2.metric("2. Dobel File 2", f"{len(df_dobel2)}")
      m3.metric("3. Sudah Peserta", f"{len(irisan)}")
      m4.metric("4. Belum Peserta", f"{len(belum)}")
      m5.metric("5. Anomali", f"{len(anomali)}")

      st.markdown("---")

      # Tombol Download 1 File Excel Multi-Sheet (5 Sheet)
      output_all = BytesIO()
      with pd.ExcelWriter(output_all, engine="openpyxl") as writer:
        df_dobel1.to_excel(
            writer, index=False, sheet_name="1_Dobel_File_1"
        )
        df_dobel2.to_excel(
            writer, index=False, sheet_name="2_Dobel_File_2"
        )
        df_sudah.to_excel(writer, index=False, sheet_name="3_Sudah_Jadi_Peserta")
        df_belum.to_excel(writer, index=False, sheet_name="4_Belum_Jadi_Peserta")
        df_anomali.to_excel(writer, index=False, sheet_name="5_Anomali")

      st.download_button(
          "📥 Download Semua 5 Laporan ke 1 File Excel",
          data=output_all.getvalue(),
          file_name="Laporan_Lengkap_Analisis_BPJS.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
          type="primary",
      )

      st.markdown("---")

      # Tab Tampilan Preview di Layar (5 Tab)
      tab1, tab2, tab3, tab4, tab5 = st.tabs([
          "📁 1. Dobel File 1",
          "📁 2. Dobel File 2",
          "🤝 3. Sudah Jadi Peserta",
          "📄 4. Belum Jadi Peserta",
          "⚠️ 5. Anomali",
      ])

      with tab1:
        st.markdown(
            "### 1. Data Terinput Dobel (Duplikat di Internal File 1)"
        )
        if not df_dobel1.empty:
          st.dataframe(df_dobel1, use_container_width=True)
        else:
          st.info("Tidak ada data dobel di File 1.")

      with tab2:
        st.markdown(
            "### 2. Data Terinput Dobel (Duplikat di Internal File 2)"
        )
        if not df_dobel2.empty:
          st.dataframe(df_dobel2, use_container_width=True)
        else:
          st.info("Tidak ada data dobel di File 2.")

      with tab3:
        st.markdown(
            "### 3. Sudah Jadi Peserta (Terdaftar di File 1 & File 2)"
        )
        if not df_sudah.empty:
          st.dataframe(df_sudah, use_container_width=True, hide_index=True)
        else:
          st.info("Tidak ada data yang cocok di kedua file.")

      with tab4:
        st.markdown(
            "### 4. Belum Jadi Peserta (Ada di File 1, Tapi Tidak Ada di File"
            " 2)"
        )
        if not df_belum.empty:
          st.dataframe(df_belum, use_container_width=True, hide_index=True)
        else:
          st.info("Semua perusahaan di File 1 sudah tercakup.")

      with tab5:
        st.markdown(
            "### 5. Anomali (Ada di File 2, Tapi Tidak Terdaftar di File 1)"
        )
        if not df_anomali.empty:
          st.dataframe(df_anomali, use_container_width=True, hide_index=True)
        else:
          st.info("Tidak ditemukan data anomali.")

      st.markdown("---")
      # Tombol Selesai untuk mereset dan mulai pemisahan file baru
      if st.button("🔄 Selesai & Mulai Pemisahan Data Baru"):
        st.rerun()

  except Exception as e:
    st.error(f"Terjadi kesalahan saat memproses file: {e}")
else:
  st.info("Silakan unggah kedua file Excel atau CSV di atas untuk memulai.")
