from io import BytesIO
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Pencocokan & Pembersihan Data BPJS",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Aplikasi Pencocokan & Pembersihan Data (WLKP vs BPJS)")
st.write(
    "Unggah file data WLKP (File 1) dan data BPJS (File 2). Sistem akan menyusun"
    " laporan lengkap dalam **1 file Excel** dengan **5 sheet utama**,"
    " termasuk penggabungan informasi Telepon dan Email."
)
st.markdown("---")

col1, col2 = st.columns(2)
with col1:
  file1 = st.file_uploader(
      "Unggah File 1 (Data Perusahaan WLKP)", type=["xlsx", "xls", "csv"], key="f1"
  )
with col2:
  file2 = st.file_uploader(
      "Unggah File 2 (Data Perusahaan BPJS)", type=["xlsx", "xls", "csv"], key="f2"
  )


# Fungsi aman pembaca file Excel / CSV
def baca_file(uploaded_file):
  nama_file = uploaded_file.name.lower()
  if nama_file.endswith(".csv"):
    try:
      return pd.read_csv(uploaded_file, encoding="utf-8")
    except:
      uploaded_file.seek(0)
      return pd.read_csv(uploaded_file, encoding="latin1")
  elif nama_file.endswith(".xls"):
    return pd.read_excel(uploaded_file, engine="xlrd")
  else:
    try:
      return pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception:
      uploaded_file.seek(0)
      try:
        return pd.read_excel(uploaded_file, engine="calamine")
      except Exception as e:
        raise ValueError(f"Format file tidak valid atau rusak ({e}).")


if file1 is not None and file2 is not None:
  try:
    df1 = baca_file(file1)
    df2 = baca_file(file2)

    st.success("Kedua file berhasil dimuat!")

    col_opt1, col_opt2 = st.columns(2)
    with col_opt1:
      kolom1 = st.selectbox(
          "Pilih Kolom Nama Perusahaan di File 1 (WLKP):", df1.columns.tolist()
      )
    with col_opt2:
      kolom2 = st.selectbox(
          "Pilih Kolom Nama Perusahaan di File 2 (BPJS):", df2.columns.tolist()
      )

    # Pilihan tambahan untuk kolom Telepon dan Email jika ingin digabungkan
    st.markdown("---")
    st.subheader("Pengaturan Kolom Kontak (Opsional)")
    col_k1, col_k2, col_k3, col_k4 = st.columns(4)
    with col_k1:
      telp1 = st.selectbox(
          "Telp File 1 (WLKP):",
          ["-- Tidak Ada --"] + df1.columns.tolist(),
          key="t1",
      )
    with col_k2:
      email1 = st.selectbox(
          "Email File 1 (WLKP):",
          ["-- Tidak Ada --"] + df1.columns.tolist(),
          key="e1",
      )
    with col_k3:
      telp2 = st.selectbox(
          "Telp File 2 (BPJS):",
          ["-- Tidak Ada --"] + df2.columns.tolist(),
          key="t2",
      )
    with col_k4:
      email2 = st.selectbox(
          "Email File 2 (BPJS):",
          ["-- Tidak Ada --"] + df2.columns.tolist(),
          key="e2",
      )

    if st.button("Proses & Buat Laporan 5 Sheet", type="primary"):
      # 1. Normalisasi teks pencocokan
      df1["clean_key"] = (
          df1[kolom1]
          .dropna()
          .astype(str)
          .str.strip()
          .str.lower()
          .str.replace(r"\s+", " ", regex=True)
          .replace(["nan", "none", ""], pd.NA)
      )
      df2["clean_key"] = (
          df2[kolom2]
          .dropna()
          .astype(str)
          .str.strip()
          .str.lower()
          .str.replace(r"\s+", " ", regex=True)
          .replace(["nan", "none", ""], pd.NA)
      )

      # 2. Sheet 1 & 2: Baris duplikat murni dari masing-masing file
      df1_duplikat = (
          df1[
              df1.duplicated(subset=["clean_key"], keep=False)
              & df1["clean_key"].notna()
          ]
          .sort_values(by=kolom1, ascending=True)
          .drop(columns=["clean_key"])
      )
      df2_duplikat = (
          df2[
              df2.duplicated(subset=["clean_key"], keep=False)
              & df2["clean_key"].notna()
          ]
          .sort_values(by=kolom2, ascending=True)
          .drop(columns=["clean_key"])
      )

      # 3. Data bersih unik dari masing-masing file
      df1_bersih_key = df1.dropna(subset=["clean_key"]).drop_duplicates(
          subset=["clean_key"], keep="first"
      )
      df2_bersih_key = df2.dropna(subset=["clean_key"]).drop_duplicates(
          subset=["clean_key"], keep="first"
      )

      # Menggabungkan informasi Telepon & Email dari kedua file secara berdampingan untuk perbandingan
      # Kita buat tabel gabungan berdasarkan clean_key
      merge_cols_1 = [kolom1]
      if telp1 != "-- Tidak Ada --":
        merge_cols_1.append(telp1)
      if email1 != "-- Tidak Ada --":
        merge_cols_1.append(email1)

      merge_cols_2 = [kolom2]
      if telp2 != "-- Tidak Ada --":
        merge_cols_2.append(telp2)
      if email2 != "-- Tidak Ada --":
        merge_cols_2.append(email2)

      # Merge data untuk melihat info lengkap berdampingan
      df_merged = pd.merge(
          df1_bersih_key[["clean_key"] + [c for c in merge_cols_1 if c != kolom1]],
          df2_bersih_key[["clean_key"] + [c for c in merge_cols_2 if c != kolom2]],
          on="clean_key",
          how="outer",
          suffixes=("_WLKP", "_BPJS"),
      )

      # Kembalikan nama perusahaan asli
      df_gabung_full = pd.merge(
          df1_bersih_key, df_merged, on="clean_key", how="inner"
      )

      set_1 = set(df1_bersih_key["clean_key"])
      set_2 = set(df2_bersih_key["clean_key"])

      kunci_sama = set_1.intersection(set_2)
      kunci_hanya_1 = set_1.difference(set_2)  # Ada di WLKP, Tidak ada di BPJS
      kunci_hanya_2 = set_2.difference(set_1)  # Ada di BPJS, Tidak ada di WLKP

      # Sheet 3: Ada di WLKP & Ada di BPJS
      df_keduanya = (
          df1_bersih_key[df1_bersih_key["clean_key"].isin(kunci_sama)]
          .sort_values(by=kolom1, ascending=True)
          .drop(columns=["clean_key"])
          .reset_index(drop=True)
      )

      # Sheet 4: Ada di WLKP tapi Tidak ada di BPJS
      df_wlkp_saja = (
          df1_bersih_key[df1_bersih_key["clean_key"].isin(kunci_hanya_1)]
          .sort_values(by=kolom1, ascending=True)
          .drop(columns=["clean_key"])
          .reset_index(drop=True)
      )

      # Sheet 5: Ada di BPJS tapi Tidak ada di WLKP
      df_bpjs_saja = (
          df2_bersih_key[df2_bersih_key["clean_key"].isin(kunci_hanya_2)]
          .sort_values(by=kolom2, ascending=True)
          .drop(columns=["clean_key"])
          .reset_index(drop=True)
      )

      jumlah_dup_1 = len(df1.dropna(subset=["clean_key"])) - len(df1_bersih_key)
      jumlah_dup_2 = len(df2.dropna(subset=["clean_key"])) - len(df2_bersih_key)

      st.markdown("---")
      st.subheader("📊 Ringkasan Hasil Analisis 5 Kategori")

      r1, r2, r3, r4, r5 = st.columns(5)
      r1.metric("Duplikat File 1", f"{jumlah_dup_1}")
      r2.metric("Duplikat File 2", f"{jumlah_dup_2}")
      r3.metric("Ada di Keduanya", f"{len(df_keduanya)}")
      r4.metric("WLKP Saja", f"{len(df_wlkp_saja)}")
      r5.metric("BPJS Saja", f"{len(df_bpjs_saja)}")

      # Membuat 1 File Excel dengan tepat 5 Sheet utama
      output_excel = BytesIO()
      with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        df1_duplikat.to_excel(
            writer, index=False, sheet_name="Duplikat File 1"
        )
        df2_duplikat.to_excel(
            writer, index=False, sheet_name="Duplikat File 2"
        )
        df_keduanya.to_excel(
            writer, index=False, sheet_name="Ada di WLKP & BPJS"
        )
        df_wlkp_saja.to_excel(
            writer, index=False, sheet_name="WLKP Tanpa BPJS"
        )
        df_bpjs_saja.to_excel(
            writer, index=False, sheet_name="BPJS Tanpa WLKP"
        )

      st.markdown("---")
      st.download_button(
          label="📥 Download Laporan Lengkap (5 Sheet Sesuai Permintaan)",
          data=output_excel.getvalue(),
          file_name="Laporan_Lengkap_Pencocokan_5_Sheet.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
          type="primary",
      )

  except Exception as e:
    st.error(f"Terjadi kesalahan saat memproses file: {e}")
else:
    st.info("Silakan unggah kedua file Excel atau CSV di atas untuk memulai.")
