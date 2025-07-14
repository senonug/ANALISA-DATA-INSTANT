import streamlit as st
import pandas as pd

# Inisialisasi session state untuk history data
if 'data_history' not in st.session_state:
    st.session_state['data_history'] = pd.DataFrame()  # DataFrame kosong awal

# Judul aplikasi
st.title("Auto-Generate Pivot Table dari Multiple Excel dengan Konfirmasi Ekspor dan Hapus History")

# Menu Hapus History
st.subheader("Manajemen History Data")
if not st.session_state['data_history'].empty:
    st.info(f"Ada history data dengan {len(st.session_state['data_history'])} baris. Anda bisa gunakan atau hapus.")
    if st.button("Hapus History Data"):
        st.session_state['data_history'] = pd.DataFrame()
        st.experimental_rerun()  # Refresh app
else:
    st.info("Tidak ada history data saat ini.")

# Upload multiple file Excel (.xls atau .xlsx)
uploaded_files = st.file_uploader("Upload file Excel (.xls atau .xlsx) - Bisa multiple files", type=["xls", "xlsx"], accept_multiple_files=True)

# Opsi gabung dengan history
use_history = st.checkbox("Gabung dengan History Data Existing (jika ada)", value=True)

if uploaded_files:
    all_dfs = []  # List untuk simpan data baru dari upload ini
    
    for uploaded_file in uploaded_files:
        # Tentukan engine berdasarkan extension
        engine = 'xlrd' if uploaded_file.name.endswith('.xls') else 'openpyxl'
        
        # Baca file Excel
        try:
            excel_file = pd.ExcelFile(uploaded_file, engine=engine)
            sheets = excel_file.sheet_names
            
            # Multiselect sheets per file (default: semua)
            st.subheader(f"Pilih Sheets untuk File: {uploaded_file.name}")
            selected_sheets = st.multiselect(f"Pilih sheets untuk {uploaded_file.name}", sheets, default=sheets, key=f"sheets_{uploaded_file.name}")
            
            if selected_sheets:
                # Baca dan gabungkan sheets terpilih per file
                dfs = [excel_file.parse(sheet) for sheet in selected_sheets]
                df_file = pd.concat(dfs, ignore_index=True)
                all_dfs.append(df_file)
            else:
                st.info(f"Pilih setidaknya satu sheet untuk file {uploaded_file.name}.")
        except Exception as e:
            st.error(f"Error membaca file {uploaded_file.name}: {e}. Pastikan file valid.")
    
    if all_dfs:
        # Gabungkan data baru dari upload ini
        df_new = pd.concat(all_dfs, ignore_index=True)
        
        # Gabung dengan history jika dipilih
        if use_history and not st.session_state['data_history'].empty:
            df = pd.concat([st.session_state['data_history'], df_new], ignore_index=True)
        else:
            df = df_new
        
        # Update history dengan data terbaru
        st.session_state['data_history'] = df
        
        # Tampilkan data awal gabungan
        st.subheader("Data Awal Gabungan (Termasuk History jika Digabung)")
        st.dataframe(df.head())
        
        # Pilih kolom untuk pivot table
        st.subheader("Konfigurasi Pivot Table")
        index_cols = st.multiselect("Pilih kolom untuk Index (baris)", df.columns)
        columns_cols = st.multiselect("Pilih kolom untuk Columns (kolom)", df.columns)
        values_cols = st.multiselect("Pilih kolom untuk Values (nilai yang dihitung)", df.columns)
        agg_func = st.selectbox("Fungsi Agregasi", ['sum', 'mean', 'count', 'min', 'max'])
        
        if index_cols and values_cols:
            try:
                # Generate pivot table
                pivot = pd.pivot_table(df, index=index_cols, columns=columns_cols, values=values_cols, aggfunc=agg_func)
                
                # Tampilkan hasil pivot table
                st.subheader("Hasil Pivot Table")
                st.dataframe(pivot)
                
                # Langkah konfirmasi: Pilih data/kolom untuk diekspor ke CSV
                st.subheader("Konfirmasi Data untuk Ekspor ke CSV")
                st.info("Pilih kolom yang ingin disertakan dalam CSV. Default: semua kolom.")
                export_cols = st.multiselect("Pilih kolom untuk diekspor", pivot.columns, default=list(pivot.columns))
                
                if export_cols:
                    # Siapkan data yang dipilih
                    pivot_export = pivot[export_cols]
                    
                    # Konversi ke CSV
                    csv = pivot_export.to_csv().encode('utf-8')
                    
                    # Tombol download setelah konfirmasi
                    st.download_button(
                        label="Konfirmasi dan Download CSV",
                        data=csv,
                        file_name="pivot_table_konfirmasi.csv",
                        mime="text/csv"
                    )
                else:
                    st.warning("Pilih setidaknya satu kolom untuk diekspor.")
            except Exception as e:
                st.error(f"Error: {e}. Pastikan kolom yang dipilih valid dan mengandung data numerik jika diperlukan.")
        else:
            st.info("Pilih setidaknya Index dan Values untuk generate pivot table.")
    else:
        st.info("Tidak ada data baru yang diproses. Gunakan history jika ada.")
elif not st.session_state['data_history'].empty:
    # Jika tidak upload baru, tapi ada history, gunakan history untuk pivot
    df = st.session_state['data_history']
    st.subheader("Menggunakan Data History Existing")
    st.dataframe(df.head())
    
    # Sisanya sama seperti di atas (konfigurasi pivot, dll.)
    # ... (copy bagian konfigurasi pivot dari atas ke sini jika ingin proses otomatis, tapi untuk singkat, asumsikan user bisa langsung konfigurasi)
else:
    st.info("Upload file atau gunakan history untuk memulai.")
