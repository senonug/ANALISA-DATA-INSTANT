import streamlit as st
import pandas as pd

# Judul aplikasi
st.title("Auto-Generate Pivot Table dari Excel dengan Konfirmasi Ekspor")

# Upload file Excel
uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type="xlsx")

if uploaded_file is not None:
    # Baca file Excel
    df = pd.read_excel(uploaded_file)
    
    # Tampilkan data awal
    st.subheader("Data Awal dari Excel")
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
