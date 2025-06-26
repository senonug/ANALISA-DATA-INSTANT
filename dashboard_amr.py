
import streamlit as st
import pandas as pd
import os
import plotly.express as px

# Fungsi logika teknis
def cek_indikator(row):
    indikator = {}

    indikator['arus_hilang'] = all([row['CURRENT_L1'] == 0, row['CURRENT_L2'] == 0, row['CURRENT_L3'] == 0])
    indikator['over_current'] = any([row['CURRENT_L1'] > 100, row['CURRENT_L2'] > 100, row['CURRENT_L3'] > 100])
    indikator['over_voltage'] = any([row['VOLTAGE_L1'] > 240, row['VOLTAGE_L2'] > 240, row['VOLTAGE_L3'] > 240])
    v = [row['VOLTAGE_L1'], row['VOLTAGE_L2'], row['VOLTAGE_L3']]
    indikator['v_drop'] = max(v) - min(v) > 10
    indikator['cos_phi_kecil'] = any([row.get(f'POWER_FACTOR_L{i}', 1) < 0.85 for i in range(1, 4)])
    indikator['active_power_negative'] = any([row.get(f'ACTIVE_POWER_L{i}', 0) < 0 for i in range(1, 4)])
    indikator['arus_kecil_teg_kecil'] = all([
        all([row['CURRENT_L1'] < 1, row['CURRENT_L2'] < 1, row['CURRENT_L3'] < 1]),
        all([row['VOLTAGE_L1'] < 180, row['VOLTAGE_L2'] < 180, row['VOLTAGE_L3'] < 180]),
        any([row.get(f'ACTIVE_POWER_L{i}', 0) > 10 for i in range(1, 4)])
    ])
    arus = [row['CURRENT_L1'], row['CURRENT_L2'], row['CURRENT_L3']]
    max_i, min_i = max(arus), min(arus)
    indikator['unbalance_I'] = (max_i - min_i) / max_i > 0.15 if max_i > 0 else False

    indikator['v_lost'] = row.get('VOLTAGE_L1', 0) == 0 or row.get('VOLTAGE_L2', 0) == 0 or row.get('VOLTAGE_L3', 0) == 0
    indikator['In_more_Imax'] = any([row['CURRENT_L1'] > 120, row['CURRENT_L2'] > 120, row['CURRENT_L3'] > 120])
    indikator['active_power_negative_siang'] = row.get('ACTIVE_POWER_SIANG', 0) < 0
    indikator['active_power_negative_malam'] = row.get('ACTIVE_POWER_MALAM', 0) < 0
    indikator['active_p_lost'] = row.get('ACTIVE_POWER_L1', 0) == 0 and row.get('ACTIVE_POWER_L2', 0) == 0 and row.get('ACTIVE_POWER_L3', 0) == 0
    indikator['current_loop'] = row.get('CURRENT_LOOP', 0) == 1
    indikator['freeze'] = row.get('FREEZE', 0) == 1

    return indikator

st.set_page_config(page_title="Dashboard TO AMR", layout="wide")
st.title("📊 Dashboard Target Operasi AMR - P2TL")
st.markdown("---")

uploaded_file = st.file_uploader("📥 Upload File Excel AMR Harian", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0)
    df = df.dropna(subset=['LOCATION_CODE'])
    df = df.copy()

    num_cols = [
        'CURRENT_L1', 'CURRENT_L2', 'CURRENT_L3',
        'VOLTAGE_L1', 'VOLTAGE_L2', 'VOLTAGE_L3',
        'ACTIVE_POWER_L1', 'ACTIVE_POWER_L2', 'ACTIVE_POWER_L3',
        'POWER_FACTOR_L1', 'POWER_FACTOR_L2', 'POWER_FACTOR_L3',
        'ACTIVE_POWER_SIANG', 'ACTIVE_POWER_MALAM', 'CURRENT_LOOP', 'FREEZE'
    ]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Gabungkan dengan data historis jika ada
    data_path = "data_harian.csv"
    if os.path.exists(data_path):
        df_hist = pd.read_csv(data_path)
        df = pd.concat([df_hist, df], ignore_index=True).drop_duplicates()

    # Simpan data hasil gabungan
    df.to_csv(data_path, index=False)

    indikator_list = df.apply(cek_indikator, axis=1)
    indikator_df = pd.DataFrame(indikator_list.tolist())
    result = pd.concat([df[['LOCATION_CODE']], indikator_df], axis=1)

    result['Jumlah Potensi TO'] = indikator_df.sum(axis=1)
    top50 = result.sort_values(by='Jumlah Potensi TO', ascending=False).head(50)

    # Ringkasan
    col1, col2, col3 = st.columns(3)
    col1.metric("📄 Total Data", len(df))
    col2.metric("🔢 Total IDPEL Unik", df['LOCATION_CODE'].nunique())
    col3.metric("🎯 Potensi Target Operasi", sum(result['Jumlah Potensi TO'] > 0))

    st.markdown("---")
    st.subheader("🏆 Top 50 Rekomendasi Target Operasi")
    st.dataframe(top50, use_container_width=True)

    st.markdown("---")
    st.subheader("📈 Visualisasi Indikator Anomali")
    indikator_counts = indikator_df.sum().sort_values(ascending=False).reset_index()
    indikator_counts.columns = ['Indikator', 'Jumlah']
    fig = px.bar(indikator_counts, x='Indikator', y='Jumlah', text='Jumlah', color='Indikator')
    st.plotly_chart(fig, use_container_width=True)

    # Fitur tambahan: Unduh hasil
    st.markdown("### 📤 Unduh Hasil Analisis")
    csv = result.to_csv(index=False).encode('utf-8')
    st.download_button("💾 Download CSV", csv, "hasil_analisis_to.csv", "text/csv")

    # Reset histori
    if st.button("🗑️ Hapus Semua Data Historis"):
        os.remove(data_path)
        st.success("Data historis berhasil dihapus.")
else:
    st.info("Silakan upload file Excel terlebih dahulu untuk memulai analisis.")
