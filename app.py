# main.py
import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO

st.set_page_config(page_title="Rekap Tiket", layout="centered")

tab1, tab2 = st.tabs(["Nominal Pengurangan", "Naik Turun Golongan"])

with tab1:
    st.header("Nominal Pengurangan (00:00 - 08:00)")
    selected_date = st.date_input("Pilih tanggal", format="DD/MM/YYYY")
    uploaded_file = st.file_uploader("Upload file Excel Tiket Summary", type=["xlsx"], key="tab1")

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        df['CETAK BOARDING PASS'] = pd.to_datetime(df['CETAK BOARDING PASS'], errors='coerce')
        df['ASAL'] = df['ASAL'].str.upper()
        cabang_list = ["MERAK", "BAKAUHENI", "KETAPANG", "GILIMANUK", "CIWANDAN", "PANJANG"]

        start_dt = datetime.combine(selected_date, datetime.min.time())
        end_dt = datetime.combine(selected_date, datetime.min.time()).replace(hour=8)
        df_filtered = df[(df['CETAK BOARDING PASS'] >= start_dt) & (df['CETAK BOARDING PASS'] < end_dt)]

        results = []
        total_all = 0
        for cabang in cabang_list:
            total_tarif = df_filtered.loc[df_filtered['ASAL'] == cabang, 'TARIF'].sum()
            total_all += total_tarif
            formatted_tarif = f"{int(total_tarif):,}".replace(",", ".") if total_tarif else "0"
            results.append({"ASAL": cabang.capitalize(), "Nominal Pengurangan": formatted_tarif})
        for _ in range(3):
            results.append({"ASAL": "", "Nominal Pengurangan": ""})
        formatted_total_all = f"{int(total_all):,}".replace(",", ".")
        results.append({"ASAL": "Total", "Nominal Pengurangan": formatted_total_all})

        result_df = pd.DataFrame(results)
        st.subheader(f"Hasil: {selected_date.strftime('%d %B %Y')}")
        st.table(result_df)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame(results).to_excel(writer, index=False)
        st.download_button("Download Excel", output.getvalue(), "rekap_pengurangan.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Unggah file dan pilih tanggal.")

with tab2:
    st.header("Nominal Naik Turun Golongan")
    date_range = st.date_input("Pilih rentang tanggal (filter)", (date(2025, 5, 5), date(2025, 5, 13)), key="tab2")
    uploaded_tsum = st.file_uploader("Upload file Excel Tiket Summary", type=["xlsx"], key="tsum")
    uploaded_inv = st.file_uploader("Upload file Excel Invoice", type=["xlsx"], key="inv")

    if uploaded_tsum and uploaded_inv and isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
        end_date += pd.Timedelta(hours=23, minutes=59, seconds=59)

        df_tsum = pd.read_excel(uploaded_tsum)
        df_inv = pd.read_excel(uploaded_inv)
        df_tsum['PEMESANAN'] = pd.to_datetime(df_tsum['PEMESANAN'], errors='coerce')
        df_inv['TANGGAL INVOICE'] = pd.to_datetime(df_inv['TANGGAL INVOICE'], errors='coerce')

        df_tsum_filtered = df_tsum[(df_tsum['PEMESANAN'] >= start_date) & (df_tsum['PEMESANAN'] <= end_date)]
        df_inv_filtered = df_inv[(df_inv['TANGGAL INVOICE'] >= start_date) & (df_inv['TANGGAL INVOICE'] <= end_date)]

        if df_tsum_filtered.empty and df_inv_filtered.empty:
            st.warning("Tidak ditemukan data dalam rentang tanggal tersebut.")
        else:
            tsum_rows = pd.DataFrame({
                'NOMER INVOICE': df_tsum_filtered['NOMOR INVOICE'],
                'HARGA': -df_tsum_filtered['TARIF'],
                'KEBERANGKATAN': df_tsum_filtered['ASAL'].astype(str).str.strip().str.upper()
            })
            inv_rows = df_inv_filtered[['NOMER INVOICE', 'HARGA', 'KEBERANGKATAN']].copy()
            inv_rows['KEBERANGKATAN'] = inv_rows['KEBERANGKATAN'].astype(str).str.strip().str.upper()

            combined = pd.concat([inv_rows, tsum_rows], ignore_index=True)
            result = combined.groupby('KEBERANGKATAN')['HARGA'].sum().reset_index()
            result = result.rename(columns={'KEBERANGKATAN': 'ASAL', 'HARGA': 'Nominal Naik Turun Golongan'})
            result['ASAL'] = result['ASAL'].str.capitalize()
            result['Nominal Naik Turun Golongan'] = result['Nominal Naik Turun Golongan'].astype(int)

            urutan_cabang = ["Merak", "Bakauheni", "Ketapang", "Gilimanuk", "Ciwandan", "Panjang"]
            ordered_result = pd.DataFrame({"ASAL": urutan_cabang})
            ordered_result = ordered_result.merge(result, on="ASAL", how="left")
            ordered_result["Nominal Naik Turun Golongan"] = ordered_result["Nominal Naik Turun Golongan"].fillna(0).astype(int)

            blank_rows = pd.DataFrame([{"ASAL": "", "Nominal Naik Turun Golongan": ""}] * 3)
            total_sum = ordered_result["Nominal Naik Turun Golongan"].sum()
            total_row = pd.DataFrame([{"ASAL": "Total", "Nominal Naik Turun Golongan": total_sum}])
            result_with_total = pd.concat([ordered_result, blank_rows, total_row], ignore_index=True)

            result_display = result_with_total.copy()
            result_display['Nominal Naik Turun Golongan'] = result_display['Nominal Naik Turun Golongan'].apply(
                lambda x: f"{int(x):,}".replace(",", ".") if isinstance(x, int) else x
            )

            st.table(result_display)

            output = BytesIO()
            result_with_total.to_excel(output, index=False, engine='openpyxl')
            st.download_button("Download Hasil ke Excel", output.getvalue(), "naik_turun_golongan.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Silakan upload kedua file Excel dan pilih rentang tanggal untuk memulai perhitungan.")