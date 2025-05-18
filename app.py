# main.py
import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO

st.set_page_config(page_title="Rekap Tiket", layout="centered")

tab1, tab2, tab3 = st.tabs(["Nominal Pengurangan", "Naik Turun Golongan", "Rek Koran vs Invoice"])

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



with tab3:
    st.header("Rek Koran vs Invoice")


import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta

st.title("Compare Rekening Koran vs Invoice - Output 10 Kolom Final")

st.header("Upload Rekening Koran (Data 1)")
file1 = st.file_uploader("Upload Excel Rekening Koran", type=["xls", "xlsx"], key="file1")

st.header("Upload Invoice (Data 2)")
file2 = st.file_uploader("Upload Excel Invoice", type=["xls", "xlsx"], key="file2")

def translate_bulan(text):
    bulan_map = {
        "JAN": "JAN", "FEB": "FEB", "MAR": "MAR", "APR": "APR", "MEI": "MAY",
        "JUN": "JUN", "JUL": "JUL", "AGU": "AUG", "SEP": "SEP", "OKT": "OCT",
        "NOV": "NOV", "DES": "DEC"
    }
    for indo, eng in bulan_map.items():
        text = text.replace(indo, eng)
    return text

def safe_strptime(s):
    try:
        return datetime.strptime(s, "%d %b %Y")
    except ValueError:
        return None

def extract_trx_range(desc):
    if pd.isnull(desc):
        return None
    match1 = re.search(r'TRX TGL ([0-9]{2} [A-Z]{3})(?:-([0-9]{2} [A-Z]{3}))? ([0-9]{4})', desc)
    if match1:
        start = f"{match1.group(1)} {match1.group(3)}"
        end = f"{match1.group(2)} {match1.group(3)}" if match1.group(2) else start
        return f"{start} - {end}" if start != end else start
    match2 = re.search(r'TRX TGL ([0-9]{2})(?:-([0-9]{2}))? ([A-Z]{3}) ([0-9]{4})', desc)
    if match2:
        hari1 = match2.group(1)
        hari2 = match2.group(2) or match2.group(1)
        bulan = match2.group(3)
        tahun = match2.group(4)
        start = f"{hari1} {bulan} {tahun}"
        end = f"{hari2} {bulan} {tahun}"
        return f"{start} - {end}" if start != end else start
    return None

def fix_invoice_summing(df1, df2):
    df2["Tanggal"] = df2["TANGGAL INVOICE"].dt.strftime("%d %b %Y").str.upper().str.strip()
    invoice_by_date = df2.groupby("Tanggal")["HARGA"].sum().to_dict()

    def sum_invoice(trx_range):
        if pd.isnull(trx_range):
            return 0
        if "-" not in trx_range:
            trx_range = translate_bulan(trx_range.strip())
            trx_date = safe_strptime(trx_range)
            if trx_date:
                key = trx_date.strftime("%d %b %Y").upper()
                return invoice_by_date.get(key, 0)
            return 0
        start_str, end_str = trx_range.split("-")
        start_date = safe_strptime(translate_bulan(start_str.strip()))
        end_date = safe_strptime(translate_bulan(end_str.strip()))
        if not start_date or not end_date:
            return 0
        date_list = [(start_date + timedelta(days=i)).strftime("%d %b %Y").upper()
                     for i in range((end_date - start_date).days + 1)]
        return sum(invoice_by_date.get(d, 0) for d in date_list)

    df1["Invoice"] = df1["Tanggal"].apply(sum_invoice)
    df1["Selisih"] = df1["Amount"] - df1["Invoice"]
    return df1

if file1 and file2:
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    df1["Post Date"] = pd.to_datetime(df1["Post Date"], dayfirst=True, errors='coerce')
    df1 = df1.dropna(subset=["Post Date", "Amount"])
    df1 = df1[(df1["Branch"].str.contains("UNIT E-CHANNEL", na=False)) & (df1["Amount"] > 100_000_000)].copy()
    df1["Tanggal"] = df1["Description"].apply(extract_trx_range)
    df1 = df1.dropna(subset=["Tanggal"])

    df2["TANGGAL INVOICE"] = pd.to_datetime(df2["TANGGAL INVOICE"], errors='coerce')
    df2 = df2.dropna(subset=["TANGGAL INVOICE", "HARGA"])

    df1 = fix_invoice_summing(df1, df2)

    df_final = df1[["Tanggal", "Post Date", "Branch", "Journal No.", "Description",
                    "Amount", "Invoice", "Selisih", "Db/Cr", "Balance"]]

    st.header("Output 10 Kolom Final")
    st.dataframe(df_final.fillna(""))

    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Compare Hasil')
    output.seek(0)

    st.download_button(
        label="Download Output Excel",
        data=output,
        file_name="hasil_compare_fixed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Silakan upload kedua file untuk melanjutkan.")