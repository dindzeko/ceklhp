import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import io

# Judul aplikasi
st.title("📊 Aplikasi Tarik Data Saham")

# Pilih Time Frame
st.subheader("Time Frame")
timeframe = st.radio("Pilih interval data:", ["30m", "60m", "1d"], index=2, horizontal=True)

# Input jumlah hari perdagangan
days = st.number_input("Masukkan jumlah hari perdagangan yang ingin diambil:", min_value=1, max_value=180, value=15)

# Input metode ticker
st.subheader("Ticker Saham")
ticker_input_method = st.radio("Pilih cara input ticker:", ["Upload Excel", "Input Manual"], horizontal=True)

tickers_list = []

if ticker_input_method == "Upload Excel":
    uploaded_file = st.file_uploader("Upload file Excel (.xlsx) yang berisi kolom 'Ticker'", type=["xlsx"])
    if uploaded_file:
        try:
            df_tickers = pd.read_excel(uploaded_file)
            if 'Ticker' not in df_tickers.columns:
                st.error("File Excel harus memiliki kolom bernama 'Ticker'")
            else:
                tickers_list = df_tickers['Ticker'].dropna().astype(str).str.strip().tolist()
                st.write("### Ticker yang ditemukan:")
                st.write(tickers_list)
        except Exception as e:
            st.error(f"Terjadi kesalahan saat membaca file: {e}")
else:
    manual_tickers = st.text_area("Masukkan daftar ticker (pisahkan dengan koma):", "BBCA.JK, TLKM.JK")
    if manual_tickers:
        tickers_list = [x.strip() for x in manual_tickers.split(",") if x.strip()]

# Tombol ambil data
if st.button("🔍 Ambil Data"):
    if not tickers_list:
        st.warning("Silakan input ticker saham terlebih dahulu.")
    else:
        with st.spinner("Mengambil data dari Yahoo Finance..."):
            end_date = datetime.today()
            start_date = end_date - timedelta(days=days * 2)  # ambil lebih banyak untuk jaga-jaga hari libur
            
            data = {}
            for ticker in tickers_list:
                try:
                    stock = yf.Ticker(ticker)
                    hist = stock.history(start=start_date, end=end_date, interval=timeframe)
                    
                    if hist.empty:
                        data[ticker] = ["No data"] * days
                        continue

                    # Ambil hanya 'Close'
                    closing_prices = hist['Close'].sort_index(ascending=False).head(days)[::-1]
                    data[ticker] = closing_prices.tolist()

                except Exception as e:
                    data[ticker] = [f"Error: {str(e)}"] * days

            # Buat DataFrame hasil
            result_df = pd.DataFrame(data)

            # Buat kolom tanggal
            if not result_df.empty:
                result_df.index = hist.index[-days:]
                result_df.reset_index(inplace=True)
                result_df.rename(columns={'index': 'Tanggal'}, inplace=True)

                st.success("✅ Data berhasil diambil!")
                st.dataframe(result_df)

                # Download Excel
                output = io.BytesIO()
                result_df.to_excel(output, index=False)
                st.download_button(
                    label="📥 Download sebagai Excel",
                    data=output,
                    file_name=f"data_saham_{timeframe}_{days}hari.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Tidak ada data yang berhasil diambil.")
