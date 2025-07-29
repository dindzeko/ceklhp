import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import numpy as np

# Judul aplikasi
st.title("📊 Aplikasi Harga Saham 15 Hari Perdagangan Terakhir")
st.write("""
Upload file Excel berisi daftar ticker saham, pilih tanggal, 
lalu ambil harga closing 15 hari perdagangan terakhir hingga tanggal tersebut dari Yahoo Finance.
""")

# Input tanggal
selected_date = st.date_input(
    "Pilih tanggal akhir (termasuk) untuk pengambilan data",
    value=datetime.today().date(),
    help="Data akan diambil dari 15 hari perdagangan terakhir hingga tanggal ini."
)

# Upload file Excel
uploaded_file = st.file_uploader("Upload file Excel (.xlsx) yang berisi kolom 'Ticker'", type=["xlsx"])

if uploaded_file:
    try:
        # Baca file Excel
        df_tickers = pd.read_excel(uploaded_file)
        
        if 'Ticker' not in df_tickers.columns:
            st.error("File Excel harus memiliki kolom bernama 'Ticker'")
        else:
            tickers = df_tickers['Ticker'].dropna().astype(str).str.strip()
            tickers_list = tickers.tolist()
            
            st.write("### Ticker yang ditemukan:")
            st.write(tickers_list)

            if st.button("🔍 Ambil Data Harga Closing"):
                with st.spinner("Mengambil data dari Yahoo Finance..."):
                    data = {}
                    # Konversi ke datetime
                    end_date = datetime.combine(selected_date, datetime.min.time()) + timedelta(days=1)  # agar inclusive
                    start_date = end_date - timedelta(days=45)  # ambil window lebih lebar karena ada libur

                    for ticker in tickers_list:
                        try:
                            stock = yf.Ticker(ticker)
                            # Ambil data harian
                            hist = stock.history(start=start_date, end=end_date, interval="1d")
                            
                            if hist.empty:
                                data[ticker] = ["No data"] * 15
                                continue

                            # Urutkan dari terbaru ke terlama, ambil 15 baris terakhir (terbaru)
                            closing_prices = hist['Close'].sort_index(ascending=False).head(15)
                            
                            # Jika kurang dari 15, isi dengan NaN
                            if len(closing_prices) < 15:
                                closing_prices = closing_prices.reindex(
                                    index=[None]* (15 - len(closing_prices)) + list(closing_prices.index)
                                ).fillna("N/A")

                            # Balik ke urutan H-15 (tertua) ke H-1 (terbaru)
                            closing_list = closing_prices[::-1].tolist()
                            data[ticker] = closing_list

                        except Exception as e:
                            data[ticker] = [f"Error: {str(e)}"] * 15

                    # Buat DataFrame dengan label hari dan tanggal perdagangan
                    result_df = pd.DataFrame(data)
                    
                    # Ubah index menjadi tanggal perdagangan
                    trading_dates = result_df.index.to_series().apply(lambda x: x.strftime('%Y-%m-%d'))
                    result_df.index = trading_dates
                    
                    # Transpose DataFrame agar ticker jadi row dan tanggal jadi kolom
                    result_df = result_df.T
                    
                    # Reset index untuk membuat ticker sebagai kolom pertama
                    result_df.reset_index(inplace=True)
                    result_df.rename(columns={'index': 'Saham'}, inplace=True)
                    
                    # Tampilkan hasil
                    st.success("Data berhasil diambil!")
                    st.write(f"### Harga Closing 15 Hari Perdagangan Terakhir hingga {selected_date}")
                    st.dataframe(result_df)

                    # Grafik
                    st.write("### Grafik Harga Closing")
                    chart_data = result_df.set_index('Saham').T
                    numeric_data = chart_data.apply(pd.to_numeric, errors='coerce')
                    if not numeric_data.isnull().all().all():
                        st.line_chart(numeric_data)
                    else:
                        st.info("Tidak ada data numerik untuk ditampilkan dalam grafik.")

                    # Download Excel
                    output_filename = f"harga_closing_15hari_hingga_{selected_date}.xlsx"
                    st.download_button(
                        label="📥 Download sebagai Excel",
                        data=result_df.to_excel(index=False),
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca file: {e}")
else:
    st.info("Silakan upload file Excel yang berisi kolom 'Ticker'.")
    st.markdown("""
    **Contoh format file Excel:**
    | Ticker   |
    |----------|
    | BBCA.JK  |
    | UNVR.JK  |
    | TLKM.JK  |
    """)
