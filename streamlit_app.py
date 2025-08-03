import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import io
import numpy as np
import time

# Konfigurasi halaman
st.set_page_config(
    page_title="Aplikasi Data Saham",
    page_icon="📈",
    layout="centered"
)

# Judul aplikasi dengan tampilan lebih menarik
st.title("📊 Aplikasi Analisis Data Saham")
st.markdown("""
<div style="background-color:#f0f2f6;padding:10px;border-radius:10px;margin-bottom:20px">
    <p style='text-align:center;font-size:16px;color:#333333'>
    Ambil data historis saham dari Yahoo Finance dan ekspor ke Excel
    </p>
</div>
""", unsafe_allow_html=True)

# Pilih Time Frame
st.subheader("⚙️ Pengaturan Data")
col1, col2 = st.columns(2)
with col1:
    timeframe = st.radio("**Pilih interval data:**", ["30m", "60m", "1d"], index=2, horizontal=True)
with col2:
    days = st.number_input("**Jumlah hari perdagangan:**", min_value=1, max_value=60, value=10)

# Input metode ticker
st.subheader("📋 Input Ticker Saham")
ticker_input_method = st.radio("Pilih cara input ticker:", ["Upload Excel", "Input Manual"], horizontal=True, label_visibility="collapsed")

tickers_list = []

# Input Ticker via Upload Excel
if ticker_input_method == "Upload Excel":
    uploaded_file = st.file_uploader("Upload file Excel (.xlsx) yang berisi kolom 'Ticker'", type=["xlsx"])
    if uploaded_file:
        try:
            df_tickers = pd.read_excel(uploaded_file)
            if 'Ticker' not in df_tickers.columns:
                st.error("❌ File Excel harus memiliki kolom bernama 'Ticker'")
            else:
                tickers_list = df_tickers['Ticker'].dropna().astype(str).str.strip().str.upper().tolist()
                st.success(f"✅ Ditemukan {len(tickers_list)} ticker")
                with st.expander("Lihat Daftar Ticker"):
                    st.write(tickers_list)
        except Exception as e:
            st.error(f"Terjadi kesalahan saat membaca file: {e}")

# Input Ticker Manual
else:
    manual_tickers = st.text_area("Masukkan daftar ticker (pisahkan dengan koma):", "BBCA.JK, TLKM.JK, BRIS.JK")
    if manual_tickers:
        tickers_list = [x.strip().upper() for x in manual_tickers.split(",") if x.strip()]
        st.info(f"ℹ️ {len(tickers_list)} ticker siap diambil")

# Tombol ambil data
if st.button("🚀 Ambil Data Saham", use_container_width=True, type="primary"):
    if not tickers_list:
        st.warning("Silakan input ticker saham terlebih dahulu.")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        end_date = datetime.today()
        start_date = end_date - timedelta(days=days * 3)  # buffer lebih besar untuk hari libur
        
        data_frames = []
        failed_tickers = []
        success_count = 0
        
        for i, ticker in enumerate(tickers_list):
            try:
                status_text.text(f"⏳ Mengambil data {ticker} ({i+1}/{len(tickers_list)})")
                progress_bar.progress((i+1) / len(tickers_list))
                
                stock = yf.Ticker(ticker)
                hist = stock.history(start=start_date, end=end_date, interval=timeframe)
                
                if hist.empty:
                    st.warning(f"⚠️ Data kosong untuk {ticker}")
                    failed_tickers.append(ticker)
                    time.sleep(0.3)
                    continue
                
                # Handle perbedaan kolom tanggal
                if 'Datetime' in hist.columns:
                    hist = hist.reset_index().rename(columns={'Datetime': 'Date'})
                else:
                    hist = hist.reset_index()
                
                # Ambil data N hari terakhir
                hist = hist.sort_values('Date', ascending=False).head(days)
                hist.insert(0, 'Ticker', ticker)
                
                # Konversi kolom tanggal ke format yang konsisten
                hist['Date'] = hist['Date'].dt.tz_localize(None)  # Hapus timezone
                hist['Date'] = hist['Date'].apply(lambda x: x.strftime('%Y-%m-%d %H:%M:%S'))
                
                data_frames.append(hist)
                success_count += 1
                time.sleep(0.3)  # Menghindari request terlalu cepat
                
            except Exception as e:
                st.error(f"❌ Gagal mengambil data {ticker}: {str(e)}")
                failed_tickers.append(ticker)
                time.sleep(0.5)

        if data_frames:
            result_df = pd.concat(data_frames, ignore_index=True)
            
            # Hapus kolom Dividends dan Stock Splits
            result_df = result_df.drop(columns=['Dividends', 'Stock Splits'], errors='ignore')
            
            st.success(f"✅ Berhasil mengambil data {success_count} dari {len(tickers_list)} ticker")
            
            if failed_tickers:
                st.warning(f"⚠️ Gagal mengambil data untuk: {', '.join(failed_tickers)}")
            
            # Tampilkan data
            with st.expander("📊 Lihat Data", expanded=True):
                st.dataframe(result_df.style.format({
                    'Open': '{:.2f}',
                    'High': '{:.2f}',
                    'Low': '{:.2f}',
                    'Close': '{:.2f}',
                    'Volume': '{:.0f}'
                }), use_container_width=True, height=400)
            
            # Download Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False)
            
            st.download_button(
                label="💾 Download Data Excel",
                data=output.getvalue(),
                file_name=f"data_saham_{timeframe}_{days}hari.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
        else:
            st.error("❌ Tidak ada data yang berhasil diambil. Silakan cek koneksi atau ticker Anda")
        
        progress_bar.empty()
        status_text.empty()
