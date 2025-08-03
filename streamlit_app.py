import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import io
import time

# Konfigurasi halaman
st.set_page_config(
    page_title="Aplikasi Data Saham Akurat",
    page_icon="📈",
    layout="centered"
)

# Judul aplikasi dengan penjelasan
st.title("📊 Aplikasi Data Saham Akurat")
st.markdown("""
<div style="background-color:#f0f2f6;padding:10px;border-radius:10px;margin-bottom:20px">
    <p style='text-align:center;font-size:16px;color:#333333'>
    Ambil data historis saham langsung dari Yahoo Finance
    </p>
</div>
""", unsafe_allow_html=True)

# Pilih Time Frame
st.subheader("⚙️ Pengaturan Data")
col1, col2 = st.columns(2)
with col1:
    timeframe = st.radio("**Interval data:**", ["30m", "60m", "1d"], index=2, horizontal=True)
with col2:
    days = st.number_input("**Jumlah hari perdagangan:**", min_value=1, max_value=60, value=10)

# Informasi penting tentang perbedaan harga
with st.expander("ℹ️ Mengapa harga bisa berbeda?", expanded=False):
    st.markdown("""
    **Perbedaan harga biasanya karena:**
    - **Harga Disesuaikan (Adjusted)**: 
        - Sudah dikoreksi untuk corporate actions (stock split, dividen)
        - Default di library yfinance
    - **Harga Aktual (Unadjusted)**:
        - Harga penutupan sesungguhnya di hari tersebut
        - Sama seperti di website Yahoo Finance
        
    Untuk analisis harga historis aktual, selalu pilih **Harga Aktual**.
    """)

# Input metode ticker
st.subheader("📋 Input Ticker Saham")
ticker_input_method = st.radio("Pilih cara input ticker:", ["Upload Excel", "Input Manual"], 
                              horizontal=True, label_visibility="collapsed")

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
    manual_tickers = st.text_area("Masukkan daftar ticker (pisahkan dengan koma):", "BBCA.JK, TLKM.JK, PTBA.JK")
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
        # Buffer lebih besar untuk intraday
        buffer_days = days * 5 if timeframe in ["30m", "60m"] else days * 3
        start_date = end_date - timedelta(days=buffer_days)
        
        data_frames = []
        failed_tickers = []
        success_count = 0
        
        for i, ticker in enumerate(tickers_list):
            try:
                status_text.text(f"⏳ Mengambil data {ticker} ({i+1}/{len(tickers_list)})")
                progress_bar.progress((i+1) / len(tickers_list))
                
                # Ambil data dengan harga aktual (unadjusted)
                stock = yf.Ticker(ticker)
                hist = stock.history(
                    start=start_date, 
                    end=end_date, 
                    interval=timeframe,
                    auto_adjust=False  # Pastikan harga tidak disesuaikan
                )
                
                if hist.empty:
                    st.warning(f"⚠️ Data kosong untuk {ticker}")
                    failed_tickers.append(ticker)
                    time.sleep(0.3)
                    continue
                
                # Reset index dan atur kolom tanggal
                hist = hist.reset_index()
                
                # Tangani perbedaan nama kolom tanggal
                if 'Datetime' in hist.columns:
                    hist.rename(columns={'Datetime': 'Date'}, inplace=True)
                
                # Pastikan kita menggunakan harga penutupan aktual
                hist['Close'] = hist['Close']  # Harga aktual
                
                # Ambil data N hari terakhir
                hist = hist.sort_values('Date', ascending=False).head(days)
                hist.insert(0, 'Ticker', ticker)
                
                # Konversi kolom tanggal
                hist['Date'] = hist['Date'].dt.tz_localize(None)
                
                # Format tanggal
                if timeframe == "1d":
                    hist['Date'] = hist['Date'].dt.strftime('%Y-%m-%d')
                else:
                    hist['Date'] = hist['Date'].dt.strftime('%Y-%m-%d %H:%M:%S')
                
                data_frames.append(hist)
                success_count += 1
                time.sleep(0.5)  # Menghindari rate limit
                
            except Exception as e:
                st.error(f"❌ Gagal mengambil data {ticker}: {str(e)}")
                failed_tickers.append(ticker)
                time.sleep(1)

        if data_frames:
            result_df = pd.concat(data_frames, ignore_index=True)
            
            # Pilih kolom yang relevan
            columns_to_keep = ['Ticker', 'Date', 'Open', 'High', 'Low', 'Close', 'Volume']
            result_df = result_df[columns_to_keep]
            
            st.success(f"✅ Berhasil mengambil data {success_count} dari {len(tickers_list)} ticker")
            
            if failed_tickers:
                st.warning(f"⚠️ Gagal mengambil data untuk: {', '.join(failed_tickers)}")
            
            # Tampilkan data
            with st.expander("📊 Lihat Data", expanded=True):
                # Format numerik
                st.dataframe(result_df.style.format({
                    'Open': '{:.2f}',
                    'High': '{:.2f}',
                    'Low': '{:.2f}',
                    'Close': '{:.2f}',
                    'Volume': '{:,.0f}'
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
            
            # Tips untuk memastikan keakuratan
            st.info("""
            **Tips untuk memastikan keakuratan data:**
            1. Bandingkan dengan data di [Yahoo Finance](https://finance.yahoo.com/)
            2. Untuk saham Indonesia, pastikan ticker diakhiri dengan .JK (contoh: PTBA.JK)
            3. Jika ada perbedaan signifikan, cek tanggal corporate actions
            """)
        else:
            st.error("❌ Tidak ada data yang berhasil diambil. Silakan cek koneksi atau ticker Anda")
        
        progress_bar.empty()
        status_text.empty()
