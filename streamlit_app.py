import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
import re
import io

# Judul aplikasi
st.title("📝 Aplikasi Rekalkulasi Tabel Dokumen Word")
st.write("Upload dokumen Word (.docx) untuk merekalkulasi tabel dan memeriksa angka dalam tanda kurung.")

# Fungsi rekalkulasi teks dalam paragraf
def recalculate_text(doc_path):
    doc = Document(doc_path)
    
    for para in doc.paragraphs:
        # Cari pola dalam teks: Rp<nilai> (<operasi>)
        match = re.search(r'Rp([\d.,]+)\s*\(([\d.,]+\s*[-+*/:]\s*[\d.,]+(?:\s*[-+*/:]\s*[\d.,]+)*)\)', para.text)
        if match:
            original_total = float(match.group(1).replace('.', '').replace(',', '.'))  # Nilai sebelum tanda kurung
            operation = match.group(2)  # Operasi di dalam tanda kurung
            
            # Validasi dan ubah operator ":" menjadi "/"
            operation = operation.replace(':', '/')  # Ganti ":" dengan "/"
            
            # Ekstrak semua angka dan operator dari operasi
            try:
                # Ubah format angka ke float dan evaluasi operasi
                cleaned_operation = re.sub(r'(\d+\.?\d*)', lambda x: str(float(x.group().replace('.', '').replace(',', '.'))), operation)
                
                # Evaluasi operasi menggunakan eval()
                recalculated_total = eval(cleaned_operation)
                
                # Bandingkan hasil rekalkulasi dengan nilai asli
                if abs(original_total - recalculated_total) > 0.01:
                    # Tambahkan hasil rekalkulasi setelah tanda kurung
                    recalculated_text = f" = Rp{recalculated_total:,.2f}".replace(',', 'temp').replace('.', ',').replace('temp', '.')
                    para.text = para.text[:match.end()] + recalculated_text
                    _highlight_discrepancy(para, match.end(), len(para.text))
            except Exception as e:
                st.warning(f"Gagal memproses operasi '{operation}': {str(e)}")
    
    return doc

# Fungsi untuk menyorot hasil rekalkulasi dengan font merah
def _highlight_discrepancy(para, start, end):
    text = para.text
    para.clear()
    
    # Tambahkan teks sebelum hasil rekalkulasi
    if start > 0:
        para.add_run(text[:start])
    
    # Tambahkan hasil rekalkulasi dengan highlight
    discrepancy_run = para.add_run(text[start:end])
    discrepancy_run.font.color.rgb = RGBColor(255, 0, 0)
    discrepancy_run.font.name = 'Times New Roman'
    discrepancy_run.font.size = Pt(11)
    discrepancy_run.underline = True
    
    # Tambahkan teks setelah hasil rekalkulasi
    if end < len(text):
        para.add_run(text[end:])

# Upload file
uploaded_file = st.file_uploader("Upload File Word (.docx)", type=["docx"])

if uploaded_file:
    try:
        # Simpan file upload sementara
        with open("input.docx", "wb") as f:
            f.write(uploaded_file.getvalue())
        
        # Proses rekalkulasi teks
        processed_doc = recalculate_text("input.docx")
        
        # Simpan hasil rekalkulasi ke BytesIO
        output = io.BytesIO()
        processed_doc.save(output)
        output.seek(0)
        
        # Tampilkan pesan sukses
        st.success("Rekalkulasi selesai! Silakan unduh hasil di bawah.")
        
        # Tombol unduh
        st.download_button(
            label="📥 Unduh Hasil Rekalkulasi",
            data=output,
            file_name="rekalkulasi.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"Terjadi kesalahan: {str(e)}")
