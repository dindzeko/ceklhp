import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import re
import io

# Judul aplikasi
st.title("📝 Aplikasi Rekalkulasi Dokumen Word")
st.write("Upload dokumen Word (.docx) untuk merekalkulasi tabel dan teks.")

# Fungsi rekalkulasi tabel
def recalculate_tables(doc):
    for table in doc.tables:
        # Inisialisasi variabel untuk rekalkulasi
        num_cols = len(table.columns)
        vertical_sums = [0.0] * num_cols  # Untuk menjumlahkan kolom angka
        
        # Proses setiap baris dalam tabel
        for row_idx, row in enumerate(table.rows):
            # Deteksi baris "Jumlah" atau "Total"
            is_total_row = any(
                keyword in cell.text.upper() for keyword in ["JUMLAH", "TOTAL"] for cell in row.cells
            )
            
            for col_idx, cell in enumerate(row.cells):
                value = cell.text.strip()
                
                # Coba ekstrak angka dari sel (termasuk angka negatif dalam tanda kurung)
                cleaned_value = value.replace('.', '').replace(',', '.')  # Bersihkan format angka
                
                # Cek apakah angka dalam tanda kurung (misalnya: "(45)")
                if re.match(r'^\((\d+\.?\d*)\)$', cleaned_value):  # Contoh: "(45)"
                    num = -float(cleaned_value[1:-1])  # Ubah ke negatif
                elif re.match(r'^-?\d+\.?\d*$', cleaned_value):  # Contoh: "-45" atau "45"
                    num = float(cleaned_value)
                else:
                    continue  # Bukan angka, lewati
                
                # Abaikan baris "Jumlah" atau "Total" dalam perhitungan
                if not is_total_row:
                    vertical_sums[col_idx] += num  # Tambahkan ke total kolom
        
        # Tambahkan baris "Rekalkulasi"
        new_row = table.add_row()
        new_row.cells[0].text = "Rekalkulasi"
        
        for col_idx in range(num_cols):
            cell = new_row.cells[col_idx]
            if col_idx >= 1:  # Kolom angka dimulai dari kolom ke-2 (indeks 1)
                cell.text = f"{vertical_sums[col_idx]:,.2f}".replace(',', 'temp').replace('.', ',').replace('temp', '.')
            _set_font(cell)

# Fungsi rekalkulasi teks dalam paragraf
def recalculate_text(doc):
    for para in doc.paragraphs:
        # Cari pola dalam teks: Rp<nilai>(<operasi>)
        match = re.search(r'Rp([\d.,]+)\.?\s*\(([^)]+)\)', para.text)
        if match:
            original_total_str = match.group(1).replace('.', '').replace(',', '.')
            original_total = float(original_total_str)
            operation = match.group(2)

            # Validasi input kosong
            if not operation.strip():
                st.warning(f"Operasi kosong ditemukan: '{para.text}'")
                continue

            # Membersihkan operator yang tidak standar
            operation = operation.replace(':', '/').replace('x', '*').replace('–', '-').replace('−', '-')
            
            # Hapus teks non-matematika (misalnya, "hari")
            operation = re.sub(r'[a-zA-Z\s]+', '', operation)

            # Proses semua nilai Rp dalam operasi
            operation_clean_rp = re.sub(
                r'Rp\s*([\d.,]+)', 
                lambda m: m.group(1).replace('.', '').replace(',', '.'), 
                operation
            )

            # Proses angka dengan pemisah ribuan (.) dan desimal (,)
            operation_clean = re.sub(
                r'(\d{1,3}(?:\.\d{3})*(?:,\d+)?)', 
                lambda m: m.group().replace('.', '').replace(',', '.'), 
                operation_clean_rp
            )

            try:
                # Evaluasi operasi menggunakan eval()
                recalculated_total = eval(operation_clean)
                recalculated_total = round(recalculated_total, 2)
                original_total = round(original_total, 2)

                # Bandingkan hasil rekalkulasi dengan nilai asli
                if original_total != recalculated_total:
                    recalculated_text = f" = Rp{recalculated_total:,.2f}".replace(',', 'temp').replace('.', ',').replace('temp', '.')
                    para.text = para.text[:match.end()] + recalculated_text
                    _highlight_discrepancy(para, match.end(), len(para.text))
            except Exception as e:
                st.warning(f"Gagal memproses operasi '{operation}': {str(e)}")

# Fungsi untuk mengatur font
def _set_font(cell):
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        for run in paragraph.runs:
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11)

# Fungsi untuk menyorot ketidaksesuaian
def _highlight_discrepancy(para, start, end):
    text = para.text
    para.clear()
    
    # Tambahkan teks sebelum discrepancy
    if start > 0:
        para.add_run(text[:start])
    
    # Tambahkan discrepancy dengan highlight
    discrepancy_run = para.add_run(text[start:end])
    discrepancy_run.font.color.rgb = RGBColor(255, 0, 0)
    discrepancy_run.font.name = 'Times New Roman'
    discrepancy_run.font.size = Pt(11)
    discrepancy_run.underline = True
    
    # Tambahkan teks setelah discrepancy
    if end < len(text):
        para.add_run(text[end:])

# Upload file
uploaded_file = st.file_uploader("Upload File Word (.docx)", type=["docx"])
if uploaded_file:
    try:
        # Simpan file upload sementara
        with open("input.docx", "wb") as f:
            f.write(uploaded_file.getvalue())
        
        # Baca dokumen Word
        doc = Document("input.docx")
        
        # Proses rekalkulasi tabel
        recalculate_tables(doc)
        
        # Proses rekalkulasi teks
        recalculate_text(doc)
        
        # Simpan hasil rekalkulasi ke BytesIO
        output = io.BytesIO()
        doc.save(output)
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
