import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import re
import io

# Judul aplikasi
st.title("📝 Aplikasi Rekalkulasi Tabel Dokumen Word")
st.write("Upload dokumen Word (.docx) untuk merekalkulasi tabel dan memeriksa angka dalam tanda kurung.")

# Fungsi rekalkulasi tabel
def recalculate_tables(doc_path):
    doc = Document(doc_path)
    
    for table in doc.tables:
        # Cari baris JUMLAH
        total_row = None
        for i, row in enumerate(table.rows):
            if 'JUMLAH' in row.cells[0].text:
                total_row = i
                break
        
        if total_row is not None:
            # Tambahkan baris Rekalkulasi
            new_row = table.add_row()
            new_row.cells[0].text = "Rekalkulasi"
            
            # Hitung ulang vertical dan horizontal
            num_cols = len(table.columns)
            vertical_sums = [0.0] * num_cols
            horizontal_sums = [0.0] * (len(table.rows) - 1)
            
            for row_idx, row in enumerate(table.rows):
                for col_idx, cell in enumerate(row.cells):
                    if row_idx == total_row or row_idx == len(table.rows) - 1:
                        continue
                    
                    # Ekstrak angka dari kolom nilai
                    if col_idx >= 2:  # Asumsi kolom angka mulai dari kolom ke-3
                        value = cell.text.replace('.', '').replace(',', '.')
                        if re.match(r'^\d+\.?\d*$', value):
                            num = float(value)
                            vertical_sums[col_idx] += num
                            horizontal_sums[row_idx] += num
            
            # Format baris Rekalkulasi
            for col_idx in range(num_cols):
                cell = new_row.cells[col_idx]
                if col_idx >= 2:
                    cell.text = f"{vertical_sums[col_idx]:,.2f}".replace(',', 'temp').replace('.', ',').replace('temp', '.')
                _set_font(cell)
                
    # Cek pola dalam teks
    for para in doc.paragraphs:
        match = re.search(r'Rp([\d.,]+)\s*\(Rp([\d.,]+)\s*\+\s*Rp([\d.,]+)\)', para.text)
        if match:
            total = float(match.group(1).replace('.', '').replace(',', '.'))
            part1 = float(match.group(2).replace('.', '').replace(',', '.'))
            part2 = float(match.group(3).replace('.', '').replace(',', '.'))
            
            if abs(total - (part1 + part2)) > 0.01:
                _highlight_discrepancy(para, match.start(), match.end())
    
    return doc

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
        
        # Proses rekalkulasi
        processed_doc = recalculate_tables("input.docx")
        
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
