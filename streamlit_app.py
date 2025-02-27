import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import re
import io

st.title("📝 Aplikasi Rekalkulasi Dokumen Word")
st.write("Upload dokumen Word (.docx) untuk merekalkulasi tabel.")

def recalculate_tables(doc):
    for table in doc.tables:
        num_cols = len(table.columns)
        vertical_sums = [0.0] * num_cols
        
        for row in table.rows:
            # Skip baris yang mengandung "JUMLAH" atau "TOTAL"
            if any(keyword in cell.text.upper() for keyword in ["JUMLAH", "TOTAL"] for cell in row.cells):
                continue
            
            for col_idx, cell in enumerate(row.cells):
                value = cell.text.strip()
                cleaned_value = value.replace('.', '').replace(',', '.')
                
                # Validasi format angka
                if re.match(r'^-?\d+\.?\d*$', cleaned_value):
                    num = float(cleaned_value)
                elif re.match(r'^\(\d+\.?\d*\)$', cleaned_value):  # Format (12345)
                    num = -float(cleaned_value[1:-1])
                else:
                    continue  # Lewati jika bukan angka valid
                
                vertical_sums[col_idx] += num
        
        # Tambahkan baris rekalkulasi
        new_row = table.add_row()
        new_row.cells[0].text = "Rekalkulasi"
        
        for col_idx in range(num_cols):
            cell = new_row.cells[col_idx]
            if col_idx > 0:
                formatted_num = f"{vertical_sums[col_idx]:,.2f}".replace(',', 'temp').replace('.', ',').replace('temp', '.')
                cell.text = formatted_num
            _set_font(cell)

def _set_font(cell):
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        for run in paragraph.runs:
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11)

uploaded_file = st.file_uploader("Upload File Word (.docx)", type=["docx"])
if uploaded_file:
    try:
        # Proses file
        doc = Document(uploaded_file)
        recalculate_tables(doc)
        
        # Simpan ke BytesIO
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        
        st.success("Rekalkulasi tabel selesai!")
        st.download_button(
            label="📥 Unduh Hasil",
            data=output,
            file_name="rekalkulasi.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"Terjadi kesalahan: {str(e)}")
