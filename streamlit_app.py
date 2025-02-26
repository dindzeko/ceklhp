import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

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
            horizontal_sums = [0.0] * len(table.rows)-1
            
            for row_idx, row in enumerate(table.rows):
                for col_idx, cell in enumerate(row.cells):
                    if row_idx == total_row or row_idx == len(table.rows)-1:
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
                self._set_font(cell)
                
    # Cek pola dalam teks
    for para in doc.paragraphs:
        match = re.search(r'Rp([\d.,]+)\s*\(Rp([\d.,]+)\s*\+\s*Rp([\d.,]+)\)', para.text)
        if match:
            total = float(match.group(1).replace('.', '').replace(',', '.'))
            part1 = float(match.group(2).replace('.', '').replace(',', '.'))
            part2 = float(match.group(3).replace('.', '').replace(',', '.'))
            
            if abs(total - (part1 + part2)) > 0.01:
                self._highlight_discrepancy(para, match.start(), match.end())
    
    doc.save('rekalkulasi.docx')

def _set_font(cell):
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        for run in paragraph.runs:
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11)

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

if __name__ == "__main__":
    recalculate_tables('input.docx')
