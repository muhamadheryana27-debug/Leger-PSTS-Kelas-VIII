import os
import re
import datetime
import pandas as pd
from lxml import etree
import zipfile
import io

# ── KONFIGURASI ──────────────────────────────────────────────────────────────
TEMPLATE_DOCX = "Print_Rapot.docx"
DATABASE_ODS  = "database.ods"
OUTPUT_DIR    = "hasil_rapot_docx"

BULAN_ID = {
    1:"Januari", 2:"Februari", 3:"Maret", 4:"April", 5:"Mei", 6:"Juni",
    7:"Juli", 8:"Agustus", 9:"September", 10:"Oktober", 11:"November", 12:"Desember"
}

# Mapping Paragraf Nilai
MAPEL_PARA = {
    'PAI'   : (41, 42, 43), 'PPKn'  : (47, 48, 49), 'B.IND' : (53, 54, 55),
    'B.ING' : (59, 60, 61), 'MTK'   : (65, 66, 67), 'IPA'   : (71, 72, 73),
    'IPS'   : (77, 78, 79), 'SENI'   : (83, 84, 85), 'PJOK'  : (89, 90, 91),
    'INF'   : (113, 114, 115), 'B.SUN' : (128, 129, 130),
}

P_KELAS, P_NAMA, P_NISN_NIS, P_JUMLAH = 13, 19, 23, 138
NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# ── FUNGSI MODIFIKASI XML ───────────────────────────────────────────────────

def set_para_text(paras, idx, text, bold=False, underline=False):
    """Mengisi teks ke paragraf tertentu dengan opsi Bold dan Underline."""
    if idx >= len(paras): return # Safety check
    p = paras[idx]
    for r in p.findall(f'{{{NS}}}r'): p.remove(r)
    
    new_r = etree.SubElement(p, f'{{{NS}}}r')
    
    if bold or underline:
        rPr = etree.SubElement(new_r, f'{{{NS}}}rPr')
        if bold:
            etree.SubElement(rPr, f'{{{NS}}}b')
        if underline:
            u = etree.SubElement(rPr, f'{{{NS}}}u')
            u.set(f'{{{NS}}}val', 'single')
            
    t = etree.SubElement(new_r, f'{{{NS}}}t')
    t.text = str(text)

def ganti_teks_placeholder(paras, placeholder, pengganti, bold=False, underline=False):
    """Mencari placeholder dan menggantinya dengan teks."""
    for p in paras:
        full_text = "".join([t.text for t in p.findall(f'.//{{{NS}}}t') if t.text])
        if placeholder in full_text:
            new_text = full_text.replace(placeholder, str(pengganti))
            for r in p.findall(f'{{{NS}}}r'): p.remove(r)
            
            new_r = etree.SubElement(p, f'{{{NS}}}r')
            if bold or underline:
                rPr = etree.SubElement(new_r, f'{{{NS}}}rPr')
                if bold: etree.SubElement(rPr, f'{{{NS}}}b')
                if underline:
                    u = etree.SubElement(rPr, f'{{{NS}}}u')
                    u.set(f'{{{NS}}}val', 'single')
                
            new_t = etree.SubElement(new_r, f'{{{NS}}}t')
            new_t.text = new_text

def angka_ke_huruf(n):
    satuan = ['', 'Satu', 'Dua', 'Tiga', 'Empat', 'Lima', 'Enam', 'Tujuh', 'Delapan', 'Sembilan', 'Sepuluh', 'Sebelas']
    try:
        n = int(round(float(n)))
        if n < 12: return satuan[n]
        if n < 20: return satuan[n-10] + ' Belas'
        if n < 100: return satuan[n//10] + ' Puluh ' + (satuan[n%10] if n%10 !=0 else '')
        if n == 100: return 'Seratus'
        return str(n)
    except: return '-'

# ── PROSES UTAMA ─────────────────────────────────────────────────────────────

def buat_file_rapot(template_bytes, row):
    now = datetime.datetime.now()
    tgl_cetak = f"Wanayasa, {now.day} {BULAN_ID[now.month]} {now.year}"
    
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as zin:
        xml_content = zin.read('word/document.xml')
    
    tree = etree.fromstring(xml_content)
    paras = tree.findall(f'.//{{{NS}}}p')

    # 1. ISI IDENTITAS
    set_para_text(paras, P_KELAS, f": {row.get('Kelas', '')}", bold=True)
    set_para_text(paras, P_NAMA, f": {str(row.get('Nama', '')).upper()}", bold=True)
    set_para_text(paras, P_NISN_NIS, f": {row.get('NIS/NISN', '')}", bold=True)

    # 2. ISI NILAI (DENGAN PERBAIKAN ERROR HANDLING)
    total = 0
    for col, (p_ang, p_hur, p_sta) in MAPEL_PARA.items():
        val = row.get(col, 0)
        if pd.notna(val) and val != '':
            try:
                # PERBAIKAN DISINI: Cek apakah data benar-benar bisa diubah jadi angka
                v = int(round(float(val)))
                set_para_text(paras, p_ang, str(v))
                set_para_text(paras, p_hur, angka_ke_huruf(v))
                set_para_text(paras, p_sta, "Tuntas" if v >= 71 else "Belum Tuntas")
                total += v
            except (ValueError, TypeError):
                # Jika data berupa teks (seperti 'VII A'), tampilkan apa adanya tanpa crash
                set_para_text(paras, p_ang, str(val))
                set_para_text(paras, p_hur, "-")
                set_para_text(paras, p_sta, "-")
                
    set_para_text(paras, P_JUMLAH, str(total))

    # 3. ISI WALI KELAS & NIP
    nama_w = str(row.get('Wali Kelas', '')).strip()
    nip_w  = str(row.get('NIP', '')).strip()
    
    if nip_w.endswith('.0'): nip_w = nip_w[:-2]
    nip_str = f"NIP. {nip_w}" if nip_w not in ['nan', '', 'None'] else "NIP. -"

    ganti_teks_placeholder(paras, "[TITI_MANGSA]", tgl_cetak)
    ganti_teks_placeholder(paras, "[NAMA_WALI]", nama_w, bold=True, underline=True)
    ganti_teks_placeholder(paras, "[NIP_WALI]", nip_str, bold=True)

    # Pack kembali menjadi DOCX
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as zin:
        with zipfile.ZipFile(output, 'w') as zout:
            for item in zin.infolist():
                if item.filename == 'word/document.xml':
                    zout.writestr(item, etree.tostring(tree))
                else:
                    zout.writestr(item, zin.read(item.filename))
    return output.getvalue()

def main():
    if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)
    print("Memulai proses pembuatan raport DOCX...")
    
    # Membaca database
    try:
        df = pd.read_excel(DATABASE_ODS, engine='odf', dtype={'NIP': str, 'NIS/NISN': str, 'Kelas': str})
    except Exception as e:
        print(f"Gagal membaca file {DATABASE_ODS}: {e}")
        return

    with open(TEMPLATE_DOCX, 'rb') as f:
        template_data = f.read()

    for i, row in df.iterrows():
        if pd.isna(row['Nama']): continue
        # Bersihkan nama file agar tidak ada karakter ilegal
        nama_bersih = str(row['Nama']).upper().replace("/", "-")
        kelas_bersih = str(row['Kelas']).replace(' ','')
        nama_file = f"{kelas_bersih}_{nama_bersih}.docx"
        
        print(f"Mencetak: {nama_file}")
        
        try:
            hasil = buat_file_rapot(template_data, row)
            with open(os.path.join(OUTPUT_DIR, nama_file), 'wb') as f:
                f.write(hasil)
        except Exception as e:
            print(f"Gagal mencetak rapot {nama_file}: {e}")

    print(f"\nSelesai! Cek folder: {OUTPUT_DIR}")

if __name__ == "__main__":
    main()