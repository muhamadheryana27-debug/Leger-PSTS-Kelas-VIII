import os
import glob
from docx import Document
from docxcompose.composer import Composer

def main():
    folder_input = "hasil_rapot_docx"  # Folder tempat file persiswa berada
    folder_output = "hasil_rapot_gabungan" # Folder untuk hasil gabungan
    
    if not os.path.exists(folder_output):
        os.makedirs(folder_output)

    # Ambil semua file docx dari folder hasil
    semua_file = glob.glob(os.path.join(folder_input, "*.docx"))
    if not semua_file:
        print(f"Tidak ada file .docx di folder {folder_input}")
        return

    # Kelompokkan file berdasarkan kelas (mengambil kata sebelum tanda underscore '_')
    kelas_dict = {}
    for f in semua_file:
        nama_file = os.path.basename(f)
        kelas = nama_file.split('_')[0]
        if kelas not in kelas_dict:
            kelas_dict[kelas] = []
        kelas_dict[kelas].append(f)

    print("=== PROSES MENGGABUNGKAN RAPORT PER KELAS ===")
    
    for kelas, daftar_file in kelas_dict.items():
        # Urutkan file berdasarkan abjad (agar urut absen)
        daftar_file.sort()
        print(f"\nMenggabungkan {len(daftar_file)} siswa untuk Kelas {kelas}...")

        # Jadikan file siswa pertama sebagai file 'Master'
        master = Document(daftar_file[0])
        composer = Composer(master)

        # Gabungkan file siswa kedua sampai terakhir ke dalam file Master
        for file_siswa in daftar_file[1:]:
            doc_lanjutan = Document(file_siswa)
            # Tambahkan pemisah halaman (Page Break) agar raport tiap siswa beda halaman
            master.add_page_break()
            composer.append(doc_lanjutan)

        # Simpan file gabungan
        nama_file_gabungan = os.path.join(folder_output, f"CETAK_SEKALIGUS_{kelas}.docx")
        composer.save(nama_file_gabungan)
        print(f"✅ Berhasil dibuat: {nama_file_gabungan}")

    print("\nSemua kelas selesai digabungkan! Silakan cek folder:", folder_output)

if __name__ == "__main__":
    main()