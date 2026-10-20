import os
from docx2pdf import convert
from tqdm import tqdm

def dapatkan_nama_pdf_unik(path_output_awal):
    """
    Memeriksa keberadaan file PDF di direktori tujuan.
    Jika file dengan nama yang sama sudah ada, fungsi ini secara iteratif 
    menambahkan akhiran penomoran seperti (1), (2), dan seterusnya hingga 
    ditemukan nama file yang unik untuk menghindari penimpaan file (overwrite).

    Parameters:
    -----------
    path_output_awal : str
        Path lengkap tujuan awal untuk file PDF.

    Returns:
    --------
    str
        Path lengkap tujuan file PDF yang sudah dipastikan unik.
    """
    # Jika file belum ada, gunakan path asli
    if not os.path.exists(path_output_awal):
        return path_output_awal

    # Pisahkan folder dan nama file
    folder, nama_file = os.path.split(path_output_awal)
    nama_dasar, ekstensi = os.path.splitext(nama_file)
    
    # Loop penomoran hingga menemukan nama file yang belum terpakai
    counter = 1
    while True:
        nama_baru = f"{nama_dasar} ({counter}){ekstensi}"
        path_baru = os.path.join(folder, nama_baru)
        if not os.path.exists(path_baru):
            return path_baru
        counter += 1

def konversi_folder_word_ke_pdf(folder_sumber, folder_tujuan=None):
    """
    Mengonversi semua dokumen Word (.docx dan .doc) dari folder sumber ke format PDF
    di folder tujuan dengan penanganan error individual dan otomatisasi penomoran file ganda.

    Parameters:
    -----------
    folder_sumber : str
        Path direktori yang berisi dokumen Word yang akan dikonversi.
    folder_tujuan : str, optional
        Path direktori tempat menyimpan hasil konversi PDF. Jika None, 
        akan disimpan di dalam folder_sumber.
    """
    # 1. Validasi keberadaan folder sumber
    if not os.path.exists(folder_sumber):
        print(f"[ERROR] Folder sumber '{folder_sumber}' tidak ditemukan.")
        return

    # Tentukan folder tujuan (buat folder baru jika belum ada)
    if folder_tujuan is None:
        folder_tujuan = folder_sumber
    else:
        os.makedirs(folder_tujuan, exist_ok=True)

    # 2. Filter file berekstensi .docx / .doc (mengabaikan file temporary Word '~$')
    semua_file = os.listdir(folder_sumber)
    file_word = [
        f for f in semua_file 
        if (f.lower().endswith('.docx') or f.lower().endswith('.doc')) 
        and not f.startswith('~$')
    ]

    # Hentikan proses jika tidak ditemukan file Word
    if not file_word:
        print("[INFO] Tidak ditemukan file Word (.docx / .doc) yang valid untuk dikonversi.")
        return

    print(f"[INFO] Ditemukan {len(file_word)} file Word. Memulai proses konversi...\n")

    # Inisialisasi penghitung statistik
    berhasil = 0
    gagal = 0

    # 3. Proses konversi file satu per satu dengan indikator kemajuan (tqdm)
    for nama_file in tqdm(file_word, desc="Progress Konversi", unit="file"):
        path_input = os.path.join(folder_sumber, nama_file)
        
        # Susun nama dasar PDF
        nama_pdf_default = os.path.splitext(nama_file)[0] + ".pdf"
        path_output_ideal = os.path.join(folder_tujuan, nama_pdf_default)

        # Cek dan dapatkan nama file unik jika terjadi duplikasi nama
        path_output_final = dapatkan_nama_pdf_unik(path_output_ideal)

        try:
            # Mengonversi file menggunakan mesin MS Word via docx2pdf
            convert(path_input, path_output_final)
            berhasil += 1
        except Exception as e:
            gagal += 1
            print(f"\n[ERROR] Gagal mengonversi '{nama_file}': {e}")

    # 4. Ringkasan laporan hasil konversi
    print("\n" + "="*40)
    print("PROSES KONVERSI SELESAI")
    print(f"• Berhasil : {berhasil} file")
    print(f"• Gagal    : {gagal} file")
    print("="*40)

if __name__ == "__main__":
    # Tentukan path direktori input dan output
    FOLDER_INPUT = "./dokumen_word"
    FOLDER_OUTPUT = "./dokumen_pdf"

    # Jalankan fungsi utama
    konversi_folder_word_ke_pdf(FOLDER_INPUT, FOLDER_OUTPUT)