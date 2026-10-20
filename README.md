# Word to PDF Batch Converter

Aplikasi berbasis Python untuk mengonversi seluruh dokumen Microsoft Word (`.docx` / `.doc`) dalam sebuah folder menjadi format PDF secara otomatis. Skrip ini dilengkapi dengan penanganan error, indikator *progress*, serta proteksi penimpaan file (*file overwrite protection*) dengan penomoran unik otomatis `(1)`, `(2)`, dan seterusnya.

---

## 📋 Prasyarat Sistem

Sebelum menjalankan program ini, pastikan sistem Anda memenuhi persyaratan berikut:

1. **Sistem Operasi**: Windows atau macOS.
2. **Microsoft Word**: Harus sudah terinstal di komputer (diperlukan oleh pustaka `docx2pdf` untuk melakukan konversi secara presisi).
3. **Python**: Versi 3.7 atau lebih baru.

---

## 🛠️ Langkah-Langkah Instalasi & Penggunaan

### Langkah 1: Kloning atau Unduh Repositori
Simpan file `word_to_pdf.py` di komputer Anda dalam folder kerja khusus.

### Langkah 2: Install Depedensi/Pustaka
Buka Terminal / Command Prompt (CMD), lalu jalankan perintah berikut untuk menginstal pustaka yang dibutuhkan:

```bash
pip install docx2pdf tqdm
```

### Langkah 3: Siapkan Folder Dokumen
Buat folder tempat menyimpan file Word Anda. Secara bawaan (default), program mencari:
- Folder Input: `./dokumen_word`
- Folder Output: `./dokumen_pdf`

*(Anda bisa membuat folder `dokumen_word` di direktori yang sama dengan skrip dan memasukkan file `.docx` ke dalamnya)*.

### Langkah 4: Konfigurasi Path (Opsional)
Jika ingin menggunakan lokasi folder lain, buka file `word_to_pdf.py` dan sesuaikan variabel di bagian paling bawah:

```python
if __name__ == "__main__":
    FOLDER_INPUT = "C:/path/ke/folder/word_anda"
    FOLDER_OUTPUT = "C:/path/ke/folder/pdf_hasil"

    konversi_folder_word_ke_pdf(FOLDER_INPUT, FOLDER_OUTPUT)
```

### Langkah 5: Jalankan Skrip
Jalankan skrip Python melalui terminal:

```bash
python word_to_pdf.py
```

---

## ⚙️ Fitur Utama

1. **Auto Unik Naming**: Jika file PDF hasil konversi sudah ada di folder tujuan, sistem tidak akan menimpanya (*overwrite*), melainkan otomatis menambahkan penomoran bertahap `(1)`, `(2)`, dst.
2. **Error Handling**: Jika ada dokumen Word yang rusak (*corrupt*) atau gagal diproses, program tidak akan terhenti/crash melainkan akan melanjutkan konversi file lainnya dan menampilkan pesan error.
3. **Filter File Sementara**: Secara otomatis mengabaikan file sementara Word (file yang berawalan `~$`).
4. **Visual Progress**: Menampilkan *progress bar* interaktif saat konversi berlangsung.