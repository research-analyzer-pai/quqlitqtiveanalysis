Q-ANALYST — COMPLETE LOCAL-FIRST STARTER
=======================================

Aplikasi pribadi berbasis Streamlit untuk Windows. Mendukung:
- proyek penelitian dengan database SQLite lokal yang persisten;
- impor beberapa file Excel sekaligus, pemilihan sheet dan pemetaan kolom;
- penggabungan unit wawancara dari banyak informan;
- AI-assisted open coding, kategori, tema, dan memo melalui OpenAI API;
- validasi/edit hasil coding oleh peneliti;
- catatan observasi;
- unggah foto dokumentasi dan metadata;
- matriks lintas informan;
- matriks triangulasi yang tersimpan;
- ekspor Excel dan Word;
- audit trail aktivitas dasar.

INSTALASI WINDOWS
1. Instal Python 3.11+ dari https://www.python.org/downloads/ dan centang Add Python to PATH.
2. Ekstrak ZIP ke folder tetap, misalnya C:\Q_ANALYST.
3. Buka CMD di folder itu.
4. Jalankan:
   py -m venv .venv
   .venv\Scripts\activate
   py -m pip install --upgrade pip
   pip install -r requirements.txt
   streamlit run app.py
5. Browser membuka http://localhost:8501 (alamat yang ditampilkan terminal).

PENGGUNAAN
- Buat/simpan proyek di Dashboard.
- Unggah banyak Excel pada tab Impor Multi-Excel.
- Untuk setiap file, pilih sheet dan kolom informan, pertanyaan, jawaban.
- Klik Simpan dan gabungkan.
- Pada AI Coding, masukkan API key sendiri lalu jalankan coding untuk data yang belum memiliki kode.
- Periksa dan validasi semua hasil.
- Tambahkan observasi dan foto.
- Isi matriks triangulasi dan ekspor.

KEAMANAN DAN KETERBATASAN
- Database SQLite dan foto disimpan di folder aplikasi lokal. Backup folder tersebut.
- AI memerlukan API key dan internet. Teks kutipan yang diproses dikirim ke penyedia AI eksternal; penggunaan API dapat dikenai biaya. Samarkan identitas informan terlebih dahulu.
- API key hanya digunakan dari input sesi dan tidak sengaja ditulis ke database.
- Tidak ada AI visual/OCR foto pada versi ini. Foto hanya disimpan beserta metadata.
- Coding AI merupakan usulan, bukan keputusan final. Peneliti wajib memvalidasi dan mencatat alasan perubahan.
- Format Excel berbeda-beda; periksa pemetaan kolom setiap file.
- Satu baris dianggap satu unit jawaban. Pecah menjadi unit makna bila jawaban berisi beberapa gagasan berbeda.
- Aplikasi ini adalah starter lokal, belum diaudit keamanan dan belum layak untuk hosting publik atau data sangat sensitif.
- Untuk backup: tutup aplikasi, salin seluruh folder Q_ANALYST beserta q_analyst.db dan uploads ke media aman.

Jika ada masalah instalasi, kirimkan teks pesan error dari CMD agar dapat diperbaiki.
