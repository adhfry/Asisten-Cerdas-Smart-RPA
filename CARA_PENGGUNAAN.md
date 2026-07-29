# Cara Menjalankan Bot Otomasi PCare BPJS

Panduan ini untuk menjalankan `app.py` — bot Selenium yang mengotomasi Pendaftaran Pasien & Pelayanan Pasien di PCare BPJS Kesehatan, lengkap dengan auto-solve Cloudflare Turnstile (EzSolver).

## 1. Persiapan Awal (cukup sekali di awal)

1. **Isi akun login PCare** di file `data_user.xlsx`:
   - Sel `B1` = Username PCare
   - Sel `B2` = Password PCare
2. **Pastikan printer sudah terpasang** (jika ingin fitur auto-print SPP/FKPP). Driver printer Epson L120 tersedia di `DriverPrinterEpson.L120.x64.NesabaMedia.exe`.
3. Cek pengaturan printer di `app.py` (baris ±28):
   ```python
   PRINTER_NAME = "EPSON L120 Series"  # sesuaikan dengan nama printer di Windows Anda
   ```
   Ganti sesuai nama printer yang muncul di **Settings > Printers & Scanners** Windows, atau isi `"Microsoft Print to PDF"` jika hanya ingin uji coba tanpa cetak fisik.

## 2. Siapkan Data Pasien

1. Buka folder **`FILE EXCEL DISINI`** di root proyek.
2. Taruh file Excel data pasien Anda di sana (format `.xlsx`).
3. Saat bot dijalankan, Anda akan diminta memilih file dan sheet mana yang akan diproses lewat menu interaktif.

## 3. Menjalankan Bot

Buka terminal di folder proyek (`d:\Project\Selenium`), lalu jalankan lewat Python virtual environment:

```powershell
env\Scripts\python.exe app.py
```

Alur yang akan terjadi:

1. **EzSolver service otomatis dinyalakan** di background (untuk auto-solve Cloudflare Turnstile di halaman login & saat pencarian pasien). Tunggu sampai muncul pesan `-> EzSolver service siap.`
2. Browser Chrome terbuka otomatis dan masuk ke halaman login PCare.
3. Pilih file Excel & sheet data pasien lewat menu (panah Atas/Bawah + Enter).
4. Pilih mode operasi:
   - **1. Pendaftaran Pasien**
   - **2. Pelayanan Pasien (Input Hasil)**
5. Username & password terisi otomatis. Cloudflare Turnstile akan diselesaikan otomatis oleh EzSolver (muncul log `-> Widget Cloudflare Turnstile terdeteksi... terselesaikan`).
6. **Captcha teks (5-6 karakter) tetap harus Anda ketik manual** di browser — ini bukan bagian yang diotomasi. Begitu terisi, bot otomatis klik "Sign In".
7. Bot akan memproses data pasien baris demi baris sesuai file Excel, sambil menampilkan progres di terminal.
8. Beberapa titik akan meminta konfirmasi Anda di terminal (misalnya jika nama pasien berbeda, data lab tidak lengkap, atau ingin lanjut ke pasien berikutnya). Ketik `y`/Enter untuk lanjut, `n` untuk berhenti, atau `a` untuk "lanjut semua tanpa tanya lagi".
9. Setelah semua baris selesai diproses, tekan **ENTER** di terminal untuk menutup browser dan mengakhiri program.

## 4. Status di Excel

Progres tiap baris ditulis otomatis ke file Excel input Anda, antara lain:
- `SUKSES` — berhasil diproses
- `FINISH` — pelayanan sudah lengkap & selesai dicetak (baris ini akan diwarnai hijau)
- `NOT COMPLETE` / `CANNOT INPUT` / `CANNOT BE INPUT` — ada kendala, cek kolom keterangan di sebelahnya untuk detail

Bot otomatis **melewati baris yang sudah `FINISH`**, jadi aman dijalankan ulang tanpa mengulang pasien yang sudah selesai.

## 5. Troubleshooting

| Masalah | Solusi |
|---|---|
| `⚠️ Folder ezsolver/ tidak ditemukan` | Pastikan folder `ezsolver/` (berisi `solver.py`, `service.py`, `clientsend.py`) ada di root proyek. |
| `⚠️ EzSolver service belum merespons /health` | Tunggu beberapa detik lebih (Chrome untuk solver butuh waktu start). Jika terus gagal, cek apakah Google Chrome terinstall dan port `8191` tidak dipakai proses lain. |
| `⚠️ Gagal solve Turnstile` | Biasanya sementara/jaringan lambat — bot tetap lanjut mencoba klik seperti biasa; kalau macet, selesaikan Turnstile manual di browser lalu lanjutkan. |
| Sesi PCare berakhir di tengah proses | Bot akan menampilkan peringatan dan menunggu Anda login ulang manual di browser, lalu tekan ENTER untuk melanjutkan. |
| Excel tidak ditemukan di folder | Pastikan file `.xlsx` (bukan `.xls`) ada di dalam folder `FILE EXCEL DISINI`. |

## 6. Menghentikan Bot

Tekan `Ctrl+C` di terminal kapan saja untuk memaksa berhenti, atau ketik `n` saat diminta konfirmasi lanjut/stop. Data yang sudah tersimpan ke Excel tidak akan hilang.
