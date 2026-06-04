# GoldCart Pro - Sistem POS Toko Emas 🏆

GoldCart Pro adalah sistem *Point of Sales* (Kasir) canggih yang dirancang khusus untuk toko emas dan perhiasan, dibangun menggunakan **Google Apps Script** dan **HTML5/JS/CSS murni** tanpa *framework* berat.

## 🌟 Fitur Utama
1. **Manajemen Transaksi Jual & Beli:** Pencatatan transaksi lengkap dengan kalkulasi harga berdasarkan kadar emas dan upah/potongan.
2. **Kalkulasi Stok Otomatis:** Perhitungan arus gramasi otomatis (Saldo Awal + Jual - Beli).
3. **Laporan & Paginasi Cerdas:** Tabel laporan transaksi dengan pembagian halaman (paginasi) otomatis dan filter tanggal untuk tutup buku.
4. **CRM Pelanggan Terintegrasi:** Melacak data dan riwayat total belanja dari masing-masing pelanggan.
5. **Cetak Nota & Integrasi WhatsApp:** Fitur mencetak struk fisik atau menyalin ringkasan nota digital yang langsung mengarah ke WhatsApp.
6. **Live Harga Emas Dunia (XAU/USD):** Pemantauan harga emas global *real-time* dengan IFrame TradingView yang anti-blokir.
7. **Tema Kustom Premium:** Tiga pilihan tema eksklusif: *Midnight Gold*, *Pearl White*, dan *Royal Sapphire* yang tersimpan otomatis (persisten).
8. **Serverless (Google Sheets Backend):** Semua data diamankan dan disimpan ke lembar kerja Google Sheets, 100% gratis tanpa biaya *hosting backend*.

## 🛠️ Cara Penggunaan (Deploy)
1. Buat **Google Sheet** baru di akun Google Anda (buat tab bernama `transaksi`, `settings`, `pelanggan`, `pengeluaran`, `harga_emas`).
2. Masuk ke menu **Ekstensi > Apps Script**.
3. *Copy* dan *paste* isi dari file `gas.gs` ke dalam `Code.gs`.
4. Buat file HTML baru di Apps Script dengan nama `index.html`, lalu *paste* seluruh kode dari file `index.html` repositori ini.
5. Klik **Deploy > New deployment**, pilih tipe **Web App**, set *Who has access* ke **Anyone**.
6. Akses link Web App yang diberikan, dan aplikasi kasir Anda siap digunakan!
