# Ringkasan Penjualan KUNUKU BABY FOOD (Per Kategori)

Script otomatisasi untuk merekap dan mengagregasi data penjualan produk mentah dari Luna POS menjadi laporan performa penjualan berbasis kategori.

## 📌 Deskripsi

Program ini mengubah laporan penjualan bulanan dari format per-produk menjadi rekapitulasi **berdasarkan 23 Kategori Produk Utama**. 

Semua data dari berbagai outlet (Hertasning, Perintis, Malang) akan digabungkan secara otomatis ke dalam satu sheet, menghilangkan rincian nama produk, dan menampilkan performa metrik inti seperti Total Qty, Diskon, Harga Jual Bersih, dan Gross Profit.

## ⚠️ Logika Bisnis & Filter Data

Script ini bertindak tidak hanya sebagai format file, melainkan memiliki validasi data yang tegas:
1. **Urutan Top 20 Terlaris:** Output akan diurutkan secara otomatis dari Kategori dengan total penjualan (Qty) terbanyak ke yang paling sedikit. Laporan ini akan membatasi dan hanya menampilkan **20 Kategori Terbaik (Top 20)**.
2. **Filter Profit (Rugi / 0% Dibuang):** Semua barang dengan Gross Profit `<= 0` (Nol atau Minus/Rugi) akan **diabaikan** secara otomatis dan tidak masuk ke dalam perhitungan total.
3. **Pengecualian MAINAN:** Produk dengan keyword "MAINAN" diabaikan.
4. **Agregasi Kategori:** Semua barang di luar 23 kategori yang sudah ditetapkan akan dibuang (contoh: barang lain-lain yang bukan fokus menu bayi/balita KUNUKU BABY FOOD).

## 🚀 Cara Penggunaan

Script ini dirancang untuk pemakaian jangka panjang dengan mendukung path file yang dinamis.

### 1. Instalasi Requirements
Pastikan Python sudah terinstal di sistem. Buka terminal/cmd dan install library `openpyxl`:
```bash
pip install openpyxl
```

### 2. Eksekusi Script
Gunakan perintah `python` diikuti nama file script, dan masukkan *path* file Excel mentahan yang diexport dari Luna POS.

**Format:**
```bash
python generate_top_bottom_sales.py <path_ke_file_excel>
```

**Contoh:**
```bash
python generate_top_bottom_sales.py "d:\MyProgram\ringkasan-penjualan-kbb\Ringkasan Penjualan Produk 01 April - 30 April 2026.xlsx"
```

### 3. Hasil Output
Script akan otomatis membuat file baru di lokasi *working directory* yang sama dengan prefix `KATEGORI_` di awal nama file. Judul bulan pada *header* Excel juga akan di-generate otomatis berdasarkan nama file yang dimasukkan.

Contoh output file:
`KATEGORI_Ringkasan Penjualan Produk 01 April - 30 April 2026.xlsx`

## 🛠 Struktur Kategori (23 List Utama)
Berikut urutan yang akan dihasilkan pada tabel akhir:
1. Bubur 9+ 100ml
2. Bubur 6+ 200ml
3. Sup
4. Bubur 6+ 100ml
5. Bubur 9+ 200ml
6. Bubur 11+ 100ml
7. Bubur 11+ 200ml
8. Lauk
9. Snack Buah
10. Rice BB Booster
11. Finger Food
12. Kaldu BB Booster
13. Abon 25ml
14. Abon 10ml
15. Rice Box
16. Pasta
17. Kremes
18. Ghee BB Booster
19. Bubur 6+ Isi 3Cup @80ml
20. Bubur 9+ Isi 3Cup @80ml
21. Bubur 11+ Isi 3Cup @80ml
22. Bubur 9+ Meal Box
23. Bubur 11+ Meal Box
