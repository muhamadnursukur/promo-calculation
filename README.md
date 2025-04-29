# 📦 Promo Calculation Engine

Skrip ini digunakan untuk menghitung bonus barang dan cashback berdasarkan data penjualan (sell out) dan aturan promo yang didefinisikan dalam file Excel. Sistem ini juga mendukung perhitungan **promo gabungan kuantitas** dan **promo gabungan value** berdasarkan kelompok produk.

---

## 📁 Struktur File

- `aturan_promo.xlsx`  
  Berisi daftar aturan promo, termasuk jenis promo, nilai bonus, batas minimum dan maksimum, serta informasi kelompok produk jika promo bersifat gabungan.

- `data_sell_out.xlsx`  
  Berisi data transaksi penjualan yang akan dihitung promonya, termasuk informasi toko, produk, jumlah, dan nilai transaksi.

- `hasil_perhitungan_promo.xlsx`  
  File output hasil perhitungan promo berdasarkan aturan dan data penjualan.

---

## ⚙️ Cara Kerja

1. **Baca Aturan Promo**
   - Skrip membaca dan mengelompokkan aturan berdasarkan produk dan area.
   - Mendukung banyak tipe promo:
     - Bonus Barang
     - Cashback
     - Gabungan Kuantitas
     - Gabungan Value
   - Mendukung penggabungan produk dalam satu kelompok.

2. **Proses Transaksi**
   - Data penjualan dikelompokkan per nota dan toko.
   - Sistem menghitung bonus/cashback langsung per produk.
   - Jika produk bagian dari promo gabungan, sistem akan cek kelengkapan produk dalam kelompok untuk hitung bonus gabungan (kuantitas dan value).

3. **Hasil Perhitungan**
   - Hasil akhir ditulis ke file `hasil_perhitungan_promo.xlsx`, lengkap dengan rincian layer dan kelipatan promo.

---

## 🛠️ Instalasi

1. Pastikan Node.js sudah terpasang.
2. Instal dependensi:

```bash
npm install xlsx
