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
```

---

## ▶️ Cara Menjalankan

```bash
node hitungPromo.js
```

Setelah selesai, file `hasil_perhitungan_promo.xlsx` akan muncul di folder yang sama.

---

## 📊 Format File Excel

### 🔹 `aturan_promo.xlsx`

| Area | Nama produk | Tipe Promo | Min | Max | Value | Produk Gabungan | Kelompok Gabungan | Wajib Gabungan |
|------|-------------|------------|-----|-----|--------|------------------|--------------------|----------------|
| Jawa | Produk A    | Bonus Barang | 10  | 20  | 1     | Tidak            |                    |                |

**Keterangan:**
- `Nama produk`: Bisa satu atau lebih nama, dipisahkan dengan `|`.
- `Tipe Promo`: Salah satu dari:
  - `Bonus Barang`
  - `Cashback`
  - `Gabungan Kuantitas`
  - `Gabungan Value`
- `Produk Gabungan`: `Ya` jika bagian dari promo gabungan.
- `Kelompok Gabungan`: Nama kelompok untuk produk gabungan.
- `Wajib Gabungan`: `Ya` jika semua produk di kelompok harus muncul dalam transaksi.

---

### 🔹 `data_sell_out.xlsx`

| Area | Nama Produk | Jumlah Karton | Value Netto | Nota | Nama_toko | ... |
|------|--------------|----------------|--------------|------|------------|-----|

**Kolom penting:**
- `Jumlah Karton`: Kuantitas produk yang dibeli.
- `Value Netto`: Nilai transaksi produk.
- `Nota` dan `Nama_toko`: Digunakan untuk mengelompokkan transaksi per nota.

---

## ✅ Output: `hasil_perhitungan_promo.xlsx`

File hasil berisi per baris transaksi dengan tambahan kolom:

| ... | totalBonusBarang | totalCashback | LayerBonus | LayerCashback | PromoGabunganQty | PromoGabunganValue | LayerGabunganQty | KelipatanGabunganQty |
|-----|------------------|---------------|-------------|----------------|-------------------|---------------------|-------------------|-----------------------|

---

## 🧠 Logika Utama

- **Bonus Barang**: Dihitung berdasarkan kelipatan minimal kuantitas.
- **Cashback**: Berlaku jika jumlah karton dalam range `min` sampai `max`.
- **Gabungan Kuantitas**: Jumlah karton dari semua produk dalam satu kelompok.
- **Gabungan Value**: Total `Value Netto` dari semua produk dalam kelompok.
- **Distribusi Bonus Gabungan**: Nilai bonus dibagi berdasarkan kontribusi value produk dalam kelompok.

---

## 🧑‍💻 Contoh Output

```bash
✅ File hasil_perhitungan_promo.xlsx berhasil dibuat.
Proses Perhitungan Promo: 328.12ms
```

---

## ✍️ Penulis

Muhamad Nur Sukur

---

## 📄 Lisensi

MIT License
