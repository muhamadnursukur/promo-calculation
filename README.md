# Promo Calculation - Excel Based

Script Node.js ini digunakan untuk membaca aturan promo dari file Excel dan memproses data transaksi penjualan (`sell out`) untuk menghitung:
- **Bonus Barang**
- **Cashback**
- **Promo Gabungan berdasarkan region dan produk tertentu**

Hasilnya akan disimpan ke file baru bernama `hasil_perhitungan_promo.xlsx`.

## 📁 Struktur File

- `aturan_promo.xlsx` - berisi aturan promo (bonus & cashback) berdasarkan produk dan area.
- `data_sell_out.xlsx` - data transaksi penjualan yang akan diproses.
- `hasil_perhitungan_promo.xlsx` - output hasil kalkulasi promo.

## 🚀 Cara Pakai

1. **Install dependency:**
```bash
npm install xlsx
```

2. **Run Program:**
```bash
node index.js
