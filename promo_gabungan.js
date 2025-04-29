const XLSX = require('xlsx');

// Fungsi untuk membaca aturan promo dari file Excel (termasuk info "Wajib Gabungan")
function bacaPromoExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const promoData = XLSX.utils.sheet_to_json(sheet);

    const promoRules = {};
    const kelompokGabungan = {};

    promoData.forEach(row => {
        const {
            Area, 'Nama produk': namaProduk, 'Tipe Promo': tipePromo,
            Min, Max, Value, 'Produk Gabungan': produkGabungan,
            'Kelompok Gabungan': kelompok, 'Wajib Gabungan': wajibGabungan
        } = row;

        if (!promoRules[namaProduk]) promoRules[namaProduk] = {};
        if (!promoRules[namaProduk][Area]) promoRules[namaProduk][Area] = [];

        promoRules[namaProduk][Area].push({
            tipePromo,
            min: Min,
            max: Max === '999999' ? Infinity : Max,
            value: Value,
            produkGabungan: produkGabungan === 'Ya',
            kelompok,
            wajibGabungan: wajibGabungan === 'Ya'
        });

        if (produkGabungan === 'Ya' && kelompok) {
            if (!kelompokGabungan[kelompok]) kelompokGabungan[kelompok] = {};
            if (!kelompokGabungan[kelompok][Area]) kelompokGabungan[kelompok][Area] = [];
            if (!kelompokGabungan[kelompok][Area].includes(namaProduk)) {
                kelompokGabungan[kelompok][Area].push(namaProduk);
            }
        }
    });

    return { promoRules, kelompokGabungan };
}

// Fungsi untuk menghitung bonus atau cashback berdasarkan aturan promo
function hitungPromo(jumlahKarton, namaProduk, area, promoRules) {
    let totalBonusBarang = 0;
    let totalCashback = 0;
    let layerBonus = [];
    let layerCashback = [];

    if (promoRules[namaProduk] && promoRules[namaProduk][area]) {
        const rules = promoRules[namaProduk][area];

        // Mengurutkan aturan promo berdasarkan kolom Min secara menurun untuk BONUS BARANG
        const sortedBonusRules = rules
            .filter(rule => rule.tipePromo === 'Bonus Barang')
            .sort((a, b) => b.min - a.min);

        let remainingKarton = jumlahKarton;

        sortedBonusRules.forEach(rule => {
            if (remainingKarton >= rule.min) {
                const kelipatan = Math.floor(remainingKarton / rule.min);
                totalBonusBarang += kelipatan * rule.value;
                layerBonus.push(`Layer ${rule.min}-${rule.max} Karton`);
                remainingKarton -= kelipatan * rule.min;
            }
        });

        if (remainingKarton > 0) {
            sortedBonusRules.forEach(rule => {
                if (remainingKarton >= rule.min) {
                    totalBonusBarang += rule.value;
                    layerBonus.push(`Layer ${rule.min}-${rule.max} Karton`);
                    remainingKarton = 0;
                }
            });
        }
    }

    // Proses CASH BACK
    if (promoRules[namaProduk] && promoRules[namaProduk][area]) {
        const rules = promoRules[namaProduk][area];

        rules.forEach(rule => {
            if (rule.tipePromo === 'Cashback') {
                if (jumlahKarton >= rule.min && jumlahKarton <= rule.max) {
                    totalCashback = rule.value * jumlahKarton;
                    layerCashback.push(`Layer ${rule.min}-${rule.max} Karton`);
                }
            }
        });
    }

    return { totalBonusBarang, totalCashback, layerBonus, layerCashback };
}

// Fungsi untuk memproses transaksi dan menghitung promo gabungan (wajib atau tidak)
function prosesTransaction(filePath, promoRules, kelompokGabungan) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);

    const hasilPerhitungan = [];
    const transaksiGabungan = {};

    rows.forEach((row, index) => {
        // Pengecekan kolom yang hilang
        const requiredColumns = ['Nama_toko', 'Nota', 'Nama Produk', 'Jumlah Karton', 'Value Netto', 'Area'];
        const missingColumns = requiredColumns.filter(col => !row[col]);

        if (missingColumns.length > 0) {
            console.error(`Data tidak lengkap pada baris ${index + 1}. Kolom yang hilang: ${missingColumns.join(', ')}`);
            return; // Lewatkan baris ini jika ada kolom yang hilang
        }

        // Pengecekan tipe data yang salah
        if (isNaN(row['Jumlah Karton']) || isNaN(row['Value Netto'])) {
            console.error(`Nilai tidak valid pada baris ${index + 1}. Pastikan Jumlah Karton dan Value Netto berupa angka.`);
            return; // Lewatkan baris ini jika data tidak valid
        }

        const key = `${row.Nama_toko}-${row.Nota}`;
        if (!transaksiGabungan[key]) transaksiGabungan[key] = [];
        transaksiGabungan[key].push(row);
    });

    Object.entries(transaksiGabungan).forEach(([key, transaksiList]) => {
        const area = transaksiList[0].Area;

        const kelompokQtyMap = {};
        const kelompokValueMap = {};
        const produkToKelompok = {};
        const produkInNota = new Set();

        // Menghitung total karton dan nilai untuk setiap kelompok produk gabungan
        for (const [kelompok, wilayah] of Object.entries(kelompokGabungan)) {
            const produkList = wilayah[area] || [];
            let totalKarton = 0;
            let totalValue = 0;

            transaksiList.forEach(row => {
                if (produkList.includes(row['Nama Produk'])) {
                    totalKarton += row['Jumlah Karton'] || 0;
                    totalValue += row['Value Netto'] || 0;
                    produkToKelompok[row['Nama Produk']] = kelompok;
                    produkInNota.add(row['Nama Produk']);
                }
            });

            if (totalKarton > 0) kelompokQtyMap[kelompok] = totalKarton;
            if (totalValue > 0) kelompokValueMap[kelompok] = totalValue;
        }

        transaksiList.forEach(row => {
            const namaProduk = row['Nama Produk'];
            const jumlahKarton = row['Jumlah Karton'];
            const valueNetto = row['Value Netto'] || 0;
            const { Periode, Region, Divisi, Distributor, Depo, Area, Unique_Code, Nama_toko, Nota, Tgl_Nota, RegFest, Qty_KTN, qtyInPcs } = row;

            const { totalBonusBarang, totalCashback, layerBonus, layerCashback } =
                hitungPromo(jumlahKarton, namaProduk, area, promoRules);

            let promoGabunganQty = 0, promoGabunganValue = 0;
            let layerGabunganQty = '', kelipatanGabunganQty = 0;
            let kelompok = produkToKelompok[namaProduk];

            if (kelompok) {
                const rules = promoRules[namaProduk][area];
                const totalKartonGab = kelompokQtyMap[kelompok] || 0;
                const totalValueGab = kelompokValueMap[kelompok] || 0;

                // Periksa apakah semua produk dalam kelompok gabungan ada di nota ini jika wajib gabungan
                const produkGabunganInNota = [...kelompokGabungan[kelompok][area]].every(produk => produkInNota.has(produk));

                for (const rule of rules) {
                    if (rule.kelompok === kelompok) {
                        if (rule.wajibGabungan) {
                            if (!produkGabunganInNota) {
                                // Jika ada produk gabungan yang hilang, tidak bisa mendapatkan promo gabungan
                                promoGabunganQty = 0;
                                promoGabunganValue = 0;
                                layerGabunganQty = '';
                                kelipatanGabunganQty = 0;
                                break;
                            }
                            if (rule.tipePromo === 'Gabungan Kuantitas' && totalKartonGab >= rule.min) {
                                kelipatanGabunganQty = Math.floor(totalKartonGab / rule.min);
                                promoGabunganQty = Math.round((valueNetto / totalValueGab) * (kelipatanGabunganQty * rule.value));
                                layerGabunganQty = `Layer ${rule.min}-${rule.max} Karton`;
                            } else if (rule.tipePromo === 'Gabungan Value' && totalValueGab >= rule.min) {
                                const kelipatanValue = Math.floor(totalValueGab / rule.min);
                                promoGabunganValue = Math.round((valueNetto / totalValueGab) * (kelipatanValue * rule.value));
                            }
                        } else {
                            // Jika tidak wajib gabungan, tetap hitung promo gabungan meski tidak semua produk ada
                            if (rule.tipePromo === 'Gabungan Kuantitas' && totalKartonGab >= rule.min) {
                                kelipatanGabunganQty = Math.floor(totalKartonGab / rule.min);
                                promoGabunganQty = Math.round((valueNetto / totalValueGab) * (kelipatanGabunganQty * rule.value));
                                layerGabunganQty = `Layer ${rule.min}-${rule.max} Karton`;
                            } else if (rule.tipePromo === 'Gabungan Value' && totalValueGab >= rule.min) {
                                const kelipatanValue = Math.floor(totalValueGab / rule.min);
                                promoGabunganValue = Math.round((valueNetto / totalValueGab) * (kelipatanValue * rule.value));
                            }
                        }
                    }
                }
            }

            hasilPerhitungan.push({
                Periode, Region, Divisi, Distributor, Depo, Area, Unique_Code,
                Nama_toko, Nota, Tgl_Nota, namaProduk, RegFest, Qty_KTN,
                valueNetto, qtyInPcs, jumlahKarton,
                totalBonusBarang, totalCashback,
                LayerBonus: layerBonus.join('; '),
                LayerCashback: layerCashback.join('; '),
                PromoGabunganQty: promoGabunganQty,
                PromoGabunganValue: promoGabunganValue,
                LayerGabunganQty: layerGabunganQty,
                KelipatanGabunganQty: kelipatanGabunganQty
            });
        });
    });

    const newWorkbook = XLSX.utils.book_new();
    const hasilWorksheet = XLSX.utils.json_to_sheet(hasilPerhitungan);
    XLSX.utils.book_append_sheet(newWorkbook, hasilWorksheet, 'Hasil Perhitungan');
    XLSX.writeFile(newWorkbook, 'hasil_perhitungan_promo.xlsx');
    console.log('✅ File hasil_perhitungan_promo.xlsx berhasil dibuat.');
}

// Membaca aturan promo dari file Excel
const { promoRules, kelompokGabungan } = bacaPromoExcel('aturan_promo.xlsx');

// Memproses file data sell out dan menghitung hasil promo
console.time('Proses Perhitungan Promo');
prosesTransaction('data_sell_out.xlsx', promoRules, kelompokGabungan);
console.timeEnd('Proses Perhitungan Promo');
