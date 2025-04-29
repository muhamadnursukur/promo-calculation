const XLSX = require('xlsx');
const path = require('path');

// Fungsi untuk membaca aturan promo dari file Excel
function bacaPromoExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const promoData = XLSX.utils.sheet_to_json(sheet);

    const promoRules = {};
    const kelompokGabungan = {};

    promoData.forEach(row => {
        const {
            Area, 'Nama produk': namaProdukRaw, 'Tipe Promo': tipePromo,
            Min, Max, Value, 'Produk Gabungan': produkGabungan,
            'Kelompok Gabungan': kelompok, 'Wajib Gabungan': wajibGabungan
        } = row;

        if (!Area || !namaProdukRaw || !tipePromo || Min == null || Value == null) return;

        const namaProduks = namaProdukRaw.includes('|') ? namaProdukRaw.split('|').map(p => p.trim()) : [namaProdukRaw];

        namaProduks.forEach(namaProduk => {
            if (!promoRules[namaProduk]) promoRules[namaProduk] = {};
            if (!promoRules[namaProduk][Area]) promoRules[namaProduk][Area] = [];

            promoRules[namaProduk][Area].push({
                tipePromo,
                min: parseFloat(Min),
                max: Max === '999999' ? Infinity : parseFloat(Max),
                value: parseFloat(Value),
                produkGabungan: produkGabungan === 'Ya',
                kelompok,
                wajibGabungan: wajibGabungan === 'Ya'
            });

            if (produkGabungan === 'Ya' && kelompok) {
                if (!kelompokGabungan[kelompok]) kelompokGabungan[kelompok] = {};
                if (!kelompokGabungan[kelompok][Area]) kelompokGabungan[kelompok][Area] = [];

                namaProduks.forEach(prod => {
                    if (!kelompokGabungan[kelompok][Area].includes(prod)) {
                        kelompokGabungan[kelompok][Area].push(prod);
                    }
                });
            }
        });
    });

    return { promoRules, kelompokGabungan };
}

function hitungPromo(jumlahKarton, namaProduk, area, promoRules) {
    let totalBonusBarang = 0;
    let totalCashback = 0;
    let layerBonus = [];
    let layerCashback = [];

    if (promoRules[namaProduk] && promoRules[namaProduk][area]) {
        const rules = promoRules[namaProduk][area];

        const sortedBonusRules = rules
            .filter(rule => rule.tipePromo === 'Bonus Barang')
            .sort((a, b) => b.min - a.min);

        let remainingKarton = jumlahKarton;

        sortedBonusRules.forEach(rule => {
            if (remainingKarton >= rule.min) {
                const kelipatan = Math.floor(remainingKarton / rule.min);
                totalBonusBarang += kelipatan * rule.value;
                layerBonus.push(`Layer ${rule.min}-${rule.max} Karton (${kelipatan}x)`);
                remainingKarton -= kelipatan * rule.min;
            }
        });

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

function prosesTransaction(filePath, promoRules, kelompokGabungan) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);

    const hasilPerhitungan = [];
    const transaksiGabungan = {};

    rows.forEach(row => {
        const key = `${row.Nama_toko}-${row.Nota}`;
        if (!transaksiGabungan[key]) transaksiGabungan[key] = [];
        transaksiGabungan[key].push(row);
    });

    Object.entries(transaksiGabungan).forEach(([key, transaksiList]) => {
        const area = transaksiList[0].Area;

        const kelompokQtyMap = {};
        const kelompokValueMap = {};
        const produkToKelompok = {};

        for (const [kelompok, wilayah] of Object.entries(kelompokGabungan)) {
            const produkList = wilayah[area] || [];
            let totalKarton = 0;
            let totalValue = 0;

            transaksiList.forEach(row => {
                if (produkList.includes(row['Nama Produk'])) {
                    const qty = row['Jumlah Karton'] || 0;
                    const value = row['Value Netto'] || 0;
                    totalKarton += qty;
                    totalValue += value;
                    produkToKelompok[row['Nama Produk']] = kelompok;
                }
            });

            if (totalKarton > 0) kelompokQtyMap[kelompok] = totalKarton;
            if (totalValue > 0) kelompokValueMap[kelompok] = totalValue;
        }

        transaksiList.forEach(row => {
            const namaProduk = row['Nama Produk'];
            const jumlahKarton = row['Jumlah Karton'];
            const valueNetto = row['Value Netto'] || 0;

            const { Periode, Region, Divisi, Distributor, Depo, Area,
                Unique_Code, Nama_toko, Nota, Tgl_Nota, RegFest,
                Qty_KTN, qtyInPcs } = row;

            const { totalBonusBarang, totalCashback, layerBonus, layerCashback } =
                hitungPromo(jumlahKarton, namaProduk, area, promoRules);

            let promoGabunganQty = 0, promoGabunganValue = 0;
            let layerGabunganQty = '', kelipatanGabunganQty = 0;
            let kelompok = produkToKelompok[namaProduk];

            if (kelompok) {
                const produkList = kelompokGabungan[kelompok][area] || [];
                const produkAdaSemua = produkList.every(prod =>
                    transaksiList.some(transaction => transaction['Nama Produk'] === prod)
                );

                if (produkAdaSemua) {
                    const rules = promoRules[namaProduk][area];
                    const totalKartonGab = kelompokQtyMap[kelompok] || 0;
                    const totalValueGab = kelompokValueMap[kelompok] || 0;

                    // ==== PROMO GABUNGAN KUANTITAS (BERLAPIS) ====
                    const applicableQtyRules = rules.filter(rule =>
                        rule.kelompok === kelompok && rule.tipePromo === 'Gabungan Kuantitas'
                    );

                    const sortedGabunganQtyRules = applicableQtyRules.sort((a, b) => b.min - a.min);
                    let remainingKartonGab = totalKartonGab;

                    sortedGabunganQtyRules.forEach(rule => {
                        if (remainingKartonGab >= rule.min) {
                            const kelipatan = Math.floor(remainingKartonGab / rule.min);
                            const bonus = Math.round((valueNetto / totalValueGab) * (kelipatan * rule.value));
                            promoGabunganQty += bonus;
                            kelipatanGabunganQty += kelipatan;
                            layerGabunganQty += `Layer ${rule.min}-${rule.max} Karton (${kelipatan}x); `;
                            remainingKartonGab -= kelipatan * rule.min;
                        }
                    });

                    // ==== PROMO GABUNGAN VALUE ====
                    rules.forEach(rule => {
                        if (rule.kelompok === kelompok && rule.tipePromo === 'Gabungan Value') {
                            if (totalValueGab >= rule.min) {
                                const kelipatanValue = Math.floor(totalValueGab / rule.min);
                                if (totalValueGab > 0 && valueNetto > 0) {
                                    promoGabunganValue += (valueNetto / totalValueGab) * (kelipatanValue * rule.value);
                                    // promoGabunganValue += Math.round((valueNetto / totalValueGab) * (kelipatanValue * rule.value));
                                }
                            }
                        }
                    });
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

// Eksekusi utama
const { promoRules, kelompokGabungan } = bacaPromoExcel('aturan_promo.xlsx');

console.time('Proses Perhitungan Promo');
prosesTransaction('data_sell_out.xlsx', promoRules, kelompokGabungan);
console.timeEnd('Proses Perhitungan Promo');
