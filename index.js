const XLSX = require('xlsx');

// Fungsi untuk membaca aturan promo dari file Excel
function bacaPromoExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const promoData = XLSX.utils.sheet_to_json(sheet);

    let promoRules = {};

    // membaca setiap baris dari aturan promo yang ada di file excel
    promoData.forEach(row => {
        const { Area, 'Nama produk': namaProduk, 'Tipe Promo': tipePromo, Min, Max, Value } = row;

        if (!promoRules[namaProduk]) {
            promoRules[namaProduk] = {};
        }
        if (!promoRules[namaProduk][Area]) {
            promoRules[namaProduk][Area] = [];
        }

        promoRules[namaProduk][Area].push({
            tipePromo: tipePromo,
            min: Min,
            max: Max === '999999' ? Infinity : Max,
            value: Value
        });
    });

    return promoRules;
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

// Fungsi untuk memproses pembelian dan menghitung promo
function prosesTransaction(filePath, promoRules) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);

    const hasilPerhitungan = [];
    const produkGabunganPerRegion = {
        'JAWA 1': ['FORTIUS 10', 'FORTIUS 30'],
        'JAWA 2': ['OKEBIS KELAPA EXTRA 28', 'MARIE SUSU 40']
    };

    const minimumValuePromo = 300000;
    const valuePerKelipatan = 10000;

    let transaksiGabungan = {};
    rows.forEach(row => {
        const key = `${row.Nama_toko}-${row.Nota}`;
        if (!transaksiGabungan[key]) {
            transaksiGabungan[key] = [];
        }
        transaksiGabungan[key].push(row);
    });

    const totalData = Object.keys(transaksiGabungan).length;
    let processedData = 0;

    Object.entries(transaksiGabungan).forEach(([key, transaksiList]) => {
        const region = transaksiList[0].Region;
        // console.log(`Memeriksa region: ${region}`); // Log untuk debugging

        const produkGabungan = produkGabunganPerRegion[region] || [];
        // console.log(`Produk gabungan untuk region ${region}: ${produkGabungan}`); // Log untuk debugging

        const produkTerlibat = transaksiList.filter(row => produkGabungan.includes(row['Nama Produk']));
        const totalValueGabungan = produkTerlibat.reduce((sum, row) => sum + (row['Value Netto'] || 0), 0);

        // console.log(`Total Value Gabungan untuk region ${region}: ${totalValueGabungan}`); // Log untuk debugging

        const kelipatan = Math.floor(totalValueGabungan / minimumValuePromo);
        const promoGabungan = kelipatan * valuePerKelipatan;

        transaksiList.forEach(row => {
            const {
                Periode, Region, Divisi, Distributor, Depo, Area, Unique_Code,
                Nama_toko, Nota, Tgl_Nota, 'Nama Produk': namaProduk, RegFest,
                Qty_KTN, 'Value Netto': valueNetto, 'Qty In Pcs': qtyInPcs, 'Jumlah Karton': jumlahKarton
            } = row;

            const valueNettoBaris = valueNetto || 0;

            let porsiCashback = 0;
            if (produkGabungan.includes(namaProduk)) {
                porsiCashback = totalValueGabungan > 0
                    ? (valueNettoBaris / totalValueGabungan) * promoGabungan
                    : 0;
            }

            let hasil = {
                Periode,
                Region,
                Divisi,
                Distributor,
                Depo,
                Area,
                Unique_Code,
                Nama_toko,
                Nota,
                Tgl_Nota,
                namaProduk,
                RegFest,
                Qty_KTN,
                valueNetto,
                qtyInPcs,
                jumlahKarton,
                totalBonusBarang: 0,
                totalCashback: 0,
                LayerBonus: '',
                LayerCashback: '',
                PromoGabungan: Math.round(porsiCashback),
                KelipatanGabungan: produkGabungan.includes(namaProduk) ? kelipatan : 0,
                TotalValueGabungan: produkGabungan.includes(namaProduk) ? totalValueGabungan : 0
            };

            if (namaProduk && jumlahKarton !== undefined && Area) {
                const { totalBonusBarang, totalCashback, layerBonus, layerCashback } =
                    hitungPromo(jumlahKarton, namaProduk, Area, promoRules);

                hasil.totalBonusBarang = totalBonusBarang;
                hasil.totalCashback = totalCashback;
                hasil.LayerBonus = layerBonus.join('; ');
                hasil.LayerCashback = layerCashback.join('; ');
            }

            hasilPerhitungan.push(hasil);
        });

        processedData++;
        if (processedData % 100 === 0 || processedData === totalData) {
            console.log(`Proses Data: ${Math.round((processedData / totalData) * 100)}% selesai (${processedData}/${totalData})`);
        }
    });

    const newWorkbook = XLSX.utils.book_new();
    const hasilWorksheet = XLSX.utils.json_to_sheet(hasilPerhitungan);
    XLSX.utils.book_append_sheet(newWorkbook, hasilWorksheet, 'Hasil Perhitungan');
    XLSX.writeFile(newWorkbook, 'hasil_perhitungan_promo.xlsx');

    console.log('✅ File hasil_perhitungan_promo.xlsx telah berhasil dibuat.');
}

// Membaca aturan promo dari file Excel
const promoRules = bacaPromoExcel('aturan_promo.xlsx');

console.time('Proses Perhitungan Promo');

// Memproses file data sell out dan menghitung hasil promo
prosesTransaction('data_sell_out.xlsx', promoRules);

console.timeEnd('Proses Perhitungan Promo');
