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
            .sort((a, b) => b.min - a.min); // Mengurutkan nilai pada kolom Min dari terbesar ke terkecil

        let remainingKarton = jumlahKarton;

        // memproses setiap lapisan untuk promo bonus barang
        sortedBonusRules.forEach(rule => {
            if (remainingKarton >= rule.min) {
                const kelipatan = Math.floor(remainingKarton / rule.min);
                totalBonusBarang += kelipatan * rule.value;
                layerBonus.push(`Layer ${rule.min}-${rule.max} Karton`);
                remainingKarton -= kelipatan * rule.min;
            }
        });

        // Jika masih ada karton tersisa dan bisa masuk ke lapisan yang lebih kecil
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

        // Proses setiap lapisan CASH BACK
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
    const totalData = rows.length;
    let processedData = 0;

    rows.forEach(row => {
        const { Periode, Region, Divisi, Distributor, Depo, Area, Unique_Code, Nama_toko, Nota, Tgl_Nota,
            'Nama Produk': namaProduk, RegFest, Qty_KTN, 'Value Netto': valueNetto, 'Qty In Pcs': qtyInPcs, 'Jumlah Karton': jumlahKarton } = row;

        if (namaProduk && jumlahKarton !== undefined && Area) {
            const { totalBonusBarang, totalCashback, layerBonus, layerCashback } = hitungPromo(jumlahKarton, namaProduk, Area, promoRules);

            // Menyiapkan hasil perhitungan
            const hasil = {
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
                totalBonusBarang,
                totalCashback,
                LayerBonus: layerBonus.join('; '), // Menggabungkan semua lapisan bonus menjadi satu kolom
                LayerCashback: layerCashback.join('; ') // Menggabungkan semua lapisan cashback menjadi satu kolom
            };

            hasilPerhitungan.push(hasil);
        }
        processedData++;

        // Log progress setiap 100 data yang diproses
        if (processedData % 100 === 0 || processedData === totalData) {
            console.log(`Proses Data: ${Math.round((processedData / totalData) * 100)}% selesai (${processedData} dari ${totalData} data)`);
        }
    });

    // Membuat workbook baru untuk hasil perhitungan
    const newWorkbook = XLSX.utils.book_new();

    // Menyiapkan data untuk ditulis ke dalam worksheet
    const hasilWorksheet = XLSX.utils.json_to_sheet(hasilPerhitungan);

    // Menambahkan worksheet ke workbook
    XLSX.utils.book_append_sheet(newWorkbook, hasilWorksheet, 'Hasil Perhitungan');

    // Menyimpan workbook baru ke file Excel
    XLSX.writeFile(newWorkbook, 'hasil_perhitungan_promo.xlsx');

    console.log('File Excel hasil perhitungan telah dibuat: hasil_perhitungan_promo.xlsx');
}

// Membaca aturan promo dari file Excel
const promoRules = bacaPromoExcel('aturan_promo.xlsx');

console.time('Proses Perhitungan Promo');

// Memproses file data sell out dan menghitung hasil promo
prosesTransaction('data_sell_out.xlsx', promoRules);

console.timeEnd('Proses Perhitungan Promo');
