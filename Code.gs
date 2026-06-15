function doGet() {
  return HtmlService.createHtmlOutputFromFile('index') // Pastikan nama file HTML Anda 'index'
    .setTitle('Kedai Es Jerukkist')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// 1. FUNGSI UNTUK MENGAMBIL DATA DASHBOARD (PERBAIKAN UTAMA)
// ==========================================
function getDashboardData(namaCabang) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!namaCabang) namaCabang = "Pusat";
  
  // 1. Ambil Data Bahan Baku & Filter per Cabang
  const sheetBahan = ss.getSheetByName('BahanBaku');
  const dataBahanRaw = sheetBahan ? sheetBahan.getDataRange().getValues() : [];
  if (dataBahanRaw.length > 0) dataBahanRaw.shift(); // Buang Header
  
  // Ambil Bahan Baku yang sesuai namaCabang
  // Format Array di frontend: item[0]=Cabang, item[1]=NamaBahan, item[2]=Stok, item[3]=Satuan
  const dataBahan = dataBahanRaw.filter(row => row[0] === namaCabang).map(row => {
  return [
    row[1], // ID Bahan (Kita jadikan Index 0) -> Kolom B di Sheet
    row[2], // Nama Bahan (Index 1) -> Kolom C di Sheet
    row[3], // Angka Stok (Index 2) -> Kolom D di Sheet
    row[4] || "Buah" // Satuan (Index 3) -> Kolom E di Sheet
  ];
});

  // 2. Ambil Data Gelas & Filter per Cabang (DIPERBAIKI AGAR SATUAN DAN ANGKA MUNCUL)
  const sheetGelas = ss.getSheetByName('GelasPlastik');
  const dataGelasRaw = sheetGelas ? sheetGelas.getDataRange().getValues() : [];
  if (dataGelasRaw.length > 0) dataGelasRaw.shift(); // Buang Header
  
  // Ambil Gelas Plastik yang sesuai namaCabang
  // Format Array di frontend: item[0]=Cabang, item[1]=NamaGelas, item[2]=Stok, item[3]=Satuan
  const dataGelas = dataGelasRaw.filter(row => row[0] === namaCabang).map(row => {
  return [
    row[1], // ID Gelas (Kita jadikan Index 0) -> Kolom B di Sheet
    row[2], // Nama Gelas (Index 1) -> Kolom C di Sheet
    row[3], // Angka Stok Gelas (Index 2) -> Kolom D di Sheet
    row[4] || ""  // Satuan Gelas (Index 3) -> Kolom E di Sheet
  ];
});

  // 3. Ambil Data Penjualan & Filter per Cabang
  const sheetJual = ss.getSheetByName('Penjualan');
  const dataJualRaw = sheetJual ? sheetJual.getDataRange().getValues() : [];
  if (dataJualRaw.length > 0) dataJualRaw.shift(); // Buang Header
  const dataJual = dataJualRaw.filter(row => row[0] === namaCabang);

  // Perhitungan Summary Box
  let totalOmzet = 0;
  let totalTerjual = 0;
  
  dataJual.forEach(row => {
    totalTerjual += Number(row[5]) || 0; // Kolom Jumlah Jual (Kolom F)
    totalOmzet += Number(row[6]) || 0;   // Kolom Total Harga (Kolom G)
  });

  // Data untuk Grafik Statistik (7 transaksi terakhir)
  let labels = [];
  let values = [];
  dataJual.slice(-7).forEach(row => {
    if (row[1]) {
      labels.push(row[1].toString().substring(0, 10)); // Kolom Waktu/Tanggal
      values.push(Number(row[5]) || 0);
    }
  });

  return {
    summary: {
      omzet: totalOmzet.toLocaleString('id-ID'),
      terjual: totalTerjual,
      stokBahan: dataBahan.length,
      stokGelas: dataGelas.length
    },
    bahan: dataBahan,
    gelas: dataGelas,
    labels: labels,
    values: values
  };
}

// ==========================================
// 2. FUNGSI UNTUK MENYIAPKAN FORM INPUT JUAL
// ==========================================
function getFormData(namaCabang) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!namaCabang) namaCabang = "Pusat";
  
  // Ambil Semua Cabang untuk drop-down filter di atas
  const sheetCabang = ss.getSheetByName('DaftarCabang');
  let semuaCabang = ["Pusat"]; 
  if (sheetCabang) {
    const dataC = sheetCabang.getDataRange().getValues();
    if (dataC.length > 1) {
      semuaCabang = dataC.slice(1).map(row => row[0]);
    }
  }

  const sheetGelas = ss.getSheetByName('GelasPlastik');
  const daftarGelas = sheetGelas ? sheetGelas.getDataRange().getValues()
                        .filter((row, i) => i > 0 && row[0] === namaCabang)
                        .map(row => row[2]) : []; // Kolom C (Nama Gelas)

  const sheetBahan = ss.getSheetByName('BahanBaku');
  const daftarBahan = sheetBahan ? sheetBahan.getDataRange().getValues()
                        .filter((row, i) => i > 0 && row[0] === namaCabang)
                        .map(row => row[2]) : []; // Kolom C (Nama Bahan)

  return { 
    daftarGelas: daftarGelas, 
    daftarBahan: daftarBahan,
    semuaCabang: semuaCabang 
  };
}

// ==========================================
// 3. FUNGSI TRANSAKSI PENJUALAN & UPDATE STOK
// ==========================================
function tambahTransaksi(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetJual = ss.getSheetByName('Penjualan');
    const sheetBahan = ss.getSheetByName('BahanBaku');
    const sheetGelas = ss.getSheetByName('GelasPlastik');

    // Simpan ke sheet Penjualan
    sheetJual.appendRow([
      data.cabang, 
      new Date(), 
      data.kasir, 
      data.jenisGelas, 
      JSON.stringify(data.pemakaianBahan), 
      data.jumlah, 
      data.total
    ]);

    // Kurangi Stok Bahan Baku di Kolom ke-4 (Kolom D) berdasarkan Cabang terkait
    const dataBahan = sheetBahan.getDataRange().getValues();
    for (let namaBahan in data.pemakaianBahan) {
      for (let i = 1; i < dataBahan.length; i++) {
        if (dataBahan[i][0] === data.cabang && dataBahan[i][2] === namaBahan) {
          let stokLama = Number(dataBahan[i][3]) || 0; // Kolom D
          let jumlahGelas = Number(data.jumlah) || 0;
          let pemakaianPerGelas = Number(data.pemakaianBahan[namaBahan]) || 0;
          sheetBahan.getRange(i + 1, 4).setValue(stokLama - (jumlahGelas * pemakaianPerGelas));
        }
      }
    }

    // Kurangi Stok Gelas di Kolom ke-4 (Kolom D) berdasarkan Cabang terkait
    const dataGelas = sheetGelas.getDataRange().getValues();
    for (let i = 1; i < dataGelas.length; i++) {
      if (dataGelas[i][0] === data.cabang && dataGelas[i][2] === data.jenisGelas) {
        let stokGelasLama = Number(dataGelas[i][3]) || 0; // Kolom D
        sheetGelas.getRange(i + 1, 4).setValue(stokGelasLama - Number(data.jumlah));
      }
    }

    return "Transaksi Cabang berhasil disimpan!";
  } catch(e) {
    return "Gagal: " + e.toString();
  }
}

// ==========================================
// FUNGSI UNTUK INPUT / TAMBAH STOK BAHAN BAKU (SISTEM AKUMULASI)
// ==========================================
function tambahBahan(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BahanBaku');
  const dataBahan = sheet.getDataRange().getValues();
  
  let itemDitemukan = false;
  
  // Looping untuk mencari apakah item sudah ada di cabang tersebut
  for (let i = 1; i < dataBahan.length; i++) {
    // Cek kesamaan Nama Cabang (Kolom A) DAN Nama Bahan (Kolom C / Index 2)
    if (dataBahan[i][0] === data.cabang && dataBahan[i][2].toString().trim().toUpperCase() === data.nama.toString().trim().toUpperCase()) {
      
      let stokLama = Number(dataBahan[i][3]) || 0; // Kolom D
      let stokTambahan = Number(data.stok) || 0;
      
      // Update nilai stok baru (Stok Lama + Stok Tambahan) di Kolom D (Baris ke i+1, Kolom 4)
      sheet.getRange(i + 1, 4).setValue(stokLama + stokTambahan);
      
      // Jika satuan diubah saat input, update juga satuannya di Kolom E
      if (data.satuan) {
        sheet.getRange(i + 1, 5).setValue(data.satuan);
      }
      
      itemDitemukan = true;
      break;
    }
  }
  
  // Jika item benar-benar belum pernah ada di cabang tersebut, buat baris baru
  if (!itemDitemukan) {
    sheet.appendRow([data.cabang, data.id, data.nama, data.stok, data.satuan]);
    return "Bahan Baku baru '" + data.nama + "' berhasil didaftarkan ke " + data.cabang;
  }
  
  return "Stok Bahan Baku '" + data.nama + "' berhasil ditambah!";
}

// ==========================================
// FUNGSI UNTUK INPUT / TAMBAH STOK BAHAN BAKU (CEK ID & NAMA)
// ==========================================
function tambahBahan(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BahanBaku');
  const dataBahan = sheet.getDataRange().getValues();
  
  let itemDitemukan = false;
  
  for (let i = 1; i < dataBahan.length; i++) {
    // Cek di cabang yang sama
    if (dataBahan[i][0] === data.cabang) {
      
      const idSama = dataBahan[i][1].toString().trim().toUpperCase() === data.id.toString().trim().toUpperCase();
      const namaSama = dataBahan[i][2].toString().trim().toUpperCase() === data.nama.toString().trim().toUpperCase();
      
      // JIKA ID SAMA ATAU NAMA SAMA, MAKA AKUMULASIKAN STOKNYA
      if (idSama || namaSama) {
        let stokLama = Number(dataBahan[i][3]) || 0;
        let stokTambahan = Number(data.stok) || 0;
        
        sheet.getRange(i + 1, 3).setValue(data.nama); // Pastikan namanya sinkron
        sheet.getRange(i + 1, 4).setValue(stokLama + stokTambahan); // Tambah Stok
        if (data.satuan) sheet.getRange(i + 1, 5).setValue(data.satuan); // Update Satuan
        
        itemDitemukan = true;
        break;
      }
    }
  }
  
  if (!itemDitemukan) {
    sheet.appendRow([data.cabang, data.id, data.nama, data.stok, data.satuan]);
    return "Bahan Baku baru '" + data.nama + "' berhasil didaftarkan ke " + data.cabang;
  }
  
  return "Stok Bahan Baku '" + data.nama + "' berhasil ditambahkan!";
}

// ==========================================
// FUNGSI UNTUK INPUT / TAMBAH STOK WADAH/VARIAN (CEK ID & NAMA)
// ==========================================
function tambahGelas(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('GelasPlastik');
  const dataGelas = sheet.getDataRange().getValues();
  
  let itemDitemukan = false;
  
  for (let i = 1; i < dataGelas.length; i++) {
    // Cek di cabang yang sama
    if (dataGelas[i][0] === data.cabang) {
      
      const idSama = dataGelas[i][1].toString().trim().toUpperCase() === data.id.toString().trim().toUpperCase();
      const namaSama = dataGelas[i][2].toString().trim().toUpperCase() === data.nama.toString().trim().toUpperCase();
      
      // JIKA ID SAMA ATAU NAMA SAMA, MAKA AKUMULASIKAN STOKNYA
      if (idSama || namaSama) {
        let stokLama = Number(dataGelas[i][3]) || 0;
        let stokTambahan = Number(data.stok) || 0;
        
        sheet.getRange(i + 1, 3).setValue(data.nama); // Pastikan namanya sinkron
        sheet.getRange(i + 1, 4).setValue(stokLama + stokTambahan); // Tambah Stok
        
        itemDitemukan = true;
        break;
      }
    }
  }
  
  if (!itemDitemukan) {
    sheet.appendRow([data.cabang, data.id, data.nama, data.stok, "Pcs"]);
    return "Wadah/Varian baru '" + data.nama + "' berhasil didaftarkan ke " + data.cabang;
  }
  
  return "Stok Wadah/Varian '" + data.nama + "' berhasil ditambahkan!";
}

// ==========================================
// 5. FUNGSI LAINNYA
// ==========================================
function simpanCabangBaru(nama) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('DaftarCabang');
  if (!sheet) {
    sheet = ss.insertSheet('DaftarCabang');
    sheet.appendRow(['Nama Cabang']);
  }
  sheet.appendRow([nama]);
  return "Cabang " + nama + " berhasil didaftarkan!";
}

function clearSemuaData(cabang) {
  if (!cabang || cabang === "" || cabang === "Memuat Cabang...") {
    return "Gagal: Pilih nama cabang yang valid terlebih dahulu!";
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetNames = ["Penjualan", "BahanBaku", "GelasPlastik", "DaftarCabang"];
    var totalTerhapus = 0;

    sheetNames.forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      if (sheet) {
        var data = sheet.getDataRange().getValues();
        for (var i = data.length - 1; i >= 1; i--) {
          var nilaiKolomA = data[i][0] ? data[i][0].toString().trim() : "";
          if (nilaiKolomA.toLowerCase() === cabang.toString().trim().toLowerCase()) {
            sheet.deleteRow(i + 1);
            totalTerhapus++;
          }
        }
      }
    });
    return "Berhasil! Data cabang '" + cabang + "' telah dibersihkan.";
  } catch (err) {
    return "Terjadi Kesalahan: " + err.toString();
  }
}
// ==========================================
// FUNGSI BARU: MENGHAPUS SATU ITEM BAHAN/WADAH TERTENTU DI CABANG TERKAIT
// ==========================================
function hapusSatuItem(cabang, jenisSheet, idItem, namaItem) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Tentukan sheet berdasarkan input dari frontend ('BahanBaku' atau 'GelasPlastik')
    const sheet = ss.getSheetByName(jenisSheet);
    
    if (!sheet) return "Gagal: Sheet tidak ditemukan.";
    
    const data = sheet.getDataRange().getValues();
    let berhasilDihapus = false;
    
    // Looping mundur dari bawah ke atas agar indeks baris tidak bergeser saat dihapus
    for (let i = data.length - 1; i >= 1; i--) {
      const cocokCabang = data[i][0] === cabang;
      const cocokId = data[i][1].toString().trim().toUpperCase() === idItem.toString().trim().toUpperCase();
      const cocokNama = data[i][2].toString().trim().toUpperCase() === namaItem.toString().trim().toUpperCase();
      
      // Jika Cabang cocok, DAN (ID cocok atau Nama cocok)
      if (cocokCabang && (cocokId || cocokNama)) {
        sheet.deleteRow(i + 1);
        berhasilDihapus = true;
        break; // Stop loop setelah ketemu dan dihapus
      }
    }
    
    if (berhasilDihapus) {
      return "Berhasil menghapus '" + namaItem + "' dari daftar!";
    } else {
      return "Gagal: Data tidak ditemukan di database.";
    }
    
  } catch(e) {
    return "Gagal: " + e.toString();
  }
}