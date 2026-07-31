const SPREADSHEET_ID = '11rk7ZKuIXxwjMLWqyec2O2l_Cmo5d5QbEtSfA-PNWIc'; 

const SHEET_TRANSAKSI = 'TRANSAKSI';
const SHEET_SETTING = 'SETTING';
const SHEET_USERS = 'USERS';
const SHEET_INVENTARIS = 'INVENTARIS';
const SHEET_PENGELUARAN = 'PENGELUARAN'; 
const SHEET_MASTER_MODAL = 'MASTER_MODAL';

const LOGO_URL = 'https://i.imgur.com/F27OCeM.png'; 

function doGet() {
  setupDatabase();
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Anatra Farm')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  if (!ss.getSheetByName(SHEET_USERS)) {
    const sheet = ss.insertSheet(SHEET_USERS);
    sheet.getRange(1, 1, 2, 3).setValues([
      ['Username', 'Password', 'Role'],
      ['admin', 'admin123', 'Super Admin']
    ]);
    sheet.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET_TRANSAKSI)) {
    const sheet = ss.insertSheet(SHEET_TRANSAKSI);
    sheet.getRange(1, 1, 1, 14).setValues([[
      'ID Transaksi', 'Tanggal', 'Kepada Yth', 'Qty', 'Nama Barang', 
      'Harga Modal @', 'Tagihan', 'Total Modal', 'Total Tagihan', 
      'DP', 'Sisa', 'Status', 'Waktu Input', 'Keterangan Modal'
    ]]);
    sheet.setFrozenRows(1);
  } else {
    // Backward-compatible: pastikan kolom "Keterangan Modal" tersedia
    // untuk sheet TRANSAKSI yang sudah ada sebelumnya.
    ensureTransaksiKeteranganModalColumn_(ss);
  }
  
  if (!ss.getSheetByName(SHEET_INVENTARIS)) {
    const sheet = ss.insertSheet(SHEET_INVENTARIS);
    sheet.getRange(1, 1, 1, 8).setValues([[
      'Tanggal', 'Nama Barang', 'Qty Teks', 'Qty Masuk (Angka)', 'Harga Modal @', 'Qty Keluar (Angka)', 'Keterangan', 'Waktu Input'
    ]]);
    sheet.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET_PENGELUARAN)) {
    const sheet = ss.insertSheet(SHEET_PENGELUARAN);
    sheet.getRange(1, 1, 1, 5).setValues([[
      'ID Pengeluaran', 'Tanggal', 'Keterangan', 'Nominal', 'Waktu Input'
    ]]);
    sheet.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET_MASTER_MODAL)) {
    const sheet = ss.insertSheet(SHEET_MASTER_MODAL);
    sheet.getRange(1, 1, 1, 5).setValues([[
      'Tanggal', 'Jenis Modal', 'Harga Modal Baru', 'Diubah Oleh', 'Waktu Input'
    ]]);
    sheet.setFrozenRows(1);
    const waktuInput = new Date();
    const tanggalHariIni = Utilities.formatDate(waktuInput, "GMT+7", "yyyy-MM-dd");
    sheet.getRange(2, 1, 4, 5).setValues([
      [tanggalHariIni, 'TELUR AYAM REMBAN', 0, 'system', waktuInput],
      [tanggalHariIni, 'TELUR AYAM BUJANG', 0, 'system', waktuInput],
      [tanggalHariIni, 'TELUR PUYUH', 0, 'system', waktuInput],
      [tanggalHariIni, 'TELUR ASIN', 0, 'system', waktuInput]
    ]);
  } else {
    // Backward-compatible: pastikan jenis modal baru (mis. TELUR ASIN) otomatis
    // tersedia untuk sheet MASTER_MODAL yang sudah ada sebelumnya.
    ensureMasterModalJenis_(ss, 'TELUR ASIN');
  }

  if (!ss.getSheetByName(SHEET_SETTING)) {
    const sheet = ss.insertSheet(SHEET_SETTING);
    const dataSetting = [
      ['Key', 'Value'],
      ['NAMA_USAHA', 'Anatra Farm'],
      ['JENIS_USAHA', 'Grosir Telur'],
      ['ALAMAT', 'Jalan Purwodadi, Perm Palm Regency Blok G 19'],
      ['HP', '0823-7175-7899'],
      ['KOTA', 'Pekanbaru'],
      ['PREFIX_NOTA', 'AF'],
      ['NOMOR_TERAKHIR', 0]
    ];
    sheet.getRange(1, 1, dataSetting.length, 2).setValues(dataSetting);
    sheet.setFrozenRows(1);
  }
}

function ensureTransaksiKeteranganModalColumn_(ss) {
  const sheet = ss.getSheetByName(SHEET_TRANSAKSI);
  if (!sheet) return;
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('Keterangan Modal') === -1) {
    sheet.getRange(1, lastCol + 1).setValue('Keterangan Modal');
  }
}

function ensureMasterModalJenis_(ss, jenisBaru) {
  const sheet = ss.getSheetByName(SHEET_MASTER_MODAL);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const search = String(jenisBaru).trim().toUpperCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toUpperCase() === search) return; // sudah ada
  }
  const waktuInput = new Date();
  const tanggalHariIni = Utilities.formatDate(waktuInput, "GMT+7", "yyyy-MM-dd");
  sheet.appendRow([tanggalHariIni, search, 0, 'system', waktuInput]);
}

function login(username, password) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && String(data[i][1]) === String(password)) {
      return { success: true, role: data[i][2], username: username };
    }
  }
  return { success: false, message: 'Username atau Password salah!' };
}

function getInitialData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetSetting = ss.getSheetByName(SHEET_SETTING);
  const settingsData = sheetSetting.getDataRange().getValues();
  let business = {};
  for(let i=1; i<settingsData.length; i++){ business[settingsData[i][0]] = settingsData[i][1]; }
  return { business: business, logoDataUrl: getLogoDataUrlFromImgur_() };
}

function getTransaksiData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_TRANSAKSI);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return JSON.stringify([]);
  
  const grouped = [];
  const map = {};
  for (let i = data.length - 1; i > 0; i--) {
    let id = data[i][0];
    if (!map[id]) {
      map[id] = {
        'ID Transaksi': id,
        'Tanggal': data[i][1],
        'Kepada Yth': data[i][2],
        'Total Modal': 0,
        'Total Jual': 0,
        'DP': data[i][9],
        'Sisa': data[i][10],
        'Status': data[i][11],
        'KeteranganModal': '',
        'DetailBarangArr': [] 
      };
      grouped.push(map[id]);
    }
    map[id]['Total Modal'] += Number(data[i][7]) || 0;
    map[id]['Total Jual'] += Number(data[i][8]) || 0;
    let qty = data[i][3] ? data[i][3] : '';
    let nama = data[i][4] ? data[i][4] : '';
    map[id]['DetailBarangArr'].push(qty + ' ' + nama);
    if (data[i][13]) map[id]['KeteranganModal'] = data[i][13];
  }
  
  grouped.forEach(g => {
    g['DetailBarang'] = g['DetailBarangArr'].join(', ');
  });
  return JSON.stringify(grouped);
}

function getDetailNota(idTransaksi) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_TRANSAKSI);
  const data = sheet.getDataRange().getValues();
  const details = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idTransaksi) {
      let qtyTeks = String(data[i][3]);
      let qNum = Number(qtyTeks.match(/[\d]+([.,][\d]+)?/)?.[0]?.replace(',', '.') || 0);
      details.push({
        qty: data[i][3], 
        qtyStok: qNum,
        namaBarang: data[i][4],
        hargaModal: data[i][5], 
        hargaJual: data[i][6], 
        jumlahHarga: data[i][8] 
      });
    }
  }
  return JSON.stringify(details);
}

function getPengeluaranData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PENGELUARAN);
  if (!sheet) return JSON.stringify([]);
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return JSON.stringify([]);
  
  let result = [];
  for (let i = data.length - 1; i > 0; i--) {
    result.push({
      id: data[i][0],
      tanggal: data[i][1],
      keterangan: data[i][2],
      nominal: Number(data[i][3]) || 0
    });
  }
  return JSON.stringify(result);
}

function savePengeluaran(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_PENGELUARAN);
    const waktuInput = new Date();
    const id = 'OUT-' + new Date().getTime().toString().slice(-6); 
    
    sheet.appendRow([id, payload.tanggal, payload.keterangan, payload.nominal, waktuInput]);
    return { success: true };
  } catch (e) {
    throw new Error('Gagal menyimpan pengeluaran: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

function getLatestModalPrice_(namaBarang) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetInv = ss.getSheetByName(SHEET_INVENTARIS);
  if (!sheetInv) return 0;
  
  const data = sheetInv.getDataRange().getValues();
  const searchName = String(namaBarang).trim().toUpperCase();
  
  for (let i = data.length - 1; i > 0; i--) {
    if (String(data[i][1]).trim().toUpperCase() === searchName && (Number(data[i][4]) || 0) > 0) {
      return Number(data[i][4]) || 0;
    }
  }
  return 0; 
}

function saveTransaksi(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetTrans = ss.getSheetByName(SHEET_TRANSAKSI);
    const sheetInv = ss.getSheetByName(SHEET_INVENTARIS);
    const waktuInput = new Date();
    
    const idTransaksi = generateNotaNumber_();
    const dp = Number(payload.dp) || 0;
    
    let invoiceTotalTagihan = 0;
    payload.items.forEach(item => {
      invoiceTotalTagihan += (Number(item.hargaJual) * Number(item.qtyNumber));
    });
    const sisa = invoiceTotalTagihan - dp;
    
    let status = 'Belum Lunas';
    if (sisa <= 0) { status = 'Lunas'; } 
    else if (dp > 0 && sisa > 0) { status = 'Cashbon'; }

    const rowsToAppend = [];
    const invRowsToAppend = [];
    
    payload.items.forEach((item) => {
      const hargaModalOtomatis = getLatestModalPrice_(item.namaBarang);
      const modalBaris = hargaModalOtomatis * Number(item.qtyNumber);
      const tagihanBaris = Number(item.hargaJual) * Number(item.qtyNumber);
      
      rowsToAppend.push([
        idTransaksi, payload.tanggal, payload.kepada, item.qty, item.namaBarang, 
        hargaModalOtomatis, item.hargaJual, modalBaris, tagihanBaris, 
        dp, sisa, status, waktuInput
      ]);
      
      invRowsToAppend.push([
        payload.tanggal, item.namaBarang, 
        "", 0, 0, item.qtyStok, `Penjualan ke: ${payload.kepada} (Nota: ${idTransaksi})`, waktuInput
      ]);
    });
    
    sheetTrans.getRange(sheetTrans.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    
    if(invRowsToAppend.length > 0) {
      sheetInv.getRange(sheetInv.getLastRow() + 1, 1, invRowsToAppend.length, invRowsToAppend[0].length).setValues(invRowsToAppend);
    }
    
    return { success: true, message: 'Nota berhasil disimpan.', id: idTransaksi };
  } catch (e) {
    throw new Error('Gagal menyimpan: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

function updateTransaksi(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetTrans = ss.getSheetByName(SHEET_TRANSAKSI);
    const sheetInv = ss.getSheetByName(SHEET_INVENTARIS);
    const waktuInput = new Date();
    const idTransaksi = payload.idTransaksi;

    const transData = sheetTrans.getDataRange().getValues();
    const itemsToReturn = [];

    for (let i = 1; i < transData.length; i++) {
      if (transData[i][0] === idTransaksi) {
        let qtyTeks = String(transData[i][3]);
        let qNum = Number(qtyTeks.match(/[\d]+([.,][\d]+)?/)?.[0]?.replace(',', '.') || 0);
        itemsToReturn.push({
          namaBarang: transData[i][4],
          qtyTeks: qtyTeks,
          qtyNum: qNum
        });
      }
    }

    if (itemsToReturn.length === 0) {
      return { success: false, message: 'Data nota tidak ditemukan.' };
    }

    const invReturnRows = [];
    itemsToReturn.forEach(item => {
      invReturnRows.push([
        payload.tanggal, 
        item.namaBarang, 
        item.qtyTeks, 
        0, 
        0, 
        -item.qtyNum, 
        `KOREKSI EDIT Nota: ${idTransaksi} (Pulihkan Stok Lama)`, 
        waktuInput
      ]);
    });

    if (invReturnRows.length > 0) {
      sheetInv.getRange(sheetInv.getLastRow() + 1, 1, invReturnRows.length, invReturnRows[0].length).setValues(invReturnRows);
    }

    for (let i = transData.length - 1; i > 0; i--) {
      if (transData[i][0] === idTransaksi) {
        sheetTrans.deleteRow(i + 1);
      }
    }

    const dp = Number(payload.dp) || 0;
    let invoiceTotalTagihan = 0;
    payload.items.forEach(item => {
      invoiceTotalTagihan += (Number(item.hargaJual) * Number(item.qtyNumber));
    });
    
    const sisa = invoiceTotalTagihan - dp;
    let status = payload.status || 'Belum Lunas';
    if (sisa <= 0) { status = 'Lunas'; } 
    else if (dp > 0 && sisa > 0 && status !== 'Lunas') { status = 'Cashbon'; }

    const rowsToAppend = [];
    const invNewRows = [];
    
    payload.items.forEach((item) => {
      const hargaModalOtomatis = getLatestModalPrice_(item.namaBarang);
      const modalBaris = hargaModalOtomatis * Number(item.qtyNumber);
      const tagihanBaris = Number(item.hargaJual) * Number(item.qtyNumber);
      
      rowsToAppend.push([
        idTransaksi, payload.tanggal, payload.kepada, item.qty, item.namaBarang, 
        hargaModalOtomatis, item.hargaJual, modalBaris, tagihanBaris, 
        dp, sisa, status, waktuInput
      ]);
      
      invNewRows.push([
        payload.tanggal, item.namaBarang, 
        "", 0, 0, item.qtyStok, `Penjualan REVISI ke: ${payload.kepada} (Nota: ${idTransaksi})`, waktuInput
      ]);
    });
    
    sheetTrans.getRange(sheetTrans.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    if(invNewRows.length > 0) {
      sheetInv.getRange(sheetInv.getLastRow() + 1, 1, invNewRows.length, invNewRows[0].length).setValues(invNewRows);
    }

    return { success: true, message: `Nota ${idTransaksi} berhasil diperbarui!` };
  } catch (e) {
    throw new Error("Gagal memperbarui nota: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

function saveBarangMasuk(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetInv = ss.getSheetByName(SHEET_INVENTARIS);
    const waktuInput = new Date();
    
    const invRowsToAppend = payload.items.map(item => [
      payload.tanggal, 
      item.namaBarang, 
      item.qtyTeks, 
      item.qtyMasuk, 
      item.hargaModal, 
      0,               
      payload.keterangan || 'Barang Masuk', 
      waktuInput
    ]);
    
    if(invRowsToAppend.length > 0) {
      sheetInv.getRange(sheetInv.getLastRow() + 1, 1, invRowsToAppend.length, invRowsToAppend[0].length).setValues(invRowsToAppend);
    }
    
    return { success: true, message: 'Stok berhasil ditambahkan.' };
  } catch (e) {
    throw new Error('Gagal menyimpan stok: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

function getDataGudang() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_INVENTARIS);
  if (!sheet) return JSON.stringify({ rekap: [], mutasi: [] });
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return JSON.stringify({ rekap: [], mutasi: [] });
  
  const mutasi = [];
  const rekapMap = {};
  for (let i = data.length - 1; i > 0; i--) { 
    let tgl = data[i][0];
    let nama = String(data[i][1]).trim();
    let namaKey = nama.toUpperCase(); 
    let qtyTeks = data[i][2];
    let masuk = Number(data[i][3]) || 0;
    let modal = Number(data[i][4]) || 0; 
    let keluar = Number(data[i][5]) || 0; 
    let ket = data[i][6]; 
    
    mutasi.push({ tgl, nama, qtyTeks, masuk, modal, keluar, ket });
    if(!rekapMap[namaKey]) {
      rekapMap[namaKey] = { nama: nama, masuk: 0, keluar: 0 };
    }
    rekapMap[namaKey].masuk += masuk;
    rekapMap[namaKey].keluar += keluar;
  }
  
  const rekapArr = Object.values(rekapMap).map(r => {
    r.sisa = r.masuk - r.keluar;
    return r;
  });
  return JSON.stringify({ mutasi: mutasi, rekap: rekapArr });
}

function updateStatusTransaksi(idTransaksi, newStatus) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_TRANSAKSI);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idTransaksi) {
        sheet.getRange(i + 1, 12).setValue(newStatus);
      }
    }
    return true;
  } catch (e) {
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

function simpanPenyesuaianStok(nama, selisihMasuk, selisihKeluar, catatan) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetInv = ss.getSheetByName(SHEET_INVENTARIS);
    const waktuInput = new Date();
    const tanggalHariIni = Utilities.formatDate(waktuInput, "GMT+7", "yyyy-MM-dd");
    
    const rowsToAppend = [];
    if (selisihMasuk !== 0) {
      rowsToAppend.push([
        tanggalHariIni, nama, "Penyesuaian", selisihMasuk, 0, 0, `Penyesuaian Masuk: ${catatan}`, waktuInput
      ]);
    }
    if (selisihKeluar !== 0) {
      rowsToAppend.push([
        tanggalHariIni, nama, "Penyesuaian", 0, 0, selisihKeluar, `Penyesuaian Terpakai: ${catatan}`, waktuInput
      ]);
    }
    
    if (rowsToAppend.length > 0) {
      sheetInv.getRange(sheetInv.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    }
    return true;
  } catch (e) {
    throw new Error('Gagal menyimpan penyesuaian stok: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

function deleteTransaksi(idTransaksi) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetTrans = ss.getSheetByName(SHEET_TRANSAKSI);
    const sheetInv = ss.getSheetByName(SHEET_INVENTARIS);
    const waktuInput = new Date();
    const tanggalHariIni = Utilities.formatDate(waktuInput, "GMT+7", "yyyy-MM-dd");

    const transData = sheetTrans.getDataRange().getValues();
    const itemsToReturn = [];
    let pelanggan = "";

    for (let i = 1; i < transData.length; i++) {
      if (transData[i][0] === idTransaksi) {
        pelanggan = transData[i][2];
        itemsToReturn.push({
          namaBarang: transData[i][4],
          qtyTeks: transData[i][3]
        });
      }
    }

    if (itemsToReturn.length === 0) {
      return { success: false, message: "Data transaksi tidak ditemukan atau sudah dihapus." };
    }

    const invRowsToAppend = [];
    itemsToReturn.forEach(item => {
      const qNum = Number(item.qtyTeks.match(/[\d]+([.,][\d]+)?/)?.[0]?.replace(',', '.') || 0);
      invRowsToAppend.push([
        tanggalHariIni, 
        item.namaBarang, 
        item.qtyTeks, 
        0, 
        0, 
        qNum, 
        `PEMBATALAN/HAPUS Nota: ${idTransaksi} (Pelanggan: ${pelanggan})`, 
        waktuInput
      ]);
    });

    if (invRowsToAppend.length > 0) {
      invRowsToAppend.forEach(r => r[5] = -r[5]); 
      sheetInv.getRange(sheetInv.getLastRow() + 1, 1, invRowsToAppend.length, invRowsToAppend[0].length).setValues(invRowsToAppend);
    }

    for (let i = transData.length - 1; i > 0; i--) {
      if (transData[i][0] === idTransaksi) {
        sheetTrans.deleteRow(i + 1);
      }
    }

    return { success: true, message: `Transaksi ${idTransaksi} berhasil dihapus!` };
  } catch (e) {
    throw new Error("Gagal menghapus transaksi di Cloud: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

/* ========================================================
   MASTER MENU MODAL (Harga Pokok Telur) + Log Perubahan Harga
   ======================================================== */

function getMasterModalData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_MASTER_MODAL);
  if (!sheet) return JSON.stringify({ current: {}, log: [] });

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return JSON.stringify({ current: {}, log: [] });

  const current = {};
  const log = [];
  for (let i = 1; i < data.length; i++) {
    const jenis = String(data[i][1]).trim().toUpperCase();
    const harga = Number(data[i][2]) || 0;
    log.push({
      tanggal: data[i][0],
      jenis: jenis,
      harga: harga,
      user: data[i][3],
      waktuInput: data[i][4]
    });
    current[jenis] = harga; // baris paling bawah = harga yang berlaku saat ini
  }
  log.reverse(); // riwayat terbaru tampil di atas

  return JSON.stringify({ current: current, log: log });
}

function saveMasterModalHarga(jenis, harga, username) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_MASTER_MODAL);
    const waktuInput = new Date();
    const tanggalHariIni = Utilities.formatDate(waktuInput, "GMT+7", "yyyy-MM-dd");

    sheet.appendRow([
      tanggalHariIni,
      String(jenis).trim().toUpperCase(),
      Number(harga) || 0,
      username || '-',
      waktuInput
    ]);
    return { success: true };
  } catch (e) {
    throw new Error('Gagal menyimpan harga modal: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

function getLatestMasterModalPrice_(jenis) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_MASTER_MODAL);
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  const searchJenis = String(jenis).trim().toUpperCase();
  for (let i = data.length - 1; i > 0; i--) {
    if (String(data[i][1]).trim().toUpperCase() === searchJenis) {
      return Number(data[i][2]) || 0;
    }
  }
  return 0;
}

// Menghitung ulang Total Modal sebuah Nota berdasarkan harga Master Modal
// (Bujang / Remban) yang dipilih user di kolom Modal pada Laporan Keuntungan.
// Hanya baris barang yang mengandung kata "TELUR" yang akan dihitung ulang.
function updateModalLaporan(idTransaksi, jenisModal) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_TRANSAKSI);
    const data = sheet.getDataRange().getValues();

    const hargaBaru = getLatestMasterModalPrice_(jenisModal);
    if (hargaBaru <= 0) {
      throw new Error('Harga Master Modal untuk "' + jenisModal + '" belum diatur. Silakan atur terlebih dahulu di menu Master Modal.');
    }

    let totalModalBaru = 0;
    let found = false;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idTransaksi) {
        const namaBarang = String(data[i][4]).toUpperCase();
        if (namaBarang.indexOf('TELUR') !== -1) {
          const qtyTeks = String(data[i][3]);
          const qNum = Number(qtyTeks.match(/[\d]+([.,][\d]+)?/)?.[0]?.replace(',', '.') || 0);
          const modalBaris = hargaBaru * qNum;
          sheet.getRange(i + 1, 6).setValue(hargaBaru);   // Harga Modal @
          sheet.getRange(i + 1, 8).setValue(modalBaris);  // Total Modal (per baris)
          totalModalBaru += modalBaris;
          found = true;
        } else {
          totalModalBaru += Number(data[i][7]) || 0;
        }
      }
    }

    if (!found) {
      throw new Error('Tidak ditemukan item Telur pada nota ini.');
    }

    return { success: true, totalModalBaru: totalModalBaru, jenisModal: jenisModal };
  } catch (e) {
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

// Submenu "Atur Modal" pada Laporan Keuntungan.
// Mendukung 2 mode:
//  - mode 'racik'  : rincian = [{jenis, qty}, ...] -> total = SUM(hargaMasterModal(jenis) * qty)
//  - mode 'manual' : nominalManual + keterangan (wajib diisi), untuk order per-butir campuran
// Total Modal hasil hitungan ditulis ke baris PERTAMA nota tsb (kolom Total Modal),
// baris lain di nota yang sama di-nol-kan supaya SUM Total Modal per nota tetap akurat,
// tanpa mengubah data Qty/Nama Barang asli (aman untuk cetak nota & data gudang).
function simpanModalLaporan(idTransaksi, payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_TRANSAKSI);
    ensureTransaksiKeteranganModalColumn_(ss);
    const data = sheet.getDataRange().getValues();

    let totalModalBaru = 0;
    let keterangan = '';

    if (payload.mode === 'manual') {
      totalModalBaru = Number(payload.nominalManual) || 0;
      const catatan = String(payload.keterangan || '').trim();
      if (totalModalBaru <= 0) throw new Error('Nominal modal manual harus lebih dari 0.');
      if (!catatan) throw new Error('Keterangan wajib diisi untuk modal manual (mis. rincian per butir).');
      keterangan = 'Manual: ' + catatan;
    } else {
      const rincian = payload.rincian || [];
      if (rincian.length === 0) throw new Error('Rincian racikan modal masih kosong.');
      const parts = [];
      rincian.forEach(r => {
        const jenis = String(r.jenis || '').trim().toUpperCase();
        if (!jenis) throw new Error('Pilih jenis telur untuk setiap baris racikan.');
        const qty = Number(r.qty) || 0;
        if (qty <= 0) throw new Error('Qty racikan "' + labelSingkatJenisModal_(jenis) + '" harus lebih dari 0.');
        const harga = getLatestMasterModalPrice_(jenis);
        if (harga <= 0) throw new Error('Harga Master Modal untuk "' + labelSingkatJenisModal_(jenis) + '" belum diatur. Silakan atur dahulu di menu Master Modal.');
        totalModalBaru += harga * qty;
        parts.push(labelSingkatJenisModal_(jenis) + ' ' + qty);
      });
      const catatanTambahan = String(payload.keterangan || '').trim();
      keterangan = 'Racikan: ' + parts.join(' + ') + (catatanTambahan ? ' (' + catatanTambahan + ')' : '');
    }

    let firstRowIndex = -1;
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idTransaksi) {
        found = true;
        if (firstRowIndex === -1) firstRowIndex = i;
      }
    }
    if (!found) throw new Error('Transaksi tidak ditemukan.');

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idTransaksi) {
        if (i === firstRowIndex) {
          sheet.getRange(i + 1, 8).setValue(totalModalBaru);   // Total Modal
          sheet.getRange(i + 1, 14).setValue(keterangan);      // Keterangan Modal
        } else {
          sheet.getRange(i + 1, 8).setValue(0);
          sheet.getRange(i + 1, 14).setValue('');
        }
      }
    }

    return { success: true, totalModalBaru: totalModalBaru, keterangan: keterangan };
  } catch (e) {
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

function labelSingkatJenisModal_(jenis) {
  const j = String(jenis).trim().toUpperCase();
  if (j === 'TELUR AYAM BUJANG') return 'Bujang';
  if (j === 'TELUR AYAM REMBAN') return 'Remban';
  if (j === 'TELUR ASIN') return 'Telur Asin';
  if (j === 'TELUR PUYUH') return 'Telur Puyuh';
  return String(jenis).split(' ').map(w => w.charAt(0) + w.slice(1).toLowerCase()).join(' ');
}

function generateNotaNumber_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SETTING);
  const values = sheet.getDataRange().getValues();
  let prefix = 'AF'; let lastNumIndex = -1; let lastNum = 0;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === 'PREFIX_NOTA') prefix = values[i][1];
    if (values[i][0] === 'NOMOR_TERAKHIR') { lastNumIndex = i + 1; lastNum = Number(values[i][1]); }
  }
  if (lastNumIndex === -1) throw new Error("Gagal membaca Data Setting.");
  const nextNum = lastNum + 1;
  sheet.getRange(lastNumIndex, 2).setValue(nextNum);
  const currentYear = new Date().getFullYear();
  return `${prefix}-${currentYear}-${String(nextNum).padStart(4, '0')}`;
}

function getLogoDataUrlFromImgur_() {
  try {
    const response = UrlFetchApp.fetch(LOGO_URL, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      return `data:${response.getBlob().getContentType()};base64,${Utilities.base64Encode(response.getBlob().getBytes())}`;
    }
  } catch (err) {
    Logger.log("Gagal mengambil logo: " + err.toString());
  }
  return '';
}
