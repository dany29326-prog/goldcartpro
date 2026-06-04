// ========== GOLDCART PRO - GOOGLE APPS SCRIPT (BACKEND LENGKAP) ==========
// PASTE kode ini ke Google Apps Script Editor
// Ganti YOUR_SHEET_ID_HERE dengan ID Google Sheet Anda

const SHEET_ID = '1UC4emkdVJO7323fOnAbKGtI0RzHAZw4V5wjLsPA0nO0'; // Ganti dengan ID Google Sheet Anda
const API_TOKEN = 'GOLDCART_SECRET_2024'; // Token untuk keamanan akses API

// ========== DO GET (UNTUK MEMBACA DATA DARI FRONTEND) ==========
function doGet(e) {
  const token = e?.parameter?.token || '';
  if (token !== API_TOKEN) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Unauthorized' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const action = e?.parameter?.action || '';
  const callback = e?.parameter?.callback || '';
  
  let result;
  
  if (action === 'getAllData') {
    result = getAllDataResponse();
  } else if (action === 'getAntam') {
    result = fetchAntamAPI();
  } else {
    result = {
      status: 'ok',
      message: 'GoldCart Pro API Running',
      produk: [],
      timestamp: Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')
    };
  }
  
  if (callback) {
    return ContentService.createTextOutput(`${callback}(${JSON.stringify(result)})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========== DO POST (UNTUK MENYIMPAN DATA DARI FRONTEND) ==========
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  
  try {
    let data = {};
    
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.data) {
      data = JSON.parse(e.parameter.data);
    } else if (e.parameter) {
      data = e.parameter;
    }
    
    if (data.token !== API_TOKEN) {
      output.setContent(JSON.stringify({ status: 'error', message: 'Unauthorized' }));
      return output;
    }
    
    const action = data.action;
    
    // Catat aktivitas
    logActivity('POST_REQUEST', action, JSON.stringify(data).substring(0, 200));
    
    // Handler untuk berbagai action
    if (action === 'saveTransaction') {
      const result = saveTransaction(data.transaction);
      output.setContent(JSON.stringify(result));
    } else if (action === 'saveExpense') {
      const result = saveExpense(data.expense);
      output.setContent(JSON.stringify(result));
    } else if (action === 'saveProduct') {
      const result = saveProduct(data.product);
      output.setContent(JSON.stringify(result));
    } else if (action === 'savePetugas') {
      const result = savePetugas(data.petugas);
      output.setContent(JSON.stringify(result));
    } else if (action === 'updateSettings') {
      const result = updateSettings(data.settings);
      output.setContent(JSON.stringify(result));
    } else if (action === 'updateHargaEmas') {
      const result = updateHargaEmas(data.kadar, data.harga_jual, data.harga_beli);
      output.setContent(JSON.stringify(result));
    } else if (action === 'payDebt') {
      const result = payDebt(data.nama);
      output.setContent(JSON.stringify(result));
    } else if (action === 'deleteExpense') {
      const result = deleteExpense(data.id);
      output.setContent(JSON.stringify(result));
    } else if (action === 'deletePetugas') {
      const result = deletePetugas(data.id);
      output.setContent(JSON.stringify(result));
    } else if (action === 'deleteProduct') {
      const result = deleteProduct(data.id);
      output.setContent(JSON.stringify(result));
    } else {
      output.setContent(JSON.stringify({ status: 'ok', message: 'API ready' }));
    }
    
  } catch(error) {
    logActivity('ERROR', 'doPost', error.stack || error.toString());
    output.setContent(JSON.stringify({ status: 'error', message: error.toString() }));
  }
  
  return output;
}

// ========== SETUP SHEET (JALANKAN FUNGSI INI SATU KALI) ==========
function setupGoldCartSheets() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    const sheetsToCreate = [
      { name: 'transaksi', headers: getTransaksiHeaders() },
      { name: 'pelanggan', headers: getPelangganHeaders() },
      { name: 'harga_emas', headers: getHargaEmasHeaders() },
      { name: 'master_produk', headers: getMasterProdukHeaders() },
      { name: 'pengeluaran', headers: getPengeluaranHeaders() },
      { name: 'petugas', headers: getPetugasHeaders() },
      { name: 'settings', headers: getSettingsHeaders() },
      { name: 'log_aktivitas', headers: getLogHeaders() }
    ];
    
    sheetsToCreate.forEach(sheet => {
      createOrUpdateSheet(ss, sheet.name, sheet.headers);
    });
    
    insertDefaultData(ss);
    
    Logger.log('✅ Semua sheet berhasil dibuat!');
    return { status: 'success', message: 'Setup selesai' };
    
  } catch(e) {
    Logger.log('❌ Error: ' + e.toString());
    return { status: 'error', message: e.toString() };
  }
}

function createOrUpdateSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    Logger.log(`✅ Sheet "${sheetName}" dibuat`);
  } else {
    const existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    let needsUpdate = false;
    
    for (let i = 0; i < headers.length; i++) {
      if (existingHeaders[i] !== headers[i]) {
        needsUpdate = true;
        break;
      }
    }
    
    if (needsUpdate) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      Logger.log(`🔄 Header sheet "${sheetName}" diperbarui`);
    }
  }
  
  // Format sheet
  sheet.setFrozenRows(1);
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#D4AF37');
  headerRange.setFontColor('#0C0E12');
}

function insertDefaultData(ss) {
  // Harga Emas Default
  let hargaEmasSheet = ss.getSheetByName('harga_emas');
  if (hargaEmasSheet.getLastRow() === 1) {
    const defaultHarga = [
      ['24', 1200000, 1050000, new Date()],
      ['22', 1100000, 960000, new Date()],
      ['18', 900000, 780000, new Date()],
      ['17', 850000, 730000, new Date()]
    ];
    defaultHarga.forEach(row => hargaEmasSheet.appendRow(row));
  }
  
  // Petugas Default
  let petugasSheet = ss.getSheetByName('petugas');
  if (petugasSheet.getLastRow() === 1) {
    petugasSheet.appendRow(['P001', 'ADMIN', true, 'Owner', new Date(), '']);
  }
  
  // Settings Default
  let settingsSheet = ss.getSheetByName('settings');
  if (settingsSheet.getLastRow() === 1) {
    const defaultSettings = [
      ['nama_toko', 'GoldCart Pro'],
      ['alamat', 'Jl. Perhiasan No. 1'],
      ['telepon', '08123456789'],
      ['footer_nota', 'Terima Kasih atas Kunjungan Anda 💎'],
      ['last_backup', Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')]
    ];
    defaultSettings.forEach(row => settingsSheet.appendRow(row));
  }
}

// ========== HEADER DEFINITIONS ==========
function getTransaksiHeaders() {
  return ['id', 'jenis', 'tanggal', 'waktu', 'petugas_id', 'petugas_nama', 'pelanggan_id', 'pelanggan_nama', 'wa', 'items', 'subtotal', 'ongkir_potongan', 'total', 'metode_bayar', 'status', 'uang_muka', 'sisa_tagihan', 'catatan'];
}
function getPelangganHeaders() {
  return ['id', 'nama', 'wa', 'alamat', 'total_jual', 'total_beli', 'total_piutang', 'terakhir_transaksi', 'created_at', 'updated_at'];
}
function getHargaEmasHeaders() {
  return ['kadar', 'harga_jual_per_gram', 'harga_beli_per_gram', 'updated_at'];
}
function getMasterProdukHeaders() {
  return ['id', 'nama_produk', 'kategori', 'kadar_emas', 'berat_gram', 'upah_maker', 'stok', 'foto_url', 'is_active'];
}
function getPengeluaranHeaders() {
  return ['id', 'tanggal', 'waktu', 'kategori', 'jumlah', 'keterangan', 'petugas_nama'];
}
function getPetugasHeaders() {
  return ['id', 'nama', 'is_active', 'jabatan', 'created_at', 'last_login'];
}
function getSettingsHeaders() {
  return ['key', 'value', 'last_updated'];
}
function getLogHeaders() {
  return ['timestamp', 'type', 'action', 'detail', 'user'];
}

// ========== GET ALL DATA RESPONSE ==========
function getAllDataResponse() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  return {
    status: 'ok',
    data: {
      transaksi: readSheetData(ss, 'transaksi'),
      pelanggan: readSheetData(ss, 'pelanggan'),
      harga_emas: readSheetData(ss, 'harga_emas'),
      master_produk: readSheetData(ss, 'master_produk'),
      pengeluaran: readSheetData(ss, 'pengeluaran'),
      petugas: readSheetData(ss, 'petugas'),
      settings: readSettings(ss)
    }
  };
}

function readSheetData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  
  const headers = rows[0];
  const data = [];
  
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0] || row[0].toString().trim() === '') continue;
    
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      let value = row[j];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      }
      obj[headers[j]] = value;
    }
    data.push(obj);
  }
  return data;
}

function readSettings(ss) {
  const sheet = ss.getSheetByName('settings');
  if (!sheet) return { nama_toko: 'GoldCart Pro' };
  
  const rows = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < rows.length; i++) {
    const key = rows[i][0];
    const value = rows[i][1];
    if (key && value) settings[key] = value;
  }
  return settings;
}

// ========== SAVE TRANSACTION ==========
function saveTransaction(trx) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('transaksi');
    
    if (!sheet) {
      sheet = ss.insertSheet('transaksi');
      sheet.appendRow(getTransaksiHeaders());
    }
    
    const rowData = [
      trx.id || 'TRX_' + Date.now(),
      trx.jenis || 'jual',
      trx.tanggal || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd'),
      trx.waktu || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss'),
      trx.petugas_id || 'P001',
      trx.petugas_nama || 'ADMIN',
      trx.pelanggan_id || '',
      trx.pelanggan_nama || 'Umum',
      trx.wa || '',
      JSON.stringify(trx.items || []),
      trx.subtotal || trx.total || 0,
      trx.ongkir_potongan || 0,
      trx.total || 0,
      trx.metode || trx.metode_bayar || 'tunai',
      trx.status || 'lunas',
      trx.uang_muka || 0,
      trx.sisa_tagihan || 0,
      trx.catatan || ''
    ];
    
    sheet.appendRow(rowData);
    updatePelanggan(ss, trx);
    logActivity('SAVE_TRANSACTION', trx.id, `${trx.jenis} - ${trx.pelanggan_nama} - Rp${trx.total}`);
    
    return { status: 'ok', message: 'Transaksi tersimpan', id: trx.id };
  } catch(e) {
    logActivity('ERROR', 'saveTransaction', e.toString());
    return { status: 'error', message: e.toString() };
  }
}

function updatePelanggan(ss, trx) {
  try {
    let sheet = ss.getSheetByName('pelanggan');
    if (!sheet) {
      sheet = ss.insertSheet('pelanggan');
      sheet.appendRow(getPelangganHeaders());
    }
    
    const nama = trx.pelanggan_nama || 'Umum';
    if (nama === 'Umum') return;
    
    const data = sheet.getDataRange().getValues();
    let foundIndex = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === nama) {
        foundIndex = i;
        break;
      }
    }
    
    const today = new Date().toISOString().slice(0,10);
    
    if (foundIndex !== -1) {
      let totalJual = Number(data[foundIndex][4]) || 0;
      let totalBeli = Number(data[foundIndex][5]) || 0;
      if (trx.jenis === 'jual') totalJual += trx.total;
      else if (trx.jenis === 'beli') totalBeli += trx.total;
      
      sheet.getRange(foundIndex + 1, 4).setValue(today);
      sheet.getRange(foundIndex + 1, 5).setValue(totalJual);
      sheet.getRange(foundIndex + 1, 6).setValue(totalBeli);
    } else {
      const newId = 'CUST_' + Date.now();
      const totalJual = trx.jenis === 'jual' ? trx.total : 0;
      const totalBeli = trx.jenis === 'beli' ? trx.total : 0;
      sheet.appendRow([newId, nama, trx.wa || '', '', totalJual, totalBeli, 0, today, today, today]);
    }
  } catch(e) {
    logActivity('ERROR', 'updatePelanggan', e.toString());
  }
}

// ========== SAVE EXPENSE ==========
function saveExpense(exp) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('pengeluaran');
    
    if (!sheet) {
      sheet = ss.insertSheet('pengeluaran');
      sheet.appendRow(getPengeluaranHeaders());
    }
    
    sheet.appendRow([
      exp.id || 'EXP_' + Date.now(),
      exp.tanggal || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd'),
      exp.waktu || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss'),
      exp.kategori || 'Operasional',
      Number(exp.jumlah) || 0,
      exp.keterangan || '',
      exp.petugas_nama || 'system'
    ]);
    
    logActivity('SAVE_EXPENSE', exp.id, 'Jumlah: Rp' + exp.jumlah);
    return { status: 'ok', message: 'Pengeluaran tersimpan' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

// ========== SAVE PRODUCT ==========
function saveProduct(product) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('master_produk');
    
    if (!sheet) {
      sheet = ss.insertSheet('master_produk');
      sheet.appendRow(getMasterProdukHeaders());
    }
    
    const data = sheet.getDataRange().getValues();
    let foundIndex = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === product.id) {
        foundIndex = i;
        break;
      }
    }
    
    const rowData = [
      product.id,
      product.nama_produk,
      product.kategori || 'cincin',
      Number(product.kadar_emas) || 22,
      Number(product.berat_gram) || 0,
      Number(product.upah_maker) || 0,
      Number(product.stok) || 0,
      product.foto_url || '',
      product.is_active === true || product.is_active === 'TRUE'
    ];
    
    if (foundIndex !== -1) {
      sheet.getRange(foundIndex + 1, 1, 1, 9).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    
    return { status: 'ok', message: 'Produk tersimpan' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

// ========== SAVE PETUGAS ==========
function savePetugas(p) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('petugas');
    
    if (!sheet) {
      sheet = ss.insertSheet('petugas');
      sheet.appendRow(getPetugasHeaders());
    }
    
    const data = sheet.getDataRange().getValues();
    let foundIndex = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === p.id) {
        foundIndex = i;
        break;
      }
    }
    
    const now = new Date();
    
    if (foundIndex !== -1) {
      sheet.getRange(foundIndex + 1, 2).setValue(p.nama);
      sheet.getRange(foundIndex + 1, 4).setValue(p.jabatan || 'Kasir');
    } else {
      sheet.appendRow([p.id, p.nama, true, p.jabatan || 'Kasir', now, '']);
    }
    
    return { status: 'ok', message: 'Petugas tersimpan' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

// ========== UPDATE SETTINGS ==========
function updateSettings(settings) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('settings');
    
    if (!sheet) {
      sheet = ss.insertSheet('settings');
      sheet.appendRow(getSettingsHeaders());
    }
    
    Object.entries(settings).forEach(([key, value]) => {
      const data = sheet.getDataRange().getValues();
      let found = false;
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(value.toString());
          sheet.getRange(i + 1, 3).setValue(new Date());
          found = true;
          break;
        }
      }
      
      if (!found) {
        sheet.appendRow([key, value.toString(), new Date()]);
      }
    });
    
    return { status: 'ok', message: 'Settings updated' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

// ========== UPDATE HARGA EMAS ==========
function updateHargaEmas(kadar, hargaJual, hargaBeli) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('harga_emas');
    if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan' };
    
    const data = sheet.getDataRange().getValues();
    let found = false;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === kadar.toString()) {
        sheet.getRange(i + 1, 2).setValue(Number(hargaJual));
        sheet.getRange(i + 1, 3).setValue(Number(hargaBeli));
        sheet.getRange(i + 1, 4).setValue(new Date());
        found = true;
        break;
      }
    }
    
    if (!found) {
      sheet.appendRow([kadar, Number(hargaJual), Number(hargaBeli), new Date()]);
    }
    
    logActivity('UPDATE_HARGA_EMAS', kadar, `Jual: ${hargaJual}, Beli: ${hargaBeli}`);
    return { status: 'ok', message: 'Harga emas diperbarui' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

// ========== PAY DEBT (PELUNASAN HUTANG) ==========
function payDebt(nama) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('transaksi');
    if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan' };
    
    const rows = sheet.getDataRange().getValues();
    let updatedCount = 0;
    
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (row[7] === nama && row[14] === 'cicil' && Number(row[16]) > 0) {
        sheet.getRange(i + 1, 15).setValue('lunas');
        sheet.getRange(i + 1, 17).setValue(0);
        updatedCount++;
      }
    }
    
    logActivity('PAY_DEBT', nama, 'Melunasi ' + updatedCount + ' transaksi');
    return { status: 'ok', message: 'Pelunasan berhasil', count: updatedCount };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

// ========== DELETE FUNCTIONS ==========
function deleteExpense(id) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('pengeluaran');
    if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan' };
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === id.toString()) {
        sheet.deleteRow(i + 1);
        logActivity('DELETE_EXPENSE', id, 'Berhasil menghapus pengeluaran');
        return { status: 'ok', message: 'Pengeluaran dihapus' };
      }
    }
    return { status: 'error', message: 'ID tidak ditemukan' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function deletePetugas(id) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('petugas');
    if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan' };
    
    if (id === 'P001') {
      return { status: 'error', message: 'Tidak bisa menghapus owner utama' };
    }
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === id.toString()) {
        sheet.getRange(i + 1, 3).setValue(false);
        logActivity('DELETE_PETUGAS', id, 'Petugas dinonaktifkan');
        return { status: 'ok', message: 'Petugas dinonaktifkan' };
      }
    }
    return { status: 'error', message: 'Petugas tidak ditemukan' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function deleteProduct(id) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('master_produk');
    if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan' };
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === id.toString()) {
        sheet.deleteRow(i + 1);
        logActivity('DELETE_PRODUCT', id, 'Produk dihapus');
        return { status: 'ok', message: 'Produk dihapus' };
      }
    }
    return { status: 'error', message: 'Produk tidak ditemukan' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

// ========== LOG ACTIVITY ==========
function logActivity(type, action, detail) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('log_aktivitas');
    
    if (!sheet) {
      sheet = ss.insertSheet('log_aktivitas');
      sheet.appendRow(getLogHeaders());
    }
    
    sheet.appendRow([new Date(), type, action, detail, Session.getActiveUser().getEmail() || 'system']);
  } catch(e) {}
}

// ========== BACKUP DATA ==========
function createBackup() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const backupName = 'GoldCart_Backup_' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd_HH-mm-ss');
    const backup = SpreadsheetApp.create(backupName);
    
    const sheets = ss.getSheets();
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const newSheet = backup.insertSheet(sheetName);
      const data = sheet.getDataRange().getValues();
      if (data.length > 0) {
        newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      }
    });
    
    logActivity('BACKUP', 'createBackup', 'Backup created: ' + backup.getUrl());
    return { status: 'ok', url: backup.getUrl() };
  } catch(e) {
    logActivity('ERROR', 'createBackup', e.toString());
    return { status: 'error', message: e.toString() };
  }
}

// ========== FUNGSI UNTUK MENJALANKAN SETUP ==========
function runSetup() {
  const result = setupGoldCartSheets();
  Logger.log(result);
}

function checkSheetsStatus() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheets = ss.getSheets();
  const result = [];
  sheets.forEach(sheet => {
    result.push({
      name: sheet.getName(),
      rows: sheet.getLastRow(),
      columns: sheet.getLastColumn()
    });
  });
  Logger.log(result);
  return result;
}

// ========== FETCH ANTAM API ==========
function fetchAntamAPI() {
  try {
    // Karena API Antam publik sering down/dijaga ketat, kita gunakan PAXG (Aset Kripto Emas Fisik) dari Indodax.
    // 1 PAXG = 1 Troy Ounce Emas = 31.103 gram. Sangat akurat dan live 24/7.
    const url = 'https://indodax.com/api/ticker/paxgidr';
    const options = { muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(url, options);
    if(response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      const paxgIdr = parseInt(data.ticker.last);
      const perGram = Math.round(paxgIdr / 31.103);
      return { status: 'ok', data: [{ sell: perGram }] };
    }
    return { status: 'error', message: 'HTTP ' + response.getResponseCode() };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}