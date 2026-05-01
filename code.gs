/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 * RentMotor Pro - Backend API (Google Apps Script)
 */

const SPREADSHEET_ID = (function() {
  try { return SpreadsheetApp.getActiveSpreadsheet().getId(); } catch (e) { return ""; }
})();
const FOLDER_ID = "YOUR_DRIVE_FOLDER_ID"; // Ganti dengan ID folder Google Drive untuk foto

function getSS() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss && ss.getId()) return ss;
  } catch (e) {
    console.warn("getActiveSpreadsheet failed, trying by ID...");
  }
  
  try {
    if (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID && SPREADSHEET_ID !== "") {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    }
  } catch (e) {
    console.error("Critical error in getSS: " + e.toString());
  }
  return null;
}

function getSheet(name) {
  const ss = getSS();
  if (!ss) throw new Error("Spreadsheet tidak terdeteksi. Hubungkan Spreadsheet atau set SPREADSHEET_ID.");
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    ensureHeaders();
    sheet = ss.getSheetByName(name);
  }
  return sheet;
}

function doGet() {
  try {
    ensureHeaders();
  } catch (e) {
    console.error("Initialization error:", e);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('RentMotor Pro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Main API Dispatcher
 * Handles requests from the client. Supports both single JSON string or (action, data) arguments.
 */
function api(request, optionalData) {
  let action, data;
  
  try {
    // Robust argument parsing
    if (typeof request === 'object' && request !== null && request.action) {
      // Case: Single object {action, data}
      action = request.action;
      data = request.data;
    } else if (typeof request === 'string' && optionalData !== undefined) {
      // Case: separate arguments (action, data)
      action = request;
      data = optionalData;
    } else if (typeof request === 'string' && (request.startsWith('{') || request.startsWith('['))) {
      // Case: JSON string
      const parsed = JSON.parse(request);
      action = parsed.action;
      data = parsed.data;
    } else if (typeof request === 'string') {
      // Case: action string only
      action = request;
      data = optionalData || null;
    }

    if (!action) throw new Error('Action parameter is missing');

    console.log(`API Executing: ${action}`, data ? "with data" : "no data");

    switch (action) {
      case 'healthCheck': return { success: true, message: 'Server Online', timestamp: new Date().toISOString(), ssConnected: !!getSS() };
      case 'getDashboardData': return getDashboardData();
      case 'getUnits': return getUnits();
      case 'getCashflow': return { success: true, data: getDataFromSheet('Cashflow') };
      case 'getExpenses': return getExpenses();
      case 'getTransactions': return getTransactions();
      case 'getSettings': return getSettings();
      case 'saveTransaction': return saveTransaction(data);
      case 'completeTransaction': return completeTransaction(data);
      case 'recordPelunasan': return recordPelunasan(data);
      case 'saveExpense': return saveExpense(data);
      case 'saveCash': return saveCash(data);
      case 'saveUnit': return saveUnit(data);
      case 'saveSettings': return saveSettings(data);
      case 'uploadPhoto': return uploadPhoto(data);
      default: throw new Error(`Action "${action}" not recognized`);
    }
  } catch (e) {
    console.error(`API Error [${action || 'unknown'}]: `, e.toString());
    return { success: false, error: e.toString(), action: action };
  }
}

function getDataFromSheet(name) {
  const sheet = getSheet(name);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length === 0 || (data.length === 1 && data[0][0] === "")) return [];
  
  const headerRow = data.shift();
  if (!headerRow) return [];
  
  const headers = headerRow.map(h => h ? h.toString().toLowerCase().trim().replace(/ /g, '_') : '');
  return data.map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      if (h) obj[h] = row[i];
    });
    return obj;
  });
}

// --- DASHBOARD ---
function getDashboardData() {
  try {
    const transactions = getDataFromSheet('Transactions');
    const units = getDataFromSheet('Units');
    const cashflow = getDataFromSheet('Cashflow');
    
    const today = new Date().toDateString();
    
    const revenue = cashflow.filter(c => c.tipe === 'Masuk').reduce((acc, c) => acc + Number(c.nominal || 0), 0);
    const totalExpenses = cashflow.filter(c => c.tipe === 'Keluar').reduce((acc, c) => acc + Number(c.nominal || 0), 0);
    
    // Calculate chart data (last 7 days)
    const chartData = [];
    for(let i = 6; i >= 0; i--) {
      const d = new Date();
      d.setDate(d.getDate() - i);
      const dateStr = d.toDateString();
      const dailyRevenue = cashflow.filter(c => c.tipe === 'Masuk' && new Date(c.tanggal).toDateString() === dateStr)
                                 .reduce((acc, c) => acc + Number(c.nominal || 0), 0);
      chartData.push({ label: d.toLocaleDateString('id-ID', { day: 'numeric', month: 'short' }), value: dailyRevenue });
    }

    const returnsToday = transactions.filter(t => t.trans_status === 'Aktif' && t.tgl_selesai && new Date(t.tgl_selesai).toDateString() === today);
    const deliveriesToday = transactions.filter(t => t.trans_status === 'Aktif' && t.tgl_mulai && new Date(t.tgl_mulai).toDateString() === today);

    const stats = {
      revenue: revenue,
      expenses: totalExpenses,
      profit: revenue - totalExpenses,
      cashBalance: cashflow.filter(c => c.metode === 'Cash').reduce((acc, c) => acc + (c.tipe === 'Masuk' ? Number(c.nominal || 0) : -Number(c.nominal || 0)), 0),
      transferBalance: cashflow.filter(c => !c.metode || c.metode.toLowerCase().includes('transfer')).reduce((acc, c) => acc + (c.tipe === 'Masuk' ? Number(c.nominal || 0) : -Number(c.nominal || 0)), 0),
      totalUnits: units.length,
      availableUnits: units.filter(u => u.status === 'Tersedia').length,
      rentedUnits: units.filter(u => u.status === 'Disewa').length,
      serviceUnits: units.filter(u => u.status === 'Servis').length,
      returnsToday: returnsToday,
      deliveriesToday: deliveriesToday,
      chartData: chartData
    };
    
    return { success: true, stats };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getExpenses() {
  return { success: true, data: getDataFromSheet('Expenses') };
}

function saveSettings(data) {
  const sheet = getSheet('Settings');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['key', 'value']);
  }
  const rows = sheet.getDataRange().getValues();
  const entries = Object.entries(data);
  
  entries.forEach(([key, val]) => {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      const rowKey = rows[i][0] ? rows[i][0].toString().toLowerCase().trim() : '';
      if (rowKey === key.toLowerCase().trim()) {
        sheet.getRange(i + 1, 2).setValue(val);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, val]);
    }
  });
  return { success: true };
}

function ensureHeaders() {
  const ss = getSS();
  if (!ss) {
    console.error("Cannot ensure headers: No Spreadsheet connected.");
    return;
  }
  const sheets = {
    'Transactions': ['id', 'customer_name', 'wa_number', 'ktp_number', 'alamat_antar', 'alamat_ambil', 'id_unit', 'tgl_mulai', 'tgl_selesai', 'durasi_hari', 'overtime_jam', 'biaya_sewa', 'biaya_overtime', 'biaya_antar', 'biaya_ambil', 'biaya_addon', 'addon_list', 'total', 'dp', 'metode_dp', 'lunas_amt', 'lunas_method', 'pay_status', 'trans_status', 'photo'],
    'Units': ['id_unit', 'nama_unit', 'plat_nomor', 'harga_per_hari', 'target_km_servis', 'status', 'last_odo'],
    'Cashflow': ['id', 'tanggal', 'kategori', 'metode', 'keterangan', 'nominal', 'tipe'],
    'Settings': ['key', 'value'],
    'Expenses': ['id', 'tanggal', 'kategori', 'metode', 'keterangan', 'nominal']
  };
  
  Object.entries(sheets).forEach(([name, headers]) => {
    let s = ss.getSheetByName(name);
    if (!s) {
      s = ss.insertSheet(name);
      s.appendRow(headers);
    } else if (s.getLastRow() === 0) {
      s.appendRow(headers);
    }
  });
}

// --- TRANSACTIONS ---
function saveTransaction(data) {
  ensureHeaders();
  const ss = getSS();
  const transSheet = ss.getSheetByName('Transactions');
  const unitSheet = ss.getSheetByName('Units');
  const cashSheet = ss.getSheetByName('Cashflow');
  
  const id = "TX" + Date.now();
  
  // 1. Record Transaction
  // Columns: A:id, B:customer_name, C:wa_number, D:ktp_number, E:alamat_antar, F:alamat_ambil, G:id_unit, H:tgl_mulai, I:tgl_selesai, J:durasi_hari, K:overtime_jam, L:biaya_sewa, M:biaya_overtime, N:biaya_antar, O:biaya_ambil, P:biaya_addon, Q:addon_list, R:total, S:dp, T:metode_dp, U:lunas_amt, V:lunas_method, W:pay_status, X:trans_status, Y:photo
  
  transSheet.appendRow([
    id, data.customer_name, data.wa_number, data.ktp_number, data.alamat_antar || '', data.alamat_ambil || '', data.id_unit, 
    data.tgl_mulai, data.tgl_selesai, data.durasi_hari, data.overtime_jam, 
    data.biaya_sewa, data.biaya_overtime, data.biaya_antar || 0, data.biaya_ambil || 0, data.biaya_addon || 0,
    data.addon_list || '', data.total, data.dp, data.metode_dp,
    0, '', // lunas_amt, lunas_method
    Number(data.total) <= Number(data.dp) ? 'Lunas' : 'DP',
    'Aktif', '' // trans_status, photo
  ]);
  
  // 2. Update Unit Status
  const unitData = unitSheet.getDataRange().getValues();
  for(let i = 1; i < unitData.length; i++) {
    if(unitData[i][0] == data.id_unit) {
      unitSheet.getRange(i+1, 6).setValue('Disewa');
      break;
    }
  }
  
  // 3. Log Cashflow (DP)
  if(data.dp > 0) {
    cashSheet.appendRow([
      "CS" + Date.now(), new Date(), 'Sewa Motor (DP)', data.metode_dp, 
      `DP Sewa ${data.id_unit} - ${id}`, data.dp, 'Masuk'
    ]);
  }
  
  return { success: true, id };
}

function completeTransaction(data) {
  const ss = getSS();
  const transSheet = ss.getSheetByName('Transactions');
  const unitSheet = ss.getSheetByName('Units');
  const cashSheet = ss.getSheetByName('Cashflow');
  
  const rows = transSheet.getDataRange().getValues();
  let unitId = '';
  
  for(let i = 1; i < rows.length; i++) {
    if(rows[i][0] == data.id_transaksi) {
      unitId = rows[i][6]; 
      
      // Update Transaction Record
      const prevLunas = Number(transSheet.getRange(i+1, 21).getValue() || 0);
      const total = Number(transSheet.getRange(i+1, 18).getValue() || 0);
      const dp = Number(transSheet.getRange(i+1, 19).getValue() || 0);
      const newLunasTotal = prevLunas + Number(data.pelunasan || 0);

      transSheet.getRange(i+1, 21).setValue(newLunasTotal); 
      transSheet.getRange(i+1, 22).setValue(data.metode_pelunasan);
      
      if (dp + newLunasTotal >= total) {
        transSheet.getRange(i+1, 23).setValue('Lunas'); 
      }
      transSheet.getRange(i+1, 24).setValue('Selesai'); 
      if(data.foto_url) transSheet.getRange(i+1, 25).setValue(data.foto_url); 
      
      // Update Unit Status & ODO
      const unitData = unitSheet.getDataRange().getValues();
      for(let j = 1; j < unitData.length; j++) {
        if(unitData[j][0] == unitId) {
          unitSheet.getRange(j+1, 6).setValue('Tersedia');
          if(data.odo_akhir) unitSheet.getRange(j+1, 8).setValue(data.odo_akhir);
          break;
        }
      }
      
      // Log Cashflow
      if(data.pelunasan > 0) {
        cashSheet.appendRow([
          "CS" + Date.now(), new Date(), 'Pelunasan Sewa', data.metode_pelunasan, 
          `Pelunasan Sewa ${unitId} - ${data.id_transaksi}`, data.pelunasan, 'Masuk'
        ]);
      }
      return { success: true };
    }
  }
  return { success: false, message: "Transaction not found" };
}

function recordPelunasan(data) {
  const ss = getSS();
  const transSheet = ss.getSheetByName('Transactions');
  const cashSheet = ss.getSheetByName('Cashflow');
  
  const rows = transSheet.getDataRange().getValues();
  
  for(let i = 1; i < rows.length; i++) {
    if(rows[i][0] == data.id_transaksi) {
      const unitId = rows[i][6];
      const total = Number(rows[i][17] || 0); // Column R
      const dp = Number(rows[i][18] || 0);    // Column S
      const prevLunasAmt = Number(rows[i][20] || 0); // Column U
      const newLunasAmt = prevLunasAmt + Number(data.amount);
      
      // Update Transaction Record
      transSheet.getRange(i+1, 21).setValue(newLunasAmt); 
      transSheet.getRange(i+1, 22).setValue(data.metode);
      
      if (dp + newLunasAmt >= total) {
        transSheet.getRange(i+1, 23).setValue('Lunas');
      }
      
      // Log Cashflow
      cashSheet.appendRow([
        "CS" + Date.now(), new Date(), 'Pelunasan Sewa', data.metode, 
        `Pelunasan Sewa ${unitId} - ${data.id_transaksi}`, data.amount, 'Masuk'
      ]);
      
      return { success: true };
    }
  }
  return { success: false, message: "Transaction not found" };
}

// --- MASTER DATA ---
function getUnits() { return { success: true, data: getDataFromSheet('Units') }; }
function getTransactions() { return { success: true, data: getDataFromSheet('Transactions') }; }
function getCashflow() { return { success: true, data: getDataFromSheet('Cashflow') }; }
function getSettings() { 
  const settingsArr = getDataFromSheet('Settings');
  const settings = {};
  
  // Robustly build settings object
  settingsArr.forEach(s => {
    // We expect columns like [key, value]
    // Normalized headers from getDataFromSheet should be 'key' and 'value'
    let k = s.key;
    let v = s.value;
    
    // Fallback search if headers are different
    if (k === undefined) {
      const keys = Object.keys(s);
      if (keys.length >= 1) k = s[keys[0]]; 
      if (keys.length >= 2) v = s[keys[1]]; 
    }
    
    if (k !== undefined && k !== null && k !== "") {
      const normalizedKey = k.toString().toLowerCase().trim();
      settings[normalizedKey] = v;
    }
  });
  
  return { success: true, data: settings }; 
}

function saveUnit(data) {
  const sheet = getSheet('Units');
  const existingData = sheet.getDataRange().getValues();
  let foundRow = -1;
  
  if (data.id_unit) {
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0] == data.id_unit) {
        foundRow = i + 1;
        break;
      }
    }
  }

  if (foundRow > -1) {
    sheet.getRange(foundRow, 2).setValue(data.nama_unit);
    sheet.getRange(foundRow, 3).setValue(data.plat_nomor);
    sheet.getRange(foundRow, 4).setValue(data.harga_per_hari);
    sheet.getRange(foundRow, 5).setValue(data.target_km_servis);
    sheet.getRange(foundRow, 6).setValue(data.status);
    if(data.last_odo) sheet.getRange(foundRow, 7).setValue(data.last_odo);
  } else {
    sheet.appendRow([
      "UT" + Date.now(), data.nama_unit, data.plat_nomor, data.harga_per_hari,
      data.target_km_servis, 'Tersedia', 0 // ID, Name, Plate, Price, ServiceTarget, Status, LastODO
    ]);
  }
  return { success: true };
}

function saveExpense(data) {
  const ss = getSS();
  const expSheet = ss.getSheetByName('Expenses');
  const cashSheet = ss.getSheetByName('Cashflow');
  
  const id = "EX" + Date.now();
  expSheet.appendRow([id, new Date(), data.kategori, data.metode, data.keterangan, data.nominal]);
  cashSheet.appendRow(["CS" + Date.now(), new Date(), data.kategori, data.metode, data.keterangan, data.nominal, 'Keluar']);
  
  return { success: true };
}

function saveCash(data) {
  const sheet = getSheet('Cashflow');
  sheet.appendRow(["CS" + Date.now(), new Date(), data.kategori, data.metode, data.keterangan, data.nominal, data.tipe]);
  return { success: true };
}

function uploadPhoto(data) {
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const contentType = data.content.substring(5, data.content.indexOf(';'));
    const bytes = Utilities.base64Decode(data.content.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, data.filename);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: file.getUrl() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
