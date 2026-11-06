/* =====================================================================
   SANTRIMAN WEBAPPS BUILDER V3 (ULTRA PRO MAX)
   WATERMARK: DIGITAL WEDDING INVITATION - BLOOMING LOVE THEME
   FILE: code.gs (Backend Logic - Enterprise Standard)
===================================================================== */


/**
 * Standard response format untuk semua fungsi
 */
function response(success, data, message, code) {
  return {
    success  : success,
    data     : data || null,
    message  : message || "",
    code     : code || (success ? 200 : 400),
    timestamp: getCurrentTimestamp()
  };
}

/**
 * Format timestamp Indonesia
 */
function getCurrentTimestamp() {
  var now = new Date();
  var d   = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  return d;
}

/**
 * Generate ID unik
 */
function generateId(prefix) {
  var now    = new Date();
  var tgl    = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd");
  var random = Math.floor(Math.random() * 90000) + 10000;
  return (prefix || "ID") + "-" + tgl + "-" + random;
}

/**
 * Log aktivitas ke sheet ActivityLog
 */
function logAktivitas(userId, username, action, modul, detail) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("ActivityLog");
    if (!sheet) return;
    
    appendRowToSheet("ActivityLog", [
      generateId("LOG"), userId, username, action, modul, detail, getCurrentTimestamp()
    ]);
  } catch(e) {
    Logger.log("logAktivitas error: " + e.toString());
  }
}

/**
 * Sanitasi input dari frontend
 */
function sanitasi(input) {
  if (input === null || input === undefined) return "";
  return input.toString().trim()
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}


/**
 * Helper: Ambil semua data sheet sebagai array of objects
 */
function getSheetData(namaSheet) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(namaSheet);
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return []; 
  
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var dataRows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  
  var result = [];
  for (var i = 0; i < dataRows.length; i++) {
    if (!dataRows[i][0]) continue; 
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = dataRows[i][j] || "";
    }
    obj._rowIndex = i + 2; 
    result.push(obj);
  }
  return result;
}

/**
 * Helper: Tambah baris baru ke sheet dengan LockService
 */
function appendRowToSheet(namaSheet, rowData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(namaSheet);
    if (!sheet) throw new Error("Sheet tidak ditemukan: " + namaSheet);
    sheet.appendRow(rowData);
    SpreadsheetApp.flush(); 
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper: Update baris tertentu berdasarkan nomor baris
 */
function updateRowInSheet(namaSheet, rowIndex, rowData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(namaSheet);
    if (!sheet) throw new Error("Sheet tidak ditemukan: " + namaSheet);
    var lastCol = sheet.getLastColumn();
    sheet.getRange(rowIndex, 1, 1, lastCol).setValues([rowData]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper: Hapus baris berdasarkan nomor baris
 */
function deleteRowInSheet(namaSheet, rowIndex) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(namaSheet);
    if (!sheet) throw new Error("Sheet tidak ditemukan: " + namaSheet);
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}


/**
 * Fungsi pembantu untuk membuat sheet jika belum ada
 */
function buatSheetJikaPerlu(ss, namaSheet, headers) {
  var sheet = ss.getSheetByName(namaSheet);
  
  if (sheet) {
    Logger.log("Sheet '" + namaSheet + "' sudah ada, skip.");
    return sheet;
  }
  
  sheet = ss.insertSheet(namaSheet);
  sheet.appendRow(headers);
  
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#1e293b");
  headerRange.setFontColor("#ffffff");
  headerRange.setHorizontalAlignment("center");
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  
  Logger.log("Sheet '" + namaSheet + "' berhasil dibuat.");
  return sheet;
}

/**
 * Setup Database Utama
 */
function setupDatabase() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    buatSheetJikaPerlu(ss, "Config",      ["key","value","updated_at"]);
    buatSheetJikaPerlu(ss, "Users",       ["id","username","password","email","role","status","full_name","phone","last_login","created_at"]);
    buatSheetJikaPerlu(ss, "Sessions",    ["token","user_id","username","role","created_at","expired_at","is_active"]);
    buatSheetJikaPerlu(ss, "ActivityLog", ["id","user_id","username","action","modul","detail","created_at"]);
    buatSheetJikaPerlu(ss, "RSVP",        ["id","nama_tamu","jumlah_hadir","status_kehadiran","pesan","created_at"]);
    buatSheetJikaPerlu(ss, "Ucapan",      ["id","nama_pengirim","ucapan","status_tampil","created_at"]);
    
    seedDemoData();
    
    SpreadsheetApp.getUi().alert(
      "✅ Setup Berhasil!\n\n" +
      "Database telah dibuat otomatis di spreadsheet ini.\n" +
      "Silakan deploy aplikasi sebagai Web App.\n\n" +
      "Akun demo:\nUsername: admin\nPassword: app123"
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert(
      "❌ Setup Gagal!\n\nError: " + e.message +
      "\n\nPastikan script dibuka dari Extensions → Apps Script di dalam Google Spreadsheet."
    );
    Logger.log("setupDatabase error: " + e.toString());
  }
}

/**
 * Seed data demo
 */
function seedDemoData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // SEED USERS
  var sheetUsers = ss.getSheetByName("Users");
  if (sheetUsers && sheetUsers.getLastRow() <= 1) {
    var now = getCurrentTimestamp();
    sheetUsers.appendRow(["USR-001", "admin", "app123", "admin@app.com", "Admin", "Active", "Administrator", "-", now, now]);
    sheetUsers.appendRow(["USR-002", "user01", "user123", "user@app.com", "User", "Active", "Demo User", "-", now, now]);
  }
  
  // SEED CONFIG
  var sheetConfig = ss.getSheetByName("Config");
  if (sheetConfig && sheetConfig.getLastRow() <= 1) {
    var now = getCurrentTimestamp();
    sheetConfig.appendRow(["app_name", "Blooming Love Invitation", now]);
    sheetConfig.appendRow(["app_version", "3.5", now]);
    sheetConfig.appendRow(["app_status", "active", now]);
  }

  // SEED UCAPAN AWAL (DUMMY)
  var sheetUcapan = ss.getSheetByName("Ucapan");
  if (sheetUcapan && sheetUcapan.getLastRow() <= 1) {
    var now = getCurrentTimestamp();
    sheetUcapan.appendRow([generateId("UCP"), "Keluarga Besar H. Mansyur", "Selamat menempuh hidup baru Ayyan & Nisa. Semoga menjadi keluarga yang Sakinah, Mawaddah, Warahmah. Aamiin.", "Tampil", now]);
  }
  
  Logger.log("seedDemoData: Selesai mengisi data demo.");
}


function simpanSession(token, dataUser) {
  var props = PropertiesService.getScriptProperties();
  var sessionData = {
    userId   : dataUser.userId,
    username : dataUser.username,
    role     : dataUser.role,
    fullName : dataUser.fullName,
    expiredAt: Date.now() + (30 * 60 * 1000) 
  };
  props.setProperty("SESSION_" + token, JSON.stringify(sessionData));
}

function cekSession(token) {
  if (!token) return { valid: false, pesan: "Token kosong" };
  var props = PropertiesService.getScriptProperties();
  var raw   = props.getProperty("SESSION_" + token);
  if (!raw)  return { valid: false, pesan: "Session tidak ada"};
  
  var session = JSON.parse(raw);
  if (Date.now() > session.expiredAt) {
    props.deleteProperty("SESSION_" + token); 
    return { valid: false, pesan: "Session kadaluarsa" };
  }
  
  session.expiredAt = Date.now() + (30 * 60 * 1000);
  props.setProperty("SESSION_" + token, JSON.stringify(session));
  return { valid: true, data: session };
}

function hapusSession(token) {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty("SESSION_" + token);
}

function login(username, password) {
  try {
    username = (username || "").toString().trim().toLowerCase();
    password = (password || "").toString().trim();
    if (!username || !password) {
      return response(false, null, "Username dan password wajib diisi");
    }
    
    var users = getSheetData("Users");
    if (!users || users.length === 0) {
      return response(false, null, "Database kosong. Jalankan setupDatabase() dulu.");
    }
    
    var user = null;
    for (var i = 0; i < users.length; i++) {
      if (users[i].username && users[i].username.toString().toLowerCase() === username) {
        user = users[i];
        break;
      }
    }
    
    if (!user) {
      return response(false, null, "Username tidak ditemukan");
    }
    if (user.status && user.status.toString() !== "Active") {
      return response(false, null, "Akun tidak aktif");
    }
    if (user.password.toString() !== password) {
      return response(false, null, "Password salah");
    }
    
    var token = Utilities.getUuid();
    simpanSession(token, {
      userId   : user.id,
      username : user.username,
      role     : user.role,
      fullName : user.full_name || user.username
    });
    
    var sheetUsers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    if (sheetUsers && user._rowIndex) {
      var headers = sheetUsers.getRange(1,1,1, sheetUsers.getLastColumn()).getValues()[0];
      var colIdx  = headers.indexOf("last_login") + 1;
      if (colIdx > 0) {
        sheetUsers.getRange(user._rowIndex, colIdx).setValue(getCurrentTimestamp());
      }
    }
    
    logAktivitas(user.id, user.username, "LOGIN", "Auth", "Login berhasil");
    
    return response(true, {
      token    : token,
      userId   : user.id,
      username : user.username,
      role     : user.role,
      fullName : user.full_name || user.username
    }, "Login berhasil");
    
  } catch(e) {
    Logger.log("login() error: " + e.toString());
    return response(false, null, "Error: " + e.message);
  }
}


/**
 * Menyimpan data RSVP dari tamu
 */
function submitRSVP(nama, jumlah, status, pesan) {
  try {
    var namaSanitized = sanitasi(nama);
    var pesanSanitized = sanitasi(pesan);
    
    if(!namaSanitized || !status) {
       return response(false, null, "Nama dan status kehadiran wajib diisi.");
    }

    var rowData = [
      generateId("RSVP"),
      namaSanitized,
      jumlah || 1,
      sanitasi(status),
      pesanSanitized,
      getCurrentTimestamp()
    ];
    
    appendRowToSheet("RSVP", rowData);
    return response(true, null, "Terima kasih, konfirmasi kehadiran Anda telah tersimpan.");
    
  } catch (e) {
    return response(false, null, "Gagal menyimpan RSVP: " + e.message);
  }
}

/**
 * Menyimpan ucapan dari tamu
 */
function submitUcapan(nama, ucapan) {
  try {
    var namaSanitized = sanitasi(nama);
    var ucapanSanitized = sanitasi(ucapan);
    
    if(!namaSanitized || !ucapanSanitized) {
       return response(false, null, "Nama dan ucapan wajib diisi.");
    }

    var rowData = [
      generateId("UCP"),
      namaSanitized,
      ucapanSanitized,
      "Tampil", // Default tampil
      getCurrentTimestamp()
    ];
    
    appendRowToSheet("Ucapan", rowData);
    return response(true, null, "Terima kasih atas doa dan ucapannya.");
    
  } catch (e) {
    return response(false, null, "Gagal menyimpan ucapan: " + e.message);
  }
}

/**
 * Mengambil daftar ucapan untuk ditampilkan di UI
 */
function getUcapanList() {
  try {
    var ucapanRaw = getSheetData("Ucapan");
    // Filter hanya yang statusnya Tampil dan urutkan dari terbaru
    var ucapanTampil = ucapanRaw.filter(function(item) {
      return item.status_tampil === "Tampil";
    });
    
    // Sort descending by created_at (asumsi format dd/MM/yyyy HH:mm:ss)
    // Untuk simplifikasi pengurutan di client side, kita lempar apa adanya (atau di reverse)
    ucapanTampil.reverse(); 
    
    return response(true, ucapanTampil, "Data ucapan berhasil dimuat.");
  } catch(e) {
    return response(false, null, "Gagal memuat ucapan: " + e.message);
  }
}


/**
 * Mengambil daftar ucapan & status kehadiran tamu dari sheet RSVP untuk ditampilkan di UI
 */
function getUnifiedList() {
  try {
    var rsvpData = getSheetData("RSVP");
    // Urutkan dari yang paling baru
    rsvpData.reverse();
    return response(true, rsvpData, "Data kehadiran & ucapan berhasil dimuat.");
  } catch(e) {
    return response(false, null, "Gagal memuat data: " + e.message);
  }
}

function doGet(e) {
  // Jika ada permintaan data ucapan / list dari client
  if (e && e.parameter && (e.parameter.action === 'getUnified' || e.parameter.action === 'getUcapan' || e.parameter.action === 'getComments')) {
    var listData = (e.parameter.action === 'getUcapan') ? getUcapanList() : getUnifiedList();
    return ContentService.createTextOutput(JSON.stringify(listData))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Jika ada permintaan simpan via GET (fallback jika POST dibatasi jaringan)
  if (e && e.parameter && (e.parameter.nama || e.parameter.action === 'submitUnified' || e.parameter.action === 'submitRSVP')) {
    return doPost(e);
  }

  // Default: Tampilkan template Web App jika dibuka langsung
  try {
    var html = HtmlService.createHtmlOutputFromFile('index')
        .setTitle('Ayyan & Nisa Wedding Invitation')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return html;
  } catch(err) {
    return ContentService.createTextOutput("Backend Wedding Invitation Ayyan & Nisa is Online.")
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function doPost(e) {
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch(parseErr) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }
    
    var nama   = data.nama || data.nama_tamu || "";
    var status = data.status || data.status_kehadiran || "Hadir";
    var jumlah = data.jumlah || data.jumlah_hadir || "1 Orang";
    var pesan  = data.pesan || data.ucapan || "";
    
    if (!nama || nama.toString().trim() === "") {
      return ContentService.createTextOutput(JSON.stringify(response(false, null, "Nama tamu wajib diisi")))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // 1. Simpan ke sheet RSVP
    var rsvpRes = submitRSVP(nama, jumlah, status, pesan);
    
    // 2. Jika ada ucapan, simpan juga ke sheet Ucapan
    if (pesan && pesan.toString().trim() !== "") {
      submitUcapan(nama, pesan);
    }
    
    return ContentService.createTextOutput(JSON.stringify(response(true, {
      nama: nama,
      status: status,
      jumlah: jumlah,
      pesan: pesan
    }, "Konfirmasi kehadiran dan ucapan Anda berhasil disimpan ke Spreadsheet.")))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify(response(false, null, "Error: " + error.toString())))
      .setMimeType(ContentService.MimeType.JSON);
  }
}