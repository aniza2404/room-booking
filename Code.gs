/* =========================================================
   SYSTEM: Room Booking (Sistem Tempahan Bilik)
   Backend — single source of truth.
   ========================================================= */

var SPREADSHEET_ID = "1IdzmMs3vjv9KS-zXPdsqrbPYDUQnU-G37H1vyzXpy5s";
var SHEET_NAME = "Bookings"; 
var LOCK_TIMEOUT_MS = 10000;
var ALLOWED_EMAIL_DOMAINS = ["ipsah.edu.my", "moe-dl.edu.my"];

// ADDED "Peralatan"
var COLS = [
  "Timestamp", "Email", "No. Telefon", "Nama", "Jabatan", "Program",
  "Tarikh Mula", "Tarikh Tamat", "Masa Mula", "Masa Tamat",
  "Tempat", "Tujuan", "Kelulusan", "Peralatan"
];

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || "index";
  var file, title;

  // ADDED ROUTING FOR CATALOGUE
  if (page === "admin") { 
    file = "admin"; 
    title = "Papan Pemuka Pentadbir"; 
  }
  else if (page === "form") { 
    file = "form"; 
    title = "Tempahan Bilik"; 
  }
  else if (page === "catalogue") { 
    file = "catalogue"; 
    title = "Katalog Peralatan"; 
  }
  else { 
    file = "index"; 
    title = "DigiBook — Sistem Tempahan Bilik dan Peralatan"; 
  }

  // Get logged in user email for the Lecturer Check
  var userEmail = Session.getActiveUser().getEmail();

  var template = HtmlService.createTemplateFromFile(file);
  template.userEmail = userEmail; 

  return template.evaluate()
    .setTitle(title.indexOf("DigiBook") === 0 ? title : title + " · DigiBook")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

function isAllowedEmail_(email) {
  var e = String(email || "").trim().toLowerCase();
  if (!e || e.indexOf("@") === -1) return false;
  var domain = e.split("@").pop();
  return ALLOWED_EMAIL_DOMAINS.some(function (d) {
    return domain === d || domain.endsWith("." + d);
  });
}

function getAdminPassword_() {
  return PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD") || "admin123";
}

function verifyAdminPassword(password) {
  return String(password || "") === String(getAdminPassword_());
}

function setAdminPassword(newPassword) {
  PropertiesService.getScriptProperties().setProperty("ADMIN_PASSWORD", String(newPassword));
  return "ADMIN_PASSWORD updated.";
}

function getSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLS);
  } else {
    var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var requiredCols = ["No. Telefon", "Kelulusan", "Peralatan"];
    requiredCols.forEach(function(colName) {
      var hasCol = header.some(function(h) { return String(h || "").trim().toLowerCase() === colName.toLowerCase(); });
      if (!hasCol) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(colName);
        header.push(colName);
      }
    });
  }
  return sh;
}

function headerIndex_(header) {
  var out = {};
  header.forEach(function (h, i) {
    var k = String(h || "").trim().toLowerCase();
    if      (k === "timestamp")    out.timestamp   = i;
    else if (k === "email")        out.email       = i;
    else if (k === "no. telefon")  out.noTel       = i;
    else if (k === "nama")         out.nama        = i;
    else if (k === "jabatan")      out.jabatan     = i;
    else if (k === "program")      out.program     = i;
    else if (k === "tarikh mula")  out.tarikhMula  = i;
    else if (k === "tarikh tamat") out.tarikhTamat = i;
    else if (k === "masa mula")    out.masaMula    = i;
    else if (k === "masa tamat")   out.masaTamat   = i;
    else if (k === "tempat")       out.tempat      = i;
    else if (k === "tujuan")       out.tujuan      = i;
    else if (k === "kelulusan")    out.kelulusan   = i;
    else if (k === "peralatan")    out.peralatan   = i; // ADDED
  });
  return out;
}

function str_(v) { return v == null ? "" : String(v).trim(); }
function pad_(n) { return String(n).padStart(2, "0"); }

function dateISO_(v) {
  if (v == null || v === "") return "";
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var s = String(v).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + "-" + m[2] + "-" + m[3];
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    var a = +m[1], b = +m[2], y = m[3];
    var d = a > 12 ? a : (b > 12 ? b : a);
    var mo = a > 12 ? b : (b > 12 ? a : b);
    return y + "-" + pad_(mo) + "-" + pad_(d);
  }
  return s;
}

function timeStr_(v) {
  if (v == null || v === "") return "";
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "HH:mm");
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})[:.](\d{2})/);
  if (m) return pad_(+m[1]) + ":" + m[2];
  if (/^\d{1,2}$/.test(s)) return pad_(+s) + ":00";
  return s;
}

function toISO_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  return str_(v);
}

function normRoom(s) { return String(s || "").toLowerCase().replace(/\s+/g, " ").trim(); }
function timesOverlap(aS, aE, bS, bE) { return aS < bE && aE > bS; }

function conflicts(a, b) {
  if (normRoom(a.tempat) !== normRoom(b.tempat)) return false;
  if (a.tarikhMula > b.tarikhTamat || b.tarikhMula > a.tarikhTamat) return false;
  return timesOverlap(a.masaMula, a.masaTamat, b.masaMula, b.masaTamat);
}

function isRejected_(r) { return String(r.kelulusan || "").toLowerCase() === "ditolak"; }

function getAllBookings() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];

  var values = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
  var idx = headerIndex_(values[0]);
  var out = [];

  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    var rec = {
      rowIndex:    i + 1,
      timestamp:   toISO_(r[idx.timestamp]),
      email:       str_(r[idx.email]),
      noTel:       str_(r[idx.noTel]),
      nama:        str_(r[idx.nama]),
      jabatan:     str_(r[idx.jabatan]),
      program:     str_(r[idx.program]),
      tarikhMula:  dateISO_(r[idx.tarikhMula]),
      tarikhTamat: dateISO_(r[idx.tarikhTamat]) || dateISO_(r[idx.tarikhMula]),
      masaMula:    timeStr_(r[idx.masaMula]),
      masaTamat:   timeStr_(r[idx.masaTamat]),
      tempat:      str_(r[idx.tempat]),
      tujuan:      str_(r[idx.tujuan]),
      kelulusan:   str_(r[idx.kelulusan]) || "Menunggu",
      peralatan:   str_(r[idx.peralatan]) // ADDED
    };
    if (rec.tarikhMula && rec.tempat && rec.masaMula && rec.masaTamat) out.push(rec);
  }
  return out;
}

function validateSlotOnly_(p) {
  if (!p) return "Data tidak lengkap.";
  if (!p.tempat) return "Sila pilih bilik.";
  if (!p.tarikhMula) return "Sila pilih tarikh.";
  if (!p.masaMula || !p.masaTamat) return "Sila pilih masa mula dan tamat.";
  if (p.masaMula >= p.masaTamat) return "Masa tamat mesti selepas masa mula.";
  return null;
}

function validateCandidate_(p) {
  if (!p) return "Data tidak lengkap.";
  if (!str_(p.nama)) return "Sila isi nama.";
  if (!isAllowedEmail_(p.email)) return "Email mesti guna domain " + ALLOWED_EMAIL_DOMAINS.join(" / ") + ".";
  if (!str_(p.noTel)) return "Sila isi no. telefon.";
  if (!str_(p.jabatan)) return "Sila isi jabatan.";
  if (!str_(p.program)) return "Sila isi program.";
  if (!str_(p.tujuan)) return "Sila isi tujuan.";
  return validateSlotOnly_(p);
}

function checkSlot(payload) {
  try {
    var err = validateSlotOnly_(payload);
    if (err) return { ok: false, error: err };

    var candidate = {
      tempat: payload.tempat, tarikhMula: payload.tarikhMula,
      tarikhTamat: payload.tarikhTamat || payload.tarikhMula,
      masaMula: payload.masaMula, masaTamat: payload.masaTamat
    };

    var rows = getAllBookings();
    var clash = rows.find(function (r) {
      if (isRejected_(r)) return false;
      return conflicts(candidate, r);
    });
    if (clash) {
      return { ok: false, error: "Slot " + clash.masaMula + "–" + clash.masaTamat + " pada " + clash.tarikhMula + " (" + clash.tempat + ") telah ditempah oleh " + (clash.nama || "pengguna lain") + "." };
    }
    return { ok: true };
  } catch (e) {
    return { ok: false, error: "Ralat: " + String(e) };
  }
}

function submitBooking(payload) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_TIMEOUT_MS)) return { ok: false, error: "Sistem sibuk. Sila cuba lagi." };

    var err = validateCandidate_(payload);
    if (err) return { ok: false, error: err };

    var candidate = {
      tempat: payload.tempat, tarikhMula: payload.tarikhMula,
      tarikhTamat: payload.tarikhTamat || payload.tarikhMula,
      masaMula: payload.masaMula, masaTamat: payload.masaTamat
    };

    var rows = getAllBookings();
    var clash = rows.find(function (r) {
      if (isRejected_(r)) return false;
      return conflicts(candidate, r);
    });
    if (clash) {
      return { ok: false, error: "Slot " + clash.masaMula + "–" + clash.masaTamat + " pada " + clash.tarikhMula + " sudah ditempah. Sila pilih masa lain." };
    }

    var sh = getSheet_();
    
    sh.appendRow([
      new Date(), 
      str_(payload.email), 
      str_(payload.noTel),
      str_(payload.nama), 
      str_(payload.jabatan),
      str_(payload.program), 
      candidate.tarikhMula, 
      candidate.tarikhTamat,
      candidate.masaMula, 
      candidate.masaTamat, 
      candidate.tempat, 
      str_(payload.tujuan), 
      "Menunggu",
      str_(payload.peralatan) // ADDED
    ]);

    return { ok: true };
  } catch (e) {
    return { ok: false, error: "Ralat: " + String(e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function setApproval(rowIndex, status, password) {
  if (!verifyAdminPassword(password)) return { ok: false, error: "Kata laluan salah." };
  if (["Menunggu", "Diluluskan", "Ditolak"].indexOf(status) === -1) return { ok: false, error: "Status tidak sah." };

  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_TIMEOUT_MS)) return { ok: false, error: "Sistem sibuk." };

    var sh = getSheet_();
    var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var idx = headerIndex_(header);
    
    var values = sh.getDataRange().getValues();
    if (rowIndex < 2 || rowIndex > values.length) return { ok: false, error: "Rekod tidak dijumpai." };

    if (status === "Diluluskan") {
      var target = {
        tempat: str_(values[rowIndex - 1][idx.tempat]), tarikhMula: dateISO_(values[rowIndex - 1][idx.tarikhMula]),
        tarikhTamat: dateISO_(values[rowIndex - 1][idx.tarikhTamat]), masaMula: timeStr_(values[rowIndex - 1][idx.masaMula]), masaTamat: timeStr_(values[rowIndex - 1][idx.masaTamat])
      };
      for (var i = 1; i < values.length; i++) {
        if (i + 1 === rowIndex) continue;
        var other = {
          tempat: str_(values[i][idx.tempat]), tarikhMula: dateISO_(values[i][idx.tarikhMula]),
          tarikhTamat: dateISO_(values[i][idx.tarikhTamat]), masaMula: timeStr_(values[i][idx.masaMula]), masaTamat: timeStr_(values[i][idx.masaTamat]), kelulusan: str_(values[i][idx.kelulusan])
        };
        if (!isRejected_(other) && conflicts(target, other)) return { ok: false, error: "Tidak boleh luluskan — bertindih dengan tempahan lain." };
      }
    }

    sh.getRange(rowIndex, idx.kelulusan + 1).setValue(status);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: "Ralat: " + String(e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function getEquipmentCatalog() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Equipments'); 
    
    if (!sheet) return []; 
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return []; 

    var headerRow = data[0];
    var idxId = -1, idxName = -1, idxImg = -1, idxDesc = -1, idxActive = -1;
    
    headerRow.forEach(function(h, i) {
      var val = String(h).trim().toLowerCase();
      if (val === 'id') idxId = i;
      else if (val === 'name') idxName = i;
      else if (val === 'image_url') idxImg = i;
      else if (val === 'description') idxDesc = i;
      else if (val === 'active') idxActive = i;
    });

    return data.slice(1)
      .filter(function(row) {
        if (idxActive !== -1) {
          const activeStr = String(row[idxActive]).trim().toUpperCase();
          if (activeStr !== 'TRUE' && activeStr !== '1' && activeStr !== 'Y' && activeStr !== 'YES' && activeStr !== '') return false;
        }
        return true;
      })
      .map(function(row) {
        var rawImageUrl = idxImg !== -1 ? String(row[idxImg] || '').trim() : '';
        
        // SUPER CONVERTER: Extracts ID and forces thumbnail format
        var directImageUrl = rawImageUrl;
        if (rawImageUrl.indexOf("drive.google.com") !== -1) {
          var match = rawImageUrl.match(/\/d\/([a-zA-Z0-9_-]+)/) || rawImageUrl.match(/id=([a-zA-Z0-9_-]+)/);
          if (match && match[1]) {
            // Using thumbnail endpoint with a large width
            directImageUrl = "https://drive.google.com/thumbnail?id=" + match[1] + "&sz=w1000";
          }
        }

        return {
          id: idxId !== -1 ? String(row[idxId] || '') : Math.random().toString(),       
          name: idxName !== -1 ? String(row[idxName] || '') : 'Peralatan',     
          imageUrl: directImageUrl, 
          desc: idxDesc !== -1 ? String(row[idxDesc] || '') : ''
        };
      });
      
  } catch (e) {
    console.error(e);
    return []; 
  }
}

function updateBooking(payload, password) {
  if (!verifyAdminPassword(password)) return { ok: false, error: "Kata laluan salah." };
  
  var err = validateSlotOnly_(payload);
  if (err) return { ok: false, error: err };

  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_TIMEOUT_MS)) return { ok: false, error: "Sistem sibuk." };

    var sh = getSheet_();
    var values = sh.getDataRange().getValues();
    var idx = headerIndex_(values[0]);
    var rowIndex = Number(payload.rowIndex);

    if (!rowIndex || rowIndex < 2 || rowIndex > values.length) return { ok: false, error: "Rekod tidak dijumpai." };

    var candidate = {
      tempat: payload.tempat, tarikhMula: dateISO_(payload.tarikhMula),
      tarikhTamat: dateISO_(payload.tarikhTamat) || dateISO_(payload.tarikhMula),
      masaMula: timeStr_(payload.masaMula), masaTamat: timeStr_(payload.masaTamat)
    };

    for (var i = 1; i < values.length; i++) {
      if (i + 1 === rowIndex) continue;
      var other = {
        tempat: str_(values[i][idx.tempat]), tarikhMula: dateISO_(values[i][idx.tarikhMula]),
        tarikhTamat: dateISO_(values[i][idx.tarikhTamat]), masaMula: timeStr_(values[i][idx.masaMula]), masaTamat: timeStr_(values[i][idx.masaTamat]), kelulusan: str_(values[i][idx.kelulusan])
      };
      if (!isRejected_(other) && conflicts(candidate, other)) return { ok: false, error: "Slot bertindih dengan tempahan lain." };
    }

    var row = values[rowIndex - 1].slice();
    row[idx.email]       = str_(payload.email);
    row[idx.noTel]       = str_(payload.noTel);
    row[idx.nama]        = str_(payload.nama);
    row[idx.jabatan]     = str_(payload.jabatan);
    row[idx.program]     = str_(payload.program);
    row[idx.tempat]      = candidate.tempat;
    row[idx.tarikhMula]  = candidate.tarikhMula;
    row[idx.tarikhTamat] = candidate.tarikhTamat;
    row[idx.masaMula]    = candidate.masaMula;
    row[idx.masaTamat]   = candidate.masaTamat;
    row[idx.tujuan]      = str_(payload.tujuan);
    row[idx.kelulusan]   = str_(payload.kelulusan) || "Menunggu";
    row[idx.peralatan]   = str_(payload.peralatan); // ADDED

    sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: "Ralat: " + String(e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function deleteBooking(rowIndex, password) {
  if (!verifyAdminPassword(password)) return { ok: false, error: "Kata laluan salah." };
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_TIMEOUT_MS)) return { ok: false, error: "Sistem sibuk." };
    var sh = getSheet_();
    var rowIndexN = Number(rowIndex);
    if (!rowIndexN || rowIndexN < 2 || rowIndexN > sh.getLastRow()) return { ok: false, error: "Rekod tidak dijumpai." };
    sh.deleteRow(rowIndexN);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: "Ralat: " + String(e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
