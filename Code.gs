/*******************************************************************
 * e-SAKIP — Backend Google Apps Script
 * Dinas Transmigrasi dan Tenaga Kerja Kabupaten Luwu Utara
 *
 * CARA PASANG
 * 1. Buka Google Sheets baru → menu Extensions → Apps Script.
 * 2. Ganti seluruh isi Code.gs dengan berkas ini, lalu Simpan.
 * 3. Jalankan fungsi `setupSpreadsheet()` satu kali (menu Run).
 *    Fungsi ini membuat semua sheet + header + 1 akun admin default.
 * 4. Deploy → New deployment → type "Web app":
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Salin URL yang berakhiran /exec.
 * 5. Tempelkan URL itu di halaman Admin dashboard (kolom "Koneksi Google Sheets").
 * 6. Ganti kata sandi admin default pada sheet `Admin`.
 *
 * ENDPOINT
 *   GET  ?action=all                      -> seluruh data (JSON)
 *   GET  ?action=sheet&sheet=Evaluasi     -> satu sheet
 *   GET  ?action=ringkasan                -> ringkasan KPI
 *   POST {action:'login', username, password}
 *   POST {action:'create', token, sheet, row}
 *   POST {action:'update', token, sheet, id, row}
 *   POST {action:'delete', token, sheet, id}
 *   POST {action:'bulk',   token, sheet, csv}
 *******************************************************************/

var SHEETS = {
  Perencanaan: ['id', 'tahun', 'jenis', 'sasaran', 'indikator', 'satuan', 'target', 'program', 'keterangan'],
  PerjanjianKinerja: ['id', 'tahun', 'bidang', 'sasaran', 'indikator', 'satuan', 'target', 'realisasi', 'anggaran'],
  RealisasiIKU: ['id', 'tahun', 'triwulan', 'indikator', 'satuan', 'target', 'realisasi', 'pagu', 'realisasi_anggaran'],
  Evaluasi: ['id', 'tahun', 'nilai', 'predikat', 'evaluator', 'tanggal', 'status', 'catatan'],
  Admin: ['username', 'password', 'nama', 'peran', 'terakhir_login']
};

var PREFIX = {
  Perencanaan: 'RC',
  PerjanjianKinerja: 'PK',
  RealisasiIKU: 'IK',
  Evaluasi: 'EV'
};

var TOKEN_TTL_MS = 8 * 60 * 60 * 1000; // 8 jam

/* ------------------------- Utilitas respons ------------------------- */

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function ok(data, extra) {
  var body = { ok: true, data: data, updatedAt: new Date().toISOString() };
  if (extra) Object.keys(extra).forEach(function (k) { body[k] = extra[k]; });
  return jsonOut(body);
}

function fail(message, code) {
  return jsonOut({ ok: false, message: message, code: code || 'error' });
}

/**
 * Apps Script Web App tidak mengizinkan header respons kustom.
 * Karena permintaan dari dashboard memakai Content-Type text/plain
 * (simple request), preflight CORS tidak terjadi dan respons
 * ContentService sudah dapat dibaca lintas origin.
 * doOptions disediakan agar klien lain tidak menerima 405.
 */
function doOptions() {
  return ContentService.createTextOutput('');
}

/* ------------------------- Akses spreadsheet ------------------------- */

function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  if (!SHEETS[name]) throw new Error('Sheet tidak dikenal: ' + name);
  var sheet = ss().getSheetByName(name);
  if (!sheet) {
    sheet = ss().insertSheet(name);
    sheet.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readSheet(name) {
  var sheet = getSheet(name);
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var header = values[0].map(String);
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row.join('').toString().trim() === '') continue;
    var obj = {};
    for (var c = 0; c < header.length; c++) {
      var v = row[c];
      obj[header[c]] = v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : v;
    }
    out.push(obj);
  }
  return out;
}

function nextId(name) {
  var rows = readSheet(name);
  var max = 0;
  rows.forEach(function (r) {
    var n = parseInt(String(r.id || '').replace(/[^0-9]/g, ''), 10);
    if (!isNaN(n) && n > max) max = n;
  });
  var num = String(max + 1);
  while (num.length < 2) num = '0' + num;
  return (PREFIX[name] || 'RW') + '-' + num;
}

/* ------------------------------- Auth ------------------------------- */

function makeToken(username) {
  var token = Utilities.getUuid();
  var cache = CacheService.getScriptCache();
  cache.put('sesi_' + token, username, TOKEN_TTL_MS / 1000);
  PropertiesService.getScriptProperties().setProperty(
    'sesi_' + token,
    JSON.stringify({ username: username, exp: Date.now() + TOKEN_TTL_MS })
  );
  return token;
}

function requireAdmin(token) {
  if (!token) throw new Error('Sesi tidak ditemukan. Silakan masuk kembali.');
  var cached = CacheService.getScriptCache().get('sesi_' + token);
  if (cached) return cached;
  var raw = PropertiesService.getScriptProperties().getProperty('sesi_' + token);
  if (!raw) throw new Error('Sesi kedaluwarsa. Silakan masuk kembali.');
  var sesi = JSON.parse(raw);
  if (sesi.exp < Date.now()) {
    PropertiesService.getScriptProperties().deleteProperty('sesi_' + token);
    throw new Error('Sesi kedaluwarsa. Silakan masuk kembali.');
  }
  return sesi.username;
}

function login(username, password) {
  var admins = readSheet('Admin');
  for (var i = 0; i < admins.length; i++) {
    var a = admins[i];
    if (
      String(a.username).trim().toLowerCase() === String(username || '').trim().toLowerCase() &&
      String(a.password) === String(password)
    ) {
      var sheet = getSheet('Admin');
      sheet.getRange(i + 2, SHEETS.Admin.indexOf('terakhir_login') + 1).setValue(new Date());
      return { token: makeToken(a.username), user: a.nama || a.username, peran: a.peran || 'admin' };
    }
  }
  throw new Error('Nama pengguna atau kata sandi salah.');
}

/* ------------------------------ Mutasi ------------------------------ */

function createRow(name, row) {
  var header = SHEETS[name];
  var sheet = getSheet(name);
  var id = row.id ? String(row.id) : nextId(name);
  var line = header.map(function (k) {
    if (k === 'id') return id;
    return row[k] === undefined || row[k] === null ? '' : row[k];
  });
  sheet.appendRow(line);
  return { id: id };
}

function updateRow(name, id, row) {
  var header = SHEETS[name];
  var sheet = getSheet(name);
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      var line = header.map(function (k, c) {
        var incoming = row[k];
        if (k === 'id') return id;
        if (incoming === undefined || incoming === null || incoming === '') return values[i][c];
        return incoming;
      });
      sheet.getRange(i + 1, 1, 1, header.length).setValues([line]);
      return { id: id };
    }
  }
  throw new Error('Baris dengan id ' + id + ' tidak ditemukan.');
}

function deleteRow(name, id) {
  var sheet = getSheet(name);
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { id: id };
    }
  }
  throw new Error('Baris dengan id ' + id + ' tidak ditemukan.');
}

function bulkCsv(name, csv) {
  var table = Utilities.parseCsv(String(csv).trim());
  if (!table.length) throw new Error('CSV kosong.');
  var head = table[0].map(function (h) { return String(h).trim(); });
  var jumlah = 0;
  for (var i = 1; i < table.length; i++) {
    if (table[i].join('').trim() === '') continue;
    var row = {};
    for (var c = 0; c < head.length; c++) row[head[c]] = table[i][c];
    createRow(name, row);
    jumlah++;
  }
  return { jumlah: jumlah };
}

/* ---------------------------- Ringkasan ---------------------------- */

function ringkasan() {
  var pk = readSheet('PerjanjianKinerja');
  var iku = readSheet('RealisasiIKU');
  var evaluasi = readSheet('Evaluasi');

  var capaian = 0;
  if (pk.length) {
    var total = 0;
    pk.forEach(function (r) {
      var t = Number(r.target) || 0;
      if (t > 0) total += Math.min((Number(r.realisasi) || 0) / t * 100, 120);
    });
    capaian = total / pk.length;
  }

  var pagu = 0, realisasiAnggaran = 0;
  iku.forEach(function (r) {
    pagu += Number(r.pagu) || 0;
    realisasiAnggaran += Number(r.realisasi_anggaran) || 0;
  });

  evaluasi.sort(function (a, b) { return Number(b.tahun) - Number(a.tahun); });
  var terbaru = evaluasi[0] || {};

  return {
    capaianKinerja: Math.round(capaian * 10) / 10,
    serapanAnggaran: pagu ? Math.round((realisasiAnggaran / pagu) * 1000) / 10 : 0,
    pagu: pagu,
    realisasiAnggaran: realisasiAnggaran,
    nilaiSakip: Number(terbaru.nilai) || 0,
    predikat: terbaru.predikat || '-',
    tahun: terbaru.tahun || new Date().getFullYear()
  };
}

/* ------------------------------ doGet ------------------------------ */

function doGet(e) {
  try {
    var params = (e && e.parameter) || {};
    var action = params.action || 'all';

    if (action === 'ringkasan') return ok(ringkasan());

    if (action === 'sheet') {
      if (!SHEETS[params.sheet]) return fail('Parameter sheet tidak valid.', 'bad_request');
      var one = {};
      one[params.sheet] = readSheet(params.sheet);
      return ok(one);
    }

    if (action === 'all') {
      return ok(
        {
          Perencanaan: readSheet('Perencanaan'),
          PerjanjianKinerja: readSheet('PerjanjianKinerja'),
          RealisasiIKU: readSheet('RealisasiIKU'),
          Evaluasi: readSheet('Evaluasi')
        },
        { ringkasan: ringkasan() }
      );
    }

    return fail('Action tidak dikenal: ' + action, 'bad_request');
  } catch (err) {
    return fail(err.message);
  }
}

/* ------------------------------ doPost ----------------------------- */

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) body = JSON.parse(e.postData.contents);
    var action = body.action;

    if (action === 'login') return ok(login(body.username, body.password));

    // Semua aksi tulis wajib membawa token admin yang sah.
    var username = requireAdmin(body.token);
    var sheet = body.sheet;
    if (!SHEETS[sheet] || sheet === 'Admin') return fail('Sheet tidak valid.', 'bad_request');

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (action === 'create') return ok(createRow(sheet, body.row || {}), { oleh: username });
      if (action === 'update') return ok(updateRow(sheet, body.id, body.row || {}), { oleh: username });
      if (action === 'delete') return ok(deleteRow(sheet, body.id), { oleh: username });
      if (action === 'bulk') return ok(bulkCsv(sheet, body.csv), { oleh: username });
    } finally {
      lock.releaseLock();
    }

    return fail('Action tidak dikenal: ' + action, 'bad_request');
  } catch (err) {
    return fail(err.message);
  }
}

/* --------------------------- Setup sekali --------------------------- */

function setupSpreadsheet() {
  Object.keys(SHEETS).forEach(function (name) {
    var sheet = getSheet(name);
    sheet.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  });

  if (!readSheet('Admin').length) {
    getSheet('Admin').appendRow(['admin', 'sakip2026', 'Admin Dinas', 'admin', '']);
  }

  if (!readSheet('PerjanjianKinerja').length) {
    getSheet('PerjanjianKinerja').appendRow([
      'PK-01', 2026, 'Ketenagakerjaan', 'Meningkatnya penempatan tenaga kerja',
      'Jumlah tenaga kerja ditempatkan', 'orang', 2500, 2560, 2400000000
    ]);
  }
  if (!readSheet('Perencanaan').length) {
    getSheet('Perencanaan').appendRow([
      'RC-01', 2026, 'RKT', 'Meningkatnya penempatan tenaga kerja',
      'Jumlah tenaga kerja ditempatkan', 'orang', 2500, 'Penempatan Tenaga Kerja', 'RKT 2026'
    ]);
  }
  if (!readSheet('RealisasiIKU').length) {
    getSheet('RealisasiIKU').appendRow([
      'IK-01', 2026, 'TW I', 'Penempatan Tenaga Kerja', 'orang', 625, 590, 600000000, 480000000
    ]);
  }
  if (!readSheet('Evaluasi').length) {
    getSheet('Evaluasi').appendRow([
      'EV-01', 2026, 84.2, 'A', 'Inspektorat Kab. Luwu Utara', '2026-07-18', 'Selesai',
      'Perbaikan pada cascading indikator kinerja.'
    ]);
  }
}
