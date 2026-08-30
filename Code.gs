/**
 * ============================================================================
 *  e-SAKIP — Dinas Transmigrasi dan Tenaga Kerja Kabupaten Luwu Utara
 *  Backend/API: Google Apps Script + Google Sheets
 * ============================================================================
 *  Cara pakai singkat:
 *  1. Buat Google Spreadsheet baru, buka Extensions > Apps Script.
 *  2. Tempel seluruh isi file ini ke code.gs (hapus isi default).
 *  3. Jalankan fungsi setupSheets() sekali (pilih fungsi di dropdown > Run)
 *     untuk membuat semua tab + header + contoh data secara otomatis.
 *  4. Isi API_KEY pada sheet "Config" (kolom B baris API_KEY) dengan
 *     string rahasia Anda sendiri — dipakai untuk validasi doPost().
 *  5. Deploy > New deployment > Type: Web app.
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     Salin URL Web App yang diberikan → tempel ke APPS_SCRIPT_URL di index.html.
 *  6. Setiap kali Anda mengubah code.gs, buat "New deployment" lagi (atau
 *     Manage deployments > edit > New version) agar perubahan live.
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// KONFIGURASI SHEET
// ---------------------------------------------------------------------------
const SHEET_RENSTRA   = 'Renstra_RKT';
const SHEET_PK         = 'Perjanjian_Kinerja';
const SHEET_REALISASI  = 'Realisasi_Kinerja_Anggaran';
const SHEET_EVALUASI   = 'Evaluasi_SAKIP';
const SHEET_CONFIG     = 'Config';

// Peta nama "logis" (dipakai di query ?sheet=...) ke nama sheet asli
const SHEET_MAP = {
  renstra:   SHEET_RENSTRA,
  pk:        SHEET_PK,
  realisasi: SHEET_REALISASI,
  evaluasi:  SHEET_EVALUASI
};

const HEADERS = {
  [SHEET_RENSTRA]: ['ID','Tahun','Sasaran_Strategis','Kode_IKU','Indikator_Kinerja','Target','Satuan','Program','Kegiatan','Anggaran_Pagu','Keterangan'],
  [SHEET_PK]: ['ID','Tahun','Sasaran_Strategis','Kode_IKU','Indikator_Kinerja','Target','Satuan','Penanggung_Jawab','Triwulan','Tanggal_PK','Status_PK'],
  [SHEET_REALISASI]: ['ID','ID_PK','Tahun','Triwulan','Kode_IKU','Indikator_Kinerja','Target','Realisasi','Satuan','Capaian_Persen','Anggaran_Pagu','Anggaran_Realisasi','Serapan_Persen','Status','Keterangan','Tanggal_Update'],
  [SHEET_EVALUASI]: ['ID','Tahun','Komponen','Bobot','Nilai','Skor','Predikat','Evaluator','Tanggal_Evaluasi','Catatan','Rekomendasi'],
  [SHEET_CONFIG]: ['Key','Value']
};

const ID_PREFIX = {
  [SHEET_RENSTRA]: 'REN',
  [SHEET_PK]: 'PK',
  [SHEET_REALISASI]: 'RLS',
  [SHEET_EVALUASI]: 'EV'
};

// ---------------------------------------------------------------------------
// SETUP — jalankan sekali secara manual dari editor Apps Script
// ---------------------------------------------------------------------------
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(HEADERS).forEach(function (sheetName) {
    let sh = ss.getSheetByName(sheetName);
    if (!sh) sh = ss.insertSheet(sheetName);
    sh.clear();
    sh.getRange(1, 1, 1, HEADERS[sheetName].length).setValues([HEADERS[sheetName]]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS[sheetName].length).setFontWeight('bold').setBackground('#0B2545').setFontColor('#FFFFFF');
  });

  // Hapus sheet default "Sheet1" jika masih ada dan kosong
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  // Isi Config default
  const cfg = ss.getSheetByName(SHEET_CONFIG);
  cfg.getRange(2, 1, 3, 2).setValues([
    ['API_KEY', 'ubah-dengan-kunci-rahasia-anda'],
    ['NAMA_OPD', 'Dinas Transmigrasi dan Tenaga Kerja Kabupaten Luwu Utara'],
    ['TAHUN_AKTIF', new Date().getFullYear()]
  ]);

  // Contoh data (boleh dihapus setelah Anda paham strukturnya)
  const renstra = ss.getSheetByName(SHEET_RENSTRA);
  renstra.appendRow(['REN-0001', 2026, 'Meningkatnya kualitas penempatan tenaga kerja', 'IKU-01', 'Persentase pencari kerja yang ditempatkan', 75, '%', 'Program Penempatan Tenaga Kerja', 'Fasilitasi penempatan kerja', 350000000, '']);
  renstra.appendRow(['REN-0002', 2026, 'Meningkatnya kualitas penyelenggaraan transmigrasi', 'IKU-02', 'Persentase kawasan transmigrasi yang berkembang', 60, '%', 'Program Pengembangan Kawasan Transmigrasi', 'Pembinaan kawasan transmigrasi', 500000000, '']);

  const pk = ss.getSheetByName(SHEET_PK);
  pk.appendRow(['PK-0001', 2026, 'Meningkatnya kualitas penempatan tenaga kerja', 'IKU-01', 'Persentase pencari kerja yang ditempatkan', 75, '%', 'Bidang Penempatan Tenaga Kerja', 'TW1', new Date(2026,0,15), 'Ditetapkan']);
  pk.appendRow(['PK-0002', 2026, 'Meningkatnya kualitas penyelenggaraan transmigrasi', 'IKU-02', 'Persentase kawasan transmigrasi yang berkembang', 60, '%', 'Bidang Transmigrasi', 'TW1', new Date(2026,0,15), 'Ditetapkan']);

  const rls = ss.getSheetByName(SHEET_REALISASI);
  upsertRealisasiRow_(rls, {
    id: 'RLS-0001', idPk: 'PK-0001', tahun: 2026, triwulan: 'TW1', kodeIku: 'IKU-01',
    indikatorKinerja: 'Persentase pencari kerja yang ditempatkan', target: 75, realisasi: 62,
    satuan: '%', anggaranPagu: 350000000, anggaranRealisasi: 210000000, keterangan: 'Sesuai rencana TW1'
  });

  const ev = ss.getSheetByName(SHEET_EVALUASI);
  ev.appendRow(['EV-0001', 2026, 'Perencanaan Kinerja', 30, 80, 24, '', 'Inspektorat Kab. Luwu Utara', new Date(2026,2,1), 'Renstra sudah selaras RPJMD', '']);
  ev.appendRow(['EV-0002', 2026, 'Pengukuran Kinerja', 30, 75, 22.5, '', 'Inspektorat Kab. Luwu Utara', new Date(2026,2,1), 'IKU sudah SMART', '']);
  ev.appendRow(['EV-0003', 2026, 'Pelaporan Kinerja', 15, 78, 11.7, '', 'Inspektorat Kab. Luwu Utara', new Date(2026,2,1), 'LKjIP tepat waktu', '']);
  ev.appendRow(['EV-0004', 2026, 'Evaluasi Internal', 10, 70, 7, '', 'Inspektorat Kab. Luwu Utara', new Date(2026,2,1), '', '']);
  ev.appendRow(['EV-0005', 2026, 'Capaian Kinerja', 15, 82, 12.3, '', 'Inspektorat Kab. Luwu Utara', new Date(2026,2,1), '', '']);
  recomputePredikat_(ss, 2026);

  SpreadsheetApp.getUi().alert('Setup selesai! Semua sheet & contoh data telah dibuat.');
}

// ---------------------------------------------------------------------------
// doGet — REST GET endpoint (JSON)
// ---------------------------------------------------------------------------
function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const action = params.action || 'dashboard';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tahun = params.tahun ? Number(params.tahun) : null;

    let payload;
    switch (action) {
      case 'dashboard':
        payload = getDashboardData_(ss, tahun);
        break;
      case 'list':
        payload = getListData_(ss, params.sheet, params, tahun);
        break;
      case 'config':
        payload = getConfig_(ss);
        break;
      default:
        return jsonResponse_({ success: false, error: 'Aksi tidak dikenal: ' + action }, 400);
    }

    return jsonResponse_({ success: true, data: payload }, 200);
  } catch (err) {
    return jsonResponse_({ success: false, error: String(err) }, 500);
  }
}

// ---------------------------------------------------------------------------
// doPost — REST POST/PUT/DELETE endpoint (JSON body dikirim sebagai text/plain
// dari frontend agar tidak memicu CORS preflight OPTIONS, lalu di-parse manual)
// ---------------------------------------------------------------------------
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};

    if (!isValidApiKey_(ss, body.apiKey)) {
      return jsonResponse_({ success: false, error: 'API key tidak valid.' }, 401);
    }

    const sheetKey = body.sheet;
    const sheetName = SHEET_MAP[sheetKey];
    if (!sheetName) {
      return jsonResponse_({ success: false, error: 'Sheet tidak dikenal: ' + sheetKey }, 400);
    }

    let result;
    switch (body.method) {
      case 'create':
        result = createRecord_(ss, sheetName, body.data || {});
        break;
      case 'update':
        result = updateRecord_(ss, sheetName, body.id, body.data || {});
        break;
      case 'delete':
        result = deleteRecord_(ss, sheetName, body.id);
        break;
      default:
        return jsonResponse_({ success: false, error: 'Method tidak dikenal: ' + body.method }, 400);
    }

    if (sheetName === SHEET_EVALUASI && body.data && body.data.Tahun) {
      recomputePredikat_(ss, Number(body.data.Tahun));
    }

    return jsonResponse_({ success: true, data: result }, 200);
  } catch (err) {
    return jsonResponse_({ success: false, error: String(err) }, 500);
  }
}

// ---------------------------------------------------------------------------
// HELPER — Respons JSON rapi (CORS pada Apps Script Web App diizinkan
// otomatis untuk request sederhana selama tidak ada custom header/Content-Type
// non text-plain pada request; lihat catatan CORS di README)
// ---------------------------------------------------------------------------
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function isValidApiKey_(ss, key) {
  const cfg = getConfig_(ss);
  if (!cfg.API_KEY || cfg.API_KEY === 'ubah-dengan-kunci-rahasia-anda') return true; // belum diset saat dev
  return key === cfg.API_KEY;
}

function getConfig_(ss) {
  const sh = ss.getSheetByName(SHEET_CONFIG);
  if (!sh) return {};
  const rows = sh.getDataRange().getValues();
  const cfg = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) cfg[rows[i][0]] = rows[i][1];
  }
  return cfg;
}

function sheetToObjects_(sh) {
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.every(function (c) { return c === '' || c === null; })) continue;
    const obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    out.push(obj);
  }
  return out;
}

function getListData_(ss, sheetKey, params, tahun) {
  const sheetName = SHEET_MAP[sheetKey];
  if (!sheetName) throw new Error('Sheet tidak dikenal: ' + sheetKey);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet belum dibuat: ' + sheetName + '. Jalankan setupSheets().');

  let rows = sheetToObjects_(sh);
  if (tahun) rows = rows.filter(function (r) { return Number(r.Tahun) === tahun; });
  if (params.triwulan) rows = rows.filter(function (r) { return r.Triwulan === params.triwulan; });
  if (params.status) rows = rows.filter(function (r) { return r.Status === params.status; });
  return rows;
}

function getDashboardData_(ss, tahun) {
  const shRealisasi = ss.getSheetByName(SHEET_REALISASI);
  const shPk = ss.getSheetByName(SHEET_PK);
  const shEval = ss.getSheetByName(SHEET_EVALUASI);
  if (!shRealisasi || !shPk || !shEval) {
    throw new Error('Sheet belum lengkap. Jalankan setupSheets() terlebih dahulu.');
  }

  let realisasi = sheetToObjects_(shRealisasi);
  let pk = sheetToObjects_(shPk);
  let evaluasi = sheetToObjects_(shEval);

  const activeYear = tahun || (realisasi.length ? Math.max.apply(null, realisasi.map(function (r) { return Number(r.Tahun); })) : new Date().getFullYear());

  realisasi = realisasi.filter(function (r) { return Number(r.Tahun) === activeYear; });
  pk = pk.filter(function (r) { return Number(r.Tahun) === activeYear; });
  evaluasi = evaluasi.filter(function (r) { return Number(r.Tahun) === activeYear; });

  const jumlahIku = realisasi.length;
  const avgCapaian = jumlahIku
    ? realisasi.reduce(function (s, r) { return s + (Number(r.Capaian_Persen) || 0); }, 0) / jumlahIku
    : 0;
  const totalPagu = realisasi.reduce(function (s, r) { return s + (Number(r.Anggaran_Pagu) || 0); }, 0);
  const totalRealisasiAnggaran = realisasi.reduce(function (s, r) { return s + (Number(r.Anggaran_Realisasi) || 0); }, 0);
  const avgSerapan = totalPagu ? (totalRealisasiAnggaran / totalPagu) * 100 : 0;

  const tercapai = realisasi.filter(function (r) { return r.Status === 'Tercapai'; }).length;
  const perluPerhatian = realisasi.filter(function (r) { return r.Status === 'Perlu Perhatian'; }).length;
  const belumTercapai = realisasi.filter(function (r) { return r.Status === 'Belum Tercapai'; }).length;

  const predikatRow = evaluasi.length ? evaluasi[evaluasi.length - 1] : null;
  const totalSkor = evaluasi.reduce(function (s, r) { return s + (Number(r.Skor) || 0); }, 0);

  return {
    tahun: activeYear,
    kpi: {
      capaianKinerjaPersen: round1_(avgCapaian),
      serapanAnggaranPersen: round1_(avgSerapan),
      predikatSakip: predikatRow ? predikatRow.Predikat : hitungPredikat_(totalSkor),
      totalSkorSakip: round1_(totalSkor),
      jumlahIku: jumlahIku
    },
    statusSebaran: { tercapai: tercapai, perluPerhatian: perluPerhatian, belumTercapai: belumTercapai },
    chartTargetRealisasi: realisasi.map(function (r) {
      return { indikator: r.Indikator_Kinerja, target: Number(r.Target) || 0, realisasi: Number(r.Realisasi) || 0, satuan: r.Satuan };
    }),
    tabelPerjanjianKinerja: pk.map(function (p) {
      const rls = realisasi.find(function (r) { return r.ID_PK === p.ID; });
      return {
        id: p.ID,
        sasaranStrategis: p.Sasaran_Strategis,
        indikator: p.Indikator_Kinerja,
        target: p.Target,
        satuan: p.Satuan,
        realisasi: rls ? rls.Realisasi : null,
        capaianPersen: rls ? round1_(Number(rls.Capaian_Persen) || 0) : null,
        status: rls ? rls.Status : 'Belum Ada Data'
      };
    }),
    evaluasiKomponen: evaluasi.map(function (r) {
      return { komponen: r.Komponen, bobot: r.Bobot, nilai: r.Nilai, skor: round1_(Number(r.Skor) || 0) };
    })
  };
}

function round1_(n) { return Math.round(n * 10) / 10; }

function hitungPredikat_(totalSkor) {
  if (totalSkor > 90) return 'AA';
  if (totalSkor > 80) return 'A';
  if (totalSkor > 70) return 'BB';
  if (totalSkor > 60) return 'B';
  if (totalSkor > 50) return 'CC';
  if (totalSkor > 30) return 'C';
  return 'D';
}

function recomputePredikat_(ss, tahun) {
  const sh = ss.getSheetByName(SHEET_EVALUASI);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idxTahun = headers.indexOf('Tahun');
  const idxSkor = headers.indexOf('Skor');
  const idxPredikat = headers.indexOf('Predikat');

  let total = 0;
  for (let r = 1; r < values.length; r++) {
    if (Number(values[r][idxTahun]) === tahun) total += Number(values[r][idxSkor]) || 0;
  }
  const predikat = hitungPredikat_(total);
  for (let r = 1; r < values.length; r++) {
    if (Number(values[r][idxTahun]) === tahun) sh.getRange(r + 1, idxPredikat + 1).setValue(predikat);
  }
}

// ---------------------------------------------------------------------------
// CRUD generik
// ---------------------------------------------------------------------------
function generateId_(sh, sheetName) {
  const prefix = ID_PREFIX[sheetName] || 'ID';
  const lastRow = sh.getLastRow();
  const n = lastRow > 1 ? lastRow - 1 + 1 : 1;
  return prefix + '-' + ('0000' + n).slice(-4) + '-' + Utilities.getUuid().slice(0, 4);
}

function createRecord_(ss, sheetName, data) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet tidak ditemukan: ' + sheetName);

  if (sheetName === SHEET_REALISASI) {
    if (!data.ID) data.ID = generateId_(sh, sheetName);
    upsertRealisasiRow_(sh, mapRealisasiInput_(data));
    return { id: data.ID };
  }

  const headers = HEADERS[sheetName];
  if (!data.ID) data.ID = generateId_(sh, sheetName);
  const row = headers.map(function (h) { return data[h] !== undefined ? data[h] : ''; });
  sh.appendRow(row);
  return { id: data.ID };
}

function updateRecord_(ss, sheetName, id, data) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet tidak ditemukan: ' + sheetName);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('ID');

  for (let r = 1; r < values.length; r++) {
    if (values[r][idCol] === id) {
      if (sheetName === SHEET_REALISASI) {
        const merged = {};
        headers.forEach(function (h, i) { merged[h] = values[r][i]; });
        Object.keys(data).forEach(function (k) { merged[k] = data[k]; });
        upsertRealisasiRow_(sh, mapRealisasiInput_(merged), r + 1);
        return { id: id, updated: true };
      }
      headers.forEach(function (h, i) {
        if (data[h] !== undefined) sh.getRange(r + 1, i + 1).setValue(data[h]);
      });
      return { id: id, updated: true };
    }
  }
  throw new Error('ID tidak ditemukan: ' + id);
}

function deleteRecord_(ss, sheetName, id) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet tidak ditemukan: ' + sheetName);
  const values = sh.getDataRange().getValues();
  const idCol = values[0].indexOf('ID');
  for (let r = 1; r < values.length; r++) {
    if (values[r][idCol] === id) {
      sh.deleteRow(r + 1);
      return { id: id, deleted: true };
    }
  }
  throw new Error('ID tidak ditemukan: ' + id);
}

// Mengisi field turunan (Capaian_Persen, Serapan_Persen, Status, Tanggal_Update)
// sebelum baris Realisasi ditulis — dipakai oleh setupSheets(), createRecord_, updateRecord_
function mapRealisasiInput_(d) {
  const target = Number(d.Target || d.target) || 0;
  const realisasi = Number(d.Realisasi || d.realisasi) || 0;
  const pagu = Number(d.Anggaran_Pagu || d.anggaranPagu) || 0;
  const realAnggaran = Number(d.Anggaran_Realisasi || d.anggaranRealisasi) || 0;
  const capaianPersen = target ? round1_((realisasi / target) * 100) : 0;
  const serapanPersen = pagu ? round1_((realAnggaran / pagu) * 100) : 0;
  let status = 'Belum Tercapai';
  if (capaianPersen >= 100) status = 'Tercapai';
  else if (capaianPersen >= 75) status = 'Perlu Perhatian';

  return {
    id: d.ID || d.id,
    idPk: d.ID_PK || d.idPk,
    tahun: d.Tahun || d.tahun,
    triwulan: d.Triwulan || d.triwulan,
    kodeIku: d.Kode_IKU || d.kodeIku,
    indikatorKinerja: d.Indikator_Kinerja || d.indikatorKinerja,
    target: target,
    realisasi: realisasi,
    satuan: d.Satuan || d.satuan,
    capaianPersen: capaianPersen,
    anggaranPagu: pagu,
    anggaranRealisasi: realAnggaran,
    serapanPersen: serapanPersen,
    status: status,
    keterangan: d.Keterangan || d.keterangan || '',
  };
}

function upsertRealisasiRow_(sh, r, rowIndex) {
  const row = [
    r.id, r.idPk, r.tahun, r.triwulan, r.kodeIku, r.indikatorKinerja, r.target, r.realisasi,
    r.satuan, r.capaianPersen, r.anggaranPagu, r.anggaranRealisasi, r.serapanPersen, r.status,
    r.keterangan, new Date()
  ];
  if (rowIndex) {
    sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
}
