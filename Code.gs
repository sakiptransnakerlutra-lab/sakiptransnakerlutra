/**
 * ============================================================================
 *  e-SAKIP — Dinas Transmigrasi dan Tenaga Kerja Kabupaten Luwu Utara
 *  Backend / API — Google Apps Script
 * ============================================================================
 *  Cara pakai singkat:
 *  1. Buat Google Spreadsheet baru, jalankan initSetup() sekali dari editor
 *     Apps Script untuk otomatis membuat semua sheet + header + 1 akun admin.
 *  2. Deploy > New deployment > Web app
 *       - Execute as   : Me
 *       - Who has access : Anyone
 *  3. Salin URL Web App, tempel ke variabel API_URL di index.html.
 *  Detail lengkap ada di PANDUAN_SETUP.md
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// KONFIGURASI
// ---------------------------------------------------------------------------
const SHEET = {
  RENSTRA: 'Renstra_RKT',
  PK: 'Perjanjian_Kinerja',
  REALISASI: 'Realisasi_IKU',
  EVALUASI: 'Evaluasi_SAKIP',
  ADMIN: 'Admin',
  LOG: 'Log_Aktivitas'
};

// Skema kolom untuk tiap sheet (urutan = urutan kolom di spreadsheet)
const SCHEMA = {
  RENSTRA: ['ID', 'Tahun', 'Sasaran_Strategis', 'Indikator_Kinerja', 'Target', 'Satuan', 'Program', 'Kegiatan', 'Anggaran_Rencana', 'Keterangan'],
  PK: ['ID', 'Tahun', 'Sasaran_Strategis', 'Indikator_Kinerja', 'Target', 'Satuan', 'Penanggung_Jawab', 'Triwulan', 'Status'],
  REALISASI: ['ID', 'Tahun', 'Triwulan', 'Sasaran_Strategis', 'Indikator_Kinerja', 'Target', 'Realisasi', 'Satuan', 'Capaian_Persen', 'Anggaran_Target', 'Anggaran_Realisasi', 'Serapan_Persen', 'Status_Capaian', 'Updated_At', 'Updated_By'],
  EVALUASI: ['ID', 'Tahun', 'Komponen', 'Nilai', 'Bobot', 'Nilai_Tertimbang', 'Predikat', 'Catatan', 'Tanggal_Evaluasi', 'Evaluator'],
  ADMIN: ['Username', 'Password_Hash', 'Nama', 'Role', 'Status'],
  LOG: ['Timestamp', 'Username', 'Aksi', 'Detail']
};

const TOKEN_TTL_SECONDS = 6 * 60 * 60; // token admin berlaku 6 jam

// ---------------------------------------------------------------------------
// SETUP AWAL — jalankan sekali secara manual dari editor Apps Script
// ---------------------------------------------------------------------------
function initSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(SCHEMA).forEach(function (key) {
    const name = SHEET[key];
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.clear();
    sheet.getRange(1, 1, 1, SCHEMA[key].length).setValues([SCHEMA[key]]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, SCHEMA[key].length).setFontWeight('bold').setBackground('#0F3D2E').setFontColor('#FFFFFF');
  });

  // Hapus sheet default "Sheet1" jika masih ada dan kosong
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  // Buat 1 akun admin default: username "admin" / password "admin123"
  // >>> SEGERA GANTI PASSWORD INI setelah login pertama kali <<<
  const adminSheet = ss.getSheetByName(SHEET.ADMIN);
  if (adminSheet.getLastRow() < 2) {
    adminSheet.appendRow(['admin', hashPassword('admin123'), 'Administrator SAKIP', 'admin', 'aktif']);
  }

  SpreadsheetApp.getUi().alert('Setup selesai. Semua sheet & akun admin default sudah dibuat.\nUsername: admin\nPassword: admin123\n\nSegera ganti password melalui menu Admin di dashboard.');
}

// ---------------------------------------------------------------------------
// CONTOH DATA (opsional) — jalankan manual dari editor Apps Script setelah
// initSetup() jika ingin melihat dashboard terisi data demo langsung.
// Mencakup 2 tahun: 2024 (tahun selesai/historis) dan 2025 (tahun berjalan)
// agar filter Tahun di dashboard bisa langsung dicoba.
// PENTING: Angka & indikator di bawah ini adalah CONTOH ILUSTRATIF berdasarkan
// tugas-fungsi umum Dinas Transmigrasi & Tenaga Kerja (bukan data resmi hasil
// Renstra/LKjIP Disnakertrans Luwu Utara).
//
// CARA MENGUBAH DATA CONTOH: dashboard membaca sheet secara langsung (live),
// jadi Anda bisa mengedit, menambah, atau menghapus baris LANGSUNG di sheet
// Renstra_RKT / Perjanjian_Kinerja / Realisasi_IKU / Evaluasi_SAKIP — cukup
// klik tombol Refresh (⟳) di dashboard atau muat ulang halaman untuk melihat
// perubahannya. Tidak perlu login admin untuk mengubah lewat spreadsheet
// (hanya perubahan lewat dashboard yang memerlukan login).
// ---------------------------------------------------------------------------
function seedContohData() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('Isi Data Contoh?', 'Ini akan menambahkan data contoh untuk tahun 2024 (selesai) dan 2025 (berjalan) ke sheet Renstra_RKT, Perjanjian_Kinerja, Realisasi_IKU, dan Evaluasi_SAKIP. Lanjutkan?', ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const INDIKATOR = [
    { sasaran: 'Meningkatnya Kualitas dan Produktivitas Tenaga Kerja', indikator: 'Persentase Tenaga Kerja Bersertifikat Kompetensi', satuan: '%', program: 'Program Pelatihan Kerja dan Produktivitas Tenaga Kerja', kegiatan: 'Pelatihan Kerja Berbasis Kompetensi', pj: 'Bidang Pelatihan Kerja dan Produktivitas' },
    { sasaran: 'Meningkatnya Penempatan Tenaga Kerja dan Perluasan Kesempatan Kerja', indikator: 'Persentase Pencari Kerja yang Ditempatkan', satuan: '%', program: 'Program Penempatan Tenaga Kerja', kegiatan: 'Bursa Kerja dan Layanan Informasi Pasar Kerja', pj: 'Bidang Penempatan Tenaga Kerja' },
    { sasaran: 'Terwujudnya Hubungan Industrial yang Harmonis', indikator: 'Persentase Perselisihan Hubungan Industrial yang Diselesaikan', satuan: '%', program: 'Program Hubungan Industrial dan Jaminan Sosial Tenaga Kerja', kegiatan: 'Mediasi dan Fasilitasi Penyelesaian Perselisihan HI', pj: 'Bidang Hubungan Industrial' },
    { sasaran: 'Meningkatnya Keberhasilan Penyelenggaraan Ketransmigrasian', indikator: 'Jumlah Transmigran yang Ditempatkan dan Dibina', satuan: 'KK', program: 'Program Perencanaan Kawasan Transmigrasi', kegiatan: 'Fasilitasi Perpindahan dan Penempatan Transmigran', pj: 'Bidang Ketransmigrasian' },
    { sasaran: 'Meningkatnya Akuntabilitas Kinerja Perangkat Daerah', indikator: 'Nilai SAKIP Perangkat Daerah', satuan: 'Nilai', program: 'Program Penunjang Urusan Pemerintahan Daerah', kegiatan: 'Penyusunan LKjIP dan Evaluasi Kinerja Internal', pj: 'Sekretariat' }
  ];

  // Tahun 2024 — SUDAH SELESAI (Triwulan/Tahunan penuh, status final)
  const data2024 = [
    { target: 30, realisasi: 32, anggaranT: 780000000, anggaranR: 765000000, status: 'Tercapai' },
    { target: 40, realisasi: 43, anggaranT: 560000000, anggaranR: 548000000, status: 'Tercapai' },
    { target: 85, realisasi: 88, anggaranT: 280000000, anggaranR: 271000000, status: 'Tercapai' },
    { target: 50, realisasi: 38, anggaranT: 1300000000, anggaranR: 980000000, status: 'Belum Tercapai' },
    { target: 78, realisasi: 79, anggaranT: 165000000, anggaranR: 160000000, status: 'Tercapai' }
  ];

  // Tahun 2025 — BERJALAN (baru sampai Triwulan II, status sebagian dalam proses)
  const data2025 = [
    { target: 35, realisasi: 22, anggaranT: 850000000, anggaranR: 410000000, status: 'Dalam Proses' },
    { target: 45, realisasi: 41, anggaranT: 620000000, anggaranR: 355000000, status: 'Dalam Proses' },
    { target: 90, realisasi: 92, anggaranT: 310000000, anggaranR: 178000000, status: 'Tercapai' },
    { target: 60, realisasi: 28, anggaranT: 1450000000, anggaranR: 520000000, status: 'Belum Tercapai' },
    { target: 80, realisasi: 76, anggaranT: 180000000, anggaranR: 95000000, status: 'Dalam Proses' }
  ];

  function seedTahun(tahun, angka, triwulanRealisasi) {
    INDIKATOR.forEach(function (it, i) {
      const d = angka[i];
      addRow(SHEET.RENSTRA, SCHEMA.RENSTRA, {
        Tahun: tahun, Sasaran_Strategis: it.sasaran, Indikator_Kinerja: it.indikator, Target: d.target, Satuan: it.satuan,
        Program: it.program, Kegiatan: it.kegiatan, Anggaran_Rencana: d.anggaranT, Keterangan: 'Contoh data'
      });
      addRow(SHEET.PK, SCHEMA.PK, {
        Tahun: tahun, Sasaran_Strategis: it.sasaran, Indikator_Kinerja: it.indikator, Target: d.target, Satuan: it.satuan,
        Penanggung_Jawab: it.pj, Triwulan: 'Tahunan', Status: d.status
      });
      const realisasiPayload = {
        Tahun: tahun, Triwulan: triwulanRealisasi, Sasaran_Strategis: it.sasaran, Indikator_Kinerja: it.indikator,
        Target: d.target, Realisasi: d.realisasi, Satuan: it.satuan, Anggaran_Target: d.anggaranT, Anggaran_Realisasi: d.anggaranR,
        Keterangan: 'Contoh data', Updated_By: 'seed'
      };
      computeRealisasiDerived(realisasiPayload);
      addRow(SHEET.REALISASI, SCHEMA.REALISASI, realisasiPayload);
    });
  }

  seedTahun(2024, data2024, 'Tahunan');
  seedTahun(2025, data2025, 'TW II');

  const evaluasiRows = [
    { Tahun: 2024, Komponen: 'Perencanaan Kinerja', Nilai: 82, Bobot: 30, Predikat: 'A', Catatan: 'Contoh data', Tanggal_Evaluasi: '2025-03-10', Evaluator: 'Inspektorat Kabupaten' },
    { Tahun: 2024, Komponen: 'Pengukuran Kinerja', Nilai: 78, Bobot: 25, Predikat: 'BB', Catatan: 'Contoh data', Tanggal_Evaluasi: '2025-03-10', Evaluator: 'Inspektorat Kabupaten' },
    { Tahun: 2024, Komponen: 'Pelaporan Kinerja', Nilai: 80, Bobot: 15, Predikat: 'A', Catatan: 'Contoh data', Tanggal_Evaluasi: '2025-03-10', Evaluator: 'Inspektorat Kabupaten' },
    { Tahun: 2024, Komponen: 'Evaluasi Internal', Nilai: 75, Bobot: 10, Predikat: 'BB', Catatan: 'Contoh data', Tanggal_Evaluasi: '2025-03-10', Evaluator: 'Inspektorat Kabupaten' },
    { Tahun: 2024, Komponen: 'Capaian Kinerja', Nilai: 79, Bobot: 20, Predikat: 'BB', Catatan: 'Contoh data', Tanggal_Evaluasi: '2025-03-10', Evaluator: 'Inspektorat Kabupaten' }
  ];
  evaluasiRows.forEach(function (r) {
    const nilai = parseFloat(r.Nilai) || 0, bobot = parseFloat(r.Bobot) || 0;
    r.Nilai_Tertimbang = Math.round(nilai * (bobot / 100) * 100) / 100;
    addRow(SHEET.EVALUASI, SCHEMA.EVALUASI, r);
  });

  ui.alert('Data contoh berhasil ditambahkan untuk tahun 2024 & 2025. Buka dashboard untuk melihatnya — gunakan filter Tahun di pojok kanan atas. Anda bisa langsung mengedit angka-angka ini di spreadsheet kapan saja; dashboard akan menampilkan perubahan setelah di-refresh.');
}

// =============================================================================
// DATA RESMI — bersumber dari LAKIP 2025 dan Renstra Distransnaker 2025-2029
// Jalankan seedDataResmiLakip2025() SEKALI SAJA (idealnya di spreadsheet yang
// masih kosong / setelah menghapus data contoh ilustratif di atas, supaya
// tidak tercampur dengan angka fiktif).
//
// SUMBER & KEAKURATAN:
// - 10 baris Realisasi_IKU & Perjanjian_Kinerja Tahun 2025 diambil langsung
//   dari LAKIP 2025: Tabel 3.2 (Target & Realisasi Kinerja), Tabel 3.6
//   (Anggaran, Realisasi, Serapan Anggaran per Sasaran Program). Anggaran
//   pada indikator yang berbagi 1 baris anggaran di dokumen asli (indikator
//   #2&#3, serta #4&#5) hanya diisi pada indikator pertama agar total
//   anggaran keseluruhan tetap sama persis dengan LAKIP (Rp 8.189.493.807
//   target / Rp 4.565.922.776 realisasi / 55,75% serapan) — lihat catatan
//   Keterangan pada baris terkait.
// - 15 baris Renstra_RKT (Tingkat Produktivitas Tenaga Kerja, Tingkat
//   Partisipasi Angkatan Kerja, Nilai SAKIP Perangkat Daerah) untuk tahun
//   2025-2029 diambil dari Renstra Distransnaker 2025-2029, Tabel 3.3.
// - 10 baris tambahan Renstra_RKT (Tahun 2025) adalah 10 Indikator Kinerja
//   Utama (IKU) daerah yang menjadi urusan Dinas, dari Renstra Tabel 4.4.
// - Evaluasi_SAKIP SENGAJA TIDAK diisi: kedua dokumen tidak memuat hasil
//   evaluasi SAKIP resmi (nilai per komponen dari Inspektorat/LHE SAKIP)
//   untuk Tahun 2025 — yang tersedia hanyalah TARGET Nilai SAKIP 2025-2029
//   (sudah dimasukkan ke Renstra_RKT). Isi Evaluasi_SAKIP secara manual
//   begitu Anda menerima LHE (Laporan Hasil Evaluasi) SAKIP dari
//   Inspektorat Kabupaten.
// =============================================================================
function seedDataResmiLakip2025() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Isi Data Resmi dari LAKIP 2025 & Renstra 2025-2029?',
    'Ini akan menambahkan data REAL (bukan contoh) ke sheet Renstra_RKT, Perjanjian_Kinerja, dan Realisasi_IKU, bersumber dari dokumen yang Anda unggah. Pastikan sheet belum berisi data contoh ilustratif yang lama (hapus dulu jika perlu) agar tidak tercampur. Lanjutkan?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // 10 Sasaran Program / Indikator Kinerja — LAKIP 2025, Tabel 3.2 & 3.6
  const IKU_2025 = [
    { sasaranProgram: 'Terkelolanya Informasi Tenaga Kerja', indikator: 'Persentase Kabupaten/Kota yang Menyusun Rencana Tenaga Kerja', satuan: '%', target: 10.00, realisasi: 0.00, anggaranT: 15650000, anggaranR: 0, pj: 'Bidang Pemberdayaan Tenaga Kerja', ket: '' },
    { sasaranProgram: 'Meningkatnya Produktivitas Tenaga Kerja', indikator: 'Persentase Tenaga Kerja yang Mendapatkan Pelatihan Berbasis Kompetensi', satuan: '%', target: 59.04, realisasi: 60.00, anggaranT: 3430764600, anggaranR: 189267200, pj: 'Bidang Pemberdayaan Tenaga Kerja / UPTD BLK', ket: '' },
    { sasaranProgram: 'Meningkatnya Produktivitas Tenaga Kerja', indikator: 'Jumlah Lulusan Pelatihan Vokasi yang Bersertifikat', satuan: 'Orang', target: 87, realisasi: 48, anggaranT: 0, anggaranR: 0, pj: 'Bidang Pemberdayaan Tenaga Kerja / UPTD BLK', ket: 'Anggaran tergabung dengan indikator "Persentase Tenaga Kerja yang Mendapatkan Pelatihan Berbasis Kompetensi" (Rp 3.430.764.600 / Rp 189.267.200, LAKIP Tabel 3.6)' },
    { sasaranProgram: 'Meningkatnya Penempatan Tenaga Kerja', indikator: 'Persentase Tenaga Kerja yang Ditempatkan di Dalam Negeri', satuan: '%', target: 5.32, realisasi: 6.49, anggaranT: 22960400, anggaranR: 22457400, pj: 'Bidang Pemberdayaan Tenaga Kerja', ket: '' },
    { sasaranProgram: 'Meningkatnya Penempatan Tenaga Kerja', indikator: 'Persentase Tenaga Kerja Informal', satuan: '%', target: 69.50, realisasi: 71.75, anggaranT: 0, anggaranR: 0, pj: 'Bidang Pemberdayaan Tenaga Kerja', ket: 'Anggaran tergabung dengan indikator "Persentase Tenaga Kerja yang Ditempatkan di Dalam Negeri" (Rp 22.960.400 / Rp 22.457.400, LAKIP Tabel 3.6)' },
    { sasaranProgram: 'Meningkatnya Pekerja Indonesia yang Terlindungi', indikator: 'Jumlah Pekerja pada Perusahaan yang Menerapkan Perlindungan Hak-Hak Pekerja dan Dialog Sosial', satuan: 'Orang', target: 1514, realisasi: 2583, anggaranT: 199090000, anggaranR: 20084900, pj: 'Bidang Hubungan Industrial', ket: '' },
    { sasaranProgram: 'Meningkatnya Pelaksanaan Transmigrasi', indikator: 'Persentase Program Transmigrasi yang Dilaksanakan', satuan: '%', target: 10.00, realisasi: 10.00, anggaranT: 41810000, anggaranR: 36596000, pj: 'Bidang Penyiapan Kawasan dan Pengembangan', ket: '' },
    { sasaranProgram: 'Meningkatnya Kualitas Pembangunan Kawasan Transmigrasi', indikator: 'Persentase Luas Kawasan Transmigrasi yang Berkembang', satuan: '%', target: 100.00, realisasi: 50.00, anggaranT: 0, anggaranR: 0, pj: 'Bidang Pengembangan Kawasan Transmigrasi', ket: 'Tidak ada alokasi anggaran tersendiri di Tahun 2025 (LAKIP Tabel 3.6)' },
    { sasaranProgram: 'Meningkatnya Pemberdayaan dan Kapasitas Transmigran dalam Pengembangan Kawasan Transmigrasi', indikator: 'Persentase Transmigran yang Dibina dan Diberdayakan', satuan: '%', target: 8.62, realisasi: 8.62, anggaranT: 141439000, anggaranR: 122800000, pj: 'Bidang Pengembangan Kawasan Transmigrasi', ket: '' },
    { sasaranProgram: 'Meningkatnya Tata Kelola Pemerintahan Perangkat Daerah', indikator: 'Persentase Capaian Program Penunjang Urusan Pemerintahan Daerah Kabupaten/Kota', satuan: '%', target: 100.00, realisasi: 83.00, anggaranT: 4337779807, anggaranR: 4174717276, pj: 'Sekretariat', ket: '' }
  ];

  const SASARAN_STRATEGIS_2025 = 'Meningkatnya Tingkat Partisipasi Angkatan Kerja';

  IKU_2025.forEach(function (it) {
    // Perjanjian Kinerja
    const capaian = (parseFloat(it.target) > 0) ? (parseFloat(it.realisasi) / parseFloat(it.target)) * 100 : 0;
    addRow(SHEET.PK, SCHEMA.PK, {
      Tahun: 2025, Sasaran_Strategis: it.sasaranProgram, Indikator_Kinerja: it.indikator,
      Target: it.target, Satuan: it.satuan, Penanggung_Jawab: it.pj, Triwulan: 'Tahunan',
      Status: capaian >= 100 ? 'Tercapai' : 'Belum Tercapai'
    });

    // Realisasi Kinerja & Anggaran (Capaian %, Serapan %, Status dihitung otomatis oleh sistem)
    const realisasiPayload = {
      Tahun: 2025, Triwulan: 'Tahunan', Sasaran_Strategis: it.sasaranProgram, Indikator_Kinerja: it.indikator,
      Target: it.target, Realisasi: it.realisasi, Satuan: it.satuan,
      Anggaran_Target: it.anggaranT, Anggaran_Realisasi: it.anggaranR,
      Keterangan: 'Sumber: LAKIP Distransnaker Tahun 2025, Tabel 3.2 & 3.6' + (it.ket ? '. ' + it.ket : ''),
      Updated_By: 'seed-resmi'
    };
    computeRealisasiDerived(realisasiPayload);
    addRow(SHEET.REALISASI, SCHEMA.REALISASI, realisasiPayload);
  });

  // 3 Indikator Tujuan/Sasaran Strategis OPD, Tahun 2025-2029 — Renstra Tabel 3.3
  const RENSTRA_STRATEGIS = [
    { indikator: 'Tingkat Produktivitas Tenaga Kerja', satuan: 'Rp/Orang', sasaran: 'Meningkatnya Kesempatan dan Produktivitas Tenaga Kerja', targets: { 2025: 57374109, 2026: 58808461, 2027: 60278673, 2028: 61785640, 2029: 63330281 } },
    { indikator: 'Tingkat Partisipasi Angkatan Kerja', satuan: '%', sasaran: 'Meningkatnya Tingkat Partisipasi Angkatan Kerja', targets: { 2025: 69.47, 2026: 70.08, 2027: 73.41, 2028: 73.39, 2029: 73.14 } },
    { indikator: 'Nilai SAKIP Perangkat Daerah', satuan: 'Nilai', sasaran: 'Meningkatnya Akuntabilitas Kinerja Perangkat Daerah', targets: { 2025: 68.55, 2026: 70.06, 2027: 71.57, 2028: 73.08, 2029: 74.59 } }
  ];
  RENSTRA_STRATEGIS.forEach(function (it) {
    Object.keys(it.targets).forEach(function (tahun) {
      addRow(SHEET.RENSTRA, SCHEMA.RENSTRA, {
        Tahun: tahun, Sasaran_Strategis: it.sasaran, Indikator_Kinerja: it.indikator, Target: it.targets[tahun], Satuan: it.satuan,
        Program: 'Renstra Dinas Transmigrasi dan Tenaga Kerja 2025-2029', Kegiatan: 'Pencapaian Tujuan dan Sasaran Strategis OPD',
        Anggaran_Rencana: 0, Keterangan: 'Sumber: Renstra Distransnaker 2025-2029, Tabel 3.3'
      });
    });
  });

  // 10 Indikator Kinerja Utama (IKU) Daerah urusan Dinas, Tahun 2025 — Renstra Tabel 4.4
  const IKU_DAERAH_2025 = [
    { indikator: 'Tingkat Pengangguran Terbuka', target: 2.30, aspek: 'Aspek Kesejahteraan Masyarakat' },
    { indikator: 'Persentase Penduduk Usia 15 Tahun ke Atas yang Bekerja menurut Pendidikan Tertinggi yang Ditamatkan (SD)', target: 34.00, aspek: 'Aspek Kesejahteraan Masyarakat' },
    { indikator: 'Persentase Penduduk Usia 15 Tahun ke Atas yang Bekerja menurut Pendidikan Tertinggi yang Ditamatkan (SMP)', target: 16.00, aspek: 'Aspek Kesejahteraan Masyarakat' },
    { indikator: 'Persentase Penduduk Usia 15 Tahun ke Atas yang Bekerja menurut Pendidikan Tertinggi yang Ditamatkan (SMA)', target: 34.50, aspek: 'Aspek Kesejahteraan Masyarakat' },
    { indikator: 'Persentase Penduduk Usia 15 Tahun ke Atas yang Bekerja menurut Pendidikan Tertinggi yang Ditamatkan (Perguruan Tinggi)', target: 15.50, aspek: 'Aspek Kesejahteraan Masyarakat' },
    { indikator: 'Cakupan Kepesertaan Jaminan Sosial Ketenagakerjaan', target: 62.00, aspek: 'Aspek Kesejahteraan Masyarakat' },
    { indikator: 'Persentase Tenaga Kerja Formal', target: 30.50, aspek: 'Aspek Kesejahteraan Masyarakat' },
    { indikator: 'Persentase Penyandang Disabilitas Bekerja di Sektor Formal', target: 0.01, aspek: 'Aspek Kesejahteraan Masyarakat' },
    { indikator: 'Tingkat Partisipasi Angkatan Kerja Perempuan', target: 54.00, aspek: 'Aspek Daya Saing Daerah' },
    { indikator: 'Tingkat Partisipasi Angkatan Kerja (RPJMD)', target: 70.00, aspek: 'Aspek Daya Saing Daerah' }
  ];
  IKU_DAERAH_2025.forEach(function (it) {
    addRow(SHEET.RENSTRA, SCHEMA.RENSTRA, {
      Tahun: 2025, Sasaran_Strategis: it.aspek, Indikator_Kinerja: it.indikator, Target: it.target, Satuan: '%',
      Program: 'Indikator Kinerja Utama Daerah (RPJMD Kab. Luwu Utara 2025-2029)', Kegiatan: '',
      Anggaran_Rencana: 0, Keterangan: 'Sumber: Renstra Distransnaker 2025-2029, Tabel 4.4'
    });
  });

  ui.alert('Data resmi dari LAKIP 2025 & Renstra 2025-2029 berhasil ditambahkan:\n- 10 baris Perjanjian Kinerja & Realisasi IKU (Tahun 2025)\n- 25 baris Renstra_RKT (Tahun 2025-2029)\n\nEvaluasi_SAKIP sengaja dibiarkan kosong karena kedua dokumen tidak memuat hasil evaluasi SAKIP resmi — isi manual saat LHE SAKIP dari Inspektorat sudah tersedia.');
}

// ---------------------------------------------------------------------------
// ENTRY POINT — GET (baca data)
// ---------------------------------------------------------------------------
function doGet(e) {
  try {
    const action = (e.parameter.action || 'getSummary').toString();
    let result;

    switch (action) {
      case 'getRenstra':
        result = getSheetData(SHEET.RENSTRA);
        break;
      case 'getPK':
        result = getSheetData(SHEET.PK);
        break;
      case 'getRealisasi':
        result = getSheetData(SHEET.REALISASI);
        break;
      case 'getEvaluasi':
        result = getSheetData(SHEET.EVALUASI);
        break;
      case 'getSummary':
        result = buildDashboardSummary(e.parameter.tahun);
        break;
      case 'getAll':
        result = {
          renstra: getSheetData(SHEET.RENSTRA),
          pk: getSheetData(SHEET.PK),
          realisasi: getSheetData(SHEET.REALISASI),
          evaluasi: getSheetData(SHEET.EVALUASI),
          summary: buildDashboardSummary(e.parameter.tahun)
        };
        break;
      default:
        return jsonOut({ success: false, message: 'Action tidak dikenal: ' + action });
    }

    return jsonOut({ success: true, action: action, data: result });
  } catch (err) {
    return jsonOut({ success: false, message: 'Terjadi kesalahan server: ' + err.message });
  }
}

// ---------------------------------------------------------------------------
// ENTRY POINT — POST (login, tambah/ubah/hapus data)
// Catatan CORS: frontend WAJIB mengirim body dengan header
// 'Content-Type': 'text/plain;charset=utf-8' (isi tetap string JSON) supaya
// browser tidak melakukan preflight OPTIONS, karena Apps Script Web App
// tidak bisa merespons preflight request.
// ---------------------------------------------------------------------------
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = (body.action || '').toString();

    // Aksi yang tidak memerlukan login
    if (action === 'login') return jsonOut(handleLogin(body));

    // Semua aksi di bawah ini wajib token admin yang valid
    const auth = verifyToken(body.token);
    if (!auth.valid) {
      return jsonOut({ success: false, message: 'Sesi tidak valid atau sudah habis. Silakan login kembali.' });
    }

    let result;
    switch (action) {
      case 'logout':
        CacheService.getScriptCache().remove('token_' + body.token);
        result = { loggedOut: true };
        break;

      case 'addRenstra': result = addRow(SHEET.RENSTRA, SCHEMA.RENSTRA, body.payload); break;
      case 'updateRenstra': result = updateRow(SHEET.RENSTRA, SCHEMA.RENSTRA, body.payload); break;
      case 'deleteRenstra': requireAdminRole(auth); result = deleteRow(SHEET.RENSTRA, body.id); break;

      case 'addPK': result = addRow(SHEET.PK, SCHEMA.PK, body.payload); break;
      case 'updatePK': result = updateRow(SHEET.PK, SCHEMA.PK, body.payload); break;
      case 'deletePK': requireAdminRole(auth); result = deleteRow(SHEET.PK, body.id); break;

      case 'addRealisasi': result = addRealisasi(body.payload, auth.username); break;
      case 'updateRealisasi': result = updateRealisasi(body.payload, auth.username); break;
      case 'deleteRealisasi': requireAdminRole(auth); result = deleteRow(SHEET.REALISASI, body.id); break;

      case 'addEvaluasi': result = addEvaluasi(body.payload); break;
      case 'updateEvaluasi': result = updateRow(SHEET.EVALUASI, SCHEMA.EVALUASI, body.payload); break;
      case 'deleteEvaluasi': requireAdminRole(auth); result = deleteRow(SHEET.EVALUASI, body.id); break;

      case 'changePassword': result = changePassword(auth.username, body.oldPassword, body.newPassword); break;

      // --- Kelola akun admin/operator (khusus role 'admin') ---
      case 'listAdminUsers': requireAdminRole(auth); result = getAdminUsers(); break;
      case 'addAdminUser': requireAdminRole(auth); result = addAdminUser(body.payload); break;
      case 'updateAdminUser': requireAdminRole(auth); result = updateAdminUser(body.payload); break;
      case 'deleteAdminUser': requireAdminRole(auth); result = deleteAdminUser(body.id, auth.username); break;

      default:
        return jsonOut({ success: false, message: 'Action tidak dikenal: ' + action });
    }

    logActivity(auth.username, action, JSON.stringify(body.payload || body.id || {}));
    return jsonOut({ success: true, action: action, data: result });
  } catch (err) {
    return jsonOut({ success: false, message: 'Terjadi kesalahan server: ' + err.message });
  }
}

// ---------------------------------------------------------------------------
// AUTH
// ---------------------------------------------------------------------------
function handleLogin(body) {
  const sheet = getSheet(SHEET.ADMIN);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const uCol = header.indexOf('Username');
  const pCol = header.indexOf('Password_Hash');
  const nCol = header.indexOf('Nama');
  const rCol = header.indexOf('Role');
  const sCol = header.indexOf('Status');

  for (let i = 1; i < data.length; i++) {
    if (data[i][uCol] === body.username) {
      if (data[i][sCol] && data[i][sCol].toString().toLowerCase() !== 'aktif') {
        return { success: false, message: 'Akun tidak aktif. Hubungi administrator.' };
      }
      if (data[i][pCol] === hashPassword(body.password || '')) {
        const token = Utilities.getUuid();
        const role = (data[i][rCol] || 'operator').toString();
        CacheService.getScriptCache().put('token_' + token, JSON.stringify({ username: body.username, role: role }), TOKEN_TTL_SECONDS);
        logActivity(body.username, 'login', 'Login berhasil (' + role + ')');
        return {
          success: true,
          token: token,
          nama: data[i][nCol],
          role: role,
          expiresIn: TOKEN_TTL_SECONDS
        };
      }
    }
  }
  return { success: false, message: 'Username atau password salah.' };
}

function verifyToken(token) {
  if (!token) return { valid: false };
  const raw = CacheService.getScriptCache().get('token_' + token);
  if (!raw) return { valid: false };
  try {
    const session = JSON.parse(raw);
    return { valid: true, username: session.username, role: session.role || 'operator' };
  } catch (e) {
    return { valid: false };
  }
}

// Hanya role 'admin' yang boleh menghapus data & mengelola akun admin/operator.
// Role 'operator' hanya boleh menambah & mengubah data (tidak menghapus).
function requireAdminRole(auth) {
  if (auth.role !== 'admin') {
    throw new Error('Aksi ini hanya diizinkan untuk akun dengan peran Admin.');
  }
}

function hashPassword(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + '::esakip-luwuutara-salt');
  return digest.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function changePassword(username, oldPassword, newPassword) {
  if (!newPassword || newPassword.length < 6) throw new Error('Password baru minimal 6 karakter.');
  const sheet = getSheet(SHEET.ADMIN);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const uCol = header.indexOf('Username');
  const pCol = header.indexOf('Password_Hash');

  for (let i = 1; i < data.length; i++) {
    if (data[i][uCol] === username) {
      if (data[i][pCol] !== hashPassword(oldPassword || '')) throw new Error('Password lama tidak sesuai.');
      sheet.getRange(i + 1, pCol + 1).setValue(hashPassword(newPassword));
      return { changed: true };
    }
  }
  throw new Error('Akun tidak ditemukan.');
}

// ---------------------------------------------------------------------------
// KELOLA AKUN ADMIN / OPERATOR (khusus role 'admin')
//   Role 'admin'    : akses penuh, termasuk hapus data & kelola akun.
//   Role 'operator' : hanya bisa tambah & ubah data (tidak bisa hapus,
//                      tidak bisa kelola akun).
// ---------------------------------------------------------------------------
function getAdminUsers() {
  const sheet = getSheet(SHEET.ADMIN);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const header = values[0];
  return values.slice(1).filter(function (r) { return r.join('') !== ''; }).map(function (row) {
    return { Username: row[header.indexOf('Username')], Nama: row[header.indexOf('Nama')], Role: row[header.indexOf('Role')], Status: row[header.indexOf('Status')] };
    // Password_Hash sengaja tidak dikirim ke frontend
  });
}

function addAdminUser(payload) {
  if (!payload.Username || !payload.Password || !payload.Nama) throw new Error('Username, Password, dan Nama wajib diisi.');
  if (payload.Password.length < 6) throw new Error('Password minimal 6 karakter.');
  const sheet = getSheet(SHEET.ADMIN);
  const data = sheet.getDataRange().getValues();
  const uCol = data[0].indexOf('Username');
  for (let i = 1; i < data.length; i++) {
    if (data[i][uCol] === payload.Username) throw new Error('Username sudah digunakan.');
  }
  sheet.appendRow([payload.Username, hashPassword(payload.Password), payload.Nama, payload.Role || 'operator', payload.Status || 'aktif']);
  return { added: true, Username: payload.Username };
}

function updateAdminUser(payload) {
  if (!payload.Username) throw new Error('Username wajib diisi.');
  const sheet = getSheet(SHEET.ADMIN);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const uCol = header.indexOf('Username');
  for (let i = 1; i < data.length; i++) {
    if (data[i][uCol] === payload.Username) {
      if (payload.Nama) sheet.getRange(i + 1, header.indexOf('Nama') + 1).setValue(payload.Nama);
      if (payload.Role) sheet.getRange(i + 1, header.indexOf('Role') + 1).setValue(payload.Role);
      if (payload.Status) sheet.getRange(i + 1, header.indexOf('Status') + 1).setValue(payload.Status);
      if (payload.Password) {
        if (payload.Password.length < 6) throw new Error('Password minimal 6 karakter.');
        sheet.getRange(i + 1, header.indexOf('Password_Hash') + 1).setValue(hashPassword(payload.Password));
      }
      return { updated: true, Username: payload.Username };
    }
  }
  throw new Error('Akun tidak ditemukan.');
}

function deleteAdminUser(username, currentUsername) {
  if (username === currentUsername) throw new Error('Tidak dapat menghapus akun yang sedang digunakan untuk login.');
  const sheet = getSheet(SHEET.ADMIN);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const uCol = header.indexOf('Username');
  const rCol = header.indexOf('Role');

  const totalAdminAktif = data.slice(1).filter(function (r) { return r[rCol] === 'admin'; }).length;

  for (let i = 1; i < data.length; i++) {
    if (data[i][uCol] === username) {
      if (data[i][rCol] === 'admin' && totalAdminAktif <= 1) {
        throw new Error('Tidak dapat menghapus satu-satunya akun dengan peran Admin.');
      }
      sheet.deleteRow(i + 1);
      return { deleted: true, Username: username };
    }
  }
  throw new Error('Akun tidak ditemukan.');
}

// ---------------------------------------------------------------------------
// CRUD GENERIK
// ---------------------------------------------------------------------------
function getSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" tidak ditemukan. Jalankan initSetup() terlebih dahulu.');
  return sheet;
}

function getSheetData(name) {
  const sheet = getSheet(name);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const header = values[0];
  const rows = values.slice(1).filter(function (r) { return r.join('') !== ''; });
  return rows.map(function (row) {
    const obj = {};
    header.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function generateId(prefix) {
  return prefix + '-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+8', 'yyyyMMddHHmmss') + Math.floor(Math.random() * 900 + 100);
}

function addRow(sheetName, schema, payload) {
  const sheet = getSheet(sheetName);
  const id = generateId(sheetName.substring(0, 3).toUpperCase());
  payload.ID = id;
  const row = schema.map(function (col) { return payload[col] !== undefined ? payload[col] : ''; });
  sheet.appendRow(row);
  return { ID: id };
}

function updateRow(sheetName, schema, payload) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  const idCol = schema.indexOf('ID');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === payload.ID) {
      const row = schema.map(function (col) { return payload[col] !== undefined ? payload[col] : data[i][schema.indexOf(col)]; });
      sheet.getRange(i + 1, 1, 1, schema.length).setValues([row]);
      return { updated: true, ID: payload.ID };
    }
  }
  throw new Error('Data dengan ID ' + payload.ID + ' tidak ditemukan.');
}

function deleteRow(sheetName, id) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { deleted: true, ID: id };
    }
  }
  throw new Error('Data dengan ID ' + id + ' tidak ditemukan.');
}

// ---------------------------------------------------------------------------
// REALISASI IKU — hitung otomatis Capaian % dan Serapan %
// ---------------------------------------------------------------------------
function computeRealisasiDerived(payload) {
  const target = parseFloat(payload.Target) || 0;
  const realisasi = parseFloat(payload.Realisasi) || 0;
  const anggaranTarget = parseFloat(payload.Anggaran_Target) || 0;
  const anggaranRealisasi = parseFloat(payload.Anggaran_Realisasi) || 0;

  payload.Capaian_Persen = target > 0 ? Math.round((realisasi / target) * 10000) / 100 : 0;
  payload.Serapan_Persen = anggaranTarget > 0 ? Math.round((anggaranRealisasi / anggaranTarget) * 10000) / 100 : 0;
  payload.Status_Capaian = payload.Capaian_Persen >= 100 ? 'Tercapai' : 'Belum Tercapai';
  payload.Updated_At = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+8', 'yyyy-MM-dd HH:mm:ss');
  return payload;
}

function addRealisasi(payload, username) {
  payload.Updated_By = username;
  computeRealisasiDerived(payload);
  return addRow(SHEET.REALISASI, SCHEMA.REALISASI, payload);
}

function updateRealisasi(payload, username) {
  payload.Updated_By = username;
  computeRealisasiDerived(payload);
  return updateRow(SHEET.REALISASI, SCHEMA.REALISASI, payload);
}

// ---------------------------------------------------------------------------
// EVALUASI SAKIP — hitung Nilai_Tertimbang otomatis
// ---------------------------------------------------------------------------
function addEvaluasi(payload) {
  const nilai = parseFloat(payload.Nilai) || 0;
  const bobot = parseFloat(payload.Bobot) || 0;
  payload.Nilai_Tertimbang = Math.round(nilai * (bobot / 100) * 100) / 100;
  return addRow(SHEET.EVALUASI, SCHEMA.EVALUASI, payload);
}

// ---------------------------------------------------------------------------
// RINGKASAN DASHBOARD
// ---------------------------------------------------------------------------
function predikatFromScore(score) {
  if (score >= 90) return 'AA';
  if (score >= 80) return 'A';
  if (score >= 70) return 'BB';
  if (score >= 60) return 'B';
  if (score >= 50) return 'CC';
  if (score >= 30) return 'C';
  return 'D';
}

function buildDashboardSummary(tahun) {
  let realisasi = getSheetData(SHEET.REALISASI);
  let pk = getSheetData(SHEET.PK);
  let evaluasi = getSheetData(SHEET.EVALUASI);

  if (tahun) {
    realisasi = realisasi.filter(function (r) { return String(r.Tahun) === String(tahun); });
    pk = pk.filter(function (r) { return String(r.Tahun) === String(tahun); });
    evaluasi = evaluasi.filter(function (r) { return String(r.Tahun) === String(tahun); });
  }

  const totalIku = realisasi.length;
  const tercapai = realisasi.filter(function (r) { return r.Status_Capaian === 'Tercapai'; }).length;

  const avgCapaian = totalIku > 0
    ? Math.round((realisasi.reduce(function (s, r) { return s + (parseFloat(r.Capaian_Persen) || 0); }, 0) / totalIku) * 100) / 100
    : 0;

  const totalAnggaranTarget = realisasi.reduce(function (s, r) { return s + (parseFloat(r.Anggaran_Target) || 0); }, 0);
  const totalAnggaranRealisasi = realisasi.reduce(function (s, r) { return s + (parseFloat(r.Anggaran_Realisasi) || 0); }, 0);
  const avgSerapan = totalAnggaranTarget > 0
    ? Math.round((totalAnggaranRealisasi / totalAnggaranTarget) * 10000) / 100
    : 0;

  // Predikat: pakai nilai evaluasi resmi jika ada, kalau tidak estimasi dari capaian
  let predikat, sumberPredikat, totalNilaiEvaluasi;
  if (evaluasi.length > 0) {
    totalNilaiEvaluasi = Math.round(evaluasi.reduce(function (s, r) { return s + (parseFloat(r.Nilai_Tertimbang) || 0); }, 0) * 100) / 100;
    predikat = predikatFromScore(totalNilaiEvaluasi);
    sumberPredikat = 'evaluasi';
  } else {
    totalNilaiEvaluasi = avgCapaian;
    predikat = predikatFromScore(avgCapaian);
    sumberPredikat = 'estimasi_capaian';
  }

  return {
    totalIku: totalIku,
    ikuTercapai: tercapai,
    ikuBelumTercapai: totalIku - tercapai,
    capaianKinerjaPersen: avgCapaian,
    serapanAnggaranPersen: avgSerapan,
    totalAnggaranTarget: totalAnggaranTarget,
    totalAnggaranRealisasi: totalAnggaranRealisasi,
    totalPK: pk.length,
    predikatSakip: predikat,
    nilaiSakip: totalNilaiEvaluasi,
    sumberPredikat: sumberPredikat,
    chartTargetVsRealisasi: realisasi.map(function (r) {
      return {
        indikator: r.Indikator_Kinerja,
        target: parseFloat(r.Target) || 0,
        realisasi: parseFloat(r.Realisasi) || 0,
        satuan: r.Satuan
      };
    })
  };
}

// ---------------------------------------------------------------------------
// LOG AKTIVITAS
// ---------------------------------------------------------------------------
function logActivity(username, action, detail) {
  try {
    const sheet = getSheet(SHEET.LOG);
    sheet.appendRow([new Date(), username, action, detail]);
  } catch (e) {
    // jangan sampai gagal logging menghentikan proses utama
  }
}

// ---------------------------------------------------------------------------
// OUTPUT JSON
// ---------------------------------------------------------------------------
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
