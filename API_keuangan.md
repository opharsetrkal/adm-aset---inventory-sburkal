// ============================================================
// ADMINISTRASI KEUANGAN INVENTORY
// Google Apps Script — Web App (Backend/API)
// 
// APA INI?
// File ini adalah "server" sederhana yang membaca data dari
// Google Sheets dan mengirimkannya ke dashboard web kita.
// 
// CARA DEPLOY:
//   1. Buka Google Sheets → Extensions → Apps Script
//   2. Hapus kode default, paste semua kode ini
//   3. Klik Save (Ctrl+S)
//   4. Jalankan fungsi testGetData() dulu untuk cek
//   5. Deploy → New Deployment → Web App
//      - Execute as  : Me
//      - Who has access: Anyone
//   6. Copy URL yang muncul → paste ke index.html
// ============================================================


// ── KONFIGURASI ───────────────────────────────────────────
// Nama spreadsheet aktif (file yang sedang dibuka)
const SS = SpreadsheetApp.getActiveSpreadsheet();

// Nama tab/sheet yang berisi data billing
// HARUS SAMA PERSIS dengan nama tab di Google Sheets
const SHEET_NAME = 'Rekap Inv';

// Daftar header kolom — urutan HARUS sama dengan kolom di Sheets
// Ini digunakan untuk membaca data dengan benar
const HEADERS = [
  'No',
  'Deskripsi Material',
  'IO/PRK',
  'No. Nodin',
  'Tgl Nodin',
  'Dari',
  'Tujuan',
  'Tgl AWB',
  'AWB',
  'Berat (Kg)',
  'No INV',
  'Service Billing',
  'Mitra',
  'Nominal (Rp)',
  'Status',
  'Keterangan'
];

// STATUS_LIST tidak lagi hardcode.
// Diambil dinamis dari nilai unik kolom Status di Sheets.
// Fungsi getStatusList() dipanggil saat getData() dijalankan.
// Dengan ini, status baru di Sheets otomatis muncul di dashboard.


// ── ENTRY POINT ───────────────────────────────────────────
// doGet() adalah fungsi yang otomatis dipanggil ketika
// ada request dari browser (seperti fetch() di JavaScript).
// 
// Parameter 'e' berisi informasi dari request tersebut,
// termasuk query parameter seperti ?bulan=2026-02
function doGet(e) {
  // Buat objek output dengan format JSON
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    // Ambil data, kirim dengan status ok: true
    const data = getData(e.parameter || {});
    output.setContent(JSON.stringify({ ok: true, ...data }));
  } catch (err) {
    // Jika ada error, kirim pesan error
    output.setContent(JSON.stringify({ ok: false, error: err.message }));
  }

  return output;
}


// ── FUNGSI UTAMA: BACA DATA ───────────────────────────────
// Fungsi ini membaca semua data dari sheet 'Rekap Inv'
// dan mengembalikannya dalam format yang siap dipakai dashboard
//
// Parameter 'params' adalah filter dari URL, contoh:
//   ?bulan=2026-02 → hanya tampilkan data bulan Feb 2026
function getData(params) {

  // Cari sheet bernama 'Rekap Inv'
  const sheet = SS.getSheetByName(SHEET_NAME);
  
  // Jika sheet tidak ditemukan, lempar error
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" tidak ditemukan. Cek nama tab di Sheets.');

  // Hitung jumlah baris yang berisi data
  const lastRow = sheet.getLastRow();

  // Jika kurang dari 2 baris (berarti hanya ada header, tidak ada data)
  // kembalikan data kosong
  if (lastRow < 2) {
    return {
      headers:       HEADERS,
      rows:          [],
      allRows:       [],
      statusOptions: [],
      mitraOptions:  [],
      monthOptions:  [],
      topRoutes:     [],
      total:         0,
      summary:       buildSummary([], [])
    };
  }

  // Ambil semua data mulai baris ke-2 (baris 1 = header)
  // getRange(baris_mulai, kolom_mulai, jumlah_baris, jumlah_kolom)
  const raw = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  // Ubah setiap baris (array) menjadi objek { 'Kolom': 'Nilai' }
  // Filter baris kosong — pastikan ada Deskripsi Material
  const allRows = raw
    .map((r, i) => rowToObj(r, i + 2))
    .filter(r => r['Deskripsi Material']);

  // ── FILTER DATA ─────────────────────────────────────────
  // Filter ini digunakan oleh reminder card "Lihat detail"
  // agar tabel langsung menampilkan data yang relevan
  let filtered = [...allRows]; // copy semua data dulu

  // Filter berdasarkan parameter URL (jika ada)
  if (params.bulan)  filtered = filtered.filter(r => monthKey(r['Tgl AWB']) === params.bulan);
  if (params.status) filtered = filtered.filter(r => r['Status'] === params.status);
  if (params.mitra)  filtered = filtered.filter(r => r['Mitra'] === params.mitra);

  // ── DROPDOWN OPTIONS ─────────────────────────────────────
  // Ambil nilai unik untuk dropdown filter di dashboard
  // [...new Set(array)] = hapus duplikat
  const mitras  = [...new Set(allRows.map(r => r['Mitra']).filter(Boolean))].sort();
  const months  = [...new Set(allRows.map(r => monthKey(r['Tgl AWB'])).filter(Boolean))].sort();
  const years   = [...new Set(allRows.map(r => yearOf(r['Tgl AWB'])).filter(Boolean))].sort();

  // ── TOP 3 RUTE ───────────────────────────────────────────
  // Hitung frekuensi setiap kombinasi Dari → Tujuan
  const routeCount = {};
  allRows.forEach(r => {
    const dari   = r['Dari'];
    const tujuan = r['Tujuan'];
    if (dari && tujuan) {
      const key        = dari + ' → ' + tujuan;
      routeCount[key]  = (routeCount[key] || 0) + 1;
    }
  });

  // Urutkan dari terbanyak, ambil 3 teratas
  const topRoutes = Object.entries(routeCount)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(([route, count]) => ({ route, count }));

  // Kembalikan semua data yang dibutuhkan dashboard
  return {
    headers:       HEADERS,
    rows:          filtered,   // data setelah filter (untuk tabel)
    allRows:       allRows,    // semua data (untuk KPI global)
    statusOptions: [...new Set(allRows.map(r => r["Status"]).filter(Boolean))],
    mitraOptions:  mitras,
    monthOptions:  months,
    yearOptions:   years,
    topRoutes:     topRoutes,
    total:         filtered.length,
    summary:       buildSummary(filtered, allRows)
  };
}


// ── FUNGSI: HITUNG SUMMARY ────────────────────────────────
// Menghitung angka-angka untuk KPI cards dan charts
//
// 'filtered' = data sesuai filter aktif (untuk Card 1 & 2)
// 'allRows'  = semua data tanpa filter (untuk Card 3-6 global)
function buildSummary(filtered, allRows) {

  // Helper: ambil nilai nominal sebagai angka
  const nom = r => parseFloat(r['Nominal (Rp)']) || 0;

  // ── KPI GLOBAL (tidak terpengaruh filter) ───────────────

  // Total nominal yang BELUM dibayar (semua status kecuali Paid)
  const globalOutstanding = allRows
    .filter(r => r['Status'] !== 'Paid')
    .reduce((total, r) => total + nom(r), 0);

  // Total nominal yang SUDAH dibayar
  const globalPaid = allRows
    .filter(r => r['Status'] === 'Paid')
    .reduce((total, r) => total + nom(r), 0);

  // Hitung aging (hari sejak Tgl AWB) untuk reminder
  // Hanya untuk yang belum Paid
  const agingBuckets = { normal: 0, r1: 0, r2: 0 };
  const r1Detail = []; // simpan detail untuk "Lihat detail" reminder
  const r2Detail = [];

  allRows.forEach(r => {
    if (r['Status'] === 'Paid') return; // skip yang sudah selesai
    const days = calcAging(r['Tgl AWB']);
    if (days > 14) {
      agingBuckets.r2++;
      r2Detail.push(r);
    } else if (days >= 8) {
      agingBuckets.r1++;
      r1Detail.push(r);
    } else {
      agingBuckets.normal++;
    }
  });

  // Hitung jumlah per status — DINAMIS dari data aktual
  // Tidak hardcode, semua nilai unik di kolom Status akan terhitung
  const byStatus = {};
  allRows.forEach(r => {
    const s = r['Status'];
    if (s) byStatus[s] = (byStatus[s] || 0) + 1;
  });

  // ── DATA UNTUK CHARTS (ikut filter aktif) ───────────────

  // Nominal per bulan (untuk line chart tren)
  const byMonth = {};
  filtered.forEach(r => {
    const k = monthKey(r['Tgl AWB']);
    if (k) byMonth[k] = (byMonth[k] || 0) + nom(r);
  });

  // Nominal per mitra (untuk bar chart)
  const byMitra = {};
  filtered.forEach(r => {
    const m = r['Mitra'];
    if (m) byMitra[m] = (byMitra[m] || 0) + nom(r);
  });

  return {
    // Global — tidak ikut filter
    globalOutstanding,
    globalPaid,
    agingBuckets,
    byStatus,
    // Filtered — ikut filter aktif
    filteredTotal:   filtered.length,
    filteredNominal: filtered.reduce((s, r) => s + nom(r), 0),
    // Charts
    byMonth,
    byMitra
  };
}


// ── FUNGSI HELPER ─────────────────────────────────────────
// Fungsi-fungsi kecil yang digunakan berulang kali

// Ubah array baris Sheets menjadi objek JavaScript
// Contoh: ['001', 'Kabel 24c', ...] → { 'No': '001', 'Deskripsi Material': 'Kabel 24c', ... }
function rowToObj(row, rowIndex) {
  const obj = { _rowIndex: rowIndex }; // simpan nomor baris untuk referensi
  HEADERS.forEach((h, i) => {
    let val = row[i];
    // Google Sheets menyimpan tanggal sebagai objek Date
    // Ubah ke format string YYYY-MM-DD agar mudah diproses
    if (val instanceof Date) {
      val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    obj[h] = (val !== undefined && val !== null) ? val : '';
  });
  return obj;
}

// Ambil format YYYY-MM dari sebuah tanggal
// Contoh: '2026-02-18' → '2026-02'
function monthKey(tgl) {
  if (!tgl) return null;
  const d = new Date(tgl);
  if (isNaN(d)) return null;
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

// Ambil tahun dari sebuah tanggal
// Contoh: '2026-02-18' → '2026'
function yearOf(tgl) {
  if (!tgl) return null;
  const d = new Date(tgl);
  return isNaN(d) ? null : String(d.getFullYear());
}

// Hitung selisih hari antara Tgl AWB dan hari ini
// Digunakan untuk menentukan level reminder
function calcAging(tgl) {
  if (!tgl) return 0;
  const d = new Date(tgl);
  if (isNaN(d)) return 0;
  // Math.abs = nilai absolut (selalu positif)
  // 864e5 = 86400000 milidetik = 1 hari
  return Math.ceil(Math.abs(new Date() - d) / 864e5);
}


// ── FUNGSI TEST ───────────────────────────────────────────
// Jalankan fungsi ini dari editor Apps Script untuk
// memastikan semuanya berjalan sebelum deploy
// Cara: pilih 'testGetData' di dropdown, klik Run
function testGetData() {
  const result = getData({});
  Logger.log('=== HASIL TEST ===');
  Logger.log('Total baris data : ' + result.total);
  Logger.log('Top 3 rute       : ' + JSON.stringify(result.topRoutes));
  Logger.log('Status options   : ' + JSON.stringify(result.statusOptions));
  Logger.log('Mitra options    : ' + JSON.stringify(result.mitraOptions));
  Logger.log('Global Paid      : Rp ' + result.summary.globalPaid.toLocaleString());
  Logger.log('Global Outstand  : Rp ' + result.summary.globalOutstanding.toLocaleString());
  Logger.log('Aging R1         : ' + result.summary.agingBuckets.r1 + ' pengiriman');
  Logger.log('Aging R2         : ' + result.summary.agingBuckets.r2 + ' pengiriman');
  Logger.log('By Status        : ' + JSON.stringify(result.summary.byStatus));
}
