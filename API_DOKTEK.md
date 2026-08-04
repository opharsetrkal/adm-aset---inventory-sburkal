/**
 * API_DOKTEK — Code.gs
 * READ-ONLY. Tidak ada operasi tulis (setValue/appendRow/clear) ke spreadsheet manapun.
 * Sumber data:
 *   tab=backlog  -> sheet "BACKLOG WIG 2 CORPORATE"          (kendala <=2025, segmen Corporate saja)
 *   tab=berjalan -> sheet "BACKLOG BERJALAN 03 JUNI 2026"    (kendala 2026)
 *
 * Kolom dibaca BY POSITION (index array), bukan by header-name lookup.
 * Alasan: kedua sheet punya header duplikat (KETERANGAN 2x di Berjalan,
 * Mon PA Matjas.WO Matjas.namaPerusahaan 2x di WIG Corporate). Baca by-name
 * akan collision (nilai kolom pertama ketimpa kolom kedua secara silent).
 *
 * Kalau kolom di sheet source digeser manual, index di bawah HARUS di-update manual juga.
 * Ini trade-off sadar: aman dari collision duplikat header, tapi rawan kalau struktur
 * kolom sheet berubah tanpa pemberitahuan.
 */

var SPREADSHEET_ID = '150D4fhP92uB2jXoq1gQl5hNclNvxyF_vQ-DjEyMd6DQ';

var SHEET_WIG_CORPORATE = 'BACKLOG WIG 2 CORPORATE';
var SHEET_BERJALAN      = 'BACKLOG BERJALAN 03 JUNI 2026';

// ── Index kolom (0-based) — BACKLOG WIG 2 CORPORATE, 51 kolom ──────────────
var COL_WIG = {
  service_id: 0,
  pa_number: 1,
  customer_name: 3,
  terminating: 5,
  namaKP: 7,
  created_at: 15,
  status_terinput: 22,
  paMatjasIDPA: 23,      // Mon PA Matjas.IDPA -- fallback ke-2 utk PA Number
  endCustomerName: 21,   // fallback ke-2 untuk nama customer
  customerMatjas: 30,    // Mon PA Matjas.CUSTOMERNAME -- fallback ke-3
  namaPTL: 32,            // Mon PA Matjas.namaPTL
  vendor: 35,             // Mon PA Matjas.WO Matjas.namaMitra
  acluan: 46,              // Keterangan
  paLama: 47                // PA LAMA -- fallback ke-3 utk PA Number
};

// ── Index kolom (0-based) — BACKLOG BERJALAN 03 JUNI 2026, 26 kolom ────────
var COL_BERJALAN = {
  service_id: 0,
  pa_number: 4,
  customer_name: 7,   // CUSTOMERNAME
  terminating: 8,
  namaPTL: 12,
  namaKP: 16,
  created_at: 3,
  status_terinput: 17,
  vendor: 14,          // namaMitra
  acluan: 22,          // KETERANGAN posisi ke-23 (1-based) -- LOCKED sesuai keputusan, bukan duplikat di posisi 26
  noPA: 23             // NO PA -- fallback ke-2 utk PA Number
  // catatan: sheet ini TIDAK punya kolom terpisah setara "Mon PA Matjas.CUSTOMERNAME"/"PA LAMA",
  // customerMatjas/paLama sengaja tidak dimapping di sini -> otomatis kosong di readSheet()
};

// ── SLA threshold, satuan HARI ──────────────────────────────────────────
var SLA_NORMAL_MAX    = 12; // 0-12 hari = Normal
var SLA_PERHATIAN_MAX = 30; // 13-30 hari = Perhatian
                             // >=31 hari = Kritis

function doGet(e) {
  try {
    var tab = (e && e.parameter && e.parameter.tab) || 'backlog';
    var result = (tab === 'berjalan') ? readSheet(SHEET_BERJALAN, COL_BERJALAN)
                                       : readSheet(SHEET_WIG_CORPORATE, COL_WIG);
    result.tab = tab;
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    var errOut = { ok: false, error: String(err && err.message ? err.message : err) };
    return ContentService.createTextOutput(JSON.stringify(errOut))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Baca kolom dengan aman -- kalau field tidak dimapping di colMap tab tertentu
// (mis. customerMatjas di BACKLOG BERJALAN), balikin '' bukan error/undefined leak.
function pick(colMap, key, r) {
  var idx = colMap[key];
  if (idx === undefined) return '';
  return r[idx];
}

// "Kosong" mencakup blank string DAN tanda strip "-" yang sering dipakai sebagai placeholder manual di sheet.
function isBlank(v) {
  if (v === null || v === undefined) return true;
  var s = String(v).trim();
  return s === '' || s === '-';
}

// Fallback berurutan: customer_name -> endCustomerName -> Mon PA Matjas.CUSTOMERNAME.
// Ambil yang PERTAMA tidak kosong/tidak "-". Kembalikan juga sumbernya untuk traceability (tooltip di UI).
function resolveCustomer(customerName, endCustomerName, customerMatjas) {
  if (!isBlank(customerName))    return { value: String(customerName).trim(),    source: 'customer_name' };
  if (!isBlank(endCustomerName)) return { value: String(endCustomerName).trim(), source: 'endCustomerName' };
  if (!isBlank(customerMatjas))  return { value: String(customerMatjas).trim(),  source: 'Mon PA Matjas.CUSTOMERNAME' };
  return { value: '', source: '' };
}

// Fallback berurutan utk PA Number. WIG: pa_number -> Mon PA Matjas.IDPA -> PA LAMA.
// Berjalan: pa_number -> NO PA (altBLabel dikosongkan krn tab ini tidak punya fallback ke-3).
function resolvePANumber(paNumber, altA, altALabel, altB, altBLabel) {
  if (!isBlank(paNumber)) return { value: String(paNumber).trim(), source: 'pa_number' };
  if (!isBlank(altA))     return { value: String(altA).trim(),     source: altALabel };
  if (!isBlank(altB))     return { value: String(altB).trim(),     source: altBLabel };
  return { value: '', source: '' };
}

function readSheet(sheetName, colMap) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet tidak ditemukan: ' + sheetName);

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    return { ok: true, rows: [], totalPA: 0, kpOptions: [], ptlOptions: [], keteranganOptions: [], tahunOptions: [], excludedTerinput: 0 };
  }

  // Baca sekali, batch — tidak ada get/setValue per-cell (hindari beban Apps Script)
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var rows = [];
  var kpSet = {}, ptlSet = {}, ketSet = {}, tahunSet = {};
  var paSet = {};
  var terinputCount = 0;
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  for (var i = 0; i < values.length; i++) {
    var r = values[i];

    var serviceId  = r[colMap.service_id];
    var paNumber   = r[colMap.pa_number];
    var namaKP     = r[colMap.namaKP];
    var createdRaw = r[colMap.created_at];
    var statusRaw  = r[colMap.status_terinput];
    var vendor     = r[colMap.vendor];
    var acluan     = r[colMap.acluan];
    var terminating   = pick(colMap, 'terminating', r);
    var namaPTL        = pick(colMap, 'namaPTL', r);
    var customerNameRaw     = pick(colMap, 'customer_name', r);
    var endCustomerNameRaw  = pick(colMap, 'endCustomerName', r);
    var customerMatjasRaw   = pick(colMap, 'customerMatjas', r);
    var customerResolved = resolveCustomer(customerNameRaw, endCustomerNameRaw, customerMatjasRaw);

    var paMatjasIDPARaw = pick(colMap, 'paMatjasIDPA', r);
    var paLamaRaw       = pick(colMap, 'paLama', r);
    var noPARaw         = pick(colMap, 'noPA', r);
    // WIG: pa_number -> Mon PA Matjas.IDPA -> PA LAMA | Berjalan: pa_number -> NO PA
    var paResolved = resolvePANumber(paNumber, paMatjasIDPARaw, 'Mon PA Matjas.IDPA', paLamaRaw || noPARaw, paLamaRaw ? 'PA LAMA' : 'NO PA');

    // Skip baris kosong total (tidak ada service_id maupun PA number hasil fallback)
    if (!serviceId && !paResolved.value) continue;

    // Filter Status Terinput = kasus SUDAH SELESAI, keluarkan dari backlog aktif.
    // Modul ini scope-nya murni kendala yang masih perlu tindak lanjut.
    var statusClean = statusRaw ? String(statusRaw).trim().toLowerCase() : '';
    if (statusClean === 'terinput') {
      terinputCount++;
      continue;
    }

    var createdAt = parseDateCell(createdRaw);
    var aging = null, slaStatus = '—';
    if (createdAt) {
      var createdMid = new Date(createdAt);
      createdMid.setHours(0, 0, 0, 0);
      aging = Math.floor((today - createdMid) / 86400000);
      if (aging < 0) aging = 0; // guard: created_at di masa depan (data entry error)
      if (aging <= SLA_NORMAL_MAX) slaStatus = 'Normal';
      else if (aging <= SLA_PERHATIAN_MAX) slaStatus = 'Perhatian';
      else slaStatus = 'Kritis';
    }

    var namaKPStr = namaKP ? String(namaKP).trim() : '—';
    var namaPTLStr = namaPTL ? String(namaPTL).trim() : '';
    var acluanStr = acluan ? String(acluan).trim() : '';
    var tahun = createdAt ? createdAt.getFullYear() : null;

    rows.push({
      service_id: serviceId ? String(serviceId).trim() : '',
      pa_number: paResolved.value || '—',
      pa_numberSource: paResolved.source,
      customer: customerResolved.value || '—',
      customerSource: customerResolved.source,
      namaKP: namaKPStr,
      created_at: createdAt ? Utilities.formatDate(createdAt, 'GMT+7', 'yyyy-MM-dd') : '',
      aging: aging,
      acluan: acluanStr,
      slaStatus: slaStatus,
      status: statusRaw ? String(statusRaw).trim() : '',
      vendor: vendor ? String(vendor).trim() : '', // kosong dibiarkan '' -- diisi manual belakangan, bukan tugas script
      terminating: terminating ? String(terminating).trim() : '',
      namaPTL: namaPTLStr
    });

    if (namaKPStr) kpSet[namaKPStr] = true;
    if (namaPTLStr) ptlSet[namaPTLStr] = true;
    if (acluanStr) ketSet[acluanStr] = true;
    if (tahun) tahunSet[tahun] = true;
    if (paResolved.value) paSet[paResolved.value] = true;
  }

  return {
    ok: true,
    rows: rows,
    totalPA: Object.keys(paSet).length,
    kpOptions: Object.keys(kpSet).sort(),
    ptlOptions: Object.keys(ptlSet).sort(),
    keteranganOptions: Object.keys(ketSet).sort(),
    tahunOptions: Object.keys(tahunSet).sort().reverse(),
    excludedTerinput: terinputCount
  };
}

// Handle cell yang sudah Date object (umum di Apps Script utk kolom bertipe tanggal)
// maupun string tanggal manual (fallback).
function parseDateCell(val) {
  if (!val) return null;
  if (Object.prototype.toString.call(val) === '[object Date]') {
    if (isNaN(val.getTime())) return null;
    return val;
  }
  var parsed = new Date(val);
  if (isNaN(parsed.getTime())) return null;
  return parsed;
}
