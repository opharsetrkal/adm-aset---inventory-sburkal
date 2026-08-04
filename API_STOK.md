// ============================================================
// Aset & Inventory SBU Reg Kalimantan
// Apps Script — Stok Material (Read-Only)
//
// CARA DEPLOY:
//   1. Buka Google Sheets stok → Extensions → Apps Script
//   2. Hapus semua code lama, paste code ini
//   3. Ganti SHEET_STOK dengan nama tab sheet kamu
//   4. Klik Save (Ctrl+S)
//   5. Jalankan testGetData() dulu — pastikan log muncul data
//   6. Deploy → New deployment
//      - Type        : Web app
//      - Execute as  : Me
//      - Who has access: Anyone
//   7. Klik Deploy → Authorize jika diminta → Allow
//   8. Copy URL → paste ke stok.html baris API_URL
// ============================================================

const SS_STOK    = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_STOK = 'MB52'; // <-- GANTI dengan nama tab sheet stok kamu

// Batch yang dianggap RUSAK — tidak masuk stok aktif
const BATCH_RUSAK_GS = ['RUSAK-TL', 'RUSAK-L'];

// ── ENTRY POINT ───────────────────────────────────────────
function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const data = getData(e && e.parameter ? e.parameter : {});
    output.setContent(JSON.stringify({ ok: true, ...data }));
  } catch(err) {
    output.setContent(JSON.stringify({ ok: false, error: err.message }));
  }
  return output;
}

// ── READ DATA ─────────────────────────────────────────────
function getData(params) {
  const sheet = SS_STOK.getSheetByName(SHEET_STOK);
  if (!sheet) throw new Error('Sheet "' + SHEET_STOK + '" tidak ditemukan. Cek nama tab.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { headers:[], rows:[], total:0 };

  // Baca header dari baris 1 — dinamis, otomatis baca kolom baru
  const actualHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => String(h).trim())
    .filter(h => h !== ''); // abaikan kolom tanpa header

  const usedCols = actualHeaders.length;

  // Baca semua data mulai baris 2
  const raw = sheet.getRange(2, 1, lastRow - 1, usedCols).getValues();

  // Cari index kolom Tgl Update untuk konversi khusus
  const tglIdx = actualHeaders.indexOf('Tgl Update');

  // Ubah tiap baris array → objek { 'Header': 'Nilai' }
  const rows = raw
    .map((r, i) => rowToObj(r, actualHeaders, i + 2, tglIdx))
    .filter(r => r['Material'] !== undefined && r['Material'] !== '');

  // Filter opsional dari URL parameter — guard jika params undefined
  params = params || {};
  let filtered = rows;
  if (params.batch) filtered = filtered.filter(r => r['Batch'] === params.batch);
  if (params.sloc)  filtered = filtered.filter(r => String(r['Storage Location']||'') === params.sloc);

  // Ambil nilai unik untuk dropdown filter — paksa String, filter kosong
  const uniq = col => [...new Set(
    filtered
      .map(r => {
        let v = r[col];
        // Potong Tgl Update jadi YYYY-MM-DD saja
        if (col === 'Tgl Update' && typeof v === 'string' && v.length > 10) {
          try { v = Utilities.formatDate(new Date(v), Session.getScriptTimeZone(), 'yyyy-MM-dd'); } catch(e){}
        }
        return v;
      })
      .filter(v => v !== null && v !== undefined && v !== '')
      .map(v => String(v))
  )].sort();

  return {
    headers:          actualHeaders,
    rows:             filtered,
    total:            filtered.length,
    batchOptions:     uniq('Batch'),
    brandOptions:     uniq('Material - Brand'),
    katSapOptions:    uniq('Kategori - SAP'),
    katTeknikOptions: uniq('Kategori - Teknik'),
    slocGudangOptions:uniq('Sloc - Gudang'),
    slocKpOptions:    uniq('Sloc - KP'),
    movingOptions:    uniq('Kategori - Moving'),
    tglUpdateOptions: uniq('Tgl Update'),
  };
}

// ── HELPER: array baris → objek ───────────────────────────
function rowToObj(row, headers, rowIndex, tglIdx) {
  const obj = { _rowIndex: rowIndex };
  headers.forEach((h, i) => {
    let val = row[i];
    // Kolom Tgl Update — paksa jadi YYYY-MM-DD apapun formatnya
    if (i === tglIdx) {
      if (val instanceof Date && !isNaN(val)) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else if (typeof val === 'string' && val.length > 10) {
        try {
          const d = new Date(val);
          if (!isNaN(d)) {
            val = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          }
        } catch(e) {}
        // Jika masih panjang, paksa dengan SpreadsheetApp date parsing
        if (typeof val === 'string' && val.length > 10) {
          try {
            // Format: "Wed Jun 10 2026 15:00:00 GMT+0800..."
            // Ambil bagian "Jun 10 2026" lalu konversi
            const parts = val.split(' ');
            if (parts.length >= 4) {
              const months = {Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,
                              Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
              const mon = months[parts[1]];
              const day = parseInt(parts[2]);
              const yr  = parseInt(parts[3]);
              if (mon && day && yr) {
                val = yr + '-' + String(mon).padStart(2,'0') + '-' + String(day).padStart(2,'0');
              }
            }
          } catch(e) {}
        }
      } else if (typeof val === 'number' && val > 0) {
        // Nilai serial Excel/Sheets
        try {
          const d = new Date((val - 25569) * 86400 * 1000);
          val = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } catch(e) {}
      }
      obj[h] = val ? String(val) : '';
      return;
    }
    // Paksa semua nilai jadi string kecuali angka (untuk kalkulasi qty & value)
    const numCols = ['Unrestricted','Value Unrestricted','Blocked','Value BlockedStock',
                     'Returns','Value Rets Blocked','In Quality Insp.','Value in QualInsp.',
                     'Restricted-Use Stock','Value Restricted','Transit and Transfer',
                     'Val. in Trans./Tfr'];
    if (numCols.includes(h)) {
      val = (val === '' || val === null || val === undefined) ? 0 : Number(val)||0;
    } else {
      val = (val !== null && val !== undefined) ? String(val) : '';
    }
    obj[h] = val;
  });
  return obj;
}

// ── TEST ──────────────────────────────────────────────────
// Pilih fungsi ini di dropdown editor lalu klik Run
function testGetData() {
  const result = getData({});
  Logger.log('=== HASIL TEST ===');
  Logger.log('Total baris      : ' + result.total);
  Logger.log('Kolom (headers)  : ' + JSON.stringify(result.headers));
  Logger.log('Batch options    : ' + JSON.stringify(result.batchOptions));
  Logger.log('Brand options    : ' + JSON.stringify(result.brandOptions));
  Logger.log('Kat SAP options  : ' + JSON.stringify(result.katSapOptions));
  Logger.log('Kat Teknik opt   : ' + JSON.stringify(result.katTeknikOptions));
  Logger.log('Sloc Gudang opt  : ' + JSON.stringify(result.slocGudangOptions));
  Logger.log('Tgl Update opt   : ' + JSON.stringify(result.tglUpdateOptions));
  Logger.log('Sample row 1     : ' + JSON.stringify(result.rows[0]));
  Logger.log('Sample row 2     : ' + JSON.stringify(result.rows[1]));
}
