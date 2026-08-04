// ============================================================
// Aset & Inventory SBU Reg Kalimantan
// Apps Script — Penyerapan Material (MB51 SBURKAL)
//
// CARA DEPLOY:
//   1. Buka Google Sheets MB51 SBURKAL → Extensions → Apps Script
//   2. Hapus semua code lama, paste code ini
//   3. Ganti SHEET_NAME jika nama tab berbeda
//   4. Jalankan testGetData() dulu — pastikan log muncul data
//   5. Deploy → New deployment → Web app → Anyone → Deploy
//   6. Copy URL → paste ke revisi-260614.html baris API_PENYERAPAN
// ============================================================

const SS_P      = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_NAME_P = 'Sheet1';

// Kata kunci batch yang dianggap RUSAK
const RUSAK_KEYWORDS = ['rusak'];

// ── ENTRY POINT ───────────────────────────────────────────
function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const data   = getPenyerapanData(params);
    output.setContent(JSON.stringify({ ok: true, ...data }));
  } catch(err) {
    output.setContent(JSON.stringify({ ok: false, error: err.message }));
  }
  return output;
}

// ── READ DATA ─────────────────────────────────────────────
function getPenyerapanData(params) {
  params = params || {};
  const sheet = SS_P.getSheetByName(SHEET_NAME_P);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME_P + '" tidak ditemukan.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { headers:[], rows:[], total:0 };

  // Header dinamis — handle duplikat dengan suffix
  const rawHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => String(h).trim())
    .filter(h => h !== '');
  const seen = {};
  const headers = rawHeaders.map(h => {
    if (!h) return h;
    if (seen[h] !== undefined) {
      seen[h]++;
      return h + '_' + seen[h];
    }
    seen[h] = 0;
    return h;
  });
  const usedCols = headers.length;

  // Baca semua data via getValues (untuk numerik)
  const raw = sheet.getRange(2, 1, lastRow - 1, usedCols).getValues();

  // Baca kolom tanggal via getDisplayValues — lebih andal untuk Date cells
  const rawDisplay = sheet.getRange(2, 1, lastRow - 1, usedCols).getDisplayValues();
  const dateColNames = ['Posting Date','Document Date','Entry Date'];
  const dateColIdxs  = headers.map((h,i)=>dateColNames.includes(h)?i:-1).filter(i=>i>=0);

  // Override nilai tanggal di raw dengan display values yang sudah format string
  dateColIdxs.forEach(ci=>{
    for(let ri=0;ri<raw.length;ri++){
      const dv = rawDisplay[ri][ci]; // format DD/MM/YYYY dari Sheets
      if(dv) raw[ri][ci] = dv;      // override ke display value
    }
  });

  // Cari index kolom kunci
  const idx = {
    qty:      headers.indexOf('Quantity'),
    amount:   headers.indexOf('Amount in LC'),
    batch:    headers.indexOf('Batch'),
    posting:  headers.indexOf('Posting Date'),
    mvtText:  headers.indexOf('Movement Type Text'),
    mvt:      headers.indexOf('Movement Type'),
    katTeknik:headers.indexOf('Kategori - Teknik'),
    katSap:   headers.indexOf('Kategori - SAP'),
    slocKp:   headers.indexOf('Sloc - KP'),
    slocGudang:headers.indexOf('Sloc - Gudang'),
    brand:    headers.indexOf('Material - Brand'),
    moving:   headers.indexOf('Kategori - Moving'),
    material: headers.indexOf('Material'),
    desc:     headers.indexOf('Material Description'),
  };

  // Convert baris → objek
  const rows = raw
    .filter(r => r[idx.material] !== '' && r[idx.material] !== null)
    .map((r, i) => rowToObjP(r, headers, idx, i + 2));

  // Fungsi cek rusak
  const isRusak = batch => {
    const b = String(batch||'').toLowerCase();
    return RUSAK_KEYWORDS.some(k => b.includes(k));
  };

  // Klasifikasi tiap baris
  rows.forEach(r => {
    const qty = parseFloat(r['Quantity']) || 0;
    const batch = r['Batch'] || '';
    if (qty < 0) {
      r['_type'] = 'penyerapan';
      r['_qty_abs'] = Math.abs(qty);
    } else if (qty > 0 && isRusak(batch)) {
      r['_type'] = 'kembali_rusak';
      r['_qty_abs'] = qty;
    } else if (qty > 0) {
      r['_type'] = 'kembali_baik';
      r['_qty_abs'] = qty;
    } else {
      r['_type'] = 'other';
      r['_qty_abs'] = 0;
    }
  });

  // Unique options untuk filter
  const uniq = col => [...new Set(
    rows.map(r => r[col])
      .filter(v => v !== null && v !== undefined && v !== '')
      .map(v => {
        let s = String(v);
        // Normalisasi DD/MM/YYYY → YYYY-MM-DD untuk sorting benar
        if (s.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
          const p = s.split('/');
          s = p[2]+'-'+p[1]+'-'+p[0];
        }
        return s;
      })
  )].sort();

  // Tabel rasio per Kat. Teknik × Sloc KP
  const rasioMap = {};
  rows.forEach(r => {
    const kt  = r['Kategori - Teknik'] || '—';
    const kp  = r['Sloc - KP']         || '—';
    const key = kt + '||' + kp;
    if (!rasioMap[key]) rasioMap[key] = {
      katTeknik: kt, slocKp: kp,
      serap: 0, kembaliBaik: 0, kembaliRusak: 0
    };
    if (r['_type'] === 'penyerapan')   rasioMap[key].serap       += r['_qty_abs'];
    if (r['_type'] === 'kembali_baik') rasioMap[key].kembaliBaik += r['_qty_abs'];
    if (r['_type'] === 'kembali_rusak')rasioMap[key].kembaliRusak+= r['_qty_abs'];
  });
  const rasioTable = Object.values(rasioMap).map(d => ({
    ...d,
    rasioKembaliBaik:  d.serap > 0 ? Math.round(d.kembaliBaik  / d.serap * 100) : 0,
    rasioKembaliRusak: d.serap > 0 ? Math.round(d.kembaliRusak / d.serap * 100) : 0,
  })).sort((a, b) => b.rasioKembaliBaik - a.rasioKembaliBaik);

  return {
    headers,
    rows,
    total:          rows.length,
    rasioTable,
    // Filter options
    mvtOptions:      uniq('Movement Type Text'),
    slocKpOptions:   uniq('Sloc - KP'),
    slocGudangOptions:uniq('Sloc - Gudang'),
    brandOptions:    uniq('Material - Brand'),
    katTeknikOptions:uniq('Kategori - Teknik'),
    katSapOptions:   uniq('Kategori - SAP'),
    batchOptions:    uniq('Batch'),
    movingOptions:   uniq('Kategori - Moving'),
    postingOptions:  uniq('Posting Date'),
  };
}

// ── HELPER: baris → objek ─────────────────────────────────
function rowToObjP(row, headers, idx, rowIndex) {
  const obj = { _rowIndex: rowIndex };
  headers.forEach((h, i) => {
    let val = row[i];
    // Konversi Date object dulu sebelum apapun
    if (val instanceof Date && !isNaN(val)) {
      val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }

    // Kolom tanggal (Posting Date, Document Date, Entry Date) → YYYY-MM-DD
    const dateCols = ['Posting Date','Document Date','Entry Date'];
    if (dateCols.includes(h)) {
      if (val instanceof Date && !isNaN(val)) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else if (typeof val === 'string' && val.length > 10) {
        try {
          // Format "Thu Jan 01 2026 00:00:00 GMT+0800..."
          const months = {Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,
                          Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
          const parts = val.split(' ');
          // Cari pola: [DayName] [MonName] [DD] [YYYY]
          if (parts.length >= 4 && months[parts[1]]) {
            const mon = months[parts[1]], day = parseInt(parts[2]), yr = parseInt(parts[3]);
            if (mon && day && yr)
              val = yr + '-' + String(mon).padStart(2,'0') + '-' + String(day).padStart(2,'0');
          } else {
            const d = new Date(val);
            if (!isNaN(d)) val = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          }
        } catch(e) {}
      }
      // Format DD/MM/YYYY → YYYY-MM-DD
      if (typeof val === 'string' && val.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
        const p = val.split('/');
        val = p[2]+'-'+p[1]+'-'+p[0];
      }
      obj[h] = val ? String(val) : '';
      return;
    }
    // Kolom numerik
    const numCols = ['Quantity','Amount in LC','Item','Movement Type','Reservation'];
    if (numCols.includes(h)) {
      val = (val === '' || val === null || val === undefined) ? 0 : Number(val) || 0;
    } else {
      val = (val !== null && val !== undefined) ? String(val) : '';
    }
    obj[h] = val;
  });
  return obj;
}

// ── TEST ──────────────────────────────────────────────────
function testGetData() {
  const result = getPenyerapanData({});
  Logger.log('=== PENYERAPAN TEST ===');
  Logger.log('Total baris      : ' + result.total);
  Logger.log('Headers          : ' + JSON.stringify(result.headers));
  Logger.log('MVT options      : ' + JSON.stringify(result.mvtOptions));
  Logger.log('Sloc KP options  : ' + JSON.stringify(result.slocKpOptions));
  Logger.log('Kat Teknik opt   : ' + JSON.stringify(result.katTeknikOptions));
  Logger.log('Posting options  : ' + JSON.stringify(result.postingOptions.slice(0,5)));
  Logger.log('Rasio table (3)  : ' + JSON.stringify(result.rasioTable.slice(0,3)));
  Logger.log('Sample row 1     : ' + JSON.stringify(result.rows[0]));
}
