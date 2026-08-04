// ============================================================
// Aset & Inventory SBU Reg Kalimantan
// Apps Script — WIG Penguatan Aset
// v2 — lazy load: summary cepat, detail on-demand
// v3 — tambah mode=all-detail untuk Resume (1 request, semua LM)
// ============================================================

const SS_WIG = SpreadsheetApp.openById('1s6MgNFTYR58uG0TTrgW3vO7KY7cRtWCg37S1XuBexa8');

const KP_LIST = ['BALIKPAPAN','KALIMANTAN BARAT','KALIMANTAN SELATAN','KALIMANTAN TENGAH','KALIMANTAN TIMUR'];

const LM_CONFIG = [
  { wig:1, lm:1, sheet:'WIG 1 LM 1 - POP',             nama:'LM 1 — POP',              part:1 },
  { wig:1, lm:2, sheet:'WIG 1 LM 2 - OLT',             nama:'LM 2 — OLT',              part:1 },
  { wig:1, lm:3, sheet:'WIG 1 LM 3 - Cable',           nama:'LM 3 — Cable',            part:1 },
  { wig:1, lm:4, sheet:'WIG 1 LM 4 - FAT',             nama:'LM 4 — FAT',              part:1 },
  { wig:1, lm:5, sheet:'WIG 1 LM 5 - JB',              nama:'LM 5 — JB',               part:1 },
  { wig:1, lm:6, sheet:'WIG 1 LM 6 & 7- Pole & Tower', nama:'LM 6 — Pole',             part:1 },
  { wig:1, lm:7, sheet:'WIG 1 LM 6 & 7- Pole & Tower', nama:'LM 7 — Tower',            part:2 },
  { wig:2, lm:1, sheet:'WIG 2 LM 1 - Backlog Retail',  nama:'LM 1 — Backlog Retail',   part:1 },
  { wig:2, lm:2, sheet:'WIG 2 LM 2 - Backlog B2B',     nama:'LM 2 — Backlog B2B',      part:1 },
  { wig:3, lm:1, sheet:'WIG 3 LM 1 - Berjalan Retail', nama:'LM 1 — Berjalan Retail',  part:1 },
  { wig:3, lm:2, sheet:'WIG 3 LM 2 - Berjalan B2B',    nama:'LM 2 — Berjalan B2B',     part:1 },
  { wig:4, lm:1, sheet:'WIG 4 LM 1 - SID BB',          nama:'LM 1 — SID Link Backbone',part:1 },
];

const WIG_NAMES = {
  1: 'WIG 1 — Meningkatkan Akurasi Data Entitas',
  2: 'WIG 2 — Kelengkapan SID Backlog',
  3: 'WIG 3 — Kelengkapan SID Berjalan',
  4: 'WIG 4 — Meningkatkan Akurasi Data Link Backbone',
};

// ── ENTRY POINT ───────────────────────────────────────────
function doGet(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const mode   = params.mode || 'summary';
    const wigNo  = parseInt(params.wig  || 0);
    const lmNo   = parseInt(params.lm   || 0);

    let data;
    if (mode === 'detail' && wigNo && lmNo) {
      // Lazy: hanya baca sheet LM yang diminta
      data = getLmDetail(wigNo, lmNo);
    } else if (mode === 'all-detail') {
      // Resume: baca semua sheet LM sekaligus (1 request, agregat total)
      data = getAllDetail();
    } else {
      // Default: hanya baca Dashboard — cepat
      data = getSummary();
    }
    out.setContent(JSON.stringify({ ok:true, ...data }));
  } catch(err) {
    out.setContent(JSON.stringify({ ok:false, error:err.message }));
  }
  return out;
}

// ── SUMMARY — hanya Dashboard ─────────────────────────────
function getSummary() {
  const dash = readDashboard();

  const wigs = [1,2,3,4].map(wigNo => ({
    wig:  wigNo,
    nama: WIG_NAMES[wigNo],
    lms:  LM_CONFIG
      .filter(c => c.wig === wigNo)
      .map(c => {
        const row = dash.find(r => r.wig===wigNo && r.lm===c.lm) || {};
        return {
          lm:           c.lm,
          nama:         c.nama,
          pctSbu:       row.pctSbu       || 0,
          totalSatuan:  row.totalSatuan  || 0,
          status:       row.status       || '—',
          pctMinggu:    row.pctMinggu    || 0,
          statusMinggu: row.statusMinggu || '—',
          pctPerKp:     row.pctPerKp     || {},
        };
      })
  }));

  const allLms  = wigs.flatMap(w => w.lms);
  const onTrack = allLms.filter(l => l.status === '✅').length;
  const offTrack= allLms.filter(l => l.status === '❌').length;

  return { mode:'summary', wigs, onTrack, offTrack, totalLm: allLms.length };
}

// ── ALL-DETAIL — baca semua 10 sheet LM sekaligus untuk Resume ──
// 1 HTTP request, GAS loop internal. Lebih ringan dari 10x fetch client.
function getAllDetail() {
  const dash = readDashboard();

  const lms = LM_CONFIG.map(cfg => {
    const detail = readLmSheet(cfg);
    const dashRow = dash.find(r => r.wig===cfg.wig && r.lm===cfg.lm) || {};
    return {
      wig:           cfg.wig,
      wigNama:       WIG_NAMES[cfg.wig],
      lm:            cfg.lm,
      nama:          cfg.nama,
      status:        dashRow.status       || '—',
      pctSbu:        dashRow.pctSbu       || 0,
      pctMinggu:     dashRow.pctMinggu    || 0,
      statusMinggu:  dashRow.statusMinggu || '—',
      totalTarget:   detail.totalTarget,
      totalReal:     detail.totalReal,
      pencapaian:    detail.pencapaian,
      targetMingguBerjalan: detail.targetMingguBerjalan,
      kpData:        detail.kpData,
      kontributor:   detail.kontributor,
    };
  });

  const totalTargetAll = lms.reduce((s,l) => s + (l.totalTarget||0), 0);
  const totalRealAll   = lms.reduce((s,l) => s + (l.totalReal||0), 0);
  const onTrack  = lms.filter(l => l.status === '✅').length;
  const offTrack = lms.filter(l => l.status === '❌').length;

  return {
    mode:'all-detail', lms,
    totalTargetAll, totalRealAll,
    onTrack, offTrack, totalLm: lms.length
  };
}

// ── DETAIL — baca satu sheet LM saja ─────────────────────
function getLmDetail(wigNo, lmNo) {
  const cfg = LM_CONFIG.find(c => c.wig===wigNo && c.lm===lmNo);
  if (!cfg) throw new Error(`LM tidak ditemukan: WIG${wigNo} LM${lmNo}`);
  const detail = readLmSheet(cfg);
  return { mode:'detail', wig:wigNo, lm:lmNo, ...detail };
}

// ── BACA DASHBOARD ────────────────────────────────────────
function readDashboard() {
  const ws = SS_WIG.getSheetByName('Dashboard');
  if (!ws) throw new Error('Sheet Dashboard tidak ditemukan');

  const rows = ws.getDataRange().getValues();
  const results = [];
  let wigNo = 0, lmNo = 0;

  for (let i = 2; i < rows.length; i++) {
    const r = rows[i];
    if (!r[2]) continue;

    const wigCell = String(r[1]||'');
    if (wigCell.match(/WIG\s+(\d+)/i)) {
      wigNo = parseInt(wigCell.match(/WIG\s+(\d+)/i)[1]);
      lmNo  = 0;
    }

    const lmCell = String(r[2]||'');
    if (!lmCell.startsWith('LM')) continue;
    lmNo++;

    const kps = ['BALIKPAPAN','KALIMANTAN BARAT','KALIMANTAN SELATAN','KALIMANTAN TENGAH','KALIMANTAN TIMUR'];
    const pctPerKp = {};
    kps.forEach((kp,j) => { pctPerKp[kp] = parseFloat(r[3+j]) || 0; });

    results.push({
      wig:          wigNo,
      lm:           lmNo,
      pctSbu:       parseFloat(r[8])  || 0,
      totalSatuan:  parseFloat(r[9])  || 0,
      status:       String(r[10]||'—'),
      pctMinggu:    parseFloat(r[17]) || 0,
      statusMinggu: String(r[18]||'—'),
      pctPerKp,
    });
  }
  return results;
}

// ── BACA SHEET LM DETAIL ──────────────────────────────────
function readLmSheet(cfg) {
  const ws = SS_WIG.getSheetByName(cfg.sheet);
  if (!ws) throw new Error('Sheet "' + cfg.sheet + '" tidak ditemukan');

  const rows = ws.getDataRange().getValues();

  // Untuk sheet gabungan LM 6 & 7, cari blok yang sesuai
  let startRow = 0;
  if (cfg.part === 2) {
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]||'').match(/LM\s*7|Tower/i)) {
        startRow = i; break;
      }
    }
  }

  // Cari baris header KP
  let headerRow = -1;
  for (let i = startRow; i < rows.length; i++) {
    if (String(rows[i][0]||'').toUpperCase() === 'KP') {
      headerRow = i; break;
    }
  }
  if (headerRow < 0) throw new Error('Header KP tidak ditemukan');

  // Baca data KP
  const kpData = [];
  let totalTarget = 0, totalReal = 0, pencapaian = null, targetMingguBerjalan = 0;
  let totalFound = false, kontributorStart = -1;

  for (let i = headerRow + 1; i < rows.length; i++) {
    const r = rows[i];
    const kpName = String(r[0]||'').toUpperCase().trim();

    if (kpName === 'KONTRIBUTOR') { kontributorStart = i; break; }
    if (kpName === 'TOTAL') {
      totalTarget = parseFloat(r[2]) || 0;
      totalReal   = parseFloat(r[3]) || 0;
      targetMingguBerjalan = parseFloat(r[5]) || 0;
      totalFound = true;
      continue;
    }
    if (kpName === 'PENCAPAIAN') { pencapaian = parseFloat(r[3]) || 0; continue; }
    if (KP_LIST.includes(kpName)) {
      const targetMingguKP = parseFloat(r[8]) || 0;
      kpData.push({
        kp:             kpName,
        pct:            parseFloat(r[1]) || 0,
        target:         parseFloat(r[2]) || 0,
        real:           parseFloat(r[3]) || 0,
        targetMinggu:   Math.max(0, parseFloat(r[5]) || 0),
        selisih:        parseFloat(r[6]) || 0,
        targetMingguKP: targetMingguKP,
        realMinggu:     targetMingguKP,
      });
      continue;
    }
    // FALLBACK: baris TOTAL/PENCAPAIAN tanpa label di kolom A (mis. sheet JB)
    // — terjadi tepat setelah baris KP terakhir, kolom A kosong tapi C/D berisi angka
    if (!kpName && kpData.length > 0) {
      const hasTargetReal = r[2] !== '' && r[2] !== null && !totalFound;
      const hasOnlyPencapaian = (r[3] !== '' && r[3] !== null) && (r[2] === '' || r[2] === null) && pencapaian === null;
      if (hasTargetReal) {
        totalTarget = parseFloat(r[2]) || 0;
        totalReal   = parseFloat(r[3]) || 0;
        targetMingguBerjalan = parseFloat(r[5]) || 0;
        totalFound = true;
      } else if (hasOnlyPencapaian) {
        pencapaian = parseFloat(r[3]) || 0;
      }
    }
  }

  // Cari blok Kontributor (jika belum ketemu via break di atas — mis. sheet tanpa loop break)
  if (kontributorStart < 0) {
    for (let i = headerRow + 1; i < rows.length; i++) {
      if (String(rows[i][0]||'').toUpperCase().trim() === 'KONTRIBUTOR') { kontributorStart = i; break; }
    }
  }

  // Baca kontributor
  const kontributor = [];
  if (kontributorStart >= 0) {
    let inKontributor = false;
    for (let i = kontributorStart; i < rows.length; i++) {
      const r = rows[i];
      const cell = String(r[0]||'').toUpperCase().trim();
      if (cell === 'KONTRIBUTOR') { inKontributor = true; continue; }
      if (cell === 'NAMA') continue;
      if (!inKontributor) continue;
      const nama = String(r[0]||'').trim();
      // Berhenti kalau baris benar-benar kosong (bukan kontributor dengan jumlah 0)
      if (!nama) break;
      const jml  = parseFloat(r[1]) || 0;
      if (jml > 0) kontributor.push({ nama, jumlah: jml });
    }
  }
  kontributor.sort((a,b) => b.jumlah - a.jumlah);

  return { kpData, totalTarget, totalReal, pencapaian, targetMingguBerjalan, kontributor };
}

// ── TEST ──────────────────────────────────────────────────
function testSummary() {
  const r = getSummary();
  Logger.log('Total LM: ' + r.totalLm + ' | On: ' + r.onTrack + ' | Off: ' + r.offTrack);
  r.wigs.forEach(w => {
    Logger.log(w.nama);
    w.lms.forEach(l => Logger.log('  ' + l.nama + ' ' + l.status + ' ' + Math.round(l.pctSbu*100) + '%'));
  });
}

function testDetail() {
  const r = getLmDetail(1, 1);
  Logger.log('KP count: ' + r.kpData.length);
  Logger.log('Kontributor: ' + JSON.stringify(r.kontributor));
}

function testAllDetail() {
  const r = getAllDetail();
  Logger.log('Total LM: ' + r.totalLm + ' | On: ' + r.onTrack + ' | Off: ' + r.offTrack);
  Logger.log('Total Target: ' + r.totalTargetAll + ' | Total Real: ' + r.totalRealAll);
  r.lms.forEach(function(l){
    Logger.log(l.wigNama + ' / ' + l.nama + ' | target=' + l.totalTarget + ' real=' + l.totalReal + ' | kontributor=' + l.kontributor.length);
    if (l.kpData && l.kpData.length) {
      l.kpData.slice(0,2).forEach(function(kp){
        Logger.log('  KP=' + kp.kp + ' targetMinggu=' + kp.targetMinggu + ' targetMingguKP=' + kp.targetMingguKP);
      });
    }
  });
}

// DEBUG: cek raw rows sheet LM 5 JB di sekitar baris TOTAL
function testWIG4() {
  // Test 1: apakah sheet bisa dibuka langsung?
  try {
    var ws = SS_WIG.getSheetByName('WIG 4 LM 1 - SID BB');
    Logger.log('Sheet found: ' + (ws ? 'YES' : 'NULL'));
    if (ws) Logger.log('Sheet name: "' + ws.getName() + '"');
  } catch(e) {
    Logger.log('Direct access error: ' + e.message);
  }

  // Test 2: list semua sheet names untuk compare
  var sheets = SS_WIG.getSheets();
  Logger.log('All sheets (' + sheets.length + '):');
  sheets.forEach(function(s){ Logger.log('  "' + s.getName() + '"'); });
}
