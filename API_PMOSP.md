// ============================================================
// Aset & Inventory SBU Reg Kalimantan
// Apps Script — Monitoring PM OSP
//
// File  : RAW DATA PM SBU RKAL OSP & P3AK 2026
// Sheet : RAW DATA
// ============================================================

const SS_PMOSP = SpreadsheetApp.openById('18rnsE0JepoF2jlvWAdKTNuj5OK6NnMJKeLQ_gXN_dMo');
const SH_PMOSP = 'RAW DATA';

// Aging laporan: Tanggal Dok - Tanggal PM
const AGING_CEPAT  = 7;   // ≤7 hari  → Cepat
const AGING_NORMAL = 14;  // 8-14 hari → Normal
                           // >14 hari  → Lambat

// ── ENTRY POINT ───────────────────────────────────────────
function doGet(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    const data = getPMOSPData();
    out.setContent(JSON.stringify({ ok: true, ...data }));
  } catch(err) {
    out.setContent(JSON.stringify({ ok: false, error: err.message }));
  }
  return out;
}

// ── MAIN ──────────────────────────────────────────────────
function getPMOSPData() {
  const ws = SS_PMOSP.getSheetByName(SH_PMOSP);
  if(!ws) throw new Error('Sheet "' + SH_PMOSP + '" tidak ditemukan');

  const lastRow = ws.getLastRow();
  const lastCol = ws.getLastColumn();
  if(lastRow < 2) return { rows:[], totalWO:0 };

  const headers = ws.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h).trim());
  const raw     = ws.getRange(2,1,lastRow-1,lastCol).getDisplayValues();

  // Index kolom — case-insensitive
  const hUp = headers.map(h => h.toUpperCase());
  const fc  = (...names) => {
    for(const n of names) {
      const i = hUp.indexOf(n.toUpperCase());
      if(i >= 0) return i;
    }
    return -1;
  };

  const idx = {
    wilayah:    fc('Wilayah'),
    serpo:      fc('Serpo'),
    noWO:       fc('No. WO'),
    tglPM:      fc('Tanggal PM'),
    bulan:      fc('Bulan'),
    tahun:      fc('Tahun'),
    title:      fc('Title'),
    pop:        fc('POP'),
    tim:        fc('Tim'),
    pm:         fc('PM'),
    status:     fc('STATUS'),
    link:       fc('LINK'),
    petugas:    fc('Petugas'),
    linkDok:    fc('Link Dokumen'),
    picDok:     fc('PIC Dokumen'),
    tglDok:     fc('Tanggal Dok'),
    weekPM:     fc('Week PM'),
  };

  const parseTgl = s => {
    if(!s || String(s).trim() === '') return null;
    try {
      const str = String(s).trim();
      // Format M/D/YYYY
      if(str.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
        const [m,d,y] = str.split('/');
        const dt = new Date(parseInt(y), parseInt(m)-1, parseInt(d));
        dt.setHours(0,0,0,0);
        return isNaN(dt) ? null : dt;
      }
      const dt = new Date(str);
      if(!isNaN(dt)) { dt.setHours(0,0,0,0); return dt; }
      return null;
    } catch(e) { return null; }
  };

  const fmtDate = d => d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';

  // Normalisasi bulan — handle berbagai format
  const BULAN_MAP = {
    'januari':'Januari','january':'Januari','jan':'Januari',
    'februari':'Februari','february':'Februari','feb':'Februari',
    'maret':'Maret','march':'Maret','mar':'Maret',
    'april':'April','apr':'April',
    'mei':'Mei','may':'Mei',
    'juni':'Juni','june':'Juni','jun':'Juni',
    'juli':'Juli','july':'Juli','jul':'Juli',
    'agustus':'Agustus','august':'Agustus','aug':'Agustus','ags':'Agustus',
    'september':'September','sep':'September','sept':'September',
    'oktober':'Oktober','october':'Oktober','oct':'Oktober','okt':'Oktober',
    'november':'November','nov':'November',
    'desember':'Desember','december':'Desember','dec':'Desember','des':'Desember',
  };
  const normBulan = s => {
    const key = String(s||'').trim().toLowerCase().replace(/\s+/g,'');
    return BULAN_MAP[key] || (s ? String(s).trim() : '');
  };

  const rows = raw
    .filter(r => r[idx.noWO] !== '' && r[idx.noWO] !== null)
    .map(r => {
      const status   = String(r[idx.status] || '').trim().toUpperCase();
      const tglPM    = parseTgl(r[idx.tglPM]);
      const tglDok   = parseTgl(r[idx.tglDok]);

      // Aging laporan: Tanggal Dok - Tanggal PM
      let agingLaporan = null;
      let agingStatus  = '—';
      if(tglPM && tglDok) {
        agingLaporan = Math.floor((tglDok - tglPM) / (1000*60*60*24));
        if(agingLaporan <= AGING_CEPAT)       agingStatus = 'Cepat';
        else if(agingLaporan <= AGING_NORMAL)  agingStatus = 'Normal';
        else                                   agingStatus = 'Lambat';
      } else if(tglPM && !tglDok) {
        agingStatus = 'Belum Tersedia';
      }

      // Normalisasi wilayah — title case + fix typo
      const wilayahRaw = String(r[idx.wilayah] || '').trim();
      // Title case sederhana: setiap kata kapital di awal
      const toTitleCase = s => s.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
      const wilayah = toTitleCase(wilayahRaw);

      // Map ke label KP standar
      const kpMap = {
        'Kalimantan Timur':   'KALIMANTAN TIMUR',
        'Kalimantan Selatan': 'KALIMANTAN SELATAN',
        'Kalimantan Barat':   'KALIMANTAN BARAT',
        'Kalimantan Tengah':  'KALIMANTAN TENGAH',
        'Balikpapan':         'BALIKPAPAN',
      };
      const kp = kpMap[wilayah] || wilayah.toUpperCase();

      return {
        noWO:        String(r[idx.noWO]     || ''),
        wilayah,  // title case normalized
        namaKP:      kp,
        serpo:       String(r[idx.serpo]    || ''),
        tglPM:       fmtDate(tglPM),
        bulan:       normBulan(String(r[idx.bulan] || '')),
        tahun:       String(r[idx.tahun]    || ''),
        title:       String(r[idx.title]    || ''),
        pop:         String(r[idx.pop]      || ''),
        tim:         String(r[idx.tim]      || ''),
        pm:          String(r[idx.pm]       || ''),
        status,
        link:        String(r[idx.link]     || ''),
        petugas:     String(r[idx.petugas]  || ''),
        linkDok:     String(r[idx.linkDok]  || ''),
        picDok:      String(r[idx.picDok]   || ''),
        tglDok:      fmtDate(tglDok),
        weekPM:      String(r[idx.weekPM]   || ''),
        agingLaporan,
        agingStatus,
      };
    });

  // ── SUMMARY ───────────────────────────────────────────
  const totalWO   = rows.length;
  const close     = rows.filter(r => r.status === 'CLOSE').length;
  const open      = rows.filter(r => r.status === 'OPEN').length;
  const progres   = rows.filter(r => r.status === 'PROGRES' || r.status === 'PROGRESS').length;

  // Aging laporan
  const cepat         = rows.filter(r => r.agingStatus === 'Cepat').length;
  const normal        = rows.filter(r => r.agingStatus === 'Normal').length;
  const lambat        = rows.filter(r => r.agingStatus === 'Lambat').length;
  const belumTersedia = rows.filter(r => r.agingStatus === 'Belum Tersedia').length;

  // Chart per wilayah
  const byWilayah = {};
  rows.forEach(r => {
    const w = r.wilayah || '—';
    if(!byWilayah[w]) byWilayah[w] = { close:0, open:0, progres:0, total:0 };
    if(r.status === 'CLOSE')                              byWilayah[w].close++;
    else if(r.status === 'OPEN')                          byWilayah[w].open++;
    else if(r.status === 'PROGRES'||r.status==='PROGRESS') byWilayah[w].progres++;
    byWilayah[w].total++;
  });
  const chartWilayah = Object.entries(byWilayah)
    .map(([w,d]) => ({ wilayah:w, ...d }))
    .sort((a,b) => b.total - a.total);

  // Chart aging laporan per wilayah
  const byWilayahAging = {};
  rows.forEach(r => {
    const w = r.wilayah || '—';
    if(!byWilayahAging[w]) byWilayahAging[w] = { cepat:0, normal:0, lambat:0, belum:0 };
    if(r.agingStatus === 'Cepat')           byWilayahAging[w].cepat++;
    else if(r.agingStatus === 'Normal')     byWilayahAging[w].normal++;
    else if(r.agingStatus === 'Lambat')     byWilayahAging[w].lambat++;
    else if(r.agingStatus === 'Belum Tersedia') byWilayahAging[w].belum++;
  });

  const uniq = field => [...new Set(rows.map(r => r[field]).filter(Boolean))].sort();

  return {
    rows,
    totalWO, close, open, progres,
    cepat, normal, lambat, belumTersedia,
    chartWilayah,
    byWilayahAging,
    wilayahOptions: uniq('wilayah'),
    serpoOptions:   uniq('serpo'),
    bulanOptions:   uniq('bulan'),
    tahunOptions:   uniq('tahun'),
    statusOptions:  [...new Set(rows.map(r=>r.status).filter(Boolean))].sort(),
    agingOptions:   ['Cepat','Normal','Lambat','Belum Tersedia'],
    agingConfig:    { cepat: AGING_CEPAT, normal: AGING_NORMAL },
  };
}

// ── TEST ──────────────────────────────────────────────────
function testGetData() {
  const r = getPMOSPData();
  Logger.log('=== PM OSP TEST ===');
  Logger.log('Total WO         : ' + r.totalWO);
  Logger.log('Close            : ' + r.close);
  Logger.log('Open             : ' + r.open);
  Logger.log('Progres          : ' + r.progres);
  Logger.log('Aging Cepat      : ' + r.cepat);
  Logger.log('Aging Normal     : ' + r.normal);
  Logger.log('Aging Lambat     : ' + r.lambat);
  Logger.log('Belum Tersedia   : ' + r.belumTersedia);
  Logger.log('Wilayah options  : ' + JSON.stringify(r.wilayahOptions));
  Logger.log('Status options   : ' + JSON.stringify(r.statusOptions));
  Logger.log('Bulan options    : ' + JSON.stringify(r.bulanOptions));
  Logger.log('Chart (top3)     : ' + JSON.stringify(r.chartWilayah.slice(0,3)));
  Logger.log('Sample row 1     : ' + JSON.stringify(r.rows[0]));
  Logger.log('Sample row 2     : ' + JSON.stringify(r.rows[1]));
}
