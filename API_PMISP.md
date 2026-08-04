// ============================================================
// Aset & Inventory SBU Reg Kalimantan
// Apps Script — Monitoring PM ISP
//
// File  : REPORT PM ISP 2026 SBU KALIMANTAN
// Sheet : PM ISP 2026
// ============================================================

const SS_PMISP = SpreadsheetApp.openById('1Vtw6BwZ-bwRt3mKZHy7r95NfRODE1JSxlN47BE-mYoM');
const SH_PMISP = 'PM ISP 2026';

// ── ENTRY POINT ───────────────────────────────────────────
function doGet(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    const data = getPMISPData();
    out.setContent(JSON.stringify({ ok: true, ...data }));
  } catch(err) {
    out.setContent(JSON.stringify({ ok: false, error: err.message }));
  }
  return out;
}

// ── MAIN ──────────────────────────────────────────────────
function getPMISPData() {
  const ws = SS_PMISP.getSheetByName(SH_PMISP);
  if(!ws) throw new Error('Sheet "' + SH_PMISP + '" tidak ditemukan');

  const lastRow = ws.getLastRow();
  const lastCol = ws.getLastColumn();
  if(lastRow < 2) return { rows:[], totalPOP:0 };

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
    popId:        fc('pop_id'),
    popName:      fc('pop_name'),
    popType:      fc('pop_type'),
    modelName:    fc('model_name'),
    namaKP:       fc('namaKP'),
    city:         fc('city'),
    provinsi:     fc('provinsi'),
    lokasiPLN:    fc('Lokasi POP.Lokasi PLN'),
    katLokasi:    fc('r10_kategori_lokasi'),
    tglPlan:      fc('Tgl Plan'),
    tglReal:      fc('Tgl Real'),
    statusPM:     fc('Status PM'),
    timPelaksana: fc('Tim Pelaksana'),
    pmKe:         fc('PM ke -'),
    picLaporan:   fc('PIC Laporan'),
    wo:           fc('WO'),
    linkLaporan:  fc('Link Laporan'),
    statusLaporan:fc('Status Laporan'),
    week:         fc('Week'),
    ujiBatere:    fc('Uji Batere'),
  };

  // Parse tanggal dari display value (d-MMM-yy atau string lain)
  const parseTgl = s => {
    if(!s || String(s).trim() === '') return null;
    try {
      const str = String(s).trim();
      // Format: 5-Jan-26 atau 7-Jan-2026
      const m = str.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$/);
      if(m) {
        const months = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
        const yr = parseInt(m[3]) + (parseInt(m[3]) < 100 ? 2000 : 0);
        const d  = new Date(yr, months[m[2].toLowerCase()], parseInt(m[1]));
        d.setHours(0,0,0,0);
        return isNaN(d) ? null : d;
      }
      // Format lain
      const d = new Date(str);
      if(!isNaN(d)) { d.setHours(0,0,0,0); return d; }
      return null;
    } catch(e) { return null; }
  };

  const fmtDate = d => d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';

  const rows = raw
    .filter(r => r[idx.popId] !== '' && r[idx.popId] !== null)
    .map(r => {
      const statusPM      = String(r[idx.statusPM]      || '').trim();
      const statusLaporan = String(r[idx.statusLaporan]  || '').trim();
      const tglPlan       = parseTgl(r[idx.tglPlan]);
      const tglReal       = parseTgl(r[idx.tglReal]);

      // Ketepatan waktu — hanya untuk Sudah Terlaksana dengan Tgl Real ada
      let ketepatanWaktu = '—';
      let selisihHari    = null;
      if(statusPM === 'Sudah Terlaksana' && tglPlan && tglReal) {
        selisihHari = Math.floor((tglReal - tglPlan) / (1000*60*60*24));
        if(selisihHari < 0)      ketepatanWaktu = 'Maju dari Jadwal';
        else if(selisihHari === 0) ketepatanWaktu = 'Tepat Waktu';
        else                      ketepatanWaktu = 'Lewat dari Jadwal';
      }

      // Sisa hari ke Tgl Plan (untuk yang Belum Terlaksana)
      let sisaHari = null;
      if(statusPM !== 'Sudah Terlaksana' && tglPlan) {
        const today = new Date(); today.setHours(0,0,0,0);
        sisaHari = Math.floor((tglPlan - today) / (1000*60*60*24));
      }

      return {
        popId:        String(r[idx.popId]         || ''),
        popName:      String(r[idx.popName]        || ''),
        popType:      String(r[idx.popType]        || ''),
        modelName:    String(r[idx.modelName]      || ''),
        namaKP:       String(r[idx.namaKP]         || ''),
        city:         String(r[idx.city]           || ''),
        provinsi:     String(r[idx.provinsi]       || ''),
        lokasiPLN:    String(r[idx.lokasiPLN]      || ''),
        katLokasi:    String(r[idx.katLokasi]      || ''),
        tglPlan:      fmtDate(tglPlan),
        tglReal:      fmtDate(tglReal),
        statusPM,
        statusLaporan,
        timPelaksana: String(r[idx.timPelaksana]   || ''),
        pmKe:         String(r[idx.pmKe]           || ''),
        picLaporan:   String(r[idx.picLaporan]     || ''),
        wo:           String(r[idx.wo]             || ''),
        linkLaporan:  String(r[idx.linkLaporan]    || ''),
        week:         String(r[idx.week]           || ''),
        ketepatanWaktu,
        selisihHari,
        sisaHari,
      };
    });

  // ── SUMMARY ───────────────────────────────────────────
  const totalPOP     = rows.length;
  const sudahPM      = rows.filter(r => r.statusPM === 'Sudah Terlaksana').length;
  const belumPM      = rows.filter(r => r.statusPM !== 'Sudah Terlaksana').length;
  const laporanReady = rows.filter(r => r.statusLaporan === 'Sudah Tersedia').length;
  const laporanBelum = rows.filter(r => r.statusLaporan !== 'Sudah Tersedia' && r.statusLaporan !== '').length;

  // Ketepatan waktu — dari yang Sudah Terlaksana
  const sudahRows   = rows.filter(r => r.statusPM === 'Sudah Terlaksana');
  const maju        = sudahRows.filter(r => r.ketepatanWaktu === 'Maju dari Jadwal').length;
  const tepat       = sudahRows.filter(r => r.ketepatanWaktu === 'Tepat Waktu').length;
  const lewat       = sudahRows.filter(r => r.ketepatanWaktu === 'Lewat dari Jadwal').length;

  // Belum terlaksana — breakdown sisa hari ke Tgl Plan
  const belumRows   = rows.filter(r => r.statusPM !== 'Sudah Terlaksana');
  const belumMendatang = belumRows.filter(r => r.sisaHari !== null && r.sisaHari >= 0).length;
  const belumLewat     = belumRows.filter(r => r.sisaHari !== null && r.sisaHari < 0).length;
  const belumTanpaPlan = belumRows.filter(r => r.sisaHari === null).length;

  // Chart per KP
  const byKP = {};
  rows.forEach(r => {
    const kp = r.namaKP || '—';
    if(!byKP[kp]) byKP[kp] = { sudah:0, belum:0, total:0 };
    if(r.statusPM === 'Sudah Terlaksana') byKP[kp].sudah++;
    else byKP[kp].belum++;
    byKP[kp].total++;
  });
  const chartKP = Object.entries(byKP)
    .map(([kp,d]) => ({ kp, ...d }))
    .sort((a,b) => b.total - a.total);

  // Chart ketepatan waktu per KP (dari sudah terlaksana)
  const byKPKetepatan = {};
  sudahRows.forEach(r => {
    const kp = r.namaKP || '—';
    if(!byKPKetepatan[kp]) byKPKetepatan[kp] = { maju:0, tepat:0, lewat:0 };
    if(r.ketepatanWaktu === 'Maju dari Jadwal')   byKPKetepatan[kp].maju++;
    else if(r.ketepatanWaktu === 'Tepat Waktu')    byKPKetepatan[kp].tepat++;
    else if(r.ketepatanWaktu === 'Lewat dari Jadwal') byKPKetepatan[kp].lewat++;
  });

  // Filter options
  const uniq = field => [...new Set(rows.map(r => r[field]).filter(Boolean))].sort();

  return {
    rows,
    totalPOP, sudahPM, belumPM,
    laporanReady, laporanBelum,
    maju, tepat, lewat,
    belumMendatang, belumLewat, belumTanpaPlan,
    chartKP,
    byKPKetepatan,
    kpOptions:     uniq('namaKP'),
    pmKeOptions:   uniq('pmKe'),
    weekOptions:   uniq('week'),
    statusPMOpts:  [...new Set(rows.map(r => r.statusPM).filter(Boolean))],
    statusLapOpts: [...new Set(rows.map(r => r.statusLaporan).filter(Boolean))],
  };
}

// ── TEST ──────────────────────────────────────────────────
function testGetData() {
  const r = getPMISPData();
  Logger.log('=== PM ISP TEST ===');
  Logger.log('Total POP        : ' + r.totalPOP);
  Logger.log('Sudah PM         : ' + r.sudahPM);
  Logger.log('Belum PM         : ' + r.belumPM);
  Logger.log('Laporan Ready    : ' + r.laporanReady);
  Logger.log('Laporan Belum    : ' + r.laporanBelum);
  Logger.log('Maju dari Jadwal : ' + r.maju);
  Logger.log('Tepat Waktu      : ' + r.tepat);
  Logger.log('Lewat dari Jadwal: ' + r.lewat);
  Logger.log('Belum-Mendatang  : ' + r.belumMendatang);
  Logger.log('Belum-Lewat Plan : ' + r.belumLewat);
  Logger.log('Belum-Tanpa Plan : ' + r.belumTanpaPlan);
  Logger.log('KP options       : ' + JSON.stringify(r.kpOptions));
  Logger.log('PM ke options    : ' + JSON.stringify(r.pmKeOptions));
  Logger.log('Status PM        : ' + JSON.stringify(r.statusPMOpts));
  Logger.log('Status Laporan   : ' + JSON.stringify(r.statusLapOpts));
  Logger.log('Chart KP (top3)  : ' + JSON.stringify(r.chartKP.slice(0,3)));
  Logger.log('Sample row 1     : ' + JSON.stringify(r.rows[0]));
  Logger.log('Sample row 2     : ' + JSON.stringify(r.rows[1]));
}
