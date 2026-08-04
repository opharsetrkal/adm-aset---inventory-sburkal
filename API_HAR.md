// ============================================================
// Opharset SBU Reg Kalimantan
// Apps Script — Monitoring HAR (SPR & SPP)
//
// File  : (spreadsheet HAR)
// Sheet : SPR, SPP
//
// CARA DEPLOY:
//   1. Buka GSheet HAR -> Extensions -> Apps Script
//   2. Paste code ini -> Save
//   3. Jalankan testGetData() -> cek log
//   4. Deploy -> Manage deployments -> Edit -> New version -> Deploy
//   5. URL tetap sama, tidak perlu update API_HAR di index.html
// ============================================================

function doGet(e) {
  var out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    var params = e ? e.parameter : {};
    var tab = (params.tab || 'spr').toLowerCase();
    var data = (tab === 'spp') ? getSPPData() : getSPRData();
    out.setContent(JSON.stringify(Object.assign({ ok: true, tab: tab }, data)));
  } catch (err) {
    out.setContent(JSON.stringify({ ok: false, error: err.message }));
  }
  return out;
}

// ══════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════
function findCol(headers) {
  var names = Array.prototype.slice.call(arguments, 1);
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim().toLowerCase();
    for (var j = 0; j < names.length; j++) {
      if (h === names[j].toLowerCase()) return i;
    }
  }
  return -1;
}

function parseTgl(val) {
  if (!val) return '';
  var s = String(val).trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    var y = d.getFullYear();
    var m = String(d.getMonth() + 1).padStart(2, '0');
    var day = String(d.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + day;
  }
  return s;
}

// Konversi string "YYYY-MM-DD" ke epoch ms murni (tanpa lewat Date() lokal)
// supaya perhitungan selisih hari tidak bergeser oleh timezone.
function dateStrToMs(s) {
  if (!s) return null;
  var m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  return Date.UTC(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
}

function todayUTCms() {
  var now = new Date();
  return Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
}

function parseNominal(val) {
  if (!val) return 0;
  var s = String(val).trim().replace(/[Rp\s]/g, '');
  if (!s) return 0;
  var lastDot   = s.lastIndexOf('.');
  var lastComma = s.lastIndexOf(',');
  if (lastComma > lastDot) {
    var afterComma  = s.slice(lastComma + 1);
    var beforeComma = s.slice(0, lastComma).replace(/[.,]/g, '');
    if (afterComma.length === 3 && /^\d+$/.test(afterComma) && /^\d+$/.test(beforeComma)) {
      s = s.replace(/,/g, ''); // EN ribuan: 29,979,400
    } else {
      s = s.replace(/\./g, '').replace(',', '.'); // ID desimal: 1.234.567,89
    }
  } else if (lastDot > lastComma) {
    var parts = s.split('.');
    var afterLast  = s.slice(lastDot + 1);
    var beforeLast = s.slice(0, lastDot).replace(/[.,]/g, '');
    if (parts.length > 2) {
      s = s.replace(/\./g, ''); // ID ribuan multi-titik: 2.702.306
    } else if (afterLast.length === 3 && beforeLast.length <= 4) {
      s = s.replace(/\./g, ''); // ID ribuan titik tunggal: 1.234
    }
    // else titik tunggal = desimal, biarkan: 24333421.623
  } else {
    s = s.replace(/[.,]/g, '');
  }
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

// Ambil angka pertama dari string durasi, mis. "30 hari" / "30 Hari Kalender" / "30" -> 30
function parseDurasi(val) {
  if (!val) return null;
  var m = String(val).match(/(\d+([.,]\d+)?)/);
  if (!m) return null;
  var n = parseFloat(m[1].replace(',', '.'));
  return isNaN(n) ? null : n;
}

// ══════════════════════════════════════════════════════════
// SPR DATA
// ══════════════════════════════════════════════════════════
function getSPRData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName('SPR');
  if (!ws) throw new Error('Sheet SPR tidak ditemukan');
  var data = ws.getDataRange().getDisplayValues();
  if (data.length < 2) return emptySPR();

  var hdr = data[0];
  var ci = {
    no:             findCol(hdr, 'NO'),
    kodeAnggaran:   findCol(hdr, 'KODE ANGGARAN'),
    noSPR:          findCol(hdr, 'NO SPR'),
    kegiatan:       findCol(hdr, 'KEGIATAN'),
    nominal:        findCol(hdr, 'NOMINAL'),
    mitra:          findCol(hdr, 'MITRA'),
    statusPengadaan:findCol(hdr, 'STATUS PENGADAAN'),
    statusPenagihan:findCol(hdr, 'STATUS PENAGIHAN'),
    keterangan:     findCol(hdr, 'KETERANGAN'),
    peruntukan:     findCol(hdr, 'PERUNTUKAN'),
    tglRilisSPR:    findCol(hdr, 'TGL RILIS SPR'),
    tglBapsBast:    findCol(hdr, 'TGL BAPS/BAST', 'TGL BAPS / BAST')
  };

  var rows = [];
  var byStatusPengadaan = {};
  var byStatusPenagihan = {};
  var byMitra = {};
  var byKodeAnggaran = {};
  var byPeruntukan = {};
  var totalNominal = 0;

  var outstandingCount = 0, outstandingValue = 0;
  var doneValue = 0, cancelValue = 0;
  var aging45Count = 0;
  var agingBuckets = { b0_30: 0, b31_45: 0, bOver45: 0 };
  var monthlyMap = {}; // key 'YYYY-MM' -> {count, nominal}
  var todayMs = todayUTCms();

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var noSPR = ci.noSPR >= 0 ? String(r[ci.noSPR]).trim() : '';
    var kegiatan = ci.kegiatan >= 0 ? String(r[ci.kegiatan]).trim() : '';
    if (!noSPR && !kegiatan) continue;

    var nominal = ci.nominal >= 0 ? parseNominal(r[ci.nominal]) : 0;
    var statusPengadaan = ci.statusPengadaan >= 0 ? String(r[ci.statusPengadaan]).trim() : '';
    var statusPenagihan = ci.statusPenagihan >= 0 ? String(r[ci.statusPenagihan]).trim() : '';
    var mitra = ci.mitra >= 0 ? String(r[ci.mitra]).trim() : '';
    var kodeAnggaran = ci.kodeAnggaran >= 0 ? String(r[ci.kodeAnggaran]).trim() : '';
    var peruntukan = ci.peruntukan >= 0 ? String(r[ci.peruntukan]).trim() : '';
    var tglRilisSPR = ci.tglRilisSPR >= 0 ? parseTgl(r[ci.tglRilisSPR]) : '';
    var tglBapsBast = ci.tglBapsBast >= 0 ? parseTgl(r[ci.tglBapsBast]) : '';

    // ── AGING: TGL RILIS SPR -> (TGL BAPS/BAST jika ada, else hari ini) ──
    var aging = null;
    if (tglRilisSPR) {
      var startMs = dateStrToMs(tglRilisSPR);
      if (startMs !== null) {
        var endMs = tglBapsBast ? dateStrToMs(tglBapsBast) : todayMs;
        if (endMs === null) endMs = todayMs;
        aging = Math.round((endMs - startMs) / 86400000);
        if (aging < 0) aging = 0;
      }
    }

    var spUp = statusPengadaan.toUpperCase();
    var stUp = statusPenagihan.toUpperCase();
    var isDone   = spUp === 'DONE';
    var isCancel = spUp === 'CANCEL';
    var isReject = spUp === 'REJECT';

    // ── OUTSTANDING: Pengadaan Done AND Penagihan bukan Cancel/Submit Invoice ke Pusat ──
    var isOutstanding = isDone &&
      stUp !== 'CANCEL' &&
      stUp.indexOf('SUBMIT INVOICE KE PUSAT') === -1;
    if (isOutstanding) { outstandingCount++; outstandingValue += nominal; }

    // ── FINANCIAL EXPOSURE (Health card kanan) ──
    if (isDone) doneValue += nominal;
    if (isCancel) cancelValue += nominal;

    // ── AGING BUCKET (exclude Done/Cancel/Reject, sejalan Priority Action) ──
    var includeInAging = !isDone && !isCancel && !isReject;
    if (includeInAging && aging !== null) {
      if (aging > 45) { agingBuckets.bOver45++; aging45Count++; }
      else if (aging >= 31) agingBuckets.b31_45++;
      else agingBuckets.b0_30++;
    }

    // ── MONTHLY TREND (by TGL RILIS SPR) ──
    if (tglRilisSPR) {
      var mKey = tglRilisSPR.slice(0, 7);
      if (!monthlyMap[mKey]) monthlyMap[mKey] = { count: 0, nominal: 0 };
      monthlyMap[mKey].count++;
      monthlyMap[mKey].nominal += nominal;
    }

    totalNominal += nominal;
    if (statusPengadaan) byStatusPengadaan[statusPengadaan] = (byStatusPengadaan[statusPengadaan] || 0) + 1;
    if (statusPenagihan) byStatusPenagihan[statusPenagihan] = (byStatusPenagihan[statusPenagihan] || 0) + 1;
    if (mitra) byMitra[mitra] = (byMitra[mitra] || 0) + nominal;
    if (kodeAnggaran) byKodeAnggaran[kodeAnggaran] = (byKodeAnggaran[kodeAnggaran] || 0) + nominal;
    if (peruntukan) byPeruntukan[peruntukan] = (byPeruntukan[peruntukan] || 0) + 1;

    rows.push({
      no: ci.no >= 0 ? String(r[ci.no]).trim() : String(i),
      kodeAnggaran: kodeAnggaran,
      noSPR: noSPR,
      kegiatan: kegiatan,
      nominal: nominal,
      mitra: mitra,
      statusPengadaan: statusPengadaan,
      statusPenagihan: statusPenagihan,
      keterangan: ci.keterangan >= 0 ? String(r[ci.keterangan]).trim() : '',
      peruntukan: peruntukan,
      tglRilisSPR: tglRilisSPR,
      tglBapsBast: tglBapsBast,
      aging: aging,
      isOutstanding: isOutstanding
    });
  }

  // ── PRIORITY ACTION: exclude Done/Cancel/Reject, sort aging DESC lalu nominal DESC, top 10 ──
  var priorityAction = rows
    .filter(function(r) {
      var sp = (r.statusPengadaan || '').toUpperCase();
      return sp !== 'DONE' && sp !== 'CANCEL' && sp !== 'REJECT' && r.aging !== null;
    })
    .sort(function(a, b) {
      if ((b.aging || 0) !== (a.aging || 0)) return (b.aging || 0) - (a.aging || 0);
      return (b.nominal || 0) - (a.nominal || 0);
    })
    .slice(0, 10)
    .map(function(r) {
      return {
        noSPR: r.noSPR, kegiatan: r.kegiatan, aging: r.aging,
        nominal: r.nominal, statusPenagihan: r.statusPenagihan,
        statusPengadaan: r.statusPengadaan, mitra: r.mitra
      };
    });

  var monthlyTrend = Object.keys(monthlyMap).sort().map(function(k) {
    return { month: k, count: monthlyMap[k].count, nominal: monthlyMap[k].nominal };
  });

  var mitraSet = {}; rows.forEach(function(r){ if(r.mitra) mitraSet[r.mitra]=1; });
  var kodeAnggaranSet = {}; rows.forEach(function(r){ if(r.kodeAnggaran) kodeAnggaranSet[r.kodeAnggaran]=1; });
  var peruntukanSet = {}; rows.forEach(function(r){ if(r.peruntukan) peruntukanSet[r.peruntukan]=1; });

  return {
    rows: rows,
    totalSPR: rows.length,
    totalNominal: totalNominal,
    byStatusPengadaan: byStatusPengadaan,
    byStatusPenagihan: byStatusPenagihan,
    byMitra: byMitra,
    byKodeAnggaran: byKodeAnggaran,
    byPeruntukan: byPeruntukan,
    statusPengadaanOpts: Object.keys(byStatusPengadaan).sort(),
    statusPenagihanOpts: Object.keys(byStatusPenagihan).sort(),
    mitraOpts: Object.keys(mitraSet).sort(),
    kodeAnggaranOpts: Object.keys(kodeAnggaranSet).sort(),
    peruntukanOpts: Object.keys(peruntukanSet).sort(),
    outstandingCount: outstandingCount,
    outstandingValue: outstandingValue,
    doneValue: doneValue,
    cancelValue: cancelValue,
    aging45Count: aging45Count,
    agingBuckets: agingBuckets,
    priorityAction: priorityAction,
    monthlyTrend: monthlyTrend
  };
}

function emptySPR() {
  return {
    rows: [], totalSPR: 0, totalNominal: 0,
    byStatusPengadaan: {}, byStatusPenagihan: {}, byMitra: {}, byKodeAnggaran: {}, byPeruntukan: {},
    statusPengadaanOpts: [], statusPenagihanOpts: [], mitraOpts: [], kodeAnggaranOpts: [], peruntukanOpts: [],
    outstandingCount: 0, outstandingValue: 0, doneValue: 0, cancelValue: 0, aging45Count: 0,
    agingBuckets: { b0_30: 0, b31_45: 0, bOver45: 0 },
    priorityAction: [], monthlyTrend: []
  };
}

// ══════════════════════════════════════════════════════════
// SPP DATA
// ══════════════════════════════════════════════════════════
function getSPPData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName('SPP');
  if (!ws) throw new Error('Sheet SPP tidak ditemukan');
  var data = ws.getDataRange().getDisplayValues();
  if (data.length < 2) return emptySPP();

  var hdr = data[0];
  var ci = {
    no:               findCol(hdr, 'NO'),
    noSPP:            findCol(hdr, 'NO SPP'),
    judulSPP:         findCol(hdr, 'JUDUL SPP'),
    kodeAnggaran:     findCol(hdr, 'KODE ANGGARAN'),
    coa:              findCol(hdr, 'COA'),
    pr:               findCol(hdr, 'PR'),
    tipeSPP:          findCol(hdr, 'TIPE SPP'),
    noSPPReff:        findCol(hdr, 'NO SPP REFF'),
    noPOSP2K:         findCol(hdr, 'NO PO/SP2K', 'NO PO / SP2K'),
    metodePengadaan:  findCol(hdr, 'METODE PENGADAAN'),
    jenisPengadaan:   findCol(hdr, 'JENIS PENGADAAN'),
    status:           findCol(hdr, 'STATUS'),
    nilaiPekerjaan:   findCol(hdr, 'NILAI PEKERJAAN'),
    updateTerbaru:    findCol(hdr, 'UPDATE TERBARU'),
    kategoriPekerjaan:findCol(hdr, 'KATEGORI PEKERJAAN'),
    tglRilisSP2K:     findCol(hdr, 'TGL RILIS SP2K'),
    durasiPekerjaan:  findCol(hdr, 'DURASI PEKERJAAN'),
    tglBapsBast:      findCol(hdr, 'TGL BAPS/BAST', 'TGL BAPS / BAST'),
    perpanjangan:     findCol(hdr, 'PERPANJANGAN')
  };

  var rows = [];
  var byStatus = {};
  var byUpdateTerbaru = {};
  var byTipeSPP = {};
  var byMetodePengadaan = {};
  var byJenisPengadaan = {};
  var byKategoriPekerjaan = {};
  var totalNilai = 0;

  var aktifCount = 0;            // = outstanding (TGL BAPS/BAST kosong)
  var outstandingValue = 0;
  var doneValue = 0, cancelValue = 0;
  var criticalCount = 0;
  var agingBuckets = { b0_30: 0, b31_45: 0, bOver45: 0 };
  var monthlyMap = {};
  var todayMs = todayUTCms();

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var noSPP = ci.noSPP >= 0 ? String(r[ci.noSPP]).trim() : '';
    var judulSPP = ci.judulSPP >= 0 ? String(r[ci.judulSPP]).trim() : '';
    if (!noSPP && !judulSPP) continue;

    var nilai = ci.nilaiPekerjaan >= 0 ? parseNominal(r[ci.nilaiPekerjaan]) : 0;
    var status = ci.status >= 0 ? String(r[ci.status]).trim() : '';
    var updateTerbaru = ci.updateTerbaru >= 0 ? String(r[ci.updateTerbaru]).trim() : '';
    var tipeSPP = ci.tipeSPP >= 0 ? String(r[ci.tipeSPP]).trim() : '';
    var metodePengadaan = ci.metodePengadaan >= 0 ? String(r[ci.metodePengadaan]).trim() : '';
    var jenisPengadaan = ci.jenisPengadaan >= 0 ? String(r[ci.jenisPengadaan]).trim() : '';
    var kategoriPekerjaan = ci.kategoriPekerjaan >= 0 ? String(r[ci.kategoriPekerjaan]).trim() : '';
    var tglRilisSP2K = ci.tglRilisSP2K >= 0 ? parseTgl(r[ci.tglRilisSP2K]) : '';
    var tglBapsBast = ci.tglBapsBast >= 0 ? parseTgl(r[ci.tglBapsBast]) : '';
    var durasiRaw = ci.durasiPekerjaan >= 0 ? String(r[ci.durasiPekerjaan]).trim() : '';
    var durasiRencana = parseDurasi(durasiRaw);
    var perpanjanganRaw = ci.perpanjangan >= 0 ? String(r[ci.perpanjangan]).trim() : '';
    var isPerpanjangan = perpanjanganRaw.toUpperCase() === 'YA';

    var statusUp = status.toUpperCase();
    var isCancel = statusUp === 'CANCEL';
    var isDone = updateTerbaru.toUpperCase().indexOf('SELESAI') !== -1;

    // ── AKTIF & OUTSTANDING: TGL BAPS/BAST kosong = belum selesai ──
    var isAktif = !tglBapsBast;
    if (isAktif) { aktifCount++; outstandingValue += nilai; }
    if (isDone) doneValue += nilai;
    if (isCancel) cancelValue += nilai;

    // ── AGING vs DURASI RENCANA ──
    var aging = null, isCritical = false;
    if (tglRilisSP2K) {
      var startMs = dateStrToMs(tglRilisSP2K);
      if (startMs !== null) {
        var endMs = tglBapsBast ? dateStrToMs(tglBapsBast) : todayMs;
        if (endMs === null) endMs = todayMs;
        aging = Math.round((endMs - startMs) / 86400000);
        if (aging < 0) aging = 0;
        if (durasiRencana !== null && aging > durasiRencana) isCritical = true;
      }
    }

    var includeInAging = !isCancel && isAktif; // exclude cancel & yang sudah selesai
    if (includeInAging && aging !== null) {
      if (isCritical) { agingBuckets.bOver45++; criticalCount++; }
      else if (durasiRencana !== null && aging >= durasiRencana * 0.7) agingBuckets.b31_45++;
      else agingBuckets.b0_30++;
    }

    // ── SISA HARI KONTRAK (untuk Perpanjangan = Ya): TGL_AKHIR_KONTRAK - hari ini ──
    // TGL_AKHIR_KONTRAK = TGL RILIS SP2K + DURASI PEKERJAAN
    var sisaHariKontrak = null;
    if (isPerpanjangan && tglRilisSP2K && durasiRencana !== null) {
      var startMsK = dateStrToMs(tglRilisSP2K);
      if (startMsK !== null) {
        var endKontrakMs = startMsK + (durasiRencana * 86400000);
        sisaHariKontrak = Math.round((endKontrakMs - todayMs) / 86400000);
      }
    }

    // ── MONTHLY TREND (by TGL RILIS SP2K) ──
    if (tglRilisSP2K) {
      var mKey = tglRilisSP2K.slice(0, 7);
      if (!monthlyMap[mKey]) monthlyMap[mKey] = { count: 0, nominal: 0 };
      monthlyMap[mKey].count++;
      monthlyMap[mKey].nominal += nilai;
    }

    totalNilai += nilai;
    if (status) byStatus[status] = (byStatus[status] || 0) + 1;
    if (updateTerbaru) byUpdateTerbaru[updateTerbaru] = (byUpdateTerbaru[updateTerbaru] || 0) + 1;
    if (tipeSPP) byTipeSPP[tipeSPP] = (byTipeSPP[tipeSPP] || 0) + 1;
    if (metodePengadaan) byMetodePengadaan[metodePengadaan] = (byMetodePengadaan[metodePengadaan] || 0) + 1;
    if (jenisPengadaan) byJenisPengadaan[jenisPengadaan] = (byJenisPengadaan[jenisPengadaan] || 0) + 1;
    if (kategoriPekerjaan) byKategoriPekerjaan[kategoriPekerjaan] = (byKategoriPekerjaan[kategoriPekerjaan] || 0) + 1;

    rows.push({
      no: ci.no >= 0 ? String(r[ci.no]).trim() : String(i),
      noSPP: noSPP,
      judulSPP: judulSPP,
      kodeAnggaran: ci.kodeAnggaran >= 0 ? String(r[ci.kodeAnggaran]).trim() : '',
      coa: ci.coa >= 0 ? String(r[ci.coa]).trim() : '',
      pr: ci.pr >= 0 ? String(r[ci.pr]).trim() : '',
      tipeSPP: tipeSPP,
      noSPPReff: ci.noSPPReff >= 0 ? String(r[ci.noSPPReff]).trim() : '',
      noPOSP2K: ci.noPOSP2K >= 0 ? String(r[ci.noPOSP2K]).trim() : '',
      metodePengadaan: metodePengadaan,
      jenisPengadaan: jenisPengadaan,
      status: status,
      nilaiPekerjaan: nilai,
      updateTerbaru: updateTerbaru,
      kategoriPekerjaan: kategoriPekerjaan,
      tglRilisSP2K: tglRilisSP2K,
      durasiPekerjaan: durasiRaw,
      durasiRencana: durasiRencana,
      tglBapsBast: tglBapsBast,
      aging: aging,
      isAktif: isAktif,
      isCritical: isCritical,
      isPerpanjangan: isPerpanjangan,
      sisaHariKontrak: sisaHariKontrak
    });
  }

  // ── PRIORITY ACTION (2 kondisi digabung, top 10 berdasarkan urgensi) ──
  // Kondisi A — Perpanjangan = Ya: urgensi renewal kontrak (sisaHariKontrak kecil/negatif)
  // Kondisi B — Perpanjangan = Tidak: urgensi percepatan penyelesaian (aging > durasiRencana)
  var priorityA = rows
    .filter(function(r) {
      var st = (r.status || '').toUpperCase();
      return st !== 'CANCEL' && r.isAktif && r.isPerpanjangan && r.sisaHariKontrak !== null;
    })
    .map(function(r) {
      var kondisi, urgencyScore;
      if (r.sisaHariKontrak < 0) { kondisi = 'Melewati Masa Kontrak'; urgencyScore = 10000 + Math.abs(r.sisaHariKontrak); }
      else if (r.sisaHariKontrak <= 30) { kondisi = 'Kritis'; urgencyScore = 5000 + (30 - r.sisaHariKontrak); }
      else if (r.sisaHariKontrak <= 60) { kondisi = 'Perhatian'; urgencyScore = 1000 + (60 - r.sisaHariKontrak); }
      else return null;
      return {
        noSPP: r.noSPP, judulSPP: r.judulSPP, kondisi: kondisi,
        keterangan: 'Perpanjangan', sisaHari: r.sisaHariKontrak,
        nilaiPekerjaan: r.nilaiPekerjaan, updateTerbaru: r.updateTerbaru,
        urgencyScore: urgencyScore
      };
    })
    .filter(function(x) { return x !== null; });

  var priorityB = rows
    .filter(function(r) {
      var st = (r.status || '').toUpperCase();
      return st !== 'CANCEL' && r.isAktif && !r.isPerpanjangan && r.aging !== null &&
        r.tglRilisSP2K && r.durasiRencana !== null && r.aging >= r.durasiRencana * 0.7;
    })
    .map(function(r) {
      var lewat = r.aging - r.durasiRencana;
      var kondisi = lewat > 0 ? 'Melewati Durasi Pekerjaan' : 'Mendekati Batas Durasi';
      return {
        noSPP: r.noSPP, judulSPP: r.judulSPP, kondisi: kondisi,
        keterangan: 'Percepatan Penyelesaian', sisaHari: r.durasiRencana - r.aging,
        nilaiPekerjaan: r.nilaiPekerjaan, updateTerbaru: r.updateTerbaru,
        urgencyScore: 3000 + lewat
      };
    });

  var priorityAction = priorityA.concat(priorityB)
    .sort(function(a, b) { return b.urgencyScore - a.urgencyScore; })
    .slice(0, 10);

  var monthlyTrend = Object.keys(monthlyMap).sort().map(function(k) {
    return { month: k, count: monthlyMap[k].count, nominal: monthlyMap[k].nominal };
  });

  return {
    rows: rows,
    totalSPP: rows.length,
    totalNilai: totalNilai,
    byStatus: byStatus,
    byUpdateTerbaru: byUpdateTerbaru,
    byTipeSPP: byTipeSPP,
    byMetodePengadaan: byMetodePengadaan,
    byJenisPengadaan: byJenisPengadaan,
    byKategoriPekerjaan: byKategoriPekerjaan,
    statusOpts: Object.keys(byStatus).sort(),
    updateTerbaruOpts: Object.keys(byUpdateTerbaru).sort(),
    tipeSPPOpts: Object.keys(byTipeSPP).sort(),
    metodePengadaanOpts: Object.keys(byMetodePengadaan).sort(),
    jenisPengadaanOpts: Object.keys(byJenisPengadaan).sort(),
    kategoriPekerjaanOpts: Object.keys(byKategoriPekerjaan).sort(),
    aktifCount: aktifCount,
    outstandingValue: outstandingValue,
    doneValue: doneValue,
    cancelValue: cancelValue,
    criticalCount: criticalCount,
    agingBuckets: agingBuckets,
    priorityAction: priorityAction,
    monthlyTrend: monthlyTrend
  };
}

function emptySPP() {
  return {
    rows: [], totalSPP: 0, totalNilai: 0,
    byStatus: {}, byUpdateTerbaru: {}, byTipeSPP: {}, byMetodePengadaan: {}, byJenisPengadaan: {}, byKategoriPekerjaan: {},
    statusOpts: [], updateTerbaruOpts: [], tipeSPPOpts: [], metodePengadaanOpts: [], jenisPengadaanOpts: [], kategoriPekerjaanOpts: [],
    aktifCount: 0, outstandingValue: 0, doneValue: 0, cancelValue: 0, criticalCount: 0,
    agingBuckets: { b0_30: 0, b31_45: 0, bOver45: 0 },
    priorityAction: [], monthlyTrend: []
  };
}

// ══════════════════════════════════════════════════════════
// TEST
// ══════════════════════════════════════════════════════════
function testGetData() {
  var spr = getSPRData();
  Logger.log('=== SPR ===');
  Logger.log('rows: ' + spr.totalSPR + ' | nominal: ' + spr.totalNominal);
  Logger.log('statusPengadaan: ' + JSON.stringify(spr.byStatusPengadaan));
  Logger.log('outstanding: count=' + spr.outstandingCount + ' value=' + spr.outstandingValue);
  Logger.log('financial exposure: done=' + spr.doneValue + ' cancel=' + spr.cancelValue + ' outstanding=' + spr.outstandingValue);
  Logger.log('aging>45: ' + spr.aging45Count + ' | buckets: ' + JSON.stringify(spr.agingBuckets));
  Logger.log('priorityAction (top10): ' + JSON.stringify(spr.priorityAction));
  Logger.log('monthlyTrend: ' + JSON.stringify(spr.monthlyTrend));

  // DEBUG: cek tglRilisSPR pada 5 row pertama yg statusnya On progress
  var onProgress = spr.rows.filter(function(r){ return r.statusPengadaan.toUpperCase() === 'ON PROGRESS OLEH MITRA'; }).slice(0,5);
  onProgress.forEach(function(r){
    Logger.log('DEBUG SPR row: noSPR=' + r.noSPR + ' | tglRilisSPR="' + r.tglRilisSPR + '" | tglBapsBast="' + r.tglBapsBast + '" | aging=' + r.aging);
  });

  var spp = getSPPData();
  Logger.log('=== SPP ===');
  Logger.log('rows: ' + spp.totalSPP + ' | nilai: ' + spp.totalNilai);
  Logger.log('status: ' + JSON.stringify(spp.byStatus));
  Logger.log('updateTerbaru: ' + JSON.stringify(spp.byUpdateTerbaru));
  Logger.log('aktif/outstanding: count=' + spp.aktifCount + ' value=' + spp.outstandingValue);
  Logger.log('financial exposure: done=' + spp.doneValue + ' cancel=' + spp.cancelValue);
  Logger.log('critical: ' + spp.criticalCount + ' | buckets: ' + JSON.stringify(spp.agingBuckets));
  Logger.log('priorityAction (top10, dual-condition): ' + JSON.stringify(spp.priorityAction));
  Logger.log('monthlyTrend: ' + JSON.stringify(spp.monthlyTrend));

  var perpanjanganCount = spp.rows.filter(function(r){return r.isPerpanjangan;}).length;
  Logger.log('SPP dengan Perpanjangan=Ya: ' + perpanjanganCount);

  // DEBUG: cek raw header SPP dan 5 row pertama
  var wsSPP = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SPP');
  var hdrSPP = wsSPP.getDataRange().getDisplayValues()[0];
  Logger.log('DEBUG SPP headers: ' + JSON.stringify(hdrSPP));
  spp.rows.slice(0,5).forEach(function(r){
    Logger.log('DEBUG SPP row: noSPP=' + r.noSPP + ' | tglRilisSP2K="' + r.tglRilisSP2K + '" | tglBapsBast="' + r.tglBapsBast + '" | durasiRaw="' + r.durasiPekerjaan + '" | durasiRencana=' + r.durasiRencana + ' | aging=' + r.aging + ' | isAktif=' + r.isAktif);
  });

  // DEBUG: cek raw header SPR
  var wsSPR = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SPR');
  var hdrSPR = wsSPR.getDataRange().getDisplayValues()[0];
  Logger.log('DEBUG SPR headers: ' + JSON.stringify(hdrSPR));
}
