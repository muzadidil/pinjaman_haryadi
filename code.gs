/**
 * Kelola Pinjaman Haryadi
 * Database = Google Sheet (sheet ini). UI = index.html (HtmlService).
 * Sheets dipakai: Pinjaman, Pembayaran, Peminjam, Akun (otomatis dibuat).
 */

var TZ = Session.getScriptTimeZone();

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Kelola Pinjaman')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .addMetaTag('mobile-web-app-capable', 'yes');
}

/* ---------- Helpers ---------- */

// Ensure sheet exists with headers. Returns the sheet.
// Kalau sheet sudah ada tapi baris 1 kosong (header terhapus), tulis ulang header otomatis.
function getSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
    return sh;
  }
  // Sheet sudah ada: perbaiki header bila baris 1 kosong.
  if (headers && headers.length) {
    var firstCell = sh.getRange(1, 1).getValue();
    if (String(firstCell).trim() === '') {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

// Create all sheets + seed default akun. Safe to call many times.
function ensureSetup_() {
  getSheet_('Pinjaman', ['ID', 'Waktu', 'Nama Peminjam', 'Akun Digunakan', 'Nominal', 'Tenor (bln)', 'Cicilan/bln', 'Jatuh Tempo Pertama', 'Jatuh Tempo Akhir', 'Akun TF']);
  getSheet_('Jadwal', ['LoanID', 'Nama Peminjam', 'Akun', 'Cicilan ke-', 'Jatuh Tempo', 'Nominal Cicilan']);
  getSheet_('Pembayaran', ['ID', 'Waktu', 'LoanID', 'Cicilan ke-', 'Nama Peminjam', 'Akun', 'Nominal Bayar', 'Catatan']);
  getSheet_('Peminjam', ['Nama']);
  var akun = getSheet_('Akun', ['Akun']);
  if (akun.getLastRow() < 2) {
    akun.getRange(2, 1, 2, 1).setValues([['SPINAM'], ['SeaBank']]); // seed default
  }
}

// Read one column (skip header) -> unique trimmed array. Batch read, no loop on sheet.
function readColumn_(sheetName, headers) {
  var sh = getSheet_(sheetName, headers);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, 1).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var v = String(values[i][0]).trim();
    if (v && out.indexOf(v) === -1) out.push(v);
  }
  return out;
}

function ribuan_(n) {
  n = Math.round(Number(n) || 0);
  var s = String(Math.abs(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (n < 0 ? '-' : '') + s;
}
function rupiah_(n) { return 'Rp ' + ribuan_(n); }
function fmtTanggal_(d) { return Utilities.formatDate(d, TZ, 'dd/MM/yyyy'); }
function fmtWaktu_(d) { return Utilities.formatDate(d, TZ, 'dd/MM/yyyy HH:mm'); }

// Jatuh tempo = tanggal + tenor bulan.
function hitungJatuhTempo_(base, tenor) {
  return new Date(base.getFullYear(), base.getMonth() + Number(tenor), base.getDate());
}

// Parse 'yyyy-MM-dd' (dari <input type=date>) -> Date lokal. Kosong/invalid -> null.
function parseTgl_(s) {
  s = String(s || '').trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/* ---------- API dipanggil dari HTML ---------- */

function getBootstrap() {
  ensureSetup_();
  return {
    peminjam: readColumn_('Peminjam', ['Nama']),
    akun: readColumn_('Akun', ['Akun'])
  };
}

function tambahPeminjam(nama) {
  nama = String(nama || '').trim();
  if (!nama) throw new Error('Nama kosong.');
  var list = readColumn_('Peminjam', ['Nama']);
  if (list.indexOf(nama) === -1) {
    getSheet_('Peminjam', ['Nama']).appendRow([nama]);
    list.push(nama);
  }
  return list;
}

function tambahAkun(akun) {
  akun = String(akun || '').trim();
  if (!akun) throw new Error('Akun kosong.');
  var list = readColumn_('Akun', ['Akun']);
  if (list.indexOf(akun) === -1) {
    getSheet_('Akun', ['Akun']).appendRow([akun]);
    list.push(akun);
  }
  return list;
}

// p: {nama, akun, nominal, tenor, cicilan, tglPertama, akunTf}
// tglPertama = tanggal jatuh tempo cicilan ke-1 (yyyy-MM-dd). Default = tgl pinjam + 1 bulan.
function simpanPinjaman(p) {
  ensureSetup_();
  var nama = String(p.nama || '').trim();
  var akun = String(p.akun || '').trim();
  var nominal = Number(p.nominal) || 0;
  var tenor = Number(p.tenor) || 0;
  var cicilan = Number(p.cicilan) || 0;
  var akunTf = String(p.akunTf || '').trim();
  if (!nama || !akun || !nominal || !tenor) throw new Error('Lengkapi: nama, akun, nominal, tenor.');

  var now = new Date();
  var loanId = now.getTime();

  // Tanggal jatuh tempo cicilan pertama
  var first = parseTgl_(p.tglPertama);
  if (!first) first = hitungJatuhTempo_(now, 1); // default: 1 bulan dari sekarang

  // Cicilan per bulan: pakai input manual; jika kosong, bagi rata.
  var perBulan = cicilan > 0 ? cicilan : Math.floor(nominal / tenor);

  // Generate jadwal: 1 baris per bulan. Cicilan terakhir menyerap sisa supaya total = nominal.
  var jadwalRows = [];
  var akumulasi = 0;
  var jtAkhir = first;
  for (var k = 1; k <= tenor; k++) {
    var due = new Date(first.getFullYear(), first.getMonth() + (k - 1), first.getDate());
    var nominalCic = (k < tenor) ? perBulan : (nominal - akumulasi); // baris terakhir = sisa
    akumulasi += (k < tenor) ? perBulan : nominalCic;
    jadwalRows.push([loanId, nama, akun, k, due, nominalCic]);
    jtAkhir = due;
  }
  // Tulis sekaligus (batch) supaya ringan.
  var jSh = getSheet_('Jadwal', []);
  jSh.getRange(jSh.getLastRow() + 1, 1, jadwalRows.length, 6).setValues(jadwalRows);

  getSheet_('Pinjaman', []).appendRow([loanId, now, nama, akun, nominal, tenor, perBulan, first, jtAkhir, akunTf]);
  return { ok: true, jatuhTempo: fmtTanggal_(first), jatuhTempoAkhir: fmtTanggal_(jtAkhir), waktu: fmtWaktu_(now) };
}

function getPinjaman() {
  var sh = getSheet_('Pinjaman', []);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 10).getValues();
  var out = [];
  for (var i = rows.length - 1; i >= 0; i--) { // newest first
    var r = rows[i];
    out.push({
      loanId: r[0],
      waktu: r[1] ? fmtWaktu_(new Date(r[1])) : '',
      nama: r[2], akun: r[3],
      nominal: Number(r[4]) || 0,
      tenor: r[5],
      cicilan: Number(r[6]) || 0,
      jatuhTempo: r[7] ? fmtTanggal_(new Date(r[7])) : '',
      jatuhTempoAkhir: r[8] ? fmtTanggal_(new Date(r[8])) : '',
      akunTf: r[9]
    });
  }
  return out;
}

// Daftar pinjaman 1 peminjam, utk dropdown di tab Pembayaran.
// Label = "dd/MM/yyyy - akun - Rp nominal". Value = loanId.
function getPinjamanByNama(nama) {
  nama = String(nama || '').trim();
  if (!nama) return [];
  var sh = getSheet_('Pinjaman', []);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 10).getValues();
  var out = [];
  for (var i = rows.length - 1; i >= 0; i--) { // newest first
    var r = rows[i];
    if (String(r[2]).trim() !== nama) continue;
    var tgl = r[1] ? fmtTanggal_(new Date(r[1])) : '';
    out.push({
      loanId: r[0],
      akun: r[3],
      nominal: Number(r[4]) || 0,
      label: tgl + ' - ' + r[3] + ' - ' + rupiah_(Number(r[4]) || 0)
    });
  }
  return out;
}

// p: {nama, loanId, ke, akun, nominal, catatan}
function simpanPembayaran(p) {
  ensureSetup_();
  var nama = String(p.nama || '').trim();
  var loanId = p.loanId ? Number(p.loanId) : '';
  var ke = p.ke ? Number(p.ke) : '';
  var akun = String(p.akun || '').trim();
  var nominal = Number(p.nominal) || 0;
  var catatan = String(p.catatan || '').trim();
  if (!nama || !nominal) throw new Error('Lengkapi: nama & nominal.');
  var now = new Date();
  getSheet_('Pembayaran', []).appendRow([now.getTime(), now, loanId, ke, nama, akun, nominal, catatan]);
  return { ok: true, waktu: fmtWaktu_(now) };
}

function getPembayaran() {
  var sh = getSheet_('Pembayaran', []);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 8).getValues();
  var out = [];
  for (var i = rows.length - 1; i >= 0; i--) {
    var r = rows[i];
    out.push({
      waktu: r[1] ? fmtWaktu_(new Date(r[1])) : '',
      loanId: r[2],
      ke: r[3],
      nama: r[4],
      akun: r[5],
      nominal: Number(r[6]) || 0,
      catatan: r[7]
    });
  }
  return out;
}

// Jadwal cicilan. nama opsional (kosong = semua). loanId opsional (filter 1 pinjaman).
// Tiap baris diberi flag lunas (true jika cicilan ke- itu sudah ada di Pembayaran).
function getJadwal(nama, loanId) {
  ensureSetup_();
  nama = String(nama || '').trim();
  loanId = loanId ? Number(loanId) : 0;

  // Set cicilan yang sudah dibayar: key = "loanId|ke".
  var paid = {};
  var bSh = getSheet_('Pembayaran', []);
  if (bSh.getLastRow() >= 2) {
    var bRows = bSh.getRange(2, 1, bSh.getLastRow() - 1, 8).getValues();
    for (var b = 0; b < bRows.length; b++) {
      var bk = Number(bRows[b][3]) || 0; // cicilan ke-
      if (bk) paid[Number(bRows[b][2]) + '|' + bk] = true;
    }
  }

  var sh = getSheet_('Jadwal', []);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 6).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var pn = String(r[1]).trim();
    if (nama && pn !== nama) continue;
    if (loanId && Number(r[0]) !== loanId) continue;
    out.push({
      loanId: r[0],
      nama: pn,
      akun: r[2],
      ke: r[3],
      jatuhTempo: r[4] ? fmtTanggal_(new Date(r[4])) : '',
      jatuhTempoSort: r[4] ? new Date(r[4]).getTime() : 0,
      nominal: Number(r[5]) || 0,
      lunas: !!paid[Number(r[0]) + '|' + (Number(r[3]) || 0)]
    });
  }
  out.sort(function (a, b) { return a.jatuhTempoSort - b.jatuhTempoSort; });
  return out;
}

// Cicilan yang BELUM dibayar untuk 1 pinjaman (loanId).
// = baris di Jadwal yang nomor 'ke'-nya belum ada di Pembayaran (loanId sama).
function getCicilanBelumLunas(loanId) {
  ensureSetup_();
  loanId = Number(loanId) || 0;
  if (!loanId) return [];

  // Kumpulkan nomor cicilan yang sudah dibayar untuk loanId ini.
  var bSh = getSheet_('Pembayaran', []);
  var paid = {};
  if (bSh.getLastRow() >= 2) {
    var bRows = bSh.getRange(2, 1, bSh.getLastRow() - 1, 8).getValues();
    for (var b = 0; b < bRows.length; b++) {
      if (Number(bRows[b][2]) === loanId) {
        var ke = Number(bRows[b][3]) || 0;
        if (ke) paid[ke] = true;
      }
    }
  }

  // Ambil jadwal pinjaman ini, buang yang sudah dibayar.
  var jSh = getSheet_('Jadwal', []);
  if (jSh.getLastRow() < 2) return [];
  var jRows = jSh.getRange(2, 1, jSh.getLastRow() - 1, 6).getValues();
  var out = [];
  for (var i = 0; i < jRows.length; i++) {
    var r = jRows[i];
    if (Number(r[0]) !== loanId) continue;
    var nomor = Number(r[3]) || 0;
    if (paid[nomor]) continue; // sudah lunas, lewati
    out.push({
      loanId: r[0],
      ke: nomor,
      akun: r[2],
      jatuhTempo: r[4] ? fmtTanggal_(new Date(r[4])) : '',
      jatuhTempoSort: r[4] ? new Date(r[4]).getTime() : 0,
      nominal: Number(r[5]) || 0,
      label: 'Cicilan ke-' + nomor + ' - ' + (r[4] ? fmtTanggal_(new Date(r[4])) : '') + ' - ' + rupiah_(Number(r[5]) || 0)
    });
  }
  out.sort(function (a, b) { return a.jatuhTempoSort - b.jatuhTempoSort; });
  return out;
}

// nama opsional. Kosong = semua peminjam.
function getLaporan(nama) {
  ensureSetup_();
  nama = String(nama || '').trim();
  var pSh = getSheet_('Pinjaman', []);
  var bSh = getSheet_('Pembayaran', []);
  var pRows = pSh.getLastRow() < 2 ? [] : pSh.getRange(2, 1, pSh.getLastRow() - 1, 10).getValues();
  var bRows = bSh.getLastRow() < 2 ? [] : bSh.getRange(2, 1, bSh.getLastRow() - 1, 8).getValues();

  var bayarMap = {};
  for (var b = 0; b < bRows.length; b++) {
    var bn = String(bRows[b][4]).trim();
    if (!bn) continue;
    bayarMap[bn] = (bayarMap[bn] || 0) + (Number(bRows[b][6]) || 0);
  }

  var map = {};
  for (var i = 0; i < pRows.length; i++) {
    var r = pRows[i];
    var pn = String(r[2]).trim();
    if (!pn) continue;
    if (nama && pn !== nama) continue;
    if (!map[pn]) map[pn] = { nama: pn, pinjaman: [], totalPinjam: 0 };
    map[pn].pinjaman.push({
      tanggal: r[1] ? fmtTanggal_(new Date(r[1])) : '',
      akun: r[3],
      nominal: Number(r[4]) || 0,
      tenor: r[5],
      cicilan: Number(r[6]) || 0,
      jatuhTempo: r[7] ? fmtTanggal_(new Date(r[7])) : ''
    });
    map[pn].totalPinjam += Number(r[4]) || 0;
  }

  var out = [];
  for (var k in map) {
    var totalBayar = bayarMap[k] || 0;
    out.push({
      nama: map[k].nama,
      pinjaman: map[k].pinjaman,
      totalPinjam: map[k].totalPinjam,
      totalBayar: totalBayar,
      sisa: map[k].totalPinjam - totalBayar
    });
  }
  out.sort(function (a, b) { return a.nama.localeCompare(b.nama); });
  return out;
}

// PDF (base64) untuk 1 peminjam, atau semua jika nama kosong.
function exportPdf(nama) {
  var data = getLaporan(nama);
  var tglCetak = fmtWaktu_(new Date());

  var html = '<html><head><meta charset="utf-8"><style>'
    + 'body{font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#1a1a1a;padding:26px;}'
    + 'h1{font-size:18px;margin:0 0 2px;}'
    + '.meta{color:#666;font-size:11px;margin-bottom:16px;}'
    + 'h2{font-size:14px;margin:18px 0 6px;border-bottom:2px solid #333;padding-bottom:3px;}'
    + 'table{width:100%;border-collapse:collapse;margin-bottom:6px;}'
    + 'th,td{border:1px solid #ccc;padding:5px 7px;text-align:left;}'
    + 'th{background:#f0f0f0;} .r{text-align:right;}'
    + '.sum{font-size:12px;margin-top:3px;} .sisa{font-weight:bold;}'
    + '</style></head><body>';
  html += '<h1>Laporan Pinjaman</h1>';
  html += '<div class="meta">Dicetak: ' + tglCetak + (nama ? ' &middot; Peminjam: ' + nama : ' &middot; Semua Peminjam') + '</div>';

  if (!data.length) {
    html += '<p>Belum ada data.</p>';
  } else {
    for (var i = 0; i < data.length; i++) {
      var d = data[i];
      html += '<h2>' + d.nama + '</h2>';
      html += '<table><tr><th>Tgl Pinjam</th><th>Akun</th><th class="r">Nominal</th><th class="r">Tenor</th><th class="r">Cicilan/bln</th><th>Jatuh Tempo</th></tr>';
      for (var j = 0; j < d.pinjaman.length; j++) {
        var p = d.pinjaman[j];
        html += '<tr><td>' + p.tanggal + '</td><td>' + p.akun + '</td><td class="r">' + rupiah_(p.nominal)
          + '</td><td class="r">' + p.tenor + ' bln</td><td class="r">' + (p.cicilan ? rupiah_(p.cicilan) : '-') + '</td><td>' + p.jatuhTempo + '</td></tr>';
      }
      html += '</table>';
      html += '<div class="sum">Total Pinjaman: ' + rupiah_(d.totalPinjam) + '</div>';
      html += '<div class="sum">Total Pembayaran: ' + rupiah_(d.totalBayar) + '</div>';
      html += '<div class="sum sisa">Sisa Tagihan: ' + rupiah_(d.sisa) + '</div>';
    }
  }
  html += '</body></html>';

  var pdf = Utilities.newBlob(html, 'text/html', 'laporan.html').getAs('application/pdf');
  var fname = 'Laporan_' + (nama ? nama.replace(/\s+/g, '_') : 'Semua')
    + '_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm') + '.pdf';
  return { base64: Utilities.base64Encode(pdf.getBytes()), filename: fname };
}