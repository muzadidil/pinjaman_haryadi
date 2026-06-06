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
function getSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Create all sheets + seed default akun. Safe to call many times.
function ensureSetup_() {
  getSheet_('Pinjaman', ['ID', 'Waktu', 'Nama Peminjam', 'Akun Digunakan', 'Nominal', 'Tenor (bln)', 'Cicilan/bln', 'Jatuh Tempo', 'Akun TF']);
  getSheet_('Pembayaran', ['ID', 'Waktu', 'Nama Peminjam', 'Nominal Bayar', 'Catatan']);
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

// p: {nama, akun, nominal, tenor, akunTf}
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
  var jt = hitungJatuhTempo_(now, tenor);
  getSheet_('Pinjaman', []).appendRow([now.getTime(), now, nama, akun, nominal, tenor, cicilan, jt, akunTf]);
  return { ok: true, jatuhTempo: fmtTanggal_(jt), waktu: fmtWaktu_(now) };
}

function getPinjaman() {
  var sh = getSheet_('Pinjaman', []);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 9).getValues();
  var out = [];
  for (var i = rows.length - 1; i >= 0; i--) { // newest first
    var r = rows[i];
    out.push({
      waktu: r[1] ? fmtWaktu_(new Date(r[1])) : '',
      nama: r[2], akun: r[3],
      nominal: Number(r[4]) || 0,
      tenor: r[5],
      cicilan: Number(r[6]) || 0,
      jatuhTempo: r[7] ? fmtTanggal_(new Date(r[7])) : '',
      akunTf: r[8]
    });
  }
  return out;
}

// p: {nama, nominal, catatan}
function simpanPembayaran(p) {
  ensureSetup_();
  var nama = String(p.nama || '').trim();
  var nominal = Number(p.nominal) || 0;
  var catatan = String(p.catatan || '').trim();
  if (!nama || !nominal) throw new Error('Lengkapi: nama & nominal.');
  var now = new Date();
  getSheet_('Pembayaran', []).appendRow([now.getTime(), now, nama, nominal, catatan]);
  return { ok: true, waktu: fmtWaktu_(now) };
}

function getPembayaran() {
  var sh = getSheet_('Pembayaran', []);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 5).getValues();
  var out = [];
  for (var i = rows.length - 1; i >= 0; i--) {
    var r = rows[i];
    out.push({
      waktu: r[1] ? fmtWaktu_(new Date(r[1])) : '',
      nama: r[2],
      nominal: Number(r[3]) || 0,
      catatan: r[4]
    });
  }
  return out;
}

// nama opsional. Kosong = semua peminjam.
function getLaporan(nama) {
  ensureSetup_();
  nama = String(nama || '').trim();
  var pSh = getSheet_('Pinjaman', []);
  var bSh = getSheet_('Pembayaran', []);
  var pRows = pSh.getLastRow() < 2 ? [] : pSh.getRange(2, 1, pSh.getLastRow() - 1, 9).getValues();
  var bRows = bSh.getLastRow() < 2 ? [] : bSh.getRange(2, 1, bSh.getLastRow() - 1, 5).getValues();

  var bayarMap = {};
  for (var b = 0; b < bRows.length; b++) {
    var bn = String(bRows[b][2]).trim();
    if (!bn) continue;
    bayarMap[bn] = (bayarMap[bn] || 0) + (Number(bRows[b][3]) || 0);
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