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

  // Cicilan per bulan: pakai input manual (ikut tagihan, sudah termasuk bunga); jika kosong, bagi rata.
  var cicilanManual = cicilan > 0;
  var perBulan = cicilanManual ? cicilan : Math.floor(nominal / tenor);

  // Bangun baris jadwal + tulis batch. Manual -> tiap cicilan = perBulan; auto -> terakhir serap sisa.
  var jadwalRows = buatJadwalRows_(loanId, nama, akun, nominal, tenor, perBulan, first, cicilanManual);
  var jtAkhir = jadwalRows[jadwalRows.length - 1][4];
  var jSh = getSheet_('Jadwal', []);
  jSh.getRange(jSh.getLastRow() + 1, 1, jadwalRows.length, 6).setValues(jadwalRows);

  // Simpan perBulan ke kolom Cicilan/bln HANYA jika manual. Auto-divide -> kosong, supaya bila
  // jadwalnya pernah dihapus lalu di-Generate ulang, tetap terdeteksi sebagai auto (bukan manual).
  getSheet_('Pinjaman', []).appendRow([loanId, now, nama, akun, nominal, tenor, (cicilanManual ? perBulan : ''), first, jtAkhir, akunTf]);
  return { ok: true, jatuhTempo: fmtTanggal_(first), jatuhTempoAkhir: fmtTanggal_(jtAkhir), waktu: fmtWaktu_(now) };
}

// Bangun array baris Jadwal untuk 1 pinjaman.
// cicilanManual = true  -> tiap cicilan = perBulan apa adanya (ikut form Cicilan/bln yang
//                          sudah termasuk bunga). Total jadwal = perBulan x tenor (boleh > pokok).
// cicilanManual = false -> bagi rata; cicilan terakhir menyerap sisa pembulatan agar
//                          total jadwal = nominal pokok persis.
// Return: [[loanId, nama, akun, ke, dueDate, nominalCic], ...]
function buatJadwalRows_(loanId, nama, akun, nominal, tenor, perBulan, first, cicilanManual) {
  var rows = [];
  var akumulasi = 0;
  for (var k = 1; k <= tenor; k++) {
    var due = new Date(first.getFullYear(), first.getMonth() + (k - 1), first.getDate());
    var nominalCic = cicilanManual
      ? perBulan                                       // ikut tagihan per bulan apa adanya
      : ((k < tenor) ? perBulan : (nominal - akumulasi)); // bagi rata: terakhir serap sisa
    akumulasi += nominalCic;
    rows.push([loanId, nama, akun, k, due, nominalCic]);
  }
  return rows;
}

// Untuk DATA LAMA yang diketik manual di sheet Pinjaman.
// - Isi kolom ID (LoanID) yang kosong dengan angka unik.
// - Buat baris Jadwal untuk pinjaman yang belum punya jadwal.
// Aman dipanggil berulang (idempoten): pinjaman yang sudah ada jadwalnya dilewati.
function generateJadwalDariSheet() {
  ensureSetup_();
  var pSh = getSheet_('Pinjaman', []);
  var last = pSh.getLastRow();
  if (last < 2) return { ok: true, dibuat: 0, idDiisi: 0, pesan: 'Tidak ada pinjaman.' };

  var pRows = pSh.getRange(2, 1, last - 1, 10).getValues();

  // LoanID yang SUDAH punya jadwal.
  var jSh = getSheet_('Jadwal', []);
  var adaJadwal = {};
  if (jSh.getLastRow() >= 2) {
    var jIds = jSh.getRange(2, 1, jSh.getLastRow() - 1, 1).getValues();
    for (var x = 0; x < jIds.length; x++) adaJadwal[Number(jIds[x][0])] = true;
  }

  var newJadwal = [];
  var idDiisi = 0, dibuat = 0;
  var baseId = new Date().getTime();

  for (var i = 0; i < pRows.length; i++) {
    var r = pRows[i];
    var nama = String(r[2]).trim();
    if (!nama) continue; // baris kosong, lewati

    var loanId = Number(r[0]);
    if (!loanId) { // ID kosong -> isi otomatis (unik: base + nomor baris)
      loanId = baseId + i;
      pSh.getRange(i + 2, 1).setValue(loanId);
      idDiisi++;
    }
    if (adaJadwal[loanId]) continue; // sudah ada jadwal

    var akun = String(r[3]).trim();
    var nominal = Number(r[4]) || 0;
    var tenor = Number(r[5]) || 0;
    var cicilanManual = (Number(r[6]) || 0) > 0; // kolom "Cicilan/bln" diisi manual?
    var perBulan = cicilanManual ? Number(r[6]) : (tenor ? Math.floor(nominal / tenor) : 0);
    var first = r[7] ? new Date(r[7]) : null;
    if (!first) { // tempo pertama kosong -> pakai waktu + 1 bln, atau hari ini + 1 bln
      var base = r[1] ? new Date(r[1]) : new Date();
      first = hitungJatuhTempo_(base, 1);
      pSh.getRange(i + 2, 8).setValue(first);
    }
    if (!nominal || !tenor) continue; // data tak lengkap, tak bisa dijadwalkan

    var rowsJ = buatJadwalRows_(loanId, nama, akun, nominal, tenor, perBulan, first, cicilanManual);
    // Update Jatuh Tempo Akhir di sheet Pinjaman.
    pSh.getRange(i + 2, 9).setValue(rowsJ[rowsJ.length - 1][4]);
    for (var y = 0; y < rowsJ.length; y++) newJadwal.push(rowsJ[y]);
    dibuat++;
  }

  if (newJadwal.length) {
    jSh.getRange(jSh.getLastRow() + 1, 1, newJadwal.length, 6).setValues(newJadwal);
  }
  return { ok: true, dibuat: dibuat, idDiisi: idDiisi,
    pesan: dibuat + ' pinjaman dibuatkan jadwal, ' + idDiisi + ' ID diisi otomatis.' };
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

  // Map loanId -> nominal pokok (utk tampil "Pinjaman Pokok" di kartu cicilan).
  var pokokMap = {};
  var pSh = getSheet_('Pinjaman', []);
  if (pSh.getLastRow() >= 2) {
    var pRows = pSh.getRange(2, 1, pSh.getLastRow() - 1, 10).getValues();
    for (var p = 0; p < pRows.length; p++) pokokMap[Number(pRows[p][0])] = Number(pRows[p][4]) || 0;
  }

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
      pokok: pokokMap[Number(r[0])] || 0,
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
  var jSh = getSheet_('Jadwal', []);
  var pRows = pSh.getLastRow() < 2 ? [] : pSh.getRange(2, 1, pSh.getLastRow() - 1, 10).getValues();
  var bRows = bSh.getLastRow() < 2 ? [] : bSh.getRange(2, 1, bSh.getLastRow() - 1, 8).getValues();
  var jRows = jSh.getLastRow() < 2 ? [] : jSh.getRange(2, 1, jSh.getLastRow() - 1, 6).getValues();

  var bayarMap = {};
  for (var b = 0; b < bRows.length; b++) {
    var bn = String(bRows[b][4]).trim();
    if (!bn) continue;
    bayarMap[bn] = (bayarMap[bn] || 0) + (Number(bRows[b][6]) || 0);
  }

  // Total tagihan per pinjaman = jumlah seluruh cicilan di Jadwal (sudah termasuk bunga).
  // Inilah yang harus dibayar peminjam, bisa lebih besar dari pokok.
  var tagihanByLoan = {};
  for (var t = 0; t < jRows.length; t++) {
    var lid = Number(jRows[t][0]);
    tagihanByLoan[lid] = (tagihanByLoan[lid] || 0) + (Number(jRows[t][5]) || 0);
  }

  var map = {};
  for (var i = 0; i < pRows.length; i++) {
    var r = pRows[i];
    var pn = String(r[2]).trim();
    if (!pn) continue;
    if (nama && pn !== nama) continue;
    if (!map[pn]) map[pn] = { nama: pn, pinjaman: [], totalPinjam: 0, totalTagihan: 0 };
    var pokok = Number(r[4]) || 0;
    var lid2 = Number(r[0]);
    // Pinjaman tanpa jadwal -> pakai pokok sebagai tagihan (fallback aman).
    var tagihan = tagihanByLoan.hasOwnProperty(lid2) ? tagihanByLoan[lid2] : pokok;
    map[pn].pinjaman.push({
      tanggal: r[1] ? fmtTanggal_(new Date(r[1])) : '',
      akun: r[3],
      nominal: pokok,
      tenor: r[5],
      cicilan: Number(r[6]) || 0,
      jatuhTempo: r[7] ? fmtTanggal_(new Date(r[7])) : ''
    });
    map[pn].totalPinjam += pokok;
    map[pn].totalTagihan += tagihan;
  }

  var out = [];
  for (var k in map) {
    var totalBayar = bayarMap[k] || 0;
    out.push({
      nama: map[k].nama,
      pinjaman: map[k].pinjaman,
      totalPinjam: map[k].totalPinjam,
      totalTagihan: map[k].totalTagihan,
      totalBayar: totalBayar,
      sisa: map[k].totalTagihan - totalBayar
    });
  }
  out.sort(function (a, b) { return a.nama.localeCompare(b.nama); });
  return out;
}

// QR code sebagai data-URI base64 (PNG). Di-embed ke PDF supaya mandiri (tak perlu internet saat dibuka).
// text = isi QR (mis. URL web app + nomor laporan). Gagal fetch -> kembalikan '' (PDF tetap jadi).
function qrDataUri_(text) {
  try {
    var url = 'https://quickchart.io/qr?margin=1&size=150&text=' + encodeURIComponent(text);
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return '';
    var b64 = Utilities.base64Encode(resp.getBlob().getBytes());
    return 'data:image/png;base64,' + b64;
  } catch (e) {
    return '';
  }
}

// Diagnostik QR: JALANKAN MANUAL dari editor Apps Script (pilih fungsi "cekQr" lalu klik Run).
// - Jika muncul dialog minta izin "Connect to an external service" -> IZINKAN. Itulah penyebab QR
//   tidak muncul di PDF (izin akses internet belum diberikan akun pemilik). Setelah diizinkan, QR muncul.
// - Jika mengembalikan "HTTP 200 ..." tanpa error, berarti pengambilan QR sudah OK.
// Sengaja TANPA try/catch agar error izin tampil jelas.
function cekQr() {
  var url = 'https://quickchart.io/qr?margin=1&size=150&text=' + encodeURIComponent('TES QR');
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var n = resp.getBlob().getBytes().length;
  var hasil = 'HTTP ' + resp.getResponseCode() + ', ukuran gambar ' + n + ' byte. '
    + (resp.getResponseCode() === 200 ? 'Izin akses internet OK — QR seharusnya bisa muncul di PDF.' : 'Layanan QR bermasalah.');
  Logger.log(hasil);
  return hasil;
}

// URL web app (untuk QR). Kosong kalau belum di-deploy sebagai web app.
function webAppUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

// PDF laporan keuangan profesional untuk 1 peminjam (atau semua jika nama kosong).
function exportPdf(nama) {
  var data = getLaporan(nama);
  var now = new Date();
  var tglCetak = fmtWaktu_(now);
  var noLaporan = 'KP-' + Utilities.formatDate(now, TZ, 'yyyyMMdd-HHmmss');

  // QR berisi info verifikasi + link app. Saat di-scan membuka web app ini.
  var qrText = (webAppUrl_() || 'Kelola Pinjaman Haryadi')
    + ' | No:' + noLaporan
    + ' | ' + (nama || 'SEMUA')
    + ' | ' + Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm');
  var qr = qrDataUri_(qrText);

  // Hitung grand total bila semua peminjam.
  var gTotPinjam = 0, gTotTagihan = 0, gTotBayar = 0, gSisa = 0;
  for (var z = 0; z < data.length; z++) { gTotPinjam += data[z].totalPinjam; gTotTagihan += (data[z].totalTagihan != null ? data[z].totalTagihan : data[z].totalPinjam); gTotBayar += data[z].totalBayar; gSisa += data[z].sisa; }

  var BLUE = '#0066AE', DARK = '#003e73', GOLD = '#F2A900';

  var css = ''
    + '@page{ margin:0; }'
    + '*{ box-sizing:border-box; }'
    + 'body{ font-family:Arial,Helvetica,sans-serif; font-size:11px; color:#1f2937; margin:0; }'
    + '.page{ padding:0 32px 90px; }'
    + '.hd{ background:linear-gradient(135deg,'+BLUE+','+DARK+'); color:#fff; padding:22px 32px; display:flex; justify-content:space-between; align-items:flex-start; }'
    + '.hd .l{ display:flex; gap:12px; align-items:center; }'
    + '.hd .logo{ width:44px; height:44px; border-radius:10px; background:'+GOLD+'; color:'+DARK+'; font-weight:800; font-size:22px; display:flex; align-items:center; justify-content:center; }'
    + '.hd h1{ font-size:19px; margin:0; letter-spacing:.3px; }'
    + '.hd .s{ font-size:10.5px; color:rgba(255,255,255,.82); margin-top:2px; }'
    + '.hd .r{ text-align:right; }'
    + '.hd .qr{ width:78px; height:78px; background:#fff; padding:4px; border-radius:8px; }'
    + '.hd .no{ font-size:9.5px; color:rgba(255,255,255,.82); margin-top:5px; }'
    + '.ribbon{ height:5px; background:'+GOLD+'; }'
    + '.metarow{ display:flex; justify-content:space-between; font-size:10px; color:#6b7280; padding:12px 0 4px; border-bottom:1px solid #e5e7eb; margin-bottom:14px; }'
    + '.cards{ display:flex; gap:10px; margin:14px 0 6px; }'
    + '.kpi{ flex:1; border:1px solid #e5e7eb; border-radius:10px; padding:11px 13px; }'
    + '.kpi .t{ font-size:9px; text-transform:uppercase; letter-spacing:.5px; color:#6b7280; }'
    + '.kpi .v{ font-size:15px; font-weight:800; margin-top:3px; color:'+DARK+'; }'
    + '.kpi.bad .v{ color:#dc2626; } .kpi.ok .v{ color:#0a8f4d; }'
    + 'h2{ font-size:13px; margin:18px 0 7px; color:'+DARK+'; border-left:4px solid '+GOLD+'; padding-left:8px; }'
    + 'table{ width:100%; border-collapse:collapse; margin-bottom:4px; }'
    + 'th,td{ padding:6px 8px; text-align:left; font-size:10px; }'
    + 'thead th{ background:'+BLUE+'; color:#fff; font-weight:600; text-transform:uppercase; font-size:9px; letter-spacing:.3px; }'
    + 'tbody tr:nth-child(even){ background:#f6f9fc; }'
    + 'tbody td{ border-bottom:1px solid #e8edf3; }'
    + '.r{ text-align:right; } .c{ text-align:center; }'
    + '.lunas{ color:#0a8f4d; font-weight:700; } .belum{ color:#dc2626; font-weight:700; }'
    + '.subt{ display:flex; justify-content:flex-end; gap:26px; font-size:10.5px; margin:6px 2px 0; }'
    + '.subt b{ color:'+DARK+'; }'
    + '.sisarow{ display:flex; justify-content:flex-end; margin-top:5px; }'
    + '.sisabox{ background:#fff4f4; border:1px solid #fecaca; border-radius:8px; padding:7px 14px; font-weight:800; color:#dc2626; font-size:12px; }'
    + '.sisabox.ok{ background:#ecfdf3; border-color:#a7f3d0; color:#0a8f4d; }'
    + '.foot{ position:fixed; bottom:0; left:0; right:0; padding:10px 32px; border-top:1px solid #e5e7eb; font-size:8.5px; color:#9ca3af; display:flex; justify-content:space-between; }'
    + '.sign{ margin-top:30px; display:flex; justify-content:flex-end; }'
    + '.sign .box{ text-align:center; font-size:10px; color:#374151; }'
    + '.sign .line{ margin-top:46px; border-top:1px solid #9ca3af; padding-top:4px; width:170px; }';

  var html = '<html><head><meta charset="utf-8"><style>' + css + '</style></head><body>';

  // ---- Header band ----
  html += '<div class="hd"><div class="l"><div class="logo">H</div>'
    + '<div><h1>LAPORAN PINJAMAN</h1><div class="s">Kelola Pinjaman &middot; atas nama Haryadi</div></div></div>'
    + '<div class="r">'
    + (qr ? '<img class="qr" src="' + qr + '">' : '')
    + '<div class="no">No. ' + noLaporan + '</div></div></div>';
  html += '<div class="ribbon"></div>';

  html += '<div class="page">';
  html += '<div class="metarow"><span>Dicetak: ' + tglCetak + '</span><span>'
    + (nama ? 'Peminjam: <b>' + nama + '</b>' : 'Cakupan: <b>Semua Peminjam</b>') + '</span></div>';

  if (!data.length) {
    html += '<p style="color:#6b7280">Belum ada data untuk ditampilkan.</p>';
  } else {
    // KPI ringkasan (grand total).
    html += '<div class="cards">'
      + '<div class="kpi"><div class="t">Pinjaman Pokok</div><div class="v">' + rupiah_(gTotPinjam) + '</div></div>'
      + '<div class="kpi"><div class="t">Total Tagihan</div><div class="v">' + rupiah_(gTotTagihan) + '</div></div>'
      + '<div class="kpi ok"><div class="t">Total Pembayaran</div><div class="v">' + rupiah_(gTotBayar) + '</div></div>'
      + '<div class="kpi ' + (gSisa > 0 ? 'bad' : 'ok') + '"><div class="t">Sisa Tagihan</div><div class="v">' + rupiah_(gSisa) + '</div></div>'
      + '</div>';

    for (var i = 0; i < data.length; i++) {
      var d = data[i];
      html += '<h2>' + d.nama + '</h2>';

      // Tabel pinjaman.
      html += '<table><thead><tr><th>Tgl Pinjam</th><th>Akun</th><th class="r">Nominal</th><th class="c">Tenor</th><th class="r">Cicilan/bln</th><th>Tempo</th></tr></thead><tbody>';
      for (var j = 0; j < d.pinjaman.length; j++) {
        var p = d.pinjaman[j];
        html += '<tr><td>' + p.tanggal + '</td><td>' + p.akun + '</td><td class="r">' + rupiah_(p.nominal)
          + '</td><td class="c">' + p.tenor + ' bln</td><td class="r">' + (p.cicilan ? rupiah_(p.cicilan) : '-') + '</td><td>' + p.jatuhTempo + '</td></tr>';
      }
      html += '</tbody></table>';

      // Jadwal cicilan + status (dari sheet Jadwal).
      var jad = getJadwal(d.nama, '');
      if (jad.length) {
        html += '<table style="margin-top:8px"><thead><tr><th class="c">Ke-</th><th>Akun</th><th>Jatuh Tempo</th><th class="r">Nominal</th><th class="c">Status</th></tr></thead><tbody>';
        for (var s = 0; s < jad.length; s++) {
          var c = jad[s];
          html += '<tr><td class="c">' + c.ke + '</td><td>' + c.akun + '</td><td>' + c.jatuhTempo + '</td><td class="r">' + rupiah_(c.nominal)
            + '</td><td class="c ' + (c.lunas ? 'lunas">LUNAS' : 'belum">BELUM') + '</td></tr>';
        }
        html += '</tbody></table>';
      }

      html += '<div class="subt"><span>Pinjaman Pokok: <b>' + rupiah_(d.totalPinjam) + '</b></span>'
        + '<span>Total Tagihan: <b>' + rupiah_(d.totalTagihan != null ? d.totalTagihan : d.totalPinjam) + '</b></span>'
        + '<span>Total Bayar: <b>' + rupiah_(d.totalBayar) + '</b></span></div>';
      html += '<div class="sisarow"><div class="sisabox ' + (d.sisa > 0 ? '' : 'ok') + '">Sisa Tagihan: ' + rupiah_(d.sisa) + '</div></div>';
    }

    // Tanda tangan.
    html += '<div class="sign"><div class="box">Hormat kami,<div class="line">( Haryadi )</div></div></div>';
  }

  html += '</div>'; // .page

  // Footer (muncul tiap halaman).
  html += '<div class="foot"><span>Dokumen ini dibuat otomatis oleh sistem Kelola Pinjaman.</span>'
    + '<span>Scan QR untuk verifikasi &middot; ' + noLaporan + '</span></div>';

  html += '</body></html>';

  var pdf = Utilities.newBlob(html, 'text/html', 'laporan.html').getAs('application/pdf');
  var fname = 'Laporan_' + (nama ? nama.replace(/\s+/g, '_') : 'Semua')
    + '_' + Utilities.formatDate(now, TZ, 'yyyyMMdd_HHmm') + '.pdf';
  return { base64: Utilities.base64Encode(pdf.getBytes()), filename: fname };
}