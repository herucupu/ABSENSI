/**
 * ABSENSI & LAPORAN HARIAN — RSI At-Tin Husada Ngawi
 * Backend Google Apps Script (database = Google Sheets, foto = Google Drive)
 *
 * LANGKAH: tempel seluruh isi file ini di Apps Script, jalankan fungsi `setup`
 * satu kali (lihat password admin di Log eksekusi), lalu Deploy > Web app
 * (Execute as: Me, Who has access: Anyone).
 */

var TZ = 'Asia/Jakarta';
var SH_KARYAWAN = 'Karyawan';
var SH_CONFIG = 'Config';
var PFX_ABSEN = 'Absensi_';
var PFX_LAPOR = 'Laporan_';
var KARYAWAN_HDR = ['id', 'nip', 'nama', 'unit', 'username', 'salt', 'hash', 'role', 'aktif', 'wajibGanti', 'createdAt',
  'email', 'hp', 'tempHash', 'tempExp', 'regId'];
var SH_FOTO = 'FotoProfil';
var FOTO_HDR = ['userId', 'foto', 'updatedAt'];
var SH_RESET = 'ResetReq';
var RESET_HDR = ['id', 'userId', 'username', 'nama', 'unit', 'createdAt', 'status'];
var MAX_FOTO_PROFIL = 45000;     // karakter base64 (foto profil 160px ~ 8-12 KB)
var API_VERSION = 10;             // dinaikkan tiap ada aksi baru; dipakai aplikasi untuk mendeteksi backend lama
var TEMP_PW_MENIT_EMAIL = 60;
var TEMP_PW_MENIT_ADMIN = 24 * 60;
var ABSEN_HDR = ['id', 'tanggal', 'userId', 'nama', 'unit', 'shift', 'jamMasuk', 'jamPulang', 'tglPulang',
  'fotoMasuk', 'fotoPulang', 'latMasuk', 'lngMasuk', 'latPulang', 'lngPulang', 'jarakMasuk', 'jarakPulang',
  'flag', 'serverMasuk', 'serverPulang', 'catatan', 'pulangId'];
var LAPOR_HDR = ['id', 'tanggal', 'userId', 'nama', 'isi', 'updatedAt', 'shift'];   // laporan per tanggal + shift (shift kosong = laporan lama)
var LAPOR_HARI = 7;              // laporan boleh diisi/diubah sampai 7 hari ke belakang
var SHIFTS = ['P', 'S', 'D', 'M', 'L'];   // P=Pagi, S=Siang, D=Middle, M=Malam, L=Libur (lama, hanya koreksi)
var SHIFTS_PICK = ['P', 'S', 'D', 'M'];
var SH_UNIT = 'Unit';
var UNIT_HDR = ['nama'];
var FLAG_CEK = ['LOKASI_JAUH', 'GPS_GAGAL', 'JAM_BEDA', 'JADWAL_L'];
var SESSION_HOURS = 12;
var SESSION_HOURS_REMEMBER = 24 * 30;
var MAX_FOTO_B64 = 450000;       // ~330 KB, foto dari aplikasi sudah dikompres (~40 KB)
var DEFAULT_RADIUS = 500;        // meter, longgar
var JAM_BEDA_MENIT = 120;
var TAIL_ROWS = 1800;            // baris terakhir yang dipindai untuk absen terbuka

/* ====================== SETUP ====================== */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SH_KARYAWAN, KARYAWAN_HDR);
  ensureColumns_(ss.getSheetByName(SH_KARYAWAN), KARYAWAN_HDR);
  ensureSheet_(ss, SH_FOTO, FOTO_HDR);
  ensureSheet_(ss, SH_RESET, RESET_HDR);
  ensureSheet_(ss, SH_UNIT, UNIT_HDR);
  var cfg = ensureSheet_(ss, SH_CONFIG, ['key', 'value']);
  if (cfg.getLastRow() < 2) {
    cfg.getRange(2, 1, 2, 2).setValues([['lokasi', '[]'], ['retensiFotoHari', '90']]);
  }
  ensureSecret_();
  ensureFolder_();
  var sh = karSheet_();
  var created = '';
  if (sh.getLastRow() < 2) {
    var pw = randomPw_();
    saveKaryawanRow_(sh, { nip: 'ADMIN', nama: 'Administrator', unit: 'IT', username: 'admin', role: 'admin', aktif: 'Y' }, pw, false);
    created = 'Akun admin dibuat. username: admin   password: ' + pw;
  } else {
    created = 'Akun sudah ada (tidak diubah).';
  }
  seedUnits_();
  setupTrigger();
  Logger.log('=== SETUP SELESAI ===');
  Logger.log(created);
  Logger.log('Sekarang: Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).');
  Logger.log('Buka juga menu Pengaturan di aplikasi untuk mengisi titik lokasi RS.');
  Logger.log('CATATAN: versi ini bisa mengirim email (password sementara). Izinkan akses Gmail/Mail bila diminta.');
}

function resetPasswordAdmin() {
  var sh = karSheet_();
  var rows = readAll_(sh);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].username === 'admin') {
      var pw = randomPw_();
      var salt = randomSalt_();
      sh.getRange(i + 2, 6, 1, 2).setValues([[salt, hash_(salt, pw)]]);
      sh.getRange(i + 2, 10).setValue('Y');
      sh.getRange(i + 2, 9).setValue('Y');
      CacheService.getScriptCache().remove('U:' + rows[i].id);
      Logger.log('Password baru untuk admin: ' + pw);
      return;
    }
  }
  Logger.log('Akun admin tidak ditemukan. Jalankan setup().');
}

function setupTrigger() {
  var ts = ScriptApp.getProjectTriggers();
  for (var i = 0; i < ts.length; i++) {
    if (ts[i].getHandlerFunction() === 'hapusFotoLama') return;
  }
  ScriptApp.newTrigger('hapusFotoLama').timeBased().everyDays(1).atHour(2).create();
}

/* ====================== UTIL ====================== */
function ensureSheet_(ss, name, hdr) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#DCE8FA');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), hdr.length).setNumberFormat('@');
  }
  return sh;
}
function ensureSecret_() {
  var p = PropertiesService.getScriptProperties();
  var s = p.getProperty('SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    p.setProperty('SECRET', s);
  }
  return s;
}
function ensureFolder_() {
  var p = PropertiesService.getScriptProperties();
  var id = p.getProperty('FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* dibuat ulang */ }
  }
  var f = DriveApp.createFolder('Absensi RSI - Foto');
  p.setProperty('FOLDER_ID', f.getId());
  return f;
}
function subFolder_(key) {
  var root = ensureFolder_();
  var it = root.getFoldersByName(key);
  return it.hasNext() ? it.next() : root.createFolder(key);
}
var KAR_CHECKED_ = false;
function ensureColumns_(sh, hdr) {
  var last = sh.getLastColumn();
  var cur = last > 0 ? sh.getRange(1, 1, 1, last).getValues()[0] : [];
  var have = {};
  for (var i = 0; i < cur.length; i++) have[str_(cur[i])] = true;
  for (var j = 0; j < hdr.length; j++) {
    if (!have[hdr[j]]) {
      last++;
      sh.getRange(1, last).setNumberFormat('@').setValue(hdr[j]).setFontWeight('bold').setBackground('#DCE8FA');
    }
  }
}
function karSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_KARYAWAN);
  if (!sh) throw fail_('Backend belum disiapkan. Jalankan fungsi setup() di Apps Script.');
  if (!KAR_CHECKED_) { ensureColumns_(sh, KARYAWAN_HDR); KAR_CHECKED_ = true; }
  return sh;
}
function fotoSheet_() { return ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SH_FOTO, FOTO_HDR); }
function resetSheet_() { return ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SH_RESET, RESET_HDR); }
function fotoOf_(userId) {
  var cache = CacheService.getScriptCache();
  var c = cache.get('P:' + userId);
  if (c !== null && c !== undefined) return c;
  var f = findRowById_(fotoSheet_(), FOTO_HDR, userId);
  var v = f ? f.o.foto : '';
  if (v && v.length < 90000) cache.put('P:' + userId, v, 3600);
  return v;
}
function setFoto_(userId, dataUrl) {
  var sh = fotoSheet_();
  var f = findRowById_(sh, FOTO_HDR, userId);
  var o = { userId: userId, foto: dataUrl, updatedAt: nowJ_() };
  if (f) writeRow_(sh, FOTO_HDR, f.row, o); else appendObj_(sh, FOTO_HDR, o);
  CacheService.getScriptCache().remove('P:' + userId);
}
function validFotoProfil_(v) {
  v = String(v || '');
  return /^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/.test(v) && v.length <= MAX_FOTO_PROFIL;
}
function validEmail_(e) { return !e || /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e); }
function normWa_(v) {
  var d = String(v || '').replace(/[^0-9]/g, '');
  if (!d) return '';
  if (d.charAt(0) === '0') d = '62' + d.substring(1);
  return d;
}
function sendMail_(to, subject, body) {
  try { MailApp.sendEmail({ to: to, subject: subject, body: body, name: 'Absensi RSI At-Tin Husada' }); return true; }
  catch (e) { return false; }
}
function plusMinutesTs_(min) { return Utilities.formatDate(new Date(Date.now() + min * 60000), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function randomTempPw_() {
  var chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  var s = '';
  for (var i = 0; i < 8; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}
function randomSalt_() { return Utilities.getUuid().replace(/-/g, '').substring(0, 16); }
function randomPw_() {
  var chars = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var s = '';
  for (var i = 0; i < 10; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}
function hex_(bytes) {
  var s = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i]; if (b < 0) b += 256;
    s += (b < 16 ? '0' : '') + b.toString(16);
  }
  return s;
}
function hash_(salt, pw) {
  return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8));
}
function nowJ_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function str_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v);
}
function addDays_(dateStr, n) {
  var p = dateStr.split('-');
  var d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n));
  return d.getUTCFullYear() + '-' + pad_(d.getUTCMonth() + 1) + '-' + pad_(d.getUTCDate());
}
function pad_(n) { return (n < 10 ? '0' : '') + n; }
function tsMs_(s) { return new Date(String(s).replace(' ', 'T') + '+07:00').getTime(); }
function minutesOf_(hhmm) { var p = hhmm.split(':'); return (+p[0]) * 60 + (+p[1]); }
function utcMin_(dateStr, hhmm) {
  var p = dateStr.split('-'); var h = hhmm.split(':');
  return Date.UTC(+p[0], +p[1] - 1, +p[2], +h[0], +h[1]) / 60000;
}
/** Kunci periode (27 -> 26) bernama bulan awal: 2026-10 = 27 Okt s/d 26 Nov 2026 */
function periodKey_(dateStr) {
  var p = dateStr.split('-'); var y = +p[0], m = +p[1], d = +p[2];
  if (d < 27) { m -= 1; if (m < 1) { m = 12; y -= 1; } }
  return y + '-' + pad_(m);
}
function periodRange_(key) {
  var p = key.split('-'); var y = +p[0], m = +p[1];
  var ny = y, nm = m + 1; if (nm > 12) { nm = 1; ny += 1; }
  return { key: key, start: y + '-' + pad_(m) + '-27', end: ny + '-' + pad_(nm) + '-26' };
}
function prevKey_(key) {
  var p = key.split('-'); var y = +p[0], m = +p[1] - 1; if (m < 1) { m = 12; y -= 1; }
  return y + '-' + pad_(m);
}
function validKey_(k) { return /^\d{4}-(0[1-9]|1[0-2])$/.test(k || ''); }
function haversine_(lat1, lng1, lat2, lng2) {
  var R = 6371000, rad = Math.PI / 180;
  var dLat = (lat2 - lat1) * rad, dLng = (lng2 - lng1) * rad;
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function fail_(msg) { var e = new Error(msg); return e; }

/* ====================== SHEET HELPERS ====================== */
function readAll_(sh) {
  var n = sh.getLastRow();
  if (n < 2) return [];
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var vals = sh.getRange(2, 1, n - 1, hdr.length).getValues();
  var res = [];
  for (var i = 0; i < vals.length; i++) res.push(rowObj_(hdr, vals[i]));
  return res;
}
function rowObj_(hdr, row) {
  var o = {};
  for (var j = 0; j < hdr.length; j++) o[hdr[j]] = str_(row[j]);
  return o;
}
/** Baca N baris terakhir. Mengembalikan [{row: nomorBaris, o: objek}] */
function readTail_(sh, hdr, n) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var start = Math.max(2, last - n + 1);
  var vals = sh.getRange(start, 1, last - start + 1, hdr.length).getValues();
  var res = [];
  for (var i = 0; i < vals.length; i++) res.push({ row: start + i, o: rowObj_(hdr, vals[i]) });
  return res;
}
function absenSheet_(key, create) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = PFX_ABSEN + key;
  var sh = ss.getSheetByName(name);
  if (!sh && create) sh = ensureSheet_(ss, name, ABSEN_HDR);
  return sh;
}
function laporSheet_(key, create) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = PFX_LAPOR + key;
  var sh = ss.getSheetByName(name);
  if (!sh && create) sh = ensureSheet_(ss, name, LAPOR_HDR);
  if (sh) ensureColumns_(sh, LAPOR_HDR);
  return sh;
}
function objToRow_(hdr, o) {
  var r = [];
  for (var j = 0; j < hdr.length; j++) r.push(o[hdr[j]] === undefined || o[hdr[j]] === null ? '' : o[hdr[j]]);
  return r;
}
function appendObj_(sh, hdr, o) {
  var row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, hdr.length).setNumberFormat('@').setValues([objToRow_(hdr, o)]);
}
function writeRow_(sh, hdr, rowNum, o) {
  sh.getRange(rowNum, 1, 1, hdr.length).setNumberFormat('@').setValues([objToRow_(hdr, o)]);
}
function findRowById_(sh, hdr, id) {
  var last = sh.getLastRow();
  if (last < 2) return null;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (str_(ids[i][0]) === id) {
      var row = i + 2;
      var vals = sh.getRange(row, 1, 1, hdr.length).getValues()[0];
      return { row: row, o: rowObj_(hdr, vals) };
    }
  }
  return null;
}

/* ====================== KARYAWAN & AUTH ====================== */
function saveKaryawanRow_(sh, d, pw, wajibGanti) {
  var salt = randomSalt_();
  var o = {
    id: Utilities.getUuid().replace(/-/g, '').substring(0, 12),
    nip: d.nip || '', nama: d.nama, unit: d.unit || '', username: String(d.username).toLowerCase(),
    salt: salt, hash: hash_(salt, pw), role: d.role || 'karyawan', aktif: d.aktif || 'Y',
    wajibGanti: wajibGanti ? 'Y' : 'N', createdAt: nowJ_(),
    email: d.email || '', hp: d.hp || '', tempHash: '', tempExp: '', regId: d.regId || ''
  };
  appendObj_(sh, KARYAWAN_HDR, o);
  return o;
}
function pubUser_(u) {
  return { id: u.id, nip: u.nip, nama: u.nama, unit: u.unit, username: u.username, role: u.role, wajibGanti: u.wajibGanti === 'Y', email: u.email || '', hp: u.hp || '' };
}
function pubUserFull_(u) { var x = pubUser_(u); x.foto = fotoOf_(u.id); return x; }
function adminCounts_() {
  var pend = 0;
  var rows = readAll_(karSheet_());
  for (var i = 0; i < rows.length; i++) if (rows[i].aktif === 'P') pend++;
  var rs = resetSheet_(), rst = 0;
  var rr = readAll_(rs);
  for (var j = 0; j < rr.length; j++) if (rr[j].status === 'baru') rst++;
  return { pending: pend, reset: rst };
}
function getUserById_(id) {
  var cache = CacheService.getScriptCache();
  var c = cache.get('U:' + id);
  if (c) return JSON.parse(c);
  var rows = readAll_(karSheet_());
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].id === id) {
      cache.put('U:' + id, JSON.stringify(rows[i]), 300);
      return rows[i];
    }
  }
  return null;
}
function makeToken_(userId, hours) {
  var exp = Math.floor(Date.now() / 1000) + hours * 3600;
  var payload = userId + '.' + exp;
  var sig = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, ensureSecret_()));
  return payload + '.' + sig;
}
function auth_(token) {
  var err = fail_('UNAUTHORIZED');
  if (!token) throw err;
  var p = String(token).split('.');
  if (p.length !== 3) throw err;
  var payload = p[0] + '.' + p[1];
  var sig = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, ensureSecret_()));
  if (sig !== p[2]) throw err;
  if (+p[1] < Math.floor(Date.now() / 1000)) throw err;
  var u = getUserById_(p[0]);
  if (!u || u.aktif !== 'Y') throw err;
  return u;
}
function isManager_(u) { return u.role === 'admin' || u.role === 'atasan'; }
/** Atasan boleh MELIHAT semua karyawan; mengoreksi tetap dibatasi unitnya (canSee_). */
function canView_(viewer) { return isManager_(viewer); }
function canSee_(viewer, unit) { return viewer.role === 'admin' || (viewer.role === 'atasan' && viewer.unit === unit); }

function tempOk_(u, pw) {
  return !!(u.tempHash && u.tempExp && hash_(u.salt, pw) === u.tempHash && tsMs_(u.tempExp) > Date.now());
}
function actLogin_(p) {
  var username = String(p.username || '').trim().toLowerCase();
  var pw = String(p.password || '');
  if (!username || !pw) throw fail_('Isi username dan password.');
  var cache = CacheService.getScriptCache();
  var fk = 'F:' + username;
  var fails = +(cache.get(fk) || 0);
  if (fails >= 5) throw fail_('Terlalu banyak percobaan gagal. Coba lagi dalam 10 menit.');
  var sh = karSheet_();
  var rows = readAll_(sh);
  var u = null;
  for (var i = 0; i < rows.length; i++) if (rows[i].username === username) { u = rows[i]; break; }
  var pwOk = !!u && hash_(u.salt, pw) === u.hash;
  var tmpOk = !!u && !pwOk && tempOk_(u, pw);
  if (!u || (u.aktif !== 'Y' && u.aktif !== 'P') || !(pwOk || tmpOk)) {
    cache.put(fk, String(fails + 1), 600);
    throw fail_('Username atau password salah.');
  }
  cache.remove(fk);
  if (u.aktif === 'P') throw fail_('Pendaftaran Anda masih menunggu persetujuan admin. Silakan coba lagi nanti.');
  if (tmpOk) {
    var f = findRowById_(sh, KARYAWAN_HDR, u.id);
    if (f) { f.o.wajibGanti = 'Y'; writeRow_(sh, KARYAWAN_HDR, f.row, f.o); u.wajibGanti = 'Y'; }
    cache.remove('U:' + u.id);
  }
  var hours = String(p.remember) === '1' ? SESSION_HOURS_REMEMBER : SESSION_HOURS;
  return { token: makeToken_(u.id, hours), user: pubUserFull_(u) };
}
function actChangePassword_(u, p) {
  var oldPw = String(p.oldPassword || ''), nw = String(p.newPassword || '');
  if (nw.length < 6) throw fail_('Password baru minimal 6 karakter.');
  if (hash_(u.salt, oldPw) !== u.hash && !tempOk_(u, oldPw)) throw fail_('Password lama salah.');
  var sh = karSheet_();
  var f = findRowById_(sh, KARYAWAN_HDR, u.id);
  var salt = randomSalt_();
  f.o.salt = salt; f.o.hash = hash_(salt, nw); f.o.wajibGanti = 'N'; f.o.tempHash = ''; f.o.tempExp = '';
  writeRow_(sh, KARYAWAN_HDR, f.row, f.o);
  CacheService.getScriptCache().remove('U:' + u.id);
  return { ok: true };
}

/* ---------- pendaftaran mandiri (publik) ---------- */
function actCekUsername_(p) {
  var un = String(p.username || '').trim().toLowerCase();
  var nip = String(p.nip || '').trim().toLowerCase();
  var rows = readAll_(karSheet_());
  var userTaken = false, nipTaken = false;
  for (var i = 0; i < rows.length; i++) {
    if (un && rows[i].username === un) userTaken = true;
    if (nip && rows[i].nip && rows[i].nip.toLowerCase() === nip) nipTaken = true;
  }
  return { userTaken: userTaken, nipTaken: nipTaken };
}
function actCekDaftar_(p) {
  var id = String(p.regId || '');
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(id)) return { status: null };
  var rows = readAll_(karSheet_());
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].regId === id) return { status: rows[i].aktif === 'P' ? 'menunggu' : (rows[i].aktif === 'Y' ? 'aktif' : 'nonaktif') };
  }
  return { status: null };
}
/* ---------- master unit ---------- */
function unitSheet_() { return ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SH_UNIT, UNIT_HDR); }
function seedUnits_() {
  var sh = unitSheet_();
  if (sh.getLastRow() >= 2) return;
  var seen = {}, list = [];
  var rows = readAll_(karSheet_());
  for (var i = 0; i < rows.length; i++) { var n = String(rows[i].unit || '').trim(); if (n && !seen[n.toLowerCase()]) { seen[n.toLowerCase()] = 1; list.push(n); } }
  if (!list.length) list.push('IT');
  list.sort();
  for (var j = 0; j < list.length; j++) sh.appendRow([list[j]]);
  CacheService.getScriptCache().remove('UNITS');
}
function getUnits_() {
  var cache = CacheService.getScriptCache();
  var c = cache.get('UNITS');
  if (c) return JSON.parse(c);
  var sh = unitSheet_();
  if (sh.getLastRow() < 2) seedUnits_();
  var rows = readAll_(sh), out = [];
  for (var i = 0; i < rows.length; i++) { var n = String(rows[i].nama || '').trim(); if (n) out.push(n); }
  out.sort(function (a, b) { return a.toLowerCase() < b.toLowerCase() ? -1 : 1; });
  cache.put('UNITS', JSON.stringify(out), 600);
  return out;
}
/** Nama unit baku (huruf besar/kecil mengikuti master) atau '' jika tidak ada di master. */
function unitCanon_(name) {
  var n = String(name || '').trim().toLowerCase(), list = getUnits_();
  for (var i = 0; i < list.length; i++) if (list[i].toLowerCase() === n) return list[i];
  return '';
}
function addUnit_(name) {
  name = String(name || '').trim().substring(0, 60);
  if (!name) return '';
  var c = unitCanon_(name);
  if (c) return c;
  unitSheet_().appendRow([name]);
  CacheService.getScriptCache().remove('UNITS');
  return name;
}
function actUnitSave_(u, p) {
  if (u.role !== 'admin') throw fail_('Hanya admin.');
  var nama = String(p.nama || '').trim().substring(0, 60), lama = String(p.lama || '').trim();
  if (!nama) throw fail_('Nama unit wajib diisi.');
  var sh = unitSheet_(), rows = readAll_(sh);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].nama).toLowerCase() === nama.toLowerCase() && String(rows[i].nama) !== lama) throw fail_('Unit sudah ada.');
  }
  if (!lama) { sh.appendRow([nama]); }
  else {
    var f = null;
    for (var r = 0; r < rows.length; r++) if (rows[r].nama === lama) f = r + 2;
    if (!f) throw fail_('Unit tidak ditemukan.');
    sh.getRange(f, 1).setValue(nama);
    if (nama !== lama) {
      var ks = karSheet_(), kr = readAll_(ks), cache = CacheService.getScriptCache();
      var col = KARYAWAN_HDR.indexOf('unit') + 1;
      for (var k = 0; k < kr.length; k++) if (kr[k].unit === lama) { ks.getRange(k + 2, col).setValue(nama); cache.remove('U:' + kr[k].id); }
    }
  }
  CacheService.getScriptCache().remove('UNITS');
  return { ok: true, units: getUnits_() };
}
function actUnitHapus_(u, p) {
  if (u.role !== 'admin') throw fail_('Hanya admin.');
  var nama = String(p.nama || '').trim();
  var kr = readAll_(karSheet_()), n = 0;
  for (var k = 0; k < kr.length; k++) if (kr[k].unit === nama) n++;
  if (n) throw fail_('Unit masih dipakai ' + n + ' karyawan. Pindahkan dulu karyawannya.');
  var sh = unitSheet_(), rows = readAll_(sh);
  for (var i = 0; i < rows.length; i++) if (rows[i].nama === nama) { sh.deleteRow(i + 2); break; }
  CacheService.getScriptCache().remove('UNITS');
  return { ok: true, units: getUnits_() };
}
/** Siapa yang sedang jaga (absen masuk belum pulang), dikelompokkan per unit. Boleh dilihat semua karyawan. */
function actJaga_(u) {
  var cache = CacheService.getScriptCache();
  var c = cache.get('JAGA');
  if (c) return JSON.parse(c);
  var now = nowJ_(), k = periodKey_(now.substring(0, 10)), keys = [k, prevKey_(k)];
  var latest = {};
  for (var ki = 0; ki < keys.length; ki++) {
    var sh = absenSheet_(keys[ki], false);
    if (!sh) continue;
    var tail = readTail_(sh, ABSEN_HDR, TAIL_ROWS);
    for (var i = 0; i < tail.length; i++) {
      var o = tail[i].o;
      if (!isOpen_(o) || o.shift === 'L') continue;
      if ((tsMs_(now) - tsMs_(o.serverMasuk)) / 3600000 > 20) continue;
      if (!latest[o.userId] || latest[o.userId].serverMasuk < o.serverMasuk) latest[o.userId] = o;
    }
  }
  var hpMap = {}, users = readAll_(karSheet_());
  for (var j = 0; j < users.length; j++) hpMap[users[j].id] = { hp: users[j].hp || '', unit: users[j].unit, aktif: users[j].aktif };
  var byUnit = {};
  var units = getUnits_();
  for (var q = 0; q < units.length; q++) byUnit[units[q]] = [];
  for (var id in latest) {
    var r = latest[id], inf = hpMap[id] || {};
    if (inf.aktif && inf.aktif !== 'Y') continue;
    var un = r.unit || inf.unit || '-';
    if (!byUnit[un]) byUnit[un] = [];
    byUnit[un].push({ userId: id, nama: r.nama, shift: r.shift, jamMasuk: r.jamMasuk, tanggal: r.tanggal, hp: inf.hp || '' });
  }
  var out = [];
  for (var name in byUnit) {
    byUnit[name].sort(function (a, b) { return a.nama < b.nama ? -1 : 1; });
    out.push({ unit: name, orang: byUnit[name] });
  }
  out.sort(function (a, b) { return a.unit.toLowerCase() < b.unit.toLowerCase() ? -1 : 1; });
  var res = { now: now, units: out };
  cache.put('JAGA', JSON.stringify(res), 20);
  return res;
}
function actInfoPublik_() { var c = getConfig_(); return { waAdmin: c.waAdmin || '', ver: API_VERSION, units: getUnits_() }; }
function actSignup_(p) {
  var cache = CacheService.getScriptCache();
  var cnt = +(cache.get('SUN') || 0);
  if (cnt >= 60) throw fail_('Terlalu banyak pendaftaran. Coba lagi nanti.');
  cache.put('SUN', String(cnt + 1), 3600);
  var regId = String(p.regId || '');
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(regId)) throw fail_('ID pendaftaran tidak valid.');
  var nama = String(p.nama || '').trim().substring(0, 80);
  var nip = String(p.nip || '').trim().substring(0, 30);
  var unit = String(p.unit || '').trim().substring(0, 60);
  var username = String(p.username || '').trim().toLowerCase();
  var pw = String(p.password || '');
  var email = String(p.email || '').trim().toLowerCase().substring(0, 100);
  var hp = normWa_(p.hp).substring(0, 20);
  if (hp.length < 9) throw fail_('No. WhatsApp wajib diisi (dipakai untuk mengirim password baru).');
  if (nama.length < 3) throw fail_('Nama minimal 3 huruf.');
  if (!nip) throw fail_('NIP wajib diisi.');
  if (!unit) throw fail_('Unit wajib dipilih.');
  var canon = unitCanon_(unit);
  if (!canon) throw fail_('Unit tidak ada di daftar. Pilih unit dari daftar.');
  unit = canon;
  if (!/^[a-z0-9._-]{3,30}$/.test(username)) throw fail_('Username 3-30 karakter (huruf kecil/angka).');
  if (pw.length < 6) throw fail_('Password minimal 6 karakter.');
  if (!validEmail_(email)) throw fail_('Format email tidak valid.');
  if (!validFotoProfil_(p.foto)) throw fail_('Foto profil wajib (JPEG kecil).');
  var sh = karSheet_();
  var rows = readAll_(sh);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].regId === regId) return { ok: true, duplicate: true };
    if (rows[i].username === username) throw fail_('Username sudah dipakai.');
    if (rows[i].nip && rows[i].nip.toLowerCase() === nip.toLowerCase()) throw fail_('NIP sudah terdaftar.');
  }
  var o = saveKaryawanRow_(sh, { nip: nip, nama: nama, unit: unit, username: username, role: 'karyawan', aktif: 'P', email: email, hp: hp, regId: regId }, pw, false);
  setFoto_(o.id, String(p.foto));
  return { ok: true };
}

/* ---------- lupa password ---------- */
function waToken_() { return PropertiesService.getScriptProperties().getProperty('WA_GATEWAY_TOKEN') || ''; }
/** Kirim WhatsApp lewat gateway (Fonnte). Mengembalikan true bila gateway menjawab status sukses. */
function sendWa_(hp, text) {
  try {
    var res = UrlFetchApp.fetch('https://api.fonnte.com/send', { method: 'post', headers: { Authorization: waToken_() }, payload: { target: hp, message: text, countryCode: '62' }, muteHttpExceptions: true });
    var j = JSON.parse(res.getContentText() || '{}');
    return j.status === true || j.status === 'true';
  } catch (e) { return false; }
}
function actLupa_(p) {
  var username = String(p.username || '').trim().toLowerCase();
  var nip = String(p.nip || '').trim().toLowerCase();
  var hpIn = normWa_(p.hp || '');
  if (!username || (!hpIn && !nip)) throw fail_('Isi username dan no. WhatsApp.');
  var cache = CacheService.getScriptCache();
  var rk = 'RS:' + username;
  var n = +(cache.get(rk) || 0);
  if (n >= 3) throw fail_('Terlalu sering meminta reset. Coba lagi dalam beberapa jam atau hubungi admin.');
  cache.put(rk, String(n + 1), 21600);
  var sh = karSheet_();
  var rows = readAll_(sh), u = null;
  for (var i = 0; i < rows.length; i++) if (rows[i].username === username) { u = rows[i]; break; }
  // jawaban selalu sama agar tidak membocorkan akun mana yang ada
  // verifikasi: no. WhatsApp yang terdaftar di akun; bila akun belum punya no. WA, memakai NIP (cadangan)
  var cocok = u && u.aktif === 'Y' && (u.hp ? (hpIn && u.hp === hpIn) : (nip && u.nip && u.nip.toLowerCase() === nip));
  if (!cocok) return { ok: true };
  var tmp = randomTempPw_();
  var f = findRowById_(sh, KARYAWAN_HDR, u.id);
  var viaEmail = false, viaWa = false;
  // Bila gateway WhatsApp aktif (token diisi admin), password sementara langsung dikirim ke WhatsApp karyawan.
  if (u.hp && waToken_()) {
    f.o.tempHash = hash_(u.salt, tmp); f.o.tempExp = plusMinutesTs_(TEMP_PW_MENIT_EMAIL);
    writeRow_(sh, KARYAWAN_HDR, f.row, f.o); cache.remove('U:' + u.id);
    viaWa = sendWa_(u.hp, 'Assalamualaikum ' + u.nama + ',\n\nPASSWORD SEMENTARA AKUN ABSENSI\n- Username: ' + u.username + '\n- Password sementara: ' + tmp + '\n- Berlaku: ' + TEMP_PW_MENIT_EMAIL + ' menit\n\nCARA MASUK\n1. Buka aplikasi Absensi, login dengan username dan password sementara di atas.\n2. Buat password baru saat diminta.\n\nPassword lama Anda tetap berlaku sampai Anda login dengan password sementara. Jika bukan Anda yang meminta, abaikan pesan ini dan jangan bagikan ke siapa pun.\n\nTerima kasih.\nRSI At-Tin Husada Ngawi');
    if (!viaWa) { f.o.tempHash = ''; f.o.tempExp = ''; writeRow_(sh, KARYAWAN_HDR, f.row, f.o); cache.remove('U:' + u.id); }
  }
  if (viaWa) return { ok: true };
  // WhatsApp diprioritaskan: bila ada no. WA, permintaan masuk antrean admin yang mengirim password sementara via WA.
  // Email hanya dipakai bila karyawan tidak punya no. WA.
  if (u.email && !u.hp) {
    f.o.tempHash = hash_(u.salt, tmp); f.o.tempExp = plusMinutesTs_(TEMP_PW_MENIT_EMAIL);
    writeRow_(sh, KARYAWAN_HDR, f.row, f.o); cache.remove('U:' + u.id);
    viaEmail = sendMail_(u.email, 'Password sementara - Absensi RSI At-Tin Husada',
      'Halo ' + u.nama + ',\n\nPermintaan reset password diterima.\n\nUsername: ' + u.username + '\nPassword sementara: ' + tmp +
      '\nBerlaku ' + TEMP_PW_MENIT_EMAIL + ' menit.\n\nLogin dengan password sementara ini, lalu Anda akan diminta membuat password baru.\n' +
      'Password lama Anda tetap berlaku sampai Anda login dengan password sementara. Jika bukan Anda yang meminta, abaikan email ini.\n\nAbsensi RSI At-Tin Husada Ngawi');
    if (!viaEmail) { f.o.tempHash = ''; f.o.tempExp = ''; writeRow_(sh, KARYAWAN_HDR, f.row, f.o); cache.remove('U:' + u.id); }
  }
  if (!viaEmail) {
    var rs = resetSheet_();
    var all = readAll_(rs);
    for (var k = 0; k < all.length; k++) if (all[k].userId === u.id && all[k].status === 'baru') return { ok: true };
    appendObj_(rs, RESET_HDR, { id: Utilities.getUuid().replace(/-/g, '').substring(0, 12), userId: u.id, username: u.username, nama: u.nama, unit: u.unit, createdAt: nowJ_(), status: 'baru' });
  }
  return { ok: true };
}

/* ---------- admin: persetujuan & reset ---------- */
function actPendingList_(u) {
  if (u.role !== 'admin') throw fail_('Hanya admin.');
  var rows = readAll_(karSheet_()), out = [];
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].aktif !== 'P') continue;
    var r = rows[i];
    out.push({ id: r.id, nama: r.nama, nip: r.nip, unit: r.unit, username: r.username, email: r.email, hp: r.hp, createdAt: r.createdAt, foto: fotoOf_(r.id) });
  }
  return { rows: out };
}
function actApprove_(u, p) {
  if (u.role !== 'admin') throw fail_('Hanya admin.');
  var sh = karSheet_();
  var f = findRowById_(sh, KARYAWAN_HDR, String(p.id || ''));
  if (!f || f.o.aktif !== 'P') throw fail_('Pendaftaran tidak ditemukan.');
  if (p.approve) {
    var role = ['admin', 'atasan', 'karyawan'].indexOf(p.role) >= 0 ? p.role : 'karyawan';
    f.o.aktif = 'Y'; f.o.role = role;
    if (p.unit) f.o.unit = addUnit_(p.unit) || f.o.unit;
    writeRow_(sh, KARYAWAN_HDR, f.row, f.o);
    CacheService.getScriptCache().remove('U:' + f.o.id);
    if (f.o.email) sendMail_(f.o.email, 'Akun disetujui - Absensi RSI At-Tin Husada', 'Halo ' + f.o.nama + ',\n\nPendaftaran Anda disetujui. Silakan login dengan username "' + f.o.username + '" dan password yang Anda buat saat mendaftar.\n\nAbsensi RSI At-Tin Husada Ngawi');
  } else {
    sh.deleteRow(f.row);
    CacheService.getScriptCache().remove('U:' + f.o.id);
    if (f.o.email) sendMail_(f.o.email, 'Pendaftaran belum dapat disetujui - Absensi RSI At-Tin Husada', 'Halo ' + f.o.nama + ',\n\nMaaf, pendaftaran Anda belum dapat disetujui. Silakan hubungi admin IT.\n\nAbsensi RSI At-Tin Husada Ngawi');
  }
  return { ok: true };
}
function actResetList_(u) {
  if (u.role !== 'admin') throw fail_('Hanya admin.');
  var rr = readAll_(resetSheet_()), users = readAll_(karSheet_()), byId = {}, out = [];
  for (var i = 0; i < users.length; i++) byId[users[i].id] = users[i];
  for (var j = 0; j < rr.length; j++) {
    if (rr[j].status !== 'baru') continue;
    var us = byId[rr[j].userId] || {};
    out.push({ id: rr[j].id, userId: rr[j].userId, username: rr[j].username, nama: rr[j].nama, unit: rr[j].unit, createdAt: rr[j].createdAt, hp: us.hp || '', email: us.email || '' });
  }
  return { rows: out };
}
/** Admin membuat password sementara (berlaku 24 jam) lalu menyampaikannya lewat WhatsApp/langsung. */
function actResetBuat_(u, p) {
  if (u.role !== 'admin') throw fail_('Hanya admin.');
  var sh = karSheet_();
  var f = findRowById_(sh, KARYAWAN_HDR, String(p.userId || ''));
  if (!f || f.o.aktif !== 'Y') throw fail_('Akun tidak ditemukan.');
  var tmp = randomTempPw_();
  f.o.tempHash = hash_(f.o.salt, tmp); f.o.tempExp = plusMinutesTs_(TEMP_PW_MENIT_ADMIN);
  writeRow_(sh, KARYAWAN_HDR, f.row, f.o);
  CacheService.getScriptCache().remove('U:' + f.o.id);
  var rs = resetSheet_(), rr = readAll_(rs);
  for (var i = 0; i < rr.length; i++) {
    if (rr[i].userId === f.o.id && rr[i].status === 'baru') { rs.getRange(i + 2, RESET_HDR.indexOf('status') + 1).setNumberFormat('@').setValue('selesai'); }
  }
  return { password: tmp, nama: f.o.nama, username: f.o.username, hp: f.o.hp || '', berlakuMenit: TEMP_PW_MENIT_ADMIN };
}

/* ---------- foto profil ---------- */
function actSaveFotoProfil_(u, p) {
  if (!validFotoProfil_(p.foto)) throw fail_('Foto profil tidak valid.');
  setFoto_(u.id, String(p.foto));
  return { ok: true };
}
/** Karyawan mengubah data pribadinya sendiri (NIP, unit, username, peran hanya bisa diubah admin). */
function actSaveProfil_(u, p) {
  var nama = String(p.nama || '').trim().substring(0, 80);
  var email = String(p.email || '').trim().toLowerCase().substring(0, 100);
  if (nama.length < 3) throw fail_('Nama minimal 3 huruf.');
  if (!validEmail_(email)) throw fail_('Format email tidak valid.');
  var sh = karSheet_();
  var f = findRowById_(sh, KARYAWAN_HDR, u.id);
  if (!f) throw fail_('Akun tidak ditemukan.');
  f.o.nama = nama; f.o.email = email; f.o.hp = normWa_(p.hp).substring(0, 20);
  writeRow_(sh, KARYAWAN_HDR, f.row, f.o);
  CacheService.getScriptCache().remove('U:' + u.id);
  return { ok: true };
}
/** Foto profil beberapa karyawan sekaligus (untuk halaman Jaga). Foto profil boleh dilihat semua karyawan yang login. */
function actFotoBanyak_(u, p) {
  var ids = String(p.ids || '').split(',').slice(0, 30), map = {};
  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i] || '').trim();
    if (/^[A-Za-z0-9_-]{6,40}$/.test(id)) map[id] = fotoOf_(id);
  }
  return { map: map };
}
function actProfilFoto_(u, p) {
  var id = String(p.id || '');
  var t = getUserById_(id);
  if (!t) throw fail_('Pengguna tidak ditemukan.');
  if (t.id !== u.id && !canView_(u)) throw fail_('Tidak punya akses.');
  return { foto: fotoOf_(id) };
}

/* ====================== CONFIG ====================== */
function getConfig_() {
  var rows = readAll_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CONFIG));
  var c = { lokasi: [], retensiFotoHari: 90, waAdmin: '', semangat: '', semangatTgl: '' };
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].key === 'lokasi') { try { c.lokasi = JSON.parse(rows[i].value) || []; } catch (e) { c.lokasi = []; } }
    if (rows[i].key === 'retensiFotoHari') c.retensiFotoHari = parseInt(rows[i].value, 10) || 90;
    if (rows[i].key === 'waAdmin') c.waAdmin = rows[i].value;
    if (rows[i].key === 'semangat') c.semangat = String(rows[i].value || '');
    if (rows[i].key === 'semangatTgl') c.semangatTgl = String(rows[i].value || '');
  }
  return c;
}
function setConfigKey_(key, value) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CONFIG);
  var rows = readAll_(sh);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].key === key) { sh.getRange(i + 2, 2).setNumberFormat('@').setValue(String(value)); return; }
  }
  sh.getRange(sh.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[key, String(value)]]);
}
/** Kata semangat di beranda (diganti admin, boleh tiap hari). */
function actSaveSemangat_(u, p) {
  if (u.role !== 'admin') throw fail_('Hanya admin.');
  var t = String(p.teks || '').replace(/\s+/g, ' ').trim().substring(0, 300);
  setConfigKey_('semangat', t);
  setConfigKey_('semangatTgl', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'));
  return { ok: true };
}
function actSemangat_() { var c = getConfig_(); return { teks: c.semangat || '', tgl: c.semangatTgl || '' }; }
function actSaveConfig_(u, p) {
  if (u.role !== 'admin') throw fail_('Hanya admin.');
  var lok = [];
  var src = p.lokasi || [];
  for (var i = 0; i < src.length && i < 20; i++) {
    var lat = parseFloat(src[i].lat), lng = parseFloat(src[i].lng), rad = parseInt(src[i].radius, 10);
    if (isNaN(lat) || isNaN(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) continue;
    lok.push({ nama: String(src[i].nama || 'Lokasi ' + (i + 1)).substring(0, 60), lat: lat, lng: lng, radius: rad > 0 ? rad : DEFAULT_RADIUS });
  }
  setConfigKey_('lokasi', JSON.stringify(lok));
  var r = parseInt(p.retensiFotoHari, 10);
  if (r >= 7 && r <= 3650) setConfigKey_('retensiFotoHari', r);
  if (p.waAdmin !== undefined) setConfigKey_('waAdmin', normWa_(p.waAdmin));
  if (typeof p.waGatewayToken === 'string' && p.waGatewayToken.trim()) {
    if (p.waGatewayToken.trim() === '-') PropertiesService.getScriptProperties().deleteProperty('WA_GATEWAY_TOKEN');
    else PropertiesService.getScriptProperties().setProperty('WA_GATEWAY_TOKEN', p.waGatewayToken.trim());
  }
  return { ok: true };
}

/* ====================== ABSENSI ====================== */
function savePhoto_(dataUrl, key, name) {
  var m = /^data:image\/(jpeg|jpg|png|webp);base64,(.+)$/.exec(String(dataUrl || ''));
  if (!m) throw fail_('Foto wajib dilampirkan.');
  if (m[2].length > MAX_FOTO_B64) throw fail_('Ukuran foto terlalu besar.');
  var bytes = Utilities.base64Decode(m[2]);
  var blob = Utilities.newBlob(bytes, 'image/' + (m[1] === 'jpg' ? 'jpeg' : m[1]), name + '.jpg');
  return subFolder_(key).createFile(blob).getId();
}
function geoCheck_(p, cfg) {
  var lat = parseFloat(p.lat), lng = parseFloat(p.lng);
  if (isNaN(lat) || isNaN(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) return { flag: 'GPS_GAGAL', jarak: '', lat: '', lng: '' };
  if (!cfg.lokasi.length) return { flag: '', jarak: '', lat: lat, lng: lng };
  var best = null, bestRad = DEFAULT_RADIUS;
  for (var i = 0; i < cfg.lokasi.length; i++) {
    var L = cfg.lokasi[i];
    var d = haversine_(lat, lng, L.lat, L.lng);
    if (best === null || d - (L.radius || DEFAULT_RADIUS) < best - bestRad) { best = d; bestRad = L.radius || DEFAULT_RADIUS; }
  }
  return { flag: best > bestRad ? 'LOKASI_JAUH' : '', jarak: Math.round(best), lat: lat, lng: lng };
}
function addFlag_(flag, f) {
  if (!f) return flag;
  var a = flag ? flag.split(',') : [];
  if (a.indexOf(f) < 0) a.push(f);
  return a.join(',');
}
function removeFlag_(flag, f) {
  var a = flag ? flag.split(',') : [];
  var r = [];
  for (var i = 0; i < a.length; i++) if (a[i] !== f && a[i]) r.push(a[i]);
  return r.join(',');
}
function hasFlag_(flag, f) { return (',' + (flag || '') + ',').indexOf(',' + f + ',') >= 0; }
function jamBeda_(jam, serverTs) {
  var sm = minutesOf_(serverTs.substring(11, 16));
  var d = Math.abs(minutesOf_(jam) - sm);
  d = Math.min(d, 1440 - d);
  return d > JAM_BEDA_MENIT;
}
function validJam_(j) { return /^([01]\d|2[0-3]):[0-5]\d$/.test(j || ''); }
function validDate_(d) { return /^\d{4}-\d{2}-\d{2}$/.test(d || ''); }
function isOpen_(o) { return o.serverMasuk && !o.serverPulang && !hasFlag_(o.flag, 'LUPA_PULANG'); }

/** Cari absen terbuka milik user (periode sekarang + sebelumnya). Menutup otomatis yang sudah > 20 jam. */
function findOpen_(userId, nowTs, autoClose) {
  var today = nowTs.substring(0, 10);
  var k = periodKey_(today);
  var keys = [k, prevKey_(k)];
  var found = null;
  for (var ki = 0; ki < keys.length; ki++) {
    var sh = absenSheet_(keys[ki], false);
    if (!sh) continue;
    var tail = readTail_(sh, ABSEN_HDR, TAIL_ROWS);
    for (var i = tail.length - 1; i >= 0; i--) {
      var o = tail[i].o;
      if (o.userId !== userId || !isOpen_(o)) continue;
      var ageH = (tsMs_(nowTs) - tsMs_(o.serverMasuk)) / 3600000;
      if (ageH > 20) {
        if (autoClose) { o.flag = addFlag_(o.flag, 'LUPA_PULANG'); writeRow_(sh, ABSEN_HDR, tail[i].row, o); }
        continue;
      }
      if (!found) found = { sh: sh, row: tail[i].row, o: o, key: keys[ki] };
    }
    if (found) break;
  }
  return found;
}

function actMasuk_(u, p) {
  var id = String(p.id || '');
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(id)) throw fail_('ID absen tidak valid.');
  var shift = String(p.shift || '').toUpperCase();
  if (SHIFTS_PICK.indexOf(shift) < 0) throw fail_('Pilih shift (Pagi, Siang, Middle atau Malam).');
  if (!validJam_(p.jam)) throw fail_('Jam masuk tidak valid.');
  var now = nowJ_(), today = now.substring(0, 10), key = periodKey_(today);
  var sh = absenSheet_(key, true);
  // idempoten: id yang sama tidak disimpan dua kali
  var dup = findRowById_(sh, ABSEN_HDR, id);
  if (dup) return { ok: true, duplicate: true };
  var open = findOpen_(u.id, now, true);
  if (open) throw fail_('Anda belum absen Pulang untuk Masuk tanggal ' + open.o.tanggal + '. Absen Pulang dulu, atau minta atasan mengoreksi.');
  var cfg = getConfig_();
  var g = geoCheck_(p, cfg);
  var fotoId = savePhoto_(p.foto, key, today + '_' + u.id + '_masuk');
  var flag = '';
  flag = addFlag_(flag, g.flag);
  if (jamBeda_(p.jam, now)) flag = addFlag_(flag, 'JAM_BEDA');
  if (shift === 'L') flag = addFlag_(flag, 'JADWAL_L');
  appendObj_(sh, ABSEN_HDR, {
    id: id, tanggal: today, userId: u.id, nama: u.nama, unit: u.unit, shift: shift,
    jamMasuk: p.jam, jamPulang: '', tglPulang: '', fotoMasuk: fotoId, fotoPulang: '',
    latMasuk: g.lat, lngMasuk: g.lng, latPulang: '', lngPulang: '', jarakMasuk: g.jarak, jarakPulang: '',
    flag: flag, serverMasuk: now, serverPulang: '', catatan: '', pulangId: ''
  });
  return { ok: true };
}

function actPulang_(u, p) {
  var id = String(p.id || '');
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(id)) throw fail_('ID absen tidak valid.');
  if (!validJam_(p.jam)) throw fail_('Jam pulang tidak valid.');
  var now = nowJ_();
  // idempoten
  var today = now.substring(0, 10), k = periodKey_(today);
  var keys = [k, prevKey_(k)];
  for (var a = 0; a < keys.length; a++) {
    var s0 = absenSheet_(keys[a], false);
    if (!s0) continue;
    var tl = readTail_(s0, ABSEN_HDR, TAIL_ROWS);
    for (var b = 0; b < tl.length; b++) if (tl[b].o.pulangId === id) return { ok: true, duplicate: true };
  }
  var open = findOpen_(u.id, now, true);
  if (!open) throw fail_('Belum ada absen Masuk yang aktif.');
  var cfg = getConfig_();
  var g = geoCheck_(p, cfg);
  var fotoId = savePhoto_(p.foto, open.key, today + '_' + u.id + '_pulang');
  var o = open.o;
  o.jamPulang = p.jam; o.tglPulang = today; o.fotoPulang = fotoId;
  o.latPulang = g.lat; o.lngPulang = g.lng; o.jarakPulang = g.jarak;
  o.serverPulang = now; o.pulangId = id;
  o.flag = addFlag_(o.flag, g.flag);
  if (jamBeda_(p.jam, now)) o.flag = addFlag_(o.flag, 'JAM_BEDA');
  if (durMenit_(o) === null) o.flag = addFlag_(o.flag, 'JAM_BEDA');
  writeRow_(open.sh, ABSEN_HDR, open.row, o);
  return { ok: true };
}

function durMenit_(o) {
  if (!o.jamMasuk || !o.jamPulang || !o.tglPulang) return null;
  var d = utcMin_(o.tglPulang, o.jamPulang) - utcMin_(o.tanggal, o.jamMasuk);
  if (d < 0 || d > 36 * 60) return null;
  return d;
}
function pubRecord_(o) {
  return {
    id: o.id, tanggal: o.tanggal, userId: o.userId, nama: o.nama, unit: o.unit, shift: o.shift,
    jamMasuk: o.jamMasuk, jamPulang: o.jamPulang, tglPulang: o.tglPulang,
    flag: o.flag, catatan: o.catatan, durasi: durMenit_(o),
    jarakMasuk: o.jarakMasuk, jarakPulang: o.jarakPulang,
    latMasuk: o.latMasuk, lngMasuk: o.lngMasuk, latPulang: o.latPulang, lngPulang: o.lngPulang,
    fotoM: !!o.fotoMasuk && o.fotoMasuk !== 'dihapus', fotoP: !!o.fotoPulang && o.fotoPulang !== 'dihapus',
    fotoDihapus: o.fotoMasuk === 'dihapus' || o.fotoPulang === 'dihapus',
    open: isOpen_(o)
  };
}

/* ====================== LAPORAN HARIAN ====================== */
function actSaveLaporan_(u, p) {
  var tgl = String(p.tanggal || '');
  if (!validDate_(tgl)) throw fail_('Tanggal tidak valid.');
  var shift = String(p.shift || '').toUpperCase();
  if (SHIFTS.indexOf(shift) < 0) throw fail_('Pilih shift untuk laporan ini.');
  var today = nowJ_().substring(0, 10);
  if (tgl > today || tgl < addDays_(today, -LAPOR_HARI)) throw fail_('Laporan hanya bisa diisi/diubah untuk ' + LAPOR_HARI + ' hari terakhir.');
  var isi = String(p.isi || '').substring(0, 4000).trim();
  if (!isi) throw fail_('Isi laporan masih kosong.');
  var sh = laporSheet_(periodKey_(tgl), true);
  var id = u.id + '_' + tgl + '_' + shift;
  var f = findRowById_(sh, LAPOR_HDR, id);
  var o = { id: id, tanggal: tgl, userId: u.id, nama: u.nama, isi: isi, updatedAt: nowJ_(), shift: shift };
  if (f) writeRow_(sh, LAPOR_HDR, f.row, o); else appendObj_(sh, LAPOR_HDR, o);
  return { ok: true };
}
/** Peta laporan: kunci userId_tanggal_shift (shift kosong = laporan lama sebelum ada pembagian shift). */
function laporanMap_(key, userId) {
  var sh = laporSheet_(key, false);
  var m = {};
  if (!sh) return m;
  var rows = readAll_(sh);
  for (var i = 0; i < rows.length; i++) {
    if (!userId || rows[i].userId === userId) m[rows[i].userId + '_' + rows[i].tanggal + '_' + (rows[i].shift || '')] = rows[i].isi;
  }
  return m;
}

/* ====================== STATE / RIWAYAT / REKAP ====================== */
function actState_(u) {
  var now = nowJ_(), today = now.substring(0, 10);
  var key = periodKey_(today);
  var open = findOpen_(u.id, now, false);
  var recs = [];
  var sh = absenSheet_(key, false);
  if (sh) {
    var tail = readTail_(sh, ABSEN_HDR, TAIL_ROWS);
    for (var i = 0; i < tail.length; i++) {
      var o = tail[i].o;
      if (o.userId === u.id && (o.tanggal === today || o.tglPulang === today)) recs.push(pubRecord_(o));
    }
  }
  var yRec = open && open.o.tanggal !== today ? pubRecord_(open.o) : null;
  var lap = laporanMap_(key, u.id);
  if (open && open.key !== key) { var l2 = laporanMap_(open.key, u.id); for (var kk in l2) lap[kk] = l2[kk]; }
  var cfg = getConfig_();
  var lapAll = {};
  for (var lk in lap) { var pp = lk.substring(u.id.length + 1); var tt = pp.substring(0, 10), ss2 = pp.substring(11); lapAll[tt + '_' + ss2] = lap[lk]; }
  return {
    me: pubUserFull_(u), now: now, today: today, period: periodRange_(key),
    open: open ? pubRecord_(open.o) : null, todayRecords: recs, olderOpen: yRec,
    laporan: lapAll, lokasi: cfg.lokasi, defaultRadius: DEFAULT_RADIUS, units: getUnits_(),
    counts: u.role === 'admin' ? adminCounts_() : null
  };
}

function summarize_(rows, users) {
  var map = {}, order = [];
  function ensure(r) {
    if (!map[r.userId]) {
      map[r.userId] = { userId: r.userId, nama: r.nama, unit: r.unit, hari: 0, menit: 0, P: 0, S: 0, D: 0, M: 0, L: 0, cek: 0, belum: 0 };
      order.push(r.userId);
    }
    return map[r.userId];
  }
  if (users) {
    for (var i = 0; i < users.length; i++) ensure({ userId: users[i].id, nama: users[i].nama, unit: users[i].unit });
  }
  for (var j = 0; j < rows.length; j++) {
    var r = rows[j], s = ensure(r);
    if (r.shift === 'L') { s.L++; }
    else if (r.jamPulang) {
      s.hari++; if (r.durasi) s.menit += r.durasi;
      if (s[r.shift] !== undefined) s[r.shift]++;
    } else if (hasFlag_(r.flag, 'LUPA_PULANG') || r.open) { s.belum++; if (s[r.shift] !== undefined) s[r.shift]++; }
    var f = r.flag || '';
    for (var k = 0; k < FLAG_CEK.length; k++) {
      if (hasFlag_(f, FLAG_CEK[k]) && !hasFlag_(f, 'DICEK')) { s.cek++; break; }
    }
  }
  var res = [];
  for (var q = 0; q < order.length; q++) res.push(map[order[q]]);
  res.sort(function (a, b) { return a.nama < b.nama ? -1 : a.nama > b.nama ? 1 : 0; });
  return res;
}

function actRiwayat_(u, p) {
  var key = validKey_(p.period) ? p.period : periodKey_(nowJ_().substring(0, 10));
  var sh = absenSheet_(key, false);
  var rows = [];
  if (sh) {
    var all = readAll_(sh);
    for (var i = 0; i < all.length; i++) if (all[i].userId === u.id) rows.push(pubRecord_(all[i]));
  }
  rows.sort(function (a, b) { return a.tanggal < b.tanggal ? 1 : a.tanggal > b.tanggal ? -1 : (a.jamMasuk < b.jamMasuk ? 1 : -1); });
  var lap = laporanMap_(key, u.id);
  return { period: periodRange_(key), rows: rows, laporan: lap, summary: summarize_(rows, null)[0] || null };
}

function actRekap_(u, p) {
  if (!isManager_(u)) throw fail_('Tidak punya akses.');
  var key = validKey_(p.period) ? p.period : periodKey_(nowJ_().substring(0, 10));
  var users = readAll_(karSheet_()).filter(function (x) {
    return x.aktif === 'Y' && canView_(u);
  });
  var sh = absenSheet_(key, false);
  var rows = [];
  if (sh) {
    var all = readAll_(sh);
    for (var i = 0; i < all.length; i++) if (canView_(u)) rows.push(pubRecord_(all[i]));
  }
  rows.sort(function (a, b) { return a.tanggal < b.tanggal ? 1 : a.tanggal > b.tanggal ? -1 : 0; });
  var lap = laporanMap_(key, null);
  var lapOut = {};
  for (var k in lap) { lapOut[k] = lap[k]; }
  var units = {};
  for (var j = 0; j < users.length; j++) units[users[j].unit] = true;
  return { period: periodRange_(key), rows: rows, laporan: lapOut, summary: summarize_(rows, users), units: Object.keys(units).sort() };
}

function actKoreksi_(u, p) {
  if (!isManager_(u)) throw fail_('Tidak punya akses.');
  var tgl = String(p.tanggal || '');
  if (!validDate_(tgl)) throw fail_('Tanggal tidak valid.');
  var sh = absenSheet_(periodKey_(tgl), false);
  if (!sh) throw fail_('Data tidak ditemukan.');
  var f = findRowById_(sh, ABSEN_HDR, String(p.id || ''));
  if (!f) throw fail_('Data tidak ditemukan.');
  var o = f.o;
  if (!canSee_(u, o.unit)) throw fail_('Tidak punya akses ke unit ini.');
  if (p.tandaiDicek) o.flag = addFlag_(o.flag, 'DICEK');
  if (p.shift) {
    var s = String(p.shift).toUpperCase();
    if (SHIFTS.indexOf(s) < 0) throw fail_('Shift tidak valid.');
    o.shift = s;
  }
  if (p.jamMasuk) { if (!validJam_(p.jamMasuk)) throw fail_('Jam masuk tidak valid.'); o.jamMasuk = p.jamMasuk; }
  if (p.jamPulang) {
    if (!validJam_(p.jamPulang)) throw fail_('Jam pulang tidak valid.');
    var tp = String(p.tglPulang || o.tglPulang || o.tanggal);
    if (!validDate_(tp)) throw fail_('Tanggal pulang tidak valid.');
    o.jamPulang = p.jamPulang; o.tglPulang = tp;
    o.flag = removeFlag_(o.flag, 'LUPA_PULANG');
    if (!o.serverPulang) o.serverPulang = nowJ_();
  }
  o.catatan = String(p.catatan !== undefined ? p.catatan : o.catatan).substring(0, 300);
  o.flag = addFlag_(o.flag, 'KOREKSI');
  o.flag = addFlag_(o.flag, 'DICEK');
  writeRow_(sh, ABSEN_HDR, f.row, o);
  return { ok: true };
}

function actFoto_(u, p) {
  var tgl = String(p.tanggal || '');
  if (!validDate_(tgl)) throw fail_('Tanggal tidak valid.');
  var sh = absenSheet_(periodKey_(tgl), false);
  if (!sh) throw fail_('Data tidak ditemukan.');
  var f = findRowById_(sh, ABSEN_HDR, String(p.id || ''));
  if (!f) throw fail_('Data tidak ditemukan.');
  var o = f.o;
  if (o.userId !== u.id && !canView_(u)) throw fail_('Tidak punya akses.');
  var fid = p.which === 'pulang' ? o.fotoPulang : o.fotoMasuk;
  if (!fid) throw fail_('Foto tidak ada.');
  if (fid === 'dihapus') throw fail_('Foto sudah dihapus otomatis (lewat masa simpan).');
  var blob = DriveApp.getFileById(fid).getBlob();
  return { dataUrl: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
}

/* ====================== MANAJEMEN KARYAWAN ====================== */
function actKaryawanList_(u) {
  if (u.role !== 'admin') throw fail_('Hanya admin.');
  var rows = readAll_(karSheet_());
  return { rows: rows.map(function (r) { return { id: r.id, nip: r.nip, nama: r.nama, unit: r.unit, username: r.username, role: r.role, aktif: r.aktif, email: r.email || '', hp: r.hp || '' }; }) };
}
function actKaryawanSave_(u, p) {
  if (u.role !== 'admin') throw fail_('Hanya admin.');
  var d = p.data || {};
  var nama = String(d.nama || '').trim(), username = String(d.username || '').trim().toLowerCase();
  if (!nama || !/^[a-z0-9._-]{3,30}$/.test(username)) throw fail_('Nama wajib diisi dan username 3-30 karakter (huruf kecil/angka).');
  var role = ['admin', 'atasan', 'karyawan'].indexOf(d.role) >= 0 ? d.role : 'karyawan';
  var sh = karSheet_();
  var rows = readAll_(sh);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].username === username && rows[i].id !== d.id) throw fail_('Username sudah dipakai.');
  }
  if (d.id) {
    var f = findRowById_(sh, KARYAWAN_HDR, d.id);
    if (!f) throw fail_('Karyawan tidak ditemukan.');
    if (f.o.id === u.id && (d.aktif === 'N' || role !== 'admin')) throw fail_('Tidak dapat menonaktifkan/menurunkan akun Anda sendiri.');
    f.o.nip = String(d.nip || ''); f.o.nama = nama; f.o.unit = addUnit_(d.unit); f.o.username = username;
    f.o.role = role; f.o.aktif = d.aktif === 'N' ? 'N' : 'Y';
    if (d.email !== undefined) { if (!validEmail_(String(d.email).trim())) throw fail_('Format email tidak valid.'); f.o.email = String(d.email).trim().toLowerCase(); }
    if (d.hp !== undefined) f.o.hp = normWa_(d.hp);
    if (d.password) {
      if (String(d.password).length < 6) throw fail_('Password minimal 6 karakter.');
      var salt = randomSalt_(); f.o.salt = salt; f.o.hash = hash_(salt, String(d.password)); f.o.wajibGanti = 'Y';
    }
    writeRow_(sh, KARYAWAN_HDR, f.row, f.o);
    CacheService.getScriptCache().remove('U:' + f.o.id);
    return { ok: true };
  }
  var pw = String(d.password || d.nip || '123456');
  saveKaryawanRow_(sh, { nip: String(d.nip || ''), nama: nama, unit: addUnit_(d.unit), username: username, role: role, aktif: 'Y', email: validEmail_(String(d.email || '').trim()) ? String(d.email || '').trim().toLowerCase() : '', hp: normWa_(d.hp) }, pw, true);
  return { ok: true, password: pw };
}
/** Impor massal: baris {nip,nama,unit,username,role,password}. Password awal default = NIP. */
function actKaryawanImport_(u, p) {
  if (u.role !== 'admin') throw fail_('Hanya admin.');
  var list = p.rows || [];
  if (list.length > 500) throw fail_('Maksimal 500 baris per impor.');
  var sh = karSheet_();
  var exist = {};
  var rows = readAll_(sh);
  for (var i = 0; i < rows.length; i++) exist[rows[i].username] = true;
  var made = [], skipped = [];
  for (var j = 0; j < list.length; j++) {
    var d = list[j];
    var nama = String(d.nama || '').trim(), un = String(d.username || '').trim().toLowerCase();
    if (!nama || !/^[a-z0-9._-]{3,30}$/.test(un) || exist[un]) { skipped.push(un || nama || ('baris ' + (j + 1))); continue; }
    var pw = String(d.password || d.nip || '123456');
    var role = ['admin', 'atasan', 'karyawan'].indexOf(d.role) >= 0 ? d.role : 'karyawan';
    saveKaryawanRow_(sh, { nip: String(d.nip || ''), nama: nama, unit: addUnit_(d.unit), username: un, role: role, aktif: 'Y', email: validEmail_(String(d.email || '').trim()) ? String(d.email || '').trim().toLowerCase() : '', hp: normWa_(d.hp) }, pw, true);
    exist[un] = true; made.push(un);
  }
  return { ok: true, dibuat: made.length, dilewati: skipped };
}

/* ====================== RETENSI FOTO ====================== */
function hapusFotoLama() {
  var cfg = getConfig_();
  var cutoff = addDays_(nowJ_().substring(0, 10), -cfg.retensiFotoHari);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var t0 = Date.now(), n = 0;
  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    if (sh.getName().indexOf(PFX_ABSEN) !== 0) continue;
    var last = sh.getLastRow();
    if (last < 2) continue;
    var vals = sh.getRange(2, 1, last - 1, ABSEN_HDR.length).getValues();
    var cm = ABSEN_HDR.indexOf('fotoMasuk'), cp = ABSEN_HDR.indexOf('fotoPulang'), ct = ABSEN_HDR.indexOf('tanggal');
    for (var i = 0; i < vals.length; i++) {
      if (Date.now() - t0 > 4.5 * 60 * 1000) { Logger.log('Waktu habis, dilanjutkan besok. Dihapus: ' + n); return; }
      var tgl = str_(vals[i][ct]);
      if (!tgl || tgl >= cutoff) continue;
      var changed = false;
      [cm, cp].forEach(function (c) {
        var fid = str_(vals[i][c]);
        if (fid && fid !== 'dihapus') {
          try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) { /* sudah hilang */ }
          sh.getRange(i + 2, c + 1).setNumberFormat('@').setValue('dihapus');
          changed = true; n++;
        }
      });
    }
  }
  Logger.log('Foto dihapus: ' + n + ' (lebih tua dari ' + cutoff + ')');
}

/* ====================== ROUTER ====================== */
function doGet(e) {
  try {
    var p = e.parameter || {};
    var a = p.action;
    if (!a) return out_({ ok: true, app: 'absensi-rsi', time: nowJ_() });
    if (a === 'ping') return out_({ ok: true, time: nowJ_() });
    if (a === 'login') return out_(actLogin_(p));
    if (a === 'cekUsername') return out_(actCekUsername_(p));
    if (a === 'cekDaftar') return out_(actCekDaftar_(p));
    if (a === 'infoPublik') return out_(actInfoPublik_());
    if (a === 'lupaPassword') {
      var lk0 = LockService.getScriptLock(); lk0.waitLock(20000);
      try { return out_(actLupa_(p)); } finally { try { lk0.releaseLock(); } catch (e0) { /* tidak terkunci */ } }
    }
    var u = auth_(p.token);
    switch (a) {
      case 'state': return out_(actState_(u));
      case 'riwayat': return out_(actRiwayat_(u, p));
      case 'rekap': return out_(actRekap_(u, p));
      case 'foto': return out_(actFoto_(u, p));
      case 'profilFoto': return out_(actProfilFoto_(u, p));
      case 'fotoBanyak': return out_(actFotoBanyak_(u, p));
      case 'changePassword': return out_(actChangePassword_(u, p));
      case 'karyawanList': return out_(actKaryawanList_(u));
      case 'pendingList': return out_(actPendingList_(u));
      case 'resetList': return out_(actResetList_(u));
      case 'resetBuat':
        var lk = LockService.getScriptLock(); lk.waitLock(20000);
        try { return out_(actResetBuat_(u, p)); } finally { try { lk.releaseLock(); } catch (e3) { /* tidak terkunci */ } }
      case 'config': { var cfgOut = getConfig_(); cfgOut.waGatewaySet = !!waToken_(); return out_(cfgOut); }
      case 'jaga': return out_(actJaga_(u));
      case 'semangat': return out_(actSemangat_());
      default: return out_({ error: 'Aksi tidak dikenal: ' + a });
    }
  } catch (err) {
    return out_({ error: String(err && err.message ? err.message : err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var p = JSON.parse(e.postData.contents);
    var r;
    if (p.action === 'signup' || p.action === 'lupaPassword') {   // publik (tanpa login)
      lock.waitLock(25000);
      r = p.action === 'signup' ? actSignup_(p) : actLupa_(p);
      return out_(r);
    }
    var u = auth_(p.token);
    lock.waitLock(25000);
    switch (p.action) {
      case 'masuk': r = actMasuk_(u, p); CacheService.getScriptCache().remove('JAGA'); break;
      case 'pulang': r = actPulang_(u, p); CacheService.getScriptCache().remove('JAGA'); break;
      case 'saveLaporan': r = actSaveLaporan_(u, p); break;
      case 'koreksi': r = actKoreksi_(u, p); CacheService.getScriptCache().remove('JAGA'); break;
      case 'saveConfig': r = actSaveConfig_(u, p); break;
      case 'saveSemangat': r = actSaveSemangat_(u, p); break;
      case 'karyawanSave': r = actKaryawanSave_(u, p); break;
      case 'karyawanImport': r = actKaryawanImport_(u, p); break;
      case 'karyawanApprove': r = actApprove_(u, p); break;
      case 'saveFotoProfil': r = actSaveFotoProfil_(u, p); break;
      case 'saveProfil': r = actSaveProfil_(u, p); break;
      case 'unitSave': r = actUnitSave_(u, p); break;
      case 'unitHapus': r = actUnitHapus_(u, p); break;
      default: r = { error: 'Aksi tidak dikenal: ' + p.action };
    }
    return out_(r);
  } catch (err) {
    return out_({ error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) { /* tidak terkunci */ }
  }
}
