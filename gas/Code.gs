/**
 * PANTAU SARANA HCGA - backend Google Apps Script + Google Sheets
 *
 * Pasang:
 *  1. Buat Google Sheet baru > Extensions > Apps Script > tempel file ini (Code.gs)
 *     dan isi appsscript.json (Project Settings > "Show appsscript.json").
 *  2. Jalankan fungsi setupSheets() satu kali (beri izin akses).
 *  3. Deploy > New deployment > Web app
 *       Execute as : Me
 *       Who has access : Anyone
 *     Salin URL berakhiran /exec, tempel ke API_URL di index.html (atau di layar login).
 *  4. (Disarankan) Trigger waktu harian untuk arsipkanDataLama().
 *
 * Catatan: simpan SHEET_ID di Script Properties bila script TIDAK terikat ke sheet-nya.
 */

const SH = {
  Drivers: ['id', 'nama', 'wa', 'pin', 'kendaraan', 'plat', 'aktif', 'duty', 'sos', 'gpsKey', 'telegram'],
  Pengawas: ['username', 'nama', 'wa', 'pin'],
  Orders: ['id', 'jenis', 'judul', 'pemohon', 'dept', 'jemputNama', 'jemputLat', 'jemputLng', 'tujuanNama', 'tujuanLat', 'tujuanLng',
    'penumpang', 'catatan', 'prioritas', 'driverId', 'status', 'dibuat', 'ditugaskan', 'diterima', 'tiba', 'berangkat', 'selesai',
    'jarakKm', 'catatanSelesai', 'alasanTolak'],
  Chat: ['id', 'driverId', 'orderId', 'dari', 'teks', 'ts', 'dibaca'],
  Track: ['ts', 'driverId', 'lat', 'lng', 'speed', 'orderId'],
  Titik: ['nama', 'lat', 'lng', 'kategori'],
  Log: ['ts', 'aktor', 'aksi', 'detail']
};
const TIME_COLS = {
  Orders: ['dibuat', 'ditugaskan', 'diterima', 'tiba', 'berangkat', 'selesai'],
  Chat: ['ts'], Track: ['ts'], Log: ['ts'], Drivers: ['sos']
};
const READ_LIMIT = { Drivers: 0, Pengawas: 0, Titik: 0, Orders: 300, Chat: 400 };   // 0 = semua baris
const ACTIVE = ['DITUGASKAN', 'DITERIMA', 'TIBA_JEMPUT', 'DALAM_PERJALANAN'];
const INPROG = ['DITERIMA', 'TIBA_JEMPUT', 'DALAM_PERJALANAN'];
const CFG = {
  TOKEN_TTL_H: 24,            // masa berlaku sesi login
  LIVE_TTL_S: 21600,          // posisi terakhir disimpan 6 jam (batas CacheService)
  TRACK_MIN_MS: 30000,        // simpan titik jejak minimal tiap 30 detik...
  TRACK_MIN_M: 50,            // ...atau tiap bergeser 50 m
  STILL_RADIUS_M: 30, STILL_SPEED: 3,
  TABLE_TTL_S: 20,
  TRACK_KEEP_DAYS: 14, CHAT_KEEP_DAYS: 90
};

/* ============ pintu masuk ============ */
function doGet(e) {
  const p = (e && e.parameter) || {};
  /* Traccar Client (iPhone) mengirim posisi lewat GET: ?id=D03-kunci&lat=..&lon=..&speed=..&bearing=..&accuracy=..&batt=.. */
  if (p.id && p.lat && p.lon) return ContentService.createTextOutput(traccarPing_(p));
  return out_({ ok: true, data: { app: 'Pantau Sarana HCGA', time: Date.now() } });
}

function doPost(e) {
  let req;
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return out_({ ok: false, error: 'Format permintaan salah' }); }
  if (req && req.update_id != null) return out_(handleTelegram_(req));      // webhook bot Telegram
  try {
    const fn = ACTIONS[req.action];
    if (!fn) throw new Error('Aksi tidak dikenal: ' + req.action);
    const data = req.action === 'login' ? fn(req) : fn(req, verify_(req.token));
    return out_({ ok: true, data: data });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err), code: err && err.code || '' });
  }
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ============ akses sheet ============ */
let _ss = null;
function ss_() {
  if (_ss) return _ss;
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  _ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!_ss) throw new Error('Spreadsheet tidak ditemukan. Isi Script Property SHEET_ID.');
  return _ss;
}
function sheet_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('Sheet "' + name + '" belum ada. Jalankan setupSheets().');
  return s;
}
function readRows_(name, lastN) {
  const sh = sheet_(name), cols = SH[name], last = sh.getLastRow();
  if (last < 2) return [];
  const first = lastN ? Math.max(2, last - lastN + 1) : 2;
  const vals = sh.getRange(first, 1, last - first + 1, cols.length).getValues(), tc = TIME_COLS[name] || [];
  return vals.map(function (r, i) {
    const o = { _row: first + i };
    cols.forEach(function (c, j) {
      let v = r[j];
      if (tc.indexOf(c) >= 0) v = v instanceof Date ? v.getTime() : (Number(v) || 0);
      o[c] = v;
    });
    return o;
  });
}
function toRow_(name, obj) {
  const tc = TIME_COLS[name] || [];
  return SH[name].map(function (c) {
    let v = obj[c]; if (v === undefined || v === null) v = '';
    if (tc.indexOf(c) >= 0) v = v ? new Date(Number(v)) : '';
    return v;
  });
}
function appendRow_(name, obj) { sheet_(name).appendRow(toRow_(name, obj)); }
function updateRow_(name, obj) { sheet_(name).getRange(obj._row, 1, 1, SH[name].length).setValues([toRow_(name, obj)]); }

/* cache tabel kecil agar polling pengawas/driver tidak membaca sheet tiap detik */
function table_(name) {
  const c = CacheService.getScriptCache(), k = 't:' + name, hit = c.get(k);
  if (hit) { try { return JSON.parse(hit); } catch (e) { /* baca ulang */ } }
  const rows = readRows_(name, READ_LIMIT[name]);
  const s = JSON.stringify(rows);
  if (s.length < 90000) c.put(k, s, CFG.TABLE_TTL_S);
  return rows;
}
function bust_() { const c = CacheService.getScriptCache(); Array.prototype.slice.call(arguments).forEach(function (n) { c.remove('t:' + n); }); }
function withLock_(fn) {
  const l = LockService.getScriptLock(); l.waitLock(15000);
  try { return fn(); } finally { l.releaseLock(); }
}

/* ============ util ============ */
function normWA_(s) { let d = String(s || '').replace(/\D/g, ''); if (d.indexOf('0') === 0) d = '62' + d.slice(1); else if (d.indexOf('8') === 0) d = '62' + d; return d; }
function isTrue_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function isActive_(d) { return !(d.aktif === false || String(d.aktif).toUpperCase() === 'FALSE'); }
function kmBetween_(a, b) {
  const R = 6371, r = Math.PI / 180, dLat = (b.lat - a.lat) * r, dLng = (b.lng - a.lng) * r;
  const x = Math.pow(Math.sin(dLat / 2), 2) + Math.cos(a.lat * r) * Math.cos(b.lat * r) * Math.pow(Math.sin(dLng / 2), 2);
  return 2 * R * Math.asin(Math.sqrt(x));
}
function startOfToday_() { const d = new Date(); d.setHours(0, 0, 0, 0); return d.getTime(); }
function strip_(o) { const c = {}; Object.keys(o).forEach(function (k) { if (k !== '_row') c[k] = o[k]; }); return c; }
function log_(who, aksi, detail) { try { appendRow_('Log', { ts: Date.now(), aktor: who.role + ':' + who.id, aksi: aksi, detail: detail || '' }); } catch (e) { /* log tidak boleh menggagalkan aksi */ } }
function fail_(msg, code) { const e = new Error(msg); if (code) e.code = code; return e; }

/* ============ autentikasi (token HMAC tanpa status) ============ */
function secret_() {
  const p = PropertiesService.getScriptProperties(); let s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
function sign_(str) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(str, secret_())); }
function makeToken_(role, id) {
  const body = role + '|' + id + '|' + (Date.now() + CFG.TOKEN_TTL_H * 3600000);
  return Utilities.base64EncodeWebSafe(body) + '.' + sign_(body);
}
function verify_(tok) {
  if (!tok) throw fail_('Silakan masuk kembali', 'AUTH');
  const parts = String(tok).split('.'); let body;
  try { body = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString(); } catch (e) { throw fail_('Sesi tidak valid', 'AUTH'); }
  if (parts.length !== 2 || sign_(body) !== parts[1]) throw fail_('Sesi tidak valid', 'AUTH');
  const f = body.split('|');
  if (Number(f[2]) < Date.now()) throw fail_('Sesi berakhir', 'AUTH');
  return { role: f[0], id: f[1] };
}
function need_(who, role) { if (who.role !== role) throw fail_('Akses ditolak'); }

/* ============ driver & posisi ============ */
function driverById_(id) { const d = table_('Drivers').filter(function (x) { return x.id === id && isActive_(x); })[0]; if (!d) throw fail_('Driver tidak ditemukan'); return d; }
function getLive_(ids) {
  const c = CacheService.getScriptCache(), got = c.getAll(ids.map(function (i) { return 'live:' + i; })), out = {};
  ids.forEach(function (i) { const v = got['live:' + i]; if (v) { try { out[i] = JSON.parse(v); } catch (e) { /* abaikan */ } } });
  return out;
}
function pubDriver_(d, L) {
  return {
    id: d.id, nama: d.nama, wa: String(d.wa || ''), kendaraan: d.kendaraan, plat: d.plat, duty: d.duty || 'OFF', sos: Number(d.sos) || 0,
    live: L ? { lat: L.lat, lng: L.lng, acc: L.acc, speed: L.speed, heading: L.heading, battery: L.battery, ts: L.ts, stillSince: L.stillSince } : null
  };
}
function activeOrderOf_(driverId, orders) {
  const mine = orders.filter(function (o) { return o.driverId === driverId && ACTIVE.indexOf(o.status) >= 0; });
  return mine.filter(function (o) { return INPROG.indexOf(o.status) >= 0; })[0] || mine[0] || null;
}
function pois_() { return table_('Titik').filter(function (p) { return p.nama && Number(p.lat); }).map(function (p) { return { nama: p.nama, lat: Number(p.lat), lng: Number(p.lng), kategori: p.kategori }; }); }
function recentOrders_() {
  const t0 = startOfToday_();
  return table_('Orders').filter(function (o) { return ['SELESAI', 'DIBATALKAN'].indexOf(o.status) < 0 || o.selesai >= t0; }).map(strip_).slice(-60);
}

/* ============ aksi ============ */
const ACTIONS = {
  login: function (req) {
    const id = String(req.id || '').trim(), pin = String(req.pin || '').trim(), role = req.role;
    const cache = CacheService.getScriptCache(), fk = 'fail:' + role + ':' + id.toLowerCase(), fails = Number(cache.get(fk) || 0);
    if (fails >= 5) throw fail_('Terlalu banyak percobaan. Coba lagi 10 menit lagi.');
    let rec = null, user = null, rid = '';
    if (role === 'pengawas') {
      rec = table_('Pengawas').filter(function (x) { return String(x.username).toLowerCase() === id.toLowerCase(); })[0];
      if (rec) { rid = rec.username; user = { id: rec.username, nama: rec.nama, wa: String(rec.wa || '') }; }
    } else if (role === 'driver') {
      rec = table_('Drivers').filter(function (x) { return normWA_(x.wa) === normWA_(id) && isActive_(x); })[0];
      if (rec) { rid = rec.id; user = { id: rec.id, nama: rec.nama, wa: String(rec.wa || '') }; }
    } else throw new Error('Peran tidak dikenal');
    if (!rec || String(rec.pin) !== pin) {
      cache.put(fk, String(fails + 1), 600);
      throw new Error(role === 'pengawas' ? 'Username atau PIN salah' : 'Nomor WA atau PIN salah');
    }
    cache.remove(fk);
    return { token: makeToken_(role, rid), role: role, user: user };
  },

  snapshot: function (req, who) {
    const now = Date.now();
    if (who.role === 'pengawas') {
      const drivers = table_('Drivers').filter(isActive_), live = getLive_(drivers.map(function (d) { return d.id; })), un = {};
      table_('Chat').forEach(function (m) { if (m.dari === 'driver' && !isTrue_(m.dibaca)) un[m.driverId] = (un[m.driverId] || 0) + 1; });
      return { now: now, drivers: drivers.map(function (d) { return pubDriver_(d, live[d.id]); }), orders: recentOrders_(), unread: un, pois: pois_() };
    }
    need_(who, 'driver');
    const d = driverById_(who.id), t0 = startOfToday_(), sv = table_('Pengawas')[0] || {};
    const mine = table_('Orders').filter(function (o) { return o.driverId === d.id && (ACTIVE.indexOf(o.status) >= 0 || (o.status === 'SELESAI' && o.selesai >= t0)); }).map(strip_);
    const unread = table_('Chat').filter(function (m) { return m.driverId === d.id && m.dari !== 'driver' && !isTrue_(m.dibaca); }).length;
    const me = pubDriver_(d, getLive_([d.id])[d.id]), bot = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_USERNAME');
    me.gpsId = d.gpsKey ? d.id + '-' + d.gpsKey : '';                   // hanya untuk driver itu sendiri (dipakai di Traccar Client)
    me.tgOk = !!d.telegram; me.tgLink = bot && me.gpsId ? 'https://t.me/' + String(bot).replace(/^@/, '') + '?start=' + me.gpsId : '';
    return { now: now, me: me, orders: mine, unread: unread, supervisor: { nama: sv.nama || 'Pengawas', wa: String(sv.wa || '') }, pois: pois_() };
  },

  ping: function (req, who) {
    need_(who, 'driver');
    const pts = Array.isArray(req.points) ? req.points : [];
    return applyPing_(who.id, pts[pts.length - 1]);
  },

  setDuty: function (req, who) {
    need_(who, 'driver');
    const s = req.state;
    if (['ON', 'BREAK', 'OFF'].indexOf(s) < 0) throw new Error('Status tidak valid');
    return withLock_(function () {
      const d = readRows_('Drivers', 0).filter(function (x) { return x.id === who.id; })[0];
      if (!d) throw new Error('Driver tidak ditemukan');
      const o = activeOrderOf_(d.id, readRows_('Orders', READ_LIMIT.Orders));
      if (s === 'OFF' && o) throw new Error('Selesaikan atau tolak order ' + o.id + ' dulu');
      if (s === 'BREAK' && o && INPROG.indexOf(o.status) >= 0) throw new Error('Tidak bisa istirahat saat order berjalan');
      d.duty = s; updateRow_('Drivers', d); bust_('Drivers');
      log_(who, 'duty', s);
      return { duty: s };
    });
  },

  createOrder: function (req, who) {
    need_(who, 'pengawas');
    const o = req.order || {};
    if (!o.jemputNama || !o.tujuanNama || !Number(o.jemputLat) || !Number(o.tujuanLat)) throw new Error('Titik jemput dan tujuan wajib diisi');
    return withLock_(function () {
      const n = {
        id: newOrderId_(), jenis: String(o.jenis || 'Lainnya').slice(0, 60), judul: String(o.judul || '').slice(0, 120), pemohon: String(o.pemohon || '').slice(0, 80), dept: String(o.dept || '').slice(0, 20),
        jemputNama: String(o.jemputNama).slice(0, 120), jemputLat: Number(o.jemputLat), jemputLng: Number(o.jemputLng),
        tujuanNama: String(o.tujuanNama).slice(0, 120), tujuanLat: Number(o.tujuanLat), tujuanLng: Number(o.tujuanLng),
        penumpang: String(o.penumpang || '').slice(0, 120), catatan: String(o.catatan || '').slice(0, 300), prioritas: o.prioritas === 'Mendesak' ? 'Mendesak' : 'Normal',
        driverId: '', status: 'BARU', dibuat: Date.now(), ditugaskan: 0, diterima: 0, tiba: 0, berangkat: 0, selesai: 0, jarakKm: 0, catatanSelesai: '', alasanTolak: ''
      };
      appendRow_('Orders', n); bust_('Orders'); log_(who, 'createOrder', n.id);
      return n;
    });
  },

  assignOrder: function (req, who) {
    need_(who, 'pengawas');
    return withLock_(function () {
      const orders = readRows_('Orders', READ_LIMIT.Orders), o = orderIn_(orders, req.orderId);
      const d = readRows_('Drivers', 0).filter(function (x) { return x.id === req.driverId && isActive_(x); })[0];
      if (!d) throw new Error('Driver tidak ditemukan');
      if (['BARU', 'DITUGASKAN'].indexOf(o.status) < 0) throw new Error('Order sudah berjalan, tidak bisa diganti driver');
      if (!d.duty || d.duty === 'OFF') throw new Error(d.nama + ' sedang tidak bertugas');
      const cur = activeOrderOf_(d.id, orders); if (cur && cur.id !== o.id) throw new Error(d.nama + ' masih memegang order ' + cur.id);
      o.driverId = d.id; o.status = 'DITUGASKAN'; o.ditugaskan = Date.now();
      updateRow_('Orders', o); bust_('Orders');
      const teks = 'Order baru: ' + o.jenis + '\n' + o.jemputNama + ' ke ' + o.tujuanNama + (o.prioritas === 'Mendesak' ? '\nPRIORITAS MENDESAK' : '');
      addMsg_(d.id, 'sistem', teks, o.id);
      notifyDriver_(d, 'Order ' + o.id + ' (' + o.jenis + ')\nJemput: ' + o.jemputNama + '\nTujuan: ' + o.tujuanNama + (o.penumpang ? '\nPenumpang/muatan: ' + o.penumpang : '') + (o.catatan ? '\nCatatan: ' + o.catatan : '') + '\nBuka aplikasi Pantau Sarana untuk menerima.');
      log_(who, 'assign', o.id + ' -> ' + d.id);
      return { ok: 1 };
    });
  },

  cancelOrder: function (req, who) {
    need_(who, 'pengawas');
    return withLock_(function () {
      const o = orderIn_(readRows_('Orders', READ_LIMIT.Orders), req.orderId);
      if (['BARU'].concat(ACTIVE).indexOf(o.status) < 0) throw new Error('Order sudah selesai');
      if (o.driverId) addMsg_(o.driverId, 'sistem', 'Order ' + o.id + ' dibatalkan pengawas', o.id);
      o.status = 'DIBATALKAN'; o.selesai = Date.now(); updateRow_('Orders', o); bust_('Orders'); log_(who, 'cancel', o.id);
      return { ok: 1 };
    });
  },

  respondOrder: function (req, who) {
    need_(who, 'driver');
    return withLock_(function () {
      const o = orderIn_(readRows_('Orders', READ_LIMIT.Orders), req.orderId);
      if (o.driverId !== who.id || o.status !== 'DITUGASKAN') throw new Error('Order ini tidak menunggu konfirmasi Anda');
      if (req.accept) { o.status = 'DITERIMA'; o.diterima = Date.now(); CacheService.getScriptCache().put('odo:' + o.id, '0', CFG.LIVE_TTL_S); }
      else {
        o.alasanTolak = String(req.reason || 'Tanpa alasan').slice(0, 120); addMsg_(who.id, 'sistem', 'Order ' + o.id + ' ditolak: ' + o.alasanTolak, o.id);
        o.status = 'BARU'; o.driverId = ''; o.ditugaskan = 0;
      }
      updateRow_('Orders', o); bust_('Orders'); log_(who, req.accept ? 'accept' : 'reject', o.id);
      return { ok: 1 };
    });
  },

  advanceOrder: function (req, who) {
    need_(who, 'driver');
    return withLock_(function () {
      const o = orderIn_(readRows_('Orders', READ_LIMIT.Orders), req.orderId), now = Date.now(), step = req.step;
      if (o.driverId !== who.id) throw new Error('Order ini bukan tugas Anda');
      if (step === 'arrive' && o.status === 'DITERIMA') { o.status = 'TIBA_JEMPUT'; o.tiba = now; }
      else if (step === 'start' && o.status === 'TIBA_JEMPUT') { o.status = 'DALAM_PERJALANAN'; o.berangkat = now; }
      else if (step === 'finish' && o.status === 'DALAM_PERJALANAN') {
        const odo = Number(CacheService.getScriptCache().get('odo:' + o.id)) || 0;
        o.status = 'SELESAI'; o.selesai = now; o.catatanSelesai = String(req.note || '').slice(0, 300); o.jarakKm = Math.round(odo * 10) / 10;
      } else throw new Error('Langkah tidak valid untuk status ' + o.status);
      updateRow_('Orders', o); bust_('Orders'); log_(who, step, o.id);
      return { ok: 1 };
    });
  },

  chat: function (req, who) {
    const id = who.role === 'driver' ? who.id : String(req.driverId || '');
    const since = Number(req.since) || 0;
    let l = table_('Chat').filter(function (m) { return m.driverId === id; });
    l = since ? l.filter(function (m) { return m.ts > since; }) : l.slice(-80);
    return l.map(function (m) { const c = strip_(m); c.dibaca = isTrue_(m.dibaca); return c; });
  },

  sendMsg: function (req, who) {
    const id = who.role === 'driver' ? who.id : String(req.driverId || ''), teks = String(req.text || '').trim().slice(0, 500);
    if (!teks) throw new Error('Pesan kosong');
    const d = driverById_(id);
    const m = addMsg_(id, who.role, teks, String(req.orderId || ''));
    if (who.role === 'pengawas') notifyTG_(d.telegram, 'Pesan dari pengawas:\n' + teks);
    return strip_(m);
  },

  markRead: function (req, who) {
    const id = who.role === 'driver' ? who.id : String(req.driverId || '');
    return withLock_(function () {
      const sh = sheet_('Chat'), col = SH.Chat.indexOf('dibaca') + 1;
      readRows_('Chat', READ_LIMIT.Chat).forEach(function (m) {
        const mine = who.role === 'pengawas' ? m.dari === 'driver' : m.dari !== 'driver';
        if (m.driverId === id && mine && !isTrue_(m.dibaca)) sh.getRange(m._row, col).setValue(true);
      });
      bust_('Chat'); return { ok: 1 };
    });
  },

  sos: function (req, who) {
    need_(who, 'driver');
    return withLock_(function () {
      const d = readRows_('Drivers', 0).filter(function (x) { return x.id === who.id; })[0];
      d.sos = Date.now(); updateRow_('Drivers', d); bust_('Drivers');
      addMsg_(d.id, 'sistem', d.nama + ' menekan tombol SOS', '');
      log_(who, 'SOS', '');
      return { ok: 1 };
    });
  },

  ackSos: function (req, who) {
    need_(who, 'pengawas');
    return withLock_(function () {
      const d = readRows_('Drivers', 0).filter(function (x) { return x.id === req.driverId; })[0];
      if (!d) throw new Error('Driver tidak ditemukan');
      d.sos = 0; updateRow_('Drivers', d); bust_('Drivers'); log_(who, 'ackSOS', d.id);
      return { ok: 1 };
    });
  },

  track: function (req, who) {
    need_(who, 'pengawas');
    const from = Math.max(Number(req.from) || 0, startOfToday_());
    return readRows_('Track', 3000).filter(function (p) { return p.driverId === req.driverId && p.ts > from; })
      .slice(-600).map(function (p) { return { lat: Number(p.lat), lng: Number(p.lng), ts: p.ts }; });
  }
};

/* simpan satu titik posisi driver (dipakai aksi ping dari aplikasi dan Traccar Client) */
function applyPing_(driverId, p) {
  if (!p || !isFinite(p.lat) || !isFinite(p.lng) || Math.abs(p.lat) > 90 || Math.abs(p.lng) > 180) throw new Error('Data posisi tidak valid');
  const d = driverById_(driverId);
  if (!d.duty || d.duty === 'OFF') return { ignored: true };
  const now = Date.now(), c = CacheService.getScriptCache(), prev = getLive_([d.id])[d.id], spd = Math.max(0, Number(p.speed) || 0);
  const pos = { lat: Number(p.lat), lng: Number(p.lng) };
  const near = !!(prev && prev.anchor && kmBetween_(prev.anchor, pos) * 1000 < CFG.STILL_RADIUS_M && spd < CFG.STILL_SPEED);
  const anchor = near ? prev.anchor : { lat: pos.lat, lng: pos.lng, ts: now };
  const L = {
    lat: pos.lat, lng: pos.lng, acc: Math.round(Number(p.acc) || 0), speed: spd, heading: Number(p.heading) || (prev ? prev.heading : 0),
    battery: p.battery == null ? (prev ? prev.battery : null) : Number(p.battery), ts: now, anchor: anchor, stillSince: near ? anchor.ts : 0, tr: prev && prev.tr
  };
  const ord = activeOrderOf_(d.id, table_('Orders')), busy = ord && INPROG.indexOf(ord.status) >= 0;
  if (busy && prev && L.acc < 50) {                         // odometer per order
    const dd = kmBetween_(prev, pos);
    if (dd > 0.005 && dd < 2) { const ok = 'odo:' + ord.id; c.put(ok, String((Number(c.get(ok)) || 0) + dd), CFG.LIVE_TTL_S); }
  }
  const tr = L.tr;                                          // jejak: disampel agar sheet tidak membengkak
  if (!tr || now - tr.ts >= CFG.TRACK_MIN_MS || kmBetween_(tr, pos) * 1000 >= CFG.TRACK_MIN_M) {
    appendRow_('Track', { ts: now, driverId: d.id, lat: pos.lat, lng: pos.lng, speed: spd, orderId: busy ? ord.id : '' });
    L.tr = { lat: pos.lat, lng: pos.lng, ts: now };
  }
  c.put('live:' + d.id, JSON.stringify(L), CFG.LIVE_TTL_S);
  return { ok: 1, ts: now };
}

function orderIn_(orders, id) { const o = orders.filter(function (x) { return x.id === id; })[0]; if (!o) throw new Error('Order tidak ditemukan'); return o; }
function addMsg_(driverId, dari, teks, orderId) {
  const m = { id: 'M' + Date.now() + Math.floor(Math.random() * 1000), driverId: driverId, orderId: orderId || '', dari: dari, teks: teks, ts: Date.now(), dibaca: false };
  appendRow_('Chat', m); bust_('Chat'); return m;
}
function newOrderId_() {
  const p = PropertiesService.getScriptProperties(), day = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMdd');
  const cur = (p.getProperty('ORD_SEQ') || '').split(':'), n = (cur[0] === day ? Number(cur[1]) : 0) + 1;
  p.setProperty('ORD_SEQ', day + ':' + n);
  return 'ORD-' + day + '-' + ('00' + n).slice(-3);
}

/* ============ iPhone: Traccar Client + Telegram ============ */
/* Satuan kecepatan dari Traccar Client: bawaan knot. Bila angka kecepatan tampak keliru (misal separuh dari speedometer),
   buat Script Property TRACCAR_SPEED berisi "ms" (meter per detik) atau "kmh". */
function traccarSpeedFactor_() {
  const u = String(PropertiesService.getScriptProperties().getProperty('TRACCAR_SPEED') || 'knot').toLowerCase();
  return u === 'ms' ? 3.6 : u === 'kmh' ? 1 : 1.852;
}
/* Traccar Client memakai OsmAnd protocol: kecepatan dalam knot. Pengenal perangkat = "<idDriver>-<gpsKey>" (lihat sheet Drivers). */
function traccarPing_(p) {
  try {
    const m = /^(.+)-([A-Za-z0-9]{6,})$/.exec(String(p.id));
    const d = m && table_('Drivers').filter(function (x) { return x.id === m[1] && isActive_(x); })[0];
    if (!d || !d.gpsKey || String(d.gpsKey) !== m[2]) return 'ID tidak dikenal';
    const r = applyPing_(d.id, { lat: Number(p.lat), lng: Number(p.lon), acc: Number(p.accuracy) || 0, speed: (Number(p.speed) || 0) * traccarSpeedFactor_(), heading: Number(p.bearing) || 0, battery: p.batt === undefined || p.batt === '' ? null : Number(p.batt) });
    return r.ignored ? 'OK (belum bertugas)' : 'OK';
  } catch (err) { return 'ERROR ' + (err && err.message || err); }
}

/* Webhook bot Telegram: driver menekan tautan t.me/<bot>?start=<idDriver>-<gpsKey> sehingga chat id tersimpan di sheet Drivers. */
function handleTelegram_(u) {
  const m = u.message || u.edited_message;
  if (!m || !m.chat) return { ok: true };
  const c = CacheService.getScriptCache(), k = 'tg:' + u.update_id;
  if (c.get(k)) return { ok: true };                                   // Telegram mengulang kiriman bila lambat
  c.put(k, '1', 3600);
  const chat = m.chat.id, mm = /^\/start(?:@\w+)?\s+(\S+)/.exec(String(m.text || '').trim()), pr = mm && /^(.+)-([A-Za-z0-9]{6,})$/.exec(mm[1]);
  if (!pr) { notifyTG_(chat, 'Halo. Untuk menerima tugas di sini, buka tautan "Hubungkan Telegram" dari aplikasi Pantau Sarana.'); return { ok: true }; }
  const d = withLock_(function () {
    const x = readRows_('Drivers', 0).filter(function (r) { return r.id === pr[1] && isActive_(r) && String(r.gpsKey) === pr[2]; })[0];
    if (x) { x.telegram = String(chat); updateRow_('Drivers', x); bust_('Drivers'); }
    return x;
  });
  notifyTG_(chat, d ? 'Terhubung. Halo ' + d.nama + ', tugas baru dan pesan pengawas akan dikirim ke sini.' : 'Kode tidak dikenal. Buka ulang tautan dari aplikasi.');
  return { ok: true };
}

function notifyDriver_(d, pesan) { notifyWA_(d.wa, pesan); notifyTG_(d.telegram, pesan); }
/* Notifikasi Telegram (OPSIONAL, gratis). Isi Script Property TELEGRAM_BOT_TOKEN dari @BotFather. */
function notifyTG_(chatId, teks) {
  const t = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!t || !chatId) return;
  try { UrlFetchApp.fetch('https://api.telegram.org/bot' + t + '/sendMessage', { method: 'post', contentType: 'application/json', payload: JSON.stringify({ chat_id: String(chatId), text: String(teks).slice(0, 3500) }), muteHttpExceptions: true }); }
  catch (e) { /* gagal kirim tidak membatalkan aksi */ }
}

/* Notifikasi WhatsApp otomatis (OPSIONAL). Isi Script Property FONNTE_TOKEN untuk mengaktifkan.
   Tanpa token, fungsi ini tidak melakukan apa pun; pengawas tetap bisa kirim manual lewat tombol WhatsApp. */
function notifyWA_(wa, pesan) {
  const t = PropertiesService.getScriptProperties().getProperty('FONNTE_TOKEN');
  if (!t || !wa) return;
  try { UrlFetchApp.fetch('https://api.fonnte.com/send', { method: 'post', headers: { Authorization: t }, payload: { target: normWA_(wa), message: pesan, countryCode: '62' }, muteHttpExceptions: true }); }
  catch (e) { /* notifikasi gagal tidak membatalkan penugasan */ }
}

/* ============ perawatan ============ */
function setupSheets() {
  const ss = ss_();
  Object.keys(SH).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, SH[name].length).setValues([SH[name]]).setFontWeight('bold'); sh.setFrozenRows(1); }
  });
  Object.keys(SH).forEach(function (name) { ss.getSheetByName(name).getRange(1, 1, 1, SH[name].length).setValues([SH[name]]); });   // kolom baru ikut ditambahkan di sheet lama
  ['Drivers', 'Pengawas'].forEach(function (n) {           // PIN dan nomor WA harus teks agar angka 0 di depan tidak hilang
    const sh = sheet_(n);
    ['pin', 'wa'].forEach(function (c) { const i = SH[n].indexOf(c); if (i >= 0) sh.getRange(2, i + 1, 500, 1).setNumberFormat('@'); });
  });
  if (sheet_('Pengawas').getLastRow() < 2) sheet_('Pengawas').appendRow(['pengawas', 'Pengawas HCGA', '081200000001', '1234']);
  if (sheet_('Drivers').getLastRow() < 2) {
    sheet_('Drivers').appendRow(['D01', 'Contoh Driver 1', '081200000101', '1111', 'Toyota Innova', 'BL 0001 AA', true, 'OFF', '']);
    sheet_('Drivers').appendRow(['D02', 'Contoh Driver 2', '081200000102', '1111', 'Toyota Hilux', 'BL 0002 AA', true, 'OFF', '']);
  }
  if (sheet_('Titik').getLastRow() < 2) {
    [['Pool Sarana HCGA', 4.1388, 96.1262, 'Pool'], ['Kantor Site', 4.1452, 96.1198, 'Kantor'], ['Mess Karyawan', 4.1347, 96.1301, 'Mess'], ['Gudang GA', 4.1421, 96.1226, 'Gudang']]
      .forEach(function (r) { sheet_('Titik').appendRow(r); });
  }
  lengkapiKunci();
  Logger.log('Selesai. GANTI PIN contoh, nama driver, dan koordinat di sheet Titik sesuai lokasi sebenarnya.');
}

/* Beri kunci GPS (kolom gpsKey) untuk driver yang belum punya. Jalankan lagi setiap menambah driver baru.
   Pengenal Traccar Client driver = id + "-" + gpsKey, contoh D03-a1b2c3d4. */
function lengkapiKunci() {
  const sh = sheet_('Drivers'), col = SH.Drivers.indexOf('gpsKey') + 1;
  readRows_('Drivers', 0).forEach(function (d) {
    if (d.id && !d.gpsKey) sh.getRange(d._row, col).setValue(Utilities.getUuid().replace(/-/g, '').slice(0, 8));
  });
  bust_('Drivers');
}

/* Pasang sebagai trigger harian: hapus jejak GPS dan chat lama agar sheet tetap ringan. */
function arsipkanDataLama() {
  [['Track', CFG.TRACK_KEEP_DAYS], ['Chat', CFG.CHAT_KEEP_DAYS], ['Log', CFG.CHAT_KEEP_DAYS]].forEach(function (x) {
    const name = x[0], cutoff = Date.now() - x[1] * 86400000, sh = sheet_(name), last = sh.getLastRow();
    if (last < 2) return;
    const col = sh.getRange(2, SH[name].indexOf('ts') + 1, last - 1, 1).getValues();
    let n = 0; while (n < col.length && (col[n][0] instanceof Date ? col[n][0].getTime() : Number(col[n][0]) || 0) < cutoff) n++;
    if (n > 0) sh.deleteRows(2, n);
  });
}
