/* ==========================================================================
   TAPACO NHÂN SỰ – TIỀN LƯƠNG · BACKEND GOOGLE APPS SCRIPT (Code.gs)
   Cách dùng: xem HUONG-DAN.md. Chạy hàm setup() một lần, rồi Triển khai → Ứng dụng web.
   ========================================================================== */

var SPREADSHEET_ID = '';            // để trống nếu Script gắn trực tiếp với Google Sheet
var TZ = 'Asia/Ho_Chi_Minh';
var SESSION_TTL = 21600;            // 6 giờ, gia hạn mỗi lần thao tác
var ADMIN_MAC_DINH = { username: 'admin', password: 'Tapaco@2026' };

function ss_() { return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }

function fmtCell_(v) {
  if (v instanceof Date) {
    var h = v.getHours() + v.getMinutes() + v.getSeconds();
    return Utilities.formatDate(v, TZ, h ? 'yyyy-MM-dd HH:mm:ss' : 'yyyy-MM-dd');
  }
  return v;
}

/* ---------- CSDL trên Google Sheets ----------
   v3.14: đọc cả bảng (hoặc chỉ phần cuối với bảng ghi thêm theo thời gian), GHI THEO DÒNG: chỉ ghi các dòng đã sửa,
   thêm dòng mới vào cuối, xóa đúng dòng bị xóa. Chỉ ghi lại cả bảng khi thứ tự dòng bị đảo / chèn giữa. */
function SheetDB_() {
  var ss = ss_(), cache = {}, heads = {}, dirty = {}, goc = {}, part = {};
  // goc[name] = { obj: [đối tượng lúc đọc], raw: [mảng ô lúc đọc], rn: [số dòng trên sheet] }
  function chuyen(name, k, cell) { var c = fmtCell_(cell); return HRCORE.isNumCol(name, k) ? (c === '' ? 0 : HRCORE.num(c)) : (c === null ? '' : String(c)); }
  function docKhoi(name, v, dongDau, h) {           // v: các dòng dữ liệu bắt đầu ở dòng sheet dongDau
    var rows = [], raw = [], rn = [];
    for (var i = 0; i < v.length; i++) {
      var r = v[i], empty = true, o = {};
      for (var j = 0; j < h.length; j++) {
        if (!h[j]) continue;
        var c = fmtCell_(r[j]);
        if (c !== '' && c !== null) empty = false;
        o[h[j]] = HRCORE.isNumCol(name, h[j]) ? (c === '' ? 0 : HRCORE.num(c)) : (c === null ? '' : String(c));
      }
      if (!empty) { rows.push(o); raw.push(r); rn.push(dongDau + i); }
    }
    return { rows: rows, raw: raw, rn: rn };
  }
  function sheet(name) { var sh = ss.getSheetByName(name); if (!sh) throw new Error('Thiếu sheet "' + name + '". Hãy chạy hàm setup().'); return sh; }
  function load(name) {
    if (cache[name] && !part[name]) return cache[name];
    var sh = sheet(name);
    if (part[name]) {                                 // đã đọc phần cuối: đọc nốt phần đầu, ghép vào CÙNG mảng (giữ tham chiếu)
      var den = part[name] - 1, h0 = heads[name];
      if (den >= 2) {
        var k0 = docKhoi(name, sh.getRange(2, 1, den - 1, h0.length).getValues(), 2, h0), arr = cache[name], g = goc[name];
        Array.prototype.unshift.apply(g.obj, k0.rows); Array.prototype.unshift.apply(g.raw, k0.raw); Array.prototype.unshift.apply(g.rn, k0.rn);
        for (var q = k0.rows.length - 1; q >= 0; q--) arr.unshift(k0.rows[q]);
      }
      delete part[name];
      return cache[name];
    }
    var v = sh.getDataRange().getValues();
    var h = (v[0] || []).map(String);
    heads[name] = h;
    var k = docKhoi(name, v.slice(1), 2, h);
    cache[name] = k.rows;
    goc[name] = { obj: k.rows.slice(), raw: k.raw, rn: k.rn };
    return k.rows;
  }
  /* bảng ghi thêm theo thời gian (vd. ChamCongGPS): chỉ đọc các dòng cuối có truong >= tu (thời điểm ghi tăng dần theo dòng) */
  function tail(name, truong, tu) {
    if (cache[name]) return cache[name];
    var sh = sheet(name), last = sh.getLastRow(), lc = Math.max(sh.getLastColumn(), 1);
    var h = sh.getRange(1, 1, 1, lc).getValues()[0].map(String), j = h.indexOf(truong);
    if (j < 0 || last < 2) return load(name);
    heads[name] = h;
    var KHUC = 1500, den = last, v = [], dau = last + 1;
    while (den >= 2) {
      var tuDong = Math.max(2, den - KHUC + 1), k = sh.getRange(tuDong, 1, den - tuDong + 1, h.length).getValues();
      v = k.concat(v); dau = tuDong; den = tuDong - 1;
      var x = null; for (var i = 0; i < k.length; i++) { var c = fmtCell_(k[i][j]); if (c !== '' && c !== null) { x = String(c); break; } }
      if (x !== null && x < tu) break;               // dòng sớm nhất của khúc đã trước mốc: các dòng phía trên còn sớm hơn
    }
    var kq = docKhoi(name, v, dau, h);
    cache[name] = kq.rows; goc[name] = { obj: kq.rows.slice(), raw: kq.raw, rn: kq.rn };
    if (dau > 2) part[name] = dau;
    return kq.rows;
  }
  function ghiCaBang(name, sh, h, rows) {
    var last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, Math.max(sh.getLastColumn(), h.length)).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, h.length).setValues(rows.map(function (r) { return dong(h, r); }));
  }
  function dong(h, r) { return h.map(function (k) { var x = r[k]; return x === undefined || x === null ? '' : x; }); }
  function kiemDong(name, sh, g) {                  // ghi muộn (khóa lấy sau khi đọc): các dòng sắp ghi / xóa vẫn phải đúng dòng trên sheet
    var h = heads[name], j = h.indexOf('id'); if (j < 0) j = h.indexOf('key'); if (j < 0) j = h.indexOf('maNV'); if (j < 0) j = 0;
    var last = sh.getLastRow(), v = last >= 2 ? sh.getRange(2, j + 1, last - 1, 1).getValues() : [];
    for (var p = 0; p < g.obj.length; p++) { var r = g.rn[p], cu = String(chuyen(name, h[j], g.raw[p][j])), tren = r - 2 < v.length ? String(chuyen(name, h[j], v[r - 2][0])) : '\u0000';
      if (cu !== tren) throw new Error('Dữ liệu "' + name + '" vừa được người khác cập nhật. Vui lòng tải lại và thao tác lại.'); }
  }
  function flushBang(name, kiem) {
    var sh = sheet(name), rows = cache[name], g = goc[name];
    if (kiem) kiemDong(name, sh, g);
    var h = heads[name].slice();
    (HRCORE.SCHEMA[name] || []).forEach(function (c) { if (h.indexOf(c) < 0) h.push(c); });
    rows.forEach(function (r) { Object.keys(r).forEach(function (k) { if (h.indexOf(k) < 0 && HRCORE.SCHEMA[name].indexOf(k) >= 0) h.push(k); }); });
    if (h.length !== heads[name].length) { sh.getRange(1, 1, 1, h.length).setValues([h]); heads[name] = h; }
    // so khớp thứ tự: các đối tượng cũ còn lại phải giữ đúng thứ tự, dòng mới chỉ nằm ở cuối
    var viTri = new Map(); g.obj.forEach(function (o, i) { viTri.set(o, i); });
    var con = [], moi = [], truoc = -1, hopLe = true;
    for (var i = 0; i < rows.length; i++) {
      var p = viTri.has(rows[i]) ? viTri.get(rows[i]) : -1;
      if (p < 0) { moi.push(rows[i]); continue; }
      if (moi.length || p <= truoc) { hopLe = false; break; }    // chèn giữa / đảo thứ tự / trùng
      con.push(p); truoc = p;
    }
    var xoa = g.obj.length - con.length;
    if (!hopLe || xoa > 30) {
      if (part[name]) load(name);                     // ghi cả bảng thì phải có đủ dữ liệu
      ghiCaBang(name, sh, h, cache[name]);
      chot(name, cache[name].map(function (r, i) { return i + 2; }));
      return;
    }
    // 1) ghi các dòng đã sửa (gộp các dòng liền nhau)
    var sua = [];
    con.forEach(function (p) {
      var o = g.obj[p], raw = g.raw[p], khac = false;
      for (var j = 0; j < h.length && !khac; j++) {
        var k = h[j]; if (!k) continue;
        var cu = raw[j] === undefined ? '' : chuyen(name, k, raw[j]);
        var mo = o[k] === undefined || o[k] === null ? '' : o[k];
        if (String(cu) !== String(mo)) khac = true;
      }
      if (khac) sua.push(p);
    });
    var nhom = [];
    sua.forEach(function (p) { var r = g.rn[p], cuoi = nhom[nhom.length - 1]; if (cuoi && cuoi.den + 1 === r) { cuoi.den = r; cuoi.ds.push(g.obj[p]); } else nhom.push({ tu: r, den: r, ds: [g.obj[p]] }); });
    nhom.forEach(function (n) { sh.getRange(n.tu, 1, n.ds.length, h.length).setValues(n.ds.map(function (r) { return dong(h, r); })); });
    // 2) xóa các dòng bị xóa (từ dưới lên, gộp liền nhau)
    var bo = [];
    if (xoa) {
      var conSet = {}; con.forEach(function (p) { conSet[p] = 1; });
      g.obj.forEach(function (o, p) { if (!conSet[p]) bo.push(g.rn[p]); });
      bo.sort(function (a, b) { return b - a; });
      for (var t = 0; t < bo.length;) { var hi = bo[t], lo = hi; t++; while (t < bo.length && bo[t] === lo - 1) { lo = bo[t]; t++; } sh.deleteRows(lo, hi - lo + 1); }
    }
    // 3) thêm dòng mới vào cuối
    var dauMoi = moi.length ? sh.getLastRow() + 1 : 0;
    if (moi.length) sh.getRange(dauMoi, 1, moi.length, h.length).setValues(moi.map(function (r) { return dong(h, r); }));
    var rn = con.map(function (p) { var r = g.rn[p], d = 0; for (var q = 0; q < bo.length; q++) if (bo[q] < r) d++; return r - d; });
    moi.forEach(function (r, i) { rn.push(dauMoi + i); });
    chot(name, rn);
  }
  function chot(name, rn) {                            // sau khi ghi: trạng thái trên sheet = dữ liệu hiện tại (để lần ghi sau trong cùng yêu cầu vẫn đúng)
    var rows = cache[name], h = heads[name];
    goc[name] = { obj: rows.slice(), raw: rows.map(function (r) { return dong(h, r); }), rn: rn };
  }
  function flush(kiem) {
    Object.keys(dirty).forEach(function (name) { if (cache[name]) flushBang(name, kiem); });
    dirty = {};
  }
  return { table: load, tail: tail, dirty: function (n) { dirty[n] = true; }, coThayDoi: function () { return Object.keys(dirty).length > 0; }, flush: flush };
}

function env_(db) {
  var cache = CacheService.getScriptCache();
  return {
    db: db,
    hash: function (s) {
      return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
        .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
    },
    uuid: function () { return Utilities.getUuid(); },
    now: function () { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); },
    today: function () { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); },
    sessions: {
      get: function (t) { var v = cache.get('s_' + t); return v ? JSON.parse(v) : null; },
      put: function (t, s) { cache.put('s_' + t, JSON.stringify(s), SESSION_TTL); },
      del: function (t) { cache.remove('s_' + t); }
    },
    attempts: {
      get: function (k) { return cache.get(k) || 0; },
      put: function (k, v) { cache.put(k, String(v), 600); }
    },
    log: function (row) {
      var sh = ss_().getSheetByName('NhatKy');
      if (sh) sh.appendRow(HRCORE.SCHEMA.NhatKy.map(function (k) { return row[k] || ''; }));
    },
    files: {
      put: function (f) {
        var blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.mime, f.name);
        var dich = null;
        if (f.chat && f.folderId) { try { dich = DriveApp.getFolderById(f.folderId); } catch (e) { dich = null; } }
        if (!dich && f.duong && f.duong.length) dich = thuMucDuong_(f.duong);      // công văn, đề xuất: thư mục nhiều cấp
        if (!dich && f.chamCong) { var cc = thuMucCon_('_CHẤM CÔNG GPS'), it = cc.getFoldersByName(f.chamCong); dich = it.hasNext() ? it.next() : cc.createFolder(f.chamCong); }   // ảnh chấm công theo tháng
        var file = (dich || (f.bantin ? thuMucCon_('_BẢNG TIN CÔNG TY') : f.chat ? thuMucCon_('_TIN NHẮN NỘI BỘ') : thuMucNV_(f.maNV, f.hoTen))).createFile(blob);
        return file.getId();                                  // giữ chế độ riêng tư mặc định, không tạo link chia sẻ
      },
      get: function (id) {
        var file = DriveApp.getFileById(id);
        return { data: Utilities.base64Encode(file.getBlob().getBytes()) };
      },
      ocr: function (id) { return ocrDrive_(id); },        // v3.37: đọc chữ ảnh / PDF bằng Google Drive (cần bật dịch vụ nâng cao "Drive API")
      thumb: function (id) {                                // v3.32: ảnh cũ chưa có ảnh thu nhỏ riêng → dùng ảnh thu nhỏ của Drive (ảnh lớn)
        var f = DriveApp.getFileById(id); if (f.getSize() <= 150000) return null;
        var b = f.getThumbnail(); return b ? { data: Utilities.base64Encode(b.getBytes()), mime: b.getContentType() || 'image/png' } : null;
      },
      del: function (id) { DriveApp.getFileById(id).setTrashed(true); },
      // thư mục con cho từng cuộc trò chuyện trong "_TIN NHẮN NỘI BỘ"
      folder: function (ten) { return thuMucCon_('_TIN NHẮN NỘI BỘ').createFolder(ten).getId(); },
      renameFolder: function (id, ten) { DriveApp.getFolderById(id).setName(ten); },
      move: function (fileId, folderId) { DriveApp.getFileById(fileId).moveTo(DriveApp.getFolderById(folderId)); },
      folderUrl: function (id) { return 'https://drive.google.com/drive/folders/' + id; },
      // tệp văn bản trong thư mục nhiều cấp: trùng tên thì ghi đè (nội dung đề xuất cập nhật theo diễn biến)
      ghiVB: function (duong, ten, nd) {
        var c = thuMucDuong_(duong), it = c.getFilesByName(ten);
        if (it.hasNext()) { var f = it.next(); f.setContent(nd); return f.getId(); }
        return c.createFile(ten, nd, MimeType.PLAIN_TEXT).getId();
      },
      duongUrl: function (duong) { return 'https://drive.google.com/drive/folders/' + thuMucDuong_(duong).getId(); }
    },
    chat: chatEnv_(cache),
    luuTru: luuTruEnv_(),
    sheetUrl: function (ten) { var ss = ss_(), sh = ss.getSheetByName(ten); return ss.getUrl() + (sh ? '#gid=' + sh.getSheetId() : ''); },   // v3.3: liên kết sheet KiemTraCong (chỉ người có quyền với tệp mới mở được)
    ai: aiEnv_(),
    tam: tamEnv_(cache),
    push: pushEnv_(),
    readLog: function (limit) {
      var sh = ss_().getSheetByName('NhatKy'), last = sh.getLastRow();
      if (last < 2) return [];
      var n = Math.min(limit, last - 1);
      var v = sh.getRange(last - n + 1, 1, n, 4).getValues();
      return v.reverse().map(function (r) { return { thoiGian: fmtCell_(r[0]), user: r[1], action: r[2], chiTiet: r[3] }; });
    }
  };
}

/* ---------- Tin nhắn: sheet TinNhan chỉ ghi thêm, đọc theo khoảng dòng (seq = số dòng) ---------- */
function chatEnv_(cache) {
  var COLS = HRCORE.SCHEMA.TinNhan, sh = null;
  function S() { if (!sh) { sh = ss_().getSheetByName('TinNhan'); if (!sh) throw new Error('Thiếu sheet "TinNhan". Hãy chạy hàm upgradeSchema().'); } return sh; }
  function rows(from, to) {
    if (to < from) return [];
    var v = S().getRange(from, 1, to - from + 1, COLS.length).getValues();
    return v.map(function (r, i) { var o = { seq: from + i }; COLS.forEach(function (k, j) { o[k] = String(fmtCell_(r[j]) == null ? '' : fmtCell_(r[j])); }); return o; });
  }
  function last() { return Math.max(1, S().getLastRow()); }
  return {
    append: function (row) {
      var n = last() + 1, rg = S().getRange(n, 1, 1, COLS.length);
      rg.setNumberFormat('@').setValues([COLS.map(function (k) { return row[k] == null ? '' : String(row[k]); })]);
      cache.put('tn_last', String(n), 21600);
      return n;
    },
    last: last,
    get: function (seq) { return seq >= 2 && seq <= last() ? rows(seq, seq)[0] : null; },
    since: function (seq, max) { var l = last(); return rows(Math.max(2, seq + 1), Math.min(l, Math.max(2, seq + 1) + max - 1)); },
    before: function (seq, n) { var to = Math.min(seq - 1, last()); return rows(Math.max(2, to - n + 1), to); },
    bump: function (users) { var o = {}, v = String(Date.now()); users.forEach(function (u) { o['cv_' + u] = v; }); if (users.length) cache.putAll(o, 21600); },
    sig: function (u) { return cache.get('cv_' + u) || '0'; },
    ghiNho: function (u, ids, l) { var o = {}; o['myht_' + u] = JSON.stringify(ids); if (l >= (+cache.get('tn_last') || 0)) o.tn_last = String(l); cache.putAll(o, 900); },   // v3.23: cho đường hỏi nhanh
    online: function (u) { cache.put('on_' + u, '1', 75); },
    onlineOf: function (us) { var g = cache.getAll(us.map(function (u) { return 'on_' + u; })); return us.filter(function (u) { return g['on_' + u]; }); },
    typing: function (ht, u, ten) {
      var k = 'tyl_' + ht, x = cache.get(k), a = x ? JSON.parse(x) : [], now = Date.now();
      a = a.filter(function (p) { return p.u !== u && now - p.t < 6000; }); a.push({ u: u, n: ten, t: now });
      cache.put(k, JSON.stringify(a), 10);
    },
    typingOf: function (hts, me) {
      // khóa gõ phím theo hội thoại: đọc danh sách người gõ đã ghi trong 6 giây
      var out = {};
      hts.forEach(function (h) { var x = cache.get('tyl_' + h); if (x) { var a = JSON.parse(x).filter(function (p) { return p.u !== me && Date.now() - p.t < 6000; }); if (a.length) out[h] = a.map(function (p) { return p.n; }); } });
      return out;
    }
  };
}

/* ---------- Bộ nhớ tạm (CacheService) cho phân tích AI nhiều bước: chia nhỏ giá trị lớn ---------- */
function tamEnv_(cache) {
  var KHUC = 30000;                                   // ký tự / khúc (giới hạn 100 KB mỗi giá trị)
  return {
    put: function (k, v, ttl) {
      var s = JSON.stringify(v), n = Math.max(1, Math.ceil(s.length / KHUC)), o = {};
      for (var i = 0; i < n; i++) o[k + '#' + i] = s.slice(i * KHUC, (i + 1) * KHUC);
      o[k] = 'n:' + n; cache.putAll(o, ttl || 21600);
    },
    get: function (k) {
      var h = cache.get(k); if (!h) return null;
      var n = +String(h).slice(2), ks = [], s = '';
      for (var i = 0; i < n; i++) ks.push(k + '#' + i);
      var g = cache.getAll(ks);
      for (i = 0; i < n; i++) { if (g[ks[i]] == null) return null; s += g[ks[i]]; }
      try { return JSON.parse(s); } catch (e) { return null; }
    }
  };
}

/* ---------- Lưu trữ nhóm chat: thư mục đặt CẠNH file Google Sheet (riêng tư, không chia sẻ) ---------- */
/* v3.19: MỘT thư mục mẹ chứa mọi dữ liệu app trên Google Drive, đặt cùng thư mục với Google Sheet (thư mục chứa app):
     <thư mục chứa Google Sheet>/TAPACO Nhân sự – Dữ liệu app (KHÔNG CHIA SẺ)/
        ├─ Tài liệu hồ sơ nhân sự (KHÔNG CHIA SẺ)   – hồ sơ CBCNV, ảnh chấm công GPS, tệp tin nhắn, bảng tin, công văn, đề xuất…
        └─ Lưu trữ nhóm chat (KHÔNG CHIA SẺ)        – lưu trữ nhóm, báo cáo phân tích AI
   Bản ≤ 3.18 tạo 2 thư mục mẹ riêng (một ở "Drive của tôi"): lần chạy đầu tự CHUYỂN vào thư mục mẹ (giữ nguyên mã tệp, liên kết cũ vẫn dùng được). */
var THU_MUC_ME_TEN = 'TAPACO Nhân sự – Dữ liệu app (KHÔNG CHIA SẺ)';
function thuMucMe_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('DATA_ROOT');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { } }
  var cha = null;
  try { var it = DriveApp.getFileById(ss_().getId()).getParents(); if (it.hasNext()) cha = it.next(); } catch (e) { }
  if (!cha) cha = DriveApp.getRootFolder();
  var co = cha.getFoldersByName(THU_MUC_ME_TEN), me = co.hasNext() ? co.next() : cha.createFolder(THU_MUC_ME_TEN);
  me.setDescription('Do app TAPACO Nhân sự quản lý: mọi tệp của app (hồ sơ CBCNV, chấm công GPS, tin nhắn, lưu trữ nhóm chat, phân tích AI…). Không chia sẻ thư mục này; truy cập qua app để có phân quyền và nhật ký.');
  props.setProperty('DATA_ROOT', me.getId());
  return me;
}
function vaoThuMucMe_(f, khoa) {                // chuyển thư mục cũ (bản ≤ 3.18) vào thư mục mẹ – làm 1 lần
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(khoa + '_GOM') === '1') return f;
  try { var me = thuMucMe_(); if (f.getId() !== me.getId()) f.moveTo(me); props.setProperty(khoa + '_GOM', '1'); } catch (e) { }
  return f;
}
/* chạy tay trong Apps Script (chọn hàm gomDuLieuDrive → Chạy) để gom ngay và xem đường dẫn */
function gomDuLieuDrive() {
  var me = thuMucMe_(); thuMucGoc_(); thuMucLuuTru_();
  Logger.log('Thư mục dữ liệu app: ' + me.getUrl());
  return me.getUrl();
}
function thuMucLuuTru_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('LT_ROOT');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return vaoThuMucMe_(f, 'LT_ROOT'); } catch (e) { } }
  var cha = thuMucMe_();
  var g = cha.createFolder('Lưu trữ nhóm chat (KHÔNG CHIA SẺ)');
  props.setProperty('LT_ROOT_GOM', '1');
  g.setDescription('Do app TAPACO Nhân sự quản lý: tin nhắn, báo cáo, tệp, ảnh của từng nhóm chat và báo cáo phân tích AI. Không chia sẻ thư mục này.');
  props.setProperty('LT_ROOT', g.getId());
  return g;
}
function luuTruEnv_() {
  function F(id) { return id ? DriveApp.getFolderById(id) : thuMucLuuTru_(); }
  return {
    thuMuc: function (chaId, ten) { var c = F(chaId), it = c.getFoldersByName(ten); return (it.hasNext() ? it.next() : c.createFolder(ten)).getId(); },
    doiTen: function (id, ten) { DriveApp.getFolderById(id).setName(ten); },
    chuyen: function (id, chaId) { DriveApp.getFolderById(id).moveTo(F(chaId)); },
    ghi: function (chaId, ten, nd, mime) {            // tạo mới hoặc ghi đè tệp văn bản cùng tên
      var c = F(chaId), it = c.getFilesByName(ten);
      if (it.hasNext()) { var f = it.next(); f.setContent(nd); return f.getId(); }
      return c.createFile(ten, nd, mime || MimeType.PLAIN_TEXT).getId();
    },
    ghiNhiPhan: function (chaId, ten, data, mime) {   // tệp nhị phân (base64): trùng tên thì thay bằng bản mới
      var c = F(chaId), it = c.getFilesByName(ten);
      while (it.hasNext()) it.next().setTrashed(true);
      return c.createFile(Utilities.newBlob(Utilities.base64Decode(data), mime, ten)).getId();
    },
    ds: function (chaId) {
      var it = F(chaId).getFiles(), out = [];
      while (it.hasNext()) { var f = it.next(); out.push({ id: f.getId(), ten: f.getName(), mime: f.getMimeType(), capNhat: Utilities.formatDate(f.getLastUpdated(), TZ, 'yyyy-MM-dd HH:mm'), kichThuoc: f.getSize() }); }
      return out;
    },
    doc: function (id) { return DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8'); },
    anh: function (id, max) {                         // ảnh cho AI / Google Docs: ảnh lớn dùng ảnh thu nhỏ của Drive
      var f = DriveApp.getFileById(id); if (!/^image\//.test(f.getMimeType())) return null;
      var b = f.getSize() <= max ? f.getBlob() : f.getThumbnail(); if (!b) return null;
      return { mime: b.getContentType() || 'image/png', data: Utilities.base64Encode(b.getBytes()) };
    },
    taoDoc: function (chaId, tieuDe, khoi, opt) { return taoDocBaoCao_(F(chaId), tieuDe, khoi, opt || {}); },
    tai: function (id) { return Utilities.base64Encode(DriveApp.getFileById(id).getBlob().getBytes()); },
    url: function (id) { return 'https://drive.google.com/drive/folders/' + (id || thuMucLuuTru_().getId()); },
    canhBao: function () {
      var f = thuMucLuuTru_(), m = [], toi = '', chu = '';
      try { toi = Session.getEffectiveUser().getEmail(); } catch (e) { }
      try { chu = f.getOwner() ? f.getOwner().getEmail() : ''; } catch (e) { }
      try { var a = f.getSharingAccess(); if (a !== DriveApp.Access.PRIVATE) m.push('đang bật chia sẻ qua đường link'); } catch (e) { }
      try { var n = f.getEditors().concat(f.getViewers()).filter(function (u) { var e = u.getEmail(); return e && e !== toi && e !== chu; }).length; if (n) m.push('có ' + n + ' người khác được cấp quyền (có thể do thư mục chứa file Google Sheet đang được chia sẻ)'); } catch (e) { }
      return m.length ? 'Thư mục lưu trữ nhóm chat ' + m.join(' và ') + '. Hãy tắt chia sẻ để bảo mật tin nhắn.' : '';
    }
  };
}
/* Google Docs báo cáo / nhật ký: khối dau, h1, h2, h3, giua, p, nho, li, bang (rong = % độ rộng cột), anh, ngat, ky do lõi dựng; opt.pdf → lưu thêm bản PDF */
function taoDocBaoCao_(thuMuc, tieuDe, khoi, opt) {
  var doc = DocumentApp.create(tieuDe), body = doc.getBody(), H = DocumentApp.ParagraphHeading, CANH = DocumentApp.HorizontalAlignment, RONG = 483;
  body.setMarginLeft(56).setMarginRight(56).setMarginTop(50).setMarginBottom(50);
  var st = {}; st[DocumentApp.Attribute.FONT_FAMILY] = 'Times New Roman'; st[DocumentApp.Attribute.FONT_SIZE] = 12; body.setAttributes(st);
  var khongVien = function (tb) { tb.setBorderWidth(0); return tb; };
  khoi.forEach(function (k) {
    try {
      if (k.t === 'dau') {
        var td = khongVien(body.appendTable([[k.x || '', k.y || '']]));
        td.getCell(0, 0).editAsText().setBold(true).setFontSize(11); td.getCell(0, 1).editAsText().setItalic(true).setFontSize(10);
        td.getCell(0, 1).getChild(0).asParagraph().setAlignment(CANH.RIGHT);
      }
      else if (k.t === 'h1') body.appendParagraph(k.x).setHeading(H.HEADING1).setAlignment(CANH.CENTER);
      else if (k.t === 'h2') body.appendParagraph(k.x).setHeading(H.HEADING2);
      else if (k.t === 'h3') body.appendParagraph(k.x).setHeading(H.HEADING3);
      else if (k.t === 'giua') body.appendParagraph(k.x).setAlignment(CANH.CENTER).editAsText().setItalic(true);
      else if (k.t === 'p') body.appendParagraph(k.x).setAlignment(CANH.JUSTIFY);
      else if (k.t === 'nho') body.appendParagraph(k.x).editAsText().setItalic(true).setFontSize(9.5).setForegroundColor('#4B5563');
      else if (k.t === 'li') body.appendListItem(k.x).setGlyphType(DocumentApp.GlyphType.BULLET);
      else if (k.t === 'ngat') body.appendPageBreak();
      else if (k.t === 'ky') {
        var tk = khongVien(body.appendTable([k.x, k.x.map(function () { return '(Ký, ghi rõ họ tên)'; }), k.x.map(function () { return '\n\n\n'; })]));
        for (var c = 0; c < k.x.length; c++) { tk.getCell(0, c).editAsText().setBold(true); tk.getCell(1, c).editAsText().setItalic(true).setFontSize(10); for (var rr = 0; rr < 3; rr++) tk.getCell(rr, c).getChild(0).asParagraph().setAlignment(CANH.CENTER); }
      }
      else if (k.t === 'bang') {
        var n = k.rows[0].length, tb = body.appendTable(k.rows.map(function (r) { var a = r.map(function (v) { return String(v == null ? '' : v); }); while (a.length < n) a.push(''); return a.slice(0, n); }));
        tb.setBorderColor('#9AA4B8');
        if (k.rong && k.rong.length === n) { var tong = k.rong.reduce(function (s2, x) { return s2 + x; }, 0); k.rong.forEach(function (w, i2) { tb.setColumnWidth(i2, Math.round(w / tong * RONG)); }); }
        for (var i = 0; i < tb.getNumRows(); i++) { var row = tb.getRow(i); for (var j = 0; j < row.getNumCells(); j++) { var cc = row.getCell(j); cc.editAsText().setFontSize(10); cc.setPaddingTop(2).setPaddingBottom(2); if (i === 0) { cc.setBackgroundColor('#DCE4F1'); cc.editAsText().setBold(true); } } }
      } else if (k.t === 'anh') {
        var img = body.appendImage(Utilities.newBlob(Utilities.base64Decode(k.data), k.mime, 'anh'));
        var w = img.getWidth(), hh = img.getHeight(); if (w > 420) { img.setWidth(420).setHeight(Math.round(hh * 420 / w)); }
        body.appendParagraph(k.x || '').editAsText().setItalic(true).setFontSize(9.5).setForegroundColor('#4A5672');
      }
    } catch (e) { body.appendParagraph('(Không chèn được một mục: ' + e.message + ')'); }
  });
  doc.saveAndClose();
  var f = DriveApp.getFileById(doc.getId()); f.moveTo(thuMuc);
  var out = { id: doc.getId(), url: doc.getUrl() };
  if (opt && opt.pdf) { try { out.pdfId = thuMuc.createFile(f.getAs('application/pdf').setName(tieuDe + '.pdf')).getId(); } catch (e) { out.loiPdf = String(e.message || e); } }
  return out;
}

/* ---------- Gọi AI (Gemini hoặc Claude). Khóa API lưu trong Thuộc tính tập lệnh, không bao giờ gửi về giao diện ---------- */
/* v3.15: 4 nhà cung cấp AI – Google Gemini, OpenAI GPT, Anthropic Claude, DeepSeek. Mỗi nhà cung cấp: khóa, model chính, model tổng hợp (tùy chọn). */
var AI_MODEL_MD = { gemini: 'gemini-flash-latest', openai: 'gpt-5-mini', claude: 'claude-haiku-4-5-20251001', deepseek: 'deepseek-chat' };
var AI_NCC_ANH = { gemini: true, openai: true, claude: true, deepseek: false };      // đọc được ảnh
var AI_OAI_URL = { openai: 'https://api.openai.com/v1', deepseek: 'https://api.deepseek.com' };
var AI_LOAI_TONG = ['tong1', 'tong2', 'dgttong'];                                     // các bước tổng hợp: dùng "model tổng hợp" nếu có
function aiVer_(x) { var m = /(\d+(?:\.\d+)?)/.exec(String(x).replace(/^[a-z]+-/, '')); return m ? parseFloat(m[1]) : 0; }
/* chọn model Gemini thay thế khi model đang dùng bị Google ngừng (lỗi 404 / "no longer available").
   ds: danh sách tên model hỗ trợ generateContent; goiY: model Google gợi ý trong thông báo lỗi; cu: model đang dùng */
function chonModelGemini_(ds, goiY, cu) {
  ds = (ds || []).map(function (x) { return String(x).replace(/^models\//, ''); }).filter(function (x) { return x && x !== cu; });
  if (goiY && goiY !== cu && (!ds.length || ds.indexOf(goiY) >= 0)) return goiY;
  function ver(x) { var m = /gemini-(\d+(?:\.\d+)?)/.exec(x); return m ? parseFloat(m[1]) : 0; }
  var on = ds.filter(function (x) { return /^gemini-\d+(\.\d+)?-flash$/.test(x); }).sort(function (a, b) { return ver(b) - ver(a); });   // bản Flash ổn định (không preview / lite / exp)
  if (on.length) return on[0];
  if (ds.indexOf('gemini-flash-latest') >= 0 && cu !== 'gemini-flash-latest') return 'gemini-flash-latest';
  var flash = ds.filter(function (x) { return /^gemini-.*flash/.test(x) && !/lite|image|tts|audio|live|embedding/.test(x); }).sort(function (a, b) { return ver(b) - ver(a); });
  if (flash.length) return flash[0];
  return cu !== 'gemini-flash-latest' ? 'gemini-flash-latest' : '';
}
/* chọn model thay thế cho OpenAI / Claude / DeepSeek: cùng dòng với model đang dùng (mini / haiku / sonnet…), bản mới nhất */
function chonModelKhac_(ncc, ds, goiY, cu) {
  ds = (ds || []).map(String).filter(function (x) { return x && x !== cu; });
  if (goiY && goiY !== cu && ds.indexOf(goiY) >= 0) return goiY;
  var by = function (re) { return ds.filter(function (x) { return re.test(x); }).sort(function (a, b) { return aiVer_(b) - aiVer_(a) || (a < b ? 1 : -1); })[0] || ''; };
  if (ncc === 'openai') {
    var la = /-nano/.test(cu) ? 'nano' : /-mini/.test(cu) ? 'mini' : '';
    return (la ? by(new RegExp('^gpt-\\d+(\\.\\d+)?-' + la + '$')) : by(/^gpt-\d+(\.\d+)?$/)) || by(/^gpt-\d+(\.\d+)?-mini$/) || by(/^gpt-\d+(\.\d+)?$/);
  }
  if (ncc === 'claude') { var dong = (/(haiku|sonnet|opus)/.exec(cu) || [, 'haiku'])[1]; return by(new RegExp(dong)) || by(/haiku/) || by(/sonnet/); }
  if (ncc === 'deepseek') return ds.indexOf('deepseek-chat') >= 0 ? 'deepseek-chat' : (ds[0] || '');
  return '';
}
function laLoiModel_(c, body) { return c === 404 || ((c === 400 || c === 403) && /no longer available|not found|not_found|does not exist|is not supported|deprecated|not available|model_not_found|invalid model/i.test(String(body))); }
/* lọc danh sách model dùng được để trò chuyện / phân tích văn bản */
function locModel_(ncc, ds) {
  return (ds || []).map(String).filter(function (x) {
    if (ncc === 'gemini') return /^gemini-/.test(x) && !/embedding|image|tts|audio|live|aqa|robotics|computer-use/.test(x);
    if (ncc === 'openai') return /^(gpt-|o\d|chatgpt-)/.test(x) && !/audio|realtime|tts|transcribe|image|embedding|search|instruct|codex|moderation|dall-e|whisper/.test(x);
    return true;
  }).sort();
}
function aiEnv_() {
  var P = PropertiesService.getScriptProperties();
  var DS = ['gemini', 'openai', 'claude', 'deepseek'];
  function cfg(ncc) {
    ncc = ncc || P.getProperty('AI_NCC') || 'gemini'; if (DS.indexOf(ncc) < 0) ncc = 'gemini';
    var N = ncc.toUpperCase();
    return { ncc: ncc, model: P.getProperty('AI_MODEL_' + N) || AI_MODEL_MD[ncc], model2: P.getProperty('AI_MODEL2_' + N) || '', key: P.getProperty('AI_KEY_' + N) || '' };
  }
  function goiHttp(url, opt) {
    for (var lan = 0; ; lan++) {
      var r;
      try { r = UrlFetchApp.fetch(url, opt); }
      catch (e) { var m = String(e && e.message || e); throw new Error(/time|thời gian/i.test(m) ? 'AI phản hồi quá 60 giây (giới hạn của Google Apps Script). Hãy chọn khoảng ngày ngắn hơn hoặc model nhanh hơn.' : 'Không gọi được AI: ' + m); }
      var c = r.getResponseCode();
      if ((c === 429 || c === 500 || c === 502 || c === 503 || c === 529) && lan < 1) { Utilities.sleep(4000); continue; }
      return r;
    }
  }
  function loi(c, body, model) {
    var m = ''; try { var j = JSON.parse(body); m = (j.error && (j.error.message || j.error.type)) || ''; } catch (e) { m = String(body).slice(0, 200); }
    if (c === 400 && /api key|API_KEY/i.test(m)) c = 401;
    if (c === 401 || c === 403) return 'Khóa API không hợp lệ hoặc chưa được cấp quyền. ' + m;
    if (c === 402) return 'Tài khoản AI hết tiền / chưa nạp tiền. ' + m;
    if (c === 404 || /no longer available/i.test(m)) return 'Model "' + model + '" không dùng được nữa và app chưa tự tìm được model thay thế. Vào Cài đặt AI: bấm "Lấy danh sách model" để chọn model đang có, hoặc để trống ô Model (dùng mặc định). ' + m;
    if (c === 429 || c === 529) return 'AI đang quá tải hoặc đã hết hạn mức sử dụng, thử lại sau ít phút. ' + m;
    return 'Lỗi AI (' + c + '): ' + m;
  }
  function dsModel(ncc, key) {                         // danh sách model của khóa API (để chọn / tự thay model bị ngừng)
    var r, out = [];
    if (ncc === 'gemini') {
      r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', { method: 'get', muteHttpExceptions: true, headers: { 'x-goog-api-key': key } });
      if (r.getResponseCode() === 200) out = (JSON.parse(r.getContentText()).models || []).filter(function (m) { return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0; }).map(function (m) { return String(m.name).replace(/^models\//, ''); });
    } else if (ncc === 'claude') {
      r = UrlFetchApp.fetch('https://api.anthropic.com/v1/models?limit=100', { method: 'get', muteHttpExceptions: true, headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' } });
      if (r.getResponseCode() === 200) out = (JSON.parse(r.getContentText()).data || []).map(function (m) { return m.id; });
    } else {
      r = UrlFetchApp.fetch(AI_OAI_URL[ncc] + '/models', { method: 'get', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + key } });
      if (r.getResponseCode() === 200) out = (JSON.parse(r.getContentText()).data || []).map(function (m) { return m.id; });
    }
    if (r && r.getResponseCode() !== 200) throw new Error(loi(r.getResponseCode(), r.getContentText(), ''));
    return out;
  }
  /* ---- từng nhà cung cấp: trả về { r: phản hồi, t: nội dung, lay: hàm lấy văn bản } ---- */
  function guiClaude(c, model, o) {
    var content = [];
    (o.anh || []).forEach(function (a) { content.push({ type: 'text', text: a.nhan }); content.push({ type: 'image', source: { type: 'base64', media_type: a.mime, data: a.data } }); });
    content.push({ type: 'text', text: o.nd });
    var r = goiHttp('https://api.anthropic.com/v1/messages', { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': c.key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: model, max_tokens: Math.min(o.maxTok || 4096, 8000), system: o.he, messages: [{ role: 'user', content: content }] }) });
    return { r: r, t: r.getContentText(), lay: function (j) { return (j.content || []).filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join(''); } };
  }
  function guiOAI(c, model, o) {                      // OpenAI và DeepSeek (cùng chuẩn Chat Completions)
    var oa = c.ncc === 'openai', user = o.nd;
    if (oa && (o.anh || []).length) { user = []; o.anh.forEach(function (a) { user.push({ type: 'text', text: a.nhan }); user.push({ type: 'image_url', image_url: { url: 'data:' + a.mime + ';base64,' + a.data } }); }); user.push({ type: 'text', text: o.nd }); }
    var body = { model: model, messages: [{ role: 'system', content: o.he + '\nTrả lời bằng một đối tượng JSON hợp lệ.' }, { role: 'user', content: user }], response_format: { type: 'json_object' } };
    if (oa) { body.max_completion_tokens = o.maxTok || 4096; if (/^(gpt-5|o\d)/.test(model)) body.reasoning_effort = 'low'; }      // giảm suy luận cho kịp 60 giây
    else { body.max_tokens = Math.min(o.maxTok || 4096, 8192); body.temperature = 0.2; }
    var url = AI_OAI_URL[c.ncc] + '/chat/completions', opt = { method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + c.key }, payload: JSON.stringify(body) };
    var r = goiHttp(url, opt), t = r.getContentText();
    for (var lan = 0; lan < 3 && r.getResponseCode() === 400; lan++) {             // model không nhận tham số nào thì bỏ tham số đó, gọi lại
      if (/reasoning_effort/i.test(t) && body.reasoning_effort) delete body.reasoning_effort;
      else if (/max_completion_tokens/i.test(t) && body.max_completion_tokens) { body.max_tokens = body.max_completion_tokens; delete body.max_completion_tokens; }
      else if (/response_format|json_object/i.test(t) && body.response_format) delete body.response_format;
      else if (/temperature/i.test(t) && body.temperature !== undefined) delete body.temperature;
      else break;
      opt.payload = JSON.stringify(body); r = goiHttp(url, opt); t = r.getContentText();
    }
    return { r: r, t: t, lay: function (j) { var m = ((j.choices || [])[0] || {}).message || {}; if (!m.content && m.refusal) throw new Error('AI từ chối trả lời: ' + m.refusal); return m.content || ''; } };
  }
  function guiGemini(c, model, o) {
    var parts = [];
    (o.anh || []).forEach(function (a) { parts.push({ text: a.nhan }); parts.push({ inlineData: { mimeType: a.mime, data: a.data } }); });
    parts.push({ text: o.nd });
    var b2 = { systemInstruction: { parts: [{ text: o.he }] }, contents: [{ role: 'user', parts: parts }],
      generationConfig: { temperature: 0.2, maxOutputTokens: o.maxTok || 4096, responseMimeType: 'application/json' } };
    if (/2\.5-flash/i.test(model)) b2.generationConfig.thinkingConfig = { thinkingBudget: 0 };          // tắt "suy nghĩ" cho nhanh (giới hạn 60 giây)
    else if (/gemini-[3-9]|flash-latest/i.test(model)) b2.generationConfig.thinkingConfig = { thinkingLevel: 'low' };
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent';
    var opt = { method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'x-goog-api-key': c.key }, payload: JSON.stringify(b2) };
    var r = goiHttp(url, opt), t = r.getContentText();
    if (r.getResponseCode() === 400 && /thinking/i.test(t)) { delete b2.generationConfig.thinkingConfig; opt.payload = JSON.stringify(b2); r = goiHttp(url, opt); t = r.getContentText(); }
    return { r: r, t: t, lay: function (j) {
      var cd = (j.candidates || [])[0];
      if (!cd) throw new Error('AI không trả kết quả' + (j.promptFeedback && j.promptFeedback.blockReason ? ' (bị chặn: ' + j.promptFeedback.blockReason + ')' : '') + '.');
      return ((cd.content || {}).parts || []).filter(function (p) { return !p.thought; }).map(function (p) { return p.text || ''; }).join('');
    } };
  }
  function gui(c, model, o) { return c.ncc === 'gemini' ? guiGemini(c, model, o) : c.ncc === 'claude' ? guiClaude(c, model, o) : guiOAI(c, model, o); }
  return {
    DS: DS,
    trangThai: function () {
      var c = cfg(), khoa = {}, md = {}, md2 = {};
      DS.forEach(function (n) { var x = cfg(n); khoa[n] = !!x.key; md[n] = x.model; md2[n] = x.model2; });
      return { ncc: c.ncc, model: c.model, model2: c.model2, coKhoa: !!c.key, khoa: khoa, md: md, md2: md2, mdMacDinh: AI_MODEL_MD, anh: AI_NCC_ANH,
        gemini: khoa.gemini, claude: khoa.claude, openai: khoa.openai, deepseek: khoa.deepseek, mdGemini: md.gemini, mdClaude: md.claude };
    },
    dat: function (o) {
      var N = o.ncc.toUpperCase(); P.setProperty('AI_NCC', o.ncc);
      if (o.model) P.setProperty('AI_MODEL_' + N, o.model); else P.deleteProperty('AI_MODEL_' + N);
      if (o.model2 !== undefined) { if (o.model2) P.setProperty('AI_MODEL2_' + N, o.model2); else P.deleteProperty('AI_MODEL2_' + N); }
      if (o.key) P.setProperty('AI_KEY_' + N, o.key);
      if (o.xoaKhoa) P.deleteProperty('AI_KEY_' + N);
    },
    dsModel: function (ncc, key) { var c = cfg(ncc); key = key || c.key; if (!key) throw new Error('Chưa có khóa API cho nhà cung cấp này.'); return locModel_(ncc, dsModel(ncc, key)); },
    goi: function (o) {
      var c = cfg(); if (!c.key) throw new Error('Chưa nhập khóa API cho ' + ({ gemini: 'Gemini', openai: 'OpenAI (GPT)', claude: 'Claude', deepseek: 'DeepSeek' })[c.ncc] + '.');
      if ((o.anh || []).length && !AI_NCC_ANH[c.ncc]) {
        // DeepSeek không đọc ảnh: nhờ nhà cung cấp khác đã có khóa; phân tích nhóm thì bỏ ảnh (chỉ là minh họa)
        var thay = DS.filter(function (n) { return AI_NCC_ANH[n] && cfg(n).key; })[0];
        if (thay) c = cfg(thay);
        else if (o.loai === 'phan') o = Object.assign({}, o, { anh: [], nd: o.nd + '\n(Ghi chú: model đang dùng không đọc ảnh – bỏ qua ảnh chứng từ.)' });
        else throw new Error(({ deepseek: 'DeepSeek' })[c.ncc] + ' không đọc được ảnh. Vào Cài đặt AI nhập thêm khóa Gemini, GPT hoặc Claude (app tự dùng cho việc đọc ảnh), hoặc chọn nhà cung cấp khác.');
      }
      var model = (AI_LOAI_TONG.indexOf(o.loai) >= 0 && c.model2) ? c.model2 : c.model, la2 = model === c.model2 && model !== c.model, doiModel = '';
      var kq = gui(c, model, o);
      if (kq.r.getResponseCode() !== 200 && laLoiModel_(kq.r.getResponseCode(), kq.t)) {
        // model bị ngừng: dò danh sách model của khóa API, chọn bản thay thế, gọi lại và lưu để lần sau dùng luôn
        var goiY = (/models\/(gemini-[\w.\-]+)/.exec(String(kq.t).replace(new RegExp('models/' + model.replace(/[.\-]/g, '\\$&'), 'g'), '')) || [])[1] || '', ds = [];
        try { ds = dsModel(c.ncc, c.key); } catch (e) { }
        var moi = c.ncc === 'gemini' ? chonModelGemini_(ds, goiY, model) : chonModelKhac_(c.ncc, ds, '', model);
        if (moi) {
          var cu = model; model = moi; kq = gui(c, model, o);
          if (kq.r.getResponseCode() === 200) { P.setProperty((la2 ? 'AI_MODEL2_' : 'AI_MODEL_') + c.ncc.toUpperCase(), moi); doiModel = cu + ' → ' + moi; }
          else model = cu;
        }
      }
      if (kq.r.getResponseCode() !== 200) throw new Error(loi(kq.r.getResponseCode(), kq.t, model));
      if (doiModel) { try { console.log('Tự đổi model ' + c.ncc + ': ' + doiModel); } catch (e) { } }
      return kq.lay(JSON.parse(kq.t));
    }
  };
}

/* ---------- Thư mục Drive chứa tài liệu hồ sơ (riêng tư) ---------- */
function thuMucGoc_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('TL_ROOT');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return vaoThuMucMe_(f, 'TL_ROOT'); } catch (e) { } }
  var root = thuMucMe_().createFolder('Tài liệu hồ sơ nhân sự (KHÔNG CHIA SẺ)');
  props.setProperty('TL_ROOT_GOM', '1');
  root.setDescription('Do app TAPACO Nhân sự quản lý. Không chia sẻ thư mục này; truy cập tài liệu qua app để có phân quyền và nhật ký.');
  props.setProperty('TL_ROOT', root.getId());
  return root;
}
function thuMucCon_(ten) {
  var root = thuMucGoc_(), it = root.getFoldersByName(ten);
  return it.hasNext() ? it.next() : root.createFolder(ten);
}
function thuMucDuong_(duong) {               // tạo / lấy thư mục theo đường nhiều cấp dưới thư mục gốc
  var f = thuMucGoc_();
  (duong || []).forEach(function (ten) {
    ten = String(ten || '').replace(/[\\/]/g, '-').trim(); if (!ten) return;
    var it = f.getFoldersByName(ten); f = it.hasNext() ? it.next() : f.createFolder(ten);
  });
  return f;
}
function thuMucNV_(maNV, hoTen) {
  var root = thuMucGoc_(), ten = maNV + ' – ' + (hoTen || '');
  var it = root.getFoldersByName(ten);
  return it.hasNext() ? it.next() : root.createFolder(ten);
}

/* v3.23: hỏi tin nhắn mới khi KHÔNG có gì thay đổi – trả lời ngay từ bộ nhớ đệm, không mở Google Sheet.
   Mỗi người mở Tin nhắn hỏi 3 giây/lần; trước đây lần nào cũng mở bảng tính (~0,5–1 s) và chiếm lượt chạy đồng thời của Apps Script,
   làm các thao tác ghi (gửi tin, lưu hồ sơ…) phải xếp hàng. Lượt có cờ online (1/5 lượt) vẫn đi đường đầy đủ để kiểm tra tài khoản. */
function pollNhanh_(req) {
  try {
    if (req.action !== 'chat.poll' || req.online || !req.token) return null;
    var c = CacheService.getScriptCache(), g = c.getAll(['s_' + req.token, 'tn_last']);
    if (!g['s_' + req.token] || !g.tn_last || String(req.cur) !== String(g.tn_last)) return null;
    var s = JSON.parse(g['s_' + req.token]), u = s.username, g2 = c.getAll(['cv_' + u, 'myht_' + u]);
    var hv = g2['cv_' + u] || '0', my = g2['myht_' + u];
    if (String(req.hv) !== String(hv) || !my) return null;
    my = JSON.parse(my);
    var ch = chatEnv_(c);
    ch.online(u);
    if (req.typing && my.indexOf(req.typing) >= 0) ch.typing(req.typing, u, s.hoTen || u);
    return { ok: true, data: { cur: +g.tn_last, hv: hv, typing: ch.typingOf(my, u), nhanh: 1 } };
  } catch (e) { return null; }
}

/* v3.32: TĂNG TỐC TIN NHẮN
   1) Đọc ảnh / tệp: kiểm tra phiên + thành viên nhóm qua bộ nhớ đệm rồi đọc thẳng Drive – không mở Google Sheet, không xếp hàng.
   2) Gửi ảnh / tệp: ghi Drive TRƯỚC khi lấy khóa ghi (ghi Drive chậm nhất) – người khác không phải chờ; lỗi thì xóa tệp đã ghi. */
function tepNhanh_(req) {
  try {
    if ((req.action !== 'chat.file' && req.action !== 'chat.anh') || !req.token) return null;
    var c = CacheService.getScriptCache(), ses = c.get('s_' + req.token); if (!ses) return null;
    var u = JSON.parse(ses).username, my = c.get('myht_' + u); if (!my) return null; my = JSON.parse(my);
    var ids = req.action === 'chat.file' ? [String(req.fileId || '')] : (Array.isArray(req.ids) ? req.ids : []).map(String).slice(0, 8);
    if (!ids.length) return null;
    var g = c.getAll(ids.map(function (i) { return 'tf_' + i; })), out = [];
    for (var i = 0; i < ids.length; i++) {
      var v = g['tf_' + ids[i]]; if (!v) return null; v = JSON.parse(v);
      if (my.indexOf(v.h) < 0) return null;
      out.push(docTepChat_(ids[i], v, !!req.thumb));
    }
    return { ok: true, data: req.action === 'chat.file' ? out[0] : { ds: out, nhanh: 1 } };
  } catch (e) { return null; }
}
function docTepChat_(fid, v, thumb) {
  var mime = v.m || '';
  if (thumb && /^image\//.test(mime)) {
    if (v.t) { var ft = DriveApp.getFileById(v.t); if (!ft.isTrashed()) return { fileId: fid, data: Utilities.base64Encode(ft.getBlob().getBytes()), mime: 'image/jpeg', name: v.n, thumb: 1 }; }
    var f0 = DriveApp.getFileById(fid); if (f0.isTrashed()) throw new Error('trashed');
    if (f0.getSize() > 150000) { var b = f0.getThumbnail(); if (b) return { fileId: fid, data: Utilities.base64Encode(b.getBytes()), mime: b.getContentType() || 'image/png', name: v.n, thumb: 1 }; }
    return { fileId: fid, data: Utilities.base64Encode(f0.getBlob().getBytes()), mime: mime, name: v.n, thumb: 0 };
  }
  var f = DriveApp.getFileById(fid); if (f.isTrashed()) throw new Error('trashed');
  return { fileId: fid, data: Utilities.base64Encode(f.getBlob().getBytes()), mime: mime || f.getMimeType(), name: v.n, thumb: 0 };
}
function taiTruoc_(req) {
  if (req.action !== 'chat.upload' || !req.token || !req.file || !req.file.data || !req.ht) return;
  var c = CacheService.getScriptCache(); if (!c.get('s_' + req.token)) return;
  var f = req.file, mime = String(f.mime || '');
  if (!/^(image\/(jpeg|png|webp|gif)|application\/|text\/plain)/.test(mime)) return;
  var fd = c.get('chfd_' + req.ht) || '', dich = null;
  if (fd) { try { dich = DriveApp.getFolderById(fd); } catch (e) { dich = null; fd = ''; } }
  if (!dich) dich = thuMucCon_('_TIN NHẮN NỘI BỘ');
  var ten = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') + '_' + String(f.name || 'tep').trim().replace(/[\\/:*?"<>|]/g, '_').slice(0, 150);
  var s64 = String(f.data), bytes = Math.floor(s64.length * 3 / 4) - (/==$/.test(s64) ? 2 : /=$/.test(s64) ? 1 : 0);
  var id = dich.createFile(Utilities.newBlob(Utilities.base64Decode(f.data), mime, ten)).getId(), th = '';
  if (f.thumb && /^image\//.test(mime) && String(f.thumb).length < 400000) { try { th = dich.createFile(Utilities.newBlob(Utilities.base64Decode(f.thumb), 'image/jpeg', ten.replace(/(\.[^.]+)?$/, '_thumb.jpg'))).getId(); } catch (e) { th = ''; } }
  req._pre = { id: id, thumbId: th, folder: fd || dich.getId(), bytes: bytes };
  delete f.data; delete f.thumb;
}
/* v3.37: OCR bằng Google Drive – ảnh / PDF chuyển sang Google Docs (Drive tự đọc chữ, tiếng Việt), lấy văn bản rồi xóa bản tạm.
   Chạy TRƯỚC khi lấy khóa ghi (đọc chữ mất vài giây). Cần: Apps Script → Dịch vụ (Services) → thêm "Drive API". */
function ocrDrive_(id) {
  if (typeof Drive === 'undefined' || !Drive.Files) throw new Error('chưa bật dịch vụ "Drive API" trong Apps Script (Dịch vụ → Drive API → Thêm)');
  var f = DriveApp.getFileById(id), blob = f.getBlob(), ten = 'OCR_tam_' + f.getName(), gid;
  if (Drive.Files.create) gid = Drive.Files.create({ name: ten, mimeType: 'application/vnd.google-apps.document' }, blob, { ocrLanguage: 'vi', fields: 'id' }).id;   // Drive API v3
  else gid = Drive.Files.insert({ title: ten, mimeType: 'application/vnd.google-apps.document' }, blob, { ocr: true, ocrLanguage: 'vi' }).id;                            // Drive API v2
  try { return DocumentApp.openById(gid).getBody().getText(); }
  finally { try { DriveApp.getFileById(gid).setTrashed(true); } catch (e) { } }
}
function ocrTruoc_(req) {
  if (req.action !== 'tl.ocr' || !req.token || !req.fileId || req.text) return;
  if (!CacheService.getScriptCache().get('s_' + req.token)) return;
  try { req._ocr = ocrDrive_(String(req.fileId)); req._ocrFid = String(req.fileId); } catch (e) { req._ocrLoi = String(e && e.message || e); }   // lỗi → máy khách đọc trên máy
}
function boTaiTruoc_(req) { if (!req._pre) return; [req._pre.id, req._pre.thumbId].forEach(function (x) { if (x) { try { DriveApp.getFileById(x).setTrashed(true); } catch (e) { } } }); }
function ghiNhoTep_(res) {
  var d = res && res.ok && res.data; if (!d || typeof d !== 'object') return;
  try {
    var c = CacheService.getScriptCache(), o = {};
    if (d._tf) Object.keys(d._tf).forEach(function (k) { o['tf_' + k] = JSON.stringify(d._tf[k]); });
    if (d._fd) Object.keys(d._fd).forEach(function (k) { o['chfd_' + k] = String(d._fd[k]); });
    if (Object.keys(o).length) c.putAll(o, 21600);
    if (d._tfXoa && d._tfXoa.length) c.removeAll(d._tfXoa.map(function (k) { return 'tf_' + k; }));
  } catch (e) { }
  delete d._tf; delete d._fd; delete d._tfXoa;
}

function handle_(req) {
  delete req._pre; delete req._ocr; delete req._ocrFid; delete req._ocrLoi;  // chỉ máy chủ được đặt (tệp đã ghi / văn bản OCR đọc trước khi lấy khóa)
  var nhanh = pollNhanh_(req) || tepNhanh_(req); if (nhanh) return nhanh;
  try { taiTruoc_(req); } catch (e) { boTaiTruoc_(req); delete req._pre; }
  ocrTruoc_(req);
  var db = SheetDB_();
  var api = HRCORE.createApi(env_(db));
  var lock = null;
  if (!api.isReadOnly(req.action)) {
    lock = LockService.getScriptLock();
    if (!lock.tryLock(25000)) { boTaiTruoc_(req); return { ok: false, error: 'Hệ thống đang bận, vui lòng thử lại sau vài giây.' }; }
  }
  var res;
  try {
    res = api.handle(req);
    if (res.ok) {
      // thao tác chỉ đọc (không giữ khóa) nhưng phát sinh ghi (vd. tự áp quyết định lương đến hạn khi mở app): giữ khóa lúc ghi
      var muon = !lock && db.coThayDoi();
      if (muon) { lock = LockService.getScriptLock(); if (!lock.tryLock(25000)) throw new Error('Hệ thống đang bận, vui lòng thử lại sau vài giây.'); }
      db.flush(muon);
    }
  } catch (e) {
    res = { ok: false, error: String(e && e.message || e) };
  } finally {
    if (lock) lock.releaseLock();
  }
  if (req._pre && !(res && res.ok)) boTaiTruoc_(req);
  ghiNhoTep_(res);
  if (res && res._push) { var ds = res._push; delete res._push; pushSau_(ds); }   // gửi thông báo đẩy sau khi đã nhả khóa
  return res;
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: 'Dữ liệu gửi lên không hợp lệ.' }); }
  return json_(handle_(req));
}

function doGet() { return json_({ ok: true, app: 'TAPACO Nhân sự – Tiền lương', version: HRCORE.VERSION }); }

/* ---------- Khởi tạo / nâng cấp cấu trúc (không xóa dữ liệu) ---------- */
function upgradeSchema() {
  var ss = ss_();
  ss.setSpreadsheetTimeZone(TZ);
  Object.keys(HRCORE.SCHEMA).forEach(function (name) {
    var cols = HRCORE.SCHEMA[name], sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, cols.length).setValues([cols]); }
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var h = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String).filter(function (x) { return x; });
    cols.forEach(function (c) { if (h.indexOf(c) < 0) h.push(c); });
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#0F1E3D').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    // cột chữ để dạng văn bản: giữ số 0 đầu (CCCD, SĐT, STK) và ngày dạng yyyy-MM-dd
    h.forEach(function (c, i) {
      var rng = sh.getRange(2, i + 1, Math.max(sh.getMaxRows() - 1, 1), 1);
      var demNgay = /^(ngayCong|nghi|ot\d|soNPT|ngayCongChuan|ngayCongHL)/.test(c);
      rng.setNumberFormat(HRCORE.isNumCol(name, c) ? (demNgay ? 'General' : '#,##0') : '@');
    });
  });
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Trang tính1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
}

function setup() {
  upgradeSchema();
  var db = SheetDB_();
  var api = HRCORE.createApi(env_(db));
  var cfgRows = db.table('CauHinh');
  if (!cfgRows.length) {
    Object.keys(HRCORE.DEFAULT_CFG).forEach(function (k) { cfgRows.push({ key: k, value: JSON.stringify(HRCORE.DEFAULT_CFG[k]) }); });
    db.dirty('CauHinh');
  }
  var users = db.table('Users');
  if (!users.some(function (u) { return u.role === 'admin'; })) {
    api.upsertUser({ username: 'system', role: 'admin' }, { username: ADMIN_MAC_DINH.username, role: 'admin', hoTen: 'Quản trị hệ thống', password: ADMIN_MAC_DINH.password });
  }
  db.flush();
  thuMucGoc_();                                                  // tạo thư mục tài liệu + xin quyền Google Drive
  Logger.log('Đã khởi tạo. Đăng nhập: ' + ADMIN_MAC_DINH.username + ' / ' + ADMIN_MAC_DINH.password + ' (bắt buộc đổi mật khẩu lần đầu).');
}

/* ---------- Nhắc hẹn, yêu cầu báo cáo định kỳ trong nhóm chat ----------
   Chạy 1 lần hàm caiDatNhacTuDong() để tạo trình kích hoạt 5 phút/lần (Apps Script → Trình kích hoạt). */
function nhacTuDong() {
  var lock = LockService.getScriptLock(), dsPush = [];
  if (!lock.tryLock(20000)) return;
  try {
    var db = SheetDB_(); var api = HRCORE.createApi(env_(db)); var n = api.chayLich(); db.flush(); if (n) Logger.log('Đã gửi ' + n + ' nhắc hẹn / yêu cầu báo cáo.');
    dsPush = api.layPush();
    // mỗi ngày một lần: lưu trữ nhóm chat có tin mới ra Drive (mỗi lượt tối đa ~50 giây, chưa xong thì lượt 5 phút sau làm tiếp)
    var P = PropertiesService.getScriptProperties(), hom = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    if (P.getProperty('LT_NGAY') !== hom) {
      var kq = api.luuTruTuDong(50000); db.flush();
      if (kq.xong) P.setProperty('LT_NGAY', hom);
      if (kq.soNgay) Logger.log('Lưu trữ nhóm chat: ' + kq.soNgay + ' ngày nhật ký' + (kq.xong ? '.' : ' (còn tiếp).'));
    }
  }
  finally { lock.releaseLock(); }
  if (dsPush.length) pushSau_(dsPush);
}
function caiDatNhacTuDong() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'nhacTuDong') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('nhacTuDong').timeBased().everyMinutes(5).create();
  Logger.log('Đã cài: nhắc hẹn và yêu cầu báo cáo trong nhóm chat chạy 5 phút/lần.');
}

/* Bản 2.8: chạy 1 lần sau khi cập nhật mã – tạo thư mục lưu trữ nhóm chat cạnh file Google Sheet, cấp quyền Drive / Google Docs / gọi AI,
   bật trình kích hoạt 5 phút (nhắc hẹn + lưu trữ hằng ngày) */
function caiDatLuuTru() {
  upgradeSchema();
  var me = gomDuLieuDrive(), g = thuMucLuuTru_(), c = luuTruEnv_().canhBao();
  Logger.log('Thư mục dữ liệu app (thư mục mẹ): ' + me);
  try { DocumentApp.getActiveDocument(); } catch (e) { }                       // để Apps Script xin quyền Google Docs
  try { UrlFetchApp.getRequest('https://generativelanguage.googleapis.com'); } catch (e) { }   // và quyền gọi dịch vụ bên ngoài (AI)
  caiDatNhacTuDong();
  Logger.log('Thư mục lưu trữ nhóm chat: ' + g.getName() + ' → ' + g.getUrl());
  Logger.log(c ? 'CẢNH BÁO: ' + c : 'Thư mục đang ở chế độ riêng tư (không chia sẻ). Tốt.');
}

/* Khôi phục quyền quản trị khi quên mật khẩu admin: chạy trong trình soạn thảo Apps Script */
function resetAdminPassword() {
  var db = SheetDB_();
  var api = HRCORE.createApi(env_(db));
  api.upsertUser({ username: 'system', role: 'admin' }, { username: ADMIN_MAC_DINH.username, role: 'admin', hoTen: 'Quản trị hệ thống', password: ADMIN_MAC_DINH.password });
  db.flush();
  Logger.log('Đã đặt lại mật khẩu ' + ADMIN_MAC_DINH.username + ' = ' + ADMIN_MAC_DINH.password);
}


/* ---------- Thông báo đẩy: Firebase Cloud Messaging (HTTP v1) ----------
   Khóa tài khoản dịch vụ Firebase lưu trong Thuộc tính tập lệnh (không ghi vào Google Sheet, không gửi về trình duyệt).
   Mã truy cập OAuth ký bằng RS256 ngay trong Apps Script, lưu đệm 55 phút. Mỗi thông báo 1 lệnh gọi; gửi song song theo lô 50. */
var PUSH_KHOA_ = 'FCM_TAI_KHOAN_DICH_VU';
function pushEnv_() {
  return {
    khoa: function () { var j = PropertiesService.getScriptProperties().getProperty(PUSH_KHOA_); if (!j) return null; var o = JSON.parse(j); return { email: o.client_email, projectId: o.project_id }; },
    luuKhoa: function (sa) {
      PropertiesService.getScriptProperties().setProperty(PUSH_KHOA_, JSON.stringify({ client_email: sa.client_email, private_key: sa.private_key, project_id: sa.project_id }));
      CacheService.getScriptCache().remove('fcm_at');
      return { email: sa.client_email, projectId: sa.project_id };
    }
  };
}
function b64u_(x) { return Utilities.base64EncodeWebSafe(x).replace(/=+$/, ''); }
function fcmMaTruyCap_() {
  var c = CacheService.getScriptCache(), at = c.get('fcm_at'); if (at) return at;
  var j = PropertiesService.getScriptProperties().getProperty(PUSH_KHOA_); if (!j) throw new Error('Chưa có khóa tài khoản dịch vụ Firebase.');
  var sa = JSON.parse(j), now = Math.floor(Date.now() / 1000);
  var dau = b64u_(JSON.stringify({ alg: 'RS256', typ: 'JWT' })), than = b64u_(JSON.stringify({ iss: sa.client_email, scope: 'https://www.googleapis.com/auth/firebase.messaging', aud: 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600 }));
  var ky = b64u_(Utilities.computeRsaSha256Signature(dau + '.' + than, sa.private_key));
  var r = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', { method: 'post', payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: dau + '.' + than + '.' + ky }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('Google từ chối khóa Firebase: ' + r.getContentText().slice(0, 200));
  at = JSON.parse(r.getContentText()).access_token; c.put('fcm_at', at, 3300);
  return at;
}
function pushGui_(ds) {
  var kq = { gui: 0, loi: 0, xoa: [] };
  if (!ds || !ds.length) return kq;
  var j = PropertiesService.getScriptProperties().getProperty(PUSH_KHOA_); if (!j) return kq;
  var pid = JSON.parse(j).project_id, at = fcmMaTruyCap_(), c = CacheService.getScriptCache();
  // cùng thiết bị + cùng cuộc trò chuyện trong 20 giây: vẫn cập nhật nội dung thông báo nhưng không đổ chuông lại (data.im = 1)
  var keys = ds.map(function (x) { return 'pc_' + b64u_(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, x.token + '|' + x.data.tag)); });
  var da = c.getAll(keys), moi = {};
  var reqs = ds.map(function (x, i) {
    var d = {}; Object.keys(x.data).forEach(function (k) { d[k] = String(x.data[k] == null ? '' : x.data[k]); });
    if (da[keys[i]]) d.im = '1'; moi[keys[i]] = '1';
    return { url: 'https://fcm.googleapis.com/v1/projects/' + pid + '/messages:send', method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + at },
      payload: JSON.stringify({ message: { token: x.token, data: d, webpush: { headers: { Urgency: 'high', TTL: '86400' } } } }), muteHttpExceptions: true };
  });
  try { c.putAll(moi, 20); } catch (e) { }
  for (var i = 0; i < reqs.length; i += 50) {
    UrlFetchApp.fetchAll(reqs.slice(i, i + 50)).forEach(function (r, k) {
      var code = r.getResponseCode(), t = code === 200 ? '' : r.getContentText();
      if (code === 200) { kq.gui++; return; }
      kq.loi++;
      if (code === 404 || code === 403 || /UNREGISTERED|registration-token-not-registered|SENDER_ID_MISMATCH|not a valid FCM registration token/i.test(t)) kq.xoa.push(ds[i + k].token);
      if (code === 401) c.remove('fcm_at');
    });
  }
  return kq;
}
function pushSau_(ds) {
  try {
    var kq = pushGui_(ds);
    if (kq.xoa.length) {                                            // thiết bị đã gỡ app / hết hạn: xóa mã khỏi sheet ThietBiPush
      var lock = LockService.getScriptLock();
      if (lock.tryLock(5000)) { try { var db = SheetDB_(); HRCORE.createApi(env_(db)).xoaTokenPush(kq.xoa); db.flush(); } finally { lock.releaseLock(); } }
    }
    if (kq.loi) Logger.log('Thông báo đẩy: gửi ' + kq.gui + ', lỗi ' + kq.loi + (kq.xoa.length ? ', xóa ' + kq.xoa.length + ' thiết bị hết hạn' : ''));
  } catch (e) { Logger.log('Thông báo đẩy lỗi: ' + e.message); }
}

/* ==========================================================================
   TAPACO NHÂN SỰ – TIỀN LƯƠNG · LÕI NGHIỆP VỤ (dùng chung Code.gs & bản dùng thử)
   Không sửa riêng ở một nơi: file này được nhúng nguyên văn vào cả 2 phía.
   ========================================================================== */
var HRCORE = (function () {
  'use strict';

  var VERSION = '3.42';

  /* Chức danh (vai trò) của tài khoản. Mỗi tài khoản (trừ Quản trị) gắn với 1 hồ sơ CBCNV và 1 bộ phận. */
  var ROLES = {
    admin: 'Quản trị hệ thống',
    giamdoc: 'Giám đốc',
    tchc: 'Tổ chức – Hành chính',
    ktluong: 'Kế toán tiền lương',
    ktcongno: 'Kế toán công nợ',
    kttruong: 'Phụ trách kế toán',
    truongphong: 'Trưởng phòng',
    cht: 'Chỉ huy trưởng công trình',
    nhanvien: 'Nhân viên'
  };
  /* mật khẩu ban đầu = tên (không dấu, chữ thường) + năm sinh: thang1990 */
  function matKhauTheoNV(nv, cfg) {
    var w = String((nv && nv.hoTen) || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().replace(/[^a-z\s]/g, ' ').trim().split(/\s+/).filter(Boolean);
    var y = /^(\d{4})-/.exec(String((nv && nv.ngaySinh) || ''));
    var mk = w.length && y ? w[w.length - 1] + y[1] : '';
    return mk.length >= 6 ? mk : String((cfg && cfg.matKhauMacDinh) || 'Tapaco@123');
  }
  var ROLE_CU = { hr: 'tchc', ketoan: 'ktluong', congno: 'ktcongno' };   // tên vai trò bản ≤ 2.2
  var ROLE_QL = ['truongphong', 'cht'];                                   // quản lý trực tiếp bộ phận
  function normRole(r) { r = String(r || '').trim(); return ROLE_CU[r] || (ROLES[r] ? r : 'nhanvien'); }
  var ROLES_MOTA = {
    admin: 'Quản lý tài khoản, chức danh, cài đặt, công thức lương. Chỉ có 1 tài khoản.',
    giamdoc: 'Toàn quyền phê duyệt. Giám đốc duyệt thì bỏ qua các bước còn lại.',
    tchc: 'Duyệt chấm công, nghỉ phép, bảng lương; lập đề xuất nâng lương, khấu trừ theo quy chế (mất mát, bồi thường…); xử lý đề nghị của CBCNV; quản lý hồ sơ, điều động.',
    ktluong: 'Lập bảng lương.',
    ktcongno: 'Chấm công, kiểm tra công nợ – tạm ứng.',
    kttruong: 'Phụ trách kế toán: kiểm soát thuế TNCN, nghĩa vụ bảo hiểm; phê duyệt cam kết thu nhập, giấy ủy quyền quyết toán, hồ sơ thuế.',
    truongphong: 'Duyệt đơn nghỉ, xem nhân sự của phòng mình.',
    cht: 'Duyệt đơn nghỉ, xem nhân sự của ban chỉ huy công trình mình.',
    nhanvien: 'Xem hồ sơ, phiếu lương, gửi đơn nghỉ và ý kiến của mình.'
  };
  function tenDangNhap(hoTen, daCo, ngaySinh) {
    var ns = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(ngaySinh || ''));
    var w = String(hoTen || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().replace(/[^a-z\s]/g, ' ').trim().split(/\s+/).filter(Boolean);
    var g = w.length ? w[w.length - 1] + w.slice(0, -1).map(function (x) { return x[0]; }).join('') : 'nv';
    while (g.length < 3) g += 'x';
    if (ns) g += ns[3] + ns[2] + ns[1].slice(2);                      // + ngày tháng năm sinh ddmmyy: thanglv201190
    var u = g, i = 2; while (daCo(u)) u = g + (i++);
    return u;
  }

  var PERM = {
    view_all: ['admin', 'giamdoc', 'tchc', 'ktluong', 'ktcongno', 'kttruong'],
    edit_nv: ['admin', 'tchc'],                       // hồ sơ, quá trình, điều động, phòng ban – công trình
    propose: ['admin', 'tchc'],                       // lập đề xuất nâng lương, khen thưởng – kỷ luật
    approve: ['admin', 'giamdoc'],                    // phê duyệt nâng lương, khen thưởng – kỷ luật
    payroll: ['admin', 'ktluong'],
    payroll_approve: ['admin', 'giamdoc'],
    users: ['admin', 'tchc'],                         // TCHC: chỉ tài khoản Nhân viên
    doimk: ['admin', 'giamdoc'],                      // đặt lại mật khẩu cho tài khoản khác
    config: ['admin'],
    log: ['admin', 'giamdoc'],
    quytrinh: ['admin', 'giamdoc'],
    tailieu: ['admin', 'giamdoc', 'tchc', 'ktluong', 'kttruong'],
    duyetThue: ['admin', 'giamdoc', 'kttruong'],      // v3.27: phê duyệt cam kết, ủy quyền quyết toán, hồ sơ thuế (Phụ trách kế toán)
    congthuc: ['admin'],
    bantin: ['admin', 'giamdoc', 'tchc'],
    congvan: ['admin', 'giamdoc', 'tchc'],            // xem toàn bộ sổ công văn đến – đi
    vanthu: ['admin', 'tchc'],                        // văn thư: đăng ký vào sổ, tải bản quét, lưu hồ sơ
    aiNhom: ['admin', 'giamdoc', 'tchc', 'ktluong', 'ktcongno'],     // Phân tích nhóm (AI): lưu trữ, phân tích, xem báo cáo mọi nhóm
    ccQuanLy: ['admin', 'giamdoc', 'tchc', 'ktluong', 'ktcongno'],   // Chấm công GPS – quản lý, kiểm tra AI, vị trí ngoài phạm vi
    danhgia: ['admin', 'giamdoc', 'tchc'],                           // Đánh giá tổng hợp CBCNV (AI)
    botqd: ['admin', 'giamdoc', 'tchc'],                             // Sửa quy định của trợ lý trong Tin nhắn
    xem_bp: ['truongphong', 'cht'],                                  // v3.38: xem nhân sự bộ phận mình (danh sách, quá trình, đơn nghỉ)
    tu_xem: ['giamdoc', 'tchc', 'kttruong', 'ktcongno', 'ktluong', 'truongphong', 'cht', 'nhanvien'],   // hồ sơ, phiếu lương… của chính mình (cố định)
    sua_pq: ['admin', 'giamdoc']                                     // v3.38: sửa bảng phân quyền
  };
  /* ===== BẢNG PHÂN QUYỀN TÍCH CHỌN ĐƯỢC (v3.8) =====
     Giám đốc, Quản trị tích chọn ở Tài khoản → Bảng phân quyền; lưu cfg.phanQuyen = { quyền: [vai trò] }.
     PERM_MD là mặc định; ô khóa (PQ_KHOA) luôn giữ để không ai tự khóa mình khỏi hệ thống. */
  var PERM_MD = JSON.parse(JSON.stringify(PERM));
  var PQ_NHOM = [
    { ten: 'Hồ sơ & nhân sự', ds: [['view_all', 'Xem toàn bộ hồ sơ, lương, quá trình'], ['edit_nv', 'Hồ sơ, quá trình công tác, phòng ban – công trình, điều động'], ['tailieu', 'Xem tài liệu đính kèm của mọi CBCNV'], ['propose', 'Lập đề xuất nâng lương, khen thưởng – kỷ luật'], ['approve', 'Phê duyệt nâng lương, khen thưởng – kỷ luật'], ['duyetThue', 'Phê duyệt cam kết thu nhập, ủy quyền quyết toán, hồ sơ thuế TNCN']] },
    { ten: 'Văn bản & thông tin', ds: [['bantin', 'Đăng Bảng tin'], ['congvan', 'Xem toàn bộ sổ công văn đến – đi'], ['vanthu', 'Văn thư: vào sổ công văn, tải bản quét, AI đọc văn bản, lưu hồ sơ']] },
    { ten: 'Giám sát & AI', ds: [['aiNhom', 'Phân tích nhóm (AI): lưu trữ, phân tích, xem báo cáo mọi nhóm'], ['ccQuanLy', 'Chấm công GPS – quản lý: lượt chấm, ảnh, bản đồ, duyệt, công tác, kiểm tra AI, vị trí ngoài phạm vi'], ['danhgia', 'Đánh giá tổng hợp CBCNV (AI)'], ['botqd', 'Sửa quy định của trợ lý trong Tin nhắn']] },
    { ten: 'Hệ thống', ds: [['users', 'Quản lý tài khoản (ngoài Quản trị: chỉ tài khoản Nhân viên)'], ['doimk', 'Đặt lại mật khẩu cho tài khoản khác'], ['log', 'Xem nhật ký thao tác'], ['quytrinh', 'Cài đặt quy trình & người duyệt'], ['config', 'Cài đặt tham số, nhóm – loại đề xuất'], ['congthuc', 'Lập công thức tính lương']] },
    { ten: 'Theo chức danh', ds: [['xem_bp', 'Xem nhân sự bộ phận mình (danh sách, quá trình công tác, đơn nghỉ của bộ phận)'], ['tu_xem', 'Hồ sơ, phiếu lương, đơn nghỉ, ý kiến của chính mình'], ['sua_pq', 'Sửa bảng phân quyền này']] }
  ];
  var PQ_SUA = []; PQ_NHOM.forEach(function (g) { g.ds.forEach(function (x) { PQ_SUA.push(x[0]); }); });
  var PQ_KHOA = { users: ['admin'], config: ['admin'], congthuc: ['admin'], quytrinh: ['admin', 'giamdoc'], view_all: ['admin', 'giamdoc'], approve: ['giamdoc'] };
  PQ_KHOA.sua_pq = ['admin', 'giamdoc']; PQ_KHOA.tu_xem = PERM.tu_xem.slice();   // v3.38: tự xem của mọi CBCNV luôn giữ (Điều 95 BLLĐ)
  var PQ_CAM = { tu_xem: ['admin'], sua_pq: ['nhanvien'], xem_bp: ['admin'] };   // ô không cấp được
  var PQ_NGUOI_SUA = ['admin', 'giamdoc'];                 // mặc định người sửa bảng phân quyền (v3.38: theo quyền sua_pq)
  function pqKhoa(k, r) { return (PQ_KHOA[k] || []).indexOf(r) >= 0; }
  function pqCam(k, r) { return (PQ_CAM[k] || []).indexOf(r) >= 0; }
  var _pqDaAp = null;
  function apDungPQ(pq) {                                  // áp cấu hình vào PERM (giữ nguyên đối tượng PERM)
    var key = JSON.stringify(pq || {}); if (key === _pqDaAp) return; _pqDaAp = key;
    Object.keys(PERM_MD).forEach(function (k) {
      var v = (PQ_SUA.indexOf(k) >= 0 && pq && Array.isArray(pq[k])) ? pq[k].filter(function (r, i, a) { return ROLES[r] && a.indexOf(r) === i && !pqCam(k, r); }) : PERM_MD[k].slice();
      (PQ_KHOA[k] || []).forEach(function (r) { if (v.indexOf(r) < 0) v.push(r); });
      PERM[k] = v;
    });
  }
  function pqLamSach(pq) {                                 // chỉ giữ quyền sửa được, vai trò hợp lệ, cộng ô khóa
    var o = {}; PQ_SUA.forEach(function (k) {
      var v = Array.isArray(pq && pq[k]) ? pq[k] : PERM_MD[k];
      v = v.filter(function (r, i, a) { return ROLES[r] && a.indexOf(r) === i && !pqCam(k, r); });
      (PQ_KHOA[k] || []).forEach(function (r) { if (v.indexOf(r) < 0) v.push(r); });
      var rk = Object.keys(ROLES); o[k] = v.sort(function (a, b) { return rk.indexOf(a) - rk.indexOf(b); });
    });
    return o;
  }
  function can(role, p) { return (PERM[p] || []).indexOf(role) >= 0; }
  function HRCORE_bcBangDoc(t) {               // v3.9: đọc giá trị "Tên: số đơn vị" thành các dòng (bỏ dòng Tổng)
    return String(t || '').split(/\n/).map(function (l) {
      var m = /^(.+?):\s*([\d.,]+)\s*(.*)$/.exec(l.trim()); return m && !/^tổng$/i.test(m[1].trim()) ? { ten: m[1].trim(), so: m[2], dv: m[3].trim() } : null;
    }).filter(Boolean);
  }

  /* ===== QUY TRÌNH CẤU HÌNH ĐƯỢC =====
     Người xử lý (nguoi) là mảng mã: 'r:<vai trò>' | 'u:<tên đăng nhập>' | 'tbp' (trưởng bộ phận/công trường của người đề nghị) */
  var BUOC_XONG = 'Đã duyệt';
  var LOAI_BUOC = {
    chamcong: 'Chấm công (nhập công, nghỉ, OT)',
    tinhluong: 'Tính lương (truy lĩnh, tính toán)',
    congno: 'Công nợ (tạm ứng, ăn ca, thu ứng, khấu trừ khác)',
    vipham: 'Kiểm tra & khấu trừ vi phạm quy chế',
    kiemtra: 'Kiểm tra / soát xét (chỉ xem, chuyển hoặc trả lại)',
    pheduyet: 'Phê duyệt cuối & khóa dữ liệu'
  };
  var FIELDS_LOAI = {
    chamcong: ['ngayCong', 'nghiHuongLuong', 'nghiKhongLuong', 'nghiOm', 'ngayLe', 'ot150', 'ot200', 'ot300', 'noiLamViec', 'ghiChu', 'bienKy', 'chiTietNgay', 'chiTietOT', 'chiTietAn'],
    tinhluong: ['truyLinh', 'tienKhoan', 'bienKy'],
    congno: ['tamUng', 'tienAnThu', 'thuUng', 'khauTruKhac', 'ghiChuCongNo', 'chiTietTamUng'],
    vipham: ['chiTietPhat', 'ghiChuTCHC']
  };
  /* Quy trình lương mặc định (TAPACO, 10/2026). Giám đốc có thể phê duyệt ở bất kỳ bước nào – bỏ qua các bước còn lại. */
  var QT_LUONG_MD = [
    { id: 'Chấm công', ten: 'Chấm công', loai: 'chamcong', nguoi: ['r:ktcongno'] },
    { id: 'Công nợ', ten: 'Kiểm tra công nợ', loai: 'congno', nguoi: ['r:ktcongno'] },
    { id: 'TCHC duyệt công', ten: 'TCHC duyệt chấm công, khấu trừ quy chế', loai: 'vipham', nguoi: ['r:tchc'] },
    { id: 'Tính lương', ten: 'Lập bảng lương', loai: 'tinhluong', nguoi: ['r:ktluong'] },
    { id: 'TCHC duyệt lương', ten: 'TCHC duyệt bảng lương', loai: 'kiemtra', nguoi: ['r:tchc'] },
    { id: 'Chờ duyệt', ten: 'Giám đốc phê duyệt', loai: 'pheduyet', nguoi: ['r:giamdoc'] }
  ];
  var QT_NGHI_MD = [
    { id: 'tbp', ten: 'Trưởng phòng / Chỉ huy trưởng duyệt', nguoi: ['tbp'], tren: 0, loai: [] },
    { id: 'tchc', ten: 'TCHC duyệt', nguoi: ['r:tchc'], tren: 0, loai: [] }
  ];
  var QT_YKIEN_MD = [
    { chuDe: 'Ngày công, nghỉ, làm thêm giờ', nguoi: ['r:tchc'] },
    { chuDe: 'Lương, phụ cấp, thưởng', nguoi: ['r:tchc', 'r:ktluong'] },
    { chuDe: 'Thuế TNCN, bảo hiểm', nguoi: ['r:ktluong'] },
    { chuDe: 'Tạm ứng, công nợ', nguoi: ['r:ktcongno'] },
    { chuDe: 'Vi phạm, khấu trừ theo quy chế', nguoi: ['r:tchc'] },
    { chuDe: 'Khác', nguoi: ['r:tchc'] }
  ];
  var LOAI_VP = ['Giảm thưởng / phụ cấp theo quy chế', 'Bồi thường thiệt hại'];
  var YK_TT = ['Mới', 'Đang xử lý', 'Đã giải quyết', 'Không chấp nhận'];

  /* Nghỉ phép / nghỉ ốm / nghỉ không lương */
  var LOAI_NGHI = ['Nghỉ phép năm', 'Nghỉ việc riêng có lương', 'Nghỉ ốm (BHXH chi trả)', 'Nghỉ không lương'];
  var NGHI_COT = { 'Nghỉ phép năm': 'nghiHuongLuong', 'Nghỉ việc riêng có lương': 'nghiHuongLuong', 'Nghỉ ốm (BHXH chi trả)': 'nghiOm', 'Nghỉ không lương': 'nghiKhongLuong' };
  var NP_TT = ['Chờ duyệt', 'Đã duyệt', 'Từ chối', 'Đã hủy'];

  /* ===== CÔNG VĂN ĐẾN – ĐI & ĐỀ XUẤT NỘI BỘ (v3.6) =====
     Sổ công văn: văn thư đăng ký vào sổ (số đến / số đi tự tăng theo năm) → tải bản quét lên Drive
     → lãnh đạo bút phê, giao người xử lý kèm hạn → người xử lý báo cáo kết quả → văn thư lưu hồ sơ.
     Đề xuất nội bộ: mỗi loại đề xuất có trình tự duyệt riêng (cấu hình ở "Quy trình & phê duyệt"),
     kèm người theo dõi và người thực hiện; nội dung, ảnh, tệp đính kèm lưu trên Drive Công ty. */
  var CV_HUONG = { den: 'Công văn đến', di: 'Công văn đi' };
  var CV_LOAI = ['Công văn', 'Quyết định', 'Thông báo', 'Tờ trình', 'Báo cáo', 'Biên bản', 'Hợp đồng', 'Hồ sơ thầu', 'Thư mời', 'Khác'];
  var CV_KHAN = ['Thường', 'Khẩn', 'Hỏa tốc'];
  var CV_TT_DEN = ['Mới đăng ký', 'Đã giao xử lý', 'Đang xử lý', 'Đã xử lý', 'Lưu hồ sơ'];
  var CV_TT_DI = ['Chờ ban hành', 'Đã ban hành', 'Lưu hồ sơ'];
  var CV_HET = 'Lưu hồ sơ';
  var CV_CS = ['Chưa trình', 'Chờ duyệt', 'Đã chia sẻ', 'Không chia sẻ'];   // duyệt chia sẻ văn bản cho người liên quan
  var CV_CS_OK = 'Đã chia sẻ';
  var TM_CV = '_CÔNG VĂN', TM_DX = '_ĐỀ XUẤT NỘI BỘ';
  var DX_TT = ['Nháp', 'Chờ duyệt', 'Đã duyệt', 'Đang thực hiện', 'Hoàn thành', 'Từ chối', 'Đã hủy'];
  var DX_MO = ['Nháp', 'Chờ duyệt'];               // người lập còn sửa, còn hủy được
  var DX_XONG = ['Hoàn thành', 'Từ chối', 'Đã hủy'];
  /* Nhóm loại đề xuất – Quản trị tạo, dùng để gom các loại đề xuất trên menu và bộ lọc. */
  var NHOM_DX_MD = [
    { id: 'taichinh', ten: 'Tài chính – thanh toán' },
    { id: 'vattu', ten: 'Vật tư – thiết bị' },
    { id: 'nhansu', ten: 'Nhân sự – chế độ' },
    { id: 'thicong', ten: 'Thi công – kỹ thuật' },
    { id: 'khac', ten: 'Khác' }
  ];
  var DX_CACH = { trinhtu: 'Duyệt theo trình tự', motnguoi: 'Chỉ cần một người duyệt' };
  var DX_GD_BUOC = { id: 'gd', ten: 'Giám đốc phê duyệt', nguoi: ['r:giamdoc'], tren: 0 };
  /* Loại đề xuất mặc định – Công ty sửa lại theo quy chế của mình tại "Quy trình & phê duyệt".
     nhom    = mã nhóm loại đề xuất
     tienTo  = tiền tố mã đề xuất (TUL → TUL0001); để trống thì mã chỉ là số thứ tự
     cach    = 'trinhtu' (qua lần lượt từng cấp) | 'motnguoi' (một trong những người được giao duyệt là xong)
     gdCuoi  = Giám đốc luôn là cấp duyệt cuối cùng của loại này
     tren    = ngưỡng số tiền (đồng) để cấp đó tham gia; để 0 nghĩa là luôn phải qua cấp đó */
  var QT_DX_MD = [
    { id: 'vattu', nhom: 'vattu', tienTo: 'VT', ten: 'Đề nghị cấp vật tư, thiết bị', tien: true, cach: 'trinhtu', gdCuoi: true, hoatDong: true,
      buoc: [{ id: 'tbp', ten: 'Trưởng bộ phận / Chỉ huy trưởng duyệt', nguoi: ['tbp'], tren: 0 },
        { id: 'tchc', ten: 'TCHC soát xét', nguoi: ['r:tchc'], tren: 0 }],
      theoDoi: ['r:tchc'], thucHien: ['r:tchc'] },
    { id: 'thanhtoan', nhom: 'taichinh', tienTo: 'TT', ten: 'Đề nghị thanh toán, tạm ứng', tien: true, cach: 'trinhtu', gdCuoi: true, hoatDong: true,
      buoc: [{ id: 'tbp', ten: 'Trưởng bộ phận / Chỉ huy trưởng duyệt', nguoi: ['tbp'], tren: 0 },
        { id: 'ketoan', ten: 'Kế toán kiểm tra', nguoi: ['r:ktcongno'], tren: 0 }],
      theoDoi: ['r:ktcongno'], thucHien: ['r:ktcongno'] },
    { id: 'tamunglg', nhom: 'taichinh', tienTo: 'TUL', ten: 'Tạm ứng lương', tien: true, cach: 'trinhtu', gdCuoi: true, hoatDong: true,
      buoc: [{ id: 'tbp', ten: 'Trưởng bộ phận / Chỉ huy trưởng duyệt', nguoi: ['tbp'], tren: 0 },
        { id: 'ketoan', ten: 'Kế toán tiền lương kiểm tra', nguoi: ['r:ktluong'], tren: 0 }],
      theoDoi: ['r:ktluong'], thucHien: ['r:ktluong'] },
    { id: 'nhansu', nhom: 'nhansu', tienTo: 'NS', ten: 'Đề nghị nhân sự (tuyển dụng, điều động, bổ nhiệm)', tien: false, cach: 'trinhtu', gdCuoi: true, hoatDong: true,
      buoc: [{ id: 'tbp', ten: 'Trưởng bộ phận / Chỉ huy trưởng duyệt', nguoi: ['tbp'], tren: 0 },
        { id: 'tchc', ten: 'TCHC soát xét', nguoi: ['r:tchc'], tren: 0 }],
      theoDoi: ['r:tchc'], thucHien: ['r:tchc'] },
    { id: 'thicong', nhom: 'thicong', tienTo: 'TC', ten: 'Đề xuất biện pháp, khối lượng phát sinh', tien: false, cach: 'trinhtu', gdCuoi: true, hoatDong: true,
      buoc: [{ id: 'tbp', ten: 'Chỉ huy trưởng công trình duyệt', nguoi: ['tbp'], tren: 0 },
        { id: 'kt', ten: 'Phòng Kỹ thuật soát xét', nguoi: ['r:truongphong'], tren: 0 }],
      theoDoi: ['r:tchc'], thucHien: [] },
    { id: 'khac', nhom: 'khac', tienTo: '', ten: 'Đề xuất, kiến nghị khác', tien: false, cach: 'trinhtu', gdCuoi: true, hoatDong: true,
      buoc: [{ id: 'tbp', ten: 'Trưởng bộ phận / Chỉ huy trưởng duyệt', nguoi: ['tbp'], tren: 0 }],
      theoDoi: ['r:tchc'], thucHien: [] }
  ];
  function kiemTraNhomDX(ds) {
    if (!Array.isArray(ds) || !ds.length) return 'Cần ít nhất 1 nhóm loại đề xuất.';
    var ids = {};
    for (var i = 0; i < ds.length; i++) {
      var g = ds[i];
      if (!String(g.ten || '').trim()) return 'Nhóm thứ ' + (i + 1) + ' chưa đặt tên.';
      if (!g.id) return 'Nhóm "' + g.ten + '" thiếu mã.';
      if (ids[g.id]) return 'Trùng mã nhóm loại đề xuất "' + g.id + '".';
      ids[g.id] = 1;
    }
    return '';
  }
  function kiemTraLoaiDX(ds, nhom) {
    if (!Array.isArray(ds) || !ds.length) return 'Cần ít nhất 1 loại đề xuất.';
    var ids = {}, tts = {}, nh = {};
    (nhom || NHOM_DX_MD).forEach(function (g) { nh[g.id] = 1; });
    for (var i = 0; i < ds.length; i++) {
      var l = ds[i];
      if (!String(l.ten || '').trim()) return 'Loại đề xuất thứ ' + (i + 1) + ' chưa đặt tên.';
      if (!l.id) return 'Loại "' + l.ten + '" thiếu mã.';
      if (ids[l.id]) return 'Trùng mã loại đề xuất "' + l.id + '".';
      ids[l.id] = 1;
      if (!l.nhom || !nh[l.nhom]) return 'Loại "' + l.ten + '" chưa chọn nhóm loại đề xuất.';
      var tt = String(l.tienTo || '').trim();
      if (tt) {
        if (!/^[A-Z0-9]{1,6}$/.test(tt)) return 'Tiền tố của loại "' + l.ten + '" chỉ gồm 1–6 chữ in hoa hoặc số (vd TUL, TTTP).';
        if (tts[tt]) return 'Trùng tiền tố "' + tt + '" giữa hai loại đề xuất.';
        tts[tt] = 1;
      }
      if (l.cach && !DX_CACH[l.cach]) return 'Loại "' + l.ten + '" có cách duyệt không hợp lệ.';
      if (!Array.isArray(l.buoc) || (!l.buoc.length && !l.gdCuoi)) return 'Loại "' + l.ten + '" chưa có cấp duyệt nào.';
      for (var j = 0; j < l.buoc.length; j++) {
        var b = l.buoc[j];
        if (!String(b.ten || '').trim()) return 'Loại "' + l.ten + '": cấp duyệt thứ ' + (j + 1) + ' chưa đặt tên.';
        if (!b.nguoi || !b.nguoi.length) return 'Loại "' + l.ten + '" – cấp "' + b.ten + '" chưa có người duyệt.';
      }
    }
    return '';
  }

  /* Các cấp duyệt đầy đủ của một loại: cấp do Quản trị khai, cộng cấp Giám đốc chốt cuối
     (Giám đốc là người duyệt cao nhất; cấp này không bị xóa nhầm khi sửa cấu hình). */
  function dxBuocDay(l) {
    var b = (l && l.buoc || []).slice();
    if (l && l.gdCuoi !== false && !b.some(function (x) { return (x.nguoi || []).indexOf('r:giamdoc') >= 0; }))
      b.push({ id: DX_GD_BUOC.id, ten: DX_GD_BUOC.ten, nguoi: DX_GD_BUOC.nguoi.slice(), tren: 0, chot: true });
    return b;
  }

  function laNguoiXuLy(s, nguoi) {
    if (!s) return false;
    if (s.role === 'admin') return true;
    return (nguoi || []).some(function (a) { return a === 'r:' + s.role || a === 'u:' + s.username; });
  }
  /* các bước lương của một kỳ: ảnh chụp lưu trong kỳ, nếu chưa có thì theo cấu hình hiện hành */
  function buocLuong(cfg, k) {
    var snap = k && parseJSON(k.quyTrinh, null);
    return Array.isArray(snap) && snap.length ? snap : (cfg.quyTrinhLuong || QT_LUONG_MD);
  }
  function viTriBuoc(steps, id) { return id === BUOC_XONG ? steps.length : steps.map(function (b) { return b.id; }).indexOf(id); }
  function normBuoc(t, steps) {
    steps = steps || QT_LUONG_MD;
    if (t === BUOC_XONG) return t;
    if (t === 'Nháp' || !t) return steps[0].id;
    if (t === 'Đã chốt') return steps[steps.length - 1].id;
    return viTriBuoc(steps, t) >= 0 ? t : steps[0].id;
  }
  function buocTheoLoai(steps, loai) { return steps.filter(function (b) { return b.loai === loai; })[0]; }
  function kiemTraQTLuong(steps) {
    if (!Array.isArray(steps) || steps.length < 3) return 'Quy trình lương cần ít nhất 3 bước.';
    var dem = {}, ids = {};
    for (var i = 0; i < steps.length; i++) {
      var b = steps[i];
      if (!b.id || !String(b.ten || '').trim()) return 'Bước ' + (i + 1) + ' thiếu tên.';
      if (ids[b.id]) return 'Trùng mã bước "' + b.id + '".';
      ids[b.id] = 1;
      if (!LOAI_BUOC[b.loai]) return 'Bước "' + b.ten + '" chưa chọn loại việc.';
      if (!b.nguoi || !b.nguoi.length) return 'Bước "' + b.ten + '" chưa có người xử lý.';
      if (b.nguoi.indexOf('tbp') >= 0) return 'Bước lương không dùng "trưởng bộ phận của người đề nghị".';
      dem[b.loai] = (dem[b.loai] || 0) + 1;
    }
    if (steps[0].loai !== 'chamcong') return 'Bước đầu tiên phải là "Chấm công".';
    if (steps[steps.length - 1].loai !== 'pheduyet') return 'Bước cuối phải là "Phê duyệt cuối & khóa dữ liệu".';
    if (dem.chamcong !== 1 || dem.tinhluong !== 1 || dem.pheduyet !== 1) return 'Quy trình phải có đúng 1 bước Chấm công, 1 bước Tính lương, 1 bước Phê duyệt cuối.';
    if ((dem.congno || 0) > 1 || (dem.vipham || 0) > 1) return 'Mỗi loại "Công nợ", "Vi phạm" chỉ có tối đa 1 bước.';
    var loais = steps.map(function (b) { return b.loai; }), itl = loais.indexOf('tinhluong');
    if (itl < 1) return 'Bước "Tính lương" phải sau "Chấm công".';
    return '';
  }

  var SCHEMA = {
    Users: ['username', 'passHash', 'salt', 'role', 'maNV', 'hoTen', 'active', 'mustChange', 'lastLogin', 'createdAt', 'boPhan', 'khoaTu', 'khoaLyDo'],
    NhanVien: ['maNV', 'hoTen', 'gioiTinh', 'ngaySinh', 'cccd', 'ngayCapCCCD', 'noiCapCCCD', 'dienThoai', 'email',
      'diaChi', 'queQuan', 'trinhDo', 'chuyenNganh', 'phongBan', 'chucVu', 'donVi', 'noiLamViec', 'ngayVao', 'loaiHD', 'soHD',
      'ngayKyHD', 'ngayHetHanHD', 'trangThai', 'ngayNghi', 'ngach', 'bac', 'ngayHuongBac', 'ngayPhepNam', 'luongChinh', 'luongBH', 'phuCap', 'anCa',
      'thamGiaBH', 'doanVien', 'mst', 'soNPT', 'camKet08', 'nganHang', 'soTK', 'ghiChu', 'capNhat', 'loaiLuong', 'lyDoKhongBH', 'congTy', 'thueTT', 'hsThue', 'hoSoHuu', 'lichSuHD'],
    QuaTrinh: ['id', 'maNV', 'loai', 'tuNgay', 'denNgay', 'phongBan', 'chucVu', 'donVi', 'soQD', 'nguoiKy',
      'noiDung', 'taoBoi', 'taoLuc'],
    LenLuong: ['id', 'maNV', 'loai', 'hieuLuc', 'ngachCu', 'bacCu', 'luongCu', 'luongBHCu', 'ngachMoi', 'bacMoi',
      'luongMoi', 'luongBHMoi', 'lyDo', 'soQD', 'trangThai', 'apDung', 'nguoiDeXuat', 'ngayDeXuat', 'nguoiDuyet',
      'ngayDuyet', 'ghiChuDuyet'],
    KhenThuong: ['id', 'maNV', 'loai', 'hinhThuc', 'ngay', 'soTien', 'chiuThue', 'kyLuong', 'noiDung', 'soQD',
      'trangThai', 'nguoiDeXuat', 'ngayDeXuat', 'nguoiDuyet', 'ngayDuyet'],
    ChamCong: ['id', 'ky', 'maNV', 'ngayCong', 'nghiHuongLuong', 'nghiKhongLuong', 'nghiOm', 'ot150', 'ot200', 'ot300',
      'tamUng', 'khauTruKhac', 'truyLinh', 'ghiChu', 'ghiChuCongNo', 'chiTietPhat', 'ghiChuTCHC', 'ngayLe', 'bienKy', 'noiLamViec', 'chiTietNgay', 'chiTietTamUng', 'tienKhoan', 'chiTietOT', 'chiTietAn', 'soBua', 'tienAnThu', 'thuUng', 'anThuTay'],
    BangLuong: ['id', 'ky', 'maNV', 'hoTen', 'phongBan', 'chucVu', 'luongChinh', 'ngayCongChuan', 'ngayCongHL',
      'luongThoiGian', 'phuCap', 'phuCapChiTiet', 'chiTiet', 'ktCongThuc', 'tienOT', 'thuong', 'truyLinh', 'viPham', 'giamThuong', 'tongThuNhap', 'luongDongBH',
      'bhxh', 'bhyt', 'bhtn', 'doanPhi', 'tnChiuThue', 'giamTru', 'tnTinhThue', 'thueTNCN', 'boiThuong', 'tamUng',
      'khauTruKhac', 'thucLinh', 'bhxhDN', 'bhytDN', 'bhtnDN', 'kpcdDN', 'ghiChu', 'ctTamUng', 'ctViPham', 'ctCong', 'congTy', 'cheDoLuat', 'cachThue', 'traThayBHCT', 'soBua', 'tienAnThu', 'thuUng', 'tongNo', 'anCaMien', 'otThue', 'otThucLinh', 'ctHD'],
    KyLuong: ['ky', 'trangThai', 'ngayCongChuan', 'nguoiTinh', 'ngayTinh', 'nguoiDuyet', 'ngayDuyet', 'soNguoi',
      'tongThuNhap', 'tongThucLinh', 'maKiemTra', 'quyTrinh', 'congThuc', 'capNhat'],
    MucHuong: ['id', 'maNV', 'hieuLuc', 'giaTri', 'cu', 'soQD', 'lyDo', 'trangThai', 'nguoiTao', 'taoLuc', 'nguoiDuyet', 'ngayDuyet', 'ghiChuDuyet'],
    NghiPhep: ['id', 'maNV', 'hoTen', 'loai', 'tuNgay', 'denNgay', 'buoi', 'soNgay', 'lyDo', 'chungTu', 'trangThai',
      'nguoiTao', 'taoLuc', 'luongDuyet', 'capDuyet', 'lichSu', 'nguoiDuyet', 'ngayDuyet', 'ghiChuDuyet'],
    TaiLieu: ['id', 'maNV', 'loai', 'ten', 'soHieu', 'ngay', 'hetHan', 'lienKet', 'nvXem', 'trangThai', 'tenFile', 'mimeType',
      'kichThuoc', 'fileId', 'nguoiTai', 'taiLuc', 'nguoiXacNhan', 'ghiChu', 'ocrId', 'ocrLuc', 'ocrNguon', 'ocrTom'],
    YKien: ['id', 'ky', 'maNV', 'hoTen', 'chuDe', 'xuLy', 'boPhan', 'trangThai', 'noiDung', 'thread', 'taoLuc', 'capNhat', 'nguoiXuLy'],
    BaiViet: ['id', 'loai', 'tieuDe', 'noiDung', 'anh', 'dinhKem', 'doiTuong', 'ghim', 'xacNhan', 'soVanBan', 'hieuLuc', 'trangThai',
      'tacGia', 'tenTacGia', 'taoLuc', 'dangLuc', 'capNhat', 'noiBat', 'noiBatDen'],
    TuongTac: ['id', 'baiId', 'loai', 'username', 'hoTen', 'noiDung', 'thoiGian'],
    HoiThoai: ['id', 'loai', 'ten', 'mau', 'moTa', 'taoBoi', 'taoLuc', 'capNhat', 'tinCuoi', 'trangThai', 'thuMuc', 'luuTru'],
    ThanhVien: ['id', 'hoiThoai', 'username', 'vaiTro', 'daDoc', 'ghim', 'tatTB', 'thamGia', 'roi'],
    TinNhan: ['id', 'hoiThoai', 'username', 'hoTen', 'loai', 'noiDung', 'tep', 'traLoi', 'thoiGian'],
    ChatMuc: ['id', 'hoiThoai', 'loai', 'tieuDe', 'noiDung', 'cauHinh', 'taoBoi', 'taoLuc', 'trangThai', 'ketTiep', 'lanCuoi', 'capNhat', 'nhacHan'],
    ChatTraLoi: ['id', 'mucId', 'ky', 'username', 'giaTri', 'thoiGian'],
    ChamCongGPS: ['id', 'maNV', 'username', 'hoTen', 'ngay', 'gio', 'loai', 'diemId', 'tenDiem', 'boPhan', 'lat', 'lng', 'doChinhXac', 'khoangCach', 'banKinh', 'anh', 'khuonMat', 'diemKM', 'thietBi', 'trangThai', 'lyDo', 'nguoiDuyet', 'duyetLuc', 'ghiChu', 'thoiGian', 'tre', 'congTac'],
    KhuonMat: ['maNV', 'username', 'hoTen', 'mau', 'anh', 'thietBi', 'dangKyLuc', 'trangThai', 'xoaBoi', 'xoaLuc'],
    CongTac: ['id', 'maNV', 'username', 'hoTen', 'boPhan', 'tuNgay', 'denNgay', 'cheDo', 'diem', 'noiDen', 'mucDich', 'phuongTien', 'trangThai', 'nguoiTao', 'nguoiDuyet', 'duyetLuc', 'lyDo', 'taoLuc', 'capNhat'],
    ThietBiPush: ['id', 'username', 'token', 'thietBi', 'tenMay', 'tuyChon', 'dangKy', 'capNhat'],
    KiemTraCC: ['id', 'tu', 'den', 'pham', 'taoBoi', 'taoLuc', 'soNguoi', 'soLuot', 'soDo', 'soVang', 'thongKe', 'aiNhanXet', 'aiLuc'],
    // v3.3: sheet "Kiểm tra công" – vị trí ngoài phạm vi chấm công phát hiện khi nhắn tin (8 cột đầu theo yêu cầu, sau đó là số liệu chi tiết)
    // v3.4: đánh giá tổng hợp CBCNV (chấm công + vị trí + nhóm công việc + AI) – đợt và từng người
    DanhGiaTH: ['id', 'tieuDe', 'tu', 'den', 'pham', 'taoBoi', 'taoLuc', 'soNguoi', 'thongKe', 'aiTong', 'aiLuc', 'docId', 'docUrl', 'pdfId', 'trangThai'],
    DanhGiaCN: ['id', 'dotId', 'maNV', 'hoTen', 'boPhan', 'chucDanh', 'username', 'chiSo', 'diemCC', 'diemTT', 'diemTN', 'diemTong', 'xepLoaiHT', 'aiXepLoai', 'aiTongQuan', 'aiTichCuc', 'aiHanChe', 'aiThieuSot', 'aiDeXuat', 'aiLuc',
      'xepLoai', 'ketLuan', 'nguoiChot', 'chotLuc', 'congBo', 'nvPhanHoi', 'phLuc', 'capNhat'],
    KiemTraCong: ['id', 'ngayGhi', 'gioGhi', 'hoTen', 'maNV', 'viTriChamCong', 'viTriNgoaiPhamVi', 'khoangCach', 'gioDo', 'boPhan', 'username', 'diemCC', 'latCC', 'lngCC', 'banKinh', 'lat', 'lng', 'doChinhXac', 'vuotBanKinh', 'banDo', 'hoiThoai', 'baoCao', 'ghiChu'],
    CanhBaoCC: ['id', 'khoa', 'dotId', 'maNV', 'hoTen', 'boPhan', 'ngay', 'nhom', 'ma', 'mucDo', 'moTa', 'luotIds', 'trangThai', 'nhacLuc', 'nguoiNhac', 'hanGT', 'giaiTrinh', 'ketQua', 'nguoiGT', 'gtLuc', 'xuLy', 'nguoiXL', 'xlLuc', 'taoLuc', 'capNhat'],
    PhieuKho: ['id', 'hoiThoai', 'soPhieu', 'loai', 'ngay', 'kho', 'doiTac', 'hangMuc', 'nguoiGiaoNhan', 'soCT', 'dong', 'anh', 'ghiChu', 'nguoiLap', 'hoTenLap', 'taoLuc', 'msgId', 'trangThai', 'lyDoHuy', 'nguoiHuy', 'capNhat'],
    TepChat: ['fileId', 'hoiThoai', 'ten', 'mime', 'kichThuoc', 'nguoiTai', 'thoiGian', 'msgId', 'seq', 'w', 'h', 'thuMuc', 'thumbId'],
    PheDuyet: ['id', 'ky', 'tu', 'den', 'tuTen', 'denTen', 'hanhDong', 'nguoi', 'vaiTro', 'thoiGian', 'ghiChu'],
    BotQuyDinh: ['id', 'nhom', 'tieuDe', 'noiDung', 'lienKet', 'vanBan', 'hieuLuc', 'trangThai', 'nguoiSua', 'capNhat'],
    BotVanBan: ['id', 'ten', 'fileId', 'mime', 'nhom', 'soHieu', 'ngayBanHanh', 'trangThai', 'soChu', 'textLuc', 'loi', 'nguoiSua', 'capNhat'],
    BotHoi: ['id', 'thoiGian', 'username', 'hoTen', 'cauHoi', 'traLoi', 'nguon', 'nhom', 'daTraLoi', 'dungAI', 'thieu', 'huuIch', 'phanHoi'],
    CongVan: ['id', 'huong', 'nam', 'so', 'soHieu', 'ngayVB', 'ngayVaoSo', 'noiGui', 'noiNhan', 'loai', 'doKhan', 'trichYeu',
      'nguoiKy', 'soBan', 'fileId', 'tenFile', 'mime', 'kichThuoc', 'butPhe', 'nguoiXuLy', 'phoiHop', 'hanXuLy', 'trangThai',
      'ketQua', 'hoanThanhLuc', 'lienKet', 'ghiChu', 'nguoiNhap', 'taoLuc', 'capNhat',
      'hanTraLoi', 'tomTat', 'aiLuc', 'aiTinCay', 'csTrangThai', 'csDeXuat', 'nguoiDoc', 'nguoiDuyetCS', 'csLuc', 'csGhiChu', 'soTN', 'noiDungOCR'],
    CongVanXL: ['id', 'cvId', 'hanhDong', 'nguoi', 'tenNguoi', 'noiDung', 'hanXuLy', 'fileId', 'tenTep', 'thoiGian'],
    CongVanTN: ['id', 'cvId', 'username', 'hoTen', 'noiDung', 'fileId', 'tenTep', 'mime', 'thoiGian'],
    DeXuat: ['id', 'ma', 'tienTo', 'so', 'loai', 'tenLoai', 'nhom', 'tenNhom', 'tieuDe', 'noiDung', 'soTien', 'donVi', 'maNV', 'hoTen', 'boPhan', 'doKhan',
      'hanCan', 'trangThai', 'capDuyet', 'luongDuyet', 'theoDoi', 'thucHien', 'lichSu', 'noiDungId', 'ketQua', 'hoanThanhLuc',
      'nguoiDuyet', 'ngayDuyet', 'ghiChuDuyet', 'nguoiTao', 'taoLuc', 'capNhat'],
    DeXuatTep: ['id', 'dxId', 'loai', 'ten', 'mime', 'kichThuoc', 'fileId', 'nguoiTai', 'taiLuc'],
    CauHinh: ['key', 'value'],
    NhatKy: ['thoiGian', 'user', 'action', 'chiTiet'],
    HopDongKhoan: ['id', 'soHD', 'maNV', 'loaiK', 'congTrinh', 'hangMuc', 'doiTo', 'tuNgay', 'denNgay', 'giaTriHD', 'hinhThuc', 'donGia', 'donViTinh', 'kieuGia', 'danhGia', 'ketLuan', 'trangThai', 'hoSo', 'ghiChu', 'nguoiTao', 'taoLuc', 'capNhat'],
    ThanhToanKhoan: ['id', 'hdId', 'maNV', 'congTrinh', 'dot', 'loai', 'kyKL', 'ngayNT', 'ngayTra', 'khoiLuong', 'giaTriNT', 'thuNhapBS', 'giamTruHD', 'tamUng', 'tuTinhThue', 'thuHoiTU', 'thuHoiTUThue', 'khauTruKhac', 'tienPhaiTra', 'tnChiuThue', 'thueSuat', 'thue', 'thucTra', 'cheDoThue', 'congThuc', 'buoc', 'trangThai', 'lichSu', 'ghiChu', 'nguoiTao', 'taoLuc', 'capNhat'],
    HoSoThue: ['id', 'maNV', 'nam', 'loai', 'mauSo', 'phienBan', 'ngayKy', 'trangThai', 'duLieu', 'ketQua', 'nguoiLap', 'ngayLap', 'nguoiKiemTra', 'ngayKiemTra', 'ghiChuDuyet', 'taiLieuId', 'tenTep', 'ghiChu', 'capNhat']
  };

  var NUM_COLS = ['luongChinh', 'luongBH', 'soNPT', 'luongCu', 'luongBHCu', 'luongMoi', 'luongBHMoi', 'soTien',
    'ngayCong', 'nghiHuongLuong', 'nghiKhongLuong', 'ot150', 'ot200', 'ot300', 'tamUng', 'khauTruKhac',
    'truyLinh', 'ngayCongChuan', 'ngayCongHL', 'luongThoiGian', 'phuCap', 'tienOT', 'thuong', 'tongThuNhap',
    'luongDongBH', 'bhxh', 'bhyt', 'bhtn', 'doanPhi', 'tnChiuThue', 'giamTru', 'tnTinhThue', 'thueTNCN',
    'thucLinh', 'bhxhDN', 'bhytDN', 'bhtnDN', 'kpcdDN', 'viPham', 'giamThuong', 'boiThuong', 'soNguoi', 'tongThucLinh', 'nghiOm', 'soNgay', 'ngayPhepNam', 'capDuyet', 'kichThuoc', 'ngayLe', 'ktCongThuc', 'anCa', 'daDoc', 'seq', 'w', 'h', 'so', 'soBan', 'soTien', 'soTN', 'tienKhoan', 'giaTriHD', 'donGia', 'khoiLuong', 'giaTriNT', 'thuNhapBS', 'giamTruHD', 'thuHoiTU', 'thuHoiTUThue', 'tienPhaiTra', 'thueSuat', 'thue', 'thucTra', 'buoc', 'dot', 'soBua', 'tienAnThu', 'thuUng', 'tongNo', 'anCaMien', 'otThue', 'otThucLinh', 'anThuTay'];

  // Bảng NhanVien.phuCap là JSON nhưng nằm trong NUM_COLS? Không: BangLuong.phuCap là số. Phân biệt theo bảng.
  function isNumCol(table, col) {
    if (col === 'phuCap') return table === 'BangLuong';
    return NUM_COLS.indexOf(col) >= 0;
  }

  var DEFAULT_CFG = {
    congTy: 'CÔNG TY CỔ PHẦN TAPACO VIỆT NAM',
    phanQuyen: {},                                  // v3.8: bảng phân quyền tích chọn (rỗng = mặc định)
    diaChiCongTy: 'Hà Nội',
    mstCongTy: '',
    luongCoSo: 2530000,              // NĐ 161/2026/NĐ-CP, từ 01/7/2026
    luongToiThieuVung: 5310000,      // NĐ 293/2025/NĐ-CP, vùng I, từ 01/01/2026
    heSoTranBHXH: 20,                // trần BHXH, BHYT = 20 × lương cơ sở
    heSoTranBHTN: 20,                // trần BHTN = 20 × lương tối thiểu vùng
    tyLeNLD: { bhxh: 8, bhyt: 1.5, bhtn: 1 },
    tyLeDN: { bhxh: 17.5, bhyt: 3, bhtn: 1, kpcd: 2, tnld: 0.5, tnldHuu: 0 }, // tnldHuu: BH TNLĐ-BNN DN đóng cho người đã nghỉ hưu (mặc định 0 – sửa nếu cơ quan BHXH yêu cầu) // tnld: BH TNLĐ-BNN DN đóng cho người đã đóng BHXH nơi khác (đã nằm trong 17,5% của người đóng đủ)
    doanPhi: 1,
    giamTruBanThan: 15500000,        // NQ 110/2025/UBTVQH15, kỳ tính thuế 2026
    giamTruNPT: 6200000,
    bieuThue: [[10000000, 5], [30000000, 10], [60000000, 20], [100000000, 30], [0, 35]], // Luật 109/2025/QH15
    // v3.25: QUY ĐỊNH THEO THỜI KỲ – mỗi kỳ lương tự áp mốc có "tu" (YYYY-MM) gần nhất ≤ kỳ; ghi đè lương cơ sở, LTTV, giảm trừ, biểu thuế, cách miễn thuế làm thêm giờ, ngưỡng khấu trừ 10%
    mocLuat: null,                    // null = MOC_LUAT_MD
    luat109Tu2026: false,             // true: kỳ 01–06/2026 áp biểu 5 bậc & miễn thuế toàn bộ tiền làm thêm giờ (Luật 109/2025 Điều 29.2, NĐ 253/2026: "áp dụng từ kỳ tính thuế năm 2026")
    ngayNghiTuan: 'cn',
    mauPhieu: null, mauBangLuong: null,   // v3.41: mẫu phiếu lương / bảng lương do người dùng sắp xếp (null = mặc định)
    cheDoBangLuong: 'gop',            // v3.32: 'gop' – 1 bảng lương gồm cả tiền làm thêm giờ; 'tach' – 2 bảng độc lập: lương theo công thời gian / lương tăng ca, thêm giờ
    anCaMienTheoCong: true, anCaMienChia: 22, anCaMienCT: '',   // v3.39: anCaMienCT = công thức mở (trống = theo 2 khóa cũ)
      // v3.32: ăn ca miễn thuế = MIN(mức tối đa, ngày công thực tế × mức tối đa ÷ 22)
    quyDinh: null,                    // v3.27: cơ sở dữ liệu quy định pháp luật (null = QUY_DINH_MD) – sửa ở Cài đặt → Quy định & biểu mẫu
    bieuMau: null,                    // v3.27: kho biểu mẫu thuế (null = BIEU_MAU_MD)
    ckBatBuocDuyet: true,             // v3.28: bảng lương chỉ tạm không khấu trừ 10% khi cam kết của năm đã ký, Phụ trách KT duyệt và có MST
    quyTrinhKhoan: null,              // v3.28: 6 bước duyệt thanh toán khoán (null = QT_KHOAN_MD)
    traThayBH: 'khong',               // v3.26: Điều 168 khoản 3 BLLĐ – trả thêm cho NLĐ không thuộc diện BH bắt buộc (người nghỉ hưu, HĐ thử việc) khoản tương đương mức DN đóng BH: 'khong' | 'co'               // ngày nghỉ hằng tuần để xếp giờ làm thêm 200%: 'cn' | 't7cn'
    thueVangLai: 10,
    nguongVangLai: 5000000,          // NĐ 253/2026/NĐ-CP Điều 50: khấu trừ 10% khi trả từ 5 triệu đồng/lần trở lên
    cachTinhCongChuan: 'codinh',      // 'codinh' | 'truCN'
    ngayCongChuan: 26,
    gioCongNgay: 8,
    otGioiHanThang: 40,              // giới hạn giờ làm thêm được miễn thuế (Điều 107 BLLĐ 2019)
    otGioiHanNam: 200,               // 200 giờ/năm; ngành, công việc được làm đến 300 giờ thì đổi thành 300
    heSoOT: { ot150: 1.5, ot200: 2, ot300: 3 },
    nghiKLToiDaBH: 14,
    chuKyNangBac: 24,
    canhBaoTruoc: 45,
    canhBaoHD: 30,
    ocrTuDong: true,                  // v3.37: tự đọc chữ (OCR) ảnh / bản scan PDF khi tải lên hồ sơ                    // v3.33: báo TCHC trước khi HĐLĐ hết hạn (ngày)
    matKhauMacDinh: 'Tapaco@123',
    ngayPhepCoBan: 12,                // Điều 113 BLLĐ 2019 (công việc bình thường)
    thamNienCongPhep: 5,              // Điều 114: cứ đủ 5 năm làm việc được thêm 1 ngày phép
    quyTrinhLuong: QT_LUONG_MD,       // các bước lập & duyệt lương
    quyTrinhNghi: QT_NGHI_MD,         // các cấp duyệt đơn nghỉ (áp dụng khi số ngày > tren, cho các loại nghỉ trong 'loai', trống = mọi loại)
    quyTrinhYKien: QT_YKIEN_MD,       // chủ đề ý kiến về lương → người tiếp nhận
    nhomDeXuat: NHOM_DX_MD,           // nhóm loại đề xuất (Quản trị tạo)
    loaiDeXuat: QT_DX_MD,             // loại đề xuất nội bộ: nhóm, tiền tố mã, trình tự duyệt, người theo dõi, người thực hiện
    truongBoPhan: {},                 // {tên phòng ban / đơn vị: tên đăng nhập trưởng}
    congThucLuong: null,              // null = dùng mẫu CT_LUONG_MD
    bienKy: null,
    congThucDongBH: '',
    loaiTaiLieu: ['Hợp đồng lao động', 'Phụ lục hợp đồng', 'Căn cước công dân', 'Mã số thuế', 'Quyết định', 'Sổ BHXH',
      'Bằng cấp – chứng chỉ', 'Sơ yếu lý lịch', 'Giấy khám sức khỏe', 'Hồ sơ người phụ thuộc', 'Hồ sơ thuế TNCN', 'Hồ sơ khoán việc', 'Khác'],
    taiLieuBatBuoc: ['Hợp đồng lao động', 'Căn cước công dân'],
    loaiNVTuTai: ['Căn cước công dân', 'Mã số thuế', 'Bằng cấp – chứng chỉ', 'Sơ yếu lý lịch', 'Giấy khám sức khỏe', 'Hồ sơ người phụ thuộc', 'Khác'],
    dungLuongToiDaMB: 10,
    noiLamViec: [                     // nơi làm việc: quyết định công chuẩn & tiền ăn ca/ngày (theo bảng lương TAPACO 2026)
      { ma: 'VP', ten: 'Văn phòng', congChuan: 24, anCa: 30000 },
      { ma: 'DA', ten: 'Dự án – công trường', congChuan: 28, anCa: 70000 }
    ],
    ngayLe: [],                       // các ngày nghỉ lễ, Tết hưởng lương (yyyy-mm-dd) – suy ra từ ngayLeDs
    ngayLeDs: [],                     // [{ngay, ten, loai: Lễ, Tết | Nghỉ bù | Công ty cho nghỉ}] – màn hình Lịch nghỉ lễ
    chuNhatTinh: 'cong',              // bảng chấm công ngày: làm Chủ nhật tính 'cong' (công thường) | 'ot' (giờ làm ngày nghỉ tuần 200%)
    loaiTamUng: ['Ứng lương', 'Ứng tiền ăn', 'Ứng xăng xe – đi lại', 'Ứng khác'],
    tyLeBoiThuongToiDa: 30,           // Điều 102 BLLĐ 2019: khấu trừ ≤ 30% lương thực trả sau BH, thuế
    danhMucViPham: [],                // [{ma, ten, mucTien, loai}] theo Nội quy / Quy chế lương thưởng của Công ty
    phuCap: [
      { ma: 'chucvu', ten: 'PC chức vụ', dongBH: true, theoCong: false, mienThue: 0 },
      { ma: 'trachnhiem', ten: 'PC trách nhiệm', dongBH: true, theoCong: false, mienThue: 0 },
      { ma: 'congtruong', ten: 'PC công trường / lưu động', dongBH: false, theoCong: true, mienThue: 0 },
      { ma: 'anca', ten: 'Tiền ăn giữa ca', dongBH: false, theoCong: true, mienThue: 730000 },
      { ma: 'dienthoai', ten: 'PC điện thoại', dongBH: false, theoCong: false, mienThue: 0 },
      { ma: 'xangxe', ten: 'PC xăng xe', dongBH: false, theoCong: false, mienThue: 0 },
      { ma: 'nhao', ten: 'PC nhà ở', dongBH: false, theoCong: false, mienThue: 0 },
      { ma: 'kpi', ten: 'Mức thưởng KPI (khi đạt 100%)', dongBH: false, theoCong: false, mienThue: 0 },
      { ma: 'chuyencan', ten: 'Tiền chuyên cần', dongBH: false, theoCong: false, mienThue: 0 }
    ],
    // v3.20: công ty thành viên – CBCNV khai báo thuộc công ty nào để hạch toán BHXH, bảng lương, thuế TNCN riêng.
    // Mã số thuế, địa chỉ để trống: Anh/Chị điền ở "Phòng ban & công trình → Công ty thành viên".
    congTyTV: [
      { ma: 'TPC', ten: 'CÔNG TY CỔ PHẦN TAPACO VIỆT NAM', tenTat: 'TAPACO', mst: '', diaChi: '', trangThai: 'Hoạt động' },
      { ma: 'TC', ten: 'Thịnh Vượng Toàn Cầu', tenTat: 'Thịnh Vượng Toàn Cầu', mst: '', diaChi: '', trangThai: 'Hoạt động' },
      { ma: 'ATP', ten: 'An Thuận Phát', tenTat: 'An Thuận Phát', mst: '', diaChi: '', trangThai: 'Hoạt động' }
    ],
    boPhan: [                         // phòng ban & ban chỉ huy công trình; CBCNV được điều động theo từng giai đoạn
      { ma: 'BGD', ten: 'Ban Giám đốc', loai: 'Phòng ban', trangThai: 'Hoạt động' },
      { ma: 'TCHC', ten: 'Phòng Tổ chức – Hành chính', loai: 'Phòng ban', trangThai: 'Hoạt động' },
      { ma: 'TCKT', ten: 'Phòng Tài chính – Kế toán', loai: 'Phòng ban', trangThai: 'Hoạt động' },
      { ma: 'KT', ten: 'Phòng Kỹ thuật', loai: 'Phòng ban', trangThai: 'Hoạt động' },
      { ma: 'KTKH', ten: 'Phòng Kinh tế – Kế hoạch', loai: 'Phòng ban', trangThai: 'Hoạt động' }
    ],
    phongBan: [],                     // tự lấy từ boPhan (giữ để tương thích)
    chucVu: ['Chủ tịch HĐQT', 'Tổng Giám đốc', 'Phó Tổng Giám đốc', 'Kế toán trưởng', 'Trưởng phòng', 'Phó phòng',
      'Chỉ huy trưởng', 'Kỹ sư', 'Kế toán viên', 'Chuyên viên', 'Nhân viên', 'Công nhân', 'Lái xe', 'Thủ kho'],
    donVi: ['Văn phòng Công ty'],
    gps: { bat: true, doChinhXac: 50, hieuChuanGiay: 25, nguongKM: 0.5, gioVao: '07:30', gioRa: '17:00', phutTre: 10, gioDuCong: 6, gioLamDuCong: 8, nghiTruaTu: '12:00', nghiTruaDen: '13:30', khoaThietBi: true, batBuocKM: true, nhanAnhNgoaiVung: true, banKinhCT: 300, banDo: '' },
    // v3.3: kiểm tra vị trí khi nhắn tin – ngoài phạm vi điểm chấm công thì ghi sheet KiemTraCong, báo TCHC, Ban Giám đốc, Kế toán
    viTriTN: { bat: true, chiGioLam: true, chuKyPhut: 15, doChinhXacToiDa: 100, dungSai: 50, baoLaiPhut: 60, boQuaVaiTro: ['admin', 'giamdoc'] },
    push: { bat: false, firebase: {}, vapid: '' },   // thông báo đẩy (Firebase Cloud Messaging); khóa bí mật lưu trong Thuộc tính tập lệnh
    diemCC: {},                       // điểm chấm công GPS theo mã bộ phận: { CT01: [{id, ten, lat, lng, banKinh, gioVao?, gioRa?}], _chung: [...] }
    thangLuong: []                    // [{ngach, bac, muc}]
  };

  var CAM_XUC = ['👍', '❤️', '😆', '😮', '😢', '🙏'];
  var BT_LOAI = ['Thông báo', 'Chính sách – Quy định', 'Tin nội bộ', 'Truyền thông – Hoạt động', 'Vinh danh – Khen thưởng'];
  var LOAI_HD = ['Không xác định thời hạn', 'Xác định thời hạn', 'Thử việc', 'HĐ dưới 3 tháng', 'Thời vụ / Khoán việc', 'Cộng tác viên (CTV)'];
  var HD_VANGLAI = ['Thử việc', 'HĐ dưới 3 tháng', 'Thời vụ / Khoán việc', 'Cộng tác viên (CTV)'];
  var TRANG_THAI_NV = ['Đang làm việc', 'Thử việc', 'Tạm nghỉ', 'Đã nghỉ việc'];
  var LOAI_QT = ['Tiếp nhận', 'Điều động', 'Bổ nhiệm', 'Miễn nhiệm', 'Thay đổi HĐLĐ', 'Đào tạo', 'Tạm nghỉ', 'Nghỉ việc', 'Thanh lý hợp đồng', 'Khác'];
  var QT_CHAM_DUT = ['Nghỉ việc', 'Thanh lý hợp đồng'];      // v3.18: chấm dứt HĐLĐ – từ ngày này không còn làm việc, tài khoản tự khóa
  var LOAI_LL = ['Nâng bậc định kỳ', 'Nâng lương trước hạn', 'Điều chỉnh lương', 'Giảm lương'];
  var LOAI_KT = ['Thưởng tháng', 'Thưởng lễ – Tết', 'Thưởng công trình / dự án', 'Lương tháng 13', 'Thưởng đột xuất', 'Kỷ luật'];
  var HT_KYLUAT = ['Khiển trách', 'Kéo dài thời hạn nâng lương ≤ 6 tháng', 'Cách chức', 'Sa thải'];

  /* ---------------- tiện ích ---------------- */
  function num(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var s = String(v).trim().replace(/\s/g, '');
    if (/^-?[\d.]+,\d+$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');   // 1.234,5
    else if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');          // 1.234.567
    else s = s.replace(/,/g, '');
    var n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }
  function R(x) { return Math.round(x); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function parseD(s) { var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || ''); return m ? { y: +m[1], m: +m[2], d: +m[3] } : null; }
  function fmtD(y, m, d) { return y + '-' + pad(m) + '-' + pad(d); }
  function dim(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }
  function addMonths(s, n) {
    var p = parseD(s); if (!p) return '';
    var t = p.y * 12 + (p.m - 1) + n, y = Math.floor(t / 12), m = t % 12 + 1;
    return fmtD(y, m, Math.min(p.d, dim(y, m)));
  }
  function addDays(s, n) {
    var p = parseD(s); if (!p) return '';
    var d = new Date(Date.UTC(p.y, p.m - 1, p.d + n));
    return fmtD(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
  }
  function kyStart(ky) { return ky + '-01'; }
  function kyEnd(ky) { var a = ky.split('-'); return fmtD(+a[0], +a[1], dim(+a[0], +a[1])); }
  function validKy(ky) { return /^\d{4}-(0[1-9]|1[0-2])$/.test(ky || ''); }
  function congChuanTruCN(ky) {
    var a = ky.split('-'), y = +a[0], m = +a[1], n = dim(y, m), c = 0;
    for (var d = 1; d <= n; d++) if (new Date(Date.UTC(y, m - 1, d)).getUTCDay() !== 0) c++;
    return c;
  }
  function parseJSON(s, dflt) {
    if (s && typeof s === 'object') return s;
    try { return s ? JSON.parse(s) : dflt; } catch (e) { return dflt; }
  }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  /* ---- âm lịch (thuật toán Hồ Ngọc Đức, múi giờ +7) → ngày lễ theo âm lịch ---- */
  var AL = (function () {
    var INT = Math.floor, PI = Math.PI;
    function jdFromDate(dd, mm, yy) { var a = INT((14 - mm) / 12), y = yy + 4800 - a, m = mm + 12 * a - 3; var jd = dd + INT((153 * m + 2) / 5) + 365 * y + INT(y / 4) - INT(y / 100) + INT(y / 400) - 32045; if (jd < 2299161) jd = dd + INT((153 * m + 2) / 5) + 365 * y + INT(y / 4) - 32083; return jd; }
    function jdToDate(jd) { var a, b, c; if (jd > 2299160) { a = jd + 32044; b = INT((4 * a + 3) / 146097); c = a - INT((b * 146097) / 4); } else { b = 0; c = jd + 32082; } var d = INT((4 * c + 3) / 1461), e = c - INT((1461 * d) / 4), m = INT((5 * e + 2) / 153); return [e - INT((153 * m + 2) / 5) + 1, m + 3 - 12 * INT(m / 10), b * 100 + d - 4800 + INT(m / 10)]; }
    function newMoon(k) {
      var T = k / 1236.85, T2 = T * T, T3 = T2 * T, dr = PI / 180;
      var Jd1 = 2415020.75933 + 29.53058868 * k + 0.0001178 * T2 - 0.000000155 * T3 + 0.00033 * Math.sin((166.56 + 132.87 * T - 0.009173 * T2) * dr);
      var M = 359.2242 + 29.10535608 * k - 0.0000333 * T2 - 0.00000347 * T3, Mpr = 306.0253 + 385.81691806 * k + 0.0107306 * T2 + 0.00001236 * T3, F = 21.2964 + 390.67050646 * k - 0.0016528 * T2 - 0.00000239 * T3;
      var C1 = (0.1734 - 0.000393 * T) * Math.sin(M * dr) + 0.0021 * Math.sin(2 * dr * M) - 0.4068 * Math.sin(Mpr * dr) + 0.0161 * Math.sin(dr * 2 * Mpr) - 0.0004 * Math.sin(dr * 3 * Mpr)
        + 0.0104 * Math.sin(dr * 2 * F) - 0.0051 * Math.sin(dr * (M + Mpr)) - 0.0074 * Math.sin(dr * (M - Mpr)) + 0.0004 * Math.sin(dr * (2 * F + M)) - 0.0004 * Math.sin(dr * (2 * F - M))
        - 0.0006 * Math.sin(dr * (2 * F + Mpr)) + 0.0010 * Math.sin(dr * (2 * F - Mpr)) + 0.0005 * Math.sin(dr * (2 * Mpr + M));
      var dt = T < -11 ? 0.001 + 0.000839 * T + 0.0002261 * T2 - 0.00000845 * T3 - 0.000000081 * T * T3 : -0.000278 + 0.000265 * T + 0.000262 * T2;
      return Jd1 + C1 - dt;
    }
    function sunLong(jdn) { var T = (jdn - 2451545.0) / 36525, T2 = T * T, dr = PI / 180, M = 357.52910 + 35999.05030 * T - 0.0001559 * T2 - 0.00000048 * T * T2, L0 = 280.46645 + 36000.76983 * T + 0.0003032 * T2; var DL = (1.914600 - 0.004817 * T - 0.000014 * T2) * Math.sin(dr * M) + (0.019993 - 0.000101 * T) * Math.sin(dr * 2 * M) + 0.000290 * Math.sin(dr * 3 * M); var L = (L0 + DL) * dr; return L - PI * 2 * INT(L / (PI * 2)); }
    var TZ = 7;
    function gSL(d) { return INT(sunLong(d - 0.5 - TZ / 24) / PI * 6); }
    function gNM(k) { return INT(newMoon(k) + 0.5 + TZ / 24); }
    function thang11(yy) { var off = jdFromDate(31, 12, yy) - 2415021, k = INT(off / 29.530588853), nm = gNM(k); if (gSL(nm) >= 9) nm = gNM(k - 1); return nm; }
    function leapOff(a11) { var k = INT((a11 - 2415021.076998695) / 29.530588853 + 0.5), last, i = 1, arc = gSL(gNM(k + i)); do { last = arc; i++; arc = gSL(gNM(k + i)); } while (arc !== last && i < 14); return i - 1; }
    function amSangDuong(dd, mm, yy) {      // ngày âm (không nhuận) → 'yyyy-mm-dd'
      var a11, b11; if (mm < 11) { a11 = thang11(yy - 1); b11 = thang11(yy); } else { a11 = thang11(yy); b11 = thang11(yy + 1); }
      var k = INT(0.5 + (a11 - 2415021.076998695) / 29.530588853), off = mm - 11; if (off < 0) off += 12;
      if (b11 - a11 > 365) { var lo = leapOff(a11); if (off >= lo) off += 1; }
      var r = jdToDate(gNM(k + off) + dd - 1); return fmtD(r[2], r[1], r[0]);
    }
    return { amSangDuong: amSangDuong };
  })();
  var amSangDuong = AL.amSangDuong;
  var LOAI_LE = ['Lễ, Tết', 'Nghỉ bù', 'Công ty cho nghỉ'];
  function thu(d) { var p = parseD(d); return p ? new Date(Date.UTC(p.y, p.m - 1, p.d)).getUTCDay() : -1; }
  /* Các ngày nghỉ lễ, Tết hưởng nguyên lương theo Điều 112 BLLĐ 2019 + Ngày Văn hóa Việt Nam 24/11 (từ 01/7/2026), kèm nghỉ bù khi trùng Chủ nhật */
  function lichLeTheoLuat(nam) {
    nam = +nam; var ds = [], add = function (d, ten) { ds.push({ ngay: d, ten: ten, loai: 'Lễ, Tết' }); };
    add(fmtD(nam, 1, 1), 'Tết Dương lịch');
    var tet = amSangDuong(1, 1, nam);
    add(addDays(tet, -1), 'Tết Nguyên đán (cuối năm Âm lịch)');
    for (var i = 0; i < 4; i++) add(addDays(tet, i), 'Tết Nguyên đán (mùng ' + (i + 1) + ')');
    add(amSangDuong(10, 3, nam), 'Giỗ Tổ Hùng Vương (10/3 Âm lịch)');
    add(fmtD(nam, 4, 30), 'Ngày Chiến thắng 30/4');
    add(fmtD(nam, 5, 1), 'Quốc tế Lao động 1/5');
    add(fmtD(nam, 9, 2), 'Quốc khánh 2/9');
    var t2 = thu(fmtD(nam, 9, 2));
    add(fmtD(nam, 9, [2, 3, 6].indexOf(t2) >= 0 ? 1 : 3), 'Quốc khánh (ngày liền kề)');
    if (nam >= 2026) add(fmtD(nam, 11, 24), 'Ngày Văn hóa Việt Nam 24/11');
    ds.sort(function (a, b) { return a.ngay < b.ngay ? -1 : 1; });
    var co = {}; ds.forEach(function (x) { co[x.ngay] = 1; });
    ds.slice().forEach(function (x) {               // trùng Chủ nhật → nghỉ bù ngày làm việc kế tiếp (Điều 111 khoản 3)
      if (thu(x.ngay) !== 0) return;
      var d = addDays(x.ngay, 1); while (co[d] || thu(d) === 0) d = addDays(d, 1);
      co[d] = 1; ds.push({ ngay: d, ten: 'Nghỉ bù ' + x.ten, loai: 'Nghỉ bù' });
    });
    return ds.sort(function (a, b) { return a.ngay < b.ngay ? -1 : 1; });
  }
  function cleanLe(list) {
    var seen = {};
    return (Array.isArray(list) ? list : []).map(function (x) {
      return typeof x === 'string' ? { ngay: x, ten: '', loai: 'Lễ, Tết' } : { ngay: String(x.ngay || ''), ten: String(x.ten || '').trim(), loai: LOAI_LE.indexOf(x.loai) >= 0 ? x.loai : 'Lễ, Tết' };
    }).filter(function (x) { if (!parseD(x.ngay) || seen[x.ngay]) return false; seen[x.ngay] = 1; return true; })
      .sort(function (a, b) { return a.ngay < b.ngay ? -1 : 1; });
  }

  /* ---- bảng chấm công theo ngày ---- */
  var KY_HIEU = {   // ký hiệu ô → số công / ngày nghỉ quy đổi
    X: { ten: 'Làm việc cả ngày', hien: 'X', cong: 1 },
    X2: { ten: 'Làm nửa ngày', hien: '½', cong: 0.5 },
    CT: { ten: 'Đi công tác', hien: 'CT', cong: 1 },
    P: { ten: 'Nghỉ phép năm', hien: 'P', hl: 1, don: 1 },
    P2: { ten: '½ phép năm + ½ công', hien: 'P½', hl: 0.5, cong: 0.5, don: 1 },
    R: { ten: 'Nghỉ việc riêng có lương', hien: 'R', hl: 1, don: 1 },
    R2: { ten: '½ việc riêng + ½ công', hien: 'R½', hl: 0.5, cong: 0.5, don: 1 },
    O: { ten: 'Nghỉ ốm (BHXH)', hien: 'Ô', om: 1, don: 1 },
    O2: { ten: '½ ốm + ½ công', hien: 'Ô½', om: 0.5, cong: 0.5, don: 1 },
    KL: { ten: 'Nghỉ không lương', hien: 'KL', kl: 1, don: 1 },
    KL2: { ten: '½ không lương + ½ công', hien: 'KL½', kl: 0.5, cong: 0.5, don: 1 },
    KP: { ten: 'Nghỉ không phép', hien: 'KP', kl: 1, kp: 1 },
    L: { ten: 'Nghỉ lễ, Tết (hưởng lương)', hien: 'L', le: 1 }
  };
  var DON_KH = { 'Nghỉ phép năm': 'P', 'Nghỉ việc riêng có lương': 'R', 'Nghỉ ốm (BHXH chi trả)': 'O', 'Nghỉ không lương': 'KL' };
  /* ô nghỉ theo đơn đã duyệt: {ngày (số): ký hiệu}; bỏ Chủ nhật và ngày lễ như cách đếm ngày nghỉ của đơn */
  function ngayTheoDon(dsDon, maNV, ky, cfg) {
    var out = {}, a = kyStart(ky), b = kyEnd(ky);
    (dsDon || []).filter(function (o) { return o.maNV === maNV && o.trangThai === 'Đã duyệt' && o.tuNgay <= b && o.denNgay >= a; }).forEach(function (o) {
      var kh = DON_KH[o.loai] || 'P', nua = o.tuNgay === o.denNgay && o.buoi && o.buoi !== 'Cả ngày';
      for (var d = o.tuNgay > a ? o.tuNgay : a, g = 0; d <= o.denNgay && d <= b && g < 40; d = addDays(d, 1), g++) {
        if (laNgayNghiLe(d, cfg)) continue;
        out[+d.slice(8)] = nua ? kh + '2' : kh;
      }
    });
    return out;
  }
  function cleanNgay(ct, ky, khoa) {
    ct = parseJSON(ct, {}); var n = dim(+ky.slice(0, 4), +ky.slice(5)), o = {};
    Object.keys(ct || {}).forEach(function (k) { var d = +k; if (d >= 1 && d <= n && KY_HIEU[ct[k]]) o[d] = ct[k]; });
    Object.keys(khoa || {}).forEach(function (d) { o[d] = khoa[d]; });
    return o;
  }
  /* tổng hợp sang các cột chấm công. Làm việc ngày lễ: vẫn hưởng lương ngày lễ + giờ làm lễ (300%).
     Làm Chủ nhật: tính công thường, hoặc giờ làm ngày nghỉ tuần (200%) nếu Cài đặt chọn. */
  function tongHopNgay(ct, ky, cfg) {
    ct = parseJSON(ct, {}); var y = +ky.slice(0, 4), m = +ky.slice(5), gio = num(cfg.gioCongNgay) || 8, le = {};
    (cfg.ngayLe || []).forEach(function (d) { if (d.slice(0, 7) === ky) le[+d.slice(8)] = 1; });
    var s = { ngayCong: 0, nghiHuongLuong: 0, nghiKhongLuong: 0, nghiOm: 0, ngayLe: 0, lamCN: 0, lamLe: 0, kp: 0, soO: 0 };
    Object.keys(ct).forEach(function (k) {
      var d = +k, h = KY_HIEU[ct[k]]; if (!h) return; s.soO++;
      var cn = new Date(Date.UTC(y, m - 1, d)).getUTCDay() === 0, cong = h.cong || 0;
      if (le[d]) { s.ngayLe += 1; s.lamLe += cong; return; }
      if (cn && cfg.chuNhatTinh === 'ot') { s.lamCN += cong; cong = 0; }
      s.ngayCong += cong; s.nghiHuongLuong += h.hl || 0; s.nghiKhongLuong += h.kl || 0; s.nghiOm += h.om || 0; s.ngayLe += h.le || 0; s.kp += h.kp || 0;
    });
    s.ot300 = s.lamLe * gio; s.ot200 = s.lamCN * gio;
    return s;
  }
  /* v3.25: BẢNG CHẤM CÔNG THÊM GIỜ – {ngày: số giờ}; loại giờ tự xếp theo ngày: lễ Tết 300%, nghỉ hằng tuần 200%, ngày thường 150% */
  function loaiOT(ky, d, cfg) {
    var y = +ky.slice(0, 4), m = +ky.slice(5), ds = ky + '-' + pad(d);
    if ((cfg.ngayLe || []).indexOf(ds) >= 0) return 'ot300';
    var w = new Date(Date.UTC(y, m - 1, d)).getUTCDay();
    if (w === 0 || (w === 6 && cfg.ngayNghiTuan === 't7cn')) return 'ot200';
    return 'ot150';
  }
  function cleanOT(ct, ky) {
    ct = parseJSON(ct, {}); var n = dim(+ky.slice(0, 4), +ky.slice(5)), o = {};
    Object.keys(ct || {}).forEach(function (k) { var d = +k, h = Math.round(num(ct[k]) * 100) / 100; if (d >= 1 && d <= n && h > 0) o[d] = Math.min(h, 24); });
    return o;
  }
  function tongHopOT(ct, ky, cfg) {
    ct = parseJSON(ct, {}); var s = { ot150: 0, ot200: 0, ot300: 0, tong: 0, soNgay: 0, vuotNgay: [] }, gio = num(cfg.gioCongNgay) || 8;
    Object.keys(ct).forEach(function (k) {
      var d = +k, h = num(ct[k]); if (!(h > 0)) return;
      var l = loaiOT(ky, d, cfg); s[l] += h; s.tong += h; s.soNgay++;
      if ((l === 'ot150' && h > gio / 2) || h > 12) s.vuotNgay.push(d);     // Điều 107 BLLĐ, Điều 60 NĐ 145/2020: ngày thường ≤ 50% giờ làm việc; tổng ≤ 12 giờ/ngày
    });
    ['ot150', 'ot200', 'ot300', 'tong'].forEach(function (k) { s[k] = Math.round(s[k] * 100) / 100; });
    return s;
  }
  // cột giờ làm thêm của dòng chấm công = giờ từ bảng chấm công ngày (làm lễ, làm Chủ nhật) + giờ từ bảng chấm công thêm giờ
  function apGioOT(o, ky, cfg) {
    var ot = parseJSON(o.chiTietOT, {}), coOT = Object.keys(ot).length > 0, th = o.chiTietNgay ? tongHopNgay(o.chiTietNgay, ky, cfg) : null;
    if (!coOT && !th) return;
    var t = coOT ? tongHopOT(ot, ky, cfg) : { ot150: 0, ot200: 0, ot300: 0 };
    if (coOT) o.ot150 = t.ot150;
    if (coOT || (th && cfg.chuNhatTinh === 'ot')) o.ot200 = (th && cfg.chuNhatTinh === 'ot' ? th.ot200 : 0) + t.ot200;
    o.ot300 = (th ? th.ot300 : (coOT ? 0 : num(o.ot300))) + t.ot300;
  }

  /* tạm ứng, vi phạm */
  var LOAI_TU_MD = ['Ứng lương', 'Ứng tiền ăn', 'Ứng xăng xe – đi lại', 'Ứng khác'];
  var NHOM_VP = ['Nghỉ không phép', 'Vi phạm hợp đồng', 'Mất mát, hư hỏng tài sản', 'Vi phạm nội quy – an toàn', 'Vi phạm khác'];
  function pick(o, keys) { var r = {}; keys.forEach(function (k) { if (o[k] !== undefined) r[k] = o[k]; }); return r; }

  /* ===== v3.20: CÔNG TY THÀNH VIÊN ===== */
  function chuanCTV(ds, c) {
    var out = [], da = {};
    (Array.isArray(ds) ? ds : []).forEach(function (x) {
      if (!x) return;
      var ma = String(x.ma || '').trim().toUpperCase().replace(/[^A-Z0-9_]/g, '').slice(0, 12), ten = String(x.ten || '').trim().slice(0, 160);
      if (!ma || !ten || da[ma]) return; da[ma] = 1;
      out.push({ ma: ma, ten: ten, tenTat: String(x.tenTat || '').trim().slice(0, 60) || ten, mst: String(x.mst || '').trim().slice(0, 20), diaChi: String(x.diaChi || '').trim().slice(0, 250), trangThai: x.trangThai === 'Ngừng' ? 'Ngừng' : 'Hoạt động' });
    });
    if (!out.length) out = JSON.parse(JSON.stringify(DEFAULT_CFG.congTyTV));
    return out;
  }
  function dsCTV(cfg) { return (cfg && Array.isArray(cfg.congTyTV) && cfg.congTyTV.length) ? cfg.congTyTV : DEFAULT_CFG.congTyTV; }
  // mã công ty của CBCNV; hồ sơ cũ chưa khai báo → công ty đầu danh mục (TAPACO)
  function congTyNV(nv, cfg) {
    var ds = dsCTV(cfg), m = String((nv && nv.congTy) || '').trim().toUpperCase();
    return ds.some(function (x) { return x.ma === m; }) ? m : ds[0].ma;
  }
  function ctvTheoMa(ma, cfg) { var ds = dsCTV(cfg); return ds.filter(function (x) { return x.ma === ma; })[0] || null; }
  function tenCTV(ma, cfg, tat) { var x = ctvTheoMa(ma, cfg); return x ? (tat ? x.tenTat || x.ten : x.ten) : (ma || ''); }
  // nhận mã hoặc tên (Excel ghi "TAPACO", "An Thuận Phát"…) → mã; '' nếu không nhận ra
  function nhanCTV(v, cfg) {
    var t = chuanTen(v); if (!t) return '';
    var ds = dsCTV(cfg);
    for (var i = 0; i < ds.length; i++) { var x = ds[i]; if (chuanTen(x.ma) === t || chuanTen(x.ten) === t || chuanTen(x.tenTat) === t) return x.ma; }
    for (var j = 0; j < ds.length; j++) { var y = ds[j], a = chuanTen(y.tenTat), b = chuanTen(y.ten); if ((a && (b.indexOf(t) >= 0 || t.indexOf(a) >= 0)) || (t.length >= 4 && b.indexOf(t) >= 0)) return y.ma; }
    return '';
  }
  // công ty hạch toán của 1 dòng bảng lương: chốt lúc tính lương; kỳ tính trước v3.20 → theo hồ sơ hiện tại
  function congTyBL(b, nvMap, cfg) {
    var m = String((b && b.congTy) || '').trim();
    if (m && ctvTheoMa(m, cfg)) return m;
    return congTyNV(nvMap && nvMap[b.maNV], cfg);
  }
  /* Thống kê nhân sự, quỹ lương, BH, thuế theo công ty và theo bộ phận (gộp các công ty).
     o: { kys: ['2026-09',…], nam: '2026', ct: '' (lọc bảng bộ phận theo 1 công ty) } */
  function tkCongTy(nvs, bls, kts, cfg, o) {
    o = o || {};
    var ds = dsCTV(cfg), nvMap = {}, kyOK = {}, nam = String(o.nam || '');
    (o.kys || []).forEach(function (k) { kyOK[k] = 1; });
    (nvs || []).forEach(function (n) { nvMap[n.maNV] = n; });
    var moiCT = function (x) { return { ma: x.ma, ten: x.ten, tenTat: x.tenTat || x.ten, dang: 0, thuViec: 0, nghi: 0, nghiNam: 0, vaoNam: 0, khen: 0, kyLuat: 0, soPhieu: 0, quy: 0, bhNLD: 0, bhDN: 0, kpcd: 0, thue: 0, thucLinh: 0, chiPhi: 0 }; };
    var CT = {}; ds.forEach(function (x) { CT[x.ma] = moiCT(x); });
    var BP = {}, bpCua = function (ten) {
      var k = ten || 'Chưa xếp bộ phận';
      if (!BP[k]) { var b = (cfg.boPhan || []).filter(function (x) { return x.ten === k; })[0]; BP[k] = { ten: k, loai: b ? b.loai : '', dang: 0, theoCT: {}, soPhieu: 0, quy: 0, bhNLD: 0, bhDN: 0, kpcd: 0, thue: 0, thucLinh: 0, chiPhi: 0 }; }
      return BP[k];
    };
    var locCT = function (ma) { return !o.ct || o.ct === ma; };
    (nvs || []).forEach(function (n) {
      var m = congTyNV(n, cfg), c = CT[m], nghi = n.trangThai === 'Đã nghỉ việc';
      if (nghi) { c.nghi++; if (nam && String(n.ngayNghi || '').slice(0, 4) === nam) c.nghiNam++; }
      else { c.dang++; if (n.trangThai === 'Thử việc' || n.loaiHD === 'Thử việc') c.thuViec++; if (locCT(m)) { var b = bpCua(n.phongBan); b.dang++; b.theoCT[m] = (b.theoCT[m] || 0) + 1; } }
      if (nam && String(n.ngayVao || '').slice(0, 4) === nam) c.vaoNam++;
    });
    (kts || []).forEach(function (k) {
      if (k.trangThai !== 'Đã duyệt' || (nam && String(k.ngay || '').slice(0, 4) !== nam)) return;
      var n = nvMap[k.maNV]; if (!n) return;
      var c = CT[congTyNV(n, cfg)]; if (k.loai === 'Kỷ luật') c.kyLuat++; else c.khen++;
    });
    var cong = function (t, b) {
      var bhN = num(b.bhxh) + num(b.bhyt) + num(b.bhtn), bhD = num(b.bhxhDN) + num(b.bhytDN) + num(b.bhtnDN);
      t.soPhieu++; t.quy += num(b.tongThuNhap); t.bhNLD += bhN; t.bhDN += bhD; t.kpcd += num(b.kpcdDN); t.thue += num(b.thueTNCN); t.thucLinh += num(b.thucLinh);
      t.chiPhi += num(b.tongThuNhap) + bhD + num(b.kpcdDN);
    };
    (bls || []).forEach(function (b) {
      if (!kyOK[b.ky]) return;
      var m = congTyBL(b, nvMap, cfg); if (!CT[m]) return;
      cong(CT[m], b);
      if (locCT(m)) cong(bpCua(b.phongBan), b);
    });
    var cts = ds.map(function (x) { return CT[x.ma]; });
    var tong = moiCT({ ma: '', ten: 'Tổng cộng' });
    cts.forEach(function (c) { Object.keys(tong).forEach(function (k) { if (typeof tong[k] === 'number') tong[k] += c[k]; }); });
    var bps = Object.keys(BP).map(function (k) { return BP[k]; }).sort(function (a, b) { return (a.loai === b.loai ? 0 : a.loai === 'Phòng ban' ? -1 : b.loai === 'Phòng ban' ? 1 : a.loai ? -1 : 1) || b.dang - a.dang || a.ten.localeCompare(b.ten); });
    return { congTy: cts, tong: tong, boPhan: bps };
  }

  /* ===== v3.26: QUYẾT TOÁN THUẾ TNCN NĂM – ước tính từ bảng lương, gợi ý cách giảm thuế hợp pháp =====
     Giảm trừ bản thân đủ 12 tháng; người phụ thuộc theo số tháng có bảng lương. Biểu thuế năm = biểu tháng × 12 theo quy định tháng 12 của năm
     (năm 2026: Luật 109/2025 – "áp dụng từ kỳ tính thuế năm 2026"). Chỉ để tham khảo: không biết thu nhập ở nơi khác. */
  function quyetToanNam(nam, bls, nvs, cfg, opt) {
    opt = opt || {};                                  // opt.gtDuNam: giảm trừ bản thân đủ 12 tháng (quyết toán); false = theo số tháng có lương (ước tính giữa năm)
    var ck = cfgTheoKy(cfg, nam + '-12'), bieuNam = (ck.bieuThue || []).map(function (b) { return [num(b[0]) * 12, b[1]]; }), nvMap = {}, g = {};
    (nvs || []).forEach(function (n) { nvMap[n.maNV] = n; });
    (opt.ttk || []).forEach(function (p) {            // v3.28: tiền khoán đã chi trong năm (theo ngày thực trả)
      if (p.trangThai !== 'Đã chi' || String(p.ngayTra || '').slice(0, 4) !== String(nam) || !(num(p.tnChiuThue) > 0 || num(p.thue) > 0)) return;
      var n0 = (nvs || []).filter(function (n) { return n.maNV === p.maNV; })[0] || {};
      var x = g[p.maNV] = g[p.maNV] || { maNV: p.maNV, hoTen: n0.hoTen || p.maNV, phongBan: n0.phongBan, soKy: 0, tnct: 0, bh: 0, daKT: 0, tongTN: 0, kyKT10: 0, kyLuyTien: 0 };
      x.tnct += num(p.tnChiuThue); x.daKT += num(p.thue); x.tongTN += num(p.tnChiuThue); x.tnKhoan = (x.tnKhoan || 0) + num(p.tnChiuThue);
    });
    (bls || []).forEach(function (b) {
      if (String(b.ky).slice(0, 4) !== String(nam)) return;
      var x = g[b.maNV] = g[b.maNV] || { maNV: b.maNV, hoTen: b.hoTen, phongBan: b.phongBan, soKy: 0, tnct: 0, bh: 0, daKT: 0, tongTN: 0, kyKT10: 0, kyLuyTien: 0 };
      x.soKy++; x.tnct += num(b.tnChiuThue); x.bh += num(b.bhxh) + num(b.bhyt) + num(b.bhtn); x.daKT += num(b.thueTNCN); x.tongTN += num(b.tongThuNhap);
      if (num(b.giamTru) > 0 || num(b.tnTinhThue) !== num(b.tnChiuThue)) x.kyLuyTien++; else x.kyKT10++;
    });
    return Object.keys(g).map(function (ma) {
      var x = g[ma], nv = nvMap[ma] || { maNV: ma }, tt = ttThue(nv, nam + '-12'), th = truongHopNV(nv);
      x.truongHop = th; x.cach = tt.cach; x.giamTruTaiCty = tt.giamTru; x.camKet = tt.camKet; x.uyQuyen = tt.uyQuyen; x.mst = nv.mst || '';
      x.giamTruNam = num(ck.giamTruBanThan) * (opt.gtDuNam === false ? x.soKy : 12) + num(nv.soNPT) * num(ck.giamTruNPT) * x.soKy;
      x.tntt = Math.max(0, x.tnct - x.bh - x.giamTruNam);
      x.phaiNop = thueLuyTien(x.tntt, bieuNam);
      x.chenh = x.daKT - x.phaiNop;                 // > 0: được hoàn / bù trừ; < 0: nộp thêm
      var gy = [];
      x.duUyQuyen = tt.cach === 'luytien' && th !== 'noikhac' && nv.trangThai !== 'Đã nghỉ việc';
      if (x.duUyQuyen && tt.uyQuyen === 'co') gy.push('Công ty quyết toán thay (đã ủy quyền – mẫu 08/UQ-QTT-TNCN); điều kiện: chỉ có thu nhập tại Công ty' + (x.chenh > 0 ? '; hoàn / bù trừ ' + Math.round(x.chenh).toLocaleString('vi-VN') + ' đ' : x.chenh < 0 ? '; khấu trừ thêm ' + Math.round(-x.chenh).toLocaleString('vi-VN') + ' đ' : ''));
      else if (x.duUyQuyen) gy.push('Đủ điều kiện ủy quyền Công ty quyết toán nếu chỉ có thu nhập tại Công ty – đề nghị cá nhân ký giấy ủy quyền');
      if (th === 'noikhac') gy.push('Có thu nhập từ 2 nơi: cá nhân tự quyết toán; giảm trừ gia cảnh chỉ tính tại 1 nơi' + (tt.giamTru !== 'co' ? ' (Công ty đang không giảm trừ – nếu cá nhân không đăng ký giảm trừ ở nơi khác thì có thể đăng ký tại Công ty để giảm thuế tạm khấu trừ)' : ''));
      if (tt.cach === 'khautru10') {
        var hopLeCK = x.tnct - x.bh <= x.giamTruNam;
        if (tt.camKet === 'co') gy.push(hopLeCK ? 'Cam kết 08 hợp lệ ước tính (thu nhập cả năm sau giảm trừ chưa đến mức chịu thuế)' + (!nv.mst ? ' – THIẾU MST: cam kết chỉ hợp lệ khi có mã số thuế' : '') : 'CẢNH BÁO: thu nhập cả năm đã vượt mức giảm trừ – cam kết 08 không còn đúng, cần khấu trừ 10% các lần trả tiếp theo');
        else if (hopLeCK && x.tnct > 0) gy.push('Nên làm cam kết 08/CK-TNCN (có MST, chỉ có thu nhập này): không bị khấu trừ 10%' + (x.daKT > 0 ? '; số đã khấu trừ ' + Math.round(x.daKT).toLocaleString('vi-VN') + ' đ được hoàn khi cá nhân tự quyết toán' : ''));
        else if (x.chenh > 0) gy.push('Cá nhân tự quyết toán (được giảm trừ gia cảnh cả năm) để được hoàn khoảng ' + Math.round(x.chenh).toLocaleString('vi-VN') + ' đ');
      }
      if (tt.cach === 'luytien' && tt.giamTru === 'co' && !num(nv.soNPT)) gy.push('Rà soát người phụ thuộc (con dưới 18 tuổi, cha mẹ hết tuổi lao động thu nhập thấp…): mỗi người giảm trừ ' + Math.round(num(ck.giamTruNPT)).toLocaleString('vi-VN') + ' đ/tháng; đăng ký trước 31/12');
      x.goiY = gy;
      return x;
    }).sort(function (a, b) { return a.maNV < b.maNV ? -1 : 1; });
  }

  /* ===================== v3.27: TAPACO PAYROLL & TAX 2026 =====================
     Cơ sở dữ liệu quy định (LegalRules), kho biểu mẫu, hồ sơ thuế từng người (hsThue), đánh giá điều kiện cam kết / ủy quyền,
     tối ưu thuế TNCN (3 phương án), ghi chú – cảnh báo thuế & BHXH trên bảng lương, bộ hàm tính lương (PayrollEngine), Net → Gross.
     Quy định "đang xác minh" chỉ để tham khảo: ứng dụng KHÔNG kết luận "đủ điều kiện" dựa trên căn cứ chưa xác minh. */
  var QD_TT = { xacminh: 'Đã xác minh (nguồn thứ cấp)', chinhthuc: 'Đã đối chiếu văn bản gốc', chuaxacminh: 'Đang xác minh', hethieuluc: 'Hết hiệu lực' };
  var QUY_DINH_MD = [
    { ma: 'BLLD_24', ten: 'HĐ dưới 1 tháng không áp dụng thử việc', vanBan: 'Bộ luật Lao động 45/2019/QH14', dieuKhoan: 'Điều 24 khoản 3', hieuLuc: '2021-01-01', hetHieuLuc: '', noiDung: 'Không áp dụng thử việc đối với người lao động giao kết hợp đồng lao động có thời hạn dưới 01 tháng.', url: '', trangThai: 'xacminh' },
    { ma: 'BLLD_25', ten: 'Thời gian thử việc tối đa', vanBan: 'Bộ luật Lao động 45/2019/QH14', dieuKhoan: 'Điều 25', hieuLuc: '2021-01-01', hetHieuLuc: '', noiDung: 'Chỉ thử việc 1 lần cho 1 công việc: ≤ 180 ngày (người quản lý doanh nghiệp); ≤ 60 ngày (chức danh cần trình độ cao đẳng trở lên); ≤ 30 ngày (trung cấp, công nhân kỹ thuật, nhân viên nghiệp vụ); ≤ 6 ngày làm việc (công việc khác).', url: 'https://thuvienphapluat.vn/lao-dong-tien-luong/nguoi-lao-dong-co-duoc-rut-ngan-thoi-gian-thu-viec-hay-khong-13865.html', trangThai: 'xacminh' },
    { ma: 'BLLD_26', ten: 'Tiền lương thử việc', vanBan: 'Bộ luật Lao động 45/2019/QH14', dieuKhoan: 'Điều 26', hieuLuc: '2021-01-01', hetHieuLuc: '', noiDung: 'Tiền lương thử việc do hai bên thỏa thuận nhưng ít nhất bằng 85% mức lương của công việc đó.', url: '', trangThai: 'xacminh' },
    { ma: 'BLLD_168', ten: 'Trả thêm khoản tương đương bảo hiểm', vanBan: 'Bộ luật Lao động 45/2019/QH14', dieuKhoan: 'Điều 168 khoản 3', hieuLuc: '2021-01-01', hetHieuLuc: '', noiDung: 'NLĐ không thuộc đối tượng tham gia BHXH, BHYT, BHTN bắt buộc thì ngoài tiền lương, NSDLĐ trả thêm cùng lúc với kỳ trả lương một khoản tiền tương đương mức NSDLĐ đóng BHXH, BHYT, BHTN, BH TNLĐ-BNN.', url: '', trangThai: 'xacminh' },
    { ma: 'BHXH_DT', ten: 'Đối tượng BHXH bắt buộc', vanBan: 'Luật BHXH 41/2024/QH15', dieuKhoan: 'Điều 2', hieuLuc: '2025-07-01', hetHieuLuc: '', noiDung: 'HĐLĐ từ đủ 01 tháng trở lên thuộc diện BHXH bắt buộc, kể cả hợp đồng tên gọi khác nhưng có nội dung về việc làm có trả công, tiền lương và sự quản lý, điều hành, giám sát của một bên. Người đang hưởng lương hưu không thuộc diện BHXH bắt buộc.', url: 'https://lsvn.vn/ky-hop-dong-lao-dong-voi-nguoi-da-nghi-huu-co-phai-dong-bhxh-theo-quy-dinh-a158829.html', trangThai: 'xacminh' },
    { ma: 'BHXH_NHIEUHD', ten: 'Người có nhiều HĐLĐ', vanBan: 'Luật BHXH 41/2024/QH15; NĐ 158/2025/NĐ-CP', dieuKhoan: '(đang xác minh số điều)', hieuLuc: '2025-07-01', hetHieuLuc: '', noiDung: 'BHXH, BHTN đóng theo HĐLĐ giao kết đầu tiên; BHYT theo HĐLĐ có mức lương cao nhất; BH TNLĐ-BNN đóng theo từng HĐLĐ.', url: 'https://luatvietnam.vn/tin-phap-luat/lam-nhieu-cong-ty-dong-bhxh-the-nao-230-16725-article.html', trangThai: 'xacminh' },
    { ma: 'TNCN_BIEU', ten: 'Biểu thuế lũy tiến 5 bậc', vanBan: 'Luật Thuế TNCN 109/2025/QH15', dieuKhoan: 'Điều 29 (khoản 2: áp dụng từ kỳ tính thuế năm 2026)', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Thu nhập tính thuế/tháng: đến 10 tr 5%; 10–30 tr 10%; 30–60 tr 20%; 60–100 tr 30%; trên 100 tr 35%. Tính lũy tiến từng phần.', url: 'https://www.ey.com/content/dam/ey-unified-site/ey-com/vi-vn/technical/tax/documents/ey-decree-253-providing-guidance-on-the-implementation-of-personal-income-tax-law-109-vi.pdf', trangThai: 'xacminh' },
    { ma: 'TNCN_GT', ten: 'Giảm trừ gia cảnh 2026', vanBan: 'NQ 110/2025/UBTVQH15', dieuKhoan: '', hieuLuc: '2026-01-01', hetHieuLuc: '', noiDung: 'Bản thân 15.500.000 đ/tháng; mỗi người phụ thuộc 6.200.000 đ/tháng; áp dụng từ kỳ tính thuế năm 2026.', url: '', trangThai: 'xacminh' },
    { ma: 'TNCN_NPT', ten: 'Người phụ thuộc', vanBan: 'NĐ 253/2026/NĐ-CP', dieuKhoan: 'Điều 47, 48', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Điều kiện, hồ sơ người phụ thuộc. Mỗi người phụ thuộc chỉ được tính giảm trừ 1 lần vào 1 người nộp thuế.', url: 'https://www.ey.com/content/dam/ey-unified-site/ey-com/vi-vn/technical/tax/documents/ey-decree-253-providing-guidance-on-the-implementation-of-personal-income-tax-law-109-vi.pdf', trangThai: 'xacminh' },
    { ma: 'TNCN_YTGD', ten: 'Giảm trừ chi phí y tế, giáo dục', vanBan: 'NĐ 253/2026/NĐ-CP', dieuKhoan: 'Điều 49', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Khám, chữa bệnh (trong danh mục BHYT chi trả) tối đa 23 tr/năm; giáo dục – đào tạo tại cơ sở trong nước tối đa 24 tr/năm; có hóa đơn, chứng từ đứng tên người nộp thuế hoặc người phụ thuộc. Có phát sinh thì cá nhân tự quyết toán.', url: 'https://www.ey.com/content/dam/ey-unified-site/ey-com/vi-vn/technical/tax/documents/ey-decree-253-providing-guidance-on-the-implementation-of-personal-income-tax-law-109-vi.pdf', trangThai: 'xacminh' },
    { ma: 'TNCN_HUUTRI', ten: 'Quỹ hưu trí tự nguyện', vanBan: 'NĐ 253/2026/NĐ-CP', dieuKhoan: 'Điều 46 khoản 2 điểm a', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Giảm trừ khoản đóng quỹ hưu trí tự nguyện tối đa 3 triệu đ/tháng.', url: 'https://asp.misa.vn/tt/kien-thuc/diem-moi-nghi-dinh-253-2026-nd-cp-ve-thue-tncn-10-thay-doi-trong-tam-tu-01-7-2026/', trangThai: 'xacminh' },
    { ma: 'TNCN_ANCA', ten: 'Tiền ăn giữa ca miễn thuế', vanBan: 'NĐ 253/2026/NĐ-CP', dieuKhoan: 'Điều 8 khoản 2 điểm k; Điều 69', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Mức miễn thuế tối đa 1.200.000 đ/người/tháng từ 01/7/2026 (trước đó ứng dụng giữ 730.000 đ – đang xác minh mức áp dụng 01/2025–06/2026).', url: 'https://asp.misa.vn/tt/kien-thuc/diem-moi-nghi-dinh-253-2026-nd-cp-ve-thue-tncn-10-thay-doi-trong-tam-tu-01-7-2026/', trangThai: 'xacminh' },
    { ma: 'TNCN_OT', ten: 'Tiền làm thêm giờ miễn thuế', vanBan: 'NĐ 253/2026/NĐ-CP', dieuKhoan: 'Điều 26', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Toàn bộ tiền lương làm thêm giờ, làm đêm trong giới hạn giờ của pháp luật lao động được miễn thuế (trước đây chỉ phần trả cao hơn).', url: 'https://thuehaiquan.tapchikinhtetaichinh.vn/tien-lam-them-gio-duoc-mien-thue-thu-nhap-ca-nhan-doanh-nghiep-phai-co-can-cu-chung-minh-164628.html', trangThai: 'xacminh' },
    { ma: 'TNCN_KT10', ten: 'Khấu trừ 10% thu nhập vãng lai', vanBan: 'NĐ 253/2026/NĐ-CP', dieuKhoan: 'Điều 50 khoản 2', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Cá nhân không ký HĐLĐ hoặc HĐLĐ dưới 3 tháng (kể cả nhận thu nhập sau khi nghỉ việc): khấu trừ 10% khi trả từ 5 triệu đ/lần (trước 01/7/2026: 2 triệu đ/lần).', url: 'https://asp.misa.vn/tt/kien-thuc/diem-moi-nghi-dinh-253-2026-nd-cp-ve-thue-tncn-10-thay-doi-trong-tam-tu-01-7-2026/', trangThai: 'xacminh' },
    { ma: 'TNCN_CAMKET', ten: 'Cam kết tạm không khấu trừ 10%', vanBan: 'TT 80/2021/TT-BTC (mẫu 08/CK-TNCN) – TT 80 được thay bằng TT 89/2026/TT-BTC từ 01/7/2026', dieuKhoan: '(đang xác minh căn cứ và mẫu hiện hành)', hieuLuc: '', hetHieuLuc: '', noiDung: 'Điều kiện theo hướng dẫn trước đây (TT 111/2013 Điều 25): cá nhân cư trú, chỉ có duy nhất thu nhập thuộc diện khấu trừ 10%, đã có MST, ước tính thu nhập cả năm sau giảm trừ gia cảnh chưa đến mức nộp thuế. Danh mục mẫu TT 89/2026 đã tra chưa thấy mẫu cam kết – cần Kế toán trưởng / cơ quan thuế xác nhận.', url: 'https://luatvietnam.vn/bieu-mau/mau-08-ck-tncn-mau-cam-ket-thue-tncn-moi-nhat-571-34450-article.html', trangThai: 'chuaxacminh' },
    { ma: 'TNCN_UQ', ten: 'Điều kiện ủy quyền quyết toán', vanBan: 'NĐ 253/2026/NĐ-CP', dieuKhoan: 'Điều 51 khoản 2 điểm a', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Cá nhân cư trú có HĐLĐ từ 3 tháng trở lên tại 1 tổ chức và thực tế đang làm việc tại đó vào thời điểm quyết toán (kể cả không đủ 12 tháng); hoặc được điều chuyển do sáp nhập, hợp nhất, chia, tách, chuyển đổi loại hình hoặc giữa các tổ chức cùng hệ thống.', url: 'https://ketoanthienung.net/cac-truong-hop-duoc-uy-quyen-quyet-toan-thue-tncn.htm', trangThai: 'xacminh' },
    { ma: 'TNCN_VL15', ten: 'Thu nhập vãng lai bình quân ≤ 15 triệu/tháng', vanBan: 'NĐ 253/2026/NĐ-CP', dieuKhoan: 'Điều 51 khoản 1 điểm b', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Thu nhập vãng lai ở nơi khác bình quân tháng trong năm không quá 15 triệu đ và đã khấu trừ 10% thì không phải quyết toán phần này (vẫn ủy quyền được), trừ khi cá nhân yêu cầu quyết toán cả phần này.', url: 'https://www.ey.com/content/dam/ey-unified-site/ey-com/vi-vn/technical/tax/documents/ey-decree-253-providing-guidance-on-the-implementation-of-personal-income-tax-law-109-vi.pdf', trangThai: 'xacminh' },
    { ma: 'TNCN_CT70', ten: 'Chuyển tiếp năm 2026', vanBan: 'NĐ 253/2026/NĐ-CP', dieuKhoan: 'Điều 70', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Thu nhập 01/01–30/6/2026 đã khai theo quy định cũ không phải khai lại; chênh lệch được điều chỉnh khi quyết toán thuế TNCN năm 2026.', url: 'https://asp.misa.vn/tt/kien-thuc/diem-moi-nghi-dinh-253-2026-nd-cp-ve-thue-tncn-10-thay-doi-trong-tam-tu-01-7-2026/', trangThai: 'xacminh' },
    { ma: 'TT89', ten: 'Hồ sơ khai thuế, biểu mẫu', vanBan: 'Thông tư 89/2026/TT-BTC ngày 30/6/2026', dieuKhoan: 'Phụ lục biểu mẫu', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Mẫu 08/UQ-QTT-TNCN (giấy ủy quyền quyết toán), 02/QTT-TNCN (cá nhân), 05/QTT-TNCN (tổ chức), bảng kê 02-1/BK-QTT-TNCN, 05-3/BK-QTT-TNCN. Hồ sơ khai thuế trước 01/7/2026 tiếp tục dùng mẫu cũ.', url: 'https://luatvietnam.vn/bieu-mau/tong-hop-bieu-mau-ve-ho-so-khai-thue-va-quan-ly-thue-tai-thong-tu-89-2026-tt-btc-571-110577-article.html', trangThai: 'xacminh' },
    { ma: 'L09_2026', ten: 'Luật sửa đổi các luật thuế', vanBan: 'Luật 09/2026/QH16 (thông qua 24/4/2026)', dieuKhoan: 'Điều 1', hieuLuc: '', hetHieuLuc: '', noiDung: 'Sửa đổi Luật Thuế TNCN về quản lý thu nhập, khai thuế, quyết toán, trách nhiệm người nộp thuế. CHƯA xác minh nội dung chi tiết tác động đến tiền lương, tiền công và ngày hiệu lực.', url: 'https://luatvietnam.vn/tin-van-ban-moi/da-co-luat-so-09-2026-qh16-sua-doi-cac-luat-thue-tncn-thue-gtgt-thue-tndn-thue-ttdb-186-109001-article.html', trangThai: 'chuaxacminh' },
    { ma: 'ND252', ten: 'Hướng dẫn Luật Quản lý thuế', vanBan: 'NĐ 252/2026/NĐ-CP', dieuKhoan: '', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Đăng ký thuế, khai thuế; tổ chức trả thu nhập đăng ký thuế thay cho cá nhân. Chưa xác minh chi tiết điều khoản liên quan tiền lương.', url: 'https://www.forvismazars.com/vn/en/insights/vietnam-updates/newsletters-alerts/vietnam-tax-legal-updates/decree-252-on-tax-administration', trangThai: 'chuaxacminh' },
    { ma: 'BLLD_13', ten: 'Định nghĩa hợp đồng lao động', vanBan: 'Bộ luật Lao động 45/2019/QH14', dieuKhoan: 'Điều 13 khoản 1', hieuLuc: '2021-01-01', hetHieuLuc: '', noiDung: 'Hợp đồng lao động là sự thỏa thuận về việc làm có trả công, tiền lương, điều kiện lao động; trường hợp hai bên thỏa thuận bằng tên gọi khác nhưng có nội dung thể hiện việc làm có trả công, tiền lương và sự quản lý, điều hành, giám sát của một bên thì được coi là hợp đồng lao động.', url: '', trangThai: 'xacminh' },
    { ma: 'ND68', ten: 'Thuế hộ, cá nhân kinh doanh', vanBan: 'NĐ 68/2026/NĐ-CP ngày 05/3/2026', dieuKhoan: '', hieuLuc: '', hetHieuLuc: '', noiDung: 'Doanh thu ≤ 500 triệu đ/năm không chịu GTGT, TNCN; trên 500 tr – 3 tỷ: thuế TNCN theo tỷ lệ % doanh thu; trên 3 tỷ: theo thu nhập (doanh thu − chi phí). Không áp khấu trừ 10% tiền công và giảm trừ bản thân cho doanh thu kinh doanh. Chưa xác minh trách nhiệm khấu trừ thay của tổ chức chi trả.', url: 'https://lsvn.vn/nghi-dinh-68-2026-nd-cp-va-buoc-chuyen-trong-quan-ly-thue-doi-voi-ho-kinh-doanh-ca-nhan-kinh-doanh-a170115.html', trangThai: 'xacminh' },
    { ma: 'TNCN_KCT', ten: 'Cá nhân không cư trú', vanBan: 'Luật Thuế TNCN 109/2025/QH15', dieuKhoan: '(đang xác minh số điều)', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Ứng dụng tạm áp 20% trên thu nhập tiền lương, tiền công của cá nhân không cư trú (theo quy định trước đây) – cần Phụ trách kế toán xác nhận mức áp dụng theo Luật 109/2025.', url: '', trangThai: 'chuaxacminh' },
    { ma: 'L109', ten: 'Luật Thuế TNCN 109/2025/QH15 – văn bản gốc', vanBan: 'Luật Thuế TNCN 109/2025/QH15', dieuKhoan: '', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Link Cổng văn bản Chính phủ do TAPACO cung cấp.', url: 'https://vanban.chinhphu.vn/?classid=1&docid=216495&pageid=27160&typegroupid=1', trangThai: 'xacminh', urlKiemTra: 'chua' },
    { ma: 'ND253', ten: 'NĐ 253/2026/NĐ-CP – văn bản gốc', vanBan: 'NĐ 253/2026/NĐ-CP ngày 30/6/2026', dieuKhoan: '', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Link Cổng thông tin Chính phủ do TAPACO cung cấp.', url: 'https://chinhphu.vn/?classid=1&docid=218684&pageid=27160&typegroupid=4', trangThai: 'xacminh', urlKiemTra: 'chua' },
    { ma: 'TT89_GOC', ten: 'TT 89/2026/TT-BTC – văn bản gốc', vanBan: 'Thông tư 89/2026/TT-BTC', dieuKhoan: '', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Link Cổng thông tin Chính phủ do TAPACO cung cấp.', url: 'https://chinhphu.vn/?classid=1&docid=218974&orggroupid=4&pageid=27160', trangThai: 'xacminh', urlKiemTra: 'chua' },
    { ma: 'BLLD_149', ten: 'Người lao động cao tuổi', vanBan: 'Bộ luật Lao động 45/2019/QH14', dieuKhoan: 'Điều 148, 149', hieuLuc: '2021-01-01', hetHieuLuc: '', noiDung: 'Người lao động cao tuổi làm việc theo HĐLĐ; không sử dụng người lao động cao tuổi làm nghề, công việc nặng nhọc, độc hại, nguy hiểm có ảnh hưởng xấu tới sức khỏe, trừ trường hợp bảo đảm các điều kiện làm việc an toàn; quan tâm chăm sóc sức khỏe.', url: '', trangThai: 'xacminh' },
    { ma: 'TNCN_LUONGHUU', ten: 'Lương hưu miễn thuế', vanBan: 'Luật Thuế TNCN 109/2025/QH15', dieuKhoan: '(đang xác minh số điều trong Luật 109)', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Tiền lương hưu do Quỹ BHXH chi trả thuộc thu nhập được miễn thuế; tiền lương, tiền công do doanh nghiệp trả cho người đang hưởng lương hưu vẫn chịu thuế.', url: '', trangThai: 'chuaxacminh' },
    { ma: 'TNCN_UQ_TUQT', ten: 'Trường hợp phải tự quyết toán', vanBan: 'NĐ 253/2026/NĐ-CP', dieuKhoan: 'Điều 51 khoản 3', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Các trường hợp cá nhân phải trực tiếp quyết toán – CHƯA xác minh nội dung chi tiết; Kế toán trưởng xác định khi xét ủy quyền.', url: '', trangThai: 'chuaxacminh' },
    { ma: 'BLLD_GOC', ten: 'Bộ luật Lao động 45/2019/QH14 – văn bản gốc', vanBan: 'Bộ luật Lao động 45/2019/QH14', dieuKhoan: '', hieuLuc: '2021-01-01', hetHieuLuc: '', noiDung: 'Link Cổng văn bản Chính phủ do TAPACO cung cấp – hệ thống chưa truy cập kiểm tra được.', url: 'https://vanban.chinhphu.vn/?docid=198540&pageid=27160', trangThai: 'chuaxacminh', urlKiemTra: 'chua' },
    { ma: 'BHXH_GOC', ten: 'Luật BHXH 41/2024/QH15 – văn bản gốc', vanBan: 'Luật BHXH 41/2024/QH15', dieuKhoan: '', hieuLuc: '2025-07-01', hetHieuLuc: '', noiDung: 'Link Cổng văn bản Chính phủ do TAPACO cung cấp – hệ thống chưa truy cập kiểm tra được.', url: 'https://vanban.chinhphu.vn/?docid=211199&pageid=27160', trangThai: 'chuaxacminh', urlKiemTra: 'chua' },
    { ma: 'TT87', ten: 'Thông tư 87/2026/TT-BTC', vanBan: 'Thông tư 87/2026/TT-BTC ngày 30/6/2026', dieuKhoan: '', hieuLuc: '2026-07-01', hetHieuLuc: '', noiDung: 'Hướng dẫn thuế TNCN đi kèm NĐ 253/2026 – CHƯA xác minh nội dung. Link do TAPACO cung cấp, chưa kiểm tra.', url: 'https://vanban.chinhphu.vn/?docid=218772&pageid=27160', trangThai: 'chuaxacminh', urlKiemTra: 'chua' },
    { ma: 'ETAX', ten: 'Cổng thuế điện tử', vanBan: 'Tổng cục Thuế / Cục Thuế', dieuKhoan: '', hieuLuc: '', hetHieuLuc: '', noiDung: 'Nộp tờ khai quyết toán 02/QTT-TNCN, đề nghị hoàn thuế.', url: 'https://thuedientu.gdt.gov.vn', trangThai: 'xacminh', urlKiemTra: 'chua' },
    { ma: 'LVL_74', ten: 'Bảo hiểm thất nghiệp', vanBan: 'Luật Việc làm 74/2025/QH15', dieuKhoan: '', hieuLuc: '', hetHieuLuc: '', noiDung: 'Đối tượng và mức đóng BHTN. Ứng dụng đang giữ tỷ lệ BHTN 1% NLĐ / 1% DN – chưa xác minh thay đổi theo luật mới.', url: '', trangThai: 'chuaxacminh' }
  ];
  var BM_TT = { hienhanh: 'Hiện hành', dangxacminh: 'Đang xác minh', hethieuluc: 'Hết hiệu lực', noibo: 'Biểu mẫu nội bộ TAPACO' };
  var BIEU_MAU_MD = [
    { ma: 'camket', ten: 'Cam kết thu nhập – không khấu trừ 10% thuế TNCN', mauSo: '08/CK-TNCN (mẫu cũ)', canCu: 'TT 80/2021/TT-BTC – đã được thay bằng TT 89/2026/TT-BTC từ 01/7/2026', apDungTu: '', phienBan: 'TT 80/2021', trangThai: 'dangxacminh', doiTuong: 'Cá nhân cư trú, HĐ dưới 3 tháng / khoán / thử việc, chỉ có 1 nguồn thu nhập thuộc diện khấu trừ 10%, đã có MST, ước tính cả năm chưa đến mức nộp thuế.', huongDan: 'Chỉ dùng để THAM KHẢO đến khi Kế toán trưởng xác nhận mẫu được cơ quan thuế chấp nhận. Ghi đúng năm áp dụng, tổng thu nhập dự kiến, số người phụ thuộc; cá nhân tự chịu trách nhiệm về cam kết.', url: 'https://luatvietnam.vn/bieu-mau/mau-08-ck-tncn-mau-cam-ket-thue-tncn-moi-nhat-571-34450-article.html', tao: 'camket' },
    { ma: 'uyquyen', ten: 'Giấy ủy quyền quyết toán thuế TNCN', mauSo: '08/UQ-QTT-TNCN', canCu: 'Thông tư 89/2026/TT-BTC ngày 30/6/2026; điều kiện: NĐ 253/2026 Điều 51', apDungTu: '2026-07-01', phienBan: 'TT 89/2026', trangThai: 'hienhanh', doiTuong: 'Cá nhân cư trú có HĐLĐ từ 3 tháng tại TAPACO, đang làm việc tại thời điểm quyết toán; không có thu nhập tiền lương nơi khác (trừ vãng lai bình quân ≤ 15 tr/tháng đã khấu trừ 10%).', huongDan: 'Bản điền sẵn theo nội dung mẫu. Đối chiếu với mẫu gốc trong phụ lục TT 89/2026 trước khi ký; nộp lại bản ký cho Phòng Kế toán trước ngày quyết toán.', url: 'https://luatvietnam.vn/bieu-mau/tong-hop-bieu-mau-ve-ho-so-khai-thue-va-quan-ly-thue-tai-thong-tu-89-2026-tt-btc-571-110577-article.html', tao: 'uyquyen' },
    { ma: 'npt', ten: 'Đăng ký người phụ thuộc', mauSo: 'Theo mẫu hiện hành (chưa xác minh số mẫu tại TT 89/2026)', canCu: 'NĐ 253/2026 Điều 47, 48', apDungTu: '', phienBan: '', trangThai: 'dangxacminh', doiTuong: 'Người nộp thuế có người phụ thuộc đủ điều kiện.', huongDan: 'Kèm giấy tờ chứng minh (giấy khai sinh, CCCD của người phụ thuộc…). Lưu bản quét vào Tài liệu → Hồ sơ người phụ thuộc.', url: '', tao: '' },
    { ma: '02qtt', ten: 'Tờ khai quyết toán thuế TNCN (cá nhân)', mauSo: '02/QTT-TNCN', canCu: 'Thông tư 89/2026/TT-BTC', apDungTu: '2026-07-01', phienBan: 'TT 89/2026', trangThai: 'hienhanh', doiTuong: 'Cá nhân cư trú có thu nhập tiền lương, tiền công tự quyết toán.', huongDan: 'Cá nhân tự lập trên eTax / ứng dụng thuế điện tử; Công ty cung cấp chứng từ khấu trừ.', url: 'https://luatvietnam.vn/bieu-mau/tong-hop-bieu-mau-ve-ho-so-khai-thue-va-quan-ly-thue-tai-thong-tu-89-2026-tt-btc-571-110577-article.html', tao: '' },
    { ma: '05qtt', ten: 'Tờ khai quyết toán thuế TNCN (tổ chức trả thu nhập)', mauSo: '05/QTT-TNCN', canCu: 'Thông tư 89/2026/TT-BTC', apDungTu: '2026-07-01', phienBan: 'TT 89/2026', trangThai: 'hienhanh', doiTuong: 'Công ty quyết toán thay cho cá nhân đã ủy quyền; kèm bảng kê 05-1, 05-2, 05-3/BK-QTT-TNCN.', huongDan: 'Lập bằng phần mềm HTKK / eTax. Số liệu lấy từ Báo cáo lương → Quyết toán TNCN năm.', url: 'https://luatvietnam.vn/bieu-mau/tong-hop-bieu-mau-ve-ho-so-khai-thue-va-quan-ly-thue-tai-thong-tu-89-2026-tt-btc-571-110577-article.html', tao: '' },
    { ma: 'ctkt', ten: 'Chứng từ khấu trừ thuế TNCN', mauSo: 'Theo quy định hiện hành về chứng từ điện tử (chưa xác minh)', canCu: 'NĐ 254/2026/NĐ-CP (hóa đơn, chứng từ điện tử) – đang xác minh', apDungTu: '', phienBan: '', trangThai: 'dangxacminh', doiTuong: 'Cá nhân bị khấu trừ thuế (vãng lai, nghỉ việc, không ủy quyền).', huongDan: 'Lập trên phần mềm chứng từ điện tử của Công ty.', url: '', tao: '' },
    { ma: 'bm01', ten: 'BM-01: Phiếu xác nhận bản chất hợp đồng khoán', mauSo: 'TAPACO-KV-01', canCu: 'Biểu mẫu nội bộ – BLLĐ Điều 13; Luật BHXH 2024 Điều 2', apDungTu: '2026-01-01', phienBan: '1.0', trangThai: 'noibo', doiTuong: 'Mọi hợp đồng khoán việc, tổ đội', huongDan: 'Ban chỉ huy trả lời 11 câu hỏi bản chất; Phụ trách KT xác nhận phân loại K1–K4.', url: '', tao: 'bm01' },
    { ma: 'bm02', ten: 'BM-02: Bản kê khai thu nhập dự kiến trong năm', mauSo: 'TAPACO-KV-02', canCu: 'Biểu mẫu nội bộ – người nhận khoán ký xác nhận', apDungTu: '2026-01-01', phienBan: '1.0', trangThai: 'noibo', doiTuong: 'Người nhận khoán', huongDan: 'Kê khai thu nhập tại TAPACO và nơi khác, người phụ thuộc, giảm trừ khác.', url: '', tao: 'bm02' },
    { ma: 'bm04', ten: 'BM-04: Bảng xác nhận khối lượng khoán', mauSo: 'TAPACO-KV-04', canCu: 'Biểu mẫu nội bộ – theo công trình, hạng mục, đợt nghiệm thu', apDungTu: '2026-01-01', phienBan: '1.0', trangThai: 'noibo', doiTuong: 'Ban chỉ huy công trình', huongDan: 'Lập cho từng đợt nghiệm thu; ký xác nhận BCH, người nhận khoán.', url: '', tao: 'bm04' },
    { ma: 'bm05', ten: 'BM-05: Bảng thanh toán tiền khoán và khấu trừ thuế', mauSo: 'TAPACO-KV-05', canCu: 'Biểu mẫu nội bộ', apDungTu: '2026-01-01', phienBan: '1.0', trangThai: 'noibo', doiTuong: 'Phòng Kế toán', huongDan: 'Gross, thuế, tạm ứng, thực nhận, số đã thanh toán theo hợp đồng.', url: '', tao: 'bm05' },
    { ma: 'bm06', ten: 'BM-06: Phiếu tổng hợp thu nhập và thuế đã khấu trừ', mauSo: 'TAPACO-KV-06', canCu: 'Biểu mẫu nội bộ – hỗ trợ cá nhân quyết toán (không thay chứng từ khấu trừ)', apDungTu: '2026-01-01', phienBan: '1.0', trangThai: 'noibo', doiTuong: 'Người nhận khoán', huongDan: 'Cấp cho cá nhân cùng chứng từ khấu trừ thuế điện tử.', url: '', tao: 'bm06' },
    { ma: 'bm07', ten: 'BM-07: Hướng dẫn quyết toán / hoàn thuế TNCN', mauSo: 'TAPACO-KV-07', canCu: 'Biểu mẫu nội bộ; tờ khai 02/QTT-TNCN (TT 89/2026)', apDungTu: '2026-01-01', phienBan: '1.0', trangThai: 'noibo', doiTuong: 'Người nhận khoán tự quyết toán', huongDan: 'Các bước tự quyết toán, đề nghị hoàn trên thuế điện tử.', url: '', tao: 'bm07' },
    { ma: 'bmh01', ten: 'Phụ lục thỏa thuận tiền trả thêm tương đương bảo hiểm', mauSo: 'TAPACO-HT-01', canCu: 'Biểu mẫu nội bộ – Điều 168 khoản 3 BLLĐ', apDungTu: '2026-01-01', phienBan: '1.0', trangThai: 'noibo', doiTuong: 'Người đang hưởng lương hưu làm việc theo HĐLĐ.', huongDan: 'Ghi rõ phương án A (cộng thêm) hoặc B (đã gồm trong tổng chi trả), căn cứ tính, tỷ lệ, số tiền; hai bên ký.', url: '', tao: 'bmh01' },
    { ma: 'bmh02', ten: 'Bảng tính lương và thuế TNCN – người hưởng lương hưu', mauSo: 'TAPACO-HT-02', canCu: 'Biểu mẫu nội bộ', apDungTu: '2026-01-01', phienBan: '1.0', trangThai: 'noibo', doiTuong: 'Người đang hưởng lương hưu.', huongDan: 'Từng kỳ: lương, trả thêm BH, BH NLĐ, giảm trừ, thu nhập tính thuế, thuế, thực nhận; lương hưu miễn thuế ghi riêng.', url: '', tao: 'bmh02' },
    { ma: 'kekhaiGT', ten: 'Phiếu kê khai thu nhập và nơi đăng ký giảm trừ bản thân', mauSo: 'TAPACO-TNCN-01', canCu: 'Biểu mẫu nội bộ – dùng để xác định giảm trừ gia cảnh chỉ tại 1 nơi', apDungTu: '2026-01-01', phienBan: '1.0', trangThai: 'noibo', doiTuong: 'Mọi CBCNV, bắt buộc với người có thu nhập ở nơi khác (trường hợp 3).', huongDan: 'CBCNV kê khai các nguồn thu nhập trong năm và nơi đăng ký giảm trừ bản thân; ký và nộp Phòng Kế toán.', url: '', tao: 'kekhaiGT' },
    { ma: 'xacnhanBH', ten: 'Phiếu xác nhận nơi tham gia BHXH khác', mauSo: 'TAPACO-BH-01', canCu: 'Biểu mẫu nội bộ – người có nhiều HĐLĐ', apDungTu: '2026-01-01', phienBan: '1.0', trangThai: 'noibo', doiTuong: 'CBCNV đang đóng BHXH tại đơn vị khác (trường hợp 3).', huongDan: 'Kèm bản sao tờ khai / xác nhận quá trình đóng BHXH (VssID) của đơn vị đang đóng.', url: '', tao: 'xacnhanBH' }
  ];
  function chuanDsMa(ds, md, truong) {                      // danh sách sửa được trong Cài đặt: giữ mục người dùng sửa, bổ sung mục mặc định mới
    var out = [], da = {};
    (Array.isArray(ds) ? ds : []).forEach(function (x) {
      if (!x || !String(x.ma || '').trim() || da[x.ma]) return; da[x.ma] = 1;
      var o = {}; truong.forEach(function (k) { o[k] = String(x[k] == null ? '' : x[k]).slice(0, 1200); }); out.push(o);
    });
    md.forEach(function (x) { if (!da[x.ma]) out.push(clone(x)); });
    return out;
  }
  var QD_TRUONG = ['ma', 'ten', 'vanBan', 'dieuKhoan', 'hieuLuc', 'hetHieuLuc', 'noiDung', 'url', 'trangThai', 'urlKiemTra'];
  var BM_TRUONG = ['ma', 'ten', 'mauSo', 'canCu', 'apDungTu', 'phienBan', 'trangThai', 'doiTuong', 'huongDan', 'url', 'tao', 'urlKiemTra'];
  function quyDinhDs(cfg) { return chuanDsMa(cfg && cfg.quyDinh, QUY_DINH_MD, QD_TRUONG).map(function (q) { if (!QD_TT[q.trangThai]) q.trangThai = 'chuaxacminh'; if (q.url && !/^https:\/\//.test(q.url)) q.url = ''; if (q.urlKiemTra !== 'chua') q.urlKiemTra = q.url ? 'da' : ''; return q; }); }
  function bieuMauDs(cfg) { return chuanDsMa(cfg && cfg.bieuMau, BIEU_MAU_MD, BM_TRUONG).map(function (b) { if (!BM_TT[b.trangThai]) b.trangThai = 'dangxacminh'; if (b.url && !/^https:\/\//.test(b.url)) b.url = ''; if (b.urlKiemTra !== 'chua') b.urlKiemTra = b.url ? 'da' : ''; return b; }); }
  function quyDinh(cfg, ma) { return quyDinhDs(cfg).filter(function (q) { return q.ma === ma; })[0] || null; }
  function bieuMau(cfg, ma) { return bieuMauDs(cfg).filter(function (b) { return b.ma === ma; })[0] || null; }

  /* ---- HỒ SƠ THUẾ – BH từng người: NhanVien.hsThue (JSON) ---- */
  var DH_QHLD = [
    ['gioLam', 'Làm theo giờ hành chính / ca do Công ty bố trí'],
    ['quanLy', 'Chịu sự quản lý, điều hành, nội quy, kỷ luật lao động của Công ty'],
    ['luongCoDinh', 'Nhận tiền cố định hằng tháng, không theo khối lượng nghiệm thu'],
    ['thuongXuyen', 'Công việc thường xuyên, liên tục từ 1 tháng trở lên'],
    ['congCu', 'Dùng công cụ, phương tiện, vật tư của Công ty để làm việc']
  ];
  var TV_VITRI = {
    quanly: { ten: 'Người quản lý doanh nghiệp', ngay: 180 },
    caodang: { ten: 'Chức danh cần trình độ cao đẳng trở lên', ngay: 60 },
    trungcap: { ten: 'Trung cấp, công nhân kỹ thuật, nhân viên nghiệp vụ', ngay: 30 },
    khac: { ten: 'Công việc khác', ngay: 6, lamViec: true }
  };
  function hsThueNV(nv) {
    var h = parseJSON(nv && nv.hsThue, {}); if (!h || typeof h !== 'object') h = {};
    var o = function (k) { return h[k] && typeof h[k] === 'object' ? h[k] : {}; }, v3 = function (x) { return x === 'co' || x === 'khong' ? x : ''; };
    var tv = o('thuViec'), nk = o('noiKhac'), hu = o('huu'), kh = o('khoan'), tk = o('tnKhac'), dk = o('duKien'), gk = o('giamTruKhac');
    return {
      cuTru: h.cuTru === 'khong' ? 'khong' : 'co',
      thuViec: { viTri: TV_VITRI[tv.viTri] ? tv.viTri : '', tuNgay: parseD(tv.tuNgay) ? tv.tuNgay : '', denNgay: parseD(tv.denNgay) ? tv.denNgay : '', luongCV: num(tv.luongCV), tyLe: num(tv.tyLe), trongHDLD: v3(tv.trongHDLD), ngayChuyenHD: parseD(tv.ngayChuyenHD) ? tv.ngayChuyenHD : '' },
      noiKhac: { donVi: String(nk.donVi || '').slice(0, 200), maDV: String(nk.maDV || '').slice(0, 40), ngayHDDau: parseD(nk.ngayHDDau) ? nk.ngayHDDau : '', bhxh: v3(nk.bhxh), bhyt: v3(nk.bhyt), bhtn: v3(nk.bhtn), taiLieu: String(nk.taiLieu || '').slice(0, 300), gtTai: nk.gtTai === 'tapaco' || nk.gtTai === 'noikhac' ? nk.gtTai : '' },
      huu: (function () { var tt = HUU_TT[hu.tinhTrang] ? hu.tinhTrang : (hu.dangHuong === 'co' ? 'dangHuong' : ''), ng = parseD(hu.ngayHuong) ? hu.ngayHuong : parseD(hu.tuNgay) ? hu.tuNgay : '';   // v3.29
        return { tinhTrang: tt, dangHuong: tt === 'dangHuong' ? 'co' : v3(hu.dangHuong), ngayHuong: ng, tuNgay: ng, soQD: String(hu.soQD || '').slice(0, 80), donViChiTra: String(hu.donViChiTra || '').slice(0, 200), luongHuu: num(hu.luongHuu), theBHYT: v3(hu.theBHYT),
          trangThaiXM: HUU_XM.indexOf(hu.trangThaiXM) >= 0 ? hu.trangThaiXM : 'Chờ', ngayXM: parseD(hu.ngayXM) ? hu.ngayXM : '', nguoiXM: String(hu.nguoiXM || '').slice(0, 80), phuongAnBH: HUU_PA[hu.phuongAnBH] ? hu.phuongAnBH : '', canCuBH: num(hu.canCuBH), congViecNguyCo: v3(hu.congViecNguyCo) }; })(),
      khoan: { kinhDoanh: v3(kh.kinhDoanh), dauHieu: (Array.isArray(kh.dauHieu) ? kh.dauHieu : []).filter(function (k) { return DH_QHLD.some(function (d) { return d[0] === k; }); }) },
      tnKhac: { luongNoiKhac: v3(tk.luongNoiKhac), vangLai: v3(tk.vangLai), vangLaiBQ: num(tk.vangLaiBQ), coChungTu: v3(tk.coChungTu), qtVangLai: v3(tk.qtVangLai), dieuChuyen: v3(tk.dieuChuyen), tnNoiKhacNam: num(tk.tnNoiKhacNam), thueNoiKhac: num(tk.thueNoiKhac), tuQTBatBuoc: v3(tk.tuQTBatBuoc) },
      duKien: { thuNhapNam: num(dk.thuNhapNam) },
      giamTruKhac: { yTe: num(gk.yTe), giaoDuc: num(gk.giaoDuc), huuTri: num(gk.huuTri) },
      nptDs: (Array.isArray(h.nptDs) ? h.nptDs : []).filter(function (p) { return p && String(p.ten || '').trim(); }).slice(0, 10).map(function (p) { var k = function (x) { return /^\d{4}-(0[1-9]|1[0-2])$/.test(String(x || '')) ? x : ''; }; return { ten: String(p.ten).slice(0, 80), quanHe: String(p.quanHe || '').slice(0, 40), tu: k(p.tu), den: k(p.den), coHoSo: p.coHoSo === 'co' ? 'co' : 'khong' }; })
    };
  }
  function soNgayLich(a, b) { var x = parseD(a), y = parseD(b); if (!x || !y) return 0; return Math.round((Date.UTC(y.y, y.m - 1, y.d) - Date.UTC(x.y, x.m - 1, x.d)) / 864e5) + 1; }
  function soNgayLamViec(a, b) {                          // trừ Chủ nhật
    var n = 0, d = a, g = 0; if (!parseD(a) || !parseD(b) || a > b) return 0;
    while (d <= b && g < 400) { var p = parseD(d); if (new Date(Date.UTC(p.y, p.m - 1, p.d)).getUTCDay() !== 0) n++; d = addDays(d, 1); g++; }
    return n;
  }
  /* kiểm tra thử việc (Điều 24, 25, 26 BLLĐ) – không mặc định 3 tháng cho mọi vị trí */
  function kiemTraThuViec(nv) {
    var h = hsThueNV(nv).thuViec, out = { coThuViec: nv.loaiHD === 'Thử việc' || truongHopNV(nv) === 'thuviec' || !!h.tuNgay, loi: [], canhBao: [], gioiHan: 0, soNgay: 0 };
    if (!out.coThuViec) return out;
    var vt = TV_VITRI[h.viTri];
    if (!vt) out.canhBao.push('Chưa khai vị trí thử việc – không mặc định thử việc 3 tháng; thời hạn tối đa phụ thuộc vị trí (Điều 25 BLLĐ).');
    var tu = h.tuNgay || nv.ngayKyHD || nv.ngayVao, den = h.denNgay || nv.ngayHetHanHD;
    if (!tu || !den) out.canhBao.push('Chưa có ngày bắt đầu / kết thúc dự kiến thử việc.');
    if (vt && tu && den) {
      out.gioiHan = vt.ngay; out.soNgay = vt.lamViec ? soNgayLamViec(tu, den) : soNgayLich(tu, den);
      if (out.soNgay > vt.ngay) out.loi.push('Thời gian thử việc vượt giới hạn cho phép: ' + out.soNgay + (vt.lamViec ? ' ngày làm việc' : ' ngày') + ' > ' + vt.ngay + (vt.lamViec ? ' ngày làm việc' : ' ngày') + ' (' + vt.ten + ', Điều 25 BLLĐ).');
    }
    if (h.tyLe > 0 && h.tyLe < 85) out.loi.push('Lương thử việc ' + h.tyLe + '% thấp hơn mức tối thiểu 85% lương công việc (Điều 26 BLLĐ).');
    else if (h.luongCV > 0 && num(nv.luongChinh) > 0 && num(nv.luongChinh) < h.luongCV * 0.85 - 0.5) out.loi.push('Lương thử việc ' + num(nv.luongChinh).toLocaleString('vi-VN') + ' đ < 85% lương công việc ' + h.luongCV.toLocaleString('vi-VN') + ' đ (Điều 26 BLLĐ).');
    if (h.trongHDLD === 'co' && truongHopNV(nv) === 'thuviec') out.loi.push('Thử việc ghi trong HĐLĐ thuộc diện BH bắt buộc: phải đóng BH ngay từ thời gian thử việc – đổi sang trường hợp 1.');
    return out;
  }
  function dauHieuQHLD(nv) {
    var h = hsThueNV(nv).khoan, ds = h.dauHieu.map(function (k) { return DH_QHLD.filter(function (d) { return d[0] === k; })[0][1]; });
    if (['Không xác định thời hạn', 'Xác định thời hạn'].indexOf(nv.loaiHD) >= 0) ds.unshift('Hồ sơ ghi loại hợp đồng là HĐLĐ (' + nv.loaiHD + ')');
    return { ds: ds, nguyCo: ds.length >= 2 };
  }
  function mstHopLe(m) { m = String(m || '').replace(/[\s.-]/g, ''); return /^\d{10}$|^\d{10}-?\d{3}$|^\d{12}$/.test(m); }

  /* ước tính thu nhập cả năm tại Công ty: đã nhận + dự kiến (bình quân 3 kỳ gần nhất × số tháng còn lại; hoặc số tự khai) */
  function uocTinhNam(nv, nam, bls, homNay, ttk) {
    nam = String(nam);
    var rows = (bls || []).filter(function (b) { return b.maNV === nv.maNV && String(b.ky).slice(0, 4) === nam; }).sort(function (a, b) { return a.ky < b.ky ? -1 : 1; });
    var x = { soKy: rows.length, daNhan: 0, bhDaDong: 0, daKT: 0, ktThuViec: 0, tnThuViec: 0, kyLuyTien: 0, kyKT10: 0, conLai: 0, bqTN: 0, bqBH: 0, duKien: 0, bhDuKien: 0, nguon: 'bangluong' };
    rows.forEach(function (b) {
      x.daNhan += num(b.tnChiuThue); x.bhDaDong += num(b.bhxh) + num(b.bhyt) + num(b.bhtn); x.daKT += num(b.thueTNCN);
      if (/Khấu trừ 10%/.test(String(b.cachThue || ''))) { x.kyKT10++; x.ktThuViec += num(b.thueTNCN); x.tnThuViec += num(b.tnChiuThue); } else x.kyLuyTien++;
    });
    (ttk || []).forEach(function (t) { if (t.maNV !== nv.maNV || t.trangThai !== 'Đã chi' || String(t.ngayTra || '').slice(0, 4) !== nam) return; x.daNhan += num(t.tnChiuThue); x.daKT += num(t.thue); x.ktKhoan = (x.ktKhoan || 0) + num(t.thue); x.tnKhoan = (x.tnKhoan || 0) + num(t.tnChiuThue); });   // v3.28: tiền khoán đã chi
    var cuoi = rows.length ? +String(rows[rows.length - 1].ky).slice(5, 7) : 0, hn = String(homNay || '').slice(0, 7);
    var tuThang = cuoi ? cuoi + 1 : (hn.slice(0, 4) === nam ? +hn.slice(5, 7) : hn.slice(0, 4) < nam ? 1 : 13);
    x.conLai = nv.trangThai === 'Đã nghỉ việc' ? 0 : Math.max(0, 13 - tuThang);
    var gan = rows.slice(-3);
    x.bqTN = gan.length ? gan.reduce(function (t, b) { return t + num(b.tnChiuThue); }, 0) / gan.length : num(nv.luongChinh);
    x.bqBH = gan.length ? gan.reduce(function (t, b) { return t + num(b.bhxh) + num(b.bhyt) + num(b.bhtn); }, 0) / gan.length : 0;
    x.duKien = x.daNhan + x.bqTN * x.conLai; x.bhDuKien = x.bhDaDong + x.bqBH * x.conLai;
    var hs = hsThueNV(nv);
    if (hs.duKien.thuNhapNam > 0) { x.duKien = hs.duKien.thuNhapNam; x.nguon = 'khai'; }
    return x;
  }
  function giamTruNam(nv, nam, cfg, hs) {
    var ck = cfgTheoKy(cfg, nam + '-12'), g = hs.giamTruKhac, moi = String(nam) >= '2026';
    var o = { banThan: num(ck.giamTruBanThan) * 12, npt: num(nv.soNPT) * num(ck.giamTruNPT) * 12, yTe: moi ? Math.min(g.yTe, 23000000) : 0, giaoDuc: moi ? Math.min(g.giaoDuc, 24000000) : 0, huuTri: Math.min(g.huuTri, (moi ? 3000000 : 1000000) * 12) };
    o.tong = o.banThan + o.npt + o.yTe + o.giaoDuc + o.huuTri; return o;
  }
  function bieuNam(cfg, nam) { return (cfgTheoKy(cfg, nam + '-12').bieuThue || []).map(function (b) { return [num(b[0]) * 12, b[1]]; }); }
  function hstCua(hst, ma, loai, nam) { return (hst || []).filter(function (r) { return r.maNV === ma && r.loai === loai && (!nam || String(r.nam) === String(nam)); }).sort(function (a, b) { return String(a.capNhat || a.ngayLap) < String(b.capNhat || b.ngayLap) ? 1 : -1; })[0] || null; }
  function coTaiLieu(tl, ma, loai) { return (tl || []).some(function (t) { return t.maNV === ma && t.loai === loai; }); }

  var DG_KQ = { du: 'Đủ điều kiện', khongdu: 'Không đủ điều kiện', bosung: 'Cần bổ sung hồ sơ', choxacminh: 'Chờ xác minh pháp lý / xác nhận của Kế toán trưởng' };
  function ketLuan(ds, choXM) {
    if (ds.some(function (d) { return d.ok === false; })) return 'khongdu';
    if (ds.some(function (d) { return d.ok === null; })) return 'bosung';
    return choXM ? 'choxacminh' : 'du';
  }
  /* ĐIỀU KIỆN CAM KẾT KHÔNG KHẤU TRỪ 10% – ctx: { nam, bls, hst, tl, cfg, homNay } */
  function evaluateTaxCommitmentEligibility(nv, ctx) {
    var cfg = ctx.cfg, nam = String(ctx.nam), hs = hsThueNV(nv), tt = ttThue(nv, nam + '-12'), u = uocTinhNam(nv, nam, ctx.bls, ctx.homNay, ctx.ttk), g = giamTruNam(nv, nam, cfg, hs), ds = [];
    var tnNK = hs.tnKhac.tnNoiKhacNam, sau = u.duKien + tnNK - u.bhDuKien - g.tong;
    ds.push({ ma: 'doiTuong', ten: 'Thuộc diện khấu trừ 10% (không HĐLĐ / HĐ dưới 3 tháng, khoán, thử việc)', ok: tt.cach === 'khautru10' ? true : false, ghi: tt.cach === 'khautru10' ? 'Kỳ 12/' + nam + ': khấu trừ 10%' : 'Đang tính lũy tiến (HĐLĐ từ 3 tháng) – không thuộc diện cam kết' });
    ds.push({ ma: 'cuTru', ten: 'Cá nhân cư trú', ok: hs.cuTru === 'co', ghi: hs.cuTru === 'co' ? 'Cư trú' : 'Không cư trú – không được cam kết' });
    ds.push({ ma: 'mst', ten: 'Mã số thuế / số định danh hợp lệ', ok: mstHopLe(nv.mst) ? true : null, ghi: nv.mst ? (mstHopLe(nv.mst) ? nv.mst : 'MST "' + nv.mst + '" sai định dạng') : 'Chưa có MST – phải đăng ký thuế trước khi cam kết' });
    var nk = hs.tnKhac.luongNoiKhac === 'co' || truongHopNV(nv) === 'noikhac', vlNK = !nk && hs.tnKhac.vangLai === 'co';     // v3.28: có lương HĐLĐ nơi khác → không đủ; chỉ có thu nhập 10% ở nơi khác → xét tổng thể, chờ xác minh
    ds.push({ ma: 'nguonKhac', ten: 'Chỉ có thu nhập thuộc diện khấu trừ 10% (không có tiền lương HĐLĐ lũy tiến)', ok: nk ? false : vlNK ? 'cho' : (hs.tnKhac.luongNoiKhac === '' || hs.tnKhac.vangLai === '') ? null : true, ghi: nk ? 'Có tiền lương HĐLĐ ở nơi khác – không được cam kết' : vlNK ? 'Có thu nhập khấu trừ 10% ở nơi khác – đã cộng vào tổng thu nhập; xét theo tổng thể, cần Phụ trách KT xác nhận' : (hs.tnKhac.luongNoiKhac === '' || hs.tnKhac.vangLai === '' ? 'Chưa kê khai thu nhập nơi khác (BM-02 / TAPACO-TNCN-01)' : 'Không có nguồn khác') });
    var nptHS = !num(nv.soNPT) || coTaiLieu(ctx.tl, nv.maNV, 'Hồ sơ người phụ thuộc');
    ds.push({ ma: 'npt', ten: 'Người phụ thuộc có hồ sơ', ok: nptHS ? true : null, ghi: num(nv.soNPT) + ' người' + (nptHS ? '' : ' – chưa có hồ sơ người phụ thuộc trong Tài liệu') });
    ds.push({ ma: 'thuNhap', ten: 'Thu nhập dự kiến cả năm sau giảm trừ chưa đến mức nộp thuế', ok: sau <= 0, ghi: 'Dự kiến ' + Math.round(u.duKien + tnNK).toLocaleString('vi-VN') + ' đ − BH ' + Math.round(u.bhDuKien).toLocaleString('vi-VN') + ' − giảm trừ ' + Math.round(g.tong).toLocaleString('vi-VN') + ' = ' + Math.round(sau).toLocaleString('vi-VN') + ' đ' + (u.nguon === 'khai' ? ' (số tự khai)' : ' (ước tính từ bảng lương)') });
    var hr = hstCua(ctx.hst, nv.maNV, 'camket', nam);
    ds.push({ ma: 'kyDuyet', ten: 'Bản cam kết đã ký và Phụ trách kế toán phê duyệt', ok: hr && hr.trangThai === 'Đã duyệt' ? true : null, ghi: hr ? hr.trangThai + (hr.nguoiKiemTra ? ' – ' + hr.nguoiKiemTra : '') : 'Chưa lập cam kết năm ' + nam });
    var bm = bieuMau(cfg, 'camket'), mauOK = bm && bm.trangThai === 'hienhanh';
    ds.push({ ma: 'mau', ten: 'Mẫu cam kết đang có hiệu lực', ok: mauOK ? true : 'cho', ghi: mauOK ? bm.mauSo + ' – ' + bm.canCu : (bm ? bm.mauSo + ': ' + BM_TT[bm.trangThai] : 'Chưa có mẫu') + ' – không tự động ghi nhận mẫu cũ là hợp lệ' });
    var kq = ketLuan(ds.filter(function (d) { return d.ok !== 'cho'; }), !mauOK || vlNK);
    return { ketQua: kq, ten: DG_KQ[kq], ds: ds, uoc: u, giamTru: g, sauGiamTru: sau, hoSo: hr };
  }
  /* ĐIỀU KIỆN ỦY QUYỀN QUYẾT TOÁN (NĐ 253/2026 Điều 51) */
  function evaluateTaxAuthorizationEligibility(nv, ctx) {
    var cfg = ctx.cfg, nam = String(ctx.nam), hs = hsThueNV(nv), tk = hs.tnKhac, ds = [], u = uocTinhNam(nv, nam, ctx.bls, ctx.homNay, ctx.ttk);
    var hdLD = ttThue(nv, nam + '-12').cach === 'luytien' || u.kyLuyTien > 0;
    ds.push({ ma: 'hd3t', ten: 'Có HĐLĐ từ 3 tháng trở lên tại TAPACO', ok: hdLD, ghi: hdLD ? (nv.loaiHD || 'HĐLĐ') + (u.kyKT10 ? '; có ' + u.kyKT10 + ' kỳ khấu trừ 10% (thử việc) – quyết toán cùng' : '') : 'Không có HĐLĐ từ 3 tháng (HĐ ' + (nv.loaiHD || '?') + ')' });
    var nghi = nv.trangThai === 'Đã nghỉ việc' && (!nv.ngayNghi || nv.ngayNghi.slice(0, 4) <= nam);
    ds.push({ ma: 'dangLam', ten: 'Đang làm việc tại TAPACO tại thời điểm quyết toán', ok: !nghi || tk.dieuChuyen === 'co', ghi: nghi ? (tk.dieuChuyen === 'co' ? 'Điều chuyển do sáp nhập / cùng hệ thống – được ủy quyền' : 'Đã nghỉ việc' + (nv.ngayNghi ? ' từ ' + nv.ngayNghi : '')) : 'Đang làm việc' });
    ds.push({ ma: 'cuTru', ten: 'Cá nhân cư trú', ok: hs.cuTru === 'co', ghi: hs.cuTru === 'co' ? 'Cư trú' : 'Không cư trú' });
    ds.push({ ma: 'mst', ten: 'Có mã số thuế', ok: mstHopLe(nv.mst) ? true : null, ghi: nv.mst || 'Chưa có MST' });
    var noiKhac = truongHopNV(nv) === 'noikhac' || tk.luongNoiKhac === 'co';
    ds.push({ ma: 'luongNoiKhac', ten: 'Không có tiền lương, tiền công từ HĐLĐ ở nơi khác', ok: noiKhac ? false : (tk.luongNoiKhac === '' ? null : true), ghi: noiKhac ? 'Có HĐLĐ / thu nhập tiền lương nơi khác – tự quyết toán' : tk.luongNoiKhac === '' ? 'Chưa kê khai' : 'Không có' });
    var vl = tk.vangLai, vlOK = vl === 'khong' ? true : vl === '' ? null : (tk.vangLaiBQ <= 15000000 && tk.coChungTu === 'co' && tk.qtVangLai !== 'co');
    ds.push({ ma: 'vangLai', ten: 'Thu nhập vãng lai nơi khác bình quân ≤ 15 tr/tháng, đã khấu trừ 10%, không yêu cầu quyết toán phần này', ok: vlOK, ghi: vl === 'co' ? 'Bình quân ' + tk.vangLaiBQ.toLocaleString('vi-VN') + ' đ/tháng' + (tk.coChungTu === 'co' ? ', có chứng từ khấu trừ' : ', CHƯA có chứng từ khấu trừ') + (tk.qtVangLai === 'co' ? ', yêu cầu quyết toán cả phần này' : '') : vl === 'khong' ? 'Không có' : 'Chưa kê khai' });
    var gk = hs.giamTruKhac, ytgd = gk.yTe > 0 || gk.giaoDuc > 0;
    ds.push({ ma: 'ytgd', ten: 'Không đề nghị giảm trừ y tế, giáo dục (Điều 49 – có thì tự quyết toán)', ok: !ytgd, ghi: ytgd ? 'Có kê khai chi phí y tế / giáo dục' : 'Không' });
    var kq = ketLuan(ds, false), hr = hstCua(ctx.hst, nv.maNV, 'uyquyen', nam);
    return { ketQua: kq, ten: kq === 'khongdu' ? 'Không đủ điều kiện – phải tự quyết toán nếu phát sinh nghĩa vụ' : kq === 'du' ? 'Được ủy quyền TAPACO quyết toán' : DG_KQ[kq], ds: ds, hoSo: hr, canKy: kq === 'du' && !(hr && hr.trangThai === 'Đã duyệt') };
  }
  /* TỐI ƯU THUẾ TNCN – 3 phương án: A khấu trừ mặc định · B không khấu trừ theo cam kết hợp lệ · C quyết toán / hoàn thuế cuối năm */
  function toiUuThue(nv, ctx) {
    var cfg = ctx.cfg, nam = String(ctx.nam), hs = hsThueNV(nv), u = uocTinhNam(nv, nam, ctx.bls, ctx.homNay, ctx.ttk), g = giamTruNam(nv, nam, cfg, hs);
    var ck = cfgTheoKy(cfg, ctx.homNay && ctx.homNay.slice(0, 4) === nam ? ctx.homNay.slice(0, 7) : nam + '-12'), tt = ttThue(nv, ck._ky || nam + '-12');
    var tnNK = hs.tnKhac.tnNoiKhacNam, thueNK = hs.tnKhac.thueNoiKhac, bn = bieuNam(cfg, nam);
    var tntt = Math.max(0, u.duKien + tnNK - u.bhDuKien - g.tong), thueNam = thueLuyTien(tntt, bn);
    var tnttCoBan = Math.max(0, u.duKien + tnNK - u.bhDuKien - g.banThan);           // chỉ giảm trừ bản thân
    var thangA;                                                                       // thuế tạm khấu trừ 1 tháng còn lại theo cách mặc định (bỏ qua cam kết)
    if (tt.cach === 'khautru10') thangA = u.bqTN >= num(ck.nguongVangLai) ? R(u.bqTN * num(ck.thueVangLai || 10) / 100) : 0;
    else if (tt.cach === 'kinhdoanh') thangA = 0;
    else thangA = thueLuyTien(Math.max(0, u.bqTN - u.bqBH - (tt.giamTru === 'co' ? num(ck.giamTruBanThan) + num(nv.soNPT) * num(ck.giamTruNPT) : 0)), ck.bieuThue || DEFAULT_CFG.bieuThue);
    var daKTMacDinh = u.daKT;                                                         // A: số phải khấu trừ nếu KHÔNG có cam kết (tính lại từng kỳ đã trả)
    if (tt.cach === 'khautru10') daKTMacDinh = (ctx.bls || []).filter(function (b) { return b.maNV === nv.maNV && String(b.ky).slice(0, 4) === nam; }).reduce(function (t, b) { var c = cfgTheoKy(cfg, b.ky); return t + (/Khấu trừ 10%/.test(String(b.cachThue || '')) && num(b.tnChiuThue) >= num(c.nguongVangLai) ? R(num(b.tnChiuThue) * num(c.thueVangLai || 10) / 100) : 0); }, 0);
    var A = daKTMacDinh + thangA * u.conLai, thangHT = tt.cach === 'khautru10' && tt.camKet === 'co' ? 0 : thangA, duKienKT = u.daKT + thangHT * u.conLai;
    var dgCK = tt.cach === 'khautru10' ? evaluateTaxCommitmentEligibility(nv, ctx) : null, dgUQ = evaluateTaxAuthorizationEligibility(nv, ctx);
    var B = dgCK && dgCK.ketQua === 'du' ? u.daKT : null;
    var tongKT = duKienKT + thueNK;                                                   // tổng sẽ bị khấu trừ trong năm theo trạng thái hiện tại
    var o = {
      nam: nam, uoc: u, giamTru: g, tnNoiKhac: tnNK, thueNoiKhac: thueNK, tntt: tntt, thueNam: thueNam,
      thueGiamTheoLuat: Math.max(0, thueLuyTien(tnttCoBan, bn) - thueNam),
      A: { ten: 'A. Khấu trừ thuế mặc định', khauTruConLai: thangA * u.conLai, thangA: thangA, tongKhauTru: A, nopThemQT: thueNam - A - thueNK },
      B: { ten: 'B. Không khấu trừ theo cam kết hợp lệ', apDung: B !== null, lyDo: dgCK ? dgCK.ten : 'Không áp dụng – đang tính lũy tiến (HĐLĐ từ 3 tháng)', tongKhauTru: B, nopThemQT: B !== null ? thueNam - B - thueNK : null },
      C: { ten: 'C. Quyết toán / hoàn thuế cuối năm', thueNam: thueNam, tongKhauTru: tongKT, hoan: Math.max(0, tongKT - thueNam), nopThem: Math.max(0, thueNam - tongKT), uyQuyen: dgUQ.ketQua === 'du' },
      chuaKhauTru: B !== null ? A - B : 0,          // số tiền chưa bị khấu trừ trong năm (tạm hoãn, KHÔNG phải giảm thuế)
      thuaHoan: Math.max(0, tongKT - thueNam),      // thuế nộp thừa được hoàn sau quyết toán
      dgCamKet: dgCK, dgUyQuyen: dgUQ, cachThue: tt, hoSoThieu: []
    };
    if (dgCK) dgCK.ds.forEach(function (d) { if (d.ok === null) o.hoSoThieu.push('Cam kết: ' + d.ten + ' – ' + d.ghi); });
    dgUQ.ds.forEach(function (d) { if (d.ok === null) o.hoSoThieu.push('Ủy quyền: ' + d.ten + ' – ' + d.ghi); });
    if (num(nv.soNPT) && !coTaiLieu(ctx.tl, nv.maNV, 'Hồ sơ người phụ thuộc')) o.hoSoThieu.push('Chứng từ người phụ thuộc (' + num(nv.soNPT) + ' người)');
    if ((hs.giamTruKhac.yTe || hs.giamTruKhac.giaoDuc)) o.hoSoThieu.push('Hóa đơn, chứng từ chi phí y tế / giáo dục đứng tên người nộp thuế hoặc người phụ thuộc');
    if (truongHopNV(nv) === 'noikhac' && hs.noiKhac.gtTai === '') o.hoSoThieu.push('Phiếu kê khai nơi đăng ký giảm trừ bản thân (TAPACO-TNCN-01)');
    o.ghiChu = ['Giảm khấu trừ tạm thời (phương án B) chỉ hoãn việc nộp – nghĩa vụ thuế cuối cùng vẫn là ' + Math.round(thueNam).toLocaleString('vi-VN') + ' đ.',
      u.nguon === 'khai' ? 'Thu nhập cả năm theo số CBCNV tự khai.' : 'Thu nhập cả năm ước tính: ' + u.soKy + ' kỳ đã có + ' + u.conLai + ' tháng × bình quân ' + Math.round(u.bqTN).toLocaleString('vi-VN') + ' đ.',
      'Biểu thuế năm = biểu tháng kỳ 12/' + nam + ' × 12; giảm trừ bản thân và người phụ thuộc tính đủ 12 tháng.' + (nam === '2026' ? ' Năm 2026: thu nhập 01–06 đã khấu trừ theo luật cũ, chênh lệch điều chỉnh khi quyết toán (NĐ 253/2026 Điều 70).' : '')];
    if (u.kyKT10 && u.kyLuyTien) o.ghiChu.push('Trong năm có ' + u.kyKT10 + ' kỳ khấu trừ 10% (thử việc, ' + Math.round(u.ktThuViec).toLocaleString('vi-VN') + ' đ) và ' + u.kyLuyTien + ' kỳ lũy tiến – tổng hợp cả hai khi quyết toán.');
    return o;
  }
  /* GHI CHÚ THUẾ – BHXH trên bảng lương: [{ ma, muc: 'do'|'vang'|'xanh', t, chiTiet, canCu: [mã quy định], hoSo, mau }] */
  function canhBaoThueBH(nv, bl, ctx) {
    var out = [], cfg = ctx.cfg, ky = bl && bl.ky || (ctx.homNay || '').slice(0, 7), nam = String(ctx.nam || ky.slice(0, 4)), hs = hsThueNV(nv), th = truongHopNV(nv), tt = ttThue(nv, ky);
    var add = function (ma, muc, t, chiTiet, canCu, hoSo, mau) { out.push({ ma: ma, muc: muc, t: t, chiTiet: chiTiet || '', canCu: canCu || [], hoSo: hoSo || '', mau: mau || '' }); };
    var c2 = Object.assign({}, ctx, { nam: nam });
    if (th === 'noikhac') {
      var nk = hs.noiKhac;
      if (!nk.donVi || !coTaiLieu(ctx.tl, nv.maNV, 'Sổ BHXH') && !nk.taiLieu) add('bhNoiKhac', 'vang', 'Nhân viên đang đóng BHXH tại đơn vị khác – cần kiểm tra hồ sơ tham gia bảo hiểm.', nk.donVi ? 'Đơn vị: ' + nk.donVi + ' – chưa có tài liệu xác nhận.' : 'Chưa khai đơn vị đang đóng BHXH, ngày HĐLĐ đầu tiên.', ['BHXH_NHIEUHD'], 'Phiếu xác nhận nơi tham gia BHXH khác + bản sao xác nhận quá trình đóng', 'xacnhanBH');
      if (nk.gtTai === 'noikhac' && tt.giamTru === 'co') add('gtTrung', 'do', 'Giảm trừ bản thân bị tính trùng tại 2 nơi.', 'Phiếu kê khai ghi đăng ký giảm trừ tại nơi khác nhưng trạng thái thuế đang giảm trừ tại TAPACO.', ['TNCN_NPT'], '', 'kekhaiGT');
      else if (tt.cach === 'luytien' && tt.giamTru !== 'co') add('gtNoiKhac', 'xanh', 'Đã áp dụng giảm trừ bản thân tại nơi khác – không áp dụng trùng.', '', ['TNCN_GT']);
    }
    if (tt.cach === 'luytien' && tt.giamTru === 'co') add('gtTapaco', 'xanh', 'Được giảm trừ bản thân tại TAPACO.', '', ['TNCN_GT']);
    if (tt.cach === 'kinhdoanh') add('kinhDoanh', 'vang', 'Cá nhân kinh doanh độc lập – không khấu trừ thuế tiền công; kiểm tra hóa đơn / nghĩa vụ thuế kinh doanh của cá nhân.', '', []);
    if (tt.cach === 'khautru10') {
      var dg = evaluateTaxCommitmentEligibility(nv, c2);
      if (tt.camKet === 'co') {
        if (dg.ketQua === 'khongdu') add('camKetSai', 'do', 'Cam kết không còn đúng điều kiện – cần khấu trừ 10% các lần trả tiếp theo.', dg.ds.filter(function (d) { return d.ok === false; }).map(function (d) { return d.ten + ': ' + d.ghi; }).join('; '), ['TNCN_CAMKET', 'TNCN_KT10'], '', 'camket');
        else if (dg.ketQua !== 'du') add('camKetCho', 'vang', 'Cam kết đã ký – chờ xác nhận hiệu lực mẫu.', dg.ds.filter(function (d) { return d.ok !== true; }).map(function (d) { return d.ten + ': ' + d.ghi; }).join('; '), ['TNCN_CAMKET'], '', 'camket');
      } else if (dg.sauGiamTru <= 0 && dg.uoc.duKien > 0) add('nenCamKet', 'vang', 'Thu nhập thấp hơn mức dự kiến phải nộp thuế – kiểm tra điều kiện cam kết.', 'Ước tính cả năm sau giảm trừ: ' + Math.round(dg.sauGiamTru).toLocaleString('vi-VN') + ' đ. Kết quả kiểm tra: ' + dg.ten + '.', ['TNCN_CAMKET', 'TNCN_KT10'], '', 'camket');
    }
    if (tt.cach === 'luytien' && th !== 'noikhac') {
      var uq = evaluateTaxAuthorizationEligibility(nv, c2);
      if (uq.ketQua === 'khongdu') add('uqKhong', 'vang', 'Không đủ điều kiện ủy quyền quyết toán.', uq.ds.filter(function (d) { return d.ok === false; }).map(function (d) { return d.ten + ': ' + d.ghi; }).join('; '), ['TNCN_UQ', 'TNCN_VL15']);
      else if (uq.canKy && uq.ketQua === 'du') add('uqKy', 'vang', 'Đủ điều kiện ủy quyền – cần ký biểu mẫu.', '', ['TNCN_UQ', 'TT89'], 'Mẫu 08/UQ-QTT-TNCN đã ký', 'uyquyen');
    }
    if (ctx.bls) {
      var to = toiUuThue(nv, c2);
      if (to.uoc.daKT > to.thueNam + 1000 && to.uoc.soKy) add('hoanThue', 'xanh', 'Thuế đã khấu trừ cao hơn thuế dự kiến cuối năm – xem xét hoàn thuế.', 'Đã khấu trừ ' + Math.round(to.uoc.daKT).toLocaleString('vi-VN') + ' đ; thuế dự kiến cả năm ' + Math.round(to.thueNam).toLocaleString('vi-VN') + ' đ.', ['TNCN_UQ', 'TNCN_CT70']);
    }
    if (th === 'khoan' || ['Thời vụ / Khoán việc', 'HĐ dưới 3 tháng'].indexOf(nv.loaiHD) >= 0) {
      var dh = dauHieuQHLD(nv);
      if (dh.nguyCo) add('khoanQHLD', 'do', 'Hợp đồng khoán có dấu hiệu quan hệ lao động.', dh.ds.join('; ') + '. Hợp đồng tên gọi khác nhưng có trả lương và sự quản lý, điều hành thuộc diện BHXH bắt buộc nếu từ đủ 1 tháng.', ['BHXH_DT']);
    }
    var tv = kiemTraThuViec(nv);
    tv.loi.forEach(function (l, i) { add('thuViec' + i, 'do', /vượt giới hạn/.test(l) ? 'Thời gian thử việc vượt giới hạn cho phép.' : l.split(' (')[0] + '.', l, ['BLLD_25', 'BLLD_26']); });
    tv.canhBao.forEach(function (l, i) { add('thuViecCB' + i, 'vang', l, '', ['BLLD_25']); });
    if (num(nv.soNPT) > 0 && !coTaiLieu(ctx.tl, nv.maNV, 'Hồ sơ người phụ thuộc')) add('nptThieu', 'vang', 'Còn thiếu chứng từ người phụ thuộc.', num(nv.soNPT) + ' người phụ thuộc đang được giảm trừ nhưng chưa có tài liệu "Hồ sơ người phụ thuộc".', ['TNCN_NPT'], 'Đăng ký người phụ thuộc + giấy tờ chứng minh', 'npt');
    if ((th === 'thuviec' || (th === 'huu' && !hs.huu.phuongAnBH)) && cfg.traThayBH !== 'co') add('d168', 'vang', 'Cần kiểm tra khoản trả thêm tương đương bảo hiểm.', 'NLĐ không thuộc diện BH bắt buộc; Cài đặt đang tắt "Trả thêm khoản tương đương BH" (Điều 168.3 BLLĐ).', ['BLLD_168']);
    if (!nv.mst && (tt.cach !== 'luytien' || num(bl && bl.thueTNCN) > 0)) add('mst', 'vang', 'Chưa có mã số thuế cá nhân.', '', ['ND252']);
    if (hs.cuTru === 'khong') add('khongCuTru', 'do', 'Cá nhân không cư trú – thuế tiền lương 20% trên thu nhập chịu thuế: ứng dụng chưa tự tính, kế toán kiểm tra.', '', []);
    if (th === 'khac') add('chuaTH', 'vang', 'Hồ sơ chưa chọn trường hợp tính lương (không đóng BH, chưa ghi lý do).', '', ['BHXH_DT']);
    if (false) add('huuXN', 'vang', 'Chưa xác nhận đang hưởng lương hưu (thời điểm bắt đầu hưởng).', 'Lương hưu do quỹ BHXH chi trả không chịu thuế; tiền lương TAPACO trả vẫn tính thuế.', ['BHXH_DT']);
    return out;
  }
  /* BỘ HÀM TÍNH LƯƠNG – tách theo nghiệp vụ, dùng chung kết quả tinhPhieuLuong */
  function traThayBHCua(p) { return parseJSON(p && p.chiTiet, []).filter(function (c) { return c.ma === 'TRATHAYBH'; }).reduce(function (t, c) { return t + num(c.gt); }, 0); }
  var PayrollEngine = {
    calculateGrossSalary: function (p) { return num(p.tongThuNhap); },
    calculateInsuranceEmployee: function (p) { return { bhxh: num(p.bhxh), bhyt: num(p.bhyt), bhtn: num(p.bhtn), doanPhi: num(p.doanPhi), tong: num(p.bhxh) + num(p.bhyt) + num(p.bhtn) }; },
    calculateInsuranceEmployer: function (p) { return { bhxh: num(p.bhxhDN), bhyt: num(p.bhytDN), bhtn: num(p.bhtnDN), kpcd: num(p.kpcdDN), tong: num(p.bhxhDN) + num(p.bhytDN) + num(p.bhtnDN) + num(p.kpcdDN) }; },
    calculateInsuranceCashAllowance: traThayBHCua,
    calculateTaxableIncome: function (p) { return { chiuThue: num(p.tnChiuThue), giamTru: num(p.giamTru), tinhThue: num(p.tnTinhThue) }; },
    calculatePersonalIncomeTax: function (p) { return num(p.thueTNCN); },
    calculateNetSalary: function (p) { return num(p.thucLinh); },
    calculateEmployerCost: function (p) { return num(p.tongThuNhap) + num(p.bhxhDN) + num(p.bhytDN) + num(p.bhtnDN) + num(p.kpcdDN); },
    calculateAnnualTaxSettlement: function (nam, bls, nvs, cfg, opt) { return quyetToanNam(nam, bls, nvs, cfg, opt); },
    evaluateTaxCommitmentEligibility: evaluateTaxCommitmentEligibility,
    evaluateTaxAuthorizationEligibility: evaluateTaxAuthorizationEligibility,
    tinh: function (nv, cc, cfg, o) { var p = tinhPhieuLuong(nv, cc, cfg, o); return { phieu: p, gross: PayrollEngine.calculateGrossSalary(p), bhNLD: PayrollEngine.calculateInsuranceEmployee(p), bhDN: PayrollEngine.calculateInsuranceEmployer(p), traThayBH: traThayBHCua(p), thue: PayrollEngine.calculateTaxableIncome(p), thueTNCN: p.thueTNCN, net: p.thucLinh, chiPhiDN: PayrollEngine.calculateEmployerCost(p) }; },
    /* Net → Gross: tìm mức lương thỏa thuận để thực lĩnh (đủ công, không tạm ứng / khấu trừ) bằng số NET mong muốn. o.bhTheoLuong: lương đóng BH = lương thỏa thuận */
    netToGross: function (nv, cc, cfg, o, net) {
      net = num(net); if (!(net > 0)) return null;
      var f = function (L) { var oo = Object.assign({}, o, { luong: L }); if (o.bhTheoLuong) oo.luongBH = L; return tinhPhieuLuong(Object.assign({}, nv, { luongChinh: L }), cc, cfg, oo); };
      var lo = 0, hi = net * 2 + 10000000, i = 0;
      while (f(hi).thucLinh < net && i++ < 20) hi *= 2;
      for (i = 0; i < 60 && hi - lo > 0.5; i++) { var m = (lo + hi) / 2; if (f(m).thucLinh < net) lo = m; else hi = m; }
      var L = Math.ceil(hi), p = f(L);
      return { luong: L, phieu: p, net: p.thucLinh, gross: p.tongThuNhap };
    }
  };

  /* ===================== v3.28: MODULE HỢP ĐỒNG KHOÁN VIỆC (CONTRACT-02) =====================
     Rule Engine: QUY_DINH_MD + mocLuat (ngưỡng theo ngày thực trả) · Calculation Engine: calculateMonthlyContractPayment, calculateContractorWithholdingTax
     Compliance Engine: classifyContractRelationship, evaluateEmploymentInsuranceRisk, validateTaxCommitmentDocuments, generateContractorTaxNotes
     Document Engine: BIEU_MAU_MD (BM-01…BM-08) + giao diện. AI KHÔNG quyết định số thuế – mọi số tiền tính bằng công thức dưới đây. */
  var LOAI_K = {
    K1: 'K1 – Cá nhân nhận khoán độc lập',
    K2: 'K2 – Tổ đội nhận khoán độc lập',
    K3: 'K3 – Hợp đồng khoán có bản chất HĐLĐ',
    K4: 'K4 – Cá nhân / hộ kinh doanh độc lập'
  };
  var HT_KHOAN = { thang: 'Khoán cố định theo tháng', ngay: 'Theo ngày / ca công', khoiluong: 'Theo khối lượng nghiệm thu', sanpham: 'Theo sản phẩm hoàn thành', giaidoan: 'Theo giai đoạn công trình' };
  /* bộ câu hỏi đánh giá bản chất: lời đáp 'tapaco' = dấu hiệu quan hệ lao động; 'canhan' = dấu hiệu khoán độc lập; '' = chưa rõ */
  var CH_BAN_CHAT = [
    { ma: 'q1', t: 'Ai quyết định giờ làm việc?', tapaco: 'TAPACO', canhan: 'Người nhận khoán', loi: 1 },
    { ma: 'q2', t: 'Ai phân công công việc hằng ngày?', tapaco: 'TAPACO', canhan: 'Người nhận khoán', loi: 1 },
    { ma: 'q3', t: 'Ai quản lý, điều hành người thực hiện?', tapaco: 'TAPACO', canhan: 'Người nhận khoán', loi: 1 },
    { ma: 'q4', t: 'Người nhận khoán có tự lựa chọn nhân công không?', tapaco: 'Không', canhan: 'Có' },
    { ma: 'q5', t: 'Người nhận khoán có tự quyết định biện pháp thực hiện không?', tapaco: 'Không', canhan: 'Có' },
    { ma: 'q6', t: 'TAPACO trả theo kết quả nghiệm thu hay trả lương cố định?', tapaco: 'Lương cố định', canhan: 'Theo nghiệm thu' },
    { ma: 'q7', t: 'Người nhận khoán có chịu trách nhiệm sửa chữa sản phẩm không đạt?', tapaco: 'Không', canhan: 'Có' },
    { ma: 'q8', t: 'Người nhận khoán có chịu rủi ro, chi phí tổ chức thực hiện?', tapaco: 'Không', canhan: 'Có' },
    { ma: 'q9', t: 'Công việc có lặp lại liên tục như một vị trí nhân viên?', tapaco: 'Có', canhan: 'Không' },
    { ma: 'q10', t: 'TAPACO có trực tiếp chấm công, xử lý kỷ luật lao động?', tapaco: 'Có', canhan: 'Không', loi: 1 },
    { ma: 'q11', t: 'Hồ sơ hợp đồng, nghiệm thu, thanh toán có nhất quán với thực tế?', tapaco: 'Không', canhan: 'Có' }
  ];
  var CH_TO_DOI = [
    { ma: 'tuCach', t: 'Đội trưởng ký hợp đồng với tư cách', o: { canhan: 'Cá nhân', daidien: 'Đại diện hợp pháp (có giấy ủy quyền)' } },
    { ma: 'nguoiNhan', t: 'Người thực sự nhận thu nhập', o: { doitruong: 'Đội trưởng', tungtv: 'Từng thành viên' } },
    { ma: 'traTrucTiep', t: 'TAPACO trả tiền', o: { tungtv: 'Trực tiếp cho từng thành viên', doitruong: 'Cho đội trưởng / chủ thể nhận khoán' } },
    { ma: 'coHoSo', t: 'Có hợp đồng, giấy ủy quyền, danh sách thành viên, chứng từ thanh toán', o: { co: 'Có đủ', khong: 'Chưa đủ' } },
    { ma: 'cheDo', t: 'Hoạt động thuộc chế độ thuế', o: { tiencong: 'Tiền công, tiền thù lao', kinhdoanh: 'Kinh doanh' } }
  ];
  var KL_PL = {
    doclap: 'Khoán độc lập có cơ sở xác nhận',
    dauhieu: 'Có dấu hiệu quan hệ lao động – cần kiểm tra',
    qhld: 'Thuộc quan hệ lao động – chuyển sang quy trình HĐLĐ',
    kinhdoanh: 'Có dấu hiệu hoạt động kinh doanh – chuyển sang nhánh thuế kinh doanh',
    thieu: 'Chưa đủ dữ liệu – tạm dừng phê duyệt chế độ thuế / BHXH'
  };
  function classifyContractRelationship(hd) {
    var dg = parseJSON(hd && hd.danhGia, {}) || {}, ld = [], dl = [], chua = [], loi = 0, lyDo = [];
    CH_BAN_CHAT.forEach(function (q) { var v = dg[q.ma]; if (v === 'tapaco') { ld.push(q); if (q.loi) loi++; } else if (v === 'canhan') dl.push(q); else chua.push(q); });
    var kd = dg.kinhDoanh === 'co' || dg.cheDo === 'kinhdoanh' || (hd && hd.loaiK === 'K4');
    var kq;
    if (kd) { kq = 'kinhdoanh'; lyDo.push(dg.kinhDoanh === 'co' ? 'Có đăng ký kinh doanh / hộ kinh doanh' : dg.cheDo === 'kinhdoanh' ? 'Tổ đội khai thuộc chế độ thuế kinh doanh' : 'Hợp đồng phân loại K4'); }
    else if (chua.length > 3) { kq = 'thieu'; lyDo.push('Còn ' + chua.length + '/11 câu hỏi chưa trả lời'); }
    else if (loi >= 3 && (dg.q6 === 'tapaco' || dg.q9 === 'tapaco')) { kq = 'qhld'; lyDo.push('TAPACO quyết định giờ làm, phân công, quản lý, chấm công / kỷ luật (' + loi + '/4 dấu hiệu quản lý, điều hành) và trả công ' + (dg.q6 === 'tapaco' ? 'cố định' : 'cho công việc lặp lại như nhân viên')); }
    else if (ld.length >= 3 || loi >= 2) { kq = 'dauhieu'; lyDo.push(ld.length + ' dấu hiệu quan hệ lao động: ' + ld.map(function (q) { return q.t.replace(/\?$/, '') + ' → ' + q.tapaco; }).join('; ')); }
    else { kq = 'doclap'; lyDo.push(dl.length + ' dấu hiệu khoán độc lập' + (ld.length ? '; ' + ld.length + ' dấu hiệu đơn lẻ không đủ kết luận quan hệ lao động' : '')); }
    if (hd && hd.loaiK === 'K3' && kq !== 'qhld') { lyDo.push('Hợp đồng đã được phân loại K3 (bản chất HĐLĐ)'); kq = 'qhld'; }
    return { ketQua: kq, ten: KL_PL[kq], lyDo: lyDo, dauHieuLD: ld.map(function (q) { return q.t; }), dauHieuDL: dl.map(function (q) { return q.t; }), chuaTraLoi: chua.map(function (q) { return q.t; }), canCu: ['BLLD_13', 'BHXH_DT'] };
  }
  function soThangHD(hd) { var a = parseD(hd.tuNgay), b = parseD(hd.denNgay); if (!a || !b) return 0; return (b.y - a.y) * 12 + (b.m - a.m) + (b.d >= a.d - 1 ? 1 : 0); }
  function evaluateEmploymentInsuranceRisk(hd) {
    var pl = classifyContractRelationship(hd), th = soThangHD(hd);
    if (pl.ketQua === 'qhld') return { muc: 'do', khongBH: false, t: 'Thuộc quan hệ lao động' + (th >= 1 || !th ? ' – thời hạn ' + (th ? th + ' tháng' : 'chưa rõ') + ': thuộc diện BHXH bắt buộc, KHÔNG xác nhận "không đóng bảo hiểm".' : ' dưới 1 tháng.'), canCu: ['BHXH_DT', 'BLLD_13'] };
    if (pl.ketQua === 'dauhieu') return { muc: 'vang', khongBH: null, t: 'Có dấu hiệu chịu sự quản lý, điều hành của TAPACO – rà soát quan hệ lao động và BHXH trước khi xác nhận không đóng BH.', canCu: ['BHXH_DT', 'BLLD_13'] };
    if (pl.ketQua === 'thieu') return { muc: 'vang', khongBH: null, t: 'Chưa đủ dữ liệu đánh giá bản chất – tạm dừng xác nhận chế độ BHXH.', canCu: ['BHXH_DT'] };
    return { muc: 'xanh', khongBH: true, t: pl.ketQua === 'kinhdoanh' ? 'Hoạt động kinh doanh độc lập – không thuộc diện BHXH bắt buộc theo hồ sơ hiện có.' : 'Khoán độc lập theo hồ sơ hiện có – không thuộc diện BHXH bắt buộc tại TAPACO.', canCu: ['BHXH_DT'] };
  }
  var CHE_DO_K = { tiencong10: 'Tiền công – khấu trừ 10% khi trả từ ngưỡng', kinhdoanh: 'Kinh doanh – không khấu trừ 10% tiền công (cá nhân/hộ tự khai theo NĐ 68/2026)', khongcutru: 'Cá nhân không cư trú – 20% (cần Phụ trách KT xác nhận)', hdld: 'Bản chất HĐLĐ – tạm tính 10%, chuyển quy trình HĐLĐ', thieu: 'Chưa đủ dữ liệu – tạm tính 10%, tạm dừng phê duyệt' };
  function classifyContractorTaxRegime(hd, nv) {
    var pl = classifyContractRelationship(hd);
    if (pl.ketQua === 'kinhdoanh') return 'kinhdoanh';
    if (pl.ketQua === 'qhld') return 'hdld';
    if (pl.ketQua === 'thieu') return 'thieu';
    if (hsThueNV(nv).cuTru === 'khong') return 'khongcutru';
    return 'tiencong10';
  }
  /* cam kết: kiểm tra hồ sơ theo năm của NGÀY THỰC TRẢ – ELIGIBLE / PENDING / LEGAL_REVIEW / EXPIRED / REVOKED */
  var CK_TT = { ELIGIBLE: 'Cam kết hợp lệ', INELIGIBLE: 'Không đủ điều kiện', PENDING: 'Chờ hồ sơ', LEGAL_REVIEW: 'Cần xác minh pháp lý', EXPIRED: 'Cam kết hết hiệu lực', REVOKED: 'Cam kết đã thu hồi', NONE: 'Chưa có cam kết' };
  function validateTaxCommitmentDocuments(nv, nam, hst, cfg) {
    var ds = (hst || []).filter(function (r) { return r.maNV === nv.maNV && r.loai === 'camket'; }), r = ds.filter(function (x) { return String(x.nam) === String(nam); }).sort(function (a, b) { return String(a.capNhat) < String(b.capNhat) ? 1 : -1; })[0];
    if (!r) return { ma: ds.some(function (x) { return String(x.nam) < String(nam); }) ? 'EXPIRED' : 'NONE', ok: false, lyDo: ds.length ? 'Cam kết năm ' + ds.map(function (x) { return x.nam; }).join(', ') + ' không áp dụng cho năm ' + nam : 'Chưa có cam kết năm ' + nam };
    if (r.trangThai === 'Thu hồi') return { ma: 'REVOKED', ok: false, lyDo: 'Cam kết đã bị thu hồi' + (r.ghiChuDuyet ? ': ' + r.ghiChuDuyet : ''), hs: r };
    if (!r.taiLieuId) return { ma: 'PENDING', ok: false, lyDo: 'Cam kết chưa có bản ký (thiếu chữ ký)', hs: r };
    if (r.trangThai === 'Chờ xác minh mẫu') return { ma: 'LEGAL_REVIEW', ok: false, lyDo: 'Mẫu cam kết đang chờ xác minh hiệu lực', hs: r };
    if (r.trangThai !== 'Đã duyệt') return { ma: 'PENDING', ok: false, lyDo: 'Cam kết chưa được phê duyệt (' + r.trangThai + ')', hs: r };
    if (!mstHopLe(nv.mst)) return { ma: 'PENDING', ok: false, lyDo: 'Chưa có mã số thuế hợp lệ – không lấy việc thiếu MST làm căn cứ miễn khấu trừ', hs: r };
    return { ma: 'ELIGIBLE', ok: true, lyDo: 'Cam kết năm ' + nam + ' đã ký, Phụ trách kế toán duyệt ' + (r.ngayKiemTra || ''), hs: r };
  }
  /* CALCULATION ENGINE – 1 lần chi trả. p: { loai: tamung|nghiemthu|quyettoan, giaTriNT, thuNhapBS, giamTruHD, tamUng, tuTinhThue, thuHoiTU, thuHoiTUThue, khauTruKhac, ngayTra } */
  function calculateContractorWithholdingTax(tn, ngayTra, cheDo, ck, cfg) {
    var ky = String(ngayTra || '').slice(0, 7), c = cfgTheoKy(cfg, ky || '2026-07'), nguong = num(c.nguongVangLai), ts = num(c.thueVangLai || 10);
    var o = { nguong: nguong, ky: ky, thueSuat: 0, thue: 0, coSo: tn, quyTac: '', canCu: [] };
    if (!ky) { o.quyTac = 'Chưa có ngày thực trả – chưa xác định được ngưỡng'; return o; }
    if (cheDo === 'kinhdoanh') { o.quyTac = 'Thu nhập kinh doanh: không khấu trừ 10% tiền công, không áp giảm trừ bản thân (NĐ 68/2026)'; o.canCu = ['ND68']; return o; }
    if (cheDo === 'khongcutru') { o.thueSuat = 20; o.thue = R(tn * 0.2); o.quyTac = 'Cá nhân không cư trú: 20% × ' + tn.toLocaleString('vi-VN') + ' (không áp ngưỡng, không giảm trừ) – cần Phụ trách KT xác nhận'; o.canCu = ['TNCN_KCT']; return o; }
    if (tn < nguong) { o.quyTac = 'Trả ' + tn.toLocaleString('vi-VN') + ' < ngưỡng ' + nguong.toLocaleString('vi-VN') + ' đ/lần (theo ngày trả ' + ky.slice(5) + '/' + ky.slice(0, 4) + '): không bắt buộc khấu trừ'; o.canCu = ['TNCN_KT10']; return o; }
    if (ck && ck.ok && cheDo === 'tiencong10') { o.quyTac = 'Có cam kết hợp lệ năm ' + ky.slice(0, 4) + ': tạm không khấu trừ'; o.canCu = ['TNCN_CAMKET']; o.camKet = true; return o; }
    o.thueSuat = ts; o.thue = R(tn * ts / 100);
    o.quyTac = ts + '% × ' + tn.toLocaleString('vi-VN') + ' = ' + o.thue.toLocaleString('vi-VN') + ' (trả ≥ ngưỡng ' + nguong.toLocaleString('vi-VN') + ' đ/lần áp dụng tại ngày trả ' + String(ngayTra).split('-').reverse().join('/') + ')' + (ck && !ck.ok && ck.ma !== 'NONE' ? '; cam kết chưa hợp lệ: ' + ck.lyDo : '') + (cheDo === 'hdld' || cheDo === 'thieu' ? ' – TẠM TÍNH' : '');
    o.canCu = ['TNCN_KT10'];
    return o;
  }
  function calculateMonthlyContractPayment(p, hd, nv, ctx) {
    var cheDo = classifyContractorTaxRegime(hd, nv), nam = String(p.ngayTra || '').slice(0, 4) || String(ctx.homNay || '').slice(0, 4);
    var ck = validateTaxCommitmentDocuments(nv, nam, ctx.hst, ctx.cfg), tu = p.loai === 'tamung';
    var gross = tu ? num(p.tamUng) : num(p.giaTriNT) + num(p.thuNhapBS) - num(p.giamTruHD);
    var tn = tu ? (p.tuTinhThue === 'co' ? gross : 0) : Math.max(0, gross - num(p.thuHoiTUThue));
    var t = tu && p.tuTinhThue !== 'co' ? { thueSuat: 0, thue: 0, quyTac: 'Tạm ứng chưa phải thu nhập đã chi trả – thuế tính khi nghiệm thu / thanh toán (khoản tạm ứng thu hồi không bị tính thuế 2 lần)', canCu: ['TNCN_KT10'], nguong: 0 } : calculateContractorWithholdingTax(tn, p.ngayTra, cheDo, ck, ctx.cfg);
    var thucTra = tu ? gross - t.thue : gross - t.thue - num(p.thuHoiTU) - num(p.khauTruKhac);
    return { tienPhaiTra: gross, tnChiuThue: tn, thueSuat: t.thueSuat, thue: t.thue, thucTra: thucTra, cheDoThue: cheDo, camKet: ck, quyTac: t.quyTac, canCu: t.canCu, nguong: t.nguong,
      congThuc: tu ? 'Tạm ứng ' + gross.toLocaleString('vi-VN') + (t.thue ? ' − thuế ' + t.thue.toLocaleString('vi-VN') : '') + ' = thực trả ' + thucTra.toLocaleString('vi-VN')
        : 'Gross = nghiệm thu ' + num(p.giaTriNT).toLocaleString('vi-VN') + ' + bổ sung ' + num(p.thuNhapBS).toLocaleString('vi-VN') + ' − giảm trừ HĐ ' + num(p.giamTruHD).toLocaleString('vi-VN') + ' = ' + gross.toLocaleString('vi-VN') + ' · TN tính khấu trừ = gross − tạm ứng đã chịu thuế thu hồi ' + num(p.thuHoiTUThue).toLocaleString('vi-VN') + ' = ' + tn.toLocaleString('vi-VN') + ' · Thuế: ' + t.quyTac + ' · Thực trả = ' + gross.toLocaleString('vi-VN') + ' − thuế ' + t.thue.toLocaleString('vi-VN') + ' − thu hồi tạm ứng ' + num(p.thuHoiTU).toLocaleString('vi-VN') + ' − khấu trừ khác ' + num(p.khauTruKhac).toLocaleString('vi-VN') + ' = ' + thucTra.toLocaleString('vi-VN') };
  }
  /* hợp đồng NET: tìm giá trị nghiệm thu để thực trả = số thỏa thuận */
  function khoanNetToGross(p, hd, nv, ctx, net) {
    net = num(net); var f = function (x) { return calculateMonthlyContractPayment(Object.assign({}, p, { giaTriNT: x }), hd, nv, ctx).thucTra; };
    var x1 = net + num(p.thuHoiTU) + num(p.khauTruKhac) - num(p.thuNhapBS) + num(p.giamTruHD);       // không phát sinh thuế (dưới ngưỡng / cam kết / kinh doanh)
    if (f(x1) >= net) return { giaTriNT: x1, kq: calculateMonthlyContractPayment(Object.assign({}, p, { giaTriNT: x1 }), hd, nv, ctx) };
    var lo = x1, hi = x1 * 2 + 1e6, i;                                                               // vùng có khấu trừ: thực trả tăng đơn điệu theo gross
    for (i = 0; i < 60 && hi - lo > 0.5; i++) { var m = (lo + hi) / 2; if (f(m) < net) lo = m; else hi = m; }
    var x = Math.floor(hi); while (f(x) < net) x++; return { giaTriNT: x, kq: calculateMonthlyContractPayment(Object.assign({}, p, { giaTriNT: x }), hd, nv, ctx) };
  }
  /* tạm ứng còn phải thu hồi theo hợp đồng – tách phần đã chịu thuế để tránh tính thuế 2 lần */
  function tamUngTon(hdId, ttk, boQua) {
    var o = { daUng: 0, daUngThue: 0, thuHoi: 0, thuHoiThue: 0 };
    (ttk || []).forEach(function (x) { if (x.hdId !== hdId || x.id === boQua || x.trangThai === 'Hủy') return;
      if (x.loai === 'tamung') { o.daUng += num(x.tamUng); if (num(x.tnChiuThue) > 0) o.daUngThue += num(x.tamUng); } else { o.thuHoi += num(x.thuHoiTU); o.thuHoiThue += num(x.thuHoiTUThue); } });
    o.con = o.daUng - o.thuHoi; o.conThue = o.daUngThue - o.thuHoiThue; return o;
  }
  function generateContractorPaymentStatement(hd, ttk) {
    var ds = (ttk || []).filter(function (x) { return x.hdId === hd.id && x.trangThai !== 'Hủy'; }), chi = function (x) { return x.trangThai === 'Đã chi'; }, NT = function (x) { return x.loai !== 'tamung'; };
    var s = function (a, f) { return a.reduce(function (t, x) { return t + num(typeof f === 'function' ? f(x) : x[f]); }, 0); };
    var tu = tamUngTon(hd.id, ttk);
    var o = { giaTriHD: num(hd.giaTriHD), daNghiemThu: s(ds.filter(NT), 'giaTriNT'), phaiTraLK: s(ds.filter(NT), 'tienPhaiTra'), daChiGross: s(ds.filter(function (x) { return chi(x) && NT(x); }), 'tienPhaiTra'),
      daChiTien: s(ds.filter(chi), 'thucTra'), tamUngChuaQT: tu.con, tnDaTra: s(ds.filter(chi), 'tnChiuThue'), thueDaKT: s(ds.filter(chi), 'thue'), soDot: ds.length };
    o.congNo = o.phaiTraLK - o.daChiGross;
    o.conLaiHD = Math.max(0, o.giaTriHD - o.daNghiemThu);
    return o;
  }
  /* DỰ BÁO THU NHẬP CẢ NĂM – A theo hợp đồng · B theo lịch sử · C theo khai báo · D tổng hợp */
  function estimateAnnualContractorIncome(nv, nam, ctx) {
    nam = String(nam); var hs = hsThueNV(nv), hds = (ctx.hdk || []).filter(function (h) { return h.maNV === nv.maNV && h.trangThai !== 'Hủy'; });
    var tt = (ctx.ttk || []).filter(function (x) { return x.maNV === nv.maNV && x.trangThai !== 'Hủy' && String(x.ngayTra || '').slice(0, 4) === nam; });
    var daTra = tt.filter(function (x) { return x.trangThai === 'Đã chi'; }), sum = function (a, k) { return a.reduce(function (t, x) { return t + num(x[k]); }, 0); };
    var bl = (ctx.bls || []).filter(function (b) { return b.maNV === nv.maNV && String(b.ky).slice(0, 4) === nam; });
    var o = { daNhanKhoan: sum(daTra, 'tnChiuThue'), daNhanLuong: sum(bl, 'tnChiuThue'), thueDaKT: sum(daTra, 'thue') + sum(bl, 'thueTNCN'), choChi: sum(tt.filter(function (x) { return x.trangThai !== 'Đã chi'; }), 'tnChiuThue'), canhBao: [] };
    o.daNhan = o.daNhanKhoan + o.daNhanLuong;
    var hn = String(ctx.homNay || '').slice(0, 7), thangNay = hn.slice(0, 4) === nam ? +hn.slice(5, 7) : hn.slice(0, 4) < nam ? 0 : 12;
    var conThang = Math.max(0, 12 - thangNay);
    // A: giá trị còn lại của hợp đồng, phân bổ theo số tháng còn hiệu lực trong năm
    o.A = o.daNhan + o.choChi + hds.reduce(function (t, h) { var st = generateContractorPaymentStatement(h, ctx.ttk), den = parseD(h.denNgay), conHD = den ? Math.max(0, (den.y - +nam) * 12 + den.m - thangNay) : conThang; var trongNam = Math.min(conThang, conHD); return t + (conHD > 0 ? st.conLaiHD * trongNam / conHD : 0); }, 0);
    // B: bình quân tháng đã phát sinh × tháng còn lại
    var thangCo = {}; daTra.forEach(function (x) { thangCo[String(x.ngayTra).slice(0, 7)] = 1; }); bl.forEach(function (b) { thangCo[b.ky] = 1; });
    var soT = Object.keys(thangCo).length; o.bqThang = soT ? o.daNhan / soT : 0; o.B = o.daNhan + o.choChi + o.bqThang * conThang;
    if (soT && soT < 3) o.canhBao.push('Dự báo theo lịch sử chỉ dựa vào ' + soT + ' tháng – sai số lớn');
    o.C = hs.duKien.thuNhapNam || 0;
    o.noiKhac = hs.tnKhac.tnNoiKhacNam; o.thueNoiKhac = hs.tnKhac.thueNoiKhac;
    o.daKhaiNoiKhac = hs.tnKhac.luongNoiKhac !== '' && hs.tnKhac.vangLai !== '';
    if (!o.daKhaiNoiKhac) o.canhBao.push('Chưa có khai báo thu nhập nơi khác – không giả định thu nhập tại TAPACO là toàn bộ thu nhập cả năm');
    o.D = (o.C > 0 ? o.C : Math.max(o.A, o.B)) + o.noiKhac;
    o.cheDo = o.C > 0 ? 'C+D' : 'A/B+D';
    o.thangThuNhap = Object.keys(thangCo).sort();
    return o;
  }
  /* GIẢM TRỪ GIA CẢNH CẢ NĂM: bản thân 12 tháng (quyết toán); người phụ thuộc theo tháng đủ điều kiện (hsThue.nptDs) */
  function calculateAnnualPersonalDeductions(nv, nam, cfg) {
    var hs = hsThueNV(nv), ck = cfgTheoKy(cfg, nam + '-12'), g = giamTruNam(nv, nam, cfg, hs), ds = hs.nptDs, th = 0, thieu = 0;
    if (ds.length) {
      ds.forEach(function (p) { for (var m = 1; m <= 12; m++) { var k = nam + '-' + (m < 10 ? '0' : '') + m; if ((!p.tu || p.tu <= k) && (!p.den || p.den >= k)) { if (p.coHoSo === 'co') th++; else thieu++; } } });
      g.npt = th * num(ck.giamTruNPT); g.tong = g.banThan + g.npt + g.yTe + g.giaoDuc + g.huuTri;
    } else th = num(nv.soNPT) * 12;
    g.thangNPT = th; g.thangNPTThieuHS = thieu; g.nguonNPT = ds.length ? 'theo tháng đủ điều kiện' : 'số NPT × 12 tháng (chưa khai tháng bắt đầu)';
    return g;
  }
  function estimateAnnualPersonalIncomeTax(thuNhap, bh, g, cfg, nam) { var tntt = Math.max(0, thuNhap - bh - g.tong); return { tntt: tntt, thue: thueLuyTien(tntt, bieuNam(cfg, nam)) }; }
  function estimatePotentialTaxRefund(thueNam, daKT, thueNoiKhac) { var t = daKT + num(thueNoiKhac); return { tongDaKT: t, hoan: Math.max(0, t - thueNam), nopThem: Math.max(0, thueNam - t) }; }
  /* QUYỀN LỢI THUẾ CỦA NGƯỜI NHẬN KHOÁN (10 chỉ tiêu) + điều kiện cam kết theo tổng thu nhập */
  function evaluateTaxCommitment(nv, ctx) {
    var nam = String(ctx.nam), hs = hsThueNV(nv), u = estimateAnnualContractorIncome(nv, nam, ctx), g = calculateAnnualPersonalDeductions(nv, nam, ctx.cfg), ds = [], tk = hs.tnKhac;
    var tong = hs.duKien.thuNhapNam > 0 ? hs.duKien.thuNhapNam + tk.tnNoiKhacNam : u.D, sau = tong - g.tong;
    ds.push({ ma: 'cuTru', ten: 'Là cá nhân cư trú', ok: hs.cuTru === 'co' ? true : false, ghi: hs.cuTru === 'co' ? 'Hồ sơ: cư trú' : 'Không cư trú – nhánh thuế riêng' });
    var hds = (ctx.hdk || []).filter(function (h) { return h.maNV === nv.maNV; }), cd = hds.map(function (h) { return classifyContractorTaxRegime(h, nv); });
    ds.push({ ma: 'dien10', ten: 'Thu nhập thuộc diện khấu trừ 10%', ok: cd.some(function (c) { return c === 'tiencong10'; }) || (!hds.length && ttThue(nv, nam + '-12').cach === 'khautru10') ? true : cd.some(function (c) { return c === 'thieu'; }) ? null : false, ghi: hds.length ? cd.map(function (c) { return CHE_DO_K[c]; }).join('; ') : ttThue(nv, nam + '-12').moTa });
    ds.push({ ma: 'mst', ten: 'Có mã số thuế', ok: mstHopLe(nv.mst) ? true : null, ghi: nv.mst || 'Chưa có MST' });
    var luongNK = tk.luongNoiKhac === 'co' || truongHopNV(nv) === 'noikhac';
    ds.push({ ma: 'noiKhac', ten: 'Thu nhập nơi khác', ok: luongNK ? false : (tk.luongNoiKhac === '' || tk.vangLai === '') ? null : tk.vangLai === 'co' ? 'xm' : true,
      ghi: luongNK ? 'Có tiền lương HĐLĐ (lũy tiến) ở nơi khác – không đáp ứng điều kiện chỉ có thu nhập thuộc diện khấu trừ 10%' : tk.vangLai === 'co' ? 'Có thu nhập khấu trừ 10% ở nơi khác (' + tk.tnNoiKhacNam.toLocaleString('vi-VN') + ' đ/năm) – đã cộng vào tổng thu nhập; cần Phụ trách KT xác nhận điều kiện theo quy định hiện hành' : (tk.luongNoiKhac === '' || tk.vangLai === '') ? 'Chưa kê khai (BM-02)' : 'Không có' });
    ds.push({ ma: 'npt', ten: 'Người phụ thuộc: điều kiện, thời gian, hồ sơ', ok: g.thangNPTThieuHS ? null : true, ghi: g.thangNPT + ' tháng NPT đủ hồ sơ (' + g.nguonNPT + ')' + (g.thangNPTThieuHS ? '; ' + g.thangNPTThieuHS + ' tháng chưa đủ hồ sơ – chưa tính' : '') });
    ds.push({ ma: 'sauGT', ten: 'Thu nhập cả năm sau giảm trừ chưa đến mức nộp thuế', ok: sau <= 0, ghi: Math.round(tong).toLocaleString('vi-VN') + ' − giảm trừ ' + Math.round(g.tong).toLocaleString('vi-VN') + ' = ' + Math.round(sau).toLocaleString('vi-VN') + ' đ' });
    var bm = bieuMau(ctx.cfg, 'camket'), mauOK = bm && bm.trangThai === 'hienhanh', hsCK0 = validateTaxCommitmentDocuments(nv, nam, ctx.hst, ctx.cfg);
    ds.push({ ma: 'mau', ten: 'Mẫu cam kết có giá trị áp dụng', ok: mauOK || (hsCK0.ok) ? true : 'xm', ghi: bm ? bm.mauSo + ' – ' + BM_TT[bm.trangThai] : '' });
    var hsCK = validateTaxCommitmentDocuments(nv, nam, ctx.hst, ctx.cfg);
    ds.push({ ma: 'kyDuyet', ten: 'Chữ ký và phê duyệt trước thanh toán', ok: hsCK.ok ? true : null, ghi: CK_TT[hsCK.ma] + ' – ' + hsCK.lyDo });
    var ma;
    if (hsCK.ma === 'REVOKED' || hsCK.ma === 'EXPIRED') ma = hsCK.ma;
    else if (ds.some(function (d) { return d.ok === false; })) ma = 'INELIGIBLE';
    else if (ds.some(function (d) { return d.ok === null && d.ma !== 'kyDuyet'; })) ma = 'PENDING';
    else if (ds.some(function (d) { return d.ok === 'xm'; })) ma = 'LEGAL_REVIEW';
    else ma = hsCK.ok ? 'ELIGIBLE' : 'ELIGIBLE';
    return { ma: ma, ten: ma === 'ELIGIBLE' ? (hsCK.ok ? 'Đủ điều kiện – cam kết đã ký, duyệt' : 'Đủ điều kiện về dữ liệu – được lập cam kết; chỉ có hiệu lực sau khi ký và duyệt') : CK_TT[ma], ds: ds, tongThuNhap: tong, giamTru: g, sauGiamTru: sau, hoSo: hsCK, uoc: u };
  }
  function evaluateTaxSettlementAuthorization(nv, ctx) {
    var dg = evaluateTaxAuthorizationEligibility(nv, ctx), coKhoan = (ctx.hdk || []).some(function (h) { return h.maNV === nv.maNV; });
    var coLD = dg.ds[0] && dg.ds[0].ok === true;
    if (!coLD) return { ketQua: 'khongdu', ten: 'Chỉ có hợp đồng khoán độc lập – không đủ điều kiện ủy quyền TAPACO quyết toán theo cơ chế HĐLĐ; cá nhân tự quyết toán nếu thuộc diện cần hoặc muốn hoàn thuế', ds: dg.ds, tuQT: true };
    var o = Object.assign({}, dg);
    if (coKhoan) o.phamVi = 'Đã chuyển sang HĐLĐ từ 3 tháng: xét đủ điều kiện Điều 51 NĐ 253/2026. Phạm vi quyết toán thay gồm thu nhập TAPACO trả; khoản khoán đã khấu trừ 10% trước khi ký HĐLĐ – Phụ trách KT xác nhận phạm vi.';
    o.chuY = 'Ủy quyền ký hợp đồng / nhận tiền thay tổ đội khác hoàn toàn với ủy quyền quyết toán thuế TNCN.';
    return o;
  }
  function quyenLoiThueKhoan(nv, ctx) {
    var nam = String(ctx.nam), ck = evaluateTaxCommitment(nv, ctx), u = ck.uoc, g = ck.giamTru, bh = (ctx.bls || []).filter(function (b) { return b.maNV === nv.maNV && String(b.ky).slice(0, 4) === nam; }).reduce(function (t, b) { return t + num(b.bhxh) + num(b.bhyt) + num(b.bhtn); }, 0);
    var tn = ck.tongThuNhap, e = estimateAnnualPersonalIncomeTax(tn, bh, g, ctx.cfg, nam), r = estimatePotentialTaxRefund(e.thue, u.thueDaKT, u.thueNoiKhac), uq = evaluateTaxSettlementAuthorization(nv, ctx);
    var thieu = ck.ds.filter(function (d) { return d.ok === null; }).map(function (d) { return d.ten + ': ' + d.ghi; }).concat(u.canhBao);
    return { nam: nam, tongThuNhap: tn, giamTru: g.tong, tntt: e.tntt, thueNam: e.thue, thueTapaco: u.thueDaKT, thueNoiKhac: u.thueNoiKhac, nopThua: r.hoan, deNghiHoan: r.hoan, nopThem: r.nopThem, hoSoThieu: thieu, camKet: ck, uyQuyen: uq, uoc: u, bh: bh,
      huongDan: r.hoan > 0 ? 'Khoản thuế đã khấu trừ là số tạm nộp. Cá nhân có thể thực hiện thủ tục quyết toán (mẫu 02/QTT-TNCN trên thuedientu.gdt.gov.vn hoặc ứng dụng eTax Mobile) để đề nghị hoàn số thuế đã nộp thừa nếu đáp ứng điều kiện; TAPACO cung cấp chứng từ khấu trừ và phiếu tổng hợp BM-06.' : r.nopThem > 0 ? 'Dự kiến phải nộp thêm khi quyết toán – cá nhân tự quyết toán (02/QTT-TNCN) hoặc ủy quyền nếu đủ điều kiện.' : 'Thuế đã khấu trừ khớp nghĩa vụ dự kiến.' };
  }
  /* ghi chú tự động cho 1 dòng thanh toán khoán */
  function generateContractorTaxNotes(x, hd, nv, ctx) {
    var out = [], add = function (ma, muc, t, canCu) { out.push({ ma: ma, muc: muc, t: t, chiTiet: '', canCu: canCu || [], hoSo: '', mau: '' }); };
    var cd = classifyContractorTaxRegime(hd, nv), pl = classifyContractRelationship(hd), ck = validateTaxCommitmentDocuments(nv, String(x.ngayTra || ctx.homNay).slice(0, 4), ctx.hst, ctx.cfg), tn = num(x.tnChiuThue), ky = String(x.ngayTra || '').slice(0, 7), ng = num(cfgTheoKy(ctx.cfg, ky || '2026-07').nguongVangLai);
    if (pl.ketQua === 'qhld') add('qhld', 'do', 'Hợp đồng khoán có bản chất quan hệ lao động – chuyển sang quy trình HĐLĐ; không xác nhận "không đóng bảo hiểm".', ['BLLD_13', 'BHXH_DT']);
    else if (pl.ketQua === 'dauhieu') add('dauhieu', 'vang', 'Hợp đồng khoán theo tháng có dấu hiệu chịu sự quản lý, điều hành của TAPACO. Cần rà soát quan hệ lao động và BHXH.', ['BLLD_13', 'BHXH_DT']);
    else if (pl.ketQua === 'thieu') add('thieu', 'vang', 'Chưa đủ dữ liệu đánh giá bản chất hợp đồng – tạm dừng phê duyệt chế độ thuế / BHXH.', ['BLLD_13']);
    if (cd === 'kinhdoanh') add('kd', 'vang', 'Hoạt động kinh doanh: không khấu trừ 10% tiền công, không áp giảm trừ bản thân; kiểm tra hóa đơn / nghĩa vụ thuế hộ, cá nhân kinh doanh.', ['ND68']);
    if (cd === 'khongcutru') add('kct', 'vang', 'Cá nhân không cư trú – nhánh thuế 20%, cần Phụ trách kế toán xác nhận.', ['TNCN_KCT']);
    if (x.loai !== 'tamung' && cd !== 'kinhdoanh' && cd !== 'khongcutru' && ky) {
      if (tn >= ng) add('kt10', ck.ok ? 'xanh' : 'xanh', 'Thu nhập thanh toán ' + Math.round(tn / 1e6 * 10) / 10 + ' triệu đồng. Thuộc diện khấu trừ 10% nếu không có cam kết hợp lệ.', ['TNCN_KT10']);
      if (ck.ok) add('ckOK', 'xanh', 'Đã có cam kết hợp lệ. Tạm thời không khấu trừ thuế đối với khoản thanh toán thuộc phạm vi áp dụng.', ['TNCN_CAMKET']);
      else if (ck.ma !== 'NONE') add('ckCho', 'vang', (ck.ma === 'EXPIRED' || ck.ma === 'REVOKED' ? CK_TT[ck.ma] + '. ' : 'Cam kết chưa được phê duyệt. ') + 'Tiếp tục khấu trừ theo quy định. (' + ck.lyDo + ')', ['TNCN_CAMKET']);
      if (tn >= ng && !num(x.thue) && !ck.ok && x.trangThai) add('chuaXL', 'do', 'Khoản thanh toán vượt ngưỡng khấu trừ nhưng chưa được xử lý thuế.', ['TNCN_KT10']);
    }
    var thang = (ctx.ttk || []).filter(function (y) { return y.maNV === x.maNV && y.trangThai !== 'Hủy' && y.loai !== 'tamung' && String(y.ngayTra || '').slice(0, 7) === ky; });
    if (ky && thang.length > 1 && thang.every(function (y) { return num(y.tnChiuThue) < ng; }) && thang.reduce(function (t, y) { return t + num(y.tnChiuThue); }, 0) >= ng)
      add('chiaNho', 'do', 'Có ' + thang.length + ' khoản trả cùng tháng đều dưới ngưỡng nhưng tổng ≥ ' + ng.toLocaleString('vi-VN') + ' đ – kiểm tra dấu hiệu chia nhỏ khoản thanh toán để tránh khấu trừ.', ['TNCN_KT10']);
    if (!mstHopLe(nv.mst)) add('mst', 'vang', 'Chưa có mã số thuế – vẫn khấu trừ theo quy định, cam kết không có hiệu lực.', ['ND252']);
    if (hd.loaiK === 'K2') {
      var dg = parseJSON(hd.danhGia, {}) || {};
      if (dg.nguoiNhan === 'tungtv' && dg.traTrucTiep !== 'tungtv') add('k2a', 'do', 'Tổ đội: thực tế từng thành viên nhận thu nhập nhưng TAPACO trả cho đội trưởng – cần hồ sơ ủy quyền nhận tiền, danh sách, chứng từ chi trả từng người.', ['TNCN_KT10']);
      if (dg.traTrucTiep === 'doitruong' && dg.nguoiNhan !== 'tungtv') add('k2b', 'vang', 'Đội trưởng nhận trọn tiền khoán: thu nhập tính cho người nhận (đội trưởng); không tự chia thu nhập cho thành viên khi chưa có căn cứ pháp lý và chứng từ hợp lệ.', ['TNCN_KT10']);
      if (dg.coHoSo !== 'co') add('k2c', 'vang', 'Tổ đội chưa đủ hợp đồng, giấy ủy quyền, danh sách thành viên, chứng từ thanh toán.', []);
    }
    if (ctx.bls || ctx.ttk) {
      var q = quyenLoiThueKhoan(nv, Object.assign({}, ctx, { nam: String(x.ngayTra || ctx.homNay).slice(0, 4) }));
      if (q.uoc.D > 0 && q.uoc.D <= q.camKet.giamTru.tong && cd === 'tiencong10' && !ck.ok) add('duoiGT', 'vang', 'Tổng tiền công dự kiến cả năm ' + Math.round(q.uoc.D / 1e6) + ' triệu đồng, thấp hơn giảm trừ gia cảnh năm ' + q.nam + '. Cần xác nhận các nguồn thu nhập khác và điều kiện cam kết.', ['TNCN_CAMKET']);
      if (q.nopThua > 0) add('hoan', 'xanh', 'Thuế đã khấu trừ cao hơn nghĩa vụ thuế dự kiến cuối năm. Cần hướng dẫn cá nhân quyết toán / hoàn thuế.', ['TNCN_CT70']);
      if (q.uyQuyen.tuQT) add('uq', 'xanh', 'Người nhận khoán chỉ có hợp đồng khoán độc lập. Không đủ điều kiện ủy quyền TAPACO quyết toán theo cơ chế HĐLĐ.', ['TNCN_UQ']);
    }
    return out;
  }
  var HS_KHOAN = [['f01', '01_Hop_dong_khoan', 'Hợp đồng khoán'], ['f02', '02_Phu_luc_hop_dong', 'Phụ lục hợp đồng'], ['f03', '03_Ho_so_ca_nhan', 'Hồ sơ cá nhân'], ['f04', '04_Khoi_luong_nghiem_thu', 'Khối lượng nghiệm thu'], ['f05', '05_Chung_tu_thanh_toan', 'Chứng từ thanh toán'], ['f06', '06_Cam_ket_thue', 'Cam kết thuế'], ['f07', '07_Chung_tu_khau_tru', 'Chứng từ khấu trừ'], ['f08', '08_Quyet_toan_hoan_thue', 'Quyết toán / hoàn thuế'], ['f09', '09_Giai_trinh_kiem_tra', 'Giải trình kiểm tra']];
  var HS_TT = ['Chưa cung cấp', 'Đã tải lên', 'Chờ kiểm tra', 'Cần bổ sung', 'Hợp lệ', 'Không hợp lệ', 'Hết hiệu lực'];
  var QT_KHOAN_MD = [
    { id: 'BCH', ten: 'Ban chỉ huy công trình: hợp đồng, phạm vi, xác nhận khối lượng nghiệm thu', nguoi: ['r:cht', 'r:truongphong'] },
    { id: 'KTKH', ten: 'Phòng Kỹ thuật – Kế hoạch: đơn giá, khối lượng, phát sinh, khấu trừ theo HĐ', nguoi: ['r:truongphong', 'r:tchc'] },
    { id: 'KT', ten: 'Phòng Kế toán: người nhận, loại HĐ, hồ sơ thuế, thuế khấu trừ, chứng từ', nguoi: ['r:ktluong', 'r:ktcongno'] },
    { id: 'PTKT', ten: 'Phụ trách kế toán: phân loại thuế, cam kết, rủi ro BHXH, hồ sơ hợp lệ', nguoi: ['r:kttruong'] },
    { id: 'GD', ten: 'Giám đốc duyệt giá trị thanh toán', nguoi: ['r:giamdoc'] },
    { id: 'TT', ten: 'Kế toán thanh toán: chi trả, hạch toán, lưu chứng từ', nguoi: ['r:ktcongno', 'r:ktluong'] }
  ];

  /* ===================== v3.29: NHÂN SỰ HƯỞNG LƯƠNG HƯU (RETIRED EMPLOYEE PAYROLL & PIT) =====================
     Xác minh tình trạng hưởng lương hưu theo thực tế (không theo tuổi), nghĩa vụ BH theo từng kỳ, khoản trả thêm tương đương BH
     (Điều 168.3 BLLĐ – phương án A chưa gồm / B đã gồm trong tổng chi trả), thuế TNCN lũy tiến chi tiết từng bậc, ủy quyền, tối ưu thuế. */
  var HUU_TT = { dangHuong: 'Đang hưởng lương hưu', duTuoi: 'Đã đủ tuổi nhưng chưa hưởng lương hưu', troCap: 'Hưởng trợ cấp khác / đã nhận BHXH một lần', chuaXacMinh: 'Chưa xác minh' };
  var HUU_XM = ['Chờ', 'Đã xác minh', 'Không hợp lệ'];
  var HUU_PA = { A: 'A – Lương thỏa thuận CHƯA gồm khoản trả thêm BH (cộng thêm)', B: 'B – Tổng chi trả thỏa thuận ĐÃ gồm khoản trả thêm BH (tách ra)' };
  var HS_HUU = [['h01', '01_HDLD_nguoi_huong_luong_huu', 'HĐLĐ với người đang hưởng lương hưu'], ['h02', '02_QD_huong_luong_huu', 'Quyết định hưởng lương hưu / tài liệu chứng minh'], ['h03', '03_Xac_nhan_BHYT', 'Xác nhận tình trạng BHYT'],
    ['h04', '04_Phu_luc_tra_them_BH', 'Phụ lục thỏa thuận tiền trả thêm tương đương BH'], ['h05', '05_Bang_tinh_luong_thue', 'Bảng tính lương và thuế hằng tháng'], ['h06', '06_Ho_so_NPT', 'Hồ sơ người phụ thuộc'],
    ['h07', '07_Ke_khai_thu_nhap_noi_khac', 'Bản kê khai thu nhập chịu thuế tại nơi khác'], ['h08', '08_Uy_quyen_quyet_toan', 'Giấy ủy quyền quyết toán (nếu đủ điều kiện)'], ['h09', '09_Cam_ket_khong_khau_tru', 'Cam kết không khấu trừ (nếu được phép)'], ['h10', '10_Chung_tu_quyet_toan_hoan_thue', 'Chứng từ khấu trừ, kết quả quyết toán, hoàn thuế']];
  /* tình trạng hưu trí áp dụng cho 1 kỳ lương: mien = không thuộc diện BHXH, BHYT, BHTN bắt buộc tại TAPACO */
  function huuTaiKy(nv, ky, cfg) {
    if (!nv || truongHopNV(nv) !== 'huu') return null;
    var h = hsThueNV(nv).huu, tt = h.tinhTrang || 'chuaXacMinh', xm = h.trangThaiXM === 'Đã xác minh';
    var o = { tinhTrang: tt, tenTT: HUU_TT[tt], xacMinh: xm, trangThaiXM: h.trangThaiXM || 'Chờ', phuongAn: h.phuongAnBH || (cfg && cfg.traThayBH === 'co' ? 'A' : ''), canCu: h.canCuBH, h: h, mien: false, tamTinh: false, truocHuu: false };
    if (tt === 'duTuoi') { o.lyDo = 'Đã đủ tuổi nhưng chưa thực sự hưởng lương hưu – không miễn BHXH; tính BH theo HĐLĐ (không dùng tuổi làm căn cứ miễn đóng).'; return o; }
    if (tt === 'troCap') { o.lyDo = 'Hưởng trợ cấp khác / đã nhận BHXH một lần, không hưởng lương hưu – đánh giá nghĩa vụ BHXH theo HĐLĐ thực tế; tạm tính như HĐLĐ đóng BH, Kế toán trưởng xác nhận.'; return o; }
    if (h.trangThaiXM === 'Không hợp lệ') { o.lyDo = 'Hồ sơ hưởng lương hưu không hợp lệ – tính như HĐLĐ đóng BH.'; return o; }
    if (tt === 'dangHuong' && h.ngayHuong && ky && ky < h.ngayHuong.slice(0, 7)) { o.truocHuu = true; o.lyDo = 'Kỳ ' + ky.slice(5) + '/' + ky.slice(0, 4) + ' trước tháng bắt đầu hưởng lương hưu (' + h.ngayHuong.slice(5, 7) + '/' + h.ngayHuong.slice(0, 4) + ') – tính như HĐLĐ đóng BH, không có khoản trả thêm.'; return o; }
    o.mien = true; o.tamTinh = !xm;
    o.lyDo = xm ? 'Đã xác minh đang hưởng lương hưu – không trích BHXH bắt buộc.' : 'Chưa đủ căn cứ xác định nghĩa vụ BHXH – tạm không trích BH theo khai báo; Kế toán xác minh trước khi khóa bảng lương.';
    return o;
  }
  /* khoản trả thêm tương đương mức DN đóng BHXH, BHYT, BHTN (Điều 168.3 BLLĐ) – căn cứ, tỷ lệ, trần theo thời kỳ */
  function traThemBH(o, cfg, pa, canCu) {
    var dn = cfg.tyLeDN || {}, tranXH = num(cfg.luongCoSo) * num(cfg.heSoTranBHXH) || Infinity, tranTN = num(cfg.luongToiThieuVung) * num(cfg.heSoTranBHTN) || Infinity;
    var f = function (b) { return { bhxh: R(Math.min(b, tranXH) * num(dn.bhxh) / 100), bhyt: R(Math.min(b, tranXH) * num(dn.bhyt) / 100), bhtn: R(Math.min(b, tranTN) * num(dn.bhtn) / 100) }; };
    var L = num(o.luong), base = null, nguon = '', S = L, t, i, rT = (num(dn.bhxh) + num(dn.bhyt) + num(dn.bhtn)) / 100;
    if (num(canCu) > 0) { base = num(canCu); nguon = 'Căn cứ tham chiếu khai trong hồ sơ hưu trí'; }
    else if (num(o.luongBH) > 0) { base = num(o.luongBH); nguon = 'Lương cơ bản (mức đóng BH tham chiếu) trong hồ sơ'; }
    if (pa === 'B') {
      if (base === null) { S = L / (1 + rT); for (i = 0; i < 40; i++) { var q = f(S); S = L - q.bhxh - q.bhyt - q.bhtn; } base = S; nguon = 'Tiền lương theo công việc, tách từ tổng chi trả thỏa thuận'; }
      t = f(base); S = L - t.bhxh - t.bhyt - t.bhtn;
    } else { if (base === null) { base = L; nguon = 'Lương thỏa thuận theo HĐLĐ'; } t = f(base); }
    var m = cfg._ky ? mocCuaKy(cfg, cfg._ky) : null;
    return { phuongAn: pa, canCu: R(base), nguon: nguon, tyLe: { bhxh: num(dn.bhxh), bhyt: num(dn.bhyt), bhtn: num(dn.bhtn), tong: Math.round(rT * 1000) / 10 }, tran: { bhxh: tranXH, bhtn: tranTN },
      bhxh: t.bhxh, bhyt: t.bhyt, bhtn: t.bhtn, tong: t.bhxh + t.bhyt + t.bhtn, luongCV: R(pa === 'B' ? S : L), tongChiTra: pa === 'B' ? L : L + t.bhxh + t.bhyt + t.bhtn,
      duoiLTT: (pa === 'B' ? S : L) < num(cfg.luongToiThieuVung), hieuLuc: m ? m.tu : '', canCuPL: 'BLLD_168',
      congThuc: 'Tiền trả thêm = ' + Math.round(base).toLocaleString('vi-VN') + ' (' + nguon + ') × (' + num(dn.bhxh) + '% + ' + num(dn.bhyt) + '% + ' + num(dn.bhtn) + '%) = ' + (t.bhxh + t.bhyt + t.bhtn).toLocaleString('vi-VN') + (base > tranXH ? ' (BHXH, BHYT tính trên trần ' + tranXH.toLocaleString('vi-VN') + ')' : '') + (pa === 'B' ? '; phương án B: tiền lương theo công việc = ' + L.toLocaleString('vi-VN') + ' − ' + (t.bhxh + t.bhyt + t.bhtn).toLocaleString('vi-VN') + ' = ' + R(S).toLocaleString('vi-VN') : '') };
  }
  /* chi tiết thuế lũy tiến từng bậc */
  function chiTietLuyTien(tntt, bieu) {
    var out = [], truoc = 0; (bieu || []).forEach(function (b, i) {
      var tran = num(b[0]) || Infinity, ts = num(b[1]); if (tntt <= truoc) { truoc = tran; return; }
      var phan = Math.min(tntt, tran) - truoc; out.push({ bac: i + 1, tu: truoc, den: tran === Infinity ? null : tran, thueSuat: ts, phan: phan, thue: R(phan * ts / 100) }); truoc = tran;
    });
    return out;
  }
  /* calculateRetireePIT(employee, payrollPeriod, legalRules) – ctx: { cfg, cc, o } */
  function calculateRetireePIT(nv, ky, ctx) {
    ctx = ctx || {}; var cfg = ctx.cfg, ck = cfgTheoKy(cfg, ky), nlv = (ck.noiLamViec || []).filter(function (x) { return x.ma === nv.noiLamViec; })[0], ncc = num(nlv && nlv.congChuan) || num(ck.ngayCongChuan) || 26;   // đủ công theo nơi làm việc
    var o = Object.assign({ luong: num(nv.luongChinh), luongBH: num(nv.luongBH), ncc: ncc, ky: ky }, ctx.o || {}), p = tinhPhieuLuong(nv, ctx.cc || { ngayCong: ncc }, cfg, o);
    var tt = ttThue(nv, ky), bt = num(p.giamTru) ? Math.min(num(p.giamTru), num(ck.giamTruBanThan)) : 0, bac = tt.cach === 'luytien' ? chiTietLuyTien(num(p.tnTinhThue), ck.bieuThue) : [];
    var hk = huuTaiKy(nv, ky, cfg), tb = parseJSON(p.traThayBHCT, null);
    return { phieu: p, ky: ky, cach: tt.cach, moTa: tt.moTa, tinhTrangHuu: hk, tnChiuThue: num(p.tnChiuThue), traThem: tb, giamTruBanThan: bt, giamTruNPT: num(p.giamTru) - bt, bhNLD: num(p.bhxh) + num(p.bhyt) + num(p.bhtn),
      tntt: num(p.tnTinhThue), bac: bac, thue: num(p.thueTNCN), thucNhan: num(p.thucLinh), luongHuu: hk ? hk.h.luongHuu : 0, tongHaiNguon: num(p.thucLinh) + (hk && hk.mien ? hk.h.luongHuu : 0),
      chiPhiTrucTiep: num(p.tongThuNhap) + num(p.bhxhDN) + num(p.bhytDN) + num(p.bhtnDN) + num(p.kpcdDN),
      congThuc: 'TN chịu thuế ' + num(p.tnChiuThue).toLocaleString('vi-VN') + ' − BH NLĐ ' + (num(p.bhxh) + num(p.bhyt) + num(p.bhtn)).toLocaleString('vi-VN') + ' − giảm trừ ' + num(p.giamTru).toLocaleString('vi-VN') + ' = ' + num(p.tnTinhThue).toLocaleString('vi-VN') + (bac.length ? '; ' + bac.map(function (b) { return 'bậc ' + b.bac + ': ' + b.phan.toLocaleString('vi-VN') + ' × ' + b.thueSuat + '% = ' + b.thue.toLocaleString('vi-VN'); }).join('; ') : '') + ' → thuế ' + num(p.thueTNCN).toLocaleString('vi-VN') };
  }
  /* ỦY QUYỀN QUYẾT TOÁN cho người hưởng lương hưu: lương hưu miễn thuế không phải thu nhập tiền lương nơi khác */
  function evaluateRetireeTaxAuthorization(nv, ctx) {
    var dg = evaluateTaxAuthorizationEligibility(nv, ctx), hs = hsThueNV(nv), ds = dg.ds.slice(), hr = dg.hoSo;
    ds.push({ ma: 'luongHuu', ten: 'Lương hưu do Quỹ BHXH chi trả (miễn thuế) không tính là thu nhập tiền lương nơi khác', ok: true, ghi: hs.huu.luongHuu ? 'Lương hưu ' + hs.huu.luongHuu.toLocaleString('vi-VN') + ' đ/tháng – không ảnh hưởng quyền ủy quyền' : 'Không ảnh hưởng quyền ủy quyền' });
    var b = hs.tnKhac.tuQTBatBuoc;
    ds.push({ ma: 'tuQT51', ten: 'Không thuộc trường hợp bắt buộc tự quyết toán (khoản 3 Điều 51 NĐ 253/2026 – Kế toán trưởng xác định)', ok: b === 'co' ? false : b === 'khong' ? true : null, ghi: b === 'co' ? 'Thuộc trường hợp phải tự quyết toán' : b === 'khong' ? 'Không thuộc' : 'Chưa xác định – nội dung khoản 3 Điều 51 đang xác minh' });
    var kq = ketLuan(ds, false);
    return { ketQua: kq, ten: kq === 'du' ? 'Được ủy quyền TAPACO quyết toán' : kq === 'khongdu' ? 'Không được ủy quyền – tự quyết toán nếu phát sinh nghĩa vụ' : 'Cần bổ sung hồ sơ', ds: ds, hoSo: hr, canKy: kq === 'du' && !(hr && hr.trangThai === 'Đã duyệt'), hoSoHopLe: !!(hr && hr.trangThai === 'Đã duyệt') };
  }
  /* PHÂN TÍCH TỐI ƯU THUẾ – người hưởng lương hưu */
  function phanTichHuuTri(nv, ctx) {
    var nam = String(ctx.nam), to = toiUuThue(nv, ctx), hs = hsThueNV(nv), h = hs.huu, ky = String(ctx.homNay || '').slice(0, 4) === nam ? String(ctx.homNay).slice(0, 7) : nam + '-12';
    var tt = ttThue(nv, ky), hk = huuTaiKy(nv, ky, ctx.cfg), uq = evaluateRetireeTaxAuthorization(nv, ctx), ds = [], add = function (ma, ten, ok, ghi) { ds.push({ ma: ma, ten: ten, ok: ok, ghi: ghi }); };
    add('gtBT', 'Đã áp dụng đủ giảm trừ bản thân', tt.cach !== 'luytien' ? 'na' : tt.giamTru === 'co' ? true : hs.noiKhac.gtTai === 'noikhac' ? true : false, tt.cach !== 'luytien' ? 'Khấu trừ 10% (HĐ dưới 3 tháng) – giảm trừ tính khi quyết toán' : tt.giamTru === 'co' ? 'Đang giảm trừ ' + num(cfgTheoKy(ctx.cfg, ky).giamTruBanThan).toLocaleString('vi-VN') + ' đ/tháng tại TAPACO' : hs.noiKhac.gtTai === 'noikhac' ? 'Đã đăng ký giảm trừ tại nơi khác – không giảm trừ trùng' : 'Chưa giảm trừ bản thân – người nghỉ hưu vẫn được giảm trừ; sửa trạng thái thuế');
    var nptThieu = hs.nptDs.filter(function (d) { return d.coHoSo !== 'co'; });
    add('npt', 'Người phụ thuộc hợp lệ đã đăng ký', nptThieu.length ? null : true, nptThieu.length ? nptThieu.length + ' người phụ thuộc chưa đủ hồ sơ: ' + nptThieu.map(function (d) { return d.ten; }).join(', ') : (num(nv.soNPT) || hs.nptDs.length) ? 'Đã đăng ký' : 'Chưa có – rà soát cha mẹ, vợ/chồng, con đủ điều kiện để giảm thuế hợp pháp');
    add('luongHuu', 'Không tính nhầm lương hưu (miễn thuế) vào thu nhập TAPACO', true, 'Lương hưu do Quỹ BHXH chi trả, không nằm trong bảng lương TAPACO');
    var bls = (ctx.bls || []).filter(function (b) { return b.maNV === nv.maNV && String(b.ky).slice(0, 4) === nam; }), lech = [];
    bls.forEach(function (b) { var c = parseJSON(b.traThayBHCT, null), hk2 = huuTaiKy(nv, b.ky, ctx.cfg); if (hk2 && hk2.mien && hk2.phuongAn && !c) lech.push(b.ky); });
    add('traThem', 'Khoản trả thêm tương đương BH tính đúng', !hk || !hk.mien ? 'na' : !hk.phuongAn ? null : lech.length ? false : true, !hk || !hk.mien ? 'Không áp dụng' : !hk.phuongAn ? 'Chưa rõ lương HĐ đã gồm tiền thay BH hay chưa – chọn phương án A / B' : lech.length ? 'Kỳ ' + lech.join(', ') + ' chưa có khoản trả thêm – tính lại lương' : 'Phương án ' + hk.phuongAn + '; đã cộng vào thu nhập chịu thuế');
    var A = to.C.tongKhauTru, B = to.thueNam, thua = Math.max(0, A - B), them = Math.max(0, B - A);
    var huong = [];
    if (hk && hk.tamTinh) huong.push('Kiểm tra hồ sơ');
    if (ds.some(function (d) { return d.ma === 'npt' && d.ok !== true; }) || ds.some(function (d) { return d.ma === 'gtBT' && d.ok === false; })) huong.push('Bổ sung giảm trừ');
    huong.push(uq.ketQua === 'du' ? 'Ủy quyền' : 'Tự quyết toán');
    return { nam: nam, ds: ds, daKTDenNay: to.uoc.daKT, A: A, B: B, nopThua: thua, nopThem: them, giamHopPhap: to.thueGiamTheoLuat, hoanDoTamKT: thua, huongXuLy: huong, uyQuyen: uq, toiUu: to, tinhTrang: hk,
      luongHuuNam: hk && hk.mien ? h.luongHuu * 12 : 0, ghiChu: 'A = thuế đã và sẽ khấu trừ cả năm (đã khấu trừ đến nay ' + Math.round(to.uoc.daKT).toLocaleString('vi-VN') + ' đ); B = thuế dự kiến phải nộp năm. Nộp thừa = max(0, A − B); nộp thêm = max(0, B − A). Số thuế được giảm hợp pháp (người phụ thuộc, y tế, giáo dục, hưu trí tự nguyện) tách riêng với số chỉ được hoàn do tạm khấu trừ thừa.' };
  }
  /* GHI CHÚ PHÁP LÝ – NHÂN SỰ HƯỞNG LƯƠNG HƯU (cột riêng trên bảng lương) */
  function ghiChuHuuTri(nv, bl, ctx) {
    var ky = (bl && bl.ky) || String(ctx.homNay || '').slice(0, 7), hk = huuTaiKy(nv, ky, ctx.cfg); if (!hk) return [];
    var out = [], add = function (ma, muc, t, chiTiet, canCu, hoSo, mau) { out.push({ ma: ma, muc: muc, t: t, chiTiet: chiTiet || '', canCu: canCu || [], hoSo: hoSo || '', mau: mau || '' }); };
    var h = hk.h, hs = hsThueNV(nv), tt = ttThue(nv, ky), tb = parseJSON(bl && bl.traThayBHCT, null), ck = cfgTheoKy(ctx.cfg, ky);
    if (hk.mien && hk.xacMinh) add('hXM', 'xanh', 'Đã xác minh đang hưởng lương hưu – không trích BHXH bắt buộc.', 'QĐ ' + (h.soQD || '…') + ' · ' + (h.donViChiTra || '') + ' · xác minh ' + (h.ngayXM || '') + ' bởi ' + (h.nguoiXM || ''), ['BHXH_DT']);
    else if (hk.tamTinh) add(h.soQD ? 'hCho' : 'hQD', 'do', h.soQD ? 'Chưa đủ căn cứ xác định nghĩa vụ BHXH – kế toán xác minh trước khi khóa bảng lương.' : 'Chưa có quyết định hưởng lương hưu – cần bổ sung.', hk.lyDo, ['BHXH_DT'], 'Quyết định hưởng lương hưu / tài liệu chứng minh (hồ sơ 02)');
    else add('hBH', hk.truocHuu ? 'xanh' : 'vang', hk.lyDo, '', ['BHXH_DT']);
    if (hk.mien && !hk.phuongAn) add('hPA', 'vang', 'Chưa rõ lương hợp đồng đã bao gồm tiền thay bảo hiểm hay chưa.', 'Chọn phương án A (chưa gồm – cộng thêm) hoặc B (đã gồm – tách ra) trong Hồ sơ thuế – BH; chưa chọn thì chưa tính khoản trả thêm.', ['BLLD_168'], 'Phụ lục thỏa thuận tiền trả thêm tương đương BH (hồ sơ 04)', 'bmh01');
    if (tb && tb.tong > 0) {
      add('hTT', 'xanh', 'Có khoản trả thêm tương đương bảo hiểm theo Điều 168 BLLĐ.', tb.congThuc, ['BLLD_168']);
      add('hTTthue', 'xanh', 'Khoản trả thêm bảo hiểm đã được cộng vào thu nhập chịu thuế.', 'Không trừ khoản tiền thay bảo hiểm khỏi thu nhập chịu thuế chỉ vì mang tên bảo hiểm.', ['BLLD_168']);
      if (tb.duoiLTT) add('hLTT', 'do', 'Tiền lương theo công việc sau khi tách khoản trả thêm thấp hơn lương tối thiểu vùng.', 'Lương theo công việc ' + tb.luongCV.toLocaleString('vi-VN') + ' < ' + num(ck.luongToiThieuVung).toLocaleString('vi-VN'), ['BLLD_168']);
    }
    if (hk.mien) add('hHuuMien', 'xanh', 'Lương hưu từ Quỹ BHXH được miễn thuế TNCN.', h.luongHuu ? 'Lương hưu ' + h.luongHuu.toLocaleString('vi-VN') + ' đ/tháng – không cộng vào thu nhập TAPACO.' : '', ['TNCN_LUONGHUU']);
    if (tt.cach === 'luytien' && tt.giamTru === 'co') add('hGT', 'xanh', 'Được áp dụng giảm trừ bản thân ' + (num(ck.giamTruBanThan) / 1e6).toLocaleString('vi-VN') + ' triệu/tháng.', 'Người nghỉ hưu, không đóng BH tại TAPACO vẫn được giảm trừ bản thân.', ['TNCN_GT']);
    if (hs.noiKhac.gtTai === 'noikhac') add('hGTNK', 'xanh', 'Đã có giảm trừ tại đơn vị khác – không giảm trừ trùng.', '', ['TNCN_GT']);
    if (tt.cach === 'khautru10') add('hNgan', 'vang', 'Hợp đồng dưới 3 tháng – kiểm tra khấu trừ 10% và điều kiện cam kết.', 'Không tự động đủ điều kiện cam kết chỉ vì đang hưởng lương hưu.', ['TNCN_KT10', 'TNCN_CAMKET'], '', 'camket');
    else {
      var uq = evaluateRetireeTaxAuthorization(nv, Object.assign({}, ctx, { nam: ky.slice(0, 4) }));
      if (uq.ketQua === 'du' && uq.canKy) add('hUQ', 'vang', 'Đủ điều kiện ủy quyền quyết toán – cần ký hồ sơ.', '', ['TNCN_UQ', 'TT89'], 'Mẫu 08/UQ-QTT-TNCN đã ký', 'uyquyen');
      else if (uq.ketQua === 'khongdu') add('hUQk', 'vang', 'Không đủ điều kiện ủy quyền – cần tự quyết toán nếu có nghĩa vụ.', uq.ds.filter(function (d) { return d.ok === false; }).map(function (d) { return d.ten + ': ' + d.ghi; }).join('; '), ['TNCN_UQ']);
    }
    if (ctx.bls) { var to = toiUuThue(nv, Object.assign({}, ctx, { nam: ky.slice(0, 4) })); if (to.thuaHoan > 1000) add('hHoan', 'xanh', 'Có khả năng được hoàn thuế đã khấu trừ thừa.', 'Dự kiến nộp thừa ' + Math.round(to.thuaHoan).toLocaleString('vi-VN') + ' đ.', ['TNCN_CT70']); }
    if (h.congViecNguyCo === 'co' || nv.noiLamViec === 'DA') add('hAT', 'vang', 'Người lao động cao tuổi làm công việc có nguy cơ cao – cần kiểm tra điều kiện an toàn.', 'Làm việc tại công trường / công việc nặng nhọc, độc hại, nguy hiểm: kiểm tra điều kiện bảo đảm an toàn, sức khỏe.', ['BLLD_149']);
    return out;
  }
  /* các NV trường hợp hưu trí của kỳ chưa đủ căn cứ BHXH – chặn khóa bảng lương */
  function huuChuaXacMinh(ky, bls, nvs, cfg) {
    var m = {}; (nvs || []).forEach(function (n) { m[n.maNV] = n; });
    return (bls || []).filter(function (b) { return b.ky === ky && m[b.maNV]; }).map(function (b) { var hk = huuTaiKy(m[b.maNV], ky, cfg); return hk && hk.tamTinh ? m[b.maNV] : null; }).filter(Boolean);
  }

  /* ===== v3.25: QUY ĐỊNH THEO THỜI KỲ (trước / sau 01/7/2026) ===== */
  var BIEU_THUE = {
    b7: { ten: '7 bậc (Luật Thuế TNCN 2007)', bac: [[5000000, 5], [10000000, 10], [18000000, 15], [32000000, 20], [52000000, 25], [80000000, 30], [0, 35]] },
    b5: { ten: '5 bậc (Luật Thuế TNCN 109/2025)', bac: [[10000000, 5], [30000000, 10], [60000000, 20], [100000000, 30], [0, 35]] }
  };
  var OT_CACH = { chenh: 'Chỉ phần trả cao hơn giờ làm việc bình thường được miễn (TT 111/2013)', toanbo: 'Toàn bộ tiền làm thêm giờ trong giới hạn giờ của luật lao động được miễn (Luật 109/2025, NĐ 253/2026 Điều 26)' };
  var MOC_LUAT_MD = [
    { tu: '2024-07', ten: 'Luật Thuế TNCN 2007 · lương cơ sở 2.340.000 (NĐ 73/2024) · lương tối thiểu vùng I 4.960.000 (NĐ 74/2024) · giảm trừ 11/4,4 triệu (NQ 954/2020)',
      luongCoSo: 2340000, luongToiThieuVung: 4960000, giamTruBanThan: 11000000, giamTruNPT: 4400000, bieu: 'b7', otCach: 'chenh', nguongVangLai: 2000000, anCaMien: 730000, nldXH: 8, nldYT: 1.5, nldTN: 1, dnXH: 17.5, dnYT: 3, dnTN: 1, kpcd: 2, tnld: 0.5, thueVL: 10 },
    { tu: '2026-01', ten: 'Từ 01/01/2026: giảm trừ 15,5/6,2 triệu (NQ 110/2025) · lương tối thiểu vùng I 5.310.000 (NĐ 293/2025) · thuế vẫn theo Luật 2007',
      luongCoSo: 2340000, luongToiThieuVung: 5310000, giamTruBanThan: 15500000, giamTruNPT: 6200000, bieu: 'b7', otCach: 'chenh', nguongVangLai: 2000000, anCaMien: 730000, nldXH: 8, nldYT: 1.5, nldTN: 1, dnXH: 17.5, dnYT: 3, dnTN: 1, kpcd: 2, tnld: 0.5, thueVL: 10 },
    { tu: '2026-07', ten: 'Từ 01/7/2026: Luật Thuế TNCN 109/2025, NĐ 253/2026 · lương cơ sở 2.530.000 (NĐ 161/2026)',
      luongCoSo: 2530000, luongToiThieuVung: 5310000, giamTruBanThan: 15500000, giamTruNPT: 6200000, bieu: 'b5', otCach: 'toanbo', nguongVangLai: 5000000, anCaMien: 1200000, nldXH: 8, nldYT: 1.5, nldTN: 1, dnXH: 17.5, dnYT: 3, dnTN: 1, kpcd: 2, tnld: 0.5, thueVL: 10 }
  ];
  var MOC_BH = ['nldXH', 'nldYT', 'nldTN', 'dnXH', 'dnYT', 'dnTN', 'kpcd', 'tnld', 'thueVL'];   // v3.33: tỷ lệ BH (Luật BHXH 2024 Đ33, 34; Luật BHYT; Luật Việc làm), tỷ lệ khấu trừ cố định – theo thời kỳ
  function chuanMoc(ds) {
    var out = [], da = {};
    (Array.isArray(ds) ? ds : []).forEach(function (m) {
      if (!m || !/^\d{4}-(0[1-9]|1[0-2])$/.test(String(m.tu || '')) || da[m.tu]) return; da[m.tu] = 1;
      out.push({ tu: m.tu, ten: String(m.ten || '').slice(0, 300), luongCoSo: num(m.luongCoSo), luongToiThieuVung: num(m.luongToiThieuVung), giamTruBanThan: num(m.giamTruBanThan), giamTruNPT: num(m.giamTruNPT),
        bieu: BIEU_THUE[m.bieu] ? m.bieu : 'b5', otCach: OT_CACH[m.otCach] ? m.otCach : 'toanbo', nguongVangLai: num(m.nguongVangLai), anCaMien: num(m.anCaMien) });
      MOC_BH.forEach(function (k) { if (m[k] !== undefined && m[k] !== null && m[k] !== '') out[out.length - 1][k] = num(m[k]); });
    });
    out.sort(function (a, b) { return a.tu < b.tu ? -1 : 1; });
    return out.length ? out : JSON.parse(JSON.stringify(MOC_LUAT_MD));
  }
  // tham số áp dụng cho 1 kỳ lương: cfg + mốc gần nhất ≤ kỳ; trả thêm cheDoLuat (ghi lên phiếu lương), otCach
  function mocCuaKy(cfg, ky) {
    var ds = chuanMoc(cfg && cfg.mocLuat), m = null;
    ds.forEach(function (x) { if (x.tu <= ky) m = x; });
    return m;
  }
  function cfgTheoKy(cfg, ky) {
    if (!cfg || !ky || cfg._ky === ky) return cfg;
    var m = mocCuaKy(cfg, ky); if (!m) return cfg;
    var c = Object.assign({}, cfg), bieu = m.bieu, ot = m.otCach, ghi = '';
    if (cfg.luat109Tu2026 && ky >= '2026-01' && ky < '2026-07' && m.bieu === 'b7') {      // tùy chọn: áp quy định tiền lương của Luật 109 cho cả kỳ tính thuế 2026
      bieu = 'b5'; ot = 'toanbo'; ghi = ' · áp biểu 5 bậc & miễn thuế làm thêm giờ theo Luật 109/2025 từ kỳ 01/2026 (Điều 29.2)';
    }
    ['luongCoSo', 'luongToiThieuVung', 'giamTruBanThan', 'giamTruNPT', 'nguongVangLai'].forEach(function (k) { if (num(m[k]) > 0) c[k] = num(m[k]); });
    if (m.nldXH !== undefined) {                 // v3.33: tỷ lệ BH, tỷ lệ khấu trừ cố định của thời kỳ
      c.tyLeNLD = { bhxh: num(m.nldXH), bhyt: num(m.nldYT), bhtn: num(m.nldTN) };
      c.tyLeDN = Object.assign({}, cfg.tyLeDN || {}, { bhxh: num(m.dnXH), bhyt: num(m.dnYT), bhtn: num(m.dnTN), kpcd: num(m.kpcd), tnld: num(m.tnld) });
    }
    if (num(m.thueVL) > 0) c.thueVangLai = num(m.thueVL);
    c.anCaMienThue = num(m.anCaMien);            // v3.27: mức ăn giữa ca miễn thuế của thời kỳ (NĐ 253/2026: 1,2 triệu từ 01/7/2026)
    c.bieuThue = BIEU_THUE[bieu].bac; c.otCach = ot; c.tenBieu = BIEU_THUE[bieu].ten;
    c.cheDoLuat = m.ten + ghi;
    c._ky = ky;
    return c;
  }

  function mergeCfg(stored) {
    var c = clone(DEFAULT_CFG);
    Object.keys(stored || {}).forEach(function (k) { if (stored[k] !== undefined && stored[k] !== null) c[k] = stored[k]; });
    c.tyLeNLD = Object.assign(clone(DEFAULT_CFG.tyLeNLD), c.tyLeNLD || {});
    c.tyLeDN = Object.assign(clone(DEFAULT_CFG.tyLeDN), c.tyLeDN || {});
    c.heSoOT = Object.assign(clone(DEFAULT_CFG.heSoOT), c.heSoOT || {});
    if (!Array.isArray(c.congThucLuong) || !c.congThucLuong.length) {
      c.congThucLuong = clone(CT_LUONG_MD);
      DEFAULT_CFG.phuCap.forEach(function (p) { if (!(c.phuCap || []).some(function (x) { return x.ma === p.ma; })) c.phuCap = (c.phuCap || []).concat([clone(p)]); });
    }
    if (!Array.isArray(c.bienKy)) c.bienKy = clone(BIEN_KY_MD);
    if (vaOTQD(c.congThucLuong)) c._vaOT = 1;          // v3.31
    if (vaANCA(c.congThucLuong)) c._vaAn = 1;          // v3.32
    if (c.cheDoBangLuong !== 'tach') c.cheDoBangLuong = 'gop';
    // bản ≤ 2.2: phòng ban là danh sách tên → chuyển thành danh mục bộ phận
    if (!stored || !Array.isArray(stored.boPhan)) {
      if (stored && Array.isArray(stored.phongBan) && stored.phongBan.length) c.boPhan = stored.phongBan.map(function (t, i) {
        return { ma: 'BP' + pad(i + 1), ten: String(t), loai: /công trường|công trình|chỉ huy/i.test(t) ? 'Công trình' : 'Phòng ban', trangThai: 'Hoạt động' };
      });
    }
    c.boPhan = (c.boPhan || []).map(function (b) { return { ma: String(b.ma || ''), ten: String(b.ten || ''), loai: b.loai === 'Công trình' ? 'Công trình' : 'Phòng ban', diaDiem: b.diaDiem || '', batDau: b.batDau || '', ketThuc: b.ketThuc || '', trangThai: b.trangThai === 'Đã kết thúc' ? 'Đã kết thúc' : 'Hoạt động', ghiChu: b.ghiChu || '' }; }).filter(function (b) { return b.ten; });
    c.phongBan = c.boPhan.map(function (b) { return b.ten; });
    c.congTyTV = chuanCTV(c.congTyTV, c);
    c.mocLuat = chuanMoc(c.mocLuat);
    c.mocLuat.forEach(function (m) {              // v3.33: mốc cũ chưa ghi tỷ lệ BH → điền theo tỷ lệ chung đang dùng
      if (m.nldXH === undefined) Object.assign(m, { nldXH: num(c.tyLeNLD.bhxh), nldYT: num(c.tyLeNLD.bhyt), nldTN: num(c.tyLeNLD.bhtn), dnXH: num(c.tyLeDN.bhxh), dnYT: num(c.tyLeDN.bhyt), dnTN: num(c.tyLeDN.bhtn), kpcd: num(c.tyLeDN.kpcd), tnld: num(c.tyLeDN.tnld) });
      if (m.thueVL === undefined) m.thueVL = num(c.thueVangLai) || 10;
    });
    c.quyDinh = quyDinhDs(c); c.bieuMau = bieuMauDs(c);                 // v3.27
    if (Array.isArray(c.loaiTaiLieu)) ['Hồ sơ thuế TNCN', 'Hồ sơ khoán việc'].forEach(function (l) { if (c.loaiTaiLieu.indexOf(l) < 0) c.loaiTaiLieu = c.loaiTaiLieu.concat([l]); });
    if (!Array.isArray(c.quyTrinhKhoan) || c.quyTrinhKhoan.length !== 6) c.quyTrinhKhoan = clone(QT_KHOAN_MD);
    // ngày lễ: bản ≤ 2.3 chỉ có danh sách ngày
    c.ngayLeDs = cleanLe(stored && Array.isArray(stored.ngayLeDs) ? stored.ngayLeDs : (c.ngayLe || []));
    c.ngayLe = c.ngayLeDs.map(function (x) { return x.ngay; });
    if (!Array.isArray(c.loaiTamUng) || !c.loaiTamUng.length) c.loaiTamUng = LOAI_TU_MD.slice();
    c.danhMucViPham = (c.danhMucViPham || []).map(function (d) { var x = clone(d); if (NHOM_VP.indexOf(x.nhom) < 0) x.nhom = x.loai === 'Bồi thường thiệt hại' ? 'Mất mát, hư hỏng tài sản' : 'Vi phạm nội quy – an toàn'; return x; });
    // tên vai trò cũ trong quy trình
    ['quyTrinhLuong', 'quyTrinhNghi', 'quyTrinhYKien'].forEach(function (k) {
      if (Array.isArray(c[k])) c[k] = c[k].map(function (b) { var x = clone(b); x.nguoi = (x.nguoi || []).map(function (a) { return /^r:/.test(a) ? 'r:' + normRole(a.slice(2)) : a; }); return x; });
    });
    if (!c.congThucDongBH) c.congThucDongBH = CT_DONGBH_MD;
    return c;
  }

  /* ---------------- thuế TNCN lũy tiến ---------------- */
  function thueLuyTien(tntt, bieu) {
    var thue = 0, truoc = 0;
    for (var i = 0; i < bieu.length; i++) {
      var tran = num(bieu[i][0]) || Infinity, ts = num(bieu[i][1]) / 100;
      if (tntt <= truoc) break;
      thue += (Math.min(tntt, tran) - truoc) * ts;
      truoc = tran;
    }
    return R(thue);
  }

  /* ---- ngày nghỉ ---- */
  function laNgayNghiLe(d, cfg) {
    var p = parseD(d); if (!p) return true;
    if (new Date(Date.UTC(p.y, p.m - 1, p.d)).getUTCDay() === 0) return true;      // Chủ nhật
    return (cfg.ngayLe || []).indexOf(d) >= 0;
  }
  function ngayNghiTrong(don, tu, den, cfg) {       // số ngày công nghỉ của đơn rơi trong [tu, den]
    var a = don.tuNgay > tu ? don.tuNgay : tu, b = don.denNgay < den ? don.denNgay : den, n = 0, guard = 0;
    if (!a || !b || a > b) return 0;
    for (var d = a; d <= b && guard < 800; d = addDays(d, 1), guard++) if (!laNgayNghiLe(d, cfg)) n++;
    if (n && don.tuNgay === don.denNgay && don.buoi && don.buoi !== 'Cả ngày') n = 0.5;
    return n;
  }
  function soNgayDon(don, cfg) { return ngayNghiTrong(don, don.tuNgay, don.denNgay, cfg); }
  function ngayNghiTheoKy(don, ky, cfg) { return ngayNghiTrong(don, kyStart(ky), kyEnd(ky), cfg); }
  function quyPhepNam(nv, nam, cfg) {
    if (num(nv.ngayPhepNam) > 0) return num(nv.ngayPhepNam);
    var vao = parseD(nv.ngayVao); if (!vao || vao.y > nam) return 0;
    var namTN = Math.max(0, nam - vao.y);                                   // số năm làm việc tính đến cuối năm
    var quy = num(cfg.ngayPhepCoBan || 12) + Math.floor(namTN / (num(cfg.thamNienCongPhep) || 5));
    var thangDau = vao.y === nam ? vao.m : 1, thangCuoi = 12;
    var nghi = nv.trangThai === 'Đã nghỉ việc' && parseD(nv.ngayNghi);
    if (nghi && nghi.y === nam) thangCuoi = nghi.m;
    var soThang = Math.max(0, thangCuoi - thangDau + 1);
    if (vao.y === nam && vao.d > 15) soThang--;                           // vào sau ngày 15 không tính tháng đó
    return soThang >= 12 ? quy : Math.floor(quy * soThang / 12 * 2) / 2;
  }
  function phepNamDaDung(maNV, nam, rows, cfg, gomCho) {
    var r = { daDung: 0, cho: 0 };
    rows.forEach(function (x) {
      if (x.maNV !== maNV || x.loai !== 'Nghỉ phép năm') return;
      var n = ngayNghiTrong(x, nam + '-01-01', nam + '-12-31', cfg);
      if (x.trangThai === 'Đã duyệt') r.daDung += n;
      else if (gomCho !== false && (x.trangThai === 'Chờ duyệt' || x.trangThai === 'Chờ GĐ duyệt')) r.cho += n;
    });
    return r;
  }

  /* v3.13.2: giờ làm thực trong ngày (chấm công GPS) = giờ ra − giờ vào − phần trùng giờ nghỉ trưa; đủ công khi ≥ "Số giờ làm đủ 1 công" (mặc định 8) − số phút cho phép trễ */
  function phutCuaGio(g) { g = String(g || ''); return /^\d{1,2}:\d{2}/.test(g) ? +g.split(':')[0] * 60 + +g.split(':')[1].slice(0, 2) : NaN; }
  function gioLamThuc(vao, ra, g) {
    var a = phutCuaGio(vao), b = phutCuaGio(ra); if (!(b > a)) return 0;
    var t1 = phutCuaGio((g || {}).nghiTruaTu), t2 = phutCuaGio((g || {}).nghiTruaDen), tru = (t2 > t1) ? Math.max(0, Math.min(b, t2) - Math.max(a, t1)) : 0;
    return (b - a - tru) / 60;
  }
  function nguongDuCong(g) { g = g || {}; return Math.max(0, (num(g.gioLamDuCong) || 8) - (num(g.phutTre) || 0) / 60); }
  /* v3.17: nhóm hồ sơ đính kèm trong lý lịch – mỗi nhóm là một thư mục con trong thư mục CBCNV trên Google Drive Công ty */
  var TL_NHOM = [
    { ma: 'hd', ten: 'Hợp đồng, quyết định', thuMuc: '1. Hợp đồng – Quyết định', loai: ['Hợp đồng lao động', 'Phụ lục hợp đồng', 'Quyết định'] },
    { ma: 'll', ten: 'Lý lịch', thuMuc: '2. Lý lịch', loai: ['Sơ yếu lý lịch'] },
    { ma: 'cn', ten: 'Hồ sơ cá nhân', thuMuc: '3. Hồ sơ cá nhân', loai: [] }          // các loại còn lại (CCCD, MST, BHXH, bằng cấp, sức khỏe, người phụ thuộc, khác…)
  ];
  function tlNhom(loai) { for (var i = 0; i < 2; i++) if (TL_NHOM[i].loai.indexOf(loai) >= 0) return TL_NHOM[i]; return TL_NHOM[2]; }
  function tenThuMucNV(nv) { return nv.maNV + ' – ' + (nv.hoTen || ''); }
  function laVangLai(nv, ky) { return ttThue(nv, ky).cach === 'khautru10'; }   // v3.26: theo trạng thái tính thuế của kỳ
  /* v3.13: loại lương theo hợp đồng. Hồ sơ cũ chưa khai: thamGiaBH = 'Không' → không đóng BH tại Công ty, còn lại → trả lương + đóng BH. */
  var LOAI_LUONG = { bh: 'Trả lương + đóng BH tại Công ty', khongbh: 'Không đóng BH tại Công ty (đóng nơi khác / đã nghỉ hưu / HĐ thử việc)', khoan: 'Lương khoán việc (không đóng BH)' };
  var LY_DO_KHONG_BH = ['Đóng BH nơi khác', 'Đã nghỉ hưu', 'Hợp đồng thử việc', 'Khoán việc, thời vụ, CTV'];
  /* v3.26: 5 TRƯỜNG HỢP TÍNH LƯƠNG – bảo hiểm và cách tính thuế mặc định của từng trường hợp */
  var TRUONG_HOP = [
    { ma: 'bh', ten: '1. HĐLĐ – đóng BH tại Công ty', loaiLuong: 'bh', lyDo: '', moTa: 'NLĐ đóng BHXH 8%, BHYT 1,5%, BHTN 1%; DN đóng 17,5% + 3% + 1% + KPCĐ 2%. Thuế: biểu lũy tiến, giảm trừ gia cảnh tại Công ty, được ủy quyền Công ty quyết toán.' },
    { ma: 'khoan', ten: '2. HĐ khoán việc theo tháng – không đóng BH', loaiLuong: 'khoan', lyDo: '', moTa: 'Không đóng BH (HĐ khoán / dịch vụ, không phải HĐLĐ). Thuế: khấu trừ 10% khi trả từ ngưỡng (2 triệu/lần trước 01/7/2026, 5 triệu/lần từ 01/7/2026); có cam kết 08 thì không khấu trừ; cá nhân tự quyết toán.' },
    { ma: 'noikhac', ten: '3. HĐLĐ – cá nhân đóng BHXH nơi khác', loaiLuong: 'khongbh', lyDo: 'Đóng BH nơi khác', moTa: 'BHXH, BHTN đóng theo HĐLĐ đầu tiên (nơi khác); DN chỉ đóng BH TNLĐ-BNN 0,5%. Thuế: biểu lũy tiến; giảm trừ gia cảnh chỉ tính tại 1 nơi (mặc định: KHÔNG giảm trừ tại Công ty); có 2 nguồn thu nhập nên không ủy quyền quyết toán được – cá nhân tự quyết toán.' },
    { ma: 'huu', ten: '4. HĐLĐ với người đã nghỉ hưu', loaiLuong: 'khongbh', lyDo: 'Đã nghỉ hưu', moTa: 'Không thuộc diện BHXH bắt buộc (khoản 7 Điều 2 Luật BHXH 2024); BHYT theo thẻ hưu. Lương hưu không chịu thuế. Thuế tiền lương: biểu lũy tiến (HĐLĐ từ 3 tháng), giảm trừ gia cảnh tại Công ty, được ủy quyền quyết toán.' },
    { ma: 'thuviec', ten: '5. HĐ thử việc – không đóng BH', loaiLuong: 'khongbh', lyDo: 'Hợp đồng thử việc', moTa: 'HĐ thử việc riêng không thuộc diện BHXH bắt buộc. Thuế: khấu trừ 10% khi trả từ ngưỡng, có cam kết 08 thì không khấu trừ. Thử việc ghi trong HĐLĐ từ 3 tháng: đổi trạng thái thuế sang "lũy tiến".' }
  ];
  function truongHopNV(nv) {
    var l = loaiLuongNV(nv), ly = String((nv && nv.lyDoKhongBH) || '');
    if (l === 'bh') return 'bh'; if (l === 'khoan') return 'khoan';
    if (ly === 'Đã nghỉ hưu') return 'huu'; if (ly === 'Hợp đồng thử việc') return 'thuviec'; if (ly === 'Đóng BH nơi khác') return 'noikhac'; if (ly === 'Khoán việc, thời vụ, CTV') return 'khoan';
    return nv && nv.loaiHD === 'Thử việc' ? 'thuviec' : 'khac';      // hồ sơ cũ chưa ghi lý do: giữ cách tính cũ (có giảm trừ) – đề nghị chọn lại trường hợp
  }
  var TH_KHAC = { ma: 'khac', ten: 'Không đóng BH tại Công ty – chưa ghi lý do (chọn lại 1 trong 5 trường hợp)', loaiLuong: 'khongbh', lyDo: '', moTa: 'Hồ sơ cũ chưa ghi lý do không đóng BH: tính như trước (không BH, thuế theo loại HĐ, có giảm trừ gia cảnh). Đề nghị chọn đúng trường hợp 3, 4 hoặc 5.' };
  /* TRẠNG THÁI TÍNH THUẾ THEO THỜI ĐIỂM – NhanVien.thueTT = [{ tu:'YYYY-MM', cach, giamTru, camKet, uyQuyen, ghiChu }]
     cach: 'tudong' | 'luytien' | 'khautru10'; giamTru: 'tudong' | 'co' | 'khong'; camKet, uyQuyen: 'tudong' | 'co' | 'khong'. Dòng có "tu" gần nhất ≤ kỳ được áp. */
  var TT_CACH = { luytien: 'Biểu lũy tiến từng phần (cá nhân cư trú, HĐLĐ từ 3 tháng)', khautru10: 'Khấu trừ 10% khi trả từ ngưỡng (HĐ dưới 3 tháng, thử việc, khoán, dịch vụ)', kinhdoanh: 'Cá nhân kinh doanh độc lập – không khấu trừ thuế tiền công (cá nhân tự khai thuế kinh doanh)' };
  function chuanTT(ds) {
    var out = [], da = {}, v3 = function (x) { return x === 'co' || x === 'khong' ? x : 'tudong'; };
    (Array.isArray(ds) ? ds : parseJSON(ds, [])).forEach(function (r) {
      if (!r || !/^\d{4}-(0[1-9]|1[0-2])$/.test(String(r.tu || '')) || da[r.tu]) return; da[r.tu] = 1;
      out.push({ tu: r.tu, cach: TT_CACH[r.cach] ? r.cach : 'tudong', giamTru: v3(r.giamTru), camKet: v3(r.camKet), uyQuyen: v3(r.uyQuyen), ghiChu: String(r.ghiChu || '').slice(0, 300) });
    });
    return out.sort(function (a, b) { return a.tu < b.tu ? -1 : 1; });
  }
  function ttThue(nv, ky) {
    var th = truongHopNV(nv), duoi3 = nv._hdDuoi3 != null ? !!nv._hdDuoi3 : (hdDs(nv).length ? hdDuoi3Thang(nv.loaiHD, nv.ngayKyHD, nv.ngayHetHanHD) : false);   // v3.33: HĐ xác định thời hạn dưới 3 tháng → khấu trừ cố định
    var hdLD = HD_VANGLAI.indexOf(nv.loaiHD) < 0 && !duoi3 && th !== 'khoan' && th !== 'thuviec';
    var hsT = hsThueNV(nv);                          // v3.27: nơi đăng ký giảm trừ (trường hợp 3), khoán kinh doanh độc lập
    var o = { truongHop: th, cach: hdLD ? 'luytien' : (th === 'khoan' && hsT.khoan.kinhDoanh === 'co' ? 'kinhdoanh' : 'khautru10'), giamTru: th === 'noikhac' ? (hsT.noiKhac.gtTai === 'tapaco' ? 'co' : 'khong') : 'co', camKet: nv.camKet08 === 'Có' ? 'co' : 'khong', uyQuyen: hdLD && th !== 'noikhac' ? 'co' : 'khong', tu: '', ghiChu: '' };
    var r = null; chuanTT(nv.thueTT).forEach(function (x) { if (!ky || x.tu <= ky) r = x; });
    if (r) { ['cach', 'giamTru', 'camKet', 'uyQuyen'].forEach(function (k) { if (r[k] !== 'tudong') o[k] = r[k]; }); o.tu = r.tu; o.ghiChu = r.ghiChu; }
    o.moTa = (o.cach === 'luytien' ? 'Lũy tiến' + (o.giamTru === 'co' ? ', giảm trừ gia cảnh tại Công ty' : ', không giảm trừ gia cảnh tại Công ty') : 'Khấu trừ 10%' + (o.camKet === 'co' ? ' – có cam kết 08 (không khấu trừ)' : '')) + (o.uyQuyen === 'co' ? ' · ủy quyền Công ty quyết toán' : ' · tự quyết toán');
    if (o.cach === 'kinhdoanh') o.moTa = 'Cá nhân kinh doanh – không khấu trừ tiền công · tự khai thuế';
    return o;
  }
  function loaiLuongNV(nv) { var l = nv && nv.loaiLuong; if (LOAI_LUONG[l]) return l; return nv && nv.thamGiaBH === 'Không' ? 'khongbh' : 'bh'; }
  /* khoản thu nhập của lương khoán việc: tiền khoán nghiệm thu trong kỳ (trống → mức khoán tháng trong hồ sơ) + thưởng + truy lĩnh */
  var CT_KHOAN = [
    { ma: 'KHOAN', ten: 'Tiền khoán việc', nhom: 'thunhap', phanLoai: 'luong', congThuc: 'IF(TIEN_KHOAN > 0, TIEN_KHOAN, LUONG_CHINH)', thue: 'toanbo', giamTruQC: false },
    { ma: 'THUONGKT', ten: 'Thưởng theo quyết định', nhom: 'thunhap', phanLoai: 'thuong', congThuc: 'THUONG', thue: 'congthuc', thueGT: 'THUONG_CHIU_THUE', giamTruQC: true },
    { ma: 'TRUYLINH', ten: 'Truy lĩnh', nhom: 'thunhap', phanLoai: 'khac', congThuc: 'TRUY_LINH', thue: 'toanbo', giamTruQC: false }
  ];

  /* ==================== BỘ TÍNH CÔNG THỨC (an toàn, không dùng eval) ====================
     Cú pháp kiểu Excel: + - * / ( ) so sánh > < >= <= = <> ; hàm MIN MAX ROUND FLOOR CEIL ABS IF AND OR NOT
     Tên biến không phân biệt hoa thường, chỉ gồm chữ không dấu, số, dấu _ */
  var _ctCache = {};
  var HAM = ['MIN', 'MAX', 'ROUND', 'FLOOR', 'CEIL', 'ABS', 'IF', 'AND', 'OR', 'NOT'];
  function phanTichCT(src) {
    src = String(src == null ? '' : src);
    if (_ctCache[src]) return _ctCache[src];
    var t = [], i = 0;
    while (i < src.length) {
      var c = src[i];
      if (/\s/.test(c)) { i++; continue; }
      if (/[0-9.]/.test(c)) { var j = i; while (j < src.length && /[0-9._]/.test(src[j])) j++; var sv = src.slice(i, j).replace(/_/g, ''); if (!/^(\d+\.?\d*|\.\d+)$/.test(sv)) throw new Error('Số không hợp lệ "' + src.slice(i, j) + '"'); t.push(['n', parseFloat(sv)]); i = j; continue; }
      if (/[A-Za-z_]/.test(c)) { var k = i; while (k < src.length && /[A-Za-z0-9_]/.test(src[k])) k++; t.push(['id', src.slice(i, k).toUpperCase()]); i = k; continue; }
      var two = src.substr(i, 2);
      if (['>=', '<=', '<>', '!=', '=='].indexOf(two) >= 0) { t.push(['op', two === '!=' ? '<>' : two === '==' ? '=' : two]); i += 2; continue; }
      if ('+-*/(),<>=;'.indexOf(c) >= 0) { t.push(['op', c === ';' ? ',' : c]); i++; continue; }
      throw new Error('Ký tự không hợp lệ "' + c + '" (dùng tên biến không dấu)');
    }
    var p = 0;
    function peek(v) { return t[p] && t[p][0] === 'op' && t[p][1] === v; }
    function need2(v) { if (!peek(v)) throw new Error('Thiếu "' + v + '"'); p++; }
    function expr() { var a = add(); while (t[p] && t[p][0] === 'op' && ['>', '<', '>=', '<=', '=', '<>'].indexOf(t[p][1]) >= 0) { var o = t[p++][1]; a = ['b', o, a, add()]; } return a; }
    function add() { var a = mul(); while (peek('+') || peek('-')) { var o = t[p++][1]; a = ['b', o, a, mul()]; } return a; }
    function mul() { var a = una(); while (peek('*') || peek('/')) { var o = t[p++][1]; a = ['b', o, a, una()]; } return a; }
    function una() { if (peek('-')) { p++; return ['u', una()]; } if (peek('+')) { p++; return una(); } return pri(); }
    function pri() {
      var x = t[p];
      if (!x) throw new Error('Công thức chưa hoàn chỉnh');
      if (x[0] === 'n') { p++; return ['n', x[1]]; }
      if (x[0] === 'id') {
        p++;
        if (peek('(')) {
          if (HAM.indexOf(x[1]) < 0) throw new Error('Không có hàm ' + x[1]);
          p++; var args = [];
          if (!peek(')')) { args.push(expr()); while (peek(',')) { p++; args.push(expr()); } }
          need2(')');
          if (x[1] === 'IF' && args.length !== 3) throw new Error('IF cần 3 tham số: IF(điều kiện, đúng, sai)');
          return ['f', x[1], args];
        }
        return ['v', x[1]];
      }
      if (peek('(')) { p++; var e = expr(); need2(')'); return e; }
      throw new Error('Không hiểu "' + x[1] + '"');
    }
    if (!t.length) throw new Error('Công thức trống');
    var ast = expr();
    if (p < t.length) throw new Error('Thừa "' + t[p][1] + '"');
    _ctCache[src] = ast;
    return ast;
  }
  function tinhAST(a, b) {
    switch (a[0]) {
      case 'n': return a[1];
      case 'v': if (!(a[1] in b)) throw new Error('Không có biến ' + a[1]); return num(b[a[1]]);
      case 'u': return -tinhAST(a[1], b);
      case 'b':
        var l = tinhAST(a[2], b), r = tinhAST(a[3], b);
        switch (a[1]) {
          case '+': return l + r; case '-': return l - r; case '*': return l * r; case '/': return r ? l / r : 0;
          case '>': return l > r ? 1 : 0; case '<': return l < r ? 1 : 0; case '>=': return l >= r ? 1 : 0; case '<=': return l <= r ? 1 : 0;
          case '=': return Math.abs(l - r) < 1e-9 ? 1 : 0; case '<>': return Math.abs(l - r) >= 1e-9 ? 1 : 0;
        }
        break;
      case 'f':
        var f = a[1], g = a[2];
        if (f === 'IF') return tinhAST(g[0], b) ? tinhAST(g[1], b) : tinhAST(g[2], b);
        var v = g.map(function (x) { return tinhAST(x, b); });
        if (f === 'MIN') return Math.min.apply(null, v);
        if (f === 'MAX') return Math.max.apply(null, v);
        if (f === 'ROUND') { var k = Math.pow(10, v[1] || 0); return Math.round(v[0] * k) / k; }
        if (f === 'FLOOR') return Math.floor(v[0]);
        if (f === 'CEIL') return Math.ceil(v[0]);
        if (f === 'ABS') return Math.abs(v[0]);
        if (f === 'AND') return v.every(function (x) { return x; }) ? 1 : 0;
        if (f === 'OR') return v.some(function (x) { return x; }) ? 1 : 0;
        if (f === 'NOT') return v[0] ? 0 : 1;
    }
    throw new Error('Biểu thức không hợp lệ');
  }
  function bienDung(a, out) {
    out = out || [];
    if (a[0] === 'v' && out.indexOf(a[1]) < 0) out.push(a[1]);
    if (a[0] === 'u') bienDung(a[1], out);
    if (a[0] === 'b') { bienDung(a[2], out); bienDung(a[3], out); }
    if (a[0] === 'f') a[2].forEach(function (x) { bienDung(x, out); });
    return out;
  }
  function tinhCT(src, bien) { return tinhAST(phanTichCT(src), bien); }

  /* biến hệ thống dùng trong công thức */
  var BIEN_HT = [
    ['LUONG_CHINH', 'Lương chính có hiệu lực trong kỳ (theo quyết định lương)'], ['LUONG_BH', 'Lương đóng BH khai trong hồ sơ (0 nếu để trống)'],
    ['CONG_CHUAN', 'Ngày công chuẩn của kỳ'], ['LUONG_NGAY', 'LUONG_CHINH / CONG_CHUAN'], ['LUONG_GIO', 'LUONG_NGAY / GIO_CONG_NGAY'], ['GIO_CONG_NGAY', 'Số giờ làm 1 ngày'],
    ['NGAY_CONG', 'Ngày công làm việc thực tế'], ['NGHI_HL', 'Ngày nghỉ phép năm, việc riêng có lương'], ['NGAY_LE', 'Ngày nghỉ lễ, Tết hưởng lương'],
    ['NGHI_KL', 'Ngày nghỉ không lương'], ['NGHI_OM', 'Ngày nghỉ ốm (BHXH chi trả)'], ['CONG_HL', 'NGAY_CONG + NGHI_HL + NGAY_LE'],
    ['OT150', 'Giờ làm thêm ngày thường'], ['OT200', 'Giờ làm thêm ngày nghỉ tuần'], ['OT300', 'Giờ làm thêm ngày lễ, Tết'],
    ['HS_OT150', 'Hệ số OT ngày thường (Cài đặt)'], ['HS_OT200', 'Hệ số OT ngày nghỉ tuần'], ['HS_OT300', 'Hệ số OT ngày lễ'],
    ['THUONG', 'Tổng khen thưởng đã duyệt gắn kỳ này'], ['THUONG_CHIU_THUE', 'Phần thưởng chịu thuế'], ['TRUY_LINH', 'Truy lĩnh (KT lương nhập)'],
    ['TIEN_KHOAN', 'Tiền khoán việc nghiệm thu trong kỳ (KT lương nhập)'], ['LA_KHOAN', '1 nếu lương khoán việc, ngược lại 0'], ['KHONG_BH', '1 nếu không đóng BH tại Công ty (đóng nơi khác / đã nghỉ hưu / khoán), ngược lại 0'],
    ['SO_NPT', 'Số người phụ thuộc'], ['THAM_NIEN_THANG', 'Số tháng làm việc tính đến cuối kỳ'], ['LA_THU_VIEC', '1 nếu HĐ thử việc, ngược lại 0'],
    ['LUONG_CO_SO', 'Lương cơ sở (Cài đặt)'], ['LUONG_TOI_THIEU_VUNG', 'Lương tối thiểu vùng (Cài đặt)'],
    ['OT_GIO', 'Tổng giờ làm thêm trong kỳ (OT150 + OT200 + OT300)'], ['OT_GIO_LUYKE', 'Giờ làm thêm các kỳ trước trong cùng năm'],
    ['OT_GIO_MIEN', 'Giờ làm thêm trong giới hạn Bộ luật Lao động (≤ giới hạn tháng và phần còn lại của giới hạn năm)'],
    ['TRA_THAY_BH', 'Khoản trả thêm tương đương mức DN đóng BHXH, BHYT, BHTN cho NLĐ không thuộc diện BH bắt buộc – người nghỉ hưu, HĐ thử việc (Điều 168.3 BLLĐ; bật ở Cài đặt) = MIN(LUONG_BH, trần) × tỷ lệ DN'],
    ['OT_TL_MIEN', 'Tỷ lệ tiền làm thêm giờ được miễn thuế TNCN (0–1). Từ kỳ 2026: OT_GIO_MIEN ÷ OT_GIO (Luật Thuế TNCN 109/2025, NĐ 253/2026); trước 2026: phần trả cao hơn giờ thường'],
    ['CONG_CHUAN_NLV', 'Công chuẩn theo nơi làm việc trong kỳ (Cài đặt → Nơi làm việc; 0 nếu chưa khai)'], ['TIEN_AN_NGAY', 'Tiền ăn ca 1 ngày: mức ăn ca riêng của cá nhân (khai báo mức hưởng), nếu không có thì theo nơi làm việc'],
    ['NGAY_DU_CA', 'Số ngày làm đủ 8 giờ (cả ngày) – căn cứ tiền ăn ca theo quy chế: đếm ô X, CT trên bảng chấm công theo ngày (nửa ngày ½, P½, R½… không tính); không có bảng theo ngày thì lấy phần nguyên ngày công'],
    ['SO_BUA', 'Tổng số bữa ăn trong kỳ theo bảng chấm ăn ca (0–3 bữa/ngày)'], ['TIEN_AN_BUA', 'Phụ cấp ăn ca 1 bữa theo nơi làm việc (Cài đặt → Nơi làm việc; 0 = tính ăn ca theo ngày)'],
    ['GIA_BUA', 'Tiền ăn thu 1 bữa theo nơi làm việc (Cài đặt → Nơi làm việc) – trừ vào lương ở cột Ăn ca của Công nợ']
  ];
  var CT_PHANLOAI = { luong: 'Lương', phucap: 'Phụ cấp', tangca: 'Tăng ca, làm thêm', thuong: 'Thưởng', khac: 'Khác' };
  var CT_VAITRO = { congchuan: 'Hiển thị làm ngày công chuẩn', conghl: 'Hiển thị làm ngày công hưởng lương' };

  /* Mẫu 1 – theo bảng lương TAPACO 2026 (file "LƯƠNG CHUẨN VY 2026.xlsx", sheet "Tinh"):
     Tổng lương thỏa thuận (H) × hệ số K ÷ công chuẩn (VP 24 / DA 28) × công tính lương (Y)
     + lương BH (I) ÷ công chuẩn × (lễ + phép + việc riêng); làm tròn trăm đồng; thử việc 85% và cộng giờ làm thêm quy ra công.
     Phần tổng được tách ra lương cơ bản (= mức đóng BH) – các phụ cấp – thưởng năng suất (phần còn lại) để in phiếu lương. */
  /* v3.31: công thức quy đổi giờ làm thêm – trước đây "IF(LA_THU_VIEC, …, 0)" nên HĐLĐ chính thức không được tính tiền làm thêm giờ */
  var OT_QD_CU2 = 'IF(LA_THU_VIEC, ROUND((OT200 * 2 + OT300 * 4 + OT150 * 1.5) / GIO_CONG_NGAY, 2), 0)';
  var OT_QD_MOI = 'IF(LA_THU_VIEC, ROUND((OT200 * 2 + OT300 * 4 + OT150 * 1.5) / GIO_CONG_NGAY, 2), ROUND((OT150 * HS_OT150 + OT200 * HS_OT200 + OT300 * HS_OT300) / GIO_CONG_NGAY, 2))';
  /* v3.32: ăn ca theo số bữa (bảng chấm ăn ca) khi nơi làm việc khai phụ cấp/bữa; nếu không, theo ngày làm đủ 8 giờ như cũ */
  var ANCA_NGAY = 'IF(NGAY_AN > 0, NGAY_AN, NGAY_DU_CA) * TIEN_AN_NGAY';
  var ANCA_BUA = 'IF(SO_BUA * TIEN_AN_BUA > 0, SO_BUA * TIEN_AN_BUA, ' + ANCA_NGAY + ')';
  function vaANCA(ds) {
    var doi = false;
    (ds || []).forEach(function (c) { if (c && c.ma === 'ANCA' && String(c.congThuc).replace(/\s+/g, '') === ANCA_NGAY.replace(/\s+/g, '')) { c.congThuc = ANCA_BUA; doi = true; } });
    return doi;
  }
  function soBuaCC(cc) {
    var ct = parseJSON(cc && cc.chiTietAn, null), t = 0;
    if (!ct || typeof ct !== 'object' || !Object.keys(ct).length) return num(cc && cc.soBua);
    Object.keys(ct).forEach(function (k) { t += Math.max(0, Math.min(3, Math.round(num(ct[k])))); });
    return t;
  }
  /* v3.32: mức ăn giữa ca miễn thuế TNCN: MIN(mức tối đa của kỳ, ngày công thực tế × mức tối đa ÷ 22) */
  function mucMienAnCa(c, cfg) {
    var mt = num(c.thueGT);
    if (c.ma === 'ANCA' && num(cfg.anCaMienThue) > 0 && (!mt || mt === 730000 || mt === 1200000)) mt = num(cfg.anCaMienThue);   // mức theo thời kỳ (730.000 trước 07/2026; 1.200.000 từ 07/2026)
    return mt;
  }
  /* v3.39: công thức mở — biến MUC_MIEN (mức tối đa của thời kỳ), GT (tiền ăn ca của tháng), NGAY_CONG, SO_BUA… và mọi biến của công thức lương.
     Kết quả kẹp trong [0; MUC_MIEN] — không vượt mức miễn thuế tối đa của thời kỳ. */
  var CT_ANCA_MD = 'MIN(MUC_MIEN, NGAY_CONG * MUC_MIEN / 22)';
  function ctMienAnCa(cfg) {
    var t = String((cfg && cfg.anCaMienCT) || '').trim(); if (t) return t;
    if (cfg && cfg.anCaMienTheoCong === false) return 'MUC_MIEN';
    var chia = num(cfg && cfg.anCaMienChia) || 22; return chia === 22 ? CT_ANCA_MD : 'MIN(MUC_MIEN, NGAY_CONG * MUC_MIEN / ' + chia + ')';
  }
  function anCaTheoCong(c, cfg) { return c.ma === 'ANCA'; }
  function mienAnCa(c, cfg, bien, gt) {
    var mt = mucMienAnCa(c, cfg); if (c.ma !== 'ANCA') return mt;
    var b = Object.assign({}, typeof bien === 'object' && bien ? bien : { NGAY_CONG: num(bien) }, { MUC_MIEN: mt, GT: num(gt) }), v;
    try { v = tinhCT(ctMienAnCa(cfg), b); } catch (e) { v = tinhCT(CT_ANCA_MD, b); }       // công thức lỗi → dùng mặc định (không giảm thuế quá mức)
    return Math.max(0, Math.min(mt, R(num(v))));
  }
  function kiemCTAnCa(cfg, src) {                       // v3.39: kiểm tra biến của công thức ăn ca miễn thuế
    var ten = { MUC_MIEN: 1, GT: 1 }, loi = [];
    BIEN_HT.forEach(function (x) { ten[x[0]] = 1; });
    (cfg.phuCap || []).forEach(function (p) { ten['PC_' + String(p.ma).toUpperCase()] = 1; });
    bienNLV(cfg).forEach(function (k) { ten[k] = 1; });
    (cfg.bienKy || []).forEach(function (v) { if (v.ma) ten[v.ma] = 1; });
    (cfg.congThucLuong || []).forEach(function (c) { if (!c.an && c.ma && c.ma !== 'ANCA') ten[c.ma] = 1; });
    try { bienDung(phanTichCT(src)).forEach(function (v) { if (!ten[v]) loi.push('Công thức ăn ca miễn thuế: chưa có biến ' + v + '.'); }); }
    catch (e) { loi.push('Công thức ăn ca miễn thuế: ' + e.message + '.'); }
    return loi;
  }
  function cleanAn(ct, ky) {
    ct = parseJSON(ct, {}); var n = dim(+ky.slice(0, 4), +ky.slice(5)), o = {};
    Object.keys(ct || {}).forEach(function (k) { var d = +k, v = Math.round(num(ct[k])); if (d >= 1 && d <= n && v > 0) o[d] = Math.min(3, v); });
    return o;
  }
  function vaOTQD(ds) {      // vá bộ công thức đang dùng (đã lưu trong Cài đặt) – chỉ khi khoản OT_QD còn đúng công thức cũ
    var doi = false;
    (ds || []).forEach(function (c) { if (c && c.ma === 'OT_QD' && String(c.congThuc).replace(/\s+/g, '') === OT_QD_CU2.replace(/\s+/g, '')) { c.congThuc = OT_QD_MOI; c.ten = 'Công quy đổi từ giờ làm thêm (thử việc: tăng ca ×1,5, CN ×2, lễ ×4; chính thức: theo hệ số làm thêm ở Cài đặt)'; doi = true; } });
    return doi;
  }
  var CT_LUONG_TAPACO = [
    { ma: 'CC', ten: 'Ngày công chuẩn (VP 24 · DA 28)', nhom: 'trunggian', vaiTro: 'congchuan', congThuc: 'IF(CONG_CHUAN_NLV > 0, CONG_CHUAN_NLV, CONG_CHUAN)' },
    { ma: 'CONG_TL', ten: 'Ngày công tính lương', nhom: 'trunggian', vaiTro: 'conghl', congThuc: 'IF(LA_THU_VIEC, NGAY_CONG, IF(NLV_DA, IF(OR(NGAY_CONG < 26, NGAY_CONG > 28), NGAY_CONG, CC), NGAY_CONG))', ghiChu: 'Thử việc: công thực tế (giờ làm thêm quy đổi tách sang khoản "Lương làm thêm giờ"). Dự án: làm 26–28 công tính đủ công chuẩn.' },
    { ma: 'TL_HD', ten: 'Tỷ lệ hưởng theo HĐ (thử việc 85%)', nhom: 'trunggian', congThuc: 'IF(LA_THU_VIEC, 0.85, 1)' },
    { ma: 'HS_HUONG', ten: 'Hệ số hưởng = tỷ lệ HĐ × K × công tính lương ÷ công chuẩn', nhom: 'trunggian', congThuc: 'TL_HD * HS_K * CONG_TL / CC' },
    { ma: 'NGAY_NGHI_CL', ten: 'Ngày nghỉ có lương (lễ + phép + việc riêng)', nhom: 'trunggian', congThuc: 'NGAY_LE + NGHI_HL' },
    { ma: 'TONG_TG', ten: 'Tổng lương thời gian (làm tròn trăm đồng)', nhom: 'trunggian', congThuc: 'ROUND(LUONG_CHINH * HS_HUONG + LUONG_BH * TL_HD / CC * NGAY_NGHI_CL, -2)' },
    { ma: 'OT_QD', ten: 'Công quy đổi từ giờ làm thêm (thử việc: tăng ca ×1,5, CN ×2, lễ ×4; chính thức: theo hệ số làm thêm ở Cài đặt)', nhom: 'trunggian', congThuc: OT_QD_MOI, ghiChu: 'v3.31: tính cho mọi CBCNV (trước đây chỉ người thử việc). Chính thức: ngày thường ×1,5, nghỉ tuần ×2, lễ Tết ×3 (Điều 98 BLLĐ; lương ngày lễ đã trả riêng).' },
    { ma: 'TONG_OT', ten: 'Tổng lương thời gian kể cả giờ làm thêm (làm tròn trăm đồng)', nhom: 'trunggian', congThuc: 'ROUND(LUONG_CHINH * TL_HD * HS_K * (CONG_TL + OT_QD) / CC + LUONG_BH * TL_HD / CC * NGAY_NGHI_CL, -2)' },
    { ma: 'LCB', ten: 'Lương cơ bản (theo mức đóng BH)', nhom: 'thunhap', phanLoai: 'luong', congThuc: 'LUONG_BH * HS_HUONG', thue: 'toanbo', giamTruQC: false },
    { ma: 'PCNO', ten: 'Tiền nhà', nhom: 'thunhap', phanLoai: 'phucap', congThuc: 'PC_NHAO * HS_HUONG', thue: 'toanbo', giamTruQC: true },
    { ma: 'PCXX', ten: 'Tiền xăng xe', nhom: 'thunhap', phanLoai: 'phucap', congThuc: 'PC_XANGXE * HS_HUONG', thue: 'toanbo', giamTruQC: true },
    { ma: 'PCDT', ten: 'Tiền điện thoại', nhom: 'thunhap', phanLoai: 'phucap', congThuc: 'PC_DIENTHOAI * HS_HUONG', thue: 'toanbo', giamTruQC: true },
    { ma: 'PCCC', ten: 'Tiền chuyên cần', nhom: 'thunhap', phanLoai: 'phucap', congThuc: 'PC_CHUYENCAN * HS_HUONG', thue: 'toanbo', giamTruQC: true },
    { ma: 'PCTN', ten: 'Phụ cấp trách nhiệm', nhom: 'thunhap', phanLoai: 'phucap', congThuc: 'PC_TRACHNHIEM * HS_HUONG', thue: 'toanbo', giamTruQC: true },
    { ma: 'LNGHI', ten: 'Lương lễ, phép, việc riêng', nhom: 'thunhap', phanLoai: 'luong', congThuc: 'LUONG_BH * TL_HD / CC * NGAY_NGHI_CL', thue: 'toanbo', giamTruQC: false },
    { ma: 'THUONGNS', ten: 'Thưởng năng suất, doanh thu', nhom: 'thunhap', phanLoai: 'thuong', congThuc: 'TONG_TG - (LCB + PCNO + PCXX + PCDT + PCCC + PCTN + LNGHI)', thue: 'toanbo', giamTruQC: true, ghiChu: 'Phần còn lại của lương thỏa thuận sau lương cơ bản và phụ cấp (đã gồm làm tròn).' },
    { ma: 'TANGCA', ten: 'Lương làm thêm giờ', nhom: 'thunhap', phanLoai: 'tangca', congThuc: 'TONG_OT - TONG_TG', thue: 'congthuc', thueGT: 'GT * (1 - OT_TL_MIEN)', giamTruQC: false, ghiChu: 'Tách riêng tiền làm thêm giờ (tổng lương không đổi). Miễn thuế TNCN phần giờ trong giới hạn luật lao động (Luật Thuế TNCN 109/2025, NĐ 253/2026); không tính vào lương đóng BH.' },
    { ma: 'ANCA', ten: 'Tiền ăn ca', nhom: 'thunhap', phanLoai: 'phucap', congThuc: ANCA_BUA, thue: 'mientoida', thueGT: 1200000, giamTruQC: false, ghiChu: 'Quy chế: chỉ ngày làm đủ 8 giờ được ăn ca (VP, DA theo mức ở Cài đặt → Nơi làm việc); làm nửa ngày không có. Có bảng chấm ăn ca và nơi làm việc khai phụ cấp/bữa: tính theo số bữa. Miễn thuế tối đa MIN(1,2 triệu; ngày công thực tế × 1,2 triệu ÷ 22).' },
    { ma: 'THUONGKT', ten: 'Thưởng theo quyết định', nhom: 'thunhap', phanLoai: 'thuong', congThuc: 'THUONG', thue: 'congthuc', thueGT: 'THUONG_CHIU_THUE', giamTruQC: true },
    { ma: 'TRUYLINH', ten: 'Truy lĩnh', nhom: 'thunhap', phanLoai: 'khac', congThuc: 'TRUY_LINH', thue: 'toanbo', giamTruQC: false }
  ];
  var BIEN_KY_TAPACO = [
    { ma: 'HS_K', ten: 'Hệ số thưởng K (1–1,2)', buoc: 'tinhluong', macDinh: 1 },
    { ma: 'NGAY_AN', ten: 'Ngày tính ăn ca (trống = số ngày làm đủ 8 giờ)', buoc: 'chamcong', macDinh: 0 }
  ];
  /* Mẫu 2 – chuẩn theo BLLĐ: lương chính ÷ công chuẩn × công, phụ cấp theo công, OT theo hệ số */
  var CT_LUONG_CHUAN = [
    { ma: 'LTG', ten: 'Lương thời gian', nhom: 'thunhap', phanLoai: 'luong', congThuc: 'LUONG_NGAY * NGAY_CONG', thue: 'toanbo', giamTruQC: false },
    { ma: 'LPHEP', ten: 'Lương nghỉ phép, việc riêng', nhom: 'thunhap', phanLoai: 'luong', congThuc: 'LUONG_NGAY * NGHI_HL', thue: 'toanbo', giamTruQC: false },
    { ma: 'LLE', ten: 'Lương lễ, Tết', nhom: 'thunhap', phanLoai: 'luong', congThuc: 'LUONG_NGAY * NGAY_LE', thue: 'toanbo', giamTruQC: false },
    { ma: 'PCCV', ten: 'Phụ cấp chức vụ', nhom: 'thunhap', phanLoai: 'phucap', congThuc: 'PC_CHUCVU * MIN(CONG_HL, CONG_CHUAN) / CONG_CHUAN', thue: 'toanbo', giamTruQC: true },
    { ma: 'PCTN', ten: 'Phụ cấp trách nhiệm', nhom: 'thunhap', phanLoai: 'phucap', congThuc: 'PC_TRACHNHIEM * MIN(CONG_HL, CONG_CHUAN) / CONG_CHUAN', thue: 'toanbo', giamTruQC: true },
    { ma: 'PCCT', ten: 'Phụ cấp công trường', nhom: 'thunhap', phanLoai: 'phucap', congThuc: 'PC_CONGTRUONG * MIN(NGAY_CONG, CONG_CHUAN) / CONG_CHUAN', thue: 'toanbo', giamTruQC: true },
    { ma: 'PCNO', ten: 'Phụ cấp nhà ở', nhom: 'thunhap', phanLoai: 'phucap', congThuc: 'PC_NHAO * MIN(CONG_HL, CONG_CHUAN) / CONG_CHUAN', thue: 'toanbo', giamTruQC: true },
    { ma: 'PCDT', ten: 'Phụ cấp điện thoại', nhom: 'thunhap', phanLoai: 'phucap', congThuc: 'PC_DIENTHOAI * MIN(CONG_HL, CONG_CHUAN) / CONG_CHUAN', thue: 'toanbo', giamTruQC: true },
    { ma: 'PCXX', ten: 'Phụ cấp xăng xe', nhom: 'thunhap', phanLoai: 'phucap', congThuc: 'PC_XANGXE * MIN(CONG_HL, CONG_CHUAN) / CONG_CHUAN', thue: 'toanbo', giamTruQC: true },
    { ma: 'ANCA', ten: 'Tiền ăn giữa ca', nhom: 'thunhap', phanLoai: 'phucap', congThuc: 'PC_ANCA * MIN(NGAY_CONG, CONG_CHUAN) / CONG_CHUAN', thue: 'mientoida', thueGT: 730000, giamTruQC: true },
    { ma: 'TANGCA', ten: 'Lương tăng ca', nhom: 'thunhap', phanLoai: 'tangca', congThuc: 'LUONG_GIO * (OT150 * HS_OT150 + OT200 * HS_OT200 + OT300 * HS_OT300)', thue: 'congthuc', thueGT: 'GT * (1 - OT_TL_MIEN)', giamTruQC: false, ghiChu: 'Miễn thuế TNCN phần giờ làm thêm trong giới hạn luật lao động (Luật Thuế TNCN 109/2025, NĐ 253/2026); giờ vượt giới hạn chịu thuế.' },
    { ma: 'LAMDEM', ten: 'Phụ cấp làm đêm', nhom: 'thunhap', phanLoai: 'tangca', congThuc: 'LUONG_GIO * OT_DEM * 0.3', thue: 'khong', giamTruQC: false },
    { ma: 'KPI', ten: 'Thưởng năng suất KPI', nhom: 'thunhap', phanLoai: 'thuong', congThuc: 'PC_KPI * KPI_PT / 100', thue: 'toanbo', giamTruQC: true },
    { ma: 'THUONGKT', ten: 'Thưởng theo quyết định', nhom: 'thunhap', phanLoai: 'thuong', congThuc: 'THUONG', thue: 'congthuc', thueGT: 'THUONG_CHIU_THUE', giamTruQC: true },
    { ma: 'TRUYLINH', ten: 'Truy lĩnh', nhom: 'thunhap', phanLoai: 'khac', congThuc: 'TRUY_LINH', thue: 'toanbo', giamTruQC: false }
  ];
  var BIEN_KY_CHUAN = [
    { ma: 'OT_DEM', ten: 'Giờ làm đêm', buoc: 'chamcong', macDinh: 0 },
    { ma: 'KPI_PT', ten: '% hoàn thành KPI', buoc: 'tinhluong', macDinh: 100 }
  ];
  var CT_MAU = {
    tapaco: { ten: 'Bảng lương TAPACO 2026 (lương thỏa thuận × K, VP/DA)', congThucLuong: CT_LUONG_TAPACO, bienKy: BIEN_KY_TAPACO, congThucDongBH: 'LUONG_BH' },
    chuan: { ten: 'Chuẩn theo BLLĐ (lương chính ÷ công chuẩn × công, OT theo hệ số)', congThucLuong: CT_LUONG_CHUAN, bienKy: BIEN_KY_CHUAN, congThucDongBH: 'IF(LUONG_BH > 0, LUONG_BH, LUONG_CHINH + PC_CHUCVU + PC_TRACHNHIEM)' }
  };
  var CT_LUONG_MD = CT_LUONG_TAPACO, BIEN_KY_MD = BIEN_KY_TAPACO;
  var CT_DONGBH_MD = CT_MAU.tapaco.congThucDongBH;

  /* nơi làm việc trong kỳ: chấm công ghi đè hồ sơ */
  function nlvTrongKy(nv, cc, cfg) {
    var ma = String((cc && cc.noiLamViec) || nv.noiLamViec || '').toUpperCase();
    var ds = cfg.noiLamViec || [], o = ds.filter(function (x) { return String(x.ma).toUpperCase() === ma; })[0];
    return { ma: o ? ma : '', muc: o || {} };
  }
  function bienNLV(cfg) { return (cfg.noiLamViec || []).map(function (x) { return 'NLV_' + String(x.ma).toUpperCase(); }); }

  function taoBien(nv, cc, cfg, o) {
    var L = num(o.luong), ncc = num(o.ncc) || num(cfg.ngayCongChuan) || 26, gio = num(cfg.gioCongNgay) || 8, h = cfg.heSoOT || {};
    var b = {
      LUONG_CHINH: L, LUONG_BH: num(o.luongBH), CONG_CHUAN: ncc, LUONG_NGAY: L / ncc, LUONG_GIO: L / ncc / gio, GIO_CONG_NGAY: gio,
      NGAY_CONG: num(cc.ngayCong), NGHI_HL: num(cc.nghiHuongLuong), NGAY_LE: num(cc.ngayLe), NGHI_KL: num(cc.nghiKhongLuong), NGHI_OM: num(cc.nghiOm),
      OT150: num(cc.ot150), OT200: num(cc.ot200), OT300: num(cc.ot300), HS_OT150: num(h.ot150) || 1.5, HS_OT200: num(h.ot200) || 2, HS_OT300: num(h.ot300) || 3,
      THUONG: num(o.thuong), THUONG_CHIU_THUE: num(o.thuongChiuThue), TRUY_LINH: num(cc.truyLinh), SO_NPT: num(nv.soNPT), TIEN_KHOAN: num(cc.tienKhoan),
      LA_KHOAN: loaiLuongNV(nv) === 'khoan' ? 1 : 0, KHONG_BH: loaiLuongNV(nv) === 'bh' ? 0 : 1,
      THAM_NIEN_THANG: 0, LA_THU_VIEC: nv.loaiHD === 'Thử việc' ? 1 : 0, LUONG_CO_SO: num(cfg.luongCoSo), LUONG_TOI_THIEU_VUNG: num(cfg.luongToiThieuVung)
    };
    b.CONG_HL = b.NGAY_CONG + b.NGHI_HL + b.NGAY_LE;
    var thB = truongHopNV(nv), dnB = cfg.tyLeDN || {};
    b.TRA_THAY_BH = o.traThayBH != null ? num(o.traThayBH) : cfg.traThayBH === 'co' && (thB === 'huu' || thB === 'thuviec') ? R(Math.min(num(o.luongBH), (num(cfg.luongCoSo) * num(cfg.heSoTranBHXH)) || num(o.luongBH)) * (num(dnB.bhxh) + num(dnB.bhyt) + num(dnB.bhtn)) / 100) : 0;
    // v3.12: giờ làm thêm được miễn thuế TNCN – trong giới hạn Điều 107 BLLĐ (mặc định 40 giờ/tháng, 200 giờ/năm; Cài đặt)
    b.OT_GIO = b.OT150 + b.OT200 + b.OT300; b.OT_GIO_LUYKE = num(o.otLuyKe);
    var ghT = num(cfg.otGioiHanThang) || 40, ghN = num(cfg.otGioiHanNam) || 200;
    b.OT_GIO_MIEN = Math.max(0, Math.min(b.OT_GIO, ghT, ghN - b.OT_GIO_LUYKE));
    if (cfg.otCach ? cfg.otCach === 'chenh' : (o.ky && o.ky < '2026-01')) { var tu = b.OT150 * (b.HS_OT150 - 1) + b.OT200 * (b.HS_OT200 - 1) + b.OT300 * (b.HS_OT300 - 1), mau = b.OT150 * b.HS_OT150 + b.OT200 * b.HS_OT200 + b.OT300 * b.HS_OT300; b.OT_TL_MIEN = mau > 0 ? Math.max(0, tu) / mau : 0; }
    else b.OT_TL_MIEN = b.OT_GIO > 0 ? b.OT_GIO_MIEN / b.OT_GIO : 0;
    var vao = parseD(nv.ngayVao), e = parseD(o.ky ? kyEnd(o.ky) : '');
    if (vao && e) b.THAM_NIEN_THANG = Math.max(0, (e.y - vao.y) * 12 + (e.m - vao.m));
    var nlv = nlvTrongKy(nv, cc, cfg);
    bienNLV(cfg).forEach(function (k) { b[k] = k === 'NLV_' + nlv.ma ? 1 : 0; });
    // v3.13.1: quy chế ăn ca – chỉ ngày làm đủ 8 giờ (cả ngày) mới có tiền ăn ca; làm nửa ngày không có
    var ctN = parseJSON(cc.chiTietNgay, null), coNgay = false, duCa = 0;
    if (ctN && typeof ctN === 'object') Object.keys(ctN).forEach(function (k) { var h = KY_HIEU[ctN[k]]; if (!h) return; coNgay = true; if ((h.cong || 0) >= 1) duCa++; });
    b.NGAY_DU_CA = coNgay ? duCa : Math.floor(b.NGAY_CONG + 1e-9);
    b.CONG_CHUAN_NLV = num(nlv.muc.congChuan); b.TIEN_AN_NGAY = num(nv.anCa) > 0 ? num(nv.anCa) : num(nlv.muc.anCa);
    b.SO_BUA = soBuaCC(cc); b.TIEN_AN_BUA = num(nlv.muc.anCaBua); b.GIA_BUA = num(nlv.muc.giaBua);      // v3.32: bảng chấm ăn ca
    var pc = parseJSON(nv.phuCap, {});
    (cfg.phuCap || []).forEach(function (p) { b['PC_' + String(p.ma).toUpperCase()] = num(pc[p.ma]); });
    var bk = parseJSON(cc.bienKy, {});
    (cfg.bienKy || BIEN_KY_MD).forEach(function (v) { b[v.ma] = (bk[v.ma] === undefined || bk[v.ma] === '') ? num(v.macDinh) : num(bk[v.ma]); });
    return b;
  }
  /* kiểm tra bộ công thức; trả về danh sách lỗi (rỗng = hợp lệ) */
  function kiemTraCT(cfg) {
    var loi = [], ten = {};
    BIEN_HT.forEach(function (x) { ten[x[0]] = 'ht'; });
    (cfg.phuCap || []).forEach(function (p) { ten['PC_' + String(p.ma).toUpperCase()] = 'pc'; });
    bienNLV(cfg).forEach(function (k) { ten[k] = 'nlv'; });
    (cfg.bienKy || []).forEach(function (v) {
      if (!/^[A-Z][A-Z0-9_]{1,24}$/.test(v.ma || '')) loi.push('Biến nhập "' + (v.ten || v.ma) + '": mã chỉ gồm chữ IN HOA không dấu, số, _ (2–25 ký tự).');
      else if (ten[v.ma]) loi.push('Biến nhập "' + v.ma + '" trùng tên biến khác.');
      else ten[v.ma] = 'ky';
      if (v.buoc !== 'chamcong' && v.buoc !== 'tinhluong') loi.push('Biến nhập "' + v.ma + '": chọn bước nhập.');
    });
    var dsCT = cfg.congThucLuong || [];
    if (!dsCT.some(function (c) { return !c.an && c.nhom === 'thunhap'; })) loi.push('Cần ít nhất 1 khoản thu nhập.');
    var vt = {};
    dsCT.forEach(function (c, i) {
      var nhan = 'Khoản ' + (i + 1) + ' "' + (c.ten || c.ma) + '"';
      if (!/^[A-Z][A-Z0-9_]{1,19}$/.test(c.ma || '')) { loi.push(nhan + ': mã chỉ gồm chữ IN HOA không dấu, số, _ (2–20 ký tự).'); return; }
      if (ten[c.ma]) { loi.push(nhan + ': mã "' + c.ma + '" trùng tên biến khác.'); return; }
      if (!String(c.ten || '').trim()) loi.push(nhan + ': thiếu tên.');
      if (c.nhom === 'trunggian' && c.vaiTro && !c.an) { if (vt[c.vaiTro]) loi.push(nhan + ': chỉ 1 khoản được "' + CT_VAITRO[c.vaiTro] + '".'); vt[c.vaiTro] = 1; }
      if (!CT_PHANLOAI[c.phanLoai] && c.nhom === 'thunhap') loi.push(nhan + ': chọn phân loại.');
      var kt = function (src, them, nhanCT) {
        try {
          bienDung(phanTichCT(src)).forEach(function (v) { if (!ten[v] && (!them || them.indexOf(v) < 0)) loi.push(nhan + ' – ' + nhanCT + ': chưa có biến ' + v + (/^PC_/.test(v) ? ' (thêm khoản "' + v.slice(3).toLowerCase() + '" vào danh mục phụ cấp hồ sơ)' : '') + '.'); });
        } catch (e) { loi.push(nhan + ' – ' + nhanCT + ': ' + e.message + '.'); }
      };
      kt(c.congThuc, null, 'công thức');
      if (c.nhom === 'thunhap' && c.thue === 'congthuc') kt(c.thueGT, ['GT'], 'phần chịu thuế');
      if (c.nhom === 'thunhap' && c.thue === 'mientoida' && !(num(c.thueGT) > 0)) loi.push(nhan + ': nhập mức miễn thuế tối đa.');
      if (!c.an) ten[c.ma] = 'ct';
    });
    try { bienDung(phanTichCT(cfg.congThucDongBH || CT_DONGBH_MD)).forEach(function (v) { if (!ten[v]) loi.push('Công thức lương đóng BH: chưa có biến ' + v + '.'); }); }
    catch (e) { loi.push('Công thức lương đóng BH: ' + e.message + '.'); }
    return loi;
  }

  /* ===== v3.33: LỊCH SỬ HỢP ĐỒNG LAO ĐỘNG – NhanVien.lichSuHD = [{ id, soHD, loaiHD, tuNgay, denNgay, luongHD, luongBH, hinhThuc, lyDo, tlId, tenFile, ghiChu, trangThai, lyDoHuy, nguoi, luc, daBao }]
     Không ghi đè: mỗi lần ký / đổi HĐ thêm 1 dòng; sửa sai thì "Hủy dòng" (giữ lại) rồi thêm dòng đúng.
     Bảng lương kỳ nào lấy HĐ có hiệu lực trong kỳ đó; đổi HĐ giữa kỳ → chia theo ngày công của từng giai đoạn. */
  var HT_HD = { bh: 'Trả lương + đóng BH tại Công ty', khongbh: 'Đóng BH nơi khác / đã về hưu (Công ty không trích)', khoan: 'Lương khoán' };
  var HD_KHONG_BH = ['Thử việc', 'Thời vụ / Khoán việc', 'Cộng tác viên (CTV)'];        // không trích BH bắt buộc dù hình thức trả nào
  var LY_DO_HD_KBH = ['Đóng BH nơi khác', 'Đã nghỉ hưu'];
  function hdTatCa(nv) { var a = parseJSON(nv && nv.lichSuHD, []); return Array.isArray(a) ? a : []; }
  function hdDs(nv) {                                    // các dòng còn hiệu lực pháp lý (chưa hủy), theo ngày bắt đầu
    return hdTatCa(nv).filter(function (h) { return h && h.trangThai !== 'Đã hủy' && parseD(h.tuNgay); })
      .sort(function (a, b) { return a.tuNgay < b.tuNgay ? -1 : a.tuNgay > b.tuNgay ? 1 : (String(a.luc) < String(b.luc) ? -1 : 1); });
  }
  function hdTaiNgay(nv, ngay) { var r = null; hdDs(nv).forEach(function (h) { if (h.tuNgay <= ngay) r = h; }); return r; }
  // HĐ dưới 3 tháng (tính theo thời hạn ghi trên HĐ) → khấu trừ 10%; từ đủ 3 tháng trở lên → biểu lũy tiến
  function hdDuoi3Thang(loaiHD, tu, den) { return loaiHD === 'HĐ dưới 3 tháng' || (loaiHD === 'Xác định thời hạn' && !!parseD(tu) && !!parseD(den) && addDays(den, 1) < addMonths(tu, 3)); }
  function hdCoBH(h) { return h.hinhThuc === 'bh' && HD_KHONG_BH.indexOf(h.loaiHD) < 0; }
  function hdMoTa(h) {
    var vl = HD_VANGLAI.indexOf(h.loaiHD) >= 0 || hdDuoi3Thang(h.loaiHD, h.tuNgay, h.denNgay);
    return { bh: hdCoBH(h) ? 'Trích BH theo lương đóng BH' : h.hinhThuc === 'khongbh' ? 'Không trích BH (' + (h.lyDo || 'đóng nơi khác / về hưu') + ')' : 'Không trích BH bắt buộc (' + (h.hinhThuc === 'khoan' ? 'lương khoán' : h.loaiHD.toLowerCase()) + ')',
      thue: vl ? 'Khấu trừ theo tỷ lệ cố định (HĐ dưới 3 tháng / thử việc / khoán / CTV)' : 'Biểu lũy tiến, có giảm trừ gia cảnh (HĐ từ 3 tháng)', vangLai: vl };
  }
  // hồ sơ "theo HĐ": loại HĐ, số, ngày ký / hết hạn, loại lương (BH) của HĐ áp dụng
  function apHD(nv, h) {
    if (!h) return nv;
    var o = Object.assign({}, nv, { loaiHD: h.loaiHD, soHD: h.soHD || '', ngayKyHD: h.tuNgay, ngayHetHanHD: h.denNgay || '' });
    if (h.hinhThuc === 'khoan') { o.loaiLuong = 'khoan'; o.lyDoKhongBH = ''; }
    else if (h.loaiHD === 'Thử việc') { o.loaiLuong = 'khongbh'; o.lyDoKhongBH = 'Hợp đồng thử việc'; }
    else if (h.loaiHD === 'Thời vụ / Khoán việc' || h.loaiHD === 'Cộng tác viên (CTV)') { o.loaiLuong = 'khongbh'; o.lyDoKhongBH = 'Khoán việc, thời vụ, CTV'; }
    else if (h.hinhThuc === 'khongbh') { o.loaiLuong = 'khongbh'; o.lyDoKhongBH = h.lyDo || 'Đóng BH nơi khác'; }
    else { o.loaiLuong = 'bh'; o.lyDoKhongBH = ''; }
    o.thamGiaBH = o.loaiLuong === 'bh' ? 'Có' : 'Không';
    o._hdDuoi3 = hdDuoi3Thang(h.loaiHD, h.tuNgay, h.denNgay) ? 1 : 0;
    return o;
  }
  // giai đoạn HĐ trong kỳ (trong thời gian làm việc): [{ hd, tu, den }] – rỗng nếu chưa có lịch sử HĐ
  function hdDoanKy(nv, ky) {
    var ds = hdDs(nv); if (!ds.length) return [];
    var a = kyStart(ky), b = kyEnd(ky);
    if (nv.ngayVao && nv.ngayVao > a) a = nv.ngayVao;
    if (nv.ngayNghi && addDays(nv.ngayNghi, -1) < b) b = addDays(nv.ngayNghi, -1);
    if (a > b) return [];
    var out = [], h0 = hdTaiNgay(nv, a);
    var moc = ds.filter(function (h) { return h.tuNgay > a && h.tuNgay <= b; });
    var cur = h0 || (moc.length ? null : ds[ds.length - 1]), tu = a;
    moc.forEach(function (h) { if (cur) out.push({ hd: cur, tu: tu, den: addDays(h.tuNgay, -1) }); cur = h; tu = h.tuNgay; });
    if (cur) out.push({ hd: cur, tu: tu, den: b });
    return out.filter(function (x) { return x.tu <= x.den; });
  }
  // đồng bộ các trường "hiện tại" của hồ sơ theo HĐ đang hiệu lực hôm nay (hồ sơ, cảnh báo, báo cáo)
  function hdDongBo(nv, homNay) {
    var h = hdTaiNgay(nv, homNay) || hdDs(nv)[0]; if (!h) return;
    var o = apHD({}, h);
    ['loaiHD', 'soHD', 'ngayKyHD', 'ngayHetHanHD', 'loaiLuong', 'lyDoKhongBH', 'thamGiaBH'].forEach(function (k) { nv[k] = o[k] == null ? '' : o[k]; });
    if (nv.trangThai === 'Thử việc' && h.loaiHD !== 'Thử việc') nv.trangThai = 'Đang làm việc';
    else if (nv.trangThai === 'Đang làm việc' && h.loaiHD === 'Thử việc') nv.trangThai = 'Thử việc';
  }
  // sự kiện mức hưởng từ HĐ: lương HĐ, lương đóng BH có hiệu lực từ ngày bắt đầu HĐ (quyết định lên lương sau ngày đó vẫn áp sau)
  function hdSuKien(nv) {
    return hdDs(nv).map(function (h) {
      var gt = {}; if (num(h.luongHD) > 0) gt.luongChinh = num(h.luongHD); if (h.luongBH !== undefined && h.luongBH !== '') gt.luongBH = num(h.luongBH);
      return { id: h.id, nguon: 'hd', hieuLuc: h.tuNgay, t: String(h.luc || ''), gt: gt, cu: {} };
    });
  }
  function hdCanhBao(nv, homNay, soNgay) {             // HĐ hiện hành sắp hết hạn / đã hết hạn mà chưa có HĐ mới
    var h = hdTaiNgay(nv, homNay); if (!h || !h.denNgay || nv.trangThai === 'Đã nghỉ việc') return null;
    if (hdDs(nv).some(function (x) { return x.tuNgay > homNay; })) return null;     // đã ký HĐ tiếp theo
    var con = soNgayLich(homNay, h.denNgay) - 1;
    return con <= (num(soNgay) || 30) ? { hd: h, con: con } : null;
  }
  /* Bảng lương nhiều giai đoạn HĐ trong 1 kỳ: mỗi giai đoạn tính theo HĐ của giai đoạn đó, ngày công tách theo bảng chấm công ngày
     (không có thì chia theo số ngày làm việc); thưởng, truy lĩnh, tạm ứng, công nợ, vi phạm, biến nhập tính ở giai đoạn cuối.
     BH: theo HĐ có hiệu lực cuối kỳ (BH đóng theo tháng). Thuế: gộp thu nhập chịu thuế cả tháng theo từng cách tính
     (lũy tiến: trừ BH, giảm trừ gia cảnh 1 lần; khấu trừ cố định: trên phần thu nhập của HĐ dưới 3 tháng / thử việc / khoán / CTV). */
  var CC_CHIA = ['ngayCong', 'nghiHuongLuong', 'nghiKhongLuong', 'nghiOm', 'ngayLe'];
  var CC_CUOI = ['tamUng', 'khauTruKhac', 'thuUng', 'tienAnThu', 'truyLinh', 'tienKhoan', 'chiTietPhat', 'chiTietTamUng', 'bienKy', 'soBua'];
  function catNgay(ct, ky, tu, den) { var o = {}; Object.keys(ct || {}).forEach(function (d) { var s = ky + '-' + pad(+d); if (s >= tu && s <= den) o[d] = ct[d]; }); return o; }
  function ccGiaiDoan(cc, ky, seg, cuoi, tongNgayLV, cfg) {
    var o = Object.assign({}, cc), ctN = parseJSON(cc.chiTietNgay, null), ctO = parseJSON(cc.chiTietOT, null), ctA = parseJSON(cc.chiTietAn, null);
    if (ctN && Object.keys(ctN).length) {
      var p = catNgay(ctN, ky, seg.tu, seg.den), th = tongHopNgay(p, ky, cfg);
      o.chiTietNgay = JSON.stringify(p); CC_CHIA.forEach(function (k) { o[k] = th[k]; });
      o.ot300 = th.ot300; if (cfg.chuNhatTinh === 'ot') o.ot200 = th.ot200;
    } else {
      var f = tongNgayLV > 0 ? soNgayLamViec(seg.tu, seg.den) / tongNgayLV : (cuoi ? 1 : 0);
      CC_CHIA.forEach(function (k) { o[k] = Math.round(num(cc[k]) * f * 100) / 100; });
    }
    if (ctO && Object.keys(ctO).length) { var po = catNgay(ctO, ky, seg.tu, seg.den); o.chiTietOT = JSON.stringify(po); if (!(ctN && Object.keys(ctN).length)) { o.ot300 = 0; o.ot200 = 0; } o.ot150 = 0; apGioOT(o, ky, cfg); }
    else if (!cuoi && !(ctN && Object.keys(ctN).length)) { o.ot150 = 0; o.ot200 = 0; o.ot300 = 0; }
    else if (!cuoi) o.ot150 = 0;
    if (ctA && Object.keys(ctA).length) o.chiTietAn = JSON.stringify(catNgay(ctA, ky, seg.tu, seg.den));
    if (!cuoi) CC_CUOI.forEach(function (k) { o[k] = /^chiTiet|bienKy/.test(k) ? '' : 0; });
    return o;
  }
  function tinhNhieuDoan(nv, cc, cfg, dv, segs, ll, mh) {
    cfg = cfgTheoKy(cfg, dv.o.ky);
    var ky = dv.o.ky, tongLV = segs.reduce(function (t, s) { return t + soNgayLamViec(s.tu, s.den); }, 0), ps = [], otTruoc = 0;
    segs.forEach(function (seg, i) {
      var cuoi = i === segs.length - 1, nvH = apHD(nv, seg.hd), m = mucHuong(nvH, ll, mh, seg.den);
      var nvI = Object.assign({}, nvH, { luongChinh: m.luongChinh, luongBH: m.luongBH, soNPT: m.soNPT, anCa: m.anCa, phuCap: JSON.stringify(m.phuCap) });
      var ccI = ccGiaiDoan(cc, ky, seg, cuoi, tongLV, cfg);
      var oI = Object.assign({}, dv.o, { luong: m.luongChinh, luongBH: m.luongBH, otLuyKe: num(dv.o.otLuyKe) + otTruoc });
      if (!cuoi) { oI.thuong = 0; oI.thuongChiuThue = 0; oI.khongBH = true; oI.traThayBH = 0; }
      var p = tinhPhieuLuong(nvI, ccI, cfg, oI);
      otTruoc += num(ccI.ot150) + num(ccI.ot200) + num(ccI.ot300);
      ps.push({ seg: seg, nv: nvI, p: p, tt: ttThue(nvI, ky), cong: num(ccI.ngayCong) });
    });
    var L = ps[ps.length - 1], out = Object.assign({}, L.p), S = function (k) { return ps.reduce(function (t, x) { return t + num(x.p[k]); }, 0); };
    ['luongThoiGian', 'phuCap', 'tienOT', 'thuong', 'truyLinh', 'giamThuong', 'ktCongThuc', 'tongThuNhap', 'tnChiuThue', 'anCaMien', 'ngayCongHL'].forEach(function (k) { out[k] = S(k); });
    out.ngayCongHL = Math.round(out.ngayCongHL * 100) / 100;
    var ct = [], vt = {};
    ps.forEach(function (x) { parseJSON(x.p.chiTiet, []).forEach(function (c) { if (vt[c.ma] === undefined) { vt[c.ma] = ct.length; ct.push(Object.assign({}, c)); } else { var y = ct[vt[c.ma]]; y.gt = num(y.gt) + num(c.gt); if (c.thue !== undefined) y.thue = num(y.thue) + num(c.thue); } }); });
    out.chiTiet = JSON.stringify(ct);
    // thuế gộp cả tháng theo cách tính của từng giai đoạn
    var bhNLD = num(L.p.bhxh) + num(L.p.bhyt) + num(L.p.bhtn), tinhThue = function (boOT) {
      var lt = 0, vl = 0, giamTru = 0, coLT = false, ckVL = false;
      ps.forEach(function (x) {
        var tn = num(x.p.tnChiuThue) - (boOT ? parseJSON(x.p.chiTiet, []).reduce(function (t, c) { return t + (c.nhom === 'thunhap' && c.phanLoai === 'tangca' ? num(c.thue) : 0); }, 0) : 0);
        if (x.tt.cach === 'luytien') { lt += tn; coLT = true; if (x.tt.giamTru === 'co') giamTru = num(cfg.giamTruBanThan) + num(x.nv.soNPT) * num(cfg.giamTruNPT); }
        else if (x.tt.cach === 'khautru10') { vl += tn; if (x.tt.camKet === 'co' && !(dv.o && dv.o.ckHopLe === false)) ckVL = true; }
      });
      var tntt = coLT ? Math.max(0, lt - bhNLD - giamTru) : 0, thL = coLT ? thueLuyTien(tntt, cfg.bieuThue || DEFAULT_CFG.bieuThue) : 0;
      var thV = (ckVL || vl < num(cfg.nguongVangLai)) ? 0 : R(vl * num(cfg.thueVangLai) / 100);
      return { thue: thL + thV, tntt: tntt + vl, giamTru: coLT ? giamTru : 0 };
    };
    var t1 = tinhThue(false), t2 = tinhThue(true);
    out.thueTNCN = t1.thue; out.tnTinhThue = t1.tntt; out.giamTru = t1.giamTru;
    var vp = parseJSON(cc.chiTietPhat, []), btYC = 0;
    (Array.isArray(vp) ? vp : []).forEach(function (x) { if (x.loai === 'Bồi thường thiệt hại') btYC += num(x.thanhTien) || num(x.mucTien) * (num(x.soLan) || 1); });
    var btMax = Math.max(0, R((out.tongThuNhap - bhNLD - out.thueTNCN) * (num(cfg.tyLeBoiThuongToiDa) || 30) / 100));
    out.boiThuong = Math.min(btYC, btMax);
    out.thucLinh = out.tongThuNhap - bhNLD - num(L.p.doanPhi) - out.thueTNCN - out.boiThuong - num(L.p.tongNo) - out.ktCongThuc;
    var otGT = ct.reduce(function (t, c) { return t + (c.nhom === 'thunhap' && c.phanLoai === 'tangca' ? num(c.gt) : 0); }, 0);
    out.otThue = Math.max(0, Math.min(otGT, t1.thue - t2.thue)); out.otThucLinh = otGT - out.otThue;
    out.cachThue = ps.map(function (x) { return x.tt.moTa.split(' · ')[0]; }).filter(function (v, i, a) { return a.indexOf(v) === i; }).join(' + ');
    out.ctHD = JSON.stringify(ps.map(function (x) { return { soHD: x.seg.hd.soHD || '', loaiHD: x.seg.hd.loaiHD, tu: x.seg.tu, den: x.seg.den, cong: x.cong, luong: num(x.nv.luongChinh), tong: num(x.p.tongThuNhap), tnct: num(x.p.tnChiuThue), cach: x.tt.cach }; }));
    out.ghiChu = 'Đổi HĐ trong kỳ: ' + ps.map(function (x) { return (x.seg.hd.soHD ? 'HĐ ' + x.seg.hd.soHD + ' ' : '') + x.seg.hd.loaiHD + ' ' + ngayVN(x.seg.tu).slice(0, 5) + '–' + ngayVN(x.seg.den).slice(0, 5) + ' (' + String(x.cong).replace('.', ',') + ' công)'; }).join(' → ') + '. BH theo HĐ cuối kỳ; thuế gộp cả tháng. ' + (L.p.ghiChu || '');
    return out;
  }

  /* ===== v3.37: OCR HỒ SƠ – ảnh chụp / bản scan PDF → văn bản (lưu riêng tệp .txt cạnh tệp gốc) → đề xuất điền lý lịch (người soát) =====
     Đọc chữ: Google Drive OCR trên Apps Script (dữ liệu ở trong Google Drive Công ty) hoặc đọc ngay trên máy (Tesseract) – KHÔNG gửi AI bên ngoài. */
  var OCR_TRUONG = [['cccd', 'Số CCCD'], ['ngaySinh', 'Ngày sinh', 'date'], ['gioiTinh', 'Giới tính', 'gt'], ['ngayCapCCCD', 'Ngày cấp CCCD', 'date'], ['noiCapCCCD', 'Nơi cấp CCCD'],
    ['queQuan', 'Quê quán'], ['diaChi', 'Địa chỉ thường trú'], ['mst', 'Mã số thuế cá nhân'], ['trinhDo', 'Trình độ'], ['chuyenNganh', 'Chuyên ngành']];
  var OCR_TL = [['soHieu', 'Số hiệu tài liệu'], ['ngay', 'Ngày ký / ngày cấp', 'date'], ['hetHan', 'Có giá trị đến', 'date']];
  function b64Utf8(str) {
    var s = unescape(encodeURIComponent(String(str || ''))), A = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/', o = '', i = 0;
    for (; i < s.length; i += 3) { var a = s.charCodeAt(i), b = s.charCodeAt(i + 1), c = s.charCodeAt(i + 2), n = (a << 16) | ((b || 0) << 8) | (c || 0);
      o += A[(n >> 18) & 63] + A[(n >> 12) & 63] + (i + 1 < s.length ? A[(n >> 6) & 63] : '=') + (i + 2 < s.length ? A[n & 63] : '='); }
    return o;
  }
  function ocrBoDau(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase(); }
  function ocrNgay(s) {                               // dd/mm/yyyy, dd-mm-yyyy, dd.mm.yyyy, "ngày dd tháng mm năm yyyy" → yyyy-mm-dd
    var m = /(\d{1,2})\s*[\/.\-]\s*(\d{1,2})\s*[\/.\-]\s*(\d{4})/.exec(s) || /ng[aà]y\s*(\d{1,2})\s*th[aá]ng\s*(\d{1,2})\s*n[aă]m\s*(\d{4})/i.exec(s);
    if (!m) return ''; var v = m[3] + '-' + pad(+m[2]) + '-' + pad(+m[1]); return parseD(v) ? v : '';
  }
  // lấy giá trị sau nhãn (cùng dòng sau dấu ":" hoặc dòng kế tiếp); nhãn so khớp không dấu
  function ocrSau(dong, re, nDong) {
    for (var i = 0; i < dong.length; i++) {
      var k = ocrBoDau(dong[i]), m = re.exec(k); if (!m) continue;
      var raw = dong[i], cut = raw.indexOf(':') >= 0 ? raw.slice(raw.lastIndexOf(':') + 1) : raw.slice(m.index + m[0].length);
      var v = cut.replace(/^[\s\/:|,.\-]+/, '').trim();
      var nhanLa = function (x) { return /[:\/]/.test(x) || /(co gia tri|date of|quoc tich|dac diem|ngay,? thang)/.test(ocrBoDau(x)); };
      if (v.length < 2) { v = []; for (var j = 1; j <= (nDong || 1) && dong[i + j] !== undefined && !nhanLa(dong[i + j]); j++) v.push(dong[i + j].trim()); v = v.join(', '); }
      else if ((nDong || 1) > 1 && /,\s*$/.test(v) && dong[i + 1] && !nhanLa(dong[i + 1])) v = v.replace(/,\s*$/, '') + ', ' + dong[i + 1].trim();   // địa chỉ xuống dòng
      return v.replace(/\s{2,}/g, ' ').replace(/^(place of (origin|residence)|date of birth|full name|sex|nationality)\s*/i, '').trim();
    }
    return '';
  }
  function ocrTach(text, loai) {
    var t = String(text || ''), dong = t.split(/\r?\n/).map(function (x) { return x.trim(); }).filter(Boolean), kd = ocrBoDau(t), g = {}, tl = {}, nhan = [];
    var laCCCD = /can cuoc|citizen identity|identity card|cmnd|chung minh nhan dan/.test(kd) || loai === 'Căn cước công dân';
    if (laCCCD) {
      nhan.push('Căn cước công dân');
      var so = /(?:^|\D)(\d{12})(?:\D|$)/.exec(t.replace(/(\d)\s(?=\d)/g, '$1')); if (so) { g.cccd = so[1]; tl.soHieu = so[1]; }
      var ns = ocrSau(dong, /(ngay,? thang,? nam sinh|date of birth)/); if (ns) g.ngaySinh = ocrNgay(ns);
      var gt = /gioi tinh[^a-z]*(nam|nu)\b/.exec(kd) || /sex[^a-z]*(nam|nu|male|female)\b/.exec(kd); if (gt) g.gioiTinh = /^(nu|female)$/.test(gt[1]) ? 'Nữ' : 'Nam';
      var qq = ocrSau(dong, /(que quan|place of origin|noi dang ky khai sinh|place of birth)/, 1); if (qq) g.queQuan = qq.slice(0, 250);
      var tr = ocrSau(dong, /(noi thuong tru|place of residence|noi cu tru)/, 2); if (tr) g.diaChi = tr.slice(0, 250);
      var hh = ocrSau(dong, /(co gia tri den|date of expiry|gia tri den)/); if (hh) tl.hetHan = ocrNgay(hh);
      var nc = ocrSau(dong, /(ngay,? thang,? nam\s*\/?\s*date|date,? month,? year|ngay cap)/); if (nc) { g.ngayCapCCCD = ocrNgay(nc); tl.ngay = g.ngayCapCCCD; }
      if (/cuc truong cuc canh sat quan ly hanh chinh ve trat tu xa hoi|cuc canh sat qlhc/.test(kd)) g.noiCapCCCD = 'Cục Cảnh sát quản lý hành chính về trật tự xã hội';
      else if (/bo cong an|ministry of public security/.test(kd)) g.noiCapCCCD = 'Bộ Công an';
      var ht = ocrSau(dong, /(ho,? chu dem va ten|ho va ten|full name)/); if (ht) g.hoTen = ht.replace(/[^A-Za-zÀ-ỹĐđ\s]/g, '').trim().slice(0, 120);
    }
    var mst = /(?:ma so thue|mst|tax code|tin)[^0-9]{0,25}(\d{10}(?:\s*-\s*\d{3})?)/.exec(kd);
    if (mst) { g.mst = mst[1].replace(/\s/g, ''); if (!laCCCD) nhan.push('Mã số thuế'); }
    var laBang = /bang tot nghiep|bang cu nhan|bang ky su|bang thac si|bang tien si|degree|diploma|chung chi/.test(kd) || /Bằng cấp|Chứng chỉ/.test(loai || '');
    if (laBang) {
      nhan.push('Bằng cấp / chứng chỉ');
      var td = [['tien si', 'Tiến sĩ'], ['thac si', 'Thạc sĩ'], ['ky su', 'Kỹ sư'], ['kien truc su', 'Kiến trúc sư'], ['cu nhan', 'Cử nhân'], ['cao dang', 'Cao đẳng'], ['trung cap', 'Trung cấp']].filter(function (x) { return kd.indexOf(x[0]) >= 0; })[0];
      if (td) g.trinhDo = td[1];
      var cn = ocrSau(dong, /(nganh dao tao|chuyen nganh|nganh|major|field of study)/); if (cn) g.chuyenNganh = cn.replace(/^(dao tao|hoc)\s*/i, '').slice(0, 150);
      var sh = ocrSau(dong, /(so hieu|so vao so|reg\.? no|so:)/); if (sh) tl.soHieu = sh.split(/\s/)[0].slice(0, 60);
    }
    if (!tl.soHieu) { var so2 = /\bs[oố]\s*[:.]?\s*([0-9A-Za-zĐđ][0-9A-Za-zĐđ\/\-.]{1,40})/.exec(t); if (so2 && /\d/.test(so2[1])) tl.soHieu = so2[1].replace(/[.,]$/, ''); }
    if (!tl.ngay) { var nk = /ng[aà]y\s*\d{1,2}\s*th[aá]ng\s*\d{1,2}\s*n[aă]m\s*\d{4}/i.exec(t); if (nk) tl.ngay = ocrNgay(nk[0]); }
    Object.keys(g).forEach(function (k) { if (!g[k]) delete g[k]; }); Object.keys(tl).forEach(function (k) { if (!tl[k]) delete tl[k]; });
    return { nv: g, tl: tl, nhan: nhan };
  }

  /* ---------------- ĐỘNG CƠ TÍNH LƯƠNG ----------------
     Thu nhập theo các khoản công thức do Quản trị lập; BH, thuế TNCN, đoàn phí, vi phạm, bồi thường, tạm ứng theo luật.
     nv: hồ sơ; cc: dòng chấm công; cfg: cấu hình; o: { luong, luongBH, ncc, thuong, thuongChiuThue, ky } */
  /* v3.29: lớp NHÂN SỰ HƯỞNG LƯƠNG HƯU – nghĩa vụ BH theo kỳ, khoản trả thêm tương đương BH (phương án A/B) */
  function tinhPhieuLuong(nv, cc, cfg, o, bienRa) {
    if (o && o.ky) cfg = cfgTheoKy(cfg, o.ky);
    var hk = huuTaiKy(nv, o && o.ky, cfg); if (!hk) return tinhPhieuLuongGoc(nv, cc, cfg, o, bienRa);
    var nv2 = nv, o2 = Object.assign({}, o || {}), tb = null;
    if (!hk.mien) nv2 = Object.assign({}, nv, { loaiLuong: 'bh', lyDoKhongBH: '', thamGiaBH: 'Có' });   // chưa thực hưởng / trước ngày hưởng → như HĐLĐ đóng BH
    if (hk.mien && hk.phuongAn && !o2.khongBH) { tb = traThemBH(o2, cfg, hk.phuongAn, hk.canCu); o2.traThayBH = tb.tong; if (hk.phuongAn === 'B') o2.luong = tb.luongCV; }
    else o2.traThayBH = 0;
    var p = tinhPhieuLuongGoc(nv2, cc, cfg, o2, bienRa);
    p.traThayBHCT = tb ? JSON.stringify(tb) : '';
    p.ghiChu = 'Hưu trí: ' + hk.lyDo + (tb ? ' Trả thêm tương đương BH (Điều 168.3 BLLĐ, phương án ' + tb.phuongAn + '): ' + tb.tong.toLocaleString('vi-VN') + ' đ, chịu thuế TNCN.' : hk.mien ? ' Chưa xác định phương án trả thêm tương đương BH.' : '') + ' ' + p.ghiChu;
    return p;
  }
  function tinhPhieuLuongGoc(nv, cc, cfg, o, bienRa) {
    if (o && o.ky) cfg = cfgTheoKy(cfg, o.ky);         // v3.25: tham số theo quy định của thời kỳ
    var bien = taoBien(nv, cc, cfg, o), L = bien.LUONG_CHINH, ncc = bien.CONG_CHUAN, kl = bien.NGHI_KL, hienCC = ncc, hienHL = bien.CONG_HL;
    var chiTiet = [], sums = { luong: 0, phucap: 0, tangca: 0, thuong: 0, khac: 0 };
    var tongTN = 0, tnct = 0, coTheGiam = 0, coTheGiamThue = 0, ktCT = 0, anCaMien = 0;
    var loaiL = loaiLuongNV(nv), lyDo = String(nv.lyDoKhongBH || '');
    (loaiL === 'khoan' ? CT_KHOAN : (cfg.congThucLuong || CT_LUONG_MD)).forEach(function (c) {
      if (c.an) return;
      var v;
      try { v = tinhCT(c.congThuc, bien); } catch (e) { throw new Error('Công thức "' + c.ten + '": ' + e.message); }
      if (!isFinite(v)) v = 0;
      if (c.nhom === 'trunggian') {                      // giá trị trung gian: không làm tròn, không cộng vào thu nhập
        bien[c.ma] = v;
        if (c.vaiTro === 'congchuan') hienCC = Math.round(v * 100) / 100;
        if (c.vaiTro === 'conghl') hienHL = Math.round(v * 100) / 100;
        return;
      }
      v = R(v);
      bien[c.ma] = v;
      if (c.nhom === 'khautru') { ktCT += v; if (v) chiTiet.push({ ma: c.ma, ten: c.ten, nhom: 'khautru', gt: v }); return; }
      var tax = v;
      if (c.thue === 'khong') tax = 0;
      else if (c.thue === 'mientoida') { var mt = mienAnCa(c, cfg, bien, v); if (c.ma === 'ANCA') anCaMien += Math.min(v, mt); tax = Math.max(0, v - mt); }   // v3.27 / v3.32: mức theo thời kỳ, theo ngày công thực tế
      else if (c.thue === 'congthuc') {
        var b2 = Object.assign({}, bien, { GT: v });
        try { tax = R(tinhCT(c.thueGT, b2)); } catch (e) { throw new Error('Phần chịu thuế "' + c.ten + '": ' + e.message); }
        tax = Math.max(0, Math.min(v, tax));
      }
      tongTN += v; tnct += tax; sums[c.phanLoai in sums ? c.phanLoai : 'khac'] += v;
      if (c.giamTruQC) { coTheGiam += v; coTheGiamThue += tax; }
      if (v) chiTiet.push({ ma: c.ma, ten: c.ten, nhom: 'thunhap', phanLoai: c.phanLoai, gt: v, thue: tax });
    });
    if (o && num(o.traThayBH) > 0 && loaiL !== 'khoan' && !(cfg.congThucLuong || CT_LUONG_MD).some(function (c) { return !c.an && /TRA_THAY_BH/.test(String(c.congThuc)); })) {   // v3.29: khoản trả thêm luôn hiện độc lập, chịu thuế
      var vT = R(num(o.traThayBH)); tongTN += vT; tnct += vT; sums.phucap += vT;
      chiTiet.push({ ma: 'TRATHAYBH', ten: 'Tiền trả thêm tương đương bảo hiểm (Điều 168.3 BLLĐ)', nhom: 'thunhap', phanLoai: 'phucap', gt: vT, thue: vT });
    }

    // vi phạm: giảm thưởng/PC theo quy chế (giảm thu nhập, tối đa bằng các khoản được phép giảm trừ) | bồi thường thiệt hại (khấu trừ sau thuế)
    var vp = parseJSON(cc.chiTietPhat, []), gtYC = 0, btYC = 0;
    (Array.isArray(vp) ? vp : []).forEach(function (x) {
      var t = num(x.thanhTien) || num(x.mucTien) * (num(x.soLan) || 1);
      if (x.loai === 'Bồi thường thiệt hại') btYC += t; else gtYC += t;
    });
    var giamThuong = Math.min(gtYC, Math.max(0, coTheGiam));
    var tong = tongTN - giamThuong;
    tnct = Math.max(0, tnct - Math.min(giamThuong, coTheGiamThue));

    if (bienRa) Object.assign(bienRa, bien);
    var duNgay = (kl + bien.NGHI_OM) < (num(cfg.nghiKLToiDaBH) || 14);
    var noiKhac = loaiL === 'khongbh' && lyDo === 'Đóng BH nơi khác', huu = loaiL === 'khongbh' && lyDo === 'Đã nghỉ hưu' && num((cfg.tyLeDN || {}).tnldHuu) > 0;
    var boBH = !!(o && o.khongBH);                   // v3.33: giai đoạn HĐ trước trong kỳ – BH tính theo HĐ cuối kỳ
    var thamGia = loaiL === 'bh' && duNgay && !boBH;
    var lbh = 0;
    if (thamGia || ((noiKhac || huu) && duNgay && !boBH)) { try { lbh = Math.max(0, R(tinhCT(cfg.congThucDongBH || CT_DONGBH_MD, bien))); } catch (e) { throw new Error('Công thức lương đóng BH: ' + e.message); } }
    var tranXH = num(cfg.luongCoSo) * num(cfg.heSoTranBHXH);
    var tranTN = num(cfg.luongToiThieuVung) * num(cfg.heSoTranBHTN);
    var bXH = Math.min(lbh, tranXH || lbh), bTN = Math.min(lbh, tranTN || lbh);
    var nld = cfg.tyLeNLD, dn = cfg.tyLeDN;
    if (!thamGia) { nld = { bhxh: 0, bhyt: 0, bhtn: 0 }; dn = { bhxh: noiKhac ? (dn.tnld === undefined ? 0.5 : num(dn.tnld)) : huu ? num(dn.tnldHuu) : 0, bhyt: 0, bhtn: 0, kpcd: 0 }; }
    var bhxh = R(bXH * num(nld.bhxh) / 100), bhyt = R(bXH * num(nld.bhyt) / 100), bhtn = R(bTN * num(nld.bhtn) / 100);
    var doanPhi = (nv.doanVien === 'Có' && lbh > 0) ? R(Math.min(lbh * num(cfg.doanPhi) / 100, num(cfg.luongCoSo) * 0.1)) : 0;
    var bhNLD = bhxh + bhyt + bhtn;

    var giamTru = 0, tntt = 0, thue = 0, tt = ttThue(nv, o && o.ky);      // v3.26: trạng thái tính thuế của kỳ
    if (tt.cach === 'kinhdoanh') { tntt = tnct; thue = 0; }      // v3.27: cá nhân kinh doanh độc lập – không khấu trừ thuế tiền công
    else if (tt.cach === 'khautru10') {
      tntt = tnct;
      var ckApDung = tt.camKet === 'co' && !(o && o.ckHopLe === false);    // v3.28: cam kết chưa duyệt / thiếu MST → tiếp tục khấu trừ
      thue = (ckApDung || tnct < num(cfg.nguongVangLai)) ? 0 : R(tnct * num(cfg.thueVangLai) / 100);
    } else {
      giamTru = tt.giamTru === 'co' ? num(cfg.giamTruBanThan) + num(nv.soNPT) * num(cfg.giamTruNPT) : 0;
      tntt = Math.max(0, tnct - bhNLD - giamTru);
      thue = thueLuyTien(tntt, cfg.bieuThue || DEFAULT_CFG.bieuThue);
    }
    // v3.32: công nợ = tạm ứng + ăn ca (tiền ăn thu theo bảng chấm ăn ca) + thu ứng + thu khấu trừ khác
    var tamUng = num(cc.tamUng), ktk = num(cc.khauTruKhac), thuUng = num(cc.thuUng);
    var coBangAn = !!(parseJSON(cc.chiTietAn, null) && Object.keys(parseJSON(cc.chiTietAn, {})).length);
    var anThu = coBangAn && !num(cc.anThuTay) ? R(num(bien.SO_BUA) * num(bien.GIA_BUA)) : num(cc.tienAnThu);   // v3.34: Công nợ sửa tay số ăn ca → dùng số đã sửa
    var tongNo = tamUng + anThu + thuUng + ktk;
    var btMax = Math.max(0, R((tong - bhNLD - thue) * (num(cfg.tyLeBoiThuongToiDa) || 30) / 100));
    var boiThuong = Math.min(btYC, btMax);
    var thucLinh = tong - bhNLD - doanPhi - thue - boiThuong - tongNo - ktCT;
    // v3.32: tách bảng lương tăng ca – thuế TNCN phân bổ cho tiền làm thêm giờ = thuế cả tháng − thuế khi không có phần chịu thuế của tiền làm thêm
    var otGT = 0, otTax = 0;
    chiTiet.forEach(function (x) { if (x.nhom === 'thunhap' && x.phanLoai === 'tangca') { otGT += num(x.gt); otTax += num(x.thue); } });
    var thue2 = thue;
    if (otTax > 0) {
      var tnct2 = Math.max(0, tnct - otTax);
      if (tt.cach === 'kinhdoanh') thue2 = 0;
      else if (tt.cach === 'khautru10') thue2 = ((tt.camKet === 'co' && !(o && o.ckHopLe === false)) || tnct2 < num(cfg.nguongVangLai)) ? 0 : R(tnct2 * num(cfg.thueVangLai) / 100);
      else thue2 = thueLuyTien(Math.max(0, tnct2 - bhNLD - giamTru), cfg.bieuThue || DEFAULT_CFG.bieuThue);
    }
    var otThue = Math.max(0, Math.min(otGT, thue - thue2));

    return {
      luongChinh: L, ngayCongChuan: hienCC, ngayCongHL: hienHL, luongThoiGian: sums.luong,
      phuCap: sums.phucap, phuCapChiTiet: '{}', chiTiet: JSON.stringify(chiTiet), tienOT: sums.tangca, thuong: sums.thuong, truyLinh: sums.khac,
      viPham: gtYC + btYC, giamThuong: giamThuong, boiThuong: boiThuong, ktCongThuc: ktCT,
      tongThuNhap: tong, luongDongBH: lbh, bhxh: bhxh, bhyt: bhyt, bhtn: bhtn, doanPhi: doanPhi,
      tnChiuThue: tnct, giamTru: giamTru, tnTinhThue: tntt, thueTNCN: thue, tamUng: tamUng, khauTruKhac: ktk, cheDoLuat: cfg.cheDoLuat || '', cachThue: tt.moTa,
      thucLinh: thucLinh, soBua: num(bien.SO_BUA), tienAnThu: anThu, thuUng: thuUng, tongNo: tongNo, anCaMien: anCaMien, otThue: otThue, otThucLinh: otGT - otThue,
      bhxhDN: R(bXH * num(dn.bhxh) / 100), bhytDN: R(bXH * num(dn.bhyt) / 100),
      bhtnDN: R(bTN * num(dn.bhtn) / 100), kpcdDN: R(bXH * num(dn.kpcd) / 100),
      ghiChu: (loaiL === 'khoan' ? 'Lương khoán việc: không đóng BH; khấu trừ ' + cfg.thueVangLai + '% khi trả từ ' + num(cfg.nguongVangLai).toLocaleString('vi-VN') + ' đ' + (nv.camKet08 === 'Có' ? ' (có cam kết 08: không khấu trừ)' : '') + '.' + (!(num(cc.tienKhoan) > 0) ? ' Chưa nhập tiền khoán kỳ này: tính theo mức khoán tháng trong hồ sơ.' : '') + (['Không xác định thời hạn', 'Xác định thời hạn'].indexOf(nv.loaiHD) >= 0 ? ' Lưu ý: hồ sơ ghi HĐLĐ ' + nv.loaiHD + ' – HĐLĐ từ 1 tháng thuộc diện BHXH bắt buộc, kiểm tra lại loại lương.' : '')
          : tt.cach === 'khautru10' ? 'Khấu trừ ' + cfg.thueVangLai + '% khi trả từ ' + num(cfg.nguongVangLai).toLocaleString('vi-VN') + ' đ (HĐ ' + nv.loaiHD + ')' + (tt.camKet === 'co' ? ' – có cam kết 08: không khấu trừ' + (!nv.mst ? ' (CHƯA CÓ MST – cam kết chỉ hợp lệ khi đã có mã số thuế)' : '') : '') + '.' : '') +
        (tt.cach === 'luytien' && tt.giamTru !== 'co' ? ' Không tính giảm trừ gia cảnh tại Công ty (đăng ký ở nơi khác).' : '') +
        (loaiL === 'khongbh' ? (lyDo === 'Đóng BH nơi khác' ? ' Đóng BH nơi khác: NLĐ không đóng; DN chỉ đóng BH TNLĐ-BNN ' + num(dn.bhxh) + '% (cột BHXH DN).' : lyDo === 'Đã nghỉ hưu' ? ' Đã nghỉ hưu: không đóng BHXH, BHYT, BHTN' + (huu ? '; DN đóng BH TNLĐ-BNN ' + num(dn.bhxh) + '% (cột BHXH DN)' : '') + '.' : lyDo === 'Hợp đồng thử việc' ? ' HĐ thử việc: không đóng BH.' : ' Không đóng BH tại Công ty (chưa ghi lý do trong hồ sơ).') : '') +
        (loaiL === 'bh' && !thamGia ? ' Nghỉ KL + ốm ≥ ' + cfg.nghiKLToiDaBH + ' ngày: không đóng BH.' : '') +
        (gtYC > giamThuong ? ' Giảm trừ quy chế vượt các khoản được phép giảm, chỉ áp ' + giamThuong.toLocaleString('vi-VN') + '.' : '') +
        (btYC > boiThuong ? ' Bồi thường vượt ' + (num(cfg.tyLeBoiThuongToiDa) || 30) + '% lương thực trả, chỉ khấu trừ ' + boiThuong.toLocaleString('vi-VN') + '; phần còn lại cần ghi khấu trừ ở kỳ sau.' : '')
    };
  }

  /* bộ nhớ tin nhắn trong trình duyệt (bản dùng thử, kiểm thử) – cùng giao diện với env.chat trên Apps Script */
  function memChat(arr, mem) {
    mem = mem || {};
    function last() { return arr.length + 1; }
    function at(q) { var r = arr[q - 2]; if (!r) return null; var o = { seq: q }; Object.keys(r).forEach(function (k) { o[k] = r[k]; }); return o; }
    return {
      append: function (row) { var o = {}; SCHEMA.TinNhan.forEach(function (k) { o[k] = row[k] == null ? '' : String(row[k]); }); arr.push(o); return arr.length + 1; },
      last: last, get: at,
      since: function (q, max) { var out = []; for (var i = Math.max(2, q + 1); i <= last() && out.length < max; i++) out.push(at(i)); return out; },
      before: function (q, n) { var to = Math.min(q - 1, last()), out = []; for (var i = Math.max(2, to - n + 1); i <= to; i++) out.push(at(i)); return out; },
      bump: function (us) { var v = String(Date.now()) + Math.random().toString(36).slice(2, 6); mem.sig = mem.sig || {}; us.forEach(function (u) { mem.sig[u] = v; }); },
      sig: function (u) { return (mem.sig || {})[u] || '0'; },
      online: function (u) { (mem.on = mem.on || {})[u] = Date.now(); },
      onlineOf: function (us) { return us.filter(function (u) { return mem.on && Date.now() - (mem.on[u] || 0) < 75000; }); },
      typing: function (ht, u, n) { mem.ty = mem.ty || {}; (mem.ty[ht] = mem.ty[ht] || {})[u] = { n: n, t: Date.now() }; },
      typingOf: function (hts, me) {
        var out = {};
        hts.forEach(function (h) { var m = (mem.ty || {})[h] || {}, a = Object.keys(m).filter(function (u) { return u !== me && Date.now() - m[u].t < 6000; }).map(function (u) { return m[u].n; }); if (a.length) out[h] = a; });
        return out;
      }
    };
  }

  /* ===== MỨC HƯỞNG THEO THỜI GIAN =====
     Mỗi CBCNV có diễn biến các mức: lương thời gian (thỏa thuận), lương cơ bản (đóng BH), từng khoản phụ cấp, số người phụ thuộc, mức ăn ca.
     Nguồn: khai báo mức hưởng (MucHuong, hiệu lực) + quyết định lên lương đã duyệt (LenLuong). Khoản không khai giữ nguyên mức trước đó. */
  var MH_KHOA = { luongChinh: 'Lương thời gian (thỏa thuận)', luongBH: 'Lương cơ bản (đóng BH)', soNPT: 'Số người phụ thuộc', anCa: 'Mức ăn ca riêng (đ/ngày)' };
  var MH_LUONG = ['luongChinh', 'luongBH'];
  function mhNguon(nv, llRows, mhRows, boQua) {
    var out = hdSuKien(nv);                          // v3.33: lương HĐ, lương đóng BH theo lịch sử HĐLĐ
    (llRows || []).forEach(function (r) {
      if (r.maNV !== nv.maNV || r.trangThai !== 'Đã duyệt' || !r.hieuLuc || r.id === boQua) return;
      var gt = {}, cu = {};
      if (num(r.luongMoi) > 0) gt.luongChinh = num(r.luongMoi);
      if (num(r.luongBHMoi) > 0) gt.luongBH = num(r.luongBHMoi);
      if (num(r.luongCu) > 0) cu.luongChinh = num(r.luongCu);
      if (num(r.luongBHCu) > 0) cu.luongBH = num(r.luongBHCu);
      out.push({ id: r.id, nguon: 'll', hieuLuc: r.hieuLuc, t: String(r.ngayDuyet || ''), gt: gt, cu: cu });
    });
    (mhRows || []).forEach(function (r) {
      if (r.maNV !== nv.maNV || r.trangThai !== 'Hiệu lực' || !r.hieuLuc || r.id === boQua) return;
      out.push({ id: r.id, nguon: 'mh', hieuLuc: r.hieuLuc, t: String(r.ngayDuyet || r.taoLuc || ''), gt: parseJSON(r.giaTri, {}), cu: parseJSON(r.cu, {}) });
    });
    out.sort(function (a, b) { return a.hieuLuc < b.hieuLuc ? -1 : a.hieuLuc > b.hieuLuc ? 1 : (a.t < b.t ? -1 : a.t > b.t ? 1 : 0); });
    return out;
  }
  function mucHuong(nv, llRows, mhRows, ngay, boQua) {
    var base = { luongChinh: num(nv.luongChinh), luongBH: num(nv.luongBH), soNPT: num(nv.soNPT), anCa: num(nv.anCa) };
    var pc = parseJSON(nv.phuCap, {});
    Object.keys(pc).forEach(function (k) { base['pc.' + k] = num(pc[k]); });
    var ds = mhNguon(nv, llRows, mhRows, boQua), keys = Object.keys(base);
    ds.forEach(function (x) { Object.keys(x.gt).forEach(function (k) { if (keys.indexOf(k) < 0) keys.push(k); }); });
    var out = { luongChinh: 0, luongBH: 0, soNPT: 0, anCa: 0, phuCap: {} };
    keys.forEach(function (k) {
      var truoc = null, sau = null;
      ds.forEach(function (x) { if (!(k in x.gt)) return; if (x.hieuLuc <= ngay) truoc = x; else if (!sau) sau = x; });
      var v = truoc ? num(truoc.gt[k]) : (sau && (k in sau.cu) ? num(sau.cu[k]) : num(base[k]));
      if (k.indexOf('pc.') === 0) { if (v) out.phuCap[k.slice(3)] = v; } else out[k] = v;
    });
    return out;
  }
  /* hồ sơ "tại kỳ": thay các mức hưởng bằng mức có hiệu lực cuối kỳ */
  function nvTaiKy(nv, llRows, mhRows, ky) {
    var m = mucHuong(nv, llRows, mhRows, kyEnd(ky));
    return { nv: Object.assign({}, nv, { luongChinh: m.luongChinh, luongBH: m.luongBH, soNPT: m.soNPT, anCa: m.anCa, phuCap: JSON.stringify(m.phuCap) }), luong: m.luongChinh, luongBH: m.luongBH, muc: m };
  }
  function luongTaiKy(nv, llRows, ky) { var m = mucHuong(nv, llRows, [], kyEnd(ky)); return { luong: m.luongChinh, luongBH: m.luongBH }; }

  /* hạn nâng bậc kế tiếp */
  function hanNangBac(nv, llRows, ktRows, cfg) {
    var moc = nv.ngayHuongBac || nv.ngayVao || '';
    llRows.forEach(function (r) {
      if (r.maNV === nv.maNV && r.trangThai === 'Đã duyệt' && (r.loai === 'Nâng bậc định kỳ' || r.loai === 'Nâng lương trước hạn') && r.hieuLuc > moc) moc = r.hieuLuc;
    });
    if (!moc) return '';
    var han = addMonths(moc, num(cfg.chuKyNangBac) || 24);
    ktRows.forEach(function (k) {
      if (k.maNV === nv.maNV && k.loai === 'Kỷ luật' && k.trangThai === 'Đã duyệt' && /Kéo dài/.test(k.hinhThuc || '') && (k.ngay || '') >= moc) han = addMonths(han, 6);
    });
    return han;
  }

  /* v3.18: ngày được chấm công: từ ngày vào làm, trước ngày chấm dứt HĐ (ngayNghi = ngày chấm dứt, kể từ ngày đó không còn làm việc) */
  function ngoaiHD(nv, ngay) { return !!((nv.ngayVao && ngay < nv.ngayVao) || (nv.ngayNghi && ngay >= nv.ngayNghi)); }
  function dangLamTrongKy(nv, ky) {
    var s = kyStart(ky), e = kyEnd(ky);
    if (nv.ngayVao && nv.ngayVao > e) return false;
    if (nv.ngayNghi && nv.ngayNghi <= s) return false;           // v3.18: chấm dứt từ ngày đầu kỳ trở về trước → không còn trong kỳ
    if (nv.trangThai === 'Đã nghỉ việc' && !nv.ngayNghi) return false;
    return true;
  }


  /* ==================== BÁO CÁO & XUẤT EXCEL CÓ CÔNG THỨC ==================== */
  /* v3.12: nhận diện bộ công thức còn tính thuế toàn bộ / theo quy định cũ đối với tiền làm thêm giờ, và vá:
     (1) mẫu TAPACO cũ: giờ làm thêm (thử việc) gộp vào công tính lương → tách thành khoản TANGCA, tổng không đổi;
     (2) khoản phân loại "tăng ca" có dùng OT150/OT200/OT300 mà chịu thuế toàn bộ hoặc theo công thức cũ → miễn theo OT_TL_MIEN. */
  var OT_GT_MOI = 'GT * (1 - OT_TL_MIEN)', OT_GT_CU = 'MIN(GT, LUONG_GIO * (OT150 + OT200 + OT300))';
  var OT_QD_CU = 'NGAY_CONG + ROUND((OT200 * 2 + OT300 * 4 + OT150 * 1.5) / GIO_CONG_NGAY, 2)';
  function otVaCT(ds) {
    ds = clone(ds || []); var doi = [], ma = {}; ds.forEach(function (c) { ma[c.ma] = c; });
    var ctl = ma.CONG_TL, tg = ma.TONG_TG;
    if (ctl && ctl.nhom === 'trunggian' && String(ctl.congThuc).indexOf(OT_QD_CU) >= 0 && tg && !ma.TANGCA && !ma.OT_QD && !ma.TONG_OT) {
      if (String(tg.congThuc).replace(/\s+/g, '') !== 'ROUND(LUONG_CHINH*HS_HUONG+LUONG_BH*TL_HD/CC*NGAY_NGHI_CL,-2)' || !ma.HS_HUONG || String(ma.HS_HUONG.congThuc).replace(/\s+/g, '') !== 'TL_HD*HS_K*CONG_TL/CC')
        doi.push({ loi: 'Khoản TONG_TG / HS_HUONG đã được sửa khác mẫu – cần tách tiền làm thêm giờ bằng tay.' });
      else {
        ctl.congThuc = String(ctl.congThuc).split(OT_QD_CU).join('NGAY_CONG');
        if (ctl.ghiChu) ctl.ghiChu = 'Thử việc: công thực tế (giờ làm thêm quy đổi tách sang khoản "Lương làm thêm giờ"). Dự án: làm 26–28 công tính đủ công chuẩn.';
        var mau = {}; CT_LUONG_TAPACO.forEach(function (c) { mau[c.ma] = c; });
        var iTG = ds.indexOf(tg); ds.splice(iTG + 1, 0, clone(mau.OT_QD), clone(mau.TONG_OT));
        var iAn = ds.map(function (c) { return c.ma; }).indexOf('ANCA'), iTK = ds.map(function (c) { return c.ma; }).indexOf('THUONGKT');
        var viTri = iAn >= 0 ? iAn : iTK >= 0 ? iTK : ds.length; ds.splice(viTri, 0, clone(mau.TANGCA));
        doi.push({ ma: 'TANGCA', mo: 'Tách tiền làm thêm giờ (thử việc) khỏi công tính lương thành khoản "Lương làm thêm giờ" – tổng lương không đổi; miễn thuế phần trong giới hạn giờ.' });
      }
    }
    ds.forEach(function (c) {
      if (c.an || c.nhom !== 'thunhap' || c.phanLoai !== 'tangca' || c.ma === 'TANGCA' && doi.some(function (d) { return d.ma === 'TANGCA'; })) return;
      var dungOT = /\bOT(150|200|300)\b/.test(String(c.congThuc)) || /\bTONG_OT\b/.test(String(c.congThuc));
      if (!dungOT) return;
      var gt = String(c.thueGT || '').replace(/\s+/g, ' ').trim();
      if (c.thue === 'toanbo' || (c.thue === 'congthuc' && gt === OT_GT_CU)) {
        c.thue = 'congthuc'; c.thueGT = OT_GT_MOI;
        doi.push({ ma: c.ma, mo: '"' + c.ten + '": phần chịu thuế đổi thành GT × (1 − OT_TL_MIEN) – miễn thuế phần giờ trong giới hạn luật lao động.' });
      }
    });
    return { ds: ds, doi: doi };
  }
  /* v3.13.1: quy chế ăn ca – vá bộ công thức đang lưu: ăn ca tính theo số ngày làm đủ 8 giờ thay cho ngày công (nửa ngày không có ăn ca) */
  var ANCA_CU = 'IF(NGAY_AN>0,NGAY_AN,NGAY_CONG)*TIEN_AN_NGAY';
  function anCaVaCT(ds, bienKy) {
    ds = clone(ds || []); bienKy = clone(bienKy || []); var doi = [];
    ds.forEach(function (c) {
      if (c.an || c.nhom !== 'thunhap') return;
      var f = String(c.congThuc || '').replace(/\s+/g, '');
      if (f === ANCA_CU || f === 'NGAY_CONG*TIEN_AN_NGAY') {
        c.congThuc = f === ANCA_CU ? 'IF(NGAY_AN > 0, NGAY_AN, NGAY_DU_CA) * TIEN_AN_NGAY' : 'NGAY_DU_CA * TIEN_AN_NGAY';
        if (!c.ghiChu) c.ghiChu = 'Quy chế: chỉ ngày làm đủ 8 giờ được ăn ca; làm nửa ngày không có.';
        doi.push({ ma: c.ma, mo: '"' + c.ten + '": tính theo số ngày làm đủ 8 giờ (NGAY_DU_CA) thay cho ngày công – làm nửa ngày không có tiền ăn ca.' });
      }
    });
    if (doi.length) bienKy.forEach(function (v) { if (v.ma === 'NGAY_AN' && /trống = ngày công/.test(v.ten || '')) v.ten = 'Ngày tính ăn ca (trống = số ngày làm đủ 8 giờ)'; });
    return { ds: ds, bienKy: bienKy, doi: doi };
  }
  function otLuyKe(maNV, ky, ccAll) {             // giờ làm thêm các kỳ trước trong cùng năm (để áp giới hạn năm)
    var t = 0; (ccAll || []).forEach(function (c) { if (c.maNV === maNV && c.ky < ky && String(c.ky).slice(0, 4) === String(ky).slice(0, 4)) t += num(c.ot150) + num(c.ot200) + num(c.ot300); });
    return t;
  }
  function dauVaoTinhLuong(nv, ky, ncc, ll, mh, kt, ccAll) {
    var segs = hdDoanKy(nv, ky);                     // v3.33: HĐ có hiệu lực trong kỳ (cuối kỳ quyết định BH, loại HĐ, thuế)
    if (segs.length) nv = apHD(nv, segs[segs.length - 1].hd);
    var th = 0, thT = 0;
    (kt || []).forEach(function (x) {
      if (x.maNV === nv.maNV && x.kyLuong === ky && x.trangThai === 'Đã duyệt' && x.loai !== 'Kỷ luật') { th += num(x.soTien); if (x.chiuThue !== 'Không') thT += num(x.soTien); }
    });
    var lk = nvTaiKy(nv, ll || [], mh || [], ky);
    return { nv: lk.nv, segs: segs, o: { luong: lk.luong, luongBH: lk.luongBH, ncc: ncc, thuong: th, thuongChiuThue: thT, ky: ky, otLuyKe: otLuyKe(nv.maNV, ky, ccAll) } };
  }
  function cotXL(i) { var s = ''; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
  /* công thức của app → công thức Excel; ref(tên biến) trả về địa chỉ ô */
  function astExcel(a, ref) {
    switch (a[0]) {
      case 'n': return String(a[1]);
      case 'v': return ref(a[1]);
      case 'u': return '(-' + astExcel(a[1], ref) + ')';
      case 'b':
        var l = astExcel(a[2], ref), r = astExcel(a[3], ref), o = a[1];
        if (o === '/') return 'IF(' + r + '=0,0,' + l + '/' + r + ')';
        if (o === '=') return '(ABS(' + l + '-' + r + ')<0.000000001)*1';
        if (o === '<>') return '(ABS(' + l + '-' + r + ')>=0.000000001)*1';
        if (['>', '<', '>=', '<='].indexOf(o) >= 0) return '(' + l + o + r + ')*1';
        return '(' + l + o + r + ')';
      case 'f':
        var f = a[1], g = a[2].map(function (x) { return astExcel(x, ref); });
        if (f === 'IF') return 'IF(' + g[0] + '<>0,' + g[1] + ',' + g[2] + ')';
        if (f === 'ROUND') return 'ROUND(' + g[0] + ',' + (g[1] || '0') + ')';
        if (f === 'FLOOR') return 'INT(' + g[0] + ')';
        if (f === 'CEIL') return '(-INT(-' + g[0] + '))';
        if (f === 'AND' || f === 'OR') return f + '(' + g.map(function (x) { return x + '<>0'; }).join(',') + ')*1';
        if (f === 'NOT') return '(' + g[0] + '=0)*1';
        return f + '(' + g.join(',') + ')';
    }
    throw new Error('Biểu thức không hợp lệ');
  }
  function ctExcel(src, ref) { return astExcel(phanTichCT(src), ref); }
  var XL_TIEN = '#,##0', XL_CONG = '#,##0.0#', XL_PT = '0.0%';
  var BIEN_TS = ['GIO_CONG_NGAY', 'HS_OT150', 'HS_OT200', 'HS_OT300', 'LUONG_CO_SO', 'LUONG_TOI_THIEU_VUNG'];
  var BIEN_DAN_XUAT = { LUONG_NGAY: 'LUONG_CHINH / CONG_CHUAN', LUONG_GIO: 'LUONG_NGAY / GIO_CONG_NGAY', CONG_HL: 'NGAY_CONG + NGHI_HL + NGAY_LE' };
  var BIEN_CO_BAN = ['LUONG_CHINH', 'LUONG_BH', 'CONG_CHUAN', 'NGAY_CONG', 'NGHI_HL', 'NGAY_LE', 'NGHI_KL', 'NGHI_OM', 'OT150', 'OT200', 'OT300', 'THUONG', 'THUONG_CHIU_THUE', 'TRUY_LINH', 'SO_NPT'];
  function tenBien(cfg) {
    var m = {}; BIEN_HT.forEach(function (b) { m[b[0]] = b[1].split(' (')[0]; });
    (cfg.phuCap || []).forEach(function (p) { m['PC_' + String(p.ma).toUpperCase()] = p.ten; });
    (cfg.bienKy || []).forEach(function (v) { m[v.ma] = v.ten; });
    (cfg.noiLamViec || []).forEach(function (x) { m['NLV_' + String(x.ma).toUpperCase()] = 'Làm tại ' + x.ten + ' (1/0)'; });
    Object.assign(m, { LUONG_CHINH: 'Lương chính', LUONG_BH: 'Lương đóng BH (hồ sơ)', CONG_CHUAN: 'Công chuẩn', NGAY_CONG: 'Ngày công', NGHI_HL: 'Nghỉ hưởng lương', NGAY_LE: 'Ngày lễ', NGHI_KL: 'Nghỉ không lương', NGHI_OM: 'Nghỉ ốm', OT150: 'Giờ tăng ca', OT200: 'Giờ làm CN', OT300: 'Giờ làm lễ', THUONG: 'Thưởng', THUONG_CHIU_THUE: 'Thưởng chịu thuế', TRUY_LINH: 'Truy lĩnh', SO_NPT: 'Số NPT', LUONG_NGAY: 'Lương ngày', LUONG_GIO: 'Lương giờ', CONG_HL: 'Công hưởng lương' });
    return m;
  }
  /* Sheet tham số: tỷ lệ BH, trần, giảm trừ, biểu thuế – các bảng tính tham chiếu tuyệt đối */
  function sheetThamSo(cfg) {
    var rows = [[{ v: 'THAM SỐ TÍNH LƯƠNG', s: 'title' }], [{ v: 'Sửa ô giá trị (cột C) thì các bảng tính tự tính lại.', s: 'note' }], [], [{ v: 'Tham số', s: 'head' }, { v: 'Mã', s: 'head' }, { v: 'Giá trị', s: 'head' }]], ref = {};
    function add(ten, ma, v, z) { rows.push([ten, { v: ma, s: 'code' }, typeof v === 'object' ? v : { v: num(v), z: z || XL_TIEN }]); ref[ma] = 'ThamSo!$C$' + rows.length; }
    var h = cfg.heSoOT || {}, nld = cfg.tyLeNLD || {}, dn = cfg.tyLeDN || {};
    add('Giờ làm mỗi ngày', 'GIO_CONG_NGAY', num(cfg.gioCongNgay) || 8, '0.##');
    add('Hệ số OT ngày thường', 'HS_OT150', num(h.ot150) || 1.5, '0.##'); add('Hệ số OT ngày nghỉ tuần', 'HS_OT200', num(h.ot200) || 2, '0.##'); add('Hệ số OT ngày lễ', 'HS_OT300', num(h.ot300) || 3, '0.##');
    add('Lương cơ sở', 'LUONG_CO_SO', cfg.luongCoSo); add('Lương tối thiểu vùng', 'LUONG_TOI_THIEU_VUNG', cfg.luongToiThieuVung);
    add('Hệ số trần BHXH, BHYT (× lương cơ sở)', 'HS_TRAN_BHXH', cfg.heSoTranBHXH, '0.##'); add('Hệ số trần BHTN (× lương tối thiểu vùng)', 'HS_TRAN_BHTN', cfg.heSoTranBHTN, '0.##');
    add('Trần lương đóng BHXH, BHYT', 'TRAN_BHXH', { f: ref.LUONG_CO_SO + '*' + ref.HS_TRAN_BHXH, v: num(cfg.luongCoSo) * num(cfg.heSoTranBHXH), z: XL_TIEN });
    add('Trần lương đóng BHTN', 'TRAN_BHTN', { f: ref.LUONG_TOI_THIEU_VUNG + '*' + ref.HS_TRAN_BHTN, v: num(cfg.luongToiThieuVung) * num(cfg.heSoTranBHTN), z: XL_TIEN });
    add('BHXH người lao động (%)', 'NLD_BHXH', nld.bhxh, '0.0#'); add('BHYT người lao động (%)', 'NLD_BHYT', nld.bhyt, '0.0#'); add('BHTN người lao động (%)', 'NLD_BHTN', nld.bhtn, '0.0#');
    add('BHXH doanh nghiệp (%)', 'DN_BHXH', dn.bhxh, '0.0#'); add('BHYT doanh nghiệp (%)', 'DN_BHYT', dn.bhyt, '0.0#'); add('BHTN doanh nghiệp (%)', 'DN_BHTN', dn.bhtn, '0.0#'); add('Kinh phí công đoàn (%)', 'DN_KPCD', dn.kpcd, '0.0#');
    add('Đoàn phí công đoàn (%)', 'DOAN_PHI', cfg.doanPhi, '0.0#');
    add('Nghỉ KL + ốm từ … ngày thì không đóng BH', 'NGUONG_NGHI_BH', num(cfg.nghiKLToiDaBH) || 14, '0');
    add('Giảm trừ bản thân (đ/tháng)', 'GT_BAN_THAN', cfg.giamTruBanThan); add('Giảm trừ mỗi người phụ thuộc', 'GT_NPT', cfg.giamTruNPT);
    add('Khấu trừ thuế HĐ dưới 3 tháng / thử việc (%)', 'THUE_VL', cfg.thueVangLai, '0.##'); add('Ngưỡng khấu trừ (đ/lần)', 'NGUONG_VL', cfg.nguongVangLai);
    add('Bồi thường tối đa (% lương thực trả)', 'TL_BOI_THUONG', num(cfg.tyLeBoiThuongToiDa) || 30, '0.##');
    rows.push([], [{ v: 'Biểu thuế lũy tiến từng phần', s: 'head' }, { v: 'Đến mức (0 = trên)', s: 'head' }, { v: 'Thuế suất %', s: 'head' }]);
    var bac = [];
    (cfg.bieuThue || DEFAULT_CFG.bieuThue).forEach(function (b, i) { rows.push(['Bậc ' + (i + 1), { v: num(b[0]), z: XL_TIEN }, { v: num(b[1]), z: '0.##' }]); bac.push({ tran: 'ThamSo!$B$' + rows.length, ts: 'ThamSo!$C$' + rows.length }); });
    ref.bac = bac;
    return { sheet: { ten: 'ThamSo', rows: rows, cols: [46, 22, 18] }, ref: ref };
  }
  function bieuThueXL(x, bac) {
    return 'ROUND(' + bac.map(function (b, i) {
      return 'MAX(0,IF(' + b.tran + '=0,' + x + ',MIN(' + x + ',' + b.tran + '))-' + (i ? bac[i - 1].tran : '0') + ')*' + b.ts + '/100';
    }).join('+') + ',0)';
  }
  /* BẢNG TÍNH LƯƠNG CHI TIẾT – mọi bước đều là công thức Excel
     ds: [{ nv (hồ sơ tại kỳ), cc, o, app (dòng bảng lương đã lưu) }]; ctl/dbh: công thức của kỳ */
  function xlBangTinh(ky, ds, cfg0, ctl, dbh, info) {
    cfg0 = cfgTheoKy(cfg0, ky);                        // v3.25: sheet ThamSo theo quy định của kỳ
    var cfg = Object.assign({}, cfg0, { congThucLuong: ctl, congThucDongBH: dbh }), ts = sheetThamSo(cfg), T = ts.ref, ten = tenBien(cfg);
    var ks = ctl.filter(function (c) { return !c.an; }), maK = {}; ks.forEach(function (c) { maK[c.ma] = c; });
    var dung = [];
    ks.forEach(function (c) { bienDung(phanTichCT(c.congThuc), dung); if (c.thue === 'congthuc' && c.thueGT) bienDung(phanTichCT(String(c.thueGT)), dung); });
    bienDung(phanTichCT(dbh), dung);
    Object.keys(BIEN_DAN_XUAT).forEach(function (k) { if (dung.indexOf(k) >= 0) bienDung(phanTichCT(BIEN_DAN_XUAT[k]), dung); });
    var nhap = BIEN_CO_BAN.slice();
    dung.forEach(function (v) { if (v !== 'GT' && !maK[v] && BIEN_TS.indexOf(v) < 0 && !BIEN_DAN_XUAT[v] && nhap.indexOf(v) < 0) nhap.push(v); });
    var dx = ['LUONG_NGAY', 'LUONG_GIO', 'CONG_HL'].filter(function (k) { return k === 'CONG_HL' || dung.indexOf(k) >= 0; });
    if (dx.indexOf('LUONG_GIO') >= 0 && dx.indexOf('LUONG_NGAY') < 0) dx.unshift('LUONG_NGAY');
    // bố trí cột
    var C = [], idx = {};
    function col(key, nhom, tieuDe, ma, z, w) { idx[key] = C.length; C.push({ key: key, nhom: nhom, ten: tieuDe, ma: ma || '', z: z, w: w || 13 }); }
    col('stt', 'Thông tin', 'STT', '', '0', 5); col('ma', 'Thông tin', 'Mã NV', '', null, 9); col('ten', 'Thông tin', 'Họ và tên', '', null, 24); col('bp', 'Thông tin', 'Bộ phận', '', null, 22); col('hd', 'Thông tin', 'Loại HĐ', '', null, 14); col('ll', 'Thông tin', 'Loại lương', '', null, 16);
    nhap.forEach(function (v) { col('v:' + v, 'Dữ liệu đầu vào', ten[v] || v, v, /^(LUONG|THUONG|TRUY|PC_|TIEN)/.test(v) ? XL_TIEN : XL_CONG); });
    col('tgbh', 'Dữ liệu đầu vào', 'Tham gia BH (1/0)', 'TG_BH', '0', 9); col('dv', 'Dữ liệu đầu vào', 'Đoàn viên (1/0)', 'DOAN_VIEN', '0', 9);
    col('vl', 'Dữ liệu đầu vào', 'HĐ dưới 3 tháng / thử việc (1/0)', 'VANG_LAI', '0', 11); col('ck', 'Dữ liệu đầu vào', 'Cam kết 08 (1/0)', 'CAM_KET_08', '0', 9); col('gtc', 'Dữ liệu đầu vào', 'Giảm trừ gia cảnh tại Công ty (1/0)', 'GT_TAI_CTY', '0', 10);
    dx.forEach(function (v) { col('v:' + v, 'Giá trị trung gian', ten[v] || v, v, v === 'CONG_HL' ? XL_CONG : XL_TIEN); });
    var tg = ks.filter(function (c) { return c.nhom === 'trunggian'; }), tn = ks.filter(function (c) { return c.nhom !== 'trunggian' && c.nhom !== 'khautru'; }), kt = ks.filter(function (c) { return c.nhom === 'khautru'; });
    tg.forEach(function (c) { col('k:' + c.ma, 'Giá trị trung gian', c.ten, c.ma, '#,##0.####', 14); });
    tn.forEach(function (c) { col('k:' + c.ma, 'Thu nhập', c.ten, c.ma, XL_TIEN, 14); });
    var coKhoan = ds.some(function (d) { return loaiLuongNV(d.nv) === 'khoan'; });
    if (coKhoan) col('tk', 'Thu nhập', 'Tiền khoán việc (lương khoán)', 'KHOAN', XL_TIEN, 14);
    [['ttn', 'Cộng thu nhập', 'TONG_TN'], ['mien', 'Thu nhập miễn thuế', 'MIEN_THUE'], ['dngt', 'Đề nghị giảm thưởng/PC (vi phạm)', 'DN_GIAM'], ['ctg', 'Khoản được phép giảm', 'CO_THE_GIAM'], ['gqc', 'Giảm trừ theo quy chế', 'GIAM_QC'], ['tong', 'TỔNG THU NHẬP', 'TONG']].forEach(function (x) { col(x[0], 'Thu nhập', x[1], x[2], XL_TIEN, 15); });
    [['tgt', 'Đóng BH tháng này (1/0)', 'DONG_BH', '0'], ['lbh', 'Lương đóng BH', 'LUONG_DONG_BH'], ['bhxh', 'BHXH', 'BHXH'], ['bhyt', 'BHYT', 'BHYT'], ['bhtn', 'BHTN', 'BHTN'], ['bh', 'Cộng BH NLĐ', 'BH_NLD'], ['dp', 'Đoàn phí', 'DOAN_PHI'],
      ['tnct', 'Thu nhập chịu thuế', 'TNCT'], ['gtgc', 'Giảm trừ gia cảnh', 'GIAM_TRU'], ['tntt', 'Thu nhập tính thuế', 'TNTT'], ['thue', 'Thuế TNCN', 'THUE']].forEach(function (x) { col(x[0], 'Bảo hiểm & thuế', x[1], x[2], x[3] || XL_TIEN, 14); });
    [['dnbt', 'Đề nghị bồi thường', 'DN_BOI_THUONG'], ['btmax', 'Bồi thường tối đa', 'BT_TOI_DA'], ['bt', 'Bồi thường khấu trừ', 'BOI_THUONG'], ['tu', 'Tạm ứng', 'TAM_UNG'], ['anthu', 'Ăn ca (tiền ăn thu)', 'AN_CA_THU'], ['thuung', 'Thu ứng', 'THU_UNG'], ['ktcn', 'Thu khấu trừ khác', 'KT_CONG_NO']].forEach(function (x) { col(x[0], 'Khấu trừ', x[1], x[2], XL_TIEN, 14); });
    kt.forEach(function (c) { col('k:' + c.ma, 'Khấu trừ', c.ten, c.ma, XL_TIEN, 14); });
    col('cn', 'Còn nhận', 'CÒN NHẬN', 'CON_NHAN', XL_TIEN, 16);
    [['dnbhxh', 'BHXH (DN)'], ['dnbhyt', 'BHYT (DN)'], ['dnbhtn', 'BHTN (DN)'], ['dnkpcd', 'KPCĐ (DN)'], ['cp', 'Tổng chi phí lương']].forEach(function (x) { col(x[0], 'Doanh nghiệp đóng', x[1], '', XL_TIEN, 14); });
    col('app', 'Đối chiếu', 'Còn nhận trên app', '', XL_TIEN, 15); col('cl', 'Đối chiếu', 'Chênh lệch', '', XL_TIEN, 12);
    var HR = 6, r0 = 7;                         // dòng tiêu đề nhóm 4, tên cột 5, mã 6, dữ liệu từ 7
    var rows = [[{ v: 'BẢNG TÍNH LƯƠNG CHI TIẾT – KỲ ' + ky.slice(5) + '/' + ky.slice(0, 4), s: 'title' }],
      [{ v: (info && info.congTy || cfg.congTy || '') + (info && info.trangThai ? ' · ' + info.trangThai : ''), s: info && info.nhap ? 'draft' : 'sub' }],
      [{ v: 'Mọi cột từ "Giá trị trung gian" trở đi là công thức Excel theo đúng công thức lương của kỳ; tham số BH, thuế ở sheet ThamSo. Cột "Chênh lệch" = 0 nghĩa là khớp bảng lương trên app.', s: 'note' }]];
    var g = [], lastN = ''; C.forEach(function (c) { g.push(c.nhom !== lastN ? { v: c.nhom, s: 'grp' } : { v: '', s: 'grp' }); lastN = c.nhom; });
    rows.push(g, C.map(function (c) { return { v: c.ten, s: 'head' }; }), C.map(function (c) { return { v: c.ma, s: 'code' }; }));
    var merges = [], st = 0; for (var i = 1; i <= C.length; i++) if (i === C.length || C[i].nhom !== C[st].nhom) { if (i - st > 1) merges.push({ s: { r: 3, c: st }, e: { r: 3, c: i - 1 } }); st = i; }
    function A(key, r) { return cotXL(idx[key]) + r; }
    ds.forEach(function (d, n) {
      var r = r0 + n, bien = {}, p = tinhPhieuLuong(d.nv, d.cc, cfg, d.o, bien), nv = d.nv, cc = d.cc;
      var ref = function (v) {
        if (v === 'GT') return '0';
        if (idx['v:' + v] !== undefined) return A('v:' + v, r);
        if (idx['k:' + v] !== undefined) return A('k:' + v, r);
        if (T[v]) return T[v];
        return '0';
      };
      var row = [];
      function put(key, cell) { row[idx[key]] = cell; }
      put('stt', n + 1); put('ma', nv.maNV); put('ten', nv.hoTen); put('bp', nv.phongBan || ''); put('hd', nv.loaiHD || ''); var LL = loaiLuongNV(nv); put('ll', LL === 'bh' ? 'Đóng BH' : LL === 'khoan' ? 'Khoán việc' : 'Không đóng BH' + (nv.lyDoKhongBH ? ' – ' + (nv.lyDoKhongBH === 'Đã nghỉ hưu' ? 'nghỉ hưu' : 'đóng nơi khác') : ''));
      nhap.forEach(function (v) { put('v:' + v, { v: num(bien[v]), z: C[idx['v:' + v]].z }); });
      var ttX = ttThue(nv, ky), tgbh = LL === 'bh' ? 1 : 0, dv = nv.doanVien === 'Có' ? 1 : 0, vl = ttX.cach === 'khautru10' ? 1 : 0, ck = ttX.camKet === 'co' ? 1 : 0, gtc = ttX.giamTru === 'co' ? 1 : 0;
      put('tgbh', tgbh); put('dv', dv); put('vl', vl); put('ck', ck); put('gtc', gtc);
      dx.forEach(function (v) { put('v:' + v, { f: ctExcel(BIEN_DAN_XUAT[v], ref), v: num(bien[v]), z: C[idx['v:' + v]].z }); });
      var txSum = [], mienSum = [], ctg = [], ctgt = [], tnV = 0, mienV = 0, ctgV = 0, ctgtV = 0;
      ks.forEach(function (c) {
        if (c.nhom === 'trunggian') { put('k:' + c.ma, { f: ctExcel(c.congThuc, ref), v: num(bien[c.ma]), z: '#,##0.####' }); return; }
        var a = A('k:' + c.ma, r), v = num(bien[c.ma]);
        put('k:' + c.ma, { f: 'ROUND(' + ctExcel(c.congThuc, ref) + ',0)', v: v, z: XL_TIEN });
        if (c.nhom === 'khautru') return;
        var tx = a, txV = v;
        if (c.thue === 'khong') { tx = '0'; txV = 0; }
        else if (c.thue === 'mientoida') { var mtV = mienAnCa(c, cfg, bien, v), mtM = mucMienAnCa(c, cfg), mtF = c.ma === 'ANCA' ? 'MAX(0,MIN(' + mtM + ',ROUND(' + ctExcel(ctMienAnCa(cfg), function (x) { return x === 'MUC_MIEN' ? String(mtM) : x === 'GT' ? a : ref(x); }) + ',0)))' : String(mtV); tx = 'MAX(0,' + a + '-' + mtF + ')'; txV = Math.max(0, v - mtV); }
        else if (c.thue === 'congthuc') {
          tx = 'MAX(0,MIN(' + a + ',ROUND(' + ctExcel(c.thueGT, function (x) { return x === 'GT' ? a : ref(x); }) + ',0)))';
          var b2 = Object.assign({}, bien, { GT: v }); txV = Math.max(0, Math.min(v, R(tinhCT(c.thueGT, b2))));
        }
        tnV += v;
        if (tx !== a) { mienSum.push('(' + a + '-' + tx + ')'); mienV += v - txV; }
        if (c.giamTruQC) { ctg.push(a); ctgt.push(tx); ctgV += v; ctgtV += txV; }
      });
      var vp = parseJSON(cc.chiTietPhat, []), gtYC = 0, btYC = 0;
      (Array.isArray(vp) ? vp : []).forEach(function (x) { var t = num(x.thanhTien) || num(x.mucTien) * (num(x.soLan) || 1); if (x.loai === 'Bồi thường thiệt hại') btYC += t; else gtYC += t; });
      put('ttn', { f: tn.length ? 'SUM(' + A('k:' + tn[0].ma, r) + ':' + A('k:' + tn[tn.length - 1].ma, r) + ')' : '0', v: tnV, z: XL_TIEN });
      put('mien', { f: mienSum.length ? mienSum.join('+') : '0', v: mienV, z: XL_TIEN });
      put('dngt', { v: gtYC, z: XL_TIEN });
      put('ctg', { f: ctg.length ? ctg.join('+') : '0', v: ctgV, z: XL_TIEN });
      put('gqc', { f: 'MIN(' + A('dngt', r) + ',MAX(0,' + A('ctg', r) + '))', v: p.giamThuong, z: XL_TIEN, s: 'bold' });
      put('tong', { f: A('ttn', r) + '-' + A('gqc', r), v: p.tongThuNhap, z: XL_TIEN, s: 'bold' });
      put('tnct', { f: 'MAX(0,' + A('ttn', r) + '-' + A('mien', r) + '-MIN(' + A('gqc', r) + ',' + (ctgt.length ? ctgt.join('+') : '0') + '))', v: p.tnChiuThue, z: XL_TIEN });
      var tgtV = tgbh && (num(bien.NGHI_KL) + num(bien.NGHI_OM)) < (num(cfg.nghiKLToiDaBH) || 14) ? 1 : 0;
      put('tgt', { f: 'IF(AND(' + A('tgbh', r) + '=1,' + ref('NGHI_KL') + '+' + ref('NGHI_OM') + '<' + T.NGUONG_NGHI_BH + '),1,0)', v: tgtV, z: '0' });
      put('lbh', { f: 'IF(' + A('tgt', r) + '=1,MAX(0,ROUND(' + ctExcel(dbh, ref) + ',0)),0)', v: p.luongDongBH, z: XL_TIEN });
      var L = A('lbh', r), capXH = 'IF(' + T.TRAN_BHXH + '>0,MIN(' + L + ',' + T.TRAN_BHXH + '),' + L + ')', capTN = 'IF(' + T.TRAN_BHTN + '>0,MIN(' + L + ',' + T.TRAN_BHTN + '),' + L + ')';
      put('bhxh', { f: 'ROUND(' + capXH + '*' + T.NLD_BHXH + '/100,0)', v: p.bhxh, z: XL_TIEN });
      put('bhyt', { f: 'ROUND(' + capXH + '*' + T.NLD_BHYT + '/100,0)', v: p.bhyt, z: XL_TIEN });
      put('bhtn', { f: 'ROUND(' + capTN + '*' + T.NLD_BHTN + '/100,0)', v: p.bhtn, z: XL_TIEN });
      put('bh', { f: 'SUM(' + A('bhxh', r) + ':' + A('bhtn', r) + ')', v: p.bhxh + p.bhyt + p.bhtn, z: XL_TIEN });
      put('dp', { f: 'IF(AND(' + A('dv', r) + '=1,' + L + '>0),ROUND(MIN(' + L + '*' + T.DOAN_PHI + '/100,' + T.LUONG_CO_SO + '*0.1),0),0)', v: p.doanPhi, z: XL_TIEN });
      put('gtgc', { f: 'IF(OR(' + A('vl', r) + '=1,' + A('gtc', r) + '=0),0,' + T.GT_BAN_THAN + '+' + ref('SO_NPT') + '*' + T.GT_NPT + ')', v: p.giamTru, z: XL_TIEN });
      put('tntt', { f: 'IF(' + A('vl', r) + '=1,' + A('tnct', r) + ',MAX(0,' + A('tnct', r) + '-' + A('bh', r) + '-' + A('gtgc', r) + '))', v: p.tnTinhThue, z: XL_TIEN });
      put('thue', { f: 'IF(' + A('vl', r) + '=1,IF(OR(' + A('ck', r) + '=1,' + A('tnct', r) + '<' + T.NGUONG_VL + '),0,ROUND(' + A('tnct', r) + '*' + T.THUE_VL + '/100,0)),' + bieuThueXL(A('tntt', r), T.bac) + ')', v: p.thueTNCN, z: XL_TIEN, s: 'bold' });
      put('dnbt', { v: btYC, z: XL_TIEN });
      var btmV = Math.max(0, R((p.tongThuNhap - (p.bhxh + p.bhyt + p.bhtn) - p.thueTNCN) * (num(cfg.tyLeBoiThuongToiDa) || 30) / 100));
      put('btmax', { f: 'MAX(0,ROUND((' + A('tong', r) + '-' + A('bh', r) + '-' + A('thue', r) + ')*' + T.TL_BOI_THUONG + '/100,0))', v: btmV, z: XL_TIEN });
      put('bt', { f: 'MIN(' + A('dnbt', r) + ',' + A('btmax', r) + ')', v: p.boiThuong, z: XL_TIEN });
      put('tu', { v: num(cc.tamUng), z: XL_TIEN }); put('anthu', { v: num(p.tienAnThu), z: XL_TIEN }); put('thuung', { v: num(p.thuUng), z: XL_TIEN }); put('ktcn', { v: num(cc.khauTruKhac), z: XL_TIEN });
      var ktRefs = kt.map(function (c) { return A('k:' + c.ma, r); });
      put('cn', { f: A('tong', r) + '-' + A('bh', r) + '-' + A('dp', r) + '-' + A('thue', r) + '-' + A('bt', r) + '-' + A('tu', r) + '-' + A('anthu', r) + '-' + A('thuung', r) + '-' + A('ktcn', r) + (ktRefs.length ? '-' + ktRefs.join('-') : ''), v: p.thucLinh, z: XL_TIEN, s: 'net' });
      var cXH = 'IF(' + T.TRAN_BHXH + '>0,MIN(' + L + ',' + T.TRAN_BHXH + '),' + L + ')';
      put('dnbhxh', { f: 'ROUND(' + cXH + '*' + T.DN_BHXH + '/100,0)', v: p.bhxhDN, z: XL_TIEN });
      put('dnbhyt', { f: 'ROUND(' + cXH + '*' + T.DN_BHYT + '/100,0)', v: p.bhytDN, z: XL_TIEN });
      put('dnbhtn', { f: 'ROUND(' + capTN + '*' + T.DN_BHTN + '/100,0)', v: p.bhtnDN, z: XL_TIEN });
      put('dnkpcd', { f: 'ROUND(' + cXH + '*' + T.DN_KPCD + '/100,0)', v: p.kpcdDN, z: XL_TIEN });
      put('cp', { f: A('tong', r) + '+SUM(' + A('dnbhxh', r) + ':' + A('dnkpcd', r) + ')', v: p.tongThuNhap + p.bhxhDN + p.bhytDN + p.bhtnDN + p.kpcdDN, z: XL_TIEN, s: 'bold' });
      // v3.13: người không đóng BH tại Công ty / lương khoán – ghi giá trị app (không công thức) cho các cột khác cách tính chung
      if (LL !== 'bh') {
        put('lbh', { v: p.luongDongBH, z: XL_TIEN });
        ['bhxh', 'bhyt', 'bhtn'].forEach(function (k) { put(k, { v: 0, z: XL_TIEN }); });
        put('dnbhxh', { v: p.bhxhDN, z: XL_TIEN }); put('dnbhyt', { v: p.bhytDN, z: XL_TIEN }); put('dnbhtn', { v: p.bhtnDN, z: XL_TIEN }); put('dnkpcd', { v: p.kpcdDN, z: XL_TIEN });
      }
      if (LL === 'khoan') {
        var ctK = {}; parseJSON(p.chiTiet, []).forEach(function (x) { ctK[x.ma] = num(x.gt); });
        ks.forEach(function (c) { put('k:' + c.ma, { v: num(ctK[c.ma]), z: c.nhom === 'trunggian' ? '#,##0.####' : XL_TIEN }); });
        put('tk', { v: num(ctK.KHOAN), z: XL_TIEN });
        put('ttn', { v: p.tongThuNhap + p.giamThuong, z: XL_TIEN }); var mienK = 0; parseJSON(p.chiTiet, []).forEach(function (x) { if (x.nhom !== 'khautru') mienK += num(x.gt) - num(x.thue); }); put('mien', { v: mienK, z: XL_TIEN }); put('ctg', { v: num(ctK.THUONGKT), z: XL_TIEN });
        put('tnct', { v: p.tnChiuThue, z: XL_TIEN });
      } else if (coKhoan) put('tk', { v: 0, z: XL_TIEN });
      put('app', { v: d.app ? num(d.app.thucLinh) : p.thucLinh, z: XL_TIEN });
      put('cl', { f: A('cn', r) + '-' + A('app', r), v: p.thucLinh - (d.app ? num(d.app.thucLinh) : p.thucLinh), z: XL_TIEN });
      for (var j = 0; j < C.length; j++) if (row[j] === undefined) row[j] = '';
      rows.push(row);
    });
    var rN = r0 + ds.length - 1, tot = C.map(function (c, j) {
      if (j === 2) return { v: 'TỔNG CỘNG', s: 'tot' };
      var khongCong = ['stt', 'ma', 'ten', 'bp', 'hd', 'll', 'tgbh', 'dv', 'vl', 'ck', 'gtc', 'tgt'].indexOf(c.key) >= 0 || /^v:(CONG_CHUAN|LUONG_NGAY|LUONG_GIO|LA_THU_VIEC|THAM_NIEN|NLV_|SO_NPT|GIO)/.test(c.key) || /^k:/.test(c.key) && maK[c.key.slice(2)] && maK[c.key.slice(2)].nhom === 'trunggian';
      if (khongCong || !ds.length) return { v: '', s: 'tot' };
      var cot = cotXL(j), v = 0; for (var q = 0; q < ds.length; q++) { var x = rows[r0 - 1 + q][j]; v += num(x && typeof x === 'object' ? x.v : x); }
      return { f: 'SUM(' + cot + r0 + ':' + cot + rN + ')', v: v, z: c.z, s: 'tot' };
    });
    rows.push(tot);
    return { ten: 'BangTinhLuong_' + ky, sheets: [{ ten: 'TinhLuong', rows: rows, cols: C.map(function (c) { return c.w; }), merges: merges, freeze: { r: 6, c: 3 }, ngang: true }, ts.sheet] };
  }
  /* các khoản thu nhập / khấu trừ có trong một tập dòng bảng lương (theo thứ tự công thức) */
  function khoanCo(bls, ctl) {
    var co = {}, ten = {}, thuTu = {};
    (ctl || []).forEach(function (c, i) { thuTu[c.ma] = i; });
    bls.forEach(function (b) { parseJSON(b.chiTiet, []).forEach(function (x) { co[x.ma] = x.nhom; ten[x.ma] = x.ten; if (thuTu[x.ma] === undefined) thuTu[x.ma] = 900 + Object.keys(thuTu).length; }); });
    var ds = Object.keys(co).sort(function (a, b) { return thuTu[a] - thuTu[b]; });
    return { tn: ds.filter(function (m) { return co[m] !== 'khautru'; }).map(function (m) { return { ma: m, ten: ten[m] }; }), kt: ds.filter(function (m) { return co[m] === 'khautru'; }).map(function (m) { return { ma: m, ten: ten[m] }; }) };
  }
  function gtKhoan(b, ma) { var x = parseJSON(b.chiTiet, []).filter(function (y) { return y.ma === ma; })[0]; return x ? num(x.gt) : 0; }
  var KHOAN_CU = [{ ma: '_ltg', ten: 'Lương thời gian', k: 'luongThoiGian' }, { ma: '_pc', ten: 'Phụ cấp', k: 'phuCap' }, { ma: '_ot', ten: 'Làm thêm giờ', k: 'tienOT' }, { ma: '_th', ten: 'Thưởng', k: 'thuong' }, { ma: '_tl', ten: 'Truy lĩnh', k: 'truyLinh' }];
  function cotThuNhap(bls, ctl) {
    var kc = khoanCo(bls, ctl), cu = bls.some(function (b) { return !b.chiTiet; });
    var tn = cu ? KHOAN_CU.map(function (x) { return { ma: x.ma, ten: x.ten, gt: function (b) { return b.chiTiet ? (x.k === 'luongThoiGian' ? num(b.luongThoiGian) : num(b[x.k])) : num(b[x.k]); } }; })
      : kc.tn.map(function (x) { return { ma: x.ma, ten: x.ten, gt: function (b) { return gtKhoan(b, x.ma); } }; });
    var kt = cu ? [] : kc.kt.map(function (x) { return { ma: x.ma, ten: x.ten, gt: function (b) { return gtKhoan(b, x.ma); } }; });
    return { tn: tn, kt: kt };
  }
  /* bảng có công thức: định nghĩa cột {ten, gt(b) | f(r, A) , z, cong (cộng ở dòng tổng)} */
  function bangCT(cols, items, r0, A0) {
    var idx = {}; cols.forEach(function (c, i) { if (c.key) idx[c.key] = i; });
    var A = function (key, r) { return cotXL(idx[key]) + r; };
    var rows = items.map(function (b, n) {
      var r = r0 + n, vals = {};
      return cols.map(function (c) {
        var v = c.gt ? c.gt(b, n) : 0;
        if (c.f) { var o = c.f(r, A, b); vals[c.key] = o.v; return { f: o.f, v: o.v, z: c.z, s: c.s }; }
        if (c.key) vals[c.key] = v;
        return c.z ? { v: v, z: c.z, s: c.s } : v;
      });
    });
    return { rows: rows, A: A, idx: idx };
  }
  function dongTong(cols, r1, r2, nhan, ham, rowsData) {
    return cols.map(function (c, j) {
      if (j === (nhan.cot != null ? nhan.cot : 1)) return { v: nhan.ten, s: 'tot' };
      if (!c.cong) return { v: '', s: 'tot' };
      var cot = cotXL(j), v = 0;
      (rowsData || []).forEach(function (row) { var x = row[j]; v += num(x && typeof x === 'object' ? x.v : x); });
      return { f: (ham || 'SUM(') + cot + r1 + ':' + cot + r2 + ')', v: v, z: c.z, s: 'tot' };
    });
  }
  /* v3.41: MẪU PHIẾU LƯƠNG / BẢNG LƯƠNG — người dùng kéo thả thứ tự, ẩn / hiện, đổi tên ở Cài đặt.
     Chỉ đổi cách trình bày; không đổi số tiền. Khối / cột "khoa" luôn hiện. Bảng lương: an = ẩn khi in / PDF, anX = ẩn trong Excel (cột ẩn, công thức vẫn đúng). */
  var MAU_PHIEU_MD = {
    tieuDe: 'PHIẾU LƯƠNG', tieuDe1: 'PHIẾU LƯƠNG THEO CÔNG THỜI GIAN', tieuDe2: 'PHIẾU LƯƠNG LÀM THÊM GIỜ, TĂNG CA',
    khoi: [['dau', 'Đầu phiếu: công ty, kỳ lương, mã NV'], ['tieude', 'Tiêu đề, họ tên – chức vụ', 1], ['the', 'Thẻ thông tin'], ['hd', 'Hợp đồng lao động'], ['cong', 'Ngày công, giờ làm thêm'],
      ['tn', 'Thu nhập', 1], ['bh', 'Các khoản trích theo lương (BH)'], ['thue', 'Thuế thu nhập cá nhân'], ['kt', 'Các khoản phải thu, khấu trừ khác'], ['cn', 'Còn nhận', 1],
      ['chu', 'Bằng chữ'], ['ghichu', 'Ghi chú'], ['ck', 'Tài khoản chuyển khoản'], ['dn', 'BH công ty đóng thêm'], ['ky', 'Chữ ký']].map(function (x) { return { k: x[0], ten: x[1], khoa: !!x[2] }; }),
    the: [['luongChinh', 'Lương chính'], ['luongDongBH', 'Lương đóng BH'], ['cong', 'Công hưởng lương'], ['npt', 'Người phụ thuộc'], ['phongBan', 'Bộ phận', 0, 1], ['chucVu', 'Chức vụ', 0, 1], ['ngayVao', 'Ngày vào làm', 0, 1], ['maSoThue', 'Mã số thuế', 0, 1], ['soTK', 'Số tài khoản', 0, 1], ['thucLinh', 'Thực lĩnh', 0, 1]]
      .map(function (x) { return { k: x[0], ten: x[1], an: !!x[3] }; }),
    ky: ['Người lao động', 'Kế toán tiền lương']
  };
  var MAU_BL_MD = {
    cot: [['stt', 'STT', 1], ['maNV', 'Mã NV', 0, 1], ['hoTen', 'Họ và tên', 1], ['chucVu', 'Chức vụ'], ['phongBan', 'Bộ phận', 0, 1], ['ngayCongHL', 'Công'], ['ngayCongChuan', 'Công chuẩn', 0, 1], ['luongChinh', 'Lương chính', 0, 1],
      ['luongThoiGian', 'Lương TG'], ['phuCap', 'Phụ cấp'], ['tienOT', 'OT'], ['thuong', 'Thưởng'], ['giamThuong', 'Giảm trừ QC'], ['tongThuNhap', 'Tổng TN'],
      ['bhNLD', 'BH NLĐ'], ['doanPhi', 'Đoàn phí', 0, 1], ['thueTNCN', 'Thuế'], ['boiThuong', 'Bồi thường'], ['tamUng', 'Tạm ứng'], ['tienAnThu', 'Ăn ca'], ['thuUng', 'Thu ứng'], ['khauTruKhac', 'KT khác'], ['tongKhauTru', 'Tổng khấu trừ', 0, 1],
      ['thucLinh', 'Thực lĩnh', 1], ['nganHang', 'Ngân hàng', 0, 1], ['soTK', 'Số tài khoản', 0, 1], ['kyNhan', 'Ký nhận']].map(function (x) { return { k: x[0], ten: x[1], khoa: !!x[2], an: !!x[3] }; }),
    ky: []        // trống = theo người duyệt của quy trình (in) / Người lập biểu – Kế toán trưởng – Giám đốc (Excel)
  };
  function mauDs(md, dl) {                 // gộp danh sách người dùng với mặc định: giữ thứ tự người dùng, thêm mục mới, bỏ mục lạ, khóa luôn hiện
    var by = {}; md.forEach(function (x) { by[x.k] = x; });
    var out = [], co = {};
    (Array.isArray(dl) ? dl : []).forEach(function (u) { var d = u && by[u.k]; if (!d || co[u.k]) return; co[u.k] = 1;
      out.push({ k: d.k, ten: String(u.ten || '').trim().slice(0, 80) || d.ten, khoa: !!d.khoa, an: d.khoa ? false : !!u.an, anX: d.khoa ? false : !!u.anX, goc: d.ten }); });
    md.forEach(function (d, i) { if (co[d.k]) return; var x = { k: d.k, ten: d.ten, khoa: !!d.khoa, an: !!d.an, anX: false, goc: d.ten };
      var ks = out.map(function (o) { return o.k; }), j = -1, t, q;
      for (t = i - 1; t >= 0 && j < 0; t--) { q = ks.indexOf(md[t].k); if (q >= 0) j = q + 1; }
      for (t = i + 1; t < md.length && j < 0; t++) { q = ks.indexOf(md[t].k); if (q >= 0) j = q; }
      out.splice(j < 0 ? out.length : j, 0, x); co[d.k] = 1; });
    return out;
  }
  function mauKy(a, md) { return (Array.isArray(a) ? a : md).map(function (x) { return String(x || '').trim().slice(0, 60); }).filter(Boolean).slice(0, 6); }
  function mauPhieu(cfg) {
    var u = (cfg && cfg.mauPhieu) || {}, t = function (k) { return String(u[k] || '').trim().slice(0, 80) || MAU_PHIEU_MD[k]; };
    return { tieuDe: t('tieuDe'), tieuDe1: t('tieuDe1'), tieuDe2: t('tieuDe2'), khoi: mauDs(MAU_PHIEU_MD.khoi, u.khoi), the: mauDs(MAU_PHIEU_MD.the, u.the), ky: mauKy(u.ky, MAU_PHIEU_MD.ky) };
  }
  function mauBL(cfg) { var u = (cfg && cfg.mauBangLuong) || {}; return { cot: mauDs(MAU_BL_MD.cot, u.cot), ky: mauKy(u.ky, MAU_BL_MD.ky) }; }
  function mauGon(m, loai) {              // lưu gọn: chỉ k, ten (nếu đổi), an
    var g = function (ds) { return ds.map(function (x) { var o = { k: x.k }; if (x.ten !== x.goc) o.ten = x.ten; if (x.an) o.an = true; if (x.anX) o.anX = true; return o; }); };
    return loai === 'phieu' ? { tieuDe: m.tieuDe, tieuDe1: m.tieuDe1, tieuDe2: m.tieuDe2, khoi: g(m.khoi), the: g(m.the), ky: m.ky } : { cot: g(m.cot), ky: m.ky };
  }
  /* v3.42: TÌM GẦN ĐÚNG (không dấu, sai 1–2 ký tự, viết tắt chữ đầu). diemTim(q, chuỗi) → 0..100 (0 = không khớp) */
  function boDau(t) { return String(t == null ? '' : t).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }
  function kc(a, b) {                     // khoảng cách Levenshtein (giới hạn chuỗi ngắn)
    if (Math.abs(a.length - b.length) > 3) return 9;
    var d = [], i, j; for (i = 0; i <= a.length; i++) d[i] = [i]; for (j = 0; j <= b.length; j++) d[0][j] = j;
    for (i = 1; i <= a.length; i++) for (j = 1; j <= b.length; j++) d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    return d[a.length][b.length];
  }
  function diemTim(q, t) {
    var Q = boDau(q), T = boDau(t); if (!Q) return 1; if (!T) return 0;
    var qd = String(q).toLowerCase().trim(), td = String(t).toLowerCase();
    var p = T.indexOf(Q);
    if (p >= 0) return Math.min(100, (td.indexOf(qd) >= 0 ? 92 : 88) + (p === 0 ? 8 : (T[p - 1] === ' ' ? 4 : 0)) - Math.min(6, p / 20));
    var qt = Q.split(' '), tt = T.split(' '), tong = 0, du = true;
    qt.forEach(function (w) {                               // mỗi từ gõ: khớp đầu từ / nằm trong / gần đúng
      var b = 0; tt.forEach(function (x) {
        var v = x.indexOf(w) === 0 ? 1 : x.indexOf(w) > 0 ? 0.8 : w.length >= 4 && /^[a-z]+$/.test(w) && /^[a-z]+$/.test(x) ? (function () { var k = Math.min(kc(w, x.slice(0, w.length)), kc(w, x)); return k <= (w.length >= 7 ? 2 : 1) ? 0.7 - k * 0.1 : 0; })() : 0;
        if (v > b) b = v; });
      if (!b) du = false; tong += b;
    });
    var tat = 0;
    if (qt.length === 1 && Q.length >= 2 && Q.length <= tt.length) {      // viết tắt chữ đầu: "nvh" → Nguyễn Văn Hợp
      var dau = tt.filter(function (x) { return /^[a-z]+$/.test(x); }).map(function (x) { return x[0]; }).join(''); if (dau.indexOf(Q) >= 0) tat = 70;
    }
    if (du) return Math.max(tat, Math.round(45 + 40 * tong / qt.length));
    if (tat) return tat;
    if (/\d/.test(Q)) return 0;                                     // số (số công văn, mã): chỉ khớp đúng chuỗi
    var i = 0, j = 0; for (; i < T.length && j < Q.length; i++) if (T[i] === Q[j]) j++;     // các ký tự theo thứ tự
    if (j === Q.length && Q.length >= 3 && Q.replace(/ /g, '').length >= 3) return Math.max(10, 35 - (T.length - Q.length) / 4);
    return 0;
  }
  function locGan(q, ds, layChu, toiDa) {             // sắp xếp danh sách theo độ gần đúng (bỏ mục không khớp)
    if (!String(q || '').trim()) return ds.slice(0, toiDa || ds.length);
    return ds.map(function (x, i) { var c = layChu(x), d = Array.isArray(c) ? Math.max.apply(null, c.map(function (y) { return diemTim(q, y); })) : diemTim(q, c); return { x: x, d: d, i: i }; })
      .filter(function (o) { return o.d > 0; }).sort(function (a, b) { return b.d - a.d || a.i - b.i; }).slice(0, toiDa || 1e9).map(function (o) { return o.x; });
  }
  function tieuDe(ten, phu, ghiChu, nhap) { var r = [[{ v: ten, s: 'title' }], [{ v: phu || '', s: nhap ? 'draft' : 'sub' }]]; r.push(ghiChu ? [{ v: ghiChu, s: 'note' }] : []); return r; }
  /* BẢNG THANH TOÁN TIỀN LƯƠNG THÁNG – theo bộ phận, cộng nhóm SUBTOTAL */
  function xlBangLuongThang(ky, bls, ctl, nvMap, info) {
    var ct = cotThuNhap(bls, ctl), cv = function (b) { return (nvMap[b.maNV] || {}).chucVu || b.chucVu || ''; };
    // v3.32: info.tach – 2 bảng lương độc lập: (1) theo công thời gian (không gồm tiền làm thêm giờ), (2) tăng ca, thêm giờ (thuế TNCN phân bổ)
    var tach = !!(info && info.tach), laOT = { _ot: 1 };
    (ctl || []).forEach(function (c) { if (c.phanLoai === 'tangca') laOT[c.ma] = 1; });
    bls.forEach(function (b) { parseJSON(b.chiTiet, []).forEach(function (x) { if (x.phanLoai === 'tangca') laOT[x.ma] = 1; }); });
    if (tach) ct.tn = ct.tn.filter(function (x) { return !laOT[x.ma]; });
    var otGTb = function (b) { return num(b.otThucLinh) + num(b.otThue) || num(b.tienOT); }, otTL = function (b) { return num(b.otThucLinh) + num(b.otThue) ? num(b.otThucLinh) : num(b.tienOT); };   // kỳ tính trước v3.32: chưa có số phân bổ
    var tongV = function (b) { return num(b.tongThuNhap) - (tach ? otGTb(b) : 0); }, tlV = function (b) { return num(b.thucLinh) - (tach ? otTL(b) : 0); };
    var cols = [{ ten: 'STT', mk: 'stt' }, { ten: 'Mã NV', mk: 'maNV', gt: function (b) { return b.maNV; } }, { ten: 'Họ và tên', mk: 'hoTen', gt: function (b) { return b.hoTen; } }, { ten: 'Chức vụ', mk: 'chucVu', gt: cv },
      { key: 'cong', mk: 'ngayCongHL', ten: 'Công HL', gt: function (b) { return num(b.ngayCongHL); }, z: XL_CONG, cong: 1 }, { ten: 'Công chuẩn', mk: 'ngayCongChuan', gt: function (b) { return num(b.ngayCongChuan); }, z: XL_CONG },
      { key: 'lc', mk: 'luongChinh', ten: 'Lương chính', gt: function (b) { return num(b.luongChinh); }, z: XL_TIEN }];
    ct.tn.forEach(function (x, i) { cols.push({ key: 'tn' + i, ten: x.ten, gt: x.gt, z: XL_TIEN, cong: 1 }); });
    cols.push({ key: 'gqc', mk: 'giamThuong', ten: 'Giảm trừ quy chế', gt: function (b) { return num(b.giamThuong); }, z: XL_TIEN, cong: 1 });
    cols.push({ key: 'tong', mk: 'tongThuNhap', ten: 'TỔNG THU NHẬP', z: XL_TIEN, cong: 1, s: 'bold', f: function (r, A, b) { return { f: (ct.tn.length ? 'SUM(' + A('tn0', r) + ':' + A('tn' + (ct.tn.length - 1), r) + ')' : '0') + '-' + A('gqc', r), v: tongV(b) }; } });
    // v3.41: mẫu bảng lương — thứ tự nhóm khấu trừ, ẩn cột (ẩn trong Excel, công thức vẫn đúng), đổi tên cột
    var mau = (info && info.mau) || mauBL(null), mi = {}; mau.cot.forEach(function (x, i) { mi[x.k] = x; x.i = i; });
    var ded = [];
    [['bhxh', 'BHXH', 'bhNLD'], ['bhyt', 'BHYT', 'bhNLD'], ['bhtn', 'BHTN', 'bhNLD'], ['doanPhi', 'Đoàn phí'], ['thueTNCN', 'Thuế TNCN'], ['boiThuong', 'Bồi thường'], ['tamUng', 'Tạm ứng'], ['tienAnThu', 'Ăn ca'], ['thuUng', 'Thu ứng'], ['khauTruKhac', 'Thu khấu trừ khác']].forEach(function (x, j) { ded.push({ key: x[0], mk: x[2] || x[0], j: j, ten: x[1], gt: function (b) { return x[0] === 'thueTNCN' && tach ? num(b.thueTNCN) - num(b.otThue) : num(b[x[0]]); }, z: XL_TIEN, cong: 1 }); });
    ct.kt.forEach(function (x, i) { ded.push({ key: 'kt' + i, mk: 'khauTruKhac', j: 100 + i, ten: x.ten, gt: x.gt, z: XL_TIEN, cong: 1 }); });
    ded.sort(function (a, b) { return ((mi[a.mk] || {}).i - (mi[b.mk] || {}).i) || (a.j - b.j); });
    ded.forEach(function (c) { cols.push(c); });
    var dauKT = ded[0].key, cuoiKT = ded[ded.length - 1].key;
    cols.push({ key: 'tkt', mk: 'tongKhauTru', ten: 'Tổng khấu trừ', z: XL_TIEN, cong: 1, f: function (r, A, b) { return { f: 'SUM(' + A(dauKT, r) + ':' + A(cuoiKT, r) + ')', v: tongV(b) - tlV(b) }; } });
    cols.push({ key: 'cn', mk: 'thucLinh', ten: 'CÒN NHẬN', z: XL_TIEN, cong: 1, s: 'net', f: function (r, A, b) { return { f: A('tong', r) + '-' + A('tkt', r), v: tlV(b) }; } });
    cols.push({ ten: 'Ký nhận', mk: 'kyNhan' });
    cols.forEach(function (c) { var m = c.mk && mi[c.mk]; if (!m) return; if (m.anX && !m.khoa && c.mk !== 'tongKhauTru') c.an = 1; if (m.ten !== m.goc && c.mk !== 'bhNLD' && !/^kt\d/.test(c.key || '')) c.ten = m.ten; });
    var rows = tieuDe((tach ? 'BẢNG 1 – THANH TOÁN TIỀN LƯƠNG THEO CÔNG THỜI GIAN THÁNG ' : 'BẢNG THANH TOÁN TIỀN LƯƠNG THÁNG ') + ky.slice(5) + '/' + ky.slice(0, 4), (info.congTy || '') + (info.trangThai ? ' · ' + info.trangThai : ''), 'Đơn vị tính: đồng. Cột Tổng thu nhập, Tổng khấu trừ, Còn nhận và các dòng cộng là công thức.', info.nhap);
    rows.push(cols.map(function (c) { return { v: c.ten, s: 'head' }; }));
    // v3.21: info.nhomCT – tách nhóm theo công ty thành viên (b.tenCongTy, thứ tự info.thuTuCT), trong mỗi công ty theo bộ phận.
    // SUBTOTAL bỏ qua các dòng SUBTOTAL lồng bên trong → cộng bộ phận, cộng công ty, tổng cộng đều đúng.
    var bpCua = function (b) { return b.phongBan || 'Chưa xếp bộ phận'; }, ctCua = function (b) { return info.nhomCT ? (b.tenCongTy || 'Chưa khai báo công ty') : ''; };
    var ctTT = info.thuTuCT || [], ctIdx = function (c) { var i = ctTT.indexOf(c); return i < 0 ? 999 : i; };
    var nhomCT = {}, thuTuCT = [];
    bls.slice().sort(function (a, b) { return (ctIdx(ctCua(a)) - ctIdx(ctCua(b))) || ctCua(a).localeCompare(ctCua(b)) || bpCua(a).localeCompare(bpCua(b)) || (a.maNV < b.maNV ? -1 : 1); }).forEach(function (b) {
      var c = ctCua(b), k = bpCua(b);
      if (!nhomCT[c]) { nhomCT[c] = { bp: {}, thuTu: [] }; thuTuCT.push(c); }
      var g = nhomCT[c]; if (!g.bp[k]) { g.bp[k] = []; g.thuTu.push(k); } g.bp[k].push(b);
    });
    var r = rows.length + 1, dau = r, stt = 0, all = [];
    var hang = function (ten, st) { return [{ v: '', s: st }, { v: ten, s: st }].concat(cols.slice(2).map(function () { return { v: '', s: st }; })); };
    thuTuCT.forEach(function (c) {
      var g = nhomCT[c], dauCT = r, allCT = [], soCT = 0;
      if (info.nhomCT) { rows.push(hang(c.toUpperCase(), 'grpc')); r++; }
      g.thuTu.forEach(function (k) {
        rows.push(hang(k, 'grp')); r++;
        var t = bangCT(cols, g.bp[k], r);
        t.rows.forEach(function (row) { row[0] = ++stt; soCT++; rows.push(row); all.push(row); allCT.push(row); });
        var r1 = r, r2 = r + g.bp[k].length - 1; r = r2 + 1;
        rows.push(dongTong(cols, r1, r2, { ten: 'Cộng ' + k, cot: 2 }, 'SUBTOTAL(9,', t.rows)); r++;
      });
      if (info.nhomCT) { var tc = dongTong(cols, dauCT, r - 1, { ten: 'CỘNG ' + c.toUpperCase() + ' (' + soCT + ' người)', cot: 2 }, 'SUBTOTAL(9,', allCT); tc.forEach(function (x) { x.s = 'totc'; }); rows.push(tc); r++; }
    });
    rows.push(dongTong(cols, dau, r - 1, { ten: 'TỔNG CỘNG' + (info.nhomCT && thuTuCT.length > 1 ? ' ' + thuTuCT.length + ' CÔNG TY' : '') + ' (' + stt + ' người)', cot: 2 }, 'SUBTOTAL(9,', all));
    var kyDs = mau.ky.length ? mau.ky : ['Người lập biểu', 'Kế toán trưởng', 'Giám đốc'], kyHien = [];
    cols.forEach(function (c, i) { if (!c.an && i > 0) kyHien.push(i); });
    var kyR = [[], []], buoc = Math.max(1, Math.floor(kyHien.length / kyDs.length));
    kyDs.forEach(function (t, n) { var c = kyHien[Math.min(kyHien.length - 1, n * buoc)]; kyR[0][c] = { v: t, s: 'sign' }; kyR[1][c] = { v: n === kyDs.length - 1 && /giám đốc/i.test(t) ? '(Ký, đóng dấu)' : '(Ký, ghi rõ họ tên)', s: 'note' }; });
    kyR.forEach(function (r) { for (var q = 0; q < r.length; q++) if (r[q] === undefined) r[q] = ''; });
    rows.push([], [], kyR[0], kyR[1]);
    var ck = tieuDe('DANH SÁCH CHUYỂN KHOẢN LƯƠNG THÁNG ' + ky.slice(5) + '/' + ky.slice(0, 4), (info.congTy || '') + (info.trangThai ? ' · ' + info.trangThai : ''), '', info.nhap);
    ck.push(['STT', 'Mã NV', 'Họ và tên', 'Ngân hàng', 'Số tài khoản', 'Số tiền'].concat(info.nhomCT ? ['Công ty'] : []).map(function (c) { return { v: c, s: 'head' }; }));
    var ds2 = bls.slice().sort(function (a, b) { return (ctIdx(ctCua(a)) - ctIdx(ctCua(b))) || ctCua(a).localeCompare(ctCua(b)) || (a.maNV < b.maNV ? -1 : 1); }), c0 = ck.length + 1;
    ds2.forEach(function (b, i) { var n = nvMap[b.maNV] || {}; ck.push([i + 1, b.maNV, b.hoTen, n.nganHang || '', { v: String(n.soTK || ''), s: '' }, { v: tlV(b), z: XL_TIEN }].concat(info.nhomCT ? [ctCua(b)] : [])); });
    ck.push([{ v: '', s: 'tot' }, { v: '', s: 'tot' }, { v: 'TỔNG CỘNG', s: 'tot' }, { v: '', s: 'tot' }, { v: '', s: 'tot' }, { f: 'SUM(F' + c0 + ':F' + (c0 + ds2.length - 1) + ')', v: bls.reduce(function (t, b) { return t + tlV(b); }, 0), z: XL_TIEN, s: 'tot' }]);
    var sh = [{ ten: 'BangLuong_' + ky.slice(5) + '_' + ky.slice(0, 4), rows: rows, cols: cols.map(function (c, i) { var w = i === 0 ? 5 : i === 2 ? 24 : i === 3 ? 18 : i === 1 ? 9 : 13; return c.an ? { wch: w, hidden: true } : w; }), freeze: { r: 4, c: 3 }, ngang: true }, { ten: 'ChuyenKhoan', rows: ck, cols: [5, 9, 26, 22, 20, 16, 24] }];
    if (tach) {
      var gio = function (b, k) { return num(parseJSON(b.ctCong, {})[k]); };
      var oc = [{ ten: 'STT' }, { ten: 'Mã NV', gt: function (b) { return b.maNV; } }, { ten: 'Họ và tên', gt: function (b) { return b.hoTen; } }, { ten: 'Bộ phận', gt: function (b) { return b.phongBan || ''; } },
        { key: 'g1', ten: 'Giờ 150%', gt: function (b) { return gio(b, 'ot150'); }, z: XL_CONG, cong: 1 }, { key: 'g2', ten: 'Giờ 200%', gt: function (b) { return gio(b, 'ot200'); }, z: XL_CONG, cong: 1 }, { key: 'g3', ten: 'Giờ 300%', gt: function (b) { return gio(b, 'ot300'); }, z: XL_CONG, cong: 1 },
        { key: 'gc', ten: 'Cộng giờ', z: XL_CONG, cong: 1, f: function (r, A, b) { return { f: 'SUM(' + A('g1', r) + ':' + A('g3', r) + ')', v: gio(b, 'ot150') + gio(b, 'ot200') + gio(b, 'ot300') }; } },
        { key: 'ot', ten: 'Tiền làm thêm giờ', gt: otGTb, z: XL_TIEN, cong: 1 }, { key: 'th', ten: 'Thuế TNCN phân bổ', gt: function (b) { return num(b.otThue); }, z: XL_TIEN, cong: 1 },
        { key: 'cn', ten: 'CÒN NHẬN', z: XL_TIEN, cong: 1, s: 'net', f: function (r, A, b) { return { f: A('ot', r) + '-' + A('th', r), v: otTL(b) }; } }, { ten: 'Ký nhận' }];
      var dsOT = bls.filter(function (b) { return otGTb(b) > 0; }).sort(function (a, b) { return (a.phongBan || '').localeCompare(b.phongBan || '') || (a.maNV < b.maNV ? -1 : 1); });
      var ro = tieuDe('BẢNG 2 – THANH TOÁN TIỀN LƯƠNG LÀM THÊM GIỜ, TĂNG CA THÁNG ' + ky.slice(5) + '/' + ky.slice(0, 4), (info.congTy || '') + (info.trangThai ? ' · ' + info.trangThai : ''), 'Thuế TNCN phân bổ = thuế cả tháng − thuế khi không có phần chịu thuế của tiền làm thêm giờ (tiền làm thêm trong giới hạn giờ được miễn thuế). Bảo hiểm, tạm ứng, công nợ trừ ở Bảng 1.', info.nhap);
      ro.push(oc.map(function (c) { return { v: c.ten, s: 'head' }; }));
      var o0 = ro.length + 1, to = bangCT(oc, dsOT, o0);
      to.rows.forEach(function (x, i) { x[0] = i + 1; ro.push(x); });
      if (dsOT.length) ro.push(dongTong(oc, o0, o0 + dsOT.length - 1, { ten: 'TỔNG CỘNG (' + dsOT.length + ' người)', cot: 2 }, null, to.rows));
      ro.push([], [], ['', { v: 'Người lập biểu', s: 'sign' }, '', '', '', { v: 'Kế toán trưởng', s: 'sign' }, '', '', '', { v: 'Giám đốc', s: 'sign' }]);
      var ck2 = tieuDe('DANH SÁCH CHUYỂN KHOẢN TIỀN LÀM THÊM GIỜ THÁNG ' + ky.slice(5) + '/' + ky.slice(0, 4), info.congTy || '', '', info.nhap);
      ck2.push(['STT', 'Mã NV', 'Họ và tên', 'Ngân hàng', 'Số tài khoản', 'Số tiền'].map(function (c) { return { v: c, s: 'head' }; }));
      var k0 = ck2.length + 1; dsOT.forEach(function (b, i) { var n = nvMap[b.maNV] || {}; ck2.push([i + 1, b.maNV, b.hoTen, n.nganHang || '', { v: String(n.soTK || ''), s: '' }, { v: otTL(b), z: XL_TIEN }]); });
      ck2.push([{ v: '', s: 'tot' }, { v: '', s: 'tot' }, { v: 'TỔNG CỘNG', s: 'tot' }, { v: '', s: 'tot' }, { v: '', s: 'tot' }, { f: 'SUM(F' + k0 + ':F' + (k0 + Math.max(1, dsOT.length) - 1) + ')', v: dsOT.reduce(function (t, b) { return t + otTL(b); }, 0), z: XL_TIEN, s: 'tot' }]);
      sh.splice(1, 0, { ten: 'TangCa_' + ky.slice(5) + '_' + ky.slice(0, 4), rows: ro, cols: [5, 9, 24, 22, 10, 10, 10, 10, 15, 14, 15, 14], freeze: { r: 4, c: 3 }, ngang: true });
      sh.push({ ten: 'ChuyenKhoan_TangCa', rows: ck2, cols: [5, 9, 26, 22, 20, 16] });
    }
    return { ten: 'BangLuong_' + ky + (tach ? '_2bang' : ''), sheets: sh };
  }
  /* bảng ma trận người × tháng */
  function maTran(ten, phu, nguoi, kys, gt, z) {
    var rows = tieuDe(ten, phu, 'Cột Cộng năm và dòng Tổng cộng là công thức.');
    var cols = ['STT', 'Mã NV', 'Họ và tên', 'Bộ phận'].concat(kys.map(function (k) { return 'T' + k.slice(5); }), ['Cộng năm']);
    rows.push(cols.map(function (c) { return { v: c, s: 'head' }; }));
    var r0 = rows.length + 1;
    nguoi.forEach(function (p, i) {
      var r = r0 + i, row = [i + 1, p.maNV, p.hoTen, p.phongBan || ''], s = 0;
      kys.forEach(function (k) { var v = gt(p.maNV, k); s += v || 0; row.push(v === null ? '' : { v: v, z: z }); });
      row.push({ f: 'SUM(' + cotXL(4) + r + ':' + cotXL(3 + kys.length) + r + ')', v: s, z: z, s: 'bold' });
      rows.push(row);
    });
    var rN = r0 + nguoi.length - 1, tot = [{ v: '', s: 'tot' }, { v: '', s: 'tot' }, { v: 'TỔNG CỘNG', s: 'tot' }, { v: '', s: 'tot' }];
    for (var j = 4; j <= 4 + kys.length; j++) { var v = 0; for (var q = 0; q < nguoi.length; q++) { var x = rows[r0 - 1 + q][j]; v += num(x && typeof x === 'object' ? x.v : x); } tot.push({ f: 'SUM(' + cotXL(j) + r0 + ':' + cotXL(j) + rN + ')', v: v, z: z, s: 'tot' }); }
    rows.push(tot);
    return { rows: rows, cols: [5, 9, 24, 22].concat(kys.map(function () { return 13; }), [15]), freeze: { r: 4, c: 3 }, ngang: true };
  }
  /* BẢNG LƯƠNG NĂM: tổng hợp theo tháng + ma trận thu nhập, thuế, còn nhận */
  function xlNam(nam, bls, kyInfo, info) {
    var kys = []; for (var m = 1; m <= 12; m++) kys.push(nam + '-' + pad(m));
    var by = {}; bls.forEach(function (b) { (by[b.ky] = by[b.ky] || []).push(b); });
    var sumK = function (k, f) { return (by[k] || []).reduce(function (s, b) { return s + f(b); }, 0); };
    var cols = [{ ten: 'Tháng', gt: function (k) { return k.slice(5) + '/' + k.slice(0, 4); } }, { ten: 'Trạng thái', gt: function (k) { return by[k] ? (kyInfo[k] || '') : 'Chưa có'; } },
      { key: 'sn', ten: 'Số người', gt: function (k) { return (by[k] || []).length; }, z: '0', cong: 1 },
      { key: 'tong', ten: 'Tổng thu nhập', gt: function (k) { return sumK(k, function (b) { return num(b.tongThuNhap); }); }, z: XL_TIEN, cong: 1 },
      { key: 'bh', ten: 'BH người lao động', gt: function (k) { return sumK(k, function (b) { return num(b.bhxh) + num(b.bhyt) + num(b.bhtn); }); }, z: XL_TIEN, cong: 1 },
      { key: 'dp', ten: 'Đoàn phí', gt: function (k) { return sumK(k, function (b) { return num(b.doanPhi); }); }, z: XL_TIEN, cong: 1 },
      { key: 'thue', ten: 'Thuế TNCN', gt: function (k) { return sumK(k, function (b) { return num(b.thueTNCN); }); }, z: XL_TIEN, cong: 1 },
      { key: 'ktk', ten: 'Tạm ứng, bồi thường, khấu trừ khác', gt: function (k) { return sumK(k, function (b) { return num(b.boiThuong) + num(b.tamUng) + num(b.tienAnThu) + num(b.thuUng) + num(b.khauTruKhac) + num(b.ktCongThuc); }); }, z: XL_TIEN, cong: 1 },
      { key: 'cn', ten: 'Còn nhận', z: XL_TIEN, cong: 1, s: 'net', f: function (r, A, k) { return { f: A('tong', r) + '-' + A('bh', r) + '-' + A('dp', r) + '-' + A('thue', r) + '-' + A('ktk', r), v: sumK(k, function (b) { return num(b.thucLinh); }) }; } },
      { key: 'bhdn', ten: 'BH doanh nghiệp đóng', gt: function (k) { return sumK(k, function (b) { return num(b.bhxhDN) + num(b.bhytDN) + num(b.bhtnDN); }); }, z: XL_TIEN, cong: 1 },
      { key: 'kpcd', ten: 'KPCĐ', gt: function (k) { return sumK(k, function (b) { return num(b.kpcdDN); }); }, z: XL_TIEN, cong: 1 },
      { key: 'cp', ten: 'Tổng chi phí lương', z: XL_TIEN, cong: 1, s: 'bold', f: function (r, A, k) { return { f: A('tong', r) + '+' + A('bhdn', r) + '+' + A('kpcd', r), v: sumK(k, function (b) { return num(b.tongThuNhap) + num(b.bhxhDN) + num(b.bhytDN) + num(b.bhtnDN) + num(b.kpcdDN); }) }; } }];
    var rows = tieuDe('TỔNG HỢP QUỸ LƯƠNG NĂM ' + nam, info.congTy || '', 'Còn nhận = Tổng thu nhập − BH − Đoàn phí − Thuế − khấu trừ khác; Tổng chi phí = Tổng thu nhập + BH DN + KPCĐ. ' + (info.ghiChu || ''));
    rows.push(cols.map(function (c) { return { v: c.ten, s: 'head' }; }));
    var r0 = rows.length + 1, t = bangCT(cols, kys, r0);
    t.rows.forEach(function (x) { rows.push(x); });
    rows.push(dongTong(cols, r0, r0 + 11, { ten: 'CẢ NĂM', cot: 0 }, null, t.rows));
    var nguoi = {}; bls.forEach(function (b) { if (!nguoi[b.maNV] || b.ky > nguoi[b.maNV].ky) nguoi[b.maNV] = b; });
    var ds = Object.keys(nguoi).map(function (m) { return nguoi[m]; }).sort(function (a, b) { return (a.phongBan || '').localeCompare(b.phongBan || '') || (a.maNV < b.maNV ? -1 : 1); });
    var bl1 = {}; bls.forEach(function (b) { bl1[b.maNV + '|' + b.ky] = b; });
    var g = function (k) { return function (ma, ky) { var b = bl1[ma + '|' + ky]; return b ? num(b[k]) : null; }; };
    var s1 = maTran('THU NHẬP TỪNG NGƯỜI NĂM ' + nam, 'Tổng thu nhập theo tháng', ds, kys, g('tongThuNhap'), XL_TIEN);
    var s2 = maTran('THUẾ TNCN ĐÃ KHẤU TRỪ NĂM ' + nam, 'Dùng đối chiếu khi quyết toán thuế', ds, kys, g('thueTNCN'), XL_TIEN);
    var s3 = maTran('THU NHẬP CHỊU THUẾ NĂM ' + nam, '', ds, kys, g('tnChiuThue'), XL_TIEN);
    var s4 = maTran('CÒN NHẬN TỪNG NGƯỜI NĂM ' + nam, 'Số chuyển khoản theo tháng', ds, kys, g('thucLinh'), XL_TIEN);
    return { ten: 'BangLuongNam_' + nam, sheets: [{ ten: 'TongHop_' + nam, rows: rows, cols: [12, 14].concat(cols.slice(2).map(function () { return 16; })), freeze: { r: 4, c: 1 }, ngang: true },
      Object.assign({ ten: 'ThuNhap' }, s1), Object.assign({ ten: 'ThueTNCN' }, s2), Object.assign({ ten: 'TNChiuThue' }, s3), Object.assign({ ten: 'ConNhan' }, s4)] };
  }
  /* BẢNG LƯƠNG CÁ NHÂN theo năm */
  function xlCaNhan(nv, nam, bls, ctl, kyInfo, info) {
    bls = bls.slice().sort(function (a, b) { return a.ky < b.ky ? -1 : 1; });
    var ct = cotThuNhap(bls, ctl);
    var cols = [{ ten: 'Tháng', gt: function (b) { return b.ky.slice(5) + '/' + b.ky.slice(0, 4); } }, { ten: 'Trạng thái', gt: function (b) { return kyInfo[b.ky] || ''; } },
      { ten: 'Công HL', gt: function (b) { return num(b.ngayCongHL); }, z: XL_CONG, cong: 1 }, { ten: 'Công chuẩn', gt: function (b) { return num(b.ngayCongChuan); }, z: XL_CONG },
      { ten: 'Lương chính', gt: function (b) { return num(b.luongChinh); }, z: XL_TIEN }];
    ct.tn.forEach(function (x, i) { cols.push({ key: 'tn' + i, ten: x.ten, gt: x.gt, z: XL_TIEN, cong: 1 }); });
    cols.push({ key: 'gqc', ten: 'Giảm trừ quy chế', gt: function (b) { return num(b.giamThuong); }, z: XL_TIEN, cong: 1 });
    cols.push({ key: 'tong', ten: 'TỔNG THU NHẬP', z: XL_TIEN, cong: 1, s: 'bold', f: function (r, A, b) { return { f: (ct.tn.length ? 'SUM(' + A('tn0', r) + ':' + A('tn' + (ct.tn.length - 1), r) + ')' : '0') + '-' + A('gqc', r), v: num(b.tongThuNhap) }; } });
    [['bhxh', 'BHXH'], ['bhyt', 'BHYT'], ['bhtn', 'BHTN'], ['doanPhi', 'Đoàn phí']].forEach(function (x) { cols.push({ key: x[0], ten: x[1], gt: function (b) { return num(b[x[0]]); }, z: XL_TIEN, cong: 1 }); });
    cols.push({ key: 'tnct', ten: 'TN chịu thuế', gt: function (b) { return num(b.tnChiuThue); }, z: XL_TIEN, cong: 1 }, { key: 'gtgc', ten: 'Giảm trừ gia cảnh', gt: function (b) { return num(b.giamTru); }, z: XL_TIEN, cong: 1 });
    cols.push({ key: 'tntt', ten: 'TN tính thuế', z: XL_TIEN, cong: 1, f: function (r, A, b) { return { f: num(b.giamTru) || !num(b.tnTinhThue) ? 'MAX(0,' + A('tnct', r) + '-' + A('bhxh', r) + '-' + A('bhyt', r) + '-' + A('bhtn', r) + '-' + A('gtgc', r) + ')' : A('tnct', r), v: num(b.tnTinhThue) }; } });
    [['thueTNCN', 'Thuế TNCN'], ['boiThuong', 'Bồi thường'], ['tamUng', 'Tạm ứng'], ['tienAnThu', 'Ăn ca'], ['thuUng', 'Thu ứng'], ['khauTruKhac', 'Thu khấu trừ khác']].forEach(function (x) { cols.push({ key: x[0], ten: x[1], gt: function (b) { return num(b[x[0]]); }, z: XL_TIEN, cong: 1 }); });
    ct.kt.forEach(function (x, i) { cols.push({ key: 'kt' + i, ten: x.ten, gt: x.gt, z: XL_TIEN, cong: 1 }); });
    cols.push({ key: 'cn', ten: 'CÒN NHẬN', z: XL_TIEN, cong: 1, s: 'net', f: function (r, A, b) { return { f: A('tong', r) + '-SUM(' + A('bhxh', r) + ':' + A('doanPhi', r) + ')-SUM(' + A('thueTNCN', r) + ':' + A(ct.kt.length ? 'kt' + (ct.kt.length - 1) : 'khauTruKhac', r) + ')', v: num(b.thucLinh) }; } });
    var rows = tieuDe('BẢNG LƯƠNG CÁ NHÂN NĂM ' + nam, (info.congTy || ''), 'Họ tên: ' + nv.hoTen + ' · Mã NV: ' + nv.maNV + ' · ' + (nv.chucVu || '') + ' · ' + (nv.phongBan || '') + ' · Số người phụ thuộc: ' + num(nv.soNPT));
    rows.push(cols.map(function (c) { return { v: c.ten, s: 'head' }; }));
    var r0 = rows.length + 1, t = bangCT(cols, bls, r0);
    t.rows.forEach(function (x) { rows.push(x); });
    if (bls.length) rows.push(dongTong(cols, r0, r0 + bls.length - 1, { ten: 'CỘNG NĂM', cot: 0 }, null, t.rows));
    return { ten: 'Luong_' + nv.maNV + '_' + nam, sheets: [{ ten: nv.maNV + '_' + nam, rows: rows, cols: [10, 12].concat(cols.slice(2).map(function () { return 13; })), freeze: { r: 4, c: 1 }, ngang: true }] };
  }
  /* CHI PHÍ LƯƠNG & BẢO HIỂM: sheet dữ liệu + tổng hợp theo bộ phận, theo tháng bằng SUMIFS */
  function xlChiPhi(tieuDeKy, bls, info) {
    var D = [['Kỳ', 'Mã NV', 'Họ và tên', 'Bộ phận', 'Tổng thu nhập', 'BHXH NLĐ', 'BHYT NLĐ', 'BHTN NLĐ', 'Đoàn phí', 'Thuế TNCN', 'Còn nhận', 'BHXH DN', 'BHYT DN', 'BHTN DN', 'KPCĐ', 'Lương đóng BH', 'Công ty']];
    var ctCua = function (b) { return b.tenCongTy || b.congTy || 'Chưa khai báo công ty'; };
    var F = ['tongThuNhap', 'bhxh', 'bhyt', 'bhtn', 'doanPhi', 'thueTNCN', 'thucLinh', 'bhxhDN', 'bhytDN', 'bhtnDN', 'kpcdDN', 'luongDongBH'];
    var ds = bls.slice().sort(function (a, b) { return a.ky < b.ky ? -1 : a.ky > b.ky ? 1 : (a.maNV < b.maNV ? -1 : 1); });
    var dl = [[{ v: 'DỮ LIỆU BẢNG LƯƠNG ' + tieuDeKy, s: 'title' }], [{ v: 'Nguồn cho các sheet tổng hợp (SUMIFS). Không chèn dòng giữa bảng.', s: 'note' }], [], D[0].map(function (c) { return { v: c, s: 'head' }; })];
    ds.forEach(function (b) { dl.push([b.ky, b.maNV, b.hoTen, b.phongBan || 'Chưa xếp bộ phận'].concat(F.map(function (k) { return { v: num(b[k]), z: XL_TIEN }; }), [ctCua(b)])); });
    var a = 5, z = 4 + Math.max(1, ds.length), rg = function (c) { return 'DuLieu!$' + c + '$' + a + ':$' + c + '$' + z; };
    var COL = { tong: 'E', bhxh: 'F', bhyt: 'G', bhtn: 'H', bhxhDN: 'L', bhytDN: 'M', bhtnDN: 'N', kpcd: 'O', cn: 'K' };
    function tongHop(ten, nhanTieuDe, khoa, colKhoa, gtKhoa) {
      var cols = [nhanTieuDe, 'Số lượt người', 'Quỹ lương (tổng thu nhập)', 'BHXH NLĐ', 'BHYT NLĐ', 'BHTN NLĐ', 'BHXH DN', 'BHYT DN', 'BHTN DN', 'KPCĐ', 'Tổng BH + KPCĐ DN', 'Tổng chi phí', 'Tỷ trọng', 'Còn nhận (chuyển khoản)'];
      var rows = tieuDe(ten, info.congTy || '', 'Mọi số là công thức SUMIFS trên sheet DuLieu.');
      rows.push(cols.map(function (c) { return { v: c, s: 'head' }; }));
      var r0 = rows.length + 1, rN = r0 + khoa.length - 1, vals = [];
      khoa.forEach(function (k, i) {
        var r = r0 + i, sel = ds.filter(function (b) { return gtKhoa(b) === k; }), s = function (f) { return sel.reduce(function (t, b) { return t + num(b[f]); }, 0); };
        var crit = rg(colKhoa) + ',$A' + r, si = function (c, f) { return { f: 'SUMIFS(' + rg(c) + ',' + crit + ')', v: s(f), z: XL_TIEN }; };
        var cp = s('tongThuNhap') + s('bhxhDN') + s('bhytDN') + s('bhtnDN') + s('kpcdDN'); vals.push(cp);
        rows.push([k, { f: 'COUNTIFS(' + crit + ')', v: sel.length, z: '0' }, si(COL.tong, 'tongThuNhap'), si(COL.bhxh, 'bhxh'), si(COL.bhyt, 'bhyt'), si(COL.bhtn, 'bhtn'), si(COL.bhxhDN, 'bhxhDN'), si(COL.bhytDN, 'bhytDN'), si(COL.bhtnDN, 'bhtnDN'), si(COL.kpcd, 'kpcdDN'),
          { f: 'SUM(G' + r + ':J' + r + ')', v: s('bhxhDN') + s('bhytDN') + s('bhtnDN') + s('kpcdDN'), z: XL_TIEN }, { f: 'C' + r + '+K' + r, v: cp, z: XL_TIEN, s: 'bold' }, { f: 'IF($L$' + (rN + 1) + '=0,0,L' + r + '/$L$' + (rN + 1) + ')', v: 0, z: XL_PT }, si(COL.cn, 'thucLinh')]);
      });
      var tongCP = vals.reduce(function (t, v) { return t + v; }, 0);
      rows.slice(r0 - 1).forEach(function (row) { row[12].v = tongCP ? row[11].v / tongCP : 0; });
      var tot = [{ v: 'TỔNG CỘNG', s: 'tot' }];
      for (var j = 1; j < cols.length; j++) { var c = cotXL(j), v = 0; for (var q = r0 - 1; q < rows.length; q++) v += num(rows[q][j].v); tot.push({ f: 'SUM(' + c + r0 + ':' + c + rN + ')', v: v, z: j === 12 ? XL_PT : j === 1 ? '0' : XL_TIEN, s: 'tot' }); }
      rows.push(tot);
      return { ten: ten.indexOf('BỘ PHẬN') >= 0 ? 'TheoBoPhan' : ten.indexOf('THEO CÔNG TY') >= 0 ? 'TheoCongTy' : 'TheoThang', rows: rows, cols: [30, 10].concat(cols.slice(2).map(function () { return 15; })), freeze: { r: 4, c: 1 }, ngang: true };
    }
    var bps = [], kys = [];
    ds.forEach(function (b) { var p = b.phongBan || 'Chưa xếp bộ phận'; if (bps.indexOf(p) < 0) bps.push(p); if (kys.indexOf(b.ky) < 0) kys.push(b.ky); });
    bps.sort(function (x, y) { return x.localeCompare(y); });
    var s1 = tongHop('CHI PHÍ LƯƠNG & BẢO HIỂM THEO BỘ PHẬN – ' + tieuDeKy, 'Bộ phận', bps, 'D', function (b) { return b.phongBan || 'Chưa xếp bộ phận'; });
    var s2 = tongHop('CHI PHÍ LƯƠNG & BẢO HIỂM THEO THÁNG – ' + tieuDeKy, 'Kỳ', kys, 'A', function (b) { return b.ky; });
    var cts = []; ds.forEach(function (b) { var c = ctCua(b); if (cts.indexOf(c) < 0) cts.push(c); });
    var s3 = tongHop('CHI PHÍ LƯƠNG & BẢO HIỂM THEO CÔNG TY THÀNH VIÊN – ' + tieuDeKy, 'Công ty', cts, 'Q', ctCua);
    return { ten: 'ChiPhiLuong_' + tieuDeKy.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/[^0-9A-Za-z]+/g, '_'), sheets: [s1, s3, s2, { ten: 'DuLieu', rows: dl, cols: [9, 9, 24, 26].concat(F.map(function () { return 14; }), [30]), freeze: { r: 4, c: 3 } }] };
  }
  /* ---- MẪU EXCEL TẠM ỨNG: tạo file cho kế toán điền, đọc lại & nhận diện ---- */
  function chuanTen(s) { return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'd').toLowerCase().replace(/[^a-z0-9]/g, ''); }
  function docTien(v) {
    if (typeof v === 'number') return isFinite(v) ? Math.round(v) : NaN;
    var s = String(v == null ? '' : v).trim().toLowerCase().replace(/\s+/g, '').replace(/đ|vnd|dong|đồng/g, '');
    if (!s || s === '-') return 0;
    var m = /^([\d.,]+)(tr|trieu|triệu|k|ng|nghin|nghìn)?$/.exec(s); if (!m) return NaN;
    var so = m[1], dv = m[2] || '', x;
    if (dv) x = parseFloat(so.replace(',', '.'));                                  // 1,5tr – 500k
    else if (/^\d{1,3}([.,]\d{3})+$/.test(so)) x = parseFloat(so.replace(/[.,]/g, ''));  // 1.500.000 – 1,500,000
    else x = parseFloat(so.replace(',', '.'));
    if (!isFinite(x)) return NaN;
    if (/^(tr|trieu|triệu)$/.test(dv)) x *= 1000000; else if (dv) x *= 1000;
    return Math.round(x);
  }
  function docNgay(v, ky) {
    if (v == null || v === '') return '';
    if (typeof v === 'number' && v > 20000 && v < 80000) { var d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000); return fmtD(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()); }
    if (v instanceof Date) return fmtD(v.getFullYear(), v.getMonth() + 1, v.getDate());
    var s = String(v).trim(), m;
    if ((m = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/.exec(s))) return fmtD(+m[3], +m[2], +m[1]);
    if ((m = /^(\d{1,2})[\/.-](\d{1,2})$/.exec(s)) && ky) return fmtD(+ky.slice(0, 4), +m[2], +m[1]);
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    return null;
  }
  /* loại tạm ứng từ tên cột / ô "Loại": khớp nguyên văn, rồi theo từ khóa, còn lại → loại cuối (Ứng khác) */
  function nhanLoaiTU(ten, loai) {
    var k = chuanTen(ten); if (!k) return null;
    for (var i = 0; i < loai.length; i++) if (chuanTen(loai[i]) === k || chuanTen(loai[i]).replace(/^ung/, '') === k.replace(/^(tamung|ung)/, '')) return loai[i];
    var tu = [[/luong/, 'luong'], [/xang|dilai|congtac|xe|xangdau/, 'xang'], [/tienan|anca|com|ungan|^an/, 'tienan'], [/khac/, 'khac']];
    for (var j = 0; j < tu.length; j++) if (tu[j][0].test(k)) { var hit = loai.filter(function (l) { return new RegExp(tu[j][1]).test(chuanTen(l)); })[0]; if (hit) return hit; }
    return null;
  }
  function xlMauTamUng(ky, nvs, loai, info) {
    var t = 'MẪU NHẬP TẠM ỨNG – KỲ LƯƠNG ' + ky.slice(5) + '/' + ky.slice(0, 4);
    var rows = [[{ v: t, s: 'title' }], [{ v: (info && info.congTy) || '', s: 'sub' }],
      [{ v: 'Nhập số tiền vào cột loại tạm ứng tương ứng (đồng, VD 1500000 hoặc 1.500.000). Không sửa cột Mã NV và dòng tiêu đề. Mỗi người nhiều lần ứng cùng loại: cộng dồn, hoặc kê từng lần ở sheet ChiTiet.', s: 'note' }]];
    var hd = ['STT', 'Mã NV', 'Họ và tên', 'Bộ phận'].concat(loai, ['Tổng tạm ứng', 'Ngày ứng', 'Ghi chú / số phiếu chi']);
    rows.push(hd.map(function (h) { return { v: h, s: 'head' }; }));
    var r0 = rows.length + 1, c1 = cotXL(4), c2 = cotXL(3 + loai.length);
    nvs.forEach(function (n, i) {
      var r = r0 + i;
      rows.push([i + 1, n.maNV, n.hoTen, n.phongBan || ''].concat(loai.map(function () { return { v: '', z: XL_TIEN }; }), [{ f: 'SUM(' + c1 + r + ':' + c2 + r + ')', v: 0, z: XL_TIEN, s: 'bold' }, { v: '', z: 'dd/mm/yyyy' }, '']));
    });
    var rN = r0 + nvs.length - 1, tot = [{ v: '', s: 'tot' }, { v: '', s: 'tot' }, { v: 'TỔNG CỘNG', s: 'tot' }, { v: '', s: 'tot' }];
    for (var j = 4; j <= 4 + loai.length; j++) tot.push({ f: 'SUM(' + cotXL(j) + r0 + ':' + cotXL(j) + Math.max(r0, rN) + ')', v: 0, z: XL_TIEN, s: 'tot' });
    tot.push({ v: '', s: 'tot' }, { v: '', s: 'tot' }); rows.push(tot);
    var ct = [[{ v: 'KÊ TỪNG LẦN TẠM ỨNG (không bắt buộc)', s: 'title' }], [{ v: 'Dùng khi một người ứng nhiều lần, cần giữ ngày từng lần. Cột Loại ghi đúng một trong: ' + loai.join('; ') + '.', s: 'note' }], [],
      ['Mã NV', 'Họ và tên', 'Loại tạm ứng', 'Ngày ứng', 'Số tiền', 'Ghi chú / số phiếu chi'].map(function (h) { return { v: h, s: 'head' }; })];
    for (var q = 0; q < 30; q++) ct.push(['', '', '', { v: '', z: 'dd/mm/yyyy' }, { v: '', z: XL_TIEN }, '']);
    var hdn = [[{ v: 'HƯỚNG DẪN', s: 'title' }], [],
      ['1. Điền sheet TamUng (mỗi người một dòng, mỗi loại một cột) và/hoặc sheet ChiTiet (mỗi lần ứng một dòng).'],
      ['2. App nhận diện người theo Mã NV; thiếu Mã NV thì theo Họ và tên (phải trùng khớp duy nhất).'],
      ['3. Số tiền: 1500000, 1.500.000, 1,5tr hoặc 500k đều được. Ô trống hoặc 0 bỏ qua.'],
      ['4. Ngày ứng: dd/mm/yyyy. Để trống được.'],
      ['5. Vào app → Chấm công & Tính lương → kỳ ' + ky.slice(5) + '/' + ky.slice(0, 4) + ' → tab Tạm ứng → "Nhập từ Excel", xem trước rồi xác nhận.'],
      [], [{ v: 'Các loại tạm ứng', s: 'head' }]].concat(loai.map(function (l) { return [l]; }));
    return { ten: 'Mau_TamUng_' + ky, sheets: [{ ten: 'TamUng', rows: rows, cols: [5, 9, 26, 24].concat(loai.map(function () { return 15; }), [16, 12, 26]), freeze: { r: 4, c: 3 } },
      { ten: 'ChiTiet', rows: ct, cols: [10, 26, 22, 12, 15, 30] }, { ten: 'HuongDan', rows: hdn, cols: [90] }] };
  }
  /* đọc file đã điền: sheets = { tên sheet: mảng hàng (aoa) }; nvs = CBCNV có trong kỳ */
  function docTamUng(sheets, nvs, loai, ky) {
    var out = [], loi = [], canhBao = [], byMa = {}, byTen = {};
    nvs.forEach(function (n) { byMa[String(n.maNV).toUpperCase()] = n; var k = chuanTen(n.hoTen); (byTen[k] = byTen[k] || []).push(n); });
    function timNV(ma, ten, vt) {
      ma = String(ma == null ? '' : ma).trim().toUpperCase();
      if (ma) { if (byMa[ma]) return byMa[ma]; loi.push(vt + ': không có mã NV "' + ma + '" trong kỳ.'); return null; }
      var ds = byTen[chuanTen(ten)] || [];
      if (ds.length === 1) return ds[0];
      if (ds.length > 1) loi.push(vt + ': có ' + ds.length + ' người tên "' + ten + '", cần ghi Mã NV.');
      else if (String(ten || '').trim()) loi.push(vt + ': không tìm thấy "' + ten + '".');
      return null;
    }
    Object.keys(sheets).forEach(function (tenSh) {
      var aoa = sheets[tenSh] || [], h = -1, map = null;
      for (var i = 0; i < Math.min(aoa.length, 15); i++) {
        var hdr = (aoa[i] || []).map(chuanTen);
        if (hdr.some(function (x) { return x === 'manv' || x === 'mnv' || x === 'hovaten' || x === 'hoten'; })) { h = i; map = hdr; break; }
      }
      if (h < 0) return;
      var cMa = map.indexOf('manv') >= 0 ? map.indexOf('manv') : map.indexOf('mnv'), cTen = map.indexOf('hovaten') >= 0 ? map.indexOf('hovaten') : map.indexOf('hoten');
      var cNgay = map.findIndex(function (x) { return /^ngay/.test(x); }), cGhi = map.findIndex(function (x) { return /^ghichu/.test(x); });
      var cLoai = map.findIndex(function (x) { return /^loai/.test(x); }), cTien = map.findIndex(function (x) { return x === 'sotien' || x === 'tien'; });
      var dai = cLoai >= 0 && cTien >= 0, cotLoai = [];
      if (!dai) (aoa[h] || []).forEach(function (x, j) {
        if ([cMa, cTen, cNgay, cGhi].indexOf(j) >= 0) return;
        var k = chuanTen(x); if (!k || /^(stt|bophan|phongban|tong|chucvu|tongtamung|kyten|kynhan)/.test(k)) return;
        var l = nhanLoaiTU(x, loai);
        if (l) cotLoai.push({ j: j, loai: l, ten: String(x) }); else canhBao.push(tenSh + ': bỏ qua cột "' + x + '" (không phải loại tạm ứng).');
      });
      if (!dai && !cotLoai.length) return;
      for (var r = h + 1; r < aoa.length; r++) {
        var row = aoa[r] || [], vt = tenSh + ' dòng ' + (r + 1);
        var ma = cMa >= 0 ? row[cMa] : '', ten = cTen >= 0 ? row[cTen] : '';
        if (/^(tong|cong)/.test(chuanTen(ten)) || /^(tong|cong)/.test(chuanTen(ma))) continue;
        var cells = dai ? [{ v: row[cTien], loai: row[cLoai] }] : cotLoai.map(function (c) { return { v: row[c.j], loai: c.loai, cot: c.ten }; });
        var coSo = cells.some(function (c) { var t = docTien(c.v); return t && t !== 0; });
        if (!coSo) { if (cells.some(function (c) { return isNaN(docTien(c.v)); })) loi.push(vt + ': số tiền "' + cells.filter(function (c) { return isNaN(docTien(c.v)); })[0].v + '" không đọc được.'); continue; }
        var nv = timNV(ma, ten, vt); if (!nv) continue;
        var ngay = cNgay >= 0 ? docNgay(row[cNgay], ky) : '';
        if (ngay === null) { canhBao.push(vt + ': ngày "' + row[cNgay] + '" không đọc được, để trống.'); ngay = ''; }
        else if (ngay && ky && ngay.slice(0, 7) !== ky) canhBao.push(vt + ': ngày ' + ngay + ' không thuộc kỳ ' + ky + '.');
        var ghi = cGhi >= 0 ? String(row[cGhi] == null ? '' : row[cGhi]).trim() : '';
        cells.forEach(function (c) {
          var t = docTien(c.v);
          if (isNaN(t)) { loi.push(vt + (c.cot ? ' cột "' + c.cot + '"' : '') + ': số tiền "' + c.v + '" không đọc được.'); return; }
          if (!t) return;
          if (t < 0) { loi.push(vt + ': số tiền âm.'); return; }
          var l = dai ? nhanLoaiTU(c.loai, loai) : c.loai;
          if (!l) { l = loai[loai.length - 1]; canhBao.push(vt + ': loại "' + (c.loai || '') + '" không nhận diện được, xếp vào ' + l + '.'); }
          out.push({ maNV: nv.maNV, hoTen: nv.hoTen, loai: l, soTien: t, ngay: ngay, ghiChu: ghi, nguon: vt });
        });
      }
    });
    return { items: out, loi: loi, canhBao: canhBao };
  }
  /* chuyển mô tả workbook → đối tượng của thư viện xlsx-js-style (SheetJS) */
  var XL_STYLE = {
    title: { font: { bold: true, sz: 14, color: { rgb: '0F1E3D' } } }, sub: { font: { italic: true, color: { rgb: '4A5672' } } }, draft: { font: { bold: true, color: { rgb: 'B91C1C' } } },
    note: { font: { italic: true, sz: 9, color: { rgb: '7B869E' } } }, code: { font: { sz: 8, color: { rgb: '7B869E' } }, alignment: { horizontal: 'center' } },
    head: { font: { bold: true, color: { rgb: '0F1E3D' } }, fill: { fgColor: { rgb: 'DCE4F1' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } },
    grp: { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '22386A' } }, alignment: { horizontal: 'left' } },
    grpc: { font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: 'C2570C' } }, alignment: { horizontal: 'left' } },
    totc: { font: { bold: true, color: { rgb: '7C2D12' } }, fill: { fgColor: { rgb: 'FDDCC2' } } },
    tot: { font: { bold: true }, fill: { fgColor: { rgb: 'FFF1E6' } } }, sign: { font: { bold: true } }, bold: { font: { bold: true } }, net: { font: { bold: true, color: { rgb: 'C2570C' } } }
  };
  var VIEN = { top: { style: 'thin', color: { rgb: 'C9D3E6' } }, bottom: { style: 'thin', color: { rgb: 'C9D3E6' } }, left: { style: 'thin', color: { rgb: 'C9D3E6' } }, right: { style: 'thin', color: { rgb: 'C9D3E6' } } };
  function sangXLSX(spec, X) {
    var wb = X.utils.book_new();
    spec.sheets.forEach(function (sh) {
      var ws = {}, maxC = 0, bang = false;
      sh.rows.forEach(function (row, r) {
        (row || []).forEach(function (c, j) {
          if (c === '' || c === null || c === undefined) return;
          var o = typeof c === 'object' ? c : { v: c }, cell = {};
          if (o.f) { cell.f = o.f; cell.t = 'n'; cell.v = num(o.v); }
          else if (typeof o.v === 'number') { cell.t = 'n'; cell.v = o.v; }
          else { if (o.v === '' && !o.s) return; cell.t = 's'; cell.v = String(o.v == null ? '' : o.v); }
          if (o.z) cell.z = o.z;
          var st = o.s && XL_STYLE[o.s] ? JSON.parse(JSON.stringify(XL_STYLE[o.s])) : {};
          if (o.s === 'head') bang = true;
          if (bang && ['title', 'sub', 'note', 'draft', 'sign'].indexOf(o.s) < 0) st.border = VIEN;
          if (o.z) st.numFmt = o.z;
          if (Object.keys(st).length) cell.s = st;
          ws[cotXL(j) + (r + 1)] = cell; if (j > maxC) maxC = j;
        });
        if ((row || []).some(function (c) { return c && c.s === 'title'; })) bang = false;
      });
      ws['!ref'] = 'A1:' + cotXL(Math.max(maxC, 0)) + Math.max(sh.rows.length, 1);
      if (sh.cols) ws['!cols'] = sh.cols.map(function (w) { return w && typeof w === 'object' ? w : { wch: w }; });
      if (sh.merges) ws['!merges'] = sh.merges;
      if (sh.freeze) ws['!freeze'] = sh.freeze;
      X.utils.book_append_sheet(wb, ws, sh.ten.slice(0, 31));
    });
    wb.Workbook = { CalcPr: { fullCalcOnLoad: true } };
    return wb;
  }

  /* khoảng cách 2 tọa độ (m) – công thức haversine */
  function khoangCachM(a1, o1, a2, o2) {
    var R = 6371000, r = Math.PI / 180, dA = (a2 - a1) * r, dO = (o2 - o1) * r;
    var h = Math.sin(dA / 2) * Math.sin(dA / 2) + Math.cos(a1 * r) * Math.cos(a2 * r) * Math.sin(dO / 2) * Math.sin(dO / 2);
    return 2 * R * Math.asin(Math.min(1, Math.sqrt(h)));
  }
  /* độ lệch 2 mẫu khuôn mặt (vector 128 chiều của face-api): < 0,5 thường là cùng một người */
  function kmKhoangCach(a, b) { var s = 0; for (var i = 0; i < 128; i++) { var d = (+a[i] || 0) - (+b[i] || 0); s += d * d; } return Math.sqrt(s); }
  /* ===== PHÂN TÍCH AI NHÓM CHAT (v2.8): lời nhắc, đọc JSON, nội dung Google Docs, Excel ===== */
  var THU_VN = ['Chủ nhật', 'Thứ hai', 'Thứ ba', 'Thứ tư', 'Thứ năm', 'Thứ sáu', 'Thứ bảy'];
  function ngayVN(d) { d = String(d || ''); return d.length >= 10 ? d.slice(8, 10) + '/' + d.slice(5, 7) + '/' + d.slice(0, 4) : d; }
  var AI_HE = 'Bạn là chuyên viên phân tích của Ban Giám đốc Công ty Cổ phần TAPACO Việt Nam (xây dựng hạ tầng giao thông, hạ tầng kỹ thuật, khu đô thị, khu công nghiệp). ' +
    'Bạn đọc nhật ký nhóm chat công trường / phòng ban để lập báo cáo quản trị.\n' +
    'Nguyên tắc bắt buộc:\n' +
    '1. Chỉ dùng thông tin có trong dữ liệu được cung cấp. Không bịa, không tự suy ra số liệu. Không chắc thì bỏ qua.\n' +
    '2. Mỗi số liệu vật tư, khối lượng phải kèm ngày, tài khoản người nêu và trích dẫn nguyên văn ngắn (≤ 120 ký tự) từ tin nhắn hoặc báo cáo.\n' +
    '3. Nhận xét khách quan, dựa trên việc làm cụ thể. Không nhận xét đời tư, sức khỏe, tôn giáo, chính trị, ngoại hình.\n' +
    '4. Tiếng Việt, ngắn gọn, văn phong báo cáo hành chính.\n' +
    '5. Chỉ trả về MỘT đối tượng JSON hợp lệ: không lời dẫn, không giải thích, không bọc trong ```.';
  var AI_KHIA = ['Nhân sự & tinh thần làm việc', 'Thời tiết & ảnh hưởng thi công', 'Kiểm soát chất lượng', 'Khối lượng & sản lượng', 'Tiến độ', 'Nhà thầu phụ', 'Nhà cung cấp',
    'An toàn lao động', 'An ninh trật tự', 'Sự cố', 'Năng lực kiểm soát của Ban chỉ huy công trường', 'Phối hợp giữa phòng ban và công trường', 'Quản lý tài chính (tạm ứng, thanh toán)',
    'Quản lý kho vật tư', 'Giao nhận vật tư', 'Quản lý thiết bị', 'Quản lý định mức thi công', 'Rủi ro cho công ty'];
  var AI_MUC = ['Tốt', 'Đạt', 'Cần chú ý', 'Nghiêm trọng', 'Chưa đủ dữ liệu'], AI_HANG = ['Rất tích cực', 'Tích cực', 'Đạt', 'Cần cải thiện', 'Ít hoạt động'];
  var AI_P_DOT = 'Đọc nhật ký trên, lập NHẬT KÝ THI CÔNG từng ngày và trích xuất thông tin. Trả về JSON đúng cấu trúc:\n' +
    '{\n' +
    ' "ngay": [ {\n' +
    '   "ngay": "YYYY-MM-DD", "tomTat": "≤ 50 từ",\n' +
    '   "thoiTiet": { "sang": "", "chieu": "", "anhHuong": "ảnh hưởng tới thi công nếu có" },\n' +
    '   "nhanLuc": { "tong": số|null, "banChiHuy": số|null, "congNhan": số|null, "thauPhu": [ { "ten": "", "soNguoi": số|null } ], "ghiChu": "" },\n' +
    '   "thietBi": [ { "ten": "", "soLuong": số|null, "tinhTrang": "hoạt động|hỏng|chờ sửa|thay thế|khác", "ghiChu": "" } ],\n' +
    '   "thiCong": [ { "hangMuc": "", "viTri": "lý trình, trục, cấu kiện", "khoiLuong": số|null, "dvt": "", "keHoach": số|null, "soVoiKH": "đạt|vượt|chậm|chưa rõ", "ghiChu": "", "trich": "nguyên văn ≤ 80 ký tự" } ],\n' +
    '   "vatTuNhap": [ { "ten": "", "dvt": "", "soLuong": số|null, "nhaCungCap": "", "trich": "" } ],\n' +
    '   "vatTuXuat": [ { "ten": "", "dvt": "", "soLuong": số|null, "choHangMuc": "", "trich": "" } ],\n' +
    '   "yeuCauVatTu": [ { "ten": "", "dvt": "", "soLuong": số|null, "username": "", "trich": "" } ],\n' +
    '   "nghiemThu": [ { "noiDung": "", "ketQua": "đạt|không đạt|chờ", "thanhPhan": "" } ],\n' +
    '   "kienNghi": [ { "username": "", "noiDung": "kiến nghị, đề xuất của cá nhân", "trich": "" } ],\n' +
    '   "anToan": { "tinhHinh": "", "baoHo": "", "viPham": [""] }, "anNinh": "", "suCo": [""], "tonTai": ["tồn tại, vướng mắc"], "keHoachMai": [""]\n' +
    ' } ],\n' +
    ' "doiTac": [ { "ten": "tên thầu phụ / nhà cung cấp", "loai": "thầu phụ|nhà cung cấp", "ngay": "YYYY-MM-DD", "noiDung": "", "danhGia": "tích cực|tiêu cực|trung tính", "trich": "" } ],\n' +
    ' "suKien": [ { "ngay": "YYYY-MM-DD", "loai": "an toàn|an ninh|chất lượng|tiến độ|thiết bị|vật tư|tài chính|định mức|phối hợp|khác", "mucDo": "thấp|trung bình|cao", "noiDung": "", "trich": "" } ],\n' +
    ' "chungTu": [ { "fileId": "mã tệp có trong nhật ký", "moTa": "ảnh/tệp thể hiện gì", "lienQuan": "hạng mục, vật tư liên quan" } ],\n' +
    ' "caNhan": [ { "username": "", "dongGop": ["việc làm cụ thể, tích cực"], "canLuuY": ["điểm cần lưu ý nếu có"] } ]\n' +
    '}\n' +
    'Quy ước: mỗi ngày có trong nhật ký là một phần tử "ngay". Mục nào nhật ký KHÔNG nêu thì để "" , [] hoặc null – tuyệt đối không suy đoán, không điền số ước lượng. ' +
    'PHẢI đọc hết khối "## Báo cáo định kỳ" của từng ngày và lấy số liệu từ đó trước: đây là số liệu do chính thành viên nộp theo mẫu, chính xác hơn tin nhắn trò chuyện. ' +
    'Ánh xạ tên trường của mẫu báo cáo sang các mục: "Nhân lực có mặt", "Số người trên công trường", "Ban chỉ huy có mặt", "Nhân lực" → nhanLuc (tong, banChiHuy, congNhan); ' +
    '"Khối lượng thực hiện trong ngày", "Việc đã hoàn thành", "Việc đang thực hiện", "Công việc chính" và mọi khối lượng ghi kèm đơn vị (m3, m2, md, tấn, cấu kiện) → thiCong (hangMuc, viTri, khoiLuong, dvt); ' +
    '"Máy, thiết bị hoạt động" → thietBi; "Vật tư nhập về" → vatTuNhap; "Vật tư sử dụng" → vatTuXuat; "Đề nghị cấp vật tư" → yeuCauVatTu; ' +
    '"Vướng mắc, đề xuất", "Khó khăn, cần hỗ trợ" → kienNghi (username là người nộp báo cáo) và tonTai; "Kế hoạch ngày mai" → keHoachMai; "Thời tiết" → thoiTiet; ' +
    '"Có sự cố an toàn lao động", "Kiểm tra bảo hộ lao động", "Mô tả sự cố", "Biện pháp khắc phục" → anToan và suCo. ' +
    'Tên người nộp ghi dạng **Họ tên** (tài_khoản, nộp lúc …): mọi việc, khối lượng, kiến nghị lấy từ báo cáo của người nào phải quy về đúng tài_khoản đó trong caNhan.dongGop và kienNghi.username. ' +
    'Khi báo cáo và tin nhắn nêu khác nhau về cùng một số liệu thì lấy theo báo cáo và ghi chênh lệch vào "ghiChu". Nhiều người cùng báo cáo một ngày thì cộng nhân lực lại, nêu cách cộng trong nhanLuc.ghiChu. ' +
    '"trich" của mục lấy từ báo cáo ghi nguyên văn giá trị trường đó. ' +
    'Ví dụ: ô "Khối lượng thực hiện trong ngày: Đổ bê tông trụ T3: 42 m³; lắp cốt thép dầm D3 60%" của tài khoản baopq phải cho ra ' +
    'thiCong = [{"hangMuc":"Đổ bê tông trụ T3","khoiLuong":42,"dvt":"m3","trich":"Đổ bê tông trụ T3: 42 m³"},{"hangMuc":"Lắp cốt thép dầm D3","khoiLuong":60,"dvt":"%","trich":"lắp cốt thép dầm D3 60%"}] ' +
    'và caNhan có {"username":"baopq","dongGop":["Đổ bê tông trụ T3 42 m3","Lắp cốt thép dầm D3 60%"]}. Mỗi việc trong ô khối lượng là MỘT phần tử thiCong, không gộp chung một dòng. ' +
    'Dòng "[số liệu máy đọc được, dùng làm chuẩn]" trong nhật ký là kết quả máy đã tách sẵn: lấy đúng các số đó, không tính lại. "keHoach" và "soVoiKH" chỉ ghi khi nhật ký có nêu kế hoạch / khối lượng dự kiến (kể cả kế hoạch nêu từ ngày trước). ' +
    'Số dùng dấu chấm thập phân (12.5). Chỉ dùng tài khoản có trong danh sách thành viên, chỉ dùng mã tệp có trong nhật ký. Tối đa 25 phần tử mỗi danh sách.';
  var AI_P_TONG1 = 'Dựa trên thống kê và nhật ký thi công đã tổng hợp ở trên, lập BỨC TRANH TỔNG THỂ cho Giám đốc. Trả về JSON:\n' +
    '{\n' +
    ' "tongThe": "250–350 từ: tình hình chung, kết quả, vấn đề nổi bật, xu hướng",\n' +
    ' "mucRuiRo": "Thấp|Trung bình|Cao", "lyDoRuiRo": "≤ 60 từ",\n' +
    ' "danhGia": [ { "khia": "tên khía cạnh", "muc": "Tốt|Đạt|Cần chú ý|Nghiêm trọng|Chưa đủ dữ liệu", "nhanXet": "≤ 90 từ, phân tích cụ thể", "danChung": ["dd/mm: dẫn chứng ngắn từ dữ liệu"], "giaiPhap": ["giải pháp xử lý cụ thể"] } ]\n' +
    '}\n' +
    '"danhGia" phải có ĐỦ và ĐÚNG THỨ TỰ các khía cạnh: ' + AI_KHIA.join('; ') + '. ' +
    'Khía cạnh không có dữ liệu ghi "Chưa đủ dữ liệu", nhận xét nêu rõ thiếu thông tin gì và đề xuất cách thu thập (ví dụ thêm mục vào mẫu báo cáo).';
  var AI_P_TONG2 = 'Dựa trên số liệu cá nhân, ghi nhận, kiến nghị, đối tác và đánh giá khía cạnh ở trên, trả về JSON:\n' +
    '{\n' +
    ' "caNhan": [ { "username": "", "xepLoai": "Rất tích cực|Tích cực|Đạt|Cần cải thiện|Ít hoạt động", "nhanXet": "≤ 60 từ: mức nhiệt tình, trách nhiệm, chất lượng thông tin", "deXuat": "≤ 25 từ: khen, nhắc nhở, đào tạo, điều chỉnh công việc…" } ],\n' +
    ' "doiTac": [ { "ten": "", "loai": "thầu phụ|nhà cung cấp", "danhGia": "Tốt|Đạt|Cần chú ý|Kém", "nhanXet": "≤ 50 từ", "deXuat": "" } ],\n' +
    ' "diemManh": [""], "tonTai": [""],\n' +
    ' "giaiPhap": [ { "vanDe": "tồn tại cần xử lý", "giaiPhap": "việc cần làm cụ thể", "phuTrach": "bộ phận / chức danh đề xuất", "thoiHan": "ngay|trong tuần|trong tháng", "uuTien": "Cao|Trung bình|Thấp" } ],\n' +
    ' "kienNghiBGD": ["việc Giám đốc cần chỉ đạo, quyết định"]\n' +
    '}\n' +
    'Đánh giá đủ mọi thành viên trong bảng cá nhân. Căn cứ trước hết vào số liệu do hệ thống tính (ngày tương tác, ngày không tương tác trừ ngày nghỉ có đơn, tỷ lệ nộp báo cáo đúng hạn, điểm tham gia), rồi đến đóng góp, kiến nghị cụ thể. ' +
    'Người có ít dữ liệu xếp "Ít hoạt động", không suy diễn. Giải pháp sắp xếp theo mức ưu tiên, tối đa 15 mục.';
  var AI_TRUONG_NK = [['thoiTiet', 'Thời tiết'], ['nhanLuc', 'Nhân lực'], ['thietBi', 'Thiết bị'], ['thiCong', 'Công tác thi công'], ['vatTuNhap', 'Vật tư nhập'], ['vatTuXuat', 'Vật tư xuất'],
    ['yeuCauVatTu', 'Yêu cầu vật tư'], ['nghiemThu', 'Nghiệm thu'], ['kienNghi', 'Kiến nghị, đề xuất'], ['anToan', 'An toàn, bảo hộ'], ['anNinh', 'An ninh'], ['suCo', 'Sự cố'], ['tonTai', 'Tồn tại, vướng mắc'], ['keHoachMai', 'Kế hoạch ngày mai']];
  function aiSL(v, dvt) { return v == null || v === '' ? '' : aiSoTxt(v) + (dvt ? ' ' + dvt : ''); }
  function aiJSON(t) {
    if (t && typeof t === 'object') return t;
    t = String(t || '').trim().replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
    try { return JSON.parse(t); } catch (e) { }
    var a = t.indexOf('{'), b = t.lastIndexOf('}');
    if (a >= 0 && b > a) { try { return JSON.parse(t.slice(a, b + 1)); } catch (e) { } }
    throw new Error('AI trả về kết quả không đúng định dạng (có thể bị cắt vì quá dài). Hãy chọn khoảng ngày ngắn hơn rồi chạy lại.');
  }
  function aiSo(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var s = String(v).replace(/\s/g, '').replace(/[^\d.,-]/g, '');
    if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
    else if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) s = s.replace(/,/g, '');
    else s = s.replace(',', '.');
    var n = parseFloat(s); return isFinite(n) ? n : null;
  }
  function aiSoTxt(v) { return v == null || v === '' ? '' : (Math.round(v * 1000) / 1000).toLocaleString('vi-VN'); }
  /* văn bản ngắn gọn cho từng ô nhật ký thi công (dùng chung Docs, Word, Excel, giao diện) */
  function nkO(n, k, kq) {
    var tn = function (u) { return aiTen(kq, u); };
    switch (k) {
      case 'thoiTiet': { var t = n.thoiTiet || {}; return [t.sang ? 'Sáng: ' + t.sang : '', t.chieu ? 'Chiều: ' + t.chieu : '', t.anhHuong ? 'Ảnh hưởng: ' + t.anhHuong : ''].filter(Boolean).join('; '); }
      case 'nhanLuc': { var l = n.nhanLuc || {}; return [l.tong != null ? 'Tổng ' + aiSoTxt(l.tong) + ' người' : '', l.banChiHuy != null ? 'Ban chỉ huy ' + aiSoTxt(l.banChiHuy) : '', l.congNhan != null ? 'công nhân ' + aiSoTxt(l.congNhan) : '',
        (l.thauPhu || []).map(function (x) { return x.ten + (x.soNguoi != null ? ' ' + aiSoTxt(x.soNguoi) + ' người' : ''); }).join(', '), l.ghiChu || '', (n.nghiPhep || []).length ? 'Nghỉ có đơn: ' + n.nghiPhep.join(', ') : ''].filter(Boolean).join('; '); }
      case 'thietBi': return (n.thietBi || []).map(function (x) { return x.ten + (x.soLuong != null ? ' ×' + aiSoTxt(x.soLuong) : '') + (x.tinhTrang ? ' (' + x.tinhTrang + ')' : '') + (x.ghiChu ? ' – ' + x.ghiChu : ''); }).join('; ');
      case 'thiCong': return (n.thiCong || []).map(function (x) { return x.hangMuc + (x.viTri ? ' ' + x.viTri : '') + (x.khoiLuong != null ? ': ' + aiSL(x.khoiLuong, x.dvt) : '') + (x.keHoach != null ? ' / KH ' + aiSL(x.keHoach, x.dvt) : '') + (x.soVoiKH && x.soVoiKH !== 'chưa rõ' ? ' (' + x.soVoiKH + ')' : '') + (x.ghiChu ? ' – ' + x.ghiChu : ''); }).join('; ');
      case 'vatTuNhap': return (n.vatTuNhap || []).map(function (x) { return x.ten + ' ' + aiSL(x.soLuong, x.dvt) + (x.nhaCungCap ? ' (' + x.nhaCungCap + ')' : ''); }).join('; ');
      case 'vatTuXuat': return (n.vatTuXuat || []).map(function (x) { return x.ten + ' ' + aiSL(x.soLuong, x.dvt) + (x.choHangMuc ? ' cho ' + x.choHangMuc : ''); }).join('; ');
      case 'yeuCauVatTu': return (n.yeuCauVatTu || []).map(function (x) { return x.ten + ' ' + aiSL(x.soLuong, x.dvt) + (x.username ? ' (' + tn(x.username) + ')' : ''); }).join('; ');
      case 'nghiemThu': return (n.nghiemThu || []).map(function (x) { return x.noiDung + (x.ketQua ? ': ' + x.ketQua : '') + (x.thanhPhan ? ' (' + x.thanhPhan + ')' : ''); }).join('; ');
      case 'kienNghi': return (n.kienNghi || []).map(function (x) { return (x.username ? tn(x.username) + ': ' : '') + x.noiDung; }).join('; ');
      case 'anToan': { var a = n.anToan || {}; return [a.tinhHinh, a.baoHo ? 'Bảo hộ: ' + a.baoHo : '', (a.viPham || []).length ? 'Vi phạm: ' + a.viPham.join(', ') : ''].filter(Boolean).join('; '); }
      case 'anNinh': return n.anNinh || '';
      default: return (n[k] || []).join('; ');
    }
  }
  /* nội dung báo cáo dạng khối (Google Docs trên Apps Script, Word / bản in trên trình duyệt dùng chung) */
  function aiKhoi(kq, anh) {
    var tk = kq.thongKe, tg = kq.tong || {}, K = [];
    var h = function (t, x) { K.push({ t: t, x: x }); }, p = function (x) { if (x) K.push({ t: 'p', x: x }); }, li = function (a, tien) { (a || []).forEach(function (x) { if (x) K.push({ t: 'li', x: (tien || '') + x }); }); };
    var bang = function (rows, rong) { if (rows.length > 1) K.push({ t: 'bang', rows: rows, rong: rong || null }); };
    var dm = function (d) { return d ? d.slice(8, 10) + '/' + d.slice(5, 7) : ''; };
    K.push({ t: 'dau', x: kq.congTy || 'CÔNG TY CỔ PHẦN TAPACO VIỆT NAM', y: 'Ngày lập: ' + ngayVN(String(kq.taoLuc).slice(0, 10)) });
    h('h1', 'BÁO CÁO PHÂN TÍCH, ĐÁNH GIÁ TỔNG THỂ');
    K.push({ t: 'giua', x: (kq.nhom.loai === 'rieng' ? 'Trò chuyện: ' : 'Nhóm: ') + kq.nhom.ten + ' · Từ ' + ngayVN(kq.tu) + ' đến ' + ngayVN(kq.den) });
    p('Người lập: ' + kq.taoBoiTen + ' · Lập lúc: ' + kq.taoLuc + ' · Phân tích: AI ' + kq.ai.ncc + ' (' + kq.ai.model + ')');
    K.push({ t: 'nho', x: 'Số liệu thống kê (tin nhắn, ngày tương tác, ngày nghỉ có đơn, báo cáo định kỳ) do hệ thống tính. Phần tổng hợp, trích xuất, nhận xét do AI lập từ nội dung trao đổi, cần đối chiếu chứng từ gốc trước khi sử dụng.' });
    h('h2', 'I. BỨC TRANH TỔNG THỂ'); p(tg.tongThe || '(AI không trả về nội dung)');
    if (tg.mucRuiRo) p('Mức độ rủi ro cho công ty: ' + tg.mucRuiRo.toUpperCase() + (tg.lyDoRuiRo ? ' – ' + tg.lyDoRuiRo : ''));
    var nopP = 0, nopD = 0, nopH = 0; tk.baoCao.forEach(function (b) { nopP += b.phaiNop; nopD += b.daNop; nopH += b.dungHan; });
    bang([['Chỉ tiêu', 'Giá trị'], ['Số ngày trong kỳ / ngày làm việc / ngày có hoạt động', tk.soNgay + ' / ' + tk.soNgayLV + ' / ' + tk.soNgayHD], ['Tin nhắn · ảnh · tệp', tk.tong.soTin + ' · ' + tk.tong.soAnh + ' · ' + tk.tong.soTep],
      ['Thành viên có hoạt động', tk.tong.soNguoi + ' / ' + tk.tong.soTV], ['Báo cáo định kỳ đã nộp / phải nộp (đúng hạn)', nopD + ' / ' + nopP + (nopP ? ' (' + Math.round(nopD / nopP * 100) + '%)' : '') + ' – ' + nopH + ' đúng hạn'],
      ['Ngày nghỉ có đơn (cả nhóm)', aiSoTxt(tk.caNhan.reduce(function (s, c) { return s + c.soNgayNghi; }, 0))], ['Lượt ngày làm việc không tương tác (cả nhóm)', String(tk.caNhan.reduce(function (s, c) { return s + c.soNgayKhongTT; }, 0))]], [60, 40]);
    var dg = tg.danhGia || [];
    if (dg.length) {
      h('h2', 'II. ĐÁNH GIÁ CHI TIẾT THEO KHÍA CẠNH');
      bang([['Khía cạnh', 'Mức']].concat(dg.map(function (x) { return [x.khia, x.muc]; })), [70, 30]);
      dg.forEach(function (x, i) {
        h('h3', (i + 1) + '. ' + x.khia + ' – ' + x.muc); p(x.nhanXet);
        if ((x.danChung || []).length) li(x.danChung, 'Dẫn chứng: ');
        if ((x.giaiPhap || []).length) li(x.giaiPhap, 'Giải pháp: ');
      });
    }
    h('h2', 'III. TỒN TẠI VÀ GIẢI PHÁP XỬ LÝ');
    if ((tg.diemManh || []).length) { h('h3', 'Điểm mạnh'); li(tg.diemManh); }
    if ((tg.tonTai || []).length) { h('h3', 'Tồn tại'); li(tg.tonTai); }
    bang([['TT', 'Vấn đề', 'Giải pháp', 'Phụ trách', 'Thời hạn', 'Ưu tiên']].concat((tg.giaiPhap || []).map(function (x, i) { return [String(i + 1), x.vanDe, x.giaiPhap, x.phuTrach, x.thoiHan, x.uuTien]; })), [5, 25, 35, 15, 10, 10]);
    if ((tg.kienNghiBGD || []).length) { h('h3', 'Kiến nghị Giám đốc chỉ đạo'); li(tg.kienNghiBGD); }
    h('h2', 'IV. ĐÁNH GIÁ NHÂN SỰ');
    K.push({ t: 'nho', x: 'Ngày LV: ngày làm việc trong kỳ kể từ khi tham gia (trừ Chủ nhật, lễ). Không TT: ngày làm việc không có tin nhắn, tệp, báo cáo, bình chọn và không nghỉ có đơn. Điểm tham gia (0–100) do hệ thống tính: 50% tỷ lệ ngày có tương tác, 30% tỷ lệ nộp báo cáo đúng hạn, 20% mức đóng góp so với người nhiều nhất.' });
    bang([['Họ tên', 'Vai trò', 'Tin/ảnh/tệp', 'Ngày LV', 'Có TT', 'Không TT', 'Nghỉ có đơn', 'BC nộp/phải (đúng hạn)', 'Điểm TG', 'Xếp loại']].concat(kq.caNhan.map(function (c) {
      return [c.hoTen, c.vaiTro, c.soTin + '/' + c.soAnh + '/' + c.soTep, String(c.soNgayLV), String(c.soNgayCoTT), String(c.soNgayKhongTT), aiSoTxt(c.soNgayNghi), c.bcPhai || c.bcNop ? c.bcNop + '/' + c.bcPhai + ' (' + c.bcDungHan + ')' : '—', String(c.diemTG), c.xepLoai];
    })), [17, 11, 9, 7, 7, 8, 8, 13, 8, 12]);
    kq.caNhan.forEach(function (c) {
      var tt = [c.nhanXet, c.deXuat ? 'Đề xuất: ' + c.deXuat : ''].filter(Boolean).join(' ');
      if (!tt && !c.dongGop.length && !c.kienNghi.length && !c.canLuuY.length) return;
      h('h3', c.hoTen + ' – ' + c.xepLoai); p(tt);
      li(c.dongGop, '✓ '); li(c.canLuuY, '⚠ ');
      li(c.kienNghi.map(function (k) { return dm(k.ngay) + ': ' + k.noiDung; }), 'Kiến nghị ');
      if (c.ngayKhongTT.length) K.push({ t: 'nho', x: 'Ngày không tương tác: ' + c.ngayKhongTT.map(dm).join(', ') + (c.soNgayKhongTT > c.ngayKhongTT.length ? '…' : '') });
    });
    h('h2', 'V. KHỐI LƯỢNG, SẢN LƯỢNG');
    var soBC = []; tk.baoCao.forEach(function (b) { b.so.forEach(function (f) { if (f.soLan) soBC.push([b.tieuDe, f.ten, aiSoTxt(f.tong), aiSoTxt(f.tb), String(f.soLan)]); }); });
    if (soBC.length) { h('h3', 'Số liệu từ báo cáo định kỳ (hệ thống cộng)'); bang([['Báo cáo', 'Chỉ tiêu', 'Tổng', 'TB / lần', 'Số lần']].concat(soBC)); }
    if (kq.khoiLuongTong.length) { h('h3', 'Khối lượng thi công theo nhật ký'); bang([['Hạng mục', 'ĐVT', 'Tổng thực hiện', 'Tổng kế hoạch nêu', 'Số ngày', 'Các ngày']].concat(kq.khoiLuongTong.map(function (g) { return [g.hangMuc, g.dvt, g.coSo ? aiSoTxt(g.tong) : '—', g.keHoach ? aiSoTxt(g.keHoach) : '', String(g.lan), g.ngay.map(dm).join(', ')]; }))); }
    if (!soBC.length && !kq.khoiLuongTong.length) p('Không ghi nhận khối lượng trong kỳ.');
    h('h2', 'VI. VẬT TƯ, THIẾT BỊ');
    if (kq.vatTuTong.length) { bang([['Vật tư', 'ĐVT', 'Nhập', 'Xuất', 'Đề nghị cấp', 'Nhập − Xuất', 'Các ngày']].concat(kq.vatTuTong.map(function (g) { return [g.ten, g.dvt, g.coNhap ? aiSoTxt(g.nhap) : '—', g.coXuat ? aiSoTxt(g.xuat) : '—', g.coDeNghi ? aiSoTxt(g.deNghi) : '—', g.chenh == null ? '—' : aiSoTxt(g.chenh), g.ngay.map(dm).join(', ')]; })));
      K.push({ t: 'nho', x: 'Nhập, Xuất, Đề nghị cấp cộng theo các lần được nêu trong nhóm. "Nhập − Xuất" là chênh lệch của những gì đã nêu, KHÔNG phải tồn kho; đối chiếu phiếu nhập kho, phiếu xuất kho để chốt số.' }); } else p('Không ghi nhận thông tin vật tư.');
    if (kq.thietBi.length) { h('h3', 'Tình trạng thiết bị được nêu'); bang([['Ngày', 'Thiết bị', 'SL', 'Tình trạng', 'Ghi chú']].concat(kq.thietBi.map(function (x) { return [ngayVN(x.ngay), x.ten, x.soLuong == null ? '' : aiSoTxt(x.soLuong), x.tinhTrang, x.ghiChu]; }))); }
    h('h2', 'VII. NHÀ THẦU PHỤ, NHÀ CUNG CẤP');
    if (kq.doiTac.length) bang([['Đơn vị', 'Loại', 'Đánh giá', 'Nhận xét', 'Đề xuất', 'Số lần nêu']].concat(kq.doiTac.map(function (x) { return [x.ten, x.loai, x.danhGia, x.nhanXet, x.deXuat, String(x.lan)]; })));
    else p('Không ghi nhận thông tin nhà thầu phụ, nhà cung cấp.');
    if (kq.suKien.length) { h('h2', 'VIII. SỰ CỐ, SỰ KIỆN ĐÁNG CHÚ Ý'); bang([['Ngày', 'Loại', 'Mức độ', 'Nội dung']].concat(kq.suKien.map(function (x) { return [ngayVN(x.ngay), x.loai, x.mucDo, x.noiDung]; })), [12, 15, 12, 61]); }
    h('h2', (kq.suKien.length ? 'IX' : 'VIII') + '. CHỨNG TỪ, HÌNH ẢNH KÈM THEO');
    if (kq.chungTu.length) bang([['Ngày', 'Giờ', 'Người gửi', 'Tệp', 'Mô tả']].concat(kq.chungTu.map(function (c) { return [ngayVN(c.ngay), c.gio, c.hoTen, c.ten, [c.moTa, c.lienQuan ? '(' + c.lienQuan + ')' : '', !c.moTa && c.ghiChu ? c.ghiChu : ''].filter(Boolean).join(' ')]; })), [11, 7, 18, 24, 40]);
    else p('Không có ảnh, tệp trong kỳ.');
    (anh || []).forEach(function (a) { K.push({ t: 'anh', mime: a.mime, data: a.data, x: a.nhan, fileId: a.fileId }); });
    K.push({ t: 'nho', x: 'Nhật ký thi công chi tiết từng ngày lập thành tài liệu riêng. Tệp gốc lưu tại thư mục "1. Tệp & ảnh", nhật ký nguyên văn tại "2. Nhật ký từng ngày" của nhóm.' });
    return K;
  }
  /* NHẬT KÝ THI CÔNG từng ngày (tài liệu riêng) */
  function nkKhoi(kq) {
    var K = [], dm = function (d) { return ngayVN(d); };
    K.push({ t: 'dau', x: kq.congTy || 'CÔNG TY CỔ PHẦN TAPACO VIỆT NAM', y: (kq.nhom.loai === 'rieng' ? 'Trò chuyện: ' : 'Nhóm: ') + kq.nhom.ten });
    K.push({ t: 'h1', x: 'NHẬT KÝ THI CÔNG' });
    K.push({ t: 'giua', x: 'Từ ngày ' + dm(kq.tu) + ' đến ngày ' + dm(kq.den) });
    K.push({ t: 'nho', x: 'Tổng hợp từ tin nhắn, báo cáo định kỳ và tệp trong nhóm. Mục để trống: nhóm không có thông tin trong ngày. Số nhân lực, ngày nghỉ có đơn và số tin nhắn do hệ thống tính; nội dung còn lại do AI tổng hợp, cần người phụ trách kiểm tra, ký xác nhận.' });
    kq.nhatKy.forEach(function (n, i) {
      if (i) K.push({ t: 'ngat' });
      K.push({ t: 'h2', x: n.thu + ', ngày ' + dm(n.ngay) });
      if (n.tomTat) K.push({ t: 'p', x: n.tomTat });
      var rows = [['Mục', 'Nội dung']];
      AI_TRUONG_NK.forEach(function (f) { if (f[0] === 'thiCong') return; var v = nkO(n, f[0], kq); rows.push([f[1], v || '—']); });
      rows.splice(4, 0, ['Trao đổi trong nhóm', n.soTin + ' tin, ' + n.soAnh + ' ảnh, ' + n.soTep + ' tệp; ' + n.soNguoi + ' người']);
      K.push({ t: 'bang', rows: rows, rong: [24, 76] });
      K.push({ t: 'h3', x: 'Công tác thi công và khối lượng' });
      if ((n.thiCong || []).length) K.push({ t: 'bang', rows: [['Hạng mục', 'Vị trí', 'KL thực hiện', 'KL kế hoạch', 'So với KH', 'Ghi chú']].concat(n.thiCong.map(function (x) { return [x.hangMuc, x.viTri, aiSL(x.khoiLuong, x.dvt), aiSL(x.keHoach, x.dvt), x.soVoiKH, x.ghiChu]; })), rong: [24, 16, 14, 14, 10, 22] });
      else K.push({ t: 'p', x: '—' });
      K.push({ t: 'ky', x: ['Người lập', 'Chỉ huy trưởng công trình'] });
    });
    return K;
  }
  function tenTepAscii(t) { return String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').replace(/[^\w-]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 40) || 'Nhom'; }
  function aiTen(kq, u) { var c = (kq.caNhan || []).filter(function (x) { return x.username === u; })[0]; return c ? c.hoTen : (u || ''); }
  /* Excel báo cáo phân tích (nhiều trang tính) */
  function xlAiBaoCao(kq) {
    var tk = kq.thongKe, tg = kq.tong || {}, H = function (a) { return a.map(function (x) { return { v: x, s: 'head' }; }); };
    var tieuDe = function (t) { return [[{ v: t, s: 'title' }], [{ v: (kq.nhom.loai === 'rieng' ? 'Trò chuyện: ' : 'Nhóm: ') + kq.nhom.ten + ' · ' + ngayVN(kq.tu) + ' – ' + ngayVN(kq.den) + ' · Lập: ' + kq.taoBoiTen + ', ' + kq.taoLuc, s: 'sub' }], []]; };
    var sh = [], r1 = tieuDe('BÁO CÁO PHÂN TÍCH, ĐÁNH GIÁ TỔNG THỂ');
    r1.push([{ v: 'BỨC TRANH TỔNG THỂ', s: 'grp' }, { v: '', s: 'grp' }, { v: '', s: 'grp' }, { v: '', s: 'grp' }, { v: '', s: 'grp' }]); r1.push([tg.tongThe || '']);
    if (tg.mucRuiRo) r1.push([{ v: 'Mức rủi ro cho công ty: ' + tg.mucRuiRo + (tg.lyDoRuiRo ? ' – ' + tg.lyDoRuiRo : ''), s: 'net' }]);
    r1.push([]); r1.push(H(['Chỉ tiêu', 'Giá trị', '', '', '']));
    var nopP = 0, nopD = 0; tk.baoCao.forEach(function (b) { nopP += b.phaiNop; nopD += b.daNop; });
    [['Số ngày trong kỳ', tk.soNgay], ['Ngày làm việc', tk.soNgayLV], ['Ngày có hoạt động', tk.soNgayHD], ['Tin nhắn', tk.tong.soTin], ['Ảnh', tk.tong.soAnh], ['Tệp', tk.tong.soTep], ['Thành viên có hoạt động', tk.tong.soNguoi], ['Tổng thành viên', tk.tong.soTV], ['Báo cáo định kỳ phải nộp', nopP], ['Báo cáo định kỳ đã nộp', nopD]].forEach(function (x) { r1.push([x[0], x[1], '', '', '']); });
    r1.push([]); r1.push(H(['Khía cạnh', 'Mức', 'Nhận xét', 'Dẫn chứng', 'Giải pháp']));
    (tg.danhGia || []).forEach(function (x) { r1.push([x.khia, x.muc, x.nhanXet, (x.danChung || []).join('\n'), (x.giaiPhap || []).join('\n')]); });
    [['Điểm mạnh', tg.diemManh], ['Tồn tại', tg.tonTai], ['Kiến nghị Giám đốc chỉ đạo', tg.kienNghiBGD]].forEach(function (g) { if ((g[1] || []).length) { r1.push([]); r1.push([{ v: g[0], s: 'bold' }]); g[1].forEach(function (x) { r1.push(['• ' + x]); }); } });
    r1.push([]); r1.push([{ v: 'Số liệu thống kê do hệ thống tính; tổng hợp, trích xuất, nhận xét do AI (' + kq.ai.ncc + ' – ' + kq.ai.model + ') lập, cần đối chiếu chứng từ gốc.', s: 'note' }]);
    sh.push({ ten: 'Tổng quan', rows: r1, cols: [36, 18, 70, 50, 50] });
    var r8 = tieuDe('TỒN TẠI VÀ GIẢI PHÁP XỬ LÝ'); r8.push(H(['TT', 'Vấn đề', 'Giải pháp', 'Phụ trách', 'Thời hạn', 'Ưu tiên']));
    (tg.giaiPhap || []).forEach(function (x, i) { r8.push([i + 1, x.vanDe, x.giaiPhap, x.phuTrach, x.thoiHan, x.uuTien]); });
    sh.push({ ten: 'Giải pháp', rows: r8, cols: [5, 45, 60, 22, 14, 12] });
    var r3 = tieuDe('ĐÁNH GIÁ NHÂN SỰ'); r3.push(H(['Họ tên', 'Tài khoản', 'Vai trò', 'Tin', 'Ảnh', 'Tệp', 'Ngày LV', 'Ngày có TT', 'Ngày không TT', 'Ngày nghỉ có đơn', 'BC phải nộp', 'BC đã nộp', 'Đúng hạn', 'Tỷ lệ nộp', 'Điểm tham gia', 'Xếp loại', 'Nhận xét', 'Đề xuất', 'Đóng góp', 'Kiến nghị', 'Ngày không TT']));
    kq.caNhan.forEach(function (c) { var r = r3.length + 1; r3.push([c.hoTen, c.username, c.vaiTro, c.soTin, c.soAnh, c.soTep, c.soNgayLV, c.soNgayCoTT, c.soNgayKhongTT, c.soNgayNghi, c.bcPhai, c.bcNop, c.bcDungHan, c.bcPhai ? { f: 'IF(K' + r + '>0,MIN(1,L' + r + '/K' + r + '),"")', v: 0, z: '0%' } : '', c.diemTG, c.xepLoai, c.nhanXet, c.deXuat, c.dongGop.join('\n'), c.kienNghi.map(function (k) { return ngayVN(k.ngay) + ': ' + k.noiDung; }).join('\n'), c.ngayKhongTT.map(ngayVN).join(', ')]); });
    sh.push({ ten: 'Nhân sự', rows: r3, cols: [24, 14, 14, 6, 6, 6, 8, 9, 10, 10, 9, 9, 9, 9, 10, 14, 50, 30, 50, 50, 40] });
    var cotNK = AI_TRUONG_NK.map(function (f) { return f[1]; });
    var r2 = tieuDe('NHẬT KÝ THI CÔNG'); r2.push(H(['Ngày', 'Thứ', 'Tin', 'Ảnh', 'Tệp', 'Tóm tắt'].concat(cotNK)));
    var d0 = r2.length + 1;
    kq.nhatKy.forEach(function (n) { r2.push([ngayVN(n.ngay), n.thu, n.soTin, n.soAnh, n.soTep, n.tomTat].concat(AI_TRUONG_NK.map(function (f) { return nkO(n, f[0], kq); }))); });
    var d1 = r2.length;
    if (kq.nhatKy.length) r2.push([{ v: 'Cộng', s: 'tot' }, { v: '', s: 'tot' }].concat(['C', 'D', 'E'].map(function (c) { return { f: 'SUM(' + c + d0 + ':' + c + d1 + ')', s: 'tot' }; })));
    sh.push({ ten: 'Nhật ký thi công', rows: r2, cols: [11, 9, 5, 5, 5, 40].concat(cotNK.map(function () { return 32; })), freeze: { xSplit: 1, ySplit: 4 } });
    var r5 = tieuDe('KHỐI LƯỢNG, SẢN LƯỢNG'), soBC = []; tk.baoCao.forEach(function (b) { b.so.forEach(function (f) { if (f.soLan) soBC.push([b.tieuDe, f.ten, f.tong, f.tb, f.soLan]); }); });
    if (soBC.length) { r5.push([{ v: 'Số liệu từ báo cáo định kỳ (hệ thống cộng)', s: 'bold' }]); r5.push(H(['Báo cáo', 'Chỉ tiêu', 'Tổng', 'Trung bình / lần', 'Số lần'])); soBC.forEach(function (x) { r5.push(x); }); r5.push([]); }
    r5.push([{ v: 'Tổng theo hạng mục (nhật ký thi công)', s: 'bold' }]); r5.push(H(['Hạng mục', 'ĐVT', 'Tổng thực hiện', 'Tổng kế hoạch nêu', 'Số ngày', 'Các ngày']));
    kq.khoiLuongTong.forEach(function (g) { r5.push([g.hangMuc, g.dvt, g.coSo ? g.tong : '', g.keHoach || '', g.lan, g.ngay.map(ngayVN).join(', ')]); });
    r5.push([]); r5.push(H(['Ngày', 'Hạng mục', 'Vị trí', 'Thực hiện', 'Kế hoạch', 'ĐVT', 'So với KH', 'Ghi chú', 'Trích dẫn']));
    kq.khoiLuong.forEach(function (x) { r5.push([ngayVN(x.ngay), x.hangMuc, x.viTri, x.khoiLuong == null ? '' : x.khoiLuong, x.keHoach == null ? '' : x.keHoach, x.dvt, x.soVoiKH, x.ghiChu, x.trich]); });
    sh.push({ ten: 'Khối lượng', rows: r5, cols: [26, 30, 18, 12, 12, 8, 10, 30, 50] });
    var r4 = tieuDe('VẬT TƯ, THIẾT BỊ'); r4.push(H(['Vật tư', 'ĐVT', 'Nhập', 'Xuất', 'Đề nghị cấp', 'Nhập − Xuất', 'Số lần nêu', 'Các ngày']));
    kq.vatTuTong.forEach(function (g) { r4.push([g.ten, g.dvt, g.coNhap ? g.nhap : '', g.coXuat ? g.xuat : '', g.coDeNghi ? g.deNghi : '', g.chenh == null ? '' : g.chenh, g.lan, g.ngay.map(ngayVN).join(', ')]); });
    r4.push([]); r4.push(H(['Ngày', 'Vật tư', 'ĐVT', 'Nhập', 'Xuất', 'Đề nghị cấp', 'Nhà cung cấp / hạng mục / người', 'Trích dẫn']));
    kq.vatTu.forEach(function (x) { var v = x.soLuong == null ? '' : x.soLuong; r4.push([ngayVN(x.ngay), x.ten, x.dvt, x.hoatDong === 'nhập' ? v : '', x.hoatDong === 'xuất' ? v : '', x.hoatDong === 'đề nghị' ? v : '', x.doiTuong || '', x.trich]); });
    if (kq.thietBi.length) { r4.push([]); r4.push(H(['Ngày', 'Thiết bị', 'Số lượng', 'Tình trạng', 'Ghi chú', '', '', ''])); kq.thietBi.forEach(function (x) { r4.push([ngayVN(x.ngay), x.ten, x.soLuong == null ? '' : x.soLuong, x.tinhTrang, x.ghiChu, '', '', '']); }); }
    sh.push({ ten: 'Vật tư, thiết bị', rows: r4, cols: [14, 30, 10, 11, 11, 13, 30, 60] });
    var r9 = tieuDe('NHÀ THẦU PHỤ, NHÀ CUNG CẤP'); r9.push(H(['Đơn vị', 'Loại', 'Đánh giá', 'Nhận xét', 'Đề xuất', 'Số lần nêu', 'Các ghi nhận']));
    kq.doiTac.forEach(function (x) { r9.push([x.ten, x.loai, x.danhGia, x.nhanXet, x.deXuat, x.lan, x.ghiNhan.map(function (g) { return ngayVN(g.ngay) + ' (' + g.danhGia + '): ' + g.noiDung; }).join('\n')]); });
    sh.push({ ten: 'Thầu phụ, NCC', rows: r9, cols: [26, 14, 12, 50, 40, 10, 70] });
    var r6 = tieuDe('CHỨNG TỪ, HÌNH ẢNH KÈM THEO'); r6.push(H(['Ngày', 'Giờ', 'Người gửi', 'Loại', 'Tên tệp', 'Ghi chú khi gửi', 'Mô tả (AI)', 'Liên quan', 'Mã tệp (Drive)']));
    kq.chungTu.forEach(function (c) { r6.push([ngayVN(c.ngay), c.gio, c.hoTen, c.loai === 'anh' ? 'Ảnh' : 'Tệp', c.ten, c.ghiChu, c.moTa, c.lienQuan, c.fileId]); });
    sh.push({ ten: 'Chứng từ', rows: r6, cols: [11, 7, 22, 6, 30, 40, 50, 24, 30] });
    if (kq.suKien.length) { var r7 = tieuDe('SỰ CỐ, SỰ KIỆN ĐÁNG CHÚ Ý'); r7.push(H(['Ngày', 'Loại', 'Mức độ', 'Nội dung', 'Trích dẫn'])); kq.suKien.forEach(function (x) { r7.push([ngayVN(x.ngay), x.loai, x.mucDo, x.noiDung, x.trich]); }); sh.push({ ten: 'Sự kiện', rows: r7, cols: [12, 14, 12, 70, 50] }); }
    return { ten: 'PhanTich_' + tenTepAscii(kq.nhom.ten) + '_' + kq.tu + '_' + kq.den, sheets: sh };
  }
  /* Excel nhật ký từng ngày (toàn văn tin nhắn + báo cáo định kỳ) */
  function xlNhatKy(nk) {
    var sh = [], r = [[{ v: 'NHẬT KÝ NHÓM: ' + nk.nhom, s: 'title' }], [{ v: 'Từ ' + ngayVN(nk.tu) + ' đến ' + ngayVN(nk.den) + ' · tin đã thu hồi không ghi nội dung', s: 'sub' }], [],
      ['Giờ', 'Người', 'Loại', 'Nội dung', 'Tệp'].map(function (x) { return { v: x, s: 'head' }; })];
    nk.ngay.forEach(function (n) {
      r.push([{ v: n.thu + ', ' + ngayVN(n.ngay) + ' – ' + n.dong.length + ' dòng', s: 'grp' }, { v: '', s: 'grp' }, { v: '', s: 'grp' }, { v: '', s: 'grp' }, { v: '', s: 'grp' }]);
      n.dong.forEach(function (d) { r.push([d.gio, d.hoTen, d.loai, d.noiDung, d.tep]); });
    });
    sh.push({ ten: 'Nhật ký', rows: r, cols: [7, 22, 14, 100, 30], freeze: { xSplit: 0, ySplit: 4 } });
    nk.bc.forEach(function (b, i) {
      var rr = [[{ v: b.tieuDe, s: 'title' }], [{ v: 'Từ ' + ngayVN(nk.tu) + ' đến ' + ngayVN(nk.den), s: 'sub' }], [], b.cot.map(function (x) { return { v: x, s: 'head' }; })], d0 = rr.length + 1;
      b.dong.forEach(function (d) { rr.push(d.map(function (v, j) { if (j === 0) return ngayVN(v); if (j >= 4 && b.so[j - 4] && v !== '' && isFinite(+v)) return +v; return v; })); });
      var d1 = rr.length;
      if (b.dong.length && b.so.some(Boolean)) rr.push([{ v: 'Cộng', s: 'tot' }, { v: '', s: 'tot' }, { v: '', s: 'tot' }, { v: '', s: 'tot' }].concat(b.so.map(function (laSo, j) { var c = cotXL(j + 4); return laSo ? { f: 'SUM(' + c + d0 + ':' + c + d1 + ')', s: 'tot' } : { v: '', s: 'tot' }; })));
      sh.push({ ten: ('BC ' + (i + 1) + ' ' + b.tieuDe).replace(/[\\\/?*\[\]:]/g, ' ').slice(0, 31), rows: rr, cols: [12, 22, 18, 9].concat(b.so.map(function (s) { return s ? 12 : 30; })) });
    });
    return { ten: 'NhatKy_' + tenTepAscii(nk.nhom) + '_' + nk.tu + '_' + nk.den, sheets: sh };
  }

  /* ==================== API ==================== */
  function createApi(env) {
    // env: { db:{table(name),dirty(name)}, hash(str), uuid(), now(), today(), sessions:{get,put,del}, log(row), attempts:{get,put} }
    var db = env.db;
    function T(n) { return db.table(n); }
    /* v3.14: lượt chấm công GPS ghi thêm theo thời gian (thoiGian tăng dần theo dòng) – chỉ cần đọc phần cuối từ mốc tu (yyyy-mm-dd).
       Lượt có ngày ≥ tu luôn được ghi lúc ≥ tu nên không sót. */
    function gpsTu(tu) { return db.tail ? db.tail('ChamCongGPS', 'thoiGian', String(tu || '')) : T('ChamCongGPS'); }
    function D(n) { db.dirty(n); }
    function err(m) { throw new Error(m); }
    function newId() { return env.uuid().replace(/-/g, '').slice(0, 12); }

    var _cfgNho = null, _cfgSig = null;                // v3.14: nhớ cấu hình trong 1 yêu cầu (getCfg gọi nhiều lần); đổi CauHinh thì tính lại
    function getCfg() {
      var rows = T('CauHinh'), sig = rows.length + '|' + rows.map(function (r) { return r.key + ':' + String(r.value).length; }).join(',');
      if (_cfgNho && _cfgSig === sig && _cfgNho.rows === rows) { var c0 = clone(_cfgNho.c); apDungPQ(c0.phanQuyen); return c0; }
      var st = {};
      rows.forEach(function (r) { st[r.key] = parseJSON(r.value, r.value); });
      var c = mergeCfg(st); apDungPQ(c.phanQuyen);
      _cfgNho = { c: clone(c), rows: rows }; _cfgSig = sig;
      return c;
    }
    function setCfg(obj) {
      var rows = T('CauHinh');
      Object.keys(obj).forEach(function (k) {
        if (!(k in DEFAULT_CFG)) return;
        var v = JSON.stringify(obj[k]);
        var r = rows.filter(function (x) { return x.key === k; })[0];
        if (r) r.value = v; else rows.push({ key: k, value: v });
      });
      _cfgNho = null; D('CauHinh');
    }

    function log(s, action, ct) {
      try { env.log({ thoiGian: env.now(), user: s ? s.username : '', action: action, chiTiet: typeof ct === 'string' ? ct : JSON.stringify(ct || '') }); } catch (e) { }
    }
    /* v3.27: nhật ký thay đổi có giá trị trước / sau (người sửa, thời gian ghi ở env.log) */
    function logDoi(s, action, ma, truoc, sau, bo) {
      var doi = [], ks = {}; bo = bo || ['capNhat', 'passHash', 'salt'];
      Object.keys(truoc || {}).concat(Object.keys(sau || {})).forEach(function (k) { ks[k] = 1; });
      Object.keys(ks).forEach(function (k) {
        if (bo.indexOf(k) >= 0) return;
        var a = truoc ? truoc[k] : undefined, b = sau ? sau[k] : undefined, ta = typeof a === 'object' ? JSON.stringify(a) : String(a == null ? '' : a), tb = typeof b === 'object' ? JSON.stringify(b) : String(b == null ? '' : b);
        if (ta !== tb) doi.push(k + ': ' + (ta.length > 300 ? ta.slice(0, 300) + '…' : ta || '∅') + ' → ' + (tb.length > 300 ? tb.slice(0, 300) + '…' : tb || '∅'));
      });
      log(s, action, (ma ? ma + ' | ' : '') + (truoc ? '' : '[tạo mới] ') + (doi.join(' ; ') || 'không đổi').slice(0, 4000));
      return doi.length;
    }

    function findNV(ma) { return T('NhanVien').filter(function (r) { return r.maNV === ma; })[0]; }
    function bpXem(s) {                      // v3.38: bộ phận được xem theo quyền xem_bp (TP/CHT: bộ phận phụ trách; chức danh khác: phòng ban của mình)
      if (!s || !can(s.role, 'xem_bp')) return '';
      if (ROLE_QL.indexOf(s.role) >= 0) return s.boPhan || '';
      var n = s.maNV && findNV(s.maNV); return n && n.trangThai !== 'Đã nghỉ việc' ? (n.phongBan || '') : '';
    }
    function hashPw(pw, salt) { return env.hash(salt + '|' + pw); }
    function publicUser(u) { return { username: u.username, role: normRole(u.role), maNV: u.maNV, hoTen: u.hoTen, boPhan: boPhanCua(u), active: u.active, mustChange: u.mustChange, lastLogin: u.lastLogin, createdAt: u.createdAt, khoaTu: u.khoaTu || '', khoaLyDo: u.khoaLyDo || '' }; }
    /* v3.18: chấm dứt hợp đồng (Nghỉ việc / Thanh lý hợp đồng trong Quá trình công tác) → khóa tài khoản từ ngày chấm dứt.
       Bản ghi chấm dứt còn hiệu lực = bản ghi mới nhất trong các bản ghi Tiếp nhận / chấm dứt (tiếp nhận lại sau đó thì hết hiệu lực). */
    function chamDutCua(maNV) {
      var ds = T('QuaTrinh').filter(function (x) { return x.maNV === maNV && (QT_CHAM_DUT.indexOf(x.loai) >= 0 || x.loai === 'Tiếp nhận'); })
        .sort(function (a, b) { return a.tuNgay < b.tuNgay ? -1 : a.tuNgay > b.tuNgay ? 1 : (a.taoLuc < b.taoLuc ? -1 : 1); });
      var c = ds[ds.length - 1]; return c && QT_CHAM_DUT.indexOf(c.loai) >= 0 ? c : null;
    }
    function khoaTK(nv, lyDo) {                       // đồng bộ khóa tài khoản theo ngày chấm dứt của hồ sơ (không khóa tài khoản Quản trị)
      var hom = env.today(), ngay = nv.ngayNghi || (nv.trangThai === 'Đã nghỉ việc' ? hom : ''), doi = false;
      T('Users').forEach(function (u) {
        if (u.maNV !== nv.maNV || normRole(u.role) === 'admin') return;
        if (ngay && u.active === 'Không' && /^Khóa bởi/.test(u.khoaLyDo || '')) { if (u.khoaTu !== ngay) { u.khoaTu = ngay; doi = true; } }   // đang khóa tay: giữ nguyên lý do
        else if (ngay && u.active !== 'Không' && /^Mở bởi/.test(u.khoaLyDo || '')) { }                                                    // Quản trị đã mở tay sau khi nghỉ: giữ mở
        else if (ngay) {
          var ly = 'Tự động: ' + (lyDo || 'chấm dứt hợp đồng') + ' từ ' + ngayVN(ngay);
          if (u.khoaTu !== ngay || !/^Tự động/.test(u.khoaLyDo || '')) { u.khoaTu = ngay; u.khoaLyDo = ly; doi = true; }
          if (ngay <= hom && u.active !== 'Không') { u.active = 'Không'; doi = true; }
        } else if (u.khoaTu || /^Tự động/.test(u.khoaLyDo || '')) {
          if (u.active === 'Không' && /^Tự động/.test(u.khoaLyDo || '')) u.active = 'Có';      // hủy / tiếp nhận lại: mở khóa (khóa tay của Quản trị giữ nguyên)
          u.khoaTu = ''; if (/^(Tự động|Mở bởi)/.test(u.khoaLyDo || '')) u.khoaLyDo = ''; doi = true;
        }
      });
      if (doi) D('Users');
    }
    function apChamDut(nv) {
      var c = chamDutCua(nv.maNV), hom = env.today();
      if (c) { nv.ngayNghi = c.tuNgay; if (c.tuNgay <= hom) nv.trangThai = 'Đã nghỉ việc'; else if (nv.trangThai === 'Đã nghỉ việc') nv.trangThai = 'Đang làm việc'; }
      else if (nv.ngayNghi || nv.trangThai === 'Đã nghỉ việc') { nv.ngayNghi = ''; if (nv.trangThai === 'Đã nghỉ việc') nv.trangThai = 'Đang làm việc'; }
      nv.capNhat = env.now(); D('NhanVien');
      khoaTK(nv, c ? c.loai.toLowerCase() : '');
    }
    function chamDutDenHan() {                        // ngày chấm dứt (khai trước) đã đến: chuyển trạng thái, khóa tài khoản
      var hom = env.today();
      var tk = {}; T('Users').forEach(function (u) { if (u.maNV && u.active !== 'Không' && !u.khoaTu && !u.khoaLyDo && normRole(u.role) !== 'admin') tk[u.maNV] = 1; });
      T('NhanVien').forEach(function (nv) {
        if (nv.ngayNghi && nv.ngayNghi <= hom && nv.trangThai !== 'Đã nghỉ việc') { nv.trangThai = 'Đã nghỉ việc'; nv.capNhat = env.now(); D('NhanVien'); khoaTK(nv, 'chấm dứt hợp đồng'); }
        else if (nv.trangThai === 'Đã nghỉ việc' && tk[nv.maNV]) khoaTK(nv, (chamDutCua(nv.maNV) || { loai: 'nghỉ việc' }).loai.toLowerCase());   // hồ sơ nghỉ việc từ trước bản 3.18: khóa tài khoản còn mở
      });
    }
    function tkBiKhoa(u) {                            // khóa tự động đến hạn (kiểm tra mỗi lần đăng nhập / thao tác)
      if (u.khoaTu && u.khoaTu <= env.today() && u.active !== 'Không' && normRole(u.role) !== 'admin') { u.active = 'Không'; D('Users'); }
      return u.active === 'Không';
    }
    /* bộ phận của tài khoản: bộ phận quản lý khai trên tài khoản, nếu trống thì bộ phận hiện tại trong hồ sơ */
    function boPhanCua(u) { if (u.boPhan) return u.boPhan; var nv = u.maNV && findNV(u.maNV); return nv ? (nv.phongBan || '') : ''; }
    function qlBPMap() {   // {bộ phận: [Trưởng phòng / CHT]} – để hiện người duyệt trước khi gửi đơn
      var m = {};
      T('Users').forEach(function (u) { var r = normRole(u.role); if (u.active === 'Không' || ROLE_QL.indexOf(r) < 0) return; var b = boPhanCua(u); (m[b] = m[b] || []).push({ username: u.username, hoTen: u.hoTen, maNV: u.maNV, role: r }); });
      return m;
    }
    function quanLyCua(tenBP, boQuaMaNV) {      // tài khoản Trưởng phòng / CHT đang quản lý bộ phận
      return T('Users').filter(function (u) { return u.active !== 'Không' && ROLE_QL.indexOf(normRole(u.role)) >= 0 && boPhanCua(u) === tenBP && (!boQuaMaNV || u.maNV !== boQuaMaNV); });
    }
    /* tên đăng nhập theo họ tên: "Lê Văn Thắng" → "thanglv" (trùng thì thêm số) */
    /* áp dụng các quyết định lên lương đã duyệt khi đến ngày hiệu lực */
    function applyDueRaises() {
      var today = env.today(), changed = false;
      T('LenLuong').filter(function (r) { return r.trangThai === 'Đã duyệt' && r.apDung !== 'Có' && r.hieuLuc && r.hieuLuc <= today; })
        .sort(function (a, b) { return a.hieuLuc < b.hieuLuc ? -1 : 1; })
        .forEach(function (r) {
          var nv = findNV(r.maNV); if (!nv) return;
          if (r.ngachMoi) nv.ngach = r.ngachMoi;
          if (r.bacMoi) nv.bac = r.bacMoi;
          if (r.loai === 'Nâng bậc định kỳ' || r.loai === 'Nâng lương trước hạn') nv.ngayHuongBac = r.hieuLuc;
          nv.capNhat = env.now();
          r.apDung = 'Có'; changed = true;
        });
      if (changed) { D('LenLuong'); D('NhanVien'); }
      T('NhanVien').forEach(function (nv) { dongBoMucHuong(nv); });
      dongBoViTri();
    }
    /* hồ sơ luôn hiển thị mức đang hưởng hôm nay */
    function dongBoMucHuong(nv) {
      var m = mucHuong(nv, T('LenLuong'), T('MucHuong'), env.today()), pcs = JSON.stringify(m.phuCap);
      var cu = JSON.stringify(parseJSON(nv.phuCap, {}));
      if (num(nv.luongChinh) === m.luongChinh && num(nv.luongBH) === m.luongBH && num(nv.soNPT) === m.soNPT && num(nv.anCa) === m.anCa && cu === pcs) return false;
      nv.luongChinh = m.luongChinh; nv.luongBH = m.luongBH; nv.soNPT = m.soNPT; nv.anCa = m.anCa; nv.phuCap = pcs; nv.capNhat = env.now();
      D('NhanVien'); return true;
    }

    function sess(req) {
      var s = req.token && env.sessions.get(req.token);
      if (!s) err('AUTH: Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.');
      var u = T('Users').filter(function (x) { return x.username === s.username; })[0];
      if (!u || tkBiKhoa(u)) err('AUTH: Tài khoản đã bị khóa' + (u && u.khoaLyDo ? ' (' + u.khoaLyDo.replace(/^Tự động: /, '') + ')' : '') + '.');
      if (normRole(u.role) !== u.role) { u.role = normRole(u.role); D('Users'); }
      s.role = u.role; s.maNV = u.maNV; s.hoTen = u.hoTen; s.boPhan = boPhanCua(u);
      env.sessions.put(req.token, s);
      return s;
    }
    function need(s, p) { if (!can(s.role, p)) err('Bạn không có quyền thực hiện thao tác này.'); }

    /* số việc chờ, dùng cho chấm đỏ trên menu (danh sách đầy đủ tải riêng qua cv.ds / dx.ds) */
    function coCV(s) {                                // có liên quan tới công văn nào không: để hiện mục trong menu
      return can(s.role, 'congvan') || T('CongVan').some(function (o) { return cvXem(s, o); });
    }
    function demCV(s) {
      return T('CongVan').filter(function (o) {
        if (o.csTrangThai === 'Chờ duyệt' && cvDuyetCSDuoc(s)) return true;
        return ['Đã giao xử lý', 'Đang xử lý'].indexOf(o.trangThai) >= 0 && laNguoiXuLy(s, cvNguoi(o));
      }).length;
    }
    function demDX(s) {
      return T('DeXuat').filter(function (o) {
        if (o.trangThai === 'Chờ duyệt') return o.maNV !== s.maNV && laNguoiXuLy(s, (dxCap(o) || {}).nguoi || []);
        if (['Đã duyệt', 'Đang thực hiện'].indexOf(o.trangThai) >= 0) return laNguoiXuLy(s, parseJSON(o.thucHien, []));
        return false;
      }).length;
    }

    /* -------- tải dữ liệu theo vai trò -------- */
    function kyCatCu() { var t = env.today(), y = +t.slice(0, 4) - 1, m = t.slice(5, 7); return y + '-' + m; }        // mốc: cùng tháng năm trước
    function kyDuyetMap() { var m = {}; T('KyLuong').forEach(function (k) { if (k.trangThai === 'Đã duyệt') m[k.ky] = 1; }); return m; }
    function loadAll(s) {
      var cfg = getCfg();
      T('KyLuong').forEach(function (k) { kyRow(k.ky, false); });
      T('NghiPhep').forEach(function (o) { chuanHoaDon(o, cfg); });
      cfg = trangTriQT(cfg);
      var ma = 'u:' + s.username;
      var giaoNP = function (o) { return parseJSON(o.luongDuyet, []).some(function (c) { return c.nguoi.indexOf(ma) >= 0; }); };
      var giaoYK = function (o) { return ykXuLy(o).indexOf(ma) >= 0; };
      if (!can(s.role, 'view_all')) {
        var me = s.maNV, bpQL = bpXem(s);
        var nvBP = bpQL ? T('NhanVien').filter(function (n) { return n.phongBan === bpQL && n.trangThai !== 'Đã nghỉ việc'; }).map(function (n) { return n.maNV; }) : [];
        var kyDuyet = T('KyLuong').filter(function (k) { return k.trangThai === 'Đã duyệt'; }).map(function (k) { return k.ky; });
        return {
          cfg: pick(cfg, ['mauPhieu', 'congTy', 'congTyTV', 'diaChiCongTy', 'phuCap', 'noiLamViec', 'boPhan', 'phongBan', 'chuKyNangBac', 'giamTruBanThan', 'giamTruNPT', 'tyLeNLD', 'ngayPhepCoBan', 'thamNienCongPhep', 'ngayLe', 'ngayLeDs', 'loaiTamUng', 'quyTrinhNghi', 'quyTrinhYKien', 'nhomDeXuat', 'loaiDeXuat', 'nghiKLToiDaBH', 'truongBoPhan', 'loaiTaiLieu', 'taiLieuBatBuoc', 'loaiNVTuTai', 'dungLuongToiDaMB', 'gps', 'diemCC', 'push', 'viTriTN', 'phanQuyen', 'cheDoBangLuong']),
          nv: T('NhanVien').filter(function (r) { return r.maNV === me; }),
          qt: T('QuaTrinh').filter(function (r) { return r.maNV === me; }),
          ll: T('LenLuong').filter(function (r) { return r.maNV === me && r.trangThai === 'Đã duyệt'; }),
          mh: T('MucHuong').filter(function (r) { return r.maNV === me && r.trangThai === 'Hiệu lực'; }),
          kt: T('KhenThuong').filter(function (r) { return r.maNV === me && r.trangThai === 'Đã duyệt'; }),
          bl: T('BangLuong').filter(function (r) { return r.maNV === me && kyDuyet.indexOf(r.ky) >= 0; }),
          ky: T('KyLuong').filter(function (k) { return k.trangThai === 'Đã duyệt'; }),
          yk: T('YKien').filter(function (r) { return r.maNV === me || giaoYK(r); }),
          np: T('NghiPhep').filter(function (r) { return r.maNV === me || giaoNP(r) || nvBP.indexOf(r.maNV) >= 0; }),
          bpXem: bpQL,
          bpNV: T('NhanVien').filter(function (n) { return nvBP.indexOf(n.maNV) >= 0; }).map(function (n) { return { maNV: n.maNV, hoTen: n.hoTen, chucVu: n.chucVu, phongBan: n.phongBan, dienThoai: n.dienThoai, ngayVao: n.ngayVao, trangThai: n.trangThai, loaiHD: n.loaiHD }; }),
          bpQT: T('QuaTrinh').filter(function (q) { return nvBP.indexOf(q.maNV) >= 0 && q.phongBan === bpQL; }),
          tl: T('TaiLieu').filter(function (r) { return r.maNV === me && r.nvXem !== 'Không'; }),
          hst: T('HoSoThue').filter(function (r) { return r.maNV === me; }),
          hdk: T('HopDongKhoan').filter(function (r) { return r.maNV === me || (bpQL && r.congTrinh === bpQL); }),
          ttk: T('ThanhToanKhoan').filter(function (r) { return r.maNV === me || (bpQL && r.congTrinh === bpQL); }),
          bt: btDs(s), tt: ttDs(s), qlBP: qlBPMap(), soCV: demCV(s), soDX: demDX(s), coCV: coCV(s),
          cc: [], users: []
        };
      }
      // v3.15: lúc đăng nhập / mở app chỉ gửi chấm công, bảng lương 12 tháng gần nhất (và kỳ chưa duyệt); các kỳ cũ hơn tải ngầm sau (dl.cu)
      var cs = kyCatCu(), dd = kyDuyetMap(), cuCC = 0, cuBL = 0, gan = function (tb) { return T(tb).filter(function (x) { var cu = x.ky < cs && dd[x.ky]; if (cu) { if (tb === 'ChamCong') cuCC++; else cuBL++; } return !cu; }); };
      var ccG = gan('ChamCong'), blG = gan('BangLuong');
      return {
        cfg: cfg,
        nv: T('NhanVien'), qt: T('QuaTrinh'), ll: T('LenLuong'), mh: T('MucHuong'), kt: T('KhenThuong'),
        cc: ccG, bl: blG, dlTu: cuCC + cuBL ? cs : '', ky: T('KyLuong'), pd: T('PheDuyet'), yk: T('YKien'), np: T('NghiPhep'),
        bt: btDs(s), tt: ttDs(s), qlBP: qlBPMap(), soCV: demCV(s), soDX: demDX(s), coCV: coCV(s),
        tl: can(s.role, 'tailieu') ? T('TaiLieu') : T('TaiLieu').filter(function (r) { return r.maNV === s.maNV && r.nvXem !== 'Không'; }),
        users: (can(s.role, 'users') || can(s.role, 'quytrinh')) ? T('Users').map(publicUser) : [],
        hst: T('HoSoThue'), hdk: T('HopDongKhoan'), ttk: T('ThanhToanKhoan')
      };
    }

    /* -------- tạo tài khoản -------- */
    function upsertUser(s, p) {
      var username = String(p.username || '').trim().toLowerCase();
      if (!/^[a-z0-9._-]{3,40}$/.test(username)) err('Tên đăng nhập 3–40 ký tự: chữ không dấu, số, dấu . _ -');
      var role = normRole(p.role || 'nhanvien');
      if (p.role && !ROLES[p.role] && !ROLE_CU[p.role]) err('Chức danh không hợp lệ.');
      if (s.role !== 'admin' && role !== 'nhanvien') err('TCHC chỉ tạo được tài khoản chức danh Nhân viên. Chức danh khác do Quản trị gán.');
      if ((role === 'nhanvien' || ROLE_QL.indexOf(role) >= 0) && !p.maNV) err('Tài khoản phải gắn với một hồ sơ CBCNV.');
      if (role === 'admin' && T('Users').some(function (x) { return normRole(x.role) === 'admin' && x.username !== username; })) err('Hệ thống chỉ có 1 tài khoản Quản trị.');
      var nv = p.maNV ? findNV(p.maNV) : null;
      if (p.maNV && !nv) err('Không tìm thấy mã nhân viên ' + p.maNV);
      var users = T('Users');
      var u = users.filter(function (x) { return x.username === username; })[0];
      if (u && s.role !== 'admin' && normRole(u.role) !== 'nhanvien') err('Không được sửa tài khoản này.');
      if (p.maNV) {
        var dup = users.filter(function (x) { return x.maNV === p.maNV && x.username !== username; })[0];
        if (dup) err('Nhân viên ' + p.maNV + ' đã có tài khoản "' + dup.username + '".');
      }
      if (!u) {
        var pw = p.password || matKhauTheoNV(nv, getCfg());
        if (String(pw).length < 6) err('Mật khẩu tối thiểu 6 ký tự.');
        var salt = newId();
        u = { username: username, passHash: hashPw(pw, salt), salt: salt, mustChange: 'Có', createdAt: env.now(), lastLogin: '' };
        users.push(u);
      } else if (p.password) {
        if (!can(s.role, 'doimk')) err('Chỉ Quản trị hoặc Giám đốc được đổi mật khẩu của tài khoản khác.');
        if (String(p.password).length < 6) err('Mật khẩu tối thiểu 6 ký tự.');
        u.salt = newId(); u.passHash = hashPw(p.password, u.salt); u.mustChange = 'Có';
      }
      if (u.username === s.username && role !== normRole(u.role)) err('Không thể tự đổi chức danh của chính mình.');
      var bp = String(p.boPhan || '').trim();
      if (bp && getCfg().phongBan.indexOf(bp) < 0) err('Bộ phận "' + bp + '" chưa có trong danh mục.');
      if (ROLE_QL.indexOf(role) >= 0 && !bp && !(nv && nv.phongBan)) err('Chọn bộ phận do ' + ROLES[role] + ' quản lý.');
      u.role = role; u.maNV = p.maNV || ''; u.hoTen = nv ? nv.hoTen : (p.hoTen || u.hoTen || '');
      u.boPhan = bp && (!nv || bp !== nv.phongBan || ROLE_QL.indexOf(role) >= 0) ? bp : '';
      var acCu = u.active; u.active = p.active === 'Không' ? 'Không' : 'Có';
        if (u.active !== acCu) {                      // v3.18: Quản trị khóa / mở tay: ghi đè khóa tự động theo ngày chấm dứt HĐ
          if (u.active === 'Không') u.khoaLyDo = 'Khóa bởi ' + (s.hoTen || s.username);
          else { var nvK = u.maNV && findNV(u.maNV); u.khoaLyDo = nvK && (nvK.ngayNghi || nvK.trangThai === 'Đã nghỉ việc') ? 'Mở bởi ' + (s.hoTen || s.username) : ''; u.khoaTu = ''; }
        }
      if (u.username === s.username) u.active = 'Có';
      D('Users');
      return u;
    }

    /* ==================== HANDLERS ==================== */
    var H = {};

    H.ping = function () { return { version: VERSION, time: env.now() }; };

    H.login = function (req) {
      var username = String(req.username || '').trim().toLowerCase();
      var key = 'fail_' + username, fails = num(env.attempts.get(key));
      if (fails >= 5) err('Đăng nhập sai quá 5 lần. Vui lòng thử lại sau 10 phút.');
      var u = T('Users').filter(function (x) { return x.username === username; })[0];
      if (!u || u.passHash !== hashPw(String(req.password || ''), u.salt)) {
        env.attempts.put(key, fails + 1);
        err('Sai tên đăng nhập hoặc mật khẩu.');
      }
      if (tkBiKhoa(u)) err('Tài khoản đã bị khóa' + (u.khoaLyDo ? ' (' + u.khoaLyDo.replace(/^Tự động: /, '') + ')' : '') + '. Liên hệ Phòng Nhân sự.');
      env.attempts.put(key, 0);
      var token = env.uuid() + env.uuid().slice(0, 8);
      if (normRole(u.role) !== u.role) { u.role = normRole(u.role); D('Users'); }
      var s = { username: u.username, role: u.role, maNV: u.maNV, hoTen: u.hoTen, boPhan: boPhanCua(u) };
      env.sessions.put(token, s);
      u.lastLogin = env.now(); D('Users');
      log(s, 'login', '');
      if (can(u.role, 'view_all')) { applyDueRaises(); chamDutDenHan(); }
      var dl = loadAll(s); dl.soCB = soCanhBao(s);
      return { token: token, user: publicUser(u), data: dl };
    };

    H.logout = function (req) { if (req.token) env.sessions.del(req.token); return true; };

    H['dl.cu'] = function (req, s) {               // v3.15: dữ liệu chấm công, bảng lương các kỳ đã duyệt cũ hơn 12 tháng (tải ngầm sau khi mở app)
      if (!can(s.role, 'view_all')) err('Bạn không có quyền.');
      var cs = String(req.truoc || kyCatCu()), dd = kyDuyetMap(), cu = function (x) { return x.ky < cs && dd[x.ky]; };
      return { truoc: cs, cc: T('ChamCong').filter(cu), bl: T('BangLuong').filter(cu) };
    };
    H.load = function (req, s) {
      if (can(s.role, 'view_all')) { applyDueRaises(); chamDutDenHan(); }
      var u = T('Users').filter(function (x) { return x.username === s.username; })[0];
      var d = loadAll(s); d.soCB = soCanhBao(s);
      return { user: publicUser(u), data: d };
    };
    function soCanhBao(s) {                         // số cảnh báo chấm công cần mình làm: QL – chờ giải trình; người kiểm tra – đã giải trình chờ kết luận
      if (CCKT_ROLES.indexOf(s.role) < 0) return 0;
      return T('CanhBaoCC').filter(function (x) { return x.trangThai === 'Đã giải trình'; }).length;
    }

    H.changePassword = function (req, s) {
      var u = T('Users').filter(function (x) { return x.username === s.username; })[0];
      if (u.passHash !== hashPw(String(req.oldPassword || ''), u.salt)) err('Mật khẩu hiện tại không đúng.');
      var p = String(req.newPassword || '');
      if (p.length < 6) err('Mật khẩu mới tối thiểu 6 ký tự.');
      if (p === req.oldPassword) err('Mật khẩu mới phải khác mật khẩu cũ.');
      u.salt = newId(); u.passHash = hashPw(p, u.salt); u.mustChange = ''; D('Users');
      log(s, 'changePassword', '');
      return true;
    };

    /* ---- nhân viên ---- */
    function maNVMoi() {
      var mx = 0; T('NhanVien').forEach(function (n) { var m = /^NV(\d+)$/i.exec(String(n.maNV || '')); if (m && +m[1] > mx) mx = +m[1]; });
      var so = String(mx + 1); while (so.length < 4) so = '0' + so; return 'NV' + so;
    }
    H['nv.save'] = function (req, s) {
      need(s, 'edit_nv');
      var r = req.row || {}, ma = String(r.maNV || '').trim().toUpperCase();
      if (!ma && req.isNew) ma = maNVMoi();                              // tự sinh NV0001, NV0002…
      if (!ma) err('Thiếu mã nhân viên.');
      var rows = T('NhanVien'), cur = findNV(ma);
      if (!String(r.hoTen || (cur && cur.hoTen) || '').trim()) err('Thiếu họ tên.');
      if (req.isNew && cur && req.maTuDong) { ma = maNVMoi(); cur = null; }   // v3.24: mã tạm do giao diện gợi ý đã bị người khác dùng → cấp mã kế tiếp
      if (req.isNew && cur) err('Mã nhân viên ' + ma + ' đã tồn tại.');
      if (!req.isNew && !cur) err('Không tìm thấy nhân viên ' + ma);
      var cols = SCHEMA.NhanVien.filter(function (c) { return c !== 'capNhat'; }), truocNV = cur ? Object.assign({}, cur) : null;
      if (cur && mhNguon(cur, T('LenLuong'), T('MucHuong')).length) {
        // đã có diễn biến mức hưởng: lương, phụ cấp, NPT, ăn ca chỉ thay đổi qua khai báo mức hưởng (giữ lịch sử)
        var pcN = function (x) { var o2 = parseJSON(x, {}), r2 = {}; Object.keys(o2).sort().forEach(function (k) { if (num(o2[k])) r2[k] = num(o2[k]); }); return JSON.stringify(r2); };
        var khac = ['luongChinh', 'luongBH', 'soNPT', 'anCa'].filter(function (c) { return r[c] !== undefined && num(r[c]) !== num(cur[c]); });
        if (r.phuCap !== undefined && pcN(r.phuCap) !== pcN(cur.phuCap)) khac.push('phuCap');
        if (khac.length) err(ma + ' đã có diễn biến mức hưởng: thay đổi lương, phụ cấp, số NPT, ăn ca bằng "Khai báo mức hưởng" (có ngày hiệu lực) để giữ lịch sử.');
        ['luongChinh', 'luongBH', 'soNPT', 'anCa', 'phuCap'].forEach(function (c) { delete r[c]; });
      }
      if (r.loaiLuong !== undefined && r.loaiLuong !== '') {
        var thM = TRUONG_HOP.filter(function (x) { return x.ma === r.loaiLuong; })[0];     // v3.26: nhận mã trường hợp (bh, khoan, noikhac, huu, thuviec)
        if (thM) { r.loaiLuong = thM.loaiLuong; r.lyDoKhongBH = thM.lyDo; }
        if (!LOAI_LUONG[r.loaiLuong]) {                                   // nhập Excel: chấp nhận tên loại ("1. Trả lương + đóng BH…", "Lương khoán việc", "Không đóng BH…")
          var tL = String(r.loaiLuong).toLowerCase().replace(/^\s*\d\s*[.)]\s*/, '');
          if (/thử việc/.test(tL) && !r.lyDoKhongBH) r.lyDoKhongBH = 'Hợp đồng thử việc';
          if (/nơi khác/.test(tL) && !r.lyDoKhongBH) r.lyDoKhongBH = 'Đóng BH nơi khác';
          if (/hưu/.test(tL) && !r.lyDoKhongBH) r.lyDoKhongBH = 'Đã nghỉ hưu';
          r.loaiLuong = /khoán/.test(tL) ? 'khoan' : /không đóng|nơi khác|nghỉ hưu|hưu|thử việc|khongbh/.test(tL) ? 'khongbh' : /đóng b|^bh$/.test(tL) ? 'bh' : r.loaiLuong;
        }
        if (r.loaiLuong === 'khongbh' && !r.lyDoKhongBH && cur && cur.lyDoKhongBH) r.lyDoKhongBH = cur.lyDoKhongBH;
        if (!LOAI_LUONG[r.loaiLuong]) err('Loại lương không hợp lệ.');
        if (r.loaiLuong === 'khongbh' && LY_DO_KHONG_BH.indexOf(r.lyDoKhongBH) < 0) err('Chọn lý do không đóng BH tại Công ty: ' + LY_DO_KHONG_BH.join(' / ') + '.');
        if (r.loaiLuong !== 'khongbh') r.lyDoKhongBH = '';
        r.thamGiaBH = r.loaiLuong === 'bh' ? 'Có' : 'Không';           // giữ cột cũ đồng bộ
      }
      if (r.congTy !== undefined && String(r.congTy).trim() !== '') {   // v3.20: công ty thành viên (mã hoặc tên khi nhập Excel)
        var mCT = nhanCTV(r.congTy, getCfg());
        if (!mCT) err('Công ty "' + r.congTy + '" chưa có trong danh mục công ty thành viên: ' + dsCTV(getCfg()).map(function (x) { return x.ma + ' – ' + x.tenTat; }).join('; ') + '.');
        r.congTy = mCT;
      }
      delete r.lichSuHD;                                                 // v3.33: chỉ thêm / hủy dòng qua hd.save, hd.huy
      if (cur && hdDs(cur).length) ['loaiHD', 'soHD', 'ngayKyHD', 'ngayHetHanHD', 'loaiLuong', 'lyDoKhongBH', 'thamGiaBH'].forEach(function (c) { delete r[c]; });   // theo lịch sử HĐLĐ
      if (r.thueTT !== undefined) r.thueTT = JSON.stringify(chuanTT(r.thueTT));   // v3.26
      if (r.hsThue !== undefined) r.hsThue = JSON.stringify(hsThueNV({ hsThue: r.hsThue }));   // v3.27
      var o = cur || {};
      cols.forEach(function (c) { if (r[c] !== undefined) o[c] = isNumCol('NhanVien', c) ? num(r[c]) : (c === 'phuCap' ? JSON.stringify(parseJSON(r[c], {})) : String(r[c]).trim()); });
      o.maNV = ma; o.capNhat = env.now();
      if (!o.trangThai) o.trangThai = 'Đang làm việc';
      if (r.trangThai !== undefined || r.ngayNghi !== undefined) {      // v3.18: sửa tay trạng thái / ngày chấm dứt trong hồ sơ
        if (o.ngayNghi && !parseD(o.ngayNghi)) err('Ngày chấm dứt hợp đồng không hợp lệ.');
        if (cur && cur.trangThai === 'Đã nghỉ việc' && r.trangThai && r.trangThai !== 'Đã nghỉ việc' && r.ngayNghi === undefined) o.ngayNghi = '';
        if (o.trangThai === 'Đã nghỉ việc' && !o.ngayNghi) o.ngayNghi = env.today();
        if (o.ngayNghi && o.ngayNghi <= env.today()) o.trangThai = 'Đã nghỉ việc';
        else if (o.trangThai === 'Đã nghỉ việc' && o.ngayNghi > env.today()) o.trangThai = 'Đang làm việc';
      }
      if (!cur) {
        rows.push(o);
        if (o.ngayVao) T('QuaTrinh').push({ id: newId(), maNV: ma, loai: 'Tiếp nhận', tuNgay: o.ngayVao, denNgay: '', phongBan: o.phongBan || '', chucVu: o.chucVu || '', donVi: o.donVi || '', soQD: '', nguoiKy: '', noiDung: 'Tiếp nhận vào làm việc', taoBoi: s.username, taoLuc: env.now() });
        D('QuaTrinh');
      }
      T('Users').forEach(function (u) { if (u.maNV === ma) u.hoTen = o.hoTen; }); D('Users');
      D('NhanVien');
      if (cur || o.ngayNghi) khoaTK(o, (chamDutCua(ma) || { loai: 'chấm dứt hợp đồng' }).loai.toLowerCase());
      if (!req.silent) logDoi(s, req.isNew ? 'nv.create' : 'nv.update', ma, truocNV, o);
      if (req.kem) return goiNV(ma);                   // v3.23: trả kèm quá trình, tài khoản của người đó – giao diện không phải tải lại toàn bộ
      return o;
    };


    /* ---- v3.33: LỊCH SỬ HỢP ĐỒNG LAO ĐỘNG (thêm dòng, không ghi đè) ---- */
    function hdTep(nv, f, h, s) {
      var cfg = getCfg();
      if (!MIME_OK.test(f.mime || '')) err('File HĐ: chỉ nhận PDF, ảnh, Word, Excel.');
      var bytes = Math.floor(String(f.data).length * 3 / 4); if (bytes > (num(cfg.dungLuongToiDaMB) || 10) * 1048576) err('File vượt ' + (num(cfg.dungLuongToiDaMB) || 10) + ' MB.');
      var name = String(f.name || 'hop-dong').replace(/[\\/:*?"<>|]/g, '_').slice(0, 150), loai = (cfg.loaiTaiLieu || []).indexOf('Hợp đồng lao động') >= 0 ? 'Hợp đồng lao động' : 'Quyết định';
      var t = { id: newId(), maNV: nv.maNV, loai: loai, ten: ((h.soHD ? 'HĐ ' + h.soHD + ' – ' : '') + h.loaiHD).slice(0, 150), soHieu: h.soHD || '', ngay: h.tuNgay, hetHan: h.denNgay || '', lienKet: '', nvXem: 'Có', trangThai: 'Đã xác nhận',
        tenFile: name, mimeType: f.mime, kichThuoc: bytes, nguoiTai: s.hoTen || s.username, taiLuc: env.now(), nguoiXacNhan: s.hoTen || s.username };
      t.fileId = env.files.put({ name: nv.maNV + '_HDLD_' + name, mime: f.mime, data: f.data, maNV: nv.maNV, hoTen: nv.hoTen, duong: [tenThuMucNV(nv), tlNhom(loai).thuMuc] });
      T('TaiLieu').push(t); D('TaiLieu');
      h.tlId = t.id; h.tenFile = name;
      return t;
    }
    function hdKyDaDuyet(maNV, tu) {                   // kỳ lương đã phê duyệt bị ảnh hưởng (không tự tính lại)
      return T('KyLuong').filter(function (k) { return kyEnd(k.ky) >= tu && normBuoc(k.trangThai, buocLuong(getCfg(), k)) === BUOC_XONG && T('BangLuong').some(function (b) { return b.ky === k.ky && b.maNV === maNV; }); }).map(function (k) { return k.ky; });
    }
    H['hd.save'] = function (req, s) {
      need(s, 'edit_nv');
      var nv = findNV(req.maNV); if (!nv) err('Không tìm thấy nhân viên.');
      var r = req.row || {}, h = {
        id: newId(), soHD: String(r.soHD || '').trim().slice(0, 60), loaiHD: String(r.loaiHD || ''), tuNgay: String(r.tuNgay || ''), denNgay: String(r.denNgay || ''),
        luongHD: R(num(r.luongHD)), luongBH: R(num(r.luongBH)), hinhThuc: String(r.hinhThuc || ''), lyDo: String(r.lyDo || ''), ghiChu: String(r.ghiChu || '').trim().slice(0, 500),
        trangThai: 'Hiệu lực', nguoi: s.hoTen || s.username, luc: env.now(), tlId: '', tenFile: ''
      };
      if (LOAI_HD.indexOf(h.loaiHD) < 0) err('Chọn loại hợp đồng.');
      if (!HT_HD[h.hinhThuc]) err('Chọn hình thức trả lương.');
      if (!parseD(h.tuNgay)) err('Nhập ngày bắt đầu hợp đồng.');
      if (h.denNgay && !parseD(h.denNgay)) err('Ngày kết thúc không hợp lệ.');
      if (h.denNgay && h.denNgay < h.tuNgay) err('Ngày kết thúc phải sau ngày bắt đầu.');
      if (!h.denNgay && h.loaiHD !== 'Không xác định thời hạn') err('HĐ ' + h.loaiHD.toLowerCase() + ' cần ngày kết thúc.');
      if (h.loaiHD === 'Không xác định thời hạn') h.denNgay = '';
      if (h.hinhThuc === 'khongbh') { if (LY_DO_HD_KBH.indexOf(h.lyDo) < 0) err('Chọn lý do không trích BH: ' + LY_DO_HD_KBH.join(' / ') + '.'); } else h.lyDo = '';
      if (!(h.luongHD > 0)) err('Nhập lương theo hợp đồng' + (h.hinhThuc === 'khoan' ? ' (mức khoán tháng)' : '') + '.');
      if (hdCoBH(h) && !(h.luongBH > 0)) err('HĐ này trích BH bắt buộc – nhập lương đóng BH.');
      if (h.loaiHD === 'Thử việc' && h.denNgay && soNgayLich(h.tuNgay, h.denNgay) > 180) err('Thời gian thử việc tối đa 180 ngày (Điều 25 BLLĐ).');
      if (nv.ngayVao && h.tuNgay < nv.ngayVao) err('Ngày bắt đầu HĐ trước ngày vào làm (' + ngayVN(nv.ngayVao) + ').');
      if (h.soHD && hdDs(nv).some(function (x) { return x.soHD === h.soHD; })) err('Số HĐ ' + h.soHD + ' đã có trong lịch sử – kiểm tra lại (dòng sai thì hủy trước khi thêm lại).');
      var tl = null; if (req.file && req.file.data) tl = hdTep(nv, req.file, h, s);
      var truoc = { lichSuHD: nv.lichSuHD || '', loaiHD: nv.loaiHD, loaiLuong: nv.loaiLuong };
      var ds = hdTatCa(nv); ds.push(h); nv.lichSuHD = JSON.stringify(ds);
      hdDongBo(nv, env.today()); nv.capNhat = env.now(); D('NhanVien');
      T('QuaTrinh').push({ id: newId(), maNV: nv.maNV, loai: 'Thay đổi HĐLĐ', tuNgay: h.tuNgay, denNgay: h.denNgay, phongBan: nv.phongBan || '', chucVu: nv.chucVu || '', donVi: nv.donVi || '', soQD: h.soHD, nguoiKy: '',
        noiDung: 'Ký ' + h.loaiHD + (h.soHD ? ' số ' + h.soHD : '') + ' · lương HĐ ' + h.luongHD.toLocaleString('vi-VN') + ' đ' + (hdCoBH(h) ? ' · đóng BH ' + h.luongBH.toLocaleString('vi-VN') + ' đ' : ' · ' + hdMoTa(h).bh.charAt(0).toLowerCase() + hdMoTa(h).bh.slice(1)), taoBoi: s.username, taoLuc: env.now() });
      D('QuaTrinh');
      logDoi(s, 'hd.save', nv.maNV, truoc, { lichSuHD: nv.lichSuHD, loaiHD: nv.loaiHD, loaiLuong: nv.loaiLuong });
      return { nv: nv, tl: tl, kyDaDuyet: hdKyDaDuyet(nv.maNV, h.tuNgay), qt: T('QuaTrinh').filter(function (x) { return x.maNV === nv.maNV; }) };
    };
    H['hd.huy'] = function (req, s) {
      need(s, 'edit_nv');
      var nv = findNV(req.maNV); if (!nv) err('Không tìm thấy nhân viên.');
      var ds = hdTatCa(nv), h = ds.filter(function (x) { return x.id === req.id; })[0]; if (!h) err('Không tìm thấy dòng hợp đồng.');
      if (h.trangThai === 'Đã hủy') err('Dòng này đã hủy.');
      var ly = String(req.lyDo || '').trim(); if (!ly) err('Nhập lý do hủy dòng (VD: nhập sai số liệu).');
      var truoc = nv.lichSuHD;
      h.trangThai = 'Đã hủy'; h.lyDoHuy = ly.slice(0, 300); h.nguoiHuy = s.hoTen || s.username; h.lucHuy = env.now();
      nv.lichSuHD = JSON.stringify(ds); hdDongBo(nv, env.today()); nv.capNhat = env.now(); D('NhanVien');
      logDoi(s, 'hd.huy', nv.maNV, { lichSuHD: truoc }, { lichSuHD: nv.lichSuHD });
      return { nv: nv, kyDaDuyet: hdKyDaDuyet(nv.maNV, h.tuNgay) };
    };
    H['hd.tep'] = function (req, s) {
      need(s, 'edit_nv');
      var nv = findNV(req.maNV); if (!nv) err('Không tìm thấy nhân viên.');
      var ds = hdTatCa(nv), h = ds.filter(function (x) { return x.id === req.id; })[0]; if (!h) err('Không tìm thấy dòng hợp đồng.');
      if (!req.file || !req.file.data) err('Chưa chọn file.');
      var tl = hdTep(nv, req.file, h, s); nv.lichSuHD = JSON.stringify(ds); nv.capNhat = env.now(); D('NhanVien');
      log(s, 'hd.tep', nv.maNV + ' ' + (h.soHD || h.loaiHD) + ': ' + tl.tenFile);
      return { nv: nv, tl: tl };
    };
    /* nhắc TCHC: HĐLĐ sắp hết hạn – mỗi HĐ báo 1 lần khi vào ngưỡng (Cài đặt → canhBaoHD ngày) */
    function hdNhac() {
      var cfg = getCfg(), hom = env.today(), n = 0, ds = [];
      T('NhanVien').forEach(function (nv) {
        var c = hdCanhBao(nv, hom, cfg.canhBaoHD); if (!c || c.hd.daBao === c.hd.denNgay) return;
        var all = hdTatCa(nv), h = all.filter(function (x) { return x.id === c.hd.id; })[0]; if (!h) return;
        h.daBao = h.denNgay; nv.lichSuHD = JSON.stringify(all); n++;
        ds.push(nv.hoTen + ' (' + h.loaiHD + (h.soHD ? ' ' + h.soHD : '') + ', ' + (c.con < 0 ? 'đã hết hạn ' + ngayVN(h.denNgay) : 'hết hạn ' + ngayVN(h.denNgay) + ' – còn ' + c.con + ' ngày') + ')');
      });
      if (n) {
        D('NhanVien');
        bao(nguoiCo(['r:tchc']), { loai: 'duyet', t: n + ' hợp đồng lao động sắp hết hạn', b: ds.slice(0, 4).join('; ') + (ds.length > 4 ? '…' : ''), tag: 'hd_han', mo: '#v=nv' });
      }
      return n;
    }

    H['nv.thue'] = function (req, s) {               // v3.26: trạng thái tính thuế theo thời điểm (TCHC, Kế toán lương)
      if (!can(s.role, 'edit_nv') && !can(s.role, 'payroll')) err('Bạn không có quyền thực hiện thao tác này.');
      var nv = findNV(req.maNV); if (!nv) err('Không tìm thấy nhân viên.');
      var truoc = { thueTT: nv.thueTT, camKet08: nv.camKet08 }, ds = chuanTT(req.thueTT || []), hs = hsThueNV(nv);
      if (hs.noiKhac.gtTai === 'noikhac') ds.forEach(function (r) { if (r.giamTru === 'co') err('Kỳ ' + r.tu.slice(5) + '/' + r.tu.slice(0, 4) + ': CBCNV đã kê khai đăng ký giảm trừ bản thân tại nơi khác – không được giảm trừ trùng tại Công ty. Sửa "Nơi đăng ký giảm trừ" trong Hồ sơ thuế nếu đã thay đổi.'); });
      nv.thueTT = JSON.stringify(ds); if (req.camKet08 !== undefined) nv.camKet08 = req.camKet08 === 'Có' ? 'Có' : 'Không';
      nv.capNhat = env.now(); D('NhanVien');
      logDoi(s, 'nv.thue', nv.maNV, truoc, { thueTT: nv.thueTT, camKet08: nv.camKet08 });
      return nv;
    };
    /* v3.27: HỒ SƠ THUẾ – BH của CBCNV (cư trú, thử việc, BH nơi khác, nghỉ hưu, khoán, thu nhập nơi khác, giảm trừ khác) */
    H['nv.hsthue'] = function (req, s) {
      if (!can(s.role, 'edit_nv') && !can(s.role, 'payroll') && !can(s.role, 'duyetThue')) err('Bạn không có quyền thực hiện thao tác này.');
      var nv = findNV(req.maNV); if (!nv) err('Không tìm thấy nhân viên.');
      var moi = hsThueNV({ hsThue: req.hsThue || {} }), cu = hsThueNV(nv), truoc = nv.hsThue || '';
      ['trangThaiXM', 'ngayXM', 'nguoiXM'].forEach(function (k) { moi.huu[k] = cu.huu[k]; });          // v3.29: xác minh hưu trí chỉ qua nv.huuxm
      if (['tinhTrang', 'ngayHuong', 'soQD', 'donViChiTra'].some(function (k) { return String(moi.huu[k]) !== String(cu.huu[k]); }) && cu.huu.trangThaiXM !== 'Chờ') { moi.huu.trangThaiXM = 'Chờ'; moi.huu.ngayXM = ''; moi.huu.nguoiXM = ''; }
      if (moi.noiKhac.gtTai === 'noikhac' && truongHopNV(nv) === 'noikhac' && chuanTT(nv.thueTT).some(function (r) { return r.giamTru === 'co'; }))
        err('Trạng thái tính thuế đang có mốc giảm trừ gia cảnh tại Công ty – sửa trạng thái thuế trước khi khai đăng ký giảm trừ ở nơi khác (tránh giảm trừ 2 nơi).');
      nv.hsThue = JSON.stringify(moi); nv.capNhat = env.now(); D('NhanVien');
      logDoi(s, 'nv.hsthue', nv.maNV, { hsThue: truoc }, { hsThue: nv.hsThue });
      return nv;
    };
    /* v3.27: CAM KẾT, ỦY QUYỀN, KÊ KHAI… – bảng HoSoThue; Nhân sự / Kế toán lương lập, Phụ trách kế toán phê duyệt */
    var HST_LOAI = { camket: 'Cam kết thu nhập (không khấu trừ 10%)', uyquyen: 'Ủy quyền quyết toán thuế TNCN', npt: 'Đăng ký người phụ thuộc', kekhaiGT: 'Kê khai thu nhập & nơi giảm trừ bản thân', xacnhanBH: 'Xác nhận nơi tham gia BHXH khác' };
    var HST_TT = ['Nháp', 'Chờ phụ trách KT duyệt', 'Đã duyệt', 'Từ chối', 'Chờ xác minh mẫu', 'Thu hồi'];
    H['hst.save'] = function (req, s) {
      if (!can(s.role, 'edit_nv') && !can(s.role, 'payroll') && !can(s.role, 'duyetThue')) err('Bạn không có quyền thực hiện thao tác này.');
      var r = req.row || {}, nv = findNV(r.maNV); if (!nv) err('Không tìm thấy nhân viên.');
      if (!HST_LOAI[r.loai]) err('Loại hồ sơ không hợp lệ.');
      if (!/^\d{4}$/.test(String(r.nam || ''))) err('Nhập năm áp dụng (YYYY).');
      if (r.ngayKy && !parseD(r.ngayKy)) err('Ngày ký không hợp lệ.');
      var rows = T('HoSoThue'), o = r.id ? rows.filter(function (x) { return x.id === r.id; })[0] : null, cfg = getCfg();
      if (r.id && !o) err('Không tìm thấy hồ sơ.');
      if (o && o.trangThai === 'Đã duyệt' && !can(s.role, 'duyetThue')) err('Hồ sơ đã được Phụ trách kế toán duyệt – chỉ Phụ trách kế toán sửa được.');
      var truoc = o ? Object.assign({}, o) : null; if (!o) { o = { id: newId(), nguoiLap: s.hoTen || s.username, ngayLap: env.now() }; rows.push(o); }
      var bm = bieuMau(cfg, r.loai) || {};
      o.maNV = nv.maNV; o.nam = String(r.nam); o.loai = r.loai; o.mauSo = String(r.mauSo || bm.mauSo || ''); o.phienBan = String(r.phienBan || bm.phienBan || '');
      o.ngayKy = r.ngayKy || ''; o.duLieu = JSON.stringify(parseJSON(r.duLieu, {}) || {}); o.ghiChu = String(r.ghiChu || '').slice(0, 500);
      if (r.ketQua !== undefined) o.ketQua = typeof r.ketQua === 'string' ? r.ketQua.slice(0, 4000) : JSON.stringify(r.ketQua).slice(0, 4000);
      o.trangThai = req.gui ? (bm.trangThai === 'dangxacminh' ? 'Chờ xác minh mẫu' : 'Chờ phụ trách KT duyệt') : (o.trangThai && o.trangThai !== 'Từ chối' && truoc ? (o.trangThai === 'Đã duyệt' ? 'Chờ phụ trách KT duyệt' : o.trangThai) : 'Nháp');
      var f = req.file;
      if (f && f.data) {                                  // bản đã ký: lưu vào Tài liệu (Drive) của CBCNV, loại "Hồ sơ thuế TNCN"
        var name = String(f.name || 'ban-ky').replace(/[\\/:*?"<>|]/g, '_').slice(0, 150);
        if (!MIME_OK.test(f.mime || '')) err('Chỉ nhận PDF, ảnh, Word.');
        var bytes = Math.floor(String(f.data).length * 3 / 4); if (bytes > (num(cfg.dungLuongToiDaMB) || 10) * 1048576) err('File vượt ' + (num(cfg.dungLuongToiDaMB) || 10) + ' MB.');
        var t = { id: newId(), maNV: nv.maNV, loai: 'Hồ sơ thuế TNCN', ten: (HST_LOAI[o.loai] + ' ' + o.nam).slice(0, 150), soHieu: o.mauSo, ngay: o.ngayKy || env.today(), hetHan: o.nam + '-12-31', lienKet: '', nvXem: 'Có', trangThai: 'Đã xác nhận', tenFile: name, mimeType: f.mime, kichThuoc: bytes, nguoiTai: s.hoTen || s.username, taiLuc: env.now(), nguoiXacNhan: s.hoTen || s.username };
        t.fileId = env.files.put({ name: nv.maNV + '_HoSoThue_' + name, mime: f.mime, data: f.data, maNV: nv.maNV, hoTen: nv.hoTen, duong: [tenThuMucNV(nv), tlNhom(t.loai).thuMuc] });
        T('TaiLieu').push(t); D('TaiLieu'); o.taiLieuId = t.id; o.tenTep = name;
      }
      o.capNhat = env.now(); D('HoSoThue');
      logDoi(s, 'hst.save', nv.maNV + ' ' + o.loai + ' ' + o.nam, truoc, o, ['capNhat', 'duLieu', 'ketQua']);
      return { hst: o, tl: o.taiLieuId ? T('TaiLieu').filter(function (x) { return x.id === o.taiLieuId; })[0] : null };
    };
    H['hst.duyet'] = function (req, s) {
      need(s, 'duyetThue');
      var o = T('HoSoThue').filter(function (x) { return x.id === req.id; })[0]; if (!o) err('Không tìm thấy hồ sơ.');
      var kq = req.kq === 'duyet' ? 'Đã duyệt' : req.kq === 'tuchoi' ? 'Từ chối' : req.kq === 'xacminh' ? 'Chờ xác minh mẫu' : req.kq === 'thuhoi' ? 'Thu hồi' : ''; if (!kq) err('Chọn kết quả.');
      if (kq === 'Thu hồi' && !String(req.ghiChu || '').trim()) err('Ghi lý do thu hồi cam kết.');
      if (kq === 'Từ chối' && !String(req.ghiChu || '').trim()) err('Ghi lý do từ chối.');
      if (kq === 'Đã duyệt' && !o.taiLieuId) err('Chưa tải bản đã ký – không duyệt được.');
      var bm = bieuMau(getCfg(), o.loai);
      if (kq === 'Đã duyệt' && bm && bm.trangThai === 'dangxacminh' && !req.xacNhanMau) err('Mẫu "' + bm.mauSo + '" đang ở trạng thái "Đang xác minh". Tích xác nhận mẫu đã được cơ quan thuế chấp nhận (hoặc cập nhật trạng thái mẫu trong Kho biểu mẫu) trước khi duyệt.');
      var truoc = { trangThai: o.trangThai };
      o.trangThai = kq; o.nguoiKiemTra = s.hoTen || s.username; o.ngayKiemTra = env.now(); o.ghiChuDuyet = String(req.ghiChu || '').slice(0, 500) + (req.xacNhanMau && kq === 'Đã duyệt' ? (req.ghiChu ? ' · ' : '') + 'Phụ trách KT xác nhận mẫu được chấp nhận' : '');
      o.capNhat = env.now(); D('HoSoThue');
      if (kq === 'Đã duyệt' && (o.loai === 'camket' || o.loai === 'uyquyen')) {   // đồng bộ trạng thái tính thuế từ kỳ 01 của năm áp dụng
        var nv = findNV(o.maNV);
        if (nv) { var ds = chuanTT(nv.thueTT), tu = o.nam + '-01', r = ds.filter(function (x) { return x.tu === tu; })[0];
          if (!r) { var cuR = null; ds.forEach(function (x) { if (x.tu <= tu) cuR = x; }); r = cuR ? Object.assign({}, cuR, { tu: tu }) : { tu: tu, cach: 'tudong', giamTru: 'tudong', camKet: 'tudong', uyQuyen: 'tudong', ghiChu: '' }; ds.push(r); }
          if (o.loai === 'camket') r.camKet = 'co'; else r.uyQuyen = 'co';
          r.ghiChu = (r.ghiChu ? r.ghiChu + '; ' : '') + (o.loai === 'camket' ? 'Cam kết ' : 'Ủy quyền ') + o.nam + ' đã duyệt';
          if (o.loai === 'camket' && !ds.some(function (x) { return x.tu === (+o.nam + 1) + '-01'; })) ds.push({ tu: (+o.nam + 1) + '-01', cach: 'tudong', giamTru: 'tudong', camKet: 'khong', uyQuyen: 'tudong', ghiChu: 'Hết năm áp dụng cam kết ' + o.nam });
          nv.thueTT = JSON.stringify(chuanTT(ds)); nv.capNhat = env.now(); D('NhanVien'); }
      }
      if (kq === 'Thu hồi' && o.loai === 'camket') {    // v3.28: thu hồi cam kết → các kỳ / lần trả từ tháng này tiếp tục khấu trừ
        var nvT = findNV(o.maNV), tuT = env.today().slice(0, 7);
        if (nvT) { var dsT = chuanTT(nvT.thueTT).filter(function (x) { return x.tu !== tuT; }), cuT = null; dsT.forEach(function (x) { if (x.tu <= tuT) cuT = x; });
          dsT.push(Object.assign({ cach: 'tudong', giamTru: 'tudong', uyQuyen: 'tudong' }, cuT || {}, { tu: tuT, camKet: 'khong', ghiChu: 'Thu hồi cam kết ' + o.nam })); nvT.thueTT = JSON.stringify(chuanTT(dsT)); nvT.capNhat = env.now(); D('NhanVien'); }
      }
      logDoi(s, 'hst.duyet', o.maNV + ' ' + o.loai + ' ' + o.nam, truoc, { trangThai: o.trangThai, ghiChuDuyet: o.ghiChuDuyet });
      return { hst: o, nv: findNV(o.maNV) };
    };
    /* ==================== v3.28: HỢP ĐỒNG KHOÁN VIỆC & THANH TOÁN KHOÁN ==================== */
    function khoanCtx() { return { cfg: getCfg(), hst: T('HoSoThue'), tl: T('TaiLieu'), bls: T('BangLuong'), hdk: T('HopDongKhoan'), ttk: T('ThanhToanKhoan'), homNay: env.today() }; }
    function laBCH(s) { return ROLE_QL.indexOf(s.role) >= 0; }
    function quyenKhoan(s, hd) {             // lập / sửa hợp đồng khoán & đợt thanh toán
      if (can(s.role, 'edit_nv') || can(s.role, 'payroll') || can(s.role, 'duyetThue') || s.role === 'admin') return true;
      if (laBCH(s) && (!hd || !hd.congTrinh || hd.congTrinh === s.boPhan)) return true;
      return false;
    }
    function nguoiBuocKhoan(s, b) { if (s.role === 'admin') return true; return (b.nguoi || []).some(function (a) { return a === 'u:' + s.username || a === 'r:' + s.role; }); }
    function tinhLaiTT(x, hd, nv, ctx) {
      var tu = tamUngTon(hd.id, ctx.ttk, x.id);
      if (x.loai !== 'tamung') x.thuHoiTUThue = Math.min(num(x.thuHoiTU), Math.max(0, tu.conThue));   // thu hồi phần tạm ứng đã chịu thuế trước → không tính thuế 2 lần
      var k = calculateMonthlyContractPayment(x, hd, nv, ctx);
      x.tienPhaiTra = k.tienPhaiTra; x.tnChiuThue = k.tnChiuThue; x.thueSuat = k.thueSuat; x.thue = k.thue; x.thucTra = k.thucTra; x.cheDoThue = k.cheDoThue; x.congThuc = k.congThuc;
      return k;
    }
    H['hdk.save'] = function (req, s) {
      var r = req.row || {}, a = T('HopDongKhoan'), o = r.id ? a.filter(function (x) { return x.id === r.id; })[0] : null;
      if (r.id && !o) err('Không tìm thấy hợp đồng.');
      if (!quyenKhoan(s, o || r)) err('Bạn không có quyền lập / sửa hợp đồng khoán của công trình này.');
      var nv = findNV(r.maNV || (o && o.maNV)); if (!nv) err('Chọn người nhận khoán (hồ sơ CBCNV). Người ngoài Công ty: tạo hồ sơ với loại HĐ "Thời vụ / Khoán việc".');
      if (!LOAI_K[r.loaiK]) err('Chọn loại K1–K4.');
      if (!String(r.soHD || '').trim()) err('Nhập số hợp đồng.');
      if (r.tuNgay && !parseD(r.tuNgay) || r.denNgay && !parseD(r.denNgay)) err('Ngày hợp đồng không hợp lệ.');
      if (laBCH(s) && s.role !== 'admin' && r.congTrinh && r.congTrinh !== s.boPhan) err('Chỉ lập hợp đồng khoán cho bộ phận / công trình của mình.');
      var truoc = o ? Object.assign({}, o) : null; if (!o) { o = { id: newId(), nguoiTao: s.hoTen || s.username, taoLuc: env.now(), trangThai: 'Đang thực hiện', hoSo: '{}' }; a.push(o); }
      ['soHD', 'loaiK', 'congTrinh', 'hangMuc', 'doiTo', 'tuNgay', 'denNgay', 'hinhThuc', 'donViTinh', 'kieuGia', 'ghiChu'].forEach(function (k) { if (r[k] !== undefined) o[k] = String(r[k]).trim().slice(0, 500); });
      o.maNV = nv.maNV; if (!HT_KHOAN[o.hinhThuc]) o.hinhThuc = 'thang'; if (o.kieuGia !== 'net') o.kieuGia = 'gross';
      ['giaTriHD', 'donGia'].forEach(function (k) { if (r[k] !== undefined) o[k] = num(r[k]); });
      if (r.danhGia !== undefined) { var dg = parseJSON(r.danhGia, {}) || {}, c = {}; CH_BAN_CHAT.concat(CH_TO_DOI).forEach(function (q) { if (dg[q.ma]) c[q.ma] = String(dg[q.ma]).slice(0, 20); }); if (dg.kinhDoanh) c.kinhDoanh = dg.kinhDoanh === 'co' ? 'co' : 'khong'; o.danhGia = JSON.stringify(c); }
      if (r.trangThai && ['Đang thực hiện', 'Tạm dừng', 'Đã thanh lý', 'Hủy'].indexOf(r.trangThai) >= 0) o.trangThai = r.trangThai;
      var pl = classifyContractRelationship(o); o.ketLuan = pl.ketQua;
      o.capNhat = env.now(); D('HopDongKhoan');
      logDoi(s, 'hdk.save', o.soHD + ' ' + nv.maNV, truoc, o, ['capNhat', 'hoSo']);
      return { hd: o, phanLoai: pl, bhxh: evaluateEmploymentInsuranceRisk(o) };
    };
    H['hdk.hoso'] = function (req, s) {           // hồ sơ 9 thư mục: tải tệp / đặt trạng thái
      var o = T('HopDongKhoan').filter(function (x) { return x.id === req.id; })[0]; if (!o) err('Không tìm thấy hợp đồng.');
      if (!quyenKhoan(s, o)) err('Bạn không có quyền.');
      var f = HS_KHOAN.filter(function (x) { return x[0] === req.thuMuc; })[0]; if (!f) err('Thư mục hồ sơ không hợp lệ.');
      var hs = parseJSON(o.hoSo, {}) || {}, cu = hs[f[0]] || { trangThai: 'Chưa cung cấp', tep: [] }, nv = findNV(o.maNV), cfg = getCfg(), t = null;
      if (req.file && req.file.data) {
        if (!MIME_OK.test(req.file.mime || '')) err('Chỉ nhận PDF, ảnh, Word, Excel.');
        var bytes = Math.floor(String(req.file.data).length * 3 / 4); if (bytes > (num(cfg.dungLuongToiDaMB) || 10) * 1048576) err('File quá lớn.');
        var name = String(req.file.name || 'tep').replace(/[\\/:*?"<>|]/g, '_').slice(0, 150), nam = String(o.tuNgay || env.today()).slice(0, 4);
        t = { id: newId(), maNV: o.maNV, loai: 'Hồ sơ khoán việc', ten: (f[2] + ' – HĐ ' + o.soHD).slice(0, 150), soHieu: o.soHD, ngay: env.today(), hetHan: '', lienKet: '', nvXem: 'Có', trangThai: 'Đã xác nhận', tenFile: name, mimeType: req.file.mime, kichThuoc: bytes, nguoiTai: s.hoTen || s.username, taiLuc: env.now(), nguoiXacNhan: s.hoTen || s.username };
        t.fileId = env.files.put({ name: o.soHD.replace(/[\\/:*?"<>|]/g, '_') + '_' + name, mime: req.file.mime, data: req.file.data, maNV: o.maNV, hoTen: nv ? nv.hoTen : '', duong: ['TAPACO', 'PAYROLL', 'NAM_' + nam, 'HOP_DONG_KHOAN', o.maNV, f[1]] });
        T('TaiLieu').push(t); D('TaiLieu'); cu.tep = (cu.tep || []).concat([t.id]); if (cu.trangThai === 'Chưa cung cấp' || cu.trangThai === 'Cần bổ sung') cu.trangThai = 'Đã tải lên';
      }
      if (req.trangThai) {
        if (HS_TT.indexOf(req.trangThai) < 0) err('Trạng thái hồ sơ không hợp lệ.');
        if ((req.trangThai === 'Hợp lệ' || req.trangThai === 'Không hợp lệ') && !can(s.role, 'payroll') && !can(s.role, 'duyetThue')) err('Chỉ Kế toán / Phụ trách kế toán xác nhận hồ sơ hợp lệ.');
        if (req.trangThai === 'Hợp lệ' && f[0] === 'f06') { var bm = bieuMau(cfg, 'camket'); if (bm && bm.trangThai !== 'hienhanh' && !req.xacNhanMau) err('Mẫu cam kết đang "' + BM_TT[bm.trangThai] + '" – "Mẫu tham khảo – cần xác minh": không xác nhận hồ sơ pháp lý hoàn tất.'); }
        cu.trangThai = req.trangThai;
      }
      cu.nguoi = s.hoTen || s.username; cu.luc = env.now(); hs[f[0]] = cu; var truoc = o.hoSo; o.hoSo = JSON.stringify(hs); o.capNhat = env.now(); D('HopDongKhoan');
      log(s, 'hdk.hoso', o.soHD + ' ' + f[1] + ': ' + (parseJSON(truoc, {})[f[0]] || { trangThai: 'Chưa cung cấp' }).trangThai + ' → ' + cu.trangThai + (t ? ' (+ ' + t.tenFile + ')' : ''));
      return { hd: o, tl: t };
    };
    H['ttk.save'] = function (req, s) {
      var r = req.row || {}, a = T('ThanhToanKhoan'), x = r.id ? a.filter(function (y) { return y.id === r.id; })[0] : null;
      if (r.id && !x) err('Không tìm thấy đợt thanh toán.');
      var hd = T('HopDongKhoan').filter(function (h) { return h.id === (r.hdId || (x && x.hdId)); })[0]; if (!hd) err('Không tìm thấy hợp đồng khoán.');
      if (!quyenKhoan(s, hd)) err('Bạn không có quyền.');
      if (x && x.buoc > 0 && x.trangThai !== 'Trả lại') err('Đợt thanh toán đang trong quy trình duyệt – không sửa được. Người duyệt có thể trả lại.');
      if (['tamung', 'nghiemthu', 'quyettoan'].indexOf(r.loai || (x && x.loai)) < 0) err('Chọn loại: tạm ứng / nghiệm thu thanh toán / quyết toán.');
      if (!parseD(r.ngayTra || (x && x.ngayTra))) err('Nhập ngày dự kiến thực trả (thuế tính theo ngày thực trả).');
      if (r.ngayNT && !parseD(r.ngayNT)) err('Ngày nghiệm thu không hợp lệ.');
      var nv = findNV(hd.maNV), truoc = x ? Object.assign({}, x) : null;
      if (!x) { x = { id: newId(), hdId: hd.id, maNV: hd.maNV, congTrinh: hd.congTrinh, nguoiTao: s.hoTen || s.username, taoLuc: env.now(), buoc: 0, trangThai: 'Nháp', lichSu: '[]', dot: a.filter(function (y) { return y.hdId === hd.id; }).length + 1 }; a.push(x); }
      ['loai', 'kyKL', 'ngayNT', 'ngayTra', 'tuTinhThue', 'ghiChu'].forEach(function (k) { if (r[k] !== undefined) x[k] = String(r[k]).trim().slice(0, 500); });
      ['khoiLuong', 'giaTriNT', 'thuNhapBS', 'giamTruHD', 'tamUng', 'thuHoiTU', 'khauTruKhac'].forEach(function (k) { if (r[k] !== undefined) x[k] = Math.max(0, num(r[k])); });
      if (x.loai === 'tamung') { x.giaTriNT = 0; x.thuNhapBS = 0; x.giamTruHD = 0; x.thuHoiTU = 0; if (!(num(x.tamUng) > 0)) err('Nhập số tiền tạm ứng.'); }
      else { x.tamUng = 0; if (!(num(x.giaTriNT) > 0)) err('Nhập giá trị nghiệm thu được xác nhận thanh toán.'); }
      var ctx = khoanCtx(), tu = tamUngTon(hd.id, ctx.ttk, x.id);
      if (num(x.thuHoiTU) > tu.con + 0.5) err('Thu hồi tạm ứng ' + num(x.thuHoiTU).toLocaleString('vi-VN') + ' vượt số tạm ứng còn lại ' + tu.con.toLocaleString('vi-VN') + ' đ.');
      var netM = num(req.thucNhanNet || r.thucNhanNet); if (hd.kieuGia === 'net' && netM > 0 && x.loai !== 'tamung') { var g = khoanNetToGross(Object.assign({}, x, { thuHoiTUThue: Math.min(num(x.thuHoiTU), tu.conThue) }), hd, nv, ctx, netM); x.giaTriNT = g.giaTriNT - num(x.thuNhapBS) + num(x.giamTruHD); }
      var k = tinhLaiTT(x, hd, nv, ctx);
      if (req.gui || r.gui) { x.buoc = 1; x.trangThai = getCfg().quyTrinhKhoan[0].id; x.lichSu = JSON.stringify(parseJSON(x.lichSu, []).concat([{ buoc: 'Lập', nguoi: s.hoTen || s.username, luc: env.now(), kq: 'gửi', ghiChu: '' }])); }
      x.capNhat = env.now(); D('ThanhToanKhoan');
      logDoi(s, 'ttk.save', hd.soHD + ' đợt ' + x.dot, truoc, x, ['capNhat', 'congThuc', 'lichSu']);
      return { tt: x, kq: k, ghiChu: generateContractorTaxNotes(x, hd, nv, ctx) };
    };
    /* 6 bước duyệt: BCH → KT-KH → Kế toán → Phụ trách KT → Giám đốc → Kế toán thanh toán (Đã chi) */
    H['ttk.move'] = function (req, s) {
      var x = T('ThanhToanKhoan').filter(function (y) { return y.id === req.id; })[0]; if (!x) err('Không tìm thấy đợt thanh toán.');
      var qt = getCfg().quyTrinhKhoan, i = num(x.buoc) - 1, b = qt[i]; if (!b || x.trangThai === 'Đã chi' || x.trangThai === 'Hủy') err('Đợt thanh toán không ở bước duyệt.');
      if (!nguoiBuocKhoan(s, b)) err('Bước "' + b.ten + '" do ' + b.nguoi.map(function (a) { return /^r:/.test(a) ? ROLES[a.slice(2)] : a.slice(2); }).join(' / ') + ' xử lý.');
      var hd = T('HopDongKhoan').filter(function (h) { return h.id === x.hdId; })[0], nv = findNV(x.maNV), ctx = khoanCtx(), ls = parseJSON(x.lichSu, []), ghi = String(req.ghiChu || '').slice(0, 500);
      if (b.id === 'BCH' && laBCH(s) && s.role !== 'admin' && hd.congTrinh && hd.congTrinh !== s.boPhan) err('Đợt thanh toán không thuộc công trình của bạn.');
      if (req.kq === 'tralai') { if (!ghi) err('Ghi lý do trả lại.'); x.buoc = 0; x.trangThai = 'Trả lại'; ls.push({ buoc: b.id, nguoi: s.hoTen || s.username, luc: env.now(), kq: 'trả lại', ghiChu: ghi }); }
      else {
        var k = tinhLaiTT(x, hd, nv, ctx), hs = parseJSON(hd.hoSo, {}) || {}, co = function (f) { return hs[f] && hs[f].trangThai !== 'Chưa cung cấp' && hs[f].trangThai !== 'Không hợp lệ' && hs[f].trangThai !== 'Hết hiệu lực'; };
        if (b.id === 'KT' || b.id === 'PTKT') {
          if (k.cheDoThue === 'hdld') err('Hợp đồng có bản chất quan hệ lao động – chuyển sang quy trình HĐLĐ (đóng BHXH, thuế lũy tiến); không duyệt chế độ khoán.');
          if (k.cheDoThue === 'thieu') err('Chưa đủ dữ liệu đánh giá bản chất hợp đồng (BM-01) – tạm dừng phê duyệt chế độ thuế / BHXH.');
          if (!co('f01')) err('Thiếu hồ sơ 01_Hop_dong_khoan.');
          if (x.loai !== 'tamung' && !co('f04')) err('Thiếu hồ sơ 04_Khoi_luong_nghiem_thu – cam kết thuế không thay thế chứng từ chứng minh khối lượng.');
          if (k.camKet.ok && !co('f06')) err('Áp dụng cam kết nhưng thư mục 06_Cam_ket_thue chưa có bản ký.');
        }
        if (b.id === 'PTKT' && k.cheDoThue === 'khongcutru' && !req.xacNhanKCT) err('Cá nhân không cư trú: tích xác nhận mức thuế 20% đã được Phụ trách kế toán kiểm tra.');
        if (b.id === 'TT' && !co('f05') && !req.boQuaCT) err('Tải chứng từ thanh toán (05_Chung_tu_thanh_toan) trước khi xác nhận đã chi.');
        ls.push({ buoc: b.id, nguoi: s.hoTen || s.username, luc: env.now(), kq: 'duyệt', ghiChu: ghi, thue: x.thue, thucTra: x.thucTra });
        if (i + 1 >= qt.length) { x.buoc = qt.length + 1; x.trangThai = 'Đã chi'; } else { x.buoc = i + 2; x.trangThai = qt[i + 1].id; }
      }
      x.lichSu = JSON.stringify(ls); x.capNhat = env.now(); D('ThanhToanKhoan');
      log(s, 'ttk.move', (hd ? hd.soHD : x.hdId) + ' đợt ' + x.dot + ': ' + b.id + ' → ' + x.trangThai + (ghi ? ' (' + ghi + ')' : '') + ' · thuế ' + num(x.thue).toLocaleString('vi-VN') + ' · thực trả ' + num(x.thucTra).toLocaleString('vi-VN'));
      return { tt: x };
    };
    H['ttk.delete'] = function (req, s) {
      var a = T('ThanhToanKhoan'), i = -1; a.forEach(function (y, j) { if (y.id === req.id) i = j; }); if (i < 0) err('Không tìm thấy đợt thanh toán.');
      var x = a[i], hd = T('HopDongKhoan').filter(function (h) { return h.id === x.hdId; })[0];
      if (!quyenKhoan(s, hd)) err('Bạn không có quyền.');
      if (num(x.buoc) > 0) err('Đợt đang trong quy trình duyệt hoặc đã chi – không xóa được.');
      a.splice(i, 1); D('ThanhToanKhoan'); logDoi(s, 'ttk.delete', (hd ? hd.soHD : '') + ' đợt ' + x.dot, x, null, ['congThuc', 'lichSu']); return true;
    };
    /* v3.27: cập nhật cơ sở dữ liệu quy định & kho biểu mẫu – Quản trị hoặc Phụ trách kế toán (người có thẩm quyền xác nhận mẫu) */
    H['qd.save'] = function (req, s) {
      if (!can(s.role, 'config') && !can(s.role, 'duyetThue')) err('Bạn không có quyền thực hiện thao tác này.');
      var c = {}, cu = getCfg();
      if (Array.isArray(req.quyDinh)) c.quyDinh = chuanDsMa(req.quyDinh, [], QD_TRUONG);
      if (Array.isArray(req.bieuMau)) c.bieuMau = chuanDsMa(req.bieuMau, [], BM_TRUONG);
      if (!c.quyDinh && !c.bieuMau) err('Không có gì để lưu.');
      (c.quyDinh || []).concat(c.bieuMau || []).forEach(function (x) { if (x.url && !/^https:\/\//.test(x.url)) err('"' + (x.ten || x.ma) + '": đường dẫn nguồn phải bắt đầu bằng https://'); });
      var doi = [];
      ['quyDinh', 'bieuMau'].forEach(function (k) {
        if (!c[k]) return; var mc = {}; (cu[k] || []).forEach(function (x) { mc[x.ma] = x; });
        c[k].forEach(function (x) { var y = mc[x.ma] || {}; Object.keys(x).forEach(function (f) { if (String(y[f] || '') !== String(x[f] || '')) doi.push(x.ma + '.' + f + ': ' + (String(y[f] || '∅').slice(0, 200)) + ' → ' + (String(x[f] || '∅').slice(0, 200))); }); });
      });
      setCfg(c); var moi = getCfg();
      log(s, 'qd.save', (doi.join(' ; ') || 'không đổi').slice(0, 4000));
      return { quyDinh: moi.quyDinh, bieuMau: moi.bieuMau };
    };
    H['hst.delete'] = function (req, s) {
      var a = T('HoSoThue'), i = -1; a.forEach(function (x, j) { if (x.id === req.id) i = j; }); if (i < 0) err('Không tìm thấy hồ sơ.');
      var o = a[i];
      if (o.trangThai === 'Đã duyệt' ? !can(s.role, 'duyetThue') : (!can(s.role, 'edit_nv') && !can(s.role, 'payroll') && !can(s.role, 'duyetThue'))) err('Bạn không có quyền xóa hồ sơ này.');
      a.splice(i, 1); D('HoSoThue'); logDoi(s, 'hst.delete', o.maNV + ' ' + o.loai + ' ' + o.nam, o, null);
      return true;
    };
    /* v3.29: xác minh tình trạng hưởng lương hưu (Kế toán lương / Phụ trách KT) */
    H['nv.huuxm'] = function (req, s) {
      if (!can(s.role, 'payroll') && !can(s.role, 'duyetThue')) err('Chỉ Kế toán lương / Phụ trách kế toán xác minh tình trạng hưu trí.');
      var nv = findNV(req.maNV); if (!nv) err('Không tìm thấy nhân viên.');
      if (HUU_XM.indexOf(req.trangThai) < 0) err('Chọn trạng thái xác minh.');
      var h = parseJSON(nv.hsThue, {}) || {}, hu = hsThueNV(nv).huu, truoc = { trangThaiXM: hu.trangThaiXM, ngayXM: hu.ngayXM, nguoiXM: hu.nguoiXM };
      if (req.trangThai === 'Đã xác minh') {
        if (hu.tinhTrang !== 'dangHuong') err('Chỉ xác minh "Đang hưởng lương hưu" khi tình trạng hưu trí đã khai là "Đang hưởng lương hưu".');
        if (!hu.ngayHuong || !hu.soQD) err('Nhập ngày bắt đầu hưởng và số quyết định hưởng lương hưu trước khi xác minh.');
        var hs = parseJSON(nv.hoSoHuu, {}) || {}, coTep = Object.keys(hs).some(function (nam) { var x = (hs[nam] || {}).h02; return x && (x.tep || []).length && x.trangThai !== 'Không hợp lệ'; });
        if (!coTep) err('Chưa có hồ sơ chứng minh (02_QD_huong_luong_huu) – tải quyết định hưởng lương hưu trước khi xác minh.');
      }
      if (req.trangThai === 'Không hợp lệ' && !String(req.ghiChu || '').trim()) err('Ghi lý do không hợp lệ.');
      h.huu = Object.assign({}, hu, { trangThaiXM: req.trangThai, ngayXM: env.today(), nguoiXM: s.hoTen || s.username, ghiChuXM: String(req.ghiChu || '').slice(0, 300) });
      nv.hsThue = JSON.stringify(hsThueNV({ hsThue: h })); var hh = parseJSON(nv.hsThue, {}); hh.huu.ghiChuXM = String(req.ghiChu || '').slice(0, 300); nv.hsThue = JSON.stringify(hh);
      nv.capNhat = env.now(); D('NhanVien');
      logDoi(s, 'nv.huuxm', nv.maNV, truoc, { trangThaiXM: req.trangThai, ngayXM: env.today(), nguoiXM: s.hoTen || s.username, ghiChu: req.ghiChu || '' });
      return nv;
    };
    /* v3.29: hồ sơ hưu trí 10 mục – Drive: NhanSu / <mã NV> / HoSoHuuTri / <năm> / <mục> */
    H['nv.hshuu'] = function (req, s) {
      if (!can(s.role, 'edit_nv') && !can(s.role, 'payroll') && !can(s.role, 'duyetThue')) err('Bạn không có quyền thực hiện thao tác này.');
      var nv = findNV(req.maNV); if (!nv) err('Không tìm thấy nhân viên.');
      var f = HS_HUU.filter(function (x) { return x[0] === req.muc; })[0]; if (!f) err('Mục hồ sơ không hợp lệ.');
      var nam = String(req.nam || env.today().slice(0, 4)); if (!/^\d{4}$/.test(nam)) err('Năm không hợp lệ.');
      var hs = parseJSON(nv.hoSoHuu, {}) || {}, y = hs[nam] = hs[nam] || {}, cu = y[f[0]] || { trangThai: 'Chưa cung cấp', tep: [] }, cfg = getCfg(), t = null, tt0 = cu.trangThai;
      if (req.file && req.file.data) {
        if (!MIME_OK.test(req.file.mime || '')) err('Chỉ nhận PDF, ảnh, Word, Excel.');
        var bytes = Math.floor(String(req.file.data).length * 3 / 4); if (bytes > (num(cfg.dungLuongToiDaMB) || 10) * 1048576) err('File quá lớn.');
        var name = String(req.file.name || 'tep').replace(/[\\/:*?"<>|]/g, '_').slice(0, 150);
        t = { id: newId(), maNV: nv.maNV, loai: 'Hồ sơ hưu trí', ten: (f[2] + ' ' + nam).slice(0, 150), soHieu: '', ngay: env.today(), hetHan: '', lienKet: '', nvXem: 'Có', trangThai: 'Đã xác nhận', tenFile: name, mimeType: req.file.mime, kichThuoc: bytes, nguoiTai: s.hoTen || s.username, taiLuc: env.now(), nguoiXacNhan: s.hoTen || s.username };
        t.fileId = env.files.put({ name: nv.maNV + '_' + f[1] + '_' + name, mime: req.file.mime, data: req.file.data, maNV: nv.maNV, hoTen: nv.hoTen, duong: ['NhanSu', nv.maNV, 'HoSoHuuTri', nam, f[1]] });
        T('TaiLieu').push(t); D('TaiLieu'); cu.tep = (cu.tep || []).concat([t.id]); if (cu.trangThai === 'Chưa cung cấp' || cu.trangThai === 'Cần bổ sung') cu.trangThai = 'Đã tải lên';
      }
      if (req.trangThai) {
        if (HS_TT.indexOf(req.trangThai) < 0) err('Trạng thái hồ sơ không hợp lệ.');
        if ((req.trangThai === 'Hợp lệ' || req.trangThai === 'Không hợp lệ') && !can(s.role, 'payroll') && !can(s.role, 'duyetThue')) err('Chỉ Kế toán / Phụ trách kế toán xác nhận hồ sơ hợp lệ.');
        if (req.trangThai === 'Hợp lệ' && (f[0] === 'h09' || f[0] === 'h08')) { var bm = bieuMau(cfg, f[0] === 'h09' ? 'camket' : 'uyquyen'); if (bm && bm.trangThai === 'dangxacminh') err('Mẫu "' + bm.mauSo + '" đang xác minh – "Mẫu tham khảo – cần xác minh": chưa xác nhận hồ sơ hợp lệ.'); }
        cu.trangThai = req.trangThai;
      }
      cu.nguoi = s.hoTen || s.username; cu.luc = env.now(); y[f[0]] = cu; nv.hoSoHuu = JSON.stringify(hs); nv.capNhat = env.now(); D('NhanVien');
      log(s, 'nv.hshuu', nv.maNV + ' ' + nam + ' ' + f[1] + ': ' + tt0 + ' → ' + cu.trangThai + (t ? ' (+ ' + t.tenFile + ')' : ''));
      return { nv: nv, tl: t };
    };
    H['nv.import'] = function (req, s) {
      need(s, 'edit_nv');
      var ok = 0, moi = 0, loi = [];
      (req.rows || []).forEach(function (r, i) {
        try {
          var ma = String(r.maNV || '').trim().toUpperCase();
          var isNew = !ma || !findNV(ma);
          H['nv.save']({ row: r, isNew: isNew, silent: true }, s);
          ok++; if (isNew) moi++;
        } catch (e) { loi.push('Dòng ' + (num(r._dong) || i + 2) + ': ' + e.message); }
      });
      log(s, 'nv.import', ok + ' hồ sơ (' + moi + ' mới), ' + loi.length + ' lỗi');
      return { ok: ok, moi: moi, loi: loi, nv: T('NhanVien'), qt: T('QuaTrinh') };
    };

    H['nv.delete'] = function (req, s) {
      if (s.role !== 'admin') err('Chỉ Quản trị được xóa hồ sơ. Với người nghỉ việc, hãy ghi quá trình "Nghỉ việc".');
      var ma = req.maNV;
      if (T('BangLuong').some(function (r) { return r.maNV === ma; })) err('Nhân viên đã có dữ liệu bảng lương, không thể xóa. Hãy chuyển trạng thái "Đã nghỉ việc".');
      T('TaiLieu').filter(function (x) { return x.maNV === ma; }).forEach(function (x) { try { env.files.del(x.fileId); } catch (e) { } });
      ['NhanVien', 'QuaTrinh', 'LenLuong', 'MucHuong', 'KhenThuong', 'ChamCong', 'NghiPhep', 'TaiLieu'].forEach(function (t) {
        var a = T(t); for (var i = a.length - 1; i >= 0; i--) if (a[i].maNV === ma) a.splice(i, 1);
        D(t);
      });
      var us = T('Users'); for (var j = us.length - 1; j >= 0; j--) if (us[j].maNV === ma) us.splice(j, 1);
      D('Users');
      log(s, 'nv.delete', ma);
      return true;
    };

    /* ---- phòng ban & công trình; điều động ---- */
    H['bp.save'] = function (req, s) {
      need(s, 'edit_nv');
      var cu = getCfg().boPhan, moi = [], ten = {}, ma = {};
      (req.boPhan || []).forEach(function (b, i) {
        var t = String(b.ten || '').trim().replace(/\s+/g, ' ');
        if (!t) return;
        if (ten[t.toLowerCase()]) err('Trùng tên bộ phận "' + t + '".');
        var m = String(b.ma || '').trim().toUpperCase().replace(/[^A-Z0-9_-]/g, '') || ((b.loai === 'Công trình' ? 'CT' : 'PB') + pad(i + 1));
        while (ma[m]) m += 'X';
        ten[t.toLowerCase()] = 1; ma[m] = 1;
        moi.push({ ma: m, ten: t.slice(0, 120), loai: b.loai === 'Công trình' ? 'Công trình' : 'Phòng ban', diaDiem: String(b.diaDiem || '').slice(0, 200), batDau: parseD(b.batDau) ? b.batDau : '', ketThuc: parseD(b.ketThuc) ? b.ketThuc : '', trangThai: b.trangThai === 'Đã kết thúc' ? 'Đã kết thúc' : 'Hoạt động', ghiChu: String(b.ghiChu || '').slice(0, 300) });
      });
      if (!moi.length) err('Cần ít nhất 1 bộ phận.');
      // đổi tên: cập nhật hồ sơ, quá trình, tài khoản, bảng tin theo mã
      var doiTen = {};
      cu.forEach(function (o) { var n = moi.filter(function (x) { return x.ma === o.ma; })[0]; if (n && n.ten !== o.ten) doiTen[o.ten] = n.ten; });
      var con = moi.map(function (x) { return x.ten; });
      cu.forEach(function (o) {
        if (doiTen[o.ten] || con.indexOf(o.ten) >= 0) return;
        var n = T('NhanVien').filter(function (v) { return v.phongBan === o.ten && v.trangThai !== 'Đã nghỉ việc'; }).length;
        if (n) err('"' + o.ten + '" còn ' + n + ' CBCNV. Điều động họ sang bộ phận khác hoặc chuyển bộ phận sang "Đã kết thúc".');
      });
      if (Object.keys(doiTen).length) {
        T('NhanVien').forEach(function (v) { if (doiTen[v.phongBan]) { v.phongBan = doiTen[v.phongBan]; } }); D('NhanVien');
        T('QuaTrinh').forEach(function (q) { if (doiTen[q.phongBan]) q.phongBan = doiTen[q.phongBan]; }); D('QuaTrinh');
        T('Users').forEach(function (u) { if (doiTen[u.boPhan]) u.boPhan = doiTen[u.boPhan]; }); D('Users');
        T('BaiViet').forEach(function (o) { var d = parseJSON(o.doiTuong, []); if (d.some(function (x) { return doiTen[x]; })) { o.doiTuong = JSON.stringify(d.map(function (x) { return doiTen[x] || x; })); } }); D('BaiViet');
      }
      setCfg({ boPhan: moi, phongBan: moi.map(function (x) { return x.ten; }) });
      log(s, 'bp.save', moi.length + ' bộ phận' + (Object.keys(doiTen).length ? '; đổi tên: ' + Object.keys(doiTen).map(function (k) { return k + ' → ' + doiTen[k]; }).join('; ') : ''));
      return { cfg: trangTriQT(getCfg()), nv: T('NhanVien'), qt: T('QuaTrinh'), users: T('Users').map(publicUser) };
    };
    // v3.20: danh mục công ty thành viên
    H['ctv.save'] = function (req, s) {
      need(s, 'edit_nv');
      var cu = dsCTV(getCfg()), moi = chuanCTV(req.congTyTV, getCfg());
      if (!Array.isArray(req.congTyTV) || !req.congTyTV.length) err('Cần ít nhất 1 công ty.');
      if (moi.length !== req.congTyTV.filter(function (x) { return x && String(x.ma || '').trim() && String(x.ten || '').trim(); }).length) err('Mỗi công ty cần có mã (không trùng) và tên.');
      var con = moi.map(function (x) { return x.ma; }), dau = cu[0].ma;
      cu.forEach(function (o) {
        if (con.indexOf(o.ma) >= 0) return;
        var n = T('NhanVien').filter(function (v) { return (v.congTy || dau) === o.ma; }).length;
        if (n) err('"' + o.tenTat + '" còn ' + n + ' hồ sơ CBCNV (kể cả đã nghỉ). Chuyển họ sang công ty khác, hoặc đặt trạng thái "Ngừng".');
        if (T('BangLuong').some(function (b) { return b.congTy === o.ma; })) err('"' + o.tenTat + '" đã có bảng lương hạch toán – không xóa được, chỉ đặt trạng thái "Ngừng".');
      });
      if (con[0] !== dau) {                                              // đổi công ty đầu danh mục: hồ sơ chưa khai báo vẫn giữ công ty cũ
        var k = 0; T('NhanVien').forEach(function (v) { if (!v.congTy) { v.congTy = dau; k++; } }); if (k) D('NhanVien');
      }
      setCfg({ congTyTV: moi });
      log(s, 'ctv.save', moi.map(function (x) { return x.ma + ' ' + x.tenTat; }).join('; '));
      return { cfg: trangTriQT(getCfg()), nv: T('NhanVien') };
    };
    H['bp.dieudong'] = function (req, s) {
      need(s, 'edit_nv');
      var b = getCfg().boPhan.filter(function (x) { return x.ten === req.boPhan; })[0];
      if (!b) err('Chọn bộ phận điều động đến.');
      if (b.trangThai === 'Đã kết thúc') err('"' + b.ten + '" đã kết thúc, không điều động đến được.');
      if (!parseD(req.tuNgay)) err('Nhập ngày bắt đầu.');
      var ds = (req.maNV || []).filter(function (m, i, a) { return a.indexOf(m) === i; });
      if (!ds.length) err('Chọn CBCNV cần điều động.');
      var ok = [];
      ds.forEach(function (m) {
        var nv = findNV(m); if (!nv) return;
        H['qt.save']({ row: { maNV: m, loai: 'Điều động', tuNgay: req.tuNgay, phongBan: b.ten, chucVu: req.chucVu || nv.chucVu || '', donVi: b.loai === 'Công trình' ? (b.diaDiem || nv.donVi || '') : (nv.donVi || ''), soQD: req.soQD || '', nguoiKy: req.nguoiKy || '', noiDung: req.noiDung || ('Điều động đến ' + b.ten) } }, s);
        ok.push(m);
      });
      log(s, 'bp.dieudong', ok.length + ' người → ' + b.ten + ' từ ' + req.tuNgay);
      return { nv: T('NhanVien'), qt: T('QuaTrinh'), users: T('Users').map(publicUser), soNguoi: ok.length };
    };
    /* vị trí hiện tại (bộ phận, chức vụ) theo quá trình công tác đã đến ngày hiệu lực */
    function dongBoViTri() {
      var today = env.today(), qt = T('QuaTrinh'), doi = false;
      T('NhanVien').forEach(function (nv) {
        var l = qt.filter(function (x) { return x.maNV === nv.maNV && x.tuNgay <= today && (x.loai === 'Điều động' || x.loai === 'Tiếp nhận' || x.loai === 'Bổ nhiệm' || x.loai === 'Miễn nhiệm' || x.loai === 'Thay đổi HĐLĐ'); })
          .sort(function (a, b) { return a.tuNgay < b.tuNgay ? 1 : a.tuNgay > b.tuNgay ? -1 : (a.taoLuc < b.taoLuc ? 1 : -1); })[0];
        if (l && l.phongBan && l.phongBan !== nv.phongBan) { chuyenQuanLy(nv, nv.phongBan, l.phongBan); nv.phongBan = l.phongBan; if (l.chucVu) nv.chucVu = l.chucVu; nv.capNhat = env.now(); doi = true; }
      });
      if (doi) D('NhanVien');
    }
    /* Trưởng phòng / CHT được điều động: bộ phận quản lý trên tài khoản đi theo */
    function chuyenQuanLy(nv, cu, moi) {
      T('Users').forEach(function (u) { if (u.maNV === nv.maNV && ROLE_QL.indexOf(normRole(u.role)) >= 0 && (!u.boPhan || u.boPhan === cu)) { u.boPhan = moi; D('Users'); } });
    }

    /* ---- quá trình công tác ---- */
    function goiNV(ma) {
      return { nv: findNV(ma), qt: T('QuaTrinh').filter(function (x) { return x.maNV === ma; }), users: T('Users').filter(function (u) { return u.maNV === ma; }).map(publicUser) };
    }
    H['qt.save'] = function (req, s) {
      need(s, 'edit_nv');
      var r = req.row || {}, nv = findNV(r.maNV);
      if (!nv) err('Không tìm thấy nhân viên.');
      if (LOAI_QT.indexOf(r.loai) < 0) err('Loại quá trình không hợp lệ.');
      if (!parseD(r.tuNgay)) err('Thiếu ngày bắt đầu.');
      if (r.denNgay && r.denNgay < r.tuNgay) err('Ngày kết thúc phải sau ngày bắt đầu.');
      var rows = T('QuaTrinh'), o = r.id ? rows.filter(function (x) { return x.id === r.id; })[0] : null;
      var isNew = !o;
      if (!o) { o = { id: newId(), taoBoi: s.username, taoLuc: env.now() }; rows.push(o); }
      ['maNV', 'loai', 'tuNgay', 'denNgay', 'phongBan', 'chucVu', 'donVi', 'soQD', 'nguoiKy', 'noiDung'].forEach(function (c) { o[c] = String(r[c] == null ? '' : r[c]).trim(); });

      if (isNew && o.loai !== 'Đào tạo' && o.loai !== 'Khác') {
        // khép lại quá trình đang mở trước đó
        rows.forEach(function (x) {
          if (x !== o && x.maNV === o.maNV && !x.denNgay && x.tuNgay < o.tuNgay && x.loai !== 'Đào tạo' && x.loai !== 'Khác') x.denNgay = addDays(o.tuNgay, -1);
        });
      }
      // cập nhật vị trí hiện tại nếu đây là quá trình mới nhất đã có hiệu lực
      var today = env.today();
      var latest = rows.filter(function (x) { return x.maNV === o.maNV && x.tuNgay <= today && x.loai !== 'Đào tạo' && x.loai !== 'Khác'; })
        .sort(function (a, b) { return a.tuNgay < b.tuNgay ? 1 : a.tuNgay > b.tuNgay ? -1 : (a.taoLuc < b.taoLuc ? 1 : -1); })[0];
      if (latest === o) {
        if (o.phongBan && o.phongBan !== nv.phongBan) { chuyenQuanLy(nv, nv.phongBan, o.phongBan); nv.phongBan = o.phongBan; }
        if (o.chucVu) nv.chucVu = o.chucVu;
        if (o.donVi) nv.donVi = o.donVi;
        if (o.loai === 'Tạm nghỉ') nv.trangThai = 'Tạm nghỉ';
        else if (nv.trangThai === 'Tạm nghỉ') nv.trangThai = 'Đang làm việc';
        nv.capNhat = env.now(); D('NhanVien');
      }
      D('QuaTrinh');
      apChamDut(nv);                                  // v3.18: Nghỉ việc / Thanh lý HĐ → ngày chấm dứt, trạng thái, khóa tài khoản (kể cả khai trước ngày)
      log(s, 'qt.save', o.maNV + ' ' + o.loai + ' ' + o.tuNgay);
      var g = goiNV(o.maNV); g.row = o; return g;
    };

    H['qt.delete'] = function (req, s) {
      need(s, 'edit_nv');
      var a = T('QuaTrinh');
      var ma = '';
      for (var i = a.length - 1; i >= 0; i--) if (a[i].id === req.id) { ma = a[i].maNV; log(s, 'qt.delete', a[i].maNV + ' ' + a[i].loai); a.splice(i, 1); }
      D('QuaTrinh');
      var nvX = ma && findNV(ma); if (nvX) apChamDut(nvX);          // xóa bản ghi chấm dứt nhập nhầm: mở lại hồ sơ, tài khoản
      return nvX ? goiNV(ma) : true;
    };

    /* ---- lên lương ---- */
    H['ll.save'] = function (req, s) {
      need(s, 'propose');
      var r = req.row || {}, nv = findNV(r.maNV);
      if (!nv) err('Không tìm thấy nhân viên.');
      if (LOAI_LL.indexOf(r.loai) < 0) err('Loại điều chỉnh không hợp lệ.');
      if (!parseD(r.hieuLuc)) err('Thiếu ngày hiệu lực.');
      if (num(r.luongMoi) <= 0) err('Mức lương mới phải lớn hơn 0.');
      var rows = T('LenLuong'), o = r.id ? rows.filter(function (x) { return x.id === r.id; })[0] : null;
      if (o && o.trangThai !== 'Chờ duyệt') err('Đề xuất đã được xử lý, không sửa được.');
      if (!o) {
        o = { id: newId(), trangThai: 'Chờ duyệt', apDung: '', nguoiDeXuat: s.username, ngayDeXuat: env.now(),
          ngachCu: nv.ngach || '', bacCu: nv.bac || '', luongCu: num(nv.luongChinh), luongBHCu: num(nv.luongBH) };
        rows.push(o);
      }
      ['maNV', 'loai', 'hieuLuc', 'ngachMoi', 'bacMoi', 'lyDo', 'soQD'].forEach(function (c) { o[c] = String(r[c] == null ? '' : r[c]).trim(); });
      o.luongMoi = num(r.luongMoi); o.luongBHMoi = num(r.luongBHMoi);
      D('LenLuong');
      log(s, 'll.save', o.maNV + ' → ' + o.luongMoi);
      return o;
    };

    H['ll.decide'] = function (req, s) {
      need(s, 'approve');
      var o = T('LenLuong').filter(function (x) { return x.id === req.id; })[0];
      if (!o) err('Không tìm thấy đề xuất.');
      if (o.trangThai !== 'Chờ duyệt') err('Đề xuất đã được xử lý.');
      o.trangThai = req.approve ? 'Đã duyệt' : 'Từ chối';
      o.nguoiDuyet = s.username; o.ngayDuyet = env.now(); o.ghiChuDuyet = String(req.note || '');
      if (req.approve && req.soQD) o.soQD = String(req.soQD);
      D('LenLuong');
      if (req.approve) applyDueRaises();
      log(s, 'll.' + (req.approve ? 'approve' : 'reject'), o.maNV);
      return { row: o, nv: findNV(o.maNV) };
    };

    H['ll.delete'] = function (req, s) {
      var a = T('LenLuong');
      for (var i = a.length - 1; i >= 0; i--) if (a[i].id === req.id) {
        if (a[i].trangThai === 'Đã duyệt' && s.role !== 'admin') err('Chỉ Quản trị được xóa quyết định đã duyệt.');
        if (a[i].trangThai !== 'Đã duyệt') need(s, 'propose');
        var nvX = findNV(a[i].maNV);
        if (nvX && a[i].trangThai === 'Đã duyệt') {
          var cuX = {}; if (num(a[i].luongCu) > 0) cuX.luongChinh = num(a[i].luongCu); if (num(a[i].luongBHCu) > 0) cuX.luongBH = num(a[i].luongBHCu);
          mhKhoiPhuc(nvX, ['luongChinh', 'luongBH'], cuX, a[i].id, a[i].hieuLuc);
        }
        log(s, 'll.delete', a[i].maNV); a.splice(i, 1);
        if (nvX) dongBoMucHuong(nvX);
      }
      D('LenLuong'); return true;
    };

    /* ---- khai báo mức hưởng theo thời gian ---- */
    function mhGiaTri(gt, cfg) {
      var out = {}, pcMa = (cfg.phuCap || []).map(function (p) { return 'pc.' + p.ma; });
      Object.keys(gt || {}).forEach(function (k) {
        var v = gt[k];
        if (v === '' || v === null || v === undefined) return;
        if (!MH_KHOA[k] && pcMa.indexOf(k) < 0) err('Khoản "' + k + '" không có trong danh mục.');
        v = num(v);
        if (v < 0) err('Giá trị không được âm.');
        if (k === 'soNPT' && (v !== Math.floor(v) || v > 20)) err('Số người phụ thuộc phải là số nguyên.');
        out[k] = v;
      });
      if (!Object.keys(out).length) err('Nhập ít nhất một mức thay đổi.');
      return out;
    }
    function mhCu(nv, o) {
      var m = mucHuong(nv, T('LenLuong'), T('MucHuong'), addDays(o.hieuLuc, -1), o.id), cu = {};
      Object.keys(parseJSON(o.giaTri, {})).forEach(function (k) { cu[k] = k.indexOf('pc.') === 0 ? num(m.phuCap[k.slice(3)]) : num(m[k]); });
      o.cu = JSON.stringify(cu);
    }
    /* bỏ một khoản khỏi diễn biến: nếu không còn mốc nào khai khoản đó thì hồ sơ trở về mức trước mốc bị bỏ */
    function mhKhoiPhuc(nv, keys, cu, boQua, hieuLuc) {
      var con = mhNguon(nv, T('LenLuong'), T('MucHuong'), boQua), pc = parseJSON(nv.phuCap, {});
      keys.forEach(function (k) {
        if (!(k in cu)) return;
        // mốc kế tiếp có khai khoản này nhận "mức cũ" của mốc bị bỏ
        var sau = con.filter(function (x) { return (k in x.gt) && x.hieuLuc >= (hieuLuc || ''); })[0];
        if (sau && sau.nguon === 'mh') { var r = T('MucHuong').filter(function (z) { return z.id === sau.id; })[0], c2 = parseJSON(r.cu, {}); c2[k] = num(cu[k]); r.cu = JSON.stringify(c2); D('MucHuong'); }
        if (sau && sau.nguon === 'll') { var l = T('LenLuong').filter(function (z) { return z.id === sau.id; })[0]; if (k === 'luongChinh') l.luongCu = num(cu[k]); if (k === 'luongBH') l.luongBHCu = num(cu[k]); D('LenLuong'); }
        if (con.some(function (x) { return k in x.gt; })) return;
        if (k.indexOf('pc.') === 0) { if (num(cu[k])) pc[k.slice(3)] = num(cu[k]); else delete pc[k.slice(3)]; }
        else nv[k] = num(cu[k]);
      });
      nv.phuCap = JSON.stringify(pc); D('NhanVien');
    }
    function mhSauThayDoi(nv, hieuLuc, s) {
      dongBoMucHuong(nv);
      var tinhLai = [];
      T('KyLuong').forEach(function (k) { var kr = kyRow(k.ky, false); if (kyEnd(k.ky) >= hieuLuc && canTinhLai(kr)) { recalc(k.ky, s); tinhLai.push(k.ky); } });
      return tinhLai;
    }
    function mhLuu(r, s) {
      if (!can(s.role, 'edit_nv') && !can(s.role, 'approve')) err('Không có quyền khai báo mức hưởng.');
      var nv = findNV(r.maNV); if (!nv) err('Không tìm thấy nhân viên ' + (r.maNV || ''));
      var hl = String(r.hieuLuc || '').slice(0, 10); if (!parseD(hl)) err('Thiếu ngày hiệu lực.');
      var gt = mhGiaTri(parseJSON(r.giaTri, {}), getCfg());
      var rows = T('MucHuong'), o = r.id ? rows.filter(function (x) { return x.id === r.id; })[0] : null;
      if (r.id && !o) err('Không tìm thấy khai báo.');
      if (o && o.trangThai === 'Hiệu lực' && s.role !== 'admin') err('Khai báo đã có hiệu lực, chỉ Quản trị được sửa. Hãy lập khai báo mới.');
      if (o && o.trangThai === 'Từ chối') err('Khai báo đã bị từ chối, hãy lập khai báo mới.');
      if (!o) { o = { id: newId(), nguoiTao: s.username, taoLuc: env.now() }; rows.push(o); }
      else if (o.trangThai === 'Hiệu lực') {
        if (o.maNV !== nv.maNV) err('Không đổi CBCNV của khai báo đã hiệu lực.');
        var bo = Object.keys(parseJSON(o.giaTri, {})).filter(function (k) { return !(k in gt); });
        if (bo.length) mhKhoiPhuc(nv, bo, parseJSON(o.cu, {}), o.id, o.hieuLuc);
      }
      o.maNV = nv.maNV; o.hieuLuc = hl; o.giaTri = JSON.stringify(gt);
      o.soQD = String(r.soQD || '').trim(); o.lyDo = String(r.lyDo || '').trim();
      var canDuyet = Object.keys(gt).some(function (k) { return MH_LUONG.indexOf(k) >= 0; }) && !can(s.role, 'approve');
      if (canDuyet) { o.trangThai = 'Chờ duyệt'; o.nguoiDuyet = ''; o.ngayDuyet = ''; }
      else { o.trangThai = 'Hiệu lực'; o.nguoiDuyet = s.username; o.ngayDuyet = env.now(); }
      mhCu(nv, o);
      D('MucHuong');
      return { o: o, nv: nv };
    }
    H['mh.save'] = function (req, s) {
      var x = mhLuu(req.row || {}, s), tinhLai = x.o.trangThai === 'Hiệu lực' ? mhSauThayDoi(x.nv, x.o.hieuLuc, s) : [];
      log(s, 'mh.save', x.o.maNV + ' từ ' + x.o.hieuLuc + ': ' + x.o.giaTri + ' (' + x.o.trangThai + ')');
      return { row: x.o, nv: x.nv, tinhLai: tinhLai };
    };
    H['mh.import'] = function (req, s) {
      var ok = 0, loi = [], ds = (req.rows || []).slice().sort(function (a, b) { return String(a.hieuLuc) < String(b.hieuLuc) ? -1 : 1; }), dsNV = {}, minHL = '9999';
      ds.forEach(function (r, i) {
        try { var x = mhLuu(r, s); ok++; if (x.o.trangThai === 'Hiệu lực') { dsNV[x.nv.maNV] = x.nv; if (x.o.hieuLuc < minHL) minHL = x.o.hieuLuc; } }
        catch (e) { loi.push((r.maNV || 'dòng ' + (i + 2)) + ' ' + (r.hieuLuc || '') + ': ' + String(e.message || e)); }
      });
      var tinhLai = [];
      Object.keys(dsNV).forEach(function (m) { dongBoMucHuong(dsNV[m]); });
      if (ok) T('KyLuong').forEach(function (k) { var kr = kyRow(k.ky, false); if (kyEnd(k.ky) >= minHL && canTinhLai(kr)) { recalc(k.ky, s); tinhLai.push(k.ky); } });
      log(s, 'mh.import', ok + ' dòng' + (loi.length ? ', lỗi ' + loi.length : ''));
      return { ok: ok, loi: loi, tinhLai: tinhLai };
    };
    H['mh.decide'] = function (req, s) {
      need(s, 'approve');
      var o = T('MucHuong').filter(function (x) { return x.id === req.id; })[0];
      if (!o) err('Không tìm thấy khai báo.');
      if (o.trangThai !== 'Chờ duyệt') err('Khai báo đã được xử lý.');
      var nv = findNV(o.maNV); if (!nv) err('Không tìm thấy hồ sơ ' + o.maNV);
      o.trangThai = req.approve ? 'Hiệu lực' : 'Từ chối'; o.nguoiDuyet = s.username; o.ngayDuyet = env.now(); o.ghiChuDuyet = String(req.note || '');
      if (req.approve) mhCu(nv, o);
      D('MucHuong');
      var tinhLai = req.approve ? mhSauThayDoi(nv, o.hieuLuc, s) : [];
      log(s, 'mh.' + (req.approve ? 'approve' : 'reject'), o.maNV + ' ' + o.hieuLuc);
      return { row: o, nv: nv, tinhLai: tinhLai };
    };
    H['mh.delete'] = function (req, s) {
      var a = T('MucHuong'), o = a.filter(function (x) { return x.id === req.id; })[0];
      if (!o) err('Không tìm thấy khai báo.');
      if (o.trangThai === 'Hiệu lực' && s.role !== 'admin') err('Chỉ Quản trị được xóa khai báo đã có hiệu lực.');
      if (o.trangThai !== 'Hiệu lực' && !can(s.role, 'edit_nv') && o.nguoiTao !== s.username) err('Không có quyền xóa.');
      var nv = findNV(o.maNV);
      if (nv && o.trangThai === 'Hiệu lực') mhKhoiPhuc(nv, Object.keys(parseJSON(o.giaTri, {})), parseJSON(o.cu, {}), o.id, o.hieuLuc);
      a.splice(a.indexOf(o), 1); D('MucHuong');
      var tinhLai = (nv && o.trangThai === 'Hiệu lực') ? mhSauThayDoi(nv, o.hieuLuc, s) : [];
      log(s, 'mh.delete', o.maNV + ' ' + o.hieuLuc + ' ' + o.giaTri);
      return { nv: nv, tinhLai: tinhLai };
    };

    /* ==================== NHẮN TIN NỘI BỘ ====================
       TinNhan: chỉ ghi thêm (append-only) qua env.chat — seq = số dòng trên sheet, dùng làm con trỏ đồng bộ.
       Thu hồi, cảm xúc ghi thành dòng sự kiện (loai thuhoi / camxuc, traLoi = id tin gốc).
       HoiThoai, ThanhVien, TepChat: bảng nhỏ qua SheetDB. Gõ phím, trực tuyến, tín hiệu đổi danh sách: bộ nhớ đệm (env.chat). */
    var CHAT_TEP_MIME = /^(image\/(jpeg|png|webp|gif)|application\/pdf|application\/(msword|vnd\.openxmlformats-officedocument\.(wordprocessingml\.document|spreadsheetml\.sheet|presentationml\.presentation))|application\/vnd\.ms-(excel|powerpoint)|application\/zip|application\/x-zip-compressed|text\/plain)$/;
    var CHAT_MAU = ['#22386A', '#C2570C', '#15803D', '#7C3AED', '#0E7490', '#B91C1C', '#A16207', '#BE185D'];
    function chatTVs(u) { return T('ThanhVien').filter(function (x) { return x.username === u && !x.roi; }); }
    function chatHT(id) { return T('HoiThoai').filter(function (x) { return x.id === id; })[0]; }
    function chatTV(htId, u) { return T('ThanhVien').filter(function (x) { return x.hoiThoai === htId && x.username === u && !x.roi; })[0]; }
    function chatCan(s, htId) {
      var h = chatHT(htId); if (!h) err('Không tìm thấy cuộc trò chuyện.');
      var tv = chatTV(htId, s.username); if (!tv) err('Bạn không còn trong cuộc trò chuyện này.');
      return { h: h, tv: tv };
    }
    function chatMem(htId) { return T('ThanhVien').filter(function (x) { return x.hoiThoai === htId && !x.roi; }); }
    function chatPub(h) {
      return { id: h.id, loai: h.loai, ten: h.ten, mau: h.mau, moTa: h.moTa, taoBoi: h.taoBoi, taoLuc: h.taoLuc, capNhat: h.capNhat, tinCuoi: parseJSON(h.tinCuoi, null), trangThai: h.trangThai,
        tv: chatMem(h.id).map(function (x) { return { username: x.username, vaiTro: x.vaiTro, daDoc: num(x.daDoc), thamGia: x.thamGia }; }) };
    }
    function chatHTDs(s) {
      var my = chatTVs(s.username), out = [];
      my.forEach(function (tv) { var h = chatHT(tv.hoiThoai); if (h) { var p = chatPub(h); p.ghim = tv.ghim === 'Có'; p.tatTB = tv.tatTB === 'Có'; p.daDoc = num(tv.daDoc); out.push(p); } });
      return out;
    }
    function chatDir() {
      var nv = {}; T('NhanVien').forEach(function (n) { nv[n.maNV] = n; });
      return T('Users').filter(function (u) { return u.active !== 'Không' && u.username !== 'system'; }).map(function (u) {
        var n = nv[u.maNV] || {};
        return { username: u.username, hoTen: u.hoTen || n.hoTen || u.username, phongBan: n.phongBan || ROLES[u.role] || '', chucVu: n.chucVu || '', maNV: u.maNV || '' };
      });
    }
    function chatBao(users) { env.chat.bump(users); }
    function chatGhi(s, htId, loai, noiDung, extra) {
      var row = { id: newId(), hoiThoai: htId, username: s.username, hoTen: s.hoTen || s.username, loai: loai, noiDung: noiDung || '', tep: '', traLoi: '', thoiGian: env.now() };
      Object.keys(extra || {}).forEach(function (k) { row[k] = extra[k]; });
      row.seq = env.chat.append(row);
      if (CHAT_PUSH.hasOwnProperty(loai)) chatDay(s, htId, row);
      if (loai !== 'camxuc') {
        var h = chatHT(htId);
        if (h) { h.capNhat = row.thoiGian; h.tinCuoi = JSON.stringify({ seq: row.seq, id: row.id, username: row.username, hoTen: row.hoTen, loai: loai, noiDung: String(noiDung || '').slice(0, 120), thoiGian: row.thoiGian }); D('HoiThoai'); }
      }
      return row;
    }
    var CHAT_PUSH = { kho: '[Phiếu kho] ', text: '', anh: '[Hình ảnh]', tep: '[Tệp] ', binhchon: '[Bình chọn] ', nhac: '[Nhắc hẹn] ', baocao: '[Yêu cầu báo cáo] ', bcnop: 'đã nộp báo cáo: ', nhactao: 'tạo nhắc hẹn: ', bctao: 'tạo yêu cầu báo cáo: ' };
    function chatDay(s, htId, m) {                     // thông báo đẩy tin nhắn: mọi thành viên (trừ người gửi, người tắt thông báo nhóm)
      if (!pushBat()) return;
      var h = chatHT(htId); if (!h || h.trangThai) return;
      var tuGui = m.loai === 'nhac' || m.loai === 'baocao';
      var ds = chatMem(htId).filter(function (t) { return t.tatTB !== 'Có' && (tuGui || t.username !== s.username); }).map(function (t) { return t.username; });
      var ten = s.hoTen || s.username, nd = m.loai === 'anh' ? CHAT_PUSH.anh : m.loai === 'tep' ? CHAT_PUSH.tep + (parseJSON(m.tep, {}).ten || '') : CHAT_PUSH[m.loai] + String(m.noiDung || '').replace(/\s+/g, ' ');
      var nhom = h.loai === 'nhom';
      bao(ds, { loai: 'tn', t: nhom ? (h.ten || 'Nhóm') : ten, b: (nhom && !tuGui ? ten + ': ' : '') + nd, tag: 'tn_' + htId, mo: '#chat=' + htId });
    }
    function chatHeThong(s, htId, nd) { return chatGhi(s, htId, 'hethong', nd); }
    function chatTen(u) { var x = T('Users').filter(function (z) { return z.username === u; })[0]; return x ? (x.hoTen || u) : u; }
    function chatThem(htId, u, vaiTro) {
      var a = T('ThanhVien'), cu = a.filter(function (x) { return x.hoiThoai === htId && x.username === u; })[0];
      if (cu) { if (!cu.roi) return false; cu.roi = ''; cu.vaiTro = vaiTro || 'thanhvien'; cu.thamGia = env.now(); cu.daDoc = env.chat.last(); }
      else a.push({ id: newId(), hoiThoai: htId, username: u, vaiTro: vaiTro || 'thanhvien', daDoc: env.chat.last(), ghim: '', tatTB: '', thamGia: env.now(), roi: '' });
      D('ThanhVien'); return true;
    }
    function chatUserOK(u) { var x = T('Users').filter(function (z) { return z.username === u; })[0]; return !!x && x.active !== 'Không'; }
    function chatTruong(s, htId) {
      var c = chatCan(s, htId);
      if (c.h.loai !== 'nhom') err('Chỉ áp dụng cho nhóm.');
      if (c.h.trangThai) err('Nhóm đã giải tán.');
      return c;
    }
    function laTruong(s, c) { return c.tv.vaiTro === 'truong' || s.role === 'admin'; }

    H['chat.init'] = function (req, s) {
      env.chat.online(s.username);
      var ht = chatHTDs(s), ids = {}, last = env.chat.last();
      ht.forEach(function (h) { ids[h.id] = 1; });
      var msgs = env.chat.before(last + 1, 3000).filter(function (m) { return ids[m.hoiThoai]; });
      if (env.lichTrongPoll) chayLich();
      if (env.chat.ghiNho) env.chat.ghiNho(s.username, ht.map(function (h) { return h.id; }), last);
      return { cur: last, hv: env.chat.sig(s.username), ht: ht, msgs: msgs, dir: chatDir(), muc: mucDs(s), _tf: tfTuMsgs(msgs) };
    };
    H['chat.poll'] = function (req, s) {
      env.chat.online(s.username);
      if (env.lichTrongPoll) chayLich();
      if (req.typing && chatTV(req.typing, s.username)) env.chat.typing(req.typing, s.username, s.hoTen || s.username);
      var last = env.chat.last(), cur = num(req.cur), hv = env.chat.sig(s.username), out = { cur: last, hv: hv };
      if (cur > last) return { reset: true };
      var ht = null;
      if (last > cur || String(req.hv) !== String(hv)) { ht = chatHTDs(s); out.ht = ht; out.muc = mucDs(s); }
      if (last > cur) {
        var ids = {}; ht.forEach(function (h) { ids[h.id] = 1; });
        var rows = env.chat.since(cur, 2000);
        out.msgs = rows.filter(function (m) { return ids[m.hoiThoai]; }); out._tf = tfTuMsgs(out.msgs);
        if (rows.length && rows[rows.length - 1].seq < last) out.cur = rows[rows.length - 1].seq;
      }
      var myHT = ht || chatTVs(s.username).map(function (t) { return { id: t.hoiThoai }; });
      out.typing = env.chat.typingOf(myHT.map(function (h) { return h.id; }), s.username);
      if (env.chat.ghiNho && out.cur === last) env.chat.ghiNho(s.username, myHT.map(function (h) { return h.id; }), last);
      if (req.online) out.online = env.chat.onlineOf(chatDir().map(function (u) { return u.username; }));
      return out;
    };
    H['chat.history'] = function (req, s) {
      chatCan(s, req.ht);
      var before = num(req.before) || env.chat.last() + 1, want = Math.min(num(req.n) || 50, 200), out = [], vong = 0;
      while (before > 2 && vong < 20) {
        var chunk = env.chat.before(before, 1000);
        if (!chunk.length) break;
        var f = chunk.filter(function (m) { return m.hoiThoai === req.ht; });
        out = f.concat(out); before = chunk[0].seq; vong++;
        if (out.filter(function (m) { return m.loai !== 'camxuc' && m.loai !== 'thuhoi'; }).length >= want) break;
      }
      return { msgs: out, before: before, het: before <= 2, _tf: tfTuMsgs(out) };
    };
    H['chat.send'] = function (req, s) {
      var c = chatCan(s, req.ht);
      if (c.h.trangThai) err('Nhóm đã giải tán, không gửi được tin nhắn.');
      if (c.h.loai === 'rieng') { var ban = chatMem(c.h.id).filter(function (x) { return x.username !== s.username; })[0]; if (ban && !chatUserOK(ban.username)) err('Tài khoản người nhận đã bị khóa.'); }
      var nd = String(req.noiDung || '').replace(/\s+$/, ''), tep = null, loai = 'text';
      if (req.tep) {
        var t = T('TepChat').filter(function (x) { return x.fileId === req.tep.fileId; })[0];
        if (!t || t.hoiThoai !== c.h.id || t.nguoiTai !== s.username) err('Tệp không hợp lệ.');
        if (t.msgId) err('Tệp đã được gửi.');
        tep = { fileId: t.fileId, ten: t.ten, mime: t.mime, kichThuoc: num(t.kichThuoc), w: num(req.tep.w), h: num(req.tep.h) };
        if (t.thumbId) tep.thumbId = t.thumbId;
        loai = /^image\//.test(t.mime) ? 'anh' : 'tep';
      }
      if (!nd.trim() && !tep) err('Tin nhắn trống.');
      if (nd.length > 5000) err('Tin nhắn tối đa 5.000 ký tự.');
      var tr = String(req.traLoi || '').slice(0, 40);
      var m = chatGhi(s, c.h.id, loai, nd, { tep: tep ? JSON.stringify(tep) : '', traLoi: tr });
      if (tep) { t.msgId = m.id; t.seq = m.seq; t.w = tep.w; t.h = tep.h; D('TepChat'); }
      c.tv.daDoc = m.seq; D('ThanhVien');
      return { msg: m };
    };
    /* v3.32: tải tệp + gửi tin trong 1 lần gọi (req.gui); ảnh kèm ảnh thu nhỏ (f.thumb) để khung chat tải nhanh.
       Trên Apps Script, tệp được ghi vào Drive TRƯỚC khi lấy khóa (req._pre do máy chủ đặt, không nhận từ máy khách) – không chặn người khác. */
    H['chat.upload'] = function (req, s) {
      var c = chatCan(s, req.ht), f = req.file || {}, cfg = getCfg(), pre = req._pre || null;
      if (c.h.trangThai) err('Nhóm đã giải tán.');
      var name = String(f.name || '').trim().replace(/[\\/:*?"<>|]/g, '_').slice(0, 150);
      if (!name || (!f.data && !pre)) err('Chưa chọn file.');
      if (!CHAT_TEP_MIME.test(f.mime || '')) err('Chỉ gửi được ảnh, PDF, Word, Excel, PowerPoint, ZIP, TXT.');
      var bytes = pre ? num(pre.bytes) : Math.floor(String(f.data).length * 3 / 4);
      if (bytes > (num(cfg.dungLuongToiDaMB) || 10) * 1048576) err('File vượt ' + (num(cfg.dungLuongToiDaMB) || 10) + ' MB.');
      var fd = chatThuMuc(c.h), id, thumbId = '', anh = /^image\//.test(f.mime);
      if (pre) {
        id = pre.id; thumbId = anh ? pre.thumbId || '' : '';
        if (fd && pre.folder !== fd && env.files.move) { try { env.files.move(id, fd); if (thumbId) env.files.move(thumbId, fd); } catch (e) { } }
      } else {
        id = env.files.put({ name: env.today() + '_' + name, mime: f.mime, data: f.data, chat: true, folderId: fd });
        if (anh && f.thumb && String(f.thumb).length < 400000) { try { thumbId = env.files.put({ name: env.today() + '_' + name.replace(/(\.[^.]+)?$/, '_thumb.jpg'), mime: 'image/jpeg', data: f.thumb, chat: true, folderId: fd }); } catch (e) { thumbId = ''; } }
      }
      var row = { fileId: id, hoiThoai: c.h.id, ten: name, mime: f.mime, kichThuoc: bytes, nguoiTai: s.username, thoiGian: env.now(), msgId: '', seq: 0, w: 0, h: 0, thuMuc: fd, thumbId: thumbId };
      T('TepChat').push(row); D('TepChat');
      var out = { fileId: id, ten: name, mime: f.mime, kichThuoc: bytes, thumbId: thumbId, _tf: {}, _fd: {} };
      out._tf[id] = { h: c.h.id, t: thumbId, m: f.mime, n: name }; if (fd) out._fd[c.h.id] = fd;
      if (req.gui) {
        if (c.h.loai === 'rieng') { var ban = chatMem(c.h.id).filter(function (x) { return x.username !== s.username; })[0]; if (ban && !chatUserOK(ban.username)) err('Tài khoản người nhận đã bị khóa.'); }
        var tep = { fileId: id, ten: name, mime: f.mime, kichThuoc: bytes, w: num(f.w), h: num(f.h) }; if (thumbId) tep.thumbId = thumbId;
        var m = chatGhi(s, c.h.id, anh ? 'anh' : 'tep', '', { tep: JSON.stringify(tep), traLoi: String(req.traLoi || '').slice(0, 40) });
        row.msgId = m.id; row.seq = m.seq; row.w = tep.w; row.h = tep.h;
        c.tv.daDoc = m.seq; D('ThanhVien');
        out.msg = m;
      }
      return out;
    };
    /* v3.32: bản đồ tệp → cuộc trò chuyện cho đường đọc ảnh nhanh của Apps Script (bộ nhớ đệm, không mở Google Sheet) */
    function tfTuMsgs(msgs) {
      var o = {}, n = 0;
      (msgs || []).forEach(function (m) {
        if (n > 400 || !m.tep || (m.loai !== 'anh' && m.loai !== 'tep' && m.loai !== 'kho')) return;
        var t = parseJSON(m.tep, null); if (!t) return;
        if (t.fileId) { o[t.fileId] = { h: m.hoiThoai, t: t.thumbId || '', m: t.mime || '', n: t.ten || '' }; n++; }
        (t.anh || []).forEach(function (a) { if (a && a.fileId) { o[a.fileId] = { h: m.hoiThoai, t: a.thumbId || '', m: 'image/jpeg', n: a.ten || '' }; n++; } });
      });
      return o;
    }
    function tepDoc(t, thumb) {
      if (thumb && /^image\//.test(t.mime)) {
        if (t.thumbId) { try { var d0 = env.files.get(t.thumbId); if (d0 && d0.data) return { fileId: t.fileId, data: d0.data, mime: 'image/jpeg', name: t.ten, thumb: 1 }; } catch (e) { } }
        else if (env.files.thumb) { try { var d1 = env.files.thumb(t.fileId); if (d1 && d1.data) return { fileId: t.fileId, data: d1.data, mime: d1.mime || 'image/png', name: t.ten, thumb: 1 }; } catch (e) { } }
      }
      var d = env.files.get(t.fileId); if (!d) err('Tệp không còn trên máy chủ.');
      return { fileId: t.fileId, data: d.data, mime: t.mime, name: t.ten, thumb: 0 };
    }
    H['chat.anh'] = function (req, s) {        // v3.32: tải gộp tối đa 8 ảnh (thu nhỏ) trong 1 lần gọi
      var ids = (Array.isArray(req.ids) ? req.ids : []).map(String).slice(0, 8), by = {}, ok = {}, out = [], tf = {};
      T('TepChat').forEach(function (x) { if (ids.indexOf(x.fileId) >= 0) by[x.fileId] = x; });
      ids.forEach(function (id) {
        var t = by[id]; if (!t) { out.push({ fileId: id, loi: 'Tệp không còn (có thể đã thu hồi).' }); return; }
        if (!(t.hoiThoai in ok)) { try { chatCan(s, t.hoiThoai); ok[t.hoiThoai] = true; } catch (e) { ok[t.hoiThoai] = false; } }
        if (!ok[t.hoiThoai]) { out.push({ fileId: id, loi: 'Bạn không còn trong cuộc trò chuyện này.' }); return; }
        try { out.push(tepDoc(t, !!req.thumb)); tf[id] = { h: t.hoiThoai, t: t.thumbId || '', m: t.mime, n: t.ten }; } catch (e) { out.push({ fileId: id, loi: e.message }); }
      });
      return { ds: out, _tf: tf };
    };
    /* thư mục Drive riêng cho từng cuộc trò chuyện: "_TIN NHẮN NỘI BỘ/<tên nhóm> [mã]" */
    function chatTenThuMuc(h) {
      var ma = ' [' + String(h.id).slice(0, 6) + ']';
      if (h.loai === 'nhom') return String(h.ten || 'Nhóm').replace(/[\\/:*?"<>|]/g, '_').slice(0, 80) + ma;
      return 'Riêng – ' + chatMem(h.id).concat(T('ThanhVien').filter(function (x) { return x.hoiThoai === h.id && x.roi; })).map(function (x) { return chatTen(x.username); })
        .filter(function (v, i, a) { return a.indexOf(v) === i; }).sort().join(' & ').slice(0, 80) + ma;
    }
    function chatThuMuc(h) {
      if (h.loai === 'nhom' && env.luuTru) { try { var lc = ltCay(h, true); if (lc) return lc.tep; } catch (e) { } }      // nhóm: "Lưu trữ nhóm chat/<nhóm>/1. Tệp & ảnh"
      if (h.thuMuc || !env.files.folder) return h.thuMuc || '';
      try { h.thuMuc = env.files.folder(chatTenThuMuc(h)); D('HoiThoai'); } catch (e) { return ''; }
      return h.thuMuc;
    }
    H['chat.files'] = function (req, s) {
      var my = {}; chatTVs(s.username).forEach(function (t) { my[t.hoiThoai] = 1; });
      if (req.ht && !my[req.ht]) err('Bạn không còn trong cuộc trò chuyện này.');
      var q = String(req.q || '').toLowerCase(), loai = req.loai, ng = String(req.nguoi || ''), tu = String(req.tu || ''), den = String(req.den || '');
      var ds = T('TepChat').filter(function (t) {
        if (!my[t.hoiThoai] || (req.ht && t.hoiThoai !== req.ht)) return false;
        var anh = /^image\//.test(t.mime);
        if (loai === 'anh' && !anh) return false;
        if (loai === 'tep' && anh) return false;
        if (ng && t.nguoiTai !== ng) return false;
        if (tu && t.thoiGian.slice(0, 10) < tu) return false;
        if (den && t.thoiGian.slice(0, 10) > den) return false;
        return !q || String(t.ten).toLowerCase().indexOf(q) >= 0;
      }).sort(function (a, b) { return a.thoiGian < b.thoiGian ? 1 : -1; });
      var tu0 = Math.max(0, num(req.tu0)), n = Math.min(num(req.n) || 120, 300);
      var out = { tong: ds.length, dungLuong: ds.reduce(function (a, t) { return a + num(t.kichThuoc); }, 0), ds: ds.slice(tu0, tu0 + n).map(function (t) {
        return { fileId: t.fileId, hoiThoai: t.hoiThoai, ten: t.ten, mime: t.mime, kichThuoc: num(t.kichThuoc), nguoiTai: t.nguoiTai, hoTen: chatTen(t.nguoiTai), thoiGian: t.thoiGian, msgId: t.msgId, seq: num(t.seq), w: num(t.w), h: num(t.h) };
      }) };
      out.coThem = tu0 + n < ds.length;
      if (s.role === 'admin' && req.ht) { var h = chatHT(req.ht); if (h && h.thuMuc && env.files.folderUrl) out.thuMucUrl = env.files.folderUrl(h.thuMuc); }
      if (s.role === 'admin') out.chuaXep = T('TepChat').filter(function (t) { return !t.thuMuc; }).length;
      return out;
    };
    H['chat.xep'] = function (req, s) {   // Quản trị: chuyển tệp cũ (bản 2.0) vào thư mục theo cuộc trò chuyện
      if (s.role !== 'admin') err('Chỉ Quản trị thực hiện được.');
      var ds = T('TepChat').filter(function (t) { return !t.thuMuc; }), xong = 0, loi = 0;
      ds.slice(0, 150).forEach(function (t) {
        var h = chatHT(t.hoiThoai); if (!h) { t.thuMuc = '-'; return; }
        var fd = chatThuMuc(h); if (!fd) { loi++; return; }
        try { if (env.files.move) env.files.move(t.fileId, fd); t.thuMuc = fd; xong++; } catch (e) { t.thuMuc = '-'; loi++; }
      });
      D('TepChat'); log(s, 'chat.xep', xong + ' tệp');
      return { xong: xong, loi: loi, con: T('TepChat').filter(function (t) { return !t.thuMuc; }).length };
    };
    H['chat.file'] = function (req, s) {
      var t = T('TepChat').filter(function (x) { return x.fileId === String(req.fileId || ''); })[0];
      if (!t) err('Tệp không còn (có thể đã thu hồi).');
      chatCan(s, t.hoiThoai);
      var r = tepDoc(t, !!req.thumb); r._tf = {}; r._tf[t.fileId] = { h: t.hoiThoai, t: t.thumbId || '', m: t.mime, n: t.ten };
      return r;
    };
    H['chat.react'] = function (req, s) {
      var c = chatCan(s, req.ht), e = String(req.emoji || '');
      if (CAM_XUC.indexOf(e) < 0) err('Biểu tượng không hợp lệ.');
      return { msg: chatGhi(s, c.h.id, 'camxuc', e, { traLoi: String(req.msgId || '').slice(0, 40) }) };
    };
    H['chat.recall'] = function (req, s) {
      var c = chatCan(s, req.ht), g = env.chat.get(num(req.seq));
      if (!g || g.id !== req.msgId || g.hoiThoai !== c.h.id) err('Không tìm thấy tin nhắn.');
      if (g.username !== s.username) err('Chỉ thu hồi được tin nhắn của mình.');
      if (g.loai === 'kho') err('Phiếu kho không thu hồi được – dùng "Hủy phiếu" trên thẻ phiếu (giữ lịch sử).');   // v3.30
      var tep = parseJSON(g.tep, null), xoaTF = [];
      if (tep && tep.fileId) {
        try { env.files.del(tep.fileId); } catch (e) { }
        var a = T('TepChat'); for (var i = a.length - 1; i >= 0; i--) if (a[i].fileId === tep.fileId) { if (a[i].thumbId) { try { env.files.del(a[i].thumbId); } catch (e) { } } a.splice(i, 1); } D('TepChat');
        xoaTF.push(tep.fileId);
      }
      var m = chatGhi(s, c.h.id, 'thuhoi', '', { traLoi: g.id });
      var h = c.h, tc = parseJSON(h.tinCuoi, null);
      if (tc && tc.id === g.id) { tc.loai = 'thuhoi'; tc.noiDung = ''; h.tinCuoi = JSON.stringify(tc); D('HoiThoai'); }
      return { msg: m, _tfXoa: xoaTF };
    };
    H['chat.read'] = function (req, s) {
      var c = chatCan(s, req.ht), q = num(req.seq);
      if (q > num(c.tv.daDoc)) { c.tv.daDoc = Math.min(q, env.chat.last()); D('ThanhVien'); chatBao(chatMem(c.h.id).map(function (x) { return x.username; })); }
      return true;
    };
    H['chat.pref'] = function (req, s) {
      var c = chatCan(s, req.ht);
      if (req.ghim !== undefined) c.tv.ghim = req.ghim ? 'Có' : '';
      if (req.tatTB !== undefined) c.tv.tatTB = req.tatTB ? 'Có' : '';
      D('ThanhVien'); chatBao([s.username]);
      return true;
    };
    H['chat.open'] = function (req, s) {
      var u = String(req.username || '').toLowerCase();
      if (u === s.username) err('Không nhắn cho chính mình.');
      if (!chatUserOK(u)) err('Tài khoản không tồn tại hoặc đã bị khóa.');
      var mine = chatTVs(s.username).map(function (x) { return x.hoiThoai; });
      var h = T('HoiThoai').filter(function (x) { return x.loai === 'rieng' && mine.indexOf(x.id) >= 0 && chatMem(x.id).some(function (m) { return m.username === u; }); })[0];
      if (!h) {
        h = { id: newId(), loai: 'rieng', ten: '', mau: '', moTa: '', taoBoi: s.username, taoLuc: env.now(), capNhat: env.now(), tinCuoi: '', trangThai: '' };
        T('HoiThoai').push(h); D('HoiThoai');
        chatThem(h.id, s.username, 'thanhvien'); chatThem(h.id, u, 'thanhvien');
        chatBao([s.username, u]);
      }
      return chatHTDs(s).filter(function (x) { return x.id === h.id; })[0];
    };
    H['chat.group'] = function (req, s) {
      var ten = String(req.ten || '').trim().slice(0, 80);
      if (!ten) err('Nhập tên nhóm.');
      var ds = (req.thanhVien || []).map(function (u) { return String(u).toLowerCase(); }).filter(function (u, i, a) { return u !== s.username && a.indexOf(u) === i && chatUserOK(u); });
      if (!ds.length) err('Chọn ít nhất 1 thành viên.');
      var h = { id: newId(), loai: 'nhom', ten: ten, mau: CHAT_MAU[Math.floor(Math.random() * CHAT_MAU.length)], moTa: String(req.moTa || '').slice(0, 500), taoBoi: s.username, taoLuc: env.now(), capNhat: env.now(), tinCuoi: '', trangThai: '' };
      T('HoiThoai').push(h); D('HoiThoai');
      chatThem(h.id, s.username, 'truong'); ds.forEach(function (u) { chatThem(h.id, u, 'thanhvien'); });
      chatHeThong(s, h.id, (s.hoTen || s.username) + ' đã tạo nhóm "' + ten + '" với ' + ds.length + ' thành viên.');
      chatBao([s.username].concat(ds)); log(s, 'chat.group', ten + ' (' + (ds.length + 1) + ' người)');
      return chatHTDs(s).filter(function (x) { return x.id === h.id; })[0];
    };
    H['chat.gupdate'] = function (req, s) {
      var c = chatTruong(s, req.ht), h = c.h, doi = [];
      if (!laQL(s, c)) err('Chỉ trưởng, phó nhóm được đổi thông tin nhóm.');
      if (req.ten !== undefined) { var t = String(req.ten).trim().slice(0, 80); if (!t) err('Tên nhóm trống.'); if (t !== h.ten) { doi.push('đổi tên nhóm thành "' + t + '"'); h.ten = t; } }
      if (req.moTa !== undefined && String(req.moTa) !== h.moTa) { h.moTa = String(req.moTa).slice(0, 500); doi.push('cập nhật mô tả nhóm'); }
      if (req.mau !== undefined && CHAT_MAU.indexOf(req.mau) >= 0) h.mau = req.mau;
      var lcT = parseJSON(h.luuTru, null);
      if (doi.length && lcT && lcT.goc && env.luuTru) { try { env.luuTru.doiTen(lcT.goc, chatTenThuMuc(h)); } catch (e) { } }
      else if (h.thuMuc && env.files.renameFolder && doi.length) { try { env.files.renameFolder(h.thuMuc, chatTenThuMuc(h)); } catch (e) { } }
      D('HoiThoai');
      if (doi.length) chatHeThong(s, h.id, (s.hoTen || s.username) + ' đã ' + doi.join(', ') + '.');
      chatBao(chatMem(h.id).map(function (x) { return x.username; }));
      return true;
    };
    H['chat.gadd'] = function (req, s) {
      var c = chatTruong(s, req.ht), them = [];
      (req.usernames || []).forEach(function (u) { u = String(u).toLowerCase(); if (chatUserOK(u) && chatThem(c.h.id, u, 'thanhvien')) them.push(u); });
      if (!them.length) err('Không có thành viên mới để thêm.');
      chatHeThong(s, c.h.id, (s.hoTen || s.username) + ' đã thêm ' + them.map(chatTen).join(', ') + ' vào nhóm.');
      chatBao(chatMem(c.h.id).map(function (x) { return x.username; }));
      return true;
    };
    function chatRoi(htId, u) {
      var tv = chatTV(htId, u); if (!tv) return;
      tv.roi = env.now(); tv.ghim = ''; D('ThanhVien');
      var con = chatMem(htId), h = chatHT(htId);
      if (!con.length) { h.trangThai = 'Đã giải tán'; D('HoiThoai'); return; }
      if (!con.some(function (x) { return x.vaiTro === 'truong'; })) { con.sort(function (a, b) { return a.thamGia < b.thamGia ? -1 : 1; })[0].vaiTro = 'truong'; }
    }
    H['chat.gremove'] = function (req, s) {
      var c = chatTruong(s, req.ht), u = String(req.username || '').toLowerCase();
      if (!laQL(s, c)) err('Chỉ trưởng, phó nhóm được mời thành viên ra khỏi nhóm.');
      if (u === s.username) err('Dùng "Rời nhóm".');
      var tv = chatTV(c.h.id, u); if (!tv) err('Người này không còn trong nhóm.');
      if (tv.vaiTro === 'truong' && c.h.taoBoi !== s.username && s.role !== 'admin') err('Không mời được trưởng nhóm khác.');
      if (tv.vaiTro === 'pho' && !laTruong(s, c)) err('Phó nhóm không mời được phó nhóm khác.');
      var ai = chatMem(c.h.id).map(function (x) { return x.username; });
      chatRoi(c.h.id, u);
      chatHeThong(s, c.h.id, (s.hoTen || s.username) + ' đã mời ' + chatTen(u) + ' ra khỏi nhóm.');
      chatBao(ai);
      return true;
    };
    H['chat.gleave'] = function (req, s) {
      var c = chatCan(s, req.ht);
      if (c.h.loai !== 'nhom') err('Không rời được cuộc trò chuyện riêng.');
      var ai = chatMem(c.h.id).map(function (x) { return x.username; });
      if (!c.h.trangThai) chatHeThong(s, c.h.id, (s.hoTen || s.username) + ' đã rời nhóm.');
      chatRoi(c.h.id, s.username); chatBao(ai);
      return true;
    };
    H['chat.grole'] = function (req, s) {
      var c = chatTruong(s, req.ht), u = String(req.username || '').toLowerCase();
      if (!laTruong(s, c)) err('Chỉ trưởng nhóm được phân quyền.');
      var tv = chatTV(c.h.id, u); if (!tv) err('Người này không còn trong nhóm.');
      var moi = ['truong', 'pho'].indexOf(req.vaiTro) >= 0 ? req.vaiTro : 'thanhvien', cu = tv.vaiTro || 'thanhvien';
      if (moi !== 'truong' && cu === 'truong' && !chatMem(c.h.id).some(function (x) { return x.vaiTro === 'truong' && x.username !== u; })) err('Nhóm cần ít nhất 1 trưởng nhóm.');
      tv.vaiTro = moi; D('ThanhVien');
      var TEN_VT = { truong: 'trưởng nhóm', pho: 'phó nhóm' };
      chatHeThong(s, c.h.id, (s.hoTen || s.username) + (moi !== 'thanhvien' ? ' đã đặt ' + chatTen(u) + ' làm ' + TEN_VT[moi] + '.' : ' đã bỏ quyền ' + TEN_VT[cu] + ' của ' + chatTen(u) + '.'));
      chatBao(chatMem(c.h.id).map(function (x) { return x.username; }));
      return true;
    };
    H['chat.gdisband'] = function (req, s) {
      var c = chatTruong(s, req.ht);
      if (c.h.taoBoi !== s.username && s.role !== 'admin') err('Chỉ người tạo nhóm hoặc Quản trị được giải tán nhóm.');
      chatHeThong(s, c.h.id, (s.hoTen || s.username) + ' đã giải tán nhóm. Thành viên vẫn xem được lịch sử, không gửi được tin mới.');
      c.h.trangThai = 'Đã giải tán'; D('HoiThoai');
      chatBao(chatMem(c.h.id).map(function (x) { return x.username; })); log(s, 'chat.gdisband', c.h.ten);
      return true;
    };


    /* ---- MỤC NHÓM: bình chọn, nhắc hẹn, yêu cầu báo cáo định kỳ ----
       ChatMuc: một dòng mỗi mục; ChatTraLoi: phiếu bầu / báo cáo đã nộp (ky = ngày báo cáo). */
    function laQL(s, c) { return c.tv.vaiTro === 'truong' || c.tv.vaiTro === 'pho' || s.role === 'admin'; }
    function mucLay(id) { var m = T('ChatMuc').filter(function (x) { return x.id === id; })[0]; if (!m || m.trangThai === 'xoa') err('Mục không còn.'); return m; }
    function mucNguoi(m) { var n = parseJSON(m.cauHinh, {}).nguoi || []; return n.length ? n.filter(function (u) { return chatTV(m.hoiThoai, u); }) : chatMem(m.hoiThoai).map(function (x) { return x.username; }); }
    function mucPub(m) {
      var o = { id: m.id, hoiThoai: m.hoiThoai, loai: m.loai, tieuDe: m.tieuDe, noiDung: m.noiDung, cauHinh: parseJSON(m.cauHinh, {}), taoBoi: m.taoBoi, taoLuc: m.taoLuc, trangThai: m.trangThai, ketTiep: m.ketTiep, lanCuoi: m.lanCuoi };
      var tl = T('ChatTraLoi').filter(function (x) { return x.mucId === m.id; });
      if (m.loai === 'baocao') { var tu = addDays(env.today(), -31); tl = tl.filter(function (x) { return x.ky >= tu; }); o.nguoi = mucNguoi(m); }
      o.traLoi = tl.map(function (x) { return { ky: x.ky, username: x.username, giaTri: parseJSON(x.giaTri, null), thoiGian: x.thoiGian }; });
      return o;
    }
    function mucDs(s) {
      var my = {}; chatTVs(s.username).forEach(function (t) { my[t.hoiThoai] = 1; });
      var tu = addDays(env.today(), -120);
      return T('ChatMuc').filter(function (m) { return my[m.hoiThoai] && m.trangThai !== 'xoa' && (m.trangThai === 'mo' || String(m.capNhat || m.taoLuc).slice(0, 10) >= tu); }).map(mucPub);
    }
    function mucMoi(s, c, loai, tieuDe, nd, ch, ketTiep) {
      var m = { id: newId(), hoiThoai: c.h.id, loai: loai, tieuDe: tieuDe, noiDung: nd || '', cauHinh: JSON.stringify(ch || {}), taoBoi: s.username, taoLuc: env.now(), trangThai: 'mo', ketTiep: ketTiep || '', lanCuoi: '', capNhat: env.now(), nhacHan: '' };
      T('ChatMuc').push(m); D('ChatMuc'); return m;
    }
    function gioOK(g) { return /^([01]\d|2[0-3]):[0-5]\d$/.test(String(g || '')); }
    function chuanLuc(v) { var s = String(v || '').replace('T', ' ').slice(0, 16); if (!/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(s)) err('Thời gian nhắc không hợp lệ.'); return s; }
    function congLuc(luc, lap) {
      var d = luc.slice(0, 10), t = luc.slice(11, 16);
      if (lap === 'ngay') d = addDays(d, 1); else if (lap === 'tuan') d = addDays(d, 7); else if (lap === 'thang') d = addMonths(d, 1); else return '';
      return d + ' ' + t;
    }
    function bcKeTiep(ch, tuNgay) {           // lần nhắc báo cáo kế tiếp từ ngày tuNgay (gồm cả ngày đó) theo các thứ đã chọn
      var thu = (ch.thu && ch.thu.length) ? ch.thu : [1, 2, 3, 4, 5, 6];
      for (var i = 0, d = tuNgay; i < 15; i++, d = addDays(d, 1)) { var w = new Date(Date.UTC(+d.slice(0, 4), +d.slice(5, 7) - 1, +d.slice(8, 10))).getUTCDay(); if (thu.indexOf(w === 0 ? 7 : w) >= 0) return d + ' ' + ch.gio; }
      return '';
    }
    function mucGhiTin(m, loai, nd, extra) {
      var u = T('Users').filter(function (x) { return x.username === m.taoBoi; })[0];
      return chatGhi({ username: m.taoBoi, hoTen: u ? (u.hoTen || m.taoBoi) : m.taoBoi }, m.hoiThoai, loai, nd, Object.assign({ traLoi: m.id }, extra || {}));
    }
    /* chạy lịch: nhắc hẹn đến giờ, phát yêu cầu báo cáo, nhắc người chưa nộp khi tới hạn */
    function chayLich() {
      var cfL = getCfg();                           // v3.8: áp bảng phân quyền
      var bay = env.now().slice(0, 16), hom = bay.slice(0, 10), n = 0, bao = {};
      if (cfL.hdNhacNgay !== hom && bay.slice(11) >= '07:00') { try { setCfg({ hdNhacNgay: hom }); hdNhac(); } catch (e) { } }   // v3.33: nhắc HĐLĐ sắp hết hạn, mỗi ngày 1 lần
      T('ChatMuc').forEach(function (m) {
        if (m.trangThai !== 'mo' || (m.loai !== 'nhac' && m.loai !== 'baocao')) return;
        var h = chatHT(m.hoiThoai); if (!h || h.trangThai) return;
        var ch = parseJSON(m.cauHinh, {});
        if (m.loai === 'nhac' && m.ketTiep && m.ketTiep <= bay) {
          mucGhiTin(m, 'nhac', m.tieuDe, { tep: JSON.stringify({ luc: m.ketTiep }) }); n++;
          var k = m.ketTiep; m.lanCuoi = k;
          if (ch.lap && ch.lap !== 'mot') { var g = 0; while (k && k <= bay && g++ < 400) k = congLuc(k, ch.lap); m.ketTiep = k; }
          else { m.ketTiep = ''; m.trangThai = 'xong'; }
          m.capNhat = env.now(); bao[m.hoiThoai] = 1;
        }
        if (m.loai === 'baocao') {
          if (m.ketTiep && m.ketTiep <= bay) {
            var kk = m.ketTiep, nx, gg = 0;
            while ((nx = bcKeTiep(ch, addDays(kk.slice(0, 10), 1))) && nx <= bay && gg++ < 400) kk = nx;     // bỏ qua các ngày đã lỡ, chỉ phát ngày gần nhất
            var ky = kk.slice(0, 10);
            if (m.lanCuoi !== ky) { mucGhiTin(m, 'baocao', m.tieuDe, { tep: JSON.stringify({ ky: ky }) }); n++; m.lanCuoi = ky; if (ch.han && ky === hom && bay >= hom + ' ' + ch.han) m.nhacHan = hom; }
            m.ketTiep = bcKeTiep(ch, addDays(ky, 1)); m.capNhat = env.now(); bao[m.hoiThoai] = 1;
          }
          if (ch.han && m.lanCuoi === hom && m.nhacHan !== hom && bay >= hom + ' ' + ch.han) {
            var nop = {}; T('ChatTraLoi').forEach(function (x) { if (x.mucId === m.id && x.ky === hom) nop[x.username] = 1; });
            var con = mucNguoi(m).filter(function (u) { return !nop[u]; });
            if (con.length) { mucGhiTin(m, 'nhac', 'Còn ' + con.length + ' người chưa hoàn thành "' + m.tieuDe + '" ngày ' + hom.slice(8, 10) + '/' + hom.slice(5, 7) + ': ' + con.map(chatTen).join(', '), { tep: JSON.stringify({ ky: hom, han: 1, con: con }) }); n++; bao[m.hoiThoai] = 1; }
            m.nhacHan = hom;
          }
        }
      });
      if (n || Object.keys(bao).length) {
        D('ChatMuc');
        Object.keys(bao).forEach(function (id) { chatBao(chatMem(id).map(function (x) { return x.username; })); });
      }
      return n;
    }
    H['chat.binhchon'] = function (req, s) {
      var c = chatTruong(s, req.ht), td = String(req.tieuDe || '').trim().slice(0, 200);
      if (!td) err('Nhập câu hỏi bình chọn.');
      var lc = (req.luaChon || []).map(function (x) { return String(x || '').trim().slice(0, 120); }).filter(function (x, i, a) { return x && a.indexOf(x) === i; });
      if (lc.length < 2) err('Cần ít nhất 2 phương án.');
      if (lc.length > 20) err('Tối đa 20 phương án.');
      var han = req.han ? chuanLuc(req.han) : '';
      var m = mucMoi(s, c, 'binhchon', td, '', { luaChon: lc.map(function (t, i) { return { id: 'o' + (i + 1), ten: t, boi: s.username }; }), nhieu: !!req.nhieu, themLC: !!req.themLC, han: han, anDanh: false });
      chatGhi(s, c.h.id, 'binhchon', td, { traLoi: m.id });
      chatBao(chatMem(c.h.id).map(function (x) { return x.username; }));
      return mucPub(m);
    };
    H['chat.bau'] = function (req, s) {
      var m = mucLay(req.mucId), c = chatCan(s, m.hoiThoai), ch = parseJSON(m.cauHinh, {});
      if (m.loai !== 'binhchon') err('Không phải bình chọn.');
      if (m.trangThai !== 'mo' || (ch.han && env.now().slice(0, 16) >= ch.han)) err('Bình chọn đã kết thúc.');
      var ids = ch.luaChon.map(function (o) { return o.id; }), chon = (req.chon || []).filter(function (x, i, a) { return ids.indexOf(x) >= 0 && a.indexOf(x) === i; });
      if (!ch.nhieu && chon.length > 1) err('Bình chọn này chỉ chọn 1 phương án.');
      var a = T('ChatTraLoi');
      for (var i = a.length - 1; i >= 0; i--) if (a[i].mucId === m.id && a[i].username === s.username) a.splice(i, 1);
      if (chon.length) a.push({ id: newId(), mucId: m.id, ky: '', username: s.username, giaTri: JSON.stringify(chon), thoiGian: env.now() });
      D('ChatTraLoi'); m.capNhat = env.now(); D('ChatMuc');
      chatBao(chatMem(c.h.id).map(function (x) { return x.username; }));
      return mucPub(m);
    };
    H['chat.themlc'] = function (req, s) {
      var m = mucLay(req.mucId), c = chatCan(s, m.hoiThoai), ch = parseJSON(m.cauHinh, {}), t = String(req.ten || '').trim().slice(0, 120);
      if (m.loai !== 'binhchon' || m.trangThai !== 'mo') err('Bình chọn đã kết thúc.');
      if (!ch.themLC && m.taoBoi !== s.username && !laQL(s, c)) err('Bình chọn không cho thêm phương án.');
      if (!t) err('Nhập phương án.');
      if (ch.luaChon.some(function (o) { return o.ten.toLowerCase() === t.toLowerCase(); })) err('Phương án đã có.');
      if (ch.luaChon.length >= 20) err('Tối đa 20 phương án.');
      ch.luaChon.push({ id: 'o' + (ch.luaChon.length + 1) + newId().slice(0, 3), ten: t, boi: s.username });
      m.cauHinh = JSON.stringify(ch); m.capNhat = env.now(); D('ChatMuc');
      chatBao(chatMem(c.h.id).map(function (x) { return x.username; }));
      return mucPub(m);
    };
    H['chat.nhac'] = function (req, s) {
      var c = chatTruong(s, req.ht), td = String(req.tieuDe || '').trim().slice(0, 300);
      if (!td) err('Nhập nội dung nhắc hẹn.');
      var luc = chuanLuc(req.luc), lap = ['mot', 'ngay', 'tuan', 'thang'].indexOf(req.lap) >= 0 ? req.lap : 'mot';
      if (lap === 'mot' && luc <= env.now().slice(0, 16)) err('Thời gian nhắc phải sau hiện tại.');
      var k = luc, g = 0; while (lap !== 'mot' && k <= env.now().slice(0, 16) && g++ < 400) k = congLuc(k, lap);
      var m = mucMoi(s, c, 'nhac', td, '', { luc: luc, lap: lap }, k);
      chatGhi(s, c.h.id, 'nhactao', td, { traLoi: m.id });
      chatBao(chatMem(c.h.id).map(function (x) { return x.username; }));
      return mucPub(m);
    };
    var BC_KIEU = ['text', 'dai', 'so', 'chon', 'ngay', 'cokhong', 'bang'];
    /* v3.9: kiểu "bang" – bảng số lượng: các dòng {ten, dv} chọn sẵn từ gợi ý (nhân lực, máy, công việc); người báo cáo điền số.
       Giá trị lưu dạng chữ "Tên: số đơn vị" mỗi dòng (+ "Tổng: …" khi cùng đơn vị) để đọc được ở mọi nơi: xem, Excel, AI phân tích. */
    var BC_DV_DEM = ['người', 'cái', 'chiếc', 'xe', 'máy', 'bộ', 'tổ', 'đội', 'công'];   // chỉ cộng Tổng khi đếm người / máy, không cộng khối lượng khác loại
    function bcSoTxt(n) { var t = String(Math.round(n * 1000) / 1000); return t.replace('.', ','); }
    function bcBangSach(f, v) {
      var ds = Array.isArray(v) ? v : bcBangDoc(String(v || ''));
      var out = [], co = {};
      ds.slice(0, 80).forEach(function (r) {
        var ten = String((r && r.ten) || '').replace(/[:\n\r]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80); if (!ten) return;
        var raw = String(r.so == null ? '' : r.so).trim(); if (raw === '') return;
        var so = num(raw.replace(/\s/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.'));
        if (!isFinite(so) || so < 0) err('"' + f.ten + '" – ' + ten + ': số lượng không hợp lệ.');
        if (!so) return;
        var dv = String(r.dv || '').replace(/[:\n\r]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 15);
        var k = ten.toLowerCase(); if (co[k]) { co[k].so += so; return; }
        co[k] = { ten: ten, so: so, dv: dv }; out.push(co[k]);
      });
      if (!out.length) return '';
      var dvs = out.map(function (r) { return r.dv; }).filter(function (x, i, a) { return a.indexOf(x) === i; });
      var txt = out.map(function (r) { return r.ten + ': ' + bcSoTxt(r.so) + (r.dv ? ' ' + r.dv : ''); });
      if (out.length > 1 && dvs.length === 1 && BC_DV_DEM.indexOf(dvs[0].toLowerCase()) >= 0) txt.push('Tổng: ' + bcSoTxt(out.reduce(function (t, r) { return t + r.so; }, 0)) + (dvs[0] ? ' ' + dvs[0] : ''));
      return txt.join('\n');
    }
    function bcBangDoc(t) { return HRCORE_bcBangDoc(t); }
    H['chat.baocao'] = function (req, s) {
      var c = chatTruong(s, req.ht);
      if (!laQL(s, c)) err('Chỉ trưởng nhóm, phó nhóm tạo được yêu cầu báo cáo.');
      var td = String(req.tieuDe || '').trim().slice(0, 200); if (!td) err('Nhập tên báo cáo.');
      var mau = (req.mau || []).map(function (f, i) {
        var ten = String(f.ten || '').trim().slice(0, 120); if (!ten) return null;
        var k = BC_KIEU.indexOf(f.kieu) >= 0 ? f.kieu : 'text';
        var lc = k === 'chon' ? (f.luaChon || []).map(function (x) { return String(x).trim(); }).filter(Boolean).slice(0, 20) : [];
        if (k === 'chon' && lc.length < 2) err('Mục "' + ten + '" cần ít nhất 2 lựa chọn.');
        var o = { id: 'f' + (i + 1), ten: ten, kieu: k, luaChon: lc, batBuoc: !!f.batBuoc, goiY: String(f.goiY || '').slice(0, 200) };
        if (k === 'bang') {
          var co = {};
          o.dong = (f.dong || []).map(function (d) { return { ten: String((d && d.ten) || '').replace(/[:\n\r]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80), dv: String((d && d.dv) || '').replace(/[:\n\r]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 15) }; })
            .filter(function (d) { var kk = d.ten.toLowerCase(); if (!d.ten || co[kk]) return false; co[kk] = 1; return true; }).slice(0, 60);
          o.them = f.them !== false;
          if (!o.dong.length && !o.them) err('Mục "' + ten + '" cần ít nhất 1 dòng (hoặc cho người báo cáo tự thêm dòng).');
        }
        return o;
      }).filter(Boolean);
      if (!mau.length) err('Mẫu báo cáo cần ít nhất 1 mục.');
      if (mau.length > 30) err('Tối đa 30 mục.');
      if (!gioOK(req.gio)) err('Giờ nhắc không hợp lệ (HH:mm).');
      if (req.han && !gioOK(req.han)) err('Hạn nộp không hợp lệ (HH:mm).');
      var thu = (req.thu || []).map(Number).filter(function (x, i, a) { return x >= 1 && x <= 7 && a.indexOf(x) === i; });
      if (!thu.length) err('Chọn ít nhất 1 ngày trong tuần.');
      var nguoi = (req.nguoi || []).map(function (u) { return String(u).toLowerCase(); }).filter(function (u) { return chatTV(c.h.id, u); });
      var ch = { mau: mau, gio: req.gio, han: req.han || '', thu: thu, nguoi: nguoi };
      var bd = req.batDau && parseD(req.batDau) ? req.batDau : env.today();
      var k = bcKeTiep(ch, bd); if (k && k <= env.now().slice(0, 16)) k = req.ngayNay ? k : bcKeTiep(ch, addDays(bd, 1));
      var m = mucMoi(s, c, 'baocao', td, String(req.moTa || '').slice(0, 500), ch, k);
      chatGhi(s, c.h.id, 'bctao', td, { traLoi: m.id });
      if (req.ngayNay && k && k <= env.now().slice(0, 16)) chayLich();
      chatBao(chatMem(c.h.id).map(function (x) { return x.username; }));
      return mucPub(m);
    };
    H['chat.nopbc'] = function (req, s) {
      var m = mucLay(req.mucId), c = chatCan(s, m.hoiThoai), ch = parseJSON(m.cauHinh, {});
      if (m.loai !== 'baocao') err('Không phải yêu cầu báo cáo.');
      if (mucNguoi(m).indexOf(s.username) < 0) err('Bạn không thuộc danh sách phải báo cáo.');
      var ky = String(req.ky || ''); if (!parseD(ky) || ky > env.today()) err('Ngày báo cáo không hợp lệ.');
      var gv = req.giaTri || {}, out = {}, tom = [];
      ch.mau.forEach(function (f) {
        var v = gv[f.id]; v = v == null ? '' : v;
        if (f.kieu === 'so') { if (String(v).trim() === '') v = ''; else { v = num(String(v).replace(/\./g, '').replace(',', '.')); } }
        else if (f.kieu === 'chon' || f.kieu === 'cokhong') { v = String(v); if (v && f.kieu === 'chon' && f.luaChon.indexOf(v) < 0) err('"' + f.ten + '": lựa chọn không hợp lệ.'); if (v && f.kieu === 'cokhong' && ['Có', 'Không'].indexOf(v) < 0) err('"' + f.ten + '": chọn Có hoặc Không.'); }
        else if (f.kieu === 'ngay') { v = String(v); if (v && !parseD(v)) err('"' + f.ten + '": ngày không hợp lệ.'); }
        else if (f.kieu === 'bang') v = bcBangSach(f, v);
        else v = String(v).trim().slice(0, 3000);
        if (f.batBuoc && (v === '' || v == null)) err('Chưa điền "' + f.ten + '".');
        out[f.id] = v; if (v !== '' && tom.length < 3) tom.push(f.ten + ': ' + String(v).slice(0, 60));
      });
      var a = T('ChatTraLoi'), cu = a.filter(function (x) { return x.mucId === m.id && x.ky === ky && x.username === s.username; })[0];
      if (cu) { cu.giaTri = JSON.stringify(out); cu.thoiGian = env.now(); }
      else a.push({ id: newId(), mucId: m.id, ky: ky, username: s.username, giaTri: JSON.stringify(out), thoiGian: env.now() });
      D('ChatTraLoi'); m.capNhat = env.now(); D('ChatMuc');
      chatGhi(s, c.h.id, 'bcnop', (cu ? 'cập nhật ' : '') + tom.join(' · '), { traLoi: m.id, tep: JSON.stringify({ ky: ky }) });
      chatBao(chatMem(c.h.id).map(function (x) { return x.username; }));
      return mucPub(m);
    };
    /* ==================== v3.30: PHIẾU NHẬP / XUẤT KHO VẬT TƯ TRONG NHÓM CHAT ====================
       Thành viên nhóm (thủ kho, chỉ huy công trường…) lập phiếu nhập hoặc xuất kho kèm ảnh chụp / ảnh tải lên (hàng hóa, biên bản, hóa đơn).
       Ảnh tải qua chat.upload (lưu vào thư mục Drive của nhóm), phiếu gửi thành thẻ trong cuộc trò chuyện; báo cáo theo kỳ, xuất Excel. */
    var KHO_LOAI = { nhap: 'Phiếu nhập kho', xuat: 'Phiếu xuất kho' };
    function khoSach(r) {
      var dong = (Array.isArray(r.dong) ? r.dong : parseJSON(r.dong, [])).map(function (d) {
        var ten = String((d && d.ten) || '').replace(/\s+/g, ' ').trim().slice(0, 120); if (!ten) return null;
        var sl = num(String(d.sl == null ? '' : d.sl).replace(/\s/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.'));
        if (!isFinite(sl) || sl <= 0) err('Vật tư "' + ten + '": nhập số lượng lớn hơn 0.');
        var dg = num(String(d.donGia == null ? '' : d.donGia).replace(/\D/g, ''));
        return { ten: ten, dv: String(d.dv || '').replace(/\s+/g, ' ').trim().slice(0, 20), sl: Math.round(sl * 1000) / 1000, donGia: dg, thanhTien: dg ? Math.round(sl * dg) : 0, quyCach: String(d.quyCach || '').trim().slice(0, 120), ghiChu: String(d.ghiChu || '').trim().slice(0, 200) };
      }).filter(Boolean);
      if (!dong.length) err('Phiếu cần ít nhất 1 dòng vật tư.');
      if (dong.length > 60) err('Tối đa 60 dòng vật tư mỗi phiếu.');
      return dong;
    }
    function khoPub(p) { var o = Object.assign({}, p); o.dong = parseJSON(p.dong, []); o.anh = parseJSON(p.anh, []); return o; }
    H['chat.kho'] = function (req, s) {
      var c = chatCan(s, req.ht), r = req.row || {};
      if (c.h.trangThai) err('Nhóm đã giải tán.');
      if (c.h.loai !== 'nhom') err('Phiếu kho chỉ lập trong nhóm (nhóm công trình / kho).');
      if (!KHO_LOAI[r.loai]) err('Chọn phiếu nhập hoặc phiếu xuất.');
      var ngay = parseD(r.ngay) ? r.ngay : ''; if (!ngay) err('Nhập ngày nhập / xuất.'); if (ngay > env.today()) err('Ngày nhập / xuất không được sau hôm nay.');
      var dong = khoSach(r);
      var anh = (Array.isArray(req.anh) ? req.anh : []).slice(0, 12).map(function (a) {
        var t = T('TepChat').filter(function (x) { return x.fileId === String((a && a.fileId) || ''); })[0];
        if (!t || t.hoiThoai !== c.h.id || t.nguoiTai !== s.username || t.msgId) err('Ảnh đính kèm không hợp lệ hoặc đã dùng – tải lại ảnh.');
        if (!/^image\//.test(t.mime)) err('"' + t.ten + '" không phải ảnh.');
        return { t: t, o: { fileId: t.fileId, ten: t.ten, mime: t.mime, w: num(a.w), h: num(a.h), moTa: String(a.moTa || '').slice(0, 120) } };
      });
      if (req.batBuocAnh && !anh.length) err('Cần ít nhất 1 ảnh (hàng hóa / biên bản / hóa đơn).');
      var a = T('PhieuKho'), ky = ngay.slice(0, 7).replace('-', ''), tien = r.loai === 'nhap' ? 'PN' : 'PX', stt = a.filter(function (x) { return x.hoiThoai === c.h.id && x.loai === r.loai && String(x.soPhieu).indexOf(tien + '-' + ky + '-') === 0; }).length + 1;
      var p = { id: newId(), hoiThoai: c.h.id, soPhieu: tien + '-' + ky + '-' + ('00' + stt).slice(-3), loai: r.loai, ngay: ngay, kho: String(r.kho || c.h.ten || '').trim().slice(0, 150),
        doiTac: String(r.doiTac || '').trim().slice(0, 200), hangMuc: String(r.hangMuc || '').trim().slice(0, 200), nguoiGiaoNhan: String(r.nguoiGiaoNhan || '').trim().slice(0, 120), soCT: String(r.soCT || '').trim().slice(0, 80),
        dong: JSON.stringify(dong), anh: JSON.stringify(anh.map(function (x) { return x.o; })), ghiChu: String(r.ghiChu || '').trim().slice(0, 1000),
        nguoiLap: s.username, hoTenLap: s.hoTen || s.username, taoLuc: env.now(), trangThai: 'Đã lập', capNhat: env.now() };
      var tong = dong.reduce(function (t, d) { return t + d.thanhTien; }, 0);
      var tom = KHO_LOAI[p.loai] + ' ' + p.soPhieu + ' – ' + dong.slice(0, 3).map(function (d) { return d.ten + ' ' + String(d.sl).replace('.', ',') + (d.dv ? ' ' + d.dv : ''); }).join('; ') + (dong.length > 3 ? '; +' + (dong.length - 3) + ' dòng' : '') + (anh.length ? ' · ' + anh.length + ' ảnh' : '');
      var m = chatGhi(s, c.h.id, 'kho', tom, { tep: JSON.stringify({ pk: p.id, so: p.soPhieu, loai: p.loai, ngay: ngay, soDong: dong.length, tong: tong, dong: dong.slice(0, 4).map(function (d) { return [d.ten, d.sl, d.dv]; }), anh: anh.slice(0, 4).map(function (x) { return x.o; }), soAnh: anh.length }) });
      p.msgId = m.id; a.push(p); D('PhieuKho');
      anh.forEach(function (x) { x.t.msgId = m.id; x.t.seq = m.seq; x.t.w = x.o.w; x.t.h = x.o.h; }); if (anh.length) D('TepChat');
      c.tv.daDoc = m.seq; D('ThanhVien');
      chatBao(chatMem(c.h.id).map(function (x) { return x.username; }));
      log(s, 'chat.kho', p.soPhieu + ' (' + c.h.ten + '): ' + dong.length + ' dòng, ' + anh.length + ' ảnh');
      return { msg: m, phieu: khoPub(p) };
    };
    H['chat.khods'] = function (req, s) {        // danh sách phiếu của nhóm theo kỳ (thành viên nhóm xem được)
      var c = chatCan(s, req.ht), tu = parseD(req.tu) ? req.tu : '', den = parseD(req.den) ? req.den : '';
      var ds = T('PhieuKho').filter(function (p) { return p.hoiThoai === c.h.id && (!tu || p.ngay >= tu) && (!den || p.ngay <= den) && (!req.loai || p.loai === req.loai) && (req.coHuy || p.trangThai !== 'Đã hủy'); })
        .sort(function (a, b) { return a.ngay < b.ngay ? 1 : a.ngay > b.ngay ? -1 : (a.taoLuc < b.taoLuc ? 1 : -1); });
      return { ds: ds.map(khoPub), ten: c.h.ten };
    };
    H['chat.khoxem'] = function (req, s) {
      var p = T('PhieuKho').filter(function (x) { return x.id === req.id; })[0]; if (!p) err('Không tìm thấy phiếu.');
      chatCan(s, p.hoiThoai); return khoPub(p);
    };
    H['chat.khohuy'] = function (req, s) {
      var p = T('PhieuKho').filter(function (x) { return x.id === req.id; })[0]; if (!p) err('Không tìm thấy phiếu.');
      var c = chatCan(s, p.hoiThoai);
      if (p.nguoiLap !== s.username && !laQL(s, c)) err('Chỉ người lập phiếu, trưởng hoặc phó nhóm hủy được phiếu.');
      if (p.trangThai === 'Đã hủy') err('Phiếu đã hủy.');
      var ly = String(req.lyDo || '').trim().slice(0, 300); if (!ly) err('Ghi lý do hủy phiếu.');
      p.trangThai = 'Đã hủy'; p.lyDoHuy = ly; p.nguoiHuy = s.hoTen || s.username; p.capNhat = env.now(); D('PhieuKho');
      chatHeThong(s, c.h.id, (s.hoTen || s.username) + ' đã hủy ' + KHO_LOAI[p.loai].toLowerCase() + ' ' + p.soPhieu + ': ' + ly);
      chatBao(chatMem(c.h.id).map(function (x) { return x.username; }));
      log(s, 'chat.khohuy', p.soPhieu + ': ' + ly);
      return khoPub(p);
    };
    H['chat.mucdong'] = function (req, s) {
      var m = mucLay(req.mucId), c = chatCan(s, m.hoiThoai);
      if (m.taoBoi !== s.username && !laQL(s, c)) err('Chỉ người tạo, trưởng hoặc phó nhóm thực hiện được.');
      var xoa = !!req.xoa, ten = { binhchon: 'bình chọn', nhac: 'nhắc hẹn', baocao: 'yêu cầu báo cáo' }[m.loai] || 'mục';
      m.trangThai = xoa ? 'xoa' : (m.trangThai === 'mo' ? 'dong' : 'mo');
      if (m.trangThai === 'mo' && m.loai === 'baocao') m.ketTiep = bcKeTiep(parseJSON(m.cauHinh, {}), addDays(env.today(), m.lanCuoi === env.today() ? 1 : 0));
      if (m.trangThai === 'mo' && m.loai === 'nhac' && !m.ketTiep) err('Nhắc hẹn đã qua, hãy tạo nhắc hẹn mới.');
      m.capNhat = env.now(); D('ChatMuc');
      chatHeThong(s, c.h.id, (s.hoTen || s.username) + (xoa ? ' đã xóa ' : m.trangThai === 'dong' ? (m.loai === 'binhchon' ? ' đã khóa ' : ' đã tạm dừng ') : ' đã mở lại ') + ten + ' "' + m.tieuDe + '".');
      chatBao(chatMem(c.h.id).map(function (x) { return x.username; }));
      return xoa ? { id: m.id, trangThai: 'xoa' } : mucPub(m);
    };
    H['chat.lich'] = function (req, s) { if (s.role !== 'admin') err('Chỉ Quản trị.'); return { daChay: chayLich() }; };

    /* ==================== LƯU TRỮ NHÓM & PHÂN TÍCH AI (v2.8) ====================
       Mỗi nhóm có một thư mục trong "TAPACO Nhân sự – Dữ liệu app/Lưu trữ nhóm chat (KHÔNG CHIA SẺ)" (thư mục mẹ đặt cạnh file Google Sheet):
         0. Thông tin nhóm.md · 1. Tệp & ảnh · 2. Nhật ký từng ngày (YYYY-MM-DD.md) · 3. Báo cáo định kỳ (*.csv) · 4. Phân tích AI (Google Docs + JSON)
       Phân tích AI chạy nhiều bước ngắn (mỗi lần gọi AI dưới 60 giây): đọc từng đợt ≤ 7 ngày → tổng hợp.
       Số đếm, tỷ lệ nộp báo cáo, tổng số liệu báo cáo định kỳ do máy tính; AI chỉ tóm tắt, trích xuất (kèm trích dẫn) và nhận xét. */
    var LT_SUB = { tep: '1. Tệp & ảnh', nk: '2. Nhật ký từng ngày', bc: '3. Báo cáo định kỳ', ai: '4. Phân tích AI' };
    var LT_ND = { text: 1, anh: 1, tep: 1 };
    function ltQuyen(s, htId) {
      var h = chatHT(htId); if (!h) err('Không tìm thấy cuộc trò chuyện.');
      if (!can(s.role, 'aiNhom')) err('Bạn chưa được phân quyền Phân tích nhóm (AI).');
      if (h.loai !== 'nhom' && s.role !== 'admin' && s.role !== 'giamdoc') err('Chỉ Giám đốc, Quản trị phân tích được trò chuyện riêng.');
      return h;
    }
    function ltCay(h, taoMoi) {               // cây thư mục lưu trữ của nhóm (tạo khi cần)
      var c = parseJSON(h.luuTru, null);
      if (c && c.goc) return c;
      if (!env.luuTru || !taoMoi) return null;
      var L = env.luuTru, goc = L.thuMuc(h.loai === 'nhom' ? '' : L.thuMuc('', '_Trò chuyện riêng'), chatTenThuMuc(h)), tep = '';
      if (h.thuMuc) { try { L.chuyen(h.thuMuc, goc); L.doiTen(h.thuMuc, LT_SUB.tep); tep = h.thuMuc; } catch (e) { tep = ''; } }
      if (!tep) tep = L.thuMuc(goc, LT_SUB.tep);
      c = { goc: goc, tep: tep, nk: L.thuMuc(goc, LT_SUB.nk), bc: L.thuMuc(goc, LT_SUB.bc), ai: L.thuMuc(goc, LT_SUB.ai) };
      h.thuMuc = tep; h.luuTru = JSON.stringify(c); D('HoiThoai');
      return c;
    }
    function ltMs(t) { return Date.parse(String(t).replace(' ', 'T') + '+07:00') || 0; }
    /* toàn bộ tin của nhóm: bỏ nội dung tin đã thu hồi, gộp cảm xúc; su = mốc thời gian mọi thay đổi và ngày bị ảnh hưởng */
    function ltTin(htId, all) {
      all = all || env.chat.before(env.chat.last() + 1, env.chat.last());
      var ds = [], idx = {}, thuHoi = {}, ngayCua = {}, su = [], cx = {};
      all.forEach(function (m) { if (m.hoiThoai !== htId) return; if (m.loai === 'thuhoi') thuHoi[m.traLoi] = 1; else ngayCua[m.id] = String(m.thoiGian).slice(0, 10); });
      all.forEach(function (m) {
        if (m.hoiThoai !== htId) return;
        if (m.loai === 'thuhoi' || m.loai === 'camxuc') {
          su.push({ tg: m.thoiGian, ngay: ngayCua[m.traLoi] || '' });
          if (m.loai === 'camxuc') { var k = m.traLoi + '|' + m.username + '|' + m.noiDung; cx[k] = !cx[k]; }
          return;
        }
        su.push({ tg: m.thoiGian, ngay: String(m.thoiGian).slice(0, 10) });
        var o = { seq: m.seq, id: m.id, username: m.username, hoTen: m.hoTen || chatTen(m.username), loai: m.loai, noiDung: thuHoi[m.id] ? '' : String(m.noiDung || ''),
          tep: thuHoi[m.id] ? null : parseJSON(m.tep, null), traLoi: m.traLoi, thoiGian: String(m.thoiGian), thuHoi: !!thuHoi[m.id] };
        ds.push(o); idx[m.id] = o;
      });
      Object.keys(cx).forEach(function (k) {
        if (!cx[k]) return; var p = k.split('|'), g = idx[p[0]]; if (!g || g.thuHoi) return;
        g.camXuc = g.camXuc || {}; g.camXuc[p[2]] = (g.camXuc[p[2]] || 0) + 1;
      });
      return { ds: ds, su: su };
    }
    function ltCtx(h, tins) {
      var muc = {}, bc = [], phat = {}, bcIds = {};
      T('ChatMuc').forEach(function (m) { if (m.hoiThoai !== h.id) return; muc[m.id] = m; if (m.loai === 'baocao' && m.trangThai !== 'xoa') { bc.push(m); bcIds[m.id] = 1; } });
      tins.forEach(function (m) { if (m.loai === 'baocao' && m.tep && m.tep.ky) (phat[m.traLoi] = phat[m.traLoi] || {})[m.tep.ky] = 1; });
      return { muc: muc, bc: bc, phat: phat, tl: T('ChatTraLoi').filter(function (x) { return bcIds[x.mucId]; }) };
    }
    function ltTheoNgay(ds) { var o = {}; ds.forEach(function (m) { var d = m.thoiGian.slice(0, 10); (o[d] = o[d] || []).push(m); }); return o; }
    function ltNgays(ds, ctx) {
      var o = {}; ds.forEach(function (m) { o[m.thoiGian.slice(0, 10)] = 1; }); ctx.tl.forEach(function (x) { if (x.ky) o[x.ky] = 1; });
      return Object.keys(o).filter(function (d) { return parseD(d); }).sort();
    }
    function ltDungHan(ch, ky, tg) { var d = String(tg).slice(0, 10); return d < ky || (d === ky && (!ch.han || String(tg).slice(11, 16) <= ch.han)); }
    function ltGiaTri(f, v) { if (v === '' || v == null) return ''; return f.kieu === 'ngay' ? ngayVN(String(v)) : String(v).replace(/\s*\n+\s*/g, ' ⏎ '); }
    function ltBauTxt(m) {
      var ch = parseJSON(m.cauHinh, {}), dem = {};
      T('ChatTraLoi').forEach(function (x) { if (x.mucId === m.id) parseJSON(x.giaTri, []).forEach(function (o) { dem[o] = (dem[o] || 0) + 1; }); });
      return (ch.luaChon || []).map(function (o) { return o.ten + ': ' + (dem[o.id] || 0); }).join('; ');
    }
    function ltDong(m, ctx) {
      var ai = m.hoTen, g = m.thoiGian.slice(11, 16), mc = ctx.muc[m.traLoi], ch = mc ? parseJSON(mc.cauHinh, {}) : {};
      if (m.thuHoi) return '- ' + g + ' · ' + ai + ': [tin nhắn đã thu hồi]';
      var cx = m.camXuc ? '  (' + Object.keys(m.camXuc).map(function (e) { return e + ' ' + m.camXuc[e]; }).join(' ') + ')' : '';
      var nd = String(m.noiDung || '').replace(/\s*\n+\s*/g, ' ⏎ '), tep = m.tep || {};
      switch (m.loai) {
        case 'text': return '- ' + g + ' · ' + ai + ': ' + nd + cx;
        case 'anh': return '- ' + g + ' · ' + ai + ': [Ảnh] ' + (tep.ten || '') + (nd ? ' – ' + nd : '') + ' (mã tệp: ' + (tep.fileId || '') + ')' + cx;
        case 'tep': return '- ' + g + ' · ' + ai + ': [Tệp] ' + (tep.ten || '') + ' (' + Math.max(1, Math.round(num(tep.kichThuoc) / 1024)) + ' KB, mã tệp: ' + (tep.fileId || '') + ')' + (nd ? ' – ' + nd : '') + cx;
        case 'hethong': return '- ' + g + ' · [Hệ thống] ' + nd;
        case 'binhchon': return '- ' + g + ' · ' + ai + ' tạo bình chọn "' + nd + '"' + (mc ? ' → kết quả: ' + ltBauTxt(mc) : '');
        case 'nhactao': return '- ' + g + ' · ' + ai + ' đặt nhắc hẹn "' + nd + '"' + (ch.luc ? ' lúc ' + ch.luc.slice(11, 16) + ' ' + ngayVN(ch.luc.slice(0, 10)) : '') + (ch.lap && ch.lap !== 'mot' ? ' (lặp ' + ({ ngay: 'hằng ngày', tuan: 'hằng tuần', thang: 'hằng tháng' }[ch.lap] || '') + ')' : '');
        case 'nhac': return '- ' + g + ' · [Nhắc hẹn] ' + nd;
        case 'bctao': return '- ' + g + ' · ' + ai + ' tạo yêu cầu báo cáo "' + nd + '"' + (ch.mau ? ' (mẫu: ' + ch.mau.map(function (f) { return f.ten; }).join(', ') + '; nhắc ' + ch.gio + (ch.han ? ', hạn ' + ch.han : '') + ')' : '');
        case 'baocao': return '- ' + g + ' · [Yêu cầu báo cáo] "' + nd + '"';
        case 'bcnop': return '- ' + g + ' · ' + ai + ' đã hoàn thành báo cáo "' + (mc ? mc.tieuDe : '') + '"' + (tep.ky && tep.ky !== m.thoiGian.slice(0, 10) ? ' cho ngày ' + ngayVN(tep.ky) : '');
        default: return '- ' + g + ' · ' + ai + ': ' + nd;
      }
    }

    /* ===== Bóc số liệu xác định từ báo cáo định kỳ (v3.5.2) =====
       Trường "Khối lượng thực hiện trong ngày" là văn bản tự do ("Đổ bê tông trụ T3: 42 m³;
       lắp cốt thép dầm D3 60%"). Máy bóc sẵn thành hạng mục + khối lượng + đơn vị để:
         (1) ghi vào nhật ký gửi AI dưới dạng đã chuẩn hóa,
         (2) làm lưới an toàn: ngày nào AI bỏ sót thì điền lại từ số liệu máy đã bóc.
       Nguyên tắc: CHỈ nhận một con số khi nó đi kèm ĐƠN VỊ hợp lệ. Không có đơn vị thì giữ nguyên
       câu chữ, không đoán — nếu không "đài móng T4" thành khối lượng 4, "dầm D3 60%" thành 3. */
    var BC_DVT = ['m3', 'm³', 'm2', 'm²', 'md', 'm1', 'm', 'km', 'ha', 'kg', 'tấn', 'tan', 'tạ', 'yến', 'khối',
      'bộ', 'bó', 'cái', 'chiếc', 'viên', 'thanh', 'cây', 'tấm', 'bao', 'cuộn', 'kiện', 'hộp', 'thùng', 'can', 'lô',
      'xe', 'chuyến', 'lít', 'l', 'công', 'người', 'ngày', 'ca', 'mẻ', 'cấu kiện', 'cọc', 'phiến', 'dầm', 'đốt',
      'tổ', 'điểm', 'vị trí', 'lần', 'trụ', 'mố', 'nhịp', 'lớp'];
    var BC_DV_RE = BC_DVT.map(function (x) { return x.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }).join('|');
    var BC_RE_CUOI = new RegExp('^(.*?)[:：]?\\s*([0-9][0-9.,]*)\\s*(%|' + BC_DV_RE + ')\\s*$', 'i');
    var BC_RE_GIUA = new RegExp('([0-9][0-9.,]*)\\s*(%|' + BC_DV_RE + ')(?=\\s|$|[,.;])', 'i');
    function bcSoVN(s) {
      var t = String(s).trim();
      if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) return parseFloat(t.replace(/\./g, '').replace(',', '.'));
      if (/^\d+,\d+$/.test(t)) return parseFloat(t.replace(',', '.'));
      var n = parseFloat(t.replace(/,/g, ''));
      return isFinite(n) ? n : null;
    }
    function bcChuSo(s) { return String(s).replace(/[^0-9a-zà-ỹ]/gi, '').length; }
    /* một ô văn bản -> danh sách { hangMuc, khoiLuong, dvt, trich } */
    function bcBocHM(txt, max) {
      var out = [];
      String(txt == null ? '' : txt).split(/[;\n\r]+/).forEach(function (p) {
        p = p.replace(/^\s*[-–•]\s*/, '').replace(/\s+/g, ' ').trim();
        if (!p) return;
        var m = BC_RE_CUOI.exec(p);
        if (m && bcChuSo(m[1]) >= 3) {
          var kl = bcSoVN(m[2]);
          if (kl != null) { out.push({ hangMuc: m[1].replace(/[:：\s]+$/, '').trim().slice(0, 160), khoiLuong: kl, dvt: String(m[3]).toLowerCase().slice(0, 20), trich: p.slice(0, 160) }); return; }
        }
        var g = BC_RE_GIUA.exec(p);
        if (g) {
          var kl2 = bcSoVN(g[1]), ten = (p.slice(0, g.index) + ' ' + p.slice(g.index + g[0].length)).replace(/\s+/g, ' ').replace(/^[:：\s-]+|[:：\s-]+$/g, '').trim();
          if (kl2 != null && bcChuSo(ten) >= 3) { out.push({ hangMuc: ten.slice(0, 160), khoiLuong: kl2, dvt: String(g[2]).toLowerCase().slice(0, 20), trich: p.slice(0, 160) }); return; }
        }
        out.push({ hangMuc: p.slice(0, 160), khoiLuong: null, dvt: '', trich: p.slice(0, 160) });
      });
      return out.slice(0, max || 15);
    }
    var BC_TR = {
      nhanLuc: /nhân\s*lực|số\s*người|quân\s*số|lao\s*động\s*có\s*mặt/i,
      banChiHuy: /ban\s*chỉ\s*huy|bch/i,
      thiCong: /khối\s*lượng|sản\s*lượng|việc\s*đã|việc\s*đang|công\s*việc|hạng\s*mục|thi\s*công|tiến\s*độ/i,
      thietBi: /thiết\s*bị|máy\s*móc|máy\s*,|xe\s*máy|máy\s*thi\s*công/i,
      vatTuNhap: /vật\s*tư\s*nhập|nhập\s*về|nhập\s*kho/i,
      vatTuXuat: /vật\s*tư\s*(sử\s*dụng|xuất)|xuất\s*kho|sử\s*dụng\s*trong\s*ngày/i,
      yeuCau: /đề\s*nghị\s*(cấp|nhập)|yêu\s*cầu\s*vật\s*tư|đề\s*nghị\s*vật\s*tư/i,
      keHoach: /kế\s*hoạch/i
    };
    /* một lần nộp báo cáo -> số liệu máy đọc được */
    function bcSoLieu(mau, gv) {
      var o = { nhanLuc: null, banChiHuy: null, thiCong: [], thietBi: [], vatTuNhap: [], vatTuXuat: [], yeuCau: [], keHoach: [] };
      (mau || []).forEach(function (f) {
        var v = gv[f.id]; if (v === '' || v == null) return;
        var ten = String(f.ten || ''), txt = String(v).trim();
        if (/^(không|không có|ko|n\/a|-)$/i.test(txt)) return;
        if (f.kieu === 'bang') {                         // v3.9: bảng số lượng – đã có sẵn tên, số, đơn vị từng dòng
          var rows = HRCORE_bcBangDoc(txt).map(function (r) { return { hangMuc: r.ten, khoiLuong: bcSoVN(r.so), dvt: r.dv, trich: r.ten + ': ' + r.so + (r.dv ? ' ' + r.dv : '') }; }).filter(function (r) { return r.khoiLuong != null; });
          var tong = rows.reduce(function (t, r) { return t + r.khoiLuong; }, 0);
          if (BC_TR.banChiHuy.test(ten)) { o.banChiHuy = tong; return; }
          if (BC_TR.nhanLuc.test(ten)) { o.nhanLuc = (o.nhanLuc || 0) + tong; o.nhanLucCT = (o.nhanLucCT || []).concat(rows); return; }
          if (BC_TR.vatTuNhap.test(ten)) { o.vatTuNhap = o.vatTuNhap.concat(rows); return; }
          if (BC_TR.vatTuXuat.test(ten)) { o.vatTuXuat = o.vatTuXuat.concat(rows); return; }
          if (BC_TR.yeuCau.test(ten)) { o.yeuCau = o.yeuCau.concat(rows); return; }
          if (BC_TR.thietBi.test(ten)) { o.thietBi = o.thietBi.concat(rows); return; }
          o.thiCong = o.thiCong.concat(rows); return;
        }
        if (BC_TR.banChiHuy.test(ten) && f.kieu === 'so') { o.banChiHuy = bcSoVN(txt); return; }
        if (BC_TR.nhanLuc.test(ten) && f.kieu === 'so') { o.nhanLuc = bcSoVN(txt); return; }
        if (BC_TR.vatTuNhap.test(ten)) { o.vatTuNhap = o.vatTuNhap.concat(bcBocHM(txt, 15)); return; }
        if (BC_TR.vatTuXuat.test(ten)) { o.vatTuXuat = o.vatTuXuat.concat(bcBocHM(txt, 15)); return; }
        if (BC_TR.yeuCau.test(ten)) { o.yeuCau = o.yeuCau.concat(bcBocHM(txt, 15)); return; }
        if (BC_TR.thietBi.test(ten)) { o.thietBi = o.thietBi.concat(bcBocHM(txt, 15)); return; }
        if (BC_TR.keHoach.test(ten)) { o.keHoach.push(txt.slice(0, 250)); return; }
        if (BC_TR.thiCong.test(ten)) { o.thiCong = o.thiCong.concat(bcBocHM(txt, 15)); return; }
      });
      return o;
    }
    function bcCoGi(o) { return o.nhanLuc != null || o.banChiHuy != null || o.thiCong.length || o.thietBi.length || o.vatTuNhap.length || o.vatTuXuat.length || o.yeuCau.length; }
    function bcSL(x) { return x.khoiLuong == null ? x.hangMuc : x.hangMuc + ': ' + x.khoiLuong + (x.dvt ? ' ' + x.dvt : ''); }
    /* số liệu máy đọc của cả một ngày, gom theo người nộp – dùng cho nhật ký và cho lưới an toàn */
    function bcNgaySo(ctx, ngay) {
      var ds = [];
      ctx.bc.forEach(function (m) {
        var ch = parseJSON(m.cauHinh, {});
        ctx.tl.filter(function (x) { return x.mucId === m.id && x.ky === ngay; }).forEach(function (x) {
          var o = bcSoLieu(ch.mau || [], parseJSON(x.giaTri, {}));
          if (bcCoGi(o)) ds.push({ username: x.username, bc: m.tieuDe, so: o });
        });
      });
      return ds;
    }
    function ltBCNgay(ctx, ngay) {
      var out = [];
      ctx.bc.forEach(function (m) {
        var ch = parseJSON(m.cauHinh, {}), tl = ctx.tl.filter(function (x) { return x.mucId === m.id && x.ky === ngay; }), phat = !!(ctx.phat[m.id] || {})[ngay];
        if (!tl.length && !phat) return;
        var phai = mucNguoi(m);
        out.push('', '### ' + m.tieuDe + ' — đã nộp ' + tl.length + '/' + phai.length);
        tl.sort(function (a, b) { return a.thoiGian < b.thoiGian ? -1 : 1; }).forEach(function (x) {
          var gv = parseJSON(x.giaTri, {}), tg = String(x.thoiGian);
          out.push('**' + chatTen(x.username) + '** (' + x.username + ', nộp lúc ' + tg.slice(11, 16) + (tg.slice(0, 10) !== ngay ? ' ngày ' + ngayVN(tg.slice(0, 10)) : '') + (ltDungHan(ch, ngay, tg) ? '' : ', trễ hạn') + ')');
          (ch.mau || []).forEach(function (f) { var v = ltGiaTri(f, gv[f.id]); if (v !== '') out.push('- ' + f.ten + ': ' + v); });
          var so = bcSoLieu(ch.mau || [], gv);          // máy bóc sẵn số liệu để AI khỏi phải tự tách
          if (bcCoGi(so)) {
            var L2 = [];
            if (so.nhanLuc != null) L2.push('nhân lực ' + so.nhanLuc + ' người' + (so.nhanLucCT && so.nhanLucCT.length ? ' (' + so.nhanLucCT.map(function (r) { return r.hangMuc + ' ' + r.khoiLuong; }).join(', ') + ')' : ''));
            if (so.banChiHuy != null) L2.push('ban chỉ huy ' + so.banChiHuy + ' người');
            if (so.thiCong.length) L2.push('khối lượng → ' + so.thiCong.map(bcSL).join(' | '));
            if (so.thietBi.length) L2.push('thiết bị → ' + so.thietBi.map(bcSL).join(' | '));
            if (so.vatTuNhap.length) L2.push('vật tư nhập → ' + so.vatTuNhap.map(bcSL).join(' | '));
            if (so.vatTuXuat.length) L2.push('vật tư xuất → ' + so.vatTuXuat.map(bcSL).join(' | '));
            if (so.yeuCau.length) L2.push('đề nghị cấp → ' + so.yeuCau.map(bcSL).join(' | '));
            if (L2.length) out.push('  [số liệu máy đọc được, dùng làm chuẩn] ' + L2.join('; '));
          }
        });
        var chua = phai.filter(function (u) { return !tl.some(function (x) { return x.username === u; }); });
        if (chua.length && phat) out.push('- Chưa nộp: ' + chua.map(chatTen).join(', '));
      });
      return out;
    }
    function ltNhatKy(h, ngay, tinNgay, ctx) {
      var nd = tinNgay.filter(function (m) { return LT_ND[m.loai] && !m.thuHoi; }), ng = {};
      nd.forEach(function (m) { ng[m.hoTen] = 1; });
      var L = ['# Nhật ký ' + (h.loai === 'nhom' ? 'nhóm "' + h.ten + '"' : 'trò chuyện ' + aiTenHT(h)) + ' – ' + THU_VN[thu(ngay)] + ', ' + ngayVN(ngay), '',
        '- Tin nhắn: ' + nd.filter(function (m) { return m.loai === 'text'; }).length + ' · Ảnh: ' + nd.filter(function (m) { return m.loai === 'anh'; }).length + ' · Tệp: ' + nd.filter(function (m) { return m.loai === 'tep'; }).length +
        ' · Người tham gia: ' + (Object.keys(ng).join(', ') || '—')];
      /* Báo cáo định kỳ đặt TRƯỚC phần trò chuyện: đây là số liệu do chính thành viên nộp
         (nhân lực, khối lượng, công việc), chính xác hơn tin nhắn. Để sau thì khi nhật ký quá dài
         phần này bị cắt mất đầu tiên — xem ltRutGon. */
      var bc = ltBCNgay(ctx, ngay);
      if (bc.length) { L.push('', '## Báo cáo định kỳ (số liệu do thành viên nộp – nguồn ưu tiên)'); L = L.concat(bc); }
      if (tinNgay.length) { L.push('', '## Diễn biến trong ngày'); tinNgay.forEach(function (m) { L.push(ltDong(m, ctx)); }); }
      return L.join('\n') + '\n';
    }
    /* Nhật ký một ngày quá dài: GIỮ NGUYÊN phần đầu (thống kê + báo cáo định kỳ) và chỉ lược bớt
       phần trò chuyện. Bản cũ cắt thẳng từ cuối nên số liệu báo cáo mất trước, tin nhắn lại còn. */
    function ltRutGon(t, max) {
      if (t.length <= max) return t;
      var k = t.indexOf('\n## Diễn biến trong ngày');
      if (k < 0 || k > max) return t.slice(0, max) + '\n…(đã rút gọn do quá dài)';
      var dau = t.slice(0, k), con = max - dau.length - 90;
      if (con < 500) return dau + '\n\n…(đã lược toàn bộ phần trò chuyện do nhật ký quá dài; giữ nguyên báo cáo định kỳ)\n';
      var cat = t.slice(k, k + con), nl = cat.lastIndexOf('\n');
      return dau + (nl > 200 ? cat.slice(0, nl) : cat) + '\n…(đã lược bớt phần trò chuyện do quá dài; giữ nguyên báo cáo định kỳ)\n';
    }
    function ltTenTep(s) { return String(s || '').replace(/[\\/:*?"<>|\r\n]/g, '_').trim().slice(0, 80) || 'Bao cao'; }
    function ltCSV(m, ctx) {
      var ch = parseJSON(m.cauHinh, {}), mau = ch.mau || [];
      var q = function (v) { v = String(v == null ? '' : v); return /[",\r\n;]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
      var rows = [['Ngày báo cáo', 'Họ tên', 'Tài khoản', 'Nộp lúc', 'Đúng hạn'].concat(mau.map(function (f) { return f.ten; }))];
      ctx.tl.filter(function (x) { return x.mucId === m.id; }).sort(function (a, b) { return a.ky < b.ky ? -1 : a.ky > b.ky ? 1 : (a.thoiGian < b.thoiGian ? -1 : 1); }).forEach(function (x) {
        var gv = parseJSON(x.giaTri, {});
        rows.push([x.ky, chatTen(x.username), x.username, x.thoiGian, ltDungHan(ch, x.ky, x.thoiGian) ? 'Có' : 'Trễ'].concat(mau.map(function (f) { return gv[f.id] == null ? '' : gv[f.id]; })));
      });
      return '\uFEFF' + rows.map(function (r) { return r.map(q).join(','); }).join('\r\n') + '\r\n';
    }
    function ltThongTin(h, ctx) {
      var VT = { truong: 'Trưởng nhóm', pho: 'Phó nhóm', thanhvien: 'Thành viên' }, L = ['# ' + (h.loai === 'nhom' ? h.ten : 'Trò chuyện riêng: ' + aiTenHT(h)), ''];
      if (h.moTa) L.push(h.moTa, '');
      L.push('- Tạo bởi: ' + chatTen(h.taoBoi) + ' lúc ' + h.taoLuc, '- Trạng thái: ' + (h.trangThai || 'Đang hoạt động'), '- Cập nhật lưu trữ: ' + env.now(), '', '## Thành viên');
      T('ThanhVien').filter(function (x) { return x.hoiThoai === h.id; }).forEach(function (x) {
        L.push('- ' + chatTen(x.username) + ' (' + x.username + ') – ' + (VT[x.vaiTro] || 'Thành viên') + ' – tham gia ' + String(x.thamGia).slice(0, 10) + (x.roi ? ' – đã rời ' + String(x.roi).slice(0, 10) : ''));
      });
      var ds = Object.keys(ctx.muc).map(function (k) { return ctx.muc[k]; }).filter(function (m) { return m.trangThai !== 'xoa'; });
      if (ds.length) {
        L.push('', '## Bình chọn, nhắc hẹn, yêu cầu báo cáo');
        ds.forEach(function (m) {
          var ch = parseJSON(m.cauHinh, {}), tt = { mo: 'đang mở', dong: 'tạm dừng/đã khóa', xong: 'đã xong' }[m.trangThai] || m.trangThai;
          L.push('- [' + ({ binhchon: 'Bình chọn', nhac: 'Nhắc hẹn', baocao: 'Báo cáo' }[m.loai] || m.loai) + '] ' + m.tieuDe + ' – ' + tt + ' – tạo bởi ' + chatTen(m.taoBoi) + ' ' + String(m.taoLuc).slice(0, 10) +
            (m.loai === 'baocao' ? ' – mẫu: ' + (ch.mau || []).map(function (f) { return f.ten; }).join(', ') : m.loai === 'binhchon' ? ' – ' + ltBauTxt(m) : ''));
        });
      }
      return L.join('\n') + '\n';
    }
    /* ghi nhật ký các ngày có thay đổi kể từ lần lưu trước (tối đa gioiHan ngày mỗi lần), CSV báo cáo, thông tin nhóm */
    function ltChay(h, gioiHan, all) {
      var c = ltCay(h, true); if (!c) err('Máy chủ chưa bật lưu trữ Google Drive.');
      var L = env.luuTru, r = ltTin(h.id, all), ctx = ltCtx(h, r.ds), ban = {};
      if (!c.dang) c.batDau = env.now();
      var luc = c.luc || '';
      r.su.forEach(function (x) { if (x.tg >= luc && x.ngay) ban[x.ngay] = 1; });          // >=: tin gửi cùng giây với lần lưu trước
      ctx.tl.forEach(function (x) { if (x.thoiGian >= luc && x.ky) ban[x.ky] = 1; });
      var ngays = Object.keys(ban).sort().filter(function (d) { return parseD(d) && (!c.dang || d > c.dang); });
      var tn = ltTheoNgay(r.ds), lam = ngays.slice(0, gioiHan || 40), ghi = 0;
      lam.forEach(function (d) {
        if (!(tn[d] || []).length && !ctx.tl.some(function (x) { return x.ky === d; })) return;
        L.ghi(c.nk, d + '.md', ltNhatKy(h, d, tn[d] || [], ctx), 'text/plain'); ghi++;
      });
      var con = ngays.length - lam.length;
      if (con > 0) c.dang = lam[lam.length - 1];
      else { c.luc = c.batDau; c.den = env.now(); delete c.dang; delete c.batDau; }
      ctx.bc.forEach(function (m) { L.ghi(c.bc, ltTenTep(m.tieuDe) + ' [' + m.id.slice(0, 4) + '].csv', ltCSV(m, ctx), 'text/csv'); });
      L.ghi(c.goc, '0. Thông tin nhóm.md', ltThongTin(h, ctx), 'text/plain');
      h.luuTru = JSON.stringify(c); D('HoiThoai');
      return { soNgay: ghi, con: con, den: c.den || '' };
    }
    /* tự động mỗi ngày (Apps Script gọi từ nhacTuDong): lưu các nhóm có tin mới */
    function luuTruTuDong(gioiHanMs) {
      getCfg();                                     // v3.8: áp bảng phân quyền
      if (!env.luuTru) return { soNgay: 0, xong: true };
      var t0 = Date.now(), all = env.chat.before(env.chat.last() + 1, env.chat.last()), n = 0, xong = true;
      T('HoiThoai').filter(function (h) { return h.loai === 'nhom'; }).forEach(function (h) {
        var c = parseJSON(h.luuTru, null), luc = (c && c.luc) || '';
        if (c && !c.dang && !all.some(function (m) { return m.hoiThoai === h.id && m.thoiGian >= luc; })) return;
        if (Date.now() - t0 > (gioiHanMs || 240000)) { xong = false; return; }
        try { var r = ltChay(h, 30, all); n += r.soNgay; if (r.con) xong = false; } catch (e) { }
      });
      return { soNgay: n, xong: xong };
    }
    H['chat.luutru'] = function (req, s) {
      var h = ltQuyen(s, req.ht);
      if (!env.luuTru) err('Máy chủ chưa bật lưu trữ Google Drive.');
      var r = ltChay(h, 25);
      if (s.role === 'admin' && env.luuTru.url) r.thuMucUrl = env.luuTru.url(parseJSON(h.luuTru, {}).goc);
      if (r.soNgay) log(s, 'chat.luutru', h.ten + ': ' + r.soNgay + ' ngày');
      return r;
    };
    H['ai.tepluu'] = function (req, s) {       // danh sách tệp đã lưu của nhóm (để xem lại)
      var h = ltQuyen(s, req.ht), c = ltCay(h, false); if (!c || !env.luuTru) return { ds: [] };
      var out = [];
      [['goc', 'Thông tin'], ['nk', 'Nhật ký từng ngày'], ['bc', 'Báo cáo định kỳ']].forEach(function (p) {
        env.luuTru.ds(c[p[0]]).forEach(function (f) { if (f.mime === 'application/vnd.google-apps.folder') return; f.nhom = p[1]; out.push(f); });
      });
      return { ds: out, thuMucUrl: s.role === 'admin' && env.luuTru.url ? env.luuTru.url(c.goc) : '' };
    };
    H['ai.doctep'] = function (req, s) {
      var h = ltQuyen(s, req.ht), c = ltCay(h, false); if (!c) err('Nhóm chưa được lưu trữ.');
      var id = String(req.id || ''), co = ['goc', 'nk', 'bc'].some(function (k) { return env.luuTru.ds(c[k]).some(function (f) { return f.id === id; }); });
      if (!co) err('Không tìm thấy tệp lưu trữ.');
      return { noiDung: env.luuTru.doc(id) };
    };

    /* ---- Phân tích AI (v2.9: nhóm và trò chuyện riêng, nhật ký thi công, 18 khía cạnh, nhân sự, giải pháp) ---- */
    function aiTV(h) {
      var VT = { truong: 'trưởng nhóm', pho: 'phó nhóm', thanhvien: h.loai === 'rieng' ? 'người trò chuyện' : 'thành viên' }, o = {};
      T('ThanhVien').filter(function (x) { return x.hoiThoai === h.id; }).forEach(function (x) {
        var u = T('Users').filter(function (z) { return z.username === x.username; })[0] || {}, nv = u.maNV ? findNV(u.maNV) : null;
        o[x.username] = { username: x.username, hoTen: chatTen(x.username), vaiTro: x.roi ? 'đã rời nhóm' : (VT[x.vaiTro] || 'thành viên'), vt: x.roi ? 'roi' : (x.vaiTro || 'thanhvien'),
          chucDanh: (nv && nv.chucVu) || ROLES[normRole(u.role)] || '', maNV: u.maNV || '', thamGia: String(x.thamGia || '').slice(0, 10), roi: String(x.roi || '').slice(0, 10) };
      });
      return o;
    }
    function ltChungTu(tins, tu, den) {
      var out = [];
      tins.forEach(function (m, i) {
        if ((m.loai !== 'anh' && m.loai !== 'tep') || m.thuHoi || !m.tep) return;
        var d = m.thoiGian.slice(0, 10); if (d < tu || d > den) return;
        var gc = m.noiDung || '';
        for (var k = Math.max(0, i - 3); !gc && k <= Math.min(tins.length - 1, i + 3); k++) {
          var x = tins[k]; if (x.loai === 'text' && x.username === m.username && !x.thuHoi && Math.abs(ltMs(x.thoiGian) - ltMs(m.thoiGian)) <= 300000) gc = x.noiDung;
        }
        out.push({ ngay: d, gio: m.thoiGian.slice(11, 16), username: m.username, hoTen: m.hoTen, loai: m.loai, ten: m.tep.ten || '', fileId: m.tep.fileId || '', mime: m.tep.mime || '', kichThuoc: num(m.tep.kichThuoc), ghiChu: String(gc).slice(0, 200), moTa: '', lienQuan: '' });
      });
      return out;
    }
    var LT_TT = { text: 1, anh: 1, tep: 1, bcnop: 1, binhchon: 1, nhactao: 1, bctao: 1 };      // tin do chính người dùng tạo = có tương tác
    /* thống kê do máy tính: tin nhắn, ngày làm việc, ngày nghỉ có đơn, ngày không tương tác, báo cáo định kỳ, điểm tham gia */
    function ltThongKe(h, tins, ctx, tu, den) {
      var tv = aiTV(h), cfg = getCfg(), cn = {}, trong = tins.filter(function (m) { var d = m.thoiGian.slice(0, 10); return d >= tu && d <= den; });
      function P(u) { return cn[u] = cn[u] || { username: u, hoTen: (tv[u] || {}).hoTen || chatTen(u), vaiTro: (tv[u] || {}).vaiTro || 'không còn trong nhóm', chucDanh: (tv[u] || {}).chucDanh || '', soTin: 0, soAnh: 0, soTep: 0, ngay: {}, camXucNhan: 0, bcPhai: 0, bcNop: 0, bcDungHan: 0, bcTre: 0, binhChon: 0 }; }
      Object.keys(tv).forEach(function (u) { if (tv[u].vt !== 'roi') P(u); });
      var tn = {}, nd = trong.filter(function (m) { return LT_ND[m.loai] && !m.thuHoi; });
      trong.forEach(function (m) { if (LT_TT[m.loai] && !m.thuHoi && (cn[m.username] || tv[m.username])) P(m.username).ngay[m.thoiGian.slice(0, 10)] = 1; });
      nd.forEach(function (m) {
        var p = P(m.username), d = m.thoiGian.slice(0, 10);
        if (m.loai === 'text') p.soTin++; else if (m.loai === 'anh') p.soAnh++; else p.soTep++;
        Object.keys(m.camXuc || {}).forEach(function (e) { p.camXucNhan += m.camXuc[e]; });
        var t = tn[d] = tn[d] || { ngay: d, soTin: 0, soAnh: 0, soTep: 0, ng: {} };
        if (m.loai === 'text') t.soTin++; else if (m.loai === 'anh') t.soAnh++; else t.soTep++; t.ng[m.username] = 1;
      });
      var baoCao = ctx.bc.map(function (m) {
        var ch = parseJSON(m.cauHinh, {}), kys = Object.keys(ctx.phat[m.id] || {}).filter(function (k) { return k >= tu && k <= den; }).sort(), ng = mucNguoi(m);
        var tl = ctx.tl.filter(function (x) { return x.mucId === m.id && x.ky >= tu && x.ky <= den; }), dh = 0;
        kys.forEach(function () { ng.forEach(function (u) { P(u).bcPhai++; }); });
        tl.forEach(function (x) { var p = P(x.username); p.bcNop++; p.ngay[String(x.thoiGian).slice(0, 10)] = 1; if (ltDungHan(ch, x.ky, x.thoiGian)) { p.bcDungHan++; dh++; } else p.bcTre++; });
        var so = (ch.mau || []).filter(function (f) { return f.kieu === 'so'; }).map(function (f) {
          var tong = 0, n = 0, tng = {};
          tl.forEach(function (x) { var v = parseJSON(x.giaTri, {})[f.id]; if (v !== '' && v != null && isFinite(+v)) { tong += +v; n++; tng[x.ky] = (tng[x.ky] || 0) + (+v); } });
          return { ten: f.ten, tong: Math.round(tong * 1000) / 1000, soLan: n, tb: n ? Math.round(tong / n * 100) / 100 : 0, theoNgay: tng };
        });
        (ch.mau || []).filter(function (f) { return f.kieu === 'bang'; }).forEach(function (f) {   // v3.9: cộng dồn từng dòng của bảng số lượng
          var g = {}, thuTu = [];
          tl.forEach(function (x) { HRCORE_bcBangDoc(parseJSON(x.giaTri, {})[f.id]).forEach(function (r) {
            var v = bcSoVN(r.so); if (v == null) return; var k = r.ten + (r.dv ? ' (' + r.dv + ')' : '');
            if (!g[k]) { g[k] = { tong: 0, n: 0, tng: {} }; thuTu.push(k); } g[k].tong += v; g[k].n++; g[k].tng[x.ky] = (g[k].tng[x.ky] || 0) + v; }); });
          thuTu.slice(0, 25).forEach(function (k) { var o = g[k]; so.push({ ten: f.ten + ' – ' + k, tong: Math.round(o.tong * 1000) / 1000, soLan: o.n, tb: o.n ? Math.round(o.tong / o.n * 100) / 100 : 0, theoNgay: o.tng }); });
        });
        var chon = (ch.mau || []).filter(function (f) { return f.kieu === 'chon' || f.kieu === 'cokhong'; }).map(function (f) {
          var dem = {}; tl.forEach(function (x) { var v = parseJSON(x.giaTri, {})[f.id]; if (v) dem[v] = (dem[v] || 0) + 1; }); return { ten: f.ten, dem: dem };
        });
        return { id: m.id, tieuDe: m.tieuDe, soKy: kys.length, nguoi: ng.length, phaiNop: kys.length * ng.length, daNop: tl.length, dungHan: dh, so: so, chon: chon };
      });
      Object.keys(ctx.muc).forEach(function (id) {
        var m = ctx.muc[id]; if (m.loai !== 'binhchon') return;
        T('ChatTraLoi').forEach(function (x) { var d = String(x.thoiGian).slice(0, 10); if (x.mucId === id && d >= tu && d <= den) { var p = P(x.username); p.binhChon++; p.ngay[d] = 1; } });
      });
      // ngày làm việc (trừ Chủ nhật, lễ), nghỉ có đơn đã duyệt
      var lv = []; for (var d = tu, g = 0; d <= den && g < 400; d = addDays(d, 1), g++) if (!laNgayNghiLe(d, cfg)) lv.push(d);
      var don = T('NghiPhep').filter(function (o) { return o.trangThai === 'Đã duyệt' && o.tuNgay <= den && o.denNgay >= tu; }), nghiNgay = {};
      var maxVol = 0; Object.keys(cn).forEach(function (u) { var p = cn[u]; maxVol = Math.max(maxVol, p.soTin + p.soAnh * 2 + p.soTep * 2 + p.bcNop * 3); });
      var caNhan = Object.keys(cn).map(function (u) {
        var p = cn[u], t = tv[u] || {}, nghi = {}, loai = {};
        don.filter(function (o) { return t.maNV && o.maNV === t.maNV; }).forEach(function (o) {
          var nua = o.tuNgay === o.denNgay && o.buoi && o.buoi !== 'Cả ngày';
          for (var x = o.tuNgay > tu ? o.tuNgay : tu, k = 0; x <= o.denNgay && x <= den && k < 400; x = addDays(x, 1), k++) {
            if (laNgayNghiLe(x, cfg)) continue; nghi[x] = nua ? 0.5 : 1; loai[o.loai] = (loai[o.loai] || 0) + (nua ? 0.5 : 1);
            (nghiNgay[x] = nghiNgay[x] || []).push(p.hoTen + (nua ? ' (½)' : ''));
          }
        });
        var dau = Object.keys(p.ngay).sort()[0] || '', vao = t.thamGia && dau && dau < t.thamGia ? dau : t.thamGia;      // có hoạt động trước ngày vào nhóm (dữ liệu nhập lại) thì tính từ ngày đầu
        var tuP = vao && vao > tu ? vao : tu, denP = t.roi && t.roi < den ? t.roi : den;
        var lvP = lv.filter(function (x) { return x >= tuP && x <= denP; }), coTT = lvP.filter(function (x) { return p.ngay[x]; });
        var khong = lvP.filter(function (x) { return !p.ngay[x] && !nghi[x]; });
        var soNghi = Object.keys(nghi).reduce(function (s, x) { return s + nghi[x]; }, 0), coSo = Math.max(1, lvP.length - soNghi);
        var vol = p.soTin + p.soAnh * 2 + p.soTep * 2 + p.bcNop * 3;
        var diem = p.bcPhai ? 50 * Math.min(1, coTT.length / coSo) + 30 * Math.min(1, p.bcDungHan / p.bcPhai) : 80 * Math.min(1, coTT.length / coSo);
        diem += maxVol ? 20 * vol / maxVol : 0;
        p.soNgayHD = Object.keys(p.ngay).length; delete p.ngay;
        p.soNgayLV = lvP.length; p.soNgayCoTT = coTT.length; p.soNgayKhongTT = khong.length; p.ngayKhongTT = khong.slice(0, 40);
        p.soNgayNghi = soNghi; p.nghiLoai = loai; p.diemTG = Math.round(Math.max(0, Math.min(100, diem)));
        p.tyLeNop = p.bcPhai ? Math.min(100, Math.round(p.bcNop / p.bcPhai * 100)) : null;
        return p;
      }).sort(function (a, b) { return b.diemTG - a.diemTG || b.soTin - a.soTin; });
      var soNgay = Math.round((ltMs(den + ' 00:00') - ltMs(tu + ' 00:00')) / 86400000) + 1;
      var theoNgay = Object.keys(tn).sort().map(function (x) { var t = tn[x]; return { ngay: x, soTin: t.soTin, soAnh: t.soAnh, soTep: t.soTep, soNguoi: Object.keys(t.ng).length }; });
      return { tu: tu, den: den, soNgay: soNgay, soNgayLV: lv.length, soNgayHD: theoNgay.length, nghiTheoNgay: nghiNgay,
        tong: { soTin: nd.filter(function (m) { return m.loai === 'text'; }).length, soAnh: nd.filter(function (m) { return m.loai === 'anh'; }).length, soTep: nd.filter(function (m) { return m.loai === 'tep'; }).length, soNguoi: caNhan.filter(function (p) { return p.soTin + p.soAnh + p.soTep + p.bcNop; }).length, soTV: Object.keys(tv).filter(function (u) { return tv[u].vt !== 'roi'; }).length },
        theoNgay: theoNgay, caNhan: caNhan, baoCao: baoCao };
    }
    function aiDot(ngays, txt, toiDa) {      // chia đợt: ≤ toiDa ngày và ≤ 45.000 ký tự để mỗi lần gọi AI kịp trong 60 giây
      var dot = [], cur = null;
      ngays.forEach(function (d) {
        var n = Math.min(txt[d].length, 45000);
        if (!cur || cur.ngay.length >= toiDa || (cur.kt + n > 45000 && cur.ngay.length)) { cur = { ngay: [], kt: 0 }; dot.push(cur); }
        cur.ngay.push(d); cur.kt += n;
      });
      return dot.map(function (c) { return { tu: c.ngay[0], den: c.ngay[c.ngay.length - 1], ngay: c.ngay }; });
    }
    function aiTVTxt(tv) { return Object.keys(tv).map(function (u) { return '- ' + u + ' = ' + tv[u].hoTen + ' (' + tv[u].vaiTro + (tv[u].chucDanh ? ', ' + tv[u].chucDanh : '') + ')'; }).join('\n'); }
    function aiGoi(o) {
      if (!env.ai) err('Máy chủ chưa bật AI.');
      return aiJSON(env.ai.goi(o));
    }
    function aiTenHT(h) { return h.loai === 'nhom' ? h.ten : chatTenThuMuc(h).replace(/ \[[^\]]*\]$/, '').replace(/^Riêng – /, ''); }
    function aiNhanHT(h) { return (h.loai === 'nhom' ? 'NHÓM: ' : 'TRÒ CHUYỆN RIÊNG GIỮA: ') + aiTenHT(h) + (h.moTa ? ' – ' + h.moTa : ''); }
    /* Lưới an toàn: ngày nào báo cáo CÓ số liệu mà AI trả về rỗng thì điền lại từ số liệu máy bóc.
       Chỉ bù khi AI để trống hoàn toàn mục đó trong ngày — không trộn vào kết quả AI đã có,
       để không nhân đôi số liệu. */
    function aiBuBC(kq, ctx, d, tv) {
      var uOK = function (u) { u = aiS(u, 60); return tv[u] ? u : ''; }, map = {}, bu = 0;
      kq.ngay.forEach(function (x) { map[x.ngay] = x; });
      d.ngay.forEach(function (n) {
        var ds = bcNgaySo(ctx, n); if (!ds.length) return;
        var x = map[n];
        if (!x) { x = aiSachNgay({ ngay: n }, uOK); kq.ngay.push(x); map[n] = x; }
        var hm = [], tb = [], vn = [], vx = [], yc = [], nl = null, bch = null, gop = {};
        ds.forEach(function (o) {
          var p = o.so;
          if (p.nhanLuc != null) nl = (nl || 0) + p.nhanLuc;
          if (p.banChiHuy != null) bch = (bch || 0) + p.banChiHuy;
          p.thiCong.forEach(function (y) { hm.push(y); (gop[o.username] = gop[o.username] || []).push(bcSL(y)); });
          p.thietBi.forEach(function (y) { tb.push(y); });
          p.vatTuNhap.forEach(function (y) { vn.push(y); });
          p.vatTuXuat.forEach(function (y) { vx.push(y); });
          p.yeuCau.forEach(function (y) { yc.push(Object.assign({ username: o.username }, y)); });
        });
        if (!x.thiCong.length && hm.length) {
          x.thiCong = hm.slice(0, 25).map(function (y) { return { hangMuc: y.hangMuc, viTri: '', khoiLuong: y.khoiLuong, dvt: y.dvt, keHoach: null, soVoiKH: 'chưa rõ', ghiChu: 'lấy từ báo cáo định kỳ', trich: y.trich }; });
          bu++;
        }
        if (!x.thietBi.length && tb.length) { x.thietBi = tb.slice(0, 20).map(function (y) { return { ten: y.hangMuc, soLuong: y.khoiLuong, tinhTrang: '', ghiChu: 'lấy từ báo cáo định kỳ' }; }); bu++; }
        if (!x.vatTuNhap.length && vn.length) { x.vatTuNhap = vn.slice(0, 25).map(function (y) { return { ten: y.hangMuc, dvt: y.dvt, soLuong: y.khoiLuong, nhaCungCap: '', trich: y.trich }; }); bu++; }
        if (!x.vatTuXuat.length && vx.length) { x.vatTuXuat = vx.slice(0, 25).map(function (y) { return { ten: y.hangMuc, dvt: y.dvt, soLuong: y.khoiLuong, choHangMuc: '', trich: y.trich }; }); bu++; }
        if (!x.yeuCauVatTu.length && yc.length) { x.yeuCauVatTu = yc.slice(0, 25).map(function (y) { return { ten: y.hangMuc, dvt: y.dvt, soLuong: y.khoiLuong, username: uOK(y.username), trich: y.trich }; }); bu++; }
        if (x.nhanLuc.tong == null && nl != null) { x.nhanLuc.tong = nl; x.nhanLuc.ghiChu = (x.nhanLuc.ghiChu ? x.nhanLuc.ghiChu + '; ' : '') + 'cộng từ báo cáo định kỳ'; bu++; }
        if (x.nhanLuc.banChiHuy == null && bch != null) x.nhanLuc.banChiHuy = bch;
        Object.keys(gop).forEach(function (u) {          // quy việc về đúng người báo cáo
          var uu = uOK(u); if (!uu) return;
          var c = kq.caNhan.filter(function (z) { return z.username === uu; })[0];
          if (!c) { c = { username: uu, dongGop: [], canLuuY: [] }; kq.caNhan.push(c); }
          if (!c.dongGop.length) { c.dongGop = gop[u].slice(0, 6).map(function (t) { return aiS(t, 300); }); bu++; }
        });
      });
      kq.ngay.sort(function (a, b) { return a.ngay < b.ngay ? -1 : 1; });
      return bu;
    }
    function aiPhanDot(h, r, ctx, d, xemAnh) {
      var tn = ltTheoNgay(r.ds), tv = aiTV(h), ncc = env.ai.trangThai().ncc;
      var nk = d.ngay.map(function (n) { return ltRutGon(ltNhatKy(h, n, tn[n] || [], ctx), 45000); }).join('\n\n');
      var ct = ltChungTu(r.ds, d.tu, d.den), anh = [];
      if (xemAnh && env.luuTru && env.luuTru.anh) {
        var ds = ct.filter(function (c) { return /^image\//.test(c.mime); }), buoc = Math.max(1, ds.length / 4);
        for (var i = 0; i < ds.length && anh.length < 4; i += buoc) {
          var c = ds[Math.floor(i)];
          try { var a = env.luuTru.anh(c.fileId, 1500000); if (a) anh.push({ mime: a.mime, data: a.data, nhan: 'Ảnh mã tệp ' + c.fileId + ' – ' + ngayVN(c.ngay) + ' ' + c.gio + ' – ' + c.hoTen + (c.ghiChu ? ' – ' + c.ghiChu : '') }); } catch (e) { }
        }
      }
      var nd = aiNhanHT(h) + '\nTHÀNH VIÊN (tài khoản = họ tên, vai trò, chức danh):\n' + aiTVTxt(tv) +
        '\nKHOẢNG NGÀY CỦA ĐỢT NÀY: ' + ngayVN(d.tu) + ' – ' + ngayVN(d.den) + (anh.length ? '\nKèm ' + anh.length + ' ảnh chứng từ (nhãn ghi mã tệp).' : '') +
        '\n\n=== NHẬT KÝ ===\n' + nk + '\n\n=== YÊU CẦU ===\n' + AI_P_DOT;
      var o = aiGoi({ he: AI_HE, nd: nd, anh: anh, maxTok: ncc === 'claude' ? 6000 : 8000, loai: 'phan', duLieu: { ngay: d.ngay, tin: r.ds.filter(function (m) { var x = m.thoiGian.slice(0, 10); return x >= d.tu && x <= d.den; }), chungTu: ct, tv: tv, ctx: { tl: ctx.tl.filter(function (x) { return x.ky >= d.tu && x.ky <= d.den; }), bc: ctx.bc } } });
      var kq = aiSachDot(o, d, ct, tv);
      aiBuBC(kq, ctx, d, tv);                 // bù lại những gì AI bỏ sót so với báo cáo định kỳ
      return kq;
    }
    function aiS(v, n) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, n || 300); }
    function aiDs(a, n) { return (Array.isArray(a) ? a : []).slice(0, n || 40); }
    function aiO(o) { return o && typeof o === 'object' && !Array.isArray(o) ? o : {}; }
    function aiChon(v, ds, md) { v = aiS(v, 40); var k = ds.filter(function (x) { return x.toLowerCase() === v.toLowerCase(); })[0]; return k || md; }
    function aiSachNgay(x, uOK) {
      x = aiO(x);
      var ds = function (a, n, f) { return aiDs(a, n).map(function (y) { return f(aiO(y)); }).filter(Boolean); };
      var lst = function (a, n, m) { return aiDs(a, n).map(function (y) { return aiS(y, m || 250); }).filter(Boolean); };
      var tt = aiO(x.thoiTiet), nl = aiO(x.nhanLuc), at = aiO(x.anToan);
      var vt = function (extra) { return function (y) { var t = aiS(y.ten, 120); if (!t) return null; var o = { ten: t, dvt: aiS(y.dvt, 20), soLuong: aiSo(y.soLuong), trich: aiS(y.trich, 160) }; extra(o, y); return o; }; };
      return {
        ngay: aiS(x.ngay, 10), tomTat: aiS(x.tomTat, 500),
        thoiTiet: { sang: aiS(tt.sang, 60), chieu: aiS(tt.chieu, 60), anhHuong: aiS(tt.anhHuong, 200) },
        nhanLuc: { tong: aiSo(nl.tong), banChiHuy: aiSo(nl.banChiHuy), congNhan: aiSo(nl.congNhan), thauPhu: ds(nl.thauPhu, 10, function (y) { var t = aiS(y.ten, 80); return t ? { ten: t, soNguoi: aiSo(y.soNguoi) } : null; }), ghiChu: aiS(nl.ghiChu, 200) },
        thietBi: ds(x.thietBi, 20, function (y) { var t = aiS(y.ten, 100); return t ? { ten: t, soLuong: aiSo(y.soLuong), tinhTrang: aiS(y.tinhTrang, 40), ghiChu: aiS(y.ghiChu, 160) } : null; }),
        thiCong: ds(x.thiCong, 25, function (y) {
          var t = aiS(y.hangMuc, 160); if (!t) return null; var kh = aiSo(y.keHoach);
          return { hangMuc: t, viTri: aiS(y.viTri, 100), khoiLuong: aiSo(y.khoiLuong), dvt: aiS(y.dvt, 20), keHoach: kh, soVoiKH: kh == null ? 'chưa rõ' : aiChon(y.soVoiKH, ['đạt', 'vượt', 'chậm', 'chưa rõ'], 'chưa rõ'), ghiChu: aiS(y.ghiChu, 160), trich: aiS(y.trich, 160) };
        }),
        vatTuNhap: ds(x.vatTuNhap, 25, vt(function (o, y) { o.nhaCungCap = aiS(y.nhaCungCap, 100); })),
        vatTuXuat: ds(x.vatTuXuat, 25, vt(function (o, y) { o.choHangMuc = aiS(y.choHangMuc, 100); })),
        yeuCauVatTu: ds(x.yeuCauVatTu, 25, vt(function (o, y) { o.username = uOK(y.username); })),
        nghiemThu: ds(x.nghiemThu, 15, function (y) { var t = aiS(y.noiDung, 200); return t ? { noiDung: t, ketQua: aiS(y.ketQua, 30), thanhPhan: aiS(y.thanhPhan, 120) } : null; }),
        kienNghi: ds(x.kienNghi, 15, function (y) { var t = aiS(y.noiDung, 300); return t ? { username: uOK(y.username), noiDung: t, trich: aiS(y.trich, 160) } : null; }),
        anToan: { tinhHinh: aiS(at.tinhHinh, 300), baoHo: aiS(at.baoHo, 150), viPham: lst(at.viPham, 8) },
        anNinh: aiS(x.anNinh, 300), suCo: lst(x.suCo, 8), tonTai: lst(x.tonTai, 10), keHoachMai: lst(x.keHoachMai, 10)
      };
    }
    var AI_SK = ['an toàn', 'an ninh', 'chất lượng', 'tiến độ', 'thiết bị', 'vật tư', 'tài chính', 'định mức', 'phối hợp', 'khác'];
    function aiSachDot(o, d, ct, tv) {
      o = aiO(o);
      var ngayOK = {}, tep = {}, da = {}; d.ngay.forEach(function (x) { ngayOK[x] = 1; }); ct.forEach(function (c) { tep[c.fileId] = c; });
      var uOK = function (u) { u = aiS(u, 60); return tv[u] ? u : ''; }, nOK = function (x) { x = aiS(x, 10); return ngayOK[x] ? x : ''; };
      var lst = function (a, n) { return aiDs(a, n || 10).map(function (x) { return aiS(x, 300); }).filter(Boolean); };
      return {
        ngay: aiDs(o.ngay, 10).map(function (x) { return aiSachNgay(x, uOK); }).filter(function (x) { if (!ngayOK[x.ngay] || da[x.ngay]) return false; da[x.ngay] = 1; return true; }),
        doiTac: aiDs(o.doiTac, 30).map(function (x) { x = aiO(x); var t = aiS(x.ten, 100); return t ? { ten: t, loai: /cung c|ncc|cung ứng/i.test(aiS(x.loai, 30)) ? 'nhà cung cấp' : 'thầu phụ', ngay: nOK(x.ngay) || d.tu, noiDung: aiS(x.noiDung, 300), danhGia: aiChon(x.danhGia, ['tích cực', 'tiêu cực', 'trung tính'], 'trung tính'), trich: aiS(x.trich, 160) } : null; }).filter(Boolean),
        suKien: aiDs(o.suKien, 30).map(function (x) { x = aiO(x); var t = aiS(x.noiDung, 300); return t ? { ngay: nOK(x.ngay) || d.tu, loai: aiChon(x.loai, AI_SK, 'khác'), mucDo: aiChon(x.mucDo, ['thấp', 'trung bình', 'cao'], 'trung bình'), noiDung: t, trich: aiS(x.trich, 160) } : null; }).filter(Boolean),
        chungTu: aiDs(o.chungTu, 80).map(function (x) { x = aiO(x); return { fileId: aiS(x.fileId, 80), moTa: aiS(x.moTa, 300), lienQuan: aiS(x.lienQuan, 160) }; }).filter(function (x) { return tep[x.fileId]; }),
        caNhan: aiDs(o.caNhan, 60).map(function (x) { x = aiO(x); return { username: uOK(x.username), dongGop: lst(x.dongGop, 6), canLuuY: lst(x.canLuuY, 4) }; }).filter(function (x) { return x.username; })
      };
    }
    function aiGop(ds, khoa, so, so2) {      // cộng theo khóa (tên + đvt + loại), giữ số lần nêu và ngày
      var o = {};
      ds.forEach(function (x) {
        var k = khoa(x), g = o[k.k] = o[k.k] || Object.assign({ tong: 0, coSo: 0, keHoach: 0, lan: 0, ngay: {} }, k.v);
        g.lan++; g.ngay[x.ngay] = 1; var v = x[so]; if (v != null && isFinite(v)) { g.tong += v; g.coSo++; }
        if (so2 && x[so2] != null && isFinite(x[so2])) g.keHoach += x[so2];
      });
      return Object.keys(o).map(function (k) { var g = o[k]; g.tong = Math.round(g.tong * 1000) / 1000; g.keHoach = Math.round(g.keHoach * 1000) / 1000; g.ngay = Object.keys(g.ngay).sort(); return g; }).sort(function (a, b) { return b.lan - a.lan; });
    }
    /* Vật tư gom theo (tên, ĐVT) và tách tổng theo từng hoạt động, để mỗi vật tư chỉ một dòng
       với cột Nhập và cột Xuất đứng cạnh nhau. chenh = nhập − xuất của những lần ĐƯỢC NÊU,
       không phải tồn kho. */
    function aiGopVT(ds) {
      var o = {};
      ds.forEach(function (x) {
        var k = String(x.ten + '|' + x.dvt).toLowerCase().replace(/\s+/g, ' ').trim();
        var g = o[k] = o[k] || { ten: x.ten, dvt: x.dvt, nhap: 0, xuat: 0, deNghi: 0, coNhap: 0, coXuat: 0, coDeNghi: 0, lan: 0, ngay: {} };
        g.lan++; g.ngay[x.ngay] = 1;
        var v = x.soLuong, co = v != null && isFinite(v);
        if (x.hoatDong === 'nhập') { if (co) { g.nhap += v; g.coNhap++; } }
        else if (x.hoatDong === 'xuất') { if (co) { g.xuat += v; g.coXuat++; } }
        else if (co) { g.deNghi += v; g.coDeNghi++; }
      });
      var l = function (v) { return Math.round(v * 1000) / 1000; };
      return Object.keys(o).map(function (k) {
        var g = o[k];
        g.nhap = l(g.nhap); g.xuat = l(g.xuat); g.deNghi = l(g.deNghi);
        g.chenh = (g.coNhap && g.coXuat) ? l(g.nhap - g.xuat) : null;   // chỉ có nghĩa khi biết CẢ nhập và xuất
        g.ngay = Object.keys(g.ngay).sort();
        return g;
      }).sort(function (a, b) { return b.lan - a.lan; });
    }
    function aiKhoaTen(s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').replace(/^(công ty|cty|tnhh|cp)\s+/g, '').trim(); }
    /* gộp kết quả các đợt + số liệu máy tính thành kết quả cuối (chưa có phần đánh giá tổng hợp) */
    function aiGopDot(parts, tk, r, j) {
      var nk = {}, doiTac = {}, suKien = [], ctAI = {}, gn = {};
      parts.forEach(function (p) {
        p.ngay.forEach(function (n) { nk[n.ngay] = n; });
        p.doiTac.forEach(function (x) { var k = aiKhoaTen(x.ten), g = doiTac[k] = doiTac[k] || { ten: x.ten, loai: x.loai, lan: 0, ghiNhan: [], danhGia: '', nhanXet: '', deXuat: '' }; g.lan++; g.ghiNhan.push({ ngay: x.ngay, noiDung: x.noiDung, danhGia: x.danhGia, trich: x.trich }); });
        suKien = suKien.concat(p.suKien);
        p.chungTu.forEach(function (c) { ctAI[c.fileId] = c; });
        p.caNhan.forEach(function (c) { var g = gn[c.username] = gn[c.username] || { dongGop: [], canLuuY: [] }; g.dongGop = g.dongGop.concat(c.dongGop); g.canLuuY = g.canLuuY.concat(c.canLuuY); });
      });
      var tnMap = {}; tk.theoNgay.forEach(function (x) { tnMap[x.ngay] = x; });
      var ngays = tk.theoNgay.map(function (x) { return x.ngay; }).concat(Object.keys(nk)).filter(function (v, k, a) { return a.indexOf(v) === k; }).sort();
      var rong = aiSachNgay({}, function () { return ''; });
      var nhatKy = ngays.map(function (n) {
        var a = nk[n] || Object.assign({}, rong, { ngay: n }), b = tnMap[n] || { soTin: 0, soAnh: 0, soTep: 0, soNguoi: 0 };
        return Object.assign({}, a, { ngay: n, thu: THU_VN[thu(n)], soTin: b.soTin, soAnh: b.soAnh, soTep: b.soTep, soNguoi: b.soNguoi, nghiPhep: tk.nghiTheoNgay[n] || [] });
      });
      var khoiLuong = [], vatTu = [], thietBi = [], kienNghi = {};
      nhatKy.forEach(function (n) {
        n.thiCong.forEach(function (x) { khoiLuong.push(Object.assign({ ngay: n.ngay }, x)); });
        n.vatTuNhap.forEach(function (x) { vatTu.push({ ngay: n.ngay, ten: x.ten, dvt: x.dvt, soLuong: x.soLuong, hoatDong: 'nhập', doiTuong: x.nhaCungCap, trich: x.trich }); });
        n.vatTuXuat.forEach(function (x) { vatTu.push({ ngay: n.ngay, ten: x.ten, dvt: x.dvt, soLuong: x.soLuong, hoatDong: 'xuất', doiTuong: x.choHangMuc, trich: x.trich }); });
        n.yeuCauVatTu.forEach(function (x) { vatTu.push({ ngay: n.ngay, ten: x.ten, dvt: x.dvt, soLuong: x.soLuong, hoatDong: 'đề nghị', doiTuong: x.username ? chatTen(x.username) : '', trich: x.trich }); });
        n.thietBi.forEach(function (x) { if (!/^hoạt động$/i.test(x.tinhTrang) || x.ghiChu) thietBi.push(Object.assign({ ngay: n.ngay }, x)); });
        n.kienNghi.forEach(function (x) { if (x.username) (kienNghi[x.username] = kienNghi[x.username] || []).push({ ngay: n.ngay, noiDung: x.noiDung, trich: x.trich }); });
      });
      var caNhan = tk.caNhan.map(function (p) { var g = gn[p.username] || {}; return Object.assign({}, p, { xepLoai: '', nhanXet: '', deXuat: '', dongGop: (g.dongGop || []).slice(0, 8), canLuuY: (g.canLuuY || []).slice(0, 5), kienNghi: (kienNghi[p.username] || []).slice(0, 15) }); });
      var chungTu = ltChungTu(r.ds, j.tu, j.den).map(function (c) { var a = ctAI[c.fileId]; if (a) { c.moTa = a.moTa; c.lienQuan = a.lienQuan; } return c; });
      return { nhatKy: nhatKy, caNhan: caNhan, khoiLuong: khoiLuong, vatTu: vatTu, thietBi: thietBi, chungTu: chungTu,
        doiTac: Object.keys(doiTac).map(function (k) { return doiTac[k]; }).sort(function (a, b) { return b.lan - a.lan; }),
        suKien: suKien.sort(function (a, b) { return a.ngay < b.ngay ? -1 : 1; }),
        khoiLuongTong: aiGop(khoiLuong, function (x) { return { k: (x.hangMuc + '|' + x.dvt).toLowerCase(), v: { hangMuc: x.hangMuc, dvt: x.dvt } }; }, 'khoiLuong', 'keHoach'),
        vatTuTong: aiGopVT(vatTu) };
    }
    function aiCaNhanTxt(tk) {
      var L = ['- Cá nhân (tài khoản | họ tên | vai trò | chức danh | tin | ảnh | tệp | ngày làm việc | ngày có tương tác | ngày KHÔNG tương tác | ngày nghỉ có đơn | báo cáo nộp/phải | đúng hạn | trễ | bình chọn | điểm tham gia 0–100):'];
      tk.caNhan.forEach(function (p) { L.push('  ' + [p.username, p.hoTen, p.vaiTro, p.chucDanh, p.soTin, p.soAnh, p.soTep, p.soNgayLV, p.soNgayCoTT, p.soNgayKhongTT, p.soNgayNghi, p.bcNop + '/' + p.bcPhai, p.bcDungHan, p.bcTre, p.binhChon, p.diemTG].join(' | ')); });
      return L;
    }
    function aiTongHopTxt(tk, g) {
      var kq0 = { caNhan: tk.caNhan }, dm = function (d) { return d.slice(8, 10) + '/' + d.slice(5, 7); };
      var L = ['THỐNG KÊ DO HỆ THỐNG TÍNH (chính xác, dùng làm căn cứ):',
        '- Khoảng ' + ngayVN(tk.tu) + ' – ' + ngayVN(tk.den) + ': ' + tk.soNgay + ' ngày, ' + tk.soNgayLV + ' ngày làm việc, có hoạt động ' + tk.soNgayHD + ' ngày; ' + tk.tong.soTin + ' tin nhắn, ' + tk.tong.soAnh + ' ảnh, ' + tk.tong.soTep + ' tệp; ' + tk.tong.soNguoi + '/' + tk.tong.soTV + ' thành viên có hoạt động.'].concat(aiCaNhanTxt(tk));
      if (tk.baoCao.length) {
        L.push('- Báo cáo định kỳ:');
        tk.baoCao.forEach(function (b) { L.push('  "' + b.tieuDe + '": phát ' + b.soKy + ' ngày, nộp ' + b.daNop + '/' + b.phaiNop + ', đúng hạn ' + b.dungHan + (b.so.length ? '; số liệu: ' + b.so.map(function (f) { return f.ten + ' tổng ' + f.tong + ' (TB ' + f.tb + '/lần, ' + f.soLan + ' lần)'; }).join('; ') : '') + (b.chon.length ? '; ' + b.chon.map(function (f) { return f.ten + ': ' + Object.keys(f.dem).map(function (k) { return k + ' ' + f.dem[k]; }).join(', '); }).join('; ') : '')); });
      }
      L.push('', 'NHẬT KÝ THI CÔNG TỪNG NGÀY (đã tổng hợp):');
      g.nhatKy.forEach(function (n) {
        var o = [dm(n.ngay) + ' ' + n.thu + ': ' + (n.tomTat || '—')];
        AI_TRUONG_NK.forEach(function (f) { var v = nkO(n, f[0], kq0); if (v) o.push(f[1] + ': ' + v); });
        L.push('- ' + o.join(' | '));
      });
      if (g.suKien.length) { L.push('', 'SỰ CỐ, SỰ KIỆN:'); g.suKien.forEach(function (x) { L.push('- ' + dm(x.ngay) + ' [' + x.loai + ', ' + x.mucDo + '] ' + x.noiDung); }); }
      if (g.doiTac.length) { L.push('', 'NHÀ THẦU PHỤ, NHÀ CUNG CẤP (ghi nhận):'); g.doiTac.forEach(function (x) { L.push('- ' + x.ten + ' (' + x.loai + ', ' + x.lan + ' lần): ' + x.ghiNhan.slice(0, 6).map(function (k) { return dm(k.ngay) + ' ' + k.danhGia + ': ' + k.noiDung; }).join('; ')); }); }
      if (g.vatTuTong.length) { L.push('', 'VẬT TƯ (cộng theo các lần được nêu; chênh lệch nhập − xuất không phải tồn kho):'); g.vatTuTong.slice(0, 50).forEach(function (x) { L.push('- ' + x.ten + (x.dvt ? ' (' + x.dvt + ')' : '') + ': nhập ' + (x.coNhap ? x.nhap : '—') + ', xuất ' + (x.coXuat ? x.xuat : '—') + (x.coDeNghi ? ', đề nghị cấp ' + x.deNghi : '') + (x.chenh == null ? '' : ', chênh lệch ' + x.chenh) + ', ' + x.lan + ' lần nêu'); }); }
      if (g.khoiLuongTong.length) { L.push('', 'KHỐI LƯỢNG THI CÔNG:'); g.khoiLuongTong.slice(0, 50).forEach(function (x) { L.push('- ' + x.hangMuc + ': thực hiện ' + (x.coSo ? x.tong + ' ' + x.dvt : 'không rõ') + (x.keHoach ? ', kế hoạch nêu ' + x.keHoach + ' ' + x.dvt : '') + ', ' + x.lan + ' ngày'); }); }
      var s = L.join('\n');
      return s.length > 110000 ? s.slice(0, 110000) + '\n…(đã rút gọn)' : s;
    }
    function aiTong2Txt(tk, g, t1) {
      var dm = function (d) { return d.slice(8, 10) + '/' + d.slice(5, 7); }, L = ['SỐ LIỆU CÁ NHÂN DO HỆ THỐNG TÍNH:'].concat(aiCaNhanTxt(tk));
      L.push('', 'GHI NHẬN, KIẾN NGHỊ CÁ NHÂN (từ nhật ký):');
      g.caNhan.forEach(function (c) { L.push('- ' + c.username + ': đóng góp: ' + (c.dongGop.join('; ') || '—') + (c.canLuuY.length ? ' | cần lưu ý: ' + c.canLuuY.join('; ') : '') + (c.kienNghi.length ? ' | kiến nghị: ' + c.kienNghi.map(function (k) { return dm(k.ngay) + ' ' + k.noiDung; }).join('; ') : '')); });
      if (g.doiTac.length) { L.push('', 'NHÀ THẦU PHỤ, NHÀ CUNG CẤP:'); g.doiTac.forEach(function (x) { L.push('- ' + x.ten + ' (' + x.loai + ', ' + x.lan + ' lần): ' + x.ghiNhan.slice(0, 8).map(function (k) { return dm(k.ngay) + ' ' + k.danhGia + ': ' + k.noiDung; }).join('; ')); }); }
      L.push('', 'ĐÁNH GIÁ KHÍA CẠNH ĐÃ LẬP:'); (t1.danhGia || []).forEach(function (x) { L.push('- ' + x.khia + ' [' + x.muc + ']: ' + x.nhanXet); });
      var tt = []; g.nhatKy.forEach(function (n) { n.tonTai.concat(n.suCo).forEach(function (x) { tt.push(dm(n.ngay) + ': ' + x); }); });
      if (tt.length) { L.push('', 'TỒN TẠI, SỰ CỐ THEO NGÀY:'); tt.slice(0, 80).forEach(function (x) { L.push('- ' + x); }); }
      var s = L.join('\n');
      return s.length > 90000 ? s.slice(0, 90000) + '\n…(đã rút gọn)' : s;
    }
    function aiSachT1(o) {
      o = aiO(o); var map = {};
      aiDs(o.danhGia, 30).forEach(function (x) { x = aiO(x); var k = aiS(x.khia, 80).toLowerCase(), kk = AI_KHIA.filter(function (a) { return a.toLowerCase() === k || a.toLowerCase().indexOf(k) === 0 || k.indexOf(a.toLowerCase().split(' (')[0]) === 0; })[0]; if (kk && !map[kk]) map[kk] = x; });
      return { tongThe: aiS(o.tongThe, 3500), mucRuiRo: aiChon(o.mucRuiRo, ['Thấp', 'Trung bình', 'Cao'], ''), lyDoRuiRo: aiS(o.lyDoRuiRo, 500),
        danhGia: AI_KHIA.map(function (k) { var x = map[k] || {}; return { khia: k, muc: aiChon(x.muc, AI_MUC, 'Chưa đủ dữ liệu'), nhanXet: aiS(x.nhanXet, 900) || 'AI không đánh giá khía cạnh này.', danChung: aiDs(x.danChung, 6).map(function (y) { return aiS(y, 250); }).filter(Boolean), giaiPhap: aiDs(x.giaiPhap, 6).map(function (y) { return aiS(y, 300); }).filter(Boolean) }; }) };
    }
    function aiSachT2(o, tv) {
      o = aiO(o); var lst = function (a, n) { return aiDs(a, n || 10).map(function (x) { return aiS(x, 400); }).filter(Boolean); };
      return {
        caNhan: aiDs(o.caNhan, 80).map(function (x) { x = aiO(x); var u = aiS(x.username, 60); return tv[u] ? { username: u, xepLoai: aiChon(x.xepLoai, AI_HANG, 'Đạt'), nhanXet: aiS(x.nhanXet, 600), deXuat: aiS(x.deXuat, 250) } : null; }).filter(Boolean),
        doiTac: aiDs(o.doiTac, 40).map(function (x) { x = aiO(x); var t = aiS(x.ten, 100); return t ? { ten: t, danhGia: aiChon(x.danhGia, ['Tốt', 'Đạt', 'Cần chú ý', 'Kém'], 'Đạt'), nhanXet: aiS(x.nhanXet, 500), deXuat: aiS(x.deXuat, 300) } : null; }).filter(Boolean),
        diemManh: lst(o.diemManh), tonTai: lst(o.tonTai, 12),
        giaiPhap: aiDs(o.giaiPhap, 20).map(function (x) { x = aiO(x); var v = aiS(x.vanDe, 300), gp = aiS(x.giaiPhap, 500); return v || gp ? { vanDe: v, giaiPhap: gp, phuTrach: aiS(x.phuTrach, 100), thoiHan: aiS(x.thoiHan, 40), uuTien: aiChon(x.uuTien, ['Cao', 'Trung bình', 'Thấp'], 'Trung bình') } : null; }).filter(Boolean)
          .sort(function (a, b) { var r = { 'Cao': 0, 'Trung bình': 1, 'Thấp': 2 }; return r[a.uuTien] - r[b.uuTien]; }),
        kienNghiBGD: lst(o.kienNghiBGD, 10)
      };
    }
    function aiTenNgan(d) { return d.slice(8, 10) + '/' + d.slice(5, 7); }
    H['ai.nhom'] = function (req, s) {
      if (!can(s.role, 'aiNhom')) err('Bạn chưa được phân quyền Phân tích nhóm (AI).');
      var tatCa = s.role === 'admin' || s.role === 'giamdoc', VT = {};       // trò chuyện riêng: chỉ Giám đốc, Quản trị
      chatTVs(s.username).forEach(function (t) { if (t.vaiTro === 'truong' || t.vaiTro === 'pho') VT[t.hoiThoai] = t.vaiTro; });
      var ds = T('HoiThoai').filter(function (h) { return h.loai === 'nhom' ? true : (tatCa && h.loai === 'rieng' && h.tinCuoi); }).map(function (h) {
        var c = parseJSON(h.luuTru, null), tc = parseJSON(h.tinCuoi, null);
        return { id: h.id, loai: h.loai, ten: aiTenHT(h), mau: h.mau, moTa: h.moTa, trangThai: h.trangThai, soTV: chatMem(h.id).length, capNhat: h.capNhat, tinCuoi: tc ? tc.thoiGian : '', daLuu: !!(c && c.goc), luuDen: c ? (c.den || '') : '', dangLuu: !!(c && c.dang), vaiTro: VT[h.id] || '' };
      }).sort(function (a, b) { return a.capNhat < b.capNhat ? 1 : -1; });
      var out = { ds: ds, ai: env.ai ? env.ai.trangThai() : null, luuTru: !!env.luuTru };
      if (s.role === 'admin' && env.luuTru) { try { out.canhBao = env.luuTru.canhBao ? env.luuTru.canhBao() : ''; out.thuMucUrl = env.luuTru.url ? env.luuTru.url('') : ''; } catch (e) { out.canhBao = 'Không kiểm tra được thư mục lưu trữ: ' + e.message; } }
      return out;
    };
    var AI_NCC_DS = ['gemini', 'openai', 'claude', 'deepseek'];
    H['ai.cauhinh'] = function (req, s) {
      if (s.role !== 'admin') err('Chỉ Quản trị cài đặt AI.');
      if (!env.ai) err('Máy chủ chưa bật AI.');
      var ncc = AI_NCC_DS.indexOf(req.ncc) >= 0 ? req.ncc : '';
      if (!ncc) err('Chọn nhà cung cấp AI: Gemini, GPT (OpenAI), Claude hoặc DeepSeek.');
      var key = String(req.key || '').trim();
      if (key && (key.length < 20 || /\s/.test(key))) err('Khóa API không hợp lệ.');
      var m1 = aiS(req.model, 80), m2 = req.model2 === undefined ? undefined : aiS(req.model2, 80);
      if ([m1, m2].some(function (m) { return m && !/^[\w.\-:\/]+$/.test(m); })) err('Tên model chỉ gồm chữ, số và . - _ : /');
      env.ai.dat({ ncc: ncc, model: m1, model2: m2, key: key, xoaKhoa: !!req.xoaKhoa });
      log(s, 'ai.cauhinh', ncc + ' ' + (m1 || 'mặc định') + (m2 ? ' / tổng hợp ' + m2 : '') + (key ? ' (cập nhật khóa)' : '') + (req.xoaKhoa ? ' (xóa khóa)' : ''));
      return env.ai.trangThai();
    };
    H['ai.models'] = function (req, s) {            // v3.15: danh sách model đang dùng được của khóa API (để chọn trong Cài đặt AI)
      if (s.role !== 'admin') err('Chỉ Quản trị.');
      if (!env.ai || !env.ai.dsModel) err('Máy chủ chưa hỗ trợ lấy danh sách model.');
      var ncc = AI_NCC_DS.indexOf(req.ncc) >= 0 ? req.ncc : ''; if (!ncc) err('Chọn nhà cung cấp AI.');
      var key = String(req.key || '').trim();
      return { ncc: ncc, ds: env.ai.dsModel(ncc, key).slice(0, 300) };
    };
    H['ai.thu'] = function (req, s) {
      if (s.role !== 'admin') err('Chỉ Quản trị.');
      if (!env.ai) err('Máy chủ chưa bật AI.');
      var t0 = Date.now(), o = aiJSON(env.ai.goi({ he: 'Chỉ trả về JSON.', nd: 'Trả về đúng JSON {"ok": true, "loiChao": "<một câu chào ngắn bằng tiếng Việt>"}', anh: [], maxTok: 200, loai: 'thu' }));
      return { ok: !!o.ok, loiChao: aiS(o.loiChao, 200), ms: Date.now() - t0, ai: env.ai.trangThai() };
    };
    /* Các bước: batdau → phan (mỗi đợt) → tong1 (tổng thể, 18 khía cạnh) → tong2 (nhân sự, đối tác, giải pháp) → lap (báo cáo: Docs + PDF + JSON) → lapnk (nhật ký thi công: Docs + PDF) */
    H['ai.phantich'] = function (req, s) {
      if (!env.ai || !env.tam) err('Máy chủ chưa bật AI.');
      if (req.buoc === 'batdau') {
        var h = ltQuyen(s, req.ht), tu = String(req.tu || ''), den = String(req.den || '');
        if (!env.ai.trangThai().coKhoa) err('Chưa cài đặt AI. Quản trị vào "Phân tích nhóm (AI)" → "Cài đặt AI" để nhập khóa API.');
        if (!parseD(tu) || !parseD(den) || tu > den) err('Khoảng ngày không hợp lệ.');
        if (den > env.today()) den = env.today();
        if (ltMs(den + ' 00:00') - ltMs(tu + ' 00:00') > 366 * 86400000) err('Mỗi lần phân tích tối đa 12 tháng.');
        if (!ltCay(h, false)) err('Chưa lưu trữ. Bấm "Lưu trữ ngay" trước.');
        var r = ltTin(h.id), ctx = ltCtx(h, r.ds), tn = ltTheoNgay(r.ds), txt = {};
        var ngays = ltNgays(r.ds, ctx).filter(function (d) { return d >= tu && d <= den; });
        if (!ngays.length) err('Không có hoạt động nào trong khoảng ngày đã chọn.');
        ngays.forEach(function (d) { txt[d] = ltNhatKy(h, d, tn[d] || [], ctx); });
        var dot = aiDot(ngays, txt, env.ai.trangThai().ncc === 'claude' ? 2 : 3), job = newId();
        env.tam.put('ai_' + job, { ht: h.id, tu: tu, den: den, dot: dot, xemAnh: !!req.xemAnh, u: s.username, nhan: env.now().replace(/[-:]/g, '').replace(' ', '-').slice(0, 13) }, 21600);
        return { job: job, dot: dot.map(function (d) { return { tu: d.tu, den: d.den, soNgay: d.ngay.length }; }), ai: env.ai.trangThai() };
      }
      var J = 'ai_' + String(req.job || ''), j = env.tam.get(J);
      if (!j || j.u !== s.username) err('Phiên phân tích đã hết hạn, hãy chạy lại.');
      var h2 = ltQuyen(s, j.ht), c2 = ltCay(h2, false), st2 = env.ai.trangThai();
      var r2 = ltTin(h2.id), ctx2 = ltCtx(h2, r2.ds), tv = aiTV(h2);
      if (req.buoc === 'phan') {
        var i = num(req.i), d = j.dot[i]; if (!d) err('Đợt không hợp lệ.');
        var kq = aiPhanDot(h2, r2, ctx2, d, j.xemAnh);
        env.tam.put(J + '_' + i, kq, 21600);
        var dem = function (k) { return kq.ngay.reduce(function (s, n) { return s + n[k].length; }, 0); };
        return { i: i, soNgay: kq.ngay.length, soThiCong: dem('thiCong'), soVatTu: dem('vatTuNhap') + dem('vatTuXuat') + dem('yeuCauVatTu'), soKienNghi: dem('kienNghi'), soSuKien: kq.suKien.length };
      }
      var parts = j.dot.map(function (d, k) { var p = env.tam.get(J + '_' + k); if (!p) err('Thiếu kết quả đợt ' + (k + 1) + ' (' + ngayVN(d.tu) + '), hãy chạy lại phân tích.'); return p; });
      var tk = ltThongKe(h2, r2.ds, ctx2, j.tu, j.den), g = aiGopDot(parts, tk, r2, j);
      var dau = aiNhanHT(h2) + '\nTHÀNH VIÊN:\n' + aiTVTxt(tv) + '\n\n';
      var dauKQ = function () { return { phienBan: 2, congTy: getCfg().congTy || '', nhom: { id: h2.id, loai: h2.loai, ten: aiTenHT(h2), moTa: h2.moTa, soTV: tk.tong.soTV }, tu: j.tu, den: j.den }; };
      if (req.buoc === 'lapnk') {        // v3.15: nhật ký thi công chỉ cần kết quả các đợt → chạy song song với bước tổng hợp
        var o2 = {};
        if (c2 && env.luuTru) {
          var kqN = Object.assign(dauKQ(), { thongKe: tk }, g);
          try { var dn = env.luuTru.taoDoc(c2.ai, 'Nhật ký thi công ' + aiTenHT(h2) + ' ' + aiTenNgan(j.tu) + '–' + ngayVN(j.den) + ' (lập ' + j.nhan + ')', nkKhoi(kqN), { pdf: true }); o2 = { nkDocId: dn.id, nkDocUrl: dn.url, nkPdfId: dn.pdfId || '' }; } catch (e) { o2 = { loiNK: String(e.message || e) }; }
        }
        env.tam.put(J + '_nk', o2, 21600);
        var kqDa = env.tam.get(J + '_kq');                // báo cáo đã lập xong trước: gắn luôn nhật ký vào tệp JSON (bước "chot" vẫn gộp lại lần cuối)
        if (kqDa) { Object.assign(kqDa, o2); env.tam.put(J + '_kq', kqDa, 21600); if (c2 && env.luuTru && kqDa.jsonId) env.luuTru.ghi(c2.ai, 'PhanTich_' + j.tu + '_' + j.den + '_' + j.nhan + '.json', JSON.stringify(kqDa), 'application/json'); }
        return o2;
      }
      if (req.buoc === 'chot') {          // gộp kết quả báo cáo + nhật ký, ghi tệp JSON của đợt phân tích
        var kqC = env.tam.get(J + '_kq'); if (!kqC) err('Chưa lập báo cáo, hãy chạy lại bước lập báo cáo.');
        Object.assign(kqC, env.tam.get(J + '_nk') || {});
        if (c2 && env.luuTru && kqC.jsonId) env.luuTru.ghi(c2.ai, 'PhanTich_' + j.tu + '_' + j.den + '_' + j.nhan + '.json', JSON.stringify(kqC), 'application/json');
        env.tam.put(J + '_kq', kqC, 21600);
        return kqC;
      }
      if (req.buoc === 'tong1') {
        var t1 = aiSachT1(aiGoi({ he: AI_HE, nd: dau + aiTongHopTxt(tk, g) + '\n\n=== YÊU CẦU ===\n' + AI_P_TONG1, anh: [], maxTok: st2.ncc === 'claude' ? 6000 : 8000, loai: 'tong1', duLieu: { thongKe: tk, gop: g } }));
        env.tam.put(J + '_t1', t1, 21600);
        return { mucRuiRo: t1.mucRuiRo, soKhia: t1.danhGia.filter(function (x) { return x.muc !== 'Chưa đủ dữ liệu'; }).length };
      }
      var t1b = env.tam.get(J + '_t1'); if (!t1b) err('Thiếu bước tổng hợp 1, hãy chạy lại.');
      if (req.buoc === 'tong2') {
        var t2 = aiSachT2(aiGoi({ he: AI_HE, nd: dau + aiTong2Txt(tk, g, t1b) + '\n\n=== YÊU CẦU ===\n' + AI_P_TONG2, anh: [], maxTok: st2.ncc === 'claude' ? 6000 : 8000, loai: 'tong2', duLieu: { thongKe: tk, gop: g, t1: t1b } }), tv);
        env.tam.put(J + '_t2', t2, 21600);
        return { soGiaiPhap: t2.giaiPhap.length };
      }
      if (req.buoc !== 'lap') err('Bước không hợp lệ.');
      var t2b = env.tam.get(J + '_t2'); if (!t2b) err('Thiếu bước tổng hợp 2, hãy chạy lại.');
      var xl = {}; t2b.caNhan.forEach(function (c) { xl[c.username] = c; });
      g.caNhan.forEach(function (p) { var a = xl[p.username] || {}; p.xepLoai = a.xepLoai || (p.soTin + p.soAnh + p.soTep + p.bcNop ? 'Đạt' : 'Ít hoạt động'); p.nhanXet = a.nhanXet || ''; p.deXuat = a.deXuat || ''; });
      var dt = {}; t2b.doiTac.forEach(function (x) { dt[aiKhoaTen(x.ten)] = x; });
      g.doiTac.forEach(function (x) { var a = dt[aiKhoaTen(x.ten)]; if (a) { x.danhGia = a.danhGia; x.nhanXet = a.nhanXet; x.deXuat = a.deXuat; } });
      var kq2 = Object.assign(dauKQ(), { taoBoi: s.username, taoBoiTen: s.hoTen || chatTen(s.username), taoLuc: env.now(),
        ai: { ncc: st2.ncc, model: st2.model, xemAnh: !!j.xemAnh }, thongKe: tk,
        tong: { tongThe: t1b.tongThe, mucRuiRo: t1b.mucRuiRo, lyDoRuiRo: t1b.lyDoRuiRo, danhGia: t1b.danhGia, diemManh: t2b.diemManh, tonTai: t2b.tonTai, giaiPhap: t2b.giaiPhap, kienNghiBGD: t2b.kienNghiBGD } }, g);
      if (c2 && env.luuTru) {
        var anhDoc = [];
        if (env.luuTru.anh) g.chungTu.filter(function (c) { return /^image\//.test(c.mime); }).sort(function (a, b) { return (b.moTa ? 1 : 0) - (a.moTa ? 1 : 0); }).slice(0, 12).forEach(function (c) {
          try { var a = env.luuTru.anh(c.fileId, 900000); if (a) anhDoc.push({ mime: a.mime, data: a.data, fileId: c.fileId, nhan: ngayVN(c.ngay) + ' ' + c.gio + ' – ' + c.hoTen + ' – ' + (c.moTa || c.ghiChu || c.ten) }); } catch (e) { }
        });
        try { var dc = env.luuTru.taoDoc(c2.ai, 'Báo cáo phân tích ' + aiTenHT(h2) + ' ' + aiTenNgan(j.tu) + '–' + ngayVN(j.den) + ' (lập ' + j.nhan + ')', aiKhoi(kq2, anhDoc), { pdf: true }); kq2.docId = dc.id; kq2.docUrl = dc.url; kq2.pdfId = dc.pdfId || ''; } catch (e) { kq2.loiDoc = String(e.message || e); }
        kq2.jsonId = env.luuTru.ghi(c2.ai, 'PhanTich_' + j.tu + '_' + j.den + '_' + j.nhan + '.json', JSON.stringify(kq2), 'application/json');
      }
      env.tam.put(J + '_kq', kq2, 21600);
      log(s, 'ai.phantich', aiTenHT(h2) + ' ' + j.tu + '→' + j.den + ' (' + st2.ncc + ')');
      return kq2;
    };
    /* Lưu tệp kết quả (Excel) vào thư mục "4. Phân tích AI" của nhóm. Trình duyệt dựng tệp
       (thư viện Excel chạy ở trình duyệt), máy chủ chỉ nhận base64 và ghi vào Drive.
       jsonId: ghi thêm mã tệp vào JSON của đợt phân tích để mở lại báo cáo cũ vẫn tải được. */
    H['ai.luutep'] = function (req, s) {
      var h = ltQuyen(s, req.ht), c = ltCay(h, false); if (!c) err('Nhóm chưa được lưu trữ.');
      if (!env.luuTru || !env.luuTru.ghiNhiPhan) err('Máy chủ chưa hỗ trợ lưu tệp nhị phân.');
      var data = String(req.data || ''); if (!data) err('Thiếu nội dung tệp.');
      if (data.length > 24000000) err('Tệp quá lớn (tối đa khoảng 18 MB).');
      var ten = ltTenTep(req.ten);
      if (!/\.(xlsx|docx|pdf)$/i.test(ten)) err('Chỉ lưu tệp .xlsx, .docx hoặc .pdf.');
      var mime = String(req.mime || '') || (/\.xlsx$/i.test(ten) ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        : /\.docx$/i.test(ten) ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' : 'application/pdf');
      var id = env.luuTru.ghiNhiPhan(c.ai, ten, data, mime);
      var khoa = /\.xlsx$/i.test(ten) ? 'xlsxId' : /\.docx$/i.test(ten) ? 'docxId' : 'pdfTaiId';
      if (req.jsonId) {        // gắn mã tệp vào JSON của đợt, bỏ qua nếu không cập nhật được
        try {
          var f = env.luuTru.ds(c.ai).filter(function (x) { return x.id === String(req.jsonId); })[0];
          if (f) { var o = parseJSON(env.luuTru.doc(f.id), null); if (o) { o[khoa] = id; env.luuTru.ghi(c.ai, f.ten, JSON.stringify(o), 'application/json'); } }
        } catch (e) { }
      }
      log(s, 'ai.luutep', aiTenHT(h) + ' – ' + ten);
      return { id: id, ten: ten, khoa: khoa };
    };
    H['ai.taifile'] = function (req, s) {    // tải PDF / tệp trong "4. Phân tích AI" về máy
      var h = ltQuyen(s, req.ht), c = ltCay(h, false); if (!c || !env.luuTru || !env.luuTru.tai) err('Không có tệp.');
      var f = env.luuTru.ds(c.ai).filter(function (x) { return x.id === String(req.id || ''); })[0];
      if (!f || /google-apps/.test(f.mime)) err('Không tìm thấy tệp.');
      return { data: env.luuTru.tai(f.id), mime: f.mime, name: f.ten };
    };
    H['ai.ds'] = function (req, s) {
      var h = ltQuyen(s, req.ht), c = ltCay(h, false); if (!c || !env.luuTru) return [];
      return env.luuTru.ds(c.ai).filter(function (f) { return /^PhanTich_.*\.json$/.test(f.ten); }).map(function (f) {
        var m = /^PhanTich_(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})_(\d{8}-\d{4})/.exec(f.ten) || [];
        return { id: f.id, ten: f.ten, tu: m[1] || '', den: m[2] || '', lap: m[3] || '', capNhat: f.capNhat };
      }).sort(function (a, b) { return a.ten < b.ten ? 1 : -1; });
    };
    H['ai.xem'] = function (req, s) {
      var h = ltQuyen(s, req.ht), c = ltCay(h, false); if (!c) err('Nhóm chưa được lưu trữ.');
      var id = String(req.id || '');
      if (!env.luuTru.ds(c.ai).some(function (f) { return f.id === id; })) err('Không tìm thấy báo cáo phân tích.');
      var o = parseJSON(env.luuTru.doc(id), null); if (!o) err('Tệp báo cáo bị hỏng.');
      return o;
    };
    H['ai.tep'] = function (req, s) {         // ảnh chứng từ cho người xem báo cáo (Giám đốc có thể không ở trong nhóm)
      var h = ltQuyen(s, req.ht), t = T('TepChat').filter(function (x) { return x.fileId === String(req.fileId || '') && x.hoiThoai === h.id; })[0];
      if (!t) err('Tệp không còn (có thể đã thu hồi).');
      var d = env.files.get(t.fileId); if (!d) err('Tệp không còn trên máy chủ.');
      return { data: d.data, mime: t.mime, name: t.ten };
    };
    H['chat.nhatky'] = function (req, s) {    // dữ liệu xuất Excel "Nhật ký từng ngày"
      var h = ltQuyen(s, req.ht), tu = String(req.tu || ''), den = String(req.den || '');
      if (!parseD(tu) || !parseD(den) || tu > den) err('Khoảng ngày không hợp lệ.');
      var r = ltTin(h.id), ctx = ltCtx(h, r.ds), tn = ltTheoNgay(r.ds), LOAI = { text: 'Tin nhắn', anh: 'Ảnh', tep: 'Tệp', hethong: 'Hệ thống', binhchon: 'Bình chọn', nhactao: 'Nhắc hẹn', nhac: 'Nhắc hẹn', bctao: 'Yêu cầu báo cáo', baocao: 'Yêu cầu báo cáo', bcnop: 'Nộp báo cáo' };
      var ngay = ltNgays(r.ds, ctx).filter(function (d) { return d >= tu && d <= den; }).map(function (d) {
        return { ngay: d, thu: THU_VN[thu(d)], dong: (tn[d] || []).map(function (m) {
          var line = ltDong(m, ctx).replace(/^- \d\d:\d\d · /, ''); if (m.hoTen && line.indexOf(m.hoTen + ': ') === 0) line = line.slice(m.hoTen.length + 2);
          return { gio: m.thoiGian.slice(11, 16), hoTen: m.loai === 'hethong' || m.loai === 'nhac' || m.loai === 'baocao' ? '' : m.hoTen, loai: m.thuHoi ? 'Đã thu hồi' : (LOAI[m.loai] || m.loai), noiDung: line, tep: m.tep && m.tep.ten ? m.tep.ten : '' };
        }) };
      });
      var bc = ctx.bc.map(function (m) {
        var ch = parseJSON(m.cauHinh, {}), mau = ch.mau || [];
        return { tieuDe: m.tieuDe, cot: ['Ngày báo cáo', 'Họ tên', 'Nộp lúc', 'Đúng hạn'].concat(mau.map(function (f) { return f.ten; })), so: mau.map(function (f) { return f.kieu === 'so'; }),
          dong: ctx.tl.filter(function (x) { return x.mucId === m.id && x.ky >= tu && x.ky <= den; }).sort(function (a, b) { return a.ky < b.ky ? -1 : a.ky > b.ky ? 1 : 0; }).map(function (x) {
            var gv = parseJSON(x.giaTri, {}); return [x.ky, chatTen(x.username), x.thoiGian, ltDungHan(ch, x.ky, x.thoiGian) ? 'Có' : 'Trễ'].concat(mau.map(function (f) { var v = gv[f.id]; return v == null ? '' : v; }));
          }) };
      });
      return { nhom: aiTenHT(h), tu: tu, den: den, ngay: ngay, bc: bc };
    };

    /* ==================== THÔNG BÁO ĐẨY (v3.1) ====================
       Mỗi sự kiện (tin nhắn, bản tin, ý kiến, đơn chờ duyệt, công tác…) xếp vào hàng đợi PUSH theo tài khoản nhận.
       Cuối lượt xử lý, hàng đợi được đổi thành danh sách { token, data } theo thiết bị đã đăng ký; Apps Script gửi qua
       Firebase Cloud Messaging sau khi nhả khóa ghi (không làm chậm người khác). Điện thoại hiện thông báo trên thanh trạng thái, có chuông. */
    var PUSH = [];
    var PUSH_LOAI = { tn: 'Tin nhắn', bt: 'Bảng tin', yk: 'Ý kiến – kiến nghị', duyet: 'Việc cần duyệt', cc: 'Chấm công – công tác' };
    function pushBat() { var p = getCfg().push || {}; return !!(p.bat && p.firebase && p.firebase.projectId); }
    function bao(users, o) {
      if (!pushBat()) return;
      var da = {};
      (users || []).forEach(function (u) { if (!u || da[u]) return; da[u] = 1; PUSH.push({ u: u, loai: o.loai, t: String(o.t || '').slice(0, 120), b: String(o.b || '').slice(0, 300), tag: String(o.tag || o.loai), mo: String(o.mo || '') }); });
    }
    function userActive() { return T('Users').filter(function (u) { return u.active !== 'Không'; }); }
    function nguoiCo(list, tru) {                     // tài khoản thuộc danh sách người xử lý ('r:vaiTro', 'u:tenDN'); Quản trị chỉ khi được ghi đích danh
      list = list || [];
      return userActive().filter(function (u) { return u.username !== tru && list.some(function (a) { return a === 'u:' + u.username || (a === 'r:' + u.role && u.role !== 'admin'); }); }).map(function (u) { return u.username; });
    }
    function userCuaNV(maNV) { var u = userActive().filter(function (x) { return x.maNV === maNV; })[0]; return u ? u.username : ''; }
    function qlCuaNV(maNV, tru) {                     // trưởng bộ phận / CHT của người này; không có thì TCHC
      var nv = findNV(maNV), bp = nv ? nv.phongBan : '';
      var ql = userActive().filter(function (u) { return u.username !== tru && ROLE_QL.indexOf(u.role) >= 0 && u.boPhan && u.boPhan === bp; }).map(function (u) { return u.username; });
      return ql.length ? ql : userActive().filter(function (u) { return u.username !== tru && can(u.role, 'edit_nv') && u.role !== 'admin'; }).map(function (u) { return u.username; });
    }
    function pushTuyChon(r) { return Object.assign({ tn: true, bt: true, yk: true, duyet: true, cc: true, xemTruoc: true }, parseJSON(r.tuyChon, {})); }
    function pushDs() {
      var tb = {}; T('ThietBiPush').forEach(function (r) { if (r.token) (tb[r.username] = tb[r.username] || []).push(r); });
      var out = [];
      PUSH.forEach(function (q) {
        (tb[q.u] || []).forEach(function (r) {
          var tc = pushTuyChon(r); if (tc[q.loai] === false) return;
          out.push({ token: r.token, u: q.u, loai: q.loai, data: { t: tc.xemTruoc ? q.t : 'TAPACO', b: tc.xemTruoc ? q.b : 'Bạn có ' + (PUSH_LOAI[q.loai] || 'thông báo').toLowerCase() + ' mới', tag: q.tag, mo: q.mo, loai: q.loai } });
        });
      });
      return out.slice(0, 500);
    }
    function xoaTokenPush(tokens) {
      var a = T('ThietBiPush'), n0 = a.length;
      for (var i = a.length - 1; i >= 0; i--) if ((tokens || []).indexOf(a[i].token) >= 0) a.splice(i, 1);
      if (a.length !== n0) D('ThietBiPush');
      return n0 - a.length;
    }
    H['push.dangky'] = function (req, s) {
      var tk = String(req.maMay || '').trim(); if (tk.length < 20 || tk.length > 600) err('Mã thiết bị nhận thông báo không hợp lệ.');
      var a = T('ThietBiPush'), r = a.filter(function (x) { return x.token === tk; })[0];
      if (!r) { r = { id: newId(), token: tk, dangKy: env.now(), tuyChon: '' }; a.push(r); }
      r.username = s.username; r.thietBi = String(req.thietBi || '').slice(0, 80); r.tenMay = String(req.tenMay || '').slice(0, 80); r.capNhat = env.now();
      if (req.tuyChon) r.tuyChon = JSON.stringify(pushTuyChon({ tuyChon: JSON.stringify(req.tuyChon) }));
      // mỗi người tối đa 8 thiết bị; một mã thiết bị chỉ thuộc 1 người (máy dùng chung: người đăng nhập sau nhận thông báo)
      var cua = a.filter(function (x) { return x.username === s.username; }).sort(function (x, y) { return x.capNhat < y.capNhat ? 1 : -1; });
      cua.slice(8).forEach(function (x) { a.splice(a.indexOf(x), 1); });
      D('ThietBiPush');
      return { id: r.id, tuyChon: pushTuyChon(r) };
    };
    H['push.huy'] = function (req, s) {
      var tk = String(req.maMay || ''), a = T('ThietBiPush'), n = 0;
      for (var i = a.length - 1; i >= 0; i--) if ((tk && a[i].token === tk) || (!tk && req.id && a[i].id === req.id && a[i].username === s.username)) { a.splice(i, 1); n++; }
      if (n) D('ThietBiPush');
      return { n: n };
    };
    H['push.tuychon'] = function (req, s) {
      var r = T('ThietBiPush').filter(function (x) { return x.token === String(req.maMay || '') && x.username === s.username; })[0];
      if (!r) err('Thiết bị này chưa bật thông báo đẩy.');
      r.tuyChon = JSON.stringify(pushTuyChon({ tuyChon: JSON.stringify(req.tuyChon || {}) })); r.capNhat = env.now(); D('ThietBiPush');
      return { tuyChon: pushTuyChon(r) };
    };
    H['push.cuatoi'] = function (req, s) {
      return { bat: pushBat(), ds: T('ThietBiPush').filter(function (x) { return x.username === s.username; }).map(function (x) { return { id: x.id, tenMay: x.tenMay, capNhat: x.capNhat, cuaMay: x.token === req.maMay, tuyChon: pushTuyChon(x) }; }) };
    };
    H['push.thu'] = function (req, s) {
      if (!pushBat()) err('Quản trị chưa cấu hình thông báo đẩy (Cài đặt → Thông báo đẩy).');
      if (!T('ThietBiPush').some(function (x) { return x.username === s.username; })) err('Máy này chưa bật thông báo đẩy.');
      bao([s.username], { loai: 'tn', t: 'TAPACO – thông báo thử', b: 'Điện thoại sẽ đổ chuông và hiện thông báo như thế này khi có tin nhắn, bản tin, kiến nghị mới.', tag: 'thu', mo: '' });
      PUSH.forEach(function (q) { q.loai = 'thu'; });            // thông báo thử không bị tắt theo loại
      return { n: T('ThietBiPush').filter(function (x) { return x.username === s.username; }).length };
    };
    H['push.trangthai'] = function (req, s) {
      need(s, 'config');
      var k = env.push && env.push.khoa ? env.push.khoa() : null;
      return { cfg: getCfg().push || {}, khoa: k, soMay: T('ThietBiPush').length, soNguoi: Object.keys(T('ThietBiPush').reduce(function (o, x) { o[x.username] = 1; return o; }, {})).length };
    };
    H['push.caidat'] = function (req, s) {
      need(s, 'config');
      var o = req.push || {}, fb = o.firebase || {}, cu = getCfg().push || {};
      var f = {}; ['apiKey', 'authDomain', 'projectId', 'storageBucket', 'messagingSenderId', 'appId'].forEach(function (k) { if (fb[k]) f[k] = String(fb[k]).trim().slice(0, 200); });
      if (o.bat && (!f.apiKey || !f.projectId || !f.messagingSenderId || !f.appId)) err('Thiếu cấu hình Firebase (apiKey, projectId, messagingSenderId, appId).');
      var vapid = String(o.vapid || '').trim(); if (o.bat && !/^[A-Za-z0-9_-]{80,100}$/.test(vapid)) err('Khóa VAPID (Web Push certificate) không hợp lệ.');
      var khoa = null;
      if (o.taiKhoanDichVu) {
        var sa = typeof o.taiKhoanDichVu === 'string' ? parseJSON(o.taiKhoanDichVu, null) : o.taiKhoanDichVu;
        if (!sa || !sa.client_email || !/BEGIN PRIVATE KEY/.test(sa.private_key || '')) err('Tệp khóa tài khoản dịch vụ (JSON) không hợp lệ.');
        if (f.projectId && sa.project_id && sa.project_id !== f.projectId) err('Khóa tài khoản dịch vụ thuộc dự án "' + sa.project_id + '", khác dự án Firebase "' + f.projectId + '".');
        if (!env.push || !env.push.luuKhoa) err('Máy chủ không lưu được khóa.');
        khoa = env.push.luuKhoa(sa);
      }
      if (o.bat && !(khoa || (env.push && env.push.khoa && env.push.khoa()))) err('Chưa có khóa tài khoản dịch vụ Firebase.');
      setCfg({ push: { bat: !!o.bat, firebase: f, vapid: vapid } }); log(s, 'push.caidat', (o.bat ? 'Bật' : 'Tắt') + ' · ' + (f.projectId || ''));
      return { push: getCfg().push, khoa: khoa || (env.push && env.push.khoa ? env.push.khoa() : null) };
    };
    /* ==================== CHẤM CÔNG GPS + KHUÔN MẶT (v3.0) ====================
       Tham khảo AMIS Chấm công, Base Checkin, Tanca: mỗi công trình / văn phòng khai báo điểm chấm công (vĩ độ, kinh độ, bán kính);
       CBCNV chấm công trên điện thoại khi trong vùng, kèm ảnh khuôn mặt; máy chủ tính lại khoảng cách, so khớp khuôn mặt đã đăng ký, ghi nhận thiết bị.
       Công chỉ ghi nhận khi đúng vị trí và có ảnh khuôn mặt. Ngoài vùng: không ghi công, chỉ lưu lượt thử để quản lý xem. */
    function gpsCfg() { return Object.assign({}, DEFAULT_CFG.gps, getCfg().gps || {}); }
    function gpsDiemDs() { var d = getCfg().diemCC || {}, out = []; Object.keys(d).forEach(function (ma) { (d[ma] || []).forEach(function (p) { out.push(Object.assign({ boPhan: ma }, p)); }); }); return out; }
    function gpsBP(ma) { return (getCfg().boPhan || []).filter(function (b) { return b.ma === ma; })[0]; }
    function gpsDiemCua(nv, ngay) {          // điểm được phép: bộ phận + điểm dùng chung + điểm riêng của người đó + điểm công tác trong ngày
      var bp = (getCfg().boPhan || []).filter(function (b) { return b.ten === nv.phongBan; })[0];
      var ds = gpsDiemDs().filter(function (p) { return p.chung || (bp && p.boPhan === bp.ma) || p.boPhan === 'nv:' + nv.maNV; }).map(function (p) { if (p.boPhan === 'nv:' + nv.maNV) p.caNhan = true; return p; });
      ctCuaNgay(nv.maNV, ngay || env.today()).forEach(function (c) {
        if (c.cheDo !== 'diem') return;
        parseJSON(c.diem, []).forEach(function (p, i) {
          if (p.ngay && p.ngay !== (ngay || env.today())) return;
          ds.push({ id: 'ct:' + c.id + ':' + i, ten: 'Công tác – ' + (p.ten || c.noiDen), lat: +p.lat, lng: +p.lng, banKinh: num(p.banKinh) || num(gpsCfg().banKinhCT) || 300, congTac: c.id, choDuyet: c.trangThai !== 'Đã duyệt', boPhan: bp ? bp.ma : '' });
        });
      });
      return ds;
    }
    /* ---- công tác: CBCNV khai báo lịch công tác (theo điểm cố định từng ngày, hoặc linh hoạt – chấm công tại nơi đang làm việc) ---- */
    var CT_TT = ['Chờ duyệt', 'Đã duyệt', 'Từ chối', 'Đã hủy'];
    function ctCuaNgay(maNV, ngay) { return T('CongTac').filter(function (c) { return c.maNV === maNV && c.tuNgay <= ngay && c.denNgay >= ngay && (c.trangThai === 'Đã duyệt' || c.trangThai === 'Chờ duyệt'); }); }
    function ctPub(c) { var o = Object.assign({}, c); o.diem = parseJSON(c.diem, []); return o; }
    function nguoiGiamSatCC(tru) {                 // người duyệt công tác, kiểm tra chấm công: TCHC + Kế toán (không qua bộ phận)
      return userActive().filter(function (u) { return u.username !== tru && u.role !== 'admin' && u.role !== 'giamdoc' && can(u.role, 'ccQuanLy'); }).map(function (u) { return u.username; });
    }
    function ctQuanLy(s, nv) { try { var pv = gpsQuanLy(s); return pv.tatCa || nv.phongBan === pv.boPhan; } catch (e) { return false; } }
    H['gps.ct.tao'] = function (req, s) {
      var r = req.row || {}, ma = r.maNV || s.maNV, nv = ma && findNV(ma);
      if (!nv) err('Tài khoản chưa gắn hồ sơ CBCNV.');
      var thay = ma !== s.maNV; if (thay && !ctQuanLy(s, nv)) err('Bạn chỉ khai báo công tác cho người trong bộ phận mình quản lý.');
      if (!parseD(r.tuNgay) || !parseD(r.denNgay)) err('Nhập ngày đi và ngày về.');
      if (r.denNgay < r.tuNgay) err('Ngày về phải từ ngày đi trở đi.');
      if (r.denNgay > addDays(r.tuNgay, 92)) err('Một lần khai báo công tác tối đa 3 tháng.');
      if (r.tuNgay < addDays(env.today(), -7) && !thay) err('Chỉ khai báo bổ sung công tác trong 7 ngày gần nhất. Quá hạn: đề nghị quản lý khai báo thay.');
      var cheDo = r.cheDo === 'linhhoat' ? 'linhhoat' : 'diem', g = gpsCfg();
      var noiDen = String(r.noiDen || '').trim().slice(0, 150), mucDich = String(r.mucDich || '').trim().slice(0, 500);
      if (!noiDen) err('Nhập nơi công tác.');
      if (mucDich.length < 5) err('Nhập mục đích / nội dung công tác.');
      var diem = (r.diem || []).slice(0, 31).map(function (p) {
        var lat = +p.lat, lng = +p.lng, bk = Math.round(+p.banKinh || num(g.banKinhCT) || 300);
        if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180 || (!lat && !lng)) err('Chọn vị trí công tác trên bản đồ.');
        if (bk < 50 || bk > 5000) err('Bán kính điểm công tác từ 50 đến 5.000 m.');
        var ng = p.ngay && parseD(p.ngay) ? String(p.ngay).slice(0, 10) : '';
        if (ng && (ng < r.tuNgay || ng > r.denNgay)) err('Ngày của điểm công tác nằm ngoài thời gian công tác.');
        return { ngay: ng, ten: String(p.ten || noiDen).trim().slice(0, 100), lat: Math.round(lat * 1e6) / 1e6, lng: Math.round(lng * 1e6) / 1e6, banKinh: bk, diaChi: String(p.diaChi || '').slice(0, 200) };
      });
      if (cheDo === 'diem' && !diem.length) err('Chọn ít nhất một điểm công tác trên bản đồ (hoặc chọn chế độ "Linh hoạt").');
      var trung = T('CongTac').filter(function (c) { return c.maNV === ma && (c.trangThai === 'Chờ duyệt' || c.trangThai === 'Đã duyệt') && c.tuNgay <= r.denNgay && c.denNgay >= r.tuNgay; })[0];
      if (trung) err('Trùng thời gian với chuyến công tác ' + ngayVN(trung.tuNgay) + ' → ' + ngayVN(trung.denNgay) + ' (' + trung.trangThai + ').');
      var c = { id: newId(), maNV: ma, username: userCuaNV(ma), hoTen: nv.hoTen, boPhan: nv.phongBan || '', tuNgay: r.tuNgay, denNgay: r.denNgay, cheDo: cheDo, diem: JSON.stringify(diem), noiDen: noiDen, mucDich: mucDich,
        phuongTien: String(r.phuongTien || '').slice(0, 60), trangThai: thay ? 'Đã duyệt' : 'Chờ duyệt', nguoiTao: s.hoTen || s.username, nguoiDuyet: thay ? (s.hoTen || s.username) : '', duyetLuc: thay ? env.now() : '', lyDo: '', taoLuc: env.now(), capNhat: env.now() };
      T('CongTac').push(c); D('CongTac'); log(s, 'gps.ct.tao', ma + ' ' + c.tuNgay + '→' + c.denNgay + ' ' + noiDen + (thay ? ' (quản lý khai báo)' : ''));
      if (thay) bao([c.username], { loai: 'cc', t: 'Bạn được cử đi công tác', b: noiDen + ' · ' + ngayVN(c.tuNgay) + (c.denNgay !== c.tuNgay ? ' → ' + ngayVN(c.denNgay) : ''), tag: 'ct_' + c.id, mo: '#v=ccgps' });
      else bao(nguoiGiamSatCC(s.username), { loai: 'duyet', t: 'Đăng ký công tác chờ duyệt', b: nv.hoTen + ': ' + noiDen + ' · ' + ngayVN(c.tuNgay) + (c.denNgay !== c.tuNgay ? ' → ' + ngayVN(c.denNgay) : ''), tag: 'ct_' + c.id, mo: '#v=gpsql' });
      return ctPub(c);
    };
    function ctCapNhatLuot(c, duyet, nguoi) {        // lượt chấm công đang chờ vì công tác chưa duyệt
      var n = 0;
      gpsTu(c.tuNgay).forEach(function (r) {
        if (r.congTac !== c.id || r.trangThai !== 'Chờ duyệt') return;
        var ly = String(r.lyDo || '').split('; ').filter(function (x) { return x && x !== 'Công tác chưa được duyệt'; });
        if (duyet) { r.lyDo = ly.join('; '); if (!ly.length) { r.trangThai = 'Hợp lệ'; r.nguoiDuyet = nguoi; r.duyetLuc = env.now(); } }
        else { r.trangThai = 'Từ chối'; r.lyDo = 'Công tác ' + c.trangThai.toLowerCase() + (c.lyDo ? ': ' + c.lyDo : ''); r.nguoiDuyet = nguoi; r.duyetLuc = env.now(); }
        n++;
      });
      if (n) D('ChamCongGPS');
      return n;
    }
    H['gps.ct.duyet'] = function (req, s) {
      var c = T('CongTac').filter(function (x) { return x.id === req.id; })[0]; if (!c) err('Không tìm thấy đăng ký công tác.');
      var nv = findNV(c.maNV); if (!nv || !ctQuanLy(s, nv)) err('Bạn không có quyền duyệt công tác của người này.');
      if (c.maNV === s.maNV) err('Không tự duyệt công tác của mình.');
      if (c.trangThai !== 'Chờ duyệt') err('Đăng ký đã được xử lý (' + c.trangThai + ').');
      var ok = req.trangThai === 'Đã duyệt', ly = String(req.lyDo || '').trim().slice(0, 300);
      if (!ok && !ly) err('Nhập lý do từ chối.');
      c.trangThai = ok ? 'Đã duyệt' : 'Từ chối'; c.lyDo = ly; c.nguoiDuyet = s.hoTen || s.username; c.duyetLuc = env.now(); c.capNhat = env.now(); D('CongTac');
      var n = ctCapNhatLuot(c, ok, c.nguoiDuyet);
      log(s, 'gps.ct.duyet', c.maNV + ' ' + c.tuNgay + ' → ' + c.trangThai);
      bao([c.username], { loai: 'cc', t: 'Công tác ' + (ok ? 'đã được duyệt' : 'bị từ chối'), b: c.noiDen + ' · ' + ngayVN(c.tuNgay) + (ly ? ' · ' + ly : ''), tag: 'ct_' + c.id, mo: '#v=ccgps' });
      return { ct: ctPub(c), luot: n };
    };
    H['gps.ct.huy'] = function (req, s) {
      var c = T('CongTac').filter(function (x) { return x.id === req.id; })[0]; if (!c) err('Không tìm thấy đăng ký công tác.');
      var nv = findNV(c.maNV), chu = c.maNV === s.maNV, ql = nv && ctQuanLy(s, nv);
      if (!chu && !ql) err('Bạn không có quyền hủy.');
      if (c.trangThai === 'Từ chối' || c.trangThai === 'Đã hủy') err('Đăng ký đã ' + c.trangThai.toLowerCase() + '.');
      if (chu && !ql && c.trangThai === 'Đã duyệt' && c.tuNgay <= env.today()) err('Chuyến công tác đã bắt đầu, đề nghị quản lý điều chỉnh.');
      c.trangThai = 'Đã hủy'; c.lyDo = String(req.lyDo || '').slice(0, 300); c.capNhat = env.now(); D('CongTac');
      ctCapNhatLuot(c, false, s.hoTen || s.username); log(s, 'gps.ct.huy', c.maNV + ' ' + c.tuNgay);
      return ctPub(c);
    };
    H['gps.ct.ds'] = function (req, s) {
      var pv = gpsQuanLy(s), tu = parseD(req.tu) ? req.tu : addDays(env.today(), -31), den = parseD(req.den) ? req.den : addDays(env.today(), 62);
      return T('CongTac').filter(function (c) { return c.tuNgay <= den && c.denNgay >= tu && gpsTrongPham(pv, c); }).sort(function (a, b) { return a.taoLuc < b.taoLuc ? 1 : -1; }).slice(0, 500).map(ctPub);
    };
    function gpsQuanLy(s) {                  // phạm vi xem / duyệt: toàn công ty hoặc bộ phận mình quản lý
      if (can(s.role, 'ccQuanLy')) return { tatCa: true };
      err('Bạn chưa được phân quyền quản lý chấm công GPS.');
    }
    function gpsTrongPham(pv, r) { if (pv.tatCa) return true; var nv = findNV(r.maNV); return !!nv && (nv.phongBan === pv.boPhan || r.boPhan === pv.boPhan); }
    function gpsKM(maNV) { return T('KhuonMat').filter(function (k) { return k.maNV === maNV && k.trangThai === 'Đang dùng'; })[0]; }
    function gpsMau(a) { if (!Array.isArray(a) || a.length !== 128) return null; var o = a.map(function (x) { return Math.round(+x * 10000) / 10000; }); return o.every(isFinite) ? o : null; }
    function gpsPhut(g) { return +g.slice(0, 2) * 60 + +g.slice(3, 5); }
    function gpsAnh(f, s, ten) {
      var d = String((f || {}).data || '');
      if (d.length < 4000) err('Thiếu ảnh khuôn mặt.');
      if (d.length > 1400000) err('Ảnh quá lớn.');
      return env.files.put({ name: ten, mime: 'image/jpeg', data: d, chamCong: env.today().slice(0, 7), maNV: s.maNV, hoTen: s.hoTen });
    }
    H['gps.cuatoi'] = function (req, s) {
      var nv = s.maNV && findNV(s.maNV); if (!nv) err('Tài khoản chưa gắn hồ sơ CBCNV nên không chấm công được.');
      var th = /^\d{4}-\d{2}$/.test(req.thang || '') ? req.thang : env.today().slice(0, 7), km = gpsKM(nv.maNV);
      var hom = env.today(), ctHom = ctCuaNgay(nv.maNV, hom);
      return { cfg: gpsCfg(), diem: gpsDiemCua(nv, hom), phongBan: nv.phongBan, homNay: hom, gioMay: env.now(),
        ctLinhHoat: ctHom.filter(function (c) { return c.cheDo === 'linhhoat'; }).map(ctPub)[0] || null,
        congTac: T('CongTac').filter(function (c) { return c.maNV === nv.maNV && c.denNgay >= addDays(hom, -62); }).sort(function (a, b) { return a.tuNgay < b.tuNgay ? 1 : -1; }).slice(0, 30).map(ctPub),
        km: km ? { dangKyLuc: km.dangKyLuc, thietBi: km.thietBi, soMau: parseJSON(km.mau, []).length } : null,
        ds: gpsTu(th + '-01').filter(function (r) { return r.maNV === nv.maNV && r.ngay.slice(0, 7) === th; }).map(function (r) { var o = Object.assign({}, r); delete o.mauKM; return o; }).sort(function (a, b) { return a.thoiGian < b.thoiGian ? 1 : -1; }) };
    };
    H['gps.dangky'] = function (req, s) {           // đăng ký khuôn mặt (một lần; đăng ký lại cần TCHC xóa mẫu cũ)
      var nv = s.maNV && findNV(s.maNV); if (!nv) err('Tài khoản chưa gắn hồ sơ CBCNV.');
      if (gpsKM(nv.maNV)) err('Bạn đã đăng ký khuôn mặt. Muốn đăng ký lại, đề nghị Tổ chức – Hành chính xóa mẫu cũ.');
      var mau = (req.mau || []).map(gpsMau).filter(Boolean);
      if (mau.length < 2) err('Cần ít nhất 2 ảnh có khuôn mặt rõ để đăng ký.');
      var id = gpsAnh(req.anh, s, 'KhuonMat_' + nv.maNV + '_' + env.today() + '.jpg');
      T('KhuonMat').push({ maNV: nv.maNV, username: s.username, hoTen: nv.hoTen, mau: JSON.stringify(mau.slice(0, 5)), anh: id, thietBi: String(req.thietBi || '').slice(0, 80), dangKyLuc: env.now(), trangThai: 'Đang dùng', xoaBoi: '', xoaLuc: '' });
      D('KhuonMat'); log(s, 'gps.dangky', nv.maNV + ' (' + mau.length + ' mẫu)');
      return { ok: true };
    };
    H['gps.cham'] = function (req, s) {
      var nv = s.maNV && findNV(s.maNV); if (!nv) err('Tài khoản chưa gắn hồ sơ CBCNV nên không chấm công được.');
      if (nv.trangThai === 'Đã nghỉ việc') err('Hồ sơ đã nghỉ việc.');
      var c = gpsCfg(); if (!c.bat) err('Công ty chưa bật chấm công GPS.');
      var hom = env.today(), GPS = gpsTu(hom), luot = GPS.filter(function (r) { return r.maNV === nv.maNV && r.ngay === hom; });
      if (luot.length >= 20) err('Đã quá số lần chấm công trong ngày.');
      var lat = +req.lat, lng = +req.lng, acc = +req.doChinhXac;
      if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) err('Chưa có vị trí. Hãy bật định vị và hiệu chuẩn lại.');
      if (!isFinite(acc) || acc > num(c.doChinhXac)) err('Định vị chưa đủ chính xác (±' + Math.round(acc || 0) + ' m, cần ≤ ' + c.doChinhXac + ' m). Hãy ra chỗ thoáng, bật chế độ vị trí chính xác cao rồi hiệu chuẩn lại.');
      var ds = gpsDiemCua(nv, hom), lh = ctCuaNgay(nv.maNV, hom).filter(function (x) { return x.cheDo === 'linhhoat'; })[0];
      if (!ds.length && !lh) err('Bộ phận "' + (nv.phongBan || '') + '" chưa khai báo điểm chấm công. Báo Tổ chức – Hành chính hoặc Chỉ huy trưởng, hoặc khai báo công tác.');
      var gan = ds.map(function (p) { return { p: p, d: Math.round(khoangCachM(lat, lng, +p.lat, +p.lng)) }; }).sort(function (a, b) { return (a.d - num(a.p.banKinh)) - (b.d - num(b.p.banKinh)); })[0];
      if ((!gan || gan.d > num(gan.p.banKinh)) && lh) {                 // công tác linh hoạt: chấm công tại nơi đang làm việc
        var bpNV = (getCfg().boPhan || []).filter(function (b) { return b.ten === nv.phongBan; })[0];
        gan = { p: { id: 'ct:' + lh.id, ten: 'Công tác – ' + lh.noiDen, banKinh: 0, congTac: lh.id, choDuyet: lh.trangThai !== 'Đã duyệt', boPhan: bpNV ? bpNV.ma : '', linhHoat: true }, d: 0 };
      }
      var trong = gan.p.linhHoat || gan.d <= num(gan.p.banKinh), loai = req.loai === 'ra' ? 'ra' : 'vao', gio = env.now().slice(11, 16), bp = gpsBP(gan.p.boPhan);
      var row = { id: newId(), maNV: nv.maNV, username: s.username, hoTen: nv.hoTen, ngay: hom, gio: gio, loai: loai, diemId: gan.p.id, tenDiem: gan.p.ten, boPhan: bp ? bp.ten : (gan.p.chung ? 'Điểm chung' : ''),
        lat: Math.round(lat * 1e6) / 1e6, lng: Math.round(lng * 1e6) / 1e6, doChinhXac: Math.round(acc), khoangCach: gan.d, banKinh: num(gan.p.banKinh), anh: '', khuonMat: '', diemKM: '', thietBi: String(req.thietBi || '').slice(0, 80),
        trangThai: '', lyDo: '', nguoiDuyet: '', duyetLuc: '', ghiChu: String(req.ghiChu || '').slice(0, 300), thoiGian: env.now(), tre: 0, congTac: gan.p.congTac || '' };
      if (gan.p.caNhan) row.boPhan = nv.phongBan || '';
      if (!trong) {
        row.trangThai = 'Sai vị trí'; row.lyDo = 'Ngoài vùng: cách "' + gan.p.ten + '" ' + gan.d + ' m (bán kính ' + gan.p.banKinh + ' m)';
        if (c.nhanAnhNgoaiVung && req.anh && String(req.anh.data || '').length > 4000) { try { row.anh = gpsAnh(req.anh, s, 'SaiViTri_' + nv.maNV + '_' + env.now().replace(/[: ]/g, '-') + '.jpg'); } catch (e) { } }
        GPS.push(row); D('ChamCongGPS');
        return { ghiNhan: false, trangThai: row.trangThai, lyDo: row.lyDo, khoangCach: gan.d, banKinh: row.banKinh, tenDiem: gan.p.ten, coTheCT: true };
      }
      if (!req.coMat) err('Ảnh chưa thấy rõ khuôn mặt. Nhìn thẳng vào camera, đủ sáng rồi chụp lại.');
      row.anh = gpsAnh(req.anh, s, 'ChamCong_' + nv.maNV + '_' + env.now().replace(/[: ]/g, '-') + '.jpg');
      var ly = [], km = gpsKM(nv.maNV), mau = gpsMau(req.mau);
      if (!km) { if (c.batBuocKM) err('Bạn chưa đăng ký khuôn mặt. Bấm "Đăng ký khuôn mặt" trước khi chấm công.'); row.khuonMat = 'Chưa đăng ký'; ly.push('Chưa đăng ký khuôn mặt'); }
      else if (!mau) { row.khuonMat = 'Chưa so khớp'; ly.push('Máy không so khớp được khuôn mặt (thiết bị không hỗ trợ / chưa tải mô hình)'); }
      else {
        var dmin = Math.min.apply(null, parseJSON(km.mau, []).map(function (m) { return kmKhoangCach(m, mau); }));
        row.diemKM = Math.round(dmin * 1000) / 1000;
        if (dmin <= num(c.nguongKM)) row.khuonMat = 'Khớp'; else { row.khuonMat = 'Không khớp'; ly.push('Khuôn mặt không khớp mẫu đăng ký (độ lệch ' + row.diemKM + ')'); }
      }
      if (c.khoaThietBi && km && km.thietBi && row.thietBi && row.thietBi !== km.thietBi) ly.push('Chấm công trên thiết bị khác thiết bị đã đăng ký');
      if (gan.p.choDuyet) ly.push('Công tác chưa được duyệt');
      if (req.nghiGia) ly.push('Tín hiệu định vị bất thường: ' + String(req.nghiGia).slice(0, 120));
      if (acc < 1) ly.push('Độ chính xác định vị bất thường (±0 m)');
      var gV = gan.p.gioVao || c.gioVao, gR = gan.p.gioRa || c.gioRa;
      if (!row.congTac) {                                                    // công tác: không tính đi trễ / về sớm
        if (loai === 'vao' && gpsPhut(gio) > gpsPhut(gV) + num(c.phutTre)) row.tre = gpsPhut(gio) - gpsPhut(gV);
        if (loai === 'ra' && gpsPhut(gio) < gpsPhut(gR)) row.tre = -(gpsPhut(gR) - gpsPhut(gio));         // số âm: về sớm
      }
      row.trangThai = ly.length ? 'Chờ duyệt' : 'Hợp lệ'; row.lyDo = ly.join('; ');
      GPS.push(row); D('ChamCongGPS');
      return { ghiNhan: true, trangThai: row.trangThai, lyDo: row.lyDo, khoangCach: gan.d, banKinh: row.banKinh, tenDiem: gan.p.ten, gio: gio, loai: loai, tre: row.tre, khuonMat: row.khuonMat, congTac: row.congTac, linhHoat: !!gan.p.linhHoat };
    };
    H['gps.ds'] = function (req, s) {
      var pv = gpsQuanLy(s), tu = parseD(req.tu) ? req.tu : env.today(), den = parseD(req.den) ? req.den : tu;
      return { ds: gpsTu(tu).filter(function (r) { return r.ngay >= tu && r.ngay <= den && gpsTrongPham(pv, r) && (!req.trangThai || r.trangThai === req.trangThai); }).sort(function (a, b) { return a.thoiGian < b.thoiGian ? 1 : -1; }).slice(0, 2000),
        km: T('KhuonMat').filter(function (k) { return k.trangThai === 'Đang dùng' && gpsTrongPham(pv, k); }).map(function (k) { return { maNV: k.maNV, hoTen: k.hoTen, dangKyLuc: k.dangKyLuc, thietBi: k.thietBi, anh: k.anh, soMau: parseJSON(k.mau, []).length }; }),
        pham: pv.tatCa ? 'Toàn công ty' : pv.boPhan,
        ct: T('CongTac').filter(function (c) { return c.tuNgay <= den && c.denNgay >= tu && (c.trangThai === 'Đã duyệt' || c.trangThai === 'Chờ duyệt') && gpsTrongPham(pv, c); }).map(ctPub),
        cb: (function () { var m = {}; T('CanhBaoCC').forEach(function (x) { if (CCKT_MO.indexOf(x.trangThai) < 0 || x.ngay < tu || x.ngay > den) return; parseJSON(x.luotIds, []).forEach(function (i) { var o = m[i] = m[i] || { mucDo: 'vang', ds: [] }; if (x.mucDo === 'do') o.mucDo = 'do'; o.ds.push(CCKT_MA[x.ma][1]); }); }); return m; })(),
        soCB: T('CanhBaoCC').filter(function (x) { return (x.trangThai === 'Đã nhắc') && (pv.tatCa || x.boPhan === pv.boPhan); }).length };
    };
    H['gps.anh'] = function (req, s) {
      var id = String(req.id || ''), r = T('ChamCongGPS').filter(function (x) { return x.anh === id; })[0] || T('KhuonMat').filter(function (x) { return x.anh === id; })[0];
      if (!r) err('Không tìm thấy ảnh.');
      if (r.maNV !== s.maNV) { var pv = gpsQuanLy(s); if (!gpsTrongPham(pv, r)) err('Không có quyền xem ảnh này.'); }
      var d = env.files.get(id); if (!d) err('Ảnh không còn trên máy chủ.');
      return { data: d.data, mime: 'image/jpeg' };
    };
    H['gps.duyet'] = function (req, s) {
      var pv = gpsQuanLy(s), tt = req.trangThai === 'Hợp lệ' ? 'Hợp lệ' : 'Từ chối', ly = String(req.lyDo || '').trim().slice(0, 300), n = 0;
      if (tt === 'Từ chối' && !ly) err('Nhập lý do từ chối.');
      (parseD(req.tu) ? gpsTu(req.tu) : T('ChamCongGPS')).forEach(function (r) {   // v3.14: giao diện gửi ngày sớm nhất của các lượt chọn → chỉ đọc phần cuối bảng
        if ((req.ids || []).indexOf(r.id) < 0 || !gpsTrongPham(pv, r)) return;
        if (r.trangThai === 'Sai vị trí') return;                                // ngoài vùng: không duyệt thành công
        if (r.maNV === s.maNV) err('Không tự duyệt chấm công của mình.');
        r.trangThai = tt; r.nguoiDuyet = s.hoTen || s.username; r.duyetLuc = env.now(); if (ly) r.ghiChu = (r.ghiChu ? r.ghiChu + ' | ' : '') + 'Duyệt: ' + ly; n++;
      });
      if (!n) err('Không có lượt chấm công nào được duyệt.');
      D('ChamCongGPS'); log(s, 'gps.duyet', tt + ' ' + n + ' lượt');
      return { n: n };
    };
    H['gps.xoakm'] = function (req, s) {
      need(s, 'edit_nv');
      var k = gpsKM(String(req.maNV || '')); if (!k) err('Người này chưa đăng ký khuôn mặt.');
      k.trangThai = 'Đã xóa'; k.xoaBoi = s.username; k.xoaLuc = env.now(); D('KhuonMat'); log(s, 'gps.xoakm', k.maNV);
      return true;
    };
    H['gps.diem'] = function (req, s) {             // khai báo điểm chấm công cho 1 bộ phận / công trình
      var ma = String(req.ma || ''), nvR = /^nv:/.test(ma) ? findNV(ma.slice(3)) : null;
      var bp = ma === '_chung' ? { ma: '_chung', ten: 'Điểm dùng chung' } : nvR ? { ma: ma, ten: 'Điểm riêng của ' + nvR.hoTen, boPhanNV: nvR.phongBan } : gpsBP(ma); if (!bp) err('Không tìm thấy bộ phận / nhân viên.');
      var duoc = can(s.role, 'edit_nv') || s.role === 'giamdoc';
      if (!duoc) err('Chỉ TCHC, Giám đốc, Quản trị khai báo được điểm chấm công.');
      var ds = (req.diem || []).map(function (p) {
        var lat = +p.lat, lng = +p.lng, r = Math.round(+p.banKinh), ten = String(p.ten || '').trim().slice(0, 80);
        if (!ten) err('Nhập tên điểm chấm công.');
        if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180 || (!lat && !lng)) err('Tọa độ "' + ten + '" không hợp lệ.');
        if (!(r >= 20 && r <= 5000)) err('Bán kính "' + ten + '" phải từ 20 đến 5.000 m.');
        var o = { id: String(p.id || newId()).slice(0, 20), ten: ten, lat: Math.round(lat * 1e6) / 1e6, lng: Math.round(lng * 1e6) / 1e6, banKinh: r, chung: ma === '_chung' };
        if (p.gioVao && /^\d\d:\d\d$/.test(p.gioVao)) o.gioVao = p.gioVao; if (p.gioRa && /^\d\d:\d\d$/.test(p.gioRa)) o.gioRa = p.gioRa;
        return o;
      });
      var all = Object.assign({}, getCfg().diemCC || {}); if (ds.length) all[ma] = ds; else delete all[ma];
      setCfg({ diemCC: all }); log(s, 'gps.diem', bp.ten + ': ' + ds.length + ' điểm');
      return { diemCC: all };
    };
    H['gps.caidat'] = function (req, s) {
      if (!can(s.role, 'edit_nv') && s.role !== 'giamdoc') err('Chỉ TCHC, Giám đốc, Quản trị.');
      var c = gpsCfg(), o = req.gps || {};
      ['doChinhXac', 'hieuChuanGiay', 'phutTre', 'gioDuCong', 'gioLamDuCong'].forEach(function (k) { if (o[k] != null && isFinite(+o[k]) && +o[k] >= 0) c[k] = +o[k]; });
      if (o.nguongKM != null && +o.nguongKM >= 0.3 && +o.nguongKM <= 0.7) c.nguongKM = +o.nguongKM;
      ['gioVao', 'gioRa', 'nghiTruaTu', 'nghiTruaDen'].forEach(function (k) { if (/^\d\d:\d\d$/.test(o[k] || '')) c[k] = o[k]; });
      if (o.nghiTruaTu === '' && o.nghiTruaDen === '') { c.nghiTruaTu = ''; c.nghiTruaDen = ''; }
      if (c.nghiTruaTu && c.nghiTruaDen && c.nghiTruaDen <= c.nghiTruaTu) err('Giờ nghỉ trưa: giờ kết thúc phải sau giờ bắt đầu.');
      if (!(num(c.gioLamDuCong) > 0 && num(c.gioLamDuCong) <= 12)) err('Số giờ làm đủ 1 công từ 1 đến 12.');
      ['bat', 'khoaThietBi', 'batBuocKM', 'nhanAnhNgoaiVung'].forEach(function (k) { if (o[k] != null) c[k] = !!o[k]; });
      if (o.banKinhCT != null) { if (!(+o.banKinhCT >= 50 && +o.banKinhCT <= 5000)) err('Bán kính điểm công tác từ 50 đến 5.000 m.'); c.banKinhCT = Math.round(+o.banKinhCT); }
      if (o.banDo != null) { var bd = String(o.banDo || '').trim(); if (bd && !/^https:\/\/[^\s]+\{z\}[^\s]*\{x\}[^\s]*\{y\}/.test(bd)) err('Địa chỉ lớp bản đồ phải dạng https://…/{z}/{x}/{y}…'); c.banDo = bd.slice(0, 300); }
      if (c.doChinhXac < 10 || c.doChinhXac > 300) err('Độ chính xác tối đa từ 10 đến 300 m.');
      setCfg({ gps: c }); log(s, 'gps.caidat', JSON.stringify(c));
      return { gps: c };
    };
    /* công theo ngày từ chấm công GPS hợp lệ: có vào và ra đủ giờ → X, thiếu ra hoặc chưa đủ giờ → ½ */
    H['gps.tonghop'] = function (req, s) {
      if (!can(s.role, 'view_all')) err('Bạn không có quyền.');
      var ky = String(req.ky || ''); if (!/^\d{4}-\d{2}$/.test(ky)) err('Kỳ không hợp lệ.');
      var c = gpsCfg(), g = {}, out = {}, ghi = [];
      gpsTu(ky + '-01').forEach(function (r) { if (r.trangThai !== 'Hợp lệ' || r.ngay.slice(0, 7) !== ky) return; var k = r.maNV + '|' + r.ngay; (g[k] = g[k] || []).push(r); });
      Object.keys(g).forEach(function (k) {
        var p = k.split('|'), ds = g[k].sort(function (a, b) { return a.gio < b.gio ? -1 : 1; }), v = ds.filter(function (r) { return r.loai === 'vao'; })[0], ra = ds.filter(function (r) { return r.loai === 'ra'; }).pop();
        var dau = (v || ds[0]).gio, cuoi = ra ? ra.gio : '', gio = cuoi ? gioLamThuc(dau, cuoi, c) : 0;
        var kh = cuoi && gio >= nguongDuCong(c) - 1e-9 ? 'X' : 'X2';
        if (ds.some(function (r) { return r.congTac; })) kh = 'CT';             // ngày công tác (đã duyệt, có chấm công) → ký hiệu CT
        (out[p[0]] = out[p[0]] || {})[+p[1].slice(8)] = kh;
        if (kh === 'X2') ghi.push({ maNV: p[0], ngay: p[1], lyDo: !cuoi ? 'thiếu chấm công ra' : (!v ? 'thiếu chấm công vào' : 'làm ' + Math.round(gio * 10) / 10 + ' giờ') });
      });
      T('CongTac').forEach(function (ct) {                                    // ngày công tác đã duyệt nhưng không có lượt chấm công hợp lệ
        if (ct.trangThai !== 'Đã duyệt') return;
        for (var d = ct.tuNgay; d <= ct.denNgay && d <= env.today(); d = addDays(d, 1)) if (d.slice(0, 7) === ky && !g[ct.maNV + '|' + d]) ghi.push({ maNV: ct.maNV, ngay: d, lyDo: 'công tác ' + ct.noiDen + ' chưa chấm công' });
      });
      return { cong: out, ghiChu: ghi };
    };
    /* ==================== KIỂM TRA CHẤM CÔNG GPS (AI) – v3.2 ====================
       Giám đốc, TCHC, Kế toán, Quản trị chạy rà soát một khoảng ngày: bộ quy tắc của hệ thống gắn cờ ĐỎ (sai phạm cần kiểm tra)
       hoặc VÀNG (cần lưu ý) theo 3 nhóm: sai phạm vi, sai người, sai thời gian. AI chỉ đọc số liệu đã tổng hợp để viết nhận xét, đề xuất.
       Cảnh báo được nhắc tới TCHC / Kế toán (v3.7.1: không qua trưởng bộ phận / CHT); TCHC / Kế toán ghi kết quả, đóng hoặc mở lại. */
    var CCKT_ROLES = { indexOf: function (r) { return can(r, 'ccQuanLy') ? 0 : -1; } };   // v3.8: theo bảng phân quyền
    var CCKT_NHOM = { pv: 'Sai phạm vi', ng: 'Sai người', tg: 'Sai thời gian' };
    var CCKT_MA = {
      PV_SAI: ['pv', 'Chấm ngoài vùng'], PV_BIEN: ['pv', 'Sát biên vùng chấm công'], PV_ACC: ['pv', 'Sai số định vị lớn'], PV_GIA: ['pv', 'Sai số định vị bất thường (nghi giả lập)'],
      PV_TOC: ['pv', 'Di chuyển bất thường giữa 2 lượt'],
      NG_KM: ['ng', 'Khuôn mặt không khớp'], NG_CSK: ['ng', 'Chưa so khớp khuôn mặt'], NG_NGUONG: ['ng', 'Khuôn mặt sát ngưỡng'], NG_TB: ['ng', 'Chấm trên thiết bị khác'],
      NG_TBCHUNG: ['ng', 'Một thiết bị chấm cho nhiều người'], NG_GIA: ['ng', 'Tín hiệu định vị bất thường'],
      TG_TRE: ['tg', 'Đi trễ'], TG_SOM: ['tg', 'Về sớm'], TG_THIEURA: ['tg', 'Thiếu chấm công ra'], TG_THIEUVAO: ['tg', 'Thiếu chấm công vào'], TG_NGAN: ['tg', 'Làm chưa đủ giờ'],
      TG_NGHI: ['tg', 'Chấm công trong ngày nghỉ có đơn'], TG_NGHILE: ['tg', 'Chấm công ngày nghỉ tuần / lễ'], TG_NHIEU: ['tg', 'Quá nhiều lượt trong ngày'], TG_CHO: ['tg', 'Chờ duyệt quá hạn'], TG_VANG: ['tg', 'Không chấm công ngày làm việc']
    };
    var CCKT_MO = ['Mới', 'Đã nhắc', 'Đã giải trình'];
    function soThuc(v) { var x = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isFinite(x) ? x : 0; }   // số thập phân (0.236) – num() coi dấu chấm là phân cách nghìn
    function ccktQuyen(s) { if (CCKT_ROLES.indexOf(s.role) < 0) err('Bạn chưa được phân quyền kiểm tra chấm công.'); }
    function ccktQL(s, cb) { return CCKT_ROLES.indexOf(s.role) >= 0; }   // v3.7.1: không qua trưởng bộ phận / CHT
    function qlBoPhan(ten, tru) {
      var ql = userActive().filter(function (u) { return u.username !== tru && ROLE_QL.indexOf(u.role) >= 0 && u.boPhan === ten; }).map(function (u) { return u.username; });
      return ql.length ? ql : userActive().filter(function (u) { return u.username !== tru && u.role === 'tchc'; }).map(function (u) { return u.username; });
    }
    function ccktQuet(tu, den, boPhan) {
      var cfg = getCfg(), c = gpsCfg(), hom = env.today(), out = [];
      var nvs = T('NhanVien').filter(function (n) { return n.trangThai !== 'Đã nghỉ việc' && (!boPhan || n.phongBan === boPhan); }), nvM = {};
      nvs.forEach(function (n) { nvM[n.maNV] = n; });
      var R = gpsTu(tu).filter(function (r) { return r.ngay >= tu && r.ngay <= den && nvM[r.maNV]; });
      var nghi = {}; T('NghiPhep').forEach(function (o) { if (o.trangThai !== 'Đã duyệt' || !nvM[o.maNV] || o.buoi === 'Sáng' || o.buoi === 'Chiều') return; for (var d = o.tuNgay; d <= o.denNgay; d = addDays(d, 1)) nghi[o.maNV + '|' + d] = o.loai; });
      var ct = {}; T('CongTac').forEach(function (o) { if (o.trangThai !== 'Đã duyệt' || !nvM[o.maNV]) return; for (var d = o.tuNgay; d <= o.denNgay; d = addDays(d, 1)) ct[o.maNV + '|' + d] = 1; });
      var km = {}; T('KhuonMat').forEach(function (k) { if (k.trangThai === 'Đang dùng' && nvM[k.maNV]) km[k.maNV] = String(k.dangKyLuc || '').slice(0, 10); });
      function co(r, ma, muc, moTa, luot, ngay, them) {
        var n = nvM[r.maNV] || {};
        out.push({ khoa: r.maNV + '|' + (ngay || r.ngay) + '|' + ma + (them || ''), maNV: r.maNV, hoTen: n.hoTen || r.hoTen, boPhan: n.phongBan || r.boPhan, ngay: ngay || r.ngay, nhom: CCKT_MA[ma][0], ma: ma, mucDo: muc, moTa: moTa, luotIds: luot || [r.id] });
      }
      var g = {}; R.forEach(function (r) { (g[r.maNV + '|' + r.ngay] = g[r.maNV + '|' + r.ngay] || []).push(r); });
      // ---- từng lượt
      R.forEach(function (r) {
        var hl = r.trangThai === 'Hợp lệ' || r.trangThai === 'Chờ duyệt', bk = num(r.banKinh), kc = num(r.khoangCach), acc = num(r.doChinhXac), ly = String(r.lyDo || '');
        if (hl && bk > 0 && kc >= 0.9 * bk) co(r, 'PV_BIEN', 'vang', 'Lượt ' + (r.loai === 'vao' ? 'vào' : 'ra') + ' ' + r.gio + ' cách tâm ' + kc + ' m / bán kính ' + bk + ' m (sát biên).');
        if (r.trangThai !== 'Sai vị trí' && acc > 0.8 * num(c.doChinhXac)) co(r, 'PV_ACC', 'vang', 'Lượt ' + r.gio + ' sai số định vị ±' + acc + ' m (ngưỡng ' + c.doChinhXac + ' m).');
        if (r.trangThai !== 'Sai vị trí' && acc > 0 && acc <= 2) co(r, 'PV_GIA', 'do', 'Lượt ' + r.gio + ' sai số chỉ ±' + acc + ' m – bất thường với điện thoại, nghi ứng dụng giả lập vị trí.');
        if (r.khuonMat === 'Không khớp') co(r, 'NG_KM', 'do', 'Lượt ' + r.gio + ' khuôn mặt không khớp mẫu (độ lệch ' + r.diemKM + ')' + (r.trangThai === 'Hợp lệ' ? ' nhưng ĐÃ ĐƯỢC DUYỆT hợp lệ bởi ' + (r.nguoiDuyet || '?') : r.trangThai === 'Chờ duyệt' ? ', đang chờ duyệt' : '') + ' – nghi chấm hộ.');
        if (r.trangThai === 'Hợp lệ' && (r.khuonMat === 'Chưa so khớp' || r.khuonMat === 'Chưa đăng ký')) co(r, 'NG_CSK', 'vang', 'Lượt ' + r.gio + ' được tính công nhưng khuôn mặt "' + r.khuonMat + '" – cần xem ảnh.');
        if (r.khuonMat === 'Khớp' && r.diemKM !== '' && soThuc(r.diemKM) >= soThuc(c.nguongKM) - 0.08) co(r, 'NG_NGUONG', 'vang', 'Lượt ' + r.gio + ' độ lệch khuôn mặt ' + r.diemKM + ' sát ngưỡng ' + c.nguongKM + '.');
        if (/thiết bị khác/.test(ly)) co(r, 'NG_TB', 'vang', 'Lượt ' + r.gio + ' chấm trên thiết bị khác thiết bị đã đăng ký.');
        if (/định vị bất thường/i.test(ly)) co(r, 'NG_GIA', 'do', 'Lượt ' + r.gio + ': ' + ly.split('; ').filter(function (x) { return /định vị bất thường/i.test(x); })[0]);
        if (r.loai === 'vao' && num(r.tre) > 0 && !r.congTac) co(r, 'TG_TRE', num(r.tre) > 60 ? 'do' : 'vang', 'Vào lúc ' + r.gio + ', trễ ' + r.tre + ' phút.');
        if (r.loai === 'ra' && num(r.tre) < 0 && hl && !r.congTac) co(r, 'TG_SOM', num(r.tre) < -60 ? 'do' : 'vang', 'Ra lúc ' + r.gio + ', về sớm ' + (-num(r.tre)) + ' phút.');
        if (r.trangThai === 'Chờ duyệt' && r.ngay < addDays(hom, -2)) co(r, 'TG_CHO', 'vang', 'Lượt ' + r.gio + ' chờ duyệt từ ' + r.ngay.slice(8) + '/' + r.ngay.slice(5, 7) + ' chưa được xử lý (' + ly + ').');
      });
      // ---- theo người – ngày
      Object.keys(g).forEach(function (k) {
        var ds = g[k].sort(function (a, b) { return a.gio < b.gio ? -1 : 1; }), p = k.split('|'), ngay = p[1], r0 = ds[0];
        var hl = ds.filter(function (r) { return r.trangThai === 'Hợp lệ' || r.trangThai === 'Chờ duyệt'; }), sai = ds.filter(function (r) { return r.trangThai === 'Sai vị trí'; });
        if (sai.length && !hl.length) co(r0, 'PV_SAI', 'do', 'Chỉ có ' + sai.length + ' lượt ngoài vùng (gần nhất cách ' + sai[sai.length - 1].khoangCach + ' m), không có lượt hợp lệ.', sai.map(function (r) { return r.id; }));
        else if (sai.length >= 3) co(r0, 'PV_SAI', 'vang', sai.length + ' lượt thử ngoài vùng trước khi chấm được.', sai.map(function (r) { return r.id; }));
        if (ds.length >= 5) co(r0, 'TG_NHIEU', 'vang', ds.length + ' lượt chấm công trong ngày.', ds.map(function (r) { return r.id; }));
        if (!hl.length) return;
        var v = hl.filter(function (r) { return r.loai === 'vao'; })[0], ra = hl.filter(function (r) { return r.loai === 'ra'; }).pop();
        if (v && !ra && ngay < hom) co(r0, 'TG_THIEURA', 'do', 'Có chấm vào ' + v.gio + ' nhưng không chấm ra – ngày chỉ được ½ công.', [v.id]);
        if (!v && ra) co(r0, 'TG_THIEUVAO', 'do', 'Có chấm ra ' + ra.gio + ' nhưng không chấm vào.', [ra.id]);
        if (v && ra) { var h = gioLamThuc(v.gio, ra.gio, c); if (h < nguongDuCong(c) - 1e-9 && !v.congTac) co(r0, 'TG_NGAN', 'vang', 'Vào ' + v.gio + ' – ra ' + ra.gio + ': làm ' + Math.round(h * 10) / 10 + ' giờ (đã trừ nghỉ trưa), chưa đủ ' + (num(c.gioLamDuCong) || 8) + ' giờ – tính nửa công, không có ăn ca.', [v.id, ra.id]); }
        if (nghi[k]) co(r0, 'TG_NGHI', 'do', 'Ngày có đơn "' + nghi[k] + '" đã duyệt nhưng vẫn có lượt chấm công hợp lệ – công và ngày nghỉ bị tính trùng.', hl.map(function (r) { return r.id; }));
        if (laNgayNghiLe(ngay, cfg) && !ct[k]) co(r0, 'TG_NGHILE', 'vang', 'Chấm công ngày ' + (new Date(ngay + 'T00:00').getDay() === 0 ? 'Chủ nhật' : 'lễ') + ' – cần xác nhận làm thêm.', hl.map(function (r) { return r.id; }));
      });
      // ---- di chuyển bất thường, thiết bị dùng chung
      var theoNV = {}; R.forEach(function (r) { if (r.trangThai !== 'Sai vị trí') (theoNV[r.maNV] = theoNV[r.maNV] || []).push(r); });
      Object.keys(theoNV).forEach(function (m) {
        var a = theoNV[m].sort(function (x, y) { return x.thoiGian < y.thoiGian ? -1 : 1; });
        for (var i = 1; i < a.length; i++) {
          var x = a[i - 1], y = a[i], d = khoangCachM(+x.lat, +x.lng, +y.lat, +y.lng) / 1000, gio = (new Date(y.thoiGian.replace(' ', 'T')) - new Date(x.thoiGian.replace(' ', 'T'))) / 3600000;
          if (d > 2 && gio > 0 && gio < 24 && d / gio > 150) co(y, 'PV_TOC', 'do', 'Lượt ' + x.gio + ' (' + x.tenDiem + ') và ' + y.gio + ' (' + y.tenDiem + ') cách ' + Math.round(d) + ' km trong ' + Math.round(gio * 60) + ' phút (' + Math.round(d / gio) + ' km/h) – không thể di chuyển kịp.', [x.id, y.id]);
        }
      });
      var tb = {}; R.forEach(function (r) { if (r.thietBi && r.trangThai !== 'Sai vị trí') (tb[r.thietBi] = tb[r.thietBi] || {})[r.maNV] = (tb[r.thietBi][r.maNV] || []).concat([r]); });
      Object.keys(tb).forEach(function (t) {
        var ms = Object.keys(tb[t]); if (ms.length < 2) return;
        var ten = ms.map(function (m) { return ((nvM[m] || {}).hoTen || m) + ' ' + tb[t][m].length + ' lượt'; });
        ms.forEach(function (m) { var a = tb[t][m].sort(function (x, y) { return x.thoiGian < y.thoiGian ? -1 : 1; });
          co(a[0], 'NG_TBCHUNG', 'do', 'Cùng một điện thoại (' + t.slice(0, 10) + '…) chấm công cho ' + ms.length + ' người: ' + ten.join(', ') + ' – nghi chấm hộ.', a.slice(0, 3).map(function (r) { return r.id; }), a[0].ngay, '|' + t.slice(0, 12)); });
      });
      // ---- không chấm công (người đã dùng chấm công GPS)
      Object.keys(km).forEach(function (m) {
        for (var d = tu > km[m] ? tu : km[m]; d <= den && d < hom; d = addDays(d, 1)) {
          var k = m + '|' + d; if (laNgayNghiLe(d, cfg) || nghi[k] || ct[k]) continue;
          var hl = (g[k] || []).filter(function (r) { return r.trangThai === 'Hợp lệ' || r.trangThai === 'Chờ duyệt'; });
          if (!hl.length && !(g[k] || []).length) co({ maNV: m, ngay: d, id: '' }, 'TG_VANG', 'vang', 'Ngày làm việc không có lượt chấm công nào, không có đơn nghỉ / công tác.', [], d);
        }
      });
      return { ds: out.slice(0, 4000), soLuot: R.length, soNguoi: Object.keys(theoNV).length };
    }
    function ccktThongKe(ds) {
      var bp = {}, ma = {}, nguoi = {}, n = { do: 0, vang: 0, pv: 0, ng: 0, tg: 0 };
      ds.forEach(function (x) {
        n[x.mucDo]++; n[x.nhom]++;
        var b = bp[x.boPhan] = bp[x.boPhan] || { boPhan: x.boPhan, do: 0, vang: 0, pv: 0, ng: 0, tg: 0, nguoi: {} }; b[x.mucDo]++; b[x.nhom]++; b.nguoi[x.maNV] = 1;
        var m = ma[x.ma] = ma[x.ma] || { ma: x.ma, ten: CCKT_MA[x.ma][1], nhom: CCKT_MA[x.ma][0], do: 0, vang: 0 }; m[x.mucDo]++;
        var p = nguoi[x.maNV] = nguoi[x.maNV] || { maNV: x.maNV, hoTen: x.hoTen, boPhan: x.boPhan, do: 0, vang: 0, ma: {} }; p[x.mucDo]++; var tl = CCKT_MA[x.ma][1]; p.ma[tl] = (p.ma[tl] || 0) + 1;
      });
      return { n: n, boPhan: Object.keys(bp).map(function (k) { var b = bp[k]; b.soNguoi = Object.keys(b.nguoi).length; delete b.nguoi; return b; }).sort(function (a, b) { return b.do - a.do || b.vang - a.vang; }),
        ma: Object.keys(ma).map(function (k) { return ma[k]; }).sort(function (a, b) { return (b.do + b.vang) - (a.do + a.vang); }),
        nguoi: Object.keys(nguoi).map(function (k) { return nguoi[k]; }).sort(function (a, b) { return b.do - a.do || b.vang - a.vang; }).slice(0, 40) };
    }
    H['cckt.chay'] = function (req, s) {
      ccktQuyen(s);
      var tu = parseD(req.tu) ? req.tu : addDays(env.today(), -30), den = parseD(req.den) ? req.den : env.today(), bp = String(req.boPhan || '');
      if (den < tu) err('Khoảng ngày không hợp lệ.'); if (den > addDays(tu, 92)) err('Mỗi lần kiểm tra tối đa 3 tháng.');
      var q = ccktQuet(tu, den, bp), a = T('CanhBaoCC'), cu = {}, dot = { id: newId(), tu: tu, den: den, pham: bp || 'Toàn công ty', taoBoi: s.hoTen || s.username, taoLuc: env.now(), soNguoi: q.soNguoi, soLuot: q.soLuot, soDo: 0, soVang: 0, thongKe: '', aiNhanXet: '', aiLuc: '' };
      a.forEach(function (x) { cu[x.khoa] = x; });
      var moi = 0, giu = 0, cua = [];
      q.ds.forEach(function (x) {
        var o = cu[x.khoa];
        if (o && (o.trangThai === 'Đã xử lý' || o.trangThai === 'Bỏ qua')) return;             // đã kết luận ở đợt trước: không báo lại
        if (o) { o.moTa = x.moTa; o.mucDo = x.mucDo; o.luotIds = JSON.stringify(x.luotIds); o.dotId = dot.id; o.capNhat = env.now(); giu++; cua.push(o); return; }
        o = Object.assign({ id: newId(), dotId: dot.id, trangThai: 'Mới', nhacLuc: '', nguoiNhac: '', giaiTrinh: '', ketQua: '', nguoiGT: '', gtLuc: '', xuLy: '', nguoiXL: '', xlLuc: '', taoLuc: env.now(), capNhat: env.now() }, x);
        o.luotIds = JSON.stringify(x.luotIds); a.push(o); cu[x.khoa] = o; moi++; cua.push(o);
      });
      // cảnh báo cũ còn mở mà lần quét này không còn (đã được sửa: duyệt / bổ sung chấm công) → tự đóng
      var tuDong = 0; a.forEach(function (x) { if (CCKT_MO.indexOf(x.trangThai) >= 0 && x.ngay >= tu && x.ngay <= den && (!bp || x.boPhan === bp) && x.dotId !== dot.id) { x.trangThai = 'Đã xử lý'; x.xuLy = 'Tự đóng: lần kiểm tra ' + env.today() + ' không còn sai sót'; x.nguoiXL = 'Hệ thống'; x.xlLuc = env.now(); tuDong++; } });
      var tk = ccktThongKe(cua); dot.soDo = tk.n.do; dot.soVang = tk.n.vang; dot.thongKe = JSON.stringify(tk);
      T('KiemTraCC').push(dot); D('KiemTraCC'); D('CanhBaoCC');
      log(s, 'cckt.chay', tu + '→' + den + ' ' + dot.pham + ': ' + tk.n.do + ' đỏ, ' + tk.n.vang + ' vàng');
      return { dot: ccktDotPub(dot), moi: moi, giu: giu, tuDong: tuDong, ds: cua.map(ccktPub) };
    };
    function ccktPub(x) { var o = Object.assign({}, x); o.luotIds = parseJSON(x.luotIds, []); o.tenMa = (CCKT_MA[x.ma] || [])[1] || x.ma; o.tenNhom = CCKT_NHOM[x.nhom] || ''; return o; }
    function ccktDotPub(d) { var o = Object.assign({}, d); o.thongKe = parseJSON(d.thongKe, {}); o.aiNhanXet = parseJSON(d.aiNhanXet, null); return o; }
    H['cckt.ds'] = function (req, s) {
      var kt = CCKT_ROLES.indexOf(s.role) >= 0;
      if (!kt) err('Kiểm tra chấm công chỉ dành cho Giám đốc, TCHC, Kế toán.');
      var tt = String(req.trangThai || ''), ds = T('CanhBaoCC').filter(function (x) {
        if (!kt && (x.boPhan !== s.boPhan || x.trangThai === 'Mới')) return false;         // quản lý bộ phận: chỉ thấy cảnh báo đã được nhắc
        if (req.dotId && x.dotId !== req.dotId && !req.tatCa) return false;
        if (tt === 'mo' ? CCKT_MO.indexOf(x.trangThai) < 0 : tt && x.trangThai !== tt) return false;
        if (req.boPhan && x.boPhan !== req.boPhan) return false;
        return true;
      }).sort(function (a, b) { return (a.mucDo === b.mucDo ? 0 : a.mucDo === 'do' ? -1 : 1) || (a.ngay < b.ngay ? 1 : -1); }).slice(0, 3000).map(ccktPub);
      var luot = {}; ds.forEach(function (x) { x.luotIds.forEach(function (i) { luot[i] = 1; }); });
      return { ds: ds, luot: T('ChamCongGPS').filter(function (r) { return luot[r.id]; }).map(function (r) { return { id: r.id, ngay: r.ngay, gio: r.gio, loai: r.loai, tenDiem: r.tenDiem, lat: r.lat, lng: r.lng, khoangCach: r.khoangCach, banKinh: r.banKinh, doChinhXac: r.doChinhXac, anh: r.anh, khuonMat: r.khuonMat, diemKM: r.diemKM, trangThai: r.trangThai, thietBi: String(r.thietBi || '').slice(0, 10), nguoiDuyet: r.nguoiDuyet }; }),
        dots: kt ? T('KiemTraCC').slice(-30).reverse().map(ccktDotPub) : [], kt: kt };
    };
    H['cckt.nhac'] = function (req, s) {
      ccktQuyen(s);
      var ids = req.ids || [], ghi = String(req.ghiChu || '').trim().slice(0, 500), han = parseD(req.han) ? req.han : addDays(env.today(), 3), bp = {};
      T('CanhBaoCC').forEach(function (x) {
        if (ids.indexOf(x.id) < 0 || CCKT_MO.indexOf(x.trangThai) < 0) return;
        x.trangThai = x.trangThai === 'Mới' ? 'Đã nhắc' : x.trangThai; x.nhacLuc = env.now(); x.nguoiNhac = s.username; x.hanGT = han; if (ghi) x.xuLy = 'Nhắc: ' + ghi; x.capNhat = env.now();
        var b = bp[x.boPhan] = bp[x.boPhan] || { do: 0, vang: 0, nv: {} }; b[x.mucDo]++; b.nv[x.maNV] = 1;
      });
      var ten = Object.keys(bp); if (!ten.length) err('Chọn cảnh báo cần nhắc (chưa xử lý).');
      D('CanhBaoCC');
      var gs = nguoiGiamSatCC(s.username), out = [];          // v3.7.1: không báo trưởng bộ phận / CHT; TCHC – Kế toán tự kết luận
      ten.forEach(function (t) {
        var b = bp[t], nv = Object.keys(b.nv).map(userCuaNV).filter(Boolean);
        bao(gs, { loai: 'duyet', t: '⚠ Kiểm tra lại chấm công – ' + t, b: b.do + ' cảnh báo đỏ, ' + b.vang + ' vàng (' + nv.length + ' người). Hạn ' + han.slice(8) + '/' + han.slice(5, 7) + (ghi ? '. ' + ghi : ''), tag: 'cckt_' + t, mo: '#v=gpsql&t=kt' });
        if (req.guiNV) nv.forEach(function (u) { bao([u], { loai: 'cc', t: 'Chấm công của bạn cần kiểm tra lại', b: 'TCHC / Kế toán sẽ liên hệ để xác nhận một số lượt chấm công. Mở app → Chấm công để xem.', tag: 'cckt_nv', mo: '#v=ccgps' }); });
        out.push({ boPhan: t, nguoi: req.guiNV ? nv.map(chatTen) : [] });
      });
      log(s, 'cckt.nhac', ten.join(', ') + ' (' + ids.length + ')');
      return { boPhan: ten, n: ids.length, nguoiNhan: out };
    };
    H['cckt.giaitrinh'] = function (req, s) {
      var kq = ['Đúng – không sai sót', 'Đã điều chỉnh / bổ sung', 'Sai sót – đề nghị xử lý theo quy chế'].indexOf(req.ketQua) >= 0 ? req.ketQua : '', nd = String(req.noiDung || '').trim().slice(0, 1000), n = 0, nhac = {};
      if (!kq) err('Chọn kết quả kiểm tra.'); if (nd.length < 5) err('Nhập nội dung giải trình.');
      T('CanhBaoCC').forEach(function (x) {
        if ((req.ids || []).indexOf(x.id) < 0) return;
        if (!ccktQL(s, x)) err('Chỉ Giám đốc, TCHC, Kế toán kết luận cảnh báo chấm công.');
        if (CCKT_MO.indexOf(x.trangThai) < 0) return;
        x.trangThai = req.dong ? 'Đã xử lý' : 'Đã giải trình'; x.giaiTrinh = nd; x.ketQua = kq; x.nguoiGT = s.hoTen || s.username; x.gtLuc = env.now(); x.capNhat = env.now(); n++;
        if (req.dong) { x.xuLy = kq; x.nguoiXL = x.nguoiGT; x.xlLuc = env.now(); } else if (x.nguoiNhac && x.nguoiNhac !== s.username) nhac[x.nguoiNhac] = x.boPhan;   // v3.7.1: TCHC / Kế toán ghi kết quả và đóng luôn
      });
      if (!n) err('Không có cảnh báo nào được cập nhật.');
      D('CanhBaoCC'); log(s, 'cckt.giaitrinh', n + ' cảnh báo: ' + kq + (req.dong ? ' (đóng)' : ''));
      Object.keys(nhac).forEach(function (u) { bao([u], { loai: 'duyet', t: 'Đã giải trình chấm công – ' + nhac[u], b: (s.hoTen || s.username) + ': ' + kq + ' (' + n + ' cảnh báo). ' + nd.slice(0, 120), tag: 'cckt_gt', mo: '#v=gpsql&t=kt' }); });
      return { n: n };
    };
    H['cckt.xuly'] = function (req, s) {
      ccktQuyen(s);
      var tt = ['Đã xử lý', 'Bỏ qua', 'Mở lại'].indexOf(req.trangThai) >= 0 ? req.trangThai : '', gc = String(req.ghiChu || '').trim().slice(0, 500), n = 0;
      if (!tt) err('Trạng thái không hợp lệ.'); if (tt !== 'Đã xử lý' && !gc) err('Nhập lý do.');
      T('CanhBaoCC').forEach(function (x) {
        if ((req.ids || []).indexOf(x.id) < 0) return;
        x.trangThai = tt === 'Mở lại' ? 'Đã nhắc' : tt; x.xuLy = gc || x.xuLy; x.nguoiXL = s.hoTen || s.username; x.xlLuc = env.now(); x.capNhat = env.now(); n++;
      });
      D('CanhBaoCC'); log(s, 'cckt.xuly', tt + ' ' + n);
      return { n: n };
    };
    var CCKT_HE = 'Bạn là chuyên viên kiểm soát nội bộ về chấm công của Công ty Cổ phần TAPACO Việt Nam (xây dựng hạ tầng). Bạn nhận số liệu đã được hệ thống tổng hợp từ chấm công bằng điện thoại (định vị GPS + ảnh khuôn mặt) và các cảnh báo do bộ quy tắc phát hiện. ' +
      'Chỉ dùng số liệu được cung cấp, không bịa thêm số, tên người hay sự việc. Viết tiếng Việt, ngắn gọn, khách quan, không kết luận ai gian lận – chỉ nêu dấu hiệu cần kiểm tra. Chỉ trả về JSON.';
    H['cckt.ai'] = function (req, s) {               // chỉ đọc: gọi AI viết nhận xét; trình duyệt lưu kết quả bằng cckt.luuai
      ccktQuyen(s);
      if (!env.ai || !env.ai.trangThai().coKhoa) err('Chưa cài đặt AI. Quản trị vào "Phân tích nhóm (AI)" → "Cài đặt AI" để nhập khóa API. Các cảnh báo vẫn dùng được không cần AI.');
      var d = T('KiemTraCC').filter(function (x) { return x.id === req.dotId; })[0]; if (!d) err('Không tìm thấy đợt kiểm tra.');
      var tk = parseJSON(d.thongKe, {}), mau = T('CanhBaoCC').filter(function (x) { return x.dotId === d.id; }).sort(function (a, b) { return a.mucDo === b.mucDo ? 0 : a.mucDo === 'do' ? -1 : 1; }).slice(0, 40)
        .map(function (x) { return { maNV: x.maNV, hoTen: x.hoTen, boPhan: x.boPhan, ngay: x.ngay, loai: CCKT_MA[x.ma][1], mucDo: x.mucDo === 'do' ? 'đỏ' : 'vàng', moTa: x.moTa }; });
      var duLieu = { khoang: d.tu + ' → ' + d.den, pham: d.pham, soNguoiCoChamCong: d.soNguoi, soLuot: d.soLuot, canhBao: tk.n, theoBoPhan: tk.boPhan, theoLoai: tk.ma, caNhanNhieuCanhBao: (tk.nguoi || []).slice(0, 15), mauCanhBao: mau };
      var nd = 'SỐ LIỆU KIỂM TRA CHẤM CÔNG (JSON):\n' + JSON.stringify(duLieu) + '\n\nHãy phân tích và trả về đúng JSON:\n{"tongQuan": "<3–5 câu: mức độ tuân thủ chung, các vấn đề nổi bật>", "mucRuiRo": "Thấp|Trung bình|Cao", ' +
        '"boPhan": [{"boPhan": "<tên đúng như số liệu>", "mucRuiRo": "Thấp|Trung bình|Cao", "nhanXet": "<1–2 câu>"}], "caNhan": [{"maNV": "<mã đúng như số liệu>", "hoTen": "", "nhanXet": "<dấu hiệu cần kiểm tra>"}], ' +
        '"nguyenNhan": ["<nguyên nhân có thể (kỹ thuật, quy trình, ý thức)>"], "deXuat": [{"viec": "<việc cần làm>", "phuTrach": "<bộ phận / chức danh>", "uuTien": "Cao|Trung bình|Thấp"}]}';
      var o = aiGoi({ he: CCKT_HE, nd: nd, anh: [], maxTok: 3000, loai: 'cckt', duLieu: duLieu });
      var bpOK = {}; (tk.boPhan || []).forEach(function (b) { bpOK[b.boPhan] = 1; }); var nvOK = {}; (tk.nguoi || []).forEach(function (p) { nvOK[p.maNV] = p.hoTen; });
      var MUC = ['Thấp', 'Trung bình', 'Cao'], S = function (v, n) { return String(v == null ? '' : v).trim().slice(0, n); };
      return { tongQuan: S(o.tongQuan, 1500), mucRuiRo: MUC.indexOf(o.mucRuiRo) >= 0 ? o.mucRuiRo : 'Trung bình',
        boPhan: (o.boPhan || []).filter(function (b) { return b && bpOK[b.boPhan]; }).slice(0, 30).map(function (b) { return { boPhan: b.boPhan, mucRuiRo: MUC.indexOf(b.mucRuiRo) >= 0 ? b.mucRuiRo : '', nhanXet: S(b.nhanXet, 400) }; }),
        caNhan: (o.caNhan || []).filter(function (p) { return p && nvOK[p.maNV]; }).slice(0, 20).map(function (p) { return { maNV: p.maNV, hoTen: nvOK[p.maNV], nhanXet: S(p.nhanXet, 300) }; }),
        nguyenNhan: (o.nguyenNhan || []).slice(0, 8).map(function (x) { return S(x, 300); }).filter(Boolean),
        deXuat: (o.deXuat || []).slice(0, 10).filter(function (x) { return x && x.viec; }).map(function (x) { return { viec: S(x.viec, 300), phuTrach: S(x.phuTrach, 80), uuTien: MUC.indexOf(x.uuTien) >= 0 ? x.uuTien : 'Trung bình' }; }),
        ai: env.ai.trangThai().ncc, luc: env.now() };
    };
    H['cckt.luuai'] = function (req, s) {
      ccktQuyen(s);
      var d = T('KiemTraCC').filter(function (x) { return x.id === req.dotId; })[0]; if (!d) err('Không tìm thấy đợt kiểm tra.');
      d.aiNhanXet = JSON.stringify(req.nx || {}).slice(0, 40000); d.aiLuc = env.now(); D('KiemTraCC');
      return ccktDotPub(d);
    };

    /* ==================== VỊ TRÍ KHI NHẮN TIN (v3.3) ====================
       Trong giờ làm việc, mỗi lần CBCNV gửi tin nhắn (tối đa 1 lần / chu kỳ), app gửi vị trí điện thoại lên máy chủ.
       Máy chủ so với các điểm được phép chấm công của người đó (bộ phận, điểm chung, điểm riêng, điểm công tác trong ngày).
       Trong phạm vi: không lưu gì. Ngoài phạm vi (vượt bán kính + dung sai): ghi một dòng vào sheet KiemTraCong ("Kiểm tra công")
       và báo TCHC, Ban Giám đốc, Kế toán (gộp theo đợt để không báo dồn dập). Vị trí không hiện cho người nhận tin nhắn.
       Không kiểm tra: ngoài giờ làm việc, Chủ nhật / ngày lễ, ngày nghỉ có đơn đã duyệt, ngày công tác linh hoạt. */
    var VT_NHAN = ['r:tchc', 'r:giamdoc', 'r:ktluong', 'r:ktcongno'];
    function vtCfg() { return Object.assign({}, DEFAULT_CFG.viTriTN, getCfg().viTriTN || {}); }
    function vtTru(t, ms) {                       // 'yyyy-mm-dd HH:MM:SS' trừ ms → cùng định dạng
      var m = /^(\d{4})-(\d\d)-(\d\d) (\d\d):(\d\d):(\d\d)/.exec(t); if (!m) return t;
      var d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]) - Math.max(0, ms || 0)), p = function (n) { return ('0' + n).slice(-2); };
      return d.getUTCFullYear() + '-' + p(d.getUTCMonth() + 1) + '-' + p(d.getUTCDate()) + ' ' + p(d.getUTCHours()) + ':' + p(d.getUTCMinutes()) + ':' + p(d.getUTCSeconds());
    }
    function vtM(m) { return String(Math.round(num(m))).replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ' m'; }   // 2025 → 2.025 m
    function vtQuyen(s) { if (CCKT_ROLES.indexOf(s.role) < 0) err('Bạn chưa được phân quyền xem vị trí ngoài phạm vi.'); }
    function vtKiemNgay(nv, hom, c, g) {          // lý do không kiểm tra hôm nay / lúc này ('' = cần kiểm tra)
      var cfg = getCfg(), gio = env.now().slice(11, 16);
      if (c.chiGioLam) {
        if (laNgayNghiLe(hom, cfg)) return 'ngay-nghi';
        if (gio < g.gioVao || gio > g.gioRa) return 'ngoai-gio';
      }
      var nghi = T('NghiPhep').some(function (o) { return o.maNV === nv.maNV && o.trangThai === 'Đã duyệt' && o.tuNgay <= hom && o.denNgay >= hom && o.buoi !== 'Sáng' && o.buoi !== 'Chiều'; });
      if (nghi) return 'nghi-phep';
      if (ctCuaNgay(nv.maNV, hom).some(function (x) { return x.cheDo === 'linhhoat'; })) return 'cong-tac-linh-hoat';
      return '';
    }
    H['vitri.bao'] = function (req, s) {
      var c = vtCfg(); if (!c.bat) return { bo: 'tat' };
      if ((c.boQuaVaiTro || []).indexOf(s.role) >= 0) return { bo: 'vai-tro' };
      var nv = s.maNV && findNV(s.maNV); if (!nv || nv.trangThai === 'Đã nghỉ việc') return { bo: 'khong-nv' };
      var g = gpsCfg(), hom = env.today(), ly = vtKiemNgay(nv, hom, c, g); if (ly) return { bo: ly };
      var lat = +req.lat, lng = +req.lng, acc = Math.round(+req.doChinhXac), tuoi = Math.max(0, +req.tuoi || 0);
      if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180 || (!lat && !lng)) err('Vị trí không hợp lệ.');
      if (!isFinite(acc) || acc > num(c.doChinhXacToiDa)) return { bo: 'kem-chinh-xac', doChinhXac: acc };
      if (tuoi > 30 * 60000) return { bo: 'vi-tri-cu' };
      var kT = 'vt_' + s.username, tr = env.tam && env.tam.get(kT);
      if (tr && tr.t && Date.now() - tr.t < Math.max(1, num(c.chuKyPhut) - 1) * 60000) return { bo: 'chu-ky' };
      if (env.tam) env.tam.put(kT, { t: Date.now() }, 21600);
      var ds = gpsDiemCua(nv, hom); if (!ds.length) return { bo: 'chua-co-diem' };
      var gan = ds.map(function (p) { return { p: p, d: Math.round(khoangCachM(lat, lng, +p.lat, +p.lng)) }; }).sort(function (a, b) { return (a.d - num(a.p.banKinh)) - (b.d - num(b.p.banKinh)); })[0];
      var bk = num(gan.p.banKinh), du = Math.max(num(c.dungSai), acc);
      if (gan.d <= bk + du) return { trongVung: true, diem: gan.p.ten, khoangCach: gan.d, banKinh: bk };
      var now = env.now(), gioDo = vtTru(now, tuoi), bpRow = gpsBP(gan.p.boPhan);
      var latCC = +(+gan.p.lat).toFixed(6), lngCC = +(+gan.p.lng).toFixed(6), la = Math.round(lat * 1e6) / 1e6, ln = Math.round(lng * 1e6) / 1e6;
      var row = { id: newId(), ngayGhi: now.slice(0, 10), gioGhi: now.slice(11, 19), hoTen: nv.hoTen, maNV: nv.maNV,
        viTriChamCong: gan.p.ten + ' (' + latCC + ', ' + lngCC + '; bán kính ' + bk + ' m)', viTriNgoaiPhamVi: la + ', ' + ln + ' (±' + acc + ' m)', khoangCach: gan.d, gioDo: gioDo,
        boPhan: nv.phongBan || (bpRow ? bpRow.ten : ''), username: s.username, diemCC: gan.p.ten, latCC: latCC, lngCC: lngCC, banKinh: bk, lat: la, lng: ln, doChinhXac: acc,
        vuotBanKinh: Math.max(0, gan.d - bk), banDo: 'https://www.google.com/maps?q=' + la + ',' + ln,
        hoiThoai: String(req.ht || '').slice(0, 40), baoCao: '', ghiChu: gan.p.congTac ? 'Điểm gần nhất là điểm công tác' + (gan.p.choDuyet ? ' (chờ duyệt)' : '') : '' };
      // gộp theo đợt: đã báo trong baoLaiPhut và chỗ đứng cách lần báo trước < 500 m → chỉ ghi, không báo lại
      var truoc = T('KiemTraCong').filter(function (x) { return x.username === s.username && x.ngayGhi === row.ngayGhi && x.baoCao === 'Đã báo'; }).pop();
      var moi = !truoc || (new Date(row.ngayGhi + 'T' + row.gioGhi) - new Date(truoc.ngayGhi + 'T' + truoc.gioGhi)) / 60000 >= num(c.baoLaiPhut) || khoangCachM(row.lat, row.lng, +truoc.lat, +truoc.lng) >= 500;
      if (moi) {
        row.baoCao = 'Đã báo';
        bao(nguoiCo(VT_NHAN, s.username), { loai: 'cc', t: '📍 Ngoài phạm vi chấm công', b: nv.hoTen + ' (' + nv.maNV + (row.boPhan ? ' · ' + row.boPhan : '') + '): cách "' + gan.p.ten + '" ' + vtM(gan.d) + ' (bán kính ' + bk + ' m) lúc ' + gioDo.slice(11, 16), tag: 'vt-' + s.username, mo: '#v=gpsql&t=vt' });
      }
      T('KiemTraCong').push(row); D('KiemTraCong');
      log(s, 'vitri.ngoaivung', row.diemCC + ' · ' + row.khoangCach + ' m · ' + row.lat + ',' + row.lng);
      return { trongVung: false, diem: gan.p.ten, khoangCach: gan.d, banKinh: bk, daBao: moi };
    };
    H['vitri.xacnhan'] = function (req, s) { log(s, 'vitri.xacnhan', 'Đã đọc thông báo kiểm tra vị trí khi nhắn tin'); return { ok: true }; };
    function vtLoc(req) {
      var tu = parseD(req.tu) ? req.tu : addDays(env.today(), -6), den = parseD(req.den) ? req.den : env.today();
      return T('KiemTraCong').filter(function (x) { return x.ngayGhi >= tu && x.ngayGhi <= den && (!req.boPhan || x.boPhan === req.boPhan) && (!req.maNV || x.maNV === req.maNV); });
    }
    H['vitri.ds'] = function (req, s) {
      vtQuyen(s);
      var ds = vtLoc(req).slice().reverse().slice(0, 1000).map(function (x) { var o = Object.assign({}, x); ['latCC', 'lngCC', 'lat', 'lng', 'banKinh', 'khoangCach', 'doChinhXac', 'vuotBanKinh'].forEach(function (k) { o[k] = soThuc(x[k]); }); return o; });
      var url = ''; try { url = env.sheetUrl ? env.sheetUrl('KiemTraCong') : ''; } catch (e) { }
      return { ds: ds, cfg: vtCfg(), url: url };
    };
    H['vitri.caidat'] = function (req, s) {
      if (!can(s.role, 'edit_nv') && s.role !== 'giamdoc') err('Chỉ TCHC, Giám đốc, Quản trị.');
      var c = vtCfg(), o = req.viTriTN || {};
      ['bat', 'chiGioLam'].forEach(function (k) { if (o[k] != null) c[k] = !!o[k]; });
      [['chuKyPhut', 5, 240], ['doChinhXacToiDa', 20, 500], ['dungSai', 0, 1000], ['baoLaiPhut', 10, 1440]].forEach(function (a) {
        if (o[a[0]] == null) return; var v = +o[a[0]]; if (!(v >= a[1] && v <= a[2])) err('Giá trị "' + a[0] + '" phải từ ' + a[1] + ' đến ' + a[2] + '.'); c[a[0]] = Math.round(v); });
      if (Array.isArray(o.boQuaVaiTro)) c.boQuaVaiTro = o.boQuaVaiTro.filter(function (r) { return Object.prototype.hasOwnProperty.call(ROLES, r); });
      setCfg({ viTriTN: c }); log(s, 'vitri.caidat', JSON.stringify(c));
      return { viTriTN: c };
    };

    /* ==================== ĐÁNH GIÁ TỔNG HỢP CBCNV (v3.4) ====================
       Kết hợp 3 nguồn cho từng người trong một khoảng ngày:
         1) Chấm công GPS (ngày công, trễ, về sớm, thiếu chấm ra, sai vị trí, khuôn mặt) và cảnh báo kiểm tra chấm công (v3.2);
         2) Vị trí ngoài phạm vi khi nhắn tin (sheet KiemTraCong, v3.3);
         3) Hoạt động trên các nhóm công việc (ngày tương tác, báo cáo định kỳ nộp đúng hạn) + nhận xét cá nhân của các báo cáo "Phân tích nhóm (AI)" đã lập.
       Hệ thống tính chỉ số và điểm theo công thức cố định (minh bạch, kiểm tra lại được). AI chỉ viết nhận xét tích cực / hạn chế / thiếu sót / đề xuất
       từ các chỉ số đó (không đọc nội dung tin nhắn, ảnh). Kết quả AI là BẢN NHÁP: Giám đốc / TCHC xem, chốt xếp loại, kết luận và quyết định công bố
       cho nhân viên; nhân viên được xem bản đã công bố và phản hồi. Chỉ Giám đốc, TCHC, Quản trị dùng. */
    var DGT_ROLES = { indexOf: function (r) { return can(r, 'danhgia') ? 0 : -1; } }, DGT_XL = ['Tốt', 'Khá', 'Đạt', 'Cần cải thiện', 'Chưa đủ dữ liệu'];
    var DGT_HE = 'Bạn là chuyên viên nhân sự của Công ty Cổ phần TAPACO Việt Nam (xây dựng hạ tầng). Bạn viết nhận xét đánh giá định kỳ cho từng CBCNV dựa HOÀN TOÀN vào chỉ số do hệ thống tính ' +
      '(chấm công GPS, cảnh báo chấm công, vị trí ngoài phạm vi khi nhắn tin, hoạt động trên nhóm công việc, nhận xét của báo cáo phân tích nhóm). Nguyên tắc: chỉ dùng số liệu được cung cấp, nêu số liệu làm dẫn chứng; ' +
      'không bịa sự việc; không suy đoán đời tư, sức khỏe, gia đình, tôn giáo, chính trị; giọng khách quan, xây dựng, tôn trọng. Sai lệch vị trí có thể do đi công tác chưa đăng ký hoặc sai số GPS: ghi là "cần làm rõ", không kết luận vi phạm. ' +
      'Số tin nhắn không phản ánh đầy đủ hiệu quả công việc: chỉ coi là mức độ tương tác, phối hợp. Người ít dữ liệu: xếp "Chưa đủ dữ liệu", không suy diễn. Viết tiếng Việt. Chỉ trả về JSON.';
    function dgtQuyen(s) { if (DGT_ROLES.indexOf(s.role) < 0) err('Bạn chưa được phân quyền Đánh giá tổng hợp CBCNV.'); }
    function dgtKhoang(req) {
      var tu = String(req.tu || ''), den = String(req.den || ''), hom = env.today();
      if (!parseD(tu) || !parseD(den) || tu > den) err('Khoảng ngày không hợp lệ.');
      if (den > hom) den = hom;
      if (tu > den) err('Khoảng ngày không hợp lệ.');
      var n = 0; for (var d = tu; d <= den && n < 400; d = addDays(d, 1)) n++;
      if (n > 186) err('Mỗi đợt đánh giá tối đa 6 tháng.');
      return { tu: tu, den: den };
    }
    function dgtXepLoai(d) { return d == null ? 'Chưa đủ dữ liệu' : d >= 85 ? 'Tốt' : d >= 70 ? 'Khá' : d >= 55 ? 'Đạt' : 'Cần cải thiện'; }
    function dgtKep(x) { return Math.round(Math.max(0, Math.min(100, x))); }
    /* nhóm chat: hoạt động từng người (không đọc nội dung) + báo cáo định kỳ + nhận xét AI của báo cáo phân tích nhóm gần nhất trùng khoảng */
    function dgtNhom(tu, den) {
      var out = {}, all = null;
      try { all = env.chat.before(env.chat.last() + 1, env.chat.last()); } catch (e) { all = []; }
      T('HoiThoai').filter(function (h) { return h.loai === 'nhom' && h.trangThai !== 'xoa' && h.trangThai !== 'Đã giải tán'; }).forEach(function (h) {
        var r, tk; try { r = ltTin(h.id, all); tk = ltThongKe(h, r.ds, ltCtx(h, r.ds), tu, den); } catch (e) { return; }
        var ngay = {}; r.ds.forEach(function (m) { var d = m.thoiGian.slice(0, 10); if (d >= tu && d <= den && !m.thuHoi && LT_TT[m.loai]) (ngay[m.username] = ngay[m.username] || {})[d] = 1; });
        var aiCN = {}, aiTen = '';
        try {
          var c = ltCay(h, false);
          if (c && c.ai && env.luuTru) {
            var f = env.luuTru.ds(c.ai).filter(function (x) { var m = /^PhanTich_(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})_/.exec(x.ten); return m && m[1] <= den && m[2] >= tu; })
              .sort(function (a, b) { return a.ten < b.ten ? 1 : -1; })[0];
            if (f) { var o = parseJSON(env.luuTru.doc(f.id), null), mf = /^PhanTich_(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})_/.exec(f.ten); if (o && Array.isArray(o.caNhan)) { aiTen = (o.tu || mf[1]) + '→' + (o.den || mf[2]); o.caNhan.forEach(function (p) { aiCN[p.username] = p; }); } }
          }
        } catch (e) { }
        tk.caNhan.forEach(function (p) {
          if (p.vaiTro === 'không còn trong nhóm' && !(p.soTin + p.soAnh + p.soTep + p.bcNop)) return;
          var o = out[p.username] = out[p.username] || { nhom: [], ngay: {}, soTin: 0, soAnh: 0, soTep: 0, bcPhai: 0, bcNop: 0, bcDungHan: 0, bcTre: 0, binhChon: 0, ngayLV: 0, ngayNghiNhom: 0 };
          o.soTin += p.soTin; o.soAnh += p.soAnh; o.soTep += p.soTep; o.bcPhai += p.bcPhai; o.bcNop += p.bcNop; o.bcDungHan += p.bcDungHan; o.bcTre += p.bcTre; o.binhChon += p.binhChon;
          o.ngayLV = Math.max(o.ngayLV, p.soNgayLV); o.ngayNghiNhom = Math.max(o.ngayNghiNhom, p.soNgayNghi);
          Object.keys(ngay[p.username] || {}).forEach(function (d) { o.ngay[d] = 1; });
          var a = aiCN[p.username];
          o.nhom.push({ ten: aiTenHT(h), soTin: p.soTin, soNgayCoTT: p.soNgayCoTT, soNgayLV: p.soNgayLV, tyLeNop: p.tyLeNop, diemTG: p.diemTG,
            ai: a ? { khoang: aiTen, xepLoai: String(a.xepLoai || '').slice(0, 30), nhanXet: String(a.nhanXet || '').slice(0, 400), deXuat: String(a.deXuat || '').slice(0, 200) } : null });
        });
      });
      return out;
    }
    function dgtTinh(tu, den, boPhan, maNV) {
      var cfg = getCfg(), hom = env.today(), denCC = den < hom ? den : addDays(hom, -1);     // hôm nay chưa hết ngày: không tính thiếu chấm ra / không có công
      var lv = []; for (var d = tu, g = 0; d <= denCC && g < 400; d = addDays(d, 1), g++) if (!laNgayNghiLe(d, cfg)) lv.push(d);
      var users = {}, bgd = {}; T('Users').forEach(function (u) { if (u.maNV && u.active !== 'Không') { users[u.maNV] = u; if (u.role === 'giamdoc' || u.role === 'admin') bgd[u.maNV] = 1; } });
      var dauGPS = {}; T('ChamCongGPS').forEach(function (r) { if (!dauGPS[r.maNV] || r.ngay < dauGPS[r.maNV]) dauGPS[r.maNV] = r.ngay; });   // ngày bắt đầu dùng chấm công GPS của từng người
      var nvs = T('NhanVien').filter(function (n) {
        if (maNV) return n.maNV === maNV;
        if (bgd[n.maNV]) return false;                 // Ban Giám đốc, Quản trị: không đưa vào đợt đánh giá
        if (n.trangThai === 'Đã nghỉ việc' && (!n.ngayNghi || n.ngayNghi < tu)) return false;
        if (n.ngayVao && n.ngayVao > den) return false;
        return (!boPhan || n.phongBan === boPhan) && (users[n.maNV] || T('ChamCongGPS').some(function (r) { return r.maNV === n.maNV; }));
      });
      if (!nvs.length) err('Không có CBCNV nào trong phạm vi đã chọn.');
      var gps = {}, cb = {}, vt = {}, don = {}, ct = {};
      T('ChamCongGPS').forEach(function (r) { if (r.ngay >= tu && r.ngay <= den) (gps[r.maNV] = gps[r.maNV] || []).push(r); });
      T('CanhBaoCC').forEach(function (x) { if (x.ngay >= tu && x.ngay <= den) (cb[x.maNV] = cb[x.maNV] || []).push(x); });
      T('KiemTraCong').forEach(function (x) { if (x.ngayGhi >= tu && x.ngayGhi <= den) (vt[x.maNV] = vt[x.maNV] || []).push(x); });
      T('NghiPhep').forEach(function (o) { if (o.trangThai === 'Đã duyệt' && o.tuNgay <= den && o.denNgay >= tu) (don[o.maNV] = don[o.maNV] || []).push(o); });
      T('CongTac').forEach(function (c) { if (c.trangThai === 'Đã duyệt' && c.tuNgay <= den && c.denNgay >= tu) (ct[c.maNV] = ct[c.maNV] || []).push(c); });
      var nhom = dgtNhom(tu, den);
      return nvs.map(function (n) {
        var u = users[n.maNV] || {}, vao = n.ngayVao && n.ngayVao > tu ? n.ngayVao : tu, ra = n.trangThai === 'Đã nghỉ việc' && n.ngayNghi && n.ngayNghi < denCC ? n.ngayNghi : denCC;
        var lvP = lv.filter(function (x) { return x >= vao && x <= ra; }), nghi = {}, ctN = {};
        (don[n.maNV] || []).forEach(function (o) { for (var x = o.tuNgay, k = 0; x <= o.denNgay && k < 400; x = addDays(x, 1), k++) if (lvP.indexOf(x) >= 0) nghi[x] = (o.buoi && o.buoi !== 'Cả ngày' && o.tuNgay === o.denNgay) ? 0.5 : 1; });
        (ct[n.maNV] || []).forEach(function (c) { for (var x = c.tuNgay, k = 0; x <= c.denNgay && k < 400; x = addDays(x, 1), k++) if (lvP.indexOf(x) >= 0) ctN[x] = 1; });
        var soNghi = Object.keys(nghi).reduce(function (s2, x) { return s2 + nghi[x]; }, 0);
        // 1. chấm công: chỉ tính từ ngày người đó bắt đầu chấm công GPS (không trừ điểm những ngày trước khi áp dụng)
        var tuCC = dauGPS[n.maNV] && dauGPS[n.maNV] > vao ? dauGPS[n.maNV] : vao, lvCC = lvP.filter(function (x) { return x >= tuCC; });
        var ds = gps[n.maNV] || [], hl = ds.filter(function (r) { return r.trangThai === 'Hợp lệ'; }), coCong = {}, vaoN = {}, raN = {}, tre = 0, phutTre = 0, veSom = 0;
        hl.forEach(function (r) { coCong[r.ngay] = 1; if (r.loai === 'vao') { vaoN[r.ngay] = 1; if (num(r.tre) > 0) { tre++; phutTre += num(r.tre); } } else { raN[r.ngay] = 1; if (num(r.tre) < 0) veSom++; } });
        var thieuRa = Object.keys(vaoN).filter(function (x) { return x <= denCC && !raN[x]; }).length;
        var coCongLV = lvCC.filter(function (x) { return coCong[x]; }).length, khongCong = lvCC.filter(function (x) { return !coCong[x] && !nghi[x]; });
        var soNghiCC = lvCC.reduce(function (s2, x) { return s2 + (nghi[x] || 0); }, 0);
        var cc = { coDuLieu: ds.length > 0, tinhTu: tuCC, soNgayLV: lvCC.length, soNgayCoCong: Object.keys(coCong).length, soNgayCoCongLV: coCongLV, soNgayNghiCoDon: soNghiCC, soNgayCongTac: Object.keys(ctN).length,
          soNgayKhongCong: khongCong.length, ngayKhongCong: khongCong.slice(0, 15), diTre: tre, phutTre: phutTre, veSom: veSom, thieuChamRa: thieuRa,
          saiViTri: ds.filter(function (r) { return r.trangThai === 'Sai vị trí'; }).length, choDuyet: ds.filter(function (r) { return r.trangThai === 'Chờ duyệt'; }).length,
          tuChoi: ds.filter(function (r) { return r.trangThai === 'Từ chối'; }).length, matKhongKhop: ds.filter(function (r) { return r.khuonMat === 'Không khớp'; }).length };
        // 2. tuân thủ: cảnh báo kiểm tra chấm công + vị trí ngoài phạm vi khi nhắn tin
        var cbs = cb[n.maNV] || [], tinh = cbs.filter(function (x) { return x.ketQua !== 'Đúng – không sai sót' && x.xuLy !== 'Bỏ qua' && x.trangThai !== 'Bỏ qua'; });
        var nhomCB = { pv: 0, ng: 0, tg: 0 }; tinh.forEach(function (x) { if (x.nhom && nhomCB[x.nhom] != null) nhomCB[x.nhom]++; });
        var vts = vt[n.maNV] || [], ngayVT = {}; vts.forEach(function (x) { ngayVT[x.ngayGhi] = 1; });
        var tt = { canhBaoDo: tinh.filter(function (x) { return x.mucDo === 'do'; }).length, canhBaoVang: tinh.filter(function (x) { return x.mucDo !== 'do'; }).length, theoNhom: nhomCB,
          chuaGiaiTrinh: cbs.filter(function (x) { return x.trangThai === 'Mới' || x.trangThai === 'Đã nhắc'; }).length, daGiaiTrinh: cbs.filter(function (x) { return x.trangThai === 'Đã giải trình' || x.trangThai === 'Đã xử lý'; }).length,
          saiSot: cbs.filter(function (x) { return /^Sai sót/.test(x.ketQua || ''); }).length, khongSaiSot: cbs.filter(function (x) { return x.ketQua === 'Đúng – không sai sót'; }).length,
          ngoaiPhamVi: vts.length, ngayNgoaiPhamVi: Object.keys(ngayVT).length, xaNhatM: vts.reduce(function (m, x) { return Math.max(m, num(soThuc(x.khoangCach)) - num(soThuc(x.banKinh))); }, 0),
          mauCanhBao: tinh.slice(0, 6).map(function (x) { return x.ngay.slice(8, 10) + '/' + x.ngay.slice(5, 7) + ' ' + (CCKT_MA[x.ma] ? CCKT_MA[x.ma][1] : x.ma) + (x.mucDo === 'do' ? ' (đỏ)' : ''); }) };
        // 3. nhóm công việc
        var g = nhom[u.username] || null, tn = null;
        if (g) {
          var coSo = Math.max(1, (lvP.length || g.ngayLV) - soNghi), ngTT = lvP.filter(function (x) { return g.ngay[x]; }).length;
          tn = { soNhom: g.nhom.length, soTin: g.soTin, soAnh: g.soAnh, soTep: g.soTep, soNgayTuongTac: ngTT, tyLeNgayTT: Math.round(Math.min(1, ngTT / coSo) * 100),
            bcPhaiNop: g.bcPhai, bcDaNop: g.bcNop, bcDungHan: g.bcDungHan, bcTre: g.bcTre, tyLeDungHan: g.bcPhai ? Math.round(Math.min(1, g.bcDungHan / g.bcPhai) * 100) : null, binhChon: g.binhChon,
            nhom: g.nhom.slice(0, 8) };
        }
        // điểm (công thức cố định – hiện cho người xem)
        var dCC = null, dTT = null, dTN = null;
        if (cc.coDuLieu && lvCC.length) { var csCC = Math.max(1, lvCC.length - soNghiCC); dCC = dgtKep(100 * Math.min(1, coCongLV / csCC) - 2 * tre - Math.min(15, phutTre / 20) - 2 * veSom - 3 * thieuRa); }
        if (cc.coDuLieu || vts.length || cbs.length) dTT = dgtKep(100 - 8 * tt.canhBaoDo - 3 * tt.canhBaoVang - 5 * cc.saiViTri - 4 * tt.ngayNgoaiPhamVi - 10 * tt.saiSot - 2 * tt.chuaGiaiTrinh);
        if (tn) dTN = dgtKep(tn.bcPhaiNop ? 0.5 * tn.tyLeNgayTT + 0.5 * tn.tyLeDungHan : tn.tyLeNgayTT);
        var w = [[dCC, 40], [dTT, 35], [dTN, 25]].filter(function (x) { return x[0] != null; }), tw = w.reduce(function (s2, x) { return s2 + x[1]; }, 0);
        var dTong = w.length >= 2 ? Math.round(w.reduce(function (s2, x) { return s2 + x[0] * x[1]; }, 0) / tw) : null;   // dưới 2 tiêu chí có dữ liệu: không tính điểm tổng
        return { maNV: n.maNV, hoTen: n.hoTen, boPhan: n.phongBan || '', chucDanh: n.chucVu || '', username: u.username || '', chiSo: { chamCong: cc, tuanThu: tt, nhom: tn, khoang: { tu: vao, den: ra } },
          diemCC: dCC, diemTT: dTT, diemTN: dTN, diemTong: dTong, xepLoaiHT: w.length >= 2 ? dgtXepLoai(dTong) : 'Chưa đủ dữ liệu' };
      });
    }
    function dgtThongKe(ds) {
      var bp = {}, xl = {}; DGT_XL.forEach(function (k) { xl[k] = 0; });
      ds.forEach(function (p) { xl[p.xepLoaiHT] = (xl[p.xepLoaiHT] || 0) + 1; var b = bp[p.boPhan] = bp[p.boPhan] || { boPhan: p.boPhan, soNguoi: 0, tong: 0, n: 0, cc: 0, ncc: 0, tt: 0, ntt: 0, tn: 0, ntn: 0 };
        b.soNguoi++; if (p.diemTong != null) { b.tong += p.diemTong; b.n++; } if (p.diemCC != null) { b.cc += p.diemCC; b.ncc++; } if (p.diemTT != null) { b.tt += p.diemTT; b.ntt++; } if (p.diemTN != null) { b.tn += p.diemTN; b.ntn++; } });
      var tb = function (s2, n) { return n ? Math.round(s2 / n) : null; }, co = ds.filter(function (p) { return p.diemTong != null; });
      return { soNguoi: ds.length, diemTB: tb(co.reduce(function (s2, p) { return s2 + p.diemTong; }, 0), co.length), xepLoai: xl,
        boPhan: Object.keys(bp).map(function (k) { var b = bp[k]; return { boPhan: b.boPhan, soNguoi: b.soNguoi, diemTB: tb(b.tong, b.n), cc: tb(b.cc, b.ncc), tt: tb(b.tt, b.ntt), tn: tb(b.tn, b.ntn) }; }).sort(function (a, b) { return (a.diemTB == null ? 999 : a.diemTB) - (b.diemTB == null ? 999 : b.diemTB); }) };
    }
    function dgtCNPub(x, full) {
      var o = { id: x.id, dotId: x.dotId, maNV: x.maNV, hoTen: x.hoTen, boPhan: x.boPhan, chucDanh: x.chucDanh, username: x.username,
        diemCC: x.diemCC === '' ? null : soThuc(x.diemCC), diemTT: x.diemTT === '' ? null : soThuc(x.diemTT), diemTN: x.diemTN === '' ? null : soThuc(x.diemTN), diemTong: x.diemTong === '' ? null : soThuc(x.diemTong),
        xepLoaiHT: x.xepLoaiHT, aiXepLoai: x.aiXepLoai, aiTongQuan: x.aiTongQuan, aiTichCuc: parseJSON(x.aiTichCuc, []), aiHanChe: parseJSON(x.aiHanChe, []), aiThieuSot: parseJSON(x.aiThieuSot, []), aiDeXuat: parseJSON(x.aiDeXuat, []), aiLuc: x.aiLuc,
        xepLoai: x.xepLoai, ketLuan: x.ketLuan, nguoiChot: x.nguoiChot, chotLuc: x.chotLuc, congBo: x.congBo === 'Có', nvPhanHoi: x.nvPhanHoi, phLuc: x.phLuc };
      if (full) o.chiSo = parseJSON(x.chiSo, {});
      return o;
    }
    function dgtDotPub(d) { return { id: d.id, tieuDe: d.tieuDe, tu: d.tu, den: d.den, pham: d.pham, taoBoi: d.taoBoi, taoLuc: d.taoLuc, soNguoi: num(d.soNguoi), thongKe: parseJSON(d.thongKe, {}), aiTong: parseJSON(d.aiTong, null), aiLuc: d.aiLuc, docUrl: d.docUrl, pdfId: d.pdfId }; }
    function dgtDot(id) { var d = T('DanhGiaTH').filter(function (x) { return x.id === id; })[0]; if (!d) err('Không tìm thấy đợt đánh giá.'); return d; }
    H['dgt.chay'] = function (req, s) {
      dgtQuyen(s);
      var k = dgtKhoang(req), bp = String(req.boPhan || ''), ds = dgtTinh(k.tu, k.den, bp, '');
      var tk = dgtThongKe(ds), now = env.now(), pham = bp || 'Toàn công ty';
      var d = { id: newId(), tieuDe: String(req.tieuDe || '').trim().slice(0, 120) || ('Đánh giá ' + pham + ' ' + ngayVN(k.tu) + ' – ' + ngayVN(k.den)), tu: k.tu, den: k.den, pham: pham,
        taoBoi: s.hoTen || s.username, taoLuc: now, soNguoi: ds.length, thongKe: JSON.stringify(tk), aiTong: '', aiLuc: '', docId: '', docUrl: '', pdfId: '', trangThai: 'Nháp' };
      T('DanhGiaTH').push(d); D('DanhGiaTH');
      var cn = T('DanhGiaCN');
      ds.forEach(function (p) {
        cn.push({ id: newId(), dotId: d.id, maNV: p.maNV, hoTen: p.hoTen, boPhan: p.boPhan, chucDanh: p.chucDanh, username: p.username, chiSo: JSON.stringify(p.chiSo).slice(0, 45000),
          diemCC: p.diemCC == null ? '' : p.diemCC, diemTT: p.diemTT == null ? '' : p.diemTT, diemTN: p.diemTN == null ? '' : p.diemTN, diemTong: p.diemTong == null ? '' : p.diemTong, xepLoaiHT: p.xepLoaiHT,
          aiXepLoai: '', aiTongQuan: '', aiTichCuc: '', aiHanChe: '', aiThieuSot: '', aiDeXuat: '', aiLuc: '', xepLoai: '', ketLuan: '', nguoiChot: '', chotLuc: '', congBo: '', nvPhanHoi: '', phLuc: '', capNhat: now });
      });
      D('DanhGiaCN'); log(s, 'dgt.chay', pham + ' ' + k.tu + '→' + k.den + ': ' + ds.length + ' người');
      return { dot: dgtDotPub(d), ds: cn.filter(function (x) { return x.dotId === d.id; }).map(function (x) { return dgtCNPub(x, false); }) };
    };
    H['dgt.ds'] = function (req, s) {
      dgtQuyen(s);
      var dots = T('DanhGiaTH').slice().sort(function (a, b) { return a.taoLuc < b.taoLuc ? 1 : -1; }).slice(0, 60).map(dgtDotPub);
      if (!req.dotId) return { dots: dots };
      var d = dgtDot(req.dotId);
      return { dots: dots, dot: dgtDotPub(d), ds: T('DanhGiaCN').filter(function (x) { return x.dotId === d.id; }).map(function (x) { return dgtCNPub(x, !!req.full); }) };
    };
    H['dgt.xem'] = function (req, s) {            // một người: đủ chỉ số
      dgtQuyen(s);
      var x = T('DanhGiaCN').filter(function (z) { return z.id === req.id; })[0]; if (!x) err('Không tìm thấy.');
      return dgtCNPub(x, true);
    };
    function dgtS(v, n) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, n); }
    function dgtA(a, n, m) { return (Array.isArray(a) ? a : []).map(function (x) { return dgtS(x, m || 260); }).filter(Boolean).slice(0, n || 5); }
    H['dgt.ai'] = function (req, s) {             // chỉ đọc: AI viết nhận xét cho tối đa 6 người (ids) hoặc tổng hợp đợt (tong); trình duyệt lưu bằng dgt.luuai
      dgtQuyen(s);
      if (!env.ai || !env.ai.trangThai().coKhoa) err('Chưa cài đặt AI. Quản trị vào "Phân tích nhóm (AI)" → "Cài đặt AI" để nhập khóa API. Điểm và chỉ số vẫn dùng được không cần AI.');
      var d = dgtDot(req.dotId), cnAll = T('DanhGiaCN').filter(function (x) { return x.dotId === d.id; });
      if (req.tong) {
        var tk = parseJSON(d.thongKe, {}), ds = cnAll.map(function (x) { return dgtCNPub(x, true); });
        var vanDe = { diTre: 0, thieuChamRa: 0, saiViTri: 0, canhBaoDo: 0, ngoaiPhamVi: 0, bcTre: 0, bcThieu: 0, nguoiKhongCongNhieu: 0 };
        ds.forEach(function (p) { var c = p.chiSo.chamCong || {}, t = p.chiSo.tuanThu || {}, g = p.chiSo.nhom || {}; vanDe.diTre += c.diTre || 0; vanDe.thieuChamRa += c.thieuChamRa || 0; vanDe.saiViTri += c.saiViTri || 0; vanDe.canhBaoDo += t.canhBaoDo || 0; vanDe.ngoaiPhamVi += t.ngoaiPhamVi || 0; vanDe.bcTre += g.bcTre || 0; vanDe.bcThieu += Math.max(0, (g.bcPhaiNop || 0) - (g.bcDaNop || 0)); if ((c.soNgayKhongCong || 0) >= 3) vanDe.nguoiKhongCongNhieu++; });
        var thap = ds.filter(function (p) { return p.diemTong != null; }).sort(function (a, b) { return a.diemTong - b.diemTong; }).slice(0, 8).map(function (p) { return { maNV: p.maNV, hoTen: p.hoTen, boPhan: p.boPhan, diemTong: p.diemTong, xepLoai: p.aiXepLoai || p.xepLoaiHT }; });
        var cao = ds.filter(function (p) { return p.diemTong != null; }).sort(function (a, b) { return b.diemTong - a.diemTong; }).slice(0, 5).map(function (p) { return { maNV: p.maNV, hoTen: p.hoTen, boPhan: p.boPhan, diemTong: p.diemTong }; });
        var duLieuT = { khoang: d.tu + ' → ' + d.den, pham: d.pham, thongKe: tk, tongVanDe: vanDe, diemCaoNhat: cao, canChuY: thap };
        var nd = 'SỐ LIỆU TỔNG HỢP ĐỢT ĐÁNH GIÁ (JSON):\n' + JSON.stringify(duLieuT) + '\n\nTrả về đúng JSON: {"tongQuan": "<4–6 câu: mặt bằng chung về chuyên cần, tuân thủ, phối hợp>", "diemManh": ["<điểm mạnh chung>"], "tonTai": ["<tồn tại, thiếu sót chung>"], ' +
          '"boPhan": [{"boPhan": "<tên đúng như số liệu>", "nhanXet": "<1–2 câu>"}], "deXuat": [{"viec": "<việc cần làm>", "phuTrach": "<bộ phận / chức danh>", "uuTien": "Cao|Trung bình|Thấp"}]}';
        var o = aiGoi({ he: DGT_HE, nd: nd, anh: [], maxTok: 3000, loai: 'dgttong', duLieu: duLieuT }), bpOK = {}; (tk.boPhan || []).forEach(function (b) { bpOK[b.boPhan] = 1; });
        var MUC = ['Cao', 'Trung bình', 'Thấp'];
        return { tong: { tongQuan: dgtS(o.tongQuan, 1800), diemManh: dgtA(o.diemManh, 6), tonTai: dgtA(o.tonTai, 6), boPhan: (o.boPhan || []).filter(function (b) { return b && bpOK[b.boPhan]; }).slice(0, 30).map(function (b) { return { boPhan: b.boPhan, nhanXet: dgtS(b.nhanXet, 400) }; }),
          deXuat: (o.deXuat || []).filter(function (x) { return x && x.viec; }).slice(0, 10).map(function (x) { return { viec: dgtS(x.viec, 300), phuTrach: dgtS(x.phuTrach, 80), uuTien: MUC.indexOf(x.uuTien) >= 0 ? x.uuTien : 'Trung bình' }; }), ai: env.ai.trangThai().ncc, luc: env.now() } };
      }
      var ids = (Array.isArray(req.ids) ? req.ids : []).slice(0, 6), chon = cnAll.filter(function (x) { return ids.indexOf(x.id) >= 0; });
      if (!chon.length) err('Chưa chọn người cần đánh giá.');
      var duLieu = chon.map(function (x) { var p = dgtCNPub(x, true); return { maNV: p.maNV, hoTen: p.hoTen, boPhan: p.boPhan, chucDanh: p.chucDanh, diem: { chuyenCan: p.diemCC, tuanThu: p.diemTT, tuongTacNhom: p.diemTN, tong: p.diemTong, xepLoaiTheoDiem: p.xepLoaiHT }, chiSo: p.chiSo }; });
      var nd2 = 'KHOẢNG ĐÁNH GIÁ: ' + d.tu + ' → ' + d.den + '\nCÁCH TÍNH ĐIỂM (0–100): chuyên cần = tỷ lệ ngày có công trên ngày làm việc trừ nghỉ có đơn, trừ điểm đi trễ, về sớm, thiếu chấm ra; tuân thủ = 100 trừ cảnh báo đỏ / vàng, sai vị trí, ngày ngoài phạm vi, sai sót đã kết luận, cảnh báo chưa giải trình; ' +
        'tương tác nhóm = tỷ lệ ngày có tương tác và tỷ lệ báo cáo định kỳ nộp đúng hạn. Điểm tổng: chuyên cần 40%, tuân thủ 35%, tương tác 25% (bỏ phần thiếu dữ liệu).\n' +
        'DỮ LIỆU TỪNG NGƯỜI (JSON):\n' + JSON.stringify(duLieu) + '\n\nTrả về đúng JSON: {"caNhan": [{"maNV": "<mã đúng như dữ liệu>", "xepLoai": "Tốt|Khá|Đạt|Cần cải thiện|Chưa đủ dữ liệu", "tongQuan": "<2–3 câu>", ' +
        '"tichCuc": ["<mặt tích cực, có số liệu>"], "hanChe": ["<hạn chế>"], "thieuSot": ["<thiếu sót cụ thể cần khắc phục; không có thì để trống>"], "deXuat": ["<đề xuất: khen, nhắc nhở, hỗ trợ, đào tạo, làm rõ…>"]}]}. Đánh giá đủ ' + chon.length + ' người.';
      var o2 = aiGoi({ he: DGT_HE, nd: nd2, anh: [], maxTok: 4000, loai: 'dgt', duLieu: duLieu }), byMa = {}; chon.forEach(function (x) { byMa[x.maNV] = x; });
      return { ds: (o2.caNhan || []).filter(function (p) { return p && byMa[p.maNV]; }).map(function (p) {
        return { id: byMa[p.maNV].id, maNV: p.maNV, xepLoai: DGT_XL.indexOf(p.xepLoai) >= 0 ? p.xepLoai : byMa[p.maNV].xepLoaiHT, tongQuan: dgtS(p.tongQuan, 700), tichCuc: dgtA(p.tichCuc), hanChe: dgtA(p.hanChe), thieuSot: dgtA(p.thieuSot), deXuat: dgtA(p.deXuat) }; }), ai: env.ai.trangThai().ncc };
    };
    H['dgt.luuai'] = function (req, s) {
      dgtQuyen(s);
      var d = dgtDot(req.dotId), now = env.now(), n = 0;
      if (req.tong) { d.aiTong = JSON.stringify(req.tong).slice(0, 40000); d.aiLuc = now; D('DanhGiaTH'); }
      (Array.isArray(req.ds) ? req.ds : []).forEach(function (p) {
        var x = T('DanhGiaCN').filter(function (z) { return z.id === p.id && z.dotId === d.id; })[0]; if (!x) return;
        x.aiXepLoai = DGT_XL.indexOf(p.xepLoai) >= 0 ? p.xepLoai : ''; x.aiTongQuan = dgtS(p.tongQuan, 700);
        x.aiTichCuc = JSON.stringify(dgtA(p.tichCuc)); x.aiHanChe = JSON.stringify(dgtA(p.hanChe)); x.aiThieuSot = JSON.stringify(dgtA(p.thieuSot)); x.aiDeXuat = JSON.stringify(dgtA(p.deXuat));
        x.aiLuc = now; x.capNhat = now; n++;
      });
      if (n) D('DanhGiaCN');
      return { n: n, dot: dgtDotPub(d) };
    };
    H['dgt.chot'] = function (req, s) {           // người có thẩm quyền chốt: xếp loại chính thức, kết luận, công bố cho nhân viên
      dgtQuyen(s);
      var x = T('DanhGiaCN').filter(function (z) { return z.id === req.id; })[0]; if (!x) err('Không tìm thấy.');
      if (req.xepLoai && DGT_XL.indexOf(req.xepLoai) < 0) err('Xếp loại không hợp lệ.');
      var cb0 = x.congBo === 'Có';
      x.xepLoai = req.xepLoai || ''; x.ketLuan = dgtS(req.ketLuan, 2000); x.nguoiChot = s.hoTen || s.username; x.chotLuc = env.now(); x.capNhat = env.now();
      x.congBo = req.congBo && x.xepLoai ? 'Có' : '';
      D('DanhGiaCN'); log(s, 'dgt.chot', x.maNV + ' ' + x.xepLoai + (x.congBo ? ' (công bố)' : ''));
      if (x.congBo && !cb0 && x.username) bao([x.username], { loai: 'cc', t: 'Đánh giá tổng hợp của bạn', b: 'Đánh giá ' + ngayVN(dgtDot(x.dotId).tu) + ' – ' + ngayVN(dgtDot(x.dotId).den) + ' đã được công bố. Bạn có thể xem và phản hồi.', tag: 'dgt-' + x.id, mo: '#v=me' });
      return dgtCNPub(x, false);
    };
    H['dgt.cuatoi'] = function (req, s) {         // nhân viên: các đánh giá đã công bố của mình (chỉ số + nhận xét, không có xếp hạng người khác)
      if (!s.maNV) return [];
      return T('DanhGiaCN').filter(function (x) { return x.maNV === s.maNV && x.congBo === 'Có'; }).map(function (x) {
        var d = T('DanhGiaTH').filter(function (z) { return z.id === x.dotId; })[0] || {}, o = dgtCNPub(x, true); o.tu = d.tu; o.den = d.den; return o; }).sort(function (a, b) { return a.den < b.den ? 1 : -1; });
    };
    H['dgt.phanhoi'] = function (req, s) {
      var x = T('DanhGiaCN').filter(function (z) { return z.id === req.id && z.maNV === s.maNV && z.congBo === 'Có'; })[0]; if (!x) err('Không tìm thấy đánh giá của bạn.');
      var nd = dgtS(req.noiDung, 2000); if (nd.length < 3) err('Nhập nội dung phản hồi.');
      x.nvPhanHoi = nd; x.phLuc = env.now(); x.capNhat = env.now(); D('DanhGiaCN'); log(s, 'dgt.phanhoi', x.dotId);
      bao(nguoiCo(['r:tchc', 'r:giamdoc'], s.username), { loai: 'cc', t: 'Phản hồi đánh giá: ' + x.hoTen, b: nd.slice(0, 120), tag: 'dgtph-' + x.id, mo: '#v=dgt' });
      return { ok: true };
    };
    H['dgt.xoa'] = function (req, s) {
      dgtQuyen(s);
      var d = dgtDot(req.dotId);
      if (T('DanhGiaCN').some(function (x) { return x.dotId === d.id && x.congBo === 'Có'; })) err('Đợt đã công bố đánh giá cho nhân viên, không xóa được. Bỏ công bố từng người trước.');
      var a = T('DanhGiaTH'), b = T('DanhGiaCN');
      a.splice(a.indexOf(d), 1); for (var i = b.length - 1; i >= 0; i--) if (b[i].dotId === d.id) b.splice(i, 1);
      D('DanhGiaTH'); D('DanhGiaCN'); log(s, 'dgt.xoa', d.tieuDe); return { ok: true };
    };
    H['dgt.baocao'] = function (req, s) {         // lập báo cáo Google Docs + PDF trong thư mục không chia sẻ "_ĐÁNH GIÁ CBCNV"
      dgtQuyen(s);
      if (!env.luuTru || !env.luuTru.taoDoc) err('Máy chủ chưa bật lưu trữ Drive (chạy caiDatLuuTru).');
      var d = dgtDot(req.dotId), ds = T('DanhGiaCN').filter(function (x) { return x.dotId === d.id; }).map(function (x) { return dgtCNPub(x, true); });
      var k = dgtKhoi(dgtDotPub(d), ds, getCfg().congTy), tm = env.luuTru.thuMuc('', '_ĐÁNH GIÁ CBCNV (KHÔNG CHIA SẺ)');
      var r = env.luuTru.taoDoc(tm, d.tieuDe + ' (lập ' + env.now().slice(0, 16).replace(/[-: ]/g, '') + ')', k, { pdf: true });
      d.docId = r.id; d.docUrl = r.url; d.pdfId = r.pdfId || ''; D('DanhGiaTH'); log(s, 'dgt.baocao', d.tieuDe);
      return { docUrl: r.url, pdfId: r.pdfId || '' };
    };
    function dgtTomTat(p) {                       // tóm tắt chỉ số một người (dùng cho báo cáo Word)
      var c = (p.chiSo || {}).chamCong || {}, t = (p.chiSo || {}).tuanThu || {}, g = (p.chiSo || {}).nhom;
      var a = [];
      a.push(c.coDuLieu ? 'Chấm công (tính từ ' + ngayVN(c.tinhTu) + '): ' + c.soNgayCoCongLV + '/' + c.soNgayLV + ' ngày làm việc có công' + (c.soNgayNghiCoDon ? ', nghỉ có đơn ' + c.soNgayNghiCoDon : '') + (c.soNgayCongTac ? ', công tác ' + c.soNgayCongTac : '') +
        ', không có công ' + c.soNgayKhongCong + '; đi trễ ' + c.diTre + ' lần (' + c.phutTre + ' phút), về sớm ' + c.veSom + ', thiếu chấm ra ' + c.thieuChamRa + ', sai vị trí ' + c.saiViTri + ', khuôn mặt không khớp ' + c.matKhongKhop + '.' : 'Chấm công GPS: chưa có dữ liệu.');
      a.push('Tuân thủ: cảnh báo đỏ ' + (t.canhBaoDo || 0) + ', vàng ' + (t.canhBaoVang || 0) + ' (chưa giải trình ' + (t.chuaGiaiTrinh || 0) + ', sai sót đã kết luận ' + (t.saiSot || 0) + '); ngoài phạm vi khi nhắn tin ' + (t.ngoaiPhamVi || 0) + ' lượt / ' + (t.ngayNgoaiPhamVi || 0) + ' ngày' + (t.xaNhatM ? ', xa nhất vượt bán kính ' + Math.round(t.xaNhatM) + ' m' : '') + '.');
      a.push(g ? 'Nhóm công việc (' + g.soNhom + ' nhóm): tương tác ' + g.soNgayTuongTac + ' ngày (' + g.tyLeNgayTT + '%), ' + g.soTin + ' tin, ' + (g.soAnh + g.soTep) + ' ảnh / tệp; báo cáo định kỳ ' + (g.bcPhaiNop ? g.bcDaNop + '/' + g.bcPhaiNop + ' (đúng hạn ' + g.tyLeDungHan + '%)' : 'không phải nộp') + '.' : 'Nhóm công việc: không tham gia nhóm nào.');
      if (g) g.nhom.forEach(function (n) { if (n.ai && (n.ai.xepLoai || n.ai.nhanXet)) a.push('Phân tích nhóm "' + n.ten + '" (' + n.ai.khoang + '): ' + n.ai.xepLoai + (n.ai.nhanXet ? ' – ' + n.ai.nhanXet : '')); });
      return a;
    }
    function dgtKhoi(d, ds, congTy) {
      var K = [], h = function (t, x) { K.push({ t: t, x: x }); }, p = function (x) { if (x) K.push({ t: 'p', x: x }); }, li = function (a, tien) { (a || []).forEach(function (x) { if (x) K.push({ t: 'li', x: (tien || '') + x }); }); };
      var dd = function (v) { return v == null || v === '' ? '–' : String(v); }, tk = d.thongKe || {}, at = d.aiTong;
      K.push({ t: 'dau', x: congTy || 'CÔNG TY CỔ PHẦN TAPACO VIỆT NAM', y: 'Ngày lập: ' + ngayVN(env.today()) });
      h('h1', 'BÁO CÁO ĐÁNH GIÁ TỔNG HỢP CBCNV');
      K.push({ t: 'giua', x: d.pham + ' · Từ ' + ngayVN(d.tu) + ' đến ' + ngayVN(d.den) + ' · ' + ds.length + ' người' });
      K.push({ t: 'nho', x: 'Chỉ số và điểm do hệ thống tính theo công thức cố định từ chấm công GPS, cảnh báo kiểm tra chấm công, vị trí ngoài phạm vi khi nhắn tin và hoạt động trên nhóm công việc (không đọc nội dung tin nhắn). ' +
        'Điểm tổng = chuyên cần 40% + tuân thủ 35% + tương tác nhóm 25% (bỏ phần thiếu dữ liệu). Nhận xét của AI là bản tham khảo; xếp loại và kết luận chính thức do người có thẩm quyền chốt. Tài liệu nội bộ – không chia sẻ.' });
      h('h2', 'I. TỔNG QUAN');
      var xl = tk.xepLoai || {}; K.push({ t: 'bang', rows: [['Số người', 'Điểm TB'].concat(DGT_XL), [dd(tk.soNguoi), dd(tk.diemTB)].concat(DGT_XL.map(function (k) { return dd(xl[k] || 0); }))], rong: [2, 2, 1.5, 1.5, 1.5, 2.2, 2.4] });
      if ((tk.boPhan || []).length) K.push({ t: 'bang', rows: [['Bộ phận', 'Số người', 'Điểm TB', 'Chuyên cần', 'Tuân thủ', 'Tương tác']].concat(tk.boPhan.map(function (b) { return [b.boPhan, b.soNguoi, dd(b.diemTB), dd(b.cc), dd(b.tt), dd(b.tn)]; })), rong: [6, 1.6, 1.6, 1.8, 1.8, 1.8] });
      if (at) {
        p(at.tongQuan); if ((at.diemManh || []).length) { h('h3', 'Điểm mạnh'); li(at.diemManh); } if ((at.tonTai || []).length) { h('h3', 'Tồn tại, thiếu sót'); li(at.tonTai); }
        (at.boPhan || []).forEach(function (b) { li([b.boPhan + ': ' + b.nhanXet]); });
        if ((at.deXuat || []).length) { h('h3', 'Đề xuất'); K.push({ t: 'bang', rows: [['Việc cần làm', 'Phụ trách', 'Ưu tiên']].concat(at.deXuat.map(function (x) { return [x.viec, x.phuTrach, x.uuTien]; })), rong: [8, 3, 1.6] }); }
      }
      h('h2', 'II. BẢNG ĐIỂM');
      K.push({ t: 'bang', rows: [['STT', 'Mã NV', 'Họ tên', 'Bộ phận', 'Chuyên cần', 'Tuân thủ', 'Tương tác', 'Tổng', 'Xếp loại']].concat(ds.map(function (x, i) { return [i + 1, x.maNV, x.hoTen, x.boPhan, dd(x.diemCC), dd(x.diemTT), dd(x.diemTN), dd(x.diemTong), x.xepLoai ? x.xepLoai + ' (chốt)' : (x.aiXepLoai || x.xepLoaiHT)]; })), rong: [0.9, 1.6, 3.4, 3.4, 1.4, 1.4, 1.4, 1.2, 2.4] });
      h('h2', 'III. ĐÁNH GIÁ TỪNG NGƯỜI');
      ds.forEach(function (x, i) {
        h('h3', (i + 1) + '. ' + x.hoTen + ' – ' + x.maNV + (x.boPhan ? ' – ' + x.boPhan : ''));
        p('Điểm: chuyên cần ' + dd(x.diemCC) + ' · tuân thủ ' + dd(x.diemTT) + ' · tương tác ' + dd(x.diemTN) + ' · tổng ' + dd(x.diemTong) + ' → theo điểm: ' + x.xepLoaiHT + (x.aiXepLoai ? ' · AI đề xuất: ' + x.aiXepLoai : '') + (x.xepLoai ? ' · XẾP LOẠI CHÍNH THỨC: ' + x.xepLoai : ''));
        dgtTomTat(x).forEach(function (t) { K.push({ t: 'nho', x: t }); });
        if (x.aiTongQuan) p('Nhận xét (AI): ' + x.aiTongQuan);
        li(x.aiTichCuc, 'Tích cực: '); li(x.aiHanChe, 'Hạn chế: '); li(x.aiThieuSot, 'Thiếu sót: '); li(x.aiDeXuat, 'Đề xuất: ');
        if (x.ketLuan) p('Kết luận của ' + (x.nguoiChot || 'người có thẩm quyền') + ': ' + x.ketLuan);
        if (x.nvPhanHoi) p('Phản hồi của nhân viên: ' + x.nvPhanHoi);
      });
      K.push({ t: 'ky', x: ['NGƯỜI LẬP', 'TỔ CHỨC – HÀNH CHÍNH', 'GIÁM ĐỐC'] });
      return K;
    }

    /* ---- BẢNG TIN CÔNG TY ---- */
    var BT_MIME_ANH = /^image\/(jpeg|png|webp|gif)$/;
    function btThay(s, o) {
      if (can(s.role, 'bantin')) return true;
      if (o.trangThai !== 'Đăng') return false;
      var dt = parseJSON(o.doiTuong, []);
      if (!dt.length) return true;
      var nv = s.maNV && findNV(s.maNV);
      return !!nv && (dt.indexOf(nv.phongBan) >= 0 || dt.indexOf(nv.donVi) >= 0);
    }
    function btDs(s) { return T('BaiViet').filter(function (o) { return btThay(s, o); }); }
    function ttDs(s) {
      var ids = {}; btDs(s).forEach(function (o) { ids[o.id] = 1; });
      return T('TuongTac').filter(function (x) { return ids[x.baiId]; });
    }
    function btLay(s, id) {
      var o = T('BaiViet').filter(function (x) { return x.id === id; })[0];
      if (!o || !btThay(s, o)) err('Không tìm thấy bài viết.');
      return o;
    }
    function btTra(o) { return { bt: o, tt: T('TuongTac').filter(function (x) { return x.baiId === o.id; }) }; }
    function btTep(list, anh) {
      return (Array.isArray(list) ? list : parseJSON(list, [])).slice(0, anh ? 20 : 10).map(function (f) {
        if (!f || !f.fileId) err('Tệp đính kèm không hợp lệ.');
        return anh ? { fileId: String(f.fileId), ten: String(f.ten || '').slice(0, 150), mime: String(f.mime || ''), w: num(f.w), h: num(f.h) }
          : { fileId: String(f.fileId), ten: String(f.ten || '').slice(0, 150), mime: String(f.mime || ''), kichThuoc: num(f.kichThuoc) };
      });
    }
    function btXoaTep(ids) { ids.forEach(function (id) { try { env.files.del(id); } catch (e) { } }); }
    function btIds(o) { return parseJSON(o.anh, []).concat(parseJSON(o.dinhKem, [])).map(function (f) { return f.fileId; }); }

    H['bt.upload'] = function (req, s) {
      need(s, 'bantin');
      var f = req.file || {}, cfg = getCfg();
      var name = String(f.name || '').trim().replace(/[\\/:*?"<>|]/g, '_').slice(0, 150);
      if (!name || !f.data) err('Chưa chọn file.');
      if (!BT_MIME_ANH.test(f.mime || '') && !MIME_OK.test(f.mime || '')) err('Chỉ nhận ảnh (JPG, PNG, WEBP, GIF), PDF, Word, Excel.');
      var bytes = Math.floor(String(f.data).length * 3 / 4);
      if (bytes > (num(cfg.dungLuongToiDaMB) || 10) * 1048576) err('File vượt ' + (num(cfg.dungLuongToiDaMB) || 10) + ' MB.');
      var id = env.files.put({ name: 'BT_' + env.today() + '_' + name, mime: f.mime, data: f.data, bantin: true });
      log(s, 'bt.upload', name);
      return { fileId: id, ten: name, mime: f.mime, kichThuoc: bytes };
    };
    H['bt.save'] = function (req, s) {
      need(s, 'bantin');
      var r = req.row || {}, a = T('BaiViet'), o = r.id ? a.filter(function (x) { return x.id === r.id; })[0] : null;
      if (r.id && !o) err('Không tìm thấy bài viết.');
      if (o && o.tacGia !== s.username && s.role !== 'admin' && s.role !== 'giamdoc') err('Chỉ người đăng, Giám đốc hoặc Quản trị được sửa bài này.');
      var tieuDe = String(r.tieuDe || '').trim();
      if (!tieuDe) err('Nhập tiêu đề.');
      if (BT_LOAI.indexOf(r.loai) < 0) err('Chọn chuyên mục.');
      var anh = btTep(r.anh, true), dk = btTep(r.dinhKem, false);
      var dt = (Array.isArray(r.doiTuong) ? r.doiTuong : parseJSON(r.doiTuong, [])).map(String).filter(Boolean);
      var cu = o ? btIds(o) : [];
      var daDang = !!(o && o.trangThai === 'Đăng');
      if (!o) { o = { id: newId(), tacGia: s.username, tenTacGia: s.hoTen || s.username, taoLuc: env.now() }; a.push(o); }
      o.loai = r.loai; o.tieuDe = tieuDe.slice(0, 200); o.noiDung = String(r.noiDung || '').slice(0, 20000);
      o.anh = JSON.stringify(anh); o.dinhKem = JSON.stringify(dk); o.doiTuong = JSON.stringify(dt);
      o.ghim = r.ghim ? 'Có' : ''; o.xacNhan = r.xacNhan ? 'Có' : '';
      o.noiBat = r.noiBat ? 'Có' : ''; o.noiBatDen = r.noiBat && parseD(r.noiBatDen) ? String(r.noiBatDen).slice(0, 10) : '';
      o.soVanBan = String(r.soVanBan || '').trim().slice(0, 60); o.hieuLuc = parseD(r.hieuLuc) ? String(r.hieuLuc).slice(0, 10) : '';
      o.trangThai = r.trangThai === 'Nháp' ? 'Nháp' : 'Đăng';
      if (o.trangThai === 'Đăng' && !o.dangLuc) o.dangLuc = env.now();
      o.capNhat = env.now();
      var moi = btIds(o); btXoaTep(cu.filter(function (id) { return moi.indexOf(id) < 0; }));
      D('BaiViet');
      log(s, 'bt.save', o.loai + ': ' + o.tieuDe + ' (' + o.trangThai + ')');
      if (o.trangThai === 'Đăng' && !daDang) bao(userActive().filter(function (u) { return u.username !== s.username && btThay({ role: 'nhanvien', maNV: u.maNV }, o); }).map(function (u) { return u.username; }),
        { loai: 'bt', t: (o.xacNhan === 'Có' ? 'Văn bản cần xác nhận đã đọc' : 'Bảng tin · ' + o.loai), b: o.tieuDe + (o.soVanBan ? ' (' + o.soVanBan + ')' : ''), tag: 'bt_' + o.id, mo: '#bt=' + o.id });
      return btTra(o);
    };
    H['bt.delete'] = function (req, s) {
      need(s, 'bantin');
      var a = T('BaiViet'), o = a.filter(function (x) { return x.id === req.id; })[0];
      if (!o) err('Không tìm thấy bài viết.');
      if (o.tacGia !== s.username && s.role !== 'admin' && s.role !== 'giamdoc') err('Chỉ người đăng, Giám đốc hoặc Quản trị được xóa bài này.');
      btXoaTep(btIds(o));
      a.splice(a.indexOf(o), 1); D('BaiViet');
      var t = T('TuongTac'); for (var i = t.length - 1; i >= 0; i--) if (t[i].baiId === o.id) t.splice(i, 1); D('TuongTac');
      log(s, 'bt.delete', o.tieuDe);
      return true;
    };
    H['bt.pin'] = function (req, s) {
      need(s, 'bantin');
      var o = btLay(s, req.id); o.ghim = req.ghim ? 'Có' : ''; D('BaiViet');
      log(s, 'bt.pin', o.tieuDe + (o.ghim ? ' (ghim)' : ' (bỏ ghim)'));
      return btTra(o);
    };
    H['bt.file'] = function (req, s) {
      var fid = String(req.fileId || ''), o = null, f = null;
      T('BaiViet').forEach(function (x) {
        if (o || !btThay(s, x)) return;
        var m = parseJSON(x.anh, []).concat(parseJSON(x.dinhKem, [])).filter(function (z) { return z.fileId === fid; })[0];
        if (m) { o = x; f = m; }
      });
      if (!o) err('Không tìm thấy tệp hoặc bạn không có quyền xem.');
      var d = env.files.get(fid); if (!d) err('Tệp không còn trên máy chủ.');
      return { data: d.data, mime: f.mime, name: f.ten };
    };
    function ttThem(s, o, loai, noiDung) {
      var x = { id: newId(), baiId: o.id, loai: loai, username: s.username, hoTen: s.hoTen || s.username, noiDung: noiDung || '', thoiGian: env.now() };
      T('TuongTac').push(x); D('TuongTac'); return x;
    }
    H['bt.like'] = function (req, s) {
      var o = btLay(s, req.id), t = T('TuongTac');
      for (var i = t.length - 1; i >= 0; i--) if (t[i].baiId === o.id && t[i].loai === 'thich' && t[i].username === s.username) { t.splice(i, 1); D('TuongTac'); return btTra(o); }
      ttThem(s, o, 'thich'); return btTra(o);
    };
    H['bt.ack'] = function (req, s) {
      var o = btLay(s, req.id);
      if (o.xacNhan !== 'Có') err('Bài này không yêu cầu xác nhận.');
      if (o.trangThai !== 'Đăng') err('Bài chưa đăng.');
      if (!T('TuongTac').some(function (x) { return x.baiId === o.id && x.loai === 'xacnhan' && x.username === s.username; })) {
        ttThem(s, o, 'xacnhan'); log(s, 'bt.ack', o.tieuDe);
      }
      return btTra(o);
    };
    H['bt.comment'] = function (req, s) {
      var o = btLay(s, req.id), nd = String(req.noiDung || '').trim();
      if (o.trangThai !== 'Đăng') err('Bài chưa đăng.');
      if (!nd) err('Nhập nội dung bình luận.');
      ttThem(s, o, 'binhluan', nd.slice(0, 2000));
      if (o.tacGia !== s.username) bao([o.tacGia], { loai: 'bt', t: (s.hoTen || s.username) + ' bình luận bài của bạn', b: nd.slice(0, 200), tag: 'btc_' + o.id, mo: '#bt=' + o.id });
      return btTra(o);
    };
    H['bt.cdel'] = function (req, s) {
      var t = T('TuongTac'), x = t.filter(function (z) { return z.id === req.id && z.loai === 'binhluan'; })[0];
      if (!x) err('Không tìm thấy bình luận.');
      if (x.username !== s.username && !can(s.role, 'bantin')) err('Bạn chỉ xóa được bình luận của mình.');
      var o = btLay(s, x.baiId);
      t.splice(t.indexOf(x), 1); D('TuongTac');
      if (x.username !== s.username) log(s, 'bt.cdel', o.tieuDe + ': ' + x.hoTen);
      return btTra(o);
    };

    /* ---- khen thưởng / kỷ luật ---- */
    H['kt.save'] = function (req, s) {
      need(s, 'propose');
      var r = req.row || {}, rows = T('KhenThuong');
      var list = r.maNVs && r.maNVs.length ? r.maNVs : [r.maNV];
      if (LOAI_KT.indexOf(r.loai) < 0) err('Loại không hợp lệ.');
      if (r.kyLuong && !validKy(r.kyLuong)) err('Kỳ lương phải có dạng YYYY-MM.');
      if (r.loai !== 'Kỷ luật' && num(r.soTien) <= 0) err('Nhập số tiền thưởng.');
      var out = [];
      list.forEach(function (ma) {
        if (!findNV(ma)) err('Không tìm thấy nhân viên ' + ma);
        var o = r.id ? rows.filter(function (x) { return x.id === r.id; })[0] : null;
        if (o && o.trangThai !== 'Chờ duyệt') err('Bản ghi đã được xử lý, không sửa được.');
        if (!o) { o = { id: newId(), trangThai: 'Chờ duyệt', nguoiDeXuat: s.username, ngayDeXuat: env.now() }; rows.push(o); }
        o.maNV = ma;
        ['loai', 'hinhThuc', 'ngay', 'kyLuong', 'noiDung', 'soQD'].forEach(function (c) { o[c] = String(r[c] == null ? '' : r[c]).trim(); });
        o.soTien = r.loai === 'Kỷ luật' ? 0 : num(r.soTien);
        o.chiuThue = r.loai === 'Kỷ luật' ? '' : (r.chiuThue === 'Không' ? 'Không' : 'Có');
        out.push(o);
      });
      D('KhenThuong');
      log(s, 'kt.save', out.length + ' bản ghi ' + r.loai);
      return out;
    };

    H['kt.decide'] = function (req, s) {
      need(s, 'approve');
      var ids = req.ids || [req.id], out = [], tinhLai = {};
      T('KhenThuong').forEach(function (o) {
        if (ids.indexOf(o.id) < 0 || o.trangThai !== 'Chờ duyệt') return;
        if (req.approve && o.kyLuong) {
          var k = T('KyLuong').filter(function (x) { return x.ky === o.kyLuong; })[0];
          if (k) kyRow(o.kyLuong, false);
          if (k && daTrinhDuyet(k)) err('Kỳ lương ' + o.kyLuong + ' đã trình/được phê duyệt, không thể đưa thêm khoản thưởng.');
          if (k && canTinhLai(k)) tinhLai[o.kyLuong] = 1;
        }
        o.trangThai = req.approve ? 'Đã duyệt' : 'Từ chối'; o.nguoiDuyet = s.username; o.ngayDuyet = env.now();
        out.push(o);
      });
      D('KhenThuong');
      Object.keys(tinhLai).forEach(function (ky) { recalc(ky, s); });
      log(s, 'kt.' + (req.approve ? 'approve' : 'reject'), out.length + ' bản ghi');
      return out;
    };

    H['kt.delete'] = function (req, s) {
      var a = T('KhenThuong');
      for (var i = a.length - 1; i >= 0; i--) if (a[i].id === req.id) {
        if (a[i].trangThai === 'Đã duyệt' && s.role !== 'admin') err('Chỉ Quản trị được xóa bản ghi đã duyệt.');
        need(s, 'propose');
        log(s, 'kt.delete', a[i].maNV); a.splice(i, 1);
      }
      D('KhenThuong'); return true;
    };

    /* ---- tài liệu đính kèm hồ sơ (file lưu riêng tư qua env.files) ---- */
    var CCCD_TEN = { truoc: 'CCCD – mặt trước', sau: 'CCCD – mặt sau' };
    var MIME_OK = /^(application\/pdf|image\/(jpeg|png|webp|heic|heif)|application\/(msword|vnd\.openxmlformats-officedocument\.(wordprocessingml\.document|spreadsheetml\.sheet))|application\/vnd\.ms-excel)$/;
    function tlMeta(r, cfg, laNV) {
      var o = {};
      ['loai', 'ten', 'soHieu', 'ngay', 'hetHan', 'lienKet', 'ghiChu'].forEach(function (k) { if (r[k] !== undefined) o[k] = String(r[k] || '').trim().slice(0, 500); });
      if (o.loai !== undefined && (cfg.loaiTaiLieu || []).indexOf(o.loai) < 0) err('Loại tài liệu không hợp lệ.');
      if (laNV && o.loai !== undefined && (cfg.loaiNVTuTai || []).indexOf(o.loai) < 0) err('Loại "' + o.loai + '" do TCHC cập nhật.');
      ['ngay', 'hetHan'].forEach(function (k) { if (o[k] && !parseD(o[k])) err('Ngày không hợp lệ.'); });
      if (!laNV && r.nvXem !== undefined) o.nvXem = r.nvXem === 'Không' ? 'Không' : 'Có';
      return o;
    }
    H['tl.upload'] = function (req, s) {
      var r = req.row || {}, f = req.file || {}, cfg = getCfg();
      var nv = findNV(r.maNV); if (!nv) err('Không tìm thấy nhân viên.');
      var ql = can(s.role, 'edit_nv'), laNV = !ql;
      if (laNV && r.maNV !== s.maNV) err('Bạn chỉ tải tài liệu cho hồ sơ của mình.');
      var mat = req.cccdMat === 'truoc' || req.cccdMat === 'sau' ? req.cccdMat : '';   // v3.10: ảnh CCCD mặt trước / mặt sau trong lý lịch
      if (mat) {
        if (!/^image\/(jpeg|png|webp)$/.test(f.mime || '')) err('Ảnh CCCD phải là ảnh JPG, PNG hoặc WEBP.');
        r = { maNV: nv.maNV, loai: 'Căn cước công dân', ten: CCCD_TEN[mat], soHieu: nv.cccd || '', ngay: parseD(nv.ngayCapCCCD) ? nv.ngayCapCCCD : '', ghiChu: 'Ảnh chụp CCCD (' + CCCD_TEN[mat].slice(7) + ')' };
        if ((cfg.loaiTaiLieu || []).indexOf(r.loai) < 0) err('Danh mục tài liệu chưa có loại "Căn cước công dân" (Cài đặt → Loại tài liệu).');
      }
      if (!r.loai) err('Chọn loại tài liệu.');
      var name = String(f.name || '').trim().replace(/[\\/:*?"<>|]/g, '_').slice(0, 150);
      if (!name || !f.data) err('Chưa chọn file.');
      if (!MIME_OK.test(f.mime || '')) err('Chỉ nhận PDF, ảnh (JPG, PNG, WEBP, HEIC), Word hoặc Excel.');
      var bytes = Math.floor(String(f.data).length * 3 / 4);
      if (bytes > (num(cfg.dungLuongToiDaMB) || 10) * 1048576) err('File vượt ' + (num(cfg.dungLuongToiDaMB) || 10) + ' MB.');
      if (r.lienKet && !/^(qt|ll|kt):/.test(r.lienKet)) err('Liên kết không hợp lệ.');
      var o = tlMeta(r, cfg, laNV);
      o.id = newId(); o.maNV = nv.maNV; o.ten = o.ten || o.loai; o.nvXem = laNV ? 'Có' : (o.nvXem || 'Có');
      o.trangThai = laNV ? 'Chờ xác nhận' : 'Đã xác nhận'; o.nguoiXacNhan = laNV ? '' : (s.hoTen || s.username);
      o.tenFile = name; o.mimeType = f.mime; o.kichThuoc = bytes; o.nguoiTai = s.hoTen || s.username; o.taiLuc = env.now();
      o.fileId = env.files.put({ name: o.maNV + '_' + o.loai + '_' + name, mime: f.mime, data: f.data, maNV: nv.maNV, hoTen: nv.hoTen, duong: [tenThuMucNV(nv), tlNhom(o.loai).thuMuc] });   // Drive: <thư mục hồ sơ>/<mã – họ tên>/<nhóm>/
      T('TaiLieu').push(o);
      var xoa = [];
      if (mat) {                                   // thay ảnh cũ cùng mặt: chỉ xóa những bản người tải được phép xóa (TCHC: mọi bản; CBCNV: bản của mình đang chờ xác nhận)
        var a = T('TaiLieu');
        for (var i = a.length - 1; i >= 0; i--) { var x = a[i];
          if (x === o || x.maNV !== nv.maNV || x.loai !== 'Căn cước công dân' || x.ten !== CCCD_TEN[mat]) continue;
          if (!ql && x.trangThai !== 'Chờ xác nhận') continue;
          try { env.files.del(x.fileId); } catch (e) { }
          xoa.push(x.id); a.splice(i, 1);
        }
      }
      D('TaiLieu');
      log(s, 'tl.upload', nv.maNV + ' ' + o.loai + ': ' + name + (mat ? ' (' + o.ten + (xoa.length ? ', thay ' + xoa.length + ' ảnh cũ' : '') + ')' : ''));
      return mat ? Object.assign({ xoa: xoa }, o) : o;
    };
    H['tl.thumuc'] = function (req, s) {           // v3.17: mở thư mục hồ sơ của CBCNV trên Google Drive (chỉ tài khoản quản lý Drive của Công ty mở được)
      if (!can(s.role, 'tailieu')) err('Bạn không có quyền.');
      var nv = findNV(req.maNV); if (!nv) err('Không tìm thấy nhân viên.');
      if (!env.files.duongUrl) err('Bản dùng thử lưu tệp trên trình duyệt – trên hệ thống thật, tệp nằm trong Google Drive của Công ty.');
      var nh = TL_NHOM.filter(function (x) { return x.ma === req.nhom; })[0];
      return { url: env.files.duongUrl(nh ? [tenThuMucNV(nv), nh.thuMuc] : [tenThuMucNV(nv)]) };
    };
    H['tl.update'] = function (req, s) {
      need(s, 'edit_nv');
      var o = T('TaiLieu').filter(function (x) { return x.id === req.id; })[0]; if (!o) err('Không tìm thấy tài liệu.');
      var m = tlMeta(req.row || {}, getCfg(), false);
      Object.keys(m).forEach(function (k) { o[k] = m[k]; });
      if (req.xacNhan && o.trangThai !== 'Đã xác nhận') { o.trangThai = 'Đã xác nhận'; o.nguoiXacNhan = s.hoTen || s.username; }
      D('TaiLieu'); log(s, 'tl.update', o.maNV + ' ' + o.loai + (req.xacNhan ? ' (xác nhận)' : ''));
      return o;
    };
    H['tl.get'] = function (req, s) {
      var o = T('TaiLieu').filter(function (x) { return x.id === req.id; })[0]; if (!o) err('Không tìm thấy tài liệu.');
      var hdK = o.loai === 'Hồ sơ khoán việc' ? T('HopDongKhoan').filter(function (h) { return o.soHieu === h.soHD && h.maNV === o.maNV; })[0] : null;   // v3.28
      if (!can(s.role, 'tailieu') && !(o.maNV === s.maNV && o.nvXem !== 'Không') && !(hdK && quyenKhoan(s, hdK))) err('Bạn không có quyền xem tài liệu này.');
      var f = env.files.get(o.fileId);
      if (!f) err('Không đọc được file (có thể đã bị xóa khỏi Drive).');
      log(s, 'tl.xem', o.maNV + ' ' + o.loai + ': ' + o.tenFile);
      return { name: o.tenFile, mime: o.mimeType, data: f.data };
    };

    /* ---- v3.37: OCR tài liệu hồ sơ ---- */
    function tlQuyen(o, s) {
      if (!can(s.role, 'tailieu') && !(o.maNV === s.maNV && o.nvXem !== 'Không')) err('Bạn không có quyền với tài liệu này.');
    }
    H['tl.ocr'] = function (req, s) {
      var o = T('TaiLieu').filter(function (x) { return x.id === req.id; })[0]; if (!o) err('Không tìm thấy tài liệu.');
      tlQuyen(o, s);
      if (!/^(image\/|application\/pdf)/.test(o.mimeType || '')) err('Chỉ đọc chữ được ảnh chụp hoặc PDF.');
      var text = '', nguon = '';
      if (typeof req._ocr === 'string' && req._ocrFid === o.fileId) { text = req._ocr; nguon = 'Google Drive OCR'; }       // Apps Script đã đọc trước khi lấy khóa
      else if (typeof req.text === 'string' && req.text.trim()) { text = req.text; nguon = 'Đọc trên máy (Tesseract)'; }
      else if (req._ocrLoi) err('OCR_CHUA_BAT: ' + req._ocrLoi);
      else if (env.files.ocr) { try { text = env.files.ocr(o.fileId); nguon = 'Google Drive OCR'; } catch (e) { err('OCR_CHUA_BAT: ' + e.message); } }
      else err('OCR_CHUA_BAT: máy chủ chưa bật đọc chữ (Drive API).');
      text = String(text || '').replace(/\r/g, '').slice(0, 200000);
      if (!text.trim()) err('Không đọc được chữ nào – ảnh mờ, nghiêng hoặc không có chữ.');
      var nv = findNV(o.maNV) || {}, ten = String(o.tenFile || o.ten || 'tai-lieu').replace(/\.[^.]+$/, '');
      if (o.ocrId) { try { env.files.del(o.ocrId); } catch (e) { } }
      o.ocrId = env.files.put({ name: o.maNV + '_' + ten + '_OCR.txt', mime: 'text/plain', data: b64Utf8(text), maNV: nv.maNV, hoTen: nv.hoTen, duong: [tenThuMucNV(nv), tlNhom(o.loai).thuMuc] });
      o.ocrLuc = env.now(); o.ocrNguon = nguon; o.ocrTom = text.replace(/\s+/g, ' ').trim().slice(0, 1500);
      D('TaiLieu');
      var dx = ocrTach(text, o.loai), hien = {};
      OCR_TRUONG.forEach(function (x) { hien[x[0]] = String(nv[x[0]] || ''); }); hien.hoTen = nv.hoTen || '';
      var tlH = {}; OCR_TL.forEach(function (x) { tlH[x[0]] = String(o[x[0]] || ''); });
      var cb = [];
      if (dx.nv.hoTen && ocrBoDau(dx.nv.hoTen).replace(/\s+/g, ' ') !== ocrBoDau(nv.hoTen).replace(/\s+/g, ' ')) cb.push('Họ tên đọc được "' + dx.nv.hoTen + '" khác hồ sơ "' + nv.hoTen + '" – kiểm tra tài liệu có đúng người.');
      if (dx.nv.cccd && T('NhanVien').some(function (x) { return x.maNV !== nv.maNV && x.cccd === dx.nv.cccd; })) cb.push('Số CCCD đọc được trùng với CBCNV khác.');
      log(s, 'tl.ocr', o.maNV + ' ' + o.loai + ': ' + o.tenFile + ' (' + nguon + ', ' + text.length + ' ký tự, đề xuất ' + Object.keys(dx.nv).length + ' ô)');   // không ghi dữ liệu cá nhân vào nhật ký
      return { tl: o, text: text, nguon: nguon, deXuat: dx, hienTai: hien, tlHien: tlH, canhBao: cb, sua: can(s.role, 'edit_nv') };
    };
    H['tl.ocrxem'] = function (req, s) {
      var o = T('TaiLieu').filter(function (x) { return x.id === req.id; })[0]; if (!o) err('Không tìm thấy tài liệu.');
      tlQuyen(o, s); if (!o.ocrId) err('Tài liệu chưa được đọc chữ.');
      var f = env.files.get(o.ocrId); if (!f) err('Không đọc được tệp văn bản OCR.');
      log(s, 'tl.ocrxem', o.maNV + ' ' + o.loai + ': ' + o.tenFile);
      return { data: f.data, ten: o.tenFile, nguon: o.ocrNguon, luc: o.ocrLuc };
    };
    H['tl.ocrluu'] = function (req, s) {
      need(s, 'edit_nv');
      var o = T('TaiLieu').filter(function (x) { return x.id === req.id; })[0]; if (!o) err('Không tìm thấy tài liệu.');
      var nv = findNV(o.maNV); if (!nv) err('Không tìm thấy nhân viên.');
      var g = req.gt || {}, t = req.tl || {}, doi = [], truoc = {}, sau = {}, v;
      Object.keys(g).forEach(function (k) {
        var dn = OCR_TRUONG.filter(function (x) { return x[0] === k; })[0]; if (!dn) err('Trường không hợp lệ: ' + k);
        v = String(g[k] == null ? '' : g[k]).trim().slice(0, 250);
        if (k === 'cccd' && v && !/^\d{12}$/.test(v)) err('Số CCCD phải đủ 12 chữ số.');
        if (k === 'cccd' && v && T('NhanVien').some(function (x) { return x.maNV !== nv.maNV && x.cccd === v; })) err('Số CCCD ' + v + ' đã có ở CBCNV khác.');
        if (k === 'mst' && v && !/^\d{10}(-\d{3})?$/.test(v)) err('Mã số thuế phải 10 số (hoặc 10 số - 3 số).');
        if ((k === 'ngaySinh' || k === 'ngayCapCCCD') && v && (!parseD(v) || v > env.today())) err(dn[1] + ' không hợp lệ: ' + v);
        if (k === 'gioiTinh' && v && ['Nam', 'Nữ'].indexOf(v) < 0) err('Giới tính chỉ Nam hoặc Nữ.');
        if (String(nv[k] || '') !== v) { truoc[k] = nv[k] || ''; nv[k] = v; sau[k] = v; doi.push(dn[1].toLowerCase()); }
      });
      Object.keys(t).forEach(function (k) {
        var dn = OCR_TL.filter(function (x) { return x[0] === k; })[0]; if (!dn) err('Trường không hợp lệ: ' + k);
        v = String(t[k] == null ? '' : t[k]).trim().slice(0, 100);
        if (dn[2] === 'date' && v && !parseD(v)) err(dn[1] + ' không hợp lệ: ' + v);
        if (String(o[k] || '') !== v) { o[k] = v; doi.push(dn[1].toLowerCase()); }
      });
      if (Object.keys(sau).length) { nv.capNhat = env.now(); D('NhanVien'); logDoi(s, 'tl.ocrluu', nv.maNV, truoc, sau); }
      D('TaiLieu');
      log(s, 'tl.ocrluu', nv.maNV + ' – cập nhật từ OCR "' + o.ten + '" (người soát): ' + (doi.join(', ') || 'không đổi'));
      return { nv: nv, tl: o, doi: doi };
    };
    H['tl.delete'] = function (req, s) {
      var a = T('TaiLieu'), o = a.filter(function (x) { return x.id === req.id; })[0]; if (!o) err('Không tìm thấy tài liệu.');
      if (!can(s.role, 'edit_nv') && !(o.maNV === s.maNV && o.trangThai === 'Chờ xác nhận')) err('Không có quyền xóa tài liệu này.');
      try { env.files.del(o.fileId); } catch (e) { }
      if (o.ocrId) { try { env.files.del(o.ocrId); } catch (e) { } }          // v3.37: xóa cả tệp văn bản OCR
      a.splice(a.indexOf(o), 1); D('TaiLieu');
      log(s, 'tl.delete', o.maNV + ' ' + o.loai + ': ' + o.tenFile);
      return true;
    };

    /* ====================== AI ĐỌC ẢNH CCCD → ĐỀ XUẤT ĐIỀN LÝ LỊCH (v3.11) ======================
       AI chỉ đề xuất; TCHC soát từng ô cạnh ảnh rồi mới lưu (nv.cccdluu). Không đọc rõ → để trống, không đoán.
       Số CCCD 12 số: 3 số đầu mã tỉnh nơi đăng ký khai sinh, số thứ 4 = giới tính + thế kỷ sinh (0/1: nam/nữ 1900–1999,
       2/3: nam/nữ 2000–2099), số thứ 5–6 = 2 số cuối năm sinh → đối chiếu chéo với ngày sinh, giới tính để báo đọc sai. */
    var CCCD_HE = 'Bạn đọc ảnh thẻ Căn cước công dân / thẻ Căn cước của Việt Nam (mặt trước và/hoặc mặt sau) để giúp phòng Tổ chức – Hành chính cập nhật lý lịch nhân sự. '
      + 'QUY TẮC BẮT BUỘC: chỉ ghi thông tin IN TRÊN THẺ mà bạn đọc rõ; chữ mờ, bị che, lóa sáng hoặc không có trên thẻ thì để chuỗi rỗng "" – tuyệt đối không suy đoán, không tự bổ sung từ hiểu biết chung '
      + '(kể cả tên cơ quan cấp thường gặp). Số CCCD chép đúng từng chữ số. Ngày ghi dạng yyyy-MM-dd. Giới tính chỉ "Nam" hoặc "Nữ". '
      + 'Ngày cấp, nơi cấp thường ở mặt sau; nếu chỉ có mặt trước thì để rỗng. Nếu ảnh không phải thẻ CCCD / Căn cước thì để rỗng tất cả và ghi lý do vào ghiChu. '
      + 'Chỉ trả về JSON: {"hoTen":"","soCCCD":"","ngaySinh":"","gioiTinh":"","queQuan":"","noiThuongTru":"","ngayCap":"","noiCap":"","coGiaTriDen":"","doTinCay":"cao|trung bình|thấp","ghiChu":""}';
    var CCCD_TRUONG = ['cccd', 'ngaySinh', 'gioiTinh', 'ngayCapCCCD', 'noiCapCCCD', 'hetHanCCCD', 'queQuan', 'diaChi'];
    function cccdBoDau(t) { return String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().replace(/[^a-z]/g, ''); }
    function cccdNgay(v, toiDa) { v = aiS(v, 10); if (!parseD(v)) return ''; return toiDa && v > toiDa ? '' : v; }
    function cccdKiem(g, hoTenHS) {                // đối chiếu chéo – trả về danh sách cảnh báo (không tự sửa)
      var cb = [], so = g.cccd || '';
      if (so && so.length === 12) {
        var d = +so.charAt(3), yy = so.slice(4, 6), the = d <= 1 ? '19' : d <= 3 ? '20' : d <= 5 ? '21' : '';
        if (d > 9 || !the) cb.push('Chữ số thứ 4 của số CCCD (' + so.charAt(3) + ') không đúng quy ước giới tính – thế kỷ.');
        else {
          var gt = d % 2 === 0 ? 'Nam' : 'Nữ';
          if (g.gioiTinh && g.gioiTinh !== gt) cb.push('Số CCCD cho biết giới tính ' + gt + ', nhưng đọc được giới tính "' + g.gioiTinh + '".');
          if (g.ngaySinh && g.ngaySinh.slice(0, 4) !== the + yy) cb.push('Số CCCD cho biết năm sinh ' + the + yy + ', nhưng đọc được ngày sinh ' + ngayVN(g.ngaySinh) + '.');
        }
      }
      if (g.ngayCapCCCD && g.ngaySinh && g.ngayCapCCCD <= g.ngaySinh) cb.push('Ngày cấp không thể trước ngày sinh.');
      if (g.hetHanCCCD && g.ngayCapCCCD && g.hetHanCCCD <= g.ngayCapCCCD) cb.push('Ngày hết hạn không thể trước ngày cấp.');
      if (g.hetHanCCCD && g.hetHanCCCD < env.today()) cb.push('Thẻ đã hết hạn từ ' + ngayVN(g.hetHanCCCD) + ' – đề nghị CBCNV làm lại thẻ.');
      if (g.hoTen && hoTenHS && cccdBoDau(g.hoTen) !== cccdBoDau(hoTenHS)) cb.push('Họ tên trên thẻ "' + g.hoTen + '" khác họ tên trong hồ sơ "' + hoTenHS + '" – kiểm tra lại có đúng ảnh của người này không.');
      return cb;
    }
    function cccdAnhHS(ma) {                        // ảnh mới nhất mỗi mặt
      var ds = T('TaiLieu').filter(function (x) { return x.maNV === ma && x.loai === 'Căn cước công dân'; }), out = [];
      ['truoc', 'sau'].forEach(function (m) {
        var x = ds.filter(function (y) { return y.ten === CCCD_TEN[m]; }).sort(function (a, b) { return a.taiLuc < b.taiLuc ? 1 : -1; })[0];
        if (x) out.push({ mat: m, o: x });
      });
      return out;
    }
    H['nv.cccdai'] = function (req, s) {
      need(s, 'edit_nv');
      var nv = findNV(req.maNV); if (!nv) err('Không tìm thấy nhân viên.');
      var anh = cccdAnhHS(nv.maNV); if (!anh.length) err('Chưa có ảnh CCCD. Tải ảnh mặt trước, mặt sau trước khi cho AI đọc.');
      var ds = anh.map(function (a) {
        if (!/^image\//.test(a.o.mimeType || '')) err('Ảnh ' + a.o.ten + ' không phải ảnh.');
        var f = env.files.get(a.o.fileId); if (!f || !f.data) err('Không đọc được ảnh ' + a.o.ten + ' (có thể đã bị xóa khỏi Drive).');
        return { nhan: (a.mat === 'truoc' ? 'Ảnh MẶT TRƯỚC thẻ:' : 'Ảnh MẶT SAU thẻ:'), mime: a.o.mimeType, data: f.data };
      });
      var o = aiGoi({ he: CCCD_HE, nd: '=== ĐỌC CCCD ===\nĐọc ' + ds.length + ' ảnh thẻ dưới đây (' + anh.map(function (a) { return a.mat === 'truoc' ? 'mặt trước' : 'mặt sau'; }).join(', ') + ') và trả JSON theo đúng mẫu.',
        anh: ds, maxTok: 700, loai: 'cccd', duLieu: { mat: anh.map(function (a) { return a.mat; }), maNV: nv.maNV } }) || {};
      var so = String(o.soCCCD || '').replace(/\D/g, ''), cb = [];
      if (o.soCCCD && so.length !== 12) { cb.push('Số đọc được "' + aiS(o.soCCCD, 30) + '" không đủ 12 chữ số – không đề xuất.'); so = ''; }
      var g = {
        hoTen: aiS(o.hoTen, 120), cccd: so, ngaySinh: cccdNgay(o.ngaySinh, env.today()),
        gioiTinh: ['Nam', 'Nữ'].indexOf(aiS(o.gioiTinh, 5)) >= 0 ? aiS(o.gioiTinh, 5) : '',
        ngayCapCCCD: cccdNgay(o.ngayCap, env.today()), noiCapCCCD: aiS(o.noiCap, 250), hetHanCCCD: cccdNgay(o.coGiaTriDen),
        queQuan: aiS(o.queQuan, 250), diaChi: aiS(o.noiThuongTru, 250)
      };
      if (o.ngayCap && !g.ngayCapCCCD) cb.push('Ngày cấp đọc được không hợp lệ hoặc ở tương lai – không đề xuất.');
      cb = cb.concat(cccdKiem(g, nv.hoTen));
      var hien = {}; CCCD_TRUONG.forEach(function (k) { hien[k] = k === 'hetHanCCCD' ? ((anh[0] || {}).o || {}).hetHan || '' : String(nv[k] || ''); }); hien.hoTen = nv.hoTen;
      var soO = CCCD_TRUONG.filter(function (k) { return g[k]; }).length;
      log(s, 'nv.cccdai', nv.maNV + ' – AI đọc ' + ds.length + ' ảnh CCCD, đề xuất ' + soO + ' ô' + (cb.length ? ', ' + cb.length + ' cảnh báo' : ''));   // không ghi số liệu cá nhân vào nhật ký
      return { maNV: nv.maNV, ai: g, hienTai: hien, canhBao: cb, soO: soO, soAnh: ds.length, mat: anh.map(function (a) { return a.mat; }),
        doTinCay: ['cao', 'trung bình', 'thấp'].indexOf(String(o.doTinCay || '').toLowerCase()) >= 0 ? String(o.doTinCay).toLowerCase() : '', ghiChu: aiS(o.ghiChu, 300), demo: !!(o.demo) };
    };
    H['nv.cccdluu'] = function (req, s) {
      need(s, 'edit_nv');
      var nv = findNV(req.maNV); if (!nv) err('Không tìm thấy nhân viên.');
      var g = req.gt || {}, doi = [], v;
      Object.keys(g).forEach(function (k) {
        if (CCCD_TRUONG.indexOf(k) < 0) err('Trường không hợp lệ: ' + k);
        v = String(g[k] == null ? '' : g[k]).trim();
        if (k === 'cccd' && v && !/^\d{12}$/.test(v)) err('Số CCCD phải đủ 12 chữ số.');
        if ((k === 'ngaySinh' || k === 'ngayCapCCCD') && v && (!parseD(v) || v > env.today())) err('Ngày không hợp lệ: ' + v);
        if (k === 'hetHanCCCD' && v && !parseD(v)) err('Ngày hết hạn không hợp lệ.');
        if (k === 'gioiTinh' && v && ['Nam', 'Nữ'].indexOf(v) < 0) err('Giới tính chỉ Nam hoặc Nữ.');
        v = v.slice(0, 250);
      });
      var ten = { cccd: 'số CCCD', ngaySinh: 'ngày sinh', gioiTinh: 'giới tính', ngayCapCCCD: 'ngày cấp', noiCapCCCD: 'nơi cấp', hetHanCCCD: 'hạn thẻ', queQuan: 'quê quán', diaChi: 'địa chỉ' }, tl = [];
      Object.keys(g).forEach(function (k) {
        v = String(g[k] == null ? '' : g[k]).trim().slice(0, 250);
        if (k === 'hetHanCCCD') { T('TaiLieu').forEach(function (x) { if (x.maNV === nv.maNV && x.loai === 'Căn cước công dân' && x.hetHan !== v) { x.hetHan = v; tl.push(x); } }); if (tl.length) doi.push(ten[k]); return; }
        if (String(nv[k] || '') !== v) { nv[k] = v; doi.push(ten[k]); }
      });
      if (g.cccd || g.ngayCapCCCD) T('TaiLieu').forEach(function (x) { if (x.maNV === nv.maNV && x.loai === 'Căn cước công dân' && /^CCCD – mặt/.test(x.ten)) { var c = false; if (g.cccd && x.soHieu !== nv.cccd) { x.soHieu = nv.cccd; c = true; } if (g.ngayCapCCCD && x.ngay !== nv.ngayCapCCCD) { x.ngay = nv.ngayCapCCCD; c = true; } if (c && tl.indexOf(x) < 0) tl.push(x); } });
      if (doi.length) { nv.capNhat = env.now(); D('NhanVien'); }
      if (tl.length) D('TaiLieu');
      log(s, 'nv.cccdluu', nv.maNV + ' – cập nhật từ ảnh CCCD (AI đọc, người soát): ' + (doi.join(', ') || 'không đổi'));
      return { nv: nv, tl: tl, doi: doi };
    };

    /* ====================== SỔ CÔNG VĂN ĐẾN – ĐI ======================
       Dữ liệu bản quét lưu trên Drive Công ty (cùng tài khoản chứa file Google Sheet),
       thư mục "_CÔNG VĂN / Đến | Đi / <năm>". Mọi lần xem đều ghi nhật ký. */
    function cvDuong(huong, nam) { return [TM_CV, huong === 'di' ? 'Đi' : 'Đến', String(nam || env.today().slice(0, 4))]; }
    function cvRow(id) { var o = T('CongVan').filter(function (x) { return x.id === id; })[0]; if (!o) err('Không tìm thấy công văn.'); return o; }
    function cvNguoi(o) { return parseJSON(o.nguoiXuLy, []).concat(parseJSON(o.phoiHop, [])); }
    function cvXem(s, o) {
      return can(s.role, 'congvan') || laNguoiXuLy(s, cvNguoi(o)) || o.nguoiNhap === (s.hoTen || s.username)
        || laNguoiXuLy(s, cvNguoiDoc(o));      // chỉ đọc được sau khi Giám đốc duyệt chia sẻ
    }
    function cvGiaoDuoc(s) { return can(s.role, 'vanthu') || s.role === 'giamdoc' || s.role === 'admin'; }
    function cvSoKeTiep(huong, nam) {
      return T('CongVan').filter(function (x) { return x.huong === huong && String(x.nam) === String(nam); })
        .reduce(function (m, x) { return Math.max(m, num(x.so)); }, 0) + 1;
    }
    function cvTen(o) { return CV_HUONG[o.huong] + ' số ' + o.so + '/' + o.nam + (o.soHieu ? ' (' + o.soHieu + ')' : ''); }
    function cvTenTep(o, ten) { return o.nam + '-' + ('000' + num(o.so)).slice(-3) + (o.soHieu ? '_' + String(o.soHieu).replace(/[\\/:*?"<>|]/g, '-') : '') + '_' + ten; }
    function cvGhiXL(o, hanhDong, s, noiDung, han, tep) {
      var r = { id: newId(), cvId: o.id, hanhDong: hanhDong, nguoi: s.username, tenNguoi: s.hoTen || s.username,
        noiDung: String(noiDung || '').slice(0, 3000), hanXuLy: han || '', fileId: (tep && tep.fileId) || '', tenTep: (tep && tep.ten) || '', thoiGian: env.now() };
      T('CongVanXL').push(r); D('CongVanXL');
      return r;
    }
    function cvTepVao(o, f, nhan) {                  // kiểm tra và ghi tệp vào thư mục công văn của năm
      var cfg = getCfg(), ten = String(f.name || '').trim().replace(/[\\/:*?"<>|]/g, '_').slice(0, 150);
      if (!ten || !f.data) err('Chưa chọn tệp.');
      if (!MIME_OK.test(f.mime || '')) err('Chỉ nhận PDF, ảnh (JPG, PNG, WEBP, HEIC), Word hoặc Excel.');
      var bytes = Math.floor(String(f.data).length * 3 / 4);
      if (bytes > (num(cfg.dungLuongToiDaMB) || 10) * 1048576) err('Tệp vượt ' + (num(cfg.dungLuongToiDaMB) || 10) + ' MB.');
      var id = env.files.put({ name: (nhan ? nhan + '_' : '') + cvTenTep(o, ten), mime: f.mime, data: f.data, duong: cvDuong(o.huong, o.nam) });
      return { fileId: id, ten: ten, mime: f.mime, kichThuoc: bytes };
    }
    function cvQuaHan(o) { return o.hanXuLy && o.hanXuLy < env.today() && ['Đã giao xử lý', 'Đang xử lý'].indexOf(o.trangThai) >= 0; }
    H['cv.ocr'] = function (req, s) {                 // v3.24: chữ nhận dạng của 1 văn bản
      var o = T('CongVan').filter(function (x) { return x.id === req.id; })[0];
      if (!o || !cvXem(s, o)) err('Không tìm thấy văn bản hoặc bạn không có quyền xem.');
      return { id: o.id, chu: o.noiDungOCR || '' };
    };
    H['cv.ds'] = function (req, s) {
      var ds = T('CongVan').filter(function (o) { return cvXem(s, o); });
      var co = {}; ds.forEach(function (o) { co[o.id] = 1; });
      var nam = String(req.nam || env.today().slice(0, 4)), url = '';
      if (can(s.role, 'congvan') && env.files.duongUrl) { try { url = env.files.duongUrl([TM_CV]); } catch (e) { url = ''; } }
      var nhe = ds.map(function (o) { if (!o.noiDungOCR) return o; var x = Object.assign({}, o); x.soChuOCR = String(o.noiDungOCR).length; x.ocrTim = String(o.noiDungOCR).slice(0, 1500); delete x.noiDungOCR; return x; });   // v3.42: đoạn đầu chữ quét để tìm   // v3.24: chữ OCR tải riêng khi mở văn bản
      return { ds: nhe, xl: T('CongVanXL').filter(function (x) { return co[x.cvId]; }),
        nam: nam, soDen: cvSoKeTiep('den', nam), soDi: cvSoKeTiep('di', nam),
        vanThu: can(s.role, 'vanthu'), giaoDuoc: cvGiaoDuoc(s), duyetCS: cvDuyetCSDuoc(s),
        coAI: !!(env.ai && env.ai.trangThai && env.ai.trangThai().coKhoa), toanBo: can(s.role, 'congvan'), thuMucUrl: url };
    };
    H['cv.save'] = function (req, s) {
      need(s, 'vanthu');
      var r = req.row || {}, moi = !req.id, o;
      if (moi) {
        var huong = r.huong === 'di' ? 'di' : 'den', nam = String(r.nam || String(r.ngayVaoSo || env.today()).slice(0, 4));
        if (!/^\d{4}$/.test(nam)) err('Năm vào sổ không hợp lệ.');
        o = { id: newId(), huong: huong, nam: nam, so: cvSoKeTiep(huong, nam), ngayVaoSo: r.ngayVaoSo || env.today(),
          trangThai: huong === 'di' ? 'Chờ ban hành' : 'Mới đăng ký', nguoiXuLy: '[]', phoiHop: '[]',
          csTrangThai: 'Chưa trình', csDeXuat: '[]', nguoiDoc: '[]', soTN: 0,
          nguoiNhap: s.hoTen || s.username, taoLuc: env.now() };
      } else {
        o = cvRow(req.id);
        if (o.trangThai === CV_HET && s.role !== 'admin') err('Công văn đã lưu hồ sơ, chỉ Quản trị sửa được.');
      }
      ['soHieu', 'ngayVB', 'ngayVaoSo', 'noiGui', 'noiNhan', 'loai', 'doKhan', 'trichYeu', 'nguoiKy', 'ghiChu', 'lienKet', 'hanTraLoi'].forEach(function (k) {
        if (r[k] !== undefined) o[k] = String(r[k] || '').trim().slice(0, 600);
      });
      if (r.tomTat !== undefined) o.tomTat = String(r.tomTat || '').trim().slice(0, 2000);
      if (r.noiDungOCR !== undefined) o.noiDungOCR = String(r.noiDungOCR || '').slice(0, 45000);   // v3.24: chữ nhận dạng từ bản quét (tra cứu)
      if (r.aiTinCay !== undefined) { o.aiTinCay = String(r.aiTinCay || '').slice(0, 20); o.aiLuc = o.aiTinCay ? env.now() : ''; }
      if (o.hanTraLoi && !parseD(o.hanTraLoi)) err('Hạn trả lời ghi trong văn bản không hợp lệ.');
      if (r.soBan !== undefined) o.soBan = Math.max(0, num(r.soBan));
      o.loai = o.loai || 'Công văn'; o.doKhan = o.doKhan || 'Thường';
      if (CV_LOAI.indexOf(o.loai) < 0) err('Loại văn bản không hợp lệ.');
      if (CV_KHAN.indexOf(o.doKhan) < 0) err('Độ khẩn không hợp lệ.');
      ['ngayVB', 'ngayVaoSo'].forEach(function (k) { if (o[k] && !parseD(o[k])) err('Ngày không hợp lệ.'); });
      if (!String(o.trichYeu || '').trim()) err('Nhập trích yếu nội dung văn bản.');
      if (o.huong === 'den' && !String(o.noiGui || '').trim()) err('Nhập nơi gửi (cơ quan, đơn vị gửi văn bản đến).');
      if (o.huong === 'di' && !String(o.noiNhan || '').trim()) err('Nhập nơi nhận văn bản đi.');
      if (req.so !== undefined && num(req.so) > 0 && num(req.so) !== num(o.so)) {
        var tr = T('CongVan').filter(function (x) { return x.id !== o.id && x.huong === o.huong && String(x.nam) === String(o.nam) && num(x.so) === num(req.so); })[0];
        if (tr) err('Số ' + req.so + '/' + o.nam + ' đã dùng cho văn bản "' + tr.trichYeu + '".');
        o.so = num(req.so);
      }
      if (req.trangThai && !moi) {
        var hl = (o.huong === 'di' ? CV_TT_DI : CV_TT_DEN);
        if (hl.indexOf(req.trangThai) < 0) err('Trạng thái không hợp lệ.');
        o.trangThai = req.trangThai;
      }
      o.capNhat = env.now();
      if (moi) T('CongVan').push(o);
      D('CongVan');
      log(s, 'cv.save', (moi ? 'đăng ký ' : 'sửa ') + cvTen(o) + ': ' + o.trichYeu);
      if (moi) cvGhiXL(o, 'Đăng ký vào sổ', s, o.trichYeu, '', null);
      return o;
    };
    H['cv.upload'] = function (req, s) {
      need(s, 'vanthu');
      var o = cvRow(req.id);
      if (o.trangThai === CV_HET && s.role !== 'admin') err('Công văn đã lưu hồ sơ.');
      var t = cvTepVao(o, req.file || {}, '');
      if (o.fileId) { try { env.files.del(o.fileId); } catch (e) { } }
      o.fileId = t.fileId; o.tenFile = t.ten; o.mime = t.mime; o.kichThuoc = t.kichThuoc; o.capNhat = env.now();
      if (o.huong === 'di' && o.trangThai === 'Chờ ban hành') o.trangThai = 'Đã ban hành';
      D('CongVan');
      log(s, 'cv.upload', cvTen(o) + ': ' + t.ten);
      return o;
    };
    H['cv.get'] = function (req, s) {
      var o = cvRow(req.id);
      if (!cvXem(s, o)) err('Bạn không có quyền xem văn bản này.');
      var fid = o.fileId, ten = o.tenFile, mime = o.mime;
      if (req.xlId) {                                 // tệp kèm theo một lần báo cáo kết quả
        var x = T('CongVanXL').filter(function (y) { return y.id === req.xlId && y.cvId === o.id; })[0];
        if (!x || !x.fileId) err('Không tìm thấy tệp.');
        fid = x.fileId; ten = x.tenTep; mime = '';
      }
      if (req.tnId) {                                 // tệp đính kèm một lời trao đổi
        var m = T('CongVanTN').filter(function (y) { return y.id === req.tnId && y.cvId === o.id; })[0];
        if (!m || !m.fileId) err('Không tìm thấy tệp.');
        fid = m.fileId; ten = m.tenTep; mime = m.mime || '';
      }
      if (!fid) err('Văn bản này chưa có bản quét trên hệ thống.');
      var f = env.files.get(fid); if (!f) err('Không đọc được tệp (có thể đã bị xóa khỏi Drive).');
      log(s, 'cv.xem', cvTen(o) + (req.xlId ? ' – tệp kết quả' : ''));
      return { name: ten, mime: mime || 'application/pdf', data: f.data };
    };
    H['cv.giao'] = function (req, s) {
      var o = cvRow(req.id);
      if (!cvGiaoDuoc(s)) err('Chỉ Giám đốc, TCHC hoặc Quản trị giao xử lý công văn.');
      if (o.trangThai === CV_HET) err('Công văn đã lưu hồ sơ.');
      var xl = (req.nguoiXuLy || []).slice(0, 10), ph = (req.phoiHop || []).slice(0, 20);
      kiemNguoi(xl, false, 'Người xử lý chính');
      if (ph.length) kiemNguoi(ph, false, 'Người phối hợp, theo dõi');
      var han = String(req.hanXuLy || '').trim();
      if (han && !parseD(han)) err('Hạn xử lý không hợp lệ.');
      if (han && han < env.today()) err('Hạn xử lý đã qua, chọn lại ngày.');
      var but = String(req.butPhe || '').trim().slice(0, 2000);
      if (!but) err('Nhập bút phê / nội dung giao việc.');
      o.butPhe = but; o.nguoiXuLy = JSON.stringify(xl); o.phoiHop = JSON.stringify(ph); o.hanXuLy = han;
      o.trangThai = 'Đã giao xử lý'; o.capNhat = env.now(); D('CongVan');
      var x = cvGhiXL(o, 'Giao xử lý', s, but + ' → ' + tenNguoi(xl) + (ph.length ? ' · phối hợp: ' + tenNguoi(ph) : ''), han, null);
      log(s, 'cv.giao', cvTen(o) + ' → ' + tenNguoi(xl) + (han ? ', hạn ' + han : ''));
      bao(nguoiCo(xl.concat(ph), s.username), { loai: 'duyet', t: 'Công văn chờ bạn xử lý',
        b: cvTen(o) + ': ' + o.trichYeu + (han ? ' · hạn ' + ngayVN(han) : ''), tag: 'cv_' + o.id, mo: '#v=cv' });
      return { cv: o, xl: x };
    };
    H['cv.baocao'] = function (req, s) {
      var o = cvRow(req.id);
      if (!cvXem(s, o)) err('Văn bản này không giao cho bạn.');
      if (o.trangThai === CV_HET) err('Công văn đã lưu hồ sơ.');
      var nd = String(req.noiDung || '').trim().slice(0, 4000);
      if (!nd) err('Nhập nội dung đã xử lý / kết quả.');
      var tep = req.file && req.file.data ? cvTepVao(o, req.file, 'KQ') : null;
      o.trangThai = req.xong ? 'Đã xử lý' : 'Đang xử lý';
      if (req.xong) { o.ketQua = nd; o.hoanThanhLuc = env.now(); }
      o.capNhat = env.now(); D('CongVan');
      var x = cvGhiXL(o, req.xong ? 'Báo cáo kết quả' : 'Cập nhật tiến độ', s, nd, '', tep);
      log(s, 'cv.baocao', cvTen(o) + ' → ' + o.trangThai);
      bao(nguoiCo(['r:tchc', 'r:giamdoc'], s.username), { loai: 'duyet', t: req.xong ? 'Công văn đã xử lý xong' : 'Công văn có cập nhật',
        b: cvTen(o) + ' · ' + (s.hoTen || s.username) + ': ' + nd.slice(0, 120), tag: 'cv_' + o.id, mo: '#v=cv' });
      return { cv: o, xl: x };
    };
    /* ---- AI đọc bản quét công văn, bóc các trường để vào sổ (v3.7) ----
       Ảnh gửi thẳng cho AI; PDF / Word chuyển tạm sang Google Docs để lấy chữ rồi xóa bản tạm ngay.
       AI CHỈ bóc thông tin có trong văn bản — máy chủ kiểm lại từng trường trước khi trả về,
       và văn thư vẫn phải soát trước khi lưu: đây là điền hộ, không phải vào sổ tự động. */
    var CV_HE = 'Bạn là văn thư của Công ty Cổ phần TAPACO Việt Nam (xây dựng hạ tầng). ' +
      'Nhiệm vụ: đọc một văn bản hành chính và bóc ra các trường để vào sổ công văn.\n' +
      'QUY TẮC BẮT BUỘC:\n' +
      '- CHỈ lấy thông tin CÓ TRONG văn bản. Tuyệt đối không suy đoán, không điền thông tin không nhìn thấy.\n' +
      '- Trường nào không thấy thì để chuỗi rỗng "". Thà để trống còn hơn đoán sai.\n' +
      '- Mọi ngày tháng trả về dạng yyyy-MM-dd. Không thấy đủ ngày tháng năm thì để rỗng.\n' +
      '- noiGui: cơ quan, đơn vị PHÁT HÀNH văn bản (thường ở góc trên bên trái và ở phần ký tên, đóng dấu).\n' +
      '- noiNhan: nơi nhận (thường sau chữ "Kính gửi:"). Nhiều nơi nhận thì ghi gọn, cách nhau bằng dấu phẩy.\n' +
      '- soHieu: số và ký hiệu văn bản, ví dụ "1234/UBND-GT".\n' +
      '- hanTraLoi: hạn mà văn bản yêu cầu trả lời, báo cáo hoặc thực hiện, CHỈ khi văn bản nêu rõ.\n' +
      '- trichYeu: đúng câu trích yếu "V/v …" của văn bản, giữ nguyên văn, không viết lại.\n' +
      '- tomTat: 2–4 câu nêu đúng việc văn bản yêu cầu và ai phải làm gì. Không thêm nhận định, không khuyến nghị.\n' +
      '- doTinCay: "cao" nếu chữ rõ và bóc được hầu hết trường; "thấp" nếu bản quét mờ, thiếu nét, đọc đoán nhiều.\n' +
      'Trả về DUY NHẤT một JSON, không kèm lời dẫn, không rào ```:\n' +
      '{"soHieu":"","ngayVB":"","noiGui":"","noiNhan":"","loai":"","doKhan":"","trichYeu":"","nguoiKy":"","hanTraLoi":"","tomTat":"","doTinCay":""}\n' +
      'loai chọn đúng một trong: ' + CV_LOAI.join(' | ') + '\n' +
      'doKhan chọn đúng một trong: ' + CV_KHAN.join(' | ');
    function cvNgay(v) { v = aiS(v, 10); return parseD(v) ? v : ''; }
    /* v3.24: ảnh, PDF → nhận dạng chữ OCR tiếng Việt (Google Drive chuyển sang Google Tài liệu, ocrLanguage=vi) → AI phân tích CHỮ đã nhận dạng.
       Không nhận dạng được (ảnh HEIC/WEBP, ảnh quá mờ…) thì ảnh gửi thẳng cho AI đọc như trước. */
    var OCR_IMG = /^image\/(jpeg|jpg|png|gif|bmp)$/i, OCR_TOI_THIEU = 40;
    function lamSachChu(t) { return String(t || '').replace(/\r/g, '').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim(); }
    function ocrDrive(fileId) {                       // trả chữ nhận dạng; lỗi → ''
      if (env.ocr) return env.ocr(fileId);
      var r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) + '/copy?supportsAllDrives=true&ocrLanguage=vi', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({ name: 'tam-ocr-' + Date.now(), mimeType: 'application/vnd.google-apps.document' })
      });
      if (r.getResponseCode() >= 300) return '';
      var id = (parseJSON(r.getContentText(), {}) || {}).id; if (!id) return '';
      try {
        var x = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id) + '/export?mimeType=text%2Fplain', { muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
        return x.getResponseCode() < 300 ? x.getContentText('UTF-8') : '';
      } finally { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { } }
    }
    function cvChuTep(f) {                           // trả về {chu, anh, nguon} để đưa cho AI
      var mime = String(f.mime || ''), anh = /^image\//.test(mime), pdf = /pdf/i.test(mime);
      var anhAI = function () { return { chu: '', anh: [{ nhan: 'Bản quét văn bản cần bóc thông tin:', mime: mime, data: f.data }], nguon: 'anh' }; };
      if (typeof UrlFetchApp === 'undefined' && !(env.ocr && (anh || pdf))) return anh ? anhAI() : { chu: '', anh: [], demo: true };   // bản dùng thử trên trình duyệt
      if (!env.files || !env.files.put) err('Máy chủ chưa bật lưu trữ Drive.');
      if (anh && !OCR_IMG.test(mime)) return anhAI();                 // HEIC, WEBP: Drive không nhận dạng được → AI đọc ảnh
      var id = env.files.put({ name: 'tam-doc-cv-' + Date.now() + '-' + String(f.name || 'vb'), mime: mime, data: f.data, duong: [TM_CV, '_tam-doc'] });
      var chu = '';
      try { chu = anh || pdf ? ocrDrive(id) : botTrichChu(id, mime); } catch (e) { chu = ''; }
      finally { try { env.files.del(id); } catch (e) { } }
      chu = lamSachChu(chu);
      if (chu.replace(/\s/g, '').length >= OCR_TOI_THIEU) return { chu: chu.slice(0, 40000), anh: [], nguon: anh || pdf ? 'ocr' : 'word' };
      if (anh) return anhAI();
      err(pdf ? 'Không nhận dạng được chữ trong PDF (bản quét quá mờ hoặc tệp có mật khẩu). Thử chụp lại rõ hơn, hoặc tải lên ảnh JPG/PNG.' : 'Không đọc được chữ trong tệp Word.');
    }
    H['cv.ai'] = function (req, s) {
      need(s, 'vanthu');
      var f = req.file || {}, cfg = getCfg();
      if (!req.chu) {
      if (!f.data) err('Chưa chọn tệp.');
      if (!MIME_OK.test(f.mime || '')) err('Chỉ đọc được PDF, ảnh (JPG, PNG, WEBP, HEIC) hoặc Word.');
      var bytes = Math.floor(String(f.data).length * 3 / 4);
      if (bytes > (num(cfg.dungLuongToiDaMB) || 10) * 1048576) err('Tệp vượt ' + (num(cfg.dungLuongToiDaMB) || 10) + ' MB.');
      }
      var huong = req.huong === 'di' ? 'di' : 'den';
      var t = req.chu ? { chu: lamSachChu(req.chu).slice(0, 40000), anh: [], nguon: 'ocr-sua' } : cvChuTep(f);   // v3.24: văn thư sửa chữ OCR rồi bấm phân tích lại
      var nd = '=== VÀO SỔ CÔNG VĂN ===\n'
        + 'Đây là ' + (huong === 'di' ? 'văn bản do Công ty TAPACO phát hành (công văn đi)' : 'văn bản gửi đến Công ty TAPACO (công văn đến)') + '.\n\n'
        + (t.chu ? (/^ocr/.test(t.nguon) ? '=== CHỮ NHẬN DẠNG (OCR) TỪ BẢN QUÉT – có thể sai vài ký tự, dấu tiếng Việt, ngắt dòng; hãy suy luận đúng nghĩa, không bịa thêm ===\n' : '=== NỘI DUNG VĂN BẢN ===\n') + t.chu : 'Nội dung văn bản nằm trong ảnh đính kèm.');
      var o = aiGoi({ he: CV_HE, nd: nd, anh: t.anh, maxTok: 1600, loai: 'congvan', duLieu: { demo: !!t.demo, ten: String(f.name || ''), huong: huong } }) || {};
      var kq = {
        soHieu: aiS(o.soHieu, 100), ngayVB: cvNgay(o.ngayVB), noiGui: aiS(o.noiGui, 250), noiNhan: aiS(o.noiNhan, 250),
        loai: CV_LOAI.indexOf(aiS(o.loai, 40)) >= 0 ? aiS(o.loai, 40) : '',
        doKhan: CV_KHAN.indexOf(aiS(o.doKhan, 20)) >= 0 ? aiS(o.doKhan, 20) : '',
        trichYeu: aiS(o.trichYeu, 600), nguoiKy: aiS(o.nguoiKy, 120),
        hanTraLoi: cvNgay(o.hanTraLoi), tomTat: aiS(o.tomTat, 1500),
        doTinCay: ['cao', 'trung bình', 'thấp'].indexOf(String(o.doTinCay || '').toLowerCase()) >= 0 ? String(o.doTinCay).toLowerCase() : ''
      };
      kq.soO = ['soHieu', 'ngayVB', 'noiGui', 'noiNhan', 'loai', 'doKhan', 'trichYeu', 'nguoiKy', 'hanTraLoi'].filter(function (k) { return kq[k]; }).length;
      kq.nguon = t.nguon || ''; if (/^ocr/.test(t.nguon)) kq.ocr = t.chu.slice(0, 40000);
      log(s, 'cv.ai', (req.chu ? 'phân tích lại chữ OCR đã sửa' : 'đọc bản quét ' + String(f.name || '') + (t.nguon === 'ocr' ? ' (OCR ' + t.chu.length + ' ký tự)' : '')) + ' – bóc được ' + kq.soO + ' trường');
      return kq;
    };

    /* ---- Giám đốc duyệt chia sẻ văn bản cho người liên quan (v3.7) ----
       Trước khi được duyệt chia sẻ, văn bản chỉ văn thư, lãnh đạo và người được giao xử lý xem được. */
    function cvDaChiaSe(o) { return o.csTrangThai === CV_CS_OK; }
    function cvNguoiDoc(o) { return cvDaChiaSe(o) ? parseJSON(o.nguoiDoc, []) : []; }
    function cvDuyetCSDuoc(s) { return s.role === 'giamdoc' || s.role === 'admin'; }
    H['cv.trinhcs'] = function (req, s) {
      var o = cvRow(req.id);
      if (!cvGiaoDuoc(s)) err('Chỉ văn thư, Giám đốc hoặc Quản trị trình chia sẻ văn bản.');
      if (o.trangThai === CV_HET) err('Công văn đã lưu hồ sơ.');
      var ds = (req.nguoiDoc || []).slice(0, 40);
      if (!ds.length) err('Chọn người cần được đọc văn bản này.');
      kiemNguoi(ds, false, 'Người được chia sẻ');
      o.csDeXuat = JSON.stringify(ds); o.csTrangThai = 'Chờ duyệt'; o.capNhat = env.now(); D('CongVan');
      var x = cvGhiXL(o, 'Trình duyệt chia sẻ', s, 'Đề nghị chia sẻ cho: ' + tenNguoi(ds) + (req.note ? ' – ' + String(req.note).trim().slice(0, 1000) : ''), '', null);
      log(s, 'cv.trinhcs', cvTen(o) + ' → ' + tenNguoi(ds));
      bao(nguoiCo(['r:giamdoc'], s.username), { loai: 'duyet', t: 'Công văn chờ duyệt chia sẻ',
        b: cvTen(o) + ': ' + o.trichYeu + ' · đề nghị chia sẻ cho ' + tenNguoi(ds), tag: 'cv_' + o.id, mo: '#v=cv' });
      return { cv: o, xl: x };
    };
    H['cv.duyetcs'] = function (req, s) {
      var o = cvRow(req.id);
      if (!cvDuyetCSDuoc(s)) err('Chỉ Giám đốc hoặc Quản trị duyệt chia sẻ văn bản.');
      if (o.trangThai === CV_HET) err('Công văn đã lưu hồ sơ.');
      var note = String(req.note || '').trim().slice(0, 1000);
      if (!req.duyet) {
        if (!note) err('Nhập lý do không chia sẻ.');
        o.csTrangThai = 'Không chia sẻ'; o.nguoiDoc = '[]';
        o.nguoiDuyetCS = s.hoTen || s.username; o.csLuc = env.now(); o.csGhiChu = note; o.capNhat = env.now(); D('CongVan');
        var x0 = cvGhiXL(o, 'Không chia sẻ', s, note, '', null);
        log(s, 'cv.duyetcs', cvTen(o) + ' → không chia sẻ');
        return { cv: o, xl: x0 };
      }
      var ds = (req.nguoiDoc && req.nguoiDoc.length ? req.nguoiDoc : parseJSON(o.csDeXuat, [])).slice(0, 40);
      if (!ds.length) err('Chọn người được đọc văn bản này.');
      kiemNguoi(ds, false, 'Người được chia sẻ');
      o.nguoiDoc = JSON.stringify(ds); o.csTrangThai = CV_CS_OK;
      o.nguoiDuyetCS = s.hoTen || s.username; o.csLuc = env.now(); o.csGhiChu = note; o.capNhat = env.now(); D('CongVan');
      var x = cvGhiXL(o, 'Duyệt chia sẻ', s, 'Chia sẻ cho: ' + tenNguoi(ds) + (note ? ' – ' + note : ''), '', null);
      log(s, 'cv.duyetcs', cvTen(o) + ' → chia sẻ cho ' + tenNguoi(ds));
      bao(nguoiCo(ds, s.username), { loai: 'duyet', t: 'Văn bản được chia sẻ cho bạn',
        b: cvTen(o) + ': ' + o.trichYeu, tag: 'cv_' + o.id, mo: '#v=cv' });
      return { cv: o, xl: x };
    };

    /* ---- Trao đổi dưới mỗi văn bản (v3.7): ai xem được văn bản thì trao đổi được ---- */
    H['cv.tn'] = function (req, s) {
      var o = cvRow(req.id);
      if (!cvXem(s, o)) err('Bạn không có quyền xem văn bản này.');
      return { ds: T('CongVanTN').filter(function (x) { return x.cvId === o.id; }) };
    };
    H['cv.gui'] = function (req, s) {
      var o = cvRow(req.id);
      if (!cvXem(s, o)) err('Bạn không có quyền trao đổi về văn bản này.');
      if (o.trangThai === CV_HET && s.role !== 'admin') err('Công văn đã lưu hồ sơ, không trao đổi thêm.');
      var nd = String(req.noiDung || '').trim().slice(0, 4000);
      var tep = req.file && req.file.data ? cvTepVao(o, req.file, 'TD') : null;
      if (!nd && !tep) err('Nhập nội dung trao đổi.');
      var m = { id: newId(), cvId: o.id, username: s.username, hoTen: s.hoTen || s.username, noiDung: nd,
        fileId: (tep && tep.fileId) || '', tenTep: (tep && tep.ten) || '', mime: (tep && tep.mime) || '', thoiGian: env.now() };
      T('CongVanTN').push(m); D('CongVanTN');
      o.soTN = T('CongVanTN').filter(function (x) { return x.cvId === o.id; }).length;
      o.capNhat = env.now(); D('CongVan');
      log(s, 'cv.gui', cvTen(o) + ': ' + nd.slice(0, 80));
      var ds = cvNguoi(o).concat(cvNguoiDoc(o)).concat(['r:tchc']);
      bao(nguoiCo(ds, s.username), { loai: 'chat', t: 'Trao đổi về ' + cvTen(o),
        b: (s.hoTen || s.username) + ': ' + (nd || '[tệp] ' + m.tenTep).slice(0, 140), tag: 'cv_' + o.id, mo: '#v=cv' });
      return m;
    };
    H['cv.chot'] = function (req, s) {
      need(s, 'vanthu');
      var o = cvRow(req.id);
      if (o.trangThai === CV_HET) err('Công văn đã lưu hồ sơ.');
      o.trangThai = CV_HET; o.capNhat = env.now(); D('CongVan');
      var x = cvGhiXL(o, 'Lưu hồ sơ', s, String(req.note || '').trim(), '', null);
      log(s, 'cv.chot', cvTen(o));
      return { cv: o, xl: x };
    };
    H['cv.delete'] = function (req, s) {
      need(s, 'vanthu');
      var a = T('CongVan'), o = cvRow(req.id);
      if (['Mới đăng ký', 'Chờ ban hành'].indexOf(o.trangThai) < 0 && s.role !== 'admin') err('Chỉ xóa được văn bản mới đăng ký, chưa giao xử lý.');
      if (o.fileId) { try { env.files.del(o.fileId); } catch (e) { } }
      var xl = T('CongVanXL');
      for (var i = xl.length - 1; i >= 0; i--) if (xl[i].cvId === o.id) { if (xl[i].fileId) { try { env.files.del(xl[i].fileId); } catch (e) { } } xl.splice(i, 1); }
      D('CongVanXL');
      var tn = T('CongVanTN');
      for (var j = tn.length - 1; j >= 0; j--) if (tn[j].cvId === o.id) { if (tn[j].fileId) { try { env.files.del(tn[j].fileId); } catch (e) { } } tn.splice(j, 1); }
      D('CongVanTN');
      a.splice(a.indexOf(o), 1); D('CongVan');
      log(s, 'cv.delete', cvTen(o) + ': ' + o.trichYeu);
      return true;
    };

    /* ====================== ĐỀ XUẤT NỘI BỘ ======================
       Nội dung, ảnh, tệp đính kèm lưu ở "_ĐỀ XUẤT NỘI BỘ / <năm> / <mã đề xuất>" trên Drive Công ty.
       Trình tự duyệt chụp lại tại thời điểm trình, nên sửa cấu hình không làm lệch đề xuất đang chạy. */
    function dxLoaiDs(cfg) { var a = cfg.loaiDeXuat; return Array.isArray(a) && a.length ? a : QT_DX_MD; }
    function dxNhomDs(cfg) { var a = cfg.nhomDeXuat; return Array.isArray(a) && a.length ? a : NHOM_DX_MD; }
    function dxLoai(cfg, id) { return dxLoaiDs(cfg).filter(function (x) { return x.id === id; })[0]; }
    function dxRow(id) { var o = T('DeXuat').filter(function (x) { return x.id === id; })[0]; if (!o) err('Không tìm thấy đề xuất.'); return o; }
    function dxDuong(o) { return [TM_DX, String(o.taoLuc || env.now()).slice(0, 4), String(o.ma || o.id)]; }   // thư mục theo năm lập, mã đề xuất không còn mang năm
    function dxCap(o) { return parseJSON(o.luongDuyet, [])[num(o.capDuyet)] || null; }
    function dxXem(s, o) {
      if (can(s.role, 'view_all')) return true;
      if (o.maNV && o.maNV === s.maNV) return true;
      var ma = 'u:' + s.username, r = 'r:' + s.role;
      var co = function (a) { return (a || []).indexOf(ma) >= 0 || (a || []).indexOf(r) >= 0; };
      if (co(parseJSON(o.theoDoi, [])) || co(parseJSON(o.thucHien, []))) return true;
      return parseJSON(o.luongDuyet, []).some(function (c) { return co(c.nguoi); });
    }
    function dxSuaDuoc(s, o) { return (o.maNV === s.maNV && DX_MO.indexOf(o.trangThai) >= 0 && num(o.capDuyet) === 0) || s.role === 'admin'; }
    /* Mã đề xuất = tiền tố của loại + số thứ tự 4 chữ số, đếm riêng cho từng tiền tố
       (TUL0191, TTTP0268…). Loại không khai tiền tố thì mã chỉ là số (0671). */
    function dxMa(tienTo) {
      tienTo = String(tienTo || '').trim().toUpperCase();
      var n = T('DeXuat').reduce(function (m, x) {
        return String(x.tienTo || '').toUpperCase() === tienTo ? Math.max(m, num(x.so)) : m;
      }, 0) + 1;
      while (T('DeXuat').some(function (x) { return x.ma === tienTo + ('000' + n).slice(-4); })) n++;
      return { tienTo: tienTo, so: n, ma: tienTo + ('000' + n).slice(-4) };
    }
    function dxCapDuyet(o, lo, cfg) {                 // cấp nào tham gia tùy ngưỡng số tiền của cấp đó
      var nop = T('Users').filter(function (u) { return u.maNV === o.maNV; }).map(function (u) { return 'u:' + u.username; });
      var day = dxBuocDay(lo);
      var caps = day.filter(function (c) { return !num(c.tren) || num(o.soTien) > num(c.tren); })
        .map(function (c) { var ng = quyTbp(c.nguoi, o.maNV, cfg); return { id: c.id, ten: c.ten, nguoi: ng, nguoiTen: tenNguoi(ng), chot: !!c.chot }; })
        .filter(function (c) { return c.chot || c.nguoi.some(function (a) { return nop.indexOf(a) < 0; }); });
      caps = caps.filter(function (c, i) { return !i || JSON.stringify(c.nguoi) !== JSON.stringify(caps[i - 1].nguoi); });
      if (lo.cach === 'motnguoi') {                   // gộp mọi người duyệt vào một cấp: ai duyệt cũng xong
        var gop = [];
        caps.forEach(function (c) { c.nguoi.forEach(function (a) { if (gop.indexOf(a) < 0) gop.push(a); }); });
        gop = gop.filter(function (a) { return nop.indexOf(a) < 0; });
        if (!gop.length) gop = quyTbp(['tbp'], o.maNV, cfg);
        caps = [{ id: 'motnguoi', ten: 'Phê duyệt', nguoi: gop, nguoiTen: tenNguoi(gop), motNguoi: true }];
      }
      if (!caps.length) { var ng0 = quyTbp(['tbp'], o.maNV, cfg); caps = [{ id: 'tbp', ten: 'Trưởng bộ phận duyệt', nguoi: ng0, nguoiTen: tenNguoi(ng0) }]; }
      return caps;
    }
    function dxNguoiHop(mac, them, nhan) {
      var out = (mac || []).slice();
      (them || []).slice(0, 20).forEach(function (a) { if (out.indexOf(a) < 0) out.push(a); });
      if (out.length) kiemNguoi(out, false, nhan);
      return out;
    }
    function dxSo(v) { return String(num(v)).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
    function dxLuuDrive(o) {                          // ghi nội dung đề xuất thành tệp văn bản trong thư mục của chính nó
      if (!env.files || !env.files.ghiVB) return;
      try {
        var L = ['ĐỀ XUẤT NỘI BỘ – ' + o.ma, '', 'Nhóm: ' + (o.tenNhom || '—'), 'Loại đề xuất: ' + o.tenLoai,
          'Bộ phận đề xuất: ' + (o.boPhan || '—'), 'Tiêu đề: ' + o.tieuDe,
          'Người đề nghị: ' + o.hoTen + (o.boPhan ? ' – ' + o.boPhan : ''), 'Lập lúc: ' + o.taoLuc, 'Độ khẩn: ' + o.doKhan];
        if (num(o.soTien)) L.push('Số tiền dự kiến: ' + dxSo(o.soTien) + ' đồng' + (o.donVi ? ' (' + o.donVi + ')' : ''));
        if (o.hanCan) L.push('Cần có trước ngày: ' + o.hanCan);
        L.push('Trạng thái: ' + o.trangThai, '', 'NỘI DUNG ĐỀ XUẤT', o.noiDung);
        var caps = parseJSON(o.luongDuyet, []);
        if (caps.length) L.push('', 'TRÌNH TỰ DUYỆT', caps.map(function (c, i) { return (i + 1) + '. ' + c.ten + ' – ' + c.nguoiTen; }).join('\n'));
        var td = parseJSON(o.theoDoi, []), th = parseJSON(o.thucHien, []);
        if (td.length) L.push('', 'NGƯỜI THEO DÕI: ' + tenNguoi(td));
        if (th.length) L.push('NGƯỜI THỰC HIỆN: ' + tenNguoi(th));
        var ls = parseJSON(o.lichSu, []);
        if (ls.length) L.push('', 'DIỄN BIẾN', ls.map(function (h) { return h.luc + ' · ' + h.cap + ' · ' + h.nguoi + ': ' + h.ketQua + (h.ghiChu ? ' – ' + h.ghiChu : ''); }).join('\n'));
        if (o.ketQua) L.push('', 'KẾT QUẢ THỰC HIỆN', o.ketQua);
        o.noiDungId = env.files.ghiVB(dxDuong(o), 'NoiDung_' + o.ma + '.txt', L.join('\n'));
        D('DeXuat');
      } catch (e) { }
    }
    function dxBao(o, s) {                            // tới lượt cấp nào thì báo người của cấp đó
      var c = dxCap(o); if (!c) return;
      bao(nguoiCo(c.nguoi, s.username).filter(function (u) { return u !== userCuaNV(o.maNV); }), { loai: 'duyet',
        t: 'Đề xuất chờ bạn duyệt', b: o.ma + ' · ' + o.hoTen + ': ' + o.tieuDe + (num(o.soTien) ? ' · ' + dxSo(o.soTien) + ' đ' : ''),
        tag: 'dx_' + o.id, mo: '#v=dx' });
    }
    function dxThongTin(s) {
      var cfg = trangTriQT(getCfg());
      return { loai: dxLoaiDs(cfg).filter(function (l) { return l.hoatDong !== false; }),
        nhom: cfg.nhomDeXuat || NHOM_DX_MD,
        boPhan: boPhanDs() };
    }
    function boPhanDs() {                             // bộ phận đề xuất: phòng ban, công trình, đơn vị đang hoạt động
      var cfg = getCfg(), out = [];
      (cfg.boPhan || []).forEach(function (b) { if (b.trangThai !== 'Ngừng' && b.ten && out.indexOf(b.ten) < 0) out.push(b.ten); });
      (cfg.donVi || []).forEach(function (t) { if (t && out.indexOf(t) < 0) out.push(t); });
      T('NhanVien').forEach(function (n) {
        [n.phongBan, n.donVi].forEach(function (t) { if (t && out.indexOf(t) < 0) out.push(t); });
      });
      return out;
    }
    H['dx.ds'] = function (req, s) {
      var ds = T('DeXuat').filter(function (o) { return dxXem(s, o); }), co = {};
      ds.forEach(function (o) { co[o.id] = 1; });
      var tt = dxThongTin(s);
      return { ds: ds, tep: T('DeXuatTep').filter(function (t) { return co[t.dxId]; }), loai: tt.loai, nhom: tt.nhom, boPhan: tt.boPhan };
    };
    H['dx.create'] = function (req, s) {
      if (!s.maNV) err('Tài khoản chưa gắn với hồ sơ CBCNV. Đề nghị TCHC gắn hồ sơ trước khi lập đề xuất.');
      var r = req.row || {}, cfg = getCfg(), lo = dxLoai(cfg, r.loai);
      if (!lo) err('Chọn loại đề xuất.');
      if (lo.hoatDong === false) err('Loại đề xuất "' + lo.ten + '" đã ngừng sử dụng.');
      var nv = findNV(s.maNV) || {}, m = dxMa(lo.tienTo);
      var nhTen = (dxNhomDs(cfg).filter(function (g) { return g.id === lo.nhom; })[0] || {}).ten || '';
      var bp = String(r.boPhan || '').trim() || nv.donVi || nv.phongBan || s.boPhan || '';
      if (!bp) err('Chọn bộ phận đề xuất.');
      if (boPhanDs().indexOf(bp) < 0) err('Bộ phận "' + bp + '" không có trong danh mục phòng ban, công trình của Công ty.');
      var o = { id: newId(), ma: m.ma, tienTo: m.tienTo, so: m.so, loai: lo.id, tenLoai: lo.ten, nhom: lo.nhom || '', tenNhom: nhTen,
        tieuDe: String(r.tieuDe || '').trim().slice(0, 300), noiDung: String(r.noiDung || '').trim().slice(0, 20000),
        soTien: lo.tien ? Math.max(0, num(r.soTien)) : 0, donVi: String(r.donVi || '').trim().slice(0, 120),
        maNV: s.maNV, hoTen: nv.hoTen || s.hoTen || s.username, boPhan: bp,
        doKhan: CV_KHAN.indexOf(r.doKhan) >= 0 ? r.doKhan : 'Thường', hanCan: String(r.hanCan || '').trim(),
        nguoiTao: s.hoTen || s.username, taoLuc: env.now(), capNhat: env.now(), lichSu: '[]', capDuyet: 0 };
      if (!o.tieuDe) err('Nhập tiêu đề đề xuất.');
      if (o.noiDung.length < 10) err('Mô tả rõ nội dung đề xuất (tối thiểu 10 ký tự).');
      if (o.hanCan && !parseD(o.hanCan)) err('Thời hạn cần có không hợp lệ.');
      if (lo.tien && !o.soTien) err('Nhập số tiền dự kiến.');
      o.luongDuyet = JSON.stringify(dxCapDuyet(o, lo, cfg));
      o.theoDoi = JSON.stringify(dxNguoiHop(lo.theoDoi, req.theoDoi, 'Người theo dõi'));
      o.thucHien = JSON.stringify(dxNguoiHop(lo.thucHien, req.thucHien, 'Người thực hiện'));
      o.trangThai = req.nhap ? 'Nháp' : 'Chờ duyệt';
      T('DeXuat').push(o); D('DeXuat');
      dxLuuDrive(o);
      log(s, 'dx.create', o.ma + ' ' + o.tenLoai + ': ' + o.tieuDe);
      if (o.trangThai === 'Chờ duyệt') dxBao(o, s);
      return o;
    };
    H['dx.update'] = function (req, s) {
      var o = dxRow(req.id), cfg = getCfg();
      if (!dxSuaDuoc(s, o)) err('Chỉ người lập sửa được, và chỉ khi chưa có cấp nào duyệt.');
      var r = req.row || {}, lo = dxLoai(cfg, r.loai || o.loai); if (!lo) err('Chọn loại đề xuất.');
      o.loai = lo.id; o.tenLoai = lo.ten; o.nhom = lo.nhom || '';
      o.tenNhom = (dxNhomDs(cfg).filter(function (g) { return g.id === lo.nhom; })[0] || {}).ten || '';
      if (r.boPhan !== undefined) {
        var bp2 = String(r.boPhan || '').trim();
        if (!bp2) err('Chọn bộ phận đề xuất.');
        if (boPhanDs().indexOf(bp2) < 0) err('Bộ phận "' + bp2 + '" không có trong danh mục phòng ban, công trình của Công ty.');
        o.boPhan = bp2;
      }
      if (r.tieuDe !== undefined) o.tieuDe = String(r.tieuDe || '').trim().slice(0, 300);
      if (r.noiDung !== undefined) o.noiDung = String(r.noiDung || '').trim().slice(0, 20000);
      if (r.donVi !== undefined) o.donVi = String(r.donVi || '').trim().slice(0, 120);
      if (r.doKhan !== undefined && CV_KHAN.indexOf(r.doKhan) >= 0) o.doKhan = r.doKhan;
      if (r.hanCan !== undefined) o.hanCan = String(r.hanCan || '').trim();
      o.soTien = lo.tien ? Math.max(0, num(r.soTien !== undefined ? r.soTien : o.soTien)) : 0;
      if (!o.tieuDe) err('Nhập tiêu đề đề xuất.');
      if (o.noiDung.length < 10) err('Mô tả rõ nội dung đề xuất (tối thiểu 10 ký tự).');
      if (o.hanCan && !parseD(o.hanCan)) err('Thời hạn cần có không hợp lệ.');
      if (lo.tien && !o.soTien) err('Nhập số tiền dự kiến.');
      o.luongDuyet = JSON.stringify(dxCapDuyet(o, lo, cfg));
      if (req.trinh && o.trangThai === 'Nháp') o.trangThai = 'Chờ duyệt';
      o.capNhat = env.now(); D('DeXuat');
      dxLuuDrive(o);
      log(s, 'dx.update', o.ma + ': ' + o.tieuDe + (req.trinh ? ' (trình duyệt)' : ''));
      if (req.trinh && o.trangThai === 'Chờ duyệt') dxBao(o, s);
      return o;
    };
    H['dx.nguoi'] = function (req, s) {
      var o = dxRow(req.id), cfg = getCfg();
      var laChu = o.maNV === s.maNV, duyet = laNguoiXuLy(s, (dxCap(o) || {}).nguoi || []);
      if (!laChu && !duyet && !can(s.role, 'view_all')) err('Không có quyền đổi người theo dõi, thực hiện của đề xuất này.');
      if (DX_XONG.indexOf(o.trangThai) >= 0) err('Đề xuất đã kết thúc.');
      if (req.theoDoi) o.theoDoi = JSON.stringify(dxNguoiHop([], req.theoDoi, 'Người theo dõi'));
      if (req.thucHien) o.thucHien = JSON.stringify(dxNguoiHop([], req.thucHien, 'Người thực hiện'));
      o.capNhat = env.now(); D('DeXuat'); dxLuuDrive(o);
      log(s, 'dx.nguoi', o.ma + ': theo dõi ' + tenNguoi(parseJSON(o.theoDoi, [])) + ' · thực hiện ' + tenNguoi(parseJSON(o.thucHien, [])));
      return o;
    };
    H['dx.upload'] = function (req, s) {
      var o = dxRow(req.id), cfg = getCfg();
      var duoc = dxSuaDuoc(s, o) || (o.maNV === s.maNV && DX_XONG.indexOf(o.trangThai) < 0)
        || laNguoiXuLy(s, parseJSON(o.thucHien, [])) || laNguoiXuLy(s, (dxCap(o) || {}).nguoi || []);
      if (!duoc) err('Không có quyền đính kèm vào đề xuất này.');
      var f = req.file || {}, ten = String(f.name || '').trim().replace(/[\\/:*?"<>|]/g, '_').slice(0, 150);
      if (!ten || !f.data) err('Chưa chọn tệp.');
      if (!MIME_OK.test(f.mime || '')) err('Chỉ nhận PDF, ảnh (JPG, PNG, WEBP, HEIC), Word hoặc Excel.');
      var bytes = Math.floor(String(f.data).length * 3 / 4);
      if (bytes > (num(cfg.dungLuongToiDaMB) || 10) * 1048576) err('Tệp vượt ' + (num(cfg.dungLuongToiDaMB) || 10) + ' MB.');
      if (T('DeXuatTep').filter(function (t) { return t.dxId === o.id; }).length >= 20) err('Mỗi đề xuất đính kèm tối đa 20 tệp.');
      var t = { id: newId(), dxId: o.id, loai: /^image\//.test(f.mime || '') ? 'anh' : 'tep', ten: ten, mime: f.mime,
        kichThuoc: bytes, fileId: env.files.put({ name: ten, mime: f.mime, data: f.data, duong: dxDuong(o) }),
        nguoiTai: s.hoTen || s.username, taiLuc: env.now() };
      T('DeXuatTep').push(t); D('DeXuatTep');
      log(s, 'dx.upload', o.ma + ': ' + ten);
      return t;
    };
    H['dx.tepget'] = function (req, s) {
      var t = T('DeXuatTep').filter(function (x) { return x.id === req.id; })[0]; if (!t) err('Không tìm thấy tệp.');
      var o = dxRow(t.dxId);
      if (!dxXem(s, o)) err('Bạn không có quyền xem tệp này.');
      var f = env.files.get(t.fileId); if (!f) err('Không đọc được tệp (có thể đã bị xóa khỏi Drive).');
      log(s, 'dx.tepxem', o.ma + ': ' + t.ten);
      return { name: t.ten, mime: t.mime, data: f.data };
    };
    H['dx.tepdel'] = function (req, s) {
      var a = T('DeXuatTep'), t = a.filter(function (x) { return x.id === req.id; })[0]; if (!t) err('Không tìm thấy tệp.');
      var o = dxRow(t.dxId);
      if (!dxSuaDuoc(s, o) && t.nguoiTai !== (s.hoTen || s.username)) err('Chỉ người tải lên hoặc người lập (khi chưa ai duyệt) xóa được tệp.');
      try { env.files.del(t.fileId); } catch (e) { }
      a.splice(a.indexOf(t), 1); D('DeXuatTep');
      log(s, 'dx.tepdel', o.ma + ': ' + t.ten);
      return true;
    };
    H['dx.decide'] = function (req, s) {
      var o = dxRow(req.id);
      if (o.trangThai !== 'Chờ duyệt') err('Đề xuất không ở trạng thái chờ duyệt (' + o.trangThai + ').');
      var caps = parseJSON(o.luongDuyet, []), ci = num(o.capDuyet), cap = caps[ci];
      if (!cap) err('Đề xuất không còn cấp duyệt.');
      var note = String(req.note || '').trim().slice(0, 2000), duocGiao = laNguoiXuLy(s, cap.nguoi);
      var vuotCap = !duocGiao && (s.role === 'giamdoc' || s.role === 'admin');
      if (o.maNV === s.maNV && s.role !== 'admin') err('Không tự duyệt đề xuất của chính mình.');
      if (!duocGiao && !vuotCap) err('Đề xuất đang chờ "' + cap.ten + '" (' + cap.nguoiTen + ').');
      var ls = parseJSON(o.lichSu, []);
      ls.push({ cap: cap.ten, nguoi: s.hoTen || s.username, luc: env.now(),
        ketQua: req.approve ? (vuotCap ? 'Duyệt (vượt cấp)' : 'Duyệt') : 'Từ chối', ghiChu: note });
      if (!req.approve) { if (!note) err('Nhập lý do từ chối.'); o.trangThai = 'Từ chối'; }
      else if (vuotCap || ci + 1 >= caps.length) o.trangThai = 'Đã duyệt';
      else o.capDuyet = ci + 1;
      if (o.trangThai !== 'Chờ duyệt') { o.nguoiDuyet = s.hoTen || s.username; o.ngayDuyet = env.now(); }
      if (note) o.ghiChuDuyet = (o.ghiChuDuyet ? o.ghiChuDuyet + ' | ' : '') + (s.hoTen || s.username) + ': ' + note;
      o.lichSu = JSON.stringify(ls); o.capNhat = env.now(); D('DeXuat');
      dxLuuDrive(o);
      log(s, 'dx.decide', o.ma + ' → ' + (o.trangThai === 'Chờ duyệt' ? 'chuyển ' + caps[o.capDuyet].ten : o.trangThai));
      if (o.trangThai === 'Chờ duyệt') dxBao(o, s);
      else {
        bao([userCuaNV(o.maNV)].filter(function (u) { return u && u !== s.username; }), { loai: 'duyet',
          t: 'Đề xuất ' + (o.trangThai === 'Đã duyệt' ? 'đã được duyệt' : 'bị từ chối'),
          b: o.ma + ': ' + o.tieuDe + (note ? ' · ' + note : ''), tag: 'dx_' + o.id, mo: '#v=mydx' });
        var th = parseJSON(o.thucHien, []);
        if (o.trangThai === 'Đã duyệt' && th.length) bao(nguoiCo(th, s.username), { loai: 'duyet',
          t: 'Đề xuất đã duyệt – chờ bạn thực hiện', b: o.ma + ': ' + o.tieuDe + (o.hanCan ? ' · cần có trước ' + ngayVN(o.hanCan) : ''),
          tag: 'dx_' + o.id, mo: '#v=dx' });
      }
      return o;
    };
    H['dx.ketqua'] = function (req, s) {
      var o = dxRow(req.id);
      if (['Đã duyệt', 'Đang thực hiện'].indexOf(o.trangThai) < 0) err('Chỉ báo cáo thực hiện khi đề xuất đã được duyệt.');
      var th = parseJSON(o.thucHien, []);
      if (!laNguoiXuLy(s, th) && !can(s.role, 'view_all')) err('Đề xuất này không giao cho bạn thực hiện.');
      var nd = String(req.noiDung || '').trim().slice(0, 4000);
      if (!nd) err('Nhập nội dung đã thực hiện.');
      var ls = parseJSON(o.lichSu, []);
      ls.push({ cap: 'Thực hiện', nguoi: s.hoTen || s.username, luc: env.now(), ketQua: req.xong ? 'Hoàn thành' : 'Đang thực hiện', ghiChu: nd });
      o.lichSu = JSON.stringify(ls);
      o.trangThai = req.xong ? 'Hoàn thành' : 'Đang thực hiện';
      o.ketQua = (o.ketQua ? o.ketQua + '\n' : '') + (s.hoTen || s.username) + ' (' + env.now() + '): ' + nd;
      if (req.xong) o.hoanThanhLuc = env.now();
      o.capNhat = env.now(); D('DeXuat');
      dxLuuDrive(o);
      log(s, 'dx.ketqua', o.ma + ' → ' + o.trangThai);
      bao([userCuaNV(o.maNV)].concat(nguoiCo(parseJSON(o.theoDoi, []), s.username)).filter(function (u) { return u && u !== s.username; }),
        { loai: 'duyet', t: req.xong ? 'Đề xuất đã hoàn thành' : 'Đề xuất đang được thực hiện', b: o.ma + ': ' + nd.slice(0, 120), tag: 'dx_' + o.id, mo: '#v=mydx' });
      return o;
    };
    H['dx.huy'] = function (req, s) {
      var o = dxRow(req.id);
      var laChu = o.maNV === s.maNV, ql = s.role === 'admin' || s.role === 'giamdoc';
      if (DX_XONG.indexOf(o.trangThai) >= 0) err('Đề xuất đã kết thúc.');
      if (!laChu && !ql) err('Chỉ người lập, Giám đốc hoặc Quản trị hủy được đề xuất.');
      if (laChu && !ql && DX_MO.indexOf(o.trangThai) < 0) err('Đề xuất đã được duyệt, đề nghị liên hệ Giám đốc nếu cần hủy.');
      var ls = parseJSON(o.lichSu, []);
      ls.push({ cap: 'Hủy', nguoi: s.hoTen || s.username, luc: env.now(), ketQua: 'Đã hủy', ghiChu: String(req.note || '').trim().slice(0, 1000) });
      o.lichSu = JSON.stringify(ls); o.trangThai = 'Đã hủy'; o.capNhat = env.now(); D('DeXuat');
      dxLuuDrive(o);
      log(s, 'dx.huy', o.ma + ': ' + o.tieuDe);
      return o;
    };

    /* ---- người xử lý: diễn giải & quy về người cụ thể ---- */
    function tenNguoi(nguoi) {
      var us = T('Users');
      return (nguoi || []).map(function (a) {
        if (a === 'tbp') return 'Trưởng phòng / Chỉ huy trưởng của người đề nghị';
        if (a.indexOf('r:') === 0) return ROLES[a.slice(2)] || a.slice(2);
        var u = us.filter(function (x) { return 'u:' + x.username === a; })[0];
        return u ? (u.hoTen || u.username) : a.slice(2);
      }).join(', ');
    }
    function quyTbp(nguoi, maNV, cfg) {      // thay 'tbp' bằng Trưởng phòng / Chỉ huy trưởng của bộ phận người đề nghị
      var nv = findNV(maNV) || {}, map = cfg.truongBoPhan || {};
      var ql = quanLyCua(nv.phongBan, maNV).map(function (u) { return 'u:' + u.username; });
      if (!ql.length) { var m = (nv.phongBan && map[nv.phongBan]) || (nv.donVi && map[nv.donVi]) || ''; if (m) ql = ['u:' + m]; }
      if (!ql.length) ql = ['r:tchc'];                    // bộ phận chưa có người quản lý: chuyển TCHC
      var out = [];
      (nguoi || []).forEach(function (a) {
        (a === 'tbp' ? ql : [a]).forEach(function (v) { if (out.indexOf(v) < 0) out.push(v); });
      });
      return out;
    }
    function trangTriQT(cfg) {               // gắn tên người xử lý để hiển thị
      ['quyTrinhLuong', 'quyTrinhNghi', 'quyTrinhYKien'].forEach(function (k) {
        cfg[k] = (cfg[k] || []).map(function (b) { var c = clone(b); c.nguoiTen = tenNguoi(b.nguoi); return c; });
      });
      var nhDX = {};
      (cfg.nhomDeXuat || NHOM_DX_MD).forEach(function (g) { nhDX[g.id] = g.ten; });
      cfg.loaiDeXuat = (cfg.loaiDeXuat || QT_DX_MD).map(function (l) {
        var c = clone(l);
        c.buoc = dxBuocDay(l).map(function (b) { var d = clone(b); d.nguoiTen = tenNguoi(b.nguoi); return d; });
        c.theoDoiTen = tenNguoi(l.theoDoi); c.thucHienTen = tenNguoi(l.thucHien);
        c.tenNhom = nhDX[l.nhom] || ''; c.cachTen = DX_CACH[l.cach || 'trinhtu'];
        return c;
      });
      return cfg;
    }

    /* ---- chấm công & bảng lương: quy trình theo cấu hình ---- */
    function kyRow(ky, create) {
      var a = T('KyLuong'), k = a.filter(function (x) { return x.ky === ky; })[0];
      if (!k && create) {
        var steps = trangTriQT(getCfg()).quyTrinhLuong;
        k = { ky: ky, trangThai: steps[0].id, quyTrinh: JSON.stringify(steps), capNhat: env.now() }; a.push(k); D('KyLuong');
      }
      if (k) k.trangThai = normBuoc(k.trangThai, buocLuong(getCfg(), k));
      return k;
    }
    function buocKy(k) { return buocLuong(getCfg(), k); }
    function buocHienTai(k) { var st = buocKy(k); return st[viTriBuoc(st, k.trangThai)] || null; }
    function ghiPD(k, tu, den, hd, s, note) {
      var st = buocKy(k), ten = function (id) { if (id === BUOC_XONG) return BUOC_XONG; var b = st[viTriBuoc(st, id)]; return b ? b.ten : id; };
      T('PheDuyet').push({ id: newId(), ky: k.ky, tu: tu, den: den, tuTen: ten(tu), denTen: ten(den), hanhDong: hd, nguoi: s.hoTen || s.username, vaiTro: ROLES[s.role] || s.role, thoiGian: env.now(), ghiChu: String(note || '').trim() });
      D('PheDuyet');
    }
    function cleanTamUng(list) {
      list = parseJSON(list, []);
      if (!Array.isArray(list)) err('Danh sách tạm ứng không hợp lệ.');
      var loai = getCfg().loaiTamUng || LOAI_TU_MD;
      return list.map(function (x) {
        var t = R(num(x.soTien)); if (t <= 0) err('Khoản tạm ứng thiếu số tiền.');
        if (x.ngay && !parseD(x.ngay)) err('Ngày tạm ứng không hợp lệ.');
        return { loai: loai.indexOf(x.loai) >= 0 ? x.loai : loai[loai.length - 1], soTien: t, ngay: String(x.ngay || ''), ghiChu: String(x.ghiChu || '').trim() };
      });
    }
    function cleanPhat(list) {
      list = parseJSON(list, []);
      if (!Array.isArray(list)) err('Danh sách vi phạm không hợp lệ.');
      return list.map(function (x) {
        var ten = String(x.ten || '').trim(); if (!ten) err('Vi phạm thiếu nội dung.');
        var loai = LOAI_VP.indexOf(x.loai) >= 0 ? x.loai : LOAI_VP[0];
        var soLan = Math.max(1, num(x.soLan) || 1), muc = num(x.mucTien);
        var tt = num(x.thanhTien) || muc * soLan;
        if (tt <= 0) err('Vi phạm "' + ten + '" chưa có số tiền.');
        var nhom = NHOM_VP.indexOf(x.nhom) >= 0 ? x.nhom : (loai === 'Bồi thường thiệt hại' ? 'Mất mát, hư hỏng tài sản' : 'Vi phạm khác');
        return { ma: String(x.ma || ''), ten: ten, nhom: nhom, loai: loai, soLan: soLan, mucTien: muc, thanhTien: tt, ghiChu: String(x.ghiChu || '') };
      });
    }
    function recalc(ky, s) {
      var k = kyRow(ky, true), cfg = getCfg();
      var ncc = num(k.ngayCongChuan) || (cfg.cachTinhCongChuan === 'truCN' ? congChuanTruCN(ky) : num(cfg.ngayCongChuan));
      k.ngayCongChuan = ncc;
      var cc = T('ChamCong').filter(function (x) { return x.ky === ky; });
      var ll = T('LenLuong'), kt = T('KhenThuong'), mh = T('MucHuong');
      var out = [], thieu = [];
      T('NhanVien').forEach(function (nv) {
        if (!dangLamTrongKy(nv, ky)) return;
        var c = cc.filter(function (x) { return x.maNV === nv.maNV; })[0];
        if (!c) { thieu.push(nv.maNV + ' ' + nv.hoTen); return; }
        var dv = dauVaoTinhLuong(nv, ky, ncc, ll, mh, kt, T('ChamCong'));
        if (cfg.ckBatBuocDuyet !== false) dv.o.ckHopLe = validateTaxCommitmentDocuments(dv.nv, ky.slice(0, 4), T('HoSoThue'), cfg).ok;   // v3.28: cam kết chỉ áp dụng khi đã ký, duyệt, có MST
        var p = dv.segs.length > 1 ? tinhNhieuDoan(nv, c, cfg, dv, dv.segs, ll, mh) : tinhPhieuLuong(dv.nv, c, cfg, dv.o);      // v3.33: đổi HĐ giữa kỳ → chia giai đoạn
        if (dv.segs.length === 1) { var sg = dv.segs[0]; p.ctHD = JSON.stringify([{ soHD: sg.hd.soHD || '', loaiHD: sg.hd.loaiHD, tu: sg.tu, den: sg.den, cong: num(c.ngayCong), luong: num(dv.o.luong), tong: num(p.tongThuNhap), tnct: num(p.tnChiuThue), cach: ttThue(dv.nv, ky).cach }]); }
        else if (!dv.segs.length) p.ctHD = '';
        p.id = ky + '|' + nv.maNV; p.ky = ky; p.maNV = nv.maNV; p.hoTen = nv.hoTen; p.phongBan = nv.phongBan; p.chucVu = nv.chucVu; p.congTy = congTyNV(nv, cfg);   // v3.20: chốt công ty hạch toán
        p.ctTamUng = c.chiTietTamUng || ''; p.ctViPham = c.chiTietPhat || '';                 // để CBCNV xem trên phiếu lương
        p.ctCong = JSON.stringify({ c: num(c.ngayCong), hl: num(c.nghiHuongLuong), kl: num(c.nghiKhongLuong), om: num(c.nghiOm), le: num(c.ngayLe), ot150: num(c.ot150), ot200: num(c.ot200), ot300: num(c.ot300), ngay: c.chiTietNgay ? 1 : 0, npt: num(dv.nv.soNPT), gtbt: num(p.giamTru) ? num(cfgTheoKy(cfg, ky).giamTruBanThan) : 0, gtnpt: num(p.giamTru) ? num(cfgTheoKy(cfg, ky).giamTruNPT) : 0, lbhHS: num(dv.o.luongBH) });
        out.push(p);
      });
      // v3.32: cập nhật TẠI CHỖ các dòng bảng lương của kỳ (trước đây xóa hết rồi thêm lại → ghi lại cả sheet BangLuong mỗi lần tính, rất chậm)
      var a = T('BangLuong'), cu = {}, con = {};
      a.forEach(function (x) { if (x.ky === ky) cu[x.id] = x; });
      out.forEach(function (p) { con[p.id] = 1; });
      for (var i = a.length - 1; i >= 0; i--) if (a[i].ky === ky && !con[a[i].id]) a.splice(i, 1);
      out = out.map(function (p) {
        var x = cu[p.id]; if (!x) { a.push(p); return p; }
        Object.keys(x).forEach(function (f) { if (!(f in p)) x[f] = ''; });
        Object.assign(x, p); return x;
      });
      k.nguoiTinh = s.username; k.ngayTinh = env.now(); k.capNhat = env.now();
      k.congThuc = JSON.stringify({ congThucLuong: cfg.congThucLuong.filter(function (c) { return !c.an; }), congThucDongBH: cfg.congThucDongBH });
      D('BangLuong'); D('KyLuong');
      return { bl: out, thieu: thieu };
    }
    function kyData(ky) {
      return {
        ky: kyRow(ky, false), cc: T('ChamCong').filter(function (x) { return x.ky === ky; }),
        bl: T('BangLuong').filter(function (x) { return x.ky === ky; }), pd: T('PheDuyet').filter(function (x) { return x.ky === ky; })
      };
    }
    function canTinhLai(k) { var st = buocKy(k), i = viTriBuoc(st, k.trangThai); return i >= 1 && i < st.length && st[i].loai !== 'pheduyet'; }
    function daTrinhDuyet(k) { var st = buocKy(k), i = viTriBuoc(st, k.trangThai); return i >= st.length - 1; }

    H['cc.save'] = function (req, s) {
      var ky = req.ky; if (!validKy(ky)) err('Kỳ lương không hợp lệ.');
      var k = kyRow(ky, true), b = buocHienTai(k), fields = b && FIELDS_LOAI[b.loai];
      if (!fields) err('Kỳ ' + ky + ' đang ở bước "' + (b ? b.ten : k.trangThai) + '", không nhập dữ liệu ở bước này.');
      if (!laNguoiXuLy(s, b.nguoi)) err('Bước "' + b.ten + '" do ' + tenNguoi(b.nguoi) + ' thực hiện.');
      if (num(req.ngayCongChuan) && (b.loai === 'chamcong' || b.loai === 'tinhluong')) k.ngayCongChuan = num(req.ngayCongChuan);
      var a = T('ChamCong'), n = ghiCC(ky, b, req.rows || []);
      D('ChamCong'); D('KyLuong'); k.capNhat = env.now();
      var thieu = [], tinh = canTinhLai(k);
      if (tinh) thieu = recalc(ky, s).thieu;
      log(s, 'cc.save', ky + ' [' + b.ten + ']: ' + n + ' dòng');
      if (!tinh) return { ky: kyRow(ky, false), cc: a.filter(function (x) { return x.ky === ky; }), pd: T('PheDuyet').filter(function (x) { return x.ky === ky; }), thieu: [] };   // v3.22: chưa tính lương → không đọc sheet BangLuong
      var out = kyData(ky); out.thieu = thieu; return out;
    };
    /* v3.32: ghi dữ liệu chấm công / công nợ của bước đang xử lý (dùng chung cho cc.save và ky.move – lưu & chuyển bước trong 1 lần gọi) */
    function ghiCC(ky, b, rows) {
      var fields = FIELDS_LOAI[b.loai], a = T('ChamCong'), n = 0, theoId = {};
      a.forEach(function (x) { if (x.ky === ky) theoId[x.id] = x; });            // v3.22: tra theo mã, không quét cả bảng cho từng người
      (rows || []).forEach(function (r) {
        if (!findNV(r.maNV)) return;
        var id = ky + '|' + r.maNV, o = theoId[id];
        if (!o) {
          if (b.loai !== 'chamcong') return;
          o = { id: id, ky: ky, maNV: r.maNV }; a.push(o); theoId[id] = o;
        }
        var coTU = parseJSON(o.chiTietTamUng, []).length > 0;
        fields.forEach(function (c) {
          if (r[c] === undefined) return;
          if (c === 'tamUng' && coTU && r.chiTietTamUng === undefined) return;        // tổng tạm ứng lấy theo bảng kê
          if (c === 'tienAnThu') {                                                    // v3.34: ăn ca – mặc định theo bảng chấm ăn ca, Công nợ sửa được; để trống = quay về số theo bảng
            var tinhA = o.chiTietAn ? R(num(o.soBua) * num(nlvTrongKy(findNV(r.maNV) || {}, o, getCfg()).muc.giaBua)) : 0;
            if (r[c] === '' || r[c] === null) { o.anThuTay = 0; o.tienAnThu = tinhA; return; }
            var vA = Math.max(0, R(num(r[c])));
            if (o.chiTietAn && vA === tinhA) { o.anThuTay = 0; o.tienAnThu = tinhA; return; }
            o.tienAnThu = vA; o.anThuTay = o.chiTietAn ? 1 : 0; return;
          }
          if (c === 'chiTietAn') {                                                    // v3.32: bảng chấm ăn ca – số bữa/ngày 0–3
            var nvA = findNV(r.maNV), an = cleanAn(r[c], ky), coTruoc = !!o.chiTietAn;
            Object.keys(an).forEach(function (d) { if (nvA && ngoaiHD(nvA, ky + '-' + pad(+d))) delete an[d]; });
            o.chiTietAn = Object.keys(an).length ? JSON.stringify(an) : '';
            o.soBua = Object.keys(an).reduce(function (t, d) { return t + an[d]; }, 0);
            if (num(o.anThuTay)) { if (!o.chiTietAn) o.anThuTay = 0; }          // đã sửa tay ở Công nợ: giữ số đã sửa
            else if (o.chiTietAn) o.tienAnThu = R(o.soBua * num(nlvTrongKy(nvA || {}, o, getCfg()).muc.giaBua)); else if (coTruoc) o.tienAnThu = 0;
            return;
          }
          if (c === 'chiTietNgay') {
            var cfg0 = getCfg(), ct = cleanNgay(r[c], ky, ngayTheoDon(T('NghiPhep'), r.maNV, ky, cfg0)), nvC = findNV(r.maNV);
            Object.keys(ct).forEach(function (d) { if (nvC && ngoaiHD(nvC, ky + '-' + pad(+d))) delete ct[d]; });   // v3.18: không chấm ngày trước khi vào làm / từ ngày chấm dứt HĐ
            var th = tongHopNgay(ct, ky, cfg0);
            o.chiTietNgay = Object.keys(ct).length ? JSON.stringify(ct) : '';
            if (o.chiTietNgay) {
              ['ngayCong', 'nghiHuongLuong', 'nghiKhongLuong', 'nghiOm', 'ngayLe', 'ot300'].forEach(function (f) { o[f] = th[f]; });
              if (cfg0.chuNhatTinh === 'ot') o.ot200 = th.ot200;
            }
            return;
          }
          if (c === 'chiTietOT') {                                                    // v3.25: bảng chấm công thêm giờ
            var nvO = findNV(r.maNV), ot = cleanOT(r[c], ky);
            Object.keys(ot).forEach(function (d) { if (nvO && ngoaiHD(nvO, ky + '-' + pad(+d))) delete ot[d]; });
            o.chiTietOT = Object.keys(ot).length ? JSON.stringify(ot) : '';
            return;
          }
          if (c === 'chiTietTamUng') {
            var tu = cleanTamUng(r[c]); o.chiTietTamUng = tu.length ? JSON.stringify(tu) : '';
            o.tamUng = tu.reduce(function (t, x) { return t + x.soTien; }, 0); return;
          }
          if (c === 'bienKy') {
            var cur = parseJSON(o.bienKy, {}), nhan = parseJSON(r.bienKy, {});
            (getCfg().bienKy || []).forEach(function (v) { if (v.buoc === b.loai && nhan[v.ma] !== undefined) cur[v.ma] = nhan[v.ma] === '' ? '' : num(nhan[v.ma]); });
            o.bienKy = JSON.stringify(cur); return;
          }
          if (c === 'noiLamViec') {
            var mnl = String(r[c] || '').trim().toUpperCase();
            if (mnl && !(getCfg().noiLamViec || []).some(function (x) { return String(x.ma).toUpperCase() === mnl; })) err('Nơi làm việc "' + mnl + '" (' + r.maNV + ') chưa có trong Cài đặt.');
            o[c] = mnl; return;
          }
          o[c] = c === 'chiTietPhat' ? JSON.stringify(cleanPhat(r[c])) : (/^ghiChu/.test(c) ? String(r[c] || '').trim() : num(r[c]));
        });
        if (r.chiTietOT !== undefined || r.chiTietNgay !== undefined || o.chiTietOT) apGioOT(o, ky, getCfg());   // v3.25: cộng giờ làm thêm từ 2 bảng
        n++;
      });
      return n;
    }

    H['bl.calc'] = function (req, s) {
      var ky = req.ky; if (!validKy(ky)) err('Kỳ lương không hợp lệ.');
      var k = kyRow(ky, true), b = buocHienTai(k);
      if (!b || b.loai !== 'tinhluong') err('Chỉ tính lương ở bước "Tính lương" (hiện tại: ' + (b ? b.ten : k.trangThai) + ').');
      if (!laNguoiXuLy(s, b.nguoi)) err('Bước "' + b.ten + '" do ' + tenNguoi(b.nguoi) + ' thực hiện.');
      if (num(req.ngayCongChuan)) k.ngayCongChuan = num(req.ngayCongChuan);
      var r = recalc(ky, s);
      log(s, 'bl.calc', ky + ': ' + r.bl.length + ' người');
      var out = kyData(ky); out.thieu = r.thieu; return out;
    };

    H['ky.move'] = function (req, s) {
      var ky = req.ky; if (!validKy(ky)) err('Kỳ lương không hợp lệ.');
      var k = kyRow(ky, true), st = buocKy(k), from = k.trangThai, to = req.to;
      var i = viTriBuoc(st, from), j = viTriBuoc(st, to), note = String(req.note || '').trim(), hd;
      if (j < 0) err('Bước không hợp lệ.');
      var b = st[i];
      if (j > i && b && FIELDS_LOAI[b.loai] && Array.isArray(req.rows) && req.rows.length && laNguoiXuLy(s, b.nguoi)) {   // v3.32: lưu & chuyển bước trong 1 lần gọi (trước đây 2 lần, mỗi lần tính lại cả bảng lương)
        if (num(req.ngayCongChuan) && (b.loai === 'chamcong' || b.loai === 'tinhluong')) k.ngayCongChuan = num(req.ngayCongChuan);
        ghiCC(ky, b, req.rows); D('ChamCong');
      }
      if (to === BUOC_XONG && j > i + 1 && i >= 0 && (s.role === 'giamdoc' || s.role === 'admin')) {
        // Giám đốc phê duyệt thẳng: bỏ qua các bước còn lại
        if (!T('ChamCong').some(function (x) { return x.ky === ky; })) err('Chưa có dữ liệu chấm công kỳ ' + ky + '.');
        var rr = recalc(ky, s);
        if (rr.thieu.length) err('Còn ' + rr.thieu.length + ' người chưa có chấm công: ' + rr.thieu.slice(0, 5).join(', ') + '.');
        if (!rr.bl.length) err('Bảng lương trống.');
        b = { nguoi: ['r:' + s.role], loai: 'pheduyet' }; i = j - 1;
      }
      if (j === i + 1) {
        if (!laNguoiXuLy(s, b.nguoi)) err('Bước "' + b.ten + '" do ' + tenNguoi(b.nguoi) + ' thực hiện.');
        if (b.loai === 'chamcong' && !T('ChamCong').some(function (x) { return x.ky === ky; })) err('Chưa có dữ liệu chấm công kỳ ' + ky + '.');
        if (b.loai !== 'chamcong' && b.loai !== 'pheduyet') {
          var r = recalc(ky, s);
          if (r.thieu.length) err('Còn ' + r.thieu.length + ' người chưa có chấm công: ' + r.thieu.slice(0, 5).join(', ') + (r.thieu.length > 5 ? '…' : '') + '. Trả lại bước Chấm công để bổ sung.');
          if (!r.bl.length) err('Bảng lương trống.');
        }
        if (to === BUOC_XONG) {
          var chuaXM = huuChuaXacMinh(ky, T('BangLuong'), T('NhanVien'), cfgTheoKy(getCfg(), ky));       // v3.29
          if (chuaXM.length) err('Chưa đủ căn cứ xác định nghĩa vụ BHXH của ' + chuaXM.length + ' người hưởng lương hưu: ' + chuaXM.slice(0, 5).map(function (n) { return n.hoTen; }).join(', ') + '. Kế toán xác minh tình trạng hưởng lương hưu (Hồ sơ → Tối ưu thuế TNCN → Nhân sự hưởng lương hưu) trước khi khóa bảng lương.');
          var bl = T('BangLuong').filter(function (x) { return x.ky === ky; }).sort(function (a, c) { return a.maNV < c.maNV ? -1 : 1; });
          k.soNguoi = bl.length;
          k.tongThuNhap = bl.reduce(function (t, x) { return t + num(x.tongThuNhap); }, 0);
          k.tongThucLinh = bl.reduce(function (t, x) { return t + num(x.thucLinh); }, 0);
          k.maKiemTra = env.hash(JSON.stringify(bl.map(function (x) { return [x.maNV, x.tongThuNhap, x.bhxh, x.thueTNCN, x.thucLinh]; }))).slice(0, 12).toUpperCase();
          k.nguoiDuyet = s.hoTen || s.username; k.ngayDuyet = env.now();
          hd = viTriBuoc(st, from) < st.length - 1 ? 'Giám đốc phê duyệt (bỏ qua bước: ' + st.slice(viTriBuoc(st, from)).filter(function (x) { return x.loai !== 'pheduyet'; }).map(function (x) { return x.ten; }).join(', ') + ')' : 'Phê duyệt';
        } else hd = i === 0 ? 'Gửi' : (st[j].loai === 'pheduyet' ? 'Trình duyệt' : 'Chuyển');
      } else if (j < i) {
        if (from === BUOC_XONG) { if (s.role !== 'admin') err('Chỉ Quản trị được mở khóa kỳ lương đã phê duyệt.'); hd = 'Mở khóa'; }
        else { if (!laNguoiXuLy(s, b.nguoi)) err('Chỉ người xử lý bước "' + b.ten + '" mới được trả lại.'); hd = 'Trả lại'; }
        if (!note) err('Nhập lý do ' + hd.toLowerCase() + '.');
        k.nguoiDuyet = ''; k.ngayDuyet = ''; k.maKiemTra = '';
      } else err('Chỉ chuyển sang bước kế tiếp hoặc trả lại bước trước.');
      k.trangThai = to; k.capNhat = env.now(); D('KyLuong');
      ghiPD(k, from, to, hd, s, note);
      log(s, 'ky.move', ky + ': ' + from + ' → ' + to + (note ? ' (' + note + ')' : ''));
      var bMoi = st[viTriBuoc(st, to)], tenKy = ky.slice(5) + '/' + ky.slice(0, 4);
      if (to === BUOC_XONG) bao(T('BangLuong').filter(function (x) { return x.ky === ky; }).map(function (x) { return userCuaNV(x.maNV); }), { loai: 'duyet', t: 'Phiếu lương tháng ' + tenKy, b: 'Bảng lương đã được phê duyệt. Bấm để xem phiếu lương của bạn.', tag: 'ky_' + ky, mo: '#v=slip' });
      else if (bMoi && bMoi.nguoi) bao(nguoiCo(bMoi.nguoi, s.username), { loai: 'duyet', t: 'Kỳ lương ' + tenKy + ' chờ bạn xử lý', b: 'Bước "' + to + '"' + (hd ? ' · ' + hd : '') + (note ? ': ' + note : ''), tag: 'ky_' + ky, mo: '#v=luong' });
      return kyData(ky);
    };

    /* ---- đơn xin nghỉ: các cấp duyệt theo cấu hình, chụp lại tại thời điểm tạo đơn ---- */
    function kyKhoa(tu, den) {
      var out = [];
      T('KyLuong').forEach(function (k) {
        var st = buocKy(k), t = normBuoc(k.trangThai, st);
        if (t !== st[0].id && kyStart(k.ky) <= den && kyEnd(k.ky) >= tu) out.push(k.ky);
      });
      return out;
    }
    function kiemTraPhep(o, cfg, boQua) {
      if (o.loai !== 'Nghỉ phép năm') return;
      var nam = +o.tuNgay.slice(0, 4);
      if (o.denNgay.slice(0, 4) !== o.tuNgay.slice(0, 4)) err('Đơn nghỉ phép năm không được kéo sang năm sau. Vui lòng tách thành 2 đơn.');
      var nv = findNV(o.maNV), quy = quyPhepNam(nv, nam, cfg);
      var dd = phepNamDaDung(o.maNV, nam, T('NghiPhep').filter(function (x) { return x !== boQua; }), cfg);
      var con = quy - dd.daDung - dd.cho;
      if (o.soNgay > con) err('Không đủ phép năm ' + nam + ': quỹ ' + quy + ' ngày, đã nghỉ ' + dd.daDung + ', đang chờ duyệt ' + dd.cho + ', còn ' + Math.max(0, con) + ' ngày. Phần vượt hãy xin "Nghỉ không lương".');
    }
    function capDuyetCho(o, cfg) {
      var qt = cfg.quyTrinhNghi || QT_NGHI_MD, nguoiNop = T('Users').filter(function (u) { return u.maNV === o.maNV; }).map(function (u) { return 'u:' + u.username; });
      var caps = qt.filter(function (c) { return o.soNgay > num(c.tren) && (!c.loai || !c.loai.length || c.loai.indexOf(o.loai) >= 0); })
        .map(function (c) { var ng = quyTbp(c.nguoi, o.maNV, cfg); return { id: c.id, ten: c.ten, nguoi: ng, nguoiTen: tenNguoi(ng) }; })
        .filter(function (c) { return c.nguoi.some(function (a) { return nguoiNop.indexOf(a) < 0; }); });   // bỏ cấp mà người duyệt chính là người nộp
      caps = caps.filter(function (c, i) { return !i || JSON.stringify(c.nguoi) !== JSON.stringify(caps[i - 1].nguoi); });  // bỏ cấp trùng người với cấp ngay trước
      if (!caps.length) { var c0 = qt[0] || QT_NGHI_MD[0], ng0 = quyTbp(c0.nguoi, o.maNV, cfg); caps = [{ id: c0.id, ten: c0.ten, nguoi: ng0, nguoiTen: tenNguoi(ng0) }]; }
      return caps;
    }
    function chuanHoaDon(o, cfg) {           // đơn tạo từ phiên bản cũ (Chờ GĐ duyệt / nguoiDuyet1)
      if (o.luongDuyet) return;
      var caps = capDuyetCho(o, cfg);
      o.luongDuyet = JSON.stringify(caps);
      o.capDuyet = o.trangThai === 'Chờ GĐ duyệt' ? caps.length - 1 : 0;
      if (o.trangThai === 'Chờ GĐ duyệt') o.trangThai = 'Chờ duyệt';
      o.lichSu = o.lichSu || '[]';
    }
    /* đơn nghỉ được duyệt / hủy khi kỳ đang chấm công: cập nhật bảng chấm công ngày đã lưu */
    function dongBoNgayDon(o) {
      var cfg = getCfg(), ky = o.tuNgay.slice(0, 7), g = 0;
      for (; ky <= o.denNgay.slice(0, 7) && g < 14; ky = addMonths(ky + '-01', 1).slice(0, 7), g++) {
        var c = T('ChamCong').filter(function (x) { return x.id === ky + '|' + o.maNV; })[0];
        if (!c || !c.chiTietNgay) continue;
        var ct = parseJSON(c.chiTietNgay, {}), khoa = ngayTheoDon(T('NghiPhep'), o.maNV, ky, cfg);
        Object.keys(ct).forEach(function (d) { if (KY_HIEU[ct[d]] && KY_HIEU[ct[d]].don && !khoa[d]) ct[d] = 'X'; });
        ct = cleanNgay(ct, ky, khoa);
        var th = tongHopNgay(ct, ky, cfg);
        c.chiTietNgay = JSON.stringify(ct);
        ['ngayCong', 'nghiHuongLuong', 'nghiKhongLuong', 'nghiOm', 'ngayLe', 'ot300'].forEach(function (f) { c[f] = th[f]; });
        if (cfg.chuNhatTinh === 'ot') c.ot200 = th.ot200;
        D('ChamCong');
      }
    }
    H['np.create'] = function (req, s) {
      var r = req.row || {}, ma = r.maNV || s.maNV;
      if (!ma) err('Tài khoản chưa gắn với hồ sơ CBCNV.');
      if (ma !== s.maNV && !can(s.role, 'edit_nv')) err('Chỉ TCHC được tạo đơn nghỉ thay cho người khác.');
      var nv = findNV(ma); if (!nv) err('Không tìm thấy nhân viên.');
      if (LOAI_NGHI.indexOf(r.loai) < 0) err('Chọn loại nghỉ.');
      if (!parseD(r.tuNgay) || !parseD(r.denNgay)) err('Nhập ngày bắt đầu và kết thúc.');
      if (r.denNgay < r.tuNgay) err('Ngày kết thúc phải từ ngày bắt đầu trở đi.');
      if (r.denNgay > addMonths(r.tuNgay, 12)) err('Một đơn nghỉ tối đa 12 tháng.');
      var cfg = getCfg();
      var o = { id: newId(), maNV: ma, hoTen: nv.hoTen, loai: r.loai, tuNgay: r.tuNgay, denNgay: r.denNgay,
        buoi: r.tuNgay === r.denNgay && (r.buoi === 'Sáng' || r.buoi === 'Chiều') ? r.buoi : 'Cả ngày',
        lyDo: String(r.lyDo || '').trim().slice(0, 1000), chungTu: String(r.chungTu || '').trim().slice(0, 300),
        trangThai: 'Chờ duyệt', nguoiTao: s.hoTen || s.username, taoLuc: env.now() };
      o.soNgay = soNgayDon(o, cfg);
      if (!o.soNgay) err('Khoảng thời gian chọn chỉ gồm Chủ nhật / ngày lễ, không cần xin nghỉ.');
      if (!o.lyDo && o.loai !== 'Nghỉ phép năm') err('Nhập lý do nghỉ.');
      var trung = T('NghiPhep').filter(function (x) {
        return x.maNV === ma && ['Chờ duyệt', 'Chờ GĐ duyệt', 'Đã duyệt'].indexOf(x.trangThai) >= 0 && x.tuNgay <= o.denNgay && x.denNgay >= o.tuNgay;
      })[0];
      if (trung) err('Trùng thời gian với đơn ' + trung.loai + ' ' + trung.tuNgay + ' → ' + trung.denNgay + ' (' + trung.trangThai + ').');
      var khoa = kyKhoa(o.tuNgay, o.denNgay);
      if (khoa.length) err('Kỳ lương ' + khoa.join(', ') + ' đã qua bước chấm công. Liên hệ TCHC để điều chỉnh.');
      kiemTraPhep(o, cfg);
      o.luongDuyet = JSON.stringify(capDuyetCho(o, cfg)); o.capDuyet = 0; o.lichSu = '[]';
      T('NghiPhep').push(o); D('NghiPhep');
      log(s, 'np.create', ma + ' ' + o.loai + ' ' + o.tuNgay + '→' + o.denNgay + ' (' + o.soNgay + ' ngày)');
      npBao(o, s);
      return o;
    };
    function npBao(o, s) {                             // đơn đến lượt cấp duyệt nào thì báo người của cấp đó
      var cap = parseJSON(o.luongDuyet, [])[num(o.capDuyet)]; if (!cap) return;
      bao(nguoiCo(cap.nguoi, s.username).filter(function (u) { return u !== userCuaNV(o.maNV); }), { loai: 'duyet', t: 'Đơn nghỉ chờ bạn duyệt', b: o.hoTen + ': ' + o.loai + ' ' + ngayVN(o.tuNgay) + (o.denNgay !== o.tuNgay ? ' → ' + ngayVN(o.denNgay) : '') + ' (' + o.soNgay + ' ngày)', tag: 'np_' + o.id, mo: '#v=np' });
    }
    H['np.decide'] = function (req, s) {
      var o = T('NghiPhep').filter(function (x) { return x.id === req.id; })[0];
      if (!o) err('Không tìm thấy đơn.');
      var cfg = getCfg(); chuanHoaDon(o, cfg);
      if (o.trangThai !== 'Chờ duyệt') err('Đơn đã được xử lý (' + o.trangThai + ').');
      var caps = parseJSON(o.luongDuyet, []), ci = num(o.capDuyet), cap = caps[ci];
      if (!cap) err('Đơn không còn cấp duyệt.');
      var note = String(req.note || '').trim(), duocGiao = laNguoiXuLy(s, cap.nguoi), vuotCap = !duocGiao && (s.role === 'giamdoc' || s.role === 'admin');
      if (o.maNV === s.maNV && s.role !== 'admin') err('Không tự duyệt đơn nghỉ của chính mình.');
      if (!duocGiao && !vuotCap) err('Đơn đang chờ "' + cap.ten + '" (' + cap.nguoiTen + ').');
      var ls = parseJSON(o.lichSu, []);
      ls.push({ cap: cap.ten, nguoi: s.hoTen || s.username, luc: env.now(), ketQua: req.approve ? (vuotCap ? 'Duyệt (vượt cấp)' : 'Duyệt') : 'Từ chối', ghiChu: note });
      if (!req.approve) {
        if (!note) err('Nhập lý do từ chối.');
        o.trangThai = 'Từ chối';
      } else {
        var khoa = kyKhoa(o.tuNgay, o.denNgay);
        if (khoa.length) err('Kỳ lương ' + khoa.join(', ') + ' đã qua bước chấm công, không duyệt thêm đơn nghỉ được. Trả kỳ lương về bước Chấm công nếu cần.');
        kiemTraPhep(o, cfg, o);
        if (vuotCap || ci + 1 >= caps.length) o.trangThai = 'Đã duyệt';
        else o.capDuyet = ci + 1;
      }
      if (o.trangThai !== 'Chờ duyệt') { o.nguoiDuyet = s.hoTen || s.username; o.ngayDuyet = env.now(); }
      if (note) o.ghiChuDuyet = (o.ghiChuDuyet ? o.ghiChuDuyet + ' | ' : '') + (s.hoTen || s.username) + ': ' + note;
      o.lichSu = JSON.stringify(ls);
      D('NghiPhep');
      if (o.trangThai === 'Đã duyệt') dongBoNgayDon(o);
      log(s, 'np.decide', o.maNV + ' ' + o.tuNgay + ' → ' + (o.trangThai === 'Chờ duyệt' ? 'chuyển ' + caps[o.capDuyet].ten : o.trangThai));
      if (o.trangThai === 'Chờ duyệt') npBao(o, s);
      else bao([userCuaNV(o.maNV)].filter(function (u) { return u !== s.username; }), { loai: 'duyet', t: 'Đơn nghỉ ' + (o.trangThai === 'Đã duyệt' ? 'đã được duyệt' : 'bị từ chối'), b: o.loai + ' ' + ngayVN(o.tuNgay) + (o.denNgay !== o.tuNgay ? ' → ' + ngayVN(o.denNgay) : '') + (note ? ' · ' + note : ''), tag: 'np_' + o.id, mo: '#v=mynp' });
      return o;
    };
    H['np.cancel'] = function (req, s) {
      var o = T('NghiPhep').filter(function (x) { return x.id === req.id; })[0];
      if (!o) err('Không tìm thấy đơn.');
      chuanHoaDon(o, getCfg());
      var laChu = o.maNV === s.maNV, ql = can(s.role, 'edit_nv') || s.role === 'giamdoc';
      if (['Từ chối', 'Đã hủy'].indexOf(o.trangThai) >= 0) err('Đơn đã kết thúc.');
      if (!ql) {
        if (!laChu) err('Không có quyền hủy đơn này.');
        if (o.trangThai === 'Đã duyệt' && o.tuNgay <= env.today()) err('Đơn đã duyệt và đã bắt đầu, liên hệ TCHC để hủy.');
      }
      var khoa = kyKhoa(o.tuNgay, o.denNgay);
      if (o.trangThai === 'Đã duyệt' && khoa.length) err('Kỳ lương ' + khoa.join(', ') + ' đã qua bước chấm công, không hủy được.');
      var daDuyet = o.trangThai === 'Đã duyệt';
      o.trangThai = 'Đã hủy';
      if (daDuyet) dongBoNgayDon(o);
      o.ghiChuDuyet = (o.ghiChuDuyet ? o.ghiChuDuyet + ' | ' : '') + 'Hủy bởi ' + (s.hoTen || s.username) + (req.note ? ': ' + String(req.note).trim() : '');
      D('NghiPhep');
      log(s, 'np.cancel', o.maNV + ' ' + o.tuNgay);
      return o;
    };

    /* ---- ý kiến CBCNV về phiếu lương: người tiếp nhận theo cấu hình ---- */
    function ykMsg(s, noiDung, extra) {
      var m = { tu: s.username, ten: s.hoTen || s.username, vaiTro: ROLES[s.role] || s.role, luc: env.now(), noiDung: noiDung };
      if (extra) m.ghiChu = extra;
      return m;
    }
    function ykXuLy(o) { var x = parseJSON(o.xuLy, null); return Array.isArray(x) && x.length ? x : ['r:' + (o.boPhan || 'hr')]; }
    H['yk.create'] = function (req, s) {
      if (!s.maNV) err('Tài khoản chưa gắn với hồ sơ CBCNV.');
      var ky = req.ky, k = kyRow(ky, false);
      if (!k || k.trangThai !== BUOC_XONG) err('Chỉ gửi ý kiến cho phiếu lương đã được phê duyệt.');
      if (!T('BangLuong').some(function (b) { return b.ky === ky && b.maNV === s.maNV; })) err('Không có phiếu lương kỳ ' + ky + '.');
      var cfg = getCfg(), cd = (cfg.quyTrinhYKien || QT_YKIEN_MD).filter(function (c) { return c.chuDe === req.chuDe; })[0];
      if (!cd) err('Chọn nội dung cần xem xét.');
      var nd = String(req.noiDung || '').trim();
      if (nd.length < 10) err('Vui lòng mô tả rõ đề nghị (tối thiểu 10 ký tự).');
      if (nd.length > 2000) err('Nội dung tối đa 2.000 ký tự.');
      var dangMo = T('YKien').filter(function (r) { return r.maNV === s.maNV && r.ky === ky && (r.trangThai === 'Mới' || r.trangThai === 'Đang xử lý'); }).length;
      if (dangMo >= 5) err('Bạn đang có 5 ý kiến chưa xử lý cho kỳ này. Vui lòng chờ phản hồi.');
      var nv = findNV(s.maNV), xl = quyTbp(cd.nguoi, s.maNV, cfg);
      var o = { id: newId(), ky: ky, maNV: s.maNV, hoTen: nv ? nv.hoTen : s.hoTen, chuDe: req.chuDe, xuLy: JSON.stringify(xl), boPhan: tenNguoi(xl),
        trangThai: 'Mới', noiDung: nd, thread: JSON.stringify([ykMsg(s, nd)]), taoLuc: env.now(), capNhat: env.now(), nguoiXuLy: '' };
      T('YKien').push(o); D('YKien');
      log(s, 'yk.create', ky + ' ' + s.maNV + ': ' + req.chuDe);
      bao(nguoiCo(xl, s.username), { loai: 'yk', t: 'Ý kiến mới về phiếu lương ' + ky.slice(5) + '/' + ky.slice(0, 4), b: o.hoTen + ' – ' + o.chuDe + ': ' + nd.slice(0, 160), tag: 'yk_' + o.id, mo: '#v=yk' });
      return o;
    };
    H['yk.reply'] = function (req, s) {
      var o = T('YKien').filter(function (r) { return r.id === req.id; })[0];
      if (!o) err('Không tìm thấy ý kiến.');
      var xl = ykXuLy(o), laChu = o.maNV && o.maNV === s.maNV;
      var xuLy = s.role === 'giamdoc' || laNguoiXuLy(s, xl);
      if (!laChu && !xuLy) err('Ý kiến này do ' + tenNguoi(xl) + ' xử lý.');
      var nd = String(req.noiDung || '').trim();
      if (nd.length > 2000) err('Nội dung tối đa 2.000 ký tự.');
      var th = parseJSON(o.thread, []), ghiChu = [];
      if (xuLy && !laChu) {
        var chuyen = req.chuyen || (req.boPhan && req.boPhan.indexOf(':') < 0 ? 'r:' + req.boPhan : req.boPhan);
        if (/^r:/.test(chuyen || '') && ROLE_CU[chuyen.slice(2)]) chuyen = 'r:' + ROLE_CU[chuyen.slice(2)];
        if (chuyen && JSON.stringify([chuyen]) !== JSON.stringify(xl)) {
          var hop = /^r:/.test(chuyen) ? (ROLES[chuyen.slice(2)] && chuyen !== 'r:nhanvien') : T('Users').some(function (u) { return 'u:' + u.username === chuyen && u.active !== 'Không'; });
          if (!hop) err('Người / bộ phận nhận chuyển không hợp lệ.');
          var tenCu = tenNguoi(xl);
          o.xuLy = JSON.stringify([chuyen]); o.boPhan = tenNguoi([chuyen]);
          ghiChu.push('Chuyển từ ' + tenCu + ' sang ' + o.boPhan);
          if (o.trangThai !== 'Mới') o.trangThai = 'Đang xử lý';
        }
        var tt = req.trangThai || (o.trangThai === 'Mới' ? 'Đang xử lý' : o.trangThai);
        if (YK_TT.indexOf(tt) < 0) err('Trạng thái không hợp lệ.');
        if ((tt === 'Đã giải quyết' || tt === 'Không chấp nhận') && !nd) err('Nhập nội dung trả lời / kết quả xử lý cho CBCNV.');
        if (tt !== o.trangThai) ghiChu.push(o.trangThai + ' → ' + tt);
        o.trangThai = tt; o.nguoiXuLy = s.hoTen || s.username;
      } else {
        if (!nd) err('Nhập nội dung.');
        if (o.trangThai === 'Đã giải quyết' || o.trangThai === 'Không chấp nhận') { o.trangThai = 'Đang xử lý'; ghiChu.push('CBCNV mở lại ý kiến'); }
      }
      if (!nd && !ghiChu.length) err('Không có gì thay đổi.');
      th.push(ykMsg(s, nd, ghiChu.join('; ')));
      o.thread = JSON.stringify(th); o.capNhat = env.now(); D('YKien');
      log(s, 'yk.reply', o.maNV + ' ' + o.ky + (ghiChu.length ? ' (' + ghiChu.join('; ') + ')' : ''));
      if (xuLy && !laChu) bao([userCuaNV(o.maNV)], { loai: 'yk', t: 'Phản hồi ý kiến phiếu lương · ' + o.trangThai, b: (s.hoTen || s.username) + ': ' + (nd || ghiChu.join('; ')).slice(0, 180), tag: 'yk_' + o.id, mo: '#v=slip' });
      else bao(nguoiCo(ykXuLy(o), s.username), { loai: 'yk', t: o.hoTen + ' trả lời ý kiến phiếu lương', b: nd.slice(0, 180), tag: 'yk_' + o.id, mo: '#v=yk' });
      return o;
    };

    /* ---- công thức lương (Quản trị) ---- */
    function chuanCT(req) {
      return {
        congThucLuong: (req.congThucLuong || []).map(function (c) {
          var nhom = c.nhom === 'khautru' || c.nhom === 'trunggian' ? c.nhom : 'thunhap';
          if (nhom === 'trunggian') return { ma: String(c.ma || '').trim().toUpperCase(), ten: String(c.ten || '').trim(), nhom: nhom, congThuc: String(c.congThuc || '').trim(),
            vaiTro: CT_VAITRO[c.vaiTro] ? c.vaiTro : '', an: !!c.an, ghiChu: String(c.ghiChu || '').slice(0, 300) };
          return { ma: String(c.ma || '').trim().toUpperCase(), ten: String(c.ten || '').trim(), nhom: nhom,
            phanLoai: CT_PHANLOAI[c.phanLoai] ? c.phanLoai : 'khac', congThuc: String(c.congThuc || '').trim(),
            thue: ['toanbo', 'khong', 'mientoida', 'congthuc'].indexOf(c.thue) >= 0 ? c.thue : 'toanbo',
            thueGT: c.thue === 'mientoida' ? num(c.thueGT) : (c.thue === 'congthuc' ? String(c.thueGT || '').trim() : ''),
            giamTruQC: !!c.giamTruQC, an: !!c.an, ghiChu: String(c.ghiChu || '').slice(0, 300) };
        }),
        bienKy: (req.bienKy || []).map(function (v) { return { ma: String(v.ma || '').trim().toUpperCase(), ten: String(v.ten || '').trim(), buoc: v.buoc, macDinh: num(v.macDinh) }; }),
        congThucDongBH: String(req.congThucDongBH || '').trim() || CT_DONGBH_MD
      };
    }
    /* khoản phụ cấp PC_xxx có trong danh mục mẫu mà công thức dùng nhưng danh mục hiện tại chưa có → tự thêm */
    function pcThieu(cfg, ct) {
      var co = {}, out = [];
      (cfg.phuCap || []).forEach(function (p) { co['PC_' + String(p.ma).toUpperCase()] = 1; });
      (ct.congThucLuong || []).forEach(function (c) {
        var ds = [];
        try { ds = bienDung(phanTichCT(c.congThuc)); } catch (e) { }
        ds.forEach(function (v) {
          if (!/^PC_/.test(v) || co[v]) return;
          var m = DEFAULT_CFG.phuCap.filter(function (p) { return 'PC_' + p.ma.toUpperCase() === v; })[0];
          if (m) { co[v] = 1; out.push(clone(m)); }
        });
      });
      return out;
    }
    H['congthuc.save'] = function (req, s) {
      need(s, 'congthuc');
      var ct = chuanCT(req), them = pcThieu(getCfg(), ct);
      if (them.length) ct.phuCap = (getCfg().phuCap || []).concat(them);
      var thu = Object.assign(getCfg(), ct), loi = kiemTraCT(thu);
      if (loi.length) err(loi.slice(0, 6).join(' | '));
      setCfg(ct);
      var tinhLai = [];
      if (req.tinhLaiKyMo) T('KyLuong').forEach(function (k) { if (canTinhLai(kyRow(k.ky, false))) { recalc(k.ky, s); tinhLai.push(k.ky); } });
      log(s, 'congthuc.save', ct.congThucLuong.length + ' khoản' + (tinhLai.length ? '; tính lại kỳ ' + tinhLai.join(', ') : ''));
      return { cfg: trangTriQT(getCfg()), tinhLai: tinhLai, data: tinhLai.length ? { bl: T('BangLuong'), ky: T('KyLuong') } : null };
    };
    H['congthuc.capnhatot'] = function (req, s) {    // v3.12: áp quy định mới về thuế TNCN tiền làm thêm giờ vào bộ công thức đang dùng
      need(s, 'congthuc');
      var cfg = getCfg(), v = otVaCT(cfg.congThucLuong || CT_LUONG_MD);
      var loi = v.doi.filter(function (d) { return d.loi; }).map(function (d) { return d.loi; });
      var doi = v.doi.filter(function (d) { return !d.loi; });
      if (req.xem) return { doi: doi, loi: loi };
      if (!doi.length) return { doi: [], loi: loi, cfg: trangTriQT(cfg) };
      var thu = Object.assign({}, cfg, { congThucLuong: v.ds }), lk = kiemTraCT(thu);
      if (lk.length) err('Không áp dụng được: ' + lk.slice(0, 4).join(' | '));
      var ct = { congThucLuong: v.ds };
      if (!(num(cfg.otGioiHanThang) > 0)) ct.otGioiHanThang = 40;
      if (!(num(cfg.otGioiHanNam) > 0)) ct.otGioiHanNam = 200;
      setCfg(ct);
      var tinhLai = [];
      if (req.tinhLaiKyMo) T('KyLuong').forEach(function (k) { if (canTinhLai(kyRow(k.ky, false))) { recalc(k.ky, s); tinhLai.push(k.ky); } });
      log(s, 'congthuc.capnhatot', doi.map(function (d) { return d.ma; }).join(', ') + ' – miễn thuế TNCN tiền làm thêm giờ (Luật 109/2025, NĐ 253/2026)' + (tinhLai.length ? '; tính lại kỳ ' + tinhLai.join(', ') : ''));
      return { doi: doi, loi: loi, cfg: trangTriQT(getCfg()), tinhLai: tinhLai, data: tinhLai.length ? { bl: T('BangLuong'), ky: T('KyLuong') } : null };
    };
    H['congthuc.capnhatanca'] = function (req, s) {  // v3.13.1: quy chế ăn ca – chỉ ngày làm đủ 8 giờ
      need(s, 'congthuc');
      var cfg = getCfg(), v = anCaVaCT(cfg.congThucLuong || CT_LUONG_MD, cfg.bienKy || BIEN_KY_MD);
      if (req.xem || !v.doi.length) return { doi: v.doi, cfg: req.xem ? undefined : trangTriQT(cfg) };
      var thu = Object.assign({}, cfg, { congThucLuong: v.ds, bienKy: v.bienKy }), lk = kiemTraCT(thu);
      if (lk.length) err('Không áp dụng được: ' + lk.slice(0, 4).join(' | '));
      setCfg({ congThucLuong: v.ds, bienKy: v.bienKy });
      var tinhLai = [];
      if (req.tinhLaiKyMo) T('KyLuong').forEach(function (k) { if (canTinhLai(kyRow(k.ky, false))) { recalc(k.ky, s); tinhLai.push(k.ky); } });
      log(s, 'congthuc.capnhatanca', v.doi.map(function (d) { return d.ma; }).join(', ') + ' – ăn ca chỉ tính ngày làm đủ 8 giờ' + (tinhLai.length ? '; tính lại kỳ ' + tinhLai.join(', ') : ''));
      return { doi: v.doi, cfg: trangTriQT(getCfg()), tinhLai: tinhLai, data: tinhLai.length ? { bl: T('BangLuong'), ky: T('KyLuong') } : null };
    };
    H['congthuc.test'] = function (req, s) {
      need(s, 'congthuc');
      var cfg = Object.assign(getCfg(), chuanCT(req)), them = pcThieu(getCfg(), cfg);
      if (them.length) cfg.phuCap = (cfg.phuCap || []).concat(them);
      var loi = kiemTraCT(cfg);
      if (loi.length) return { loi: loi };
      var nv = findNV(req.maNV); if (!nv) err('Chọn CBCNV để chạy thử.');
      var ky = validKy(req.ky) ? req.ky : env.today().slice(0, 7), k = T('KyLuong').filter(function (x) { return x.ky === ky; })[0];
      var ncc = num(k && k.ngayCongChuan) || (cfg.cachTinhCongChuan === 'truCN' ? congChuanTruCN(ky) : num(cfg.ngayCongChuan));
      var cc = T('ChamCong').filter(function (x) { return x.ky === ky && x.maNV === nv.maNV; })[0] || { ngayCong: ncc };
      var th = 0, thT = 0;
      T('KhenThuong').forEach(function (x) { if (x.maNV === nv.maNV && x.kyLuong === ky && x.trangThai === 'Đã duyệt' && x.loai !== 'Kỷ luật') { th += num(x.soTien); if (x.chiuThue !== 'Không') thT += num(x.soTien); } });
      var lk = nvTaiKy(nv, T('LenLuong'), T('MucHuong'), ky), o = { luong: lk.luong, luongBH: lk.luongBH, ncc: ncc, thuong: th, thuongChiuThue: thT, ky: ky, otLuyKe: otLuyKe(nv.maNV, ky, T('ChamCong')) };
      var bien = {}, p = tinhPhieuLuong(lk.nv, cc, cfg, o, bien);
      return { loi: [], ky: ky, coChamCong: !!cc.id, bien: bien, phieu: p, pcThem: them.map(function (x) { return x.ten; }) };
    };

    /* ---- cài đặt quy trình & người phê duyệt (Quản trị, Giám đốc) ---- */
    function kiemNguoi(nguoi, choTbp, ten) {
      if (!Array.isArray(nguoi) || !nguoi.length) err('"' + ten + '" chưa có người xử lý.');
      var us = T('Users');
      nguoi.forEach(function (a, i) { if (/^r:/.test(a) && ROLE_CU[a.slice(2)]) nguoi[i] = 'r:' + ROLE_CU[a.slice(2)]; });
      nguoi.forEach(function (a) {
        if (a === 'tbp') { if (!choTbp) err('"' + ten + '": không dùng trưởng bộ phận ở đây.'); return; }
        if (/^r:/.test(a)) { if (!ROLES[a.slice(2)] || a === 'r:nhanvien') err('"' + ten + '": vai trò không hợp lệ.'); return; }
        if (/^u:/.test(a)) { if (!us.some(function (u) { return 'u:' + u.username === a; })) err('"' + ten + '": không tìm thấy tài khoản ' + a.slice(2) + '.'); return; }
        err('"' + ten + '": người xử lý không hợp lệ.');
      });
    }
    H['quytrinh.save'] = function (req, s) {
      need(s, 'quytrinh');
      var out = {}, ghi = [];
      if (req.quyTrinhLuong) {
        var st = req.quyTrinhLuong.map(function (b) { return { id: String(b.id || newId()), ten: String(b.ten || '').trim(), loai: b.loai, nguoi: b.nguoi || [] }; });
        var e = kiemTraQTLuong(st); if (e) err(e);
        st.forEach(function (b) {
          kiemNguoi(b.nguoi, false, b.ten);
          if ((b.loai === 'chamcong' || b.loai === 'tinhluong' || b.loai === 'congno' || b.loai === 'vipham') && b.nguoi.some(function (a) {
            return /^u:/.test(a) && T('Users').some(function (u) { return 'u:' + u.username === a && u.role === 'nhanvien'; });
          })) err('Bước "' + b.ten + '" nhập/sửa dữ liệu lương: chỉ giao cho tài khoản có vai trò quản lý (không phải CBCNV).');
        });
        out.quyTrinhLuong = st; ghi.push('lương');
      }
      if (req.quyTrinhNghi) {
        var qn = req.quyTrinhNghi.map(function (c) { return { id: String(c.id || newId()), ten: String(c.ten || '').trim(), nguoi: c.nguoi || [], tren: Math.max(0, num(c.tren)), loai: (c.loai || []).filter(function (l) { return LOAI_NGHI.indexOf(l) >= 0; }) }; });
        if (!qn.length) err('Quy trình duyệt nghỉ cần ít nhất 1 cấp.');
        qn.forEach(function (c) { if (!c.ten) err('Có cấp duyệt nghỉ chưa đặt tên.'); kiemNguoi(c.nguoi, true, c.ten); });
        out.quyTrinhNghi = qn; ghi.push('nghỉ phép');
      }
      if (req.quyTrinhYKien) {
        var qy = req.quyTrinhYKien.map(function (c) { return { chuDe: String(c.chuDe || '').trim(), nguoi: c.nguoi || [] }; });
        if (!qy.length) err('Cần ít nhất 1 chủ đề ý kiến.');
        var seen = {};
        qy.forEach(function (c) { if (!c.chuDe) err('Có chủ đề ý kiến chưa đặt tên.'); if (seen[c.chuDe]) err('Trùng chủ đề "' + c.chuDe + '".'); seen[c.chuDe] = 1; kiemNguoi(c.nguoi, true, c.chuDe); });
        out.quyTrinhYKien = qy; ghi.push('ý kiến');
      }
      if (req.nhomDeXuat) {
        need(s, 'config');                            // Quản trị tạo nhóm, loại đề xuất và cài người duyệt
        var qn2 = req.nhomDeXuat.map(function (g) { return { id: String(g.id || newId()).trim(), ten: String(g.ten || '').trim() }; });
        var en = kiemTraNhomDX(qn2); if (en) err(en);
        out.nhomDeXuat = qn2; ghi.push('nhóm loại đề xuất');
      }
      if (req.loaiDeXuat) {
        need(s, 'config');
        var qd = req.loaiDeXuat.map(function (l) {
          return { id: String(l.id || newId()).trim(), ten: String(l.ten || '').trim(), tien: !!l.tien,
            nhom: String(l.nhom || '').trim(), tienTo: String(l.tienTo || '').trim().toUpperCase(),
            cach: DX_CACH[l.cach] ? l.cach : 'trinhtu', gdCuoi: l.gdCuoi !== false, hoatDong: l.hoatDong !== false,
            buoc: (l.buoc || []).filter(function (b) { return !b.chot; })
              .map(function (b) { return { id: String(b.id || newId()), ten: String(b.ten || '').trim(), nguoi: b.nguoi || [], tren: Math.max(0, num(b.tren)) }; }),
            theoDoi: l.theoDoi || [], thucHien: l.thucHien || [] };
        });
        var ed = kiemTraLoaiDX(qd, out.nhomDeXuat || getCfg().nhomDeXuat || NHOM_DX_MD); if (ed) err(ed);
        qd.forEach(function (l) {
          l.buoc.forEach(function (b) { kiemNguoi(b.nguoi, true, l.ten + ' – ' + b.ten); });
          if (l.theoDoi.length) kiemNguoi(l.theoDoi, false, l.ten + ' – người theo dõi');
          if (l.thucHien.length) kiemNguoi(l.thucHien, false, l.ten + ' – người thực hiện');
        });
        out.loaiDeXuat = qd; ghi.push('loại đề xuất');
      }
      if (req.truongBoPhan) {
        var tb = {};
        Object.keys(req.truongBoPhan).forEach(function (k) {
          var u = String(req.truongBoPhan[k] || '').trim(); if (!u) return;
          if (!T('Users').some(function (x) { return x.username === u; })) err('Không tìm thấy tài khoản ' + u + ' (trưởng ' + k + ').');
          tb[k] = u;
        });
        out.truongBoPhan = tb; ghi.push('trưởng bộ phận');
      }
      setCfg(out);
      var capNhatKy = [], giuKy = [];
      if (out.quyTrinhLuong && req.apDungKyDangXuLy) {
        var moi = trangTriQT(getCfg()).quyTrinhLuong;
        T('KyLuong').forEach(function (k) {
          if (normBuoc(k.trangThai, buocKy(k)) === BUOC_XONG) return;
          if (viTriBuoc(moi, k.trangThai) >= 0) { k.quyTrinh = JSON.stringify(moi); capNhatKy.push(k.ky); } else giuKy.push(k.ky);
        });
        D('KyLuong');
      }
      log(s, 'quytrinh.save', ghi.join(', ') + (capNhatKy.length ? '; áp dụng kỳ ' + capNhatKy.join(', ') : ''));
      return { cfg: trangTriQT(getCfg()), ky: T('KyLuong'), capNhatKy: capNhatKy, giuKy: giuKy };
    };

    /* ---- tài khoản ---- */
    H['user.save'] = function (req, s) {
      need(s, 'users');
      var u = upsertUser(s, req.row || {});
      log(s, 'user.save', u.username + ' (' + u.role + ')');
      return publicUser(u);
    };

    H['user.bulk'] = function (req, s) {
      need(s, 'users');
      var users = T('Users'), out = [];
      T('NhanVien').forEach(function (nv) {
        if (nv.trangThai === 'Đã nghỉ việc') return;
        if (users.some(function (u) { return u.maNV === nv.maNV; })) return;
        var un = tenDangNhap(nv.hoTen, function (x) { return users.some(function (u) { return u.username === x; }); }, nv.ngaySinh);
        var mk = req.password || matKhauTheoNV(nv, getCfg());
        out.push(Object.assign(publicUser(upsertUser(s, { username: un, role: 'nhanvien', maNV: nv.maNV, password: mk })), { mkBanDau: mk }));
      });
      log(s, 'user.bulk', out.length + ' tài khoản');
      return out;
    };

    H['user.resetpw'] = function (req, s) {
      need(s, 'doimk');
      var u = T('Users').filter(function (x) { return x.username === req.username; })[0];
      if (!u) err('Không tìm thấy tài khoản.');
      if (u.username === s.username) err('Đổi mật khẩu của chính mình ở mục "Đổi mật khẩu".');
      if (normRole(u.role) === 'admin' && s.role !== 'admin') err('Chỉ Quản trị đổi được mật khẩu tài khoản Quản trị.');
      var mk = String(req.password || '') || matKhauTheoNV(u.maNV ? findNV(u.maNV) : null, getCfg());
      if (mk.length < 6) err('Mật khẩu tối thiểu 6 ký tự.');
      u.salt = newId(); u.passHash = hashPw(mk, u.salt); u.mustChange = 'Có'; D('Users');
      if (env.sessions.xoaCua) env.sessions.xoaCua(u.username);
      log(s, 'user.resetpw', u.username + (req.password ? ' (mật khẩu mới)' : ' (về mặc định)'));
      return { username: u.username, matKhau: req.password ? '' : mk };
    };
    H['user.delete'] = function (req, s) {
      need(s, 'users');
      if (req.username === s.username) err('Không thể xóa tài khoản đang đăng nhập.');
      var a = T('Users');
      for (var i = a.length - 1; i >= 0; i--) if (a[i].username === req.username) {
        if (s.role !== 'admin' && normRole(a[i].role) !== 'nhanvien') err('Không có quyền xóa tài khoản này.');
        a.splice(i, 1);
      }
      D('Users'); log(s, 'user.delete', req.username); return true;
    };

    /* ---- cấu hình & nhật ký ---- */
    H['le.save'] = function (req, s) {
      if (!can(s.role, 'config') && !can(s.role, 'edit_nv')) err('Bạn không có quyền sửa lịch nghỉ lễ.');
      var ds = cleanLe(req.ds || []);
      setCfg({ ngayLeDs: ds, ngayLe: ds.map(function (x) { return x.ngay; }) });
      log(s, 'le.save', ds.length + ' ngày nghỉ lễ');
      return getCfg();
    };
    H['pq.save'] = function (req, s) {
      if (!can(s.role, 'sua_pq')) err('Tài khoản không có quyền sửa bảng phân quyền.');
      var cu = pqLamSach(getCfg().phanQuyen), moi = req.macDinh ? pqLamSach(null) : pqLamSach(req.pq || {}), doi = [];
      PQ_SUA.forEach(function (k) {
        Object.keys(ROLES).forEach(function (r) { var a = cu[k].indexOf(r) >= 0, b = moi[k].indexOf(r) >= 0; if (a !== b) doi.push((b ? '+' : '−') + ROLES[r] + ': ' + k); });
      });
      setCfg({ phanQuyen: req.macDinh ? {} : moi }); _pqDaAp = null; getCfg();
      log(s, 'pq.save', req.macDinh ? 'Về mặc định' + (doi.length ? ' (' + doi.length + ' thay đổi)' : '') : (doi.join('; ') || 'không đổi'));
      return { pq: pqLamSach(getCfg().phanQuyen), doi: doi.length };
    };
    H['cfg.save'] = function (req, s) {
      need(s, 'config');
      if (req.cfg) delete req.cfg.phanQuyen;
      if (req.cfg && req.cfg.mauPhieu != null) req.cfg.mauPhieu = mauGon(mauPhieu({ mauPhieu: req.cfg.mauPhieu }), 'phieu');          // v3.41
      if (req.cfg && req.cfg.mauBangLuong != null) req.cfg.mauBangLuong = mauGon(mauBL({ mauBangLuong: req.cfg.mauBangLuong }), 'bl');
      if (req.cfg && req.cfg.anCaMienCT != null) {                       // v3.39: công thức ăn ca miễn thuế
        req.cfg.anCaMienCT = String(req.cfg.anCaMienCT).trim().slice(0, 500);
        if (req.cfg.anCaMienCT) { var lAC = kiemCTAnCa(Object.assign({}, getCfg(), req.cfg), req.cfg.anCaMienCT); if (lAC.length) err(lAC.join(' ')); }
      }           // bảng phân quyền chỉ sửa qua pq.save
      var cfTruoc = getCfg(), truocK = {}, sauK = {};
      Object.keys(req.cfg || {}).forEach(function (k) { truocK[k] = cfTruoc[k]; });
      setCfg(req.cfg || {});
      var cf = getCfg();                                // v3.26: bật "trả thêm khoản tương đương BH" → thêm khoản vào công thức lương nếu chưa có
      if (cf.traThayBH === 'co' && Array.isArray(cf.congThucLuong) && !cf.congThucLuong.some(function (c) { return !c.an && /TRA_THAY_BH/.test(String(c.congThuc)); })) {
        setCfg({ congThucLuong: cf.congThucLuong.concat([{ ma: 'TRATHAYBH', ten: 'Khoản trả thêm tương đương mức đóng BH (Điều 168 BLLĐ)', nhom: 'thunhap', phanLoai: 'phucap', congThuc: 'TRA_THAY_BH', thue: 'toanbo', giamTruQC: false, ghiChu: 'Người nghỉ hưu, HĐ thử việc không đóng BH: trả thêm = mức DN đóng BHXH, BHYT, BHTN; chịu thuế TNCN.' }]) });
      }
      var cfSau = getCfg(); Object.keys(req.cfg || {}).forEach(function (k) { sauK[k] = cfSau[k]; });
      logDoi(s, 'cfg.save', '', truocK, sauK, ['congThucLuong']);           // v3.27: ghi giá trị trước / sau
      return cfSau;
    };

    H['log.list'] = function (req, s) {
      need(s, 'log');
      return env.readLog ? env.readLog(num(req.limit) || 300) : [];
    };


    /* ==================== BOT CHÍNH SÁCH (v3.5) ====================
       Trợ lý tra cứu quy định – chế độ – biểu mẫu – chính sách công ty, đặt trong màn hình Tin nhắn.
       Hai nguồn trả lời, đều do phòng TCHC quản lý trong app:
         (1) BotQuyDinh – mục quy định ngắn TCHC tự nhập (câu hỏi thường gặp, trích yếu chế độ)
         (2) BotVanBan  – văn bản gốc trên Drive đã đăng ký (quy chế, quyết định); máy trích chữ, lưu đệm
       Bot KHÔNG tự biết gì về công ty: không có nguồn thì trả lời "chưa có quy định trong hệ thống"
       và hướng người hỏi sang TCHC. Không được bịa số liệu – CBCNV sẽ làm theo câu trả lời. */
    var BOT_NHOM = ['Lương & phụ cấp', 'Phép & nghỉ', 'Chấm công & giờ làm việc', 'Bảo hiểm & chế độ',
      'Biểu mẫu & hồ sơ', 'Quy định & kỷ luật', 'Đào tạo & bổ nhiệm', 'An toàn lao động', 'Khác'];
    var BOT_HE = 'Bạn là trợ lý tra cứu quy định nội bộ của Công ty Cổ phần TAPACO Việt Nam (xây dựng hạ tầng), trả lời CBCNV của công ty. ' +
      'Bạn CHỈ được trả lời dựa trên phần "NGUỒN" trong câu hỏi – đó là quy định và văn bản do công ty ban hành. ' +
      'Nguyên tắc tuyệt đối: KHÔNG được nêu bất kỳ con số, mức tiền, số ngày, tỷ lệ, thời hạn, số hiệu văn bản hay tên biểu mẫu nào không có trong NGUỒN. ' +
      'Không suy ra từ luật chung, từ thông lệ ngành hay từ công ty khác. Không dùng kiến thức ngoài NGUỒN, kể cả khi bạn tin là đúng. ' +
      'Nếu NGUỒN không đủ căn cứ: đặt daTraLoi = false, traLoi nói rõ công ty chưa có quy định này trong hệ thống và đề nghị liên hệ phòng TCHC, ' +
      'và ghi vào thieu nội dung quy định mà công ty cần bổ sung. Tuyệt đối không đưa con số phỏng đoán, không nói "thường là", "thông thường". ' +
      'Nếu NGUỒN chỉ trả lời được một phần: trả lời phần có căn cứ, nói rõ phần nào chưa có quy định. ' +
      'Mỗi câu trả lời phải ghi căn cứ vào nguon (đúng tiêu đề quy định hoặc số hiệu – tên văn bản đã dùng, không tự đặt tên mới). ' +
      'Việc riêng của từng người (lương cá nhân, số ngày phép còn lại, tranh chấp, kỷ luật cụ thể) thì không trả lời bằng số liệu mà hướng dẫn liên hệ TCHC hoặc trưởng bộ phận. ' +
      'Không suy đoán, không bình luận về sức khỏe, đời tư, gia đình, tôn giáo, chính trị của bất kỳ ai. ' +
      'Giọng lễ phép, ngắn gọn, rõ ràng; xưng "Anh/Chị" với người hỏi. Viết tiếng Việt. Chỉ trả về JSON.';
    var BOT_P = 'Trả về JSON đúng dạng: {"daTraLoi":true/false,"traLoi":"câu trả lời cho người hỏi","nguon":["tiêu đề quy định hoặc số hiệu – tên văn bản đã dùng"],"nhom":"' +
      BOT_NHOM.join('" hoặc "') + '","thieu":"nếu daTraLoi=false thì nêu ngắn gọn quy định công ty cần bổ sung, nếu không thì để rỗng"}. ' +
      'nguon phải là tên có thật trong NGUỒN. Nếu không dùng được nguồn nào thì daTraLoi=false và nguon=[].';
    var BOT_ROLES_QD = { indexOf: function (r) { return can(r, 'botqd') ? 0 : -1; } };
    var BOT_TMUC = '_QUY ĐỊNH CÔNG TY (KHÔNG CHIA SẺ)';
    /* âm tiết không mang chủ đề — bỏ khi so khớp, nếu không "công ty", "không", "thế nào"
       sẽ khớp với mọi quy định và trợ lý lấy nhầm căn cứ */
    var BOT_DUNG = { cua: 1, la: 1, co: 1, the: 1, nao: 1, gi: 1, va: 1, cho: 1, duoc: 1, bao: 1, toi: 1, minh: 1,
      khi: 1, thi: 1, ma: 1, voi: 1, trong: 1, den: 1, nhu: 1, sao: 1, can: 1, phai: 1, se: 1, da: 1, dang: 1,
      hay: 1, hoac: 1, ve: 1, moi: 1, nay: 1, do: 1, day: 1, mot: 1, nhung: 1, cac: 1, vay: 1, a: 1, ah: 1,
      em: 1, anh: 1, chi: 1, xin: 1, hoi: 1, khong: 1, ty: 1, ra: 1, vao: 1, lam: 1, nguoi: 1, muon: 1 };

    function botQuyen(s) { if (BOT_ROLES_QD.indexOf(s.role) < 0) err('Chỉ Giám đốc, TCHC hoặc Quản trị được sửa quy định của trợ lý.'); }
    function botS(v, n) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, n || 300); }
    function botHieuLuc(x) { return String(x.trangThai || 'hieuluc') !== 'het'; }
    /* bỏ dấu tiếng Việt để so khớp từ khóa (người hỏi hay gõ không dấu) */
    function botKhongDau(s) {
      return String(s == null ? '' : s).toLowerCase()
        .replace(/[àáạảãâầấậẩẫăằắặẳẵ]/g, 'a').replace(/[èéẹẻẽêềếệểễ]/g, 'e')
        .replace(/[ìíịỉĩ]/g, 'i').replace(/[òóọỏõôồốộổỗơờớợởỡ]/g, 'o')
        .replace(/[ùúụủũưừứựửữ]/g, 'u').replace(/[ỳýỵỷỹ]/g, 'y').replace(/đ/g, 'd');
    }
    function botTu(s) {
      return botKhongDau(s).replace(/[^a-z0-9\s]/g, ' ').split(/\s+/)
        .filter(function (w) { return w.length > 1 && !BOT_DUNG[w]; });
    }
    /* Độ liên quan giữa câu hỏi và một đoạn văn bản.
       So khớp theo TỪNG ÂM TIẾT trọn vẹn, không so chuỗi con: so chuỗi con cho kết quả sai
       (vd "trợ" khớp trong "trong", "công" khớp trong "công ty") khiến trợ lý lấy nhầm quy định
       không liên quan rồi trả lời sai. Cụm 2 âm tiết liền nhau được tính nặng hơn vì tiếng Việt
       ghép từ theo cụm ("xăng xe", "phép năm", "công trường").
       Trả về: d = điểm, phu = tỉ lệ âm tiết của câu hỏi tìm thấy trong văn bản (0..1). */
    function botKhop(tuHoi, vanBan) {
      var tv = botTu(vanBan), co = {}, bg = {}, i;
      for (i = 0; i < tv.length; i++) co[tv[i]] = 1;
      for (i = 1; i < tv.length; i++) bg[tv[i - 1] + ' ' + tv[i]] = 1;
      var d = 0, thay = {}, khop = 0;
      for (i = 0; i < tuHoi.length; i++) {
        var w = tuHoi[i];
        if (co[w] && !thay[w]) { thay[w] = 1; khop++; d += 2; }
      }
      for (i = 1; i < tuHoi.length; i++) if (bg[tuHoi[i - 1] + ' ' + tuHoi[i]]) d += 5;
      var rieng = Object.keys(thay).length, tong = 0, dem = {};
      for (i = 0; i < tuHoi.length; i++) if (!dem[tuHoi[i]]) { dem[tuHoi[i]] = 1; tong++; }
      return { d: d, phu: tong ? rieng / tong : 0 };
    }
    function botDiem(tuHoi, vanBan) { return botKhop(tuHoi, vanBan).d; }
    var BOT_PHU_MIN = 0.6, BOT_KHOP_MIN = 2;

    function botQDPub(x) {
      return { id: x.id, nhom: x.nhom, tieuDe: x.tieuDe, noiDung: x.noiDung, lienKet: x.lienKet, vanBan: x.vanBan,
        hieuLuc: x.hieuLuc, trangThai: x.trangThai || 'hieuluc', nguoiSua: x.nguoiSua, capNhat: x.capNhat };
    }
    function botVBPub(x) {
      return { id: x.id, ten: x.ten, fileId: x.fileId, mime: x.mime, nhom: x.nhom, soHieu: x.soHieu,
        ngayBanHanh: x.ngayBanHanh, trangThai: x.trangThai || 'hieuluc', soChu: num(x.soChu), textLuc: x.textLuc,
        loi: x.loi, nguoiSua: x.nguoiSua, capNhat: x.capNhat };
    }
    function botVBNhan(x) { return (x.soHieu ? x.soHieu + ' – ' : '') + x.ten; }

    /* ---- trích chữ từ văn bản trên Drive, lưu đệm thành tệp .txt trong thư mục riêng ---- */
    function botChuyenDoc(fileId) {          // PDF / Word → Google Docs tạm để lấy chữ, đọc xong xóa ngay
      var r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) + '/copy?supportsAllDrives=true', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({ name: 'tam-bot-' + Date.now(), mimeType: 'application/vnd.google-apps.document' })
      });
      if (r.getResponseCode() >= 300) err('Không chuyển được văn bản sang Google Docs để đọc chữ. Đề nghị mở tệp trên Drive, chọn "Mở bằng → Google Tài liệu", lưu lại rồi đăng ký tệp Google Tài liệu đó.');
      var id = (parseJSON(r.getContentText(), {}) || {}).id;
      if (!id) err('Không chuyển được văn bản sang Google Docs để đọc chữ.');
      var txt = '';
      try { txt = DocumentApp.openById(id).getBody().getText(); }
      finally { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { } }
      return txt;
    }
    function botTrichChu(fileId, mime) {
      if (mime === 'application/vnd.google-apps.document') return DocumentApp.openById(fileId).getBody().getText();
      if (/^text\/|json|csv|markdown/.test(String(mime))) return env.luuTru.doc(fileId);
      return botChuyenDoc(fileId);
    }
    function botNapChu(v, buoc) {            // buoc = true: bắt trích lại dù đã có đệm
      if (!env.luuTru) err('Máy chủ chưa bật lưu trữ Drive.');
      var tm = env.luuTru.thuMuc('', BOT_TMUC), dem = env.luuTru.thuMuc(tm, '_trich-xuat');
      var ten = 'vb-' + v.id + '.txt';
      if (!buoc && v.textLuc && num(v.soChu) > 0) {
        var co = env.luuTru.ds(dem).filter(function (f) { return f.ten === ten; })[0];
        if (co) { try { return env.luuTru.doc(co.id); } catch (e) { } }
      }
      var txt = botTrichChu(v.fileId, v.mime) || '';
      txt = String(txt).replace(/\r/g, '').replace(/\n{3,}/g, '\n\n').trim();
      env.luuTru.ghi(dem, ten, txt, 'text/plain');
      v.soChu = txt.length; v.textLuc = env.now(); v.loi = '';
      return txt;
    }

    H['bot.ds'] = function (req, s) {
      var qd = T('BotQuyDinh'), vb = T('BotVanBan'), qt = BOT_ROLES_QD.indexOf(s.role) >= 0;
      var tt = env.ai ? env.ai.trangThai() : null;
      return {
        nhom: BOT_NHOM, quanTri: qt,
        aiBat: !!(tt && tt.coKhoa), aiNcc: tt ? tt.ncc : '',
        soQD: qd.filter(botHieuLuc).length, soVB: vb.filter(botHieuLuc).length,
        goiY: qd.filter(botHieuLuc).slice(0, 300).map(function (x) { return { id: x.id, nhom: x.nhom, tieuDe: x.tieuDe }; }),
        qd: qt ? qd.map(botQDPub) : [],
        vb: qt ? vb.map(botVBPub) : vb.filter(botHieuLuc).map(function (x) { return { id: x.id, ten: x.ten, soHieu: x.soHieu, nhom: x.nhom }; }),
        tmuc: qt && env.luuTru ? env.luuTru.url(env.luuTru.thuMuc('', BOT_TMUC)) : ''
      };
    };

    /* hỏi bot: chỉ đọc, không giữ khóa ghi (gọi AI chạy lâu). Trình duyệt ghi nhật ký sau bằng bot.ghi */
    H['bot.hoi'] = function (req, s) {
      var hoi = botS(req.hoi, 500);
      if (hoi.length < 5) err('Xin nhập câu hỏi rõ hơn (ít nhất 5 ký tự).');
      var qd = T('BotQuyDinh').filter(botHieuLuc), vb = T('BotVanBan').filter(botHieuLuc);
      if (!qd.length && !vb.length) {
        return { daTraLoi: false, nguon: [], nhom: '', dungAI: false, thieu: 'Chưa nhập bất kỳ quy định nào',
          traLoi: 'Hệ thống chưa có quy định nào được nhập. Đề nghị Anh/Chị liên hệ phòng TCHC để được trả lời, và đề nghị TCHC nhập quy chế công ty vào mục "Quy định công ty" của trợ lý.' };
      }
      var tuHoi = botTu(hoi);

      /* chọn nguồn liên quan: toàn bộ mục quy định ngắn (xếp theo điểm), và tối đa 3 văn bản gốc liên quan nhất */
      var qdXep = qd.map(function (x) {
        var k = botKhop(tuHoi, x.tieuDe + ' ' + x.nhom + ' ' + x.noiDung);
        return { x: x, d: k.d, phu: k.phu };
      }).sort(function (a, b) { return b.d - a.d; });
      var dungQD = qdXep.filter(function (o) { return o.d > 0; }).slice(0, 25);
      if (!dungQD.length) dungQD = qdXep.slice(0, 8);          // không khớp từ nào: vẫn đưa vài mục để AI tự xét
      var txtQD = dungQD.map(function (o, i) {
        return (i + 1) + ') [' + (o.x.nhom || 'Khác') + '] ' + o.x.tieuDe +
          (o.x.vanBan ? '\n   Căn cứ: ' + o.x.vanBan : '') + (o.x.hieuLuc ? '\n   Hiệu lực từ: ' + o.x.hieuLuc : '') +
          '\n   Nội dung: ' + botS(o.x.noiDung, 3000) + (o.x.lienKet ? '\n   Liên kết: ' + o.x.lienKet : '');
      }).join('\n\n');

      var txtVB = '', dsVBDung = [];
      if (vb.length) {
        var vbXep = vb.map(function (x) { return { x: x, d: botDiem(tuHoi, botVBNhan(x) + ' ' + x.nhom) * 2 }; });
        var con = 52000 - txtQD.length; if (con < 4000) con = 4000;
        vbXep.sort(function (a, b) { return b.d - a.d; });
        for (var i = 0; i < vbXep.length && dsVBDung.length < 3 && con > 1500; i++) {
          var v = vbXep[i].x, t = '';
          try { t = botNapChu(v, false); } catch (e) { continue; }
          if (!t) continue;
          /* văn bản dài: chỉ lấy các đoạn liên quan đến câu hỏi */
          if (t.length > con) {
            var doan = t.split(/\n\s*\n/), xp = doan.map(function (p, k) { return { p: p, k: k, d: botDiem(tuHoi, p) }; })
              .sort(function (a, b) { return b.d - a.d || a.k - b.k; }).filter(function (o) { return o.d > 0; }).slice(0, 40)
              .sort(function (a, b) { return a.k - b.k; });
            t = (xp.length ? xp.map(function (o) { return o.p; }).join('\n…\n') : t).slice(0, con);
          }
          dsVBDung.push(botVBNhan(v));
          txtVB += (txtVB ? '\n\n' : '') + '--- VĂN BẢN: ' + botVBNhan(v) + (v.ngayBanHanh ? ' (ban hành ' + v.ngayBanHanh + ')' : '') + ' ---\n' + t;
          con -= t.length;
        }
      }

      /* chưa cấu hình AI: tra cứu trực tiếp, trả nguyên văn mục quy định khớp nhất (an toàn, không sinh nội dung mới) */
      if (!env.ai || !env.ai.trangThai().coKhoa) {
        /* Chưa bật AI thì chỉ tra cứu máy móc, nên phải chặt: so câu hỏi với TIÊU ĐỀ mục quy định
           (TCHC viết tiêu đề đúng như câu CBCNV hay hỏi), yêu cầu phủ ≥ 60% và ít nhất 2 âm tiết khớp.
           So với toàn bộ nội dung sẽ lọt những khớp vô nghĩa ("quy", "lao", "động", "công ty")
           và trợ lý trả lời sai chủ đề — nguy hiểm hơn là không trả lời. */
        var khop = dungQD.map(function (o) {
          return { x: o.x, k: botKhop(tuHoi, o.x.tieuDe + ' ' + o.x.nhom) };
        }).filter(function (o) {
          return o.k.phu >= BOT_PHU_MIN && o.k.d >= BOT_KHOP_MIN * 2;
        }).sort(function (a, b) { return b.k.d - a.k.d; }).slice(0, 3);
        if (!khop.length) return { daTraLoi: false, dungAI: false, nguon: [], nhom: '', thieu: hoi,
          traLoi: 'Trợ lý chưa tìm thấy quy định nào phù hợp với câu hỏi này (máy chủ chưa bật AI nên chỉ tra cứu theo từ khóa). Đề nghị Anh/Chị liên hệ phòng TCHC.' };
        return { daTraLoi: true, dungAI: false, nhom: khop[0].x.nhom,
          nguon: khop.map(function (o) { return o.x.tieuDe; }),
          traLoi: khop.map(function (o) { return '• ' + o.x.tieuDe + '\n' + o.x.noiDung + (o.x.vanBan ? '\n(Căn cứ: ' + o.x.vanBan + ')' : ''); }).join('\n\n') +
            '\n\n(Máy chủ chưa bật AI – đây là nội dung quy định được trích nguyên văn theo từ khóa.)' };
      }

      var nd = 'CÂU HỎI CỦA CBCNV: ' + hoi +
        '\n\n=== NGUỒN: MỤC QUY ĐỊNH DO PHÒNG TCHC NHẬP ===\n' + (txtQD || '(không có)') +
        (txtVB ? '\n\n=== NGUỒN: VĂN BẢN GỐC CỦA CÔNG TY ===\n' + txtVB : '') +
        '\n\n=== YÊU CẦU ===\n' + BOT_P;
      var o = aiGoi({ he: BOT_HE, nd: nd, anh: [], maxTok: 2000, loai: 'bot',
        duLieu: { hoi: hoi, soQD: dungQD.length, vb: dsVBDung } });

      /* ---- Chốt an toàn chống bịa (bắt buộc: CBCNV sẽ làm theo câu trả lời) ----
         1) Căn cứ AI nêu phải có thật trong NGUỒN đã cấp – loại mọi tên văn bản AI tự đặt.
         2) Không còn căn cứ nào ⇒ coi như chưa có quy định.
         3) Khi chưa có quy định thì BỎ HẲN phần chữ AI sinh ra và dùng câu trả lời chuẩn:
            chữ không có căn cứ thì không kiểm chứng được, rất dễ lọt con số bịa ra ngoài. */
      var hopLe = dungQD.map(function (p) { return p.x.tieuDe; }).concat(dsVBDung);
      var hlK = hopLe.map(function (x) { return botKhongDau(x); });
      var nguon = aiDs(o && o.nguon, 6).map(function (x) { return botS(x, 200); }).filter(Boolean)
        .filter(function (x) {
          var k = botKhongDau(x);
          return hlK.some(function (h) { return h === k || (k.length >= 6 && (h.indexOf(k) >= 0 || k.indexOf(h) >= 0)); });
        });
      var tl = botS(o && o.traLoi, 4000), dtl = !!(o && o.daTraLoi) && !!nguon.length && !!tl;
      var thieu = botS(o && o.thieu, 300) || hoi;
      if (!dtl) tl = 'Công ty chưa có quy định về nội dung này trong hệ thống của trợ lý. Đề nghị Anh/Chị liên hệ phòng TCHC để được trả lời chính thức.';
      return { daTraLoi: dtl, traLoi: tl, nguon: dtl ? nguon : [], dungAI: true,
        nhom: dtl ? aiChon(o && o.nhom, BOT_NHOM, '') : '', thieu: dtl ? '' : thieu };
    };

    /* ghi nhật ký câu hỏi (trình duyệt gọi ngay sau bot.hoi) – để TCHC biết CBCNV đang vướng gì, thiếu quy định gì */
    H['bot.ghi'] = function (req, s) {
      var a = T('BotHoi');
      a.push({ id: newId(), thoiGian: env.now(), username: s.username, hoTen: chatTen(s.username),
        cauHoi: botS(req.hoi, 500), traLoi: botS(req.traLoi, 1000), nguon: botS((req.nguon || []).join('; '), 400),
        nhom: botS(req.nhom, 60), daTraLoi: req.daTraLoi ? 'Có' : '', dungAI: req.dungAI ? 'Có' : '',
        thieu: botS(req.thieu, 300), huuIch: '', phanHoi: '' });
      if (a.length > 3000) a.splice(0, a.length - 2500);
      D('BotHoi');
      return { ok: true, id: a[a.length - 1].id };
    };
    H['bot.danhgia'] = function (req, s) {
      var x = T('BotHoi').filter(function (z) { return z.id === req.id; })[0]; if (!x) err('Không tìm thấy câu hỏi.');
      if (x.username !== s.username) err('Chỉ người đặt câu hỏi được đánh giá câu trả lời.');
      x.huuIch = req.huuIch === 'co' ? 'Có' : req.huuIch === 'khong' ? 'Không' : '';
      x.phanHoi = botS(req.phanHoi, 500); D('BotHoi');
      return { ok: true };
    };

    H['bot.luu'] = function (req, s) {
      botQuyen(s);
      var tieuDe = botS(req.tieuDe, 200), noiDung = botS(req.noiDung, 6000);
      if (tieuDe.length < 5) err('Tiêu đề quy định quá ngắn.');
      if (noiDung.length < 10) err('Nội dung quy định quá ngắn.');
      var a = T('BotQuyDinh'), x = req.id ? a.filter(function (z) { return z.id === req.id; })[0] : null;
      if (req.id && !x) err('Không tìm thấy mục quy định.');
      if (!x) { x = { id: newId() }; a.push(x); }
      x.nhom = aiChon(req.nhom, BOT_NHOM, 'Khác'); x.tieuDe = tieuDe; x.noiDung = noiDung;
      x.lienKet = botS(req.lienKet, 500); x.vanBan = botS(req.vanBan, 300);
      x.hieuLuc = botS(req.hieuLuc, 10); x.trangThai = req.trangThai === 'het' ? 'het' : 'hieuluc';
      x.nguoiSua = s.username; x.capNhat = env.now();
      D('BotQuyDinh'); log(s, 'bot.luu', tieuDe);
      return botQDPub(x);
    };
    H['bot.xoa'] = function (req, s) {
      botQuyen(s);
      var a = T('BotQuyDinh'), x = a.filter(function (z) { return z.id === req.id; })[0]; if (!x) err('Không tìm thấy mục quy định.');
      a.splice(a.indexOf(x), 1); D('BotQuyDinh'); log(s, 'bot.xoa', x.tieuDe);
      return { ok: true };
    };

    /* đăng ký văn bản gốc trên Drive cho bot đọc: nhận fileId (hoặc đường dẫn Drive), trích chữ ngay để báo lỗi sớm */
    H['bot.vb.luu'] = function (req, s) {
      botQuyen(s);
      var raw = botS(req.fileId, 200), m = raw.match(/[-\w]{25,}/);
      var fid = m ? m[0] : raw;
      if (!fid) err('Thiếu mã tệp (fileId) hoặc đường dẫn tệp trên Drive.');
      var a = T('BotVanBan'), x = req.id ? a.filter(function (z) { return z.id === req.id; })[0] : null;
      if (req.id && !x) err('Không tìm thấy văn bản.');
      if (!x && a.filter(function (z) { return z.fileId === fid; }).length) err('Văn bản này đã được đăng ký.');
      var f;
      try { f = DriveApp.getFileById(fid); } catch (e) { err('Không mở được tệp trên Drive. Kiểm tra mã tệp và quyền truy cập của tài khoản chạy app.'); }
      if (!x) { x = { id: newId() }; a.push(x); }
      x.fileId = fid; x.mime = f.getMimeType();
      x.ten = botS(req.ten, 200) || f.getName();
      x.nhom = aiChon(req.nhom, BOT_NHOM, 'Khác'); x.soHieu = botS(req.soHieu, 100);
      x.ngayBanHanh = botS(req.ngayBanHanh, 10); x.trangThai = req.trangThai === 'het' ? 'het' : 'hieuluc';
      x.nguoiSua = s.username; x.capNhat = env.now();
      try { botNapChu(x, true); }
      catch (e) { x.loi = botS(e && e.message || e, 300); x.soChu = 0; x.textLuc = ''; D('BotVanBan'); err(x.loi); }
      D('BotVanBan'); log(s, 'bot.vb.luu', x.ten);
      return botVBPub(x);
    };
    H['bot.vb.xoa'] = function (req, s) {
      botQuyen(s);
      var a = T('BotVanBan'), x = a.filter(function (z) { return z.id === req.id; })[0]; if (!x) err('Không tìm thấy văn bản.');
      a.splice(a.indexOf(x), 1); D('BotVanBan'); log(s, 'bot.vb.xoa', x.ten);
      return { ok: true };
    };

    /* nhật ký + báo cáo lỗ hổng: câu nào bot không trả lời được, gom theo nội dung còn thiếu */
    H['bot.nhatky'] = function (req, s) {
      botQuyen(s);
      var a = T('BotHoi').slice().sort(function (x, y) { return x.thoiGian < y.thoiGian ? 1 : -1; });
      var ds = a.slice(0, Math.min(num(req.limit) || 300, 1000)).map(function (x) {
        return { id: x.id, thoiGian: x.thoiGian, hoTen: x.hoTen, username: x.username, cauHoi: x.cauHoi,
          traLoi: botS(x.traLoi, 400), nguon: x.nguon, nhom: x.nhom, daTraLoi: x.daTraLoi === 'Có',
          dungAI: x.dungAI === 'Có', huuIch: x.huuIch, phanHoi: x.phanHoi };
      });
      /* Gom nội dung còn thiếu theo độ giống nhau của tập âm tiết (Jaccard ≥ 0,6), không gom theo
         chuỗi đầu: "có hỗ trợ tiền xăng xe không" và "có hỗ trợ tiền xăng xe đi công trường không"
         là một nhu cầu, tách ra hai dòng sẽ làm TCHC tưởng là hai việc khác nhau. */
      var nhomT = [], nhom = {}, nguoi = {};
      function gomThieu(txt) {
        var t = botTu(txt), s = {}, i, j;
        for (i = 0; i < t.length; i++) s[t[i]] = 1;
        var ks = Object.keys(s); if (!ks.length) return;
        for (i = 0; i < nhomT.length; i++) {
          var g = nhomT[i], chung = 0;
          for (j = 0; j < ks.length; j++) if (g.tu[ks[j]]) chung++;
          var hop = g.n + ks.length - chung;
          if (hop && chung / hop >= 0.6) {
            g.lan++;
            if (ks.length > g.n) { g.mau = txt; g.n = ks.length; for (j = 0; j < ks.length; j++) g.tu[ks[j]] = 1; }
            return;
          }
        }
        if (nhomT.length < 200) nhomT.push({ mau: txt, lan: 1, n: ks.length, tu: s });
      }
      a.forEach(function (x) {
        nhom[x.nhom || 'Khác'] = (nhom[x.nhom || 'Khác'] || 0) + 1;
        nguoi[x.username] = (nguoi[x.username] || 0) + 1;
        if (x.daTraLoi !== 'Có') gomThieu(x.thieu || x.cauHoi);
      });
      return { ds: ds, tong: a.length,
        soTraLoiDuoc: a.filter(function (x) { return x.daTraLoi === 'Có'; }).length,
        soHuuIch: a.filter(function (x) { return x.huuIch === 'Có'; }).length,
        soKhongHuuIch: a.filter(function (x) { return x.huuIch === 'Không'; }).length,
        thieu: nhomT.map(function (g) { return { mau: g.mau, lan: g.lan }; }).sort(function (x, y) { return y.lan - x.lan; }).slice(0, 30),
        theoNhom: Object.keys(nhom).map(function (k) { return { nhom: k, lan: nhom[k] }; }).sort(function (x, y) { return y.lan - x.lan; }),
        soNguoi: Object.keys(nguoi).length };
    };

    var READ_ONLY = ['tl.ocrxem', 'chat.anh', 'chat.khods', 'chat.khoxem', 'login', 'load', 'dl.cu', 'tl.get', 'tl.thumuc', 'nv.cccdai', 'cv.ds', 'cv.get', 'cv.ai', 'cv.tn', 'dx.ds', 'dx.tepget',
      'ping', 'log.list', 'cv.ocr', 'chat.init', 'chat.poll', 'chat.history', 'chat.file', 'chat.files',
      'ai.nhom', 'ai.thu', 'ai.models', 'ai.phantich', 'ai.taifile', 'gps.cuatoi', 'gps.ds', 'gps.anh', 'gps.tonghop', 'gps.ct.ds', 'ai.ds', 'ai.xem', 'ai.tep', 'ai.tepluu', 'ai.doctep', 'chat.nhatky', 'push.cuatoi', 'push.trangthai', 'cckt.ds', 'cckt.ai', 'vitri.ds', 'dgt.ds', 'dgt.xem', 'dgt.ai', 'dgt.cuatoi', 'bot.ds', 'bot.hoi', 'bot.nhatky'];   // phân tích AI chạy lâu: không giữ khóa ghi (chỉ ghi Drive, bộ nhớ đệm)

    function handle(req) {
      try {
        req = req || {};
        var fn = H[req.action];
        if (!fn) err('Thao tác không hợp lệ: ' + req.action);
        getCfg();                                   // v3.8: áp bảng phân quyền đã lưu trước mọi kiểm tra quyền
        var s = (req.action === 'login' || req.action === 'ping' || req.action === 'logout') ? null : sess(req);
        PUSH.length = 0;
        var data = fn(req, s), out = { ok: true, data: data };
        if (PUSH.length) { var dsP = pushDs(); if (dsP.length) out._push = dsP; }
        return out;
      } catch (e) {
        PUSH.length = 0;
        var m = String(e && e.message || e);
        return { ok: false, error: m.replace(/^AUTH: /, ''), auth: /^AUTH: /.test(m) };
      }
    }

    return { handle: handle, isReadOnly: function (a) { return READ_ONLY.indexOf(a) >= 0; }, getCfg: getCfg, upsertUser: upsertUser, chayLich: chayLich, luuTruTuDong: luuTruTuDong, layPush: function () { var d = PUSH.length ? pushDs() : []; PUSH.length = 0; return d; }, xoaTokenPush: xoaTokenPush };
  }

  return {
    VERSION: VERSION, ROLES: ROLES, khoangCachM: khoangCachM, kmKhoangCach: kmKhoangCach, aiKhoi: aiKhoi, nkKhoi: nkKhoi, nkO: nkO, AI_TRUONG_NK: AI_TRUONG_NK, AI_KHIA: AI_KHIA, xlAiBaoCao: xlAiBaoCao, xlNhatKy: xlNhatKy, aiJSON: aiJSON, aiSo: aiSo, aiSoTxt: aiSoTxt, AI_HE: AI_HE, ngayVN: ngayVN, THU_VN: THU_VN, otVaCT: otVaCT, anCaVaCT: anCaVaCT, TL_NHOM: TL_NHOM, tlNhom: tlNhom, gioLamThuc: gioLamThuc, nguongDuCong: nguongDuCong, bcBangDoc: HRCORE_bcBangDoc, PERM: PERM, PERM_MD: PERM_MD, PQ_NHOM: PQ_NHOM, PQ_SUA: PQ_SUA, PQ_KHOA: PQ_KHOA, PQ_NGUOI_SUA: PQ_NGUOI_SUA, pqKhoa: pqKhoa, PQ_CAM: PQ_CAM, pqCam: pqCam, apDungPQ: apDungPQ, pqLamSach: pqLamSach, can: can, phanTichCT: phanTichCT, tinhCT: tinhCT, boDau: boDau, diemTim: diemTim, locGan: locGan, MAU_PHIEU_MD: MAU_PHIEU_MD, MAU_BL_MD: MAU_BL_MD, mauPhieu: mauPhieu, mauBL: mauBL, mauGon: mauGon, ctMienAnCa: ctMienAnCa, CT_ANCA_MD: CT_ANCA_MD, kiemCTAnCa: kiemCTAnCa, mienAnCa: mienAnCa, bienDung: bienDung, kiemTraCT: kiemTraCT,
    BIEN_HT: BIEN_HT, matKhauTheoNV: matKhauTheoNV, tenBien: tenBien, xlMauTamUng: xlMauTamUng, docTamUng: docTamUng, docTien: docTien, docNgay: docNgay, nhanLoaiTU: nhanLoaiTU, dauVaoTinhLuong: dauVaoTinhLuong, astExcel: astExcel, ctExcel: ctExcel, cotXL: cotXL, xlBangTinh: xlBangTinh, xlBangLuongThang: xlBangLuongThang, xlNam: xlNam, xlCaNhan: xlCaNhan, xlChiPhi: xlChiPhi, sangXLSX: sangXLSX, khoanCo: khoanCo, amSangDuong: amSangDuong, lichLeTheoLuat: lichLeTheoLuat, cleanLe: cleanLe, LOAI_LE: LOAI_LE, KY_HIEU: KY_HIEU, DON_KH: DON_KH, ngayTheoDon: ngayTheoDon, cleanNgay: cleanNgay, tongHopNgay: tongHopNgay, NHOM_VP: NHOM_VP, thu: thu, ROLE_QL: ROLE_QL, normRole: normRole, tenDangNhap: tenDangNhap, ROLES_MOTA: ROLES_MOTA, BT_LOAI: BT_LOAI, CAM_XUC: CAM_XUC, memChat: memChat, CT_PHANLOAI: CT_PHANLOAI, CT_VAITRO: CT_VAITRO, CT_MAU: CT_MAU, nlvTrongKy: nlvTrongKy, bienNLV: bienNLV, CT_LUONG_MD: CT_LUONG_MD, BIEN_KY_MD: BIEN_KY_MD, CT_DONGBH_MD: CT_DONGBH_MD, BUOC_XONG: BUOC_XONG, LOAI_BUOC: LOAI_BUOC, FIELDS_LOAI: FIELDS_LOAI,
    buocLuong: buocLuong, viTriBuoc: viTriBuoc, buocTheoLoai: buocTheoLoai, laNguoiXuLy: laNguoiXuLy, kiemTraQTLuong: kiemTraQTLuong,
    CV_HUONG: CV_HUONG, CV_LOAI: CV_LOAI, CV_KHAN: CV_KHAN, CV_TT_DEN: CV_TT_DEN, CV_TT_DI: CV_TT_DI, CV_HET: CV_HET,
    CV_CS: CV_CS, CV_CS_OK: CV_CS_OK,
    DX_TT: DX_TT, DX_MO: DX_MO, DX_XONG: DX_XONG, QT_DX_MD: QT_DX_MD, NHOM_DX_MD: NHOM_DX_MD, DX_CACH: DX_CACH,
    kiemTraLoaiDX: kiemTraLoaiDX, kiemTraNhomDX: kiemTraNhomDX, dxBuocDay: dxBuocDay,
    QT_LUONG_MD: QT_LUONG_MD, QT_NGHI_MD: QT_NGHI_MD, QT_YKIEN_MD: QT_YKIEN_MD, LOAI_VP: LOAI_VP, YK_TT: YK_TT, LOAI_NGHI: LOAI_NGHI, NGHI_COT: NGHI_COT, NP_TT: NP_TT,
    laNgayNghiLe: laNgayNghiLe, soNgayDon: soNgayDon, ngayNghiTheoKy: ngayNghiTheoKy, quyPhepNam: quyPhepNam, phepNamDaDung: phepNamDaDung, normBuoc: normBuoc, SCHEMA: SCHEMA, isNumCol: isNumCol,
    DEFAULT_CFG: DEFAULT_CFG, LOAI_HD: LOAI_HD, LOAI_LUONG: LOAI_LUONG, LY_DO_KHONG_BH: LY_DO_KHONG_BH, loaiLuongNV: loaiLuongNV, TRANG_THAI_NV: TRANG_THAI_NV, LOAI_QT: LOAI_QT,
    LOAI_LL: LOAI_LL, LOAI_KT: LOAI_KT, HT_KYLUAT: HT_KYLUAT,
    num: num, addMonths: addMonths, addDays: addDays, kyEnd: kyEnd, validKy: validKy, congChuanTruCN: congChuanTruCN,
    mergeCfg: mergeCfg, parseJSON: parseJSON, thueLuyTien: thueLuyTien, tinhPhieuLuong: tinhPhieuLuong,
    luongTaiKy: luongTaiKy, mucHuong: mucHuong, nvTaiKy: nvTaiKy, mhNguon: mhNguon, MH_KHOA: MH_KHOA, MH_LUONG: MH_LUONG, hanNangBac: hanNangBac, dangLamTrongKy: dangLamTrongKy, ngoaiHD: ngoaiHD, QT_CHAM_DUT: QT_CHAM_DUT, laVangLai: laVangLai,
    congTyNV: congTyNV, congTyBL: congTyBL, tenCTV: tenCTV, nhanCTV: nhanCTV, dsCTV: dsCTV, tkCongTy: tkCongTy, chuanCTV: chuanCTV,
    cfgTheoKy: cfgTheoKy, mocCuaKy: mocCuaKy, MOC_LUAT_MD: MOC_LUAT_MD, BIEU_THUE: BIEU_THUE, OT_CACH: OT_CACH, loaiOT: loaiOT, cleanOT: cleanOT, tongHopOT: tongHopOT, apGioOT: apGioOT, cleanAn: cleanAn, soBuaCC: soBuaCC, ocrTach: ocrTach, OCR_TRUONG: OCR_TRUONG, OCR_TL: OCR_TL, b64Utf8: b64Utf8, HT_HD: HT_HD, HD_KHONG_BH: HD_KHONG_BH, LY_DO_HD_KBH: LY_DO_HD_KBH, hdDs: hdDs, hdTatCa: hdTatCa, hdTaiNgay: hdTaiNgay, hdDoanKy: hdDoanKy, apHD: apHD, hdMoTa: hdMoTa, hdCoBH: hdCoBH, hdDuoi3Thang: hdDuoi3Thang, hdCanhBao: hdCanhBao, tinhNhieuDoan: tinhNhieuDoan, MOC_BH: MOC_BH, HD_VANGLAI: HD_VANGLAI, mienAnCa: mienAnCa, mucMienAnCa: mucMienAnCa, ANCA_BUA: ANCA_BUA,
    TRUONG_HOP: TRUONG_HOP, TH_KHAC: TH_KHAC, truongHopNV: truongHopNV, ttThue: ttThue, chuanTT: chuanTT, TT_CACH: TT_CACH, quyetToanNam: quyetToanNam,
    HUU_TT: HUU_TT, HUU_XM: HUU_XM, HUU_PA: HUU_PA, HS_HUU: HS_HUU, huuTaiKy: huuTaiKy, traThemBH: traThemBH, chiTietLuyTien: chiTietLuyTien, calculateRetireePIT: calculateRetireePIT, evaluateRetireeTaxAuthorization: evaluateRetireeTaxAuthorization, phanTichHuuTri: phanTichHuuTri, ghiChuHuuTri: ghiChuHuuTri, huuChuaXacMinh: huuChuaXacMinh, tinhPhieuLuongGoc: tinhPhieuLuongGoc,
    QUY_DINH_MD: QUY_DINH_MD, QD_TT: QD_TT, BIEU_MAU_MD: BIEU_MAU_MD, BM_TT: BM_TT, quyDinhDs: quyDinhDs, bieuMauDs: bieuMauDs, quyDinh: quyDinh, bieuMau: bieuMau, DH_QHLD: DH_QHLD, TV_VITRI: TV_VITRI, hsThueNV: hsThueNV, kiemTraThuViec: kiemTraThuViec, dauHieuQHLD: dauHieuQHLD, mstHopLe: mstHopLe,
    LOAI_K: LOAI_K, HT_KHOAN: HT_KHOAN, CH_BAN_CHAT: CH_BAN_CHAT, CH_TO_DOI: CH_TO_DOI, KL_PL: KL_PL, CHE_DO_K: CHE_DO_K, CK_TT: CK_TT, HS_KHOAN: HS_KHOAN, HS_TT: HS_TT, QT_KHOAN_MD: QT_KHOAN_MD,
    classifyContractRelationship: classifyContractRelationship, evaluateEmploymentInsuranceRisk: evaluateEmploymentInsuranceRisk, classifyContractorTaxRegime: classifyContractorTaxRegime, calculateMonthlyContractPayment: calculateMonthlyContractPayment, calculateContractorWithholdingTax: calculateContractorWithholdingTax, khoanNetToGross: khoanNetToGross, tamUngTon: tamUngTon,
    estimateAnnualContractorIncome: estimateAnnualContractorIncome, calculateAnnualPersonalDeductions: calculateAnnualPersonalDeductions, evaluateTaxCommitment: evaluateTaxCommitment, validateTaxCommitmentDocuments: validateTaxCommitmentDocuments, estimateAnnualPersonalIncomeTax: estimateAnnualPersonalIncomeTax, estimatePotentialTaxRefund: estimatePotentialTaxRefund, evaluateTaxSettlementAuthorization: evaluateTaxSettlementAuthorization, quyenLoiThueKhoan: quyenLoiThueKhoan, generateContractorTaxNotes: generateContractorTaxNotes, generateContractorPaymentStatement: generateContractorPaymentStatement,
    uocTinhNam: uocTinhNam, giamTruNam: giamTruNam, DG_KQ: DG_KQ, evaluateTaxCommitmentEligibility: evaluateTaxCommitmentEligibility, evaluateTaxAuthorizationEligibility: evaluateTaxAuthorizationEligibility, toiUuThue: toiUuThue, canhBaoThueBH: canhBaoThueBH, PayrollEngine: PayrollEngine, traThayBHCua: traThayBHCua, soNgayLamViec: soNgayLamViec,
    createApi: createApi
  };
})();
