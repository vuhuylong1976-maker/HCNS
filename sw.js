/* TAPACO Nhân sự – Service Worker: chỉ cache giao diện; dữ liệu luôn lấy từ Apps Script */
var CACHE = 'tapaco-hr-v352';
var SHELL = ['./', './index.html', './manifest.webmanifest', './icon-180.png', './icon-192.png', './icon-512.png'];
self.addEventListener('install', function (e) { self.skipWaiting(); e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(SHELL).catch(function () {}); })); });
self.addEventListener('activate', function (e) { e.waitUntil(caches.keys().then(function (k) { return Promise.all(k.map(function (x) { if (x !== CACHE && x.indexOf('tapaco-chat-anh') !== 0) return caches.delete(x); })); }).then(function () { return self.clients.claim(); })); });
self.addEventListener('fetch', function (e) {
  var req = e.request; if (req.method !== 'GET') return;
  if (/script\.google\.com|googleusercontent\.com|raw\.githubusercontent\.com/.test(req.url)) return;
  e.respondWith(fetch(req).then(function (res) {
    if (res && res.status === 200 && res.type !== 'opaque') { var cp = res.clone(); caches.open(CACHE).then(function (c) { c.put(req, cp); }); }
    return res;
  }).catch(function () { return caches.match(req); }));
});
/* thông báo đẩy (Firebase Cloud Messaging): app đang mở trên màn hình → app tự hiện banner + âm báo;
   app chạy nền / đã tắt → hiện thông báo hệ thống (chuông, rung, nổi trên thanh trạng thái) */
self.addEventListener('push', function (e) {
  var p = {}; try { p = e.data ? e.data.json() : {}; } catch (x) { p = { data: { t: 'TAPACO', b: e.data ? e.data.text() : '' } }; }
  var d = p.data || p, n = p.notification || {};
  var t = d.t || n.title || 'TAPACO', b = d.b || n.body || '', tag = d.tag || 'tapaco', mo = d.mo || '';
  var ios = /iPhone|iPad|iPod/.test(self.navigator.userAgent);
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (cs) {
    var dangXem = cs.filter(function (c) { return c.visibilityState === 'visible' && c.focused; })[0];
    if (dangXem && !ios) { dangXem.postMessage({ type: 'push', t: t, b: b, tag: tag, mo: mo, loai: d.loai || '' }); return; }
    return self.registration.showNotification(t, { body: b, tag: tag, renotify: d.im !== '1', silent: false, icon: 'icon-192.png', badge: 'icon-192.png',
      vibrate: d.im === '1' ? undefined : [220, 90, 220, 90, 220], requireInteraction: d.loai === 'duyet', timestamp: Date.now(), data: { mo: mo } });
  }));
});
/* bấm thông báo: mở app đúng màn hình (cuộc trò chuyện, bài viết, việc cần duyệt…) */
self.addEventListener('notificationclick', function (e) {
  var dt = e.notification.data || {}, mo = dt.mo || (dt.ht !== undefined ? '#chat=' + dt.ht : ''); e.notification.close();
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (cs) {
    for (var i = 0; i < cs.length; i++) { if ('focus' in cs[i]) { cs[i].postMessage({ type: 'mo', mo: mo }); return cs[i].focus(); } }
    return self.clients.openWindow('./index.html' + mo);
  }));
});
self.addEventListener('pushsubscriptionchange', function (e) {
  e.waitUntil(self.clients.matchAll({ type: 'window' }).then(function (cs) { cs.forEach(function (c) { c.postMessage({ type: 'pushDoi' }); }); }));
});
