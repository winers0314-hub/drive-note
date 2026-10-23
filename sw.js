// 설치한 앱이 빨리 열리고, 인터넷이 잠깐 끊겨도 화면은 뜨도록 앱 파일을 보관
// 항상 새 버전을 먼저 받아 오고(고친 내용이 바로 반영), 못 받으면 보관한 것을 씀
const CACHE = 'es-note-v12';
const LIBS = [
  'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
  'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
];
const SHELL = ['./', 'index.html', 'css/app.css', 'config.js', 'manifest.webmanifest', 'icons/icon.svg', 'icons/icon-192.png', 'icons/icon-512.png',
  'js/app.js', 'js/doc.js', 'js/onenote.js', 'js/docx-import.js', 'js/editor.js', 'js/hwpx.js', 'js/hwpx-template.js', 'js/store.js', 'js/xlsx.js', ...LIBS];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => Promise.all(SHELL.map(u => c.add(u).catch(() => {})))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== 'GET') return;
  if (url.origin !== location.origin && !LIBS.includes(req.url)) return; // 구글 로그인·드라이브는 건드리지 않음
  if (url.pathname.startsWith('/api/')) return; // 로그인 서버 응답은 보관하지 않음
  // 앱 파일은 매번 서버에 새 버전이 있는지 확인
  const net = url.origin === location.origin ? fetch(req, { cache: 'no-cache' }) : fetch(req);
  e.respondWith(net.then(res => {
    if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
    return res;
  }).catch(() => caches.match(req, { ignoreSearch: true }).then(r => r || caches.match('index.html'))));
});
