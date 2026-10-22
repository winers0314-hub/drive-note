// 에스 노트 — 화면과 동작
import { DriveStore, LocalStore, newIndex, uid, SECTION_COLORS } from './store.js';
import { setupEditor, exec, formatBlock, fontSize, insertHTML, insertText, toggleChecklist, tableHTML, tableOp, addFile, caretIn, hideImgBar, sizeText } from './editor.js';
import { sanitize, parseHTML, toLines } from './doc.js';
import { toHwpx } from './hwpx.js';
import { toXlsx } from './xlsx.js';
import { msLogin, msSignedIn, listNotebooks, listPages, pageContent, convertPage } from './onenote.js';

const $ = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const CLIENT_ID = (window.NOTE_CONFIG && window.NOTE_CONFIG.googleClientId || '').trim();
const phone = () => matchMedia('(max-width: 599px)').matches;                          // 휴대폰: 한 칸씩
const fold = () => matchMedia('(min-width: 600px) and (max-width: 1023px)').matches;  // 폴드 펼침·태블릿

// ---------- 설정 (이 기기에 저장) ----------
const DEFAULTS = { auto: true, autoNumber: true, autoBullet: true, autoCheck: true, shrink: true, exportToDrive: true, autoExport: 'none', theme: 'system', defaultLines: '', browserSpell: true };
let S = { ...DEFAULTS };
try { S = { ...DEFAULTS, ...JSON.parse(localStorage.getItem('dn-settings') || '{}') }; } catch {}
const saveSettings = () => { try { localStorage.setItem('dn-settings', JSON.stringify(S)); } catch {} applySettings(); };
function applySettings() {
  if (S.theme === 'system') document.documentElement.removeAttribute('data-theme'); else document.documentElement.dataset.theme = S.theme;
  const q = $('autoQuick'); q.classList.toggle('is-on', S.auto); q.querySelector('b').textContent = S.auto ? '켬' : '끔';
  const sp = $('spellLive'); sp.classList.toggle('is-on', S.browserSpell); sp.querySelector('b').textContent = S.browserSpell ? '켬' : '끔';
  $('editor').setAttribute('spellcheck', S.browserSpell ? 'true' : 'false');
  updateLinesUI();
  applyWrap();
  $('convertNote').textContent = store?.kind === 'drive'
    ? (S.exportToDrive ? '바꾼 파일은 내려받기 되고, 드라이브 "에스 노트/내보내기" 폴더에도 저장돼요.' : '바꾼 파일은 이 기기에 내려받기 돼요.') + (S.autoExport !== 'none' ? ` · 자동 변환 켬(${AUTO_LABEL[S.autoExport]})` : '')
    : '바꾼 파일은 이 기기에 내려받기 돼요. 한글 파일은 .hwpx 형식이라 한글 2014 이후 버전에서 열려요.';
}
const AUTO_LABEL = { none: '안 함', hwpx: '한글', xlsx: '엑셀', both: '한글+엑셀' };

// ---------- 상태 ----------
let store = null, ix = null;
let cur = { nb: null, sec: null, page: null };
let dirty = false, editedSinceOpen = false, saveTimer = 0, indexTimer = 0, saving = Promise.resolve(), indexPending = false;
let searchMode = null; // {q, ids}
const imgURL = new Map(); // fid → blob URL

const allSections = () => ix.notebooks.flatMap(nb => nb.sections.map(sec => ({ nb, sec })));
const allPages = () => allSections().flatMap(({ nb, sec }) => sec.pages.map(p => ({ nb, sec, p })));
const findPage = id => allPages().find(x => x.p.id === id);
function remember() { try { localStorage.setItem('dn-last', JSON.stringify({ nb: cur.nb?.id, sec: cur.sec?.id, page: cur.page?.id })); } catch {} }

// ---------- 상태 표시·알림 ----------
function status(kind, text) {
  const m = { ok: '✓ 저장됨', saving: '저장 중…', edit: '편집 중', offline: '⚠ 오프라인 · 연결되면 저장', auth: '⚠ 다시 로그인 필요', err: '⚠ 저장 실패' };
  $('sync').textContent = text || m[kind] || '';
  $('sync').title = store?.kind === 'drive' ? '구글 드라이브에 저장' : '이 기기에만 저장';
}
let toastTimer = 0;
function toast(msg, actions = [], ms = 4000) {
  const t = $('toast');
  t.innerHTML = `<span>${esc(msg)}</span>` + actions.map((a, i) => `<button data-i="${i}">${esc(a.label)}</button>`).join('');
  t.hidden = false;
  t.onclick = e => { const b = e.target.closest('button'); if (!b) return; t.hidden = true; actions[+b.dataset.i].run(); };
  clearTimeout(toastTimer); if (ms) toastTimer = setTimeout(() => { t.hidden = true; }, ms);
}
function modal(html, onReady) {
  $('modal').innerHTML = html; $('modal').hidden = false; $('shade').hidden = false;
  onReady?.($('modal'));
  const f = $('modal').querySelector('[autofocus]'); if (f) f.focus();
}
function closeModal() { $('modal').hidden = true; $('shade').hidden = true; $('modal').innerHTML = ''; }
$('shade').onclick = () => { if (!$('modal').dataset.locked) closeModal(); };
const ask = (title, label, value = '') => new Promise(res => {
  modal(`<h2>${esc(title)}</h2><input type="text" id="askIn" value="${esc(value)}" placeholder="${esc(label)}" autofocus><div class="row"><button class="btn" id="askNo">취소</button><button class="btn primary" id="askOk">확인</button></div>`, m => {
    const ok = () => { const v = m.querySelector('#askIn').value.trim(); closeModal(); res(v || null); };
    m.querySelector('#askOk').onclick = ok; m.querySelector('#askNo').onclick = () => { closeModal(); res(null); };
    m.querySelector('#askIn').onkeydown = e => { if (e.key === 'Enter') ok(); };
    setTimeout(() => m.querySelector('#askIn').select(), 0);
  });
});
const confirmBox = (title, body, yes = '삭제', danger = true) => new Promise(res => {
  modal(`<h2>${esc(title)}</h2><p>${body}</p><div class="row"><button class="btn" id="cNo">취소</button><button class="btn ${danger ? 'danger' : 'primary'}" id="cYes">${esc(yes)}</button></div>`, m => {
    m.querySelector('#cYes').onclick = () => { closeModal(); res(true); };
    m.querySelector('#cNo').onclick = () => { closeModal(); res(false); };
  });
});

// ---------- 오류 처리 ----------
function handleErr(e, retry) {
  console.error(e);
  if (e.auth) {
    status('auth');
    toast('구글 로그인이 만료됐어요. 다시 로그인하면 이어서 저장해요.', [{ label: '로그인', run: async () => { try { await store.signIn(); status('ok'); retry?.(); } catch (er) { toast(er.message); } } }], 0);
  } else if (e.offline) { status('offline'); setTimeout(() => retry?.(), 15000); }
  else { status('err'); toast(e.message || '저장하지 못했어요.', retry ? [{ label: '다시 시도', run: retry }] : [], 8000); }
}
addEventListener('online', () => { if (dirty) savePageNow(); if (indexPending) saveIndexNow(); });

// ---------- 저장 ----------
function scheduleIndexSave() { indexPending = true; clearTimeout(indexTimer); indexTimer = setTimeout(saveIndexNow, 700); }
function saveIndexNow() {
  clearTimeout(indexTimer);
  if (!indexPending) return saving;
  indexPending = false; ix.updated = Date.now();
  const snapshot = JSON.parse(JSON.stringify(ix));
  saving = saving.then(async () => {
    try { status('saving'); await store.saveIndex(snapshot); if (!dirty) status('ok'); }
    catch (e) { indexPending = true; handleErr(e, saveIndexNow); }
  });
  return saving;
}
function serialize() {
  const c = $('editor').cloneNode(true);
  c.querySelectorAll('img[data-fid]').forEach(i => i.removeAttribute('src'));
  c.querySelectorAll('.sel').forEach(i => i.classList.remove('sel'));
  return c.innerHTML;
}
function plain(html) { return toLines(parseHTML(html)).lines.join(' ').replace(/\s+/g, ' ').trim(); }
function markDirty() {
  if (!cur.page) return;
  dirty = true; editedSinceOpen = true; status('edit');
  clearTimeout(saveTimer); saveTimer = setTimeout(savePageNow, 1200);
  saveDraftSoon();
}
let draftTimer = 0;
function saveDraftSoon() {
  clearTimeout(draftTimer);
  const p = cur.page;
  draftTimer = setTimeout(() => { try { localStorage.setItem('dn-draft-' + p.id, JSON.stringify({ t: Date.now(), title: $('pageTitle').value, html: serialize() })); } catch {} }, 400);
}
function savePageNow() {
  clearTimeout(saveTimer);
  if (!dirty || !cur.page) return saving;
  const p = cur.page, html = serialize(), at = Date.now();
  dirty = false;
  p.title = $('pageTitle').value.trim(); p.updated = at; p.snip = plain(html).slice(0, 80);
  renderPages();
  saving = saving.then(async () => {
    try {
      status('saving');
      await store.savePage(p, html);
      try { const d = JSON.parse(localStorage.getItem('dn-draft-' + p.id) || 'null'); if (d && d.t <= at + 500) localStorage.removeItem('dn-draft-' + p.id); } catch {}
      scheduleIndexSave();
      if (!dirty) status('ok');
    } catch (e) { if (cur.page === p) dirty = true; handleErr(e, () => { if (cur.page === p) { dirty = true; savePageNow(); } }); }
  });
  return saving;
}
async function flush() { savePageNow(); saveIndexNow(); await saving; }
addEventListener('beforeunload', e => { if (dirty || indexPending) { savePageNow(); saveIndexNow(); e.preventDefault(); e.returnValue = ''; } });
addEventListener('keydown', e => { if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); flush(); } });

// ---------- 다른 기기에서 바뀐 것 반영 ----------
let lastCheck = 0;
async function checkRemote() {
  if (!store || store.kind !== 'drive' || !store.signedIn || Date.now() - lastCheck < 5000 || indexPending) return;
  lastCheck = Date.now();
  try {
    if (!(await store.remoteChanged())) return;
    const fresh = await store.loadIndex();
    if (!fresh || indexPending) return;
    const before = cur.page ? { id: cur.page.id, updated: cur.page.updated } : null;
    ix = fresh; reselect(cur.nb?.id, cur.sec?.id, cur.page?.id);
    renderAll();
    if (before && cur.page && cur.page.id === before.id && cur.page.updated > before.updated && !dirty) await openPage(cur.page, { keepScroll: true, silent: true });
    toast('다른 기기에서 바뀐 내용을 불러왔어요.');
  } catch (e) { if (!e.offline) console.warn(e); }
}
addEventListener('focus', checkRemote);
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') checkRemote(); else flush(); });

// ---------- 그림·파일 ----------
async function fileURL(fid) {
  if (imgURL.has(fid)) return imgURL.get(fid);
  const b = await store.getFile(fid); if (!b) return '';
  const u = URL.createObjectURL(b); imgURL.set(fid, u); return u;
}
async function shrink(file) {
  if (!S.shrink || file.type === 'image/gif') return file;
  try {
    const bmp = await createImageBitmap(file), max = 1600;
    if (bmp.width <= max && bmp.height <= max && file.size < 1.5e6) return file;
    const k = Math.min(1, max / Math.max(bmp.width, bmp.height));
    const c = document.createElement('canvas'); c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
    const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.drawImage(bmp, 0, 0, c.width, c.height);
    const b = await new Promise(r => c.toBlob(r, 'image/jpeg', 0.85));
    return b && b.size < file.size ? new File([b], file.name.replace(/\.\w+$/, '') + '.jpg', { type: 'image/jpeg' }) : file;
  } catch { return file; }
}
async function upload(file, isImage) {
  try {
    const f = isImage ? await shrink(file) : file;
    toast(`${isImage ? '그림' : '파일'} 올리는 중… (${sizeText(f.size)})`, [], 0);
    const fid = await store.putFile(f, f.name || (isImage ? 'image.jpg' : 'file'));
    const url = URL.createObjectURL(f); imgURL.set(fid, url);
    $('toast').hidden = true;
    return { fid, url };
  } catch (e) { $('toast').hidden = true; handleErr(e); return null; }
}
async function openFile(fid, name) {
  try {
    toast('파일 여는 중…', [], 0);
    const b = await store.getFile(fid); $('toast').hidden = true;
    if (!b) return toast('파일을 찾지 못했어요.');
    download(b, name || 'file');
  } catch (e) { $('toast').hidden = true; handleErr(e); }
}
function download(blob, name) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 30000);
}

// ---------- 그리기 ----------
function reselect(nbId, secId, pageId) {
  cur.nb = ix.notebooks.find(n => n.id === nbId) || ix.notebooks[0] || null;
  cur.sec = cur.nb ? (cur.nb.sections.find(s => s.id === secId) || cur.nb.sections[0] || null) : null;
  cur.page = cur.sec ? (cur.sec.pages.find(p => p.id === pageId) || null) : null;
}
function renderAll() { renderNotebook(); renderSections(); renderPages(); }
const NB_COLORS = ['#4dab9a', '#529cca', '#e255a1', '#ffa344', '#9a6dd7', '#ff7369', '#937264', '#e9c03a'];
// 펼쳐 둔 전자 필기장 (컴퓨터 화면의 트리)
let openNbs = new Set(), curCollapsed = false; // 지금 전자 필기장은 늘 펼치되, 화살표로 접으면 접힌 채로
try { openNbs = new Set(JSON.parse(localStorage.getItem('dn-open-nb') || '[]')); } catch {}
const saveOpenNbs = () => { try { localStorage.setItem('dn-open-nb', JSON.stringify([...openNbs])); } catch {} };
// 전자 필기장 칸: 컴퓨터에서는 전자 필기장 아래에 섹션이 펼쳐지는 트리, 폴드에서는 전자 필기장만 (섹션은 옆 칸)
function renderNotebook() {
  $('nbName').textContent = cur.nb ? cur.nb.name : '전자 필기장 없음';
  $('secHead').textContent = cur.nb ? cur.nb.name : '';
  if (cur.nb && !curCollapsed) openNbs.add(cur.nb.id);
  $('nbList').innerHTML = ix.notebooks.map((n, i) => {
    const open = openNbs.has(n.id);
    return `<li data-id="${n.id}" class="nb-item${n === cur.nb ? ' on' : ''}${open ? ' open' : ''}"><button class="chev" data-chev aria-label="${open ? '접기' : '펼치기'}">›</button><span class="nbi" style="background:${NB_COLORS[i % NB_COLORS.length]}"></span><span class="name">${esc(n.name)}</span><button class="more" data-more="nb" aria-label="전자 필기장 메뉴">⋯</button></li>` +
      (open ? n.sections.map(s => `<li data-sec="${s.id}" data-nb="${n.id}" class="tree-sec${s === cur.sec ? ' on' : ''}"><span class="bar" style="background:${s.color}"></span><span class="name">${esc(s.name)}</span><button class="more" data-more="sec" aria-label="섹션 메뉴">⋯</button></li>`).join('') +
        `<li class="tree-sec tree-add" data-addsec="${n.id}"><span class="bar"></span><span class="name">＋ 섹션 추가</span></li>` : '');
  }).join('');
}
function renderSections() {
  const ul = $('sectionList');
  ul.innerHTML = (cur.nb?.sections || []).map(s => `<li data-id="${s.id}" class="${s === cur.sec ? 'on' : ''}" draggable="true" style="border-left-color:${s.color}"><span class="dot" style="background:${s.color}"></span><span class="name">${esc(s.name)}</span><button class="more" data-more="sec" aria-label="섹션 메뉴">⋯</button></li>`).join('');
  renderNotebook();
}
const fmtDate = t => t ? new Date(t).toLocaleDateString('ko-KR', { year: 'numeric', month: 'long', day: 'numeric' }) : '';
const fmtFull = t => t ? new Date(t).toLocaleString('ko-KR', { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long', hour: 'numeric', minute: '2-digit' }) : '';
function renderPages() {
  const ul = $('pageList'), head = $('pagesHead');
  if (searchMode) {
    const hits = searchMode.ids.map(findPage).filter(Boolean);
    head.innerHTML = `<span>검색 결과 ${hits.length}개 · “${esc(searchMode.q)}”</span><button id="endSearch">닫기</button>`;
    $('endSearch').onclick = endSearch;
    ul.innerHTML = hits.map(({ sec, p }) => `<li data-id="${p.id}" class="${p === cur.page ? 'on' : ''}"><div class="row"><span class="name">${esc(p.title || '제목 없음')}</span></div><span class="snip">${esc(sec.name)} · ${esc(p.snip || '')}</span></li>`).join('') || '<li class="none"><span class="snip">찾는 내용이 없어요.</span></li>';
    $('addPage').hidden = true;
    return;
  }
  $('addPage').hidden = !cur.sec;
  head.innerHTML = cur.sec ? `<span style="color:${cur.sec.color}">■</span><span>${esc(cur.sec.name)}</span>` : '';
  ul.innerHTML = (cur.sec?.pages || []).map(p => `<li data-id="${p.id}" class="${p === cur.page ? 'on' : ''}${p.level ? ' sub' : ''}" draggable="true"><div class="row"><span class="name">${esc(p.title || '제목 없음')}</span><button class="more" data-more="page" aria-label="페이지 메뉴">⋯</button></div><span class="date">${fmtDate(p.updated || p.created)}</span></li>`).join('');
}
function showEmpty(html) { $('page').hidden = true; $('empty').hidden = false; $('empty').innerHTML = html; }
function setView(v) {
  $('shell').dataset.view = v;
  $('navBack').hidden = !phone() || v === 'sections';
}
$('navBack').onclick = () => setView($('shell').dataset.view === 'page' ? 'pages' : 'sections');

// 목록 접기·펴기 (컴퓨터: ☰ / 전체 페이지 보기, 폴드: 글을 누르면 접히고 ☰로 다시 펼침)
let navOff = false;
try { navOff = localStorage.getItem('dn-nav-off') === '1'; } catch {}
function applyNav() {
  const off = navOff && !phone();
  $('shell').classList.toggle('nav-off', off);
  $('navToggle').title = off ? '목록 펴기' : '목록 접기 (전체 페이지 보기)';
  $('navToggle').classList.toggle('on', !off);
}
function setNav(off) { navOff = off; try { localStorage.setItem('dn-nav-off', off ? '1' : '0'); } catch {} applyNav(); }
$('navToggle').onclick = () => setNav(!navOff);
$('fullView').onclick = () => setNav(true);
// 글만 보기: 메뉴·목록을 모두 숨기고 제목과 글만
function setFocus(on) {
  document.body.classList.toggle('focus', on); $('focusExit').hidden = !on;
  if (on && !cur.page) toast('페이지를 열면 글만 크게 볼 수 있어요.');
}
$('focusBtn').onclick = () => setFocus(true);
$('focusView').onclick = () => setFocus(true);
$('focusExit').onclick = () => setFocus(false);
addEventListener('keydown', e => { if (e.key === 'Escape' && document.body.classList.contains('focus') && $('modal').hidden) setFocus(false); });
// 옆으로 길게 쓰기 (줄 바꿈 안 함)
function applyWrap() {
  $('page').classList.toggle('nowrap', !!S.nowrap);
  const b = $('nowrapBtn'); b.classList.toggle('is-on', !!S.nowrap); b.querySelector('b').textContent = S.nowrap ? '켬' : '끔';
}
$('nowrapBtn').onclick = () => { S.nowrap = !S.nowrap; saveSettings(); toast(S.nowrap ? '줄을 바꾸지 않고 옆으로 이어 써요. 옆으로 스크롤해서 보세요.' : '창 너비에 맞춰 줄을 바꿔요.'); };
$('canvas').addEventListener('click', () => { if (fold() && !navOff && cur.page) setNav(true); });
applyNav();

// ---------- 페이지 열기 ----------
async function openPage(p, { keepScroll = false, silent = false, focusTitle = false } = {}) {
  if (cur.page && cur.page !== p) { leavePage(cur.page); }
  else if (dirty) savePageNow();
  hideImgBar();
  cur.page = p; remember(); renderPages();
  if (!p) { showEmpty(cur.sec ? '<h2>페이지가 없어요</h2><p>왼쪽 아래 <b>＋ 페이지 추가</b>를 눌러 첫 페이지를 만드세요.</p>' : '<h2>섹션이 없어요</h2><p><b>＋ 섹션 추가</b>로 섹션을 먼저 만드세요.</p>'); return; }
  $('empty').hidden = true; $('page').hidden = false;
  const top = $('canvas').scrollTop;
  $('spell').hidden = true; $('ai').hidden = true;
  applyLines(p);
  $('pageTitle').value = p.title || '';
  $('pageDate').textContent = fmtFull(p.created);
  $('editor').innerHTML = '';
  if (!silent) status('', '불러오는 중…');
  let html = '';
  try { html = await store.loadPage(p); }
  catch (e) { handleErr(e, () => openPage(p)); return; }
  if (cur.page !== p) return;
  // 저장되지 않은 임시 내용이 있으면 되살리기 제안
  let draft = null; try { draft = JSON.parse(localStorage.getItem('dn-draft-' + p.id) || 'null'); } catch {}
  $('editor').innerHTML = sanitize(html) || '<p><br></p>';
  dirty = false; editedSinceOpen = false;
  status('ok');
  hydrate();
  if (draft && draft.t > (p.updated || 0) + 1000 && draft.html !== html) {
    toast('저장되지 않은 내용이 있어요.', [{ label: '되살리기', run: () => { $('editor').innerHTML = sanitize(draft.html); if (draft.title) $('pageTitle').value = draft.title; hydrate(); markDirty(); } }, { label: '버리기', run: () => localStorage.removeItem('dn-draft-' + p.id) }], 0);
  }
  if (keepScroll) $('canvas').scrollTop = top; else $('canvas').scrollTop = 0;
  if (phone()) setView('page');
  if (focusTitle) { $('pageTitle').focus(); }
}
function hydrate() {
  $('editor').querySelectorAll('img[data-fid]').forEach(async img => { const u = await fileURL(img.dataset.fid).catch(() => ''); if (u) img.src = u; else img.alt = '그림을 불러오지 못했어요'; });
}
function leavePage(p) {
  savePageNow();
  if (editedSinceOpen && S.autoExport !== 'none' && store.kind === 'drive') autoExport(p);
  editedSinceOpen = false;
}

// ---------- 전자 필기장 · 섹션 · 페이지 ----------
function selectSection(sec) {
  if (cur.page) leavePage(cur.page);
  cur.sec = sec; cur.page = null; searchMode = null; $('search').value = '';
  renderSections(); renderPages();
  const first = sec?.pages[0] || null;
  openPage(first);
  if (phone()) setView('pages');
}
async function addSection() {
  const name = await ask('새 섹션', '섹션 이름', '새 섹션'); if (!name) return;
  const sec = { id: uid(), name, color: SECTION_COLORS[cur.nb.sections.length % SECTION_COLORS.length], pages: [] };
  cur.nb.sections.push(sec); scheduleIndexSave(); selectSection(sec);
}
function addPage(after = cur.page, level) {
  if (!cur.sec) return;
  const now = Date.now(), p = { id: uid(), title: '', created: now, updated: now, level: level ?? 0, ...(S.defaultLines ? { lines: S.defaultLines } : {}) };
  const i = after ? cur.sec.pages.indexOf(after) : -1;
  if (i >= 0) { let j = i + 1; if (level === undefined) { while (j < cur.sec.pages.length && cur.sec.pages[j].level) j++; } cur.sec.pages.splice(j, 0, p); }
  else cur.sec.pages.push(p);
  searchMode = null;
  scheduleIndexSave();
  openPage(p, { focusTitle: true });
}
async function addNotebook() {
  const name = await ask('새 전자 필기장', '전자 필기장 이름', '새 전자 필기장'); if (!name) return;
  const nb = { id: uid(), name, sections: [{ id: uid(), name: '새 섹션', color: SECTION_COLORS[1], pages: [] }] };
  ix.notebooks.push(nb); scheduleIndexSave();
  if (cur.page) leavePage(cur.page);
  cur.nb = nb; cur.sec = nb.sections[0]; cur.page = null; renderAll(); openPage(null);
}
async function trashPages(pages) {
  for (const p of pages) { try { await store.deletePage(p); } catch (e) { console.warn(e); } localStorage.removeItem('dn-draft-' + p.id); }
}
const trashNote = () => store.kind === 'drive' ? ' 드라이브 휴지통으로 옮겨져서 30일 안에는 되살릴 수 있어요.' : ' 되살릴 수 없어요.';

// 목록 클릭
$('sectionList').addEventListener('click', e => {
  const li = e.target.closest('li[data-id]'); if (!li) return;
  const sec = cur.nb.sections.find(s => s.id === li.dataset.id);
  if (e.target.closest('[data-more]')) { e.stopPropagation(); sectionMenu(sec, e.target.closest('[data-more]')); return; }
  if (sec !== cur.sec) selectSection(sec); else if (phone()) setView('pages');
});
$('pageList').addEventListener('click', e => {
  const li = e.target.closest('li[data-id]'); if (!li) return;
  const hit = findPage(li.dataset.id); if (!hit) return;
  if (e.target.closest('[data-more]')) { e.stopPropagation(); pageMenu(hit.p, e.target.closest('[data-more]')); return; }
  if (searchMode && (hit.sec !== cur.sec)) { if (cur.page) leavePage(cur.page); cur.nb = hit.nb; cur.sec = hit.sec; cur.page = null; renderNotebook(); renderSections(); }
  if (hit.p !== cur.page) openPage(hit.p); else if (phone()) setView('page');
});
['sectionList', 'pageList'].forEach(id => $(id).addEventListener('contextmenu', e => {
  const li = e.target.closest('li[data-id]'); if (!li) return; e.preventDefault();
  if (id === 'sectionList') sectionMenu(cur.nb.sections.find(s => s.id === li.dataset.id), { x: e.clientX, y: e.clientY });
  else { const h = findPage(li.dataset.id); if (h) pageMenu(h.p, { x: e.clientX, y: e.clientY }); }
}));
$('addSection').onclick = () => cur.nb && addSection();
$('addPage').onclick = () => addPage(null);
$('pageTitle').addEventListener('input', () => { if (cur.page) { cur.page.title = $('pageTitle').value; const li = $('pageList').querySelector(`li[data-id="${cur.page.id}"] .name`); if (li) li.textContent = cur.page.title || '제목 없음'; markDirty(); } });
$('pageTitle').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); const ed = $('editor'); ed.focus(); const r = document.createRange(); r.selectNodeContents(ed); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); } });

// 메뉴
function showCtx(html, at) {
  const c = $('ctx'); c.innerHTML = html; c.hidden = false;
  const r = at.getBoundingClientRect ? at.getBoundingClientRect() : { left: at.x, bottom: at.y, top: at.y };
  const w = c.offsetWidth, h = c.offsetHeight;
  c.style.left = Math.max(8, Math.min(innerWidth - w - 8, r.left)) + 'px';
  c.style.top = (r.bottom + h + 8 > innerHeight ? Math.max(8, r.top - h - 4) : r.bottom + 4) + 'px';
  return c;
}
document.addEventListener('mousedown', e => {
  if (!$('ctx').hidden && !e.target.closest('#ctx')) $('ctx').hidden = true;
  if (!$('nbMenu').hidden && !e.target.closest('.nb')) $('nbMenu').hidden = true;
  if (!e.target.closest('.drop')) document.querySelectorAll('.pop').forEach(p => { p.hidden = true; });
});
function sectionMenu(sec, at) {
  const i = cur.nb.sections.indexOf(sec);
  const c = showCtx(`<button data-a="rename">이름 바꾸기</button><div class="colors">${SECTION_COLORS.map(col => `<button data-color="${col}" style="background:${col}" aria-label="색"></button>`).join('')}</div><hr>
    <button data-a="up" ${i ? '' : 'disabled'}>위로</button><button data-a="down" ${i < cur.nb.sections.length - 1 ? '' : 'disabled'}>아래로</button>
    <button data-a="move">다른 전자 필기장으로 옮기기</button><hr><button data-a="del" class="danger">섹션 삭제</button>`, at);
  c.onclick = async e => {
    const b = e.target.closest('button'); if (!b) return; c.hidden = true;
    if (b.dataset.color) { sec.color = b.dataset.color; }
    else if (b.dataset.a === 'rename') { const n = await ask('섹션 이름 바꾸기', '섹션 이름', sec.name); if (!n) return; sec.name = n; }
    else if (b.dataset.a === 'up' || b.dataset.a === 'down') { const j = i + (b.dataset.a === 'up' ? -1 : 1); cur.nb.sections.splice(i, 1); cur.nb.sections.splice(j, 0, sec); }
    else if (b.dataset.a === 'move') { const to = await pickNotebook(cur.nb); if (!to) return; cur.nb.sections.splice(cur.nb.sections.indexOf(sec), 1); to.sections.push(sec); if (cur.sec === sec) { cur.nb = to; renderNotebook(); } }
    else if (b.dataset.a === 'del') {
      if (!(await confirmBox(`“${sec.name}” 섹션 삭제`, `이 섹션의 페이지 ${sec.pages.length}개도 함께 지워져요.${trashNote()}`))) return;
      cur.nb.sections.splice(cur.nb.sections.indexOf(sec), 1);
      if (cur.sec === sec) { dirty = false; cur.page = null; cur.sec = cur.nb.sections[0] || null; openPage(cur.sec?.pages[0] || null); }
      trashPages(sec.pages);
    }
    scheduleIndexSave(); renderSections(); renderPages();
  };
}
function pageMenu(p, at) {
  const sec = cur.sec.pages.includes(p) ? cur.sec : findPage(p.id).sec, i = sec.pages.indexOf(p);
  const c = showCtx(`<button data-a="sub">${p.level ? '하위 페이지 해제' : '하위 페이지로 만들기'}</button><button data-a="child">아래에 하위 페이지 추가</button>
    <button data-a="up" ${i ? '' : 'disabled'}>위로</button><button data-a="down" ${i < sec.pages.length - 1 ? '' : 'disabled'}>아래로</button>
    <button data-a="move">다른 섹션으로 옮기기</button><button data-a="copy">복사본 만들기</button><hr><button data-a="del" class="danger">페이지 삭제</button>`, at);
  c.onclick = async e => {
    const b = e.target.closest('button'); if (!b) return; c.hidden = true;
    const a = b.dataset.a;
    if (a === 'sub') p.level = p.level ? 0 : 1;
    else if (a === 'child') { addPage(p, 1); return; }
    else if (a === 'up' || a === 'down') { sec.pages.splice(i, 1); sec.pages.splice(i + (a === 'up' ? -1 : 1), 0, p); }
    else if (a === 'move') {
      const to = await pickSection(sec); if (!to) return;
      if (p === cur.page) savePageNow();
      sec.pages.splice(sec.pages.indexOf(p), 1); p.level = 0; to.pages.push(p);
      if (p === cur.page) openPage(sec.pages[0] || null);
    }
    else if (a === 'copy') {
      if (p === cur.page) await flush();
      const html = p === cur.page ? serialize() : await store.loadPage(p).catch(() => '');
      const now = Date.now(), np = { id: uid(), title: (p.title || '제목 없음') + ' (복사본)', created: now, updated: now, level: p.level, snip: p.snip };
      sec.pages.splice(i + 1, 0, np);
      try { await store.savePage(np, html); } catch (er) { handleErr(er); }
    }
    else if (a === 'del') {
      if (!(await confirmBox(`“${p.title || '제목 없음'}” 페이지 삭제`, `이 페이지를 지울까요?${trashNote()}`))) return;
      sec.pages.splice(sec.pages.indexOf(p), 1);
      if (p === cur.page) { dirty = false; editedSinceOpen = false; cur.page = null; openPage(sec.pages[Math.min(i, sec.pages.length - 1)] || null); }
      trashPages([p]);
    }
    scheduleIndexSave(); renderPages();
  };
}
function pickSection(from) {
  return new Promise(res => {
    const opts = allSections().filter(x => x.sec !== from);
    if (!opts.length) { toast('옮길 다른 섹션이 없어요.'); return res(null); }
    modal(`<h2>어느 섹션으로 옮길까요?</h2>${opts.map((x, k) => `<button class="choice" data-k="${k}"><b><span style="color:${x.sec.color}">■</span> ${esc(x.sec.name)}</b><small>${esc(x.nb.name)}</small></button>`).join('')}<div class="row"><button class="btn" id="pNo">취소</button></div>`, m => {
      m.onclick = e => { const b = e.target.closest('[data-k]'); if (b) { closeModal(); res(opts[+b.dataset.k].sec); } else if (e.target.id === 'pNo') { closeModal(); res(null); } };
    });
  });
}
function pickNotebook(from) {
  return new Promise(res => {
    const opts = ix.notebooks.filter(n => n !== from);
    if (!opts.length) { toast('다른 전자 필기장이 없어요. 먼저 새 전자 필기장을 만드세요.'); return res(null); }
    modal(`<h2>어느 전자 필기장으로 옮길까요?</h2>${opts.map((n, k) => `<button class="choice" data-k="${k}"><b>${esc(n.name)}</b><small>섹션 ${n.sections.length}개</small></button>`).join('')}<div class="row"><button class="btn" id="pNo">취소</button></div>`, m => {
      m.onclick = e => { const b = e.target.closest('[data-k]'); if (b) { closeModal(); res(opts[+b.dataset.k]); } else if (e.target.id === 'pNo') { closeModal(); res(null); } };
    });
  });
}

// 전자 필기장 메뉴
$('nbBtn').onclick = () => {
  const m = $('nbMenu');
  if (!m.hidden) { m.hidden = true; return; }
  m.innerHTML = ix.notebooks.map(n => `<button data-nb="${n.id}" class="${n === cur.nb ? 'on' : ''}">▤ ${esc(n.name)}</button>`).join('') +
    `<hr><button data-a="new">＋ 새 전자 필기장</button><button data-a="rename">이 전자 필기장 이름 바꾸기</button><button data-a="del" class="danger">이 전자 필기장 삭제</button>`;
  m.hidden = false;
};
$('nbMenu').onclick = async e => {
  const b = e.target.closest('button'); if (!b) return; $('nbMenu').hidden = true;
  if (b.dataset.nb) selectNotebook(ix.notebooks.find(n => n.id === b.dataset.nb));
  else if (b.dataset.a === 'new') addNotebook();
  else if (b.dataset.a === 'rename') renameNotebook(cur.nb);
  else if (b.dataset.a === 'del') deleteNotebook(cur.nb);
};
function selectNotebook(nb) {
  if (!nb || nb === cur.nb) return;
  curCollapsed = false;
  if (cur.page) leavePage(cur.page);
  cur.nb = nb; cur.sec = nb.sections[0] || null; cur.page = null; searchMode = null; renderAll(); openPage(cur.sec?.pages[0] || null);
  if (phone()) setView('sections');
}
async function renameNotebook(nb) {
  const n = await ask('전자 필기장 이름 바꾸기', '이름', nb.name);
  if (n) { nb.name = n; scheduleIndexSave(); renderNotebook(); }
}
async function deleteNotebook(nb) {
  if (ix.notebooks.length <= 1) return toast('전자 필기장이 하나뿐이라 지울 수 없어요.');
  const pages = nb.sections.flatMap(s => s.pages);
  if (!(await confirmBox(`“${nb.name}” 전자 필기장 삭제`, `섹션 ${nb.sections.length}개, 페이지 ${pages.length}개가 함께 지워져요.${trashNote()}`))) return;
  ix.notebooks.splice(ix.notebooks.indexOf(nb), 1);
  if (nb === cur.nb) { dirty = false; editedSinceOpen = false; cur.nb = ix.notebooks[0]; cur.sec = cur.nb.sections[0] || null; cur.page = null; openPage(cur.sec?.pages[0] || null); }
  scheduleIndexSave(); renderAll(); trashPages(pages);
}
// 전자 필기장 칸 (넓은 화면)
$('nbList').addEventListener('click', e => {
  const li = e.target.closest('li'); if (!li) return;
  const more = e.target.closest('[data-more]');
  if (li.dataset.addsec) { selectNotebook(ix.notebooks.find(n => n.id === li.dataset.addsec)); addSection(); return; }
  if (li.dataset.sec) { // 트리의 섹션
    const nb = ix.notebooks.find(n => n.id === li.dataset.nb), sec = nb?.sections.find(s => s.id === li.dataset.sec); if (!sec) return;
    if (nb !== cur.nb) { if (cur.page) leavePage(cur.page); cur.nb = nb; renderNotebook(); }
    if (more) { e.stopPropagation(); sectionMenu(sec, more); return; }
    if (sec !== cur.sec) selectSection(sec);
    return;
  }
  const nb = ix.notebooks.find(n => n.id === li.dataset.id); if (!nb) return;
  if (more) { e.stopPropagation(); notebookMenu(nb, more); return; }
  if (e.target.closest('[data-chev]')) { // 화살표: 펼치기·접기만
    if (openNbs.has(nb.id)) openNbs.delete(nb.id); else openNbs.add(nb.id);
    if (nb === cur.nb) curCollapsed = !openNbs.has(nb.id);
    saveOpenNbs(); renderNotebook(); return;
  }
  curCollapsed = false; openNbs.add(nb.id); saveOpenNbs();
  if (nb === cur.nb) renderNotebook(); else selectNotebook(nb);
});
$('nbList').addEventListener('contextmenu', e => {
  const li = e.target.closest('li'); if (!li) return; e.preventDefault();
  if (li.dataset.sec) { const nb = ix.notebooks.find(n => n.id === li.dataset.nb); if (nb !== cur.nb) { if (cur.page) leavePage(cur.page); cur.nb = nb; renderNotebook(); } const sec = nb.sections.find(s => s.id === li.dataset.sec); if (sec) sectionMenu(sec, { x: e.clientX, y: e.clientY }); return; }
  if (li.dataset.id) notebookMenu(ix.notebooks.find(n => n.id === li.dataset.id), { x: e.clientX, y: e.clientY });
});
$('addNotebook').onclick = () => addNotebook();
function notebookMenu(nb, at) {
  const c = showCtx(`<button data-a="rename">이름 바꾸기</button><hr><button data-a="del" class="danger">전자 필기장 삭제</button>`, at);
  c.onclick = e => { const b = e.target.closest('button'); if (!b) return; c.hidden = true; if (b.dataset.a === 'rename') renameNotebook(nb); else deleteNotebook(nb); };
}

// 끌어서 순서 바꾸기 · 페이지를 섹션으로 옮기기
let dragItem = null;
['sectionList', 'pageList'].forEach(id => {
  const ul = $(id);
  ul.addEventListener('dragstart', e => { const li = e.target.closest('li[data-id]'); if (!li || searchMode) return e.preventDefault(); dragItem = { kind: id, id: li.dataset.id }; e.dataTransfer.effectAllowed = 'move'; });
  ul.addEventListener('dragover', e => {
    const li = e.target.closest('li[data-id]'); if (!dragItem || !li) return;
    e.preventDefault(); ul.querySelectorAll('.drop-before, .drop-into').forEach(x => x.classList.remove('drop-before', 'drop-into'));
    li.classList.add(dragItem.kind === 'pageList' && id === 'sectionList' ? 'drop-into' : 'drop-before');
  });
  ul.addEventListener('dragleave', e => { const li = e.target.closest('li'); if (li) li.classList.remove('drop-before', 'drop-into'); });
  ul.addEventListener('drop', e => {
    const li = e.target.closest('li[data-id]'); if (!dragItem || !li) return; e.preventDefault();
    li.classList.remove('drop-before', 'drop-into');
    if (dragItem.kind === 'sectionList' && id === 'sectionList') {
      const a = cur.nb.sections.findIndex(s => s.id === dragItem.id), sec = cur.nb.sections[a]; if (!sec || sec.id === li.dataset.id) return;
      cur.nb.sections.splice(a, 1); cur.nb.sections.splice(cur.nb.sections.findIndex(s => s.id === li.dataset.id), 0, sec); renderSections();
    } else if (dragItem.kind === 'pageList' && id === 'pageList') {
      const a = cur.sec.pages.findIndex(p => p.id === dragItem.id), p = cur.sec.pages[a]; if (!p || p.id === li.dataset.id) return;
      cur.sec.pages.splice(a, 1); cur.sec.pages.splice(cur.sec.pages.findIndex(x => x.id === li.dataset.id), 0, p); renderPages();
    } else if (dragItem.kind === 'pageList' && id === 'sectionList') {
      const to = cur.nb.sections.find(s => s.id === li.dataset.id); if (!to || to === cur.sec) return;
      const p = cur.sec.pages.find(x => x.id === dragItem.id); if (!p) return;
      if (p === cur.page) { savePageNow(); }
      cur.sec.pages.splice(cur.sec.pages.indexOf(p), 1); p.level = 0; to.pages.push(p);
      if (p === cur.page) openPage(cur.sec.pages[0] || null); else renderPages();
      toast(`“${p.title || '제목 없음'}”을(를) ${to.name} 섹션으로 옮겼어요.`);
    }
    dragItem = null; scheduleIndexSave();
  });
  ul.addEventListener('dragend', () => { dragItem = null; document.querySelectorAll('.drop-before, .drop-into').forEach(x => x.classList.remove('drop-before', 'drop-into')); });
});

// ---------- 검색 ----------
let searchTimer = 0;
$('search').addEventListener('input', () => { clearTimeout(searchTimer); const q = $('search').value.trim(); if (!q) return endSearch(); searchTimer = setTimeout(() => runSearch(q), 450); });
$('search').addEventListener('keydown', e => { if (e.key === 'Escape') { $('search').value = ''; endSearch(); } if (e.key === 'Enter') { clearTimeout(searchTimer); const q = $('search').value.trim(); if (q) runSearch(q); } });
async function runSearch(q) {
  await flush().catch(() => {});
  const k = q.toLowerCase(), list = allPages();
  const byTitle = list.filter(x => (x.p.title || '').toLowerCase().includes(k) || (x.p.snip || '').toLowerCase().includes(k)).map(x => x.p.id);
  searchMode = { q, ids: byTitle }; renderPages(); if (phone()) setView('pages');
  try {
    const more = await store.search(q, list.map(x => x.p));
    if (searchMode?.q !== q) return;
    searchMode.ids = [...new Set([...byTitle, ...more])]; renderPages();
  } catch (e) { if (!e.offline) handleErr(e); }
}
function endSearch() { searchMode = null; renderPages(); }

// ---------- 리본 ----------
document.querySelectorAll('.ribbon-tabs button').forEach(b => b.onclick = () => {
  document.querySelectorAll('.ribbon-tabs button').forEach(x => x.setAttribute('aria-selected', x === b));
  document.querySelectorAll('.panel').forEach(p => { p.hidden = p.dataset.panel !== b.dataset.tab; });
  updateTableOps();
});
// 리본 버튼을 눌러도 본문 선택이 풀리지 않게
$('ribbon').addEventListener('mousedown', e => { if (e.target.closest('button') && !e.target.closest('.big')) e.preventDefault(); });
let savedRange = null;
document.addEventListener('selectionchange', () => {
  const s = getSelection();
  if (s.rangeCount && $('editor').contains(s.anchorNode)) { savedRange = s.getRangeAt(0).cloneRange(); updateTableOps(); }
});
function restoreRange() {
  if (!savedRange || !$('editor').contains(savedRange.startContainer)) return;
  $('editor').focus(); const s = getSelection(); s.removeAllRanges(); s.addRange(savedRange);
}
function updateTableOps() { $('tableOps').classList.toggle('off', !caretIn('td, th')); }
const needPage = () => { if (!cur.page) { toast('먼저 페이지를 열거나 만드세요.'); return false; } return true; };

document.querySelectorAll('[data-cmd]').forEach(b => b.onclick = () => { if (!needPage()) return; restoreRange(); exec(b.dataset.cmd); });
document.querySelectorAll('[data-block]').forEach(b => b.onclick = () => { if (!needPage()) return; restoreRange(); formatBlock(b.dataset.block); });
$('fontSize').onchange = () => { const v = $('fontSize').value; $('fontSize').value = ''; if (!v || !needPage()) return; restoreRange(); fontSize(+v); };
$('btnCheck').onclick = () => { if (!needPage()) return; restoreRange(); toggleChecklist(); };
$('autoQuick').onclick = () => { S.auto = !S.auto; saveSettings(); toast(S.auto ? '자동 목록을 켰어요. “1. ” “- ” “[] ”를 입력하면 목록이 돼요.' : '자동 목록을 껐어요. “1. ”을 입력해도 글자 그대로 남아요.'); };

// 색 고르기
const TEXT_COLORS = ['#000000', '#e03131', '#f08c00', '#2f9e44', '#1971c2', '#7048e8', '#c2255c', '#0c8599', '#868e96', '#5c940d', '#a61e4d', '#ffffff'];
const HL_COLORS = ['#fff176', '#b9f6ca', '#b3e5fc', '#f8bbd0', '#ffe0b2', '#e1bee7', '#ffcdd2', '#c5cae9', '#dcedc8', '#f0f4c3', '#d7ccc8', ''];
$('colorPop').innerHTML = TEXT_COLORS.map(c => `<button class="sw" data-c="${c}" style="background:${c}" title="${c}"></button>`).join('');
$('hlPop').innerHTML = HL_COLORS.map(c => `<button class="sw${c ? '' : ' none'}" data-c="${c}" style="${c ? 'background:' + c : ''}" title="${c ? c : '형광펜 없애기'}"></button>`).join('');
document.querySelectorAll('[data-pop]').forEach(b => b.onclick = () => {
  const p = $(b.dataset.pop), open = p.hidden;
  document.querySelectorAll('.pop').forEach(x => { x.hidden = true; });
  if (!open) return;
  p.hidden = false; const r = b.getBoundingClientRect();
  p.style.left = Math.max(8, Math.min(innerWidth - p.offsetWidth - 8, r.left)) + 'px'; p.style.top = r.bottom + 4 + 'px';
});
$('colorPop').onclick = e => { const c = e.target.closest('[data-c]'); if (!c || !needPage()) return; $('colorPop').hidden = true; $('colorBar').style.background = c.dataset.c; restoreRange(); exec('foreColor', c.dataset.c); };
$('hlPop').onclick = e => { const c = e.target.closest('[data-c]'); if (!c || !needPage()) return; $('hlPop').hidden = true; if (c.dataset.c) $('hlBar').style.background = c.dataset.c; restoreRange(); exec('hiliteColor', c.dataset.c); };
// 표 크기 고르기
$('tgrid').innerHTML = Array.from({ length: 64 }, (_, i) => `<i data-r="${Math.floor(i / 8) + 1}" data-c="${i % 8 + 1}"></i>`).join('');
$('tgrid').addEventListener('mouseover', e => { const t = e.target.closest('i'); if (!t) return; const r = +t.dataset.r, c = +t.dataset.c; $('tgrid').querySelectorAll('i').forEach(x => x.classList.toggle('on', +x.dataset.r <= r && +x.dataset.c <= c)); $('tlabel').textContent = `${r} × ${c} 표`; });
$('tgrid').addEventListener('click', e => { const t = e.target.closest('i'); if (!t || !needPage()) return; $('tablePop').hidden = true; restoreRange(); insertHTML(tableHTML(+t.dataset.r, +t.dataset.c)); });
document.querySelectorAll('[data-table]').forEach(b => b.onclick = () => { restoreRange(); if (!tableOp(b.dataset.table)) toast('표 안을 먼저 누르세요.'); });
$('btnImage').onclick = () => needPage() && $('pickImage').click();
$('btnFile').onclick = () => needPage() && $('pickFile').click();
$('pickImage').onchange = async () => { const fs = [...$('pickImage').files]; $('pickImage').value = ''; restoreRange(); for (const f of fs) await addFile(f); };
$('pickFile').onchange = async () => { const fs = [...$('pickFile').files]; $('pickFile').value = ''; restoreRange(); for (const f of fs) await addFile(f); };
$('btnDate').onclick = () => { if (!needPage()) return; restoreRange(); insertText(new Date().toLocaleDateString('ko-KR', { year: 'numeric', month: 'long', day: 'numeric', weekday: 'short' })); };
$('btnTime').onclick = () => { if (!needPage()) return; restoreRange(); insertText(new Date().toLocaleTimeString('ko-KR', { hour: 'numeric', minute: '2-digit' })); };
$('btnHr').onclick = () => { if (!needPage()) return; restoreRange(); insertHTML('<hr><p><br></p>'); };
$('btnLink').onclick = async () => {
  if (!needPage()) return;
  const r = savedRange && savedRange.cloneRange(), text = r ? r.toString() : '';
  let url = await ask('링크 넣기', 'https://…', text.match(/^https?:\/\//) ? text : 'https://'); if (!url) return;
  if (!/^(https?:|mailto:)/i.test(url)) url = 'https://' + url;
  savedRange = r; restoreRange();
  if (text) exec('createLink', url); else insertHTML(`<a href="${esc(url)}" target="_blank" rel="noopener">${esc(url)}</a>&nbsp;`);
};
$('openSettings').onclick = () => settingsModal();
$('openStorage').onclick = () => storageModal();
$('account').onclick = () => storageModal();

// ---------- 노트선 (보기 탭) ----------
function applyLines(p) {
  $('editor').className = 'editor' + (p && p.lines ? ' lines-' + p.lines : '');
  updateLinesUI();
  applyWrap();
}
function updateLinesUI() {
  const now = cur.page?.lines || '';
  document.querySelectorAll('[data-lines]').forEach(b => b.classList.toggle('on', b.dataset.lines === now));
  const d = $('linesDefault'), on = !!S.defaultLines;
  d.classList.toggle('is-on', on);
  d.querySelector('b').textContent = on ? { narrow: '좁은 줄', wide: '넓은 줄', grid: '모눈' }[S.defaultLines] : '끔';
}
document.querySelectorAll('[data-lines]').forEach(b => b.onclick = () => {
  if (!needPage()) return;
  if (b.dataset.lines) cur.page.lines = b.dataset.lines; else delete cur.page.lines;
  applyLines(cur.page); scheduleIndexSave();
});
$('linesDefault').onclick = () => {
  if (S.defaultLines) { S.defaultLines = ''; saveSettings(); return toast('새 페이지는 노트선 없이 만들어요.'); }
  if (!cur.page?.lines) return toast('먼저 이 페이지의 노트선(좁은 줄·넓은 줄·모눈)을 고르세요.');
  S.defaultLines = cur.page.lines; saveSettings();
  toast('앞으로 새로 만드는 페이지에 이 노트선을 써요.');
};

// ---------- 맞춤법 검사 (검토 탭) ----------
const SPELL_API = (window.NOTE_CONFIG && window.NOTE_CONFIG.spellApi) || 'api/spellcheck';
const SPELL_BLOCKS = 'p, h1, h2, h3, h4, h5, h6, li, td, th, blockquote, pre';
const norm = t => t.replace(/ /g, ' ');
// 블록이 직접 가진 글자 (안쪽 블록과 첨부 파일 이름은 빼고)
function ownTextNodes(block) {
  const ed = $('editor'), w = document.createTreeWalker(block, NodeFilter.SHOW_TEXT), out = [];
  for (let n; (n = w.nextNode());) {
    if (n.parentElement.closest('.att')) continue;
    const c = n.parentElement.closest(SPELL_BLOCKS);
    if (c === block || (!c && block === ed)) out.push(n);
  }
  return out;
}
function spellBlocks() {
  const ed = $('editor');
  return [ed, ...ed.querySelectorAll(SPELL_BLOCKS)]
    .map(el => ({ el, text: norm(ownTextNodes(el).map(n => n.nodeValue).join('')) }))
    .filter(b => b.text.trim().length > 1);
}
function findRange(el, needle) {
  const nodes = ownTextNodes(el), i = norm(nodes.map(n => n.nodeValue).join('')).indexOf(needle);
  if (i < 0 || !nodes.length) return null;
  const at = (k, start) => {
    let acc = 0;
    for (const n of nodes) { const L = n.nodeValue.length; if (start ? k < acc + L : k <= acc + L) return [n, k - acc]; acc += L; }
    const l = nodes[nodes.length - 1]; return [l, l.nodeValue.length];
  };
  const r = document.createRange(); r.setStart(...at(i, true)); r.setEnd(...at(i + needle.length, false));
  return r;
}
function selectRange(r) { const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
function applyFix(it) {
  if (!it.block.el.isConnected) return false;
  const r = findRange(it.block.el, it.wrong); if (!r) return false;
  $('editor').focus(); selectRange(r);
  document.execCommand('insertText', false, it.right); // Ctrl+Z로 되돌릴 수 있게
  markDirty(); return true;
}
async function runSpell() {
  if (!needPage()) return;
  if (store?.kind !== 'drive') return toast('맞춤법 검사는 구글 드라이브로 로그인했을 때 쓸 수 있어요.');
  if (!store.signedIn) return handleErr({ auth: true }, runSpell);
  const blocks = spellBlocks();
  if (!blocks.length) return toast('검사할 글이 없어요.');
  showSpell(null);
  // 한 번에 보내는 양을 나눠서
  const groups = [[]]; let len = 0;
  for (const b of blocks) {
    if (len + b.text.length > 4500 && groups[groups.length - 1].length) { groups.push([]); len = 0; }
    groups[groups.length - 1].push(b); len += Math.min(b.text.length, 2000);
  }
  const items = [];
  try {
    for (const g of groups) {
      const r = await fetch(SPELL_API, { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + store.token }, body: JSON.stringify({ p: g.map(b => b.text.slice(0, 2000)) }) });
      if (r.status === 503) { $('spell').hidden = true; return toast('맞춤법 검사 서버에 OpenAI 키가 설정돼 있지 않아요.', [], 10000); }
      if (r.status === 401) { $('spell').hidden = true; return toast('이 계정은 맞춤법 검사를 쓸 수 없어요. 다시 로그인해 보세요.', [], 10000); }
      if (r.status === 404) { $('spell').hidden = true; return toast('맞춤법 검사 서버를 찾지 못했어요. config.js의 spellApi를 확인해 주세요.', [], 8000); }
      const j = await r.json().catch(() => ({}));
      if (!r.ok) throw new Error('맞춤법 검사 오류: ' + (j.detail || j.error || r.status));
      items.push(...(j.items || []).map(it => ({ ...it, block: g[it.p] })).filter(it => it.block));
    }
  } catch (e) { $('spell').hidden = true; return handleErr(e && e.message ? e : new Error('맞춤법 검사를 하지 못했어요.')); }
  showSpell(items);
}
function showSpell(items) {
  const el = $('spell'); el.hidden = false;
  const head = (title, all) => `<div class="spell-head"><b>${title}</b>${all ? '<button class="all" id="spAll">모두 고치기</button>' : ''}<button id="spClose" aria-label="닫기">✕</button></div>`;
  if (!items) {
    el.innerHTML = head('맞춤법 검사') + '<div class="spell-empty">검사하는 중…</div>';
    $('spClose').onclick = () => { el.hidden = true; };
    return;
  }
  el.innerHTML = head(items.length ? `고칠 곳 ${items.length}개` : '맞춤법 검사', items.length > 1) +
    `<div class="spell-list">${items.length ? items.map((it, i) => `<div class="spell-item" data-i="${i}"><div class="fix"><del>${esc(it.wrong)}</del> → <ins>${esc(it.right)}</ins></div>${it.why ? `<small>${esc(it.why)}</small>` : ''}<div class="acts"><button class="ok" data-a="fix">고치기</button><button data-a="skip">무시</button></div></div>`).join('') : '<div class="spell-empty">틀린 곳을 찾지 못했어요 👍</div>'}</div>`;
  $('spClose').onclick = () => { el.hidden = true; };
  const done = (box, msg) => {
    box.classList.add('done'); box.querySelector('.acts').innerHTML = `<small>${msg}</small>`;
    if (!el.querySelector('.spell-item:not(.done)')) el.querySelector('.spell-head b').textContent = '다 확인했어요';
  };
  $('spAll')?.addEventListener('click', () => {
    let n = 0;
    el.querySelectorAll('.spell-item:not(.done)').forEach(box => { if (applyFix(items[+box.dataset.i])) { n++; done(box, '고쳤어요'); } else done(box, '이미 바뀌어서 건너뜀'); });
    toast(`${n}곳을 고쳤어요. 되돌리려면 Ctrl+Z를 누르세요.`);
  });
  el.querySelector('.spell-list').onclick = e => {
    const box = e.target.closest('.spell-item'); if (!box || box.classList.contains('done')) return;
    const it = items[+box.dataset.i], a = e.target.closest('[data-a]')?.dataset.a;
    if (a === 'fix') return applyFix(it) ? done(box, '고쳤어요') : done(box, '이미 바뀌어서 건너뜀');
    if (a === 'skip') return done(box, '무시했어요');
    // 항목을 누르면 본문에서 그 자리를 보여 줌
    const r = it.block.el.isConnected && findRange(it.block.el, it.wrong);
    if (r) { it.block.el.scrollIntoView({ block: 'center', behavior: 'smooth' }); $('editor').focus(); selectRange(r); }
  };
}
$('spellRun').onclick = runSpell;
$('spellLive').onclick = () => { S.browserSpell = !S.browserSpell; saveSettings(); };
// ---------- AI 도우미 (검토 탭): 페이지에 대해 묻기, 페이지 전체 고치기 ----------
const AI_API = (window.NOTE_CONFIG && window.NOTE_CONFIG.aiApi) || 'api/note-ai';
const AI_PRESETS = [
  ['주제별로 정리', '내용을 하나도 빼지 말고 주제별로 묶어서, 주제마다 제목(h2)을 붙여 정리해 주세요. 같은 내용이 여러 번 나오면 하나로 합쳐 주세요.'],
  ['문장 다듬기', '뜻과 내용은 그대로 두고 맞춤법, 띄어쓰기, 어색한 문장을 자연스럽게 다듬어 주세요.'],
  ['맨 위에 요약 넣기', '맨 위에 "요약" 제목(h2)과 핵심 내용 3~5줄 글머리 기호 요약을 넣고, 나머지 내용은 그대로 두세요.'],
  ['표로 정리', '서로 비교하거나 나열된 정보가 있으면 표로 정리해 주세요. 나머지 내용은 그대로 두세요.'],
];
// 그림·첨부 파일은 AI에 보내지 않고 ⟦그림1⟧ 같은 표시로 바꿨다가 되돌림
let aiKeep = [];
function pageForAI() {
  const c = $('editor').cloneNode(true);
  aiKeep = [];
  c.querySelectorAll('img, .att').forEach(el => {
    const k = el.cloneNode(true); if (k.tagName === 'IMG' && k.dataset.fid) k.removeAttribute('src'); k.classList.remove('sel');
    const token = `⟦${el.tagName === 'IMG' ? '그림' : '첨부'}${aiKeep.length + 1}⟧`;
    aiKeep.push({ token, html: k.outerHTML });
    el.replaceWith(token);
  });
  return c.innerHTML;
}
function restoreKept(html) {
  let out = sanitize(html), missing = 0;
  for (const k of aiKeep) {
    if (out.includes(k.token)) out = out.replace(k.token, k.html);
    else { out += `<p>${k.html}</p>`; missing++; }
  }
  return { html: out, missing };
}
async function aiCall(mode, prompt) {
  const r = await fetch(AI_API, { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + store.token }, body: JSON.stringify({ mode, prompt, title: $('pageTitle').value, html: pageForAI() }) });
  const j = await r.json().catch(() => ({}));
  if (r.ok) return j;
  const msg = { 503: 'AI 서버에 OpenAI 키가 설정돼 있지 않아요.', 401: '이 계정은 AI 도우미를 쓸 수 없어요. 다시 로그인해 보세요.', 413: '페이지가 너무 길어요. 페이지를 나눠서 해 보세요.', 404: 'AI 서버를 찾지 못했어요. config.js의 aiApi를 확인해 주세요.' }[r.status];
  throw new Error(msg || 'AI 오류: ' + (j.detail || j.error || r.status));
}
function answerHTML(text) {
  const lines = text.split('\n'); let html = '', list = false;
  for (const l of lines) {
    const m = l.match(/^\s*[-•*]\s+(.*)$/);
    if (m) { if (!list) { html += '<ul>'; list = true; } html += `<li>${esc(m[1])}</li>`; continue; }
    if (list) { html += '</ul>'; list = false; }
    if (l.trim()) html += `<p>${esc(l)}</p>`;
  }
  return html + (list ? '</ul>' : '');
}
function openAI() {
  if (!needPage()) return;
  if (store?.kind !== 'drive') return toast('AI 도우미는 구글 드라이브로 로그인했을 때 쓸 수 있어요.');
  $('spell').hidden = true;
  const el = $('ai'); el.hidden = false;
  el.innerHTML = `<div class="spell-head"><b>✨ AI 도우미</b><button id="aiClose" aria-label="닫기">✕</button></div>
    <div class="ai-body">
      <div class="ai-label">이 페이지를 바로 고치기</div>
      <div class="ai-chips">${AI_PRESETS.map((p, i) => `<button data-preset="${i}">${esc(p[0])}</button>`).join('')}</div>
      <div class="ai-label">직접 쓰기</div>
      <textarea id="aiPrompt" rows="3" placeholder="궁금한 점을 묻거나, 원하는 걸 적어 주세요.&#10;예) 이 내용으로 발표 순서를 짜 줘&#10;예) 날짜순으로 다시 정리해 줘"></textarea>
      <div class="ai-acts"><button class="btn" id="aiAsk" title="Ctrl+Enter">질문하기</button><button class="btn primary" id="aiEdit">이대로 페이지 고치기</button></div>
      <div id="aiOut"></div>
    </div>`;
  $('aiClose').onclick = () => { el.hidden = true; };
  const promptBox = $('aiPrompt');
  promptBox.onkeydown = e => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); run('ask', promptBox.value.trim()); } };
  el.querySelector('.ai-chips').onclick = e => { const b = e.target.closest('[data-preset]'); if (b) run('edit', AI_PRESETS[+b.dataset.preset][1]); };
  $('aiAsk').onclick = () => run('ask', promptBox.value.trim());
  $('aiEdit').onclick = () => run('edit', promptBox.value.trim());
  promptBox.focus();

  const page = cur.page;
  async function run(mode, prompt) {
    if (!prompt) { promptBox.focus(); return toast('무엇을 할지 적어 주세요.'); }
    if (!store.signedIn) return handleErr({ auth: true }, () => run(mode, prompt));
    const out = $('aiOut');
    el.querySelectorAll('button').forEach(b => { if (b.id !== 'aiClose') b.disabled = true; });
    out.innerHTML = `<div class="spell-empty">AI가 ${mode === 'edit' ? '페이지를 고치는' : '생각하는'} 중…<br><small>긴 페이지는 30초~1분쯤 걸려요.</small></div>`;
    try {
      const j = await aiCall(mode, prompt);
      if (cur.page !== page || el.hidden) return;
      if (mode === 'ask') {
        out.innerHTML = `<div class="ai-answer">${answerHTML(j.answer || '')}</div><div class="ai-acts"><button class="btn" id="aiCopy">복사</button><button class="btn primary" id="aiInsert">페이지 끝에 넣기</button></div>`;
        $('aiCopy').onclick = async () => { try { await navigator.clipboard.writeText(j.answer || ''); toast('복사했어요.'); } catch { toast('복사하지 못했어요.'); } };
        $('aiInsert').onclick = () => {
          // 목록·표 안이 아니라 페이지 맨 아래 새 줄에 넣기
          const ed = $('editor'), p = document.createElement('p'); p.innerHTML = '<br>'; ed.append(p);
          const r = document.createRange(); r.selectNodeContents(p);
          ed.focus(); selectRange(r); document.execCommand('insertHTML', false, answerHTML(j.answer || '')); markDirty();
          toast('페이지 끝에 넣었어요.');
        };
      } else {
        const { html, missing } = restoreKept(j.html || '');
        out.innerHTML = `${j.note ? `<p class="ai-note">${esc(j.note)}</p>` : ''}${missing ? `<p class="ai-note warn">그림·첨부 ${missing}개는 맨 아래에 붙였어요.</p>` : ''}
          <div class="ai-label">미리 보기</div><div class="ai-preview editor">${html}</div>
          <div class="ai-acts"><button class="btn" id="aiDiscard">버리기</button><button class="btn primary" id="aiApply">페이지에 적용</button></div>
          <p class="ai-hint">적용한 뒤에도 Ctrl+Z로 되돌릴 수 있어요.</p>`;
        out.querySelectorAll('img[data-fid]').forEach(async img => { const u = await fileURL(img.dataset.fid).catch(() => ''); if (u) img.src = u; });
        $('aiDiscard').onclick = () => { out.innerHTML = ''; };
        $('aiApply').onclick = () => {
          if (cur.page !== page) return toast('다른 페이지로 바뀌어서 적용하지 않았어요.');
          const ed = $('editor'), r = document.createRange(); r.selectNodeContents(ed);
          ed.focus(); selectRange(r); document.execCommand('insertHTML', false, html);
          hydrate(); markDirty();
          out.innerHTML = '<p class="ai-note">페이지에 적용했어요. 되돌리려면 Ctrl+Z를 누르세요.</p>';
        };
      }
    } catch (e) { out.innerHTML = `<p class="ai-note warn">${esc(e.message || 'AI 도우미가 답하지 못했어요.')}</p>`; }
    finally { el.querySelectorAll('button').forEach(b => { b.disabled = false; }); }
  }
}
$('aiOpen').onclick = openAI;


// ---------- 한글·엑셀 변환 ----------
async function pagesFor(scope, page = cur.page) {
  const list = scope === 'page' ? [page] : cur.sec.pages;
  const out = [];
  for (const p of list) {
    const html = p === cur.page ? serialize() : await store.loadPage(p);
    out.push({ title: p.title || '제목 없음', date: fmtFull(p.created), created: fmtFull(p.created), updated: fmtFull(p.updated), html });
  }
  return out;
}
const fileBase = s => (s || '노트').replace(/[\\/:*?"<>|]/g, ' ').trim().slice(0, 60) || '노트';
async function convert(fmt, pages, whole) {
  return fmt === 'hwpx' ? toHwpx(pages, fid => store.getFile(fid)) : toXlsx(pages, whole);
}
document.querySelectorAll('[data-export]').forEach(b => b.onclick = async () => {
  const fmt = b.dataset.export, scope = b.dataset.scope;
  if (scope === 'page' && !needPage()) return;
  if (scope === 'section' && !cur.sec?.pages.length) return toast('이 섹션에 페이지가 없어요.');
  try {
    toast(`${fmt === 'hwpx' ? '한글' : '엑셀'} 파일로 바꾸는 중…`, [], 0);
    await flush();
    const pages = await pagesFor(scope);
    const blob = await convert(fmt, pages, scope === 'section');
    const name = fileBase(scope === 'page' ? cur.page.title : `${cur.nb.name} - ${cur.sec.name}`) + '.' + fmt;
    download(blob, name);
    let msg = `“${name}” 파일을 내려받았어요.`;
    if (store.kind === 'drive' && S.exportToDrive) { await store.saveExport(blob, name); msg += ' 드라이브 “내보내기” 폴더에도 저장했어요.'; }
    toast(msg, [], 6000);
  } catch (e) { $('toast').hidden = true; handleErr(e); }
});
// 페이지를 닫을 때 드라이브에 자동으로 변환해 두기
async function autoExport(p) {
  const fmts = S.autoExport === 'both' ? ['hwpx', 'xlsx'] : [S.autoExport];
  try {
    await saving;
    const pages = [{ title: p.title || '제목 없음', date: fmtFull(p.created), created: fmtFull(p.created), updated: fmtFull(p.updated), html: await store.loadPage(p) }];
    p.exp ||= {};
    for (const f of fmts) p.exp[f] = await store.saveExport(await convert(f, pages, false), fileBase(p.title) + '.' + f, p.exp[f]);
    scheduleIndexSave();
  } catch (e) { console.warn('자동 변환 실패', e); }
}

// ---------- 설정 창 ----------
function settingsModal() {
  const sw = (k, label, hint, cls = '') => `<label class="sw-row ${cls}"><span>${label}${hint ? `<small>${hint}</small>` : ''}</span><input type="checkbox" class="switch" data-k="${k}" ${S[k] ? 'checked' : ''}></label>`;
  const drive = store?.kind === 'drive';
  modal(`<h2>설정</h2>
    <div class="set"><h3>자동 서식</h3>
      ${sw('auto', '자동 목록 사용', '줄 맨 앞에 아래 글자를 쓰고 띄어쓰기를 하면 목록으로 바뀌어요. 끄면 쓴 그대로 남아요.')}
      ${sw('autoNumber', '“1. ” “1) ” → 번호 목록', '', 'sub-sw')}
      ${sw('autoBullet', '“- ” “* ” → 글머리 기호', '', 'sub-sw')}
      ${sw('autoCheck', '“[] ” → 할 일(체크박스)', '', 'sub-sw')}
      <p style="font-size:12.5px">바뀐 직후 <b>백스페이스</b>를 누르면 그 한 번만 원래 글자로 되돌아가요.</p></div>
    <div class="set"><h3>그림</h3>${sw('shrink', '큰 사진은 줄여서 올리기', '긴 쪽 1600픽셀로 줄여 저장공간을 아껴요. 끄면 원본 그대로 올려요.')}</div>
    <div class="set"><h3>한글·엑셀 변환</h3>
      ${sw('exportToDrive', '바꾼 파일을 드라이브에도 저장', drive ? '“에스 노트/내보내기” 폴더에 저장돼요.' : '구글 드라이브에 연결하면 쓸 수 있어요.')}
      <div class="sw-row" style="cursor:default"><span>자동 변환<small>${drive ? '페이지를 고치고 다른 페이지로 넘어갈 때, 그 페이지를 드라이브 “내보내기” 폴더에 자동으로 변환해 둬요.' : '구글 드라이브에 연결하면 쓸 수 있어요.'}</small></span></div>
      <div class="radio-row">${Object.entries(AUTO_LABEL).map(([v, l]) => `<label><input type="radio" name="autoExport" value="${v}" ${S.autoExport === v ? 'checked' : ''} ${drive ? '' : 'disabled'}>${l}</label>`).join('')}</div></div>
    <div class="set"><h3>화면</h3><div class="radio-row">${[['system', '기기 설정 따라'], ['light', '밝게'], ['dark', '어둡게']].map(([v, l]) => `<label><input type="radio" name="theme" value="${v}" ${S.theme === v ? 'checked' : ''}>${l}</label>`).join('')}</div></div>
    <div class="row"><button class="btn primary" id="setDone">닫기</button></div>`, m => {
    const sync = () => m.querySelectorAll('.sub-sw input').forEach(i => { i.disabled = !S.auto; });
    sync();
    if (!drive) m.querySelector('[data-k="exportToDrive"]').disabled = true;
    m.addEventListener('change', e => {
      const t = e.target;
      if (t.dataset.k) S[t.dataset.k] = t.checked;
      if (t.name === 'autoExport') S.autoExport = t.value;
      if (t.name === 'theme') S.theme = t.value;
      saveSettings(); sync();
    });
    m.querySelector('#setDone').onclick = closeModal;
  });
}

// ---------- 저장 위치 · 계정 ----------
function storageModal() {
  const drive = store?.kind === 'drive';
  modal(`<h2>저장 위치</h2>
    ${drive ? `<p>지금 노트는 <b>구글 드라이브</b>의 “에스 노트” 폴더에 저장돼요. 이 앱은 자기가 만든 파일만 볼 수 있어요.</p>
      <button class="choice" id="stOpen"><b>드라이브에서 폴더 열기</b><small>그림·첨부 파일·내보낸 파일을 드라이브에서 직접 볼 수 있어요.</small></button>
      <button class="choice" id="stMigrate"><b>이 기기에 있던 노트를 드라이브로 옮기기</b><small>체험 모드로 쓴 노트가 있으면 드라이브로 복사해요.</small></button>
      <button class="choice" id="stOut"><b>로그아웃</b><small>노트는 드라이브에 그대로 남아요.</small></button>`
    : `<p>지금 노트는 <b>이 기기 브라우저 안</b>에만 저장돼요 (체험 모드). 다른 기기에서는 보이지 않고, 브라우저 기록을 지우면 사라질 수 있어요.</p>
      <button class="choice" id="stDrive" ${CLIENT_ID ? '' : 'disabled'}><b>구글 드라이브로 바꾸기</b><small>${CLIENT_ID ? '로그인하면 이 기기의 노트를 드라이브로 옮길 수 있어요.' : '먼저 config.js에 구글 클라이언트 ID를 넣어야 해요. README.md를 보세요.'}</small></button>`}
    <div class="row"><button class="btn" id="stClose">닫기</button></div>`, m => {
    m.querySelector('#stClose').onclick = closeModal;
    m.querySelector('#stOpen')?.addEventListener('click', () => window.open(`https://drive.google.com/drive/folders/${store.root}`, '_blank', 'noopener'));
    m.querySelector('#stOut')?.addEventListener('click', async () => { await flush().catch(() => {}); store.signOut(); localStorage.removeItem('dn-mode'); location.reload(); });
    m.querySelector('#stMigrate')?.addEventListener('click', () => { closeModal(); migrate(); });
    m.querySelector('#stDrive')?.addEventListener('click', async () => { await flush().catch(() => {}); localStorage.setItem('dn-mode', 'drive'); closeModal(); startDrive(true, true); });
  });
}
// 이 기기(체험) 노트 → 드라이브
async function migrate() {
  const local = new LocalStore(); await local.init();
  const lix = await local.loadIndex();
  const pages = lix ? lix.notebooks.flatMap(n => n.sections.flatMap(s => s.pages)) : [];
  if (!pages.length) return toast('이 기기에 옮길 노트가 없어요.');
  if (!(await confirmBox('이 기기 노트를 드라이브로 옮기기', `페이지 ${pages.length}개를 드라이브로 복사해요. 이 기기의 노트는 지워지지 않아요.`, '옮기기', false))) return;
  try {
    let n = 0;
    for (const nb of lix.notebooks) {
      const copy = { id: uid(), name: nb.name + ' (이 기기에서 옮김)', sections: [] };
      for (const sec of nb.sections) {
        const s2 = { ...sec, id: uid(), pages: [] };
        for (const p of sec.pages) {
          toast(`옮기는 중… ${++n} / ${pages.length}`, [], 0);
          let html = await local.loadPage(p);
          const doc = parseHTML(html);
          for (const el of doc.querySelectorAll('[data-fid]')) {
            const b = await local.getFile(el.dataset.fid);
            if (b) el.dataset.fid = await store.putFile(b, el.dataset.name || 'image');
          }
          const p2 = { ...p, id: uid(), fid: undefined, exp: undefined };
          await store.savePage(p2, doc.innerHTML);
          s2.pages.push(p2);
        }
        copy.sections.push(s2);
      }
      ix.notebooks.push(copy);
    }
    scheduleIndexSave(); await flush();
    renderAll(); toast('다 옮겼어요. 전자 필기장 목록에서 “(이 기기에서 옮김)”을 찾으세요.', [], 8000);
  } catch (e) { handleErr(e); }
}

// ---------- 시작 ----------
async function load() {
  status('', '불러오는 중…');
  await store.init();
  ix = await store.loadIndex();
  if (!ix) { ix = newIndex(); await store.saveIndex(ix); }
  let last = {}; try { last = JSON.parse(localStorage.getItem('dn-last') || '{}'); } catch {}
  reselect(last.nb, last.sec, last.page);
  if (!cur.page && cur.sec) cur.page = cur.sec.pages[0] || null;
  renderAll(); applySettings(); setView(phone() && !cur.page ? 'sections' : 'page');
  await openPage(cur.page);
  $('account').textContent = store.kind === 'drive' ? '☁ 구글 드라이브' : '💾 이 기기 (체험)';
  status('ok');
}
async function startLocal() {
  store = new LocalStore();
  try { await load(); } catch (e) { handleErr(e); }
}
async function startDrive(interactive, offerMigrate) {
  store = new DriveStore(CLIENT_ID);
  if (!store.restore()) {
    if (!interactive) return driveGate();
    try { await store.signIn(); } catch (e) { toast(e.message, [], 8000); return driveGate(); }
  }
  try {
    await load();
    if (offerMigrate) migrate();
  } catch (e) {
    if (e.auth) return driveGate();
    handleErr(e, () => location.reload());
  }
}
// 새로 열 때마다 구글 로그인 버튼 한 번 (팝업은 버튼을 눌러야 열려서)
function driveGate() {
  showEmpty(`<h2>구글 드라이브에 연결해 주세요</h2><p>노트는 내 구글 드라이브의 “에스 노트” 폴더에 저장돼요.</p><div class="btns"><button class="btn primary" id="gateGo">구글로 로그인</button><button class="btn" id="gateLocal">이 기기에서 체험하기</button></div>`);
  $('account').textContent = '로그인 필요'; status('', '');
  $('gateGo').onclick = () => startDrive(true);
  $('gateLocal').onclick = () => { localStorage.setItem('dn-mode', 'local'); startLocal(); };
}
function welcome() {
  $('modal').dataset.locked = '1';
  modal(`<div class="welcome"><span class="logo">S</span><h2>에스 노트에 오신 걸 환영해요</h2><p>노트를 어디에 저장할까요? 나중에 설정에서 바꿀 수 있어요.</p>
    <button class="choice" id="wDrive" ${CLIENT_ID ? '' : 'disabled'}><b>☁ 구글 드라이브에 저장 (추천)</b><small>${CLIENT_ID ? '휴대폰·컴퓨터 어디서나 같은 노트를 볼 수 있어요. 내 드라이브의 “에스 노트” 폴더만 사용해요.' : '아직 구글 연결 설정이 안 돼 있어요. README.md의 순서대로 config.js에 클라이언트 ID를 넣으면 쓸 수 있어요.'}</small></button>
    <button class="choice" id="wLocal"><b>💾 이 기기에만 저장 (체험)</b><small>로그인 없이 바로 써 볼 수 있어요. 다른 기기와 연동되지 않아요.</small></button></div>`, m => {
    m.querySelector('#wDrive').onclick = () => { delete $('modal').dataset.locked; closeModal(); localStorage.setItem('dn-mode', 'drive'); startDrive(true); };
    m.querySelector('#wLocal').onclick = () => { delete $('modal').dataset.locked; closeModal(); localStorage.setItem('dn-mode', 'local'); startLocal(); };
  });
}

// ---------- 원노트에서 가져오기 ----------
const MS_ID = (window.NOTE_CONFIG && window.NOTE_CONFIG.msClientId || '').trim();
$('importOneNote').onclick = () => {
  if (!store || !ix) return toast('먼저 저장 위치를 고르세요.');
  modal(`<h2>원노트에서 가져오기</h2><p>원노트의 원래 노트는 바뀌지 않아요. 두 가지 방법 중에 고르세요.</p>
    <button class="choice" id="onAcct" ${MS_ID ? '' : 'disabled'}><b>☁ 마이크로소프트 계정 연결</b><small>${MS_ID ? '개인 계정 원노트를 전자 필기장째 가져와요. 그림·첨부 파일·할 일·하위 페이지까지 옮겨져요. 학교·회사 계정은 막혀 있을 수 있어요.' : '아직 마이크로소프트 연결 설정(config.js의 msClientId)이 안 돼 있어요.'}</small></button>
    <button class="choice" id="onWord"><b>📄 Word 파일(.docx)로 가져오기</b><small>원노트에서 섹션마다 <b>파일 → 내보내기 → 섹션 → Word 문서</b>로 저장한 파일들을 골라요. 학교 계정 노트도 돼요.</small></button>
    <p style="font-size:12.5px">펜으로 쓴 손글씨는 원노트가 내보내 주지 않아서 빠질 수 있어요.</p>
    <div class="row"><button class="btn" id="onNo">취소</button></div>`, m => {
    m.querySelector('#onNo').onclick = closeModal;
    m.querySelector('#onWord').onclick = wordImport;
    m.querySelector('#onAcct').onclick = async () => {
      try { if (!msSignedIn()) await msLogin(MS_ID); } catch (e) { return toast(e.message, [], 8000); }
      pickNotebooks();
    };
  });
};
// Word 파일 하나 = 섹션 하나, 원노트 페이지마다 페이지 하나
function wordImport() {
  modal(`<h2>Word 파일로 가져오기</h2>
    <p>원노트에서 내보낸 Word 파일(.docx)을 한 번에 여러 개 골라도 돼요. 파일 하나가 섹션 하나가 되고, 원노트 페이지마다 페이지가 하나씩 생겨요.</p>
    <label style="display:block;margin-top:12px;font-weight:600">넣을 전자 필기장 이름<input type="text" id="wNb" value="원노트에서 가져옴"></label>
    <div class="row"><button class="btn" id="wNo">취소</button><button class="btn primary" id="wPick">Word 파일 고르기</button></div>`, m => {
    m.querySelector('#wNo').onclick = closeModal;
    m.querySelector('#wPick').onclick = () => {
      const name = m.querySelector('#wNb').value.trim() || '원노트에서 가져옴';
      const inp = document.createElement('input'); inp.type = 'file'; inp.accept = '.docx'; inp.multiple = true;
      inp.onchange = () => runWordImport([...inp.files].filter(f => /\.docx$/i.test(f.name)), name);
      inp.click();
    };
  });
}
async function runWordImport(files, nbName) {
  if (!files.length) return toast('.docx 파일을 골라 주세요.');
  const { docxToPages } = await import('./docx-import.js');
  $('modal').dataset.locked = '1';
  modal(`<h2>Word 파일에서 가져오는 중</h2><div class="progress"><i id="onBar"></i></div><div class="prog-text" id="onText"></div><p style="font-size:12.5px">이 창을 닫지 마세요.</p>`);
  const show = (k, text) => { $('onBar').style.width = Math.round(k / files.length * 100) + '%'; $('onText').textContent = text; };
  let nb = ix.notebooks.find(n => n.name === nbName);
  if (!nb) { nb = { id: uid(), name: nbName, sections: [] }; ix.notebooks.push(nb); }
  let ok = 0, imgs = 0; const failed = [];
  for (const [i, f] of files.entries()) {
    show(i, `${i + 1} / ${files.length} · ${f.name} 읽는 중…`);
    try {
      const pages = await docxToPages(f, (b, n) => store.putFile(b, n), () => show(i, `${i + 1} / ${files.length} · ${f.name} · 그림 ${++imgs}개 올리는 중…`));
      const sec = { id: uid(), name: f.name.replace(/\.docx$/i, ''), color: SECTION_COLORS[nb.sections.length % SECTION_COLORS.length], pages: [] };
      for (const pg of pages) {
        const when = pg.created || Date.now(), p = { id: uid(), title: pg.title, created: when, updated: when, level: 0, snip: plain(pg.html).slice(0, 80) };
        await store.savePage(p, pg.html);
        sec.pages.push(p); ok++;
      }
      nb.sections.push(sec);
      scheduleIndexSave(); await saveIndexNow();
    } catch (e) {
      console.warn('Word 가져오기 실패', f.name, e);
      if (e.auth) { failed.push(...files.slice(i).map(x => x.name)); break; }
      failed.push(f.name);
    }
  }
  scheduleIndexSave(); await saveIndexNow();
  delete $('modal').dataset.locked; closeModal();
  if (cur.page) leavePage(cur.page);
  cur.nb = nb; cur.sec = nb.sections[0] || null; cur.page = null; renderAll(); openPage(cur.sec?.pages[0] || null);
  toast(`페이지 ${ok}개를 가져와서 “${nb.name}” 전자 필기장에 넣었어요.${failed.length ? ` 가져오지 못한 파일: ${failed.join(', ')}` : ''}`, [], 12000);
  if (failed.length && !ok) handleErr(new Error('Word 파일을 읽지 못했어요. 원노트에서 Word 문서(.docx)로 내보낸 파일인지 확인해 주세요.'));
}
async function pickNotebooks() {
  modal(`<h2>원노트 전자 필기장을 불러오는 중…</h2><div class="progress"><i style="width:35%"></i></div>`);
  let nbs;
  try { nbs = await listNotebooks(); } catch (e) { closeModal(); return handleErr(e); }
  if (!nbs.length) { closeModal(); return toast('원노트에 전자 필기장이 없어요.'); }
  const done = new Set(allPages().map(x => x.p.onId).filter(Boolean));
  modal(`<h2>가져올 전자 필기장을 고르세요</h2><p>${done.size ? `이미 가져온 페이지 ${done.size}개는 건너뛰고 나머지만 가져와요.` : '고른 전자 필기장의 모든 섹션과 페이지를 가져와요.'}</p>
    <div class="on-tools"><span id="onCount"></span><button class="btn" id="onAll">전체 선택</button><button class="btn" id="onNone">전체 해제</button></div>
    <div class="on-list">${nbs.map((nb, i) => `<label><input type="checkbox" data-i="${i}" checked><span><b>${esc(nb.name)}</b><small>섹션 ${nb.sections.length}개 · ${esc(nb.sections.slice(0, 4).map(s => s.name).join(', '))}${nb.sections.length > 4 ? ' …' : ''}</small></span></label>`).join('')}</div>
    <div class="row"><button class="btn" id="onNo">취소</button><button class="btn primary" id="onGo">가져오기</button></div>`, m => {
    const boxes = [...m.querySelectorAll('.on-list input')];
    const count = () => { const n = boxes.filter(b => b.checked).length; m.querySelector('#onCount').textContent = `${nbs.length}개 중 ${n}개 선택`; m.querySelector('#onGo').disabled = !n; };
    m.querySelector('#onAll').onclick = () => { boxes.forEach(b => { b.checked = true; }); count(); };
    m.querySelector('#onNone').onclick = () => { boxes.forEach(b => { b.checked = false; }); count(); };
    m.querySelector('.on-list').onchange = count;
    count();
    m.querySelector('#onNo').onclick = closeModal;
    m.querySelector('#onGo').onclick = () => {
      const sel = boxes.filter(b => b.checked).map(b => nbs[+b.dataset.i]);
      if (sel.length) runImport(sel, done);
    };
  });
}
async function runImport(nbs, done) {
  let cancel = false, authStop = false;
  $('modal').dataset.locked = '1';
  modal(`<h2>원노트에서 가져오는 중</h2><div class="progress"><i id="onBar"></i></div><div class="prog-text" id="onText">페이지 목록을 세는 중…</div>
    <p style="font-size:12.5px">이 창을 닫지 마세요. 중간에 멈춰도 다시 [원노트에서 가져오기]를 누르면 이어서 가져와요.</p>
    <div class="row"><button class="btn" id="onStop">멈추기</button></div>`, m => {
    m.querySelector('#onStop').onclick = e => { cancel = true; e.target.disabled = true; $('onText').textContent = '지금 페이지까지만 가져오고 멈출게요…'; };
  });
  const show = (k, n, text) => { $('onBar').style.width = (n ? Math.round(k / n * 100) : 3) + '%'; $('onText').textContent = text; };
  const finish = msg => { delete $('modal').dataset.locked; closeModal(); renderAll(); if (cur.sec) renderPages(); toast(msg, [], 12000); };
  // 1) 섹션마다 페이지 목록
  const plan = [];
  try {
    for (const nb of nbs) for (const s of nb.sections) {
      if (cancel) break;
      show(0, 0, `페이지 목록을 세는 중… ${nb.name} › ${s.name}`);
      plan.push({ nb, s, pages: await listPages(s.id) });
    }
  } catch (e) { finish('원노트 목록을 읽지 못했어요.'); return handleErr(e); }
  const total = plan.reduce((a, x) => a + x.pages.filter(p => !done.has(p.id)).length, 0);
  // 2) 페이지 하나씩: 내용 → 그림·첨부를 내 저장소로 → 페이지 저장
  let k = 0, ok = 0, fail = 0;
  for (const { nb, s, pages } of plan) {
    if (cancel) break;
    let myNb = ix.notebooks.find(x => x.onId === nb.id);
    if (!myNb) { myNb = { id: uid(), name: nb.name, onId: nb.id, sections: [] }; ix.notebooks.push(myNb); }
    let mySec = myNb.sections.find(x => x.onId === s.id);
    if (!mySec) { mySec = { id: uid(), name: s.name, onId: s.id, color: SECTION_COLORS[myNb.sections.length % SECTION_COLORS.length], pages: [] }; myNb.sections.push(mySec); }
    for (const op of pages) {
      if (cancel) break;
      if (done.has(op.id)) continue;
      show(k, total, `${++k} / ${total} · ${nb.name} › ${s.name} › ${op.title || '제목 없음'}`);
      try {
        const html = await convertPage(await pageContent(op.id), (blob, name) => store.putFile(blob, name));
        const p = { id: uid(), onId: op.id, title: op.title || '', created: Date.parse(op.createdDateTime) || Date.now(), updated: Date.parse(op.lastModifiedDateTime) || Date.now(), level: op.level ? 1 : 0, snip: plain(html).slice(0, 80) };
        await store.savePage(p, html);
        mySec.pages.push(p); done.add(op.id); ok++;
        if (ok % 5 === 0) { scheduleIndexSave(); await saveIndexNow(); }
      } catch (e) {
        console.warn('원노트 페이지 가져오기 실패', op.title, e);
        if (e.auth) { authStop = cancel = true; break; } // 구글 로그인이 끝남 → 멈추고 다시 로그인
        fail++;
      }
    }
  }
  scheduleIndexSave(); await saveIndexNow();
  if (authStop) { finish(`${ok}개를 가져오고 멈췄어요. 구글에 다시 로그인한 뒤 [원노트에서 가져오기]를 누르면 이어서 가져와요.`); return handleErr({ auth: true }); }
  finish(cancel ? `${ok}개를 가져오고 멈췄어요. 다시 [원노트에서 가져오기]를 누르면 이어서 가져와요.`
    : `원노트 페이지 ${ok}개를 가져왔어요.${fail ? ` ${fail}개는 가져오지 못했어요. 다시 [원노트에서 가져오기]를 누르면 그것만 다시 시도해요.` : ''} 전자 필기장 목록에서 확인하세요.`);
}

// ---------- 프로그램으로 설치 (PWA) ----------
let installEvt = null;
const installed = () => matchMedia('(display-mode: standalone)').matches || navigator.standalone;
// 새 버전을 올리면 설치한 앱도 알아서 바뀜: 창으로 돌아올 때·30분마다 확인 → 바뀌면 저장 후 새로고침
if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost')) {
  const hadController = !!navigator.serviceWorker.controller;
  navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).then(reg => {
    const check = () => reg.update().catch(() => {});
    addEventListener('focus', check); setInterval(check, 30 * 60e3);
  }).catch(() => {});
  let reloading = false;
  navigator.serviceWorker.addEventListener('controllerchange', async () => {
    if (!hadController || reloading) return;
    reloading = true;
    const go = async () => { await flush().catch(() => {}); location.reload(); };
    if (!dirty && $('modal').hidden && document.activeElement !== $('editor')) go();
    else toast('새 버전이 나왔어요.', [{ label: '지금 바꾸기', run: go }], 0);
  });
}
addEventListener('beforeinstallprompt', e => { e.preventDefault(); installEvt = e; });
addEventListener('appinstalled', () => { installEvt = null; $('installApp').hidden = true; toast('설치했어요. 바탕화면이나 시작 메뉴의 “에스 노트”로 열 수 있어요.', [], 6000); });
if (installed()) $('installApp').hidden = true;
$('installApp').onclick = async () => {
  if (installEvt) { installEvt.prompt(); const r = await installEvt.userChoice; if (r.outcome === 'accepted') installEvt = null; return; }
  modal(`<h2>프로그램으로 설치하기</h2>
    <div class="set"><h3>💻 컴퓨터 (크롬·엣지)</h3><p>주소창 오른쪽 끝의 <b>설치 아이콘(⊕ 또는 모니터 모양)</b>을 누르거나, 메뉴 <b>⋮ → 앱 설치</b>(엣지는 <b>⋯ → 앱 → 이 사이트를 앱으로 설치</b>)를 누르세요. 바탕화면과 시작 메뉴에 “에스 노트”가 생겨요.</p></div>
    <div class="set"><h3>📱 안드로이드 (크롬)</h3><p>오른쪽 위 <b>⋮ → 홈 화면에 추가 → 설치</b>를 누르세요.</p></div>
    <div class="set"><h3>📱 아이폰·아이패드 (사파리)</h3><p>아래 <b>공유 버튼(□↑) → 홈 화면에 추가</b>를 누르세요.</p></div>
    <p>이미 설치했다면 바탕화면·시작 메뉴·홈 화면에서 “에스 노트”를 찾으세요.</p>
    <div class="row"><button class="btn primary" id="insOk">확인</button></div>`, m => { m.querySelector('#insOk').onclick = closeModal; });
};

setupEditor($('editor'), {
  onChange: markDirty,
  settings: () => S,
  uploadImage: f => upload(f, true),
  uploadFile: f => upload(f, false),
  openFile,
  onAuto: (kind, undo) => toast(kind === 'ol' ? '번호 목록으로 바꿨어요.' : kind === 'ul' ? '글머리 기호로 바꿨어요.' : '할 일 목록으로 바꿨어요.',
    [{ label: '되돌리기', run: undo }, { label: '자동 목록 끄기', run: () => { S.auto = false; saveSettings(); undo(); toast('자동 목록을 껐어요. 설정이나 홈 탭에서 다시 켤 수 있어요.'); } }]),
});
applySettings();
matchMedia('(max-width: 599px)').addEventListener('change', () => { setView($('shell').dataset.view); applyNav(); });
const mode = localStorage.getItem('dn-mode');
if (mode === 'drive' && CLIENT_ID) startDrive(false);
else if (mode === 'local') startLocal();
else welcome();
