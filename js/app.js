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
  updateSpellAttr();
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

const openNbs_ = () => ix.notebooks.filter(n => !n.closed); // 닫지 않은 전자 필기장
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
// 지우기는 "삭제하겠습니다"를 그대로 입력해야 됨
const DEL_WORD = '삭제하겠습니다';
const confirmDelete = (title, body) => new Promise(res => {
  modal(`<h2>${esc(title)}</h2><p>${body}</p><p>지우려면 아래 칸에 <b>${DEL_WORD}</b>를 그대로 입력하세요.</p><input type="text" id="delIn" placeholder="${DEL_WORD}" autocomplete="off" autofocus><div class="row"><button class="btn" id="cNo">취소</button><button class="btn danger" id="cYes" disabled>삭제</button></div>`, m => {
    const inp = m.querySelector('#delIn'), yes = m.querySelector('#cYes');
    const ok = () => inp.value.replace(/\s/g, '') === DEL_WORD;
    inp.oninput = () => { yes.disabled = !ok(); };
    inp.onkeydown = e => { if (e.key === 'Enter' && ok()) yes.click(); };
    yes.onclick = () => { if (!ok()) return; closeModal(); res(true); };
    m.querySelector('#cNo').onclick = () => { closeModal(); res(false); };
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
addEventListener('online', async () => {
  if (dirty) savePageNow(); if (indexPending) saveIndexNow();
  if (store?.kind === 'drive' && ix) {
    store.offline = false; status('ok');
    try { const n = await store.flushPending(id => findPage(id)?.p); if (n) { scheduleIndexSave(); toast(`인터넷이 다시 연결돼서 페이지 ${n}개를 드라이브에 올렸어요.`); } } catch {}
  }
});
addEventListener('offline', () => { if (store?.kind === 'drive') status('', '⚠ 오프라인 · 이 기기에 저장 중'); });

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
  const ed = $('editor');
  if (!ed.querySelector('img, .sel')) return ed.innerHTML; // 그림이 없으면 그대로
  const c = ed.cloneNode(true);
  c.querySelectorAll('img[data-fid]').forEach(i => i.removeAttribute('src'));
  c.querySelectorAll('.sel').forEach(i => i.classList.remove('sel'));
  return c.innerHTML;
}
// 브라우저 빨간 밑줄은 글이 아주 길면 꺼서 입력이 느려지지 않게
function updateSpellAttr() { $('editor').setAttribute('spellcheck', S.browserSpell && $('editor').textContent.length < 30000 ? 'true' : 'false'); }
function plain(html) { return toLines(parseHTML(html)).lines.join(' ').replace(/\s+/g, ' ').trim(); }
const pageSnip = () => $('editor').innerText.slice(0, 300).replace(/\s+/g, ' ').trim().slice(0, 80);
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
  draftTimer = setTimeout(() => { try { const html = serialize(); if (html.length < 1.5e6) localStorage.setItem('dn-draft-' + p.id, JSON.stringify({ t: Date.now(), title: $('pageTitle').value, html })); } catch {} }, 1500);
}
function savePageNow() {
  clearTimeout(saveTimer);
  if (!dirty || !cur.page) return saving;
  const p = cur.page, html = serialize(), at = Date.now();
  dirty = false;
  p.title = $('pageTitle').value.trim(); p.updated = at; p.snip = pageSnip(); updateSpellAttr();
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
  if (!store || store.kind !== 'drive' || !navigator.onLine || Date.now() - lastCheck < 5000 || indexPending) return;
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
// 전자 필기장 칸 찾기: 전자 필기장 이름이나 섹션 이름만 찾음 (페이지 내용은 X). 띄어쓰기·대소문자 무시
let nbQ = '';
const normQ = t => String(t || '').replace(/\s+/g, '').toLowerCase();
function markQ(name) {
  if (!nbQ) return esc(name);
  // 띄어쓰기를 건너뛰며 맞는 곳을 찾아 굵게 표시
  const chars = [...name], idx = [];
  chars.forEach((c, k) => { if (!/\s/.test(c)) idx.push(k); });
  const flat = idx.map(k => chars[k]).join('').toLowerCase(), at = flat.indexOf(nbQ);
  if (at < 0) return esc(name);
  const a = idx[at], b = idx[at + nbQ.length - 1] + 1;
  return esc(chars.slice(0, a).join('')) + '<mark>' + esc(chars.slice(a, b).join('')) + '</mark>' + esc(chars.slice(b).join(''));
}
function renderNotebook() {
  $('nbName').textContent = cur.nb ? cur.nb.name : '전자 필기장 없음';
  $('secHead').textContent = cur.nb ? cur.nb.name : '';
  if (cur.nb && !curCollapsed) openNbs.add(cur.nb.id);
  if (nbQ) {
    // 찾는 중: 이름이 맞는 전자 필기장은 섹션을 모두, 아니면 이름이 맞는 섹션만 펼쳐서 보여 줌
    const html = ix.notebooks.map((n, i) => {
      if (n.closed) return '';
      const hit = normQ(n.name).includes(nbQ), secs = hit ? n.sections : n.sections.filter(s => normQ(s.name).includes(nbQ));
      if (!hit && !secs.length) return '';
      return `<li data-id="${n.id}" class="nb-item open${n === cur.nb ? ' on' : ''}"><button class="chev" data-chev aria-label="접기">›</button><span class="nbi" style="background:${NB_COLORS[i % NB_COLORS.length]}"></span><span class="name">${markQ(n.name)}</span><button class="more" data-more="nb" aria-label="전자 필기장 메뉴">⋯</button></li>` +
        secs.map(s => `<li data-sec="${s.id}" data-nb="${n.id}" class="tree-sec${s === cur.sec ? ' on' : ''}"><span class="bar" style="background:${s.color}"></span><span class="name">${markQ(s.name)}</span><button class="more" data-more="sec" aria-label="섹션 메뉴">⋯</button></li>`).join('');
    }).join('');
    $('nbList').innerHTML = html || '<li class="none"><span class="name">찾는 전자 필기장·섹션이 없어요.</span></li>';
    return;
  }
  $('nbList').innerHTML = ix.notebooks.map((n, i) => {
    if (n.closed) return '';
    const open = openNbs.has(n.id);
    return `<li data-id="${n.id}" class="nb-item${n === cur.nb ? ' on' : ''}${open ? ' open' : ''}"><button class="chev" data-chev aria-label="${open ? '접기' : '펼치기'}">›</button><span class="nbi" style="background:${NB_COLORS[i % NB_COLORS.length]}"></span><span class="name">${esc(n.name)}</span><button class="more" data-more="nb" aria-label="전자 필기장 메뉴">⋯</button></li>` +
      (open ? n.sections.map(s => `<li data-sec="${s.id}" data-nb="${n.id}" class="tree-sec${s === cur.sec ? ' on' : ''}"><span class="bar" style="background:${s.color}"></span><span class="name">${esc(s.name)}</span><button class="more" data-more="sec" aria-label="섹션 메뉴">⋯</button></li>`).join('') +
        `<li class="tree-sec tree-add" data-addsec="${n.id}"><span class="bar"></span><span class="name">＋ 섹션 추가</span></li>` : '');
  }).join('');
}
function setNbFilter(v) {
  $('nbFilter').value = v; nbQ = normQ(v); $('nbFilterX').hidden = !v;
  if (ix) renderNotebook();
}
$('nbFilter').addEventListener('input', () => setNbFilter($('nbFilter').value));
$('nbFilter').addEventListener('keydown', e => { if (e.key === 'Escape') setNbFilter(''); });
$('nbFilterX').onclick = () => { setNbFilter(''); $('nbFilter').focus(); };
const DRAG_ATTR = matchMedia('(pointer: fine)').matches ? ' draggable="true"' : '';
function renderSections() {
  const ul = $('sectionList');
  ul.innerHTML = (cur.nb?.sections || []).map(s => `<li data-id="${s.id}" class="${s === cur.sec ? 'on' : ''}"${DRAG_ATTR} style="border-left-color:${s.color}"><span class="dot" style="background:${s.color}"></span><span class="name">${esc(s.name)}</span><button class="more" data-more="sec" aria-label="섹션 메뉴">⋯</button></li>`).join('');
  renderNotebook();
}
const fmtDate = t => t ? new Date(t).toLocaleDateString('ko-KR', { year: 'numeric', month: 'long', day: 'numeric' }) : '';
const fmtFull = t => t ? new Date(t).toLocaleString('ko-KR', { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long', hour: 'numeric', minute: '2-digit' }) : '';
function renderPages() {
  const ul = $('pageList'), head = $('pagesHead');
  if (searchMode) {
    const hits = searchMode.ids.map(findPage).filter(Boolean);
    head.innerHTML = searchMode.loading ? `<span>✨ AI가 찾는 중… “${esc(searchMode.q)}”</span><button id="endSearch">닫기</button>`
      : `<span>${searchMode.ai ? '✨ AI 검색' : '검색'} 결과 ${hits.length}개 · “${esc(searchMode.q)}”</span><button id="endSearch">닫기</button>`;
    $('endSearch').onclick = endSearch;
    ul.innerHTML = hits.map(({ sec, p }) => `<li data-id="${p.id}" class="${p === cur.page ? 'on' : ''}"><div class="row"><span class="name">${esc(p.title || '제목 없음')}</span></div><span class="snip">${searchMode.why?.[p.id] ? `<span class="why">✨ ${esc(searchMode.why[p.id])}</span> · ` : ''}${esc(sec.name)} · ${esc(p.snip || '')}</span></li>`).join('') || (searchMode.loading ? '' : `<li class="none"><span class="snip">${searchMode.ai ? 'AI도 맞는 페이지를 찾지 못했어요.' : '찾는 내용이 없어요. 뜻으로 찾고 싶으면 검색창 오른쪽 [일반]을 눌러 ✨ AI 검색으로 바꿔 보세요.'}</span></li>`);
    $('addPage').hidden = true;
    return;
  }
  $('addPage').hidden = !cur.sec;
  head.innerHTML = cur.sec ? `<span style="color:${cur.sec.color}">■</span><span>${esc(cur.sec.name)}</span>` : '';
  ul.innerHTML = (cur.sec?.pages || []).map(p => `<li data-id="${p.id}" class="${p === cur.page ? 'on' : ''}${p.level ? ' sub' : ''}"${DRAG_ATTR}><div class="row"><span class="name">${esc(p.title || '제목 없음')}</span><button class="more" data-more="page" aria-label="페이지 메뉴">⋯</button></div><span class="date">${fmtDate(p.updated || p.created)}</span></li>`).join('');
}
function showEmpty(html) { $('page').hidden = true; $('empty').hidden = false; $('empty').innerHTML = html; }
function setView(v) {
  $('shell').dataset.view = v;
  $('navBack').hidden = !phone() || v === 'sections';
}
$('navBack').onclick = () => setView($('shell').dataset.view === 'page' ? 'pages' : 'sections');

// 목록 접기·펴기 (☰ / 전체 페이지 보기). 글을 눌러도 저절로 접히지 않음
let navOff = false;
try { navOff = localStorage.getItem('dn-nav-off') === '1'; } catch {}
function applyNav() {
  const off = navOff && !phone();
  $('shell').classList.toggle('nav-off', off);
  $('navToggle').title = off ? '목록 펴기' : '목록 접기 (전체 페이지 보기)';
  $('navToggle').classList.toggle('on', !off);
}
function setNav(off) {
  // 목록을 접고 펴면 글 너비가 바뀌어 줄이 다시 나뉨 → 보고 있던 줄이 화면 같은 자리에 오도록 맞춤
  const cv = $('canvas'), top = cv.getBoundingClientRect().top;
  const anchor = [...$('editor').children].find(el => el.getBoundingClientRect().bottom > top + 1);
  const before = anchor ? anchor.getBoundingClientRect().top : 0;
  navOff = off; try { localStorage.setItem('dn-nav-off', off ? '1' : '0'); } catch {} applyNav();
  if (anchor && cv.scrollTop > 0) cv.scrollTop += anchor.getBoundingClientRect().top - before;
}
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
// (예전: 폴드에서 글을 누르면 목록이 저절로 접혔음 → 화면이 바뀌고 글 위치가 움직여 어지러워서 없앰. ☰로만 접고 폄)
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
  applyLines(p); applyPageWidth(p);
  $('pageTitle').value = p.title || '';
  $('pageDate').textContent = fmtFull(p.created);
  // 바로 열리면 깜빡임 없이 바꾸고, 0.12초 넘게 걸릴 때만 비우고 "불러오는 중" 표시
  const waitT = setTimeout(() => { if (cur.page === p) { $('editor').innerHTML = ''; if (!silent) status('', '불러오는 중…'); } }, 120);
  let html = '';
  // 옛 보관본을 먼저 보여 주고, 새 내용이 오면 (아직 안 고쳤을 때) 바꿔 끼움
  const onFresh = h => {
    if (cur.page !== p || dirty || editedSinceOpen) return;
    const st = $('canvas').scrollTop; $('editor').innerHTML = sanitize(h) || '<p><br></p>'; hydrate(); updateSpellAttr(); $('canvas').scrollTop = st;
  };
  try { html = await store.loadPage(p, onFresh); }
  catch (e) { clearTimeout(waitT); handleErr(e, () => openPage(p)); return; }
  clearTimeout(waitT);
  if (cur.page !== p) return;
  // 저장되지 않은 임시 내용이 있으면 되살리기 제안
  let draft = null; try { draft = JSON.parse(localStorage.getItem('dn-draft-' + p.id) || 'null'); } catch {}
  $('editor').innerHTML = sanitize(html) || '<p><br></p>';
  updateSpellAttr();
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
  prefetchSection(sec);
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
    <button data-a="move">다른 전자 필기장으로 옮기기</button><button data-a="dup">섹션 복사본 만들기</button><button data-a="copyTo">다른 전자 필기장에 복사</button><hr><button data-a="del" class="danger">섹션 삭제</button>`, at);
  c.onclick = async e => {
    const b = e.target.closest('button'); if (!b) return; c.hidden = true;
    if (b.dataset.color) { sec.color = b.dataset.color; }
    else if (b.dataset.a === 'rename') { const n = await ask('섹션 이름 바꾸기', '섹션 이름', sec.name); if (!n) return; sec.name = n; }
    else if (b.dataset.a === 'up' || b.dataset.a === 'down') { const j = i + (b.dataset.a === 'up' ? -1 : 1); cur.nb.sections.splice(i, 1); cur.nb.sections.splice(j, 0, sec); }
    else if (b.dataset.a === 'move') { const to = await pickNotebook(cur.nb); if (!to) return; cur.nb.sections.splice(cur.nb.sections.indexOf(sec), 1); to.sections.push(sec); if (cur.sec === sec) { cur.nb = to; renderNotebook(); } }
    else if (b.dataset.a === 'dup') { await copySection(sec, cur.nb); return; }
    else if (b.dataset.a === 'copyTo') { const to = await pickNotebook(cur.nb, '복사'); if (!to) return; await copySection(sec, to); return; }
    else if (b.dataset.a === 'del') {
      if (!(await confirmDelete(`“${sec.name}” 섹션 삭제`, `이 섹션의 페이지 ${sec.pages.length}개도 함께 지워져요.${trashNote()}`))) return;
      cur.nb.sections.splice(cur.nb.sections.indexOf(sec), 1);
      if (cur.sec === sec) { dirty = false; cur.page = null; cur.sec = cur.nb.sections[0] || null; openPage(cur.sec?.pages[0] || null); }
      trashPages(sec.pages);
    }
    scheduleIndexSave(); renderSections(); renderPages();
  };
}
function pageMenu(p, at) {
  const sec = cur.sec.pages.includes(p) ? cur.sec : findPage(p.id).sec, i = sec.pages.indexOf(p);
  const c = showCtx(`<button data-a="rename">이름 바꾸기</button><button data-a="sub">${p.level ? '하위 페이지 해제' : '하위 페이지로 만들기'}</button><button data-a="child">아래에 하위 페이지 추가</button>
    <button data-a="up" ${i ? '' : 'disabled'}>위로</button><button data-a="down" ${i < sec.pages.length - 1 ? '' : 'disabled'}>아래로</button>
    <button data-a="move">다른 섹션으로 옮기기</button><button data-a="copy">복사본 만들기</button><button data-a="copyTo">다른 섹션에 복사</button><hr><button data-a="del" class="danger">페이지 삭제</button>`, at);
  c.onclick = async e => {
    const b = e.target.closest('button'); if (!b) return; c.hidden = true;
    const a = b.dataset.a;
    if (a === 'rename') {
      const n = await ask('페이지 이름 바꾸기', '페이지 제목', p.title || ''); if (n === null) return;
      p.title = n;
      if (p === cur.page) { $('pageTitle').value = n; markDirty(); }
    }
    else if (a === 'sub') p.level = p.level ? 0 : 1;
    else if (a === 'child') { addPage(p, 1); return; }
    else if (a === 'up' || a === 'down') { sec.pages.splice(i, 1); sec.pages.splice(i + (a === 'up' ? -1 : 1), 0, p); }
    else if (a === 'move') {
      const to = await pickSection(sec); if (!to) return;
      if (p === cur.page) savePageNow();
      sec.pages.splice(sec.pages.indexOf(p), 1); p.level = 0; to.pages.push(p);
      if (p === cur.page) openPage(sec.pages[0] || null);
    }
    else if (a === 'copy') {
      const np = await copyPage(p, (p.title || '제목 없음') + ' (복사본)'); if (!np) return;
      sec.pages.splice(i + 1, 0, np);
    }
    else if (a === 'copyTo') {
      const to = await pickSection(sec, '복사'); if (!to) return;
      const np = await copyPage(p, p.title || ''); if (!np) return;
      np.level = 0; to.pages.push(np);
      toast(`“${p.title || '제목 없음'}”을(를) ${to.name} 섹션에 복사했어요.`);
    }
    else if (a === 'del') {
      if (!(await confirmDelete(`“${p.title || '제목 없음'}” 페이지 삭제`, `이 페이지를 지울까요?${trashNote()}`))) return;
      sec.pages.splice(sec.pages.indexOf(p), 1);
      if (p === cur.page) { dirty = false; editedSinceOpen = false; cur.page = null; openPage(sec.pages[Math.min(i, sec.pages.length - 1)] || null); }
      trashPages([p]);
    }
    scheduleIndexSave(); renderPages();
  };
}
// 페이지 복사: 내용을 새 페이지 파일로 저장 (그림·첨부 파일은 원래 파일을 같이 씀 — 페이지를 지워도 파일은 남음)
async function copyPage(p, title) {
  try {
    if (p === cur.page) await flush();
    const html = p === cur.page ? serialize() : await store.loadPage(p);
    const now = Date.now(), np = { id: uid(), title, created: now, updated: now, level: p.level || 0, snip: p.snip };
    await store.savePage(np, html || '');
    return np;
  } catch (er) { handleErr(er); return null; }
}
// 섹션 복사: 같은 전자 필기장이면 바로 아래에 "(복사본)", 다른 전자 필기장이면 맨 아래에 같은 이름으로
async function copySection(sec, toNb) {
  const same = toNb.sections.includes(sec);
  const ns = { ...sec, id: uid(), name: same ? sec.name + ' (복사본)' : sec.name, pages: [] };
  for (let k = 0; k < sec.pages.length; k++) {
    toast(`섹션 복사 중… ${k + 1} / ${sec.pages.length}`);
    const np = await copyPage(sec.pages[k], sec.pages[k].title || ''); if (!np) return;
    np.level = sec.pages[k].level || 0; ns.pages.push(np);
  }
  if (same) toNb.sections.splice(toNb.sections.indexOf(sec) + 1, 0, ns); else toNb.sections.push(ns);
  scheduleIndexSave(); renderSections(); renderPages();
  toast(`“${sec.name}” 섹션을 ${same ? '복사했어요' : toNb.name + '에 복사했어요'}. (페이지 ${ns.pages.length}개)`);
}
function pickSection(from, verb = '옮기') {
  return new Promise(res => {
    const opts = allSections().filter(x => x.sec !== from && !x.nb.closed);
    if (!opts.length) { toast('다른 섹션이 없어요. 먼저 섹션을 만드세요.'); return res(null); }
    modal(`<h2>어느 섹션으로 ${verb === '복사' ? '복사할' : '옮길'}까요?</h2>${opts.map((x, k) => `<button class="choice" data-k="${k}"><b><span style="color:${x.sec.color}">■</span> ${esc(x.sec.name)}</b><small>${esc(x.nb.name)}</small></button>`).join('')}<div class="row"><button class="btn" id="pNo">취소</button></div>`, m => {
      m.onclick = e => { const b = e.target.closest('[data-k]'); if (b) { closeModal(); res(opts[+b.dataset.k].sec); } else if (e.target.id === 'pNo') { closeModal(); res(null); } };
    });
  });
}
function pickNotebook(from, verb = '옮기') {
  return new Promise(res => {
    const opts = openNbs_().filter(n => n !== from);
    if (!opts.length) { toast('다른 전자 필기장이 없어요. 먼저 새 전자 필기장을 만드세요.'); return res(null); }
    modal(`<h2>어느 전자 필기장으로 ${verb === '복사' ? '복사할' : '옮길'}까요?</h2>${opts.map((n, k) => `<button class="choice" data-k="${k}"><b>${esc(n.name)}</b><small>섹션 ${n.sections.length}개</small></button>`).join('')}<div class="row"><button class="btn" id="pNo">취소</button></div>`, m => {
      m.onclick = e => { const b = e.target.closest('[data-k]'); if (b) { closeModal(); res(opts[+b.dataset.k]); } else if (e.target.id === 'pNo') { closeModal(); res(null); } };
    });
  });
}

// 전자 필기장 메뉴
$('nbBtn').onclick = () => {
  const m = $('nbMenu');
  if (!m.hidden) { m.hidden = true; return; }
  m.innerHTML = openNbs_().map(n => `<button data-nb="${n.id}" class="${n === cur.nb ? 'on' : ''}">▤ ${esc(n.name)}</button>`).join('') +
    `<hr><button data-a="new">＋ 새 전자 필기장</button><button data-a="open">📂 전자 필기장 열기·가져오기</button><button data-a="rename">이 전자 필기장 이름 바꾸기</button><button data-a="close">이 전자 필기장 닫기</button><button data-a="del" class="danger">이 전자 필기장 삭제</button>`;
  m.hidden = false;
};
$('nbMenu').onclick = async e => {
  const b = e.target.closest('button'); if (!b) return; $('nbMenu').hidden = true;
  if (b.dataset.nb) selectNotebook(ix.notebooks.find(n => n.id === b.dataset.nb));
  else if (b.dataset.a === 'new') addNotebook();
  else if (b.dataset.a === 'rename') renameNotebook(cur.nb);
  else if (b.dataset.a === 'open') openNotebookModal();
  else if (b.dataset.a === 'close') closeNotebook(cur.nb);
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
// 닫기: 지우지 않고 목록에서만 감춤 ([전자 필기장 열기·가져오기]에서 다시 열기)
function closeNotebook(nb) {
  if (openNbs_().length <= 1) return toast('열린 전자 필기장이 하나뿐이라 닫을 수 없어요.');
  nb.closed = true;
  if (nb === cur.nb) { const next = openNbs_()[0]; cur.nb = null; selectNotebook(next); }
  scheduleIndexSave(); renderAll();
  toast(`“${nb.name}”을(를) 닫았어요. [전자 필기장 열기·가져오기]에서 다시 열 수 있어요.`, [], 6000);
}
function openNotebookModal() {
  const closed = ix.notebooks.filter(n => n.closed);
  modal(`<h2>전자 필기장 열기·가져오기</h2>
    ${closed.length ? `<div class="ai-label">닫아 둔 전자 필기장</div>${closed.map(n => `<button class="choice" data-open="${n.id}"><b>▤ ${esc(n.name)}</b><small>섹션 ${n.sections.length}개 · 페이지 ${n.sections.reduce((a, s) => a + s.pages.length, 0)}개</small></button>`).join('')}` : '<p>닫아 둔 전자 필기장이 없어요.</p>'}
    <div class="ai-label">가져오기</div>
    <button class="choice" id="opOneNote"><b>원노트에서 가져오기</b><small>마이크로소프트 계정 연결 또는 Word 파일로 가져와요.</small></button>
    <div class="row"><button class="btn" id="opClose">닫기</button></div>`, m => {
    m.querySelector('#opClose').onclick = closeModal;
    m.querySelector('#opOneNote').onclick = () => { closeModal(); $('importOneNote').click(); };
    m.querySelectorAll('[data-open]').forEach(b => b.onclick = () => {
      const nb = ix.notebooks.find(n => n.id === b.dataset.open); nb.closed = false;
      closeModal(); scheduleIndexSave(); if (nb === cur.nb) renderAll(); else selectNotebook(nb);
      toast(`“${nb.name}”을(를) 다시 열었어요.`);
    });
  });
}
async function deleteNotebook(nb) {
  if (openNbs_().length <= 1 && !nb.closed) return toast('열린 전자 필기장이 하나뿐이라 지울 수 없어요.');
  const pages = nb.sections.flatMap(s => s.pages);
  if (!(await confirmDelete(`“${nb.name}” 전자 필기장 삭제`, `섹션 ${nb.sections.length}개, 페이지 ${pages.length}개가 함께 지워져요.${trashNote()}`))) return;
  ix.notebooks.splice(ix.notebooks.indexOf(nb), 1);
  if (nb === cur.nb) { dirty = false; editedSinceOpen = false; cur.nb = openNbs_()[0]; cur.sec = cur.nb.sections[0] || null; cur.page = null; openPage(cur.sec?.pages[0] || null); }
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
$('openNotebook').onclick = () => openNotebookModal();
// 배열에서 item을 target 앞(또는 뒤)으로 옮김
function moveInArray(arr, item, target, after) {
  if (item === target) return;
  arr.splice(arr.indexOf(item), 1);
  arr.splice(arr.indexOf(target) + (after ? 1 : 0), 0, item);
}
function notebookMenu(nb, at) {
  const vis = openNbs_(), k = vis.indexOf(nb);
  const c = showCtx(`<button data-a="rename">이름 바꾸기</button><button data-a="up" ${k > 0 ? '' : 'disabled'}>위로</button><button data-a="down" ${k < vis.length - 1 ? '' : 'disabled'}>아래로</button><button data-a="close">전자 필기장 닫기</button><hr><button data-a="del" class="danger">전자 필기장 삭제</button>`, at);
  c.onclick = e => {
    const b = e.target.closest('button'); if (!b) return; c.hidden = true;
    if (b.dataset.a === 'up' || b.dataset.a === 'down') { const other = vis[k + (b.dataset.a === 'up' ? -1 : 1)]; moveInArray(ix.notebooks, nb, other, b.dataset.a === 'down'); scheduleIndexSave(); renderNotebook(); }
    else if (b.dataset.a === 'rename') renameNotebook(nb); else if (b.dataset.a === 'close') closeNotebook(nb); else deleteNotebook(nb);
  };
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

// ---------- 꾹 눌러 위아래로 옮기기 (휴대폰·태블릿) ----------
// 꾹 누르면 진동 → 그대로 끌어서 놓으면 그 자리로. 꾹 눌렀다가 움직이지 않고 떼면 메뉴
function setupLongPressMove(ul, { canDrag, drop }) {
  let st = null, timer = 0, skipCtx = false;
  const clear = () => {
    clearTimeout(timer); timer = 0;
    if (st?.li) st.li.classList.remove('lp-lift');
    ul.querySelectorAll('.lp-before, .lp-after').forEach(x => x.classList.remove('lp-before', 'lp-after'));
    st = null;
  };
  const target = (x, y) => {
    const el = document.elementFromPoint(x, y)?.closest('li');
    if (!el || !ul.contains(el) || el === st.li || el.classList.contains('none')) return null;
    const r = el.getBoundingClientRect();
    return { el, after: y > r.top + r.height / 2 };
  };
  ul.addEventListener('touchstart', e => {
    if (e.touches.length !== 1) return clear();
    const li = e.target.closest('li'); if (!li || !canDrag(li) || e.target.closest('[data-more], [data-chev]')) return;
    const t = e.touches[0];
    st = { li, x: t.clientX, y: t.clientY, active: false };
    timer = setTimeout(() => {
      if (!st) return;
      st.active = true; skipCtx = true;
      navigator.vibrate?.(30);
      li.classList.add('lp-lift');
    }, 450);
  }, { passive: true });
  ul.addEventListener('touchmove', e => {
    if (!st) return;
    const t = e.touches[0];
    if (!st.active) { if (Math.hypot(t.clientX - st.x, t.clientY - st.y) > 10) clear(); return; }
    e.preventDefault(); // 끄는 동안 화면이 밀리지 않게
    st.moved = true;
    // 목록 위·아래 끝 근처면 목록을 굴림
    const box = ul.getBoundingClientRect();
    if (t.clientY < box.top + 40) ul.scrollTop -= 10; else if (t.clientY > box.bottom - 40) ul.scrollTop += 10;
    ul.querySelectorAll('.lp-before, .lp-after').forEach(x => x.classList.remove('lp-before', 'lp-after'));
    const tg = target(t.clientX, t.clientY); st.tg = tg;
    if (tg) tg.el.classList.add(tg.after ? 'lp-after' : 'lp-before');
  }, { passive: false });
  ul.addEventListener('touchend', e => {
    if (!st) return;
    if (!st.active) return clear();
    e.preventDefault(); // 놓을 때 누르기(click)로 처리되지 않게
    const { li, tg, moved, x, y } = st;
    clear();
    if (moved && tg) { drop(li, tg.el, tg.after); scheduleIndexSave(); }
    else if (!moved) { skipCtx = false; li.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: x, clientY: y })); }
    setTimeout(() => { skipCtx = false; }, 400);
  }, { passive: false });
  ul.addEventListener('touchcancel', clear);
  // 꾹 누를 때 휴대폰이 스스로 띄우는 메뉴(contextmenu)는 막음 — 떼었을 때 위에서 직접 띄움
  ul.addEventListener('contextmenu', e => { if (skipCtx) { e.preventDefault(); e.stopImmediatePropagation(); } }, true);
}
// 페이지: 같은 섹션 안에서 순서
setupLongPressMove($('pageList'), {
  canDrag: li => !searchMode && !!li.dataset.id,
  drop: (li, to, after) => {
    const a = cur.sec.pages.find(p => p.id === li.dataset.id), b = cur.sec.pages.find(p => p.id === to.dataset.id);
    if (a && b) moveInArray(cur.sec.pages, a, b, after);
    renderPages();
  },
});
// 섹션 (폴드·휴대폰의 섹션 칸): 같은 전자 필기장 안에서 순서
setupLongPressMove($('sectionList'), {
  canDrag: li => !!li.dataset.id,
  drop: (li, to, after) => {
    const a = cur.nb.sections.find(x => x.id === li.dataset.id), b = cur.nb.sections.find(x => x.id === to.dataset.id);
    if (a && b) moveInArray(cur.nb.sections, a, b, after);
    renderSections();
  },
});
// 전자 필기장 칸(트리): 전자 필기장끼리 순서, 섹션은 다른 전자 필기장으로도 옮길 수 있음
setupLongPressMove($('nbList'), {
  canDrag: li => !nbQ && (li.classList.contains('nb-item') || (li.dataset.sec && !li.dataset.addsec)),
  drop: (li, to, after) => {
    const nbOf = el => ix.notebooks.find(n => n.id === (el.dataset.nb || el.dataset.addsec || el.dataset.id));
    if (li.classList.contains('nb-item')) { // 전자 필기장 옮기기: 놓은 줄이 속한 전자 필기장 앞·뒤로
      const a = nbOf(li), b = nbOf(to); if (!a || !b || a === b) return;
      moveInArray(ix.notebooks, a, b, to.classList.contains('nb-item') ? after : true);
    } else { // 섹션 옮기기
      const fromNb = nbOf(li), sec = fromNb?.sections.find(x => x.id === li.dataset.sec), toNb = nbOf(to); if (!sec || !toNb) return;
      fromNb.sections.splice(fromNb.sections.indexOf(sec), 1);
      const tsec = to.dataset.sec ? toNb.sections.find(x => x.id === to.dataset.sec) : null;
      if (tsec) toNb.sections.splice(toNb.sections.indexOf(tsec) + (after ? 1 : 0), 0, sec);
      else if (to.classList.contains('nb-item') && after) toNb.sections.unshift(sec);
      else toNb.sections.push(sec);
      if (cur.sec === sec && toNb !== cur.nb) cur.nb = toNb;
      if (fromNb !== toNb) { openNbs.add(toNb.id); saveOpenNbs(); }
    }
    renderNotebook(); renderSections();
  },
});

// ---------- 검색 ----------
let searchTimer = 0;
// 일반 검색: 적는 대로 바로 (무료, 기기 안에서). AI 검색: AI 모드로 바꾼 뒤 Enter를 눌렀을 때만 (한 번마다 요금)
let aiSearchOn = false;
function setSearchMode(on) {
  aiSearchOn = on;
  const b = $('searchMode'), box = $('search');
  b.textContent = on ? '✨ AI' : '일반';
  b.classList.toggle('on', on);
  box.closest('.search').classList.toggle('ai', on);
  box.placeholder = on ? '찾고 싶은 내용을 말하듯 적고 Enter (AI 요금 발생)' : '모든 노트 검색';
}
$('searchMode').onclick = () => {
  setSearchMode(!aiSearchOn);
  clearTimeout(searchTimer);
  const q = $('search').value.trim();
  if (!aiSearchOn && q) runSearch(q); // 일반으로 돌아오면 바로 무료 검색
  if (aiSearchOn) toast('✨ AI 검색: 적고 Enter를 누를 때만 AI가 찾아요. 한 번 찾을 때마다 요금이 들어요.');
  $('search').focus();
};
$('search').addEventListener('input', () => {
  clearTimeout(searchTimer);
  const q = $('search').value.trim();
  if (!q) return endSearch();
  if (!aiSearchOn) searchTimer = setTimeout(() => runSearch(q), 450);
});
$('search').addEventListener('keydown', e => {
  if (e.key === 'Escape') { $('search').value = ''; endSearch(); }
  if (e.key === 'Enter') {
    clearTimeout(searchTimer);
    const q = $('search').value.trim(); if (!q) return;
    if (aiSearchOn) { if (!(searchMode?.ai && searchMode.q === q)) runAISearch(q); } // 같은 걸 또 누르면 다시 요금 들지 않게
    else runSearch(q);
  }
});
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
// 한 번에 한 단계씩 작게·크게 (휴대폰에서 목록을 열지 않아도 됨)
const FONT_STEPS = [9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];
function stepFont(d) {
  if (!needPage()) return; restoreRange();
  // 고른 글의 첫 글자 크기를 기준으로
  const s = getSelection(); if (!s.rangeCount) return;
  const rg = s.getRangeAt(0); let n = rg.startContainer;
  if (n.nodeType === 1) n = n.childNodes[rg.startOffset] || n;
  while (n && n.nodeType === 1 && n.firstChild) n = n.firstChild;
  const el = n && (n.nodeType === 1 ? n : n.parentElement);
  const pt = el ? parseFloat(getComputedStyle(el).fontSize) * 0.75 : 11.25;
  let i = d > 0 ? FONT_STEPS.findIndex(x => x > pt + 0.1) : FONT_STEPS.length - 1 - [...FONT_STEPS].reverse().findIndex(x => x < pt - 0.1);
  if (i < 0 || i >= FONT_STEPS.length) return;
  fontSize(FONT_STEPS[i]);
}
$('fontDown').onclick = () => stepFont(-1);
$('fontUp').onclick = () => stepFont(1);
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

// ---------- 글 너비 (원노트처럼 오른쪽 손잡이를 끌어서) ----------
function applyPageWidth(p) {
  const w = p && p.width;
  $('page').style.width = w ? w + 'px' : '';
  $('page').classList.toggle('fixed-w', !!w);
}
(() => {
  const h = $('widthHandle');
  h.addEventListener('pointerdown', e => {
    if (!cur.page) return;
    e.preventDefault();
    const x0 = e.clientX, w0 = $('page').getBoundingClientRect().width;
    document.body.classList.add('resizing'); h.classList.add('active');
    const move = ev => { cur.page.width = Math.round(Math.min(4000, Math.max(320, w0 + ev.clientX - x0))); applyPageWidth(cur.page); };
    const up = () => { removeEventListener('pointermove', move); removeEventListener('pointerup', up); removeEventListener('pointercancel', up); document.body.classList.remove('resizing'); h.classList.remove('active'); scheduleIndexSave(); };
    addEventListener('pointermove', move); addEventListener('pointerup', up); addEventListener('pointercancel', up);
  });
  h.addEventListener('dblclick', () => { if (!cur.page) return; delete cur.page.width; applyPageWidth(cur.page); scheduleIndexSave(); toast('글 너비를 창 너비에 맞췄어요.'); });
})();

// ---------- 휴대폰 뒤로 가기: 글을 쓰는 중이면 키보드와 커서만 닫음 ----------
(() => {
  let editing = false, ignorePop = false;
  const touch = () => matchMedia('(pointer: coarse)').matches;
  const inEdit = () => [$('editor'), $('pageTitle')].includes(document.activeElement);
  ['editor', 'pageTitle'].forEach(id => $(id).addEventListener('focus', () => { if (touch() && !editing) { editing = true; history.pushState({ esn: 'edit' }, ''); } }));
  document.addEventListener('focusout', () => setTimeout(() => {
    if (editing && !inEdit()) { editing = false; ignorePop = true; history.back(); } // 다른 곳을 눌러 빠져나간 경우 뒤로 가기 기록 정리
  }, 0));
  addEventListener('popstate', () => {
    if (ignorePop) { ignorePop = false; return; }
    if (editing) { editing = false; document.activeElement?.blur(); getSelection().removeAllRanges(); }
  });
})();

// ---------- 본문을 눌렀을 때 화면이 저절로 움직이지 않게 ----------
// 휴대폰·폴드는 키보드가 올라올 때 브라우저가 스스로 화면을 굴림 → 누른 자리(커서)가 보이는 한 원래 자리로 되돌림
(() => {
  const cv = $('canvas');
  let guard = null;
  const caretRect = () => {
    const s = getSelection(); if (!s.rangeCount) return null;
    const r = s.getRangeAt(0).cloneRange(); r.collapse(true);
    let b = r.getBoundingClientRect();
    if (!b.height) { const n = r.startContainer; const el = n.nodeType === 1 ? (n.childNodes[r.startOffset] || n) : n.parentElement; b = el?.getBoundingClientRect?.() || b; }
    return b.height ? b : null;
  };
  const restore = () => {
    if (!guard || performance.now() > guard.until) { guard = null; return; }
    const doc = document.scrollingElement;
    if (doc.scrollTop) doc.scrollTop = 0; // 화면 전체가 밀린 경우
    const d = cv.scrollTop - guard.top; if (!d) return;
    const c = caretRect(); if (!c) return;
    const box = cv.getBoundingClientRect();
    const bottom = Math.min(box.bottom, (visualViewport?.height || innerHeight) + (visualViewport?.offsetTop || 0)) - 8;
    const y = c.top + d; // 원래 자리로 되돌렸을 때 커서 위치
    if (y >= box.top && y + c.height <= bottom) cv.scrollTop = guard.top; // 그래도 커서가 보이면 되돌림
  };
  const arm = () => { guard = { top: cv.scrollTop, until: performance.now() + 900 }; };
  cv.addEventListener('touchstart', e => { if ($('editor').contains(e.target)) arm(); }, { passive: true });
  cv.addEventListener('touchmove', () => { guard = null; }, { passive: true }); // 손가락으로 굴리는 건 그대로
  cv.addEventListener('mousedown', e => { if ($('editor').contains(e.target)) arm(); });
  cv.addEventListener('wheel', () => { guard = null; }, { passive: true });
  cv.addEventListener('scroll', restore, { passive: true });
  addEventListener('scroll', restore, { passive: true });
  visualViewport?.addEventListener('resize', () => requestAnimationFrame(restore));
})();

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
  d.querySelector('b').textContent = on ? { narrow: '좁게', medium: '중간', wide: '넓게', grid: '모눈' }[S.defaultLines] : '끔';
}
document.querySelectorAll('[data-lines]').forEach(b => b.onclick = () => {
  if (!needPage()) return;
  if (b.dataset.lines) cur.page.lines = b.dataset.lines; else delete cur.page.lines;
  applyLines(cur.page); scheduleIndexSave();
  $('linesPop').hidden = true;
});

// ---------- 칸 너비 조절 (컴퓨터): 전자 필기장·섹션 칸과 페이지 칸의 경계를 끌어서 ----------
const COL_DEFAULT = { nb: 240, pages: 230 };
let colW = {};
try { colW = JSON.parse(localStorage.getItem('dn-colw') || '{}'); } catch {}
function applyColW() { for (const k in COL_DEFAULT) $('shell').style.setProperty('--w-' + k, (colW[k] || COL_DEFAULT[k]) + 'px'); }
document.querySelectorAll('.resizer').forEach(r => {
  const k = r.dataset.col;
  r.addEventListener('pointerdown', e => {
    e.preventDefault();
    const x0 = e.clientX, w0 = r.parentElement.getBoundingClientRect().width;
    document.body.classList.add('resizing'); r.classList.add('active');
    // 마우스가 경계선을 벗어나도 계속 따라가도록 창 전체에서 받음
    const move = ev => { colW[k] = Math.round(Math.min(560, Math.max(150, w0 + ev.clientX - x0))); applyColW(); };
    const up = () => {
      removeEventListener('pointermove', move); removeEventListener('pointerup', up); removeEventListener('pointercancel', up);
      document.body.classList.remove('resizing'); r.classList.remove('active');
      try { localStorage.setItem('dn-colw', JSON.stringify(colW)); } catch {}
    };
    addEventListener('pointermove', move); addEventListener('pointerup', up); addEventListener('pointercancel', up);
  });
  r.addEventListener('dblclick', () => { delete colW[k]; applyColW(); try { localStorage.setItem('dn-colw', JSON.stringify(colW)); } catch {} });
});
applyColW();
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
  try { await store.getToken(); } catch (e) { return handleErr(e, runSpell); }
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
      const r = await fetch(SPELL_API, { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + await store.getToken() }, body: JSON.stringify({ p: g.map(b => b.text.slice(0, 2000)) }) });
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
function pageForAI(root = $('editor')) {
  const c = root.cloneNode(true);
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
// AI 도우미 성능: 기본(Haiku, 저렴) / 꼼꼼(Sonnet) / 최고(Opus). 맞춤법·AI 검색은 늘 기본
const AI_TIERS = [['fast', '기본', '빠르고 저렴 · 고치기 1번 약 10원'], ['good', '꼼꼼', '정리·고치기가 더 꼼꼼 · 고치기 1번 약 200원'], ['best', '최고', '가장 똑똑함 · 고치기 1번 약 400원']];
let aiTier = 'fast';
try { if (AI_TIERS.some(t => t[0] === localStorage.getItem('dn-ai-tier'))) aiTier = localStorage.getItem('dn-ai-tier'); } catch {}
async function aiCall(mode, prompt, root, extra) {
  const r = await fetch(AI_API, { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + await store.getToken() }, body: JSON.stringify({ mode, prompt, title: $('pageTitle').value, html: mode === 'search' ? '' : pageForAI(root), ...(mode === 'search' ? {} : { tier: aiTier }), ...extra }) });
  const j = await r.json().catch(() => ({}));
  if (r.ok) return j;
  const msg = { 503: 'AI 서버에 AI 키가 설정돼 있지 않아요.', 401: '이 계정은 AI 도우미를 쓸 수 없어요. 다시 로그인해 보세요.', 413: '페이지가 너무 길어요. 페이지를 나눠서 해 보세요.', 404: 'AI 서버를 찾지 못했어요. config.js의 aiApi를 확인해 주세요.' }[r.status];
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
      <div class="ai-label">성능</div>
      <div class="ai-tier" id="aiTier">${AI_TIERS.map(t => `<button data-tier="${t[0]}" class="${t[0] === aiTier ? 'on' : ''}">${t[1]}</button>`).join('')}</div>
      <div class="ai-hint left" id="aiTierHint">${AI_TIERS.find(t => t[0] === aiTier)[2]}</div>
      <div class="ai-label">이 페이지를 바로 고치기</div>
      <div class="ai-chips">${AI_PRESETS.map((p, i) => `<button data-preset="${i}">${esc(p[0])}</button>`).join('')}</div>
      <div class="ai-label">직접 쓰기</div>
      <textarea id="aiPrompt" rows="3" placeholder="궁금한 점을 묻거나, 원하는 걸 적어 주세요.&#10;예) 이 내용으로 발표 순서를 짜 줘&#10;예) 날짜순으로 다시 정리해 줘"></textarea>
      <div class="ai-acts"><button class="btn" id="aiAsk" title="Ctrl+Enter">질문하기</button><button class="btn primary" id="aiEdit">이대로 페이지 고치기</button></div>
      <div id="aiOut"></div>
    </div>`;
  $('aiClose').onclick = () => { el.hidden = true; };
  $('aiTier').onclick = e => {
    const b = e.target.closest('[data-tier]'); if (!b) return;
    aiTier = b.dataset.tier; try { localStorage.setItem('dn-ai-tier', aiTier); } catch {}
    $('aiTier').querySelectorAll('button').forEach(x => x.classList.toggle('on', x === b));
    $('aiTierHint').textContent = AI_TIERS.find(t => t[0] === aiTier)[2];
  };
  const promptBox = $('aiPrompt');
  promptBox.onkeydown = e => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); run('ask', promptBox.value.trim()); } };
  el.querySelector('.ai-chips').onclick = e => { const b = e.target.closest('[data-preset]'); if (b) run('edit', AI_PRESETS[+b.dataset.preset][1]); };
  $('aiAsk').onclick = () => run('ask', promptBox.value.trim());
  $('aiEdit').onclick = () => run('edit', promptBox.value.trim());
  promptBox.focus();

  const page = cur.page;
  el._run = run; // 검토 중 [이대로 다시 고치기]에서 씀
  async function run(mode, prompt, base) {
    if (!prompt) { promptBox.focus(); return toast('무엇을 할지 적어 주세요.'); }
    try { await store.getToken(); } catch (e) { return handleErr(e, () => run(mode, prompt)); }
    const out = $('aiOut');
    el.querySelectorAll('button').forEach(b => { if (b.id !== 'aiClose') b.disabled = true; });
    const keepOrig = base ? out._orig : $('editor').cloneNode(true); // 처음 페이지 모습 (바뀐 부분 비교용)
    out.innerHTML = `<div class="spell-empty">AI가 ${mode === 'edit' ? '페이지를 고치는' : '생각하는'} 중…<br><small>긴 페이지는 30초~1분쯤 걸려요.</small></div>`;
    out._orig = keepOrig;
    try {
      const j = await aiCall(mode, prompt, base);
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
      } else showReview(out, j, page);
    } catch (e) { out.innerHTML = `<p class="ai-note warn">${esc(e.message || 'AI 도우미가 답하지 못했어요.')}</p>`; }
    finally { el.querySelectorAll('button').forEach(b => { b.disabled = false; }); }
  }
}
$('aiOpen').onclick = openAI;

// ---- 고친 결과 검토: 바뀐 부분 보기, 직접 고치기, 추가 요청, 적용 ----
function showReview(out, j, page) {
  const { html, missing } = restoreKept(j.html || '');
  out.innerHTML = `${j.note ? `<p class="ai-note">${esc(j.note)}</p>` : ''}${missing ? `<p class="ai-note warn">그림·첨부 ${missing}개는 맨 아래에 붙였어요.</p>` : ''}
    <div class="ai-tabs"><button data-t="diff" class="on">바뀐 부분</button><button data-t="edit">고친 결과 · 직접 고치기</button></div>
    <div class="ai-diff" id="aiDiff"></div>
    <div class="ai-preview editor" id="aiPrev" contenteditable="true" spellcheck="false" hidden>${html}</div>
    <div class="ai-label">더 고칠 점이 있으면</div>
    <textarea id="aiMore" rows="2" placeholder="예) 둘째 문단은 원래대로 둬 줘 / 요약을 더 짧게"></textarea>
    <div class="ai-acts"><button class="btn" id="aiMoreGo">이대로 다시 고치기</button></div>
    <div class="ai-acts sep"><button class="btn" id="aiDiscard">버리기</button><button class="btn primary" id="aiApply">검토 끝 · 페이지에 적용</button></div>
    <p class="ai-hint">적용하기 전에는 페이지가 바뀌지 않아요. 적용한 뒤에도 Ctrl+Z로 되돌릴 수 있어요.</p>`;
  const prev = out.querySelector('#aiPrev'), diff = out.querySelector('#aiDiff');
  prev.querySelectorAll('img[data-fid]').forEach(async img => { const u = await fileURL(img.dataset.fid).catch(() => ''); if (u) img.src = u; });
  const showDiff = () => { diff.innerHTML = diffHTML(out._orig, prev); };
  showDiff();
  out.querySelector('.ai-tabs').onclick = e => {
    const t = e.target.closest('[data-t]'); if (!t) return;
    out.querySelectorAll('.ai-tabs button').forEach(x => x.classList.toggle('on', x === t));
    prev.hidden = t.dataset.t !== 'edit'; diff.hidden = t.dataset.t !== 'diff';
    if (t.dataset.t === 'diff') showDiff(); else prev.focus();
  };
  // 추가 요청: 지금 검토 중인 결과를 바탕으로 다시 고침
  out.querySelector('#aiMoreGo').onclick = () => {
    const more = out.querySelector('#aiMore').value.trim(); if (!more) return out.querySelector('#aiMore').focus();
    const base = prev.cloneNode(true); base.querySelectorAll('img[data-fid]').forEach(i => i.removeAttribute('src'));
    $('ai')._run('edit', more, base);
  };
  out.querySelector('#aiDiscard').onclick = () => { out.innerHTML = '<p class="ai-note">버렸어요. 페이지는 그대로예요.</p>'; };
  out.querySelector('#aiApply').onclick = () => {
    if (cur.page !== page) return toast('다른 페이지로 바뀌어서 적용하지 않았어요.');
    const c = prev.cloneNode(true); c.querySelectorAll('img[data-fid]').forEach(i => i.removeAttribute('src'));
    const ed = $('editor'), r = document.createRange(); r.selectNodeContents(ed);
    ed.focus(); selectRange(r); document.execCommand('insertHTML', false, sanitize(c.innerHTML));
    hydrate(); markDirty();
    out.innerHTML = '<p class="ai-note">검토한 내용을 페이지에 적용했어요. 되돌리려면 Ctrl+Z를 누르세요.</p>';
  };
}
// 문단 단위로 원래 페이지와 고친 결과를 비교 (같은 문단은 흐리게, 바뀐 낱말은 표시)
function blockTexts(root) {
  const B = 'p, h1, h2, h3, h4, li, td, th, blockquote, pre';
  return [...root.querySelectorAll(B)].filter(el => !el.querySelector(B)).map(el => el.textContent.replace(/\s+/g, ' ').trim()).filter(Boolean).slice(0, 800);
}
function lcsPairs(a, b, eq) {
  const n = a.length, m = b.length, dp = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) dp[i][j] = eq(a[i], b[j]) ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
  const ops = []; let i = 0, j = 0;
  while (i < n && j < m) { if (eq(a[i], b[j])) { ops.push(['=', a[i]]); i++; j++; } else if (dp[i + 1][j] >= dp[i][j + 1]) ops.push(['-', a[i++]]); else ops.push(['+', b[j++]]); }
  while (i < n) ops.push(['-', a[i++]]); while (j < m) ops.push(['+', b[j++]]);
  return ops;
}
function wordDiff(a, b) {
  const A = a.split(/(\s+)/), Bw = b.split(/(\s+)/);
  if (A.length * Bw.length > 60000) return [`<del>${esc(a)}</del>`, `<ins>${esc(b)}</ins>`];
  const ops = lcsPairs(A, Bw, (x, y) => x === y);
  const side = (skip, mark, tag) => { let h = '', run = ''; for (const [t, w] of ops) { if (t === skip) continue; if (t === mark) { run += w; continue; } if (run) { h += run.trim() ? `<${tag}>${esc(run)}</${tag}>` : esc(run); run = ''; } h += esc(w); } if (run) h += `<${tag}>${esc(run)}</${tag}>`; return h; };
  return [side('+', '-', 'del'), side('-', '+', 'ins')];
}
const similar = (a, b) => { const A = bigramsOf(a), B = bigramsOf(b); if (!A.size || !B.size) return false; let c = 0; A.forEach(g => { if (B.has(g)) c++; }); return c / Math.min(A.size, B.size) >= 0.35; };
function diffHTML(origRoot, newRoot) {
  const ops = lcsPairs(blockTexts(origRoot), blockTexts(newRoot), (x, y) => x === y);
  let html = '', ch = 0, add = 0, del = 0;
  for (let k = 0; k < ops.length; k++) {
    const [t, txt] = ops[k];
    if (t === '=') { html += `<p class="d-same">${esc(txt)}</p>`; continue; }
    if (t === '-' && ops[k + 1]?.[0] === '+' && similar(txt, ops[k + 1][1])) { const [x, y] = wordDiff(txt, ops[k + 1][1]); html += `<div class="d-chg"><p class="d-del">${x}</p><p class="d-add">${y}</p></div>`; ch++; k++; continue; }
    if (t === '-') { html += `<p class="d-del whole"><del>${esc(txt)}</del></p>`; del++; }
    else { html += `<p class="d-add whole"><ins>${esc(txt)}</ins></p>`; add++; }
  }
  const sum = ch + add + del ? `바뀐 문단 ${ch}개 · 새로 생긴 ${add}개 · 빠진 ${del}개` : '글 내용은 그대로예요 (모양·순서만 바뀌었을 수 있어요)';
  return `<p class="d-sum">${sum} <span>· <del>빨강</del> 빠짐 · <ins>초록</ins> 새로 씀</span></p>${html}`;
}

// ---------- AI로 찾기: 말로 찾고 싶은 내용을 쓰면 노트 전체에서 맞는 페이지를 찾음 ----------
const bigramsOf = t => { const s2 = t.replace(/\s+/g, '').toLowerCase(), out = new Set(); for (let i = 0; i < s2.length - 1; i++) out.add(s2.slice(i, i + 2)); return out; };
async function pageText(p) {
  const c = store.kind === 'drive' ? await store.cachedPage(p) : { html: await store.loadPage(p) };
  return c ? parseHTML(c.html).textContent.replace(/\s+/g, ' ').trim() : (p.snip || '');
}
async function runAISearch(q) {
  if (!q) return toast('찾고 싶은 내용을 검색창에 적어 주세요.');
  if (store?.kind !== 'drive') return toast('AI로 찾기는 구글 드라이브로 로그인했을 때 쓸 수 있어요.');
  try { await store.getToken(); } catch (e) { return handleErr(e, () => runAISearch(q)); }
  await flush().catch(() => {});
  searchMode = { q, ids: [], ai: true, loading: true }; renderPages(); if (phone()) setView('pages');
  // 1) 앱에서 먼저 후보를 고름 (글자 겹침 점수) → 2) AI가 그중 맞는 걸 고르고 이유를 붙임
  const qb = bigramsOf(q), list = allPages().filter(x => !x.nb.closed), scored = [];
  for (const x of list) {
    const text = await pageText(x.p), title = x.p.title || '';
    const tb = bigramsOf(title), bb = bigramsOf(text);
    let score = 0; qb.forEach(g => { if (tb.has(g)) score += 3; if (bb.has(g)) score += 1; });
    // 가장 많이 겹치는 곳 근처를 잘라서 보냄
    let at = 0, best = -1; const low = text.toLowerCase();
    for (const g of qb) { const i = low.indexOf(g); if (i >= 0 && (best < 0 || i < best)) best = i; }
    if (best > 0) at = Math.max(0, best - 150);
    scored.push({ x, score, text: text.slice(at, at + 600) });
  }
  scored.sort((a, b) => b.score - a.score);
  const cands = scored.filter(c => c.score > 0).slice(0, 50);
  if (cands.length < 15) cands.push(...scored.filter(c => c.score === 0).slice(0, 15 - cands.length)); // 낱말이 안 겹쳐도 뜻으로 찾을 수 있게
  try {
    const j = await aiCall('search', q, null, { items: cands.map(c => ({ id: c.x.p.id, title: c.x.p.title || '제목 없음', path: c.x.nb.name + ' › ' + c.x.sec.name, text: c.text })) });
    if (searchMode?.q !== q) return;
    searchMode = { q, ai: true, ids: j.results.map(r => r.id), why: Object.fromEntries(j.results.map(r => [r.id, r.why])) };
    renderPages();
  } catch (e) { searchMode = null; renderPages(); handleErr(e); }
}


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
      ${sw('autoBullet', '“- ” “* ” “ㅁ ” “: ” → 글머리 기호', '', 'sub-sw')}
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
    m.querySelector('#stOut')?.addEventListener('click', async () => { await flush().catch(() => {}); await store.signOut(); localStorage.removeItem('dn-mode'); location.reload(); });
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
  if (interactive) {
    try { await store.signIn(); } catch (e) { toast(e.message, [], 8000); return driveGate(); }
    if (store.noRefresh) toast('로그인을 계속 유지하려면 한 번만 더 해 주세요: 구글 계정 설정 → 보안 → 타사 앱 연결에서 “에스 노트”를 지운 뒤 다시 로그인하면 돼요.', [], 15000);
  } else if (!store.restore()) {
    // 켤 때마다 로그인하지 않도록: 로그인 유지 서버에서 저절로 다시 로그인
    try { await store.getToken(); }
    catch (e) { if (!e.offline) return driveGate(); } // 인터넷이 끊겼으면 기기에 보관한 노트로 열기
  }
  try {
    await load();
    if (offerMigrate) migrate();
    afterDriveLoad();
  } catch (e) {
    if (e.auth) return driveGate();
    if (e.offline) return offlineGate();
    handleErr(e, () => location.reload());
  }
}
// 인터넷이 끊긴 채로 켰는데 기기에 보관한 노트도 없을 때
function offlineGate() {
  showEmpty('<h2>인터넷 연결을 기다리는 중이에요</h2><p>연결되면 저절로 노트를 불러와요. 로그인은 그대로 유지돼요.</p>');
  status('offline');
  addEventListener('online', () => startDrive(false), { once: true });
}
// 드라이브에서 열고 난 뒤: 끊겼을 때 써 둔 페이지 올리기, 같은 섹션 페이지 미리 받아 두기
async function afterDriveLoad() {
  if (store.offline) { status('', '⚠ 오프라인 · 이 기기에 저장 중'); return; }
  try { const n = await store.flushPending(id => findPage(id)?.p); if (n) { scheduleIndexSave(); toast(`인터넷이 끊겼을 때 쓴 페이지 ${n}개를 드라이브에 올렸어요.`); } } catch {}
  prefetchSection(cur.sec);
}
// 모든 페이지를 뒤에서 미리 받아 둠: 고른 섹션 먼저, 그다음 나머지 (한 번 받으면 다음부터 바로 열림)
let prefetching = null;
async function prefetchSection(sec) {
  if (store?.kind !== 'drive' || !navigator.onLine || !ix) return;
  const token = prefetching = {};
  const rest = allPages().filter(x => !x.nb.closed && x.sec !== sec).map(x => x.p);
  const queue = [...(sec ? sec.pages : []), ...rest];
  const worker = async () => {
    while (queue.length) {
      if (prefetching !== token) return; // 다른 섹션으로 바뀌면 그 섹션부터 다시
      const p = queue.shift();
      const c = await store.cachedPage(p);
      if (c && (c.pending || (c.updated >= (p.updated || 0) && c.fid === p.fid))) continue;
      if (!p.fid) continue;
      try { await store.fetchPage(p); } catch (e) { if (e.offline || e.auth) return; }
    }
  };
  await Promise.all([worker(), worker(), worker()]); // 세 개씩 함께
}
// 페이지를 누르는 순간(손을 떼기 전)부터 받기 시작
$('pageList').addEventListener('pointerdown', e => {
  const li = e.target.closest('li[data-id]'); if (!li || store?.kind !== 'drive') return;
  const h = findPage(li.dataset.id); if (h && h.p.fid) store.loadPage(h.p).catch(() => {});
});
// 새로 열 때마다 구글 로그인 버튼 한 번 (팝업은 버튼을 눌러야 열려서)
function driveGate() {
  showEmpty(`<h2>구글 드라이브에 연결해 주세요</h2><p>노트는 내 구글 드라이브의 “에스 노트” 폴더에 저장돼요.</p><div class="btns"><button class="btn primary" id="gateGo">구글로 로그인</button><button class="btn" id="gateLocal">이 기기에서 체험하기</button></div>`);
  $('account').textContent = '로그인 필요'; status('', '');
  $('gateGo').onclick = () => startDrive(true);
  $('gateLocal').onclick = () => { localStorage.setItem('dn-mode', 'local'); startLocal(); };
}
function welcome() {
  $('modal').dataset.locked = '1';
  modal(`<div class="welcome"><img class="logo-img big" src="icons/icon-192.png" alt=""><h2>에스 노트에 오신 걸 환영해요</h2><p>노트를 어디에 저장할까요? 나중에 설정에서 바꿀 수 있어요.</p>
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
