// 원노트 → 에스 노트 가져오기 (마이크로소프트 Graph, 원노트 읽기 권한만 사용)
// 로그인: 팝업 + PKCE (라이브러리 없이). 돌아오는 주소는 ms-auth.html
import { sanitize } from './doc.js';

const AUTH = 'https://login.microsoftonline.com/common/oauth2/v2.0';
const GRAPH = 'https://graph.microsoft.com/v1.0';
const SCOPE = 'Notes.Read offline_access';
const redirectUri = () => location.origin + location.pathname.replace(/[^/]*$/, '') + 'ms-auth.html';
const b64url = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const sleep = ms => new Promise(r => setTimeout(r, ms));

let tok = null; // {access, refresh, exp}

// ---------- 로그인 ----------
// 버튼을 누른 순간 창부터 열어 둬야 팝업 차단에 걸리지 않음
export async function msLogin(clientId) {
  const pop = window.open('about:blank', 'ms-login', 'width=520,height=700');
  if (!pop) throw new Error('로그인 창이 열리지 않았어요. 팝업 차단을 풀어 주세요.');
  const verifier = b64url(crypto.getRandomValues(new Uint8Array(32)));
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  const state = b64url(crypto.getRandomValues(new Uint8Array(12)));
  try { localStorage.removeItem('ms-auth-result'); } catch {}
  pop.location.href = `${AUTH}/authorize?` + new URLSearchParams({
    client_id: clientId, response_type: 'code', redirect_uri: redirectUri(), response_mode: 'query',
    scope: SCOPE, state, code_challenge: challenge, code_challenge_method: 'S256', prompt: 'select_account',
  });
  const search = await waitForRedirect(pop);
  const p = new URLSearchParams(search);
  if (p.get('error')) throw new Error(p.get('error') === 'access_denied' ? '원노트 읽기 권한을 허락하지 않았어요.' : (p.get('error_description') || p.get('error')));
  if (p.get('state') !== state || !p.get('code')) throw new Error('로그인 정보를 확인하지 못했어요. 다시 시도해 주세요.');
  await tokenRequest(clientId, { grant_type: 'authorization_code', code: p.get('code'), redirect_uri: redirectUri(), code_verifier: verifier });
}
// ms-auth.html이 결과를 보내는 길 세 가지(메시지, 저장소, 창 주소) 중 먼저 오는 것
function waitForRedirect(pop) {
  return new Promise((res, rej) => {
    let closedAt = 0;
    const done = v => { clearInterval(t); removeEventListener('message', onMsg); removeEventListener('storage', onStore); try { pop.close(); } catch {} res(v); };
    const onMsg = e => { if (e.origin === location.origin && e.data && typeof e.data.msAuth === 'string') done(e.data.msAuth); };
    const onStore = e => { if (e.key === 'ms-auth-result' && e.newValue) done(e.newValue); };
    addEventListener('message', onMsg); addEventListener('storage', onStore);
    const t = setInterval(() => {
      try { if (pop.location.origin === location.origin && pop.location.search) return done(pop.location.search); } catch {}
      try { const v = localStorage.getItem('ms-auth-result'); if (v) { localStorage.removeItem('ms-auth-result'); return done(v); } } catch {}
      if (pop.closed) { closedAt ||= Date.now(); if (Date.now() - closedAt > 1500) { clearInterval(t); removeEventListener('message', onMsg); removeEventListener('storage', onStore); rej(new Error('로그인 창이 닫혔어요.')); } }
    }, 300);
  });
}
async function tokenRequest(clientId, body) {
  const r = await fetch(`${AUTH}/token`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams({ client_id: clientId, scope: SCOPE, ...body }) });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error('마이크로소프트 로그인 오류: ' + (j.error_description || j.error || r.status));
  tok = { access: j.access_token, refresh: j.refresh_token || tok?.refresh, exp: Date.now() + (j.expires_in - 120) * 1000, clientId };
}
export const msSignedIn = () => !!tok;

// ---------- Graph 요청 (요청이 많으면 잠깐 쉬었다가 다시) ----------
async function g(url, raw) {
  for (let i = 0; ; i++) {
    if (tok.exp < Date.now() && tok.refresh) await tokenRequest(tok.clientId, { grant_type: 'refresh_token', refresh_token: tok.refresh });
    let r;
    try { r = await fetch(url.startsWith('http') ? url : GRAPH + url, { headers: { Authorization: 'Bearer ' + tok.access } }); }
    catch { if (i < 5) { await sleep(3000 * (i + 1)); continue; } throw new Error('인터넷 연결이 끊겼어요.'); }
    if ((r.status === 429 || r.status === 503 || r.status === 504) && i < 8) { await sleep((+r.headers.get('Retry-After') || 2 ** i) * 1000); continue; }
    if (r.status === 401 && tok.refresh && i < 2) { await tokenRequest(tok.clientId, { grant_type: 'refresh_token', refresh_token: tok.refresh }); continue; }
    if (!r.ok) {
      const j = await r.json().catch(() => null);
      throw new Error(`원노트 오류 ${r.status}: ${j?.error?.message || ''}`.trim());
    }
    return raw ? r : r.json();
  }
}
async function all(url) {
  const out = [];
  for (let next = url; next;) { const j = await g(next); out.push(...(j.value || [])); next = j['@odata.nextLink']; }
  return out;
}

// ---------- 목록 ----------
// → [{id, name, sections:[{id, name, pages: n?}]}]
export async function listNotebooks() {
  const [nbs, secs] = await Promise.all([
    all('/me/onenote/notebooks?$select=id,displayName&$orderby=displayName'),
    all('/me/onenote/sections?$select=id,displayName&$expand=parentNotebook($select=id),parentSectionGroup($select=id,displayName)'),
  ]);
  return nbs.map(nb => ({
    id: nb.id, name: nb.displayName,
    sections: secs.filter(s => s.parentNotebook?.id === nb.id)
      .map(s => ({ id: s.id, name: (s.parentSectionGroup ? s.parentSectionGroup.displayName + ' / ' : '') + s.displayName })),
  }));
}
export async function listPages(sectionId) {
  const ps = await all(`/me/onenote/sections/${sectionId}/pages?pagelevel=true&$top=100&$select=id,title,createdDateTime,lastModifiedDateTime,level,order`);
  return ps.sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
}

export async function pageContent(pageId) { return (await g(`/me/onenote/pages/${pageId}/content`, true)).text(); }

// ---------- 페이지 내용 바꾸기 ----------
const esc = s => s.replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const TAG_MARK = { important: '⭐ ', question: '❓ ', critical: '❗ ', idea: '💡 ', remember: '📌 ' };
const absPos = el => { const st = el.getAttribute?.('style') || ''; return /position:\s*absolute/.test(st) ? { top: parseFloat(st.match(/top:\s*([\d.]+)/)?.[1] || 0), left: parseFloat(st.match(/left:\s*([\d.]+)/)?.[1] || 0) } : null; };

// 원노트 HTML → 에스 노트 HTML. putFile(blob, name) → fid
export async function convertPage(html, putFile) {
  const body = new DOMParser().parseFromString(html, 'text/html').body;
  // 원노트는 글상자를 좌표로 배치함 → 위에서 아래, 왼쪽에서 오른쪽 순서로 늘어놓기
  const boxes = [...body.children].map((el, i) => ({ el, i, p: absPos(el) }));
  if (boxes.some(b => b.p)) boxes.sort((a, b) => (a.p?.top ?? 0) - (b.p?.top ?? 0) || (a.p?.left ?? 0) - (b.p?.left ?? 0) || a.i - b.i).forEach(b => body.append(b.el));

  // 할 일 태그 → 체크 목록, 다른 태그는 기호로
  for (const el of [...body.querySelectorAll('[data-tag]')]) {
    const tags = el.getAttribute('data-tag').split(',').map(t => t.trim());
    const todo = tags.find(t => t.startsWith('to-do'));
    const mark = tags.map(t => TAG_MARK[t] || '').join('');
    if (mark) el.prepend(mark);
    if (!todo) continue;
    const done = todo.includes('completed');
    if (el.tagName === 'LI') { el.parentNode.classList.add('checklist'); if (done) el.setAttribute('data-checked', ''); continue; }
    const prev = el.previousElementSibling;
    const ul = prev && prev.tagName === 'UL' && prev.classList.contains('checklist') && prev.dataset.fromTodo ? prev : Object.assign(document.createElement('ul'), { className: 'checklist' });
    if (ul !== prev) { ul.dataset.fromTodo = '1'; el.before(ul); }
    const li = document.createElement('li'); if (done) li.setAttribute('data-checked', '');
    li.append(...el.childNodes); ul.append(li); el.remove();
  }
  body.querySelectorAll('ul[data-from-todo]').forEach(u => u.removeAttribute('data-from-todo'));

  // 그림 → 내 저장소로 옮기기
  for (const img of [...body.querySelectorAll('img')]) {
    const src = img.getAttribute('data-fullres-src') || img.getAttribute('src') || '';
    const w = parseFloat(img.getAttribute('width') || img.style.width) || 0;
    img.removeAttribute('style'); img.removeAttribute('width'); img.removeAttribute('height');
    if (w) img.style.width = Math.min(100, Math.round(w / 6.24)) + '%'; // 원노트 본문 폭 약 624px 기준
    if (!src.startsWith(GRAPH) && !src.startsWith('https://graph.microsoft.com/')) continue;
    try {
      const blob = await (await g(src, true)).blob();
      img.setAttribute('data-fid', await putFile(blob, (img.getAttribute('alt') || 'image').slice(0, 40) + '.' + ((blob.type.split('/')[1] || 'png').replace('jpeg', 'jpg'))));
      img.removeAttribute('src');
    } catch { img.replaceWith('[그림을 가져오지 못했어요]'); }
  }
  // 첨부 파일
  for (const ob of [...body.querySelectorAll('object')]) {
    const name = ob.getAttribute('data-attachment') || '첨부 파일';
    try {
      const blob = await (await g(ob.getAttribute('data'), true)).blob();
      const fid = await putFile(new Blob([blob], { type: ob.getAttribute('type') || blob.type }), name);
      const size = blob.size < 1048576 ? Math.max(1, Math.round(blob.size / 1024)) + 'KB' : (blob.size / 1048576).toFixed(1) + 'MB';
      const span = document.createElement('span');
      span.innerHTML = `<span class="att" contenteditable="false" data-fid="${fid}" data-name="${esc(name)}" data-size="${blob.size}">📎 ${esc(name)} <small>${size}</small></span>`;
      ob.replaceWith(span.firstChild);
    } catch { ob.replaceWith(`[첨부 파일 ${name}을(를) 가져오지 못했어요]`); }
  }
  // 동영상 등 → 링크
  body.querySelectorAll('iframe').forEach(f => { const u = f.getAttribute('data-original-src') || f.getAttribute('src') || ''; const a = document.createElement('a'); a.href = u; a.textContent = u; f.replaceWith(a); });
  return sanitize(body.innerHTML);
}
