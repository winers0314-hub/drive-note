// 저장소 두 가지 — 같은 함수를 가짐
//  DriveStore: 내 구글 드라이브의 "에스 노트" 폴더 (이 앱이 만든 파일만 접근: drive.file 권한)
//  LocalStore: 이 기기 브라우저 안 (체험용)
//
// 드라이브 폴더 구성
//  에스 노트/index.json       전자 필기장·섹션·페이지 목록
//  에스 노트/페이지/*.html     페이지 내용 (페이지마다 파일 하나)
//  에스 노트/첨부 파일/*       그림·첨부 파일
//  에스 노트/내보내기/*        한글·엑셀로 바꾼 파일

export const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
export const SECTION_COLORS = ['#529cca', '#4dab9a', '#ffa344', '#e255a1', '#9a6dd7', '#ff7369', '#937264', '#9b9a97']; // 노션 색

export function newIndex() {
  return {
    v: 1, updated: Date.now(),
    notebooks: [{ id: uid(), name: '내 전자 필기장', sections: [{ id: uid(), name: '빠른 노트', color: SECTION_COLORS[0], pages: [] }] }],
  };
}

export class AuthError extends Error { constructor() { super('auth'); this.auth = true; } }

// ---------- 이 기기 (IndexedDB) ----------
function openDB() {
  return new Promise((res, rej) => {
    const r = indexedDB.open('drive-note', 1);
    r.onupgradeneeded = () => ['kv', 'pages', 'files'].forEach(n => r.result.createObjectStore(n));
    r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
  });
}
function req(db, store, mode, fn) {
  return new Promise((res, rej) => {
    const t = db.transaction(store, mode), r = fn(t.objectStore(store));
    t.oncomplete = () => res(r ? r.result : undefined); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error);
  });
}
export class LocalStore {
  kind = 'local';
  async init() { this.db = await openDB(); }
  async loadIndex() { return (await req(this.db, 'kv', 'readonly', s => s.get('index'))) || null; }
  async saveIndex(ix) { await req(this.db, 'kv', 'readwrite', s => s.put(ix, 'index')); }
  async loadPage(p) { return (await req(this.db, 'pages', 'readonly', s => s.get(p.id))) || ''; }
  async savePage(p, html) { await req(this.db, 'pages', 'readwrite', s => s.put(html, p.id)); }
  async deletePage(p) { await req(this.db, 'pages', 'readwrite', s => s.delete(p.id)); }
  async putFile(blob, name) { const id = 'f_' + uid(); await req(this.db, 'files', 'readwrite', s => s.put({ blob, name }, id)); return id; }
  async getFile(id) { const r = await req(this.db, 'files', 'readonly', s => s.get(id)); return r ? r.blob : null; }
  async saveExport() { return null; } // 이 기기 모드에서는 내려받기만
  async remoteChanged() { return false; }
  async search(q, pages) {
    const hit = [], k = q.toLowerCase();
    for (const p of pages) { const h = await this.loadPage(p); if (h.replace(/<[^>]+>/g, ' ').toLowerCase().includes(k)) hit.push(p.id); }
    return hit;
  }
}

// ---------- 구글 드라이브 ----------
const API = 'https://www.googleapis.com/drive/v3', UP = 'https://www.googleapis.com/upload/drive/v3';
const SCOPE = 'https://www.googleapis.com/auth/drive.file';
const AUTH_API = 'api/auth'; // 이 사이트의 로그인 유지 서버 (구글 갱신 열쇠를 안전한 쿠키로 보관)
const q = s => s.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
const safeName = s => (s || '제목 없음').replace(/[\\/:*?"<>|]/g, ' ').trim().slice(0, 80) || '제목 없음';
const offlineErr = () => { const e = new Error('인터넷 연결이 끊겼어요.'); e.offline = true; return e; };

// 기기 안 보관함: 페이지·그림·목록을 저장해 두고 바로 열기, 인터넷이 끊겨도 보기·쓰기
function openCache() {
  return new Promise((res, rej) => {
    const r = indexedDB.open('es-note-cache', 1);
    r.onupgradeneeded = () => ['pages', 'files', 'kv'].forEach(n => r.result.createObjectStore(n));
    r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
  });
}
class Cache {
  constructor() { this.db = openCache().catch(() => null); }
  async get(store, key) { const db = await this.db; if (!db) return null; try { return (await req(db, store, 'readonly', s => s.get(key))) ?? null; } catch { return null; } }
  async put(store, key, val) { const db = await this.db; if (!db) return; try { await req(db, store, 'readwrite', s => s.put(val, key)); } catch {} }
  async del(store, key) { const db = await this.db; if (!db) return; try { await req(db, store, 'readwrite', s => s.delete(key)); } catch {} }
  async keys(store) { const db = await this.db; if (!db) return []; try { return await req(db, store, 'readonly', s => s.getAllKeys()); } catch { return []; } }
}

export class DriveStore {
  kind = 'drive';
  constructor(clientId) { this.clientId = clientId; this.token = null; this.exp = 0; this.cache = new Cache(); }

  // 로그인 유지 서버가 준비돼 있는지 (Vercel에 GOOGLE_CLIENT_SECRET이 있으면)
  async serverReady() {
    if (this._ready !== undefined) return this._ready;
    try { const r = await fetch(AUTH_API + '?health=1', { cache: 'no-store' }); this._ready = r.ok && (await r.json()).configured === true; }
    catch { this._ready = false; }
    return this._ready;
  }
  setToken(t, expiresIn) {
    this.token = t; this.exp = Date.now() + (expiresIn - 120) * 1000;
    try { localStorage.setItem('dn-token', JSON.stringify({ t: this.token, e: this.exp })); } catch {}
  }
  // 구글 로그인 창 — 버튼을 누를 때 불러야 팝업이 열림. 한 번 하면 다음부터는 저절로 로그인
  async signIn() {
    if (!window.google?.accounts?.oauth2) throw new Error('구글 로그인 스크립트를 불러오지 못했어요. 인터넷 연결을 확인해 주세요.');
    const errText = e => e.type === 'popup_closed' ? '로그인 창이 닫혔어요.' : e.type === 'popup_failed_to_open' ? '로그인 창이 열리지 않았어요. 팝업 차단을 풀어 주세요.' : '로그인하지 못했어요.';
    if (await this.serverReady()) {
      const code = await new Promise((res, rej) => google.accounts.oauth2.initCodeClient({
        client_id: this.clientId, scope: SCOPE, ux_mode: 'popup',
        callback: r => r.error ? rej(new Error(r.error_description || r.error)) : res(r),
        error_callback: e => rej(new Error(errText(e))),
      }).requestCode());
      if (!google.accounts.oauth2.hasGrantedAllScopes(code, SCOPE)) throw new Error('드라이브 사용 권한에 체크해 주셔야 노트를 저장할 수 있어요.');
      const r = await fetch(AUTH_API, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ code: code.code }) });
      const j = await r.json().catch(() => ({}));
      if (!r.ok) throw new Error('로그인 처리 오류: ' + (j.detail || j.error || r.status));
      this.setToken(j.access_token, j.expires_in);
      this.noRefresh = !!j.noRefresh;
      return;
    }
    // 서버 설정 전: 예전 방식 (1시간마다 다시 로그인)
    await new Promise((res, rej) => google.accounts.oauth2.initTokenClient({
      client_id: this.clientId, scope: SCOPE,
      callback: r => {
        if (r.error) return rej(new Error(r.error_description || r.error));
        if (!google.accounts.oauth2.hasGrantedAllScopes(r, SCOPE)) return rej(new Error('드라이브 사용 권한에 체크해 주셔야 노트를 저장할 수 있어요.'));
        this.setToken(r.access_token, r.expires_in); res();
      },
      error_callback: e => rej(new Error(errText(e))),
    }).requestAccessToken({ prompt: '' }));
  }
  // 저장해 둔 출입증(1시간짜리)이 있으면 바로 씀
  restore() {
    try { const o = JSON.parse(localStorage.getItem('dn-token')); if (o && o.e > Date.now()) { this.token = o.t; this.exp = o.e; return true; } } catch {}
    return false;
  }
  // 출입증이 끝났으면 로그인 유지 서버에서 새로 받아 옴 (창·팝업 없이)
  async refresh() {
    let r;
    try { r = await fetch(AUTH_API, { cache: 'no-store', credentials: 'same-origin' }); } catch { throw offlineErr(); }
    if (r.status === 401 || r.status === 404 || r.status === 503) throw new AuthError();
    if (!r.ok) throw new Error('로그인 갱신 오류 ' + r.status);
    const j = await r.json(); this.setToken(j.access_token, j.expires_in);
  }
  async getToken() {
    if (this.token && Date.now() < this.exp) return this.token;
    this._refreshing ||= this.refresh().finally(() => { this._refreshing = null; });
    await this._refreshing;
    return this.token;
  }
  get signedIn() { return !!this.token && Date.now() < this.exp; }
  async signOut() {
    try { await fetch(AUTH_API, { method: 'DELETE' }); } catch {}
    try { if (this.token) google.accounts.oauth2.revoke(this.token, () => {}); } catch {}
    this.token = null; try { localStorage.removeItem('dn-token'); localStorage.removeItem('dn-dirs'); } catch {}
  }

  async api(url, opt = {}, raw, retried) {
    const token = await this.getToken();
    let r;
    try { r = await fetch(url, { ...opt, headers: { ...(opt.headers || {}), Authorization: 'Bearer ' + token } }); }
    catch { throw offlineErr(); }
    if (r.status === 401) {
      this.token = null; try { localStorage.removeItem('dn-token'); } catch {}
      if (!retried) return this.api(url, opt, raw, true); // 서버에서 새 출입증 받아 한 번 더
      throw new AuthError();
    }
    if (!r.ok) { const t = await r.text(); const e = new Error(`드라이브 오류 ${r.status}: ${t.slice(0, 160)}`); e.status = r.status; throw e; }
    if (raw) return r;
    return r.status === 204 ? null : r.json();
  }
  async list(query, fields = 'files(id,name,modifiedTime)') {
    return (await this.api(`${API}/files?q=${encodeURIComponent(query)}&fields=${encodeURIComponent(fields)}&spaces=drive&pageSize=100`)).files;
  }
  async folder(name, parent) {
    const f = await this.list(`name='${q(name)}' and mimeType='application/vnd.google-apps.folder' and trashed=false and '${parent || 'root'}' in parents`);
    if (f.length) return f[0].id;
    return (await this.api(`${API}/files?fields=id`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name, mimeType: 'application/vnd.google-apps.folder', parents: [parent || 'root'] }) })).id;
  }
  // 새 파일(fileId 없음) 또는 내용 바꾸기(fileId 있음)
  upload(meta, blob, fileId) {
    if (blob.size > 4.5e6) return this.uploadLarge(meta, blob, fileId); // 한 번에 보내기는 5MB까지
    const b = 'dn' + uid();
    const body = new Blob([`--${b}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n--${b}\r\nContent-Type: ${blob.type || 'application/octet-stream'}\r\n\r\n`, blob, `\r\n--${b}--`]);
    const url = `${UP}/files${fileId ? '/' + fileId : ''}?uploadType=multipart&fields=id,modifiedTime`;
    return this.api(url, { method: fileId ? 'PATCH' : 'POST', headers: { 'Content-Type': `multipart/related; boundary=${b}` }, body });
  }
  // 큰 파일: 업로드 주소를 먼저 받고 내용을 보냄
  async uploadLarge(meta, blob, fileId) {
    const r = await this.api(`${UP}/files${fileId ? '/' + fileId : ''}?uploadType=resumable&fields=id,modifiedTime`, {
      method: fileId ? 'PATCH' : 'POST',
      headers: { 'Content-Type': 'application/json; charset=UTF-8', 'X-Upload-Content-Type': blob.type || 'application/octet-stream' },
      body: JSON.stringify(meta),
    }, true);
    let put;
    try { put = await fetch(r.headers.get('Location'), { method: 'PUT', body: blob }); } catch { throw offlineErr(); }
    if (!put.ok) { const e = new Error(`드라이브 업로드 오류 ${put.status}`); e.status = put.status; throw e; }
    return put.json();
  }
  async download(id) { return (await this.api(`${API}/files/${id}?alt=media`, {}, true)).blob(); }

  // 앱 폴더: "에스 노트" (예전 이름 "드라이브 노트" 폴더가 있으면 이름만 바꿔서 그대로 씀)
  async rootFolder() {
    const find = name => this.list(`name='${q(name)}' and mimeType='application/vnd.google-apps.folder' and trashed=false and 'root' in parents`);
    const now = await find('에스 노트');
    if (now.length) return now[0].id;
    const old = await find('드라이브 노트');
    if (old.length) {
      await this.api(`${API}/files/${old[0].id}?fields=id`, { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name: '에스 노트' }) });
      return old[0].id;
    }
    return this.folder('에스 노트');
  }
  // 폴더 위치는 한 번 찾으면 기억 (켤 때마다 찾지 않아서 빨리 열림)
  async init() {
    try { const d = JSON.parse(localStorage.getItem('dn-dirs') || 'null'); if (d && d.root) { Object.assign(this, d); return; } } catch {}
    this.root = await this.rootFolder();
    [this.pagesDir, this.filesDir] = await Promise.all([this.folder('페이지', this.root), this.folder('첨부 파일', this.root)]);
    try { localStorage.setItem('dn-dirs', JSON.stringify({ root: this.root, pagesDir: this.pagesDir, filesDir: this.filesDir })); } catch {}
  }
  // 목록: 인터넷이 끊겼으면 기기에 보관한 마지막 목록으로
  async loadIndex() {
    let f;
    try { f = (await this.list(`name='index.json' and '${this.root}' in parents and trashed=false`))[0]; }
    catch (e) {
      if (e.status === 404) { localStorage.removeItem('dn-dirs'); throw e; } // 폴더가 지워졌으면 다시 찾기
      const c = await this.cache.get('kv', 'index'); if (c && (e.offline || e.auth)) { this.offline = true; this.indexId = c.id; this.indexTime = c.time; return c.ix; }
      throw e;
    }
    this.offline = false;
    if (!f) return null;
    const c = await this.cache.get('kv', 'index');
    if (c && c.id === f.id && c.time === f.modifiedTime) { this.indexId = f.id; this.indexTime = f.modifiedTime; return c.ix; }
    this.indexId = f.id; this.indexTime = f.modifiedTime;
    const ix = JSON.parse(await (await this.download(f.id)).text());
    this.cache.put('kv', 'index', { id: f.id, time: f.modifiedTime, ix });
    return ix;
  }
  async saveIndex(ix) {
    this.cache.put('kv', 'index', { id: this.indexId, time: this.indexTime, ix }); // 먼저 기기에
    const blob = new Blob([JSON.stringify(ix)], { type: 'application/json' });
    const r = await this.upload(this.indexId ? { name: 'index.json' } : { name: 'index.json', parents: [this.root] }, blob, this.indexId);
    this.indexId = r.id; this.indexTime = r.modifiedTime;
    this.cache.put('kv', 'index', { id: r.id, time: r.modifiedTime, ix });
  }
  // 다른 기기에서 목록이 바뀌었는지
  async remoteChanged() {
    if (!this.indexId) return false;
    const r = await this.api(`${API}/files/${this.indexId}?fields=modifiedTime`);
    return r.modifiedTime !== this.indexTime;
  }
  // 페이지: 기기에 보관한 것이 최신이면 바로 (드라이브에서 받지 않음)
  async loadPage(p) {
    const c = await this.cache.get('pages', p.id);
    if (c && (c.pending || (c.updated >= (p.updated || 0) && c.fid === p.fid))) return c.html;
    if (!p.fid) return c ? c.html : '';
    try {
      const html = await (await this.download(p.fid)).text();
      this.cache.put('pages', p.id, { html, updated: p.updated || 0, fid: p.fid });
      return html;
    } catch (e) { if (c && (e.offline || e.auth)) return c.html; throw e; }
  }
  cachedPage(p) { return this.cache.get('pages', p.id); }
  // 저장: 먼저 기기에 (인터넷이 끊겨도 남음) → 드라이브에 올리면 "보냄" 표시
  async savePage(p, html) {
    await this.cache.put('pages', p.id, { html, updated: p.updated || Date.now(), fid: p.fid, pending: true, title: p.title });
    const name = safeName(p.title) + '.html', blob = new Blob([html], { type: 'text/html' });
    const r = await this.upload(p.fid ? { name } : { name, parents: [this.pagesDir], appProperties: { pageId: p.id } }, blob, p.fid);
    p.fid = r.id;
    await this.cache.put('pages', p.id, { html, updated: p.updated || Date.now(), fid: p.fid });
  }
  // 인터넷이 끊겼을 때 써 둔 페이지들을 다시 올림. findPage(id) → 페이지 정보
  async flushPending(findPage) {
    let n = 0;
    for (const id of await this.cache.keys('pages')) {
      const c = await this.cache.get('pages', id); if (!c || !c.pending) continue;
      const p = findPage(id); if (!p) { await this.cache.del('pages', id); continue; }
      await this.savePage(p, c.html); n++;
    }
    return n;
  }
  // 지우면 드라이브 휴지통으로 (30일 안에 되살릴 수 있음)
  async trash(id) { if (id) await this.api(`${API}/files/${id}?fields=id`, { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ trashed: true }) }); }
  async deletePage(p) { this.cache.del('pages', p.id); await this.trash(p.fid); }
  async putFile(blob, name) {
    const id = (await this.upload({ name: safeName(name), parents: [this.filesDir] }, blob)).id;
    if (blob.size < 25e6) this.cache.put('files', id, blob);
    return id;
  }
  // 그림·첨부: 한 번 받은 건 기기에 보관 (파일 내용은 바뀌지 않음)
  async getFile(id) {
    const c = await this.cache.get('files', id); if (c) return c;
    const b = await this.download(id);
    if (b.size < 25e6) this.cache.put('files', id, b);
    return b;
  }
  async saveExport(blob, name, prevId) {
    this.exportDir ||= await this.folder('내보내기', this.root);
    try { return (await this.upload(prevId ? { name } : { name, parents: [this.exportDir] }, blob, prevId)).id; }
    catch (e) { if (prevId && e.status === 404) return (await this.upload({ name, parents: [this.exportDir] }, blob)).id; throw e; }
  }
  async search(text) {
    const f = await this.list(`fullText contains '${q(text)}' and '${this.pagesDir}' in parents and trashed=false`, 'files(id,appProperties)');
    return f.map(x => x.appProperties?.pageId).filter(Boolean);
  }
}
