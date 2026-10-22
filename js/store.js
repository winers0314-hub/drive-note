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
const q = s => s.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
const safeName = s => (s || '제목 없음').replace(/[\\/:*?"<>|]/g, ' ').trim().slice(0, 80) || '제목 없음';

export class DriveStore {
  kind = 'drive';
  constructor(clientId) { this.clientId = clientId; this.token = null; this.exp = 0; }

  // 구글 로그인 창 — 반드시 버튼을 누를 때 불러야 팝업이 열림
  signIn() {
    return new Promise((res, rej) => {
      if (!window.google?.accounts?.oauth2) return rej(new Error('구글 로그인 스크립트를 불러오지 못했어요. 인터넷 연결을 확인해 주세요.'));
      const tc = google.accounts.oauth2.initTokenClient({
        client_id: this.clientId, scope: SCOPE,
        callback: r => {
          if (r.error) return rej(new Error(r.error_description || r.error));
          if (!google.accounts.oauth2.hasGrantedAllScopes(r, SCOPE)) return rej(new Error('드라이브 사용 권한에 체크해 주셔야 노트를 저장할 수 있어요.'));
          this.token = r.access_token; this.exp = Date.now() + (r.expires_in - 120) * 1000;
          try { sessionStorage.setItem('dn-token', JSON.stringify({ t: this.token, e: this.exp })); } catch {}
          res();
        },
        error_callback: e => rej(new Error(e.type === 'popup_closed' ? '로그인 창이 닫혔어요.' : e.type === 'popup_failed_to_open' ? '로그인 창이 열리지 않았어요. 팝업 차단을 풀어 주세요.' : '로그인하지 못했어요.')),
      });
      tc.requestAccessToken({ prompt: '' });
    });
  }
  restore() {
    try { const o = JSON.parse(sessionStorage.getItem('dn-token')); if (o && o.e > Date.now()) { this.token = o.t; this.exp = o.e; return true; } } catch {}
    return false;
  }
  get signedIn() { return !!this.token && Date.now() < this.exp; }
  signOut() {
    try { if (this.token) google.accounts.oauth2.revoke(this.token, () => {}); } catch {}
    this.token = null; try { sessionStorage.removeItem('dn-token'); } catch {}
  }

  async api(url, opt = {}, raw) {
    if (!this.signedIn) throw new AuthError();
    let r;
    try { r = await fetch(url, { ...opt, headers: { ...(opt.headers || {}), Authorization: 'Bearer ' + this.token } }); }
    catch { const e = new Error('인터넷 연결이 끊겼어요.'); e.offline = true; throw e; }
    if (r.status === 401) { this.token = null; throw new AuthError(); }
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
    const put = await fetch(r.headers.get('Location'), { method: 'PUT', body: blob });
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
  async init() {
    this.root = await this.rootFolder();
    [this.pagesDir, this.filesDir] = await Promise.all([this.folder('페이지', this.root), this.folder('첨부 파일', this.root)]);
  }
  async loadIndex() {
    const f = (await this.list(`name='index.json' and '${this.root}' in parents and trashed=false`))[0];
    if (!f) return null;
    this.indexId = f.id; this.indexTime = f.modifiedTime;
    return JSON.parse(await (await this.download(f.id)).text());
  }
  async saveIndex(ix) {
    const blob = new Blob([JSON.stringify(ix)], { type: 'application/json' });
    const r = await this.upload(this.indexId ? { name: 'index.json' } : { name: 'index.json', parents: [this.root] }, blob, this.indexId);
    this.indexId = r.id; this.indexTime = r.modifiedTime;
  }
  // 다른 기기에서 목록이 바뀌었는지
  async remoteChanged() {
    if (!this.indexId) return false;
    const r = await this.api(`${API}/files/${this.indexId}?fields=modifiedTime`);
    return r.modifiedTime !== this.indexTime;
  }
  async loadPage(p) { return p.fid ? (await this.download(p.fid)).text() : ''; }
  async savePage(p, html) {
    const name = safeName(p.title) + '.html', blob = new Blob([html], { type: 'text/html' });
    const r = await this.upload(p.fid ? { name } : { name, parents: [this.pagesDir], appProperties: { pageId: p.id } }, blob, p.fid);
    p.fid = r.id;
  }
  // 지우면 드라이브 휴지통으로 (30일 안에 되살릴 수 있음)
  async trash(id) { if (id) await this.api(`${API}/files/${id}?fields=id`, { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ trashed: true }) }); }
  async deletePage(p) { await this.trash(p.fid); }
  async putFile(blob, name) { return (await this.upload({ name: safeName(name), parents: [this.filesDir] }, blob)).id; }
  async getFile(id) { return this.download(id); }
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
