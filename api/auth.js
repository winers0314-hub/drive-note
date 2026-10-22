// 로그인 유지 서버 — 구글 "갱신 열쇠(refresh token)"를 암호화해서 이 사이트 전용 쿠키에 보관하고,
// 앱이 켜질 때마다 새 출입증(1시간짜리 access token)을 내준다. 그래서 껐다 켜도 다시 로그인하지 않아도 됨.
//
// Vercel 환경변수
//   GOOGLE_CLIENT_SECRET  구글 클라우드 → 클라이언트 → "클라이언트 보안 비밀번호" (없으면 503 → 앱은 예전 방식으로 로그인)
//   GOOGLE_CLIENT_ID      (선택) config.js와 같은 값
//   COOKIE_SECRET         (선택) 쿠키 암호화 열쇠. 없으면 클라이언트 보안 비밀번호로 만듦
//
//   POST   { code }  처음 로그인: 구글 코드 → 출입증 + 갱신 열쇠(쿠키)
//   GET              다시 로그인: 쿠키의 갱신 열쇠 → 새 출입증
//   DELETE           로그아웃: 갱신 열쇠 폐기, 쿠키 지움

const crypto = require('crypto');
const CLIENT_ID = process.env.GOOGLE_CLIENT_ID || '815675421175-eg16pbjm5jvbtv1rr5ot3gda8uqt27mu.apps.googleusercontent.com';
const SECRET = process.env.GOOGLE_CLIENT_SECRET;
const KEY = crypto.createHash('sha256').update('es-note:' + (process.env.COOKIE_SECRET || SECRET || '')).digest();
const COOKIE = 'esn_rt';
const MAX_AGE = 60 * 60 * 24 * 400; // 400일

const seal = text => {
  const iv = crypto.randomBytes(12), c = crypto.createCipheriv('aes-256-gcm', KEY, iv);
  const enc = Buffer.concat([c.update(text, 'utf8'), c.final()]);
  return Buffer.concat([iv, c.getAuthTag(), enc]).toString('base64url');
};
const open = v => {
  try {
    const b = Buffer.from(v, 'base64url'), d = crypto.createDecipheriv('aes-256-gcm', KEY, b.subarray(0, 12));
    d.setAuthTag(b.subarray(12, 28));
    return Buffer.concat([d.update(b.subarray(28)), d.final()]).toString('utf8');
  } catch { return null; }
};
const readCookie = req => (String(req.headers.cookie || '').split(/;\s*/).find(c => c.startsWith(COOKIE + '=')) || '').slice(COOKIE.length + 1);
const setCookie = (res, val, age) => res.setHeader('Set-Cookie', `${COOKIE}=${val}; Path=/api/auth; Max-Age=${age}; HttpOnly; Secure; SameSite=Strict`);

async function google(params) {
  const r = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: CLIENT_ID, client_secret: SECRET, ...params }),
  });
  return { ok: r.ok, j: await r.json().catch(() => ({})) };
}

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method === 'GET' && (req.query || {}).health === '1') return res.status(200).json({ ok: true, configured: !!SECRET });
  if (!SECRET) return res.status(503).json({ error: 'not-configured' });
  try {
    if (req.method === 'POST') {
      const b = typeof req.body === 'string' ? JSON.parse(req.body) : (req.body || {});
      if (!b.code) return res.status(400).json({ error: 'no-code' });
      const { ok, j } = await google({ code: String(b.code), grant_type: 'authorization_code', redirect_uri: 'postmessage' });
      if (!ok || !j.access_token) return res.status(400).json({ error: 'exchange-failed', detail: j.error_description || j.error || '' });
      if (j.refresh_token) setCookie(res, seal(j.refresh_token), MAX_AGE);
      return res.status(200).json({ access_token: j.access_token, expires_in: j.expires_in || 3600, noRefresh: !j.refresh_token && !readCookie(req) });
    }
    if (req.method === 'GET') {
      const rt = open(readCookie(req));
      if (!rt) return res.status(401).json({ error: 'no-session' });
      const { ok, j } = await google({ refresh_token: rt, grant_type: 'refresh_token' });
      if (!ok || !j.access_token) {
        if (j.error === 'invalid_grant') setCookie(res, '', 0); // 열쇠가 끝났거나 취소됨 → 다시 로그인
        return res.status(401).json({ error: 'refresh-failed', detail: j.error || '' });
      }
      if (j.refresh_token) setCookie(res, seal(j.refresh_token), MAX_AGE);
      return res.status(200).json({ access_token: j.access_token, expires_in: j.expires_in || 3600 });
    }
    if (req.method === 'DELETE') {
      const rt = open(readCookie(req));
      if (rt) await fetch('https://oauth2.googleapis.com/revoke?token=' + encodeURIComponent(rt), { method: 'POST' }).catch(() => {});
      setCookie(res, '', 0);
      return res.status(200).json({ ok: true });
    }
    return res.status(405).json({ error: 'method-not-allowed' });
  } catch (e) {
    return res.status(500).json({ error: 'server', detail: String(e && e.message || e).slice(0, 200) });
  }
};
