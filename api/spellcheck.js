// 맞춤법 검사 API — 페이지의 문단들을 받아 맞춤법·띄어쓰기·문법 오류를 찾는다. (OpenAI / ChatGPT)
//
// Vercel 환경변수
//   OPENAI_API_KEY   OpenAI API 키 (없으면 503 not-configured → 앱이 안내)
//   SPELL_AI_MODEL   (선택) 쓸 모델 이름
//   ALLOWED_EMAILS   (선택) 쓸 수 있는 구글 계정, 쉼표로 여러 개. 비우면 이 앱으로 로그인한 계정 누구나
//   GOOGLE_CLIENT_ID (선택) 앱의 구글 클라이언트 ID (config.js와 같은 값)
// 인증: 앱이 받은 구글 로그인 토큰을 확인해서, 다른 사람이 API 비용을 쓰지 못하게 함

const API_KEY = process.env.OPENAI_API_KEY;
const CLIENT_ID = process.env.GOOGLE_CLIENT_ID || '815675421175-eg16pbjm5jvbtv1rr5ot3gda8uqt27mu.apps.googleusercontent.com';
const ALLOWED = (process.env.ALLOWED_EMAILS || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
const MODEL_CANDIDATES = [process.env.SPELL_AI_MODEL, 'gpt-5-mini', 'gpt-4.1-mini', 'gpt-4o-mini', 'gpt-4.1'].filter(Boolean);
const UNAVAILABLE = new Set();
const MAX_CHARS = 6000;

// 구글 토큰이 이 앱에서 발급된 것인지, 허용된 계정인지
async function checkUser(token) {
  if (!token) return null;
  try {
    const info = await (await fetch('https://oauth2.googleapis.com/tokeninfo?access_token=' + encodeURIComponent(token))).json();
    if (!info || info.error || (info.aud !== CLIENT_ID && info.azp !== CLIENT_ID)) return null;
    if (!ALLOWED.length) return { email: '' };
    const about = await (await fetch('https://www.googleapis.com/drive/v3/about?fields=user(emailAddress)', { headers: { Authorization: 'Bearer ' + token } })).json();
    const email = String(about && about.user && about.user.emailAddress || '').toLowerCase();
    return ALLOWED.includes(email) ? { email } : null;
  } catch { return null; }
}

async function generateAI(messages) {
  let lastErr = '';
  for (const model of MODEL_CANDIDATES) {
    if (UNAVAILABLE.has(model)) continue;
    let resp;
    try {
      resp = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${API_KEY}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ model, max_completion_tokens: 4000, response_format: { type: 'json_object' }, messages }),
      });
    } catch (e) { lastErr = String(e && e.message || e); continue; }
    if (resp.ok) {
      const data = await resp.json();
      const text = (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content || '').trim();
      if (text) return { model, text };
      lastErr = 'empty-response'; continue;
    }
    const errText = await resp.text();
    lastErr = `${resp.status}: ${errText}`.slice(0, 300);
    if (resp.status === 404 || /model.*(not|does not) exist|do not have access|unknown model/i.test(errText)) UNAVAILABLE.add(model);
  }
  return { error: lastErr || 'all-models-failed' };
}

const SYSTEM = `당신은 한국어 맞춤법 검사기입니다. 사용자의 노트 문단들에서 맞춤법, 띄어쓰기, 명백한 문법 오류만 찾습니다.

규칙:
- 입력은 {"p":[문단 문자열, ...]} 형식이고, 문단 번호는 0부터 셉니다.
- 오류마다 {"p": 문단 번호, "wrong": "틀린 부분", "right": "고친 표현", "why": "짧은 이유"}를 만듭니다.
- "wrong"은 그 문단에 글자 그대로 들어 있는 부분이어야 합니다. 틀린 낱말 앞뒤로 한두 낱말을 붙여 그 문단에서 한 번만 나오게 합니다.
- "why"는 15자 안팎으로 씁니다. 예: "띄어쓰기", "맞춤법(되/돼)", "조사 오류".
- 내용, 말투, 어순은 바꾸지 않습니다. 구어체·줄임말·이름·전문 용어·영어·숫자·기호는 틀린 것으로 보지 않습니다.
- 확실한 오류만 냅니다. 오류가 없으면 빈 목록을 냅니다.
- 반드시 JSON 하나로만 답합니다: {"items":[...]}`;

module.exports = async (req, res) => {
  const configured = !!API_KEY;
  if (req.method === 'GET' && (req.query || {}).health === '1') return res.status(200).json({ ok: true, configured });
  if (!configured) return res.status(503).json({ error: 'not-configured' });
  if (req.method !== 'POST') return res.status(405).json({ error: 'method-not-allowed' });
  try {
    const token = String(req.headers.authorization || '').replace(/^Bearer\s+/i, '');
    if (!(await checkUser(token))) return res.status(401).json({ error: 'unauthorized' });
    const b = typeof req.body === 'string' ? JSON.parse(req.body) : (req.body || {});
    const paras = Array.isArray(b.p) ? b.p.map(x => String(x || '').slice(0, 2000)) : [];
    if (!paras.length) return res.status(200).json({ ok: true, items: [] });
    if (paras.join('').length > MAX_CHARS) return res.status(413).json({ error: 'too-long' });
    const out = await generateAI([{ role: 'system', content: SYSTEM }, { role: 'user', content: JSON.stringify({ p: paras }) }]);
    if (out.error) return res.status(502).json({ error: 'ai-error', detail: out.error });
    let items = [];
    try { items = JSON.parse(out.text).items || []; } catch { return res.status(502).json({ error: 'ai-bad-json' }); }
    // 실제로 그 문단에 있는 부분만 돌려줌
    items = items.filter(it => it && Number.isInteger(it.p) && paras[it.p] && typeof it.wrong === 'string' && it.wrong && typeof it.right === 'string' && it.wrong !== it.right && paras[it.p].includes(it.wrong))
      .map(it => ({ p: it.p, wrong: it.wrong, right: it.right, why: String(it.why || '').slice(0, 40) }));
    return res.status(200).json({ ok: true, items, model: out.model });
  } catch (e) {
    return res.status(500).json({ error: 'server', detail: String(e && e.message || e).slice(0, 300) });
  }
};
