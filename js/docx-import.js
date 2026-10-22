// 원노트에서 내보낸 Word 파일(.docx) → 에스 노트 페이지들
// 원노트는 섹션을 Word로 내보낼 때 페이지마다 쪽을 나누고, 맨 앞에 제목과 날짜를 넣음
import { parseHTML, sanitize } from './doc.js';

const MAMMOTH = 'https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.8.0/mammoth.browser.min.js';
const MARK = '⁣쪽나눔⁣';
let loading = null;
function loadMammoth() {
  if (window.mammoth) return Promise.resolve();
  return loading ||= new Promise((res, rej) => {
    const s = document.createElement('script'); s.src = MAMMOTH;
    s.onload = () => res(); s.onerror = () => { loading = null; rej(new Error('Word 파일을 읽는 도구를 불러오지 못했어요. 인터넷 연결을 확인해 주세요.')); };
    document.head.append(s);
  });
}

// 날짜 줄: "2024년 3월 5일 화요일  오후 3:21" / "Tuesday, March 5, 2024 3:21 PM"
function parseDateLine(t) {
  t = t.replace(/\s+/g, ' ').trim();
  let m = t.match(/^(\d{4})년 ?(\d{1,2})월 ?(\d{1,2})일(?: ?[가-힣]요일)?(?: ?(오전|오후) ?(\d{1,2}):(\d{2}))?$/);
  if (m) { let h = +(m[5] || 0); if (m[4] === '오후' && h < 12) h += 12; if (m[4] === '오전' && h === 12) h = 0; return new Date(+m[1], m[2] - 1, +m[3], h, +(m[6] || 0)).getTime(); }
  if (t.length < 60 && /\d{4}/.test(t) && !isNaN(Date.parse(t))) return Date.parse(t);
  return 0;
}
const b64Blob = (b64, type) => { const bin = atob(b64), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return new Blob([u], { type }); };

// file → [{title, created, html}]
export async function docxToPages(file, putFile, onImage) {
  await loadMammoth();
  const M = window.mammoth;
  const res = await M.convertToHtml({ arrayBuffer: await file.arrayBuffer() }, {
    // 쪽 나눔을 표시 글자로 바꿔 두었다가 나중에 페이지를 나눔
    transformDocument: M.transforms.run(run => run.children.some(c => c.type === 'break' && c.breakType === 'page')
      ? { ...run, children: run.children.map(c => c.type === 'break' && c.breakType === 'page' ? { type: 'text', value: MARK } : c) } : run),
    convertImage: M.images.imgElement(async image => {
      const type = image.contentType || 'image/png';
      const blob = b64Blob(await image.read('base64'), type);
      onImage?.();
      const fid = await putFile(blob, 'image.' + (type.split('/')[1] || 'png').replace('jpeg', 'jpg'));
      return { src: 'fid:' + fid };
    }),
    styleMap: ["p[style-name='Title'] => h1:fresh", "p[style-name='제목'] => h1:fresh"],
  });

  // 페이지 나누기
  const body = parseHTML(res.value), groups = [[]];
  const strip = el => { const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT); let n; while ((n = w.nextNode())) if (n.nodeValue.includes(MARK)) n.nodeValue = n.nodeValue.split(MARK).join(''); };
  for (const el of [...body.childNodes]) {
    const t = el.textContent || '';
    if (!t.includes(MARK)) { groups[groups.length - 1].push(el); continue; }
    const before = t.slice(0, t.indexOf(MARK)).trim();
    strip(el);
    const keep = (el.textContent || '').trim() || el.querySelector?.('img, table');
    if (before) { if (keep) groups[groups.length - 1].push(el); groups.push([]); }
    else { groups.push(keep ? [el] : []); }
  }

  const pages = [];
  for (const g of groups) {
    const box = document.createElement('div'); box.append(...g);
    const blocks = [...box.children];
    if (!box.textContent.trim() && !box.querySelector('img, table')) continue;
    // 맨 앞 줄 = 제목, 그다음 줄이 날짜면 만든 날짜로
    let title = '', created = 0;
    const first = blocks.find(b => b.textContent.trim() || b.querySelector('img'));
    if (first && !first.querySelector('img, table') && first.textContent.trim().length <= 120) {
      title = first.textContent.replace(/\s+/g, ' ').trim();
      const next = first.nextElementSibling;
      if (next && (created = parseDateLine(next.textContent))) next.remove();
      else if (next) { const nn = next.textContent.trim(); if (!nn) { const nx = next.nextElementSibling; if (nx && (created = parseDateLine(nx.textContent))) { nx.remove(); } } }
      first.remove();
    }
    box.querySelectorAll('img[src^="fid:"]').forEach(img => { img.setAttribute('data-fid', img.getAttribute('src').slice(4)); img.removeAttribute('src'); });
    // 앞뒤 빈 문단 정리
    while (box.firstElementChild && !box.firstElementChild.textContent.trim() && !box.firstElementChild.querySelector('img, table')) box.firstElementChild.remove();
    pages.push({ title, created, html: sanitize(box.innerHTML) });
  }
  if (!pages.length) pages.push({ title: file.name.replace(/\.docx$/i, ''), created: 0, html: '' });
  if (pages.length === 1 && !pages[0].title) pages[0].title = file.name.replace(/\.docx$/i, '');
  return pages;
}
