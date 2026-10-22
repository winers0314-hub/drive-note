// 노트 HTML을 읽는 공용 함수 (한글·엑셀 변환, 붙여넣기 정리에서 함께 씀)

export function parseHTML(html) {
  return new DOMParser().parseFromString(`<body>${html || ''}</body>`, 'text/html').body;
}

const ALIGN = { left: 'LEFT', center: 'CENTER', right: 'RIGHT', justify: 'JUSTIFY' };
export function blockAlign(el) {
  const m = (el.getAttribute('style') || '').match(/text-align:\s*(\w+)/) || [, el.getAttribute('align')];
  return ALIGN[(m[1] || '').toLowerCase()] || '';
}

// 색 표기 통일: rgb(…) / #abc → #AABBCC
export function hexColor(c) {
  if (!c) return '';
  c = c.trim().toLowerCase();
  if (c === 'transparent' || c === 'inherit' || c === 'initial') return '';
  let m = c.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)$/);
  if (m) { if (m[4] !== undefined && +m[4] === 0) return ''; return '#' + [m[1], m[2], m[3]].map(x => (+x).toString(16).padStart(2, '0')).join('').toUpperCase(); }
  m = c.match(/^#([0-9a-f]{3})$/); if (m) return '#' + m[1].split('').map(x => x + x).join('').toUpperCase();
  m = c.match(/^#([0-9a-f]{6})$/); if (m) return '#' + m[1].toUpperCase();
  const NAMED = { black: '#000000', red: '#FF0000', blue: '#0000FF', green: '#008000', yellow: '#FFFF00', white: '#FFFFFF', gray: '#808080', orange: '#FFA500' };
  return NAMED[c] || '';
}
const SIZE_KW = { 'x-small': 7.5, small: 10, medium: 12, large: 13.5, 'x-large': 18, 'xx-large': 24, 'xxx-large': 36 };
const HEADING = { h1: 20, h2: 16, h3: 13, h4: 12, h5: 11, h6: 10 };

// 요소가 더하는 글자 모양 → {b,i,u,s,color,shade,size(pt)}
export function styleOf(el, st) {
  const tag = el.tagName.toLowerCase(), s = { ...st };
  if (tag === 'b' || tag === 'strong') s.b = true;
  if (tag === 'i' || tag === 'em') s.i = true;
  if (tag === 'u' || tag === 'ins') s.u = true;
  if (tag === 's' || tag === 'strike' || tag === 'del') s.s = true;
  if (tag === 'mark') s.shade = '#FFF176';
  if (tag === 'a') { s.u = true; s.color = s.color || '#1F5FBF'; }
  if (HEADING[tag]) { s.size = HEADING[tag]; s.b = true; }
  if (tag === 'font' && el.getAttribute('color')) s.color = hexColor(el.getAttribute('color')) || s.color;
  const css = el.getAttribute('style') || '';
  for (const decl of css.split(';')) {
    const i = decl.indexOf(':'); if (i < 0) continue;
    const k = decl.slice(0, i).trim().toLowerCase(), v = decl.slice(i + 1).trim().toLowerCase();
    if (k === 'font-weight') s.b = v === 'bold' || +v >= 600;
    else if (k === 'font-style') s.i = v === 'italic';
    else if (k === 'text-decoration' || k === 'text-decoration-line') { if (v.includes('underline')) s.u = true; if (v.includes('line-through')) s.s = true; }
    else if (k === 'color') s.color = hexColor(v) || s.color;
    else if (k === 'background-color' || k === 'background') { const h = hexColor(v); if (h) s.shade = h; }
    else if (k === 'font-size') {
      let m;
      if ((m = v.match(/^([\d.]+)pt$/))) s.size = +m[1];
      else if ((m = v.match(/^([\d.]+)px$/))) s.size = Math.round(+m[1] * 0.75 * 2) / 2;
      else if (SIZE_KW[v]) s.size = SIZE_KW[v];
    }
  }
  return s;
}

// 엑셀용: 페이지 내용을 줄 단위로 (표는 따로 꺼냄)
export function toLines(root) {
  const lines = [], tables = [];
  let cur = '';
  const end = () => { if (cur.trim()) lines.push(cur.replace(/\s+$/, '')); cur = ''; };
  const walk = (node, prefix) => {
    for (const n of node.childNodes) {
      if (n.nodeType === 3) { cur += n.nodeValue.replace(/[\r\n]+/g, '').replace(/ /g, ' '); continue; }
      if (n.nodeType !== 1) continue;
      const tag = n.tagName.toLowerCase();
      if (tag === 'br') { end(); continue; }
      if (tag === 'img') { end(); lines.push('[그림]'); continue; }
      if (tag === 'hr') { end(); continue; }
      if (n.classList.contains('att')) { cur += '[첨부] ' + (n.dataset.name || n.textContent.trim()); continue; }
      if (tag === 'table') { end(); tables.push(tableRows(n)); lines.push(`[표 ${tables.length}]`); continue; }
      if (tag === 'ul' || tag === 'ol') {
        end();
        const ordered = tag === 'ol', check = n.classList.contains('checklist');
        let k = ordered ? (parseInt(n.getAttribute('start'), 10) || 1) : 0;
        for (const li of n.children) {
          if (li.tagName !== 'LI') continue;
          cur = prefix + (check ? (li.hasAttribute('data-checked') ? '☑ ' : '☐ ') : ordered ? (k++) + '. ' : '• ');
          walk(li, prefix + '    '); end();
        }
        continue;
      }
      if (/^(p|div|h[1-6]|blockquote|pre|li)$/.test(tag)) { end(); walk(n, prefix); end(); continue; }
      walk(n, prefix);
    }
  };
  walk(root, '');
  end();
  return { lines, tables };
}
export function tableRows(table) {
  return [...table.querySelectorAll(':scope > tr, :scope > tbody > tr, :scope > thead > tr')]
    .map(tr => [...tr.children].map(td => td.innerText !== undefined && td.isConnected ? td.innerText.trim() : td.textContent.replace(/\s+/g, ' ').trim()));
}

// 붙여넣기·불러오기 때 허용하는 태그와 속성만 남긴다 (스크립트·이벤트 제거)
const KEEP = new Set(['p', 'div', 'br', 'span', 'b', 'strong', 'i', 'em', 'u', 's', 'strike', 'del', 'ins', 'mark', 'font', 'a', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'blockquote', 'pre', 'code', 'hr', 'table', 'thead', 'tbody', 'tr', 'td', 'th', 'img', 'sub', 'sup']);
const DROP = new Set(['script', 'style', 'iframe', 'object', 'embed', 'link', 'meta', 'title', 'head', 'noscript', 'svg', 'math', 'form', 'input', 'button', 'select', 'textarea', 'video', 'audio']);
const CSS_OK = /^(color|background-color|font-weight|font-style|text-decoration|text-decoration-line|font-size|text-align|margin-left|width)$/;
export function sanitize(html, { keepFiles = true } = {}) {
  const body = parseHTML(html);
  const clean = node => {
    for (const n of [...node.childNodes]) {
      if (n.nodeType === 8) { n.remove(); continue; }
      if (n.nodeType !== 1) continue;
      const tag = n.tagName.toLowerCase();
      if (DROP.has(tag)) { n.remove(); continue; }
      clean(n);
      if (!KEEP.has(tag)) { n.replaceWith(...n.childNodes); continue; }
      for (const a of [...n.attributes]) {
        const name = a.name.toLowerCase();
        const ok = (name === 'style') || (name === 'href' && tag === 'a' && /^(https?:|mailto:)/i.test(a.value)) ||
          (name === 'src' && tag === 'img' && /^(https?:|data:image\/|blob:)/i.test(a.value)) ||
          (name === 'start' && tag === 'ol') || (name === 'color' && tag === 'font') || (name === 'align') ||
          (name === 'class' && /^(checklist|att)$/.test(a.value)) || (name === 'data-checked' && tag === 'li') ||
          (keepFiles && /^data-(fid|name|size)$/.test(name)) || (name === 'contenteditable' && n.classList.contains('att'));
        if (!ok) n.removeAttribute(a.name);
      }
      if (n.hasAttribute('style')) {
        const css = n.getAttribute('style').split(';').map(d => d.trim()).filter(d => { const k = d.split(':')[0]?.trim().toLowerCase(); return CSS_OK.test(k) && (k !== 'width' || tag === 'img') && !/url\(|expression/i.test(d); }).join('; ');
        if (css) n.setAttribute('style', css); else n.removeAttribute('style');
      }
      if (tag === 'a') { n.setAttribute('target', '_blank'); n.setAttribute('rel', 'noopener'); }
    }
  };
  clean(body);
  return body.innerHTML;
}
