// 페이지 편집기: 서식 명령, 자동 서식(끄고 켤 수 있음), 할 일 목록, 표, 그림·파일, 붙여넣기
import { sanitize } from './doc.js';

let ed, opt, lastAuto = null, imgBar, selImg = null;
const sel = () => window.getSelection();
const inEditor = n => n && ed.contains(n.nodeType === 1 ? n : n.parentNode);
const changed = () => opt.onChange();

export function setupEditor(editor, options) {
  ed = editor; opt = options;
  document.execCommand('defaultParagraphSeparator', false, 'p');
  ed.addEventListener('input', onInput);
  ed.addEventListener('compositionend', () => setTimeout(() => tryAuto(), 0)); // 한글 입력(ㅁ 등) 뒤 띄어쓰기
  ed.addEventListener('keydown', onKey);
  ed.addEventListener('click', onClick);
  ed.addEventListener('paste', onPaste);
  ed.addEventListener('dragover', e => { if ([...e.dataTransfer.types].includes('Files')) e.preventDefault(); });
  ed.addEventListener('drop', onDrop);
  imgBar = document.getElementById('imgBar');
  imgBar.addEventListener('mousedown', e => e.preventDefault());
  imgBar.addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b || !selImg) return;
    if (b.dataset.w === 'del') { selImg.remove(); hideImgBar(); }
    else { selImg.style.width = b.dataset.w + '%'; placeImgBar(); }
    changed();
  });
  document.addEventListener('selectionchange', () => opt.onSelection?.());
  setupLongPressHighlight();
  window.addEventListener('resize', () => selImg && placeImgBar());
  document.getElementById('canvas').addEventListener('scroll', () => selImg && placeImgBar(), { passive: true });
}

// ---------- 서식 명령 ----------
export function exec(cmd, val) {
  ed.focus();
  const css = cmd === 'foreColor' || cmd === 'hiliteColor';
  document.execCommand('styleWithCSS', false, css);
  if (cmd === 'hiliteColor' && !val) val = 'transparent';
  document.execCommand(cmd, false, val ?? null);
  document.execCommand('styleWithCSS', false, false);
  changed();
}
export function formatBlock(tag) { ed.focus(); document.execCommand('formatBlock', false, tag); changed(); }
export function fontSize(pt) {
  ed.focus();
  document.execCommand('styleWithCSS', false, false);
  document.execCommand('fontSize', false, '7');
  const made = [];
  ed.querySelectorAll('font[size="7"]').forEach(f => {
    const s = document.createElement('span'); s.style.fontSize = pt + 'pt';
    s.append(...f.childNodes); f.replaceWith(s); made.push(s);
    s.querySelectorAll('span[style*="font-size"]').forEach(x => { x.style.fontSize = ''; if (!x.getAttribute('style')) x.replaceWith(...x.childNodes); }); // 안쪽 예전 크기는 지움
  });
  // 바꾼 글을 계속 골라 둠 (크기를 한 단계씩 더 바꿀 수 있게)
  if (made.length) { const r = document.createRange(); r.setStartBefore(made[0]); r.setEndAfter(made[made.length - 1]); const s = sel(); s.removeAllRanges(); s.addRange(r); }
  changed();
}
export function insertHTML(html) { ed.focus(); document.execCommand('insertHTML', false, html); changed(); }
export function insertText(t) { ed.focus(); document.execCommand('insertText', false, t); changed(); }

// 할 일 목록: 지금 줄을 체크박스 목록으로 (이미 그렇다면 해제)
export function toggleChecklist() {
  ed.focus();
  const ul = closest('ul');
  if (ul && ul.classList.contains('checklist')) { document.execCommand('insertUnorderedList'); changed(); return; }
  if (!ul) document.execCommand('insertUnorderedList');
  const u = closest('ul'); if (u) u.classList.add('checklist');
  changed();
}
function closest(sel_) {
  const s = sel(); if (!s.rangeCount) return null;
  let n = s.anchorNode; if (!inEditor(n)) return null;
  n = n.nodeType === 1 ? n : n.parentNode;
  const el = n.closest(sel_); return el && ed.contains(el) ? el : null;
}
export const caretIn = s => closest(s);

// ---------- 자동 서식 ----------
// "1. " "1) " → 번호 목록, "- " → 대시 목록, "* " "• " "ㅁ " → 동그라미 목록, ": " → 쌍점 목록, "[] " "[ ] " → 할 일
function onInput(e) {
  if (selImg && !ed.contains(selImg)) hideImgBar();
  changed();
  if (e.inputType === 'insertParagraph') { // 할 일에서 새 줄은 체크 해제 상태로
    const li = closest('ul.checklist > li'); if (li) li.removeAttribute('data-checked');
    splitContinuation();
    return;
  }
  if (e.isComposing) return; // 한글 조합 중에는 건드리지 않음 (조합이 끝나면 compositionend에서 다시 봄)
  if (!/^insert(Text|CompositionText)$/.test(e.inputType) || !/[ \u00a0]$/.test(e.data || '')) return;
  tryAuto();
}
function tryAuto() {
  const set = opt.settings();
  if (!set.auto) return;
  const s = sel(); if (!s.rangeCount || !s.isCollapsed) return;
  if (closest('td, th, pre')) return;
  const li0 = closest('li');
  const block = li0 || closest('p, div, h1, h2, h3, blockquote') || ed;
  const r = document.createRange(); r.setStart(block, 0); r.setEnd(s.anchorNode, s.anchorOffset);
  const typed = r.toString().replace(/\u00a0/g, ' ');
  // "1. " → 번호, "- " → 대시(-), "* " "• " "ㅁ " → 동그라미, ": " → 쌍점(:), "[] " → 할 일
  let kind = null, start = 1, m;
  if (set.autoNumber && (m = typed.match(/^(\d{1,3})[.)] $/))) { kind = 'ol'; start = +m[1]; }
  else if (set.autoBullet && typed === '- ') kind = 'dash';
  else if (set.autoBullet && typed === ': ') kind = 'colon';
  else if (set.autoBullet && /^[*•ㅁ] $/.test(typed)) kind = 'ul';
  else if (set.autoCheck && /^\[ ?\] $/.test(typed)) kind = 'check';
  if (!kind) return;
  if (li0) { // 목록 안의 새 줄: 그 줄만 다른 목록으로 (예: "2." 줄에서 "- " → 대시 줄)
    const pl = li0.parentNode;
    if (listKind(pl) === kind) return;
    s.removeAllRanges(); s.addRange(r); document.execCommand('delete');
    switchListItem(li0, kind, start);
    changed(); return;
  }
  s.removeAllRanges(); s.addRange(r);
  document.execCommand('delete');
  document.execCommand(kind === 'ol' ? 'insertOrderedList' : 'insertUnorderedList');
  const list = closest(kind === 'ol' ? 'ol' : 'ul');
  // 크롬이 목록을 <p> 안에 넣는 경우가 있어 꺼냄
  if (list && list.parentNode.tagName === 'P') {
    const p = list.parentNode, li = closest('li');
    p.replaceWith(...[...p.childNodes].filter(n => n === list || n.textContent.trim()));
    if (li) { const r2 = document.createRange(); r2.selectNodeContents(li); r2.collapse(true); s.removeAllRanges(); s.addRange(r2); }
  }
  if (list && kind === 'ol' && start !== 1) list.setAttribute('start', start);
  // 크롬이 바로 위의 다른 종류 목록에 합쳐 버린 경우: 이 줄만 떼어 냄
  const liNew = closest('li');
  if (list && liNew && list.children.length > 1 && listKind(list) !== kind && !(kind === 'ul' && listKind(list) === 'ul')) {
    switchListItem(liNew, kind, start);
  } else {
    if (list && kind === 'check') list.classList.add('checklist');
    if (list && kind === 'dash') list.classList.add('dash');
    if (list && kind === 'colon') list.classList.add('colon');
  }
  lastAuto = { li: closest('li'), text: typed, kind };
  opt.onAuto?.(kind, undoAuto);
  changed();
}
const listKind = l => l.tagName === 'OL' ? 'ol' : l.classList.contains('checklist') ? 'check' : l.classList.contains('dash') ? 'dash' : l.classList.contains('colon') ? 'colon' : 'ul';
// 목록의 한 줄만 다른 종류의 목록으로 떼어 냄 (앞뒤 줄은 원래 목록 그대로)
function switchListItem(li, kind, start) {
  const list = li.parentNode, items = [...list.children], i = items.indexOf(li), after = items.slice(i + 1);
  const nl = document.createElement(kind === 'ol' ? 'ol' : 'ul');
  if (kind === 'dash' || kind === 'colon') nl.className = kind; if (kind === 'check') nl.className = 'checklist';
  if (kind === 'ol' && start !== 1) nl.setAttribute('start', start);
  list.after(nl); nl.append(li);
  if (after.length) {
    const rest = list.cloneNode(false);
    if (list.tagName === 'OL') rest.setAttribute('start', (parseInt(list.getAttribute('start'), 10) || 1) + i);
    nl.after(rest); rest.append(...after);
  }
  if (!list.children.length) list.remove();
  if (!li.textContent && !li.querySelector('br')) li.innerHTML = '<br>';
  const r = document.createRange(); r.selectNodeContents(li); r.collapse(true);
  const s = sel(); s.removeAllRanges(); s.addRange(r);
}
// 자동으로 바뀐 것 되돌리기 (바로 다음 백스페이스 또는 알림의 [되돌리기])
export function undoAuto() {
  const a = lastAuto; lastAuto = null;
  if (!a || !a.li || !ed.contains(a.li)) return;
  const r = document.createRange(); r.selectNodeContents(a.li); r.collapse(true);
  const s = sel(); s.removeAllRanges(); s.addRange(r);
  ed.focus();
  a.li.parentNode.classList.remove('dash', 'colon');
  document.execCommand(a.kind === 'ol' ? 'insertOrderedList' : 'insertUnorderedList');
  document.execCommand('insertText', false, a.text);
  changed();
}

function onKey(e) {
  if (lastAuto && e.key === 'Backspace') {
    const li = closest('li'), s = sel();
    if (li && li === lastAuto.li && s.isCollapsed && atStart(li)) { e.preventDefault(); undoAuto(); return; }
  }
  if (e.key !== 'Shift' && e.key !== 'Control') lastAuto = e.key === 'Backspace' ? lastAuto : null;
  if (e.key === 'Backspace' && !e.shiftKey && !e.ctrlKey && !e.metaKey && listBackspace()) { e.preventDefault(); changed(); return; }
  if (e.key === 'Tab') {
    const td = closest('td, th');
    if (td) { e.preventDefault(); moveCell(td, e.shiftKey ? -1 : 1); return; }
    if (closest('li')) { e.preventDefault(); document.execCommand(e.shiftKey ? 'outdent' : 'indent'); changed(); return; }
    if (!e.shiftKey) { e.preventDefault(); document.execCommand('insertText', false, '    '); }
  }
  if ((e.key === 'Delete' || e.key === 'Backspace') && selImg) { e.preventDefault(); selImg.remove(); hideImgBar(); changed(); }
}
// 목록 줄 맨 앞에서 지우기 (원노트처럼)
//  1번째: 기호(- 1.)만 없어지고 글은 윗줄 글과 같은 자리에 맞춰짐 (윗줄 항목에 이어 쓰는 줄)
//  2번째: 그 줄이 목록에서 빠져 맨 왼쪽으로. 커서는 그 줄에 그대로
function listBackspace() {
  const s = sel(); if (!s.rangeCount || !s.isCollapsed) return false;
  const li = closest('li'); if (!li) return false;
  let n = s.anchorNode;
  if (n === li) n = li.childNodes[s.anchorOffset] || li.lastChild;
  while (n && n.parentNode !== li) n = n.parentNode;
  if (!n) return false;
  const isBlock = x => x.nodeType === 1 && /^(DIV|P)$/.test(x.tagName);
  const hasBefore = x => { for (let p = x.previousSibling; p; p = p.previousSibling) if (p.textContent.trim() || (p.nodeType === 1 && p.tagName !== 'BR')) return true; return false; };
  if (isBlock(n) && hasBefore(n)) { // 이어 쓰는 줄의 맨 앞 → 목록에서 빼기
    if (!atStart(n)) return false;
    liftOut(li, n); return true;
  }
  if (!atStart(li)) return false;
  const prev = li.previousElementSibling;
  if (prev && prev.tagName === 'LI') { // 윗 항목에 이어 쓰는 줄로
    const line = document.createElement('div');
    while (li.firstChild && !isBlock(li.firstChild) && !/^(UL|OL)$/.test(li.firstChild.nodeName)) line.append(li.firstChild);
    if (!line.textContent && !line.querySelector('img, br')) line.innerHTML = '<br>';
    prev.append(line, ...li.childNodes); li.remove();
    caretAt(line); return true;
  }
  liftOut(li, null); return true; // 목록의 첫 줄 → 바로 맨 왼쪽으로
}
// 이어 쓰는 줄에서 엔터: 새 줄은 다음 항목(기호 있는 줄)이 됨
function splitContinuation() {
  const li = closest('li'), s = sel(); if (!li || !s.rangeCount) return;
  let n = s.anchorNode; if (n === li) return;
  while (n && n.parentNode !== li) n = n.parentNode;
  // 크롬이 새 항목을 <li><div>…</div></li>로 만든 경우: div를 풀어 보통 항목으로
  if (n && n === li.firstChild && li.childNodes.length === 1 && n.nodeType === 1 && n.tagName === 'DIV') {
    n.replaceWith(...(n.childNodes.length ? n.childNodes : [document.createElement('br')])); caretAt(li); return;
  }
  if (!n || n.nodeType !== 1 || !/^(DIV|P)$/.test(n.tagName) || !n.previousSibling) return;
  const prevBlock = n.previousElementSibling;
  if (!prevBlock || !/^(DIV|P)$/.test(prevBlock.tagName)) return; // 이어 쓰는 줄에서 엔터를 쳐서 생긴 새 줄일 때만
  const nli = document.createElement('li');
  for (let x = n; x; ) { const nx = x.nextSibling; nli.append(x); x = nx; }
  const first = nli.firstChild; // 첫 줄은 감싼 div를 풀어 li 글로
  if (first && first.nodeType === 1 && /^(DIV|P)$/.test(first.tagName)) first.replaceWith(...(first.childNodes.length ? first.childNodes : [document.createElement('br')]));
  li.after(nli); caretAt(nli);
}
// li 안의 줄(block, 없으면 li 전체)을 목록 밖 맨 왼쪽 줄로 꺼냄. 아래 항목들은 새 목록으로 이어짐
function liftOut(li, block) {
  const list = li.parentNode, items = [...list.children], i = items.indexOf(li);
  const out = [];
  if (block) {
    for (let x = block; x; ) { const nx = x.nextSibling; out.push(x); x = nx; }
  } else {
    const p = document.createElement('p');
    while (li.firstChild && !/^(DIV|P|UL|OL)$/.test(li.firstChild.nodeName)) p.append(li.firstChild);
    out.push(p, ...li.childNodes);
  }
  const lines = out.map(x => {
    if (x.nodeType === 1 && /^(P|DIV|UL|OL)$/.test(x.tagName)) return x;
    const p = document.createElement('p'); p.append(x); return p;
  }).map(x => { if (x.tagName === 'DIV') { const p = document.createElement('p'); p.append(...x.childNodes); x.remove(); return p; } return x; });
  lines.forEach(x => { if (/^P$/.test(x.tagName) && !x.textContent && !x.querySelector('img, br')) x.innerHTML = '<br>'; });
  const keepLi = !!block;
  if (!keepLi) li.remove();
  const after = items.slice(i + 1);
  let tail = list;
  if (lines.length) { tail.after(...lines); tail = lines[lines.length - 1]; }
  if (after.length) {
    const rest = list.cloneNode(false);
    if (list.tagName === 'OL') rest.setAttribute('start', (parseInt(list.getAttribute('start'), 10) || 1) + i + (keepLi ? 1 : 0));
    tail.after(rest); rest.append(...after);
  }
  if (!list.children.length) list.remove();
  caretAt(lines[0]);
}
function caretAt(el) {
  if (!el) return;
  const r = document.createRange(); r.selectNodeContents(el); r.collapse(true);
  const s = sel(); s.removeAllRanges(); s.addRange(r);
}
function atStart(el) {
  const s = sel(); const r = document.createRange(); r.selectNodeContents(el); r.setEnd(s.anchorNode, s.anchorOffset);
  return r.toString().length === 0;
}

// ---------- 클릭: 체크박스, 링크, 그림, 첨부 ----------
function onClick(e) {
  const li = e.target.closest('ul.checklist > li');
  if (li && e.target === li && e.clientX - li.getBoundingClientRect().left < 26) {
    e.preventDefault();
    li.toggleAttribute('data-checked'); changed(); return;
  }
  const a = e.target.closest('a[href]');
  if (a && (e.ctrlKey || e.metaKey)) { e.preventDefault(); window.open(a.href, '_blank', 'noopener'); return; }
  const att = e.target.closest('.att');
  if (att) { e.preventDefault(); opt.openFile(att.dataset.fid, att.dataset.name); return; }
  if (e.target.tagName === 'IMG') { selImg = e.target; selImg.classList.add('sel'); placeImgBar(); return; }
  hideImgBar();
}
function placeImgBar() {
  ed.querySelectorAll('img.sel').forEach(i => i !== selImg && i.classList.remove('sel'));
  const r = selImg.getBoundingClientRect();
  imgBar.hidden = false;
  imgBar.style.top = Math.max(8, r.top - imgBar.offsetHeight - 6) + 'px';
  imgBar.style.left = Math.max(8, Math.min(window.innerWidth - imgBar.offsetWidth - 8, r.left)) + 'px';
}
export function hideImgBar() { if (selImg) selImg.classList.remove('sel'); selImg = null; if (imgBar) imgBar.hidden = true; }

// ---------- 꾹 눌러 끌면 형광펜 (휴대폰·태블릿) ----------
// 시작점을 꾹 누르면 진동 → 그대로 끌면 노란 형광펜이 칠해짐
function setupLongPressHighlight() {
  const pv = document.getElementById('hlPreview');
  let st = null;
  const at = (x, y) => document.caretRangeFromPoint ? document.caretRangeFromPoint(x, y) : null;
  const rangeTo = (x, y) => {
    const end = at(x, y); if (!end || !ed.contains(end.startContainer)) return null;
    const r = document.createRange();
    if (st.start.compareBoundaryPoints(Range.START_TO_START, end) <= 0) { r.setStart(st.start.startContainer, st.start.startOffset); r.setEnd(end.startContainer, end.startOffset); }
    else { r.setStart(end.startContainer, end.startOffset); r.setEnd(st.start.startContainer, st.start.startOffset); }
    return r;
  };
  const draw = r => { pv.innerHTML = r ? [...r.getClientRects()].map(b => `<i style="left:${b.left}px;top:${b.top}px;width:${b.width}px;height:${b.height}px"></i>`).join('') : ''; };
  const stop = () => { if (st) clearTimeout(st.timer); st = null; ed.classList.remove('hl-mode'); pv.innerHTML = ''; };
  ed.addEventListener('touchstart', e => {
    if (e.touches.length !== 1) return stop();
    const t = e.touches[0];
    st = { x: t.clientX, y: t.clientY, active: false, range: null };
    st.timer = setTimeout(() => {
      if (!st) return;
      st.start = at(st.x, st.y); if (!st.start || !ed.contains(st.start.startContainer)) return stop();
      st.active = true; ed.classList.add('hl-mode'); navigator.vibrate?.(30);
      getSelection().removeAllRanges();
    }, 450);
  }, { passive: true });
  ed.addEventListener('touchmove', e => {
    if (!st) return;
    const t = e.touches[0];
    if (!st.active) { if (Math.hypot(t.clientX - st.x, t.clientY - st.y) > 10) stop(); return; } // 그냥 스크롤
    e.preventDefault();
    st.range = rangeTo(t.clientX, t.clientY); draw(st.range);
  }, { passive: false });
  const end = () => {
    if (!st) return;
    const r = st.active && st.range && !st.range.collapsed ? st.range : null;
    stop();
    if (!r) return;
    ed.focus(); const s = sel(); s.removeAllRanges(); s.addRange(r);
    exec('hiliteColor', '#fff176');
    s.collapseToEnd();
  };
  ed.addEventListener('touchend', end); ed.addEventListener('touchcancel', stop);
  ed.addEventListener('contextmenu', e => { if (st?.active) e.preventDefault(); });
}

// ---------- 표 ----------
export function tableHTML(rows, cols) {
  const tr = '<tr>' + '<td><br></td>'.repeat(cols) + '</tr>';
  return `<table><tbody>${tr.repeat(rows)}</tbody></table><p><br></p>`;
}
function moveCell(td, dir) {
  const cells = [...td.closest('table').querySelectorAll('td, th')];
  let i = cells.indexOf(td) + dir;
  if (i >= cells.length) { tableOp('rowBelow'); i = cells.indexOf(td) + 1; return moveCell(td, 1); }
  if (i < 0) return;
  const r = document.createRange(); r.selectNodeContents(cells[i]); r.collapse(false);
  const s = sel(); s.removeAllRanges(); s.addRange(r);
}
export function tableOp(op) {
  const td = closest('td, th'); if (!td) return false;
  const tr = td.parentNode, table = td.closest('table'), ci = [...tr.children].indexOf(td);
  const empty = tag => { const c = document.createElement(tag); c.innerHTML = '<br>'; return c; };
  if (op === 'rowAbove' || op === 'rowBelow') {
    const n = document.createElement('tr'); [...tr.children].forEach(c => n.append(empty(c.tagName.toLowerCase())));
    op === 'rowAbove' ? tr.before(n) : tr.after(n);
  } else if (op === 'colLeft' || op === 'colRight') {
    table.querySelectorAll('tr').forEach(r => { const ref = r.children[ci]; const c = empty(ref ? ref.tagName.toLowerCase() : 'td'); if (!ref) r.append(c); else op === 'colLeft' ? ref.before(c) : ref.after(c); });
  } else if (op === 'delRow') {
    if (table.querySelectorAll('tr').length <= 1) table.remove(); else tr.remove();
  } else if (op === 'delCol') {
    if (tr.children.length <= 1) table.remove(); else table.querySelectorAll('tr').forEach(r => r.children[ci]?.remove());
  } else if (op === 'delTable') table.remove();
  else if (op === 'header') { // 첫 줄을 머리글(굵게)로
    const first = table.querySelector('tr'); [...first.children].forEach(c => { const n = document.createElement(c.tagName === 'TH' ? 'td' : 'th'); n.append(...c.childNodes); c.replaceWith(n); });
  }
  changed(); return true;
}

// ---------- 붙여넣기·끌어 놓기 ----------
async function onPaste(e) {
  const files = [...(e.clipboardData?.files || [])];
  if (files.length) { e.preventDefault(); for (const f of files) await addFile(f); return; }
  const html = e.clipboardData?.getData('text/html');
  if (html) {
    e.preventDefault();
    const clean = sanitize(html.replace(/<!--StartFragment-->|<!--EndFragment-->/g, ''), { keepFiles: false });
    if (clean.length > 60000) insertFast(clean); // 아주 큰 글은 한 번에 넣어서 멈추지 않게
    else document.execCommand('insertHTML', false, clean);
    changed();
    return;
  }
  const text = e.clipboardData?.getData('text/plain') || '';
  if (text.length > 60000) {
    e.preventDefault();
    const esc = t => t.replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
    insertFast(text.split(/\r?\n/).map(l => `<p>${esc(l) || '<br>'}</p>`).join(''));
    changed();
  }
}
function insertFast(html) {
  const s = sel(); if (!s.rangeCount) return;
  const r = s.getRangeAt(0); r.deleteContents();
  const frag = r.createContextualFragment(html), last = frag.lastChild;
  r.insertNode(frag);
  if (last) { const r2 = document.createRange(); r2.setStartAfter(last); r2.collapse(true); s.removeAllRanges(); s.addRange(r2); }
}
async function onDrop(e) {
  const files = [...(e.dataTransfer?.files || [])];
  if (!files.length) return;
  e.preventDefault();
  const pos = document.caretRangeFromPoint ? document.caretRangeFromPoint(e.clientX, e.clientY) : null;
  if (pos) { const s = sel(); s.removeAllRanges(); s.addRange(pos); }
  for (const f of files) await addFile(f);
}
// 그림은 페이지에 바로 보이게, 그 밖의 파일은 첨부로
export async function addFile(file) {
  const s = sel(), range = s.rangeCount && inEditor(s.anchorNode) ? s.getRangeAt(0).cloneRange() : null;
  const isImg = /^image\/(png|jpeg|gif|webp|bmp)$/.test(file.type);
  const res = isImg ? await opt.uploadImage(file) : await opt.uploadFile(file);
  if (!res) return;
  ed.focus();
  if (range) { s.removeAllRanges(); s.addRange(range); }
  else { const r = document.createRange(); r.selectNodeContents(ed); r.collapse(false); s.removeAllRanges(); s.addRange(r); }
  const esc = t => t.replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  if (isImg) document.execCommand('insertHTML', false, `<img data-fid="${res.fid}" src="${res.url}" alt="">`);
  else document.execCommand('insertHTML', false, `<span class="att" contenteditable="false" data-fid="${res.fid}" data-name="${esc(file.name)}" data-size="${file.size}">📎 ${esc(file.name)} <small>${sizeText(file.size)}</small></span>&nbsp;`);
  changed();
}
export const sizeText = n => n < 1024 ? n + 'B' : n < 1048576 ? (n / 1024).toFixed(0) + 'KB' : (n / 1048576).toFixed(1) + 'MB';
