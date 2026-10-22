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
  ed.querySelectorAll('font[size="7"]').forEach(f => {
    const s = document.createElement('span'); s.style.fontSize = pt + 'pt';
    s.append(...f.childNodes); f.replaceWith(s);
  });
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
// "1. " "1) " → 번호 목록, "- " "* " "• " → 글머리 기호, "[] " "[ ] " → 할 일
function onInput(e) {
  if (selImg && !ed.contains(selImg)) hideImgBar();
  changed();
  if (e.inputType === 'insertParagraph') { // 할 일에서 새 줄은 체크 해제 상태로
    const li = closest('ul.checklist > li'); if (li) li.removeAttribute('data-checked');
    return;
  }
  if (e.inputType !== 'insertText' || (e.data !== ' ' && e.data !== ' ')) return;
  const set = opt.settings();
  if (!set.auto) return;
  const s = sel(); if (!s.rangeCount || !s.isCollapsed) return;
  if (closest('li, td, th, pre')) return;
  const block = closest('p, div, h1, h2, h3, blockquote') || ed;
  const r = document.createRange(); r.setStart(block, 0); r.setEnd(s.anchorNode, s.anchorOffset);
  const typed = r.toString().replace(/ /g, ' ');
  let kind = null, start = 1, m;
  if (set.autoNumber && (m = typed.match(/^(\d{1,3})[.)] $/))) { kind = 'ol'; start = +m[1]; }
  else if (set.autoBullet && /^[-*•] $/.test(typed)) kind = 'ul';
  else if (set.autoCheck && /^\[ ?\] $/.test(typed)) kind = 'check';
  if (!kind) return;
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
  if (list && kind === 'check') list.classList.add('checklist');
  lastAuto = { li: closest('li'), text: typed, kind };
  opt.onAuto?.(kind, undoAuto);
  changed();
}
// 자동으로 바뀐 것 되돌리기 (바로 다음 백스페이스 또는 알림의 [되돌리기])
export function undoAuto() {
  const a = lastAuto; lastAuto = null;
  if (!a || !a.li || !ed.contains(a.li)) return;
  const r = document.createRange(); r.selectNodeContents(a.li); r.collapse(true);
  const s = sel(); s.removeAllRanges(); s.addRange(r);
  ed.focus();
  document.execCommand(a.kind === 'ol' ? 'insertOrderedList' : 'insertUnorderedList');
  document.execCommand('insertText', false, a.text);
  changed();
}

function onKey(e) {
  if (lastAuto && e.key === 'Backspace') {
    const li = closest('li'), s = sel();
    if (li === lastAuto.li && s.isCollapsed && atStart(li)) { e.preventDefault(); undoAuto(); return; }
  }
  if (e.key !== 'Shift' && e.key !== 'Control') lastAuto = e.key === 'Backspace' ? lastAuto : null;
  if (e.key === 'Tab') {
    const td = closest('td, th');
    if (td) { e.preventDefault(); moveCell(td, e.shiftKey ? -1 : 1); return; }
    if (closest('li')) { e.preventDefault(); document.execCommand(e.shiftKey ? 'outdent' : 'indent'); changed(); return; }
    if (!e.shiftKey) { e.preventDefault(); document.execCommand('insertText', false, '    '); }
  }
  if ((e.key === 'Delete' || e.key === 'Backspace') && selImg) { e.preventDefault(); selImg.remove(); hideImgBar(); changed(); }
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
    document.execCommand('insertHTML', false, clean);
    changed();
  }
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
