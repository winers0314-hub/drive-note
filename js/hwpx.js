// 노트 페이지(HTML) → 한글 문서(.hwpx)
// 글자 모양(굵게·기울임·밑줄·취소선·색·형광펜·크기), 문단(정렬·들여쓰기), 목록, 할 일, 표, 그림을 옮긴다.
import * as T from './hwpx-template.js';
import { parseHTML, blockAlign, styleOf } from './doc.js';

const PX = 75;            // 1px = 75 HWPUNIT (7200 per inch / 96 dpi)
const PAGE_W = 42520;     // 본문 폭 (A4, 좌우 여백 30mm)
const esc = s => s.replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

class Doc {
  constructor() {
    this.chars = new Map(); this.charXml = [];
    this.paras = new Map(); this.paraXml = [];
    this.bins = [];       // {id, ext, mime, data}
    this.uid = 1500000000;
  }
  // 글자 모양 하나 = charPr 하나 (같은 모양은 재사용)
  char(st) {
    const height = Math.round((st.size || 10) * 100);
    const color = st.color || '#000000', shade = st.shade || 'none';
    const key = [height, color, shade, !!st.b, !!st.i, !!st.u, !!st.s].join('|');
    if (key === '1000|#000000|none|false|false|false|false') return 0; // 기본 글자 모양
    if (this.chars.has(key)) return this.chars.get(key);
    const id = T.CHAR_BASE + this.charXml.length;
    this.charXml.push(`<hh:charPr id="${id}" height="${height}" textColor="${color}" shadeColor="${shade}" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="2">${T.CHARPR_INNER}` +
      (st.i ? '<hh:italic/>' : '') + (st.b ? '<hh:bold/>' : '') +
      (st.u ? '<hh:underline type="BOTTOM" shape="SOLID" color="#000000"/>' : '') +
      (st.s ? '<hh:strikeout shape="SOLID" color="#000000"/>' : '') + '</hh:charPr>');
    this.chars.set(key, id);
    return id;
  }
  // 문단 모양 = 정렬 + 왼쪽 여백
  para(align = 'JUSTIFY', left = 0) {
    left = Math.max(0, Math.round(left));
    const key = align + '|' + left;
    if (key === 'JUSTIFY|0') return 0;
    if (this.paras.has(key)) return this.paras.get(key);
    const id = T.PARA_BASE + this.paraXml.length;
    this.paraXml.push(T.PARAPR0.replace('id="0"', `id="${id}"`)
      .replace(/<hh:align horizontal="\w+"/, `<hh:align horizontal="${align}"`)
      .replace(/<hc:left value="\d+"/, `<hc:left value="${left}"`));
    this.paras.set(key, id);
    return id;
  }
  header() {
    return T.HEADER.replace('{{CHAR_CNT}}', T.CHAR_BASE + this.charXml.length).replace('{{CHARPRS}}', this.charXml.join(''))
      .replace('{{PARA_CNT}}', T.PARA_BASE + this.paraXml.length).replace('{{PARAPRS}}', this.paraXml.join(''));
  }
}

// 블록 단위로 문단을 모은다: {align, left, prefix, runs:[{text, st}], pageBreak} | {table} | {image}
class Writer {
  constructor(doc, images) { this.doc = doc; this.images = images; this.out = []; this.cur = null; }
  start(ctx) { this.flush(); this.cur = { align: ctx.align, left: ctx.left, runs: [], prefix: ctx.prefix || '', pageBreak: ctx.pageBreak }; ctx.prefix = ''; ctx.pageBreak = false; }
  flush() { if (this.cur) { this.out.push(this.paraXml(this.cur)); this.cur = null; } }
  text(t, st, ctx) {
    if (!this.cur) this.start(ctx);
    const last = this.cur.runs[this.cur.runs.length - 1];
    if (last && sameSt(last.st, st)) last.text += t; else this.cur.runs.push({ text: t, st });
  }
  paraXml(p) {
    const d = this.doc;
    const runs = p.runs.filter(r => r.text.length);
    if (p.prefix) runs.unshift({ text: p.prefix, st: runs[0] ? runs[0].st : {} });
    const body = runs.length ? runs.map(r => `<hp:run charPrIDRef="${d.char(r.st)}"><hp:t>${esc(r.text)}</hp:t></hp:run>`).join('') : '<hp:run charPrIDRef="0"><hp:t/></hp:run>';
    return `<hp:p id="0" paraPrIDRef="${d.para(p.align, p.left)}" styleIDRef="0" pageBreak="${p.pageBreak ? 1 : 0}" columnBreak="0" merged="0">${body}</hp:p>`;
  }
  raw(xml, ctx) { this.flush(); this.out.push(`<hp:p id="0" paraPrIDRef="${this.doc.para(ctx.align, ctx.left)}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0">${xml}<hp:t/></hp:run></hp:p>`); }
}
const sameSt = (a, b) => ['b', 'i', 'u', 's', 'color', 'shade', 'size'].every(k => (a[k] || '') === (b[k] || ''));

// DOM 순회
function walk(node, st, ctx, w) {
  for (const n of node.childNodes) {
    if (n.nodeType === 3) {
      const t = n.nodeValue.replace(/[\r\n]+/g, '').replace(/ /g, ' ').replace(/\t/g, '    ');
      if (!t) continue;
      if (!w.cur && !t.trim()) continue; // 블록 사이 공백
      w.text(t, st, ctx);
      continue;
    }
    if (n.nodeType !== 1) continue;
    const tag = n.tagName.toLowerCase();
    if (tag === 'br') { if (w.cur) w.flush(); else w.start(ctx), w.flush(); continue; }
    if (tag === 'table') { w.flush(); w.raw(tableXml(n, st, w), ctx); continue; }
    if (tag === 'img') { const x = picXml(n, w); if (x) w.raw(x, ctx); else w.text('[그림]', st, ctx); continue; }
    if (tag === 'hr') { w.flush(); w.start(ctx); w.text('─'.repeat(30), { ...st, color: '#999999' }, ctx); w.flush(); continue; }
    if (n.classList.contains('att')) { w.text('[첨부] ' + (n.dataset.name || n.textContent.trim()), { ...st, color: '#1f5fbf' }, ctx); continue; }
    if (tag === 'ul' || tag === 'ol') { listXml(n, st, ctx, w); continue; }
    const s2 = styleOf(n, st);
    if (/^(p|div|h[1-6]|blockquote|pre|li)$/.test(tag)) {
      w.flush();
      const c2 = { ...ctx, align: blockAlign(n) || ctx.align, left: ctx.left + (tag === 'blockquote' ? 40 * PX : 0) + indentOf(n) };
      w.start(c2);
      walk(n, s2, c2, w);
      w.flush();
      continue;
    }
    walk(n, s2, ctx, w);
  }
}
const indentOf = el => { const m = (el.getAttribute('style') || '').match(/margin-left:\s*(\d+)px/); return m ? +m[1] * PX : 0; };

function listXml(list, st, ctx, w) {
  w.flush();
  const ordered = list.tagName === 'OL', check = list.classList.contains('checklist');
  let n = ordered ? (parseInt(list.getAttribute('start'), 10) || 1) : 0;
  const depth = (ctx.depth || 0) + 1;
  for (const li of list.children) {
    if (li.tagName !== 'LI') continue;
    const prefix = check ? (li.hasAttribute('data-checked') ? '☑ ' : '☐ ') : ordered ? (n++) + '. ' : list.classList.contains('dash') ? '- ' : (depth % 2 ? '• ' : '◦ ');
    const c2 = { ...ctx, depth, left: ctx.left + depth * 1500, align: blockAlign(li) || ctx.align, prefix };
    const s2 = check && li.hasAttribute('data-checked') ? { ...styleOf(li, st), s: true, color: '#888888' } : styleOf(li, st);
    w.start(c2);
    // 목록 안의 목록은 한 단계 더 들여서
    walk(li, s2, c2, w);
    w.flush();
  }
}

function tableXml(table, st, w) {
  const rows = [...table.querySelectorAll(':scope > tr, :scope > tbody > tr, :scope > thead > tr')];
  const cols = Math.max(1, ...rows.map(r => r.children.length));
  const cw = Math.floor(PAGE_W / cols), rh = 282;
  const trs = rows.map((tr, r) => {
    const cells = [...tr.children];
    while (cells.length < cols) cells.push(null);
    return '<hp:tr>' + cells.map((td, c) => {
      const sub = new Writer(w.doc, w.images);
      if (td) walk(td, td.tagName === 'TH' ? { ...st, b: true } : st, { align: blockAlign(td) || 'LEFT', left: 0 }, sub);
      sub.flush();
      if (!sub.out.length) sub.out.push(sub.paraXml({ runs: [], align: 'LEFT', left: 0 }));
      return `<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="3"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${sub.out.join('')}</hp:subList><hp:cellAddr colAddr="${c}" rowAddr="${r}"/><hp:cellSpan colSpan="1" rowSpan="1"/><hp:cellSz width="${cw}" height="${rh}"/><hp:cellMargin left="510" right="510" top="141" bottom="141"/></hp:tc>`;
    }).join('') + '</hp:tr>';
  }).join('');
  return `<hp:tbl id="${w.doc.uid++}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" rowCnt="${rows.length}" colCnt="${cols}" cellSpacing="0" borderFillIDRef="3" noAdjust="0"><hp:sz width="${cw * cols}" widthRelTo="ABSOLUTE" height="${rh * rows.length}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="283" right="283" top="283" bottom="283"/><hp:inMargin left="510" right="510" top="141" bottom="141"/>${trs}</hp:tbl>`;
}

function picXml(img, w) {
  const im = w.images.get(img.getAttribute('data-fid') || img.getAttribute('src'));
  if (!im) return '';
  const pct = parseFloat((img.getAttribute('style') || '').match(/width:\s*([\d.]+)%/)?.[1] || '100') / 100;
  let width = im.w * PX, height = im.h * PX;
  const maxW = PAGE_W * Math.min(1, pct);
  if (width > maxW || pct < 1) { height = Math.round(height * maxW / width); width = Math.round(maxW); }
  const id = 'image' + im.n;
  return `<hp:pic id="${w.doc.uid++}" zOrder="0" numberingType="PICTURE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" instid="${w.doc.uid++}" reverse="0"><hp:offset x="0" y="0"/><hp:orgSz width="${width}" height="${height}"/><hp:curSz width="0" height="0"/><hp:flip horizontal="0" vertical="0"/><hp:rotationInfo angle="0" centerX="0" centerY="0" rotateimage="1"/><hp:renderingInfo><hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/></hp:renderingInfo><hp:imgRect><hc:pt0 x="0" y="0"/><hc:pt1 x="${width}" y="0"/><hc:pt2 x="${width}" y="${height}"/><hc:pt3 x="0" y="${height}"/></hp:imgRect><hp:imgClip left="0" right="0" top="0" bottom="0"/><hp:inMargin left="0" right="0" top="0" bottom="0"/><hc:img binaryItemIDRef="${id}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/><hp:effects/><hp:sz width="${width}" widthRelTo="ABSOLUTE" height="${height}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="0" right="0" top="0" bottom="0"/><hp:shapeComment>그림입니다.</hp:shapeComment></hp:pic>`;
}

// 그림 파일을 한글이 읽는 형식(png/jpg/gif/bmp)으로 준비
async function loadImages(roots, getFile) {
  const map = new Map(); let n = 0;
  for (const root of roots) for (const img of root.querySelectorAll('img')) {
    const key = img.getAttribute('data-fid') || img.getAttribute('src');
    if (!key || map.has(key)) continue;
    try {
      let blob = img.getAttribute('data-fid') ? await getFile(img.getAttribute('data-fid')) : await (await fetch(key)).blob();
      if (!blob) continue;
      const bmp = await createImageBitmap(blob);
      let ext = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/bmp': 'bmp' }[blob.type];
      if (!ext) { // webp 등은 png로 바꿔서
        const c = document.createElement('canvas'); c.width = bmp.width; c.height = bmp.height; c.getContext('2d').drawImage(bmp, 0, 0);
        blob = await new Promise(r => c.toBlob(r, 'image/png')); ext = 'png';
      }
      map.set(key, { n: ++n, ext, mime: blob.type, w: bmp.width, h: bmp.height, data: blob });
    } catch { /* 못 읽은 그림은 [그림]으로 */ }
  }
  return map;
}

// pages: [{title, date, html}] → Blob(.hwpx)
export async function toHwpx(pages, getFile) {
  const doc = new Doc();
  const roots = pages.map(p => parseHTML(p.html));
  const images = await loadImages(roots, getFile);
  const w = new Writer(doc, images);
  pages.forEach((p, i) => {
    const base = { align: 'JUSTIFY', left: 0 };
    w.start({ ...base, pageBreak: i > 0 });
    w.text(p.title || '제목 없음', { size: 18, b: true }, base); w.flush();
    if (p.date) { w.start({ ...base }); w.text(p.date, { size: 9, color: '#777777' }, base); w.flush(); }
    w.start({ ...base }); w.flush();
    walk(roots[i], {}, { ...base }, w);
    w.flush();
  });
  const section = T.SEC_OPEN + w.out.join('').replace(/^(<hp:p [^>]*>)/, (m) => m + T.SECPR_RUN) + '</hs:sec>';
  const now = new Date();
  const items = [...images.values()].map(im => `<opf:item id="image${im.n}" href="BinData/image${im.n}.${im.ext}" media-type="${im.mime}" isEmbeded="1"/>`).join('');
  const hpf = T.HPF.replace('{{TITLE}}', esc(pages[0]?.title || '노트')).replaceAll('{{ISO}}', now.toISOString().replace(/\.\d+Z$/, 'Z'))
    .replace('{{DATE}}', esc(now.toLocaleString('ko-KR'))).replace('{{ITEMS}}', items);
  const zip = new JSZip();
  zip.file('mimetype', 'application/hwp+zip', { compression: 'STORE' });
  zip.file('version.xml', T.VERSION);
  zip.file('Contents/header.xml', doc.header());
  zip.file('Contents/section0.xml', section);
  zip.file('Contents/content.hpf', hpf);
  for (const im of images.values()) zip.file(`BinData/image${im.n}.${im.ext}`, im.data);
  zip.file('settings.xml', T.SETTINGS);
  zip.file('META-INF/container.xml', T.CONTAINER);
  zip.file('META-INF/container.rdf', T.CONTAINER_RDF);
  zip.file('META-INF/manifest.xml', T.MANIFEST);
  zip.file('Preview/PrvText.txt', roots.map((r, i) => (pages[i].title || '') + '\n' + r.textContent).join('\n').slice(0, 1000));
  return zip.generateAsync({ type: 'blob', mimeType: 'application/hwp+zip', compression: 'DEFLATE', createFolders: false });
}
