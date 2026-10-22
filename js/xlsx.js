// 노트 페이지(HTML) → 엑셀(.xlsx)
// 시트 구성: [내용] 페이지 글을 한 줄씩 · [표1], [표2]… 페이지 안의 표 · (섹션 전체일 때) [페이지 목록]
import { parseHTML, toLines } from './doc.js';

const num = v => { const t = v.replace(/,/g, ''); return /^-?\d+(\.\d+)?$/.test(t) && t.length < 16 ? +t : v; };
function sheetName(name, used) {
  let base = name.replace(/[\[\]:*?/\\]/g, ' ').trim().slice(0, 28) || '시트', n = base, k = 2;
  while (used.has(n)) n = `${base.slice(0, 25)} (${k++})`;
  used.add(n); return n;
}
function addSheet(wb, used, name, rows, widths) {
  const ws = XLSX.utils.aoa_to_sheet(rows);
  if (widths) ws['!cols'] = widths.map(wch => ({ wch }));
  XLSX.utils.book_append_sheet(wb, ws, sheetName(name, used));
}

// pages: [{title, date, created, updated, html}], whole: 섹션 전체 변환이면 true
export function toXlsx(pages, whole) {
  const wb = XLSX.utils.book_new(), used = new Set();
  const parsed = pages.map(p => ({ p, ...toLines(parseHTML(p.html)) }));
  if (whole) {
    addSheet(wb, used, '페이지 목록', [['제목', '만든 날', '고친 날', '첫 줄'],
      ...parsed.map(({ p, lines }) => [p.title || '제목 없음', p.created || '', p.updated || '', lines[0] || ''])], [30, 18, 18, 60]);
  }
  const rows = whole ? [['페이지', '내용']] : [['내용']];
  parsed.forEach(({ p, lines }) => {
    if (!lines.length) { if (whole) rows.push([p.title || '제목 없음', '']); return; }
    lines.forEach((l, i) => rows.push(whole ? [i ? '' : p.title || '제목 없음', l] : [l]));
  });
  addSheet(wb, used, '내용', rows, whole ? [24, 90] : [100]);
  let t = 0;
  parsed.forEach(({ p, tables }) => tables.forEach(tb => {
    t++;
    addSheet(wb, used, whole ? `표${t} ${p.title || ''}` : `표${t}`, tb.map(r => r.map(num)),
      Array.from({ length: Math.max(1, ...tb.map(r => r.length)) }, (_, c) => Math.min(50, Math.max(8, ...tb.map(r => String(r[c] || '').length * 2)))));
  }));
  const data = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  return new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
