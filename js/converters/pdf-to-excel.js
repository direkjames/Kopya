import { openPdf, closePdf } from '../lib/loader.js';
import { baseName, tick } from '../lib/util.js';
import { pageItems, gridRows } from '../lib/pdftext.js';
import { writeXlsx } from '../lib/xlsx-writer.js';

// "1,234.50", "-12", "(300)" become numbers; dates, IDs with leading zeros stay text.
function toValue(s) {
  const t = s.trim();
  if (!/^[(-]?[₱$€£¥]?\s?\d{1,3}(,\d{3})*(\.\d+)?\)?$|^[(-]?[₱$€£¥]?\s?\d+(\.\d+)?\)?$/.test(t)) return t;
  if (/^0\d/.test(t.replace(/^[(-]/, ''))) return t;
  const neg = /^\(.*\)$/.test(t) || t.startsWith('-');
  const num = Number(t.replace(/[(),\s₱$€£¥-]/g, ''));
  if (!Number.isFinite(num) || t.replace(/\D/g, '').length > 15) return t;
  return neg ? -num : num;
}

// Columns = horizontal bands where cells from multi-cell rows overlap.
function columnBands(rows) {
  const spans = [];
  for (const r of rows) if (r.cells.length > 1) for (const c of r.cells) spans.push([c.x0 - 1.5, c.x1 + 1.5]);
  if (!spans.length) return [];
  spans.sort((a, b) => a[0] - b[0]);
  const bands = [spans[0].slice()];
  for (const s of spans.slice(1)) {
    const b = bands[bands.length - 1];
    if (s[0] <= b[1]) b[1] = Math.max(b[1], s[1]);
    else bands.push(s.slice());
  }
  return bands;
}

function mergeLonelyBands(bands, rows) {
  const hit = (r, b) => r.cells.some((c) => Math.min(b[1], c.x1) - Math.max(b[0], c.x0) > 0);
  const multi = rows.filter((r) => r.cells.length > 1);
  let i = 0;
  while (i < bands.length - 1) {
    const a = bands[i], b = bands[i + 1];
    const together = multi.some((r) => hit(r, a) && hit(r, b));
    const countA = multi.filter((r) => hit(r, a)).length;
    const countB = multi.filter((r) => hit(r, b)).length;
    // Only merge when one side is sparse, like a lone header.
    if (!together && Math.min(countA, countB) <= Math.max(2, Math.max(countA, countB) * 0.15)) {
      bands.splice(i, 2, [a[0], b[1]]);
    } else i++;
  }
  return bands;
}

function pageTable(rows) {
  const bands = mergeLonelyBands(columnBands(rows), rows);
  const out = [];
  for (const r of rows) {
    if (!bands.length || r.cells.length === 1 && !bands.some((b) => r.cells[0].x0 >= b[0] && r.cells[0].x1 <= b[1])) {
      // Titles, notes and other free text go in the first column.
      out.push([toValue(r.cells.map((c) => c.text).join(' '))]);
      continue;
    }
    const line = new Array(bands.length).fill('');
    for (const c of r.cells) {
      let best = 0, bestOverlap = -Infinity;
      bands.forEach((b, i) => {
        const ov = Math.min(b[1], c.x1) - Math.max(b[0], c.x0);
        const score = ov > 0 ? ov : -Math.abs((b[0] + b[1]) / 2 - (c.x0 + c.x1) / 2);
        if (score > bestOverlap) { bestOverlap = score; best = i; }
      });
      line[best] = line[best] ? `${line[best]} ${c.text}` : c.text;
    }
    out.push(line.map((v) => (v === '' ? '' : toValue(v))));
  }
  return out;
}

export async function convert(file, { report, options }) {
  const oneSheet = options.layout === 'single';
  const pdf = await openPdf(file);
  const n = pdf.numPages;
  const sheets = [];
  let cellCount = 0;
  for (let i = 1; i <= n; i++) {
    report((i - 1) / n * 0.9, `Reading page ${i} of ${n}`);
    const page = await pdf.getPage(i);
    const { items } = await pageItems(page);
    const rows = gridRows(items);
    const table = pageTable(rows);
    cellCount += table.reduce((s, r) => s + r.filter((v) => v !== '').length, 0);
    sheets.push({ name: `Page ${i}`, rows: table });
    page.cleanup();
    await tick();
  }
  await closePdf(pdf);
  if (!cellCount) throw new Error('No text was found in this PDF. It may be a scanned image, which needs OCR before its tables can be read.');
  report(0.95, 'Writing the spreadsheet');
  let out = sheets;
  if (oneSheet) {
    const rows = [];
    sheets.forEach((s, i) => { if (i) rows.push([]); rows.push(...s.rows); });
    out = [{ name: 'All pages', rows }];
  }
  const blob = await writeXlsx(out);
  return {
    blob, name: `${baseName(file.name)}.xlsx`,
    detail: oneSheet ? `${n} page${n > 1 ? 's' : ''} in one sheet` : `${n} sheet${n > 1 ? 's' : ''}, one per page`,
  };
}
