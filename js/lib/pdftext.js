// Turns pdf.js text items into lines, paragraphs (reading flow) or rows (grid).
import { getPdfjs } from './loader.js';

const PUA = { '\uf0b7': '•', '\uf0a7': '▪', '\uf0d8': '➢', '\uf0fc': '✓', '\uf076': '❖', '\uf06e': '■', '\uf06c': '●', '\uf0e0': '➔', '\uf0a8': '◆' };
const fixPua = (t) => t.replace(/[\uf000-\uf0ff]/g, (c) => PUA[c] || '•');
const BULLET_RE = /^\s*([•●▪◦■□➢➤►‣⁃✓❖◆➔\-–—*]|\(?\d{1,3}[.)]|\(?[a-zA-Z][.)]|[ivxIVX]{1,4}[.)])\s+/;

async function fontStyles(page) {
  const styles = new Map();
  try {
    await page.getOperatorList();
  } catch { return styles; }
  return {
    get(name) {
      if (styles.has(name)) return styles.get(name);
      let st = { bold: false, italic: false };
      try {
        if (page.commonObjs.has(name)) {
          const f = page.commonObjs.get(name);
          const n = `${f.name || ''} ${f.loadedName || ''}`;
          st = {
            bold: !!f.bold || /bold|black|heavy|semibold|demi/i.test(n),
            italic: !!f.italic || /italic|oblique/i.test(n),
          };
        }
      } catch { /* font not resolved */ }
      styles.set(name, st);
      return st;
    },
  };
}

// Raw positioned items, top-left origin in PDF points.
export async function pageItems(page, { withFonts = false } = {}) {
  const { Util } = await getPdfjs();
  const viewport = page.getViewport({ scale: 1 });
  const content = await page.getTextContent();
  const fonts = withFonts ? await fontStyles(page) : null;
  const items = [];
  for (const it of content.items) {
    if (typeof it.str !== 'string' || !it.str.length) continue;
    const tx = Util.transform(viewport.transform, it.transform);
    const size = Math.hypot(tx[2], tx[3]) || it.height || 10;
    const scaleX = Math.hypot(tx[0], tx[1]) / (Math.hypot(it.transform[0], it.transform[1]) || 1);
    const w = (it.width || 0) * (scaleX || 1);
    const st = fonts ? fonts.get(it.fontName) : { bold: false, italic: false };
    items.push({
      str: fixPua(it.str), x0: tx[4], x1: tx[4] + w, y: tx[5], size,
      bold: st.bold, italic: st.italic, eol: !!it.hasEOL,
    });
  }
  return { width: viewport.width, height: viewport.height, items };
}

function addRun(line, it, withSpace) {
  const runs = line.runs;
  const text = (withSpace ? ' ' : '') + it.str;
  const last = runs[runs.length - 1];
  if (last && last.bold === it.bold && last.italic === it.italic) last.text += text;
  else runs.push({ text, bold: it.bold, italic: it.italic });
  line.x1 = Math.max(line.x1, it.x1);
  const n = it.str.trim().length;
  line.w += n; line.ws += n * it.size;
  if (line.w) line.size = line.ws / line.w;
}

// Lines in content-stream order (keeps multi-column reading order).
export function flowLines(items) {
  const lines = [];
  let cur = null;
  for (const it of items) {
    if (!it.str.trim() && !cur) continue;
    const tol = 0.5 * Math.min(it.size, cur ? cur.size : it.size);
    const sameLine = cur && Math.abs(it.y - cur.y) < tol && it.x0 >= cur.x1 - it.size * 1.5;
    if (sameLine && !it.str.trim()) {
      // Blank spacer items: remember a space but don't let them hide column gaps.
      cur.pendingSpace = true;
      continue;
    }
    if (sameLine) {
      const gap = it.x0 - cur.x1;
      const lastText = cur.runs[cur.runs.length - 1].text;
      if (gap > it.size * 1.6 && it.str.trim() && lastText.trim()) {
        // A wide gap means table columns: keep them apart with a tab.
        cur.runs[cur.runs.length - 1].text = lastText.replace(/\s+$/, '');
        addRun(cur, { ...it, str: `\t${it.str.replace(/^\s+/, '')}` }, false);
        cur.stops.push(it.x0);
        cur.pendingSpace = false;
      } else {
        const needSpace = (cur.pendingSpace || gap > it.size * 0.18) && !/\s$/.test(lastText) && !/^\s/.test(it.str);
        addRun(cur, it, needSpace);
        cur.pendingSpace = false;
      }
    } else {
      if (!it.str.trim()) continue;
      cur = { y: it.y, x0: it.x0, x1: it.x1, size: it.size, runs: [], stops: [], w: 0, ws: 0 };
      addRun(cur, { ...it, str: it.str.replace(/^\s+/, '') }, false);
      lines.push(cur);
    }
  }
  for (const l of lines) {
    l.text = l.runs.map((r) => r.text).join('').replace(/\s+$/, '');
    const chars = l.runs.reduce((n, r) => n + r.text.trim().length, 0) || 1;
    l.boldShare = l.runs.reduce((n, r) => n + (r.bold ? r.text.trim().length : 0), 0) / chars;
  }
  return lines.filter((l) => l.text.trim());
}

function median(arr) {
  if (!arr.length) return 0;
  const s = [...arr].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)];
}

// Most common text size, weighted by characters.
export function bodySize(lines) {
  const counts = new Map();
  for (const l of lines) {
    const k = Math.round(l.size * 2) / 2;
    counts.set(k, (counts.get(k) || 0) + l.text.length);
  }
  let best = 11, n = -1;
  for (const [k, v] of counts) if (v > n) { best = k; n = v; }
  return best;
}

// Groups lines into paragraphs with heading levels.
export function paragraphs(lines, pageWidth, body) {
  const gaps = [];
  for (let i = 1; i < lines.length; i++) {
    const a = lines[i - 1], b = lines[i];
    const g = b.y - a.y;
    if (g > 0 && Math.abs(a.size - b.size) < 0.5) gaps.push(g / a.size);
  }
  const lineGap = Math.min(Math.max(median(gaps) || 1.2, 1.0), 1.8);
  const right = Math.max(...lines.map((l) => l.x1), 0);
  const left = Math.min(...lines.map((l) => l.x0), pageWidth);

  const paras = [];
  let p = null;
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];
    const prev = lines[i - 1];
    let start = !p;
    if (p && prev) {
      const gap = (l.y - prev.y) / prev.size;
      const sizeChange = Math.abs(l.size - prev.size) / prev.size > 0.15;
      const ended = /[.?!:"”’)]$/.test(prev.text) && prev.x1 < right - prev.size * 4;
      const indented = l.x0 > prev.x0 + l.size * 1.2 && /[.?!:"”’]$/.test(prev.text);
      start = gap < -0.1 || gap > lineGap * 1.35 + 0.15 || sizeChange || ended || indented
        || BULLET_RE.test(l.text) || Math.abs(l.x0 - prev.x0) > l.size * 6
        || l.stops.length > 0 || prev.stops.length > 0;
    }
    if (start) {
      p = { lines: [], size: l.size, x0: l.x0, y: l.y };
      paras.push(p);
    }
    p.lines.push(l);
  }

  return paras.map((para) => {
    const runs = [];
    para.lines.forEach((l, i) => {
      l.runs.forEach((r, j) => {
        let text = r.text;
        if (j === l.runs.length - 1) text = text.replace(/\s+$/, '');
        if (i > 0 && j === 0) {
          const last = runs[runs.length - 1];
          if (last && /[A-Za-z]-$/.test(last.text) && /^[a-z]/.test(text)) last.text = last.text.slice(0, -1);
          else if (last && !/\s$/.test(last.text)) text = ' ' + text;
        }
        const last = runs[runs.length - 1];
        if (last && last.bold === r.bold && last.italic === r.italic) last.text += text;
        else runs.push({ text, bold: r.bold, italic: r.italic });
      });
    });
    const text = runs.map((r) => r.text).join('');
    const ratio = para.size / body;
    const short = text.length < 160 && para.lines.length <= 3;
    let heading = 0;
    if (short && ratio >= 1.6) heading = 1;
    else if (short && ratio >= 1.2) heading = 2;
    else if (short && ratio >= 0.98 && para.lines.every((l) => l.boldShare > 0.9) && text.length < 90 && !/[.,;]$/.test(text)) heading = 3;
    const l0 = para.lines[0];
    const mid = (l0.x0 + Math.max(...para.lines.map((l) => l.x1))) / 2;
    const centered = Math.abs(mid - pageWidth / 2) < pageWidth * 0.04
      && l0.x0 > left + l0.size * 2 && para.lines.length <= 4;
    const tabbed = para.lines.length === 1 && l0.stops.length > 0;
    return {
      runs, text, size: para.size, heading: tabbed ? 0 : heading, centered: tabbed ? false : centered, tabStops: tabbed ? l0.stops.map((x) => x - left) : null,
      bullet: BULLET_RE.test(text), y: para.y, indent: Math.max(0, para.x0 - left),
    };
  });
}

// Rows of cells by geometry (for tables / spreadsheets).
export function gridRows(items) {
  const parts = [];
  for (const it of items) {
    if (!it.str.trim()) continue;
    // Split items that hide several cells behind runs of spaces.
    const segs = it.str.split(/(\s{2,}|\t)/);
    if (segs.length === 1) { parts.push(it); continue; }
    const per = (it.x1 - it.x0) / Math.max(it.str.length, 1);
    let pos = 0;
    for (const s of segs) {
      if (s.trim() && !/^(\s{2,}|\t)$/.test(s)) {
        const lead = s.length - s.trimStart().length;
        parts.push({ ...it, str: s.trim(), x0: it.x0 + (pos + lead) * per, x1: it.x0 + (pos + lead + s.trim().length) * per });
      }
      pos += s.length;
    }
  }
  parts.sort((a, b) => a.y - b.y || a.x0 - b.x0);
  const rows = [];
  for (const it of parts) {
    const row = rows[rows.length - 1];
    if (row && Math.abs(it.y - row.y) < Math.min(it.size, row.size) * 0.45) {
      row.items.push(it);
      row.size = Math.max(row.size, it.size);
    } else rows.push({ y: it.y, size: it.size, items: [it] });
  }
  for (const row of rows) {
    row.items.sort((a, b) => a.x0 - b.x0);
    const cells = [];
    for (const it of row.items) {
      const c = cells[cells.length - 1];
      const gap = c ? it.x0 - c.x1 : Infinity;
      if (c && gap < it.size * 0.75) {
        c.text += (gap > it.size * 0.15 && !/\s$/.test(c.text) ? ' ' : '') + it.str;
        c.x1 = Math.max(c.x1, it.x1);
      } else cells.push({ text: it.str, x0: it.x0, x1: it.x1 });
    }
    row.cells = cells.map((c) => ({ ...c, text: c.text.trim() }));
  }
  return rows;
}
