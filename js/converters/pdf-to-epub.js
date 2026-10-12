import { openPdf, closePdf } from '../lib/loader.js';
import { baseName, canvasToBlob, escapeXml, renderPdfPage, safeScale, tick } from '../lib/util.js';
import { pageItems, flowLines, paragraphs, bodySize } from '../lib/pdftext.js';
import { writeEpub } from '../lib/epub-writer.js';

const norm = (t) => t.trim().toLowerCase().replace(/\d+/g, '#').replace(/\s+/g, ' ');

// Drops running headers, footers and page numbers that repeat across pages.
function stripFurniture(pages) {
  const zone = (l, p) => l.y < p.height * 0.09 || l.y > p.height * 0.91;
  const counts = new Map();
  for (const p of pages) {
    const seen = new Set();
    for (const l of p.lines) if (zone(l, p)) seen.add(norm(l.text));
    for (const k of seen) counts.set(k, (counts.get(k) || 0) + 1);
  }
  const limit = Math.max(2, pages.length * 0.4);
  for (const p of pages) {
    p.lines = p.lines.filter((l) => !(zone(l, p) && (counts.get(norm(l.text)) >= limit
      || /^\W*(page\s*)?\d{1,4}(\s*(of|\/)\s*\d+)?\W*$/i.test(l.text.trim()))));
  }
}

function runsHtml(runs) {
  return runs.map((r) => {
    let t = escapeXml(r.text).replace(/\t/g, '\u2003\u2003');
    if (r.italic) t = `<i>${t}</i>`;
    if (r.bold) t = `<b>${t}</b>`;
    return t;
  }).join('');
}

export async function convert(file, { report }) {
  const pdf = await openPdf(file);
  const n = pdf.numPages;
  const pages = [];
  const images = [];
  let cover = null;

  for (let i = 1; i <= n; i++) {
    report((i - 1) / n * 0.85, `Reading page ${i} of ${n}`);
    const page = await pdf.getPage(i);
    const data = await pageItems(page, { withFonts: true });
    const lines = flowLines(data.items);
    const chars = lines.reduce((s, l) => s + l.text.trim().length, 0);
    const entry = { height: data.height, width: data.width, lines, index: i };
    if (i === 1 || chars < 15) {
      const canvas = await renderPdfPage(page, safeScale(page, 1400 / data.height, 4e6));
      const bytes = new Uint8Array(await (await canvasToBlob(canvas, 'image/jpeg', 0.85)).arrayBuffer());
      canvas.width = canvas.height = 0;
      if (i === 1) { cover = 'cover.jpg'; images.push({ name: cover, data: bytes }); }
      if (chars < 15) {
        entry.image = i === 1 ? cover : `page-${i}.jpg`;
        if (i !== 1) images.push({ name: entry.image, data: bytes });
      }
    }
    pages.push(entry);
    page.cleanup();
    await tick();
  }

  let meta = {};
  try { meta = (await pdf.getMetadata()).info || {}; } catch { /* no metadata */ }
  await closePdf(pdf);
  report(0.9, 'Arranging chapters');

  stripFurniture(pages);
  const body = bodySize(pages.flatMap((p) => p.lines));

  // Flow of blocks across the whole book.
  const blocks = [];
  for (const p of pages) {
    if (p.image) {
      if (p.index !== 1) blocks.push({ kind: 'img', src: p.image, page: p.index });
      continue;
    }
    const paras = paragraphs(p.lines, p.width, body);
    paras.forEach((para, j) => {
      const prev = blocks[blocks.length - 1];
      // Join a paragraph that continues across a page break.
      if (j === 0 && prev && prev.kind === 'p' && !prev.heading && !para.heading && !para.bullet && !para.tabStops && !prev.tabStops
        && !/[.?!:"”’)]$/.test(prev.text) && /^[a-z(“"']/.test(para.text)) {
        const glue = /[A-Za-z]-$/.test(prev.text) ? '' : ' ';
        if (!glue) prev.runs[prev.runs.length - 1].text = prev.runs[prev.runs.length - 1].text.slice(0, -1);
        prev.runs.push({ ...para.runs[0], text: glue + para.runs[0].text }, ...para.runs.slice(1));
        prev.text += glue + para.text;
        return;
      }
      blocks.push({ kind: 'p', ...para, page: p.index });
    });
  }

  const h1 = blocks.filter((b) => b.heading === 1).length;
  const h2 = blocks.filter((b) => b.heading === 2).length;
  const splitLevel = h1 >= 2 ? 1 : h2 >= 2 ? 2 : 0;
  const title = (meta.Title && String(meta.Title).trim()) || baseName(file.name);

  const chapters = [];
  let cur = null;
  const startChapter = (name) => { cur = { title: name, parts: [] }; chapters.push(cur); };
  for (const b of blocks) {
    if (!cur) startChapter(splitLevel || n <= 10 ? title : 'Pages 1–10');
    if (splitLevel && b.kind === 'p' && b.heading && b.heading <= splitLevel && cur.parts.length) startChapter(b.text.slice(0, 80));
    else if (!splitLevel && b.page > chapters.length * 10 && cur.parts.length) startChapter(`Pages ${chapters.length * 10 + 1}–${Math.min(n, chapters.length * 10 + 10)}`);
    if (splitLevel && b.kind === 'p' && b.heading && b.heading <= splitLevel && cur.parts.length === 0) cur.title = b.text.slice(0, 80);
    if (b.kind === 'img') cur.parts.push(`<div class="page"><img src="images/${b.src}" alt="Page ${b.page}"/></div>`);
    else if (b.heading) cur.parts.push(`<h${b.heading}${b.centered ? ' class="c"' : ''}>${runsHtml(b.runs)}</h${b.heading}>`);
    else cur.parts.push(`<p${b.centered ? ' class="c"' : b.bullet ? ' class="li"' : ''}>${runsHtml(b.runs)}</p>`);
  }
  if (!chapters.length) startChapter(title);
  if (!blocks.length) throw new Error('No text was found in this PDF. It may be a scanned image, which needs OCR to become an e-book.');

  report(0.95, 'Packing the EPUB');
  const blob = await writeEpub({
    title, author: meta.Author ? String(meta.Author).trim() : '',
    chapters: chapters.map((c) => ({ title: c.title, html: c.parts.join('\n') || '<p></p>' })),
    images, cover,
  });
  return { blob, name: `${baseName(file.name)}.epub`, detail: `${chapters.length} chapter${chapters.length > 1 ? 's' : ''}, ${n} source page${n > 1 ? 's' : ''}` };
}
