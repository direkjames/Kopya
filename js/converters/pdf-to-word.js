import { openPdf, closePdf, getDocx } from '../lib/loader.js';
import { baseName, canvasToBlob, renderPdfPage, safeScale, tick } from '../lib/util.js';
import { pageItems, flowLines, paragraphs, bodySize } from '../lib/pdftext.js';

const TW = 20; // twips per point

export async function convert(file, { report, options }) {
  const exact = options.mode === 'exact';
  const docx = await getDocx();
  const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, AlignmentType, Tab, TabStopType } = docx;
  const pdf = await openPdf(file);
  const n = pdf.numPages;

  const pages = [];
  for (let i = 1; i <= n; i++) {
    report((i - 1) / n * 0.9, `Reading page ${i} of ${n}`);
    const page = await pdf.getPage(i);
    const data = await pageItems(page, { withFonts: !exact });
    const lines = exact ? [] : flowLines(data.items);
    const chars = lines.reduce((s, l) => s + l.text.trim().length, 0);
    const entry = { width: data.width, height: data.height, lines };
    if (exact || chars < 15) {
      const canvas = await renderPdfPage(page, safeScale(page, 150 / 72));
      entry.image = new Uint8Array(await (await canvasToBlob(canvas, 'image/jpeg', 0.88)).arrayBuffer());
      canvas.width = canvas.height = 0;
    }
    pages.push(entry);
    page.cleanup();
    await tick();
  }
  await closePdf(pdf);
  report(0.92, 'Building the Word document');

  const allLines = pages.flatMap((p) => p.lines);
  const body = bodySize(allLines);
  let scanned = 0;
  const sections = [];

  pages.forEach((pg, idx) => {
    const children = [];
    const xs = pg.lines.map((l) => l.x0);
    const xe = pg.lines.map((l) => l.x1);
    const clamp = (v) => Math.min(Math.max(v, 28), 90);
    let margin = { top: 54, bottom: 54, left: 54, right: 54 };
    if (pg.image) {
      margin = exact ? { top: 0, bottom: 0, left: 0, right: 0 } : { top: 18, bottom: 18, left: 18, right: 18 };
      if (!exact) scanned++;
      const w = pg.width - margin.left - margin.right;
      const h = pg.height - margin.top - margin.bottom - (exact ? 2 : 0);
      const fit = Math.min(w / pg.width, h / pg.height);
      children.push(new Paragraph({
        spacing: { before: 0, after: 0, line: 240 },
        children: [new ImageRun({
          type: 'jpg', data: pg.image,
          transformation: { width: Math.floor(pg.width * fit * 96 / 72), height: Math.floor(pg.height * fit * 96 / 72) },
        })],
      }));
    } else {
      if (xs.length) {
        margin.left = clamp(Math.min(...xs) - 2);
        margin.right = clamp(pg.width - Math.max(...xe) - 2);
        margin.top = clamp(Math.min(...pg.lines.map((l) => l.y - l.size)) - 4);
      }
      for (const para of paragraphs(pg.lines, pg.width, body)) {
        const size = Math.round(para.size * 2) / 2;
        const opts = {
          children: para.runs.map((r) => new TextRun({
            children: r.text.split('\t').flatMap((t, k) => (k ? [new Tab(), t] : [t])).filter((x) => x !== ''),
            bold: r.bold || (para.heading > 0 && para.heading < 3), italics: r.italic, size: Math.round(size * 2),
          })),
          spacing: { after: Math.round(size * (para.heading ? 0.5 : 0.6) * TW), line: 264 },
        };
        if (para.tabStops) opts.tabStops = para.tabStops.map((x) => ({ type: TabStopType.LEFT, position: Math.round(x * TW) }));
        if (para.heading) opts.heading = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3][para.heading - 1];
        if (para.centered) opts.alignment = AlignmentType.CENTER;
        else if (para.indent > 18) opts.indent = { left: Math.round(Math.min(para.indent, 200) * TW) };
        children.push(new Paragraph(opts));
      }
      if (!children.length) children.push(new Paragraph({ children: [] }));
    }
    sections.push({
      properties: {
        page: {
          size: { width: Math.round(pg.width * TW), height: Math.round(pg.height * TW) },
          margin: { top: margin.top * TW, bottom: margin.bottom * TW, left: margin.left * TW, right: margin.right * TW, header: 0, footer: 0 },
        },
      },
      children,
    });
    if (idx % 20 === 0) report(0.92 + 0.05 * idx / pages.length, 'Building the Word document');
  });

  const heading = (size) => ({ run: { color: '000000', font: 'Calibri', size: Math.round(size * 2) }, paragraph: { spacing: { before: 120 } } });
  const doc = new Document({
    creator: 'Kopya',
    styles: {
      default: {
        document: { run: { font: 'Calibri', size: Math.round(body * 2) } },
        heading1: heading(body * 1.8), heading2: heading(body * 1.4), heading3: heading(body),
      },
    },
    sections,
  });
  const blob = await Packer.toBlob(doc);
  let detail = exact ? `${n} pages kept as pictures` : `${n} page${n > 1 ? 's' : ''} of editable text`;
  if (!exact && scanned) detail += `; ${scanned} page${scanned > 1 ? 's' : ''} had no text and ${scanned > 1 ? 'were' : 'was'} added as a picture`;
  return { blob, name: `${baseName(file.name)}.docx`, detail };
}
