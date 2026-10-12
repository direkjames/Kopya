import { getJSZip } from '../lib/loader.js';
import { baseName } from '../lib/util.js';
import { DocxRenderer, defaultFont } from '../lib/docx-render.js';
import { countPages, createStage, pdfBuilder, renderFlow, waitForImages } from '../lib/paginate.js';

const CSS = `
.kp-docx{color:#000;white-space:pre-wrap;tab-size:48px;-moz-tab-size:48px;overflow-wrap:break-word;text-align:left}
.kp-docx p{margin:0}
.kp-docx table{white-space:pre-wrap}
.kp-docx td>p:last-child{padding-bottom:0!important}
.kp-docx img{display:inline-block}
.kp-docx a{color:inherit}
`;

export async function convert(file, { report, options }) {
  const head = new Uint8Array(await file.slice(0, 4).arrayBuffer());
  if (head[0] === 0xd0 && head[1] === 0xcf) {
    throw new Error('This is the older .doc format. Open it in Word or Google Docs and save it as .docx, then try again.');
  }
  const JSZip = await getJSZip();
  let zip;
  try { zip = await JSZip.loadAsync(file); } catch { throw new Error("This file couldn't be opened as a Word document. It may be damaged."); }
  report(0.05, 'Reading the document');
  const r = new DocxRenderer(zip);
  try { await r.load(); } catch { throw new Error("This file doesn't look like a Word document (.docx)."); }

  const stage = createStage(CSS);
  try {
    const blocks = await r.body_();
    const page = r.page;
    const contentW = page.width - page.margin.left - page.margin.right;
    const font = defaultFont(r);
    const root = document.createElement('div');
    root.className = 'kp-docx';
    root.style.cssText = `width:${contentW}px;display:flow-root;font-family:${font.family};font-size:${font.size}pt;`;
    for (const b of blocks) root.appendChild(b);
    stage.appendChild(root);
    report(0.12, 'Laying out pages');
    await waitForImages(root);
    if (document.fonts && document.fonts.ready) await document.fonts.ready;

    const hf = {
      header: await r.headerFooter('header', 'default'), footer: await r.headerFooter('footer', 'default'),
      headerFirst: r.titlePg ? await r.headerFooter('header', 'first') : null,
      footerFirst: r.titlePg ? await r.headerFooter('footer', 'first') : null,
    };
    const needsTotal = Object.values(hf).some((x) => x && x.querySelector('.kp-pages'));
    const total = needsTotal ? countPages(root, page) : 0;

    const place = (frame, box, no, isHeader) => {
      if (!box) return;
      const el = box.cloneNode(true);
      el.className = 'kp-docx';
      el.style.cssText = `position:absolute;left:${page.margin.left}px;width:${contentW}px;font-family:${font.family};font-size:${font.size}pt;${isHeader ? `top:${page.headerAt}px` : `bottom:${page.footerAt}px`}`;
      el.querySelectorAll('.kp-page').forEach((s) => { s.textContent = String(no); });
      el.querySelectorAll('.kp-pages').forEach((s) => { s.textContent = String(total); });
      frame.appendChild(el);
    };
    const decorate = (frame, no) => {
      const first = no === 1 && r.titlePg;
      place(frame, first ? hf.headerFirst : hf.header, no, true);
      place(frame, first ? hf.footerFirst : hf.footer, no, false);
    };

    const builder = await pdfBuilder();
    const pixelRatio = options.quality === 'small' ? 1.5 : 2;
    const made = await renderFlow({
      root, stage, page, builder, pixelRatio, quality: 0.88, decorate, totalPages: total,
      onPage: (k) => report(0.15 + 0.8 * Math.min(1, k / Math.max(total || k + 2, 1)), `Drawing page ${k}`),
    });
    report(0.97, 'Saving the PDF');
    const blob = await builder.save();
    const notes = [...r.warnings];
    return { blob, name: `${baseName(file.name)}.pdf`, detail: `${made} page${made > 1 ? 's' : ''}`, warnings: notes };
  } finally {
    stage.remove();
    r.release();
  }
}
