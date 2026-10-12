import { getJSZip } from '../lib/loader.js';
import { baseName, tick } from '../lib/util.js';
import { PptxRenderer } from '../lib/pptx-render.js';
import { createStage, pdfBuilder, snapshot, waitForImages } from '../lib/paginate.js';

export async function convert(file, { report, options }) {
  const head = new Uint8Array(await file.slice(0, 4).arrayBuffer());
  if (head[0] === 0xd0 && head[1] === 0xcf) {
    throw new Error('This is the older .ppt format. Open it in PowerPoint or Google Slides and save it as .pptx, then try again.');
  }
  const JSZip = await getJSZip();
  let zip;
  try { zip = await JSZip.loadAsync(file); } catch { throw new Error("This file couldn't be opened as a PowerPoint file. It may be damaged."); }
  const r = new PptxRenderer(zip);
  try { await r.load(); } catch { throw new Error("This file doesn't look like a PowerPoint presentation (.pptx)."); }
  if (!r.slides.length) throw new Error('This presentation has no slides.');

  const stage = createStage('');
  const builder = await pdfBuilder();
  const target = options.quality === 'small' ? 1280 : 1920;
  const pixelRatio = Math.min(3, Math.max(1, target / r.width));
  const wPt = r.cx / 12700;
  const hPt = r.cy / 12700;
  let made = 0, skipped = 0;
  try {
    for (let i = 0; i < r.slides.length; i++) {
      report(i / r.slides.length * 0.96, `Drawing slide ${i + 1} of ${r.slides.length}`);
      const { el, hidden } = await r.render(i);
      if (hidden && options.hidden !== 'include') { skipped++; continue; }
      stage.appendChild(el);
      await waitForImages(el);
      const canvas = await snapshot(el, r.width, r.height, pixelRatio);
      await builder.addCanvas(canvas, wPt, hPt, 0.9);
      el.remove();
      made++;
      await tick();
    }
    if (!made) throw new Error('Every slide in this presentation is hidden, so there was nothing to convert.');
    report(0.98, 'Saving the PDF');
    const blob = await builder.save();
    let detail = `${made} slide${made > 1 ? 's' : ''}`;
    if (skipped) detail += `; ${skipped} hidden slide${skipped > 1 ? 's' : ''} left out`;
    return { blob, name: `${baseName(file.name)}.pdf`, detail, warnings: [...r.warnings] };
  } finally {
    stage.remove();
    r.release();
  }
}
