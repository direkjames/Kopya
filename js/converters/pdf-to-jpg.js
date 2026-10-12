import { openPdf, closePdf, getJSZip } from '../lib/loader.js';
import { baseName, canvasToBlob, renderPdfPage, safeScale } from '../lib/util.js';

export async function convert(file, { report, options }) {
  const dpi = Number(options.dpi) || 150;
  const pdf = await openPdf(file);
  const name = baseName(file.name);
  const n = pdf.numPages;
  const pad = String(n).length;
  const images = [];
  for (let i = 1; i <= n; i++) {
    report((i - 1) / n, `Drawing page ${i} of ${n}`);
    const page = await pdf.getPage(i);
    const canvas = await renderPdfPage(page, safeScale(page, dpi / 72));
    images.push({ name: `${name}-page-${String(i).padStart(pad, '0')}.jpg`, blob: await canvasToBlob(canvas, 'image/jpeg', 0.92) });
    canvas.width = canvas.height = 0;
    page.cleanup();
  }
  await closePdf(pdf);
  if (images.length === 1) return { blob: images[0].blob, name: `${name}.jpg` };
  report(0.98, 'Packing images into a ZIP');
  const JSZip = await getJSZip();
  const zip = new JSZip();
  for (const img of images) zip.file(img.name, img.blob);
  const blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
  return { blob, name: `${name}-jpg.zip`, detail: `${n} pages, one JPG each` };
}
