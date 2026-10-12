export const baseName = (name) => name.replace(/\.[^./\\]+$/, '') || 'document';

export const escapeXml = (s) => String(s)
  .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;').replace(/'/g, '&apos;');

export const escapeHtml = escapeXml;

// Lets the browser repaint (progress bar) during long loops.
export const tick = () => new Promise((r) => setTimeout(r, 0));

export function canvasToBlob(canvas, type = 'image/jpeg', quality = 0.9) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((b) => (b ? resolve(b) : reject(new Error('The page was too large to draw. Try a lower quality setting.'))), type, quality);
  });
}

// Draws one pdf.js page onto a white canvas.
export async function renderPdfPage(page, scale) {
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.floor(viewport.width));
  canvas.height = Math.max(1, Math.floor(viewport.height));
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: ctx, viewport, canvas }).promise;
  return canvas;
}

// Keeps a scale from producing canvases browsers refuse to allocate.
export function safeScale(page, wanted, maxPixels = 24e6, maxSide = 14000) {
  const vp = page.getViewport({ scale: 1 });
  let s = wanted;
  s = Math.min(s, maxSide / vp.width, maxSide / vp.height);
  s = Math.min(s, Math.sqrt(maxPixels / (vp.width * vp.height)));
  return s;
}

// Resolves a relative path inside a zip (EPUB, DOCX, PPTX).
export function resolvePath(fromFile, rel) {
  if (!rel) return rel;
  rel = rel.split('#')[0];
  try { rel = decodeURIComponent(rel); } catch { /* keep as is */ }
  if (rel.startsWith('/')) return rel.slice(1);
  const parts = fromFile.split('/');
  parts.pop();
  for (const seg of rel.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.' && seg !== '') parts.push(seg);
  }
  return parts.join('/');
}

export function parseXml(text) {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('bad xml');
  return doc;
}

// Namespace-agnostic child helpers for OOXML.
export const kids = (el, name) => (el ? Array.from(el.children).filter((c) => c.localName === name) : []);
export const kid = (el, name) => (el ? Array.from(el.children).find((c) => c.localName === name) || null : null);
export const desc = (el, name) => (el ? Array.from(el.getElementsByTagNameNS('*', name)) : []);
export function attr(el, name) {
  if (!el) return null;
  for (const a of el.attributes) if (a.localName === name) return a.value;
  return null;
}

export const mimeFromPath = (p) => {
  const ext = p.split('.').pop().toLowerCase();
  return ({
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
    svg: 'image/svg+xml', webp: 'image/webp', bmp: 'image/bmp', avif: 'image/avif',
  })[ext] || 'application/octet-stream';
};

export const isBrowserImage = (p) => /\.(png|jpe?g|gif|svg|webp|bmp|avif)$/i.test(p);
