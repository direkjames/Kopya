// Lays out flowing HTML, finds page breaks between lines, and draws pages into a PDF.
import { getHtmlToImage, getPDFLib } from './loader.js';
import { canvasToBlob, tick } from './util.js';

export function createStage(css) {
  const stage = document.createElement('div');
  stage.className = 'kp-stage';
  stage.setAttribute('aria-hidden', 'true');
  stage.style.cssText = 'position:absolute;left:-30000px;top:0;pointer-events:none;contain:layout style;';
  if (css) {
    const style = document.createElement('style');
    style.textContent = css;
    stage.appendChild(style);
  }
  document.body.appendChild(stage);
  return stage;
}

export async function waitForImages(root) {
  const imgs = Array.from(root.querySelectorAll('img'));
  await Promise.all(imgs.map((img) => (img.complete ? null : new Promise((r) => {
    img.onload = img.onerror = r;
    setTimeout(r, 8000);
  }))));
  for (const img of imgs) {
    if (!img.naturalWidth) img.remove();
    else if (img.decode) await img.decode().catch(() => {});
  }
}

const ATOMIC = 'img,svg,canvas,video,tr,hr,[data-kp-keep]';

// Returns [[start, end], ...] in px relative to the root's top.
export function paginate(root, pageHeight) {
  const top = root.getBoundingClientRect().top;
  const atoms = [];
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const range = document.createRange();
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (!n.nodeValue.trim()) continue;
    if (n.parentElement && n.parentElement.closest(ATOMIC)) continue;
    range.selectNodeContents(n);
    for (const r of range.getClientRects()) if (r.height > 0.5) atoms.push([r.top - top, r.bottom - top]);
  }
  for (const el of root.querySelectorAll(ATOMIC)) {
    if (el.parentElement && el.parentElement.closest(ATOMIC)) continue;
    const r = el.getBoundingClientRect();
    if (r.height > 0.5) atoms.push([r.top - top, r.bottom - top]);
  }
  const blocks = Array.from(root.children).map((c) => c.getBoundingClientRect());
  const total = Math.max(root.scrollHeight, ...blocks.map((b) => b.bottom - top), ...atoms.map((a) => a[1]), 1);

  // Spans no break may fall inside.
  const spans = atoms.map(([a, b]) => [a + 0.75, b - 0.75]).filter(([a, b]) => b > a).sort((x, y) => x[0] - y[0]);
  const merged = [];
  for (const s of spans) {
    const last = merged[merged.length - 1];
    if (last && s[0] <= last[1]) last[1] = Math.max(last[1], s[1]);
    else merged.push(s.slice());
  }
  const blocked = (y) => {
    let lo = 0, hi = merged.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (y < merged[mid][0]) hi = mid - 1;
      else if (y > merged[mid][1]) lo = mid + 1;
      else return true;
    }
    return false;
  };
  const cands = Array.from(new Set([...atoms.flatMap((a) => a), ...blocks.flatMap((b) => [b.top - top, b.bottom - top])]
    .map((v) => Math.round(v * 4) / 4))).sort((a, b) => a - b);
  const forced = Array.from(root.querySelectorAll('[data-kp-break]'))
    .map((el) => el.getBoundingClientRect().top - top).sort((a, b) => a - b);
  const starts = atoms.map((a) => a[0]).sort((a, b) => a - b);
  const hasContent = (s, e) => atoms.some(([a, b]) => b > s + 0.5 && a < e - 0.5);

  const slices = [];
  let s = 0;
  let guard = 0;
  while (s < total - 1 && guard++ < 5000) {
    const limit = s + pageHeight;
    let e = forced.find((f) => f > s + 1 && f <= limit);
    if (e === undefined) {
      if (limit >= total) e = total;
      else {
        e = limit;
        for (let i = cands.length - 1; i >= 0; i--) {
          const c = cands[i];
          if (c > limit) continue;
          if (c <= s + pageHeight * 0.2) break;
          if (!blocked(c)) { e = c; break; }
        }
      }
    }
    if (hasContent(s, e) || !slices.length) slices.push([s, e]);
    // Skip blank space at the top of the next page.
    const next = starts.find((a) => a >= e - 0.5);
    const isForced = forced.some((f) => Math.abs(f - e) < 1);
    s = next !== undefined && !isForced && next - e < pageHeight * 0.5 ? Math.max(e, next - 2) : e;
    if (next === undefined && !forced.some((f) => f > e) && atoms.every(([, b]) => b <= e + 0.5)) break;
  }
  return { slices, blocks: blocks.map((b) => [b.top - top, b.bottom - top]) };
}

// Collects pages into a PDF.
export async function pdfBuilder() {
  const { PDFDocument } = await getPDFLib();
  const doc = await PDFDocument.create();
  doc.setProducer('Kopya');
  doc.setCreator('Kopya');
  return {
    doc,
    async addCanvas(canvas, wPt, hPt, quality = 0.88) {
      const bytes = new Uint8Array(await (await canvasToBlob(canvas, 'image/jpeg', quality)).arrayBuffer());
      const img = await doc.embedJpg(bytes);
      const page = doc.addPage([wPt, hPt]);
      page.drawImage(img, { x: 0, y: 0, width: wPt, height: hPt });
      canvas.width = canvas.height = 0;
    },
    async save() {
      const bytes = await doc.save();
      return new Blob([bytes], { type: 'application/pdf' });
    },
  };
}

export async function snapshot(el, w, h, pixelRatio) {
  const h2i = await getHtmlToImage();
  return h2i.toCanvas(el, {
    width: w, height: h, pixelRatio, backgroundColor: '#ffffff',
    skipFonts: true, cacheBust: false, skipAutoScale: true,
  });
}

// Draws every slice of `root` as a page.
// page: { width, height, margin: {top,right,bottom,left} } in px.
export async function renderFlow({ root, stage, page, builder, pixelRatio, quality, decorate, onPage, firstPageNo = 1, totalPages }) {
  const contentW = page.width - page.margin.left - page.margin.right;
  const contentH = page.height - page.margin.top - page.margin.bottom;
  const { slices, blocks } = paginate(root, contentH);
  const kids = Array.from(root.children);
  for (let i = 0; i < slices.length; i++) {
    const [s, e] = slices[i];
    const frame = document.createElement('div');
    frame.style.cssText = `position:relative;width:${page.width}px;height:${page.height}px;background:#fff;overflow:hidden;`;
    const view = document.createElement('div');
    view.style.cssText = `position:absolute;left:${page.margin.left}px;top:${page.margin.top}px;width:${contentW}px;height:${Math.ceil(e - s)}px;overflow:hidden;`;
    const inner = document.createElement('div');
    inner.className = root.className;
    inner.style.cssText = root.style.cssText;
    inner.style.display = 'flow-root';
    inner.style.width = `${contentW}px`;
    inner.style.margin = '0';
    let first = -1;
    kids.forEach((k, j) => {
      const [bt, bb] = blocks[j];
      if (bb > s && bt < e) {
        if (first < 0) first = j;
        inner.appendChild(k.cloneNode(true));
      }
    });
    view.appendChild(inner);
    frame.appendChild(view);
    stage.appendChild(frame);
    if (first >= 0 && inner.firstElementChild) {
      const c = inner.firstElementChild.getBoundingClientRect().top - inner.getBoundingClientRect().top;
      inner.style.transform = `translateY(${-(s - blocks[first][0]) - c}px)`;
    }
    if (decorate) decorate(frame, firstPageNo + i, totalPages);
    await waitForImages(frame);
    const canvas = await snapshot(frame, page.width, page.height, pixelRatio);
    await builder.addCanvas(canvas, page.width * 0.75, page.height * 0.75, quality);
    frame.remove();
    if (onPage) onPage(i + 1, slices.length);
    await tick();
  }
  return slices.length;
}

export function countPages(root, page) {
  return paginate(root, page.height - page.margin.top - page.margin.bottom).slices.length;
}
