// Image reading, resizing and writing for the image converter.
// JPG, PNG, WebP, BMP, GIF and SVG go through the browser's own canvas.
// HEIC and AVIF writing load a small WebAssembly codec only when needed.
import { getLibheif, encodeAvif } from './loader.js';
import { canvasToBlob, tick } from './util.js';

export const MAX_FILE_BYTES = 250 * 1024 * 1024; // refuse bigger files instead of freezing the tab
export const MAX_SOURCE_PIXELS = 150e6;          // refuse absurd images (decompression bombs)
const MAX_SIDE = 16384;                          // most browsers can't make a canvas bigger than this
const PIXEL_BUDGET = { avif: 36e6, bmp: 60e6, default: 64e6 };

export const FORMATS = {
  jpg: { mime: 'image/jpeg', ext: '.jpg', label: 'JPG', opaque: true },
  png: { mime: 'image/png', ext: '.png', label: 'PNG', opaque: false },
  webp: { mime: 'image/webp', ext: '.webp', label: 'WebP', opaque: false },
  avif: { mime: 'image/avif', ext: '.avif', label: 'AVIF', opaque: false },
  bmp: { mime: 'image/bmp', ext: '.bmp', label: 'BMP', opaque: true },
};

const extOf = (name) => (name.match(/\.([^.]+)$/) || ['', ''])[1].toLowerCase();

function sniff(bytes, file) {
  const ext = extOf(file.name);
  const ascii = (from, to) => String.fromCharCode(...bytes.subarray(from, to));
  if (ascii(4, 8) === 'ftyp') {
    const brand = ascii(8, 12);
    if (/^(heic|heix|hevc|hevx|heim|heis|mif1|msf1)$/.test(brand)) return 'heic';
    if (brand === 'avif' || brand === 'avis') return 'bitmap';
  }
  if (ext === 'svg' || ext === 'svgz') return 'svg';
  if (ext === 'heic' || ext === 'heif') return 'heic';
  const head = new TextDecoder().decode(bytes.subarray(0, 512)).trimStart().toLowerCase();
  if (head.startsWith('<svg') || (head.startsWith('<?xml') && head.includes('<svg'))) return 'svg';
  return 'bitmap';
}

// ---------- reading ----------

async function decodeBitmap(file) {
  let bitmap;
  try {
    bitmap = await createImageBitmap(file, { imageOrientation: 'from-image' });
  } catch {
    try { bitmap = await createImageBitmap(file); } catch { bitmap = null; }
  }
  if (!bitmap) {
    if (extOf(file.name) === 'avif') throw new Error('This browser can’t open AVIF files. Update to a recent Chrome, Edge, Firefox or Safari 16 or newer.');
    throw new Error('This file couldn’t be read as an image. It may be damaged or in a format this browser doesn’t support.');
  }
  return { drawable: bitmap, width: bitmap.width, height: bitmap.height, close: () => bitmap.close() };
}

const px = (v) => {
  const m = /^\s*([\d.]+)\s*(px)?\s*$/i.exec(v || '');
  return m ? parseFloat(m[1]) : null;
};

async function decodeSvg(file) {
  if (file.size > 20 * 1048576) throw new Error('This SVG is larger than 20 MB, which is more than a browser can safely draw.');
  let text = new TextDecoder().decode(new Uint8Array(await file.arrayBuffer()));
  if (/^\x1f\x8b/.test(text)) throw new Error('Compressed SVG (.svgz) files aren’t supported. Unzip it to .svg first.');
  const doc = new DOMParser().parseFromString(text, 'image/svg+xml');
  const root = doc.documentElement;
  if (!root || root.localName !== 'svg' || doc.getElementsByTagName('parsererror').length) {
    throw new Error('This SVG couldn’t be read. It may be damaged.');
  }
  let w = px(root.getAttribute('width'));
  let h = px(root.getAttribute('height'));
  const vb = (root.getAttribute('viewBox') || '').trim().split(/[\s,]+/).map(Number);
  if ((!w || !h) && vb.length === 4 && vb[2] > 0 && vb[3] > 0) {
    if (w) h = w * vb[3] / vb[2];
    else if (h) w = h * vb[2] / vb[3];
    else { w = vb[2]; h = vb[3]; }
  }
  if (!w || !h) { w = w || 300; h = h || 150; }
  // Browsers need explicit pixel sizes to draw an SVG at a size we choose.
  root.setAttribute('width', String(w));
  root.setAttribute('height', String(h));
  text = new XMLSerializer().serializeToString(root);
  const url = URL.createObjectURL(new Blob([text], { type: 'image/svg+xml' }));
  try {
    const img = new Image();
    img.decoding = 'async';
    await new Promise((resolve, reject) => {
      img.onload = resolve;
      img.onerror = () => reject(new Error('This SVG couldn’t be drawn. It may use features the browser doesn’t support.'));
      img.src = url;
    });
    return { drawable: img, width: w, height: h, vector: true, close: () => {} };
  } finally {
    URL.revokeObjectURL(url);
  }
}

async function decodeHeic(file) {
  const libheif = await getLibheif();
  const data = new Uint8Array(await file.arrayBuffer());
  let images;
  try {
    images = new libheif.HeifDecoder().decode(data);
  } catch {
    throw new Error('This HEIC file couldn’t be read. It may be damaged, or use a codec Kopya can’t decode.');
  }
  if (!images || !images.length) throw new Error('This HEIC file has no picture in it that Kopya can read.');
  const image = images.find((i) => { try { return i.is_primary(); } catch { return false; } }) || images[0];
  const freeAll = () => { for (const i of images) { try { i.free(); } catch { /* already freed */ } } };
  try {
    const width = image.get_width();
    const height = image.get_height();
    if (!width || !height || width * height > MAX_SOURCE_PIXELS) throw new Error('This image is too large to convert in the browser.');
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext('2d');
    if (!ctx) throw new Error('This image is too large for your browser to draw.');
    const imageData = ctx.createImageData(width, height);
    const filled = await new Promise((resolve) => image.display(imageData, resolve));
    if (!filled) throw new Error('This HEIC file couldn’t be decoded. It may be damaged.');
    ctx.putImageData(imageData, 0, 0);
    return { drawable: canvas, width, height, close: () => { canvas.width = canvas.height = 0; } };
  } finally {
    freeAll();
  }
}

export async function readImage(file) {
  if (file.size > MAX_FILE_BYTES) throw new Error(`This file is larger than ${Math.round(MAX_FILE_BYTES / 1048576)} MB, which is more than a browser can safely convert.`);
  if (!file.size) throw new Error('This file is empty.');
  const head = new Uint8Array(await file.slice(0, 512).arrayBuffer());
  const kind = sniff(head, file);
  let img;
  if (kind === 'svg') img = await decodeSvg(file);
  else if (kind === 'heic') {
    try {
      img = await decodeHeic(file);
    } catch (heicError) {
      // Some files are named .heic but hold an ordinary picture (or Safari can open them itself).
      try { img = await decodeBitmap(file); } catch { throw heicError; }
    }
  } else img = await decodeBitmap(file);
  img.warnings = img.warnings || [];
  if (!img.vector && img.width * img.height > MAX_SOURCE_PIXELS) {
    img.close();
    throw new Error('This image is too large to convert in the browser.');
  }
  if (extOf(file.name) === 'gif' && new TextDecoder('latin1').decode(head).includes('NETSCAPE2.0')) {
    img.warnings.push('This GIF is animated. Only its first frame was converted.');
  }
  return img;
}

// ---------- resizing ----------

// Works out the output size. `maxSide` of 0 means "keep the original size".
export function targetSize(img, maxSide, format) {
  const warnings = [];
  let scale = 1;
  const longest = Math.max(img.width, img.height);
  if (maxSide) scale = img.vector ? maxSide / longest : Math.min(1, maxSide / longest);
  else if (img.vector && longest < 1024) scale = 1024 / longest; // vector art stays sharp when enlarged
  let w = img.width * scale;
  let h = img.height * scale;
  const budget = PIXEL_BUDGET[format] || PIXEL_BUDGET.default;
  const limit = Math.min(MAX_SIDE / Math.max(w, h), Math.sqrt(budget / (w * h)), 1.0000001);
  if (limit < 1) {
    w *= limit;
    h *= limit;
    warnings.push(`The image was reduced to ${Math.round(w)} × ${Math.round(h)} pixels because this format can’t safely handle a larger picture in a browser.`);
  }
  return { width: Math.max(1, Math.round(w)), height: Math.max(1, Math.round(h)), warnings };
}

function newCanvas(w, h) {
  const c = document.createElement('canvas');
  c.width = w;
  c.height = h;
  const ctx = c.getContext('2d', { willReadFrequently: false });
  if (!ctx) { c.width = c.height = 0; return null; }
  return { canvas: c, ctx };
}

// Draws the image at the target size on a new canvas. Big reductions are done
// in halving steps, which keeps fine detail from turning into jagged noise.
export function renderImage(img, width, height, background) {
  let target = newCanvas(width, height);
  let note = null;
  while (!target && width > 256) {
    width = Math.floor(width * 0.7);
    height = Math.floor(height * 0.7);
    target = newCanvas(width, height);
    note = `The image was reduced to ${width} × ${height} pixels because your browser couldn’t draw anything larger.`;
  }
  if (!target) throw new Error('This image is too large for your browser to draw.');
  const { canvas, ctx } = target;
  if (background) { ctx.fillStyle = background; ctx.fillRect(0, 0, width, height); }
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';

  let src = img.drawable;
  let sw = img.width;
  let sh = img.height;
  const temps = [];
  if (!img.vector) {
    while (sw / 2 >= width && sh / 2 >= height) {
      const nw = Math.max(width, Math.floor(sw / 2));
      const nh = Math.max(height, Math.floor(sh / 2));
      const step = newCanvas(nw, nh);
      if (!step) break;
      step.ctx.imageSmoothingQuality = 'high';
      step.ctx.drawImage(src, 0, 0, sw, sh, 0, 0, nw, nh);
      temps.push(step.canvas);
      src = step.canvas;
      sw = nw;
      sh = nh;
    }
  }
  ctx.drawImage(src, 0, 0, sw, sh, 0, 0, width, height);
  for (const t of temps) t.width = t.height = 0;
  return { canvas, width, height, note };
}

// ---------- writing ----------

// 24-bit BMP, written in strips so a big picture doesn't need two copies in memory.
async function encodeBmp(canvas) {
  const { width, height } = canvas;
  const stride = (width * 3 + 3) & ~3;
  const size = 54 + stride * height;
  if (size >= 0x7fffffff) throw new Error('This image is too large to save as BMP.');
  const out = new Uint8Array(size);
  const dv = new DataView(out.buffer);
  out[0] = 0x42; out[1] = 0x4d;
  dv.setUint32(2, size, true);
  dv.setUint32(10, 54, true);
  dv.setUint32(14, 40, true);
  dv.setInt32(18, width, true);
  dv.setInt32(22, height, true);
  dv.setUint16(26, 1, true);
  dv.setUint16(28, 24, true);
  dv.setUint32(34, stride * height, true);
  dv.setInt32(38, 2835, true);
  dv.setInt32(42, 2835, true);
  const ctx = canvas.getContext('2d');
  const strip = Math.max(1, Math.floor(4e6 / (width * 4)));
  for (let y0 = 0; y0 < height; y0 += strip) {
    const rows = Math.min(strip, height - y0);
    const data = ctx.getImageData(0, y0, width, rows).data;
    for (let r = 0; r < rows; r++) {
      let o = 54 + (height - 1 - (y0 + r)) * stride;
      let i = r * width * 4;
      for (let x = 0; x < width; x++, i += 4) {
        out[o++] = data[i + 2];
        out[o++] = data[i + 1];
        out[o++] = data[i];
      }
    }
    if (y0 && (y0 / strip) % 16 === 0) await tick();
  }
  return new Blob([out], { type: 'image/bmp' });
}

async function encodeCanvas(canvas, mime, quality) {
  const blob = await canvasToBlob(canvas, mime, quality);
  if (blob.type !== mime) {
    throw new Error(`This browser can’t save ${mime.split('/')[1].toUpperCase()} images. Try a recent Chrome, Edge or Firefox.`);
  }
  return blob;
}

// `level` is { q: 0-1 for JPG/WebP, avif: 0-100 for AVIF }.
export async function encodeImage(canvas, format, level, report) {
  const f = FORMATS[format];
  if (format === 'bmp') return encodeBmp(canvas);
  if (format === 'avif') {
    const { width, height } = canvas;
    report(0.55, 'Encoding AVIF (this one takes a while)');
    await tick();
    const rgba = canvas.getContext('2d').getImageData(0, 0, width, height).data;
    const pixels = width * height;
    const speed = pixels > 8e6 ? 9 : pixels > 2e6 ? 8 : 6;
    // The encoder can't report progress, so ease the bar forward while it works.
    let shown = 0.55;
    const timer = setInterval(() => { shown += (0.95 - shown) * 0.04; report(shown, 'Encoding AVIF (this one takes a while)'); }, 500);
    try {
      const bytes = await encodeAvif(rgba, width, height, {
        quality: level.avif, qualityAlpha: -1, denoiseLevel: 0, tileRowsLog2: 0, tileColsLog2: 0,
        speed, subsample: 1, chromaDeltaQ: false, sharpness: 0, tune: 0, enableSharpYUV: false, bitDepth: 8,
      });
      return new Blob([bytes], { type: f.mime });
    } finally {
      clearInterval(timer);
    }
  }
  return encodeCanvas(canvas, f.mime, level.q);
}
