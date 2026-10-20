// Lazy-loads the vendored libraries only when a converter needs them,
// so the first page load stays small.
const base = new URL('../../vendor/', import.meta.url).href;
const scripts = new Map();

function loadScript(file, globalName) {
  if (window[globalName]) return Promise.resolve(window[globalName]);
  if (!scripts.has(file)) {
    scripts.set(file, new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = base + file;
      s.onload = () => window[globalName]
        ? resolve(window[globalName])
        : reject(new Error(`${file} loaded but ${globalName} is missing`));
      s.onerror = () => {
        scripts.delete(file);
        reject(new Error(`Couldn't load ${file}. Check your connection and try again.`));
      };
      document.head.appendChild(s);
    }));
  }
  return scripts.get(file);
}

export const getJSZip = () => loadScript('jszip.min.js', 'JSZip');
export const getPDFLib = () => loadScript('pdf-lib.min.js', 'PDFLib');
export const getHtmlToImage = () => loadScript('html-to-image.js', 'htmlToImage');
export const getDocx = () => loadScript('docx.min.js', 'docx');
export const getLame = () => loadScript('lame.min.js', 'lamejs');

// HEIC/HEIF decoder (libheif). Loaded only when someone drops a HEIC photo.
let heifPromise;
export function getLibheif() {
  if (!heifPromise) {
    heifPromise = Promise.all([
      loadScript('heif/libheif.js', 'libheif'),
      fetch(`${base}heif/libheif.wasm`).then((r) => {
        if (!r.ok) throw new Error('Couldn’t load the HEIC decoder. Check your connection and try again.');
        return r.arrayBuffer();
      }),
    ]).then(([factory, wasmBinary]) => factory({ wasmBinary }))
      .catch((e) => { heifPromise = null; throw e; });
  }
  return heifPromise;
}

// AVIF encoder, run in its own worker. Resolves with the encoded bytes.
export function encodeAvif(rgba, width, height, options) {
  return new Promise((resolve, reject) => {
    let worker;
    try {
      worker = new Worker(base + 'avif/worker.js', { type: 'module' });
    } catch {
      reject(new Error('This browser can’t run the AVIF encoder. Try a recent Chrome, Edge, Firefox or Safari.'));
      return;
    }
    const done = (fn, value) => { worker.terminate(); fn(value); };
    worker.onmessage = (e) => {
      if (e.data.ok) done(resolve, e.data.bytes);
      else done(reject, new Error(`AVIF encoding failed: ${e.data.message}`));
    };
    worker.onerror = () => done(reject, new Error('The AVIF encoder couldn’t start. Check your connection and try again.'));
    worker.postMessage({ buffer: rgba.buffer, width, height, options }, [rgba.buffer]);
  });
}

let pdfjsPromise;
export function getPdfjs() {
  if (!pdfjsPromise) {
    pdfjsPromise = import(base + 'pdfjs/pdf.min.mjs').then((lib) => {
      lib.GlobalWorkerOptions.workerSrc = base + 'pdfjs/pdf.worker.min.mjs';
      return lib;
    }).catch((e) => { pdfjsPromise = null; throw e; });
  }
  return pdfjsPromise;
}

export async function openPdf(file) {
  const pdfjs = await getPdfjs();
  const data = new Uint8Array(await file.arrayBuffer());
  try {
    return await pdfjs.getDocument({
      data,
      cMapUrl: base + 'pdfjs/cmaps/',
      cMapPacked: true,
      standardFontDataUrl: base + 'pdfjs/standard_fonts/',
      wasmUrl: base + 'pdfjs/wasm/',
      iccUrl: base + 'pdfjs/iccs/',
      isEvalSupported: false,
    }).promise;
  } catch (e) {
    if (e && e.name === 'PasswordException') {
      throw new Error('This PDF is password-protected. Remove the password first, then try again.');
    }
    throw new Error("This file couldn't be read as a PDF. It may be damaged.");
  }
}

export async function closePdf(pdf) {
  try {
    if (typeof pdf.destroy === 'function') await pdf.destroy();
    else if (pdf.loadingTask) await pdf.loadingTask.destroy();
  } catch { /* already closed */ }
}
