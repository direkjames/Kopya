import { getJSZip } from '../lib/loader.js';
import { attr, baseName, isBrowserImage, mimeFromPath, parseXml, resolvePath, tick } from '../lib/util.js';
import { createStage, pdfBuilder, renderFlow, snapshot, waitForImages } from '../lib/paginate.js';

const MM = 96 / 25.4;
const SIZES = { a5: [148, 210, 15, 11], a4: [210, 297, 20, 12], letter: [215.9, 279.4, 20, 12] };

const BOOK_CSS = (fontPt, maxImg) => `
.kp-book{font-family:Georgia,"Times New Roman",Times,serif;font-size:${fontPt}pt;line-height:1.5;color:#111;text-align:left;overflow-wrap:break-word;hyphens:auto}
.kp-book p,.kp-book div{margin:0 0 .55em}
.kp-book h1,.kp-book h2,.kp-book h3,.kp-book h4,.kp-book h5,.kp-book h6{font-family:"Helvetica Neue",Arial,sans-serif;line-height:1.25;margin:.2em 0 .7em;break-after:avoid}
.kp-book h1{font-size:1.7em}.kp-book h2{font-size:1.4em}.kp-book h3{font-size:1.2em}.kp-book h4,.kp-book h5,.kp-book h6{font-size:1em}
.kp-book>h1:first-child,.kp-book>h2:first-child{margin-top:1.6em}
.kp-book img{display:block;max-width:100%;max-height:${maxImg}px;height:auto;margin:.4em auto;object-fit:contain}
.kp-book span img,.kp-book p img{display:inline-block;margin:0;vertical-align:middle}
.kp-book blockquote{margin:0 0 .8em 1.5em;font-style:italic}
.kp-book ul,.kp-book ol{margin:0 0 .8em;padding-left:1.6em}
.kp-book li{margin:0 0 .25em}
.kp-book pre{white-space:pre-wrap;font-size:.85em;font-family:Menlo,Consolas,monospace}
.kp-book table{border-collapse:collapse;margin:0 0 .8em;font-size:.9em;max-width:100%}
.kp-book td,.kp-book th{border:1px solid #aaa;padding:.2em .4em;vertical-align:top}
.kp-book hr{border:0;border-top:1px solid #bbb;margin:1em 20%}
.kp-book a{color:inherit;text-decoration:none}
.kp-book figcaption{font-size:.85em;text-align:center;color:#444}
`;

const DROP = new Set(['script', 'style', 'link', 'meta', 'title', 'head', 'iframe', 'object', 'embed', 'form', 'input', 'button', 'select', 'textarea', 'audio', 'video', 'noscript', 'template']);
const KEEP = new Set(['p', 'div', 'section', 'article', 'aside', 'header', 'footer', 'nav', 'main', 'center', 'span', 'em', 'i', 'strong', 'b', 'u', 's', 'sub', 'sup', 'small', 'big', 'br', 'hr', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'dl', 'dt', 'dd', 'blockquote', 'pre', 'code', 'table', 'thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'caption', 'figure', 'figcaption', 'img', 'cite', 'q', 'abbr', 'mark', 'del', 'ins', 'var', 'kbd', 'samp', 'tt']);
const CONTAINERS = new Set(['div', 'section', 'article', 'aside', 'header', 'footer', 'nav', 'main', 'center', 'figure', 'body']);
const BLOCKS = new Set(['p', 'div', 'section', 'article', 'aside', 'header', 'footer', 'nav', 'main', 'center', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'dl', 'blockquote', 'pre', 'table', 'figure', 'hr', 'figcaption']);
const STYLE_PROPS = ['text-align', 'font-style', 'font-weight', 'font-variant', 'text-indent'];

function classRules(cssTexts) {
  const rules = new Map();
  for (const css of cssTexts) {
    let sheet;
    try { sheet = new CSSStyleSheet(); sheet.replaceSync(css.replace(/@import[^;]+;/g, '')); } catch { continue; }
    for (const rule of sheet.cssRules) {
      if (!rule.selectorText || !rule.style) continue;
      const decl = {};
      for (const p of STYLE_PROPS) { const v = rule.style.getPropertyValue(p); if (v) decl[p] = v; }
      const br = rule.style.getPropertyValue('page-break-before') || rule.style.getPropertyValue('break-before');
      if (/always|page|left|right/.test(br)) decl.pageBreak = true;
      if (!Object.keys(decl).length) continue;
      for (const sel of rule.selectorText.split(',')) {
        const m = sel.trim().match(/^([a-z0-9]+)?\.([\w-]+)$/i);
        if (m) rules.set(`${(m[1] || '*').toLowerCase()}.${m[2]}`, { ...(rules.get(`${(m[1] || '*').toLowerCase()}.${m[2]}`) || {}), ...decl });
      }
    }
  }
  return rules;
}

class EpubBook {
  constructor(zip) { this.zip = zip; this.urls = new Map(); }

  async text(path) {
    const f = this.zip.file(path) || this.zip.file(decodeURIComponent(path));
    return f ? f.async('string') : null;
  }

  async imageUrl(path) {
    if (this.urls.has(path)) return this.urls.get(path);
    const f = this.zip.file(path);
    let url = null;
    if (f && isBrowserImage(path)) {
      const data = await f.async('uint8array');
      url = URL.createObjectURL(new Blob([data], { type: mimeFromPath(path) }));
    }
    this.urls.set(path, url);
    return url;
  }

  release() { for (const u of this.urls.values()) if (u) URL.revokeObjectURL(u); }

  async open() {
    const container = parseXml(await this.text('META-INF/container.xml') || '');
    const rootfile = container.getElementsByTagNameNS('*', 'rootfile')[0];
    this.opfPath = attr(rootfile, 'full-path');
    const opf = parseXml(await this.text(this.opfPath));
    const items = new Map();
    for (const it of opf.getElementsByTagNameNS('*', 'item')) {
      items.set(attr(it, 'id'), { href: resolvePath(this.opfPath, attr(it, 'href')), type: attr(it, 'media-type') || '', props: attr(it, 'properties') || '' });
    }
    this.spine = Array.from(opf.getElementsByTagNameNS('*', 'itemref'))
      .map((r) => items.get(attr(r, 'idref')))
      .filter((it) => it && /html|xml/.test(it.type) && !/\bnav\b/.test(it.props));
    const md = (n) => { const e = opf.getElementsByTagNameNS('*', n)[0]; return e ? e.textContent.trim() : ''; };
    this.title = md('title');
    this.author = md('creator');
    let cover = [...items.values()].find((it) => /cover-image/.test(it.props));
    if (!cover) {
      const meta = Array.from(opf.getElementsByTagNameNS('*', 'meta')).find((m) => attr(m, 'name') === 'cover');
      if (meta) cover = items.get(attr(meta, 'content'));
    }
    this.cover = cover && isBrowserImage(cover.href) ? cover.href : null;
    const cssFiles = [...items.values()].filter((it) => it.type === 'text/css');
    this.rules = classRules(await Promise.all(cssFiles.map((c) => this.text(c.href).then((t) => t || ''))));
  }

  async chapter(item) {
    const raw = await this.text(item.href);
    if (!raw) return null;
    let doc = new DOMParser().parseFromString(raw, 'application/xhtml+xml');
    if (doc.getElementsByTagName('parsererror').length) doc = new DOMParser().parseFromString(raw, 'text/html');
    const body = doc.body || doc.getElementsByTagNameNS('*', 'body')[0];
    if (!body) return null;
    const holder = document.createElement('div');
    for (const n of Array.from(body.childNodes)) {
      const c = await this.clean(n, item.href);
      if (c) holder.appendChild(c);
    }
    return holder;
  }

  styleFor(el, tag) {
    const out = {};
    for (const cls of (attr(el, 'class') || '').split(/\s+/).filter(Boolean)) {
      Object.assign(out, this.rules.get(`*.${cls}`) || {}, this.rules.get(`${tag}.${cls}`) || {});
    }
    const inline = attr(el, 'style');
    if (inline) {
      const probe = document.createElement('span');
      probe.style.cssText = inline;
      for (const p of STYLE_PROPS) { const v = probe.style.getPropertyValue(p); if (v) out[p] = v; }
      if (/page-break-before\s*:\s*always|break-before\s*:\s*page/.test(inline)) out.pageBreak = true;
    }
    const align = attr(el, 'align');
    if (align && !out['text-align']) out['text-align'] = align;
    return out;
  }

  async clean(node, from) {
    if (node.nodeType === 3) return document.createTextNode(node.nodeValue);
    if (node.nodeType !== 1) return null;
    const tag = node.localName.toLowerCase();
    if (DROP.has(tag)) return null;
    if (tag === 'svg') {
      // Cover pages often wrap a single image in SVG.
      const image = node.getElementsByTagNameNS('*', 'image')[0];
      if (!image) return null;
      const href = image.getAttributeNS('http://www.w3.org/1999/xlink', 'href') || attr(image, 'href');
      const url = href && await this.imageUrl(resolvePath(from, href));
      if (!url) return null;
      const img = document.createElement('img');
      img.src = url;
      return img;
    }
    if (tag === 'img' || tag === 'image') {
      const src = attr(node, 'src') || node.getAttributeNS('http://www.w3.org/1999/xlink', 'href') || attr(node, 'href');
      const url = src && !/^(https?:|data:)/.test(src) ? await this.imageUrl(resolvePath(from, src)) : (src && src.startsWith('data:') ? src : null);
      if (!url) {
        const alt = attr(node, 'alt');
        return alt ? document.createTextNode(`[${alt}]`) : null;
      }
      const img = document.createElement('img');
      img.src = url;
      img.alt = attr(node, 'alt') || '';
      return img;
    }
    const el = document.createElement(KEEP.has(tag) ? (tag === 'center' ? 'div' : tag) : 'span');
    if (!KEEP.has(tag) && !Array.from(node.childNodes).length) return null;
    for (const a of ['colspan', 'rowspan', 'start']) { const v = attr(node, a); if (v) el.setAttribute(a, v); }
    const st = this.styleFor(node, tag);
    if (tag === 'center') st['text-align'] = 'center';
    for (const p of STYLE_PROPS) if (st[p]) el.style.setProperty(p, st[p]);
    if (st.pageBreak) el.dataset.pb = '1';
    for (const n of Array.from(node.childNodes)) {
      const c = await this.clean(n, from);
      if (c) el.appendChild(c);
    }
    return el;
  }
}

// Unwraps nested containers so each top-level child is a short block.
function flatten(holder) {
  const out = [];
  let loose = null;
  const flushLoose = () => { if (loose && loose.textContent.trim() || loose && loose.querySelector('img')) out.push(loose); loose = null; };
  const walk = (parent, inherited) => {
    for (const n of Array.from(parent.childNodes)) {
      const tag = n.nodeType === 1 ? n.localName : '#text';
      const isBlock = n.nodeType === 1 && (BLOCKS.has(tag) || tag === 'img' && parent === holder);
      if (!isBlock) {
        if (n.nodeType === 3 && !n.nodeValue.trim() && !loose) continue;
        if (!loose) { loose = document.createElement('p'); if (inherited) loose.style.textAlign = inherited; }
        loose.appendChild(n);
        continue;
      }
      flushLoose();
      const hasBlockKids = CONTAINERS.has(tag) && Array.from(n.children).some((c) => BLOCKS.has(c.localName) || c.localName === 'img');
      if (hasBlockKids) {
        if (n.dataset.pb) { const mark = document.createElement('div'); mark.dataset.kpBreak = '1'; out.push(mark); }
        walk(n, n.style.textAlign || inherited);
        flushLoose();
      } else {
        if (inherited && !n.style.textAlign && n.style) n.style.textAlign = inherited;
        if (n.dataset && n.dataset.pb) n.dataset.kpBreak = '1';
        out.push(n);
      }
    }
  };
  walk(holder, '');
  flushLoose();
  return out;
}

export async function convert(file, { report, options }) {
  const [wMm, hMm, mMm, fontPt] = SIZES[options.pageSize] || SIZES.a5;
  const page = {
    width: Math.round(wMm * MM), height: Math.round(hMm * MM),
    margin: { top: Math.round(mMm * MM), bottom: Math.round((mMm + 4) * MM), left: Math.round(mMm * MM), right: Math.round(mMm * MM) },
  };
  const pixelRatio = options.quality === 'small' ? 1.5 : 2.2;

  const JSZip = await getJSZip();
  let zip;
  try { zip = await JSZip.loadAsync(file); } catch { throw new Error("This file couldn't be opened as an EPUB. It may be damaged or DRM-protected."); }
  const book = new EpubBook(zip);
  try { await book.open(); } catch { throw new Error("This EPUB's table of contents couldn't be read. The file may be damaged."); }
  if (zip.file('META-INF/encryption.xml') && /EncryptedData/.test(await book.text('META-INF/encryption.xml') || '')
    && !book.spine.length) throw new Error('This e-book is DRM-protected, so it can’t be converted.');

  const contentH = page.height - page.margin.top - page.margin.bottom;
  const stage = createStage(BOOK_CSS(fontPt, contentH - 20));
  const builder = await pdfBuilder();
  let pageNo = 0;
  try {
    if (book.cover) {
      const url = await book.imageUrl(book.cover);
      if (url) {
        const frame = document.createElement('div');
        frame.style.cssText = `width:${page.width}px;height:${page.height}px;background:#fff;display:flex;align-items:center;justify-content:center;`;
        frame.innerHTML = `<img src="${url}" style="max-width:100%;max-height:100%;object-fit:contain">`;
        stage.appendChild(frame);
        await waitForImages(frame);
        await builder.addCanvas(await snapshot(frame, page.width, page.height, pixelRatio), page.width * 0.75, page.height * 0.75, 0.9);
        frame.remove();
        pageNo++;
      }
    }

    const decorate = (frame, no) => {
      const f = document.createElement('div');
      f.textContent = String(no);
      f.style.cssText = `position:absolute;left:0;right:0;bottom:${Math.round(page.margin.bottom * 0.45)}px;text-align:center;font:9pt Georgia,serif;color:#777;`;
      frame.appendChild(f);
    };

    const total = book.spine.length;
    for (let i = 0; i < total; i++) {
      report(i / total * 0.97, `Laying out chapter ${i + 1} of ${total}`);
      const holder = await book.chapter(book.spine[i]);
      if (!holder) continue;
      const blocks = flatten(holder);
      const text = blocks.map((b) => b.textContent).join('').trim();
      const imgs = blocks.reduce((n, b) => n + (b.localName === 'img' ? 1 : b.querySelectorAll('img').length), 0);
      if (!text && !imgs) continue;
      // Skip a cover page that only repeats the cover image.
      if (i === 0 && book.cover && !text && imgs === 1) continue;
      const root = document.createElement('div');
      root.className = 'kp-book';
      root.style.width = `${page.width - page.margin.left - page.margin.right}px`;
      root.style.display = 'flow-root';
      for (const b of blocks) root.appendChild(b);
      stage.appendChild(root);
      await waitForImages(root);
      const made = await renderFlow({
        root, stage, page, builder, pixelRatio, quality: 0.84, decorate, firstPageNo: pageNo + 1,
        onPage: (k) => report(Math.min(0.97, (i + 0.5) / total * 0.97), `Drawing page ${pageNo + k}`),
      });
      pageNo += made;
      root.remove();
      await tick();
    }
    if (!pageNo) throw new Error('No readable pages were found in this EPUB.');
    report(0.98, 'Saving the PDF');
    builder.doc.setTitle(book.title || baseName(file.name));
    if (book.author) builder.doc.setAuthor(book.author);
    const blob = await builder.save();
    return { blob, name: `${baseName(file.name)}.pdf`, detail: `${pageNo} pages, ${options.pageSize === 'a4' ? 'A4' : options.pageSize === 'letter' ? 'Letter' : 'A5 book'} size` };
  } finally {
    stage.remove();
    book.release();
  }
}
