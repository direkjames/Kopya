// Renders a .docx into HTML blocks that the paginator can lay out as pages.
import { attr, isBrowserImage, kid, kids, mimeFromPath, parseXml, resolvePath } from './util.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const tw2pt = (v) => Number(v || 0) / 20;
const tw2px = (v) => Number(v || 0) / 15;
const emu2px = (v) => Number(v || 0) / 9525;

const HIGHLIGHT = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000',
  darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
};
const SERIF = /times|georgia|cambria|garamond|book|serif|palatino|minion|baskerville|caladea|century/i;
const MONO = /courier|consolas|mono|menlo/i;
const FALLBACK = { calibri: 'Carlito', cambria: 'Caladea', arial: '"Liberation Sans"', 'times new roman': '"Liberation Serif"', 'courier new': '"Liberation Mono"' };
const PUA_BULLETS = { '': '•', '': '▪', '': '➢', '': '✓', '': '❖', '': '◆', '': '■', '': '➔', o: '◦', '§': '▪', '': '•' };

export function fontStack(name) {
  if (!name) return null;
  const fb = FALLBACK[name.toLowerCase()];
  const generic = MONO.test(name) ? 'monospace' : SERIF.test(name) ? 'serif' : 'sans-serif';
  return `"${name}",${fb ? `${fb},` : ''}${generic}`;
}
const lineBase = (font) => (/calibri|carlito/i.test(font || '') ? 1.22 : /cambria/i.test(font || '') ? 1.17 : 1.15);

function onOff(el) {
  if (!el) return undefined;
  const v = el.getAttributeNS(W_NS, 'val');
  return !(v === '0' || v === 'false' || v === 'off' || v === 'none');
}
const wv = (el, name = 'val') => (el ? el.getAttributeNS(W_NS, name) : null);

function parseBorder(el) {
  if (!el) return undefined;
  const val = wv(el);
  if (!val || val === 'none' || val === 'nil') return 'none';
  const sz = Math.max(0.5, Number(wv(el, 'sz') || 4) / 8);
  const color = wv(el, 'color');
  const c = !color || color === 'auto' ? '#000' : `#${color}`;
  return `${val === 'double' ? Math.max(sz * 3, 2.25) : sz}pt ${val === 'double' ? 'double' : val === 'dotted' ? 'dotted' : /dash/.test(val) ? 'dashed' : 'solid'} ${c}`;
}

function parseRPr(el, theme) {
  const r = {};
  if (!el) return r;
  for (const k of ['b', 'i', 'strike', 'dstrike', 'caps', 'smallCaps', 'vanish']) {
    const v = onOff(kid(el, k));
    if (v !== undefined) r[k] = v;
  }
  const u = kid(el, 'u');
  if (u) r.u = wv(u) !== 'none';
  const sz = kid(el, 'sz');
  if (sz) r.sz = Number(wv(sz)) / 2;
  const color = kid(el, 'color');
  if (color && wv(color) && wv(color) !== 'auto') r.color = `#${wv(color)}`;
  const hl = kid(el, 'highlight');
  if (hl && HIGHLIGHT[wv(hl)]) r.highlight = HIGHLIGHT[wv(hl)];
  const shd = kid(el, 'shd');
  if (shd && wv(shd, 'fill') && !/auto/.test(wv(shd, 'fill'))) r.highlight = `#${wv(shd, 'fill')}`;
  const va = kid(el, 'vertAlign');
  if (va) r.vertAlign = wv(va);
  const f = kid(el, 'rFonts');
  if (f) {
    const name = wv(f, 'ascii') || wv(f, 'hAnsi');
    const th = wv(f, 'asciiTheme') || wv(f, 'hAnsiTheme');
    if (name) r.font = name;
    else if (th) r.font = /major/.test(th) ? theme.major : theme.minor;
  }
  const rs = kid(el, 'rStyle');
  if (rs) r.rStyle = wv(rs);
  return r;
}

function parsePPr(el) {
  const p = {};
  if (!el) return p;
  const jc = kid(el, 'jc');
  if (jc) p.jc = wv(jc);
  const sp = kid(el, 'spacing');
  if (sp) {
    p.spacing = {};
    for (const k of ['before', 'after', 'line', 'lineRule', 'beforeAutospacing', 'afterAutospacing']) {
      const v = wv(sp, k);
      if (v !== null && v !== '') p.spacing[k] = v;
    }
  }
  const ind = kid(el, 'ind');
  if (ind) {
    p.ind = {};
    for (const [k, alt] of [['left', 'start'], ['right', 'end'], ['firstLine'], ['hanging']]) {
      const v = wv(ind, k) ?? (alt ? wv(ind, alt) : null);
      if (v !== null && v !== '') p.ind[k] = Number(v);
    }
  }
  const num = kid(el, 'numPr');
  if (num) {
    const id = wv(kid(num, 'numId'));
    p.numPr = { numId: id, ilvl: Number(wv(kid(num, 'ilvl')) || 0) };
  }
  const pbb = onOff(kid(el, 'pageBreakBefore'));
  if (pbb !== undefined) p.pageBreakBefore = pbb;
  const cs = onOff(kid(el, 'contextualSpacing'));
  if (cs !== undefined) p.contextualSpacing = cs;
  const shd = kid(el, 'shd');
  if (shd && wv(shd, 'fill') && !/auto/.test(wv(shd, 'fill'))) p.shd = `#${wv(shd, 'fill')}`;
  const bdr = kid(el, 'pBdr');
  if (bdr) p.bdr = { top: parseBorder(kid(bdr, 'top')), bottom: parseBorder(kid(bdr, 'bottom')), left: parseBorder(kid(bdr, 'left')), right: parseBorder(kid(bdr, 'right')) };
  const ps = kid(el, 'pStyle');
  if (ps) p.pStyle = wv(ps);
  return p;
}

function mergeP(a, b) {
  const out = { ...a, ...b };
  if (a.spacing || b.spacing) out.spacing = { ...(a.spacing || {}), ...(b.spacing || {}) };
  if (a.ind || b.ind) {
    out.ind = { ...(a.ind || {}), ...(b.ind || {}) };
    // firstLine and hanging exclude each other.
    if (b.ind && b.ind.firstLine !== undefined) delete out.ind.hanging;
    if (b.ind && b.ind.hanging !== undefined) delete out.ind.firstLine;
  }
  return out;
}

function toRoman(n) {
  const map = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let s = '';
  for (const [v, r] of map) while (n >= v) { s += r; n -= v; }
  return s;
}
function toLetter(n) { let s = ''; while (n > 0) { s = String.fromCharCode(97 + ((n - 1) % 26)) + s; n = Math.floor((n - 1) / 26); } return s; }
function fmtNum(n, fmt) {
  switch (fmt) {
    case 'lowerLetter': return toLetter(n);
    case 'upperLetter': return toLetter(n).toUpperCase();
    case 'lowerRoman': return toRoman(n);
    case 'upperRoman': return toRoman(n).toUpperCase();
    case 'decimalZero': return String(n).padStart(2, '0');
    case 'none': return '';
    default: return String(n);
  }
}

export class DocxRenderer {
  constructor(zip) { this.zip = zip; this.urls = []; this.footnoteRefs = []; this.warnings = new Set(); }

  async xml(path) {
    const f = this.zip.file(path);
    return f ? parseXml(await f.async('string')) : null;
  }

  async rels(path) {
    const p = path.replace(/([^/]+)$/, '_rels/$1.rels');
    const doc = await this.xml(p);
    const map = new Map();
    if (doc) for (const r of doc.getElementsByTagNameNS('*', 'Relationship')) {
      map.set(attr(r, 'Id'), { target: attr(r, 'TargetMode') === 'External' ? null : resolvePath(path, attr(r, 'Target')), type: attr(r, 'Type') });
    }
    return map;
  }

  async load() {
    const root = await this.rels('');
    const main = [...root.values()].find((r) => /officeDocument$/.test(r.type));
    this.docPath = main ? main.target : 'word/document.xml';
    this.doc = await this.xml(this.docPath);
    if (!this.doc) throw new Error('missing document');
    this.docRels = await this.rels(this.docPath);
    const byType = (t) => [...this.docRels.values()].find((r) => r.type.endsWith(`/${t}`));

    this.theme = { minor: 'Calibri', major: 'Calibri Light' };
    const themeRel = byType('theme');
    const theme = themeRel && await this.xml(themeRel.target);
    if (theme) {
      const fs = theme.getElementsByTagNameNS('*', 'fontScheme')[0];
      const latin = (n) => { const e = fs && kid(kid(fs, n), 'latin'); return e && attr(e, 'typeface'); };
      this.theme = { minor: latin('minorFont') || 'Calibri', major: latin('majorFont') || 'Calibri Light' };
    }

    this.styles = new Map();
    this.defP = {}; this.defR = { sz: 10, font: 'Times New Roman' };
    const stRel = byType('styles');
    const styles = await this.xml(stRel ? stRel.target : 'word/styles.xml');
    if (styles) {
      const dd = styles.getElementsByTagNameNS('*', 'docDefaults')[0];
      if (dd) {
        this.defR = { ...this.defR, ...parseRPr(kid(kid(dd, 'rPrDefault'), 'rPr'), this.theme) };
        this.defP = parsePPr(kid(kid(dd, 'pPrDefault'), 'pPr'));
      }
      for (const s of styles.getElementsByTagNameNS('*', 'style')) {
        const id = wv(s, 'styleId');
        const tblPr = kid(s, 'tblPr');
        const cond = {};
        for (const c of kids(s, 'tblStylePr')) cond[wv(c, 'type')] = { rPr: parseRPr(kid(c, 'rPr'), this.theme), shd: (() => { const sh = kid(kid(c, 'tcPr'), 'shd'); return sh && wv(sh, 'fill') && wv(sh, 'fill') !== 'auto' ? `#${wv(sh, 'fill')}` : null; })() };
        this.styles.set(id, {
          type: wv(s, 'type'), basedOn: wv(kid(s, 'basedOn')), isDefault: wv(s, 'default') === '1',
          pPr: parsePPr(kid(s, 'pPr')), rPr: parseRPr(kid(s, 'rPr'), this.theme),
          tblBorders: tblPr && kid(tblPr, 'tblBorders'), cond,
        });
        if (wv(s, 'type') === 'paragraph' && wv(s, 'default') === '1') this.defaultPStyle = id;
      }
    }

    this.nums = new Map();
    const numRel = byType('numbering');
    const numbering = numRel && await this.xml(numRel.target);
    if (numbering) {
      const abstract = new Map();
      for (const a of numbering.getElementsByTagNameNS('*', 'abstractNum')) {
        const lvls = {};
        for (const l of kids(a, 'lvl')) {
          lvls[Number(wv(l, 'ilvl'))] = {
            start: Number(wv(kid(l, 'start')) || 1), fmt: wv(kid(l, 'numFmt')) || 'decimal',
            text: wv(kid(l, 'lvlText')) ?? '', pPr: parsePPr(kid(l, 'pPr')), rPr: parseRPr(kid(l, 'rPr'), this.theme),
            suff: wv(kid(l, 'suff')) || 'tab',
          };
        }
        abstract.set(wv(a, 'abstractNumId'), lvls);
      }
      for (const n of numbering.getElementsByTagNameNS('*', 'num')) {
        const lvls = abstract.get(wv(kid(n, 'abstractNumId'))) || {};
        const copy = {};
        for (const k of Object.keys(lvls)) copy[k] = { ...lvls[k] };
        for (const o of kids(n, 'lvlOverride')) {
          const so = kid(o, 'startOverride');
          if (so && copy[wv(o, 'ilvl')]) copy[wv(o, 'ilvl')].start = Number(wv(so));
        }
        this.nums.set(wv(n, 'numId'), { lvls: copy, counters: [] });
      }
    }

    const fnRel = byType('footnotes');
    this.footnotes = new Map();
    const fn = fnRel && await this.xml(fnRel.target);
    if (fn) {
      this.fnRels = await this.rels(fnRel.target);
      for (const f of fn.getElementsByTagNameNS('*', 'footnote')) {
        const t = wv(f, 'type');
        if (!t || t === 'normal') this.footnotes.set(wv(f, 'id'), f);
      }
    }

    const body = this.doc.getElementsByTagNameNS('*', 'body')[0];
    this.body = body;
    const sect = kid(body, 'sectPr') || [...body.getElementsByTagNameNS('*', 'sectPr')].pop();
    const pgSz = kid(sect, 'pgSz');
    const pgMar = kid(sect, 'pgMar');
    const w = Number(wv(pgSz, 'w')) || 11906;
    const h = Number(wv(pgSz, 'h')) || 16838;
    const m = (k, d) => Math.max(0, Math.abs(Number(wv(pgMar, k) ?? d)));
    this.page = {
      width: Math.round(tw2px(w)), height: Math.round(tw2px(h)),
      margin: { top: tw2px(m('top', 1440)), bottom: tw2px(m('bottom', 1440)), left: tw2px(m('left', 1440)), right: tw2px(m('right', 1440)) },
      headerAt: tw2px(m('header', 720)), footerAt: tw2px(m('footer', 720)),
    };
    const cols = kid(sect, 'cols');
    if (cols && Number(wv(cols, 'num')) > 1) this.warnings.add('Text in columns was laid out as a single column.');
    this.titlePg = onOff(kid(sect, 'titlePg')) || false;
    this.hf = {};
    for (const kind of ['header', 'footer']) {
      for (const ref of kids(sect, `${kind}Reference`)) {
        const rel = this.docRels.get(ref.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'));
        if (rel && rel.target) this.hf[`${kind}-${wv(ref, 'type') || 'default'}`] = rel.target;
      }
    }
  }

  styleChain(id) {
    const chain = [];
    let s = id;
    const seen = new Set();
    while (s && this.styles.has(s) && !seen.has(s)) { seen.add(s); chain.unshift(this.styles.get(s)); s = this.styles.get(s).basedOn; }
    return chain;
  }

  effP(direct, ctx) {
    let p = { ...this.defP };
    if (ctx.tblStyle) for (const s of this.styleChain(ctx.tblStyle)) p = mergeP(p, s.pPr);
    for (const s of this.styleChain(direct.pStyle || this.defaultPStyle)) p = mergeP(p, s.pPr);
    return mergeP(p, direct);
  }

  effR(pStyle, direct, ctx, extra) {
    let r = { ...this.defR };
    if (ctx.tblStyle) for (const s of this.styleChain(ctx.tblStyle)) Object.assign(r, s.rPr);
    for (const s of this.styleChain(pStyle || this.defaultPStyle)) Object.assign(r, s.rPr);
    if (extra) Object.assign(r, extra);
    if (direct.rStyle) for (const s of this.styleChain(direct.rStyle)) Object.assign(r, s.rPr);
    Object.assign(r, direct);
    return r;
  }

  async imageUrl(rid, rels) {
    const rel = (rels || this.docRels).get(rid);
    if (!rel || !rel.target) return null;
    if (!isBrowserImage(rel.target)) { this.warnings.add('Some pictures use formats browsers can’t draw (such as EMF/WMF) and were left out.'); return null; }
    const f = this.zip.file(rel.target);
    if (!f) return null;
    const url = URL.createObjectURL(new Blob([await f.async('uint8array')], { type: mimeFromPath(rel.target) }));
    this.urls.push(url);
    return url;
  }

  release() { for (const u of this.urls) URL.revokeObjectURL(u); }

  runStyle(span, r) {
    const s = span.style;
    if (r.font) s.fontFamily = fontStack(r.font);
    if (r.sz) s.fontSize = `${r.sz}pt`;
    if (r.b) s.fontWeight = '700';
    if (r.i) s.fontStyle = 'italic';
    const deco = [r.u && 'underline', (r.strike || r.dstrike) && 'line-through'].filter(Boolean);
    if (deco.length) s.textDecoration = deco.join(' ');
    if (r.color) s.color = r.color;
    if (r.highlight) s.backgroundColor = r.highlight;
    if (r.caps) s.textTransform = 'uppercase';
    if (r.smallCaps) s.fontVariant = 'small-caps';
    if (r.vertAlign === 'superscript' || r.vertAlign === 'subscript') {
      s.verticalAlign = r.vertAlign === 'superscript' ? 'super' : 'sub';
      s.fontSize = `${(r.sz || 11) * 0.65}pt`;
      s.lineHeight = '0';
    }
  }

  // Block-level children (body, cell, header, text box) → array of DOM blocks.
  async blocks(container, ctx = {}) {
    const out = [];
    for (const el of Array.from(container.children)) {
      const n = el.localName;
      if (n === 'p') out.push(...await this.paragraph(el, ctx));
      else if (n === 'tbl') out.push(await this.table(el, ctx));
      else if (n === 'sdt') out.push(...await this.blocks(kid(el, 'sdtContent') || el, ctx));
      else if (n === 'customXml' || n === 'ins' || n === 'moveTo') out.push(...await this.blocks(el, ctx));
      else if (n === 'AlternateContent') out.push(...await this.blocks(kid(el, 'Choice') || el, ctx));
    }
    return out;
  }

  levelDef(pPr) {
    const np = pPr.numPr;
    if (!np || !np.numId || np.numId === '0') return null;
    const num = this.nums.get(np.numId);
    return (num && num.lvls[np.ilvl]) || null;
  }

  marker(pPr) {
    const np = pPr.numPr;
    if (!np || !np.numId || np.numId === '0') return null;
    const num = this.nums.get(np.numId);
    if (!num) return null;
    const lvl = num.lvls[np.ilvl];
    if (!lvl) return null;
    const c = num.counters;
    for (let i = 0; i < np.ilvl; i++) if (c[i] === undefined && num.lvls[i]) c[i] = num.lvls[i].start;
    c[np.ilvl] = c[np.ilvl] === undefined ? lvl.start : c[np.ilvl] + 1;
    c.length = np.ilvl + 1;
    let text;
    if (lvl.fmt === 'bullet') text = Array.from(lvl.text).map((ch) => PUA_BULLETS[ch] || (ch.charCodeAt(0) >= 0xf000 ? '•' : ch)).join('') || '•';
    else text = lvl.text.replace(/%(\d)/g, (_, d) => { const i = Number(d) - 1; const l = num.lvls[i]; return l ? fmtNum(c[i] ?? l.start, l.fmt) : ''; });
    return { text, lvl };
  }

  paraStyle(el, p, ctx, prevStyle) {
    const s = el.style;
    const sp = p.spacing || {};
    let before = sp.beforeAutospacing === '1' ? 14 : tw2pt(sp.before);
    let after = sp.afterAutospacing === '1' ? 14 : tw2pt(sp.after);
    if (ctx.noBefore) before = 0;
    s.paddingTop = `${before}pt`;
    s.paddingBottom = `${after}pt`;
    const ind = p.ind || {};
    const left = tw2pt(ind.left);
    const right = tw2pt(ind.right);
    if (left) s.paddingLeft = `${left}pt`;
    if (right) s.paddingRight = `${right}pt`;
    if (ind.hanging) s.textIndent = `${-tw2pt(ind.hanging)}pt`;
    else if (ind.firstLine) s.textIndent = `${tw2pt(ind.firstLine)}pt`;
    const jc = p.jc;
    if (jc === 'center') s.textAlign = 'center';
    else if (jc === 'right' || jc === 'end') s.textAlign = 'right';
    else if (jc === 'both' || jc === 'distribute') s.textAlign = 'justify';
    if (p.shd) s.backgroundColor = p.shd;
    if (p.bdr) {
      for (const side of ['top', 'bottom', 'left', 'right']) if (p.bdr[side] && p.bdr[side] !== 'none') s[`border${side[0].toUpperCase()}${side.slice(1)}`] = p.bdr[side];
    }
  }

  lineHeight(p, r) {
    const sp = p.spacing || {};
    const line = Number(sp.line || 240);
    if (sp.lineRule === 'exact') return `${tw2pt(line)}pt`;
    if (sp.lineRule === 'atLeast') return `${Math.max(tw2pt(line), (r.sz || 11) * lineBase(r.font))}pt`;
    return String((line / 240) * lineBase(r.font));
  }

  async paragraph(el, ctx) {
    const direct = parsePPr(kid(el, 'pPr'));
    let p = this.effP(direct, ctx);
    // Numbering level indents sit between the style and the paragraph's own settings.
    const lvlDef = this.levelDef(p);
    if (lvlDef && lvlDef.pPr.ind) p = { ...p, ind: { ...(p.ind || {}), ...lvlDef.pPr.ind, ...(direct.ind || {}) } };
    if (direct.ind && direct.ind.firstLine !== undefined && p.ind) delete p.ind.hanging;
    const markR = parseRPr(kid(kid(el, 'pPr'), 'rPr'), this.theme);
    const baseR = this.effR(p.pStyle, {}, ctx);
    const out = [];
    const make = (first) => {
      const node = document.createElement('p');
      this.paraStyle(node, p, first ? ctx : { ...ctx, noBefore: true }, ctx.prevStyle);
      node.style.lineHeight = this.lineHeight(p, baseR);
      this.runStyle(node, { ...baseR, ...markR, b: baseR.b, i: baseR.i, u: false, strike: false, highlight: null, vertAlign: null });
      node.style.margin = '0';
      return node;
    };
    if (p.pageBreakBefore) out.push(this.breakMark());
    let cur = make(true);
    const mk = this.marker(p);
    if (mk) {
      const span = document.createElement('span');
      const r = this.effR(p.pStyle, {}, ctx, mk.lvl.rPr);
      this.runStyle(span, { ...r, u: false });
      if (mk.lvl.rPr.font && /symbol|wingdings/i.test(mk.lvl.rPr.font)) span.style.fontFamily = 'sans-serif';
      span.textContent = mk.text;
      const hang = p.ind && p.ind.hanging ? tw2pt(p.ind.hanging) : 0;
      span.style.display = 'inline-block';
      span.style.textIndent = '0';
      if (mk.lvl.suff === 'tab' && hang) span.style.minWidth = `${hang}pt`;
      else if (mk.lvl.suff !== 'nothing') span.textContent += ' ';
      span.style.whiteSpace = 'pre';
      cur.appendChild(span);
    }
    out.push(cur);
    const state = { fieldStack: [], cur, split: () => { out.push(this.breakMark()); cur = make(false); state.cur = cur; out.push(cur); } };
    await this.inline(el, p, ctx, state);
    for (const node of out) if (node.localName === 'p' && !node.textContent.length && !node.querySelector('img,div')) node.appendChild(document.createTextNode('​'));
    const styleId = p.pStyle || this.defaultPStyle;
    const firstP = out.find((n) => n.localName === 'p');
    if (ctx.prevNode && ctx.prevStyle === styleId && (p.contextualSpacing || ctx.prevContextual) && firstP) {
      ctx.prevNode.style.paddingBottom = '0pt';
      firstP.style.paddingTop = '0pt';
    }
    ctx.prevStyle = styleId;
    ctx.prevNode = [...out].reverse().find((n) => n.localName === 'p') || null;
    ctx.prevContextual = !!p.contextualSpacing;
    // Section break that starts a new page.
    const sect = kid(kid(el, 'pPr'), 'sectPr');
    if (sect && ctx.top) {
      const type = wv(kid(sect, 'type'));
      if (!type || /nextPage|oddPage|evenPage/.test(type)) out.push(this.breakMark());
    }
    return out;
  }

  breakMark() {
    const d = document.createElement('div');
    d.dataset.kpBreak = '1';
    d.style.height = '0';
    return d;
  }

  async inline(container, p, ctx, state) {
    for (const el of Array.from(container.children)) {
      const n = el.localName;
      if (n === 'r') await this.run(el, p, ctx, state);
      else if (n === 'hyperlink' || n === 'smartTag' || n === 'ins' || n === 'moveTo' || n === 'customXml' || n === 'fldSimple' || n === 'bdo' || n === 'dir') {
        if (n === 'fldSimple' && /PAGE\b/.test(wv(el, 'instr') || '') && ctx.hf) {
          const sp = document.createElement('span');
          sp.className = /NUMPAGES/.test(wv(el, 'instr')) ? 'kp-pages' : 'kp-page';
          sp.textContent = '1';
          state.cur.appendChild(sp);
          continue;
        }
        await this.inline(el, p, ctx, state);
      } else if (n === 'sdt') await this.inline(kid(el, 'sdtContent') || el, p, ctx, state);
      else if (n === 'AlternateContent') await this.inline(kid(el, 'Choice') || el, p, ctx, state);
    }
  }

  async run(el, p, ctx, state) {
    const rPr = parseRPr(kid(el, 'rPr'), this.theme);
    const r = this.effR(p.pStyle, rPr, ctx, ctx.condR);
    if (r.vanish) return;
    const fs = state.fieldStack;
    const hidden = () => fs.some((f) => f.phase === 'instr' || f.hide);
    let span = null;
    const getSpan = () => {
      if (!span) { span = document.createElement('span'); this.runStyle(span, r); state.cur.appendChild(span); }
      return span;
    };
    for (const c of Array.from(el.children)) {
      const n = c.localName;
      if (n === 'fldChar') {
        const t = wv(c, 'fldCharType');
        if (t === 'begin') fs.push({ phase: 'instr', instr: '' });
        else if (t === 'separate' && fs.length) {
          const f = fs[fs.length - 1];
          f.phase = 'result';
          if (ctx.hf && /\b(NUMPAGES|PAGE)\b/.test(f.instr)) {
            const sp = document.createElement('span');
            sp.className = /NUMPAGES/.test(f.instr) ? 'kp-pages' : 'kp-page';
            sp.textContent = '1';
            this.runStyle(sp, r);
            state.cur.appendChild(sp);
            f.hide = true;
          }
        } else if (t === 'end') {
          const f = fs.pop();
          if (f && f.phase === 'instr' && ctx.hf && /\b(NUMPAGES|PAGE)\b/.test(f.instr)) {
            const sp = document.createElement('span');
            sp.className = /NUMPAGES/.test(f.instr) ? 'kp-pages' : 'kp-page';
            this.runStyle(sp, r);
            state.cur.appendChild(sp);
          }
        }
        span = null;
        continue;
      }
      if (n === 'instrText') { if (fs.length) fs[fs.length - 1].instr += c.textContent; continue; }
      if (hidden()) continue;
      if (n === 't') getSpan().appendChild(document.createTextNode(c.textContent));
      else if (n === 'tab') getSpan().appendChild(document.createTextNode('\t'));
      else if (n === 'noBreakHyphen') getSpan().appendChild(document.createTextNode('‑'));
      else if (n === 'softHyphen') getSpan().appendChild(document.createTextNode('­'));
      else if (n === 'sym') {
        const ch = String.fromCharCode(parseInt(wv(c, 'char') || '2022', 16));
        getSpan().appendChild(document.createTextNode(PUA_BULLETS[ch] || (ch.charCodeAt(0) >= 0xf000 ? '•' : ch)));
      } else if (n === 'br') {
        const type = wv(c, 'type');
        if (type === 'page' && ctx.top) { state.split(); span = null; } else getSpan().appendChild(document.createElement('br'));
      } else if (n === 'cr') getSpan().appendChild(document.createElement('br'));
      else if (n === 'drawing') await this.drawing(c, state, ctx);
      else if (n === 'pict' || n === 'object') await this.vml(c, state);
      else if (n === 'AlternateContent') {
        const choice = kid(c, 'Choice');
        const d = choice && choice.getElementsByTagNameNS('*', 'drawing')[0];
        if (d) await this.drawing(d, state, ctx);
        else { const fb = kid(c, 'Fallback'); const pict = fb && fb.getElementsByTagNameNS('*', 'pict')[0]; if (pict) await this.vml(pict, state); }
        span = null;
      } else if (n === 'footnoteReference' && this.footnotes.has(wv(c, 'id'))) {
        this.footnoteRefs.push(wv(c, 'id'));
        const sup = document.createElement('sup');
        sup.textContent = String(this.footnoteRefs.length);
        sup.style.fontSize = '0.65em';
        sup.style.lineHeight = '0';
        state.cur.appendChild(sup);
        span = null;
      }
    }
  }

  async drawing(d, state, ctx) {
    const holder = kid(d, 'inline') || kid(d, 'anchor');
    if (!holder) return;
    const ext = kid(holder, 'extent');
    const w = emu2px(attr(ext, 'cx'));
    const h = emu2px(attr(ext, 'cy'));
    const blip = holder.getElementsByTagNameNS('*', 'blip')[0];
    const rels = ctx.rels;
    if (blip) {
      const rid = blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');
      const url = await this.imageUrl(rid, rels);
      if (!url) return;
      const img = document.createElement('img');
      img.src = url;
      img.style.cssText = `width:${w}px;height:${h}px;max-width:100%;object-fit:contain;vertical-align:bottom;text-indent:0;`;
      if (holder.localName === 'anchor') { img.style.display = 'block'; img.style.margin = '4px auto'; }
      state.cur.appendChild(img);
      return;
    }
    const txbx = holder.getElementsByTagNameNS('*', 'txbxContent')[0];
    if (txbx) {
      const box = document.createElement('div');
      box.dataset.kpKeep = '1';
      box.style.cssText = `display:inline-block;width:${Math.min(w, 2000)}px;max-width:100%;vertical-align:top;text-indent:0;white-space:normal;`;
      for (const b of await this.blocks(txbx, { ...ctx, top: false })) box.appendChild(b);
      state.cur.appendChild(box);
      return;
    }
    if (holder.getElementsByTagNameNS('*', 'chart').length) this.warnings.add('Charts aren’t drawn yet and were left out.');
  }

  async vml(pict, state) {
    const im = pict.getElementsByTagNameNS('*', 'imagedata')[0];
    if (!im) return;
    const rid = im.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    const url = await this.imageUrl(rid);
    if (!url) return;
    const shape = im.parentElement;
    const style = (shape && attr(shape, 'style')) || '';
    const dim = (k) => { const m = style.match(new RegExp(`${k}:([\\d.]+)(pt|px|in)?`)); if (!m) return null; const v = Number(m[1]); return m[2] === 'in' ? v * 96 : m[2] === 'px' ? v : v * 4 / 3; };
    const img = document.createElement('img');
    img.src = url;
    const w = dim('width'); const h = dim('height');
    img.style.cssText = `${w ? `width:${w}px;` : ''}${h ? `height:${h}px;` : ''}max-width:100%;object-fit:contain;vertical-align:bottom;`;
    state.cur.appendChild(img);
  }

  async table(tbl, ctx) {
    const tblPr = kid(tbl, 'tblPr');
    const styleId = wv(kid(tblPr, 'tblStyle'));
    const chain = this.styleChain(styleId);
    const borders = {};
    for (const src of [...chain.map((s) => s.tblBorders), kid(tblPr, 'tblBorders')]) {
      if (!src) continue;
      for (const side of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'start', 'end']) {
        const b = kid(src, side);
        if (b) borders[side === 'start' ? 'left' : side === 'end' ? 'right' : side] = parseBorder(b);
      }
    }
    const cond = Object.assign({}, ...chain.map((s) => s.cond));
    const look = kid(tblPr, 'tblLook');
    const firstRowOn = look ? (wv(look, 'firstRow') === '1' || (parseInt(wv(look, 'val') || '0', 16) & 0x20) > 0) : true;

    const cellMar = { left: 7.2, right: 7.2, top: 0, bottom: 0 };
    for (const src of [kid(tblPr, 'tblCellMar')]) {
      if (!src) continue;
      for (const [side, alt] of [['left', 'start'], ['right', 'end'], ['top'], ['bottom']]) {
        const b = kid(src, side) || (alt && kid(src, alt));
        if (b) cellMar[side] = tw2px(wv(b, 'w'));
      }
    }
    const table = document.createElement('table');
    table.style.cssText = 'border-collapse:collapse;table-layout:fixed;text-indent:0;margin:0;';
    const jc = wv(kid(tblPr, 'jc'));
    if (jc === 'center') { table.style.marginLeft = 'auto'; table.style.marginRight = 'auto'; }
    const ind = kid(tblPr, 'tblInd');
    if (ind && wv(ind, 'type') !== 'pct') table.style.marginLeft = `${tw2px(wv(ind, 'w')) - cellMar.left}px`;
    const grid = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) => tw2px(wv(g, 'w')));
    const totalW = grid.reduce((a, b) => a + b, 0);
    if (grid.length) {
      const cg = document.createElement('colgroup');
      for (const g of grid) { const col = document.createElement('col'); col.style.width = `${g}px`; cg.appendChild(col); }
      table.appendChild(cg);
      table.style.width = `${totalW}px`;
    } else table.style.width = '100%';
    table.style.maxWidth = '100%';

    const rows = kids(tbl, 'tr');
    const ncols = grid.length || Math.max(...rows.map((r) => kids(r, 'tc').reduce((s, c) => s + (Number(wv(kid(kid(c, 'tcPr'), 'gridSpan'))) || 1), 0)), 1);
    // Resolve vertical merges into rowspans.
    const layout = rows.map((tr) => {
      let col = Number(wv(kid(kid(tr, 'trPr'), 'gridBefore'))) || 0;
      return kids(tr, 'tc').map((tc) => {
        const tcPr = kid(tc, 'tcPr');
        const span = Number(wv(kid(tcPr, 'gridSpan'))) || 1;
        const vm = kid(tcPr, 'vMerge');
        const cell = { tc, tcPr, col, span, vMerge: vm ? (wv(vm) === 'restart' ? 'restart' : 'continue') : null, rowspan: 1 };
        col += span;
        return cell;
      });
    });
    layout.forEach((row, ri) => row.forEach((cell) => {
      if (cell.vMerge !== 'restart') return;
      for (let k = ri + 1; k < layout.length; k++) {
        const below = layout[k].find((c) => c.col === cell.col);
        if (below && below.vMerge === 'continue') { cell.rowspan++; below.skip = true; } else break;
      }
    }));

    const tbody = document.createElement('tbody');
    for (let ri = 0; ri < rows.length; ri++) {
      const tr = document.createElement('tr');
      const trPr = kid(rows[ri], 'trPr');
      const hgt = kid(trPr, 'trHeight');
      if (hgt) tr.style.height = `${tw2px(wv(hgt, 'val'))}px`;
      const isFirst = ri === 0 && firstRowOn && cond.firstRow;
      for (const cell of layout[ri]) {
        if (cell.skip) continue;
        const td = document.createElement('td');
        if (cell.span > 1) td.colSpan = cell.span;
        if (cell.rowspan > 1) td.rowSpan = cell.rowspan;
        const lastRow = ri + cell.rowspan - 1 === rows.length - 1;
        const tcB = kid(cell.tcPr, 'tcBorders');
        const side = (name, alt, tableSide) => {
          const own = tcB && (kid(tcB, name) || (alt && kid(tcB, alt)));
          return own ? parseBorder(own) : tableSide;
        };
        const bt = side('top', null, ri === 0 ? borders.top : borders.insideH);
        const bb = side('bottom', null, lastRow ? borders.bottom : borders.insideH);
        const bl = side('left', 'start', cell.col === 0 ? borders.left : borders.insideV);
        const br = side('right', 'end', cell.col + cell.span >= ncols ? borders.right : borders.insideV);
        td.style.borderTop = bt || 'none';
        td.style.borderBottom = bb || 'none';
        td.style.borderLeft = bl || 'none';
        td.style.borderRight = br || 'none';
        td.style.padding = `${cellMar.top}px ${cellMar.right}px ${cellMar.bottom}px ${cellMar.left}px`;
        const va = wv(kid(cell.tcPr, 'vAlign'));
        td.style.verticalAlign = va === 'center' ? 'middle' : va === 'bottom' ? 'bottom' : 'top';
        const shd = kid(cell.tcPr, 'shd');
        const fill = shd && wv(shd, 'fill');
        if (fill && fill !== 'auto') td.style.backgroundColor = `#${fill}`;
        else if (isFirst && cond.firstRow.shd) td.style.backgroundColor = cond.firstRow.shd;
        const cctx = { ...ctx, top: false, tblStyle: styleId, condR: isFirst ? cond.firstRow.rPr : null, prevStyle: undefined };
        for (const b of await this.blocks(cell.tc, cctx)) td.appendChild(b);
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    ctx.prevStyle = undefined;
    ctx.prevNode = null;
    return table;
  }

  async body_() {
    const ctx = { top: true, rels: this.docRels };
    const blocks = await this.blocks(this.body, ctx);
    if (this.footnoteRefs.length) {
      const sep = document.createElement('div');
      sep.style.cssText = 'border-top:0.75pt solid #000;width:33%;margin:12pt 0 6pt;height:0;';
      blocks.push(sep);
      for (let i = 0; i < this.footnoteRefs.length; i++) {
        const fn = this.footnotes.get(this.footnoteRefs[i]);
        const parts = await this.blocks(fn, { top: false, rels: this.fnRels });
        if (parts[0]) {
          const n = document.createElement('sup');
          n.textContent = `${i + 1} `;
          parts[0].insertBefore(n, parts[0].firstChild);
        }
        blocks.push(...parts);
      }
    }
    // Drop trailing page-break markers.
    while (blocks.length && blocks[blocks.length - 1].dataset && blocks[blocks.length - 1].dataset.kpBreak) blocks.pop();
    return blocks;
  }

  async headerFooter(kind, type) {
    const path = this.hf[`${kind}-${type}`];
    if (!path) return null;
    const doc = await this.xml(path);
    if (!doc) return null;
    const rels = await this.rels(path);
    const blocks = await this.blocks(doc.documentElement, { top: false, hf: true, rels });
    if (!blocks.length) return null;
    const box = document.createElement('div');
    blocks.forEach((b) => box.appendChild(b));
    if (!box.textContent.replace(/​/g, '').trim() && !box.querySelector('img')) return null;
    return box;
  }
}

export function defaultFont(renderer) {
  return { family: fontStack(renderer.defR.font) || 'sans-serif', size: renderer.defR.sz || 11 };
}
