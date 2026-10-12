// Renders .pptx slides as absolutely positioned HTML, one element per slide.
import { attr, isBrowserImage, kid, kids, mimeFromPath, parseXml, resolvePath } from './util.js';
import { fontStack } from './docx-render.js';

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const EMU = 9525; // per CSS px
const px = (v) => Number(v || 0) / EMU;
const rid = (el, name = 'embed') => (el ? el.getAttributeNS(R_NS, name) : null);

const PRST_COLORS = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080', darkGray: 'A9A9A9', lightGray: 'D3D3D3', orange: 'FFA500', purple: '800080' };
const SYMBOL_MAP = { '': '▪', '': '•', '': '➢', '': '✓', '': '❖', '': '●', '': '■', '': '➔', '': '◆', 'Ø': '➢', 'ü': '✓', 'q': '❑', 'v': '❖', 'n': '■', 'l': '●', '§': '▪', 'Ÿ': '•', 'ð': '➔' };

// ---------- colour helpers ----------
const hexToRgb = (h) => [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
const clamp = (v, a = 0, b = 255) => Math.min(b, Math.max(a, v));
function rgbToHsl([r, g, b]) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  return [h, s, l];
}
function hslToRgb([h, s, l]) {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t) => {
    if (t < 0) t += 1; if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255];
}
const toLin = (x) => { const c = x / 255; return c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4; };
const fromLin = (v) => { const c = Math.min(1, Math.max(0, v)); return 255 * (c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055); };
const css = (c) => (c ? (c.a < 1 ? `rgba(${c.rgb.map(Math.round).join(',')},${c.a.toFixed(3)})` : `rgb(${c.rgb.map(Math.round).join(',')})`) : 'transparent');

export class PptxRenderer {
  constructor(zip) { this.zip = zip; this.parts = new Map(); this.urls = new Map(); this.warnings = new Set(); }

  async part(path) {
    if (this.parts.has(path)) return this.parts.get(path);
    const f = this.zip.file(path);
    let res = null;
    if (f) {
      const xml = parseXml(await f.async('string'));
      const relsFile = this.zip.file(path.replace(/([^/]+)$/, '_rels/$1.rels'));
      const rels = new Map();
      if (relsFile) {
        const rx = parseXml(await relsFile.async('string'));
        for (const r of rx.getElementsByTagNameNS('*', 'Relationship')) {
          rels.set(attr(r, 'Id'), { target: attr(r, 'TargetMode') === 'External' ? null : resolvePath(path, attr(r, 'Target')), type: attr(r, 'Type') || '' });
        }
      }
      res = { xml, rels, path };
    }
    this.parts.set(path, res);
    return res;
  }

  relOfType(part, suffix) { return part && [...part.rels.values()].find((r) => r.type.endsWith(suffix)); }

  async load() {
    let presPath = 'ppt/presentation.xml';
    const relsXml = this.zip.file('_rels/.rels');
    if (relsXml) {
      const rx = parseXml(await relsXml.async('string'));
      for (const r of rx.getElementsByTagNameNS('*', 'Relationship')) if (/officeDocument$/.test(attr(r, 'Type'))) presPath = resolvePath('', attr(r, 'Target'));
    }
    const pres = await this.part(presPath);
    if (!pres) throw new Error('missing presentation');
    const sz = pres.xml.getElementsByTagNameNS('*', 'sldSz')[0];
    this.cx = Number(attr(sz, 'cx')) || 12192000;
    this.cy = Number(attr(sz, 'cy')) || 6858000;
    this.width = Math.round(this.cx / EMU);
    this.height = Math.round(this.cy / EMU);
    this.slides = [];
    for (const s of pres.xml.getElementsByTagNameNS('*', 'sldId')) {
      const rel = pres.rels.get(rid(s, 'id'));
      if (rel && rel.target) this.slides.push(rel.target);
    }
    this.defaultText = pres.xml.getElementsByTagNameNS('*', 'defaultTextStyle')[0] || null;
    const ts = this.zip.file('ppt/tableStyles.xml');
    this.tableStyles = new Map();
    if (ts) {
      const x = parseXml(await ts.async('string'));
      for (const st of x.getElementsByTagNameNS('*', 'tblStyle')) this.tableStyles.set(attr(st, 'styleId'), st);
    }
  }

  // ---------- context (slide → layout → master → theme) ----------
  async context(slidePath) {
    const slide = await this.part(slidePath);
    const layoutRel = this.relOfType(slide, '/slideLayout');
    const layout = layoutRel ? await this.part(layoutRel.target) : null;
    const masterRel = this.relOfType(layout, '/slideMaster');
    const master = masterRel ? await this.part(masterRel.target) : null;
    const themeRel = this.relOfType(master, '/theme');
    const theme = themeRel ? await this.part(themeRel.target) : null;
    const scheme = {};
    const fonts = { major: 'Calibri Light', minor: 'Calibri' };
    let lnStyles = [];
    if (theme) {
      const cs = theme.xml.getElementsByTagNameNS('*', 'clrScheme')[0];
      if (cs) for (const c of cs.children) {
        const s = kid(c, 'srgbClr');
        const sys = kid(c, 'sysClr');
        scheme[c.localName] = s ? attr(s, 'val') : sys ? (attr(sys, 'lastClr') || (attr(sys, 'val') === 'window' ? 'FFFFFF' : '000000')) : '000000';
      }
      const fs = theme.xml.getElementsByTagNameNS('*', 'fontScheme')[0];
      const latin = (n) => { const e = fs && kid(kid(fs, n), 'latin'); return e && attr(e, 'typeface'); };
      fonts.major = latin('majorFont') || fonts.major;
      fonts.minor = latin('minorFont') || fonts.minor;
      const lsl = theme.xml.getElementsByTagNameNS('*', 'lnStyleLst')[0];
      lnStyles = lsl ? kids(lsl, 'ln') : [];
    }
    const clrMap = {};
    const cm = master && master.xml.getElementsByTagNameNS('*', 'clrMap')[0];
    if (cm) for (const a of cm.attributes) clrMap[a.localName] = a.value;
    const ovr = slide.xml.getElementsByTagNameNS('*', 'clrMapOvr')[0];
    const om = ovr && kid(ovr, 'overrideClrMapping');
    if (om) for (const a of om.attributes) clrMap[a.localName] = a.value;
    const txStyles = master && master.xml.getElementsByTagNameNS('*', 'txStyles')[0];
    return { slide, layout, master, theme, scheme, fonts, clrMap, lnStyles, txStyles, number: 0 };
  }

  color(el, ctx, phClr) {
    if (!el) return null;
    const c = Array.from(el.children).find((x) => /^(srgbClr|schemeClr|sysClr|prstClr|scrgbClr|hslClr)$/.test(x.localName));
    if (!c) return null;
    let hex = '000000';
    const v = attr(c, 'val');
    switch (c.localName) {
      case 'srgbClr': hex = v; break;
      case 'sysClr': hex = attr(c, 'lastClr') || (v === 'window' ? 'FFFFFF' : '000000'); break;
      case 'prstClr': hex = PRST_COLORS[v] || '000000'; break;
      case 'scrgbClr': hex = ['r', 'g', 'b'].map((k) => Math.round(clamp(Number(attr(c, k)) / 100000, 0, 1) * 255).toString(16).padStart(2, '0')).join(''); break;
      case 'schemeClr': {
        if (v === 'phClr') { if (phClr) return this.applyMods(phClr, c); hex = ctx.scheme.accent1 || '4472C4'; break; }
        const mapped = ctx.clrMap[v] || ({ bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' })[v] || v;
        hex = ctx.scheme[mapped] || ctx.scheme[v] || '000000';
        break;
      }
      default: break;
    }
    return this.applyMods({ rgb: hexToRgb(hex.padStart(6, '0')), a: 1 }, c);
  }

  applyMods(base, c) {
    let rgb = base.rgb.slice();
    let a = base.a;
    for (const m of c.children) {
      const val = Number(attr(m, 'val')) / 100000;
      switch (m.localName) {
        case 'alpha': a = val; break;
        case 'lumMod': { const h = rgbToHsl(rgb); h[2] = clamp(h[2] * val, 0, 1); rgb = hslToRgb(h); break; }
        case 'lumOff': { const h = rgbToHsl(rgb); h[2] = clamp(h[2] + val, 0, 1); rgb = hslToRgb(h); break; }
        case 'satMod': { const h = rgbToHsl(rgb); h[1] = clamp(h[1] * val, 0, 1); rgb = hslToRgb(h); break; }
        case 'tint': rgb = rgb.map((x) => fromLin(1 - (1 - toLin(x)) * val)); break;
        case 'shade': rgb = rgb.map((x) => fromLin(toLin(x) * val)); break;
        default: break;
      }
    }
    return { rgb, a };
  }

  async imageUrl(part, id) {
    const rel = part && part.rels.get(id);
    if (!rel || !rel.target) return null;
    if (this.urls.has(rel.target)) return this.urls.get(rel.target);
    let url = null;
    const f = this.zip.file(rel.target);
    if (f && isBrowserImage(rel.target)) url = URL.createObjectURL(new Blob([await f.async('uint8array')], { type: mimeFromPath(rel.target) }));
    else if (f) this.warnings.add('Some pictures use formats browsers can’t draw (such as EMF/WMF/TIFF) and were left out.');
    this.urls.set(rel.target, url);
    return url;
  }

  release() { for (const u of this.urls.values()) if (u) URL.revokeObjectURL(u); }

  // Returns a CSS background value, or null for no fill.
  async fill(spPr, ctx, part, styleFill) {
    const f = spPr && Array.from(spPr.children).find((x) => /^(noFill|solidFill|gradFill|blipFill|pattFill|grpFill)$/.test(x.localName));
    if (!f) return styleFill || null;
    if (f.localName === 'noFill') return null;
    if (f.localName === 'solidFill') return css(this.color(f, ctx));
    if (f.localName === 'pattFill') return css(this.color(kid(f, 'fgClr'), ctx));
    if (f.localName === 'gradFill') {
      const stops = kids(kid(f, 'gsLst'), 'gs').map((g) => ({ pos: Number(attr(g, 'pos')) / 1000, c: this.color(g, ctx) })).sort((a, b) => a.pos - b.pos);
      if (!stops.length) return null;
      if (stops.length === 1) return css(stops[0].c);
      const lin = kid(f, 'lin');
      const path = kid(f, 'path');
      const list = stops.map((s) => `${css(s.c)} ${s.pos}%`).join(',');
      if (path) return `radial-gradient(circle,${list})`;
      const ang = lin ? Number(attr(lin, 'ang')) / 60000 : 90;
      return `linear-gradient(${ang + 90}deg,${list})`;
    }
    if (f.localName === 'blipFill') {
      const url = await this.imageUrl(part, rid(kid(f, 'blip')));
      if (!url) return null;
      return `url("${url}") center/${kid(f, 'tile') ? 'auto' : '100% 100%'} ${kid(f, 'tile') ? 'repeat' : 'no-repeat'}`;
    }
    return null;
  }

  line(ln, ctx, styleLine) {
    const out = styleLine ? { ...styleLine } : null;
    if (!ln) return out;
    if (kid(ln, 'noFill')) return null;
    const res = out || { width: 0.75 * 4 / 3, color: 'rgb(0,0,0)' };
    const w = attr(ln, 'w');
    if (w) res.width = Math.max(px(w), 0.5);
    const sf = kid(ln, 'solidFill');
    if (sf) res.color = css(this.color(sf, ctx));
    else if (!out && !kid(ln, 'gradFill')) return w ? res : null;
    const dash = attr(kid(ln, 'prstDash'), 'val');
    if (dash && dash !== 'solid') res.dash = /dot/i.test(dash) && !/dash/i.test(dash) ? 'dot' : 'dash';
    const head = kid(ln, 'headEnd');
    const tail = kid(ln, 'tailEnd');
    if (head && attr(head, 'type') && attr(head, 'type') !== 'none') res.head = true;
    if (tail && attr(tail, 'type') && attr(tail, 'type') !== 'none') res.tail = true;
    return res;
  }

  styleRefs(sp, ctx) {
    const st = kid(sp, 'style');
    if (!st) return {};
    const fr = kid(st, 'fillRef');
    const lr = kid(st, 'lnRef');
    const fo = kid(st, 'fontRef');
    const out = {};
    if (fr && Number(attr(fr, 'idx')) > 0) out.fill = css(this.color(fr, ctx));
    if (lr && Number(attr(lr, 'idx')) > 0) {
      const tmpl = ctx.lnStyles[Number(attr(lr, 'idx')) - 1];
      out.line = { width: tmpl && attr(tmpl, 'w') ? Math.max(px(attr(tmpl, 'w')), 0.5) : 1, color: css(this.color(lr, ctx)) };
    }
    if (fo) out.fontColor = this.color(fo, ctx);
    return out;
  }

  // ---------- placeholders ----------
  phInfo(sp) {
    const nv = sp && Array.from(sp.children).find((c) => /^nv(Sp|Pic|GraphicFrame|GrpSp|CxnSp)Pr$/.test(c.localName));
    const ph = nv && kid(nv, 'nvPr') && kid(kid(nv, 'nvPr'), 'ph');
    if (!ph) return null;
    return { type: attr(ph, 'type') || 'obj', idx: attr(ph, 'idx') };
  }

  findPh(part, info) {
    if (!part || !info) return null;
    const shapes = Array.from(part.xml.getElementsByTagNameNS('*', 'sp'));
    const withInfo = shapes.map((s) => ({ s, i: this.phInfo(s) })).filter((x) => x.i);
    const norm = (t) => (t === 'ctrTitle' ? 'title' : t === 'subTitle' || t === 'obj' ? 'body' : t);
    if (info.idx) {
      const m = withInfo.find((x) => x.i.idx === info.idx);
      if (m) return m.s;
    }
    const m = withInfo.find((x) => x.i.type === info.type) || withInfo.find((x) => norm(x.i.type) === norm(info.type));
    return m ? m.s : null;
  }

  // ---------- text ----------
  readLvl(el) {
    if (!el) return {};
    const p = {};
    for (const a of ['marL', 'indent']) if (attr(el, a) !== null) p[a] = px(attr(el, a));
    if (attr(el, 'algn')) p.algn = attr(el, 'algn');
    const ls = kid(el, 'lnSpc');
    if (ls) { const pct = kid(ls, 'spcPct'); const pts = kid(ls, 'spcPts'); p.lnSpc = pct ? { pct: Number(attr(pct, 'val')) / 100000 } : { pt: Number(attr(pts, 'val')) / 100 }; }
    for (const k of ['spcBef', 'spcAft']) {
      const s = kid(el, k);
      if (s) { const pct = kid(s, 'spcPct'); const pts = kid(s, 'spcPts'); p[k] = pct ? { pct: Number(attr(pct, 'val')) / 100000 } : { pt: Number(attr(pts, 'val')) / 100 }; }
    }
    if (kid(el, 'buNone')) p.bu = { type: 'none' };
    const bc = kid(el, 'buChar');
    if (bc) p.bu = { type: 'char', char: attr(bc, 'char') };
    const ba = kid(el, 'buAutoNum');
    if (ba) p.bu = { type: 'auto', scheme: attr(ba, 'type'), start: Number(attr(ba, 'startAt') || 1) };
    if (kid(el, 'buFont')) p.buFont = attr(kid(el, 'buFont'), 'typeface');
    if (kid(el, 'buClr')) p.buClr = kid(el, 'buClr');
    if (kid(el, 'buSzPct')) p.buSz = Number(attr(kid(el, 'buSzPct'), 'val')) / 100000;
    const d = kid(el, 'defRPr');
    if (d) p.rPr = this.readR(d);
    return p;
  }

  readR(el) {
    const r = {};
    if (!el) return r;
    if (attr(el, 'sz')) r.sz = Number(attr(el, 'sz')) / 100;
    for (const k of ['b', 'i']) if (attr(el, k) !== null) r[k] = attr(el, k) === '1' || attr(el, k) === 'true';
    if (attr(el, 'u')) r.u = attr(el, 'u') !== 'none';
    if (attr(el, 'strike')) r.strike = attr(el, 'strike') !== 'noStrike';
    if (attr(el, 'baseline')) r.baseline = Number(attr(el, 'baseline')) / 1000;
    if (attr(el, 'cap')) r.cap = attr(el, 'cap');
    if (attr(el, 'spc')) r.spc = Number(attr(el, 'spc')) / 100;
    const sf = kid(el, 'solidFill');
    if (sf) r.fill = sf;
    if (kid(el, 'noFill')) r.fill = 'none';
    const gf = kid(el, 'gradFill');
    if (gf) { const gs = gf.getElementsByTagNameNS('*', 'gs')[0]; if (gs) r.fill = gs; }
    const hl = kid(el, 'highlight');
    if (hl) r.highlight = hl;
    const latin = kid(el, 'latin');
    if (latin && attr(latin, 'typeface')) r.font = attr(latin, 'typeface');
    return r;
  }

  lstLevels(lst) {
    const out = {};
    if (!lst) return out;
    for (let i = 1; i <= 9; i++) { const l = kid(lst, `lvl${i}pPr`); if (l) out[i - 1] = this.readLvl(l); }
    const def = kid(lst, 'defPPr');
    if (def) out.def = this.readLvl(def);
    return out;
  }

  mergeLvl(a, b) {
    if (!b) return a;
    const out = { ...a, ...b };
    out.rPr = { ...(a.rPr || {}), ...(b.rPr || {}) };
    return out;
  }

  // Text style chain for a shape, level by level.
  textChain(sp, ctx, phLayout, phMaster, info) {
    const chain = [];
    if (info) {
      const t = info.type;
      const txs = ctx.txStyles;
      const which = t === 'title' || t === 'ctrTitle' ? 'titleStyle' : /^(body|subTitle|obj)$/.test(t) || !t ? 'bodyStyle' : 'otherStyle';
      if (txs) chain.push(this.lstLevels(kid(txs, which)));
    } else {
      chain.push(this.lstLevels(this.defaultText));
    }
    for (const s of [phMaster, phLayout, sp]) {
      const tb = s && kid(s, 'txBody');
      if (tb) chain.push(this.lstLevels(kid(tb, 'lstStyle')));
    }
    return chain;
  }

  fontName(name, ctx) {
    if (!name) return null;
    if (name.startsWith('+mj')) return ctx.fonts.major;
    if (name.startsWith('+mn')) return ctx.fonts.minor;
    return name;
  }

  async text(txBody, sp, ctx, chain, bodyPrs, defaultColor, w, h) {
    const bodyPr = bodyPrs.reduce((acc, b) => {
      if (b) for (const a of b.attributes) acc.attrs[a.localName] = a.value;
      if (b && kid(b, 'normAutofit')) acc.autofit = kid(b, 'normAutofit');
      if (b && kid(b, 'noAutofit')) acc.autofit = null;
      return acc;
    }, { attrs: {}, autofit: null });
    const a = bodyPr.attrs;
    const ins = (k, d) => px(a[k] !== undefined ? a[k] : d);
    const box = document.createElement('div');
    const anchor = a.anchor || 't';
    const vert = a.vert && a.vert !== 'horz';
    box.style.cssText = `position:absolute;inset:0;display:flex;flex-direction:column;justify-content:${anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start'};padding:${ins('tIns', 45720)}px ${ins('rIns', 91440)}px ${ins('bIns', 45720)}px ${ins('lIns', 91440)}px;box-sizing:border-box;overflow:visible;`;
    if (a.wrap === 'none') box.style.whiteSpace = 'nowrap';
    if (vert) { box.style.writingMode = 'vertical-rl'; if (a.vert === 'vert270') box.style.transform = 'rotate(180deg)'; }
    let fontScale = 1, lnRed = 0;
    if (bodyPr.autofit) {
      fontScale = attr(bodyPr.autofit, 'fontScale') ? Number(attr(bodyPr.autofit, 'fontScale')) / 100000 : 1;
      lnRed = attr(bodyPr.autofit, 'lnSpcReduction') ? Number(attr(bodyPr.autofit, 'lnSpcReduction')) / 100000 : 0;
    }
    const counters = [];
    let any = false;
    for (const p of kids(txBody, 'p')) {
      const pPrEl = kid(p, 'pPr');
      const lvl = Number(attr(pPrEl, 'lvl') || 0);
      let pp = {};
      for (const c of chain) pp = this.mergeLvl(this.mergeLvl(pp, c.def), c[lvl]);
      pp = this.mergeLvl(pp, this.readLvl(pPrEl));
      const div = document.createElement('div');
      const st = div.style;
      st.margin = '0';
      st.textAlign = ({ ctr: 'center', r: 'right', just: 'justify', dist: 'justify' })[pp.algn] || 'left';
      st.paddingLeft = `${Math.max(0, pp.marL || 0)}px`;
      st.textIndent = `${pp.indent || 0}px`;
      const runs = Array.from(p.children).filter((c) => /^(r|br|fld)$/.test(c.localName));
      const endR = this.readR(kid(p, 'endParaRPr'));
      const firstR = runs.find((r) => r.localName !== 'br');
      const baseR = { ...(pp.rPr || {}), ...(firstR ? this.readR(kid(firstR, 'rPr')) : endR) };
      const baseSz = (baseR.sz || 18) * fontScale;
      st.fontSize = `${baseSz}pt`;
      const lsp = pp.lnSpc;
      st.lineHeight = lsp && lsp.pt ? `${lsp.pt}pt` : String(Math.max(0.6, (lsp && lsp.pct ? lsp.pct : 1) - lnRed) * 1.2);
      const spc = (s) => (!s ? 0 : s.pt !== undefined ? s.pt : s.pct * baseSz * 1.2);
      st.paddingTop = `${spc(pp.spcBef)}pt`;
      st.paddingBottom = `${spc(pp.spcAft)}pt`;
      if (!runs.length || runs.every((r) => r.localName === 'br')) {
        div.appendChild(document.createTextNode('​'));
        box.appendChild(div);
        counters.length = lvl;
        continue;
      }
      // Bullet
      if (pp.bu && pp.bu.type !== 'none' && runs.some((r) => (r.textContent || '').trim())) {
        const b = document.createElement('span');
        let ch;
        if (pp.bu.type === 'char') {
          ch = pp.bu.char || '•';
          ch = SYMBOL_MAP[ch] || (ch.charCodeAt(0) >= 0xf000 ? '•' : ch);
        } else {
          counters.length = lvl + 1;
          counters[lvl] = counters[lvl] === undefined ? pp.bu.start : counters[lvl] + 1;
          ch = autoNum(counters[lvl], pp.bu.scheme);
        }
        b.textContent = ch;
        const bc = pp.buClr ? this.color(pp.buClr, ctx) : null;
        const rc = baseR.fill && baseR.fill !== 'none' ? this.color(baseR.fill.localName === 'gs' ? baseR.fill : baseR.fill, ctx) : defaultColor;
        b.style.cssText = `display:inline-block;text-indent:0;min-width:${Math.max(0, -(pp.indent || 0))}px;color:${css(bc || rc)};${pp.buSz ? `font-size:${baseSz * pp.buSz}pt;` : ''}${pp.buFont && !/symbol|wingdings/i.test(pp.buFont) ? `font-family:${fontStack(this.fontName(pp.buFont, ctx))};` : 'font-family:Arial,sans-serif;'}`;
        div.appendChild(b);
      } else if (!pp.bu || pp.bu.type === 'none') counters.length = lvl;
      for (const r of runs) {
        if (r.localName === 'br') { div.appendChild(document.createElement('br')); continue; }
        const rp = { ...(pp.rPr || {}), ...this.readR(kid(r, 'rPr')) };
        let t = (kid(r, 't') || {}).textContent || '';
        if (r.localName === 'fld' && /slidenum/.test(attr(r, 'type') || '')) t = String(ctx.number);
        if (!t) continue;
        any = any || !!t.trim();
        const span = document.createElement('span');
        span.textContent = t;
        const s = span.style;
        const sz = (rp.sz || 18) * fontScale;
        s.fontSize = `${sz}pt`;
        const font = this.fontName(rp.font, ctx) || ctx.fonts.minor;
        s.fontFamily = fontStack(font);
        if (rp.b) s.fontWeight = '700';
        if (rp.i) s.fontStyle = 'italic';
        const deco = [rp.u && 'underline', rp.strike && 'line-through'].filter(Boolean);
        if (deco.length) s.textDecoration = deco.join(' ');
        if (rp.cap === 'all') s.textTransform = 'uppercase';
        if (rp.cap === 'small') s.fontVariant = 'small-caps';
        if (rp.spc) s.letterSpacing = `${rp.spc}pt`;
        if (rp.baseline) { s.verticalAlign = rp.baseline > 0 ? 'super' : 'sub'; s.fontSize = `${sz * 0.65}pt`; }
        let col = defaultColor;
        if (rp.fill === 'none') col = { rgb: [0, 0, 0], a: 0 };
        else if (rp.fill) col = this.color(rp.fill, ctx) || col;
        s.color = css(col);
        if (rp.highlight) s.backgroundColor = css(this.color(rp.highlight, ctx));
        div.appendChild(span);
      }
      box.appendChild(div);
    }
    void sp; void w; void h;
    return any ? box : null;
  }

  // ---------- shapes ----------
  xfrm(el) {
    const x = el && (kid(el, 'xfrm'));
    if (!x) return null;
    const off = kid(x, 'off'); const ext = kid(x, 'ext');
    const chOff = kid(x, 'chOff'); const chExt = kid(x, 'chExt');
    return {
      x: px(attr(off, 'x')), y: px(attr(off, 'y')), w: px(attr(ext, 'cx')), h: px(attr(ext, 'cy')),
      rot: Number(attr(x, 'rot') || 0) / 60000, flipH: attr(x, 'flipH') === '1', flipV: attr(x, 'flipV') === '1',
      chX: chOff ? px(attr(chOff, 'x')) : null, chY: chOff ? px(attr(chOff, 'y')) : null,
      chW: chExt ? px(attr(chExt, 'cx')) : null, chH: chExt ? px(attr(chExt, 'cy')) : null,
    };
  }

  place(el, b) {
    el.style.position = 'absolute';
    el.style.left = `${b.x}px`;
    el.style.top = `${b.y}px`;
    el.style.width = `${Math.max(b.w, 0)}px`;
    el.style.height = `${Math.max(b.h, 0)}px`;
    const t = [];
    if (b.rot) t.push(`rotate(${b.rot}deg)`);
    if (b.flipH) t.push('scaleX(-1)');
    if (b.flipV) t.push('scaleY(-1)');
    if (t.length) el.style.transform = t.join(' ');
  }

  async shapes(tree, ctx, part, map, out, opts = {}) {
    for (const el of Array.from(tree.children)) {
      const n = el.localName;
      try {
        if (n === 'sp' || n === 'cxnSp') {
          const info = this.phInfo(el);
          if (info && opts.skipPlaceholders) continue;
          const node = await this.shape(el, ctx, part, map, info);
          if (node) out.appendChild(node);
        } else if (n === 'pic') {
          const node = await this.picture(el, ctx, part, map);
          if (node) out.appendChild(node);
        } else if (n === 'grpSp') {
          const g = this.xfrm(kid(el, 'grpSpPr'));
          if (!g) continue;
          const sx = g.chW ? g.w / g.chW : 1;
          const sy = g.chH ? g.h / g.chH : 1;
          const inner = (b) => ({ ...b, x: g.x + (b.x - (g.chX || 0)) * sx, y: g.y + (b.y - (g.chY || 0)) * sy, w: b.w * sx, h: b.h * sy });
          const sub = (b) => inner(map(b));
          // Map child coordinates through this group, then through any outer group.
          const outer = map;
          const composed = (b) => outer(inner(b));
          void sub;
          if (g.rot || g.flipH || g.flipV) {
            const wrap = document.createElement('div');
            const ob = outer({ x: g.x, y: g.y, w: g.w, h: g.h, rot: g.rot, flipH: g.flipH, flipV: g.flipV });
            this.place(wrap, ob);
            const local = (b) => { const m = composed(b); return { ...m, x: m.x - ob.x, y: m.y - ob.y }; };
            await this.shapes(el, ctx, part, local, wrap, opts);
            out.appendChild(wrap);
          } else await this.shapes(el, ctx, part, composed, out, opts);
        } else if (n === 'graphicFrame') {
          const node = await this.frame(el, ctx, part, map);
          if (node) out.appendChild(node);
        } else if (n === 'AlternateContent') {
          await this.shapes(kid(el, 'Choice') || el, ctx, part, map, out, opts);
        }
      } catch (e) {
        console.warn('Kopya: skipped a shape', e);
      }
    }
  }

  async shape(sp, ctx, part, map, info) {
    const spPr = kid(sp, 'spPr');
    let phL = null, phM = null;
    if (info && part === ctx.slide) { phL = this.findPh(ctx.layout, info); phM = this.findPh(ctx.master, info); }
    else if (info && part === ctx.layout) phM = this.findPh(ctx.master, info);
    let box = this.xfrm(spPr) || this.xfrm(kid(phL, 'spPr')) || this.xfrm(kid(phM, 'spPr'));
    if (!box) return null;
    box = map(box);
    const refs = this.styleRefs(sp, ctx);
    const geom = kid(spPr, 'prstGeom');
    const cust = kid(spPr, 'custGeom');
    const prst = geom ? attr(geom, 'prst') : cust ? 'cust' : 'rect';
    const bg = await this.fill(spPr, ctx, part, refs.fill) ?? (phL || phM ? await this.fill(kid(phL || phM, 'spPr'), ctx, part, null) : null);
    const ln = this.line(kid(spPr, 'ln'), ctx, refs.line);
    const el = document.createElement('div');
    this.place(el, box);
    const isLine = /^(line|straightConnector1|bentConnector\d|curvedConnector\d)$/.test(prst) || sp.localName === 'cxnSp';
    if (isLine) {
      if (ln) el.appendChild(lineSvg(box, ln));
    } else if (prst === 'rect' || prst === 'roundRect' || prst === 'ellipse' || prst === 'snip1Rect' || prst === 'flowChartProcess' || prst === 'flowChartAlternateProcess') {
      const s = el.style;
      if (bg) s.background = bg;
      if (ln) { s.border = `${ln.width}px ${ln.dash === 'dot' ? 'dotted' : ln.dash ? 'dashed' : 'solid'} ${ln.color}`; s.boxSizing = 'border-box'; }
      if (prst === 'ellipse') s.borderRadius = '50%';
      if (prst === 'roundRect' || prst === 'flowChartAlternateProcess') {
        const gd = geom && geom.getElementsByTagNameNS('*', 'gd')[0];
        const adj = gd ? Number((attr(gd, 'fmla') || '').split(' ')[1] || 16667) / 100000 : 0.16667;
        s.borderRadius = `${Math.min(box.w, box.h) * adj}px`;
      }
    } else {
      const svg = geomSvg(prst, box, bg, ln, cust, geom);
      if (svg) el.appendChild(svg);
      else if (bg || ln) { el.style.background = bg || ''; if (ln) el.style.border = `${ln.width}px solid ${ln.color}`; }
    }
    const tb = kid(sp, 'txBody');
    if (tb) {
      const chain = this.textChain(sp, ctx, phL, phM, info);
      if (refs.fontColor && !info && chain[0]) {
        // A shape's own font colour (from its style) beats the presentation-wide default.
        const base = {};
        for (const [k, v] of Object.entries(chain[0])) base[k] = { ...v, rPr: { ...(v.rPr || {}) } };
        for (const v of Object.values(base)) delete v.rPr.fill;
        chain[0] = base;
      }
      const defColor = refs.fontColor || this.color(null, ctx) || this.defaultTextColor(ctx, chain);
      const bodies = [kid(kid(phM, 'txBody'), 'bodyPr'), kid(kid(phL, 'txBody'), 'bodyPr'), kid(tb, 'bodyPr')];
      const text = await this.text(tb, sp, ctx, chain, bodies, defColor, box.w, box.h);
      if (text) {
        const ins = textRect(prst, box.w, box.h);
        if (ins) text.style.inset = `${ins[1]}px ${ins[2]}px ${ins[3]}px ${ins[0]}px`;
        // Text never flips with the shape.
        if (box.flipH || box.flipV) text.style.transform = `${box.flipH ? 'scaleX(-1)' : ''} ${box.flipV ? 'scaleY(-1)' : ''}`;
        el.appendChild(text);
      }
    }
    if (!el.childNodes.length && !el.style.background && !el.style.border) return null;
    return el;
  }

  defaultTextColor(ctx, chain) {
    for (let i = chain.length - 1; i >= 0; i--) {
      const c = chain[i];
      const l = c[0] || c.def;
      if (l && l.rPr && l.rPr.fill && l.rPr.fill !== 'none') return this.color(l.rPr.fill, ctx);
    }
    const tx1 = ctx.scheme[ctx.clrMap.tx1 || 'dk1'] || '000000';
    return { rgb: hexToRgb(tx1), a: 1 };
  }

  async picture(pic, ctx, part, map) {
    const spPr = kid(pic, 'spPr');
    const info = this.phInfo(pic);
    let box = this.xfrm(spPr);
    if (!box && info) box = this.xfrm(kid(this.findPh(ctx.layout, info), 'spPr')) || this.xfrm(kid(this.findPh(ctx.master, info), 'spPr'));
    if (!box) return null;
    box = map(box);
    const bf = kid(pic, 'blipFill');
    const url = await this.imageUrl(part, rid(kid(bf, 'blip')));
    if (!url) return null;
    const el = document.createElement('div');
    this.place(el, box);
    el.style.overflow = 'hidden';
    const geom = attr(kid(spPr, 'prstGeom'), 'prst');
    if (geom === 'ellipse') el.style.borderRadius = '50%';
    if (geom === 'roundRect') el.style.borderRadius = `${Math.min(box.w, box.h) * 0.16667}px`;
    const src = kid(bf, 'srcRect');
    const c = (k) => (src && attr(src, k) ? Number(attr(src, k)) / 100000 : 0);
    const l = c('l'), t = c('t'), r = c('r'), b = c('b');
    const img = document.createElement('img');
    img.src = url;
    const iw = box.w / Math.max(0.01, 1 - l - r);
    const ih = box.h / Math.max(0.01, 1 - t - b);
    img.style.cssText = `position:absolute;left:${-l * iw}px;top:${-t * ih}px;width:${iw}px;height:${ih}px;max-width:none;`;
    el.appendChild(img);
    const ln = this.line(kid(spPr, 'ln'), ctx, null);
    if (ln) { const o = document.createElement('div'); o.style.cssText = `position:absolute;inset:0;border:${ln.width}px solid ${ln.color};border-radius:inherit;`; el.appendChild(o); }
    return el;
  }

  async frame(gf, ctx, part, map) {
    let box = this.xfrm(gf);
    if (!box) return null;
    box = map(box);
    const data = gf.getElementsByTagNameNS('*', 'graphicData')[0];
    const uri = attr(data, 'uri') || '';
    if (/table/.test(uri)) return this.table(data.getElementsByTagNameNS('*', 'tbl')[0], ctx, box);
    if (/diagram/.test(uri)) {
      const ids = data.getElementsByTagNameNS('*', 'relIds')[0];
      const dmRel = ids && part.rels.get(rid(ids, 'dm'));
      const dm = dmRel && await this.part(dmRel.target);
      const ext = dm && dm.xml.getElementsByTagNameNS('*', 'dataModelExt')[0];
      const drawRel = ext && part.rels.get(attr(ext, 'relId'));
      const drawing = drawRel && await this.part(drawRel.target);
      const tree = drawing && drawing.xml.getElementsByTagNameNS('*', 'spTree')[0];
      if (!tree) { this.warnings.add('A SmartArt graphic couldn’t be drawn.'); return null; }
      const wrap = document.createElement('div');
      this.place(wrap, box);
      await this.shapes(tree, ctx, drawing, (b) => b, wrap);
      return wrap;
    }
    if (/chart/.test(uri)) {
      this.warnings.add('Charts aren’t drawn yet and were left out.');
      return null;
    }
    // OLE objects and others: draw their preview picture when present.
    const pic = data.getElementsByTagNameNS('*', 'pic')[0];
    if (pic) {
      const bf = kid(pic, 'blipFill');
      const url = await this.imageUrl(part, rid(kid(bf, 'blip')));
      if (url) { const el = document.createElement('img'); el.src = url; this.place(el, box); return el; }
    }
    return null;
  }

  tableStyle(id, ctx) {
    const st = this.tableStyles.get(id);
    if (!st) {
      if (id === '{2D5ABB26-0587-4C30-8999-92F81FD0307C}') return null; // "No Style, No Grid"
      // PowerPoint's built-in default (Medium Style 2): themed header, banded rows, white rules.
      const acc = ctx.scheme.accent1 || '4472C4';
      const mix = (t) => css({ rgb: hexToRgb(acc).map((x) => fromLin(1 - (1 - toLin(x)) * t)), a: 1 });
      const white = { width: 1.33, color: 'rgb(255,255,255)' };
      const lt1 = { rgb: hexToRgb(ctx.scheme.lt1 || 'FFFFFF'), a: 1 };
      return {
        whole: { fill: mix(0.2), color: null, bold: false, borders: { left: white, right: white, top: white, bottom: white, insideH: white, insideV: white } },
        band1H: { fill: mix(0.4), color: null, bold: false, borders: {} },
        firstRow: { fill: css({ rgb: hexToRgb(acc), a: 1 }), color: lt1, bold: true, borders: { bottom: { width: 4, color: 'rgb(255,255,255)' } } },
        lastRow: { fill: css({ rgb: hexToRgb(acc), a: 1 }), color: lt1, bold: true, borders: {} },
        firstCol: { fill: null, color: null, bold: true, borders: {} },
      };
    }
    const part = (name) => {
      const p = kid(st, name);
      if (!p) return null;
      const tc = kid(p, 'tcStyle');
      const tx = kid(p, 'tcTxStyle');
      const fill = tc && kid(tc, 'fill');
      const bdr = tc && kid(tc, 'tcBdr');
      const borders = {};
      if (bdr) for (const side of ['left', 'right', 'top', 'bottom', 'insideH', 'insideV']) {
        const s = kid(bdr, side);
        const ln = s && kid(s, 'ln');
        if (ln) borders[side] = this.line(ln, ctx, null);
      }
      return {
        fill: fill ? css(this.color(kid(fill, 'solidFill'), ctx)) : null,
        color: tx ? this.color(tx, ctx) : null,
        bold: tx ? attr(tx, 'b') === 'on' : false,
        borders,
      };
    };
    return { whole: part('wholeTbl'), band1H: part('band1H'), firstRow: part('firstRow'), lastRow: part('lastRow'), firstCol: part('firstCol') };
  }

  async table(tbl, ctx, box) {
    if (!tbl) return null;
    const tblPr = kid(tbl, 'tblPr');
    const styleIdEl = tblPr && kid(tblPr, 'tableStyleId');
    const ts = styleIdEl ? this.tableStyle(styleIdEl.textContent.trim(), ctx) : null;
    const flags = { firstRow: attr(tblPr, 'firstRow') === '1', bandRow: attr(tblPr, 'bandRow') === '1', lastRow: attr(tblPr, 'lastRow') === '1', firstCol: attr(tblPr, 'firstCol') === '1' };
    const wrap = document.createElement('div');
    this.place(wrap, { ...box, h: box.h });
    wrap.style.height = 'auto';
    const table = document.createElement('table');
    table.style.cssText = 'border-collapse:collapse;table-layout:fixed;width:100%;';
    const cols = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) => px(attr(g, 'w')));
    const cg = document.createElement('colgroup');
    for (const w of cols) { const c = document.createElement('col'); c.style.width = `${w}px`; cg.appendChild(c); }
    table.appendChild(cg);
    const rows = kids(tbl, 'tr');
    const chain = [this.lstLevels(this.defaultText)];
    rows.forEach((tr, ri) => {
      const row = document.createElement('tr');
      row.style.height = `${px(attr(tr, 'h'))}px`;
      table.appendChild(row);
      row._cells = kids(tr, 'tc').map((tc, ci) => ({ tc, ri, ci }));
    });
    const trs = Array.from(table.querySelectorAll('tr'));
    for (const row of trs) {
      for (const { tc, ri, ci } of row._cells) {
        if (attr(tc, 'hMerge') === '1' || attr(tc, 'vMerge') === '1') continue;
        const td = document.createElement('td');
        if (attr(tc, 'gridSpan')) td.colSpan = Number(attr(tc, 'gridSpan'));
        if (attr(tc, 'rowSpan')) td.rowSpan = Number(attr(tc, 'rowSpan'));
        const tcPr = kid(tc, 'tcPr');
        const parts = [];
        if (ts) {
          parts.push(ts.whole);
          if (flags.bandRow && ts.band1H && (ri - (flags.firstRow ? 1 : 0)) % 2 === 0 && !(flags.firstRow && ri === 0)) parts.push(ts.band1H);
          if (flags.firstCol && ci === 0 && ts.firstCol) parts.push(ts.firstCol);
          if (flags.firstRow && ri === 0 && ts.firstRow) parts.push(ts.firstRow);
          if (flags.lastRow && ri === rows.length - 1 && ts.lastRow) parts.push(ts.lastRow);
        }
        const eff = parts.filter(Boolean).reduce((a, p) => ({
          fill: p.fill || a.fill, color: p.color || a.color, bold: p.bold || a.bold, borders: { ...a.borders, ...p.borders },
        }), { fill: null, color: null, bold: false, borders: {} });
        const own = await this.fill(tcPr, ctx, null, null);
        const fillEl = tcPr && Array.from(tcPr.children).find((x) => /Fill$/.test(x.localName));
        td.style.background = fillEl ? (own || 'transparent') : (eff.fill || 'transparent');
        const side = (k, tableKey) => {
          const ln = tcPr && kid(tcPr, k);
          const l = ln ? this.line(ln, ctx, null) : eff.borders[tableKey] || null;
          return l ? `${l.width}px solid ${l.color}` : 'none';
        };
        td.style.borderLeft = side('lnL', ci === 0 ? 'left' : 'insideV');
        td.style.borderRight = side('lnR', 'right');
        td.style.borderTop = side('lnT', ri === 0 ? 'top' : 'insideH');
        td.style.borderBottom = side('lnB', ri === rows.length - 1 ? 'bottom' : 'insideH');
        const m = (k, d) => px(attr(tcPr, k) ?? d);
        td.style.padding = `${m('marT', 45720)}px ${m('marR', 91440)}px ${m('marB', 45720)}px ${m('marL', 91440)}px`;
        const anchor = attr(tcPr, 'anchor');
        td.style.verticalAlign = anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top';
        const tb = kid(tc, 'txBody');
        if (tb) {
          const defColor = eff.color || this.defaultTextColor(ctx, []);
          let cellChain = chain;
          if (eff.color) {
            const base = {};
            for (const [k, v] of Object.entries(chain[0])) base[k] = { ...v, rPr: { ...(v.rPr || {}) } };
            for (const v of Object.values(base)) delete v.rPr.fill;
            cellChain = [base];
          }
          const content = await this.text(tb, tc, ctx, cellChain, [kid(tb, 'bodyPr')], defColor, 0, 0);
          if (content) {
            content.style.position = 'static';
            content.style.padding = '0';
            content.style.display = 'block';
            if (eff.bold) content.querySelectorAll('span').forEach((s) => { if (!s.style.fontWeight) s.style.fontWeight = '700'; });
            td.appendChild(content);
          }
        }
        row.appendChild(td);
      }
      delete row._cells;
    }
    wrap.appendChild(table);
    return wrap;
  }

  async background(ctx) {
    for (const part of [ctx.slide, ctx.layout, ctx.master]) {
      if (!part) continue;
      const bg = part.xml.getElementsByTagNameNS('*', 'bg')[0];
      if (!bg) continue;
      const pr = kid(bg, 'bgPr');
      if (pr) return (await this.fill(pr, ctx, part, null)) || '#fff';
      const ref = kid(bg, 'bgRef');
      if (ref) return css(this.color(ref, ctx));
    }
    return '#fff';
  }

  async render(index) {
    const ctx = await this.context(this.slides[index]);
    ctx.number = index + 1;
    const hidden = attr(ctx.slide.xml.documentElement, 'show') === '0';
    const el = document.createElement('div');
    el.style.cssText = `position:relative;width:${this.width}px;height:${this.height}px;overflow:hidden;background:${await this.background(ctx)};`;
    const showMaster = (p) => attr(p.xml.documentElement, 'showMasterSp') !== '0';
    const tree = (p) => p && p.xml.getElementsByTagNameNS('*', 'spTree')[0];
    const id = (b) => b;
    if (ctx.master && showMaster(ctx.slide) && (!ctx.layout || showMaster(ctx.layout))) await this.shapes(tree(ctx.master), ctx, ctx.master, id, el, { skipPlaceholders: true });
    if (ctx.layout && showMaster(ctx.slide)) await this.shapes(tree(ctx.layout), ctx, ctx.layout, id, el, { skipPlaceholders: true });
    await this.shapes(tree(ctx.slide), ctx, ctx.slide, id, el);
    return { el, hidden };
  }
}

// Text area inside non-rectangular shapes: [left, top, right, bottom] in px.
function textRect(prst, w, h) {
  const a = Math.min(w, h);
  switch (prst) {
    case 'ellipse': return [w * 0.146, h * 0.146, w * 0.146, h * 0.146];
    case 'chevron': return [a * 0.5, 0, a * 0.5, 0];
    case 'homePlate': case 'pentagon': return [0, 0, a * 0.5, 0];
    case 'diamond': case 'flowChartDecision': return [w * 0.25, h * 0.25, w * 0.25, h * 0.25];
    case 'triangle': return [w * 0.25, h * 0.5, w * 0.25, 0];
    case 'hexagon': case 'octagon': return [a * 0.25, a * 0.1, a * 0.25, a * 0.1];
    case 'parallelogram': case 'trapezoid': return [a * 0.25, 0, a * 0.25, 0];
    case 'rightArrow': case 'leftArrow': return [0, h * 0.25, 0, h * 0.25];
    default: return null;
  }
}

function autoNum(n, scheme = 'arabicPeriod') {
  const roman = (k) => { const m = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]; let s = ''; for (const [v, r] of m) while (k >= v) { s += r; k -= v; } return s; };
  const alpha = (k) => { let s = ''; while (k > 0) { s = String.fromCharCode(97 + ((k - 1) % 26)) + s; k = Math.floor((k - 1) / 26); } return s; };
  let core = String(n);
  if (/^alphaLc/.test(scheme)) core = alpha(n);
  else if (/^alphaUc/.test(scheme)) core = alpha(n).toUpperCase();
  else if (/^romanLc/.test(scheme)) core = roman(n);
  else if (/^romanUc/.test(scheme)) core = roman(n).toUpperCase();
  if (/ParenBoth$/.test(scheme)) return `(${core}) `;
  if (/ParenR$/.test(scheme)) return `${core}) `;
  if (/Plain$/.test(scheme)) return `${core} `;
  return `${core}. `;
}

const SVGNS = 'http://www.w3.org/2000/svg';

function lineSvg(box, ln) {
  const svg = document.createElementNS(SVGNS, 'svg');
  const w = Math.max(box.w, 1), h = Math.max(box.h, 1);
  svg.setAttribute('width', w); svg.setAttribute('height', h);
  svg.setAttribute('viewBox', `0 0 ${w} ${h}`);
  svg.style.cssText = 'position:absolute;left:0;top:0;overflow:visible;';
  const id = `ar${Math.random().toString(36).slice(2, 8)}`;
  if (ln.head || ln.tail) {
    const defs = document.createElementNS(SVGNS, 'defs');
    defs.innerHTML = `<marker id="${id}" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="4" markerHeight="4" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="${ln.color}"/></marker>`;
    svg.appendChild(defs);
  }
  const l = document.createElementNS(SVGNS, 'line');
  l.setAttribute('x1', 0); l.setAttribute('y1', 0); l.setAttribute('x2', box.w); l.setAttribute('y2', box.h);
  l.setAttribute('stroke', ln.color); l.setAttribute('stroke-width', ln.width);
  if (ln.dash) l.setAttribute('stroke-dasharray', ln.dash === 'dot' ? `${ln.width},${ln.width * 2}` : `${ln.width * 4},${ln.width * 3}`);
  if (ln.tail) l.setAttribute('marker-end', `url(#${id})`);
  if (ln.head) l.setAttribute('marker-start', `url(#${id})`);
  svg.appendChild(l);
  return svg;
}

function polygon(prst, w, h) {
  const P = (pts) => `M${pts.map((p) => p.join(' ')).join('L')}Z`;
  const a = Math.min(w, h);
  switch (prst) {
    case 'triangle': return P([[w / 2, 0], [w, h], [0, h]]);
    case 'rtTriangle': return P([[0, 0], [w, h], [0, h]]);
    case 'diamond': case 'flowChartDecision': return P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case 'parallelogram': case 'flowChartInputOutput': return P([[a * 0.25, 0], [w, 0], [w - a * 0.25, h], [0, h]]);
    case 'trapezoid': return P([[a * 0.25, 0], [w - a * 0.25, 0], [w, h], [0, h]]);
    case 'pentagon': case 'homePlate': return P([[0, 0], [w - a * 0.5, 0], [w, h / 2], [w - a * 0.5, h], [0, h]]);
    case 'chevron': return P([[0, 0], [w - a * 0.5, 0], [w, h / 2], [w - a * 0.5, h], [0, h], [a * 0.5, h / 2]]);
    case 'hexagon': return P([[a * 0.25, 0], [w - a * 0.25, 0], [w, h / 2], [w - a * 0.25, h], [a * 0.25, h], [0, h / 2]]);
    case 'octagon': { const c = a * 0.29; return P([[c, 0], [w - c, 0], [w, c], [w, h - c], [w - c, h], [c, h], [0, h - c], [0, c]]); }
    case 'rightArrow': return P([[0, h * 0.25], [w - a * 0.5, h * 0.25], [w - a * 0.5, 0], [w, h / 2], [w - a * 0.5, h], [w - a * 0.5, h * 0.75], [0, h * 0.75]]);
    case 'leftArrow': return P([[w, h * 0.25], [a * 0.5, h * 0.25], [a * 0.5, 0], [0, h / 2], [a * 0.5, h], [a * 0.5, h * 0.75], [w, h * 0.75]]);
    case 'upArrow': return P([[w * 0.25, h], [w * 0.25, a * 0.5], [0, a * 0.5], [w / 2, 0], [w, a * 0.5], [w * 0.75, a * 0.5], [w * 0.75, h]]);
    case 'downArrow': return P([[w * 0.25, 0], [w * 0.75, 0], [w * 0.75, h - a * 0.5], [w, h - a * 0.5], [w / 2, h], [0, h - a * 0.5], [w * 0.25, h - a * 0.5]]);
    case 'plus': case 'mathPlus': { const t = a * 0.25; return P([[w / 2 - t, 0], [w / 2 + t, 0], [w / 2 + t, h / 2 - t], [w, h / 2 - t], [w, h / 2 + t], [w / 2 + t, h / 2 + t], [w / 2 + t, h], [w / 2 - t, h], [w / 2 - t, h / 2 + t], [0, h / 2 + t], [0, h / 2 - t], [w / 2 - t, h / 2 - t]]); }
    case 'star5': {
      const pts = [];
      for (let i = 0; i < 10; i++) {
        const r = i % 2 ? 0.38 : 1;
        const ang = -Math.PI / 2 + i * Math.PI / 5;
        pts.push([w / 2 + Math.cos(ang) * r * w / 2, h / 2 + Math.sin(ang) * r * h / 2]);
      }
      return P(pts);
    }
    case 'flowChartTerminator': return `M${h / 2} 0H${w - h / 2}A${h / 2} ${h / 2} 0 0 1 ${w - h / 2} ${h}H${h / 2}A${h / 2} ${h / 2} 0 0 1 ${h / 2} 0Z`;
    case 'can': case 'flowChartMagneticDisk': { const r = h * 0.1; return `M0 ${r}A${w / 2} ${r} 0 0 1 ${w} ${r}V${h - r}A${w / 2} ${r} 0 0 1 0 ${h - r}Z`; }
    case 'wedgeRectCallout': case 'wedgeRoundRectCallout': return P([[0, 0], [w, 0], [w, h], [w * 0.4, h], [w * 0.2, h * 1.25], [w * 0.25, h], [0, h]]);
    default: return `M0 0H${w}V${h}H0Z`;
  }
}

function customPath(cust, w, h) {
  const out = [];
  for (const path of cust.getElementsByTagNameNS('*', 'path')) {
    const pw = Number(attr(path, 'w')) || w * EMU;
    const ph = Number(attr(path, 'h')) || h * EMU;
    const sx = w / (pw || 1), sy = h / (ph || 1);
    const pt = (el) => `${Number(attr(el, 'x')) * sx} ${Number(attr(el, 'y')) * sy}`;
    for (const c of path.children) {
      const pts = Array.from(c.getElementsByTagNameNS('*', 'pt'));
      if (c.localName === 'moveTo') out.push(`M${pt(pts[0])}`);
      else if (c.localName === 'lnTo') out.push(`L${pt(pts[0])}`);
      else if (c.localName === 'cubicBezTo') out.push(`C${pts.map(pt).join(' ')}`);
      else if (c.localName === 'quadBezTo') out.push(`Q${pts.map(pt).join(' ')}`);
      else if (c.localName === 'close') out.push('Z');
    }
  }
  return out.join('');
}

function geomSvg(prst, box, bg, ln, cust) {
  const w = Math.max(box.w, 1), h = Math.max(box.h, 1);
  const d = prst === 'cust' && cust ? customPath(cust, w, h) : polygon(prst, w, h);
  if (!d) return null;
  const svg = document.createElementNS(SVGNS, 'svg');
  svg.setAttribute('width', w); svg.setAttribute('height', h);
  svg.setAttribute('viewBox', `0 0 ${w} ${h}`);
  svg.style.cssText = 'position:absolute;left:0;top:0;overflow:visible;';
  const path = document.createElementNS(SVGNS, 'path');
  path.setAttribute('d', d);
  let fill = 'none';
  if (bg && !/gradient|url\(/.test(bg)) fill = bg;
  else if (bg) {
    // Gradients/pictures in odd shapes: use the first colour.
    const m = bg.match(/rgba?\([^)]+\)/);
    fill = m ? m[0] : 'none';
  }
  path.setAttribute('fill', fill);
  if (ln) { path.setAttribute('stroke', ln.color); path.setAttribute('stroke-width', ln.width); }
  svg.appendChild(path);
  return svg;
}
