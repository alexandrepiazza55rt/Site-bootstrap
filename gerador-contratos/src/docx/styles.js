// Resolução de estilos: docDefaults -> estilo de parágrafo -> numeração -> formatação direta.
import { NS, wKid, wKids, wAttr, wNum, onOff, isW } from './xml.js';

export class Styles {
  constructor(stylesDoc, themeDoc) {
    this.byId = new Map();
    this.defaultPara = null;
    this.defaultChar = null;
    this.defaultTable = null;
    this.docRPr = null;
    this.docPPr = null;
    this.theme = { minor: 'Calibri', major: 'Calibri Light' };
    if (themeDoc) {
      const minor = themeDoc.getElementsByTagNameNS(NS.a, 'minorFont')[0];
      const major = themeDoc.getElementsByTagNameNS(NS.a, 'majorFont')[0];
      const latin = (el) => el && el.getElementsByTagNameNS(NS.a, 'latin')[0]?.getAttribute('typeface');
      this.theme.minor = latin(minor) || this.theme.minor;
      this.theme.major = latin(major) || this.theme.major;
    }
    if (!stylesDoc) return;
    const root = stylesDoc.documentElement;
    const dd = wKid(root, 'docDefaults');
    if (dd) {
      this.docRPr = wKid(wKid(dd, 'rPrDefault'), 'rPr');
      this.docPPr = wKid(wKid(dd, 'pPrDefault'), 'pPr');
    }
    for (const s of wKids(root, 'style')) {
      const id = wAttr(s, 'styleId');
      const type = wAttr(s, 'type');
      const dv = wAttr(s, 'default');
      const def = dv === '1' || dv === 'true' || dv === 'on';
      const info = {
        id, type,
        name: wAttr(wKid(s, 'name'), 'val') || id,
        basedOn: wAttr(wKid(s, 'basedOn'), 'val'),
        pPr: wKid(s, 'pPr'),
        rPr: wKid(s, 'rPr'),
        tblPr: wKid(s, 'tblPr'),
        tcPr: wKid(s, 'tcPr'),
        el: s,
      };
      this.byId.set(id, info);
      if (def) {
        if (type === 'paragraph') this.defaultPara = id;
        else if (type === 'character') this.defaultChar = id;
        else if (type === 'table') this.defaultTable = id;
      }
    }
  }

  /** Cadeia de estilos da raiz até o estilo pedido. */
  chain(id) {
    const out = [];
    const seen = new Set();
    let cur = id;
    while (cur && this.byId.has(cur) && !seen.has(cur)) {
      seen.add(cur);
      const s = this.byId.get(cur);
      out.unshift(s);
      cur = s.basedOn;
    }
    return out;
  }

  paraStyleId(pPr) {
    return wAttr(wKid(pPr, 'pStyle'), 'val') || this.defaultPara;
  }

  isHeadingStyle(id) {
    for (const s of this.chain(id)) {
      if (/^heading \d/i.test(s.name) || /^t[ií]tulo ?\d/i.test(s.name)) return true;
      if (s.pPr && wKid(s.pPr, 'outlineLvl')) return true;
    }
    return false;
  }

  fontName(rFonts) {
    if (!rFonts) return null;
    const direct = wAttr(rFonts, 'ascii') || wAttr(rFonts, 'hAnsi');
    if (direct) return direct;
    const theme = wAttr(rFonts, 'asciiTheme') || wAttr(rFonts, 'hAnsiTheme');
    if (theme) return /major/i.test(theme) ? this.theme.major : this.theme.minor;
    return wAttr(rFonts, 'cs') || null;
  }
}

// --------------------------------------------------------- parágrafo
export function defaultParaProps() {
  return {
    jc: 'left', indLeft: 0, indRight: 0, firstLine: 0,
    before: 0, after: 0, line: 240, lineRule: 'auto', beforeAuto: false, afterAuto: false,
    borders: {}, shd: null, tabs: [], keepNext: false, keepLines: false, pageBreakBefore: false,
    widowControl: true, contextualSpacing: false, numId: null, ilvl: 0, outlineLvl: null, styleId: null,
  };
}

export function applyPPr(pp, pPr) {
  if (!pPr) return pp;
  for (let c = pPr.firstElementChild; c; c = c.nextElementSibling) {
    if (!isW(c)) continue;
    switch (c.localName) {
      case 'jc': {
        const v = wAttr(c, 'val');
        pp.jc = v === 'start' ? 'left' : v === 'end' ? 'right' : v;
        break;
      }
      case 'ind': {
        const l = wNum(c, 'left', wNum(c, 'start', null));
        const r = wNum(c, 'right', wNum(c, 'end', null));
        if (l !== null) pp.indLeft = l;
        if (r !== null) pp.indRight = r;
        const h = wNum(c, 'hanging', null);
        const f = wNum(c, 'firstLine', null);
        if (h !== null) pp.firstLine = -h;
        else if (f !== null) pp.firstLine = f;
        break;
      }
      case 'spacing': {
        const b = wNum(c, 'before', null); if (b !== null) pp.before = b;
        const a = wNum(c, 'after', null); if (a !== null) pp.after = a;
        const ln = wNum(c, 'line', null); if (ln !== null) pp.line = ln;
        const lr = wAttr(c, 'lineRule'); if (lr) pp.lineRule = lr;
        const ba = wAttr(c, 'beforeAutospacing'); if (ba !== null) pp.beforeAuto = ba !== '0' && ba !== 'false';
        const aa = wAttr(c, 'afterAutospacing'); if (aa !== null) pp.afterAuto = aa !== '0' && aa !== 'false';
        break;
      }
      case 'pBdr': {
        pp.borders = { ...pp.borders };
        for (const side of wKids(c)) pp.borders[side.localName] = readBorder(side);
        break;
      }
      case 'shd': pp.shd = readShd(c); break;
      case 'tabs': {
        let tabs = pp.tabs.slice();
        for (const t of wKids(c, 'tab')) {
          const pos = wNum(t, 'pos', 0);
          const val = wAttr(t, 'val') || 'left';
          tabs = tabs.filter((x) => x.pos !== pos);
          if (val !== 'clear') tabs.push({ pos, val, leader: wAttr(t, 'leader') || 'none' });
        }
        pp.tabs = tabs.sort((a, b) => a.pos - b.pos);
        break;
      }
      case 'keepNext': pp.keepNext = onOff(c); break;
      case 'keepLines': pp.keepLines = onOff(c); break;
      case 'pageBreakBefore': pp.pageBreakBefore = onOff(c); break;
      case 'widowControl': pp.widowControl = onOff(c); break;
      case 'contextualSpacing': pp.contextualSpacing = onOff(c); break;
      case 'outlineLvl': pp.outlineLvl = wNum(c, 'val', null); break;
      case 'numPr': {
        const id = wAttr(wKid(c, 'numId'), 'val');
        const lvl = wAttr(wKid(c, 'ilvl'), 'val');
        if (id !== null) pp.numId = id;
        if (lvl !== null) pp.ilvl = parseInt(lvl, 10) || 0;
        break;
      }
      default: break;
    }
  }
  return pp;
}

export function readBorder(el) {
  const val = wAttr(el, 'val');
  if (!val || val === 'nil' || val === 'none') return null;
  return {
    val,
    sz: wNum(el, 'sz', 4), // oitavos de ponto
    space: wNum(el, 'space', 0),
    color: normColor(wAttr(el, 'color')),
  };
}

export function readShd(el) {
  const fill = wAttr(el, 'fill');
  if (!fill || fill === 'auto') return null;
  return '#' + fill;
}

export function normColor(c) {
  if (!c || c === 'auto') return '#000000';
  return /^[0-9a-f]{6}$/i.test(c) ? '#' + c : c;
}

// ------------------------------------------------------------- run
export function defaultRunProps() {
  return {
    font: 'Times New Roman', sz: 20, b: false, i: false, u: null, strike: false, dstrike: false,
    color: null, highlight: null, shd: null, caps: false, smallCaps: false, vertAlign: null,
    spacing: 0, vanish: false,
  };
}

export function applyRPr(rp, rPr, styles) {
  if (!rPr) return rp;
  for (let c = rPr.firstElementChild; c; c = c.nextElementSibling) {
    if (!isW(c)) continue;
    switch (c.localName) {
      case 'rFonts': { const f = styles.fontName(c); if (f) rp.font = f; break; }
      case 'sz': rp.sz = wNum(c, 'val', rp.sz); break;
      case 'b': rp.b = onOff(c); break;
      case 'i': rp.i = onOff(c); break;
      case 'u': { const v = wAttr(c, 'val'); rp.u = !v || v === 'none' ? (v ? null : 'single') : v; break; }
      case 'strike': rp.strike = onOff(c); break;
      case 'dstrike': rp.dstrike = onOff(c); break;
      case 'color': { const v = wAttr(c, 'val'); rp.color = v && v !== 'auto' ? '#' + v : null; break; }
      case 'highlight': { const v = wAttr(c, 'val'); rp.highlight = v && v !== 'none' ? v : null; break; }
      case 'shd': rp.shd = readShd(c); break;
      case 'caps': rp.caps = onOff(c); break;
      case 'smallCaps': rp.smallCaps = onOff(c); break;
      case 'vertAlign': { const v = wAttr(c, 'val'); rp.vertAlign = v === 'baseline' ? null : v; break; }
      case 'spacing': rp.spacing = wNum(c, 'val', 0); break;
      case 'vanish': rp.vanish = onOff(c); break;
      default: break;
    }
  }
  return rp;
}

/** Propriedades efetivas do parágrafo (sem numeração). */
export function paraBaseProps(styles, pPr, tableStyleId) {
  const pp = defaultParaProps();
  applyPPr(pp, styles.docPPr);
  if (tableStyleId) for (const s of styles.chain(tableStyleId)) applyPPr(pp, s.pPr);
  const sid = styles.paraStyleId(pPr);
  pp.styleId = sid;
  for (const s of styles.chain(sid)) applyPPr(pp, s.pPr);
  return pp;
}

/** Propriedades de run herdadas do estilo do parágrafo (antes da formatação direta). */
export function runBaseProps(styles, paraStyleId, tableStyleId) {
  const rp = defaultRunProps();
  applyRPr(rp, styles.docRPr, styles);
  if (tableStyleId) for (const s of styles.chain(tableStyleId)) applyRPr(rp, s.rPr, styles);
  for (const s of styles.chain(paraStyleId)) applyRPr(rp, s.rPr, styles);
  return rp;
}

export function runProps(styles, base, rPr) {
  const rp = { ...base };
  const rs = wAttr(wKid(rPr, 'rStyle'), 'val');
  if (rs) for (const s of styles.chain(rs)) applyRPr(rp, s.rPr, styles);
  applyRPr(rp, rPr, styles);
  return rp;
}

// ------------------------------------------------------------ fontes
const FONT_STACKS = {
  'courier new': "'Courier New', Courier, 'Liberation Mono', 'Nimbus Mono PS', 'Cousine', monospace",
  courier: "'Courier New', Courier, 'Liberation Mono', monospace",
  arial: "Arial, 'Liberation Sans', Helvetica, 'Arimo', sans-serif",
  helvetica: "Helvetica, Arial, 'Liberation Sans', sans-serif",
  'times new roman': "'Times New Roman', 'Liberation Serif', Times, 'Tinos', serif",
  times: "'Times New Roman', Times, 'Liberation Serif', serif",
  calibri: "Calibri, Carlito, 'Segoe UI', 'Helvetica Neue', Arial, sans-serif",
  'calibri light': "'Calibri Light', Calibri, Carlito, 'Segoe UI Light', sans-serif",
  cambria: "Cambria, Caladea, Georgia, serif",
  georgia: 'Georgia, serif',
  verdana: 'Verdana, Geneva, sans-serif',
  tahoma: 'Tahoma, Verdana, sans-serif',
  'segoe ui': "'Segoe UI', system-ui, sans-serif",
  garamond: "Garamond, 'EB Garamond', 'Times New Roman', serif",
  'book antiqua': "'Book Antiqua', Palatino, 'Palatino Linotype', serif",
  'century gothic': "'Century Gothic', 'URW Gothic', Futura, sans-serif",
  'trebuchet ms': "'Trebuchet MS', sans-serif",
  'bookman old style': "'Bookman Old Style', 'URW Bookman', serif",
  aptos: "Aptos, Calibri, Carlito, 'Segoe UI', sans-serif",
};

const LINE_FACTOR = {
  'courier new': 1.1328, courier: 1.1328, arial: 1.149, helvetica: 1.149, 'times new roman': 1.149, times: 1.149,
  calibri: 1.2207, 'calibri light': 1.2207, cambria: 1.1719, georgia: 1.1362, verdana: 1.2153, tahoma: 1.2069,
  'segoe ui': 1.33, garamond: 1.12, 'book antiqua': 1.17, 'century gothic': 1.2266, 'trebuchet ms': 1.1611,
  'bookman old style': 1.17, aptos: 1.2,
};

export function cssFont(name) {
  const k = (name || '').toLowerCase();
  if (FONT_STACKS[k]) return FONT_STACKS[k];
  const quoted = /\s/.test(name) ? `'${name}'` : name;
  return `${quoted}, 'Times New Roman', serif`;
}

export function lineFactor(name) {
  return LINE_FACTOR[(name || '').toLowerCase()] || 1.17;
}

export function isMono(name) {
  return /courier|mono|consolas/i.test(name || '');
}
