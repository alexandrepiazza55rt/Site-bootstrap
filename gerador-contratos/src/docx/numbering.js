// Numeração automática (listas multinível) no mesmo comportamento do Word.
import { wKid, wKids, wAttr, wNum } from './xml.js';

export class Numbering {
  constructor(numDoc, styles) {
    this.abstracts = new Map(); // abstractNumId -> {levels[], styleLink, numStyleLink}
    this.nums = new Map(); // numId -> {abstractId, overrides: Map(ilvl -> {start, lvl})}
    this.styles = styles;
    if (!numDoc) return;
    const root = numDoc.documentElement;
    for (const an of wKids(root, 'abstractNum')) {
      const id = wAttr(an, 'abstractNumId');
      const levels = [];
      for (const l of wKids(an, 'lvl')) levels[wNum(l, 'ilvl', 0)] = readLevel(l);
      this.abstracts.set(id, {
        levels,
        numStyleLink: wAttr(wKid(an, 'numStyleLink'), 'val'),
        styleLink: wAttr(wKid(an, 'styleLink'), 'val'),
      });
    }
    for (const n of wKids(root, 'num')) {
      const overrides = new Map();
      for (const o of wKids(n, 'lvlOverride')) {
        const ilvl = wNum(o, 'ilvl', 0);
        const so = wKid(o, 'startOverride');
        const lvlEl = wKid(o, 'lvl');
        overrides.set(ilvl, { start: so ? wNum(so, 'val', null) : null, lvl: lvlEl ? readLevel(lvlEl) : null });
      }
      this.nums.set(wAttr(n, 'numId'), { abstractId: wAttr(wKid(n, 'abstractNumId'), 'val'), overrides });
    }
  }

  abstractFor(numId) {
    const num = this.nums.get(String(numId));
    if (!num) return null;
    let abs = this.abstracts.get(num.abstractId);
    // numStyleLink: aponta para um estilo de numeração que referencia outro num
    if (abs && abs.numStyleLink && this.styles) {
      const st = this.styles.byId.get(abs.numStyleLink);
      const np = st && st.pPr && wKid(st.pPr, 'numPr');
      const linked = np && wAttr(wKid(np, 'numId'), 'val');
      const ln = linked && this.nums.get(linked);
      if (ln && this.abstracts.get(ln.abstractId)) return { id: ln.abstractId, abs: this.abstracts.get(ln.abstractId), num };
    }
    return abs ? { id: num.abstractId, abs, num } : null;
  }

  level(numId, ilvl) {
    const a = this.abstractFor(numId);
    if (!a) return null;
    const ov = a.num.overrides.get(ilvl);
    return (ov && ov.lvl) || a.abs.levels[ilvl] || null;
  }

  createCounter() {
    return new NumberingCounter(this);
  }

  hasNum(numId) {
    return this.nums.has(String(numId));
  }
}

function readLevel(l) {
  return {
    start: wNum(wKid(l, 'start'), 'val', 1),
    numFmt: wAttr(wKid(l, 'numFmt'), 'val') || 'decimal',
    lvlText: wAttr(wKid(l, 'lvlText'), 'val') ?? '',
    lvlJc: wAttr(wKid(l, 'lvlJc'), 'val') || 'left',
    suff: wAttr(wKid(l, 'suff'), 'val') || 'tab',
    restart: wKid(l, 'lvlRestart') ? wNum(wKid(l, 'lvlRestart'), 'val', null) : null,
    isLgl: !!wKid(l, 'isLgl'),
    pPr: wKid(l, 'pPr'),
    rPr: wKid(l, 'rPr'),
  };
}

/** Percorre parágrafos em ordem e calcula o texto do número de cada um. */
export class NumberingCounter {
  constructor(numbering) {
    this.n = numbering;
    this.counters = new Map(); // abstractId -> number[]
    this.usedNums = new Set();
  }

  next(numId, ilvl) {
    const a = this.n.abstractFor(numId);
    if (!a) return null;
    const lvl = this.n.level(numId, ilvl);
    if (!lvl) return null;
    let c = this.counters.get(a.id);
    if (!c) { c = []; this.counters.set(a.id, c); }
    const key = String(numId);
    if (!this.usedNums.has(key)) {
      this.usedNums.add(key);
      // startOverride reinicia a lista na primeira utilização desta instância
      for (const [l, ov] of a.num.overrides) if (ov.start !== null) c[l] = ov.start - 1;
    }
    const start = lvl.start;
    c[ilvl] = (c[ilvl] === undefined ? start - 1 : c[ilvl]) + 1;
    // níveis inferiores reiniciam (respeitando lvlRestart)
    for (let k = ilvl + 1; k < 9; k++) {
      const lk = this.n.level(numId, k);
      const r = lk ? lk.restart : null;
      if (r === 0) continue;
      if (r === null || ilvl < r) c[k] = undefined;
    }
    const text = formatLevelText(lvl, (k) => {
      const lk = this.n.level(numId, k);
      const v = c[k] === undefined ? (lk ? lk.start : 1) : c[k];
      return { v, fmt: lvl.isLgl ? 'decimal' : (lk ? lk.numFmt : 'decimal') };
    });
    return { text, lvl, value: c[ilvl] };
  }

  /** Avança o contador sem produzir texto (para manter estado). */
  peek(numId, ilvl) {
    return this.next(numId, ilvl);
  }
}

export function formatLevelText(lvl, valueOf) {
  if (lvl.numFmt === 'bullet') return bulletChar(lvl.lvlText);
  if (lvl.numFmt === 'none') return '';
  return lvl.lvlText.replace(/%([1-9])/g, (_, d) => {
    const { v, fmt } = valueOf(parseInt(d, 10) - 1);
    return formatNumber(v, fmt);
  });
}

function bulletChar(t) {
  const map = { '': '•', '': '▪', '': '➢', '': '❖', '': '✓', '': '•', '': '■', '': '➔', o: '◦', '': '' };
  return t.split('').map((ch) => (ch in map ? map[ch] : ch.charCodeAt(0) >= 0xf000 ? '•' : ch)).join('') || '•';
}

export function formatNumber(n, fmt) {
  switch (fmt) {
    case 'decimalZero': return n < 10 ? '0' + n : String(n);
    case 'lowerLetter': return letters(n).toLowerCase();
    case 'upperLetter': return letters(n);
    case 'lowerRoman': return roman(n).toLowerCase();
    case 'upperRoman': return roman(n);
    case 'ordinal': return n + 'º';
    case 'none': return '';
    default: return String(n);
  }
}

function letters(n) {
  if (n <= 0) return '';
  const ch = String.fromCharCode(65 + ((n - 1) % 26));
  return ch.repeat(Math.floor((n - 1) / 26) + 1);
}

function roman(n) {
  if (n <= 0) return String(n);
  const t = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let s = '';
  for (const [v, r] of t) while (n >= v) { s += r; n -= v; }
  return s;
}
