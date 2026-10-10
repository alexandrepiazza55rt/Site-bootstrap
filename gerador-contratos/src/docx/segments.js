// Modelo de edição de um parágrafo: lista linear de "segmentos" (texto, campo, tab, quebra, bruto).
// Mantém o w:pPr intacto e preserva a formatação (w:rPr) de cada trecho.
import { NS, isW, wKid, wEl, textEl, wAttr, RPR_ORDER, setOrdered, removeKid } from './xml.js';

export const FIELD_RE = /\{\{\s*([A-Za-z0-9_]+(?:@\d+)?)\s*\}\}/g;

const SKIP_IN_P = new Set(['pPr', 'proofErr', 'del', 'moveFrom', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd']);
const UNWRAP_IN_P = new Set(['hyperlink', 'ins', 'moveTo', 'smartTag', 'customXml', 'dir', 'bdo']);

/** Lê um w:p em segmentos. */
export function parseParagraph(p) {
  const segs = [];
  walk(p, segs);
  return segs;
}

function walk(container, segs) {
  for (let c = container.firstElementChild; c; c = c.nextElementSibling) {
    if (c.namespaceURI !== NS.w) { segs.push({ k: 'raw', node: c }); continue; }
    const n = c.localName;
    if (n === 'r') parseRun(c, segs);
    else if (SKIP_IN_P.has(n)) continue;
    else if (UNWRAP_IN_P.has(n)) walk(c, segs);
    else if (n === 'sdt') { const sc = wKid(c, 'sdtContent'); if (sc) walk(sc, segs); }
    else segs.push({ k: 'raw', node: c });
  }
}

function parseRun(r, segs) {
  const rPr = wKid(r, 'rPr');
  for (let c = r.firstElementChild; c; c = c.nextElementSibling) {
    if (c.namespaceURI !== NS.w) { segs.push({ k: 'raw', node: wrapRaw(r, rPr, c), rPr }); continue; }
    switch (c.localName) {
      case 'rPr': case 'lastRenderedPageBreak': case 'softHyphen': case 'delText': break;
      case 't': pushText(segs, c.textContent, rPr); break;
      case 'tab': segs.push({ k: 'tab', rPr }); break;
      case 'br': case 'cr': segs.push({ k: 'br', rPr, type: wAttr(c, 'type') || 'textWrapping' }); break;
      case 'noBreakHyphen': pushText(segs, '‑', rPr); break;
      default: segs.push({ k: 'raw', node: wrapRaw(r, rPr, c), rPr });
    }
  }
}

function wrapRaw(r, rPr, child) {
  const run = r.ownerDocument.createElementNS(NS.w, 'w:r');
  if (rPr) run.appendChild(rPr.cloneNode(true));
  run.appendChild(child.cloneNode(true));
  return run;
}

function pushText(segs, text, rPr) {
  if (!text) return;
  FIELD_RE.lastIndex = 0;
  let last = 0;
  let m;
  while ((m = FIELD_RE.exec(text))) {
    if (m.index > last) segs.push({ k: 't', text: text.slice(last, m.index), rPr });
    segs.push({ k: 'f', id: m[1], rPr });
    last = m.index + m[0].length;
  }
  if (last < text.length) segs.push({ k: 't', text: text.slice(last), rPr });
}

const keyCache = new WeakMap();
export function rPrKey(rPr) {
  if (!rPr) return '';
  let k = keyCache.get(rPr);
  if (k === undefined) {
    k = new XMLSerializer().serializeToString(rPr).replace(/\s+xmlns:\w+="[^"]*"/g, '').replace(/\s+w:rsid\w*="[^"]*"/g, '');
    keyCache.set(rPr, k);
  }
  return k;
}

/** Reescreve o conteúdo do parágrafo a partir dos segmentos (mantém w:pPr). */
export function writeParagraph(p, segs) {
  const doc = p.ownerDocument;
  const pPr = wKid(p, 'pPr');
  while (p.firstChild) p.removeChild(p.firstChild);
  if (pPr) p.appendChild(pPr);
  let cur = null; // run de texto corrente
  let curKey = null;
  const newRun = (rPr) => {
    const r = doc.createElementNS(NS.w, 'w:r');
    if (rPr && rPr.firstElementChild) r.appendChild(doc.importNode(rPr, true));
    p.appendChild(r);
    return r;
  };
  for (const s of segs) {
    if (s.k === 't') {
      if (!s.text) continue;
      const key = rPrKey(s.rPr);
      if (cur && key === curKey && cur.lastElementChild && isW(cur.lastElementChild, 't')) {
        const t = cur.lastElementChild;
        t.textContent += s.text;
        if (/^\s|\s$|\s\s/.test(t.textContent)) t.setAttributeNS(NS.xml, 'xml:space', 'preserve');
      } else {
        cur = newRun(s.rPr); curKey = key;
        cur.appendChild(textEl(doc, s.text));
      }
    } else if (s.k === 'f') {
      const r = newRun(s.rPr);
      r.appendChild(textEl(doc, '{{' + s.id + '}}'));
      cur = null;
    } else if (s.k === 'tab' || s.k === 'br') {
      const key = rPrKey(s.rPr);
      if (!(cur && key === curKey)) { cur = newRun(s.rPr); curKey = key; }
      const el = wEl(doc, s.k === 'tab' ? 'tab' : 'br');
      if (s.k === 'br' && s.type && s.type !== 'textWrapping') el.setAttributeNS(NS.w, 'w:type', s.type);
      cur.appendChild(el);
    } else if (s.k === 'raw') {
      p.appendChild(doc.importNode(s.node, true));
      cur = null;
    }
  }
}

// ------------------------------------------------- posições lineares
export function segLen(s) {
  return s.k === 't' ? s.text.length : 1;
}

export function totalLen(segs) {
  let n = 0;
  for (const s of segs) n += segLen(s);
  return n;
}

export function linearPos(segs, segIndex, offset) {
  let n = 0;
  for (let i = 0; i < segIndex && i < segs.length; i++) n += segLen(segs[i]);
  return n + offset;
}

/** Segmentos no intervalo linear [a, b). */
export function sliceSegs(segs, a, b) {
  const out = [];
  let pos = 0;
  for (const s of segs) {
    const len = segLen(s);
    const s0 = pos, s1 = pos + len;
    pos = s1;
    if (s1 <= a || s0 >= b) { if (len === 0 && s0 >= a && s0 < b) out.push(s); continue; }
    if (s.k === 't') {
      const from = Math.max(0, a - s0), to = Math.min(len, b - s0);
      if (to > from) out.push({ ...s, text: s.text.slice(from, to) });
    } else out.push(s);
  }
  return out;
}

export function replaceRange(segs, a, b, insert) {
  const len = totalLen(segs);
  return [...sliceSegs(segs, 0, a), ...insert, ...sliceSegs(segs, b, len)];
}

export function segsText(segs, { fields = 'token' } = {}) {
  let s = '';
  for (const g of segs) {
    if (g.k === 't') s += g.text;
    else if (g.k === 'f') s += fields === 'token' ? '{{' + g.id + '}}' : fields === 'char' ? '￼' : '';
    else if (g.k === 'tab') s += '\t';
    else if (g.k === 'br') s += '\n';
    else if (fields === 'char') s += '￼';
  }
  return s;
}

/** Texto "plano" do parágrafo para buscas (campos como {{id}}). */
export function paragraphText(p) {
  return segsText(parseParagraph(p));
}

/** rPr predominante (do primeiro trecho de texto) — usado para novos textos. */
export function dominantRPr(segs, pPr) {
  for (const s of segs) if ((s.k === 't' || s.k === 'f') && (s.k === 'f' || s.text.trim())) return s.rPr || null;
  const markRPr = pPr && wKid(pPr, 'rPr');
  return markRPr || null;
}

/** Cria um rPr derivado com negrito/itálico/sublinhado ajustados. */
export function withFormatting(doc, rPr, { b, i, u }, inherited) {
  const out = rPr ? rPr.cloneNode(true) : doc.createElementNS(NS.w, 'w:rPr');
  const setToggle = (name, csName, want, inh) => {
    if (want === undefined) return;
    removeKid(out, name); removeKid(out, csName);
    if (!!want !== !!inh) {
      const el = wEl(doc, name);
      if (!want) el.setAttributeNS(NS.w, 'w:val', '0');
      setOrdered(out, el, RPR_ORDER);
      const el2 = wEl(doc, csName);
      if (!want) el2.setAttributeNS(NS.w, 'w:val', '0');
      setOrdered(out, el2, RPR_ORDER);
    }
  };
  setToggle('b', 'bCs', b, inherited?.b);
  setToggle('i', 'iCs', i, inherited?.i);
  if (u !== undefined) {
    removeKid(out, 'u');
    const inhU = !!inherited?.u;
    if (u !== inhU) {
      const el = wEl(doc, 'u');
      el.setAttributeNS(NS.w, 'w:val', u ? 'single' : 'none');
      setOrdered(out, el, RPR_ORDER);
    }
  }
  return out.firstElementChild ? out : null;
}
