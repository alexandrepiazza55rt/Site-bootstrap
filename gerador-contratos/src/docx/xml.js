// Utilitários para manipular WordprocessingML com DOMParser/XMLSerializer.

export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  v: 'urn:schemas-microsoft-com:vml',
  xml: 'http://www.w3.org/XML/1998/namespace',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
};

export function parseXml(str) {
  const doc = new DOMParser().parseFromString(str, 'application/xml');
  const err = doc.getElementsByTagName('parsererror')[0];
  if (err) throw new Error('XML inválido: ' + err.textContent.slice(0, 200));
  return doc;
}

export function serializeXml(doc) {
  let s = new XMLSerializer().serializeToString(doc);
  if (!s.startsWith('<?xml')) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
  return s;
}

/** true se o elemento é w:<local> */
export function isW(el, local) {
  return !!el && el.nodeType === 1 && el.namespaceURI === NS.w && (local === undefined || el.localName === local);
}

/** Filhos-elemento w:<local> (ou todos os filhos-elemento se local omitido). */
export function wKids(el, local) {
  const out = [];
  if (!el) return out;
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
    if (local === undefined ? c.nodeType === 1 : isW(c, local)) out.push(c);
  }
  return out;
}

export function wKid(el, local) {
  if (!el) return null;
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (isW(c, local)) return c;
  return null;
}

/** Descendentes w:<local> em ordem de documento. */
export function wAll(el, local) {
  return el ? Array.from(el.getElementsByTagNameNS(NS.w, local)) : [];
}

export function wAttr(el, name) {
  if (!el) return null;
  const v = el.getAttributeNS(NS.w, name);
  if (v !== null && v !== '') return v;
  // alguns geradores omitem o prefixo
  const raw = el.getAttribute('w:' + name);
  return raw === '' ? null : raw;
}

export function setWAttr(el, name, value) {
  if (value === null || value === undefined) el.removeAttributeNS(NS.w, name);
  else el.setAttributeNS(NS.w, 'w:' + name, String(value));
}

export function wNum(el, name, def = null) {
  const v = wAttr(el, name);
  if (v === null) return def;
  const n = parseFloat(v);
  return Number.isFinite(n) ? n : def;
}

/** Valor booleano de propriedades como <w:b/>, <w:b w:val="0"/>. */
export function onOff(el) {
  if (!el) return null;
  const v = wAttr(el, 'val');
  if (v === null) return true;
  return !(v === '0' || v === 'false' || v === 'off' || v === 'none');
}

export function wEl(doc, local, attrs) {
  const el = doc.createElementNS(NS.w, 'w:' + local);
  if (attrs) for (const k in attrs) if (attrs[k] !== undefined && attrs[k] !== null) setWAttr(el, k, attrs[k]);
  return el;
}

export function textEl(doc, text) {
  const t = wEl(doc, 't');
  if (/^\s|\s$|\s\s/.test(text)) t.setAttributeNS(NS.xml, 'xml:space', 'preserve');
  t.textContent = text;
  return t;
}

/** Ordem dos filhos de w:rPr exigida pelo schema. */
export const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
  'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
  'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl',
  'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];

export const PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
  'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
  'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
  'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
  'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];

/** Insere/substitui o filho w:<local> respeitando a ordem do schema. */
export function setOrdered(parent, child, order) {
  const local = child.localName;
  const existing = wKid(parent, local);
  if (existing) { parent.replaceChild(child, existing); return child; }
  const idx = order.indexOf(local);
  for (let c = parent.firstElementChild; c; c = c.nextElementSibling) {
    const ci = order.indexOf(c.localName);
    if (ci > idx) { parent.insertBefore(child, c); return child; }
  }
  parent.appendChild(child);
  return child;
}

export function removeKid(parent, local) {
  const k = wKid(parent, local);
  if (k) parent.removeChild(k);
}

export function ensureKid(parent, local, order) {
  return wKid(parent, local) || setOrdered(parent, wEl(parent.ownerDocument, local), order);
}

export function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

export function randomParaId() {
  // 8 dígitos hex, valor < 0x80000000 (exigência do Word)
  const n = 1 + Math.floor(Math.random() * 0x7ffffffe);
  return n.toString(16).toUpperCase().padStart(8, '0');
}

export function getParaId(p) {
  return p.getAttributeNS(NS.w14, 'paraId') || null;
}

export function setParaId(p, id) {
  p.setAttributeNS(NS.w14, 'w14:paraId', id || randomParaId());
}
