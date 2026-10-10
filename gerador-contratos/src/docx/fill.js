// Expansão de linhas repetidas (listas) e preenchimento final dos campos para exportação.
import { NS, isW, wAll, textEl, wEl, parseXml, serializeXml, setParaId } from './xml.js';
import { TOKEN_RE } from '../fields/compute.js';

/**
 * Para cada grupo-lista, a unidade de repetição é a linha de tabela (w:tr) — ou o parágrafo —
 * que contém campos do grupo. Retorna Map(elemento -> {grupo, count}).
 */
export function findListUnits(doc, resolver) {
  const listFields = new Map(); // campoId -> grupoId
  for (const c of resolver.campos) if (resolver.isList(c)) listFields.set(c.id, c.grupo);
  const units = new Map();
  if (!listFields.size) return units;
  for (const t of wAll(doc.documentElement, 't')) {
    const txt = t.textContent;
    if (!txt.includes('{{')) continue;
    TOKEN_RE.lastIndex = 0;
    let m;
    while ((m = TOKEN_RE.exec(txt))) {
      const g = listFields.get(m[1]);
      if (!g || m[2] !== undefined) continue;
      let unit = t;
      while (unit && !isW(unit, 'tr')) unit = unit.parentNode;
      if (!unit) { unit = t; while (unit && !isW(unit, 'p')) unit = unit.parentNode; }
      if (unit && !units.has(unit)) units.set(unit, { grupo: g, count: Math.max(1, resolver.itemsOf(g).length) });
    }
  }
  units.listFields = listFields;
  return units;
}

/** Clona cada unidade repetida uma vez por item, trocando {{campo}} por {{campo@i}}. */
export function expandLists(doc, resolver) {
  const units = findListUnits(doc, resolver);
  for (const [unit, { grupo: g, count: n }] of units) {
    const parent = unit.parentNode;
    for (let i = 0; i < n; i++) {
      const clone = unit.cloneNode(true);
      for (const t of wAll(clone, 't')) {
        t.textContent = t.textContent.replace(TOKEN_RE, (all, id, idx) => (idx === undefined && units.listFields.get(id) === g ? `{{${id}@${i}}}` : all));
      }
      if (i > 0) for (const p of (isW(clone, 'p') ? [clone] : wAll(clone, 'p'))) setParaId(p);
      parent.insertBefore(clone, unit);
    }
    parent.removeChild(unit);
  }
  return doc;
}

/** Substitui {{campos}} pelos valores finais (quebras de linha viram w:br). */
export function fillTokens(doc, resolver) {
  for (const t of wAll(doc.documentElement, 't')) {
    const txt = t.textContent;
    if (!txt.includes('{{')) continue;
    TOKEN_RE.lastIndex = 0;
    const out = txt.replace(TOKEN_RE, (all, id, idx) => {
      const r = resolver.get(id, idx === undefined ? undefined : +idx);
      return r.unknown ? all : r.text || '';
    });
    if (!out.includes('\n')) {
      t.textContent = out;
      t.setAttributeNS(NS.xml, 'xml:space', 'preserve');
      continue;
    }
    const run = t.parentNode;
    const lines = out.split('\n');
    lines.forEach((line, i) => {
      if (i > 0) run.insertBefore(wEl(doc, 'br'), t);
      const nt = textEl(doc, line);
      nt.setAttributeNS(NS.xml, 'xml:space', 'preserve');
      run.insertBefore(nt, t);
    });
    run.removeChild(t);
  }
  return doc;
}

/** document.xml final (string) a partir do XML de trabalho. */
export function buildFinalXml(xml, resolver) {
  const doc = parseXml(xml);
  expandLists(doc, resolver);
  fillTokens(doc, resolver);
  removeEmptyRuns(doc);
  return serializeXml(doc);
}

function removeEmptyRuns(doc) {
  for (const t of wAll(doc.documentElement, 't')) {
    if (t.textContent === '' && t.parentNode && isW(t.parentNode, 'r')) {
      const r = t.parentNode;
      r.removeChild(t);
      if (!Array.from(r.children).some((c) => !isW(c, 'rPr'))) r.parentNode.removeChild(r);
    }
  }
}

/** Atualiza docProps/core.xml (título, autor, datas). */
export function coreProps(original, { titulo, autor }) {
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const esc = (s) => String(s || '').replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  let xml = typeof original === 'string' && original.includes('coreProperties') ? original : null;
  if (!xml) {
    xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>';
  }
  const set = (tag, value, attrs = '') => {
    const re = new RegExp(`<${tag}[^>]*>[\\s\\S]*?</${tag}>|<${tag}[^>]*/>`);
    const el = `<${tag}${attrs}>${esc(value)}</${tag}>`;
    xml = re.test(xml) ? xml.replace(re, el) : xml.replace('</cp:coreProperties>', el + '</cp:coreProperties>');
  };
  if (titulo) set('dc:title', titulo);
  if (autor) set('cp:lastModifiedBy', autor);
  set('dcterms:modified', now, ' xsi:type="dcterms:W3CDTF"');
  return xml;
}
