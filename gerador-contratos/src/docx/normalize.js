// Limpeza de documentos importados: aceita alterações controladas, remove marcas de revisão,
// junta runs fragmentados, conserta {{campos}} quebrados e garante w14:paraId únicos.
import { NS, isW, wKid, wAll, setParaId, getParaId, randomParaId } from './xml.js';
import { rPrKey } from './segments.js';

const RSID_ATTRS = ['rsidR', 'rsidRPr', 'rsidRDefault', 'rsidP', 'rsidDel', 'rsidSect', 'rsidTr'];

export function normalizeDocument(doc) {
  const body = doc.getElementsByTagNameNS(NS.w, 'body')[0];
  if (!body) throw new Error('Documento sem corpo (w:body).');
  ensureW14(doc);

  // 1) alterações controladas
  for (const name of ['del', 'moveFrom']) for (const el of wAll(body, name)) el.parentNode && el.parentNode.removeChild(el);
  for (const name of ['ins', 'moveTo']) for (const el of wAll(body, name)) unwrap(el);
  for (const name of ['rPrChange', 'pPrChange', 'sectPrChange', 'tblPrChange', 'trPrChange', 'tcPrChange', 'tblGridChange', 'numberingChange',
    'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd', 'proofErr', 'lastRenderedPageBreak']) {
    for (const el of wAll(body, name)) el.parentNode && el.parentNode.removeChild(el);
  }
  // 2) controles de conteúdo, smart tags
  for (const el of wAll(body, 'sdt')) {
    const content = wKid(el, 'sdtContent');
    if (content) { while (content.firstChild) el.parentNode.insertBefore(content.firstChild, el); }
    el.parentNode.removeChild(el);
  }
  for (const name of ['smartTag', 'customXml']) for (const el of wAll(body, name)) unwrap(el);

  // 3) atributos rsid
  for (const el of body.getElementsByTagName('*')) {
    for (const a of RSID_ATTRS) if (el.hasAttributeNS(NS.w, a)) el.removeAttributeNS(NS.w, a);
  }

  // 4) runs e campos
  const seen = new Set();
  for (const p of wAll(body, 'p')) {
    mergeRuns(p);
    repairTokens(p);
    let id = getParaId(p);
    if (!id || seen.has(id)) { id = randomParaId(); setParaId(p, id); }
    seen.add(id);
  }
  return doc;
}

export function ensureW14(doc) {
  const root = doc.documentElement;
  if (!root.getAttribute('xmlns:w14')) root.setAttributeNS('http://www.w3.org/2000/xmlns/', 'xmlns:w14', NS.w14);
  if (!root.getAttribute('xmlns:mc')) root.setAttributeNS('http://www.w3.org/2000/xmlns/', 'xmlns:mc', NS.mc);
  const ign = root.getAttributeNS(NS.mc, 'Ignorable') || '';
  if (!ign.split(/\s+/).includes('w14')) root.setAttributeNS(NS.mc, 'mc:Ignorable', (ign + ' w14').trim());
}

function unwrap(el) {
  const parent = el.parentNode;
  if (!parent) return;
  while (el.firstChild) parent.insertBefore(el.firstChild, el);
  parent.removeChild(el);
}

const SIMPLE = new Set(['t', 'tab', 'br', 'rPr']);

function isSimpleRun(r) {
  for (let c = r.firstElementChild; c; c = c.nextElementSibling) if (!isW(c) || !SIMPLE.has(c.localName)) return false;
  return true;
}

/** Junta runs adjacentes com a mesma formatação (em qualquer nível: p, hyperlink...). */
export function mergeRuns(container) {
  let prev = null;
  for (let c = container.firstElementChild; c;) {
    const next = c.nextElementSibling;
    if (isW(c, 'r') && isSimpleRun(c)) {
      if (prev && rPrKey(wKid(prev, 'rPr')) === rPrKey(wKid(c, 'rPr'))) {
        for (let k = c.firstElementChild; k;) {
          const nk = k.nextElementSibling;
          if (!isW(k, 'rPr')) {
            const last = prev.lastElementChild;
            if (isW(k, 't') && last && isW(last, 't')) {
              last.textContent += k.textContent;
              last.setAttributeNS(NS.xml, 'xml:space', 'preserve');
            } else prev.appendChild(k);
          }
          k = nk;
        }
        container.removeChild(c);
      } else prev = c;
    } else if (isW(c, 'hyperlink') || isW(c, 'fldSimple')) {
      mergeRuns(c); prev = null;
    } else if (isW(c, 'bookmarkStart') || isW(c, 'bookmarkEnd')) {
      // marcadores não quebram a junção de runs
    } else prev = null;
    c = next;
  }
}

/** {{campo}} dividido entre runs diferentes -> move o token inteiro para o primeiro run. */
export function repairTokens(p) {
  const ts = wAll(p, 't');
  if (ts.length < 2) return;
  let full = '';
  const offs = [];
  for (const t of ts) { offs.push(full.length); full += t.textContent; }
  if (!full.includes('{{')) return;
  const re = /\{\{[^{}]{1,80}?\}\}/g;
  let m;
  const edits = [];
  while ((m = re.exec(full))) edits.push([m.index, m.index + m[0].length, m[0]]);
  for (let e = edits.length - 1; e >= 0; e--) {
    const [ms, me, token] = edits[e];
    let i = ts.length - 1;
    while (i > 0 && offs[i] > ms) i--;
    let j = i;
    while (j < ts.length - 1 && offs[j + 1] < me) j++;
    if (i === j) continue;
    const ti = ts[i], tj = ts[j];
    const tail = tj.textContent.slice(me - offs[j]);
    ti.textContent = ti.textContent.slice(0, ms - offs[i]) + token;
    ti.setAttributeNS(NS.xml, 'xml:space', 'preserve');
    for (let k = i + 1; k < j; k++) ts[k].textContent = '';
    tj.textContent = tail;
    tj.setAttributeNS(NS.xml, 'xml:space', 'preserve');
  }
  for (const t of ts) {
    if (t.textContent === '' && t.parentNode) {
      const r = t.parentNode;
      r.removeChild(t);
      if (isW(r, 'r') && !Array.from(r.children).some((c) => !isW(c, 'rPr'))) r.parentNode && r.parentNode.removeChild(r);
    }
  }
}

/** Converte {{Nome do Campo}} livres em ids normalizados; retorna os campos encontrados. */
export function collectTokens(doc) {
  const found = new Map();
  for (const t of wAll(doc.documentElement, 't')) {
    const txt = t.textContent;
    if (!txt.includes('{{')) continue;
    const out = txt.replace(/\{\{\s*([^{}]{1,80}?)\s*\}\}/g, (_, raw) => {
      const id = slugId(raw);
      if (!found.has(id)) found.set(id, raw.trim());
      return '{{' + id + '}}';
    });
    if (out !== txt) t.textContent = out;
  }
  return found; // id -> rótulo original
}

export function slugId(s) {
  const base = String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase()
    .replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 40);
  return /^[a-z]/.test(base) ? base : 'campo_' + (base || Math.random().toString(36).slice(2, 7));
}
