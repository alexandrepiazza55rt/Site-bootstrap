// Operações estruturais no documento: inserir, duplicar, mover, excluir e converter cláusulas.
import { NS, isW, wKid, wKids, wAttr, wEl, setOrdered, PPR_ORDER, removeKid, setParaId, getParaId } from './xml.js';
import { paraBaseProps } from './styles.js';
import { parseParagraph, segsText, dominantRPr } from './segments.js';

export function bodyOf(doc) {
  return doc.getElementsByTagNameNS(NS.w, 'body')[0];
}

/** Blocos de nível superior (parágrafos e tabelas) em ordem. */
export function topBlocks(doc) {
  return wKids(bodyOf(doc)).filter((el) => isW(el, 'p') || isW(el, 'tbl'));
}

export function isTopLevel(node) {
  return node && node.parentNode && isW(node.parentNode, 'body');
}

/** Número/nível de um parágrafo (direto ou herdado do estilo). */
export function numInfo(ctx, p) {
  const pPr = wKid(p, 'pPr');
  const pp = paraBaseProps(ctx.styles, pPr, null);
  let numId = pp.numId, ilvl = pp.ilvl;
  const np = pPr && wKid(pPr, 'numPr');
  if (np) {
    const d = wAttr(wKid(np, 'numId'), 'val');
    const l = wAttr(wKid(np, 'ilvl'), 'val');
    if (d !== null) numId = d;
    if (l !== null) ilvl = parseInt(l, 10) || 0;
  }
  if (!numId || numId === '0' || !ctx.numbering.hasNum(numId)) return { numId: null, ilvl: 0, styleId: pp.styleId, heading: ctx.styles.isHeadingStyle(pp.styleId) || pp.outlineLvl !== null };
  const a = ctx.numbering.abstractFor(numId);
  return { numId, ilvl, abstractId: a ? a.id : null, lvl: ctx.numbering.level(numId, ilvl), styleId: pp.styleId, heading: ctx.styles.isHeadingStyle(pp.styleId) };
}

/** Classificação amigável do tipo de parágrafo. */
export function kindOf(ctx, p) {
  const n = numInfo(ctx, p);
  if (!n.numId) {
    if (n.heading) return { key: 's:' + n.styleId, label: 'Título', icon: 'heading', level: -1 };
    return { key: 's:' + (n.styleId || ''), label: 'Texto', icon: 'pilcrow', level: 99 };
  }
  const lt = (n.lvl && n.lvl.lvlText) || '';
  const fmt = (n.lvl && n.lvl.numFmt) || 'decimal';
  let label, icon;
  if (lt.includes('§')) { label = 'Parágrafo (§)'; icon = 'section'; }
  else if (fmt === 'bullet') { label = 'Marcador'; icon = 'list'; }
  else if (/letter/i.test(fmt)) { label = 'Alínea (a, b, c)'; icon = 'list-letters'; }
  else if (/roman/i.test(fmt)) { label = 'Inciso (I, II, III)'; icon = 'list-ordered'; }
  else if (n.ilvl === 0 && (n.heading || /^%1\.?$/.test(lt))) { label = 'Título de seção'; icon = 'heading'; }
  else if ((lt.match(/%/g) || []).length >= 2) { label = 'Cláusula / subitem'; icon = 'list-ordered'; }
  else { label = 'Item numerado'; icon = 'list-ordered'; }
  return { key: `n:${n.abstractId}:${n.ilvl}`, label, icon, level: n.ilvl, numId: n.numId, ilvl: n.ilvl };
}

/** Tipos de parágrafo usados no documento (para o menu "Inserir"). */
export function documentKinds(ctx, doc) {
  const seen = new Map();
  for (const p of doc.getElementsByTagNameNS(NS.w, 'p')) {
    if (!isTopLevel(p)) continue;
    const k = kindOf(ctx, p);
    if (!seen.has(k.key)) seen.set(k.key, { ...k, count: 0 });
    seen.get(k.key).count++;
  }
  return Array.from(seen.values()).sort((a, b) => a.level - b.level);
}

/** Último parágrafo do tipo `key` antes de `ref` (ou o primeiro do documento). */
export function prototypeFor(ctx, doc, key, ref) {
  let best = null;
  for (const p of doc.getElementsByTagNameNS(NS.w, 'p')) {
    if (!isTopLevel(p)) continue;
    if (kindOf(ctx, p).key === key) {
      if (!best) best = p;
      if (ref && (ref.compareDocumentPosition(p) & Node.DOCUMENT_POSITION_PRECEDING)) best = p;
      else if (ref) break;
    }
  }
  return best;
}

/** Novo parágrafo vazio com a formatação de `proto` (como o Enter do Word). */
export function newParagraphLike(proto) {
  const doc = proto.ownerDocument;
  const p = doc.createElementNS(NS.w, 'w:p');
  const pPr = wKid(proto, 'pPr');
  const segs = parseParagraph(proto);
  const rPr = dominantRPr(segs, pPr);
  if (pPr) {
    const np = pPr.cloneNode(true);
    removeKid(np, 'sectPr');
    removeKid(np, 'pageBreakBefore');
    if (rPr) {
      const markR = rPr.cloneNode(true);
      // a marca de parágrafo usa w:rPr dentro de w:pPr
      setOrdered(np, markR, PPR_ORDER);
    }
    p.appendChild(np);
  } else if (rPr) {
    const np = wEl(doc, 'pPr');
    np.appendChild(rPr.cloneNode(true));
    p.appendChild(np);
  }
  setParaId(p);
  return p;
}

export function insertAfter(ref, node) {
  ref.parentNode.insertBefore(node, ref.nextSibling);
  return node;
}

export function duplicate(node) {
  const c = node.cloneNode(true);
  const ps = isW(c, 'p') ? [c] : Array.from(c.getElementsByTagNameNS(NS.w, 'p'));
  for (const p of ps) setParaId(p);
  insertAfter(node, c);
  return c;
}

/** Remove um bloco (garante que células de tabela não fiquem vazias). */
export function removeBlock(node) {
  const parent = node.parentNode;
  if (!parent) return;
  if (isW(node, 'p') && isW(parent, 'tc') && wKids(parent, 'p').length + wKids(parent, 'tbl').length <= 1) {
    for (const c of Array.from(node.childNodes)) if (!isW(c, 'pPr')) node.removeChild(c);
    return;
  }
  // não remover o último parágrafo do corpo antes do sectPr
  if (isW(parent, 'body') && topBlocks(node.ownerDocument).length <= 1) {
    for (const c of Array.from(node.childNodes)) if (!isW(c, 'pPr')) node.removeChild(c);
    return;
  }
  parent.removeChild(node);
  // uma tabela não pode ser seguida diretamente pelo sectPr no Word: garante parágrafo
  const body = bodyOf(node.ownerDocument);
  const last = body && body.lastElementChild;
  const prev = last && isW(last, 'sectPr') ? last.previousElementSibling : last;
  if (prev && isW(prev, 'tbl')) body.insertBefore(wEl(node.ownerDocument, 'p'), prev.nextSibling);
}

/** Hierarquia: -1 título sem número, 0..8 nível da numeração, 99 texto/tabela. */
export function rank(ctx, node) {
  if (!isW(node, 'p')) return 99;
  return kindOf(ctx, node).level;
}

/** Bloco + seus subordinados (ex.: seção com suas cláusulas; cláusula com seus § e alíneas). */
export function sectionRange(ctx, node) {
  const r = rank(ctx, node);
  if (r === 99 || r === -1) return [node];
  const out = [node];
  let lastNumbered = 0;
  let peer = false;
  for (let n = node.nextElementSibling; n && !isW(n, 'sectPr'); n = n.nextElementSibling) {
    const rn = rank(ctx, n);
    if (rn <= r) { peer = true; break; }
    out.push(n);
    if (rn !== 99) lastNumbered = out.length - 1;
  }
  if (!peer) out.length = lastNumbered + 1;
  return out;
}

/** Move um bloco (com seus subordinados) para cima/baixo. */
export function moveBlock(ctx, node, dir) {
  const group = sectionRange(ctx, node);
  const r = rank(ctx, node);
  const parent = node.parentNode;
  if (dir < 0) {
    let y = group[0].previousElementSibling;
    if (r !== 99) while (y && rank(ctx, y) > r) y = y.previousElementSibling;
    if (!y) return false;
    for (const g of group) parent.insertBefore(g, y);
    return true;
  }
  const n = group[group.length - 1].nextElementSibling;
  if (!n || isW(n, 'sectPr')) return false;
  let after = n;
  if (r !== 99 && rank(ctx, n) === r) { const ng = sectionRange(ctx, n); after = ng[ng.length - 1]; }
  const ref = after.nextSibling;
  for (const g of group) parent.insertBefore(g, ref);
  return true;
}

/** Converte o tipo do parágrafo (copia w:pPr do protótipo, mantendo o texto). */
export function convertKind(p, proto) {
  const old = wKid(p, 'pPr');
  const markRPr = old && wKid(old, 'rPr');
  const np = wKid(proto, 'pPr') ? wKid(proto, 'pPr').cloneNode(true) : wEl(p.ownerDocument, 'pPr');
  removeKid(np, 'sectPr');
  if (markRPr) setOrdered(np, markRPr.cloneNode(true), PPR_ORDER);
  if (old) p.replaceChild(np, old);
  else p.insertBefore(np, p.firstChild);
}

/** Alinhamento do parágrafo. */
export function setAlign(p, jc) {
  let pPr = wKid(p, 'pPr');
  if (!pPr) { pPr = wEl(p.ownerDocument, 'pPr'); p.insertBefore(pPr, p.firstChild); }
  const el = wEl(p.ownerDocument, 'jc');
  el.setAttributeNS(NS.w, 'w:val', jc);
  setOrdered(pPr, el, PPR_ORDER);
}

export function setPageBreakBefore(p, on) {
  let pPr = wKid(p, 'pPr');
  if (!pPr) { pPr = wEl(p.ownerDocument, 'pPr'); p.insertBefore(pPr, p.firstChild); }
  removeKid(pPr, 'pageBreakBefore');
  if (on) setOrdered(pPr, wEl(p.ownerDocument, 'pageBreakBefore'), PPR_ORDER);
}

export function hasPageBreakBefore(p) {
  const pPr = wKid(p, 'pPr');
  const el = pPr && wKid(pPr, 'pageBreakBefore');
  return !!el && wAttr(el, 'val') !== '0';
}

export function alignOf(ctx, p) {
  const pp = paraBaseProps(ctx.styles, wKid(p, 'pPr'), null);
  const pPr = wKid(p, 'pPr');
  const jc = pPr && wKid(pPr, 'jc');
  return jc ? wAttr(jc, 'val') : pp.jc;
}

/** Texto curto para listas/estrutura. */
export function blockSummary(node, max = 90) {
  if (isW(node, 'tbl')) return 'Tabela';
  const t = segsText(parseParagraph(node), { fields: 'token' }).replace(/\s+/g, ' ').trim();
  return t.length > max ? t.slice(0, max - 1) + '…' : t;
}

/** Mapa de paraId -> texto (para comparar contrato x modelo). */
export function paraTextMap(doc) {
  const m = new Map();
  for (const p of doc.getElementsByTagNameNS(NS.w, 'p')) {
    const id = getParaId(p);
    if (id) m.set(id, segsText(parseParagraph(p)));
  }
  return m;
}

/** Marcas de alteração em relação ao modelo. */
export function diffMarks(doc, templateMap) {
  const marks = new Map();
  let removed = 0;
  const present = new Set();
  for (const p of doc.getElementsByTagNameNS(NS.w, 'p')) {
    const id = getParaId(p);
    if (!id || !templateMap.has(id)) { if (segsText(parseParagraph(p)).trim()) marks.set(p, 'new'); continue; }
    present.add(id);
    if (templateMap.get(id) !== segsText(parseParagraph(p))) marks.set(p, 'edited');
  }
  for (const [id, txt] of templateMap) if (!present.has(id) && txt.trim()) removed++;
  return { marks, removed };
}
