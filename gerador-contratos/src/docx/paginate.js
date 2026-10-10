// Paginação: distribui os blocos renderizados em páginas, quebrando parágrafos por linha
// e tabelas por linha de tabela, com cabeçalho/rodapé repetidos em cada página.
import { layoutTabs, fixPageAnchors } from './render.js';

const PX_PER_PT = 96 / 72;

export function paginate({ blocks, section, headerHtml, firstHeaderHtml, footerHtml, firstFooterHtml, host }) {
  const s = section;
  const pageW = s.w / 20, pageH = s.h / 20, left = s.left / 20, contentW = s.contentW / 20;
  host.style.width = contentW + 'pt';
  host.innerHTML = '';

  const measurePart = (html, topPt) => {
    if (!html) return { el: null, h: 0 };
    const el = document.createElement('div');
    el.className = 'page-hf';
    el.style.width = contentW + 'pt';
    el.innerHTML = html;
    host.appendChild(el);
    layoutTabs(el);
    fixPageAnchors(el, topPt);
    const h = flowHeight(el) / PX_PER_PT;
    host.removeChild(el);
    return { el, h };
  };
  const hdr = measurePart(headerHtml, s.header / 20);
  const hdr1 = firstHeaderHtml !== undefined ? measurePart(firstHeaderHtml, s.header / 20) : hdr;
  const ftr = measurePart(footerHtml, 0);
  const ftr1 = firstFooterHtml !== undefined ? measurePart(firstFooterHtml, 0) : ftr;

  const bodyTop = Math.max(s.top / 20, s.header / 20 + Math.max(hdr.h, hdr1.h));
  const bodyBottom = Math.max(s.bottom / 20, s.footer / 20 + Math.max(ftr.h, ftr1.h));
  const contentH = pageH - bodyTop - bodyBottom;
  const limitPx = contentH * PX_PER_PT;

  for (const b of blocks) host.appendChild(b);
  layoutTabs(host);

  const heights = new WeakMap();
  const heightOf = (el) => {
    if (!heights.has(el)) heights.set(el, el.getBoundingClientRect().height);
    return heights.get(el);
  };

  const pages = [];
  let body = null;
  let used = 0;
  const newPage = () => {
    const idx = pages.length;
    const page = document.createElement('div');
    page.className = 'page';
    page.style.width = pageW + 'pt';
    page.style.height = pageH + 'pt';
    const h = idx === 0 ? hdr1 : hdr;
    const f = idx === 0 ? ftr1 : ftr;
    if (h.el) {
      const he = h.el.cloneNode(true);
      he.className = 'page-header';
      Object.assign(he.style, { top: s.header / 20 + 'pt', left: left + 'pt', width: contentW + 'pt' });
      page.appendChild(he);
    }
    body = document.createElement('div');
    body.className = 'page-body';
    Object.assign(body.style, { top: bodyTop + 'pt', left: left + 'pt', width: contentW + 'pt', height: contentH + 'pt' });
    page.appendChild(body);
    if (f.el) {
      const fe = f.el.cloneNode(true);
      fe.className = 'page-footer';
      Object.assign(fe.style, { bottom: s.footer / 20 + 'pt', left: left + 'pt', width: contentW + 'pt' });
      page.appendChild(fe);
    }
    pages.push(page);
    used = 0;
  };
  newPage();

  const queue = blocks.slice();
  let guard = 0;
  while (queue.length && guard++ < 20000) {
    const b = queue.shift();
    if (b.dataset.pbb && used > 0) newPage();
    const h = heightOf(b);
    if (used + h <= limitPx + 0.5) { body.appendChild(b); used += h; continue; }

    const avail = limitPx - used;
    let parts = null;
    if (b.classList.contains('w-p') && !b.classList.contains('keep-lines')) parts = splitParagraph(b, avail);
    else if (b.classList.contains('w-tblwrap')) parts = splitTable(b, avail);
    if (parts) {
      host.insertBefore(parts[1], b);
      host.insertBefore(parts[0], parts[1]);
      host.removeChild(b);
      heights.set(parts[0], avail);
      body.appendChild(parts[0]);
      newPage();
      queue.unshift(parts[1]);
      continue;
    }
    if (used > 0) {
      // "manter com o próximo": leva junto títulos que ficariam sozinhos no fim da página
      const pulled = [];
      while (body.lastElementChild && body.lastElementChild.classList.contains('keep-next') && body.childElementCount > 1) {
        pulled.unshift(body.lastElementChild);
        body.removeChild(body.lastElementChild);
      }
      newPage();
      for (const p of pulled) { body.appendChild(p); used += heightOf(p); }
      queue.unshift(b);
      continue;
    }
    body.appendChild(b);
    used += h;
  }

  // numeração de páginas em cabeçalhos/rodapés
  pages.forEach((page, i) => {
    for (const el of page.querySelectorAll('.page-header [data-pagefield], .page-footer [data-pagefield]')) {
      el.textContent = el.dataset.pagefield === 'numpages' ? String(pages.length) : String(i + 1);
    }
  });
  host.innerHTML = '';
  return { pages, bodyTop, contentH };
}

function flowHeight(el) {
  // altura do fluxo (ignora flutuantes absolutos)
  return el.getBoundingClientRect().height;
}

// ------------------------------------------------- quebra de parágrafo
function textNodes(root) {
  const out = [];
  const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
    acceptNode: (n) => (n.parentElement && n.parentElement.closest('.w-anchor') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
  });
  while (w.nextNode()) if (w.currentNode.length) out.push(w.currentNode);
  return out;
}

function lineBoxes(inner) {
  const r = document.createRange();
  r.selectNodeContents(inner);
  const rects = Array.from(r.getClientRects()).filter((x) => x.height > 2 && x.width > 0);
  rects.sort((a, b) => a.top - b.top);
  const lines = [];
  for (const rc of rects) {
    const mid = (rc.top + rc.bottom) / 2;
    const line = lines.find((l) => mid >= l.top && mid <= l.bottom);
    if (line) { line.top = Math.min(line.top, rc.top); line.bottom = Math.max(line.bottom, rc.bottom); }
    else lines.push({ top: rc.top, bottom: rc.bottom });
  }
  lines.sort((a, b) => a.top - b.top);
  return lines;
}

function charRect(range, node, i) {
  range.setStart(node, i);
  range.setEnd(node, i + 1);
  const rs = range.getClientRects();
  return rs.length ? rs[rs.length - 1] : range.getBoundingClientRect();
}

/** Primeira posição (nó, offset) cujo caractere começa em y >= yTop. */
function firstCharFrom(nodes, yTop) {
  const range = document.createRange();
  for (const n of nodes) {
    const last = charRect(range, n, n.length - 1);
    if (last.top < yTop) continue;
    let lo = 0, hi = n.length - 1, ans = n.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const rc = charRect(range, n, mid);
      if (rc.top >= yTop) { ans = mid; hi = mid - 1; } else lo = mid + 1;
    }
    return { node: n, offset: ans };
  }
  return null;
}

export function splitParagraph(el, availPx) {
  const inner = el.querySelector('.w-pi');
  if (!inner) return null;
  const elRect = el.getBoundingClientRect();
  const lines = lineBoxes(inner);
  if (lines.length < 2) return null;
  const limit = elRect.top + availPx;
  let k = 0;
  while (k < lines.length && lines[k].bottom <= limit + 0.5) k++;
  if (k === 0 || k >= lines.length) return null;
  if (el.dataset.widow) {
    if (k < 2) return null;
    if (lines.length - k < 2) k = lines.length - 2;
    if (k < 2) return null;
  }
  const pos = firstCharFrom(textNodes(inner), lines[k].top - 1);
  if (!pos) return null;
  // evita quebrar dentro de um campo (span atômico)
  const rA = document.createRange();
  rA.setStart(inner, 0);
  rA.setEnd(pos.node, pos.offset);
  const rB = document.createRange();
  rB.setStart(pos.node, pos.offset);
  rB.setEnd(inner, inner.childNodes.length);
  const fragA = rA.cloneContents();
  const fragB = rB.cloneContents();
  if (!fragA.textContent.trim()) return null;

  const a = el.cloneNode(false);
  const ia = inner.cloneNode(false);
  ia.appendChild(fragA);
  a.appendChild(ia);
  const b = el.cloneNode(false);
  const ib = inner.cloneNode(false);
  ib.appendChild(fragB);
  b.appendChild(ib);

  a.style.paddingBottom = '0pt';
  b.style.paddingTop = '0pt';
  ia.style.borderBottom = 'none';
  ib.style.borderTop = 'none';
  ib.style.textIndent = '0pt';
  if (inner.dataset.jc === 'both' || inner.dataset.jc === 'distribute') ia.style.textAlignLast = 'justify';
  a.dataset.split = 'a';
  b.dataset.split = 'b';
  b.removeAttribute('data-pbb');
  a.classList.remove('keep-next');

  // corrige deslocamentos (data-o) do run que continua na parte B
  const runsA = ia.querySelectorAll('.w-r[data-s]');
  const lastA = runsA[runsA.length - 1];
  const firstB = ib.querySelector('.w-r[data-s]');
  if (lastA && firstB && lastA.dataset.s === firstB.dataset.s && !lastA.dataset.f) {
    firstB.dataset.o = String((+lastA.dataset.o || 0) + lastA.textContent.length);
  }
  return [a, b];
}

// -------------------------------------------------- quebra de tabela
export function splitTable(wrap, availPx) {
  const table = wrap.querySelector('table');
  const tbody = table && table.tBodies[0];
  if (!tbody) return null;
  const rows = Array.from(tbody.rows);
  const top = wrap.getBoundingClientRect().top;
  const limit = top + availPx;
  let headerCount = 0;
  while (headerCount < rows.length && rows[headerCount].dataset.header) headerCount++;
  let k = headerCount;
  while (k < rows.length && rows[k].getBoundingClientRect().bottom <= limit + 0.5) k++;
  if (k <= headerCount || k >= rows.length) return null;

  const mk = () => {
    const w = wrap.cloneNode(false);
    const t = table.cloneNode(false);
    const cg = table.querySelector('colgroup');
    if (cg) t.appendChild(cg.cloneNode(true));
    const tb = document.createElement('tbody');
    t.appendChild(tb);
    w.appendChild(t);
    return [w, tb];
  };
  const [a, ta] = mk();
  const [b, tb] = mk();
  rows.slice(0, k).forEach((r) => ta.appendChild(r.cloneNode(true)));
  rows.slice(0, headerCount).forEach((r) => { const c = r.cloneNode(true); c.classList.add('repeat-header'); tb.appendChild(c); });
  rows.slice(k).forEach((r) => tb.appendChild(r.cloneNode(true)));
  b.dataset.split = 'b';
  a.dataset.split = 'a';
  return [a, b];
}
