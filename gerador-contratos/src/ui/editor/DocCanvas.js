// Área do documento: páginas reais (A4/Carta...) renderizadas a partir do .docx, com seleção e campos clicáveis.
import { useEffect, useLayoutEffect, useRef, useState } from 'preact/hooks';
import { html, cx } from '../kit.js';
import { Renderer, sectionProps, fixPageAnchors } from '../../docx/render.js';
import { paginate } from '../../docx/paginate.js';
import { parseParagraph, linearPos } from '../../docx/segments.js';

export function buildPages({ ctx, doc, resolve, listUnits, marks, mode = 'screen' }) {
  const r = new Renderer(ctx, { mode, resolve, listUnits, marks });
  const sec = sectionProps(doc);
  const blocks = r.renderBody(doc, sec);
  const part = (id) => (id ? ctx.part(id) : null);
  const hdr = part(sec.headers.default);
  const ftr = part(sec.footers.default);
  const hdr1 = sec.titlePg ? part(sec.headers.first) : hdr;
  const ftr1 = sec.titlePg ? part(sec.footers.first) : ftr;
  const hh = (p) => (p ? r.renderPart(p, sec) : '');
  let host = document.getElementById('measure-host');
  if (!host) {
    host = document.createElement('div');
    host.id = 'measure-host';
    host.className = 'measure-host';
    document.body.appendChild(host);
  }
  const res = paginate({
    blocks, section: sec, host,
    headerHtml: hh(hdr), firstHeaderHtml: sec.titlePg ? hh(hdr1) : undefined,
    footerHtml: hh(ftr), firstFooterHtml: sec.titlePg ? hh(ftr1) : undefined,
  });
  return { ...res, nodes: r.nodes, section: sec };
}

export function DocCanvas({
  ctx, doc, version, resolve, listUnits, marks, highlight, selected, focusField, zoom, templateMode, showChanges,
  onSelect, onFieldClick, onTextSelect, onPages, onZoomComputed,
}) {
  const scrollRef = useRef();
  const pagesRef = useRef();
  const nodesRef = useRef([]);
  const [autoZoom, setAutoZoom] = useState(1);
  const pageWRef = useRef(595);

  // renderização + paginação
  useLayoutEffect(() => {
    if (!ctx || !doc || !pagesRef.current) return;
    let out;
    try {
      out = buildPages({ ctx, doc, resolve, listUnits, marks });
    } catch (e) {
      console.error(e);
      pagesRef.current.innerHTML = `<div class="render-error">Não foi possível desenhar o documento: ${String(e.message || e)}</div>`;
      return;
    }
    nodesRef.current = out.nodes;
    pageWRef.current = out.section.w / 20;
    const el = pagesRef.current;
    el.replaceChildren(...out.pages);
    out.pages.forEach((p, i) => {
      const n = document.createElement('div');
      n.className = 'page-num';
      n.textContent = `${i + 1} / ${out.pages.length}`;
      p.appendChild(n);
      const body = p.querySelector('.page-body');
      if (body) fixPageAnchors(body, out.bodyTop);
    });
    applySelection();
    applyFocus(false);
    onPages && onPages(out.pages.length);
    computeAuto();
  }, [ctx, doc, version, resolve, listUnits, marks]);

  const revealRef = useRef(false);
  const applySelection = () => {
    const el = pagesRef.current;
    if (!el) return;
    for (const s of el.querySelectorAll('.is-selected')) s.classList.remove('is-selected');
    if (!selected) return;
    const nodes = nodesRef.current;
    let first = null;
    for (const b of el.querySelectorAll('[data-xi]')) {
      if (nodes[+b.dataset.xi] === selected && (b.classList.contains('w-p') || b.classList.contains('w-tblwrap'))) { b.classList.add('is-selected'); first = first || b; }
    }
    if (first && revealRef.current && scrollRef.current) {
      revealRef.current = false;
      const r = first.getBoundingClientRect();
      const c = scrollRef.current.getBoundingClientRect();
      if (r.bottom < c.top + 20 || r.top > c.bottom - 40) first.scrollIntoView({ block: 'center', behavior: 'smooth' });
    }
  };
  useEffect(() => { revealRef.current = true; applySelection(); }, [selected]);

  const applyFocus = (scroll) => {
    const el = pagesRef.current;
    if (!el) return;
    for (const s of el.querySelectorAll('.w-fld.is-focus')) s.classList.remove('is-focus');
    if (!focusField) return;
    const hits = Array.from(el.querySelectorAll(`.w-fld[data-f="${CSS.escape(focusField.id)}"]`)).filter((f) => {
      if (focusField.idx === undefined) return true;
      const it = f.closest('[data-item]');
      return !it || +it.dataset.item === focusField.idx;
    });
    hits.forEach((h) => h.classList.add('is-focus'));
    if (scroll && hits[0] && scrollRef.current) {
      const r = hits[0].getBoundingClientRect();
      const c = scrollRef.current.getBoundingClientRect();
      if (r.top < c.top + 40 || r.bottom > c.bottom - 40) hits[0].scrollIntoView({ block: 'center', behavior: 'smooth' });
    }
  };
  useEffect(() => applyFocus(true), [focusField]);

  // zoom automático (ajusta à largura)
  const computeAuto = () => {
    const sc = scrollRef.current;
    if (!sc) return;
    const avail = sc.clientWidth - 56;
    const pagePx = pageWRef.current * (96 / 72);
    const z = Math.max(0.45, Math.min(1.25, avail / pagePx));
    setAutoZoom(Math.round(z * 100) / 100);
  };
  useEffect(() => {
    const sc = scrollRef.current;
    if (!sc || typeof ResizeObserver === 'undefined') return undefined;
    const ro = new ResizeObserver(() => computeAuto());
    ro.observe(sc);
    return () => ro.disconnect();
  }, []);
  const z = zoom === 'auto' ? autoZoom : zoom;
  useEffect(() => { onZoomComputed && onZoomComputed(z); }, [z]);

  const onClick = (e) => {
    const sel = window.getSelection();
    if (sel && !sel.isCollapsed && sel.toString().trim()) return;
    const f = e.target.closest('.w-fld');
    if (f && onFieldClick) {
      const it = f.closest('[data-item]');
      onFieldClick(f.dataset.f, it ? +it.dataset.item : undefined, f);
    }
    const b = e.target.closest('.w-p, .w-tblwrap');
    if (b) onSelect(nodesRef.current[+b.dataset.xi], { el: b, dbl: e.detail >= 2, field: !!f });
    else if (e.target.closest('.page')) onSelect(null);
  };

  const onMouseUp = () => {
    if (!onTextSelect) return;
    setTimeout(() => {
      const sel = window.getSelection();
      if (!sel || sel.isCollapsed || !sel.rangeCount) { onTextSelect(null); return; }
      const info = selectionInfo(sel, nodesRef.current);
      if (!info) { onTextSelect(null); return; }
      onTextSelect(info);
    }, 0);
  };

  return html`<div class=${cx('canvas-scroll', templateMode && 'tpl')} ref=${scrollRef} onClick=${onClick} onMouseUp=${onMouseUp}>
    <div class=${cx('canvas-pages', 'doc-interactive', highlight && 'hl', showChanges && 'show-changes')} ref=${pagesRef} style=${{ zoom: z }} />
  </div>`;
}

// ------------------------------------------------- seleção de texto -> posição no parágrafo
function boundaryPos(container, offset, isEnd) {
  // devolve {seg, off} (off em caracteres dentro do segmento) a partir de um ponto do DOM
  let node = container;
  if (node.nodeType === 3) {
    const span = node.parentElement && node.parentElement.closest('[data-s]');
    if (!span) return null;
    if (span.classList.contains('w-fld')) return { seg: +span.dataset.s, off: isEnd ? 1 : 0, field: true };
    return { seg: +span.dataset.s, off: (+span.dataset.o || 0) + offset };
  }
  // elemento: olha o filho no offset
  const child = node.childNodes[offset];
  if (child) {
    const el = child.nodeType === 1 ? (child.matches('[data-s]') ? child : child.querySelector('[data-s]')) : child.parentElement.closest('[data-s]');
    if (el) return { seg: +el.dataset.s, off: el.classList.contains('w-fld') ? 0 : +el.dataset.o || 0 };
  }
  const all = node.nodeType === 1 ? node.querySelectorAll('[data-s]') : [];
  const last = all[all.length - 1] || (node.closest && node.closest('[data-s]'));
  if (last) return { seg: +last.dataset.s, off: last.classList.contains('w-fld') ? 1 : (+last.dataset.o || 0) + last.textContent.length };
  return null;
}

export function selectionInfo(sel, nodes) {
  const range = sel.getRangeAt(0);
  const el = (n) => (n.nodeType === 3 ? n.parentElement : n);
  const pA = el(range.startContainer).closest('.w-p');
  const pB = el(range.endContainer).closest('.w-p');
  if (!pA || !pB) return null;
  const node = nodes[+pA.dataset.xi];
  if (!node || node !== nodes[+pB.dataset.xi]) return { error: 'Selecione um trecho dentro de um único parágrafo.' };
  const a = boundaryPos(range.startContainer, range.startOffset, false);
  const b = boundaryPos(range.endContainer, range.endOffset, true);
  if (!a || !b) return null;
  const segs = parseParagraph(node);
  let start = linearPos(segs, a.seg, a.off);
  let end = linearPos(segs, b.seg, b.off);
  if (end < start) [start, end] = [end, start];
  // ignora espaços nas bordas
  const text = sel.toString();
  const lead = text.length - text.trimStart().length;
  const trail = text.length - text.trimEnd().length;
  if (end - start > lead + trail) { start += lead; end -= trail; }
  if (end <= start) return null;
  const rect = range.getBoundingClientRect();
  return { node, start, end, text: text.trim(), rect, item: pA.closest('[data-item]') ? +pA.closest('[data-item]').dataset.item : undefined };
}
