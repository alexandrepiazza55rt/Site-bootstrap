// Editor de texto rico de um parágrafo: mantém a formatação original de cada trecho (w:rPr),
// permite negrito/itálico/sublinhado, campos {{...}} como "chips", Enter divide o parágrafo.
import { useEffect, useRef, useState } from 'preact/hooks';
import { html, cx, Icon, IconButton, Menu } from '../kit.js';
import { parseParagraph, dominantRPr, withFormatting } from '../../docx/segments.js';
import { runBaseProps, runProps, paraBaseProps } from '../../docx/styles.js';
import { wKid } from '../../docx/xml.js';
import { escapeHtml as esc } from '../../docx/xml.js';

/** Converte segmentos em HTML editável. */
function segsToHtml(segs, ctx, p, label) {
  const pPr = wKid(p, 'pPr');
  const pp = paraBaseProps(ctx.styles, pPr, null);
  const base = runBaseProps(ctx.styles, pp.styleId, null);
  let out = '';
  segs.forEach((s, i) => {
    const rp = s.rPr !== undefined ? runProps(ctx.styles, base, s.rPr) : base;
    const st = `font-weight:${rp.b ? 700 : 400};font-style:${rp.i ? 'italic' : 'normal'};text-decoration:${rp.u ? 'underline' : 'none'}`;
    if (s.k === 't') out += `<span data-s="${i}" style="${st}">${esc(s.text)}</span>`;
    else if (s.k === 'f') out += `<span class="chip" contenteditable="false" data-f="${esc(s.id)}" data-s="${i}" style="${st}">${esc(label(s.id))}</span>`;
    else if (s.k === 'tab') out += `<span class="chip chip-tab" contenteditable="false" data-s="${i}" title="Tabulação">⇥</span>`;
    else if (s.k === 'br') out += `<br data-s="${i}">`;
    else if (s.k === 'raw') {
      const name = s.node.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'drawing').length ? 'imagem' : 'objeto';
      if (/bookmark|proofErr|commentRange|permStart|permEnd/.test(s.node.localName || '')) out += `<span class="chip chip-hidden" contenteditable="false" data-s="${i}"></span>`;
      else out += `<span class="chip chip-raw" contenteditable="false" data-s="${i}">◆ ${name}</span>`;
    }
  });
  return out || '';
}

/** Lê o DOM do editor de volta para segmentos (preservando rPr de origem). */
export function htmlToSegs(root, origSegs, ctx, p) {
  const doc = p.ownerDocument;
  const pPr = wKid(p, 'pPr');
  const pp = paraBaseProps(ctx.styles, pPr, null);
  const inherited = runBaseProps(ctx.styles, pp.styleId, null);
  const defRPr = dominantRPr(origSegs, pPr);
  const out = [];
  const fmtOf = (el, fmt) => {
    const f = { ...fmt };
    const tag = el.tagName;
    if (tag === 'B' || tag === 'STRONG') f.b = true;
    if (tag === 'I' || tag === 'EM') f.i = true;
    if (tag === 'U') f.u = true;
    const st = el.style;
    if (st) {
      if (st.fontWeight) f.b = st.fontWeight === 'bold' || +st.fontWeight >= 600;
      if (st.fontStyle) f.i = st.fontStyle === 'italic';
      const td = st.textDecorationLine || st.textDecoration;
      if (td) f.u = /underline/.test(td);
    }
    return f;
  };
  const effOf = (rPr) => runProps(ctx.styles, inherited, rPr);
  const walk = (node, fmt, baseSeg) => {
    for (let ch = node.firstChild; ch; ch = ch.nextSibling) {
      if (ch.nodeType === 3) {
        const text = ch.nodeValue.replace(/​/g, '').replace(/ /g, ' ');
        if (!text) continue;
        const rPr = baseSeg ? baseSeg.rPr : defRPr;
        const eff = effOf(rPr);
        const want = { b: fmt.b ?? eff.b, i: fmt.i ?? eff.i, u: fmt.u ?? !!eff.u };
        const same = want.b === eff.b && want.i === eff.i && want.u === !!eff.u;
        out.push({ k: 't', text, rPr: same ? rPr : withFormatting(doc, rPr, want, inherited) });
        continue;
      }
      if (ch.nodeType !== 1) continue;
      const el = ch;
      const sIdx = el.dataset && el.dataset.s !== undefined ? +el.dataset.s : null;
      const seg = sIdx !== null ? origSegs[sIdx] : null;
      if (el.dataset && el.dataset.f) {
        const rPr = seg ? seg.rPr : baseSeg ? baseSeg.rPr : defRPr;
        const f2 = fmtOf(el, fmt);
        const eff = effOf(rPr);
        const want = { b: f2.b ?? eff.b, i: f2.i ?? eff.i, u: f2.u ?? !!eff.u };
        const same = want.b === eff.b && want.i === eff.i && want.u === !!eff.u;
        out.push({ k: 'f', id: el.dataset.f, rPr: same ? rPr : withFormatting(doc, rPr, want, inherited) });
        continue;
      }
      if (el.classList.contains('chip-tab')) { out.push({ k: 'tab', rPr: seg ? seg.rPr : baseSeg ? baseSeg.rPr : defRPr }); continue; }
      if (el.classList.contains('chip-raw') || el.classList.contains('chip-hidden')) { if (seg) out.push(seg); continue; }
      if (el.tagName === 'BR') {
        if (!el.nextSibling && el.parentNode === root) continue; // <br> final artificial do contenteditable
        out.push({ k: 'br', rPr: seg ? seg.rPr : baseSeg ? baseSeg.rPr : defRPr });
        continue;
      }
      const isBlock = el.tagName === 'DIV' || el.tagName === 'P';
      if (isBlock && out.length) out.push({ k: 't', text: ' ', rPr: baseSeg ? baseSeg.rPr : defRPr });
      walk(el, fmtOf(el, fmt), seg && seg.k === 't' ? seg : baseSeg);
    }
  };
  walk(root, {}, null);
  return out;
}

export function ParagraphEditor({ node, ctx, campos, version, onChange, onSplit, onMergeUp, onPasteLines, autoFocus, placeholder }) {
  const ref = useRef();
  const segsRef = useRef([]);
  const [fmt, setFmt] = useState({});
  const label = (id) => {
    const base = id.replace(/@\d+$/, '');
    const c = campos.find((x) => x.id === base);
    return c ? c.rotulo : id;
  };

  // reconstrói quando muda o parágrafo (ou após desfazer)
  useEffect(() => {
    const segs = parseParagraph(node);
    segsRef.current = segs;
    if (ref.current) {
      ref.current.innerHTML = segsToHtml(segs, ctx, node, label);
      if (autoFocus) {
        ref.current.focus();
        placeCaret(ref.current, autoFocus === 'start');
      }
    }
  }, [node, version]);

  const emit = () => {
    const segs = htmlToSegs(ref.current, segsRef.current, ctx, node);
    onChange(segs);
  };

  const exec = (cmd) => {
    ref.current.focus();
    document.execCommand('styleWithCSS', false, true);
    document.execCommand(cmd);
    emit();
    updateFmt();
  };

  const updateFmt = () => {
    try {
      setFmt({ b: document.queryCommandState('bold'), i: document.queryCommandState('italic'), u: document.queryCommandState('underline') });
    } catch { /* ignore */ }
  };

  const insertField = (c) => {
    const el = ref.current;
    el.focus();
    const sel = window.getSelection();
    let range = sel.rangeCount ? sel.getRangeAt(0) : null;
    if (!range || !el.contains(range.startContainer)) { range = document.createRange(); range.selectNodeContents(el); range.collapse(false); }
    range.deleteContents();
    const chip = document.createElement('span');
    chip.className = 'chip';
    chip.contentEditable = 'false';
    chip.dataset.f = c.id;
    chip.textContent = c.rotulo;
    range.insertNode(chip);
    const space = document.createTextNode('​');
    chip.after(space);
    range.setStartAfter(space);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    emit();
  };

  const onKeyDown = (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      if (!onSplit) return;
      const el = ref.current;
      const sel = window.getSelection();
      if (!sel.rangeCount) return;
      const r = sel.getRangeAt(0);
      r.deleteContents();
      const tail = document.createRange();
      tail.setStart(r.startContainer, r.startOffset);
      tail.setEnd(el, el.childNodes.length);
      const frag = tail.extractContents();
      const tmp = document.createElement('div');
      tmp.appendChild(frag);
      const a = htmlToSegs(el, segsRef.current, ctx, node);
      const b = htmlToSegs(tmp, segsRef.current, ctx, node);
      onSplit(a, b);
      return;
    }
    if (e.key === 'Enter' && e.shiftKey) {
      e.preventDefault();
      document.execCommand('insertLineBreak');
      emit();
      return;
    }
    if (e.key === 'Backspace' && onMergeUp) {
      const sel = window.getSelection();
      if (sel.isCollapsed && atStart(ref.current, sel)) {
        e.preventDefault();
        onMergeUp(htmlToSegs(ref.current, segsRef.current, ctx, node));
      }
    }
    if ((e.ctrlKey || e.metaKey) && ['b', 'i', 'u'].includes(e.key.toLowerCase())) {
      e.preventDefault();
      exec({ b: 'bold', i: 'italic', u: 'underline' }[e.key.toLowerCase()]);
    }
  };

  const onPaste = (e) => {
    e.preventDefault();
    const text = (e.clipboardData || window.clipboardData).getData('text/plain') || '';
    const lines = text.replace(/\r/g, '').split(/\n+/).map((l) => l.replace(/\t/g, ' ')).filter((l) => l.trim());
    if (!lines.length) return;
    document.execCommand('insertText', false, lines[0]);
    emit();
    if (lines.length > 1 && onPasteLines) onPasteLines(lines.slice(1));
  };

  return html`<div class="pe">
    <div class="pe-toolbar">
      <${IconButton} icon="bold" title="Negrito (Ctrl+B)" active=${fmt.b} size=${16} onMouseDown=${(e) => { e.preventDefault(); exec('bold'); }} />
      <${IconButton} icon="italic" title="Itálico (Ctrl+I)" active=${fmt.i} size=${16} onMouseDown=${(e) => { e.preventDefault(); exec('italic'); }} />
      <${IconButton} icon="underline" title="Sublinhado (Ctrl+U)" active=${fmt.u} size=${16} onMouseDown=${(e) => { e.preventDefault(); exec('underline'); }} />
      <span class="pe-sep" />
      <${Menu} trigger=${({ toggle }) => html`<button type="button" class="pe-insert" onMouseDown=${(e) => e.preventDefault()} onClick=${toggle}><${Icon} name="braces" size=${15} /> Inserir campo <${Icon} name="chevron-down" size=${13} /></button>`}
        items=${fieldMenu(campos, insertField)} />
    </div>
    <div class="pe-area" ref=${ref} contenteditable="true" spellcheck="true" lang="pt-BR" data-placeholder=${placeholder || 'Digite o texto…'}
      onInput=${emit} onKeyDown=${onKeyDown} onKeyUp=${updateFmt} onMouseUp=${updateFmt} onPaste=${onPaste} />
    <div class="pe-hint"><kbd>Enter</kbd> novo parágrafo · <kbd>Shift</kbd>+<kbd>Enter</kbd> quebra de linha · <kbd>Backspace</kbd> no início junta com o anterior</div>
  </div>`;
}

function fieldMenu(campos, onPick) {
  const items = [];
  const byGroup = new Map();
  for (const c of campos) {
    if (!byGroup.has(c.grupo)) byGroup.set(c.grupo, []);
    byGroup.get(c.grupo).push(c);
  }
  if (!campos.length) return [{ label: 'Nenhum campo criado no modelo', disabled: true }];
  for (const [, list] of byGroup) {
    items.push('-');
    for (const c of list) items.push({ label: c.rotulo, hint: '{{' + c.id + '}}', onClick: () => onPick(c) });
  }
  return items.slice(1);
}

function placeCaret(el, start) {
  const r = document.createRange();
  r.selectNodeContents(el);
  r.collapse(!!start);
  const s = window.getSelection();
  s.removeAllRanges();
  s.addRange(r);
}

function atStart(root, sel) {
  const r = sel.getRangeAt(0);
  const pre = document.createRange();
  pre.selectNodeContents(root);
  pre.setEnd(r.startContainer, r.startOffset);
  return pre.toString().replace(/​/g, '') === '' && !pre.cloneContents().querySelector('.chip');
}

export { cx };
