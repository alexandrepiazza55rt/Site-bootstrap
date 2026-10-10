// Estrutura do documento: seções e cláusulas numeradas, com busca e navegação.
import { useMemo, useState } from 'preact/hooks';
import { html, cx, Icon } from '../kit.js';
import { NS, isW } from '../../docx/xml.js';
import { numInfo, kindOf, blockSummary } from '../../docx/ops.js';
import { parseParagraph, segsText } from '../../docx/segments.js';

export function Outline({ ctx, doc, version, selected, marks, onPick, campos }) {
  const [q, setQ] = useState('');
  const items = useMemo(() => {
    const out = [];
    const counter = ctx.numbering.createCounter();
    const body = doc.getElementsByTagNameNS(NS.w, 'body')[0];
    for (let n = body.firstElementChild; n; n = n.nextElementSibling) {
      if (isW(n, 'tbl')) { out.push({ node: n, label: 'Tabela', num: '', level: 2, kind: 'tbl', text: '' }); continue; }
      if (!isW(n, 'p')) continue;
      const info = numInfo(ctx, n);
      const r = info.numId ? counter.next(info.numId, info.ilvl) : null;
      const text = segsText(parseParagraph(n)).replace(/\{\{(\w+)(@\d+)?\}\}/g, (_, id) => '[' + ((campos.find((c) => c.id === id) || {}).rotulo || id) + ']');
      if (!text.trim()) continue;
      const k = kindOf(ctx, n);
      out.push({ node: n, num: r ? r.text : '', text: text.replace(/\s+/g, ' ').trim(), level: k.level, kind: k.label });
    }
    return out;
  }, [ctx, doc, version, campos]);

  const term = q.trim().toLowerCase();
  const norm = (s) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
  const filtered = term ? items.filter((it) => norm(it.text).includes(norm(term)) || it.num.includes(term)) : items;
  const hits = term ? filtered.length : 0;

  return html`<div class="outline">
    <div class="outline-search">
      <${Icon} name="search" size=${15} />
      <input class="input input-plain" placeholder="Buscar no texto do contrato…" value=${q} onInput=${(e) => setQ(e.currentTarget.value)} />
      ${term && html`<span class="muted small">${hits} ${hits === 1 ? 'trecho' : 'trechos'}</span>`}
    </div>
    <ol class="outline-list">
      ${filtered.map((it, i) => {
        const mk = marks && marks.get(it.node);
        const isHead = it.level === 0 || it.level === -1;
        const snippet = term ? highlight(it.text, term) : it.text.length > 110 ? it.text.slice(0, 108) + '…' : it.text;
        return html`<li key=${i} class=${cx('ol-item', 'lv' + Math.max(0, Math.min(it.level, 4)), isHead && 'is-head', it.level === 99 && 'is-text', selected === it.node && 'is-active')}
            onClick=${() => onPick(it.node)} title=${it.kind}>
          ${it.num && html`<span class="ol-num">${it.num}</span>`}
          <span class="ol-text">${snippet}</span>
          ${mk && html`<span class=${'ol-mark ' + mk} title=${mk === 'new' ? 'Trecho novo neste contrato' : 'Texto alterado em relação ao modelo'} />`}
        </li>`;
      })}
    </ol>
  </div>`;
}

function highlight(text, term) {
  const norm = (s) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
  const i = norm(text).indexOf(norm(term));
  if (i < 0) return text.slice(0, 110);
  const start = Math.max(0, i - 40);
  const pre = (start > 0 ? '…' : '') + text.slice(start, i);
  const mid = text.slice(i, i + term.length);
  const post = text.slice(i + term.length, i + term.length + 60) + (i + term.length + 60 < text.length ? '…' : '');
  return html`${pre}<mark>${mid}</mark>${post}`;
}

export { blockSummary };
