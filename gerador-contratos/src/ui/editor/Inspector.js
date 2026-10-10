// Painel lateral do bloco selecionado: editar texto, tipo, alinhamento e ações (inserir, mover, excluir...).
import { html, cx, Icon, IconButton, Button, Menu } from '../kit.js';
import { ParagraphEditor } from './ParagraphEditor.js';
import { NS, isW } from '../../docx/xml.js';
import { kindOf, numInfo, alignOf, hasPageBreakBefore, isTopLevel } from '../../docx/ops.js';

export function numberLabel(ctx, doc, node) {
  const counter = ctx.numbering.createCounter();
  for (const p of doc.getElementsByTagNameNS(NS.w, 'p')) {
    const info = numInfo(ctx, p);
    const r = info.numId ? counter.next(info.numId, info.ilvl) : null;
    if (p === node) return r ? r.text : '';
  }
  return '';
}

function ancestor(node, local) {
  let n = node.parentNode;
  while (n && n.nodeType === 1) { if (isW(n, local)) return n; n = n.parentNode; }
  return null;
}

export function Inspector({
  node, ctx, doc, campos, kinds, extVersion, mode, edited, listRow, onText, onSplit, onMergeUp, onPasteLines, onAction, onClose, focusText,
}) {
  if (!node) return null;
  const isPara = isW(node, 'p');
  const kind = isPara ? kindOf(ctx, node) : { label: 'Tabela', icon: 'layers' };
  const num = isPara ? numberLabel(ctx, doc, node) : '';
  const table = isPara ? ancestor(node, 'tbl') : node;
  const inTable = isPara && !!table;
  const top = isTopLevel(node);
  const align = isPara ? alignOf(ctx, node) : null;
  const pbb = isPara && hasPageBreakBefore(node);
  const tableTop = table && isTopLevel(table) ? table : null;

  const insertItems = [
    { header: 'Inserir abaixo' },
    ...kinds.map((k) => ({ icon: k.icon, label: k.label, hint: k.sample ? 'ex.: ' + k.sample : undefined, onClick: () => onAction('insert', k.key) })),
    '-',
    { icon: 'library', label: 'Cláusula da biblioteca…', onClick: () => onAction('library-insert') },
  ];

  return html`<aside class="inspector" aria-label="Editar trecho">
    <header class="insp-head">
      <span class="insp-kind"><${Icon} name=${kind.icon || 'pilcrow'} size=${16} /></span>
      <div class="insp-titles">
        <strong>${num ? `${kind.label} ${num}` : kind.label}</strong>
        <small>${inTable ? 'Dentro de uma tabela' : mode === 'modelo' ? 'Editando o modelo' : 'Somente neste contrato'}${edited ? ' · alterado' : ''}</small>
      </div>
      <${IconButton} icon="x" title="Fechar (Esc)" onClick=${onClose} />
    </header>

    <div class="insp-body">
      ${isPara && html`<${ParagraphEditor} node=${node} ctx=${ctx} campos=${campos} version=${extVersion} autoFocus=${focusText}
        onChange=${onText} onSplit=${top || inTable ? onSplit : null} onMergeUp=${onMergeUp} onPasteLines=${onPasteLines} />`}
      ${listRow && html`<div class="insp-note"><${Icon} name="layers" size=${15} /> Esta linha se repete automaticamente para cada item da lista “${listRow}”.</div>`}

      ${isPara && html`<div class="insp-section">
        <div class="insp-label">Alinhamento</div>
        <div class="seg-ctl">
          ${[['left', 'align-left', 'À esquerda'], ['center', 'align-center', 'Centralizado'], ['right', 'align-right', 'À direita'], ['both', 'align-justify', 'Justificado']].map(([v, ic, t]) => html`
            <button type="button" class=${cx(align === v || (v === 'left' && align === 'start') ? 'is-active' : '')} title=${t} onClick=${() => onAction('align', v)}><${Icon} name=${ic} size=${16} /></button>`)}
        </div>
        ${top && html`<label class="switch-row">
          <input type="checkbox" checked=${pbb} onChange=${(e) => onAction('pagebreak', e.currentTarget.checked)} />
          <span class="switch" /><span>Começar em uma nova página</span>
        </label>`}
      </div>`}

      ${isPara && top && kinds.length > 1 && html`<div class="insp-section">
        <div class="insp-label">Tipo do parágrafo</div>
        <select class="input" value=${kind.key} onChange=${(e) => onAction('convert', e.currentTarget.value)}>
          ${kinds.map((k) => html`<option value=${k.key}>${k.label}${k.sample ? ` (${k.sample})` : ''}</option>`)}
        </select>
        <p class="insp-tip">A numeração (1.1, §1º, a...) é refeita automaticamente.</p>
      </div>`}

      <div class="insp-section">
        <div class="insp-label">${inTable ? 'Tabela' : 'Ações'}</div>
        <div class="insp-actions">
          ${(top || !inTable) && isPara && html`<${Menu} trigger=${({ toggle }) => html`<${Button} icon="plus" size="sm" variant="primary" onClick=${toggle}>Inserir abaixo</${Button}>`} items=${insertItems} />`}
          ${(top || tableTop) && html`<${Button} icon="copy" size="sm" onClick=${() => onAction('duplicate', inTable ? 'table' : undefined)}>Duplicar</${Button}>`}
          ${(top || tableTop) && html`<${Button} icon="arrow-up" size="sm" onClick=${() => onAction('move', -1)} title="Mover para cima (leva junto os subitens)">Subir</${Button}>`}
          ${(top || tableTop) && html`<${Button} icon="arrow-down" size="sm" onClick=${() => onAction('move', 1)} title="Mover para baixo (leva junto os subitens)">Descer</${Button}>`}
          ${isPara && top && html`<${Button} icon="bookmark-plus" size="sm" onClick=${() => onAction('library-save')}>Salvar na biblioteca</${Button}>`}
          ${edited && mode === 'contrato' && html`<${Button} icon="reset" size="sm" onClick=${() => onAction('restore')}>Restaurar texto do modelo</${Button}>`}
          ${(top || tableTop) && html`<${Button} icon="trash" size="sm" variant="danger-ghost" onClick=${() => onAction('delete', inTable ? 'table' : undefined)}>${inTable ? 'Excluir tabela' : 'Excluir'}</${Button}>`}
        </div>
        ${!inTable && top && kind.level >= 0 && kind.level < 99 && html`<p class="insp-tip"><${Icon} name="info" size=${13} /> Mover e excluir levam junto os parágrafos subordinados (§, alíneas).</p>`}
      </div>
    </div>
  </aside>`;
}
