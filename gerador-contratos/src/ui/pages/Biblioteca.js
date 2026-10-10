// Biblioteca de cláusulas reutilizáveis.
import { useState } from 'preact/hooks';
import { html, Icon, IconButton, Empty, Badge, toast, confirmDialog, promptDialog, relTime } from '../kit.js';
import { listClausulas, saveClausula, deleteClausula } from '../../store/repo.js';
import { useCollection } from './Dashboard.js';
import { norm } from '../dialogs.js';

export function Biblioteca() {
  const list = useCollection(listClausulas);
  const [q, setQ] = useState('');
  const [open, setOpen] = useState(null);
  if (!list) return html`<div class="page-center"><span class="spinner lg" /></div>`;
  const shown = list.filter((c) => !q || norm(c.titulo + ' ' + c.texto + ' ' + (c.categoria || '')).includes(norm(q)));
  const pretty = (t) => t.replace(/\{\{(\w+)(@\d+)?\}\}/g, '[$1]');
  return html`<div class="page">
    <header class="page-head">
      <div><h1>Biblioteca de cláusulas</h1><p class="muted">Cláusulas prontas para inserir em qualquer contrato, com a formatação original.</p></div>
    </header>
    ${!list.length ? html`<${Empty} icon="library" title="Sua biblioteca está vazia">
        Em um contrato, clique numa cláusula e escolha <b>Salvar na biblioteca</b> (ex.: “Seguro obrigatório”, “Multa por atraso diferenciada”).
        Depois, use <b>Inserir abaixo › Cláusula da biblioteca</b> em qualquer contrato.</${Empty}>`
      : html`
      <div class="toolbar"><div class="search-box"><${Icon} name="search" size=${16} /><input class="input input-plain" placeholder="Buscar cláusula…" value=${q} onInput=${(e) => setQ(e.currentTarget.value)} /></div></div>
      <div class="lib-list">
        ${shown.map((c) => html`<article class="lib-card" key=${c.id}>
          <header>
            <span class="lib-ic"><${Icon} name="section" size=${16} /></span>
            <div class="lib-titles"><strong>${c.titulo}</strong><small>${c.origem ? 'de ' + c.origem + ' · ' : ''}salva ${relTime(c.criadoEm)}${c.xml.length > 1 ? ` · ${c.xml.length} parágrafos` : ''}</small></div>
            ${c.categoria && html`<${Badge} tone="gray">${c.categoria}</${Badge}>`}
            <${IconButton} icon="copy" title="Copiar texto" size=${16} onClick=${async () => { await navigator.clipboard.writeText(pretty(c.texto)); toast.ok('Texto copiado.'); }} />
            <${IconButton} icon="pencil" title="Renomear" size=${16} onClick=${async () => { const t = await promptDialog({ title: 'Renomear cláusula', label: 'Nome', value: c.titulo }); if (t) { c.titulo = t; await saveClausula(c); } }} />
            <${IconButton} icon="trash" title="Excluir" size=${16} onClick=${async () => { if (await confirmDialog({ title: 'Excluir da biblioteca?', message: c.titulo, ok: 'Excluir', danger: true })) await deleteClausula(c.id); }} />
          </header>
          <p class=${open === c.id ? 'lib-text open' : 'lib-text'} onClick=${() => setOpen(open === c.id ? null : c.id)}>${pretty(c.texto)}</p>
        </article>`)}
      </div>`}
  </div>`;
}
