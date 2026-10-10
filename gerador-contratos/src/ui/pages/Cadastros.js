// Cadastros reutilizáveis (clientes, equipamentos...) salvos a partir dos contratos.
import { useState } from 'preact/hooks';
import { html, cx, Icon, Button, IconButton, Empty, toast, confirmDialog, openModal, ModalHead, relTime } from '../kit.js';
import { listCadastros, saveCadastro, deleteCadastro } from '../../store/repo.js';
import { useCollection } from './Dashboard.js';
import { norm } from '../dialogs.js';

function editRecord(rec) {
  return openModal(({ close }) => {
    const Body = () => {
      const [r, setR] = useState(structuredClone(rec));
      const keys = Object.keys(r.valores || {});
      return html`<form onSubmit=${(e) => { e.preventDefault(); close(r); }}>
        <${ModalHead} title="Editar cadastro" sub=${r.tipo} icon="user" close=${close} />
        <div class="modal-body">
          <label class="field"><span class="field-label">Nome de exibição</span><input class="input" value=${r.titulo} onInput=${(e) => setR({ ...r, titulo: e.currentTarget.value })} /></label>
          ${keys.map((k) => html`<label class="field" key=${k}><span class="field-label">${(r.rotulos && r.rotulos[k]) || k}</span>
            ${String(r.valores[k]).length > 60 ? html`<textarea class="input" rows="2" value=${r.valores[k]} onInput=${(e) => setR({ ...r, valores: { ...r.valores, [k]: e.currentTarget.value } })} />`
              : html`<input class="input" value=${r.valores[k]} onInput=${(e) => setR({ ...r, valores: { ...r.valores, [k]: e.currentTarget.value } })} />`}</label>`)}
        </div>
        <div class="modal-foot"><${Button} variant="ghost" onClick=${() => close(undefined)}>Cancelar</${Button}><${Button} variant="primary" type="submit">Salvar</${Button}></div>
      </form>`;
    };
    return html`<${Body} />`;
  }, { size: 'sm' });
}

export function Cadastros() {
  const list = useCollection(listCadastros);
  const [tipo, setTipo] = useState(null);
  const [q, setQ] = useState('');
  if (!list) return html`<div class="page-center"><span class="spinner lg" /></div>`;
  const tipos = Array.from(new Set(list.map((r) => r.tipo))).sort();
  const t = tipo && tipos.includes(tipo) ? tipo : tipos[0];
  const shown = list.filter((r) => r.tipo === t && (!q || norm(r.titulo + ' ' + r.subtitulo + ' ' + Object.values(r.valores || {}).join(' ')).includes(norm(q))));

  return html`<div class="page">
    <header class="page-head">
      <div><h1>Cadastros</h1><p class="muted">Clientes, equipamentos e outros dados que você reaproveita entre contratos.</p></div>
    </header>
    ${!list.length ? html`<${Empty} icon="users" title="Nenhum cadastro ainda">
        Dentro de um contrato, no grupo <b>Locatário</b> ou <b>Equipamentos</b>, abra o menu <b>⋯</b> e escolha <b>Salvar estes dados no cadastro</b>.
        Na próxima vez, é só escolher o cliente da lista.</${Empty}>`
      : html`
      <div class="chips">${tipos.map((k) => html`<button type="button" class=${cx('chip-filter', k === t && 'is-active')} onClick=${() => setTipo(k)}>${k} <span>${list.filter((r) => r.tipo === k).length}</span></button>`)}</div>
      <div class="toolbar"><div class="search-box"><${Icon} name="search" size=${16} /><input class="input input-plain" placeholder=${'Buscar em ' + t + '…'} value=${q} onInput=${(e) => setQ(e.currentTarget.value)} /></div></div>
      <div class="rec-grid">
        ${shown.map((r) => html`<article class="rec-card" key=${r.id}>
          <div class="rec-head"><span class="rec-av">${(r.titulo || '?').trim().slice(0, 1).toUpperCase()}</span>
            <div><strong>${r.titulo}</strong>${r.subtitulo && html`<small>${r.subtitulo}</small>`}</div>
            <span class="spacer" />
            <${IconButton} icon="pencil" title="Editar" size=${16} onClick=${async () => { const n = await editRecord(r); if (n) { await saveCadastro(n); toast.ok('Cadastro atualizado.'); } }} />
            <${IconButton} icon="trash" title="Excluir" size=${16} onClick=${async () => { if (await confirmDialog({ title: 'Excluir cadastro?', message: r.titulo, ok: 'Excluir', danger: true })) await deleteCadastro(r.id); }} />
          </div>
          <dl class="rec-vals">${Object.entries(r.valores || {}).slice(0, 5).map(([k, v]) => html`<div key=${k}><dt>${(r.rotulos && r.rotulos[k]) || k}</dt><dd>${v}</dd></div>`)}</dl>
          <small class="muted">atualizado ${relTime(r.atualizadoEm)}</small>
        </article>`)}
      </div>`}
  </div>`;
}
