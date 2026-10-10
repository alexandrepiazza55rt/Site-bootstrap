// Lista de contratos com busca, filtros, ações rápidas e exportação para planilha.
import { useState } from 'preact/hooks';
import { html, cx, Icon, Button, IconButton, Badge, Menu, Empty, toast, confirmDialog, downloadBlob, relTime, brl, dataBR } from '../kit.js';
import { listContratos, listModelos, deleteContrato, createContrato, tituloContrato, nomeArquivo, saveContrato } from '../../store/repo.js';
import { navigate } from '../router.js';
import { statusInfo } from '../editor/Editor.js';
import { useCollection, prazoBadge } from './Dashboard.js';
import { buildDocx } from '../../export.js';
import { FieldResolver } from '../../fields/compute.js';
import { norm } from '../dialogs.js';

const FILTROS = [['todos', 'Todos'], ['rascunho', 'Rascunhos'], ['finalizado', 'Finalizados'], ['assinado', 'Assinados'], ['encerrado', 'Encerrados'], ['cancelado', 'Cancelados']];

export function Contratos() {
  const contratos = useCollection(listContratos);
  const modelos = useCollection(listModelos);
  const [q, setQ] = useState('');
  const [f, setF] = useState('todos');
  const [modelo, setModelo] = useState('');
  if (!contratos || !modelos) return html`<div class="page-center"><span class="spinner lg" /></div>`;

  const list = contratos.filter((c) => (f === 'todos' || c.status === f) && (!modelo || c.modeloId === modelo)
    && (!q || norm([tituloContrato(c), c.modeloNome, c.resumo?.valorTexto, JSON.stringify(c.valores || {})].join(' ')).includes(norm(q))));
  const count = (s) => (s === 'todos' ? contratos.length : contratos.filter((c) => c.status === s).length);

  const baixarWord = async (c) => {
    try {
      const r = new FieldResolver(c, c.valores, c.itens);
      const blob = await buildDocx({ pkgBuf: c.pkg, docXml: c.docXml, resolver: r, titulo: tituloContrato(c) });
      downloadBlob(blob, nomeArquivo(c, 'docx'));
    } catch (e) { toast.err('Falha ao gerar: ' + e.message); }
  };
  const exportCsv = () => {
    const rows = [['Número', 'Cliente', 'Modelo', 'Status', 'Valor', 'Início', 'Término', 'Criado em', 'Atualizado em']];
    for (const c of list) rows.push([c.resumo?.numero || '', c.resumo?.cliente || '', c.modeloNome || '', statusInfo(c.status)[1], c.resumo?.valor != null ? String(c.resumo.valor).replace('.', ',') : '', dataBR(c.resumo?.inicio), dataBR(c.resumo?.termino), new Date(c.criadoEm).toLocaleDateString('pt-BR'), new Date(c.atualizadoEm).toLocaleDateString('pt-BR')]);
    const csv = '﻿' + rows.map((r) => r.map((v) => `"${String(v ?? '').replace(/"/g, '""').replace(/^—$/, '')}"`).join(';')).join('\r\n');
    downloadBlob(new Blob([csv], { type: 'text/csv;charset=utf-8' }), `contratos-${new Date().toISOString().slice(0, 10)}.csv`);
  };

  return html`<div class="page">
    <header class="page-head">
      <div><h1>Contratos</h1><p class="muted">${contratos.length} contrato(s) salvos neste computador</p></div>
      <div class="head-actions">
        <${Button} icon="download" variant="ghost" onClick=${exportCsv} disabled=${!list.length}>Planilha</${Button}>
        <${Button} icon="plus" variant="primary" onClick=${() => navigate('/novo')}>Novo contrato</${Button}>
      </div>
    </header>

    <div class="toolbar">
      <div class="search-box"><${Icon} name="search" size=${16} /><input class="input input-plain" placeholder="Buscar por cliente, número, CNPJ, equipamento…" value=${q} onInput=${(e) => setQ(e.currentTarget.value)} /></div>
      ${modelos.length > 1 && html`<select class="input w-auto" value=${modelo} onChange=${(e) => setModelo(e.currentTarget.value)}>
        <option value="">Todos os modelos</option>${modelos.map((m) => html`<option value=${m.id}>${m.nome}</option>`)}</select>`}
    </div>
    <div class="chips">
      ${FILTROS.map(([k, n]) => html`<button type="button" class=${cx('chip-filter', f === k && 'is-active')} onClick=${() => setF(k)}>${n} <span>${count(k)}</span></button>`)}
    </div>

    ${!contratos.length ? html`<${Empty} icon="files" title="Nenhum contrato ainda" action=${html`<${Button} icon="plus" variant="primary" onClick=${() => navigate('/novo')}>Criar o primeiro</${Button}>`}>Crie um contrato a partir de um modelo. Ele fica salvo aqui para editar, duplicar e baixar quando quiser.</${Empty}>`
      : !list.length ? html`<${Empty} icon="search" title="Nada encontrado">Tente outra busca ou filtro.</${Empty}>`
      : html`<div class="table-card"><table class="table">
        <thead><tr><th>Contrato</th><th>Valor</th><th>Vigência</th><th>Status</th><th>Atualizado</th><th /></tr></thead>
        <tbody>
          ${list.map((c) => {
            const st = statusInfo(c.status);
            return html`<tr key=${c.id} onClick=${() => navigate('/contrato/' + c.id)}>
              <td><div class="cell-main"><span class="row-ic sm" style=${{ background: c.cor || '#2563eb' }}><${Icon} name="file-text" size=${13} /></span>
                <div><strong>${c.resumo?.numero || '—'}</strong> <span>${c.resumo?.cliente || html`<i class="muted">cliente não informado</i>`}</span><small>${c.modeloNome}</small></div></div></td>
              <td class="nowrap">${c.resumo?.valor != null ? brl(c.resumo.valor) : '—'}</td>
              <td class="nowrap">${c.resumo?.inicio ? dataBR(c.resumo.inicio) : '—'} ${c.resumo?.termino && html`<span class="muted">→</span>`} ${c.resumo?.termino && ['finalizado', 'assinado'].includes(c.status) ? prazoBadge(c.resumo.termino) : c.resumo?.termino ? dataBR(c.resumo.termino) : ''}</td>
              <td><${Badge} tone=${st[2]}>${st[1]}</${Badge}></td>
              <td class="muted nowrap">${relTime(c.atualizadoEm)}</td>
              <td class="actions" onClick=${(e) => e.stopPropagation()}>
                <${Menu} align="right" trigger=${({ toggle }) => html`<${IconButton} icon="more" title="Ações" onClick=${toggle} />`} items=${[
                  { icon: 'pencil', label: 'Abrir', onClick: () => navigate('/contrato/' + c.id) },
                  { icon: 'download', label: 'Baixar Word', onClick: () => baixarWord(c) },
                  { icon: 'copy-plus', label: 'Duplicar (renovação)', onClick: async () => { const n = await createContrato(c.modeloId, { fromId: c.id }); toast.ok('Contrato duplicado com novo número.'); navigate('/contrato/' + n.id); } },
                  { header: 'Status' },
                  ...[['assinado', 'Marcar como assinado', 'badge'], ['encerrado', 'Marcar como encerrado', 'check'], ['cancelado', 'Marcar como cancelado', 'x']].filter(([k]) => k !== c.status).map(([k, n, ic]) => ({ icon: ic, label: n, onClick: () => { c.status = k; saveContrato(c, 'Status: ' + statusInfo(k)[1]); } })),
                  '-',
                  { icon: 'trash', label: 'Excluir', danger: true, onClick: async () => { if (await confirmDialog({ title: 'Excluir contrato?', message: `${tituloContrato(c)} será excluído definitivamente.`, ok: 'Excluir', danger: true })) { await deleteContrato(c.id); toast('Contrato excluído.'); } } },
                ]} />
              </td>
            </tr>`;
          })}
        </tbody>
      </table></div>`}
  </div>`;
}
