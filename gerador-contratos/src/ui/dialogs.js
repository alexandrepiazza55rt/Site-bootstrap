// Diálogos compartilhados: novo contrato, cadastros, biblioteca de cláusulas, pendências e dica de PDF.
import { useEffect, useMemo, useState } from 'preact/hooks';
import { html, cx, Icon, Button, Badge, ModalHead, openModal, relTime } from './kit.js';
import { listModelos, listCadastros, listClausulas, listContratos, tituloContrato } from '../store/repo.js';

const norm = (s) => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();

export function newContractDialog({ modeloId } = {}) {
  return openModal(({ close }) => {
    const Body = () => {
      const [modelos, setModelos] = useState(null);
      const [contratos, setContratos] = useState([]);
      const [sel, setSel] = useState(modeloId || null);
      const [base, setBase] = useState('');
      useEffect(() => {
        listModelos().then((l) => { setModelos(l); if (!sel && l[0]) setSel(l[0].id); });
        listContratos().then(setContratos);
      }, []);
      const doContratos = contratos.filter((c) => c.modeloId === sel);
      return html`<div>
        <${ModalHead} title="Novo contrato" sub="Escolha o modelo. O número do contrato é gerado automaticamente." icon="file-plus" close=${close} />
        <div class="modal-body">
          ${!modelos ? html`<div class="center pad"><span class="spinner" /></div>` : !modelos.length ? html`<p class="muted">Nenhum modelo cadastrado. Importe um modelo do Word primeiro.</p>` : html`
            <div class="model-pick">
              ${modelos.map((m) => html`<button type="button" class=${cx('model-opt', sel === m.id && 'is-active')} onClick=${() => { setSel(m.id); setBase(''); }}>
                <span class="model-swatch" style=${{ background: m.cor || '#2563eb' }}><${Icon} name="file-text" size=${18} /></span>
                <span class="model-opt-text"><strong>${m.nome}</strong><small>${m.campos.length} campos · atualizado ${relTime(m.atualizadoEm)}</small></span>
                ${sel === m.id && html`<${Icon} name="check-circle" size=${18} class="ok-ic" />`}
              </button>`)}
            </div>
            ${doContratos.length > 0 && html`<label class="field mt12"><span class="field-label">Começar a partir de… <small>(opcional)</small></span>
              <select class="input" value=${base} onChange=${(e) => setBase(e.currentTarget.value)}>
                <option value="">Modelo em branco</option>
                ${doContratos.slice(0, 50).map((c) => html`<option value=${c.id}>Cópia de: ${tituloContrato(c)}</option>`)}
              </select>
              <small class="muted">Copia os dados e as cláusulas editadas de um contrato anterior — ótimo para renovações.</small></label>`}
          `}
        </div>
        <div class="modal-foot">
          <${Button} variant="ghost" onClick=${() => close(undefined)}>Cancelar</${Button}>
          <${Button} variant="primary" icon="arrow-right" disabled=${!sel} onClick=${() => close({ modeloId: sel, fromId: base || undefined })}>Criar contrato</${Button}>
        </div>
      </div>`;
    };
    return html`<${Body} />`;
  }, { size: 'md' });
}

export function pickRecordDialog({ tipo, titulo }) {
  return openModal(({ close }) => {
    const Body = () => {
      const [list, setList] = useState(null);
      const [q, setQ] = useState('');
      useEffect(() => { listCadastros().then((l) => setList(l.filter((r) => r.tipo === tipo))); }, []);
      const shown = (list || []).filter((r) => !q || norm(r.titulo + ' ' + (r.subtitulo || '')).includes(norm(q)));
      return html`<div>
        <${ModalHead} title=${titulo || 'Escolher do cadastro'} sub=${tipo} icon="users" close=${close} />
        <div class="modal-body">
          <div class="outline-search"><${Icon} name="search" size=${15} /><input class="input input-plain" autoFocus placeholder="Buscar…" value=${q} onInput=${(e) => setQ(e.currentTarget.value)} /></div>
          ${!list ? html`<div class="center pad"><span class="spinner" /></div>`
            : !list.length ? html`<div class="empty small"><p>Nenhum registro em <b>${tipo}</b> ainda. Ao preencher um contrato, use “Salvar estes dados no cadastro”.</p></div>`
            : html`<ul class="pick-list">${shown.map((r) => html`<li key=${r.id}><button type="button" onClick=${() => close(r)}>
                <strong>${r.titulo}</strong>${r.subtitulo && html`<small>${r.subtitulo}</small>`}<${Icon} name="chevron-right" size=${16} />
              </button></li>`)}</ul>`}
        </div>
      </div>`;
    };
    return html`<${Body} />`;
  }, { size: 'sm' });
}

export function libraryPickDialog() {
  return openModal(({ close }) => {
    const Body = () => {
      const [list, setList] = useState(null);
      const [q, setQ] = useState('');
      const [sel, setSel] = useState(null);
      useEffect(() => { listClausulas().then(setList); }, []);
      const shown = (list || []).filter((c) => !q || norm(c.titulo + ' ' + c.texto + ' ' + (c.categoria || '')).includes(norm(q)));
      return html`<div>
        <${ModalHead} title="Inserir cláusula da biblioteca" sub="Cláusulas salvas para reaproveitar em qualquer contrato" icon="library" close=${close} />
        <div class="modal-body lib-pick">
          <div class="outline-search"><${Icon} name="search" size=${15} /><input class="input input-plain" autoFocus placeholder="Buscar cláusula…" value=${q} onInput=${(e) => setQ(e.currentTarget.value)} /></div>
          ${!list ? html`<div class="center pad"><span class="spinner" /></div>`
            : !list.length ? html`<div class="empty small"><p>Sua biblioteca está vazia. Selecione uma cláusula em qualquer contrato e clique em <b>Salvar na biblioteca</b>.</p></div>`
            : html`<ul class="pick-list">${shown.map((c) => html`<li key=${c.id}><button type="button" class=${cx(sel === c && 'is-active')} onClick=${() => setSel(c)} onDblClick=${() => close(c)}>
                <strong>${c.titulo}</strong><small>${c.texto.slice(0, 160)}${c.texto.length > 160 ? '…' : ''}</small>
                ${c.categoria && html`<${Badge} tone="gray">${c.categoria}</${Badge}>`}
              </button></li>`)}</ul>`}
        </div>
        <div class="modal-foot">
          <${Button} variant="ghost" onClick=${() => close(undefined)}>Cancelar</${Button}>
          <${Button} variant="primary" icon="plus" disabled=${!sel} onClick=${() => close(sel)}>Inserir</${Button}>
        </div>
      </div>`;
    };
    return html`<${Body} />`;
  }, { size: 'md' });
}

export function missingDialog({ missing, formato }) {
  return openModal(({ close }) => html`<div>
    <${ModalHead} title="Alguns campos obrigatórios estão vazios" icon="alert" close=${close} />
    <div class="modal-body">
      <ul class="missing-list">
        ${missing.slice(0, 12).map((m) => html`<li><button type="button" class="link-btn" onClick=${() => close({ goto: m })}>
          <${Icon} name="alert-circle" size=${14} /> ${m.campo.rotulo}${m.idx !== undefined ? ` (item ${m.idx + 1})` : ''}</button></li>`)}
        ${missing.length > 12 && html`<li class="muted">e mais ${missing.length - 12}…</li>`}
      </ul>
      <p class="muted small">Campos vazios saem em branco no ${formato}.</p>
    </div>
    <div class="modal-foot">
      <${Button} variant="ghost" onClick=${() => close({ goto: missing[0] })}>Preencher agora</${Button}>
      <${Button} variant="primary" onClick=${() => close({ go: true })}>Gerar mesmo assim</${Button}>
    </div>
  </div>`, { size: 'sm' });
}

export function pdfTipDialog() {
  return openModal(({ close }) => {
    const Body = () => {
      const [never, setNever] = useState(false);
      return html`<div>
        <${ModalHead} title="Gerar PDF" icon="file-down" close=${close} />
        <div class="modal-body">
          <div class="pdf-tip">
            <div class="pdf-step"><span>1</span><p>A janela de impressão vai abrir com o contrato já paginado.</p></div>
            <div class="pdf-step"><span>2</span><p>Em <b>Destino</b>, escolha <b>“Salvar como PDF”</b>.</p></div>
            <div class="pdf-step"><span>3</span><p>Deixe <b>Margens: Padrão</b> e <b>Escala: 100%</b>, e desmarque “Cabeçalhos e rodapés”. Clique em <b>Salvar</b>.</p></div>
          </div>
          <p class="muted small">O PDF sai com texto selecionável e qualidade de impressão (não é imagem). Se preferir, baixe em Word e use “Salvar como PDF” no Word.</p>
          <label class="check"><input type="checkbox" checked=${never} onChange=${(e) => setNever(e.currentTarget.checked)} /> Não mostrar novamente</label>
        </div>
        <div class="modal-foot">
          <${Button} variant="ghost" onClick=${() => close(undefined)}>Cancelar</${Button}>
          <${Button} variant="primary" icon="printer" onClick=${() => close({ never })}>Continuar</${Button}>
        </div>
      </div>`;
    };
    return html`<${Body} />`;
  }, { size: 'sm' });
}

export function useMemoNorm(list, q, fn) {
  return useMemo(() => list.filter((x) => !q || norm(fn(x)).includes(norm(q))), [list, q]);
}

export { norm };
