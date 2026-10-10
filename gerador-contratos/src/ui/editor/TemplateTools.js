// Ferramentas do modo "modelo": gerenciar campos/grupos, criar campo a partir da seleção e detecção automática.
import { useMemo, useState } from 'preact/hooks';
import { html, cx, Icon, IconButton, Button, Badge, ModalHead, openModal, confirmDialog, toast } from '../kit.js';
import { TIPOS, CAIXAS, PAPEIS, isCalculado, novoCampo, sugerirTipo } from '../../fields/types.js';
import { FieldResolver, parseFormula, tokensIn } from '../../fields/compute.js';
import { slugId } from '../../docx/normalize.js';
import { NS } from '../../docx/xml.js';
import { parseParagraph, writeParagraph, replaceRange, segsText, sliceSegs } from '../../docx/segments.js';
import { dataBRparaISO } from '../../br/docs.js';

// ------------------------------------------------------- utilitários de texto do documento
/** Texto linear de cada parágrafo (campos/objetos = 1 caractere). */
export function linearTexts(doc) {
  const out = [];
  for (const p of doc.getElementsByTagNameNS(NS.w, 'p')) {
    const segs = parseParagraph(p);
    out.push({ p, segs, text: segsText(segs, { fields: 'char' }) });
  }
  return out;
}

/** Substitui intervalos [start,end) por campos; ocorrências agrupadas por parágrafo. */
export function applyFieldRanges(occurrences) {
  const byP = new Map();
  for (const o of occurrences) {
    if (!byP.has(o.p)) byP.set(o.p, []);
    byP.get(o.p).push(o);
  }
  for (const [p, list] of byP) {
    let segs = parseParagraph(p);
    list.sort((a, b) => b.start - a.start);
    let lastStart = Infinity;
    for (const o of list) {
      if (o.end > lastStart) continue; // sobreposição
      const inside = sliceSegs(segs, o.start, o.end);
      const first = inside.find((s) => s.k === 't') || inside[0];
      segs = replaceRange(segs, o.start, o.end, [{ k: 'f', id: o.id, rPr: first ? first.rPr : null }]);
      lastStart = o.start;
    }
    writeParagraph(p, segs);
  }
}

export function countTokens(xml) {
  const m = new Map();
  const re = /\{\{([A-Za-z0-9_]+)\}\}/g;
  let r;
  while ((r = re.exec(xml))) m.set(r[1], (m.get(r[1]) || 0) + 1);
  return m;
}

export function uniqueId(base, campos) {
  let id = slugId(base);
  if (!campos.some((c) => c.id === id)) return id;
  let i = 2;
  while (campos.some((c) => c.id === id + '_' + i)) i++;
  return id + '_' + i;
}

// ------------------------------------------------------- gerenciador de campos (painel)
export function FieldsManager({ campos, grupos, xml, onEditField, onNewField, onDeleteField, onEditGroup, onNewGroup, onDetect, onFocusField }) {
  const counts = useMemo(() => countTokens(xml), [xml]);
  const [q, setQ] = useState('');
  const norm = (s) => String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
  const match = (c) => !q || norm(c.rotulo + ' ' + c.id).includes(norm(q));
  const all = [...grupos, ...(campos.some((c) => !grupos.find((g) => g.id === c.grupo)) ? [{ id: '__sem', nome: 'Sem grupo' }] : [])];
  return html`<div class="fm">
    <div class="fm-top">
      <${Button} icon="wand" variant="primary" size="sm" onClick=${onDetect}>Detectar campos</${Button}>
      <${Button} icon="plus" size="sm" onClick=${() => onNewField()}>Campo</${Button}>
      <${Button} icon="folder" size="sm" variant="ghost" onClick=${onNewGroup}>Grupo</${Button}>
    </div>
    <div class="fm-tip"><${Icon} name="click" size=${15} /> Dica: selecione um trecho no documento e clique em <b>Transformar em campo</b>.</div>
    <div class="outline-search"><${Icon} name="search" size=${15} /><input class="input input-plain" placeholder="Filtrar campos…" value=${q} onInput=${(e) => setQ(e.currentTarget.value)} /></div>
    ${all.map((g) => {
      const list = campos.filter((c) => (g.id === '__sem' ? !grupos.find((x) => x.id === c.grupo) : c.grupo === g.id) && match(c));
      if (!list.length && q) return null;
      return html`<section class="fm-group" key=${g.id}>
        <header>
          <${Icon} name=${g.icone || 'folder'} size=${15} />
          <strong>${g.nome}</strong>
          ${g.lista && html`<${Badge} tone="violet" icon="layers">lista</${Badge}>`}
          <span class="spacer" />
          ${g.id !== '__sem' && html`<${IconButton} icon="pencil" size=${14} title="Editar grupo" onClick=${() => onEditGroup(g)} />`}
          ${g.id !== '__sem' && html`<${IconButton} icon="plus" size=${14} title="Novo campo neste grupo" onClick=${() => onNewField({ grupo: g.id })} />`}
        </header>
        ${list.length ? html`<ul>
          ${list.map((c) => {
            const n = counts.get(c.id) || 0;
            return html`<li key=${c.id} class="fm-field" onClick=${() => onEditField(c)}>
              <div class="fm-main">
                <span class="fm-label">${c.rotulo}${c.obrigatorio && html`<i class="req">*</i>`}</span>
                <code>{{${c.id}}}</code>
              </div>
              <span class=${cx('fm-type', isCalculado(c) && 'calc')}>${(TIPOS[c.tipo] || {}).nome || c.tipo}</span>
              <button type="button" class=${cx('fm-count', !n && 'zero')} title=${n ? `Aparece ${n}× no texto — clique para localizar` : 'Não aparece no texto'}
                onClick=${(e) => { e.stopPropagation(); if (n) onFocusField(c.id); }}>${n}×</button>
              <${IconButton} icon="trash" size=${14} title="Excluir campo" onClick=${(e) => { e.stopPropagation(); onDeleteField(c); }} />
            </li>`;
          })}
        </ul>` : html`<div class="fm-empty">Nenhum campo neste grupo.</div>`}
      </section>`;
    })}
  </div>`;
}

// ------------------------------------------------------- diálogo de campo
const TIPO_GRUPOS = ['Entrada', 'Documentos', 'Números', 'Datas', 'Automáticos'];

export function fieldDialog({ campo, campos, grupos, isNew, testValues }) {
  return openModal(({ close }) => html`<${FieldDialogBody} campo=${campo} campos=${campos} grupos=${grupos} isNew=${isNew} testValues=${testValues} close=${close} />`, { size: 'md' });
}

function FieldDialogBody({ campo, campos, grupos, isNew, testValues, close }) {
  const [c, setC] = useState(() => novoCampo({ ...campo }));
  const [idTouched, setIdTouched] = useState(!isNew);
  const set = (patch) => setC((x) => {
    const n = { ...x, ...patch };
    if (isNew && !idTouched && patch.rotulo !== undefined) n.id = uniqueId(patch.rotulo || 'campo', campos.filter((k) => k !== campo));
    return n;
  });
  const others = campos.filter((k) => k.id !== campo.id);
  const idClash = others.some((k) => k.id === c.id);
  const idBad = !/^[a-z][a-z0-9_]*$/.test(c.id || '');
  let formulaErr = null;
  if (c.tipo === 'calculo') { try { parseFormula(c.formula || ''); } catch (e) { formulaErr = e.message; } }
  const preview = useMemo(() => {
    try {
      const model = { campos: [...others, c], grupos };
      const r = new FieldResolver(model, testValues || {}, {});
      const v = r.get(c.id);
      return v.empty ? (v.error || '(vazio)') : v.text;
    } catch (e) { return e.message; }
  }, [c, testValues]);
  const numericFields = others.filter((k) => ['moeda', 'numero', 'calculo', 'indice'].includes(k.tipo));
  const dateFields = others.filter((k) => ['data', 'dataCalculada'].includes(k.tipo));
  const srcFields = c.tipo === 'dataExtenso' ? dateFields : c.tipo === 'dataCalculada' ? dateFields : [...numericFields, ...dateFields];
  const valid = c.rotulo.trim() && !idClash && !idBad && !formulaErr;

  return html`<form onSubmit=${(e) => { e.preventDefault(); if (valid) close(c); }}>
    <${ModalHead} title=${isNew ? 'Novo campo' : 'Editar campo'} sub=${'Aparece no texto como {{' + (c.id || '…') + '}}'} icon="braces" close=${close} />
    <div class="modal-body grid2">
      <label class="field span2"><span class="field-label">Nome do campo (como aparece no formulário)</span>
        <input class="input" value=${c.rotulo} autoFocus onInput=${(e) => set({ rotulo: e.currentTarget.value })} placeholder="Ex.: Nome do cliente" /></label>
      <label class="field"><span class="field-label">Tipo</span>
        <select class="input" value=${c.tipo} onChange=${(e) => set({ tipo: e.currentTarget.value })}>
          ${TIPO_GRUPOS.map((g) => html`<optgroup label=${g}>${Object.entries(TIPOS).filter(([, t]) => t.grupo === g).map(([k, t]) => html`<option value=${k}>${t.nome}</option>`)}</optgroup>`)}
        </select></label>
      <label class="field"><span class="field-label">Grupo no formulário</span>
        <select class="input" value=${c.grupo} onChange=${(e) => set({ grupo: e.currentTarget.value })}>
          ${grupos.map((g) => html`<option value=${g.id}>${g.nome}${g.lista ? ' (lista)' : ''}</option>`)}
        </select></label>

      ${!isCalculado(c) && c.tipo !== 'sequencial' && html`
        <label class="field"><span class="field-label">Valor padrão <small>(opcional)</small></span>
          ${c.tipo === 'data'
            ? html`<select class="input" value=${c.padrao === 'hoje' ? 'hoje' : c.padrao ? 'fixo' : ''} onChange=${(e) => set({ padrao: e.currentTarget.value === 'hoje' ? 'hoje' : '' })}>
                <option value="">Sem padrão</option><option value="hoje">Data de hoje</option></select>`
            : html`<input class="input" value=${c.padrao || ''} onInput=${(e) => set({ padrao: e.currentTarget.value })} />`}
        </label>
        <label class="field"><span class="field-label">Dica no formulário</span>
          <input class="input" value=${c.dica || ''} onInput=${(e) => set({ dica: e.currentTarget.value })} placeholder="Ex.: 00.000.000/0000-00" /></label>`}

      ${c.tipo === 'lista' && html`<label class="field span2"><span class="field-label">Opções (uma por linha)</span>
        <textarea class="input" rows="4" value=${(c.opcoes || []).join('\n')} onInput=${(e) => set({ opcoes: e.currentTarget.value.split('\n').map((s) => s.trim()).filter(Boolean) })} /></label>`}

      ${c.tipo === 'numero' && html`
        <label class="field"><span class="field-label">Mínimo de dígitos</span>
          <input class="input" type="number" min="0" max="8" value=${c.digitos || 0} onInput=${(e) => set({ digitos: +e.currentTarget.value || 0 })} />
          <small class="muted">Ex.: 2 → “06”</small></label>
        <label class="field"><span class="field-label">Casas decimais</span>
          <input class="input" type="number" min="0" max="6" placeholder="automático" value=${c.decimais ?? ''} onInput=${(e) => set({ decimais: e.currentTarget.value === '' ? '' : +e.currentTarget.value })} /></label>`}

      ${(c.tipo === 'moeda' || (c.tipo === 'calculo' && c.formato === 'moeda')) && html`<label class="check span2">
        <input type="checkbox" checked=${!!c.simbolo} onChange=${(e) => set({ simbolo: e.currentTarget.checked })} /> Incluir “R$ ” antes do valor</label>`}

      ${c.tipo === 'data' && html`<label class="field"><span class="field-label">Formato no texto</span>
        <select class="input" value=${c.formato || ''} onChange=${(e) => set({ formato: e.currentTarget.value })}>
          <option value="">10/10/2026</option><option value="extenso">10 de outubro de 2026</option></select></label>`}

      ${(c.tipo === 'extenso' || c.tipo === 'dataExtenso' || c.tipo === 'dataCalculada') && html`<label class="field"><span class="field-label">${c.tipo === 'dataCalculada' ? 'Data inicial' : 'Escrever por extenso o campo'}</span>
        <select class="input" value=${c.origem || ''} onChange=${(e) => set({ origem: e.currentTarget.value })}>
          <option value="">Selecione…</option>
          ${srcFields.map((k) => html`<option value=${k.id}>${k.rotulo}</option>`)}
        </select></label>`}
      ${c.tipo === 'extenso' && html`<label class="check"><input type="checkbox" checked=${!!c.feminino} onChange=${(e) => set({ feminino: e.currentTarget.checked })} /> Feminino (“duas”, “duzentas”)</label>`}
      ${(c.tipo === 'dataExtenso' || (c.tipo === 'data' && c.formato === 'extenso') || (c.tipo === 'dataCalculada' && c.formato === 'extenso')) && html`<label class="check"><input type="checkbox" checked=${!!c.mesMaiusculo} onChange=${(e) => set({ mesMaiusculo: e.currentTarget.checked })} /> Mês com inicial maiúscula (“Outubro”)</label>`}
      ${c.tipo === 'dataCalculada' && html`
        <label class="field"><span class="field-label">Somar meses</span>
          <select class="input" value=${Number.isFinite(+c.meses) && c.meses !== '' ? '__n' : c.meses || ''} onChange=${(e) => set({ meses: e.currentTarget.value === '__n' ? 1 : e.currentTarget.value })}>
            <option value="">Nenhum</option><option value="__n">Número fixo…</option>
            ${numericFields.map((k) => html`<option value=${k.id}>Campo: ${k.rotulo}</option>`)}
          </select>
          ${Number.isFinite(+c.meses) && c.meses !== '' && html`<input class="input mt4" type="number" value=${c.meses} onInput=${(e) => set({ meses: +e.currentTarget.value })} />`}
        </label>
        <label class="field"><span class="field-label">Somar dias</span>
          <input class="input" type="number" value=${c.dias || ''} placeholder="0" onInput=${(e) => set({ dias: e.currentTarget.value })} /></label>
        <label class="field"><span class="field-label">Formato no texto</span>
          <select class="input" value=${c.formato || ''} onChange=${(e) => set({ formato: e.currentTarget.value })}>
            <option value="">14/04/2027</option><option value="extenso">14 de abril de 2027</option></select></label>`}

      ${c.tipo === 'calculo' && html`
        <label class="field span2"><span class="field-label">Fórmula</span>
          <input class="input mono" value=${c.formula || ''} onInput=${(e) => set({ formula: e.currentTarget.value })} placeholder="Ex.: valor_mensal * prazo_meses" />
          ${formulaErr ? html`<small class="bad">${formulaErr}</small>` : html`<small class="muted">Use os códigos dos campos, + − × ÷ ( ), e funções: soma(), arred(x, 2), min(), max(). Para somar uma lista: soma(item_qtd * item_valor).</small>`}
          <div class="chips-row">${numericFields.slice(0, 14).map((k) => html`<button type="button" class="chip-btn" onClick=${() => set({ formula: ((c.formula || '') + ' ' + k.id).trim() })}>${k.id}</button>`)}</div>
        </label>
        <label class="field"><span class="field-label">Formato do resultado</span>
          <select class="input" value=${c.formato || 'numero'} onChange=${(e) => set({ formato: e.currentTarget.value })}>
            <option value="moeda">Valor (R$ 0,00)</option><option value="numero">Número</option><option value="inteiro">Número inteiro</option></select></label>
        <label class="check"><input type="checkbox" checked=${c.editavel !== false} onChange=${(e) => set({ editavel: e.currentTarget.checked })} /> Permitir digitar outro valor</label>`}

      ${c.tipo === 'sequencial' && html`<label class="field span2"><span class="field-label">Formato do número</span>
        <input class="input mono" value=${c.formato || '{n:4}/{ano}'} onInput=${(e) => set({ formato: e.currentTarget.value })} />
        <small class="muted">{n:4} = número com 4 dígitos, {ano} = ano atual. Ex.: 0063/2026</small></label>`}

      ${!['indice', 'calculo', 'dataCalculada', 'extenso', 'dataExtenso'].includes(c.tipo) && html`<label class="field"><span class="field-label">Maiúsculas/minúsculas</span>
        <select class="input" value=${c.caixa || ''} onChange=${(e) => set({ caixa: e.currentTarget.value })}>${CAIXAS.map(([k, n]) => html`<option value=${k}>${n}</option>`)}</select></label>`}
      ${c.tipo === 'extenso' && html`<label class="field"><span class="field-label">Maiúsculas/minúsculas</span>
        <select class="input" value=${c.caixa || ''} onChange=${(e) => set({ caixa: e.currentTarget.value })}>${CAIXAS.map(([k, n]) => html`<option value=${k}>${n}</option>`)}</select></label>`}

      <details class="span2 adv">
        <summary>Opções avançadas</summary>
        <div class="grid2">
          <label class="field"><span class="field-label">Código no texto</span>
            <input class=${cx('input mono', (idClash || idBad) && 'is-invalid')} value=${c.id} onInput=${(e) => { setIdTouched(true); setC({ ...c, id: e.currentTarget.value.toLowerCase().replace(/[^a-z0-9_]/g, '_') }); }} />
            ${idClash && html`<small class="bad">Já existe um campo com este código.</small>`}
            ${!isNew && c.id !== campo.id && !idClash && html`<small class="muted">O texto do modelo será atualizado automaticamente.</small>`}</label>
          <label class="field"><span class="field-label">Papel (para consultas e painel)</span>
            <select class="input" value=${c.papel || ''} onChange=${(e) => set({ papel: e.currentTarget.value })}>${PAPEIS.map(([k, n]) => html`<option value=${k}>${n}</option>`)}</select></label>
          <label class="field"><span class="field-label">Largura no formulário</span>
            <select class="input" value=${c.largura || 'full'} onChange=${(e) => set({ largura: e.currentTarget.value })}>
              <option value="full">Inteira</option><option value="half">Metade</option><option value="third">Um terço</option></select></label>
          <label class="field"><span class="field-label">Se ficar vazio, mostrar</span>
            <input class="input" value=${c.vazio || ''} placeholder="(nada)" onInput=${(e) => set({ vazio: e.currentTarget.value })} /></label>
        </div>
      </details>
      ${!isCalculado(c) && html`<label class="check span2"><input type="checkbox" checked=${!!c.obrigatorio} onChange=${(e) => set({ obrigatorio: e.currentTarget.checked })} /> Obrigatório (avisa antes de gerar o documento)</label>`}
      <div class="preview-box span2"><span class="muted small">Prévia com os dados de teste:</span> <strong>${preview}</strong></div>
    </div>
    <div class="modal-foot">
      <${Button} variant="ghost" onClick=${() => close(undefined)}>Cancelar</${Button}>
      <${Button} variant="primary" type="submit" disabled=${!valid}>${isNew ? 'Criar campo' : 'Salvar'}</${Button}>
    </div>
  </form>`;
}

// ------------------------------------------------------- grupo
export function groupDialog({ grupo, isNew }) {
  return openModal(({ close }) => {
    const Body = () => {
      const [g, setG] = useState({ id: '', nome: '', icone: 'folder', lista: false, recolhido: false, itemNome: '', cadastro: '', ...grupo });
      const icons = ['folder', 'file-text', 'user', 'users', 'building', 'truck', 'package', 'coins', 'calendar', 'map-pin', 'tag', 'scale', 'briefcase', 'id'];
      return html`<form onSubmit=${(e) => { e.preventDefault(); if (g.nome.trim()) close({ ...g, id: g.id || slugId(g.nome) }); }}>
        <${ModalHead} title=${isNew ? 'Novo grupo de campos' : 'Editar grupo'} icon="folder" close=${close} />
        <div class="modal-body">
          <label class="field"><span class="field-label">Nome do grupo</span><input class="input" autoFocus value=${g.nome} onInput=${(e) => setG({ ...g, nome: e.currentTarget.value })} placeholder="Ex.: Fiador" /></label>
          <div class="field"><span class="field-label">Ícone</span><div class="icon-pick">${icons.map((ic) => html`<button type="button" class=${cx(g.icone === ic && 'is-active')} onClick=${() => setG({ ...g, icone: ic })}><${Icon} name=${ic} size=${17} /></button>`)}</div></div>
          <label class="check"><input type="checkbox" checked=${g.lista} onChange=${(e) => setG({ ...g, lista: e.currentTarget.checked })} /> <span><b>Lista de itens</b> — a linha da tabela (ou o parágrafo) com estes campos se repete para cada item. Ideal para equipamentos, parcelas, produtos.</span></label>
          ${g.lista && html`<label class="field"><span class="field-label">Nome de cada item</span><input class="input" value=${g.itemNome} placeholder="Ex.: Equipamento" onInput=${(e) => setG({ ...g, itemNome: e.currentTarget.value })} /></label>`}
          <label class="field"><span class="field-label">Cadastro reutilizável <small>(opcional)</small></span><input class="input" value=${g.cadastro} placeholder="Ex.: Clientes, Equipamentos, Fiadores" onInput=${(e) => setG({ ...g, cadastro: e.currentTarget.value })} />
            <small class="muted">Permite salvar e reaproveitar estes dados em outros contratos.</small></label>
          <label class="check"><input type="checkbox" checked=${g.recolhido} onChange=${(e) => setG({ ...g, recolhido: e.currentTarget.checked })} /> Começar recolhido no formulário</label>
        </div>
        <div class="modal-foot">
          ${!isNew && html`<${Button} variant="danger-ghost" icon="trash" onClick=${() => close({ __delete: true })}>Excluir grupo</${Button}>`}
          <span class="spacer" />
          <${Button} variant="ghost" onClick=${() => close(undefined)}>Cancelar</${Button}>
          <${Button} variant="primary" type="submit" disabled=${!g.nome.trim()}>Salvar</${Button}>
        </div>
      </form>`;
    };
    return html`<${Body} />`;
  }, { size: 'sm' });
}

// ------------------------------------------------------- criar campo a partir da seleção
export function createFromSelectionDialog({ info, doc, campos, grupos }) {
  const texts = linearTexts(doc);
  const occ = [];
  const needle = info.text;
  if (needle.length >= 2) {
    for (const t of texts) {
      let i = t.text.indexOf(needle);
      while (i >= 0) {
        if (!(t.p === info.node && i === info.start)) occ.push({ p: t.p, start: i, end: i + needle.length });
        i = t.text.indexOf(needle, i + needle.length);
      }
    }
  }
  const sug = sugerirTipo(needle);
  return openModal(({ close }) => {
    const Body = () => {
      const [modo, setModo] = useState('novo');
      const [rotulo, setRotulo] = useState('');
      const [tipo, setTipo] = useState(sug.tipo);
      const [grupo, setGrupo] = useState(grupos[0] ? grupos[0].id : 'geral');
      const [existente, setExistente] = useState(campos[0] ? campos[0].id : '');
      const [todas, setTodas] = useState(occ.length > 0);
      const [padrao, setPadrao] = useState(false);
      const ok = modo === 'novo' ? rotulo.trim() : existente;
      const submit = (e) => {
        e.preventDefault();
        if (!ok) return;
        let campo;
        if (modo === 'novo') {
          let pad = '';
          if (padrao) pad = sug.tipo === 'data' ? dataBRparaISO(needle) || '' : needle;
          campo = novoCampo({ id: uniqueId(rotulo, campos), rotulo: rotulo.trim(), tipo, grupo, caixa: sug.caixa || '', digitos: sug.digitos || 0, padrao: pad,
            formato: sug.extensoData ? 'extenso' : undefined });
        } else campo = campos.find((c) => c.id === existente);
        close({ campo, isNew: modo === 'novo', ranges: [{ p: info.node, start: info.start, end: info.end }, ...(todas ? occ : [])] });
      };
      return html`<form onSubmit=${submit}>
        <${ModalHead} title="Transformar em campo" sub=${'“' + (needle.length > 60 ? needle.slice(0, 58) + '…' : needle) + '”'} icon="braces" close=${close} />
        <div class="modal-body">
          <div class="seg-tabs">
            <button type="button" class=${cx(modo === 'novo' && 'is-active')} onClick=${() => setModo('novo')}>Criar novo campo</button>
            <button type="button" class=${cx(modo === 'existe' && 'is-active')} disabled=${!campos.length} onClick=${() => setModo('existe')}>Usar campo existente</button>
          </div>
          ${modo === 'novo' ? html`<div class="grid2">
            <label class="field span2"><span class="field-label">Nome do campo</span>
              <input class="input" autoFocus value=${rotulo} placeholder="Ex.: Nome do locatário" onInput=${(e) => setRotulo(e.currentTarget.value)} /></label>
            <label class="field"><span class="field-label">Tipo</span>
              <select class="input" value=${tipo} onChange=${(e) => setTipo(e.currentTarget.value)}>
                ${Object.entries(TIPOS).filter(([, t]) => !t.calculado).map(([k, t]) => html`<option value=${k}>${t.nome}</option>`)}
              </select></label>
            <label class="field"><span class="field-label">Grupo</span>
              <select class="input" value=${grupo} onChange=${(e) => setGrupo(e.currentTarget.value)}>${grupos.map((g) => html`<option value=${g.id}>${g.nome}</option>`)}</select></label>
            <label class="check span2"><input type="checkbox" checked=${padrao} onChange=${(e) => setPadrao(e.currentTarget.checked)} /> Usar o texto selecionado como valor padrão <small class="muted">(bom para dados que quase nunca mudam)</small></label>
          </div>` : html`<label class="field"><span class="field-label">Campo</span>
            <select class="input" value=${existente} onChange=${(e) => setExistente(e.currentTarget.value)}>${campos.map((c) => html`<option value=${c.id}>${c.rotulo} — {{${c.id}}}</option>`)}</select></label>`}
          ${occ.length > 0 && html`<label class="check highlight-check"><input type="checkbox" checked=${todas} onChange=${(e) => setTodas(e.currentTarget.checked)} />
            Substituir também as outras <b>${occ.length}</b> ocorrência(s) deste mesmo texto no documento</label>`}
        </div>
        <div class="modal-foot">
          <${Button} variant="ghost" onClick=${() => close(undefined)}>Cancelar</${Button}>
          <${Button} variant="primary" type="submit" icon="check" disabled=${!ok}>Transformar</${Button}>
        </div>
      </form>`;
    };
    return html`<${Body} />`;
  }, { size: 'sm' });
}

// ------------------------------------------------------- detecção automática
const MESES = 'janeiro|fevereiro|março|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro';
const PARTES = /(LOCADOR[A]?|LOCAT[ÁA]RI[OA]|CONTRATANTE|CONTRATAD[OA]|COMPRADOR[A]?|VENDEDOR[A]?|COMODANTE|COMODAT[ÁA]RI[OA]|OUTORGANTE|OUTORGAD[OA]|FIADOR[A]?|PRESTADOR[A]?|TOMADOR[A]?|CEDENTE|CESSION[ÁA]RI[OA]|EMPREGADOR[A]?|EMPREGAD[OA])/i;

const DETECTORS = [
  { tipo: 'cnpj', nome: 'CNPJ', re: /\b\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}\b/g },
  { tipo: 'cpf', nome: 'CPF', re: /\b\d{3}\.\d{3}\.\d{3}-\d{2}\b/g },
  { tipo: 'cep', nome: 'CEP', re: /\b\d{2}\.\d{3}-\d{3}\b|\b\d{5}-\d{3}\b/g },
  { tipo: 'moeda', nome: 'Valor', re: /R\$\s?(\d{1,3}(?:\.\d{3})*,\d{2})(\s*\(([^)]{4,120})\))?/g, valueGroup: 1, extensoGroup: 3 },
  { tipo: 'data', nome: 'Data', re: /\b\d{2}\/\d{2}\/\d{4}\b/g },
  { tipo: 'data', nome: 'Data por extenso', re: new RegExp(`\\b\\d{1,2}\\s+de\\s+(?:${MESES})\\s+de\\s+\\d{4}\\b`, 'gi'), extenso: true },
  { tipo: 'telefone', nome: 'Telefone', re: /\(\d{2}\)\s?9?\.?\s?\d{4}-?\d{4}/g },
  { tipo: 'email', nome: 'E-mail', re: /\b[\w.+-]+@[\w-]+\.[\w.]+\b/g },
  { tipo: 'texto', nome: 'Espaço para preencher', re: /_{5,}|\.{6,}/g, blank: true },
];

export function detectSuggestions(doc, campos) {
  const texts = linearTexts(doc);
  const byValue = new Map();
  const add = (key, s) => {
    if (!byValue.has(key)) byValue.set(key, { ...s, occ: [] });
    byValue.get(key).occ.push(...s.occ);
  };
  for (const t of texts) {
    for (const d of DETECTORS) {
      d.re.lastIndex = 0;
      let m;
      while ((m = d.re.exec(t.text))) {
        let start = m.index, end = m.index + m[0].length;
        let valor = m[0];
        if (d.valueGroup) { start = m.index + m[0].indexOf(m[d.valueGroup]); end = start + m[d.valueGroup].length; valor = m[d.valueGroup]; }
        if (t.text.slice(start, end).includes('￼')) continue;
        const before = t.text.slice(Math.max(0, start - 160), start);
        const parte = (before.match(new RegExp(PARTES.source, 'gi')) || []).pop();
        const ctx = contextLabel(before);
        const label = labelFor(d, parte, ctx);
        const key = d.blank ? 'blank:' + t.p.getAttribute('w14:paraId') + ':' + start : d.tipo + ':' + valor;
        add(key, { tipo: d.tipo, nome: label, valor, extenso: d.extenso, blank: d.blank, occ: [{ p: t.p, start, end }] });
        if (d.extensoGroup && m[d.extensoGroup]) {
          const ex = m[d.extensoGroup];
          const es = m.index + m[0].lastIndexOf(ex);
          add('ext:' + valor + ':' + ex.toLowerCase(), { tipo: 'extenso', nome: label + ' por extenso', valor: ex, origemKey: d.tipo + ':' + valor, caixa: caseOf(ex), occ: [{ p: t.p, start: es, end: es + ex.length }] });
        }
      }
    }
    // nomes das partes em MAIÚSCULAS: "LOCATÁRIO: EMPRESA X LTDA,"
    const pm = /(LOCADOR[A]?|LOCAT[ÁA]RI[OA]|CONTRATANTE|CONTRATAD[OA]|COMPRADOR[A]?|VENDEDOR[A]?|COMODANTE|COMODAT[ÁA]RI[OA]|FIADOR[A]?|PRESTADOR[A]?|TOMADOR[A]?)\s*:\s*([A-ZÀ-Ý0-9][A-ZÀ-Ý0-9 .&\-/]{3,90}?)(?=\s*,|\s*$)/g;
    let m;
    while ((m = pm.exec(t.text))) {
      const name = m[2].trim();
      const start = m.index + m[0].indexOf(m[2]);
      if (t.text.slice(start, start + name.length).includes('￼')) continue;
      add('nome:' + name, { tipo: 'texto', nome: 'Nome — ' + cap(m[1]), valor: name, caixa: 'maiusculas', papel: 'nome', occ: [{ p: t.p, start, end: start + name.length }] });
    }
  }
  // outras ocorrências do mesmo texto (ex.: nome repetido na assinatura)
  const list = Array.from(byValue.values());
  for (const s of list) {
    if (s.blank || s.valor.length < 6) continue;
    for (const t of texts) {
      let i = t.text.indexOf(s.valor);
      while (i >= 0) {
        if (!s.occ.some((o) => o.p === t.p && o.start <= i && o.end >= i + s.valor.length) && !t.text.slice(i, i + s.valor.length).includes('￼')) s.occ.push({ p: t.p, start: i, end: i + s.valor.length });
        i = t.text.indexOf(s.valor, i + s.valor.length);
      }
    }
  }
  // rótulos únicos
  const used = new Map();
  for (const s of list) {
    const n = (used.get(s.nome) || 0) + 1;
    used.set(s.nome, n);
    if (n > 1) s.nome = `${s.nome} ${n}`;
    s.on = !s.blank;
  }
  return list.filter((s) => !(campos || []).some((c) => c.padrao && c.padrao === s.valor));
}

function contextLabel(before) {
  const b = before.toLowerCase();
  if (/representad[oa] (por|pel[oa])[^,]*$/.test(b)) return 'representante';
  if (/resid\w* [^,]*$/.test(b)) return 'residência';
  if (/sede\b[^,]*$/.test(b)) return 'sede';
  return '';
}

function labelFor(d, parte, ctx) {
  let n = d.nome;
  if (d.tipo === 'cpf' && ctx === 'representante') n = 'CPF do representante';
  if (parte) n += ' — ' + cap(parte);
  return n;
}

const cap = (s) => s[0].toUpperCase() + s.slice(1).toLowerCase();
function caseOf(s) {
  const letters = s.replace(/[^A-Za-zÀ-ÿ]/g, '');
  if (letters === letters.toUpperCase()) return 'maiusculas';
  if (letters === letters.toLowerCase()) return 'minusculas';
  if (/^[A-ZÀ-Ý]/.test(s) && /\s[A-ZÀ-Ý]/.test(s)) return 'titulo';
  return '';
}

export function detectDialog({ doc, campos, grupos }) {
  const sugs = detectSuggestions(doc, campos);
  return openModal(({ close }) => {
    const Body = () => {
      const [items, setItems] = useState(sugs);
      const [grupo, setGrupo] = useState(grupos[0] ? grupos[0].id : 'geral');
      const upd = (i, patch) => setItems(items.map((s, k) => (k === i ? { ...s, ...patch } : s)));
      const n = items.filter((s) => s.on).length;
      return html`<div>
        <${ModalHead} title="Detectar campos automaticamente" sub="Encontrei estes dados no texto. Marque os que devem virar campos preenchíveis." icon="wand" close=${close} />
        <div class="modal-body">
          ${!items.length ? html`<div class="empty small"><p>Nenhum CPF, CNPJ, CEP, valor, data ou telefone encontrado no texto (que ainda não seja campo).</p></div>`
            : html`<div class="detect-list">
              ${items.map((s, i) => html`<label class=${cx('detect-item', s.on && 'on')} key=${i}>
                <input type="checkbox" checked=${s.on} onChange=${(e) => upd(i, { on: e.currentTarget.checked })} />
                <div class="detect-main">
                  <input class="input input-sm" value=${s.nome} onInput=${(e) => upd(i, { nome: e.currentTarget.value })} onClick=${(e) => e.preventDefault()} />
                  <div class="detect-meta"><${Badge} tone=${s.tipo === 'extenso' ? 'green' : 'blue'}>${(TIPOS[s.tipo] || {}).nome || s.tipo}</${Badge}>
                    <span class="detect-val">“${s.valor.length > 70 ? s.valor.slice(0, 68) + '…' : s.valor}”</span>
                    ${s.occ.length > 1 && html`<span class="muted small">${s.occ.length}×</span>`}</div>
                </div>
              </label>`)}
            </div>`}
          ${items.length > 0 && html`<label class="field mt8"><span class="field-label">Colocar os novos campos no grupo</span>
            <select class="input" value=${grupo} onChange=${(e) => setGrupo(e.currentTarget.value)}>${grupos.map((g) => html`<option value=${g.id}>${g.nome}</option>`)}</select></label>`}
        </div>
        <div class="modal-foot">
          <span class="muted small">${n} selecionado(s)</span><span class="spacer" />
          <${Button} variant="ghost" onClick=${() => close(undefined)}>Cancelar</${Button}>
          <${Button} variant="primary" icon="wand" disabled=${!n} onClick=${() => close({ items: items.filter((s) => s.on), grupo })}>Criar ${n} campo(s)</${Button}>
        </div>
      </div>`;
    };
    return html`<${Body} />`;
  }, { size: 'md' });
}

/** Aplica sugestões aceitas: cria campos e troca o texto por {{campos}}. */
export function applySuggestions(result, campos) {
  const novos = [];
  const keyToId = new Map();
  const ranges = [];
  const all = [...campos];
  for (const s of result.items.filter((x) => x.tipo !== 'extenso')) {
    const id = uniqueId(s.nome, all);
    const c = novoCampo({ id, rotulo: s.nome, tipo: s.tipo, grupo: result.grupo, caixa: s.caixa || '', papel: s.papel || '',
      formato: s.extenso ? 'extenso' : undefined, mesMaiusculo: s.extenso && /[A-Z]/.test(s.valor.replace(/^\d+\s+de\s+/, '')[0] || '') });
    all.push(c);
    novos.push(c);
    keyToId.set(s.tipo + ':' + s.valor, id);
    for (const o of s.occ) ranges.push({ ...o, id });
  }
  for (const s of result.items.filter((x) => x.tipo === 'extenso')) {
    const origem = keyToId.get(s.origemKey) || '';
    const id = uniqueId(s.nome, all);
    const c = novoCampo({ id, rotulo: s.nome, tipo: 'extenso', grupo: result.grupo, origem, caixa: s.caixa || '' });
    all.push(c);
    novos.push(c);
    for (const o of s.occ) ranges.push({ ...o, id });
  }
  applyFieldRanges(ranges);
  return novos;
}

export { tokensIn, confirmDialog, toast };
