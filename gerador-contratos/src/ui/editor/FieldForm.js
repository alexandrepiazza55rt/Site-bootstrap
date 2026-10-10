// Formulário de preenchimento dos campos (agrupado, com máscaras, consultas e itens repetidos).
import { useState, useEffect, useRef, useMemo } from 'preact/hooks';
import { html, cx, Icon, IconButton, Button, Badge, toast, Menu } from '../kit.js';
import { TIPOS, isCalculado } from '../../fields/types.js';
import {
  mascaraCPF, mascaraCNPJ, mascaraCpfCnpj, mascaraCEP, mascaraTelefone, validaCPF, validaCNPJ, soDigitos,
  parseMoeda, formataMoeda,
} from '../../br/docs.js';
import { consultaCNPJ, consultaCEP, enderecoCompleto, UFS } from '../../br/apis.js';

const MASKS = { cpf: mascaraCPF, cnpj: mascaraCNPJ, cpfCnpj: mascaraCpfCnpj, cep: mascaraCEP, telefone: mascaraTelefone };

function validity(campo, v) {
  if (!v) return null;
  const d = soDigitos(v);
  if (campo.tipo === 'cpf') return d.length < 11 ? null : validaCPF(v) ? 'ok' : 'CPF inválido — confira os números';
  if (campo.tipo === 'cnpj') return d.length < 14 ? null : validaCNPJ(v) ? 'ok' : 'CNPJ inválido — confira os números';
  if (campo.tipo === 'cpfCnpj') {
    if (d.length === 11) return validaCPF(v) ? 'ok' : 'CPF inválido';
    if (d.length === 14) return validaCNPJ(v) ? 'ok' : 'CNPJ inválido';
    return null;
  }
  if (campo.tipo === 'email') return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v) ? 'ok' : 'E-mail parece incompleto';
  return null;
}

/** Um campo de entrada. */
export function FieldInput({ campo, value, resolved, onChange, onLookup, onFocus, autoFocus, compact }) {
  const [busy, setBusy] = useState(false);
  const ref = useRef();
  useEffect(() => { if (autoFocus && ref.current) ref.current.focus(); }, [autoFocus]);
  const t = campo.tipo;
  const v = value ?? '';
  const ph = campo.padrao && t !== 'data' ? `Padrão: ${campo.padrao}` : campo.dica || '';
  const val = validity(campo, v);
  const common = {
    ref, id: 'f-' + campo.id + (compact ? '-' + compact : ''), class: cx('input', val && val !== 'ok' && 'is-invalid'),
    onFocus: () => onFocus && onFocus(),
  };
  let input;
  if (t === 'textoLongo') {
    input = html`<textarea ...${common} rows=${2} value=${v} placeholder=${ph}
      onInput=${(e) => { autosize(e.currentTarget); onChange(e.currentTarget.value); }} ref=${(el) => { ref.current = el; if (el) autosize(el); }} />`;
  } else if (t === 'lista') {
    const opts = campo.opcoes || [];
    input = html`<select ...${common} value=${v || campo.padrao || ''} onChange=${(e) => onChange(e.currentTarget.value)}>
      ${!campo.padrao && html`<option value="">Selecione…</option>`}
      ${opts.map((o) => html`<option value=${o}>${o}</option>`)}
      ${v && !opts.includes(v) && html`<option value=${v}>${v}</option>`}
    </select>`;
  } else if (t === 'data') {
    input = html`<input ...${common} type="date" value=${v || (campo.padrao === 'hoje' ? '' : campo.padrao || '')} onInput=${(e) => onChange(e.currentTarget.value)} />`;
  } else if (t === 'moeda' || (t === 'calculo' && campo.formato === 'moeda')) {
    const calc = t === 'calculo';
    input = html`<div class="input-affix">
      <span class="affix">R$</span>
      <input ...${common} inputmode="decimal" value=${v} placeholder=${calc ? (resolved && !resolved.empty ? resolved.text.replace(/^R\$\s?/, '') : 'calculado') : ph || '0,00'}
        onInput=${(e) => onChange(e.currentTarget.value.replace(/[^\d.,]/g, ''))}
        onBlur=${(e) => { const n = parseMoeda(e.currentTarget.value); if (e.currentTarget.value && Number.isFinite(n)) onChange(formataMoeda(n)); }} />
    </div>`;
  } else if (t === 'numero' || t === 'calculo') {
    const calc = t === 'calculo';
    input = html`<input ...${common} inputmode="decimal" value=${v} placeholder=${calc ? (resolved && !resolved.empty ? resolved.text : 'calculado') : ph}
      onInput=${(e) => onChange(e.currentTarget.value.replace(/[^\d.,-]/g, ''))} />`;
  } else {
    const mask = MASKS[t];
    input = html`<input ...${common} type=${t === 'email' ? 'email' : 'text'} value=${v} placeholder=${ph}
      inputmode=${mask && t !== 'telefone' ? 'numeric' : t === 'telefone' ? 'tel' : undefined}
      autocomplete="off" spellcheck=${t === 'texto' || t === 'textoLongo'}
      onInput=${(e) => {
        const raw = e.currentTarget.value;
        const nv = mask ? mask(raw) : raw;
        if (mask && nv !== raw) e.currentTarget.value = nv;
        onChange(nv);
      }} />`;
  }
  const canLookup = onLookup && (t === 'cnpj' || t === 'cep' || (t === 'cpfCnpj' && soDigitos(v).length === 14));
  const doLookup = async () => {
    setBusy(true);
    try { await onLookup(campo, v); } catch (e) { toast.err(e.message); } finally { setBusy(false); }
  };
  return html`<div class="fi-row">
    ${input}
    ${canLookup && html`<button type="button" class="lookup-btn" disabled=${busy} title=${t === 'cep' ? 'Buscar endereço pelo CEP' : 'Buscar dados na Receita Federal'} onClick=${doLookup}>
      ${busy ? html`<span class="spinner" />` : html`<${Icon} name="search" size=${15} />`}<span>${t === 'cep' ? 'Buscar' : 'Buscar dados'}</span>
    </button>`}
    ${val && html`<div class=${cx('fi-hint', val === 'ok' ? 'ok' : 'bad')}>${val === 'ok' ? html`<${Icon} name="check" size=${13} /> válido` : val}</div>`}
  </div>`;
}

function autosize(el) {
  el.style.height = 'auto';
  el.style.height = Math.min(el.scrollHeight + 2, 220) + 'px';
}

/** Linha "fx" mostrando campos calculados derivados. */
function Derived({ list, resolver, idx, onFieldClick }) {
  if (!list.length) return null;
  return html`<div class="derived">
    ${list.map((c) => {
      const r = resolver.get(c.id, idx);
      return html`<button type="button" class="derived-item" title=${'Aparece no contrato como {{' + c.id + '}}'} onClick=${() => onFieldClick && onFieldClick(c.id, idx)}>
        <${Icon} name="wand" size=${13} />
        <span class="derived-label">${c.rotulo}</span>
        <span class=${cx('derived-val', r.empty && 'empty')}>${r.empty ? '—' : r.text}</span>
      </button>`;
    })}
  </div>`;
}

const widthClass = (c) => (c.largura === 'half' ? 'w-half' : c.largura === 'third' ? 'w-third' : 'w-full');

export function FieldForm(props) {
  const { campos, grupos, valores, itens, resolver, used, onValue, onItems, onBulk, focus, onFocusField, mode, onRecords, onFieldClick } = props;
  const [showAll, setShowAll] = useState(false);
  const [collapsed, setCollapsed] = useState(() => new Set(grupos.filter((g) => g.recolhido).map((g) => g.id)));
  const rootRef = useRef();

  const visible = (c) => showAll || !used || used.has(c.id);
  const derivedOf = useMemo(() => {
    const m = new Map();
    for (const c of campos) {
      if (c.origem && (c.tipo === 'extenso' || c.tipo === 'dataExtenso' || c.tipo === 'dataCalculada')) {
        if (!m.has(c.origem)) m.set(c.origem, []);
        m.get(c.origem).push(c);
      }
    }
    return m;
  }, [campos]);

  // foco vindo do documento
  useEffect(() => {
    if (!focus || !rootRef.current) return;
    const c = campos.find((x) => x.id === focus.id);
    if (c) {
      setCollapsed((s) => { if (!s.has(c.grupo)) return s; const n = new Set(s); n.delete(c.grupo); return n; });
    }
    requestAnimationFrame(() => {
      const target = c && c.origem && isCalculado(c) ? c.origem : focus.id;
      const sel = '#f-' + target + (focus.idx !== undefined ? '-' + (focus.idx + 1) : '');
      const el = rootRef.current && (rootRef.current.querySelector(sel) || rootRef.current.querySelector('#f-' + target) || rootRef.current.querySelector(`[data-campo="${focus.id}"]`));
      if (el) {
        el.scrollIntoView({ block: 'center', behavior: 'smooth' });
        if (el.focus) el.focus({ preventScroll: true });
        const row = el.closest('.fld');
        if (row) { row.classList.remove('flash'); void row.offsetWidth; row.classList.add('flash'); }
      }
    });
  }, [focus]);

  // consultas (CNPJ/CEP) preenchem campos do mesmo grupo pelo "papel"
  const lookup = async (campo, v, idx) => {
    const same = campos.filter((c) => c.grupo === campo.grupo);
    const byPapel = (p) => same.find((c) => c.papel === p);
    const patch = {};
    if (campo.tipo === 'cep') {
      const e = await consultaCEP(v);
      const end = byPapel('endereco');
      if (end) patch[end.id] = enderecoCompleto({ logradouro: e.logradouro, numero: '', complemento: '', bairro: e.bairro, cep: e.cep, cidade: e.cidade, uf: e.uf });
      if (byPapel('cidade')) patch[byPapel('cidade').id] = e.cidade;
      if (byPapel('uf')) patch[byPapel('uf').id] = e.uf;
      if (byPapel('estado')) patch[byPapel('estado').id] = e.estado || UFS[e.uf] || '';
      if (!Object.keys(patch).length) { toast.warn(`${e.logradouro}, ${e.bairro} — ${e.cidade}-${e.uf}. Nenhum campo de endereço marcado neste grupo para preencher.`); return; }
      onBulk(patch, idx, campo.grupo);
      toast.ok(`${e.logradouro || 'Endereço'} — ${e.cidade}-${e.uf}. Complete o número.`, { title: 'Endereço encontrado' });
      return;
    }
    const d = await consultaCNPJ(v);
    const set = (papel, val) => { const f = byPapel(papel); if (f && val) patch[f.id] = val; };
    set('nome', d.nome);
    set('endereco', enderecoCompleto(d));
    set('cep', d.cep);
    set('cidade', d.cidade);
    set('uf', d.uf);
    set('estado', UFS[d.uf]);
    set('telefone', d.telefone);
    set('email', d.email);
    set('representante', d.representante);
    const prev = {};
    for (const k in patch) prev[k] = (idx !== undefined ? (itens[campo.grupo] || [])[idx]?.[k] : valores[k]) ?? '';
    onBulk(patch, idx, campo.grupo);
    const n = Object.keys(patch).length;
    toast.ok(`${d.nome}${d.situacao && d.situacao !== 'ATIVA' ? ` — situação: ${d.situacao}` : ''}. ${n} campo(s) preenchido(s)${d.representante ? ', incluindo o representante' : ''}.`, {
      title: 'Dados da Receita Federal',
      type: d.situacao && d.situacao !== 'ATIVA' ? 'warn' : 'ok',
      action: { label: 'Desfazer', onClick: () => onBulk(prev, idx, campo.grupo) },
    });
  };

  const renderField = (c, value, setValue, idx) => {
    if (c.tipo === 'indice') return null;
    const r = resolver.get(c.id, idx);
    const der = (derivedOf.get(c.id) || []).filter(visible);
    const isCalc = isCalculado(c) && c.tipo !== 'calculo';
    if (isCalc) {
      if (c.origem && campos.some((x) => x.id === c.origem && visible(x))) return null; // aparece junto da origem
      return html`<div class=${cx('fld', widthClass(c))} key=${c.id} data-campo=${c.id}>
        <label class="fld-label"><span>${c.rotulo}</span><span class="fx" title="Calculado automaticamente">fx</span></label>
        <div class=${cx('computed', r.empty && 'empty')} onClick=${() => onFieldClick && onFieldClick(c.id, idx)}>${r.empty ? (r.error || 'Aguardando dados') : r.text}</div>
      </div>`;
    }
    const required = c.obrigatorio && r.empty;
    const manual = c.tipo === 'calculo' && r.manual;
    return html`<div class=${cx('fld', widthClass(c), required && 'is-missing')} key=${c.id + (idx ?? '')}>
      <label class="fld-label" for=${'f-' + c.id + (idx !== undefined ? '-' + (idx + 1) : '')}>
        <span>${c.rotulo}${c.obrigatorio && html`<i class="req" title="Obrigatório">*</i>`}</span>
        ${c.tipo === 'calculo' && html`<span class="fx" title=${c.ajuda || 'Calculado automaticamente: ' + c.formula}>fx</span>`}
        ${manual && html`<button type="button" class="link-btn" onClick=${() => setValue('')}><${Icon} name="reset" size=${12} /> recalcular</button>`}
      </label>
      <${FieldInput} campo=${c} value=${value} resolved=${r} compact=${idx !== undefined ? idx + 1 : undefined}
        onChange=${setValue} onFocus=${() => onFocusField && onFocusField(c.id, idx)}
        onLookup=${(c.tipo === 'cnpj' || c.tipo === 'cep' || c.tipo === 'cpfCnpj') && mode !== 'campos' ? (cc, vv) => lookup(cc, vv, idx) : null} />
      ${c.ajuda && c.tipo === 'calculo' && html`<div class="fld-help">${c.ajuda}</div>`}
      <${Derived} list=${der} resolver=${resolver} idx=${idx} onFieldClick=${onFieldClick} />
    </div>`;
  };

  const toggle = (gid) => setCollapsed((s) => { const n = new Set(s); if (n.has(gid)) n.delete(gid); else n.add(gid); return n; });

  const groups = grupos.map((g) => ({ g, list: campos.filter((c) => c.grupo === g.id && visible(c)) }))
    .concat([{ g: { id: '__sem', nome: 'Outros campos', icone: 'braces' }, list: campos.filter((c) => !grupos.some((g) => g.id === c.grupo) && visible(c)) }])
    .filter((x) => x.list.length);
  const hiddenCount = used ? campos.filter((c) => !used.has(c.id) && !isCalculado(c)).length : 0;

  return html`<div class="form" ref=${rootRef}>
    ${groups.map(({ g, list }) => {
      const isCol = collapsed.has(g.id);
      const req = list.filter((c) => c.obrigatorio && !isCalculado(c));
      let filled = 0, total = 0;
      if (g.lista) {
        const its = itens[g.id] || [];
        for (const c of req) { total += Math.max(1, its.length); its.forEach((_, i) => { if (!resolver.get(c.id, i).empty) filled++; }); }
      } else for (const c of req) { total++; if (!resolver.get(c.id).empty) filled++; }
      return html`<section class=${cx('fgroup', isCol && 'is-collapsed')} key=${g.id}>
        <header class="fgroup-head" onClick=${() => toggle(g.id)}>
          <span class="fgroup-icon"><${Icon} name=${g.icone || 'braces'} size=${16} /></span>
          <h3>${g.nome}</h3>
          ${total > 0 && html`<span class=${cx('fgroup-count', filled === total && 'done')}>${filled === total ? html`<${Icon} name="check" size=${12} />` : null}${filled}/${total}</span>`}
          ${g.cadastro && onRecords && html`<span class="fgroup-actions" onClick=${(e) => e.stopPropagation()}>
            <${Menu} align="right" trigger=${({ toggle: t }) => html`<${IconButton} icon="more" title="Cadastros" onClick=${t} size=${16} />`}
              items=${[
                { header: g.cadastro },
                { icon: 'search', label: g.lista ? `Adicionar ${(g.itemNome || 'item').toLowerCase()} do cadastro` : 'Preencher com um cadastro salvo', onClick: () => onRecords('pick', g) },
                !g.lista && { icon: 'save', label: 'Salvar estes dados no cadastro', onClick: () => onRecords('save', g) },
              ]} />
          </span>`}
          <${Icon} name="chevron-down" size=${16} class="fgroup-chevron" />
        </header>
        ${!isCol && html`<div class="fgroup-body">
          ${g.lista ? html`<${ListGroup} g=${g} list=${list} itens=${itens[g.id] || []} onItems=${(arr) => onItems(g.id, arr)}
              renderField=${renderField} onRecords=${onRecords} />`
            : html`<div class="fgrid">${list.map((c) => renderField(c, valores[c.id], (v) => onValue(c.id, v)))}</div>`}
        </div>`}
      </section>`;
    })}
    ${used && hiddenCount > 0 && html`<button type="button" class="link-btn show-all" onClick=${() => setShowAll(!showAll)}>
      <${Icon} name=${showAll ? 'eye-off' : 'eye'} size=${14} />
      ${showAll ? 'Ocultar campos que não aparecem no texto' : `Mostrar ${hiddenCount} campo(s) que não aparecem no texto`}
    </button>`}
  </div>`;
}

function ListGroup({ g, list, itens, onItems, renderField, onRecords }) {
  const nome = g.itemNome || 'Item';
  const add = () => onItems([...itens, Object.fromEntries(list.filter((c) => c.padrao && c.tipo !== 'indice').map((c) => [c.id, c.padrao]))]);
  const set = (i, id, v) => onItems(itens.map((it, k) => (k === i ? { ...it, [id]: v } : it)));
  const move = (i, d) => {
    const j = i + d;
    if (j < 0 || j >= itens.length) return;
    const a = itens.slice();
    [a[i], a[j]] = [a[j], a[i]];
    onItems(a);
  };
  const shown = itens.length ? itens : [];
  return html`<div class="list-group">
    ${shown.map((it, i) => html`<div class="list-item" key=${i}>
      <div class="list-item-head">
        <span class="list-item-n">${i + 1}</span>
        <strong>${nome} ${i + 1}</strong>
        <span class="spacer" />
        <${IconButton} icon="arrow-up" title="Mover para cima" size=${15} disabled=${i === 0} onClick=${() => move(i, -1)} />
        <${IconButton} icon="arrow-down" title="Mover para baixo" size=${15} disabled=${i === itens.length - 1} onClick=${() => move(i, 1)} />
        <${IconButton} icon="copy" title="Duplicar" size=${15} onClick=${() => onItems([...itens.slice(0, i + 1), { ...it }, ...itens.slice(i + 1)])} />
        <${IconButton} icon="trash" title="Remover" size=${15} onClick=${() => onItems(itens.filter((_, k) => k !== i))} />
      </div>
      <div class="fgrid">${list.map((c) => renderField(c, it[c.id], (v) => set(i, c.id, v), i))}</div>
    </div>`)}
    ${!itens.length && html`<div class="list-empty"><${Icon} name="package" size=${18} /> Nenhum ${nome.toLowerCase()} adicionado ainda.</div>`}
    <div class="list-actions">
      <${Button} icon="plus" size="sm" onClick=${add}>Adicionar ${nome.toLowerCase()}</${Button}>
      ${g.cadastro && onRecords && html`<${Button} icon="search" size="sm" variant="ghost" onClick=${() => onRecords('pick', g)}>Do cadastro</${Button}>`}
    </div>
  </div>`;
}

export function fieldTypeLabel(t) {
  return (TIPOS[t] && TIPOS[t].nome) || t;
}

export { Badge };
