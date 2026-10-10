// Início: visão geral, vencimentos e contratos recentes.
import { useEffect, useState } from 'preact/hooks';
import { html, cx, Icon, Button, Badge, useStore, relTime, brl, dataBR } from '../kit.js';
import { appConfig } from '../state.js';
import { listContratos, listModelos, tituloContrato } from '../../store/repo.js';
import { onChange } from '../../store/db.js';
import { navigate } from '../router.js';
import { statusInfo } from '../editor/Editor.js';
import { isoHoje, diasEntre } from '../../br/docs.js';

export function useCollection(loader) {
  const [data, setData] = useState(null);
  useEffect(() => {
    let alive = true;
    const load = () => loader().then((d) => alive && setData(d));
    load();
    const off = onChange(() => load());
    return () => { alive = false; off(); };
  }, []);
  return data;
}

function saudacao() {
  const h = new Date().getHours();
  return h < 12 ? 'Bom dia' : h < 18 ? 'Boa tarde' : 'Boa noite';
}

export function prazoBadge(termino) {
  if (!termino) return null;
  const d = diasEntre(isoHoje(), termino);
  if (d === null) return null;
  if (d < 0) return html`<${Badge} tone="red" icon="alert">venceu há ${-d} dia${-d === 1 ? '' : 's'}</${Badge}>`;
  if (d === 0) return html`<${Badge} tone="red" icon="alert">vence hoje</${Badge}>`;
  if (d <= 30) return html`<${Badge} tone="amber" icon="clock">vence em ${d} dia${d === 1 ? '' : 's'}</${Badge}>`;
  return html`<${Badge} tone="gray">até ${dataBR(termino)}</${Badge}>`;
}

export function Dashboard() {
  const cfg = useStore(appConfig);
  const contratos = useCollection(listContratos);
  const modelos = useCollection(listModelos);
  if (!contratos || !modelos) return html`<div class="page-center"><span class="spinner lg" /></div>`;

  const hoje = isoHoje();
  const vigentes = contratos.filter((c) => ['finalizado', 'assinado'].includes(c.status) && (!c.resumo?.termino || c.resumo.termino >= hoje));
  const valor = vigentes.reduce((a, c) => a + (c.resumo?.valor || 0), 0);
  const rascunhos = contratos.filter((c) => c.status === 'rascunho');
  const mes = hoje.slice(0, 7);
  const doMes = contratos.filter((c) => (c.criadoEm || '').slice(0, 7) === mes);
  const venc = contratos
    .filter((c) => c.resumo?.termino && !['cancelado', 'encerrado'].includes(c.status))
    .map((c) => ({ c, d: diasEntre(hoje, c.resumo.termino) }))
    .filter((x) => x.d !== null && x.d <= 60 && x.d >= -45)
    .sort((a, b) => a.d - b.d);
  const recentes = contratos.slice(0, 6);

  return html`<div class="page">
    <header class="page-head">
      <div>
        <h1>${saudacao()}${cfg.usuario ? ', ' + cfg.usuario.split(' ')[0] : ''}!</h1>
        <p class="muted">${new Date().toLocaleDateString('pt-BR', { weekday: 'long', day: 'numeric', month: 'long' })} · ${cfg.empresa || 'ContratoFácil'}</p>
      </div>
      <div class="head-actions">
        <${Button} icon="plus" variant="primary" size="lg" onClick=${() => navigate('/novo')}>Novo contrato</${Button}>
      </div>
    </header>

    <section class="kpis">
      <div class="kpi"><span class="kpi-ic blue"><${Icon} name="file-check" size=${20} /></span><div><strong>${vigentes.length}</strong><span>contratos vigentes</span></div></div>
      <div class="kpi"><span class="kpi-ic green"><${Icon} name="coins" size=${20} /></span><div><strong>${brl(valor)}</strong><span>em contratos vigentes</span></div></div>
      <div class="kpi"><span class="kpi-ic amber"><${Icon} name="calendar-clock" size=${20} /></span><div><strong>${venc.filter((x) => x.d >= 0 && x.d <= 30).length}</strong><span>vencem em 30 dias</span></div></div>
      <div class="kpi"><span class="kpi-ic violet"><${Icon} name="pencil" size=${20} /></span><div><strong>${rascunhos.length}</strong><span>rascunhos · ${doMes.length} criados este mês</span></div></div>
    </section>

    ${!contratos.length && html`<section class="card hero-card">
      <div class="hero-text">
        <h2>Vamos criar seu primeiro contrato?</h2>
        <p>O modelo <b>${modelos[0] ? modelos[0].nome : 'de locação'}</b> já está pronto, com todos os campos configurados.</p>
        <ol class="steps">
          <li><span>1</span> Clique em <b>Novo contrato</b></li>
          <li><span>2</span> Digite o CNPJ do cliente e clique em <b>Buscar dados</b></li>
          <li><span>3</span> Informe equipamento, valores e datas — o resto é automático</li>
          <li><span>4</span> Baixe em <b>Word</b> ou <b>PDF</b></li>
        </ol>
        <${Button} icon="plus" variant="primary" onClick=${() => navigate('/novo')}>Novo contrato</${Button}>
      </div>
      <div class="hero-art" aria-hidden="true"><div class="paper p1" /><div class="paper p2"><i /><i /><i /><b /><i /><i /></div></div>
    </section>`}

    <div class="dash-grid">
      <section class="card">
        <header class="card-head"><h2><${Icon} name="calendar-clock" size=${18} /> Vencimentos</h2><span class="muted small">próximos 60 dias</span></header>
        ${venc.length ? html`<ul class="rows">
          ${venc.slice(0, 8).map(({ c }) => html`<li key=${c.id} class="row" onClick=${() => navigate('/contrato/' + c.id)}>
            <span class="row-ic" style=${{ background: c.cor || '#2563eb' }}><${Icon} name="file-text" size=${15} /></span>
            <div class="row-main"><strong>${tituloContrato(c)}</strong><small>${c.modeloNome} · término ${dataBR(c.resumo.termino)}</small></div>
            ${prazoBadge(c.resumo.termino)}
          </li>`)}
        </ul>` : html`<div class="empty small"><div class="empty-icon"><${Icon} name="calendar" size=${22} /></div><p>Nenhum contrato vencendo nos próximos 60 dias.</p></div>`}
      </section>

      <section class="card">
        <header class="card-head"><h2><${Icon} name="history" size=${18} /> Recentes</h2><a class="link-btn" href="#/contratos">Ver todos <${Icon} name="arrow-right" size=${14} /></a></header>
        ${recentes.length ? html`<ul class="rows">
          ${recentes.map((c) => {
            const st = statusInfo(c.status);
            return html`<li key=${c.id} class="row" onClick=${() => navigate('/contrato/' + c.id)}>
              <span class="row-ic" style=${{ background: c.cor || '#2563eb' }}><${Icon} name="file-text" size=${15} /></span>
              <div class="row-main"><strong>${tituloContrato(c)}</strong><small>${c.resumo?.valorTexto ? 'R$ ' + c.resumo.valorTexto + ' · ' : ''}editado ${relTime(c.atualizadoEm)}</small></div>
              <${Badge} tone=${st[2]}>${st[1]}</${Badge}>
            </li>`;
          })}
        </ul>` : html`<div class="empty small"><div class="empty-icon"><${Icon} name="files" size=${22} /></div><p>Seus contratos aparecerão aqui.</p></div>`}
      </section>

      <section class="card">
        <header class="card-head"><h2><${Icon} name="layers" size=${18} /> Modelos</h2><a class="link-btn" href="#/modelos">Gerenciar <${Icon} name="arrow-right" size=${14} /></a></header>
        <ul class="rows">
          ${modelos.slice(0, 5).map((m) => html`<li key=${m.id} class="row" onClick=${() => navigate('/novo?modelo=' + m.id)}>
            <span class="row-ic" style=${{ background: m.cor || '#2563eb' }}><${Icon} name="layers" size=${15} /></span>
            <div class="row-main"><strong>${m.nome}</strong><small>${m.campos.length} campos · ${contratos.filter((c) => c.modeloId === m.id).length} contrato(s)</small></div>
            <span class="row-go"><${Icon} name="plus" size=${16} /></span>
          </li>`)}
        </ul>
      </section>

      <section class="card tips">
        <header class="card-head"><h2><${Icon} name="lightbulb" size=${18} /> Dicas rápidas</h2></header>
        <ul class="tip-list">
          <li><${Icon} name="search" size=${15} /> No campo CNPJ, clique em <b>Buscar dados</b>: razão social, endereço e sócio-administrador são preenchidos automaticamente.</li>
          <li><${Icon} name="click" size=${15} /> Clique em qualquer cláusula do documento para editá-la só naquele contrato.</li>
          <li><${Icon} name="copy-plus" size=${15} /> Renovação? Abra o contrato antigo e use <b>Duplicar contrato</b>.</li>
          <li><${Icon} name="library" size=${15} /> Guarde cláusulas especiais na <b>Biblioteca</b> e reutilize em qualquer contrato.</li>
          <li><${Icon} name="keyboard" size=${15} /> Atalhos: <kbd>Ctrl</kbd>+<kbd>Z</kbd> desfaz, <kbd>Ctrl</kbd>+<kbd>P</kbd> gera o PDF.</li>
        </ul>
      </section>
    </div>
  </div>`;
}

export { cx };
