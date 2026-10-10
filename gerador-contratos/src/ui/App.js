// Estrutura do aplicativo: menu lateral, rotas e boas-vindas.
import { useEffect, useState } from 'preact/hooks';
import { html, cx, Icon, Button, ModalHost, ToastHost, useStore, openModal, ModalHead, toast } from './kit.js';
import { routeStore, navigate, match } from './router.js';
import { appConfig, applyTheme } from './state.js';
import { ensureSeed, getConfig, saveConfig, createContrato } from '../store/repo.js';
import { newContractDialog } from './dialogs.js';
import { onChange } from '../store/db.js';
import { Dashboard } from './pages/Dashboard.js';
import { Contratos } from './pages/Contratos.js';
import { Modelos } from './pages/Modelos.js';
import { Cadastros } from './pages/Cadastros.js';
import { Biblioteca } from './pages/Biblioteca.js';
import { Ajustes } from './pages/Ajustes.js';
import { Editor } from './editor/Editor.js';

const NAV = [
  { path: '/', icon: 'home', label: 'Início' },
  { path: '/contratos', icon: 'files', label: 'Contratos' },
  { path: '/modelos', icon: 'layers', label: 'Modelos' },
  { path: '/cadastros', icon: 'users', label: 'Cadastros' },
  { path: '/biblioteca', icon: 'library', label: 'Biblioteca de cláusulas' },
  { path: '/ajustes', icon: 'settings', label: 'Ajustes' },
];

let installEvt = null;
window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); installEvt = e; window.dispatchEvent(new Event('cf-install')); });

function Sidebar({ path, cfg, open, onClose }) {
  const [canInstall, setCanInstall] = useState(!!installEvt);
  useEffect(() => {
    const h = () => setCanInstall(!!installEvt);
    window.addEventListener('cf-install', h);
    return () => window.removeEventListener('cf-install', h);
  }, []);
  const active = (p) => (p === '/' ? path === '/' : path.startsWith(p));
  return html`<nav class=${cx('sidebar', open && 'is-open')} aria-label="Menu principal">
    <div class="brand" onClick=${() => navigate('/')}>
      <span class="brand-mark"><svg viewBox="0 0 32 32" width="30" height="30" aria-hidden="true"><rect x="3" y="2" width="22" height="28" rx="4" fill="currentColor" opacity=".18"/><rect x="7" y="5" width="22" height="25" rx="4" fill="currentColor"/><path d="M12 13h12M12 18h12M12 23h7" stroke="#fff" stroke-width="2.2" stroke-linecap="round"/></svg></span>
      <span class="brand-text"><strong>ContratoFácil</strong><small>${cfg.empresa || 'Gerador de contratos'}</small></span>
    </div>
    <${Button} icon="plus" variant="primary" class="nav-cta" onClick=${() => { onClose(); navigate('/novo'); }}>Novo contrato</${Button}>
    <ul class="nav">
      ${NAV.map((n) => html`<li key=${n.path}><a href=${'#' + n.path} class=${cx(active(n.path) && 'is-active')} onClick=${onClose}>
        <${Icon} name=${n.icon} size=${18} /><span>${n.label}</span></a></li>`)}
    </ul>
    <div class="nav-foot">
      ${canInstall && html`<button type="button" class="install-btn" onClick=${async () => { installEvt.prompt(); const r = await installEvt.userChoice; if (r.outcome === 'accepted') toast.ok('App instalado! Procure o ícone ContratoFácil no seu computador.'); installEvt = null; setCanInstall(false); }}>
        <${Icon} name="download" size=${16} /> Instalar como aplicativo</button>`}
      <div class="nav-safe"><${Icon} name="shield" size=${14} /> Dados salvos só neste computador</div>
    </div>
  </nav>`;
}

function welcome(cfg) {
  return openModal(({ close }) => {
    const Body = () => {
      const [nome, setNome] = useState(cfg.usuario || '');
      const [empresa, setEmpresa] = useState(cfg.empresa || '');
      return html`<form onSubmit=${(e) => { e.preventDefault(); close({ nome: nome.trim(), empresa: empresa.trim() }); }}>
        <div class="welcome-hero">
          <div class="welcome-badge"><${Icon} name="sparkles" size=${26} /></div>
          <h2>Boas-vindas ao ContratoFácil</h2>
          <p>Seus contratos com a formatação original do Word, preenchidos em minutos.</p>
        </div>
        <div class="modal-body">
          <div class="welcome-feats">
            <div><${Icon} name="text-input" size=${18} /><span><b>Preencha só os campos</b> — nome, CNPJ, valores, datas. Valores por extenso e datas de término são calculados sozinhos.</span></div>
            <div><${Icon} name="search" size=${18} /><span><b>Digite o CNPJ</b> e o sistema busca razão social, endereço e representante na Receita.</span></div>
            <div><${Icon} name="edit" size=${18} /><span><b>Edite qualquer cláusula</b> clicando no texto. A numeração se ajusta sozinha.</span></div>
            <div><${Icon} name="file-down" size=${18} /><span><b>Baixe em Word ou PDF</b> com 100% da formatação do seu modelo.</span></div>
          </div>
          <div class="grid2 mt12">
            <label class="field"><span class="field-label">Como podemos te chamar?</span><input class="input" autoFocus value=${nome} placeholder="Seu nome" onInput=${(e) => setNome(e.currentTarget.value)} /></label>
            <label class="field"><span class="field-label">Empresa</span><input class="input" value=${empresa} onInput=${(e) => setEmpresa(e.currentTarget.value)} /></label>
          </div>
        </div>
        <div class="modal-foot"><${Button} variant="primary" type="submit" iconRight="arrow-right">Começar</${Button}></div>
      </form>`;
    };
    return html`<${Body} />`;
  }, { size: 'md', persistent: true });
}

function NovoRedirect() {
  useEffect(() => {
    (async () => {
      const q = new URLSearchParams((location.hash.split('?')[1]) || '');
      const r = await newContractDialog({ modeloId: q.get('modelo') || undefined });
      if (!r) { navigate('/contratos'); return; }
      try {
        const c = await createContrato(r.modeloId, { fromId: r.fromId });
        navigate('/contrato/' + c.id);
      } catch (e) { toast.err(e.message); navigate('/contratos'); }
    })();
  }, []);
  return html`<div class="page-center"><span class="spinner lg" /></div>`;
}

export function App() {
  const { path } = useStore(routeStore);
  const cfg = useStore(appConfig);
  const [ready, setReady] = useState(false);
  const [fatal, setFatal] = useState(null);
  const [menu, setMenu] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        await ensureSeed();
        const c = await getConfig();
        appConfig.set(c);
        applyTheme(c.tema);
        setReady(true);
        if (!c.onboarded) {
          const r = await welcome(c);
          const next = await saveConfig({ onboarded: true, usuario: r ? r.nome : '', empresa: (r && r.empresa) || c.empresa });
          appConfig.set(next);
        }
      } catch (e) {
        console.error(e);
        setFatal(e.message || String(e));
      }
    })();
    const off = onChange(async (store) => { if (store === 'config') appConfig.set(await getConfig()); });
    const mq = window.matchMedia ? matchMedia('(prefers-color-scheme: dark)') : null;
    const onMq = () => applyTheme(appConfig.get().tema);
    if (mq && mq.addEventListener) mq.addEventListener('change', onMq);
    return () => { off(); if (mq && mq.removeEventListener) mq.removeEventListener('change', onMq); };
  }, []);

  useEffect(() => { setMenu(false); window.scrollTo(0, 0); }, [path]);

  if (fatal) return html`<div class="page-center"><div class="empty"><div class="empty-icon"><${Icon} name="alert" size=${28} /></div><h3>Não foi possível iniciar</h3><p>${fatal}</p></div></div>`;
  if (!ready) return html`<div class="page-center"><span class="spinner lg" /></div>`;

  let page;
  let full = false;
  let m;
  if ((m = match('/contrato/:id', path))) { page = html`<${Editor} key=${'c' + m.id} kind="contrato" id=${m.id} />`; full = true; }
  else if ((m = match('/modelo/:id', path))) { page = html`<${Editor} key=${'m' + m.id} kind="modelo" id=${m.id} />`; full = true; }
  else if (path.startsWith('/contratos')) page = html`<${Contratos} />`;
  else if (path.startsWith('/modelos')) page = html`<${Modelos} />`;
  else if (path.startsWith('/cadastros')) page = html`<${Cadastros} />`;
  else if (path.startsWith('/biblioteca')) page = html`<${Biblioteca} />`;
  else if (path.startsWith('/ajustes')) page = html`<${Ajustes} />`;
  else if (path.startsWith('/novo')) page = html`<${NovoRedirect} key=${Date.now()} />`;
  else page = html`<${Dashboard} />`;

  return html`<div class=${cx('app', full && 'is-full')}>
    ${!full && html`<${Sidebar} path=${path} cfg=${cfg} open=${menu} onClose=${() => setMenu(false)} />`}
    ${!full && html`<header class="mobile-bar"><button type="button" class="icon-btn" aria-label="Menu" onClick=${() => setMenu(!menu)}><${Icon} name="menu" /></button><strong>ContratoFácil</strong></header>`}
    ${menu && html`<div class="scrim" onClick=${() => setMenu(false)} />`}
    <div class="main">${page}</div>
    <${ModalHost} />
    <${ToastHost} />
  </div>`;
}

export { ModalHead };
