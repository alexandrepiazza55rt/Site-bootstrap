// Ajustes: perfil, aparência, numeração, backup e dados.
import { useEffect, useState } from 'preact/hooks';
import { html, cx, Icon, Button, useStore, toast, confirmDialog, downloadBlob, pickFile } from '../kit.js';
import { appConfig, applyTheme } from '../state.js';
import { saveConfig, exportBackup, importBackup, wipeAll, formatNumero } from '../../store/repo.js';
import { storageEstimate, persistStorage } from '../../store/db.js';

/* global __VERSION__ */
const VERSION = typeof __VERSION__ !== 'undefined' ? __VERSION__ : 'dev';

export function Ajustes() {
  const cfg = useStore(appConfig);
  const [est, setEst] = useState(null);
  const [persist, setPersist] = useState(null);
  const ano = String(new Date().getFullYear());
  useEffect(() => {
    storageEstimate().then(setEst);
    if (navigator.storage && navigator.storage.persisted) navigator.storage.persisted().then(setPersist);
  }, []);
  const upd = async (patch) => { const n = await saveConfig(patch); appConfig.set(n); if (patch.tema) applyTheme(patch.tema); };
  const prox = (cfg.seq && cfg.seq[ano]) || 1;

  return html`<div class="page narrow">
    <header class="page-head"><div><h1>Ajustes</h1><p class="muted">Preferências deste computador.</p></div></header>

    <section class="card set">
      <h2><${Icon} name="user" size=${18} /> Perfil</h2>
      <div class="grid2">
        <label class="field"><span class="field-label">Seu nome</span><input class="input" value=${cfg.usuario || ''} onChange=${(e) => upd({ usuario: e.currentTarget.value })} /></label>
        <label class="field"><span class="field-label">Empresa</span><input class="input" value=${cfg.empresa || ''} onChange=${(e) => upd({ empresa: e.currentTarget.value })} /></label>
      </div>
    </section>

    <section class="card set">
      <h2><${Icon} name="monitor" size=${18} /> Aparência</h2>
      <div class="seg-tabs w-auto">
        ${[['auto', 'Automático', 'monitor'], ['claro', 'Claro', 'sun'], ['escuro', 'Escuro', 'moon']].map(([k, n, ic]) => html`
          <button type="button" class=${cx((cfg.tema || 'auto') === k && 'is-active')} onClick=${() => upd({ tema: k })}><${Icon} name=${ic} size=${15} /> ${n}</button>`)}
      </div>
      <label class="switch-row"><input type="checkbox" checked=${cfg.destacarCampos !== false} onChange=${(e) => upd({ destacarCampos: e.currentTarget.checked })} /><span class="switch" /><span>Destacar campos preenchíveis no documento</span></label>
    </section>

    <section class="card set">
      <h2><${Icon} name="hash" size=${18} /> Numeração dos contratos</h2>
      <p class="muted small">Cada contrato novo recebe o próximo número automaticamente. O número volta para 1 a cada ano.</p>
      <div class="grid2">
        <label class="field"><span class="field-label">Próximo número em ${ano}</span>
          <input class="input" type="number" min="1" value=${prox} onChange=${(e) => { const v = Math.max(1, parseInt(e.currentTarget.value, 10) || 1); upd({ seq: { ...cfg.seq, [ano]: v } }); toast.ok('Próximo contrato: ' + formatNumero('{n:4}/{ano}', v, ano)); }} /></label>
        <div class="field"><span class="field-label">Próximo contrato sairá como</span><div class="computed big">${formatNumero('{n:4}/{ano}', prox, ano)}</div></div>
      </div>
    </section>

    <section class="card set">
      <h2><${Icon} name="database" size=${18} /> Backup</h2>
      <p class="muted small">Seus contratos ficam salvos <b>somente neste navegador</b>. Faça backup com frequência (ex.: toda sexta) e guarde o arquivo no Google Drive, OneDrive ou pen drive. O backup também serve para levar tudo para outro computador.</p>
      <div class="btn-row">
        <${Button} icon="download" variant="primary" onClick=${async () => { downloadBlob(await exportBackup(), `backup-contratofacil-${new Date().toISOString().slice(0, 10)}.json`); upd({ ultimoBackup: new Date().toISOString() }); toast.ok('Backup baixado.'); }}>Fazer backup agora</${Button}>
        <${Button} icon="upload" onClick=${async () => {
          const f = await pickFile('.json,application/json');
          if (!f) return;
          const subst = await confirmDialog({ title: 'Restaurar backup', message: 'Deseja SUBSTITUIR tudo pelo backup? Escolha “Mesclar” para apenas adicionar o que falta.', ok: 'Substituir tudo', cancel: 'Mesclar' });
          try { const n = await importBackup(f, { substituir: !!subst }); toast.ok(`Backup restaurado: ${n.contratos} contrato(s), ${n.modelos} modelo(s).`); setTimeout(() => location.reload(), 900); } catch (e) { toast.err(e.message); }
        }}>Restaurar backup</${Button}>
      </div>
      ${cfg.ultimoBackup && html`<p class="muted small mt8"><${Icon} name="check-circle" size=${13} /> Último backup: ${new Date(cfg.ultimoBackup).toLocaleString('pt-BR')}</p>`}
    </section>

    <section class="card set">
      <h2><${Icon} name="shield" size=${18} /> Privacidade e armazenamento</h2>
      <ul class="plain-list">
        <li><${Icon} name="lock" size=${15} /> Nenhum contrato é enviado para a internet. Só as consultas de CNPJ/CEP (BrasilAPI e ViaCEP) usam a internet, enviando apenas o número consultado.</li>
        <li><${Icon} name="drive" size=${15} /> Espaço usado: ${est && est.usage ? (est.usage / 1048576).toFixed(1) + ' MB' : '—'}${est && est.quota ? ' de ' + Math.round(est.quota / 1048576 / 1024) + ' GB disponíveis' : ''}.</li>
        <li><${Icon} name=${persist ? 'check-circle' : 'info'} size=${15} /> ${persist ? 'Armazenamento protegido: o navegador não apaga os dados sozinho.' : html`O navegador pode limpar dados se faltar espaço. <button type="button" class="link-btn" onClick=${async () => { const ok = await persistStorage(); setPersist(ok); toast[ok ? 'ok' : 'warn'](ok ? 'Armazenamento protegido.' : 'O navegador não permitiu. Instalar como aplicativo ajuda.'); }}>Proteger armazenamento</button>`}</li>
      </ul>
      <div class="btn-row mt8">
        <${Button} icon="trash" variant="danger-ghost" onClick=${async () => {
          if (!(await confirmDialog({ title: 'Apagar TODOS os dados?', message: 'Contratos, modelos, cadastros e biblioteca serão apagados deste computador. Faça um backup antes!', ok: 'Apagar tudo', danger: true }))) return;
          if (!(await confirmDialog({ title: 'Tem certeza mesmo?', message: 'Não será possível desfazer.', ok: 'Sim, apagar', danger: true }))) return;
          await wipeAll(); location.hash = '/'; location.reload();
        }}>Apagar todos os dados</${Button}>
      </div>
    </section>

    <section class="card set">
      <h2><${Icon} name="keyboard" size=${18} /> Atalhos do editor</h2>
      <div class="kbd-grid">
        <span><kbd>Ctrl</kbd>+<kbd>Z</kbd></span><span>Desfazer alteração no texto</span>
        <span><kbd>Ctrl</kbd>+<kbd>Y</kbd></span><span>Refazer</span>
        <span><kbd>Ctrl</kbd>+<kbd>S</kbd></span><span>Salvar agora (já salva sozinho)</span>
        <span><kbd>Ctrl</kbd>+<kbd>P</kbd></span><span>Gerar PDF / imprimir</span>
        <span><kbd>Ctrl</kbd>+<kbd>B</kbd> <kbd>I</kbd> <kbd>U</kbd></span><span>Negrito, itálico, sublinhado (no editor de cláusula)</span>
        <span><kbd>Enter</kbd></span><span>Novo parágrafo com a mesma formatação</span>
        <span><kbd>Esc</kbd></span><span>Fechar o painel da cláusula</span>
      </div>
    </section>
    <p class="muted small center">ContratoFácil · versão ${VERSION}</p>
  </div>`;
}
