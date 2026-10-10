// Modelos de contrato: importar do Word, editar, duplicar, baixar.
import { useState } from 'preact/hooks';
import { html, cx, Icon, Button, IconButton, Menu, Empty, toast, confirmDialog, promptDialog, downloadBlob, relTime, pickFile } from '../kit.js';
import { listModelos, listContratos, importModelo, deleteModelo, duplicateModelo, createContrato, saveModelo } from '../../store/repo.js';
import { navigate } from '../router.js';
import { useCollection } from './Dashboard.js';
import { buildDocx } from '../../export.js';

export function Modelos() {
  const modelos = useCollection(listModelos);
  const contratos = useCollection(listContratos);
  const [drag, setDrag] = useState(false);
  const [busy, setBusy] = useState(false);
  if (!modelos || !contratos) return html`<div class="page-center"><span class="spinner lg" /></div>`;

  const importar = async (file) => {
    if (!file) return;
    if (!/\.docx$/i.test(file.name)) { toast.err('Envie um arquivo do Word no formato .docx (no Word: Arquivo › Salvar como › Documento do Word).'); return; }
    setBusy(true);
    try {
      const m = await importModelo(file);
      toast.ok(`Modelo “${m.nome}” importado${m.campos.length ? ` com ${m.campos.length} campo(s) {{…}}` : ''}. Agora marque os campos.`, { title: 'Pronto!' });
      navigate('/modelo/' + m.id);
    } catch (e) {
      console.error(e);
      toast.err('Não consegui ler este arquivo: ' + e.message);
    } finally { setBusy(false); }
  };

  const onDrop = (e) => {
    e.preventDefault();
    setDrag(false);
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    importar(f);
  };

  return html`<div class="page" onDragOver=${(e) => { e.preventDefault(); setDrag(true); }} onDragLeave=${(e) => { if (e.currentTarget === e.target) setDrag(false); }} onDrop=${onDrop}>
    <header class="page-head">
      <div><h1>Modelos</h1><p class="muted">A base de cada contrato — com a formatação do seu Word.</p></div>
      <div class="head-actions">
        <${Button} icon="upload" variant="primary" loading=${busy} onClick=${async () => importar(await pickFile('.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document'))}>Importar do Word</${Button}>
      </div>
    </header>

    <div class="model-grid">
      ${modelos.map((m) => {
        const n = contratos.filter((c) => c.modeloId === m.id).length;
        return html`<article class="model-card" key=${m.id}>
          <div class="model-top" style=${{ '--mc': m.cor || '#2563eb' }}>
            <div class="model-paper"><i /><i /><b /><i /><i /><b /><i /></div>
            <${Menu} align="right" class="model-menu" trigger=${({ toggle }) => html`<${IconButton} icon="more" title="Mais" onClick=${toggle} />`} items=${[
              { icon: 'pencil', label: 'Editar modelo', onClick: () => navigate('/modelo/' + m.id) },
              { icon: 'copy', label: 'Duplicar modelo', onClick: async () => { const c = await duplicateModelo(m.id); toast.ok('Modelo duplicado.'); navigate('/modelo/' + c.id); } },
              { icon: 'download', label: 'Baixar .docx com {{campos}}', onClick: async () => downloadBlob(await buildDocx({ pkgBuf: m.pkg, docXml: m.docXml, resolver: null, titulo: m.nome }), `Modelo - ${m.nome}.docx`) },
              { icon: 'pencil', label: 'Renomear', onClick: async () => { const nome = await promptDialog({ title: 'Renomear modelo', label: 'Nome', value: m.nome }); if (nome) { m.nome = nome; await saveModelo(m); } } },
              '-',
              { icon: 'trash', label: 'Excluir modelo', danger: true, onClick: async () => {
                if (await confirmDialog({ title: `Excluir “${m.nome}”?`, message: n ? `Os ${n} contrato(s) já criados continuam funcionando normalmente.` : 'Esta ação não pode ser desfeita.', ok: 'Excluir', danger: true })) { await deleteModelo(m.id); toast('Modelo excluído.'); }
              } },
            ]} />
          </div>
          <div class="model-body">
            <h3>${m.nome}</h3>
            ${m.descricao && html`<p class="muted small">${m.descricao}</p>`}
            <div class="model-stats"><span><${Icon} name="braces" size=${14} /> ${m.campos.length} campos</span><span><${Icon} name="files" size=${14} /> ${n} contrato(s)</span><span><${Icon} name="clock" size=${14} /> ${relTime(m.atualizadoEm)}</span></div>
          </div>
          <div class="model-foot">
            <${Button} icon="plus" variant="primary" size="sm" onClick=${async () => { const c = await createContrato(m.id); navigate('/contrato/' + c.id); }}>Novo contrato</${Button}>
            <${Button} icon="pencil" size="sm" variant="ghost" onClick=${() => navigate('/modelo/' + m.id)}>Editar</${Button}>
          </div>
        </article>`;
      })}
      <button type="button" class=${cx('model-card add', drag && 'is-drag')} onClick=${async () => importar(await pickFile('.docx'))}>
        <span class="add-ic"><${Icon} name="upload" size=${26} /></span>
        <strong>Importar outro modelo</strong>
        <span class="muted small">Arraste um arquivo .docx aqui ou clique para escolher. Toda a formatação do Word é mantida.</span>
      </button>
    </div>

    <section class="card how">
      <header class="card-head"><h2><${Icon} name="lightbulb" size=${18} /> Como transformar um contrato do Word em modelo</h2></header>
      <div class="how-grid">
        <div><span class="how-n">1</span><p><b>Importe o .docx</b> do contrato que você já usa — pode ser um contrato preenchido.</p></div>
        <div><span class="how-n">2</span><p>Clique em <b>Detectar campos</b>: CPF, CNPJ, valores, datas e nomes viram campos automaticamente.</p></div>
        <div><span class="how-n">3</span><p>Para o resto, <b>selecione o texto</b> no documento e clique em <b>Transformar em campo</b>.</p></div>
        <div><span class="how-n">4</span><p>Pronto! Crie contratos novos só preenchendo o formulário.</p></div>
      </div>
      <p class="muted small">Também funciona escrever no próprio Word <code>{{nome_do_cliente}}</code> onde quiser um campo: ao importar, eles são reconhecidos.</p>
    </section>
    ${drag && html`<div class="drop-overlay"><${Icon} name="upload" size=${36} /><strong>Solte o arquivo .docx para importar</strong></div>`}
    ${!modelos.length && html`<${Empty} icon="layers" title="Nenhum modelo">Importe um contrato do Word para começar.</${Empty}>`}
  </div>`;
}
