// Editor de contrato / modelo: formulário + páginas + inspetor, com desfazer, salvamento automático e exportação.
import { useEffect, useMemo, useRef, useState, useCallback } from 'preact/hooks';
import { html, cx, Icon, IconButton, Button, Menu, toast, confirmDialog, promptDialog, downloadBlob, useDebounced, Badge, openModal, ModalHead } from '../kit.js';
import { DocContext } from '../../docx/context.js';
import { parseXml, serializeXml, getParaId, setParaId, NS, isW, wKid, wAll } from '../../docx/xml.js';
import { parseParagraph, writeParagraph, segsText, dominantRPr } from '../../docx/segments.js';
import {
  documentKinds, prototypeFor, newParagraphLike, insertAfter, duplicate, removeBlock, moveBlock, convertKind, setAlign,
  setPageBreakBefore, sectionRange, rank, kindOf, diffMarks, isTopLevel, topBlocks,
} from '../../docx/ops.js';
import { findListUnits } from '../../docx/fill.js';
import { FieldResolver, usedFieldIds } from '../../fields/compute.js';
import { isCalculado } from '../../fields/types.js';
import { DocCanvas } from './DocCanvas.js';
import { FieldForm } from './FieldForm.js';
import { Inspector, numberLabel } from './Inspector.js';
import { Outline } from './Outline.js';
import {
  FieldsManager, fieldDialog, groupDialog, createFromSelectionDialog, detectDialog, applySuggestions, applyFieldRanges,
} from './TemplateTools.js';
import { pickRecordDialog, libraryPickDialog, missingDialog, pdfTipDialog, newContractDialog } from '../dialogs.js';
import {
  getContrato, getModelo, saveContrato, saveModelo, deleteContrato, createContrato, nomeArquivo, saveCadastro,
  saveClausula, getConfig, saveConfig, tituloContrato, resumoContrato,
} from '../../store/repo.js';
import { buildDocx, printDocument } from '../../export.js';
import { navigate } from '../router.js';
import { appConfig } from '../state.js';

const STATUS = [
  ['rascunho', 'Rascunho', 'gray'], ['finalizado', 'Finalizado', 'blue'], ['assinado', 'Assinado', 'green'],
  ['encerrado', 'Encerrado', 'violet'], ['cancelado', 'Cancelado', 'red'],
];
export const statusInfo = (s) => STATUS.find((x) => x[0] === s) || STATUS[0];

export function Editor({ kind, id }) {
  const isTpl = kind === 'modelo';
  const [rec, setRec] = useState(null);
  const [error, setError] = useState(null);
  const [ctx, setCtx] = useState(null);
  const [doc, setDoc] = useState(null);
  const [docVer, setDocVer] = useState(0);
  const [extVer, setExtVer] = useState(0);
  const [valores, setValores] = useState({});
  const [itens, setItens] = useState({});
  const [campos, setCampos] = useState([]);
  const [grupos, setGrupos] = useState([]);
  const [selected, setSelected] = useState(null);
  const [focusText, setFocusText] = useState(false);
  const [focusField, setFocusField] = useState(null);
  const [formFocus, setFormFocus] = useState(null);
  const [leftTab, setLeftTab] = useState(isTpl ? 'campos' : 'preencher');
  const [leftOpen, setLeftOpen] = useState(true);
  const [zoom, setZoom] = useState('auto');
  const [zoomVal, setZoomVal] = useState(1);
  const [highlight, setHighlight] = useState(true);
  const [showChanges, setShowChanges] = useState(false);
  const [saveState, setSaveState] = useState('saved');
  const [pages, setPages] = useState(0);
  const [textSel, setTextSel] = useState(null);
  const [tplDoc, setTplDoc] = useState(null);
  const [busy, setBusy] = useState(null);
  const undo = useRef({ undo: [], redo: [] });
  const [, setUndoTick] = useState(0);
  const editSession = useRef(null);
  const latest = useRef({});
  latest.current = { rec, doc, valores, itens, campos, grupos };

  // ------------------------------------------------------------- carregar
  useEffect(() => {
    let alive = true;
    let c = null;
    (async () => {
      try {
        const r = isTpl ? await getModelo(id) : await getContrato(id);
        if (!r) throw new Error(isTpl ? 'Modelo não encontrado.' : 'Contrato não encontrado.');
        c = await DocContext.fromBuffer(r.pkg);
        if (!alive) return;
        const d = parseXml(r.docXml);
        setRec(r); setCtx(c); setDoc(d);
        setValores(isTpl ? r.testeValores || {} : r.valores || {});
        setItens(isTpl ? r.testeItens || {} : r.itens || {});
        setCampos(r.campos || []); setGrupos(r.grupos || []);
        if (!isTpl && r.modeloId) getModelo(r.modeloId).then((m) => { if (alive && m) setTplDoc(parseXml(m.docXml)); });
        const cfg = await getConfig();
        setHighlight(cfg.destacarCampos !== false);
      } catch (e) {
        console.error(e);
        if (alive) setError(e.message || String(e));
      }
    })();
    return () => { alive = false; if (c) c.dispose(); };
  }, [kind, id]);

  // ------------------------------------------------------------- salvar
  const doSave = useCallback(async (acao) => {
    const L = latest.current;
    if (!L.rec || !L.doc) return;
    setSaveState('saving');
    try {
      L.rec.docXml = serializeXml(L.doc);
      L.rec.campos = L.campos;
      L.rec.grupos = L.grupos;
      if (isTpl) { L.rec.testeValores = L.valores; L.rec.testeItens = L.itens; await saveModelo(L.rec); }
      else { L.rec.valores = L.valores; L.rec.itens = L.itens; await saveContrato(L.rec, acao); }
      setSaveState('saved');
    } catch (e) {
      console.error(e);
      setSaveState('error');
      toast.err('Não foi possível salvar: ' + e.message);
    }
  }, [isTpl]);
  const save = useDebounced(doSave, 700);
  const touch = () => { setSaveState('dirty'); save(); };
  useEffect(() => () => { if (latest.current.rec) save.flush(); }, []);
  useEffect(() => {
    const h = (e) => { if (saveState !== 'saved') { save.flush(); e.preventDefault(); e.returnValue = ''; } };
    window.addEventListener('beforeunload', h);
    return () => window.removeEventListener('beforeunload', h);
  }, [saveState]);

  // ------------------------------------------------------------- derivados
  const resolver = useMemo(() => new FieldResolver({ campos, grupos }, valores, itens), [campos, grupos, valores, itens]);
  const resumo = useMemo(() => (rec && !isTpl ? resumoContrato({ ...rec, valores, itens, campos, grupos }) : null), [rec, valores, itens, campos, grupos]);
  const [canvasState, setCanvasState] = useState(null);
  const pushCanvas = useDebounced((s) => setCanvasState(s), 140);
  const xmlNow = useMemo(() => (doc ? serializeXml(doc) : ''), [doc, docVer]);
  const used = useMemo(() => usedFieldIds({ campos }, xmlNow), [campos, xmlNow]);
  const baseMap = useMemo(() => {
    if (isTpl || !rec) return null;
    return new Map(Object.entries(rec.baseTextos || {}));
  }, [rec]);
  const diff = useMemo(() => (baseMap && doc ? diffMarks(doc, baseMap) : { marks: new Map(), removed: 0 }), [baseMap, doc, docVer]);
  useEffect(() => {
    if (!doc) return;
    const listUnits = findListUnits(doc, resolver);
    const resolve = (tok, itemIdx) => {
      const m = /^([A-Za-z0-9_]+)(?:@(\d+))?$/.exec(tok);
      const c = m && resolver.byId.get(m[1]);
      if (!m || !c) return { text: '{{' + tok + '}}', empty: true, label: '⚠ ' + tok + ' (campo inexistente)' };
      const idx = m[2] !== undefined ? +m[2] : resolver.isList(c) ? itemIdx : undefined;
      const r = resolver.get(m[1], idx);
      if (r.empty) return { ...r, label: c.rotulo };
      return r;
    };
    const st = { resolve, listUnits, marks: diff.marks, ver: docVer, doc };
    if (!canvasState || canvasState.doc !== doc) setCanvasState(st);
    else pushCanvas(st);
  }, [doc, docVer, resolver, diff]);

  const kinds = useMemo(() => {
    if (!ctx || !doc) return [];
    const ks = documentKinds(ctx, doc).filter((k) => k.level !== 99 || true);
    // exemplo de numeração para cada tipo
    for (const k of ks) {
      const proto = prototypeFor(ctx, doc, k.key, null);
      k.sample = proto ? numberLabel(ctx, doc, proto) : '';
    }
    return ks;
  }, [ctx, doc, docVer]);

  // ------------------------------------------------------------- desfazer
  const pushUndo = (label) => {
    const u = undo.current;
    u.undo.push({ xml: serializeXml(latest.current.doc), sel: selected && isW(selected, 'p') ? getParaId(selected) : null, label });
    if (u.undo.length > 80) u.undo.shift();
    u.redo = [];
    setUndoTick((n) => n + 1);
  };
  const changed = ({ external = false, acao } = {}) => {
    setDocVer((v) => v + 1);
    if (external) setExtVer((v) => v + 1);
    setSaveState('dirty');
    if (acao) { save.cancel(); setTimeout(() => doSave(acao), 0); } else save();
  };
  const restoreSnap = (snap) => {
    const d = parseXml(snap.xml);
    setDoc(d);
    let sel = null;
    if (snap.sel) for (const p of d.getElementsByTagNameNS(NS.w, 'p')) if (getParaId(p) === snap.sel) { sel = p; break; }
    setSelected(sel);
    editSession.current = null;
    setDocVer((v) => v + 1);
    setExtVer((v) => v + 1);
    touch();
  };
  const doUndo = () => {
    const u = undo.current;
    const s = u.undo.pop();
    if (!s) return;
    u.redo.push({ xml: serializeXml(latest.current.doc), sel: selected && isW(selected, 'p') ? getParaId(selected) : null, label: s.label });
    restoreSnap(s);
    setUndoTick((n) => n + 1);
    toast('Desfeito: ' + (s.label || 'alteração'));
  };
  const doRedo = () => {
    const u = undo.current;
    const s = u.redo.pop();
    if (!s) return;
    u.undo.push({ xml: serializeXml(latest.current.doc), sel: selected && isW(selected, 'p') ? getParaId(selected) : null, label: s.label });
    restoreSnap(s);
    setUndoTick((n) => n + 1);
  };

  // ------------------------------------------------------------- valores
  const setValue = (fid, v) => { setValores((o) => ({ ...o, [fid]: v })); touch(); };
  const setItems = (gid, arr) => { setItens((o) => ({ ...o, [gid]: arr })); touch(); };
  const bulk = (patch, idx, gid) => {
    if (idx !== undefined) setItens((o) => ({ ...o, [gid]: (o[gid] || []).map((it, k) => (k === idx ? { ...it, ...patch } : it)) }));
    else setValores((o) => ({ ...o, ...patch }));
    touch();
  };

  // ------------------------------------------------------------- seleção
  const select = (node, opts = {}) => {
    if (node !== selected) editSession.current = null;
    setSelected(node);
    setFocusText(opts.focus || false);
    setTextSel(null);
  };

  const onCanvasSelect = (node, info) => {
    if (!node) { select(null); return; }
    if (info && info.field && !isTpl) return; // clique em campo: foca o formulário
    select(node, { focus: info && info.dbl ? 'end' : false });
  };

  const onFieldClick = (fid, idx) => {
    const base = fid.replace(/@\d+$/, '');
    if (isTpl) {
      const c = campos.find((x) => x.id === base);
      if (c) editField(c);
      return;
    }
    setLeftOpen(true);
    setLeftTab('preencher');
    setFormFocus({ id: base, idx, t: Date.now() });
    setFocusField({ id: base, idx });
  };

  // ------------------------------------------------------------- edição de texto
  const beginSession = (label) => {
    if (editSession.current !== selected) { pushUndo(label); editSession.current = selected; }
  };
  const onText = (segs) => {
    if (!selected) return;
    beginSession('Editar texto');
    writeParagraph(selected, segs);
    changed();
  };
  const onSplit = (a, b) => {
    pushUndo('Novo parágrafo');
    writeParagraph(selected, a);
    const np = selected.cloneNode(false);
    const pPr = wKid(selected, 'pPr');
    if (pPr) { const c = pPr.cloneNode(true); const s = wKid(c, 'sectPr'); if (s) c.removeChild(s); np.appendChild(c); }
    setParaId(np);
    writeParagraph(np, b.length ? b : []);
    insertAfter(selected, np);
    editSession.current = np;
    setSelected(np);
    setFocusText('start');
    changed({ external: true });
  };
  const onMergeUp = (segs) => {
    let prev = selected.previousElementSibling;
    while (prev && !isW(prev, 'p')) { if (isW(prev, 'tbl')) return; prev = prev.previousElementSibling; }
    if (!prev) return;
    pushUndo('Juntar parágrafos');
    writeParagraph(prev, [...parseParagraph(prev), ...segs]);
    selected.parentNode.removeChild(selected);
    editSession.current = prev;
    setSelected(prev);
    setFocusText('end');
    changed({ external: true });
  };
  const onPasteLines = (lines) => {
    pushUndo('Colar parágrafos');
    let ref = selected;
    const segs0 = parseParagraph(selected);
    const rPr = dominantRPr(segs0, wKid(selected, 'pPr'));
    for (const line of lines) {
      const np = selected.cloneNode(false);
      const pPr = wKid(selected, 'pPr');
      if (pPr) np.appendChild(pPr.cloneNode(true));
      setParaId(np);
      writeParagraph(np, [{ k: 't', text: line, rPr }]);
      insertAfter(ref, np);
      ref = np;
    }
    changed({ external: true });
    toast.ok(`${lines.length} parágrafo(s) colado(s) com a mesma formatação.`);
  };

  // ------------------------------------------------------------- ações de bloco
  const tableOf = (n) => { let x = n; while (x && !isW(x, 'tbl')) x = x.parentNode; return x; };
  const blockAction = async (action, arg) => {
    if (!selected) return;
    const node = arg === 'table' ? tableOf(selected) : selected;
    switch (action) {
      case 'insert': {
        const proto = prototypeFor(ctx, doc, arg, node) || node;
        pushUndo('Inserir parágrafo');
        const np = newParagraphLike(proto);
        const k = kindOf(ctx, np);
        const r0 = rank(ctx, node);
        let after = node;
        if (isTopLevel(node) && k.level <= r0 && r0 !== 99) { const g = sectionRange(ctx, node); after = g[g.length - 1]; }
        insertAfter(after, np);
        editSession.current = np;
        setSelected(np);
        setFocusText('end');
        changed({ external: true });
        toast.ok(`${k.label} inserido(a). Digite o texto ao lado.`);
        break;
      }
      case 'duplicate': {
        pushUndo('Duplicar');
        const c = duplicate(node);
        setSelected(isW(c, 'p') ? c : selected);
        changed({ external: true });
        break;
      }
      case 'move': {
        pushUndo('Mover');
        const target = isTopLevel(node) ? node : tableOf(node);
        if (!target || !moveBlock(ctx, target, arg)) { undo.current.undo.pop(); toast('Não há para onde mover.'); return; }
        changed({ external: true });
        requestAnimationFrame(() => {
          const el = document.querySelector('.canvas-pages .is-selected');
          if (el) el.scrollIntoView({ block: 'center', behavior: 'smooth' });
        });
        break;
      }
      case 'delete': {
        const group = arg === 'table' ? [node] : isTopLevel(node) ? sectionRange(ctx, node) : [node];
        if (group.length > 1) {
          const ok = await confirmDialog({ title: 'Excluir este trecho?', message: `Serão removidos ${group.length} parágrafos (o item selecionado e seus subitens). Você pode desfazer com Ctrl+Z.`, ok: 'Excluir', danger: true });
          if (!ok) return;
        }
        pushUndo('Excluir');
        const after = group[group.length - 1].nextElementSibling;
        const before = group[0].previousElementSibling;
        group.forEach(removeBlock);
        const next = after && (isW(after, 'p') || isW(after, 'tbl')) ? after : before && (isW(before, 'p') || isW(before, 'tbl')) ? before : null;
        setSelected(null);
        void next;
        changed({ external: true });
        toast('Trecho excluído.', { action: { label: 'Desfazer', onClick: doUndo } });
        break;
      }
      case 'align': pushUndo('Alinhamento'); setAlign(node, arg); changed({ external: true }); break;
      case 'pagebreak': pushUndo('Quebra de página'); setPageBreakBefore(node, arg); changed({ external: true }); break;
      case 'convert': {
        const proto = prototypeFor(ctx, doc, arg, node);
        if (!proto) return;
        pushUndo('Mudar tipo');
        convertKind(node, proto);
        changed({ external: true });
        break;
      }
      case 'library-save': {
        const group = isTopLevel(node) ? sectionRange(ctx, node) : [node];
        const text = group.map((p) => (isW(p, 'p') ? segsText(parseParagraph(p)) : '[tabela]')).join('\n');
        const num = numberLabel(ctx, doc, node);
        const sug = text.replace(/\{\{\w+(@\d+)?\}\}/g, '…').replace(/\s+/g, ' ').trim().slice(0, 60);
        const titulo = await promptDialog({ title: 'Salvar na biblioteca de cláusulas', label: 'Nome para encontrar depois', value: sug, icon: 'bookmark-plus', ok: 'Salvar' });
        if (!titulo) return;
        await saveClausula({ titulo, texto: text, xml: group.map((g) => new XMLSerializer().serializeToString(g)), origem: rec.modeloNome || rec.nome, numero: num, categoria: kindOf(ctx, node).label });
        toast.ok(`Salvo na biblioteca${group.length > 1 ? ` (${group.length} parágrafos)` : ''}.`);
        break;
      }
      case 'library-insert': {
        const c = await libraryPickDialog();
        if (!c) return;
        pushUndo('Inserir da biblioteca');
        let ref = isTopLevel(node) ? sectionRange(ctx, node).slice(-1)[0] : node;
        let first = null;
        for (const x of c.xml) {
          const frag = parseXml(`<w:root xmlns:w="${NS.w}" xmlns:w14="${NS.w14}" xmlns:r="${NS.r}" xmlns:mc="${NS.mc}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}">${x}</w:root>`);
          const el = doc.importNode(frag.documentElement.firstElementChild, true);
          fixNumbering(el, node);
          for (const p of isW(el, 'p') ? [el] : wAll(el, 'p')) setParaId(p);
          insertAfter(ref, el);
          ref = el;
          first = first || el;
        }
        setSelected(first && isW(first, 'p') ? first : null);
        changed({ external: true });
        toast.ok('Cláusula inserida. A numeração foi ajustada automaticamente.');
        break;
      }
      case 'restore': {
        const pid = getParaId(node);
        const src = tplDoc && Array.from(tplDoc.getElementsByTagNameNS(NS.w, 'p')).find((p) => getParaId(p) === pid);
        if (!src) { toast.warn('O texto original não está mais disponível no modelo.'); return; }
        pushUndo('Restaurar texto');
        const el = doc.importNode(src, true);
        node.parentNode.replaceChild(el, node);
        editSession.current = null;
        setSelected(el);
        changed({ external: true });
        toast.ok('Texto do modelo restaurado.');
        break;
      }
      default: break;
    }
  };

  /** Ajusta numId de parágrafos vindos de outro documento. */
  const fixNumbering = (el, near) => {
    const ps = isW(el, 'p') ? [el] : wAll(el, 'p');
    for (const p of ps) {
      const pPr = wKid(p, 'pPr');
      const np = pPr && wKid(pPr, 'numPr');
      const idEl = np && wKid(np, 'numId');
      const nid = idEl && idEl.getAttributeNS(NS.w, 'val');
      if (nid && !ctx.numbering.hasNum(nid)) {
        // usa a lista do parágrafo numerado mais próximo
        let ref = near;
        let found = null;
        while (ref && !found) {
          const pp = isW(ref, 'p') && wKid(ref, 'pPr');
          const rn = pp && wKid(pp, 'numPr');
          const ri = rn && wKid(rn, 'numId');
          if (ri && ctx.numbering.hasNum(ri.getAttributeNS(NS.w, 'val'))) found = ri.getAttributeNS(NS.w, 'val');
          ref = ref.previousElementSibling;
        }
        if (found) idEl.setAttributeNS(NS.w, 'w:val', found);
        else pPr.removeChild(np);
      }
    }
  };

  // ------------------------------------------------------------- campos (modelo)
  const renameToken = (oldId, newId) => {
    for (const t of wAll(doc.documentElement, 't')) {
      if (t.textContent.includes('{{' + oldId + '}}')) t.textContent = t.textContent.split('{{' + oldId + '}}').join('{{' + newId + '}}');
    }
  };
  const editField = async (c, isNew = false) => {
    const res = await fieldDialog({ campo: c, campos, grupos, isNew, testValues: valores });
    if (!res) return;
    if (isNew) { setCampos([...campos, res]); touch(); toast.ok(`Campo “${res.rotulo}” criado. Use “Inserir campo” no texto para posicioná-lo.`); return; }
    let next = campos.map((x) => (x === c ? res : x));
    if (res.id !== c.id) {
      pushUndo('Renomear campo');
      renameToken(c.id, res.id);
      next = next.map((x) => {
        const y = { ...x };
        if (y.origem === c.id) y.origem = res.id;
        if (y.meses === c.id) y.meses = res.id;
        if (y.formula) y.formula = y.formula.replace(new RegExp('\\b' + c.id + '\\b', 'g'), res.id);
        return y;
      });
      changed({ external: true });
    }
    setCampos(next);
    touch();
  };
  const deleteField = async (c) => {
    const n = (xmlNow.match(new RegExp('\\{\\{' + c.id + '\\}\\}', 'g')) || []).length;
    const deps = campos.filter((x) => x.origem === c.id || (x.formula && new RegExp('\\b' + c.id + '\\b').test(x.formula)));
    const ok = await confirmDialog({
      title: `Excluir o campo “${c.rotulo}”?`, danger: true, ok: 'Excluir campo',
      message: html`<p class="muted">${n ? html`Ele aparece <b>${n}×</b> no texto; essas ocorrências serão trocadas por ${c.padrao ? html`“${c.padrao}”` : 'nada (removidas)'}.` : 'Ele não aparece no texto.'}
        ${deps.length ? html`<br />Atenção: ${deps.map((d) => d.rotulo).join(', ')} dependem dele.` : ''}</p>`,
    });
    if (!ok) return;
    pushUndo('Excluir campo');
    for (const t of wAll(doc.documentElement, 't')) {
      if (t.textContent.includes('{{' + c.id + '}}')) t.textContent = t.textContent.split('{{' + c.id + '}}').join(c.padrao || '');
    }
    setCampos(campos.filter((x) => x !== c));
    changed({ external: true });
  };
  const editGroup = async (g, isNew) => {
    const res = await groupDialog({ grupo: g, isNew });
    if (!res) return;
    if (res.__delete) {
      const n = campos.filter((c) => c.grupo === g.id).length;
      if (n) { toast.warn(`Mova ou exclua os ${n} campo(s) deste grupo antes.`); return; }
      setGrupos(grupos.filter((x) => x.id !== g.id));
    } else if (isNew) {
      let gid = res.id;
      while (grupos.some((x) => x.id === gid)) gid += '_2';
      setGrupos([...grupos, { ...res, id: gid }]);
    } else setGrupos(grupos.map((x) => (x.id === g.id ? { ...res, id: g.id } : x)));
    touch();
  };
  const createFromSelection = async () => {
    const info = textSel;
    setTextSel(null);
    if (!info || info.error) return;
    const res = await createFromSelectionDialog({ info, doc, campos, grupos: grupos.length ? grupos : [{ id: 'geral', nome: 'Dados do contrato' }] });
    if (!res) return;
    pushUndo('Criar campo');
    if (!grupos.length) setGrupos([{ id: 'geral', nome: 'Dados do contrato', icone: 'file-text' }]);
    applyFieldRanges(res.ranges.map((r) => ({ ...r, id: res.campo.id })));
    if (res.isNew) setCampos([...campos, res.campo]);
    window.getSelection().removeAllRanges();
    changed({ external: true });
    toast.ok(`${res.ranges.length > 1 ? res.ranges.length + ' trechos viraram' : 'Trecho virou'} o campo “${res.campo.rotulo}”.`, { action: { label: 'Desfazer', onClick: doUndo } });
  };
  const detect = async () => {
    const res = await detectDialog({ doc, campos, grupos: grupos.length ? grupos : [{ id: 'geral', nome: 'Dados do contrato' }] });
    if (!res) return;
    pushUndo('Detectar campos');
    if (!grupos.length) setGrupos([{ id: 'geral', nome: 'Dados do contrato', icone: 'file-text' }]);
    const novos = applySuggestions(res, campos);
    setCampos([...campos, ...novos]);
    changed({ external: true });
    toast.ok(`${novos.length} campo(s) criado(s). Revise nomes e tipos na aba Campos.`);
  };

  // ------------------------------------------------------------- cadastros
  const onRecords = async (op, g) => {
    const tipo = g.cadastro;
    const gCampos = campos.filter((c) => c.grupo === g.id && !isCalculado(c));
    if (op === 'save') {
      const vals = {};
      for (const c of gCampos) if (valores[c.id]) vals[c.id] = valores[c.id];
      if (!Object.keys(vals).length) { toast.warn('Preencha os dados antes de salvar no cadastro.'); return; }
      const nome = gCampos.find((c) => c.papel === 'nome') || gCampos[0];
      const doc_ = gCampos.find((c) => c.papel === 'documento');
      await saveCadastro({ tipo, titulo: valores[nome.id] || 'Sem nome', subtitulo: doc_ ? valores[doc_.id] || '' : '', valores: vals, rotulos: Object.fromEntries(gCampos.map((c) => [c.id, c.rotulo])) });
      toast.ok(`Salvo em ${tipo}.`);
      return;
    }
    const r = await pickRecordDialog({ tipo, titulo: g.lista ? `Adicionar ${(g.itemNome || 'item').toLowerCase()}` : 'Preencher com cadastro' });
    if (!r) return;
    const vals = {};
    for (const c of gCampos) if (r.valores[c.id] !== undefined) vals[c.id] = r.valores[c.id];
    if (g.lista) setItems(g.id, [...(itens[g.id] || []), vals]);
    else { setValores((o) => ({ ...o, ...vals })); touch(); }
    toast.ok(`Dados de “${r.titulo}” aplicados.`);
  };

  // ------------------------------------------------------------- exportar
  const checkMissing = async (formato) => {
    if (isTpl) return true;
    const miss = resolver.missing(used);
    if (!miss.length) return true;
    const r = await missingDialog({ missing: miss, formato });
    if (!r) return false;
    if (r.goto) { onFieldClick(r.goto.campo.id, r.goto.idx); return false; }
    return !!r.go;
  };
  const fileTitle = (ext) => (isTpl ? `Modelo - ${rec.nome}.${ext}` : nomeArquivo({ ...rec, resumo }, ext));
  const exportWord = async (raw = false) => {
    if (!raw && !(await checkMissing('Word'))) return;
    setBusy('docx');
    try {
      await doSave();
      const blob = await buildDocx({ pkgBuf: rec.pkg, docXml: serializeXml(doc), resolver: raw ? null : resolver, titulo: isTpl ? rec.nome : tituloContrato(rec), autor: appConfig.get().usuario });
      downloadBlob(blob, raw ? `Modelo - ${rec.nome}.docx` : fileTitle('docx'));
      if (!isTpl && rec.status === 'rascunho') { rec.status = 'finalizado'; setRec({ ...rec }); }
      if (!isTpl) doSave('Word gerado');
      toast.ok(raw ? 'Modelo baixado com os campos {{…}} — pode editar no Word e importar de volta.' : 'Documento Word gerado com a formatação original.', { title: 'Download pronto' });
    } catch (e) {
      console.error(e);
      toast.err('Falha ao gerar o Word: ' + e.message);
    } finally { setBusy(null); }
  };
  const exportPdf = async () => {
    if (!(await checkMissing('PDF'))) return;
    const cfg = await getConfig();
    if (!cfg.pdfDicaVista) {
      const r = await pdfTipDialog();
      if (!r) return;
      if (r.never) saveConfig({ pdfDicaVista: true });
    }
    setBusy('pdf');
    try {
      await doSave();
      await printDocument({ ctx, docXml: serializeXml(doc), resolver, titulo: fileTitle('pdf').replace(/\.pdf$/, '') });
      if (!isTpl) { if (rec.status === 'rascunho') { rec.status = 'finalizado'; setRec({ ...rec }); } doSave('PDF gerado'); }
    } catch (e) {
      console.error(e);
      toast.err('Falha ao preparar o PDF: ' + e.message);
    } finally { setBusy(null); }
  };

  // ------------------------------------------------------------- teclado
  useEffect(() => {
    const onKey = (e) => {
      const tag = (e.target.tagName || '').toLowerCase();
      const typing = tag === 'input' || tag === 'textarea' || tag === 'select' || e.target.isContentEditable;
      const mod = e.ctrlKey || e.metaKey;
      if (mod && e.key.toLowerCase() === 's') { e.preventDefault(); save.flush(); toast.ok('Salvo.'); return; }
      if (mod && e.key.toLowerCase() === 'p') { e.preventDefault(); exportPdf(); return; }
      if (typing) return;
      if (mod && e.key.toLowerCase() === 'z' && !e.shiftKey) { e.preventDefault(); doUndo(); }
      else if (mod && (e.key.toLowerCase() === 'y' || (e.key.toLowerCase() === 'z' && e.shiftKey))) { e.preventDefault(); doRedo(); }
      else if (e.key === 'Escape') select(null);
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });

  // ------------------------------------------------------------- render
  if (error) return html`<div class="page-center"><div class="empty"><div class="empty-icon"><${Icon} name="alert" size=${28} /></div><h3>${error}</h3>
    <${Button} icon="arrow-left" onClick=${() => navigate(isTpl ? '/modelos' : '/contratos')}>Voltar</${Button}></div></div>`;
  if (!rec || !ctx || !doc) return html`<div class="page-center"><span class="spinner lg" /><p class="muted">Abrindo ${isTpl ? 'modelo' : 'contrato'}…</p></div>`;

  const st = statusInfo(rec.status);
  const miss = isTpl ? [] : resolver.missing(used);
  const reqTotal = isTpl ? 0 : countRequired(campos, grupos, itens, used);
  const pct = reqTotal ? Math.round(((reqTotal - miss.length) / reqTotal) * 100) : 100;
  const u = undo.current;
  const selRow = selected ? listRowName(selected, canvasState, grupos, campos) : null;

  return html`<div class=${cx('editor', isTpl && 'is-tpl', !leftOpen && 'left-closed', selected && 'has-inspector')}>
    <header class="ed-top">
      <${IconButton} icon="arrow-left" title="Voltar" onClick=${() => { save.flush(); navigate(isTpl ? '/modelos' : '/contratos'); }} />
      <div class="ed-title">
        ${isTpl ? html`<input class="title-input" value=${rec.nome} onInput=${(e) => { rec.nome = e.currentTarget.value; setRec({ ...rec }); touch(); }} aria-label="Nome do modelo" />`
          : html`<strong>${(resumo && resumo.numero) || 'Contrato'}${resumo && resumo.cliente ? ' · ' + resumo.cliente : html` <span class="muted">· cliente não informado</span>`}</strong>`}
        <small>${isTpl ? html`<${Icon} name="layers" size=${12} /> Modelo · ${campos.length} campos` : rec.modeloNome}</small>
      </div>
      ${!isTpl && html`<${Menu} trigger=${({ toggle }) => html`<button type="button" class=${'status-pill tone-' + st[2]} onClick=${toggle}><span class="dot" />${st[1]}<${Icon} name="chevron-down" size=${13} /></button>`}
        items=${STATUS.map(([k, n]) => ({ label: n, icon: k === rec.status ? 'check' : undefined, onClick: () => { rec.status = k; setRec({ ...rec }); doSave('Status: ' + n); } }))} />`}
      <span class=${'save-ind ' + saveState} title="Salvamento automático">
        ${saveState === 'saving' ? html`<span class="spinner sm" /> Salvando…` : saveState === 'dirty' ? html`<${Icon} name="clock" size=${14} /> Alterado` : saveState === 'error' ? html`<${Icon} name="alert" size=${14} /> Erro ao salvar` : html`<${Icon} name="check-circle" size=${14} /> Salvo`}
      </span>
      <span class="spacer" />
      <div class="ed-tools">
        <${IconButton} icon="undo" title=${'Desfazer' + (u.undo.length ? ': ' + u.undo[u.undo.length - 1].label : '') + ' (Ctrl+Z)'} disabled=${!u.undo.length} onClick=${doUndo} />
        <${IconButton} icon="redo" title="Refazer (Ctrl+Y)" disabled=${!u.redo.length} onClick=${doRedo} />
        <span class="tool-sep" />
        <${IconButton} icon="zoom-out" title="Diminuir zoom" onClick=${() => setZoom(Math.max(0.4, Math.round(((zoom === 'auto' ? zoomVal : zoom) - 0.1) * 10) / 10))} />
        <button type="button" class="zoom-val" title="Ajustar à largura" onClick=${() => setZoom('auto')}>${Math.round(zoomVal * 100)}%</button>
        <${IconButton} icon="zoom-in" title="Aumentar zoom" onClick=${() => setZoom(Math.min(2, Math.round(((zoom === 'auto' ? zoomVal : zoom) + 0.1) * 10) / 10))} />
        <span class="tool-sep" />
        <${IconButton} icon="highlighter" title="Destacar campos no texto" active=${highlight} onClick=${() => { setHighlight(!highlight); saveConfig({ destacarCampos: !highlight }); }} />
        ${!isTpl && html`<${IconButton} icon="compare" title=${`Mostrar alterações em relação ao modelo (${diff.marks.size} trecho(s)${diff.removed ? `, ${diff.removed} removido(s)` : ''})`} active=${showChanges} onClick=${() => setShowChanges(!showChanges)} />`}
      </div>
      <div class="ed-actions">
        ${isTpl && html`<${Button} icon="wand" variant="ghost" onClick=${detect}>Detectar campos</${Button}>`}
        <${Button} icon="download" loading=${busy === 'docx'} onClick=${() => exportWord(false)}>${isTpl ? 'Testar Word' : 'Word'}</${Button}>
        <${Button} icon="file-down" variant="primary" loading=${busy === 'pdf'} onClick=${exportPdf}>PDF</${Button}>
        <${Menu} align="right" trigger=${({ toggle }) => html`<${IconButton} icon="more-v" title="Mais opções" onClick=${toggle} />`}
          items=${isTpl ? [
            { icon: 'file-plus', label: 'Novo contrato com este modelo', onClick: async () => { await doSave(); const c = await createContrato(rec.id); navigate('/contrato/' + c.id); } },
            { icon: 'download', label: 'Baixar modelo (.docx com {{campos}})', hint: 'Edite no Word e importe de volta', onClick: () => exportWord(true) },
            { icon: 'printer', label: 'Imprimir', kbd: 'Ctrl+P', onClick: exportPdf },
          ] : [
            { icon: 'copy-plus', label: 'Duplicar contrato', hint: 'Novo número, mesmos dados', onClick: async () => { await doSave(); const c = await createContrato(rec.modeloId, { fromId: rec.id }); toast.ok('Contrato duplicado.'); navigate('/contrato/' + c.id); } },
            { icon: 'printer', label: 'Imprimir', kbd: 'Ctrl+P', onClick: exportPdf },
            { icon: 'history', label: 'Histórico', onClick: () => showHistory(rec) },
            '-',
            { icon: 'trash', label: 'Excluir contrato', danger: true, onClick: async () => { if (await confirmDialog({ title: 'Excluir este contrato?', message: 'Esta ação não pode ser desfeita.', ok: 'Excluir', danger: true })) { save.cancel(); await deleteContrato(rec.id); toast('Contrato excluído.'); navigate('/contratos'); } } },
          ]} />
      </div>
    </header>

    <div class="ed-body">
      <aside class="ed-left">
        <div class="ed-tabs">
          ${isTpl && html`<button type="button" class=${cx(leftTab === 'campos' && 'is-active')} onClick=${() => setLeftTab('campos')}><${Icon} name="braces" size=${15} /> Campos</button>`}
          <button type="button" class=${cx(leftTab === 'preencher' && 'is-active')} onClick=${() => setLeftTab('preencher')}><${Icon} name="text-input" size=${15} /> ${isTpl ? 'Testar' : 'Preencher'}</button>
          <button type="button" class=${cx(leftTab === 'estrutura' && 'is-active')} onClick=${() => setLeftTab('estrutura')}><${Icon} name="tree" size=${15} /> Cláusulas</button>
          <${IconButton} icon="panel-left" title="Recolher painel" class="ed-collapse" onClick=${() => setLeftOpen(false)} />
        </div>
        ${!isTpl && leftTab === 'preencher' && html`<div class="progress-row" title="Campos obrigatórios preenchidos">
          <div class="progress"><div class="progress-bar" style=${{ width: pct + '%' }} /></div>
          <span class=${cx('progress-txt', pct === 100 && 'done')}>${pct === 100 ? html`<${Icon} name="check-circle" size=${14} /> Tudo preenchido` : `${reqTotal - miss.length} de ${reqTotal} obrigatórios`}</span>
        </div>`}
        <div class="ed-left-body">
          ${leftTab === 'preencher' && html`<${FieldForm} campos=${campos} grupos=${grupos} valores=${valores} itens=${itens} resolver=${resolver}
            used=${used} onValue=${setValue} onItems=${setItems} onBulk=${bulk} focus=${formFocus} mode=${isTpl ? 'teste' : 'contrato'}
            onFocusField=${(fid, idx) => setFocusField({ id: fid, idx })} onRecords=${isTpl ? null : onRecords} onFieldClick=${(fid, idx) => setFocusField({ id: fid, idx, t: Date.now() })} />`}
          ${leftTab === 'estrutura' && html`<${Outline} ctx=${ctx} doc=${doc} version=${docVer} selected=${selected} marks=${diff.marks} campos=${campos}
            onPick=${(n) => { select(n); requestAnimationFrame(() => { const el = document.querySelector('.canvas-pages .is-selected'); if (el) el.scrollIntoView({ block: 'center', behavior: 'smooth' }); }); }} />`}
          ${leftTab === 'campos' && html`<${FieldsManager} campos=${campos} grupos=${grupos} xml=${xmlNow}
            onEditField=${(c) => editField(c)} onNewField=${(pre) => editField({ grupo: (grupos[0] || {}).id || 'geral', ...(pre || {}) }, true)}
            onDeleteField=${deleteField} onEditGroup=${(g) => editGroup(g, false)} onNewGroup=${() => editGroup({}, true)} onDetect=${detect}
            onFocusField=${(fid) => setFocusField({ id: fid, t: Date.now() })} />`}
        </div>
      </aside>
      ${!leftOpen && html`<button type="button" class="left-reopen" title="Abrir painel" onClick=${() => setLeftOpen(true)}><${Icon} name="panel-right" size=${18} /></button>`}

      <main class="ed-canvas">
        ${canvasState && html`<${DocCanvas} ctx=${ctx} doc=${canvasState.doc} version=${canvasState.ver} resolve=${canvasState.resolve} listUnits=${canvasState.listUnits}
          marks=${canvasState.marks} highlight=${highlight} selected=${selected} focusField=${focusField} zoom=${zoom} templateMode=${isTpl}
          showChanges=${showChanges} onSelect=${onCanvasSelect} onFieldClick=${onFieldClick} onTextSelect=${isTpl ? setTextSel : null}
          onPages=${setPages} onZoomComputed=${setZoomVal} />`}
        <div class="canvas-status">
          <span><${Icon} name="files" size=${13} /> ${pages} ${pages === 1 ? 'página' : 'páginas'}</span>
          ${!isTpl && diff.marks.size > 0 && html`<button type="button" class="link-btn" onClick=${() => setShowChanges(!showChanges)}><${Icon} name="compare" size=${13} /> ${diff.marks.size} trecho(s) diferente(s) do modelo</button>`}
          <span class="muted">${isTpl ? 'Clique num trecho para editar · selecione um texto para criar campo' : 'Clique num trecho para editar · clique num campo para preenchê-lo'}</span>
        </div>
        ${textSel && html`<div class="sel-pop" style=${{ left: Math.min(window.innerWidth - 120, Math.max(120, textSel.rect.left + textSel.rect.width / 2)) + 'px', top: Math.min(window.innerHeight - 60, Math.max(70, textSel.rect.top - 48)) + 'px' }}>
          ${textSel.error ? html`<span><${Icon} name="info" size=${14} /> ${textSel.error}</span>`
            : html`<button type="button" onMouseDown=${(e) => e.preventDefault()} onClick=${createFromSelection}><${Icon} name="braces" size=${15} /> Transformar em campo</button>`}
        </div>`}
      </main>

      ${selected && html`<${Inspector} key=${getParaId(selected) || 'tbl'} node=${selected} ctx=${ctx} doc=${doc} campos=${campos} kinds=${kinds}
        extVersion=${extVer} mode=${isTpl ? 'modelo' : 'contrato'} edited=${!isTpl && diff.marks.get(selected) === 'edited' && !!tplDoc}
        listRow=${selRow} focusText=${focusText} onText=${onText} onSplit=${onSplit} onMergeUp=${onMergeUp} onPasteLines=${onPasteLines}
        onAction=${blockAction} onClose=${() => select(null)} />`}
    </div>
  </div>`;
}

function countRequired(campos, grupos, itens, used) {
  const gl = new Map(grupos.map((g) => [g.id, g]));
  let n = 0;
  for (const c of campos) {
    if (!c.obrigatorio || isCalculado(c) || (used && !used.has(c.id))) continue;
    const g = gl.get(c.grupo);
    n += g && g.lista ? Math.max(1, (itens[g.id] || []).length) : 1;
  }
  return n;
}

function listRowName(node, canvasState, grupos, campos) {
  if (!canvasState || !canvasState.listUnits) return null;
  let n = node;
  while (n && n.nodeType === 1) {
    const u = canvasState.listUnits.get(n);
    if (u) { const g = grupos.find((x) => x.id === u.grupo); return g ? g.nome : u.grupo; }
    n = n.parentNode;
  }
  void campos;
  return null;
}

function showHistory(rec) {
  openModal(({ close }) => html`<div>
    <${ModalHead} title="Histórico do contrato" icon="history" close=${close} />
    <div class="modal-body"><ol class="timeline">${(rec.historico || []).slice().reverse().map((e) => html`<li><span class="tl-dot" /><div><strong>${e.acao}</strong><small>${new Date(e.em).toLocaleString('pt-BR')}</small></div></li>`)}</ol></div>
  </div>`, { size: 'sm' });
}

export { topBlocks, Badge };
