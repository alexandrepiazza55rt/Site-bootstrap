// Regras de negócio sobre os dados: modelos, contratos, cadastros, biblioteca e backup.
import { db, uid, persistStorage } from './db.js';
import { DocxPackage } from '../docx/pkg.js';
import { parseXml, serializeXml } from '../docx/xml.js';
import { normalizeDocument, collectTokens } from '../docx/normalize.js';
import { paraTextMap } from '../docx/ops.js';
import { FieldResolver } from '../fields/compute.js';
import { novoCampo, sugerirTipo, isCalculado } from '../fields/types.js';
import { isoHoje } from '../br/docs.js';
import { seedTemplate, SEED_ID } from '../seed/rentloc.js';

const now = () => new Date().toISOString();

// ------------------------------------------------------------ config
const DEFAULT_CONFIG = {
  key: 'app', usuario: '', empresa: 'Rent Loc Equipamentos', tema: 'auto', seq: {}, onboarded: false,
  destacarCampos: true, zoom: 'auto', autor: '',
};

export async function getConfig() {
  const c = await db.get('config', 'app');
  return { ...DEFAULT_CONFIG, ...(c || {}) };
}

export async function saveConfig(patch) {
  const c = await getConfig();
  const next = { ...c, ...patch, key: 'app' };
  await db.put('config', next);
  return next;
}

// ------------------------------------------------------------ seed
export async function ensureSeed() {
  const cfg = await db.get('config', 'app');
  if (cfg && cfg.seeded) return false;
  const existing = await db.get('modelos', SEED_ID);
  if (!existing) {
    const buf = seedTemplate.docx.buffer.slice(seedTemplate.docx.byteOffset, seedTemplate.docx.byteOffset + seedTemplate.docx.byteLength);
    const m = await modeloFromBuffer(buf, { id: SEED_ID, nome: seedTemplate.nome, origem: 'seed' });
    m.descricao = seedTemplate.descricao;
    m.cor = seedTemplate.cor;
    m.grupos = seedTemplate.grupos;
    m.campos = seedTemplate.campos;
    m.config = seedTemplate.config;
    await db.put('modelos', m);
  }
  const ano = String(new Date().getFullYear());
  await saveConfig({ seeded: true, seq: { ...(cfg && cfg.seq), [ano]: (cfg && cfg.seq && cfg.seq[ano]) || (ano === '2026' ? 63 : 1) } });
  persistStorage();
  return true;
}

// ------------------------------------------------------------ modelos
export async function modeloFromBuffer(buf, { id, nome, origem = 'import' } = {}) {
  const pkg = await DocxPackage.load(buf);
  const doc = parseXml(pkg.get(pkg.mainPath()));
  normalizeDocument(doc);
  const tokens = collectTokens(doc);
  const campos = [];
  for (const [cid, label] of tokens) {
    const sug = sugerirTipo('');
    campos.push(novoCampo({ id: cid, rotulo: prettyLabel(label), grupo: 'geral', ...sug }));
  }
  return {
    id: id || uid('m_'),
    nome: nome || 'Novo modelo',
    descricao: '',
    cor: pickColor(),
    origem,
    grupos: [{ id: 'geral', nome: 'Dados do contrato', icone: 'file-text' }],
    campos,
    config: {},
    pkg: buf,
    docXml: serializeXml(doc),
    testeValores: {},
    testeItens: {},
    criadoEm: now(),
    atualizadoEm: now(),
  };
}

function prettyLabel(s) {
  const t = String(s).replace(/_/g, ' ').trim();
  return t ? t[0].toUpperCase() + t.slice(1) : 'Campo';
}

const COLORS = ['#2563eb', '#7c3aed', '#059669', '#d97706', '#dc2626', '#0891b2', '#db2777', '#4f46e5'];
function pickColor() { return COLORS[Math.floor(Math.random() * COLORS.length)]; }

export async function importModelo(file, nome) {
  const buf = await file.arrayBuffer();
  const m = await modeloFromBuffer(buf, { nome: nome || file.name.replace(/\.docx$/i, '') });
  await db.put('modelos', m);
  return m;
}

export const listModelos = () => db.all('modelos').then((l) => l.sort((a, b) => (b.atualizadoEm || '').localeCompare(a.atualizadoEm || '')));
export const getModelo = (id) => db.get('modelos', id);
export async function saveModelo(m) {
  m.atualizadoEm = now();
  await db.put('modelos', m);
  return m;
}
export const deleteModelo = (id) => db.del('modelos', id);

export async function duplicateModelo(id) {
  const m = await getModelo(id);
  const c = { ...structuredClone(m), id: uid('m_'), nome: m.nome + ' (cópia)', origem: 'copia', criadoEm: now(), atualizadoEm: now() };
  await db.put('modelos', c);
  return c;
}

// ------------------------------------------------------------ numeração
export function formatNumero(formato, n, ano) {
  return String(formato || '{n:4}/{ano}')
    .replace(/\{n(?::(\d+))?\}/, (_, d) => String(n).padStart(d ? +d : 1, '0'))
    .replace(/\{ano\}/, String(ano));
}

export async function peekNumero(formato) {
  const cfg = await getConfig();
  const ano = new Date().getFullYear();
  const n = (cfg.seq && cfg.seq[ano]) || 1;
  return formatNumero(formato, n, ano);
}

async function takeNumero(formato) {
  const cfg = await getConfig();
  const ano = new Date().getFullYear();
  const n = (cfg.seq && cfg.seq[ano]) || 1;
  await saveConfig({ seq: { ...cfg.seq, [ano]: n + 1 } });
  return formatNumero(formato, n, ano);
}

// ------------------------------------------------------------ contratos
export async function createContrato(modeloId, { fromId } = {}) {
  const m = await getModelo(modeloId);
  if (!m) throw new Error('Modelo não encontrado.');
  const base = fromId ? await getContrato(fromId) : null;
  const valores = base ? { ...base.valores } : {};
  const itens = base ? structuredClone(base.itens || {}) : {};
  if (!base) {
    // padrões já preenchidos (visíveis e editáveis no formulário)
    const listas = new Set((m.grupos || []).filter((g) => g.lista).map((g) => g.id));
    for (const c of m.campos) {
      if (listas.has(c.grupo) || !c.padrao || isCalculado(c) || c.tipo === 'sequencial') continue;
      valores[c.id] = c.padrao === 'hoje' && c.tipo === 'data' ? isoHoje() : c.padrao;
    }
  }
  const seqCampo = (base ? base.campos : m.campos).find((c) => c.tipo === 'sequencial');
  if (seqCampo) valores[seqCampo.id] = await takeNumero(seqCampo.formato);
  const docXml = base ? base.docXml : m.docXml;
  const c = {
    id: uid('c_'),
    modeloId: m.id,
    modeloNome: m.nome,
    cor: m.cor,
    status: 'rascunho',
    valores,
    itens,
    docXml,
    pkg: base ? base.pkg : m.pkg,
    campos: structuredClone(base ? base.campos : m.campos),
    grupos: structuredClone(base ? base.grupos : m.grupos),
    config: structuredClone((base ? base.config : m.config) || {}),
    baseTextos: Object.fromEntries(paraTextMap(parseXml(m.docXml))),
    historico: [{ em: now(), acao: base ? `Criado a partir do contrato ${base.resumo?.numero || ''}`.trim() : 'Criado' }],
    criadoEm: now(),
    atualizadoEm: now(),
  };
  c.resumo = resumoContrato(c);
  await db.put('contratos', c);
  return c;
}

export const getContrato = (id) => db.get('contratos', id);
export const listContratos = () => db.all('contratos').then((l) => l.sort((a, b) => (b.atualizadoEm || '').localeCompare(a.atualizadoEm || '')));

export async function saveContrato(c, acao) {
  c.atualizadoEm = now();
  c.resumo = resumoContrato(c);
  if (acao) {
    c.historico = c.historico || [];
    c.historico.push({ em: now(), acao });
    if (c.historico.length > 80) c.historico = c.historico.slice(-80);
  }
  await db.put('contratos', c);
  return c;
}

export const deleteContrato = (id) => db.del('contratos', id);

export function resumoContrato(c) {
  const r = new FieldResolver(c, c.valores, c.itens);
  const cfg = c.config || {};
  const pick = (id) => (id ? r.get(id) : null);
  const num = pick(cfg.campoNumero || (c.campos.find((x) => x.tipo === 'sequencial') || {}).id);
  const cli = pick(cfg.campoCliente || (c.campos.find((x) => x.papel === 'nome') || {}).id);
  const val = pick(cfg.campoValor || (c.campos.find((x) => x.papel === 'valor') || {}).id);
  const ini = pick(cfg.campoInicio || (c.campos.find((x) => x.papel === 'inicio') || {}).id);
  const ter = pick(cfg.campoTermino || (c.campos.find((x) => x.papel === 'termino') || {}).id);
  return {
    numero: num && !num.empty ? num.text : '',
    cliente: cli && !cli.empty ? cli.text : '',
    valor: val && Number.isFinite(val.num) ? val.num : null,
    valorTexto: val && !val.empty ? val.text : '',
    inicio: ini && ini.iso ? ini.iso : null,
    termino: ter && ter.iso ? ter.iso : null,
  };
}

export function tituloContrato(c) {
  const r = c.resumo || {};
  const parts = [r.numero, r.cliente].filter(Boolean);
  return parts.length ? parts.join(' · ') : c.modeloNome || 'Contrato sem título';
}

export function nomeArquivo(c, ext) {
  const r = c.resumo || {};
  const base = ['Contrato', r.numero && r.numero.replace(/\//g, '-'), r.cliente].filter(Boolean).join(' - ') || 'Contrato';
  return base.replace(/[\\/:*?"<>|]+/g, '').replace(/\s+/g, ' ').trim().slice(0, 120) + '.' + ext;
}

// ------------------------------------------------------------ cadastros
export const listCadastros = () => db.all('cadastros').then((l) => l.sort((a, b) => (a.titulo || '').localeCompare(b.titulo || '', 'pt-BR')));
export const deleteCadastro = (id) => db.del('cadastros', id);

export async function saveCadastro(rec) {
  rec.id = rec.id || uid('r_');
  rec.atualizadoEm = now();
  await db.put('cadastros', rec);
  return rec;
}

// ------------------------------------------------------------ biblioteca de cláusulas
export const listClausulas = () => db.all('clausulas').then((l) => l.sort((a, b) => (b.criadoEm || '').localeCompare(a.criadoEm || '')));
export const deleteClausula = (id) => db.del('clausulas', id);
export async function saveClausula(c) {
  c.id = c.id || uid('k_');
  c.criadoEm = c.criadoEm || now();
  await db.put('clausulas', c);
  return c;
}

// ------------------------------------------------------------ backup
function toB64(buf) {
  const bytes = new Uint8Array(buf);
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

function fromB64(s) {
  const bin = atob(s);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out.buffer;
}

function encode(v) {
  if (v instanceof ArrayBuffer) return { __b64: toB64(v) };
  if (ArrayBuffer.isView(v)) return { __b64: toB64(v.buffer.slice(v.byteOffset, v.byteOffset + v.byteLength)) };
  if (Array.isArray(v)) return v.map(encode);
  if (v && typeof v === 'object') { const o = {}; for (const k in v) o[k] = encode(v[k]); return o; }
  return v;
}

function decode(v) {
  if (v && typeof v === 'object' && !Array.isArray(v) && typeof v.__b64 === 'string' && Object.keys(v).length === 1) return fromB64(v.__b64);
  if (Array.isArray(v)) return v.map(decode);
  if (v && typeof v === 'object') { const o = {}; for (const k in v) o[k] = decode(v[k]); return o; }
  return v;
}

export async function exportBackup() {
  const data = { app: 'ContratoFácil', versao: 1, exportadoEm: now(), dados: {} };
  for (const s of ['modelos', 'contratos', 'cadastros', 'clausulas', 'config']) data.dados[s] = encode(await db.all(s));
  return new Blob([JSON.stringify(data)], { type: 'application/json' });
}

export async function importBackup(file, { substituir = false } = {}) {
  const txt = await file.text();
  let data;
  try { data = JSON.parse(txt); } catch { throw new Error('Arquivo de backup inválido.'); }
  if (!data || data.app !== 'ContratoFácil' || !data.dados) throw new Error('Este arquivo não é um backup do ContratoFácil.');
  const counts = {};
  for (const s of ['modelos', 'contratos', 'cadastros', 'clausulas', 'config']) {
    if (substituir) await db.clear(s);
    const items = decode(data.dados[s] || []);
    counts[s] = items.length;
    for (const it of items) await db.put(s, it);
  }
  return counts;
}

export async function wipeAll() {
  for (const s of ['modelos', 'contratos', 'cadastros', 'clausulas', 'config']) await db.clear(s);
}
