// Motor de campos: resolve valores digitados, padrões, extensos, datas calculadas e fórmulas.
import { moedaExtenso, inteiroExtenso, numeroExtenso, dataExtenso } from '../br/extenso.js';
import {
  formataMoeda, parseMoeda, formataNumero, parseNumero, formataData, somaData, isoHoje, aplicaCaixa, dataBRparaISO,
} from '../br/docs.js';

export const TOKEN_RE = /\{\{([A-Za-z0-9_]+)(?:@(\d+))?\}\}/g;

/** ids de campos usados num XML de documento. */
export function tokensIn(xml) {
  const out = new Set();
  TOKEN_RE.lastIndex = 0;
  let m;
  while ((m = TOKEN_RE.exec(xml))) out.add(m[1]);
  return out;
}

// --------------------------------------------------------- fórmulas
const astCache = new Map();

export function parseFormula(src) {
  if (astCache.has(src)) return astCache.get(src);
  const toks = String(src || '').match(/\d+(?:\.\d+)?|[A-Za-z_][A-Za-z0-9_]*|[-+*/(),%]|\S/g) || [];
  let i = 0;
  const peek = () => toks[i];
  const next = () => toks[i++];
  const expr = () => {
    let n = term();
    while (peek() === '+' || peek() === '-') { const op = next(); n = { op, a: n, b: term() }; }
    return n;
  };
  const term = () => {
    let n = unary();
    while (peek() === '*' || peek() === '/') { const op = next(); n = { op, a: n, b: unary() }; }
    return n;
  };
  const unary = () => {
    if (peek() === '-') { next(); return { op: 'neg', a: unary() }; }
    if (peek() === '+') { next(); return unary(); }
    let n = primary();
    while (peek() === '%') { next(); n = { op: '/', a: n, b: { num: 100 } }; }
    return n;
  };
  const primary = () => {
    const t = next();
    if (t === undefined) throw new Error('Fórmula incompleta');
    if (t === '(') {
      const n = expr();
      if (next() !== ')') throw new Error('Falta fechar parênteses');
      return n;
    }
    if (/^\d/.test(t)) return { num: parseFloat(t) };
    if (/^[A-Za-z_]/.test(t)) {
      if (peek() === '(') {
        next();
        const args = [];
        if (peek() !== ')') {
          args.push(expr());
          while (peek() === ',') { next(); args.push(expr()); }
        }
        if (next() !== ')') throw new Error('Falta fechar parênteses em ' + t);
        return { fn: t.toLowerCase(), args };
      }
      return { id: t };
    }
    throw new Error('Símbolo inesperado: ' + t);
  };
  let ast;
  if (!toks.length) ast = { num: 0 };
  else {
    ast = expr();
    if (i < toks.length) throw new Error('Símbolo inesperado: ' + toks[i]);
  }
  astCache.set(src, ast);
  return ast;
}

export function formulaIds(src) {
  const out = new Set();
  try {
    const walk = (n) => {
      if (!n) return;
      if (n.id) out.add(n.id);
      if (n.a) walk(n.a);
      if (n.b) walk(n.b);
      if (n.args) n.args.forEach(walk);
    };
    walk(parseFormula(src));
  } catch { /* fórmula inválida */ }
  return out;
}

// ---------------------------------------------------------- resolver
export class FieldResolver {
  /**
   * @param {{campos:Array, grupos:Array}} model
   * @param {Object} valores  id -> string
   * @param {Object} itens    grupoId -> [{id -> string}]
   */
  constructor(model, valores = {}, itens = {}) {
    this.campos = model.campos || [];
    this.grupos = model.grupos || [];
    this.valores = valores || {};
    this.itens = itens || {};
    this.byId = new Map(this.campos.map((c) => [c.id, c]));
    this.grupoById = new Map(this.grupos.map((g) => [g.id, g]));
    this.cache = new Map();
    this.stack = new Set();
  }

  isList(campo) {
    const g = campo && this.grupoById.get(campo.grupo);
    return !!(g && g.lista);
  }

  itemsOf(grupoId) {
    return this.itens[grupoId] || [];
  }

  /** Resolve "id" ou "id@3". */
  token(tok) {
    const m = /^([A-Za-z0-9_]+)(?:@(\d+))?$/.exec(tok);
    if (!m) return { text: tok, empty: true, label: tok };
    return this.get(m[1], m[2] !== undefined ? +m[2] : undefined);
  }

  get(id, idx) {
    const key = idx === undefined ? id : id + '@' + idx;
    if (this.cache.has(key)) return this.cache.get(key);
    if (this.stack.has(key)) return { text: '', empty: true, error: 'Referência circular', label: id };
    this.stack.add(key);
    let r;
    try { r = this.compute(id, idx); } catch (e) { r = { text: '', empty: true, error: e.message }; }
    this.stack.delete(key);
    const campo = this.byId.get(id);
    r.campo = campo;
    r.label = campo ? campo.rotulo || id : id;
    if (!r.empty && campo && campo.caixa && typeof r.text === 'string') r.text = aplicaCaixa(r.text, campo.caixa);
    if (r.empty) r.text = campo && campo.vazio ? campo.vazio : '';
    this.cache.set(key, r);
    return r;
  }

  num(id, idx) {
    const r = this.get(id, idx);
    return Number.isFinite(r.num) ? r.num : 0;
  }

  rawInput(campo, idx) {
    let raw;
    if (this.isList(campo)) {
      const it = this.itemsOf(campo.grupo)[idx === undefined ? -1 : idx];
      raw = it ? it[campo.id] : undefined;
    } else raw = this.valores[campo.id];
    if (raw === undefined || raw === null || String(raw).trim() === '') {
      raw = campo.padrao;
      if (raw === 'hoje' && campo.tipo === 'data') raw = isoHoje();
    }
    return raw === undefined || raw === null ? '' : String(raw);
  }

  compute(id, idx) {
    const c = this.byId.get(id);
    if (!c) return { text: '{{' + id + '}}', empty: true, unknown: true };
    const list = this.isList(c);
    if (list && idx === undefined && c.tipo !== 'calculo') {
      // campo de lista usado fora da linha repetida: junta os valores
      const items = this.itemsOf(c.grupo);
      const parts = items.map((_, i) => this.get(id, i)).filter((r) => !r.empty).map((r) => r.text);
      return { text: parts.join('; '), empty: !parts.length, num: items.reduce((a, _, i) => a + this.num(id, i), 0) };
    }
    switch (c.tipo) {
      case 'indice': return { text: String((idx || 0) + 1), num: (idx || 0) + 1, computed: true };
      case 'moeda': {
        const raw = this.rawInput(c, idx);
        if (!raw) return { empty: true };
        const n = parseMoeda(raw);
        if (!Number.isFinite(n)) return { text: raw, num: NaN };
        return { text: (c.simbolo ? 'R$ ' : '') + formataMoeda(n), num: n };
      }
      case 'numero': {
        const raw = this.rawInput(c, idx);
        if (!raw) return { empty: true };
        const n = parseNumero(raw);
        if (!Number.isFinite(n)) return { text: raw, num: NaN };
        return { text: formataNumero(n, { digitos: +c.digitos || 0, decimais: c.decimais === '' || c.decimais === undefined ? null : c.decimais, milhar: c.milhar !== false }) + (c.sufixo || ''), num: n };
      }
      case 'data': {
        const raw = this.rawInput(c, idx);
        if (!raw) return { empty: true };
        const iso = /^\d{4}-\d{2}-\d{2}/.test(raw) ? raw.slice(0, 10) : dataBRparaISO(raw);
        if (!iso) return { text: raw };
        return { text: c.formato === 'extenso' ? dataExtenso(iso, { mesMaiusculo: !!c.mesMaiusculo }) : formataData(iso), iso, num: Date.parse(iso) / 86400000 };
      }
      case 'extenso': {
        const src = this.byId.get(c.origem);
        if (!src) return { empty: true, error: 'Campo de origem não definido' };
        const sidx = this.isList(src) ? idx : undefined;
        const s = this.get(c.origem, sidx);
        if (s.empty) return { empty: true, computed: true };
        let t;
        const st = src.tipo === 'calculo' ? (src.formato || 'numero') : src.tipo;
        if (s.iso) t = dataExtenso(s.iso, { mesMaiusculo: !!c.mesMaiusculo, diaExtenso: true });
        else if (st === 'moeda') t = moedaExtenso(s.num);
        else if (Number.isFinite(s.num)) t = Number.isInteger(s.num) ? inteiroExtenso(s.num, !!c.feminino) : numeroExtenso(s.num, !!c.feminino);
        else t = '';
        return { text: t, computed: true, empty: !t };
      }
      case 'dataExtenso': {
        const s = c.origem ? this.get(c.origem, idx) : { iso: isoHoje() };
        if (!s.iso) return { empty: true, computed: true };
        return { text: dataExtenso(s.iso, { mesMaiusculo: !!c.mesMaiusculo }), iso: s.iso, computed: true };
      }
      case 'dataCalculada': {
        const base = this.get(c.origem, idx);
        if (!base.iso) return { empty: true, computed: true };
        const val = (v) => (v === undefined || v === '' ? 0 : Number.isFinite(+v) ? +v : this.num(v, idx));
        const iso = somaData(base.iso, val(c.meses), val(c.dias));
        return { text: c.formato === 'extenso' ? dataExtenso(iso, { mesMaiusculo: !!c.mesMaiusculo }) : formataData(iso), iso, computed: true };
      }
      case 'calculo': {
        const override = list ? this.rawInput({ ...c, padrao: '' }, idx) : this.valores[c.id];
        let n;
        let manual = false;
        if (c.editavel !== false && override !== undefined && override !== null && String(override).trim() !== '') {
          n = c.formato === 'moeda' ? parseMoeda(override) : parseNumero(override);
          manual = true;
        } else {
          const prev = this.sawEmpty;
          this.sawEmpty = false;
          n = this.evaluate(parseFormula(c.formula || '0'), idx);
          const saw = this.sawEmpty;
          this.sawEmpty = prev;
          if (saw) return { empty: true, computed: true };
        }
        if (!Number.isFinite(n)) return { empty: true, computed: true, error: 'Resultado inválido' };
        const text = c.formato === 'moeda' ? (c.simbolo ? 'R$ ' : '') + formataMoeda(n)
          : formataNumero(n, { digitos: +c.digitos || 0, decimais: c.decimais === '' || c.decimais === undefined ? (c.formato === 'inteiro' ? 0 : null) : c.decimais });
        return { text, num: n, computed: !manual, manual };
      }
      default: {
        const raw = this.rawInput(c, idx);
        if (!raw) return { empty: true };
        return { text: raw };
      }
    }
  }

  evaluate(n, idx) {
    if (n.num !== undefined) return n.num;
    if (n.id) {
      const c = this.byId.get(n.id);
      if (!c) throw new Error('Campo desconhecido na fórmula: ' + n.id);
      const r = this.get(n.id, this.isList(c) ? idx : undefined);
      if (r.empty) this.sawEmpty = true;
      return Number.isFinite(r.num) ? r.num : 0;
    }
    if (n.op) {
      const a = this.evaluate(n.a, idx);
      if (n.op === 'neg') return -a;
      const b = this.evaluate(n.b, idx);
      if (n.op === '+') return a + b;
      if (n.op === '-') return a - b;
      if (n.op === '*') return a * b;
      if (n.op === '/') return b === 0 ? 0 : a / b;
    }
    if (n.fn) {
      const f = n.fn;
      if (f === 'soma' || f === 'contar' || f === 'media') {
        const g = this.listGroupIn(n.args[0]);
        if (!g) return f === 'contar' ? 0 : this.evaluate(n.args[0] || { num: 0 }, idx);
        const items = this.itemsOf(g);
        if (f === 'contar') return items.length;
        if (!items.length) this.sawEmpty = true;
        let s = 0;
        items.forEach((_, i) => { s += this.evaluate(n.args[0], i); });
        return f === 'media' ? (items.length ? s / items.length : 0) : s;
      }
      const args = n.args.map((a) => this.evaluate(a, idx));
      if (f === 'arred') { const d = args.length > 1 ? args[1] : 2; const p = 10 ** d; return Math.round(args[0] * p) / p; }
      if (f === 'min') return Math.min(...args);
      if (f === 'max') return Math.max(...args);
      if (f === 'abs') return Math.abs(args[0]);
      if (f === 'teto') return Math.ceil(args[0]);
      if (f === 'piso') return Math.floor(args[0]);
      throw new Error('Função desconhecida: ' + f);
    }
    return 0;
  }

  listGroupIn(ast) {
    let found = null;
    const walk = (n) => {
      if (!n || found) return;
      if (n.id) { const c = this.byId.get(n.id); if (c && this.isList(c)) found = c.grupo; }
      if (n.a) walk(n.a);
      if (n.b) walk(n.b);
      if (n.args) n.args.forEach(walk);
    };
    walk(ast);
    return found;
  }

  /** Campos de entrada obrigatórios ainda vazios. */
  missing(usedIds) {
    const out = [];
    for (const c of this.campos) {
      if (!c.obrigatorio || (usedIds && !usedIds.has(c.id))) continue;
      if (this.isList(c)) {
        const items = this.itemsOf(c.grupo);
        if (!items.length) { out.push({ campo: c }); continue; }
        items.forEach((_, i) => { if (this.get(c.id, i).empty) out.push({ campo: c, idx: i }); });
      } else if (this.get(c.id).empty) out.push({ campo: c });
    }
    return out;
  }
}

/** Ids usados no documento + dependências (origem, fórmulas, meses...). */
export function usedFieldIds(model, xml) {
  const used = tokensIn(xml);
  const byId = new Map((model.campos || []).map((c) => [c.id, c]));
  const queue = Array.from(used);
  while (queue.length) {
    const id = queue.shift();
    const c = byId.get(id);
    if (!c) continue;
    const deps = [];
    if (c.origem) deps.push(c.origem);
    if (c.tipo === 'calculo' && c.formula) deps.push(...formulaIds(c.formula));
    if (c.tipo === 'dataCalculada') { if (c.meses && !Number.isFinite(+c.meses)) deps.push(c.meses); if (c.dias && !Number.isFinite(+c.dias)) deps.push(c.dias); }
    for (const d of deps) if (!used.has(d)) { used.add(d); queue.push(d); }
  }
  return used;
}
