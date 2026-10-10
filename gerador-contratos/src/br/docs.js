// Máscaras, validações e formatações brasileiras (CPF, CNPJ, CEP, telefone, moeda, datas).

export const soDigitos = (s) => String(s || '').replace(/\D+/g, '');

export function validaCPF(v) {
  const d = soDigitos(v);
  if (d.length !== 11 || /^(\d)\1{10}$/.test(d)) return false;
  const calc = (len) => {
    let sum = 0;
    for (let i = 0; i < len; i++) sum += +d[i] * (len + 1 - i);
    const r = (sum * 10) % 11;
    return r === 10 ? 0 : r;
  };
  return calc(9) === +d[9] && calc(10) === +d[10];
}

export function validaCNPJ(v) {
  const d = soDigitos(v);
  if (d.length !== 14 || /^(\d)\1{13}$/.test(d)) return false;
  const calc = (len) => {
    const pesos = len === 12 ? [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2] : [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];
    let sum = 0;
    for (let i = 0; i < len; i++) sum += +d[i] * pesos[i];
    const r = sum % 11;
    return r < 2 ? 0 : 11 - r;
  };
  return calc(12) === +d[12] && calc(13) === +d[13];
}

export function mascaraCPF(v) {
  const d = soDigitos(v).slice(0, 11);
  return d.replace(/^(\d{3})(\d)/, '$1.$2').replace(/^(\d{3})\.(\d{3})(\d)/, '$1.$2.$3').replace(/\.(\d{3})(\d{1,2})$/, '.$1-$2');
}

export function mascaraCNPJ(v) {
  const d = soDigitos(v).slice(0, 14);
  return d.replace(/^(\d{2})(\d)/, '$1.$2').replace(/^(\d{2})\.(\d{3})(\d)/, '$1.$2.$3')
    .replace(/\.(\d{3})(\d)/, '.$1/$2').replace(/(\d{4})(\d{1,2})$/, '$1-$2');
}

export function mascaraCpfCnpj(v) {
  return soDigitos(v).length > 11 ? mascaraCNPJ(v) : mascaraCPF(v);
}

export function mascaraCEP(v, comPonto = true) {
  const d = soDigitos(v).slice(0, 8);
  if (comPonto) return d.replace(/^(\d{2})(\d)/, '$1.$2').replace(/^(\d{2})\.(\d{3})(\d)/, '$1.$2-$3');
  return d.replace(/^(\d{5})(\d)/, '$1-$2');
}

export function mascaraTelefone(v) {
  const d = soDigitos(v).slice(0, 11);
  if (d.length <= 10) return d.replace(/^(\d{2})(\d)/, '($1) $2').replace(/(\d{4})(\d{1,4})$/, '$1-$2');
  return d.replace(/^(\d{2})(\d)/, '($1) $2').replace(/(\d)(\d{4})(\d{4})$/, '$1.$2-$3');
}

const NF2 = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

export function formataMoeda(n) {
  const v = Number(n);
  if (!Number.isFinite(v)) return '';
  return NF2.format(v);
}

/** "270.000,00" | "270000" | "R$ 1.234,5" -> 270000 */
export function parseMoeda(s) {
  if (typeof s === 'number') return s;
  let t = String(s || '').replace(/[^\d,.-]/g, '');
  if (!t) return NaN;
  if (t.includes(',')) t = t.replace(/\./g, '').replace(',', '.');
  else if ((t.match(/\./g) || []).length > 1 || /\.\d{3}$/.test(t)) t = t.replace(/\./g, '');
  return parseFloat(t);
}

/** Máscara de digitação de moeda: dígitos -> "1.234,56" (centavos à direita). */
export function mascaraMoedaDigitando(s) {
  const d = soDigitos(s).replace(/^0+(?=\d)/, '');
  if (!d) return '';
  const n = parseInt(d, 10) / 100;
  return NF2.format(n);
}

export function formataNumero(n, { digitos = 0, decimais = null, milhar = true } = {}) {
  const v = Number(n);
  if (!Number.isFinite(v)) return '';
  if (decimais === null || decimais === undefined || decimais === '') {
    const isInt = Number.isInteger(v);
    if (isInt && digitos) return String(Math.abs(v)).padStart(digitos, '0').replace(/^/, v < 0 ? '-' : '');
    return new Intl.NumberFormat('pt-BR', { useGrouping: milhar, maximumFractionDigits: 6 }).format(v);
  }
  return new Intl.NumberFormat('pt-BR', { useGrouping: milhar, minimumFractionDigits: +decimais, maximumFractionDigits: +decimais }).format(v);
}

export function parseNumero(s) {
  if (typeof s === 'number') return s;
  const t = String(s || '').trim();
  if (!t) return NaN;
  if (/^-?\d+(\.\d+)?$/.test(t)) return parseFloat(t);
  return parseMoeda(t);
}

export function isoHoje() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

export function formataData(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || '');
  return m ? `${m[3]}/${m[2]}/${m[1]}` : '';
}

/** "14/10/2026" -> "2026-10-14" */
export function dataBRparaISO(s) {
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(s || '').trim());
  if (!m) return null;
  return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
}

/** Soma meses/dias preservando o dia (ajusta fim de mês). */
export function somaData(iso, meses = 0, dias = 0) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || '');
  if (!m) return '';
  let y = +m[1], mo = +m[2] - 1 + (Number(meses) || 0), d = +m[3];
  y += Math.floor(mo / 12);
  mo = ((mo % 12) + 12) % 12;
  const ultimo = new Date(Date.UTC(y, mo + 1, 0)).getUTCDate();
  const dt = new Date(Date.UTC(y, mo, Math.min(d, ultimo)));
  dt.setUTCDate(dt.getUTCDate() + (Number(dias) || 0));
  return dt.toISOString().slice(0, 10);
}

export function diasEntre(isoA, isoB) {
  const a = Date.parse(isoA), b = Date.parse(isoB);
  if (!Number.isFinite(a) || !Number.isFinite(b)) return null;
  return Math.round((b - a) / 86400000);
}

const MINUSCULAS = new Set(['de', 'da', 'do', 'das', 'dos', 'e', 'em', 'a', 'o', 'as', 'os', 'para', 'por', 'com']);

export function aplicaCaixa(texto, caixa) {
  const s = String(texto ?? '');
  switch (caixa) {
    case 'maiusculas': return s.toLocaleUpperCase('pt-BR');
    case 'minusculas': return s.toLocaleLowerCase('pt-BR');
    case 'titulo':
      return s.toLocaleLowerCase('pt-BR').replace(/[\p{L}\p{N}]+/gu, (w, idx) => (idx > 0 && MINUSCULAS.has(w) ? w : w[0].toLocaleUpperCase('pt-BR') + w.slice(1)));
    case 'frase': return s ? s[0].toLocaleUpperCase('pt-BR') + s.slice(1) : s;
    default: return s;
  }
}

export function moneyFromText(s) {
  return parseMoeda(s);
}
