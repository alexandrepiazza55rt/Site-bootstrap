// Números e valores por extenso em português do Brasil.

const UNID = ['zero', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove', 'dez', 'onze', 'doze',
  'treze', 'quatorze', 'quinze', 'dezesseis', 'dezessete', 'dezoito', 'dezenove'];
const UNID_F = { um: 'uma', dois: 'duas' };
const DEZ = ['', '', 'vinte', 'trinta', 'quarenta', 'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa'];
const CENT = ['', 'cento', 'duzentos', 'trezentos', 'quatrocentos', 'quinhentos', 'seiscentos', 'setecentos', 'oitocentos', 'novecentos'];
const CENT_F = { duzentos: 'duzentas', trezentos: 'trezentas', quatrocentos: 'quatrocentas', quinhentos: 'quinhentas', seiscentos: 'seiscentas', setecentos: 'setecentas', oitocentos: 'oitocentas', novecentos: 'novecentas' };
const ESCALAS = [
  ['', ''], ['mil', 'mil'], ['milhão', 'milhões'], ['bilhão', 'bilhões'], ['trilhão', 'trilhões'], ['quatrilhão', 'quatrilhões'],
];

function ate999(n, fem) {
  if (n === 0) return '';
  if (n === 100) return 'cem';
  const c = Math.floor(n / 100), r = n % 100;
  const parts = [];
  if (c) parts.push(fem && CENT_F[CENT[c]] ? CENT_F[CENT[c]] : CENT[c]);
  if (r) {
    if (r < 20) parts.push(fem && UNID_F[UNID[r]] ? UNID_F[UNID[r]] : UNID[r]);
    else {
      const d = Math.floor(r / 10), u = r % 10;
      parts.push(u ? `${DEZ[d]} e ${fem && UNID_F[UNID[u]] ? UNID_F[UNID[u]] : UNID[u]}` : DEZ[d]);
    }
  }
  return parts.join(' e ');
}

/** Inteiro por extenso. fem=true para "duas", "duzentas"... */
export function inteiroExtenso(n, fem = false) {
  n = Math.floor(Math.abs(Number(n) || 0));
  if (n === 0) return 'zero';
  const grupos = [];
  let x = n;
  while (x > 0) { grupos.push(x % 1000); x = Math.floor(x / 1000); }
  const partes = [];
  for (let i = grupos.length - 1; i >= 0; i--) {
    const g = grupos[i];
    if (!g) continue;
    let txt;
    if (i === 0) txt = ate999(g, fem);
    else if (i === 1) txt = g === 1 ? 'mil' : `${ate999(g, fem)} mil`;
    else txt = `${ate999(g, false)} ${g === 1 ? ESCALAS[i][0] : ESCALAS[i][1]}`;
    partes.push({ txt, g });
  }
  // "e" antes de grupo < 100 ou centena redonda; vírgula nos demais casos
  let out = partes[0].txt;
  for (let k = 1; k < partes.length; k++) {
    const p = partes[k];
    out += (p.g < 100 || p.g % 100 === 0 ? ' e ' : ', ') + p.txt;
  }
  return out;
}

/** Valor monetário por extenso: 1234.5 -> "mil, duzentos e trinta e quatro reais e cinquenta centavos". */
export function moedaExtenso(valor, { moeda = ['real', 'reais'], centavo = ['centavo', 'centavos'] } = {}) {
  const v = Math.round(Math.abs(Number(valor) || 0) * 100);
  const inteiro = Math.floor(v / 100);
  const cent = v % 100;
  const partes = [];
  if (inteiro > 0) {
    let t = inteiroExtenso(inteiro);
    // "um milhão de reais", "dois milhões de reais"
    const deDe = inteiro >= 1e6 && inteiro % 1e6 === 0;
    t += (deDe ? ' de ' : ' ') + (inteiro === 1 ? moeda[0] : moeda[1]);
    partes.push(t);
  }
  if (cent > 0) {
    let t = inteiroExtenso(cent) + ' ' + (cent === 1 ? centavo[0] : centavo[1]);
    if (inteiro === 0) t += ' de ' + moeda[0];
    partes.push(t);
  }
  if (!partes.length) return 'zero ' + moeda[1];
  return partes.join(' e ');
}

/** Número (com decimais) por extenso: 2,5 -> "dois vírgula cinco". */
export function numeroExtenso(valor, fem = false) {
  const n = Number(valor) || 0;
  const s = String(Math.abs(n));
  if (!s.includes('.')) return (n < 0 ? 'menos ' : '') + inteiroExtenso(Math.abs(n), fem);
  const [i, d] = s.split('.');
  const dec = d.replace(/0+$/, '');
  return `${n < 0 ? 'menos ' : ''}${inteiroExtenso(+i, fem)} vírgula ${dec.split('').every((c) => c === '0') ? 'zero' : (dec.startsWith('0') ? dec.split('').map((c) => inteiroExtenso(+c)).join(' ') : inteiroExtenso(+dec))}`;
}

export function percentualExtenso(valor) {
  return numeroExtenso(valor) + ' por cento';
}

const MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];

/** "2026-10-10" -> "10 de outubro de 2026" */
export function dataExtenso(iso, { mesMaiusculo = false, diaExtenso = false } = {}) {
  const d = parseISO(iso);
  if (!d) return '';
  let mes = MESES[d.m - 1];
  if (mesMaiusculo) mes = mes[0].toUpperCase() + mes.slice(1);
  const dia = diaExtenso ? (d.d === 1 ? 'primeiro' : inteiroExtenso(d.d)) : String(d.d).padStart(2, '0');
  return `${dia} de ${mes} de ${d.y}`;
}

export function parseISO(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || '');
  if (!m) return null;
  return { y: +m[1], m: +m[2], d: +m[3] };
}

export function mesNome(m) {
  return MESES[m - 1];
}
