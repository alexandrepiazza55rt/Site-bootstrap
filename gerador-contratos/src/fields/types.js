// Catálogo de tipos de campo.

export const TIPOS = {
  texto: { nome: 'Texto', icone: 'type', grupo: 'Entrada' },
  textoLongo: { nome: 'Texto longo', icone: 'align-left', grupo: 'Entrada' },
  cpf: { nome: 'CPF', icone: 'id', grupo: 'Documentos' },
  cnpj: { nome: 'CNPJ', icone: 'building', grupo: 'Documentos' },
  cpfCnpj: { nome: 'CPF ou CNPJ', icone: 'id', grupo: 'Documentos' },
  rg: { nome: 'RG', icone: 'id', grupo: 'Documentos' },
  cep: { nome: 'CEP', icone: 'map-pin', grupo: 'Documentos' },
  telefone: { nome: 'Telefone', icone: 'phone', grupo: 'Entrada' },
  email: { nome: 'E-mail', icone: 'mail', grupo: 'Entrada' },
  moeda: { nome: 'Valor (R$)', icone: 'coins', grupo: 'Números' },
  numero: { nome: 'Número', icone: 'hash', grupo: 'Números' },
  data: { nome: 'Data', icone: 'calendar', grupo: 'Datas' },
  lista: { nome: 'Lista de opções', icone: 'list', grupo: 'Entrada' },
  sequencial: { nome: 'Nº do contrato (automático)', icone: 'hash', grupo: 'Automáticos' },
  extenso: { nome: 'Por extenso (automático)', icone: 'wand', grupo: 'Automáticos', calculado: true },
  dataExtenso: { nome: 'Data por extenso (automático)', icone: 'wand', grupo: 'Automáticos', calculado: true },
  dataCalculada: { nome: 'Data calculada (automático)', icone: 'calendar-plus', grupo: 'Automáticos', calculado: true },
  calculo: { nome: 'Cálculo / fórmula', icone: 'calculator', grupo: 'Automáticos', calculado: true },
  indice: { nome: 'Nº do item (automático)', icone: 'hash', grupo: 'Automáticos', calculado: true },
};

export const CAIXAS = [
  ['', 'Como digitado'],
  ['maiusculas', 'MAIÚSCULAS'],
  ['minusculas', 'minúsculas'],
  ['titulo', 'Primeiras Maiúsculas'],
  ['frase', 'Primeira letra maiúscula'],
];

export const PAPEIS = [
  ['', '—'],
  ['nome', 'Nome / Razão social'],
  ['documento', 'CPF/CNPJ'],
  ['endereco', 'Endereço completo'],
  ['cidade', 'Cidade'],
  ['uf', 'Estado (UF)'],
  ['estado', 'Estado (por extenso)'],
  ['cep', 'CEP'],
  ['telefone', 'Telefone'],
  ['email', 'E-mail'],
  ['representante', 'Representante legal'],
  ['termino', 'Data de término (vencimentos)'],
  ['inicio', 'Data de início'],
  ['valor', 'Valor total do contrato'],
];

export const isCalculado = (c) => !!(TIPOS[c.tipo] && TIPOS[c.tipo].calculado);

export function novoCampo(dados = {}) {
  return {
    id: '', rotulo: '', grupo: 'geral', tipo: 'texto', obrigatorio: false, padrao: '', dica: '', caixa: '',
    largura: 'full', ...dados,
  };
}

/** Sugere tipo e caixa a partir de um texto selecionado no documento. */
export function sugerirTipo(texto) {
  const t = String(texto || '').trim();
  if (/^\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}$/.test(t)) return { tipo: 'cnpj' };
  if (/^\d{3}\.\d{3}\.\d{3}-\d{2}$/.test(t)) return { tipo: 'cpf' };
  if (/^\d{2}\.?\d{3}-\d{3}$/.test(t)) return { tipo: 'cep' };
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(t)) return { tipo: 'data' };
  if (/^(R\$\s?)?\d{1,3}(\.\d{3})*,\d{2}$/.test(t)) return { tipo: 'moeda' };
  if (/^\(?\d{2}\)?\s?9?\.?\d{4}-?\d{4}$/.test(t)) return { tipo: 'telefone' };
  if (/^[\w.+-]+@[\w-]+\.[\w.]+$/.test(t)) return { tipo: 'email' };
  if (/^\d+$/.test(t)) return { tipo: 'numero', digitos: t.length > 1 && t.startsWith('0') ? t.length : 0 };
  if (/^\d{1,2} de [a-zç]+ de \d{4}$/i.test(t)) return { tipo: 'data', extensoData: true };
  const letras = t.replace(/[^A-Za-zÀ-ÿ]/g, '');
  if (letras.length > 3 && letras === letras.toUpperCase()) return { tipo: 'texto', caixa: 'maiusculas' };
  return { tipo: 'texto' };
}
