// Consultas públicas: CNPJ (BrasilAPI / Receita Federal) e CEP (ViaCEP).
import { soDigitos, mascaraCEP, mascaraTelefone, aplicaCaixa } from './docs.js';

async function getJSON(url, ms = 12000) {
  const ctl = new AbortController();
  const t = setTimeout(() => ctl.abort(), ms);
  try {
    const res = await fetch(url, { signal: ctl.signal, headers: { Accept: 'application/json' } });
    return res;
  } catch (e) {
    if (e.name === 'AbortError') throw new Error('A consulta demorou demais. Verifique sua internet e tente de novo.');
    throw new Error('Sem conexão com a internet para fazer a consulta.');
  } finally {
    clearTimeout(t);
  }
}

const titulo = (s) => aplicaCaixa(String(s || '').trim(), 'titulo');

export async function consultaCNPJ(cnpj) {
  const d = soDigitos(cnpj);
  if (d.length !== 14) throw new Error('Digite os 14 números do CNPJ.');
  const res = await getJSON(`https://brasilapi.com.br/api/cnpj/v1/${d}`);
  if (res.status === 404) throw new Error('CNPJ não encontrado na Receita Federal.');
  if (res.status === 429) throw new Error('Muitas consultas seguidas. Aguarde um minuto e tente novamente.');
  if (!res.ok) throw new Error('Serviço da Receita indisponível no momento. Tente novamente em instantes.');
  const j = await res.json();
  const tipo = j.descricao_tipo_de_logradouro ? titulo(j.descricao_tipo_de_logradouro) + ' ' : '';
  const socios = (j.qsa || []).map((s) => ({ nome: s.nome_socio, qualificacao: s.qualificacao_socio || '' }));
  const adm = socios.find((s) => /administrador|diretor|presidente|titular/i.test(s.qualificacao)) || socios[0];
  return {
    nome: j.razao_social || '',
    fantasia: j.nome_fantasia || '',
    logradouro: (tipo + titulo(j.logradouro)).trim(),
    numero: j.numero || '',
    complemento: titulo(j.complemento),
    bairro: titulo(j.bairro),
    cep: j.cep ? mascaraCEP(j.cep) : '',
    cidade: titulo(j.municipio),
    uf: j.uf || '',
    telefone: j.ddd_telefone_1 ? mascaraTelefone(j.ddd_telefone_1) : '',
    email: (j.email || '').toLowerCase(),
    representante: adm ? adm.nome : '',
    socios,
    situacao: j.descricao_situacao_cadastral || '',
    abertura: j.data_inicio_atividade || '',
  };
}

export async function consultaCEP(cep) {
  const d = soDigitos(cep);
  if (d.length !== 8) throw new Error('Digite os 8 números do CEP.');
  const res = await getJSON(`https://viacep.com.br/ws/${d}/json/`);
  if (!res.ok) throw new Error('Serviço de CEP indisponível no momento.');
  const j = await res.json();
  if (j.erro) throw new Error('CEP não encontrado.');
  return {
    logradouro: j.logradouro || '', complemento: j.complemento || '', bairro: j.bairro || '',
    cidade: j.localidade || '', uf: j.uf || '', estado: j.estado || '', cep: mascaraCEP(d),
  };
}

/** "Avenida X, nº 255, Sala B, Bairro, CEP 78.557-512, Sinop-MT" */
export function enderecoCompleto(e) {
  const num = e.numero ? (/^s\/?n$/i.test(String(e.numero).replace(/\s/g, '')) ? 'S/N' : 'nº ' + e.numero) : '';
  const parts = [e.logradouro, num, e.complemento, e.bairro, e.cep && 'CEP ' + e.cep, [e.cidade, e.uf].filter(Boolean).join('-')];
  return parts.filter((x) => x && String(x).trim()).join(', ');
}

export const UFS = {
  AC: 'Acre', AL: 'Alagoas', AP: 'Amapá', AM: 'Amazonas', BA: 'Bahia', CE: 'Ceará', DF: 'Distrito Federal',
  ES: 'Espírito Santo', GO: 'Goiás', MA: 'Maranhão', MT: 'Mato Grosso', MS: 'Mato Grosso do Sul', MG: 'Minas Gerais',
  PA: 'Pará', PB: 'Paraíba', PR: 'Paraná', PE: 'Pernambuco', PI: 'Piauí', RJ: 'Rio de Janeiro', RN: 'Rio Grande do Norte',
  RS: 'Rio Grande do Sul', RO: 'Rondônia', RR: 'Roraima', SC: 'Santa Catarina', SP: 'São Paulo', SE: 'Sergipe', TO: 'Tocantins',
};
