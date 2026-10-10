// Modelo inicial: Contrato de Locação de Bem Móvel sem Operação (Rent Loc).
import docx from './rentloc-locacao.docx';

export const SEED_ID = 'modelo-rentloc-locacao';

export const grupos = [
  { id: 'contrato', nome: 'Contrato', icone: 'file-text' },
  { id: 'locatario', nome: 'Locatário (cliente)', icone: 'user', cadastro: 'Clientes' },
  { id: 'equipamentos', nome: 'Equipamentos locados', icone: 'truck', lista: true, cadastro: 'Equipamentos', itemNome: 'Equipamento' },
  { id: 'valores', nome: 'Valores e prazo', icone: 'coins' },
  { id: 'compra', nome: 'Opção de compra', icone: 'tag' },
  { id: 'destinacao', nome: 'Destinação do bem', icone: 'map-pin' },
  { id: 'testemunhas', nome: 'Testemunhas', icone: 'users', recolhido: true },
  { id: 'locador', nome: 'Locador (sua empresa)', icone: 'building', recolhido: true },
];

const c = (id, rotulo, grupo, tipo, extra = {}) => ({ id, rotulo, grupo, tipo, obrigatorio: false, padrao: '', dica: '', caixa: '', largura: 'full', ...extra });

export const campos = [
  // contrato
  c('numero_contrato', 'Nº do contrato', 'contrato', 'sequencial', { largura: 'half', obrigatorio: true, formato: '{n:4}/{ano}' }),
  c('data_assinatura', 'Data da assinatura', 'contrato', 'data', { largura: 'half', padrao: 'hoje', obrigatorio: true }),
  c('cidade_assinatura', 'Cidade da assinatura', 'contrato', 'texto', { padrao: 'Sinop', largura: 'half' }),
  c('data_assinatura_extenso', 'Data da assinatura por extenso', 'contrato', 'dataExtenso', { origem: 'data_assinatura', mesMaiusculo: true }),

  // locatário
  c('locatario_cnpj', 'CNPJ', 'locatario', 'cnpj', { obrigatorio: true, papel: 'documento', largura: 'half', dica: '00.000.000/0000-00' }),
  c('locatario_nome', 'Razão social', 'locatario', 'texto', { obrigatorio: true, papel: 'nome', caixa: 'maiusculas', dica: 'Ex.: EMPRESA EXEMPLO LTDA' }),
  c('locatario_endereco', 'Endereço completo', 'locatario', 'textoLongo', { obrigatorio: true, papel: 'endereco', dica: 'Avenida, nº, sala, bairro, CEP, cidade-UF' }),
  c('locatario_representante', 'Representante legal', 'locatario', 'texto', { obrigatorio: true, papel: 'representante', caixa: 'maiusculas', largura: 'full' }),
  c('locatario_rep_cpf', 'CPF do representante', 'locatario', 'cpf', { obrigatorio: true, largura: 'half' }),
  c('locatario_rep_rg', 'RG / órgão emissor', 'locatario', 'rg', { largura: 'half', dica: 'Ex.: 1234567 SSP/MT' }),
  c('locatario_rep_profissao', 'Profissão', 'locatario', 'texto', { padrao: 'empresário', largura: 'half', caixa: 'minusculas' }),
  c('locatario_rep_estado_civil', 'Estado civil', 'locatario', 'lista', {
    padrao: 'casado', largura: 'half', opcoes: ['solteiro', 'solteira', 'casado', 'casada', 'divorciado', 'divorciada', 'viúvo', 'viúva', 'em união estável'],
  }),
  c('locatario_rep_endereco', 'Endereço do representante', 'locatario', 'textoLongo', { dica: 'Rua, nº, bairro, CEP, cidade-UF' }),

  // equipamentos (lista)
  c('item_n', 'Item', 'equipamentos', 'indice'),
  c('item_descricao', 'Descrição do equipamento', 'equipamentos', 'textoLongo', { obrigatorio: true, caixa: 'maiusculas', dica: 'Ex.: ESCAVADEIRA HIDRÁULICA MODELO 320D2L, ANO 2014, MARCA CATERPILLAR, SÉRIE ...' }),
  c('item_qtd', 'Qtd.', 'equipamentos', 'numero', { padrao: '1', largura: 'third' }),
  c('item_valor_mes', 'Valor mensal (R$)', 'equipamentos', 'moeda', { obrigatorio: true, largura: 'third' }),
  c('item_hora_extra', 'Hora extra (R$)', 'equipamentos', 'moeda', { largura: 'third' }),

  // valores e prazo
  c('prazo_meses', 'Prazo (meses)', 'valores', 'numero', { padrao: '6', digitos: 2, obrigatorio: true, largura: 'third' }),
  c('horas_mes', 'Horas por mês', 'valores', 'numero', { padrao: '220', largura: 'third' }),
  c('dia_vencimento', 'Dia do vencimento', 'valores', 'numero', { padrao: '10', largura: 'third' }),
  c('data_inicio', 'Início da locação', 'valores', 'data', { obrigatorio: true, largura: 'half', papel: 'inicio' }),
  c('data_termino', 'Término (calculado)', 'valores', 'dataCalculada', { origem: 'data_inicio', meses: 'prazo_meses', largura: 'half', papel: 'termino' }),
  c('valor_total', 'Valor global da locação (R$)', 'valores', 'calculo', {
    formula: 'soma(item_qtd * item_valor_mes) * prazo_meses', formato: 'moeda', editavel: true, papel: 'valor',
    ajuda: 'Calculado: soma dos valores mensais × prazo. Você pode digitar outro valor.',
  }),
  c('valor_total_extenso', 'Valor global por extenso', 'valores', 'extenso', { origem: 'valor_total', caixa: 'maiusculas' }),
  c('valor_entrega', 'Valor pago na entrega (R$)', 'valores', 'moeda', { obrigatorio: true, largura: 'half' }),
  c('valor_entrega_extenso', 'Valor na entrega por extenso', 'valores', 'extenso', { origem: 'valor_entrega', caixa: 'minusculas' }),

  // opção de compra
  c('valor_compra', 'Valor total para compra (R$)', 'compra', 'moeda', { largura: 'half' }),
  c('valor_compra_extenso', 'Valor de compra por extenso', 'compra', 'extenso', { origem: 'valor_compra', caixa: 'titulo' }),
  c('valor_compra_entrada', 'Entrada (R$)', 'compra', 'moeda', { largura: 'half' }),
  c('valor_compra_entrada_extenso', 'Entrada por extenso', 'compra', 'extenso', { origem: 'valor_compra_entrada', caixa: 'titulo' }),
  c('valor_compra_restante', 'Restante (calculado)', 'compra', 'calculo', { formula: 'valor_compra - valor_compra_entrada', formato: 'moeda', editavel: true, largura: 'half' }),
  c('valor_compra_restante_extenso', 'Restante por extenso', 'compra', 'extenso', { origem: 'valor_compra_restante', caixa: 'minusculas' }),
  c('compra_parcelas', 'Nº de parcelas', 'compra', 'numero', { padrao: '3', largura: 'half' }),
  c('compra_parcelas_extenso', 'Nº de parcelas por extenso', 'compra', 'extenso', { origem: 'compra_parcelas', feminino: true, caixa: 'minusculas' }),
  c('valor_compra_parcela', 'Valor de cada parcela (calculado)', 'compra', 'calculo', { formula: 'arred(valor_compra_restante / compra_parcelas, 2)', formato: 'moeda', editavel: true, largura: 'half' }),
  c('valor_compra_parcela_extenso', 'Parcela por extenso', 'compra', 'extenso', { origem: 'valor_compra_parcela', caixa: 'titulo' }),

  // destinação
  c('destinacao_servico', 'Serviço a ser executado', 'destinacao', 'texto', { obrigatorio: true, dica: 'Ex.: desmate, terraplenagem', caixa: 'minusculas' }),
  c('destinacao_local', 'Local de uso do bem', 'destinacao', 'texto', { obrigatorio: true, dica: 'Ex.: Município de Sinop-MT' }),

  // testemunhas
  c('testemunha1_nome', 'Testemunha 1 — nome', 'testemunhas', 'texto', { vazio: '______________________', largura: 'half' }),
  c('testemunha1_cpf', 'Testemunha 1 — CPF', 'testemunhas', 'cpf', { vazio: '______________________', largura: 'half' }),
  c('testemunha2_nome', 'Testemunha 2 — nome', 'testemunhas', 'texto', { vazio: '______________________', largura: 'half' }),
  c('testemunha2_cpf', 'Testemunha 2 — CPF', 'testemunhas', 'cpf', { vazio: '______________________', largura: 'half' }),

  // locador (padrões da empresa)
  c('locador_nome', 'Razão social', 'locador', 'texto', { padrao: 'RENT LOC EQUIPAMENTOS LTDA', caixa: 'maiusculas' }),
  c('locador_cnpj', 'CNPJ', 'locador', 'cnpj', { padrao: '26.844.547/0001-67', largura: 'half' }),
  c('locador_telefone', 'Telefone', 'locador', 'telefone', { padrao: '(66) 9.9679-9008', largura: 'half' }),
  c('locador_endereco', 'Endereço da sede', 'locador', 'textoLongo', { padrao: 'Avenida Bruno Martini, S/N, KM 7.5, Chácara 08, Condomínio São Lucas, CEP 78.550-970' }),
  c('locador_cidade', 'Cidade', 'locador', 'texto', { padrao: 'Sinop-MT', largura: 'half' }),
  c('locador_estado', 'Estado', 'locador', 'texto', { padrao: 'Mato Grosso', largura: 'half' }),
  c('locador_representante', 'Administrador(a)', 'locador', 'texto', { padrao: 'MIRIAN SCHWINGEL', caixa: 'maiusculas' }),
  c('locador_rep_cpf', 'CPF do(a) administrador(a)', 'locador', 'cpf', { padrao: '580.556.731-87', largura: 'half' }),
  c('locador_rep_endereco', 'Endereço do(a) administrador(a)', 'locador', 'textoLongo', { padrao: 'Rua Berna, 446, Residencial Bella Suíça, CEP 78.556-598, em Sinop-MT' }),
  c('locador_banco', 'Dados bancários (boleto/PIX)', 'locador', 'texto', { padrao: 'Santander S.A. 033 - AG 4168 C/C 13.004848-3' }),
];

export const seedTemplate = {
  id: SEED_ID,
  nome: 'Locação de Bem Móvel sem Operação',
  descricao: 'Contrato de locação de máquinas e equipamentos (Rent Loc), com tabela de equipamentos, opção de compra e cláusulas de rescisão.',
  cor: '#2563eb',
  grupos,
  campos,
  config: { campoCliente: 'locatario_nome', campoValor: 'valor_total', campoInicio: 'data_inicio', campoTermino: 'data_termino', campoNumero: 'numero_contrato' },
  docx, // Uint8Array
};
