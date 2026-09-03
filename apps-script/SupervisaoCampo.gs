/**
 * ============================================================================
 * MÓDULO: SUPERVISÃO DE CAMPO
 * Terceiro script do projeto. Cole como novo arquivo (Arquivo > + > Script,
 * nome "SupervisaoCampo") e rode  configurarSupervisao()  uma vez.
 *
 * DUAS PARTES
 *  1. Projetos em andamento — abas PROJ_CARDS e PROJ_META, nesta planilha.
 *     Supervisor e administrativo criam e editam; técnico só vê.
 *  2. Status RDO — apenas LEITURA da planilha do RDO, que é outra planilha.
 *     Nada é escrito lá.
 *
 * SOBRE AS COLUNAS DA PLANILHA DO RDO
 *  Os nomes das colunas são descobertos pelo cabeçalho, comparando com a lista
 *  de apelidos em COLUNAS_RDO. A comparação ignora maiúsculas, acentos e
 *  espaços, então "Data_exp", "DATA EXP" e "data exp" dão no mesmo.
 *
 *  Se alguma coluna não for reconhecida, use  Portal > Conferir colunas do RDO
 *  para ver o cabeçalho real e acrescente o nome na lista certa abaixo.
 * ============================================================================
 */

/* ---------------------------------------------------------------------------
   ENDEREÇOS DAS PLANILHAS EXTERNAS

   Ficam nas Propriedades do Script, não no código. O código vai para o GitHub;
   as propriedades, não. Assim o endereço das planilhas não fica público junto
   com o site.

   COMO PREENCHER (no editor do Apps Script):
     Configurações do projeto (ícone de engrenagem, na barra da esquerda)
     > Propriedades do script > Adicionar propriedade

     ID_RDO          link ou ID da planilha do banco de dados do RDO
     ABA_RDO         nome da aba dos relatórios (padrão: Relatorios)
     ID_MINIMASTER   link ou ID da planilha MINI MASTER
     ABA_MINIMASTER  nome da aba dos técnicos (vazio = primeira aba)
     ID_INPUTS       link ou ID da planilha "Banco de inputs"
     ABA_INPUTS      nome da aba dos tipos de reparo (padrão: ATV POR HR)

   Pode colar o link inteiro do navegador: o ID é extraído sozinho.
   Depois, rode verConfiguracao() para conferir se está lendo as duas planilhas.
   --------------------------------------------------------------------------- */
function prop_(chave, padrao) {
  const v = PropertiesService.getScriptProperties().getProperty(chave);
  return (v === null || v === '') ? (padrao || '') : v;
}

function idPlanilhaRDO_() {
  // extrairId_ aceita tanto o ID quanto a URL inteira: quem preenche a
  // propriedade pelo painel do Apps Script normalmente cola o link do navegador.
  const id = extrairId_(prop_('ID_RDO'));
  if (!id) throw new Error('A planilha do RDO ainda não foi configurada. ' +
                           'Crie a propriedade ID_RDO em Configurações do projeto > ' +
                           'Propriedades do script.');
  return id;
}
function abaNomeRDO_()        { return prop_('ABA_RDO', 'Relatorios'); }
function idPlanilhaMiniMaster_() {
  const id = extrairId_(prop_('ID_MINIMASTER'));
  if (!id) throw new Error('A planilha MINI MASTER ainda não foi configurada. ' +
                           'Crie a propriedade ID_MINIMASTER em Configurações do projeto > ' +
                           'Propriedades do script.');
  return id;
}
function abaNomeMiniMaster_() { return prop_('ABA_MINIMASTER', ''); }

function idPlanilhaInputs_() {
  const id = extrairId_(prop_('ID_INPUTS'));
  if (!id) throw new Error('A planilha "Banco de inputs" ainda não foi configurada. ' +
                           'Crie a propriedade ID_INPUTS em Configurações do projeto > ' +
                           'Propriedades do script.');
  return id;
}
function abaNomeInputs_() { return prop_('ABA_INPUTS', 'ATV POR HR'); }

/* Colunas da aba ATV POR HR, descobertas pelo cabeçalho (sem acento, minúsculo).
   Acrescente apelidos aqui se o cabeçalho da planilha mudar. */
const COLUNAS_INPUTS = {
  tipo:        ['tipo_de_reparo', 'tipo_reparo', 'tiporeparo', 'tipo'],
  obrigatoria: ['atividade_obrigatoria', 'atividade_obrig', 'obrigatoria', 'atv_obrigatoria']
};

/* Apelidos aceitos para cada coluna. Acrescente aqui se o cabeçalho mudar.
   A comparação ignora maiúsculas, acentos e espaços. */
const COLUNAS_RDO = {
  data:    ['data_exp', 'dataexp', 'data', 'data_relatorio', 'data_do_relatorio', 'data_execucao'],
  parque:  ['parque', 'parque_eolico', 'nome_parque', 'usina', 'complexo', 'pe'],
  cliente: ['cliente', 'contratante', 'empresa'],
  link:    ['link_pdf', 'linkpdf', 'link', 'pdf', 'url_pdf', 'arquivo', 'link_relatorio'],
  tipo:    ['tipo_reparo', 'tiporeparo', 'tipo_de_reparo', 'tipo'],
  equipe:  ['equipe', 'numero_equipe', 'n_equipe', 'num_equipe', 'time'],
  autor:   ['matricula_login', 'responsavel', 'tecnico', 'supervisor', 'criado_por', 'usuario'],
  equipes: ['matriculas', 'equipe_matriculas'],
  turbina: ['turbina', 'aerogerador', 'wtg', 'maquina', 'torre'],
  blade:   ['blade', 'pa', 'pá'],
  local:   ['local'],
  avanco:  ['avanco_reparo', 'avanco', 'progresso'],
  fim:     ['reparo_finalizado', 'finalizado']
};

const ABA_PROJ_CARDS = 'PROJ_CARDS';
const ABA_PROJ_META  = 'PROJ_META';
const COLUNAS_PROJ = ['id', 'codigo', 'status', 'rev', 'json', 'atualizado_em', 'atualizado_por'];

const CACHE_RDO_SEGUNDOS = 120;   // releitura da planilha do RDO


/* ===========================================================================
   INSTALAÇÃO
   =========================================================================== */

/* Aceita a URL inteira do Google Sheets ou só o ID. */
function extrairId_(texto) {
  const t = String(texto || '').trim();
  const m = t.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : t;
}

/**
 * Mostra o que está configurado e se dá para ler as duas planilhas.
 * Roda pelo botão Executar do editor, sem precisar de interface.
 * O resultado sai em "Registro de execução".
 */
function verConfiguracao() {
  const props = PropertiesService.getScriptProperties();
  let txt = 'CONFIGURAÇÃO ATUAL\n';
  txt += '  ID_RDO         : ' + (props.getProperty('ID_RDO') || '(não configurado)') + '\n';
  txt += '  ABA_RDO        : ' + abaNomeRDO_() + '\n';
  txt += '  ID_MINIMASTER  : ' + (props.getProperty('ID_MINIMASTER') || '(não configurado)') + '\n';
  txt += '  ABA_MINIMASTER : ' + (abaNomeMiniMaster_() || '(primeira aba)') + '\n';
  txt += '  ID_INPUTS      : ' + (props.getProperty('ID_INPUTS') || '(não configurado)') + '\n';
  txt += '  ABA_INPUTS     : ' + abaNomeInputs_() + '\n\n';

  try {
    const d = diagnosticoColunasRDO_();
    txt += 'RDO: ' + d.linhas + ' relatórios.\n';
    Object.keys(d.mapa).forEach(function (campo) {
      txt += '  ' + campo + ': ' + (d.mapa[campo] === -1
        ? 'NÃO ENCONTRADA' : '"' + d.cabecalho[d.mapa[campo]] + '"') + '\n';
    });
    txt += '\n  Cabeçalho completo: ' + d.cabecalho.filter(String).join(' | ') + '\n';

    txt += '\n  Parques encontrados (usados como chave do casamento):\n';
    amostraParques_(15).forEach(function (nome) { txt += '    ' + nome + '\n'; });
  } catch (e) {
    txt += 'RDO: ERRO — ' + e.message + '\n';
  }

  try {
    const t = lerTecnicos_();
    txt += '\nMINI MASTER: ' + t.length + ' técnicos.';
    if (t.length) txt += '\n  Primeiro: ' + t[0].nome + ' — ' + t[0].matricula;
    if (t.length > 1) txt += '\n  Último:   ' + t[t.length-1].nome + ' — ' + t[t.length-1].matricula;
  } catch (e) {
    txt += '\nMINI MASTER: ERRO — ' + e.message;
  }

  try {
    const tr = lerTiposReparo_();
    txt += '\n\nBANCO DE INPUTS: ' + tr.length + ' tipos de reparo ' +
           '(linhas com Atividade obrigatória = Não).';
    tr.slice(0, 10).forEach(function (n) { txt += '\n    ' + n; });
    if (tr.length > 10) txt += '\n    ... e mais ' + (tr.length - 10) + '.';
  } catch (e) {
    txt += '\n\nBANCO DE INPUTS: ERRO — ' + e.message;
  }

  Logger.log('\n' + txt + '\n');
  avisar_(txt);
  return txt;
}


/** Menu: Portal > Configurar planilhas */
function menuConfigurarPlanilhas() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const r1 = ui.prompt('Planilha do RDO',
    'Cole o link (ou o ID) da planilha do banco de dados do RDO.\n\n' +
    'Atual: ' + (props.getProperty('ID_RDO') || '(não configurada)'),
    ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const idRdo = extrairId_(r1.getResponseText());
  if (!idRdo) { ui.alert('Link vazio.'); return; }

  const r2 = ui.prompt('Aba do RDO',
    'Nome da aba onde ficam os relatórios.\n\nAtual: ' + abaNomeRDO_(),
    ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const abaRdo = r2.getResponseText().trim() || abaNomeRDO_();

  const r3 = ui.prompt('Planilha MINI MASTER',
    'Cole o link (ou o ID) da MINI MASTER — coluna A: matrícula, coluna B: nome.\n\n' +
    'Atual: ' + (props.getProperty('ID_MINIMASTER') || '(não configurada)'),
    ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK) return;
  const idMini = extrairId_(r3.getResponseText());

  const r4 = ui.prompt('Aba da MINI MASTER',
    'Nome da aba com os técnicos. Deixe em branco para usar a primeira aba.\n\n' +
    'Atual: ' + (abaNomeMiniMaster_() || '(primeira aba)'),
    ui.ButtonSet.OK_CANCEL);
  if (r4.getSelectedButton() !== ui.Button.OK) return;

  const r5 = ui.prompt('Planilha Banco de inputs',
    'Cole o link (ou o ID) da planilha "Banco de inputs" — de onde saem os tipos de reparo.\n\n' +
    'Atual: ' + (props.getProperty('ID_INPUTS') || '(não configurada)'),
    ui.ButtonSet.OK_CANCEL);
  if (r5.getSelectedButton() !== ui.Button.OK) return;
  const idInputs = extrairId_(r5.getResponseText());

  const r6 = ui.prompt('Aba do Banco de inputs',
    'Nome da aba com as atividades.\n\nAtual: ' + abaNomeInputs_(),
    ui.ButtonSet.OK_CANCEL);
  if (r6.getSelectedButton() !== ui.Button.OK) return;

  props.setProperties({
    ID_RDO: idRdo,
    ABA_RDO: abaRdo,
    ID_MINIMASTER: idMini,
    ABA_MINIMASTER: r4.getResponseText().trim(),
    ID_INPUTS: idInputs,
    ABA_INPUTS: r6.getResponseText().trim() || abaNomeInputs_()
  }, false);

  CacheService.getScriptCache().removeAll(['rdo_dados', 'minimaster_tecnicos', 'inputs_tipos_reparo']);

  let msg = 'Salvo nas Propriedades do Script.\n\n';
  try {
    const d = diagnosticoColunasRDO_();
    msg += 'RDO: ' + d.linhas + ' relatórios lidos.\n';
    msg += d.faltando.length ? 'Colunas não reconhecidas: ' + d.faltando.join(', ') + '\n'
                             : 'Colunas do RDO: todas reconhecidas.\n';
  } catch (e) { msg += 'RDO: ERRO — ' + e.message + '\n'; }
  try {
    msg += 'MINI MASTER: ' + lerTecnicos_().length + ' técnicos.\n';
  } catch (e) { msg += 'MINI MASTER: ERRO — ' + e.message + '\n'; }
  try {
    msg += 'BANCO DE INPUTS: ' + lerTiposReparo_().length + ' tipos de reparo.\n';
  } catch (e) { msg += 'BANCO DE INPUTS: ERRO — ' + e.message + '\n'; }

  msg += '\nConfira a separação de parque e equipe em Portal > Conferir colunas do RDO.';
  avisar_(msg);
}


function configurarSupervisao() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  criarAba_(ss, ABA_PROJ_CARDS, COLUNAS_PROJ);
  criarAba_(ss, ABA_PROJ_META, ['chave', 'valor']);

  const meta = ss.getSheetByName(ABA_PROJ_META);
  if (meta.getLastRow() < 2) {
    meta.getRange(2, 1, 2, 2).setValues([['rev', 1], ['nextSeq', 1]]);
  }
  ss.getSheetByName(ABA_PROJ_CARDS).getRange('E1')
    .setNote('Ficha completa do projeto. Editar à mão aqui quebra o módulo.');

  avisar_('Abas de projetos criadas.\n\n' +
          'Próximo passo: informe os endereços das planilhas em\n' +
          'Configurações do projeto > Propriedades do script:\n' +
          '  ID_RDO, ABA_RDO, ID_MINIMASTER, ABA_MINIMASTER,\n' +
          '  ID_INPUTS, ABA_INPUTS\n\n' +
          'Depois rode verConfiguracao() para conferir.');
}


/* ===========================================================================
   DIAGNÓSTICO DAS COLUNAS
   =========================================================================== */

/** Normaliza um cabeçalho: sem acento, minúsculo, separadores virando "_". */
function normalizarCab_(t) {
  return String(t == null ? '' : t)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().trim()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function abaRDO_() {
  const ss = SpreadsheetApp.openById(idPlanilhaRDO_());
  const nome = abaNomeRDO_();
  const aba = ss.getSheetByName(nome);
  if (!aba) throw new Error('A aba "' + nome + '" não existe na planilha do RDO.');
  return aba;
}

/**
 * Descobre em que coluna está cada campo, comparando o cabeçalho com os
 * apelidos de COLUNAS_RDO. Devolve -1 para o campo que não achou.
 */
function detectarColunas_(cabecalho) {
  const norm = cabecalho.map(normalizarCab_);
  const mapa = {};
  Object.keys(COLUNAS_RDO).forEach(function (campo) {
    mapa[campo] = -1;
    // primeiro, nome exatamente igual
    for (let i = 0; i < COLUNAS_RDO[campo].length; i++) {
      const idx = norm.indexOf(COLUNAS_RDO[campo][i]);
      if (idx > -1) { mapa[campo] = idx; return; }
    }
    // depois, cabeçalho que contém o apelido (ex.: "data_exp_final")
    for (let i = 0; i < COLUNAS_RDO[campo].length; i++) {
      const alvo = COLUNAS_RDO[campo][i];
      const idx = norm.findIndex(function (c) { return c.indexOf(alvo) > -1; });
      if (idx > -1) { mapa[campo] = idx; return; }
    }
  });
  return mapa;
}


/** Primeiros valores distintos da coluna Parque. */
function amostraParques_(quantos) {
  const aba = abaRDO_();
  const ultima = Math.min(aba.getLastRow(), 400);
  const cab = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  const mapa = detectarColunas_(cab);
  if (mapa.parque === -1) return [];

  const col = aba.getRange(2, mapa.parque + 1, Math.max(0, ultima - 1), 1).getValues();
  const vistos = {}, saida = [];
  for (let i = 0; i < col.length && saida.length < quantos; i++) {
    const bruto = String(col[i][0] || '').trim().replace(/\s+/g, ' ');
    if (!bruto || vistos[bruto]) continue;
    vistos[bruto] = true;
    saida.push(bruto);
  }
  return saida;
}

function diagnosticoColunasRDO_() {
  const aba = abaRDO_();
  const cab = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  const mapa = detectarColunas_(cab);
  const essenciais = ['data', 'parque', 'link'];
  return {
    cabecalho: cab,
    mapa: mapa,
    faltando: essenciais.filter(function (c) { return mapa[c] === -1; }),
    linhas: Math.max(0, aba.getLastRow() - 1)
  };
}

/** Menu: Portal > Conferir colunas do RDO */
function menuColunasRDO() {
  let d;
  try {
    d = diagnosticoColunasRDO_();
  } catch (e) {
    avisar_('Não consegui ler a planilha do RDO.\n\n' + e.message);
    return;
  }

  let txt = 'Planilha do RDO — aba "' + ABA_RDO + '"\n';
  txt += d.linhas + ' relatórios.\n\n';
  txt += 'RECONHECIDAS:\n';
  Object.keys(d.mapa).forEach(function (campo) {
    txt += '  ' + campo.padEnd(9) + (d.mapa[campo] === -1
      ? '(não encontrada)'
      : '-> coluna "' + d.cabecalho[d.mapa[campo]] + '"') + '\n';
  });
  txt += '\nCABEÇALHO COMPLETO DA PLANILHA:\n' +
         d.cabecalho.filter(String).join('  |  ');

  if (d.faltando.length) {
    txt += '\n\nFALTA reconhecer: ' + d.faltando.join(', ') +
           '\nAcrescente o nome real na lista COLUNAS_RDO, no topo do arquivo SupervisaoCampo.gs.';
  }

  // Como a coluna Parque está sendo dividida em parque + equipe.
  // Confira estas linhas: é daqui que sai o casamento com os projetos.
  try {
    txt += '\n\nPARQUES ENCONTRADOS (chave do casamento com os projetos):\n';
    amostraParques_(15).forEach(function (nome) { txt += '  ' + nome + '\n'; });
  } catch (e) { /* já avisou acima */ }

  try {
    const t = lerTecnicos_();
    txt += '\n\nMINI MASTER: ' + t.length + ' técnicos.' +
           (t.length ? '\n  Exemplo: ' + t[0].nome + ' — ' + t[0].matricula : '');
  } catch (e) {
    txt += '\n\nMINI MASTER: ERRO — ' + e.message;
  }

  avisar_(txt);
}


/* ===========================================================================
   LEITURA DO RDO
   =========================================================================== */

/* Avanço pode vir como "60%", como 60 ou como 0,6 (fração formatada em %).
   Devolve sempre um texto com o símbolo, para o card não ter que adivinhar. */
function percentual_(exibido, cru) {
  const e = String(exibido == null ? '' : exibido).trim();
  if (!e) return '';
  if (e.indexOf('%') > -1) return e;                 // já vem formatado
  const n = Number(String(cru).replace(',', '.'));
  if (!isNaN(n) && String(cru) !== '') {
    return (n > 0 && n <= 1 ? Math.round(n * 100) : Math.round(n)) + '%';
  }
  return e;
}

function soNumero_(t) {
  const d = String(t == null ? '' : t).replace(/\D/g, '');
  return d ? String(Number(d)) : '';
}

function chaveTexto_(t) {
  return String(t == null ? '' : t)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

/** Converte para 'yyyy-MM-dd' seja Date, seja texto em vários formatos. */
function dataISO_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);   // dd/mm/aaaa
  if (m) {
    let ano = m[3].length === 2 ? '20' + m[3] : m[3];
    return ano + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/** Lê a planilha do RDO inteira, com cache curto. */
function lerRDO_() {
  const cache = CacheService.getScriptCache();
  const guardado = cache.get('rdo_dados');
  if (guardado) {
    try { return JSON.parse(guardado); } catch (e) { /* cache ruim: relê */ }
  }

  const aba = abaRDO_();
  const faixa = aba.getDataRange();
  const dados = faixa.getValues();
  if (dados.length < 2) return [];

  /* Lê também o que a planilha MOSTRA, não só o valor cru.
     Motivo concreto: uma turbina digitada como "5/5" o Sheets converte em data,
     e o valor cru volta como "Tue May 05 2026 00:00:00 GMT-0300", que era o que
     aparecia no card. O texto exibido é o que a pessoa vê na planilha. */
  const exibido = faixa.getDisplayValues();
  const txt = function (l, c) {
    if (c === -1) return '';
    const v = exibido[l][c];
    return String(v == null ? '' : v).trim();
  };

  const mapa = detectarColunas_(dados[0]);
  const linhas = [];
  for (let l = 1; l < dados.length; l++) {
    const linha = dados[l];
    // A data segue pelo valor cru, que é Date de verdade; o resto pelo exibido.
    const data = mapa.data > -1 ? (dataISO_(linha[mapa.data]) || dataISO_(txt(l, mapa.data))) : '';
    const celulaParque = txt(l, mapa.parque);
    if (!data && !celulaParque) continue;

    // A célula inteira é a chave, de propósito. O número no fim ("SANTO
    // AGOSTINHO 1") é a equipe, mas parque e equipe são tratados juntos como um
    // rótulo só: separar exigia adivinhar onde termina o nome, o que dava
    // resultado inconsistente ("OITIS 1" ficava inteiro, "SÃO FERNANDO 1" era
    // dividido). Mantendo junto, os dois lados usam exatamente o mesmo texto e
    // cada dupla parque+equipe vira uma linha própria na cobrança — que é o
    // comportamento desejado.
    const parque = celulaParque.replace(/\s+/g, ' ');

    linhas.push({
      data: data,
      parque: parque,
      cliente: txt(l, mapa.cliente),
      link:    txt(l, mapa.link),
      tipo:    txt(l, mapa.tipo),
      autor:   txt(l, mapa.autor),
      equipes: txt(l, mapa.equipes),
      turbina: txt(l, mapa.turbina),
      blade:   txt(l, mapa.blade),
      local:   txt(l, mapa.local),
      avanco:  percentual_(txt(l, mapa.avanco), linha[mapa.avanco]),
      fim:     txt(l, mapa.fim)
    });
  }

  try {
    cache.put('rdo_dados', JSON.stringify(linhas), CACHE_RDO_SEGUNDOS);
  } catch (e) {
    // passou do limite do cache: segue sem, só fica mais lento
  }
  return linhas;
}

function ehFimDeSemana_(iso) {
  const p = String(iso).split('-');
  if (p.length !== 3) return false;
  const d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  return d.getDay() === 0 || d.getDay() === 6;
}


/* ===========================================================================
   AÇÕES DO STATUS RDO
   =========================================================================== */

/**
 * { data, cliente, parque } -> status do dia + relatórios que batem com o filtro.
 *
 * A cobrança sai dos projetos EM ANDAMENTO: cada um deveria ter um RDO na data.
 * Sábado e domingo entram como "não obrigatório" — o relatório pode existir,
 * mas a falta não é cobrada.
 */
function acaoRdoStatus_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  let linhas;
  try {
    linhas = lerRDO_();
  } catch (e) {
    return { ok: false, motivo: 'RDO_INDISPONIVEL', detalhe: e.message };
  }

  const data = dataISO_(p.data) || dataISO_(new Date());
  const fds = ehFimDeSemana_(data);

  // ---- cobrança: projetos em andamento x relatórios do dia
  const doDia = linhas.filter(function (r) { return r.data === data; });

  // O casamento é pelo nome do parque, exatamente como está na planilha do RDO.
  // Por isso o projeto escolhe o parque de uma lista tirada da própria planilha,
  // em vez de digitar: assim os dois lados usam a mesma grafia.
  const porParque = {};
  doDia.forEach(function (r) {
    if (!porParque[chaveTexto_(r.parque)]) porParque[chaveTexto_(r.parque)] = r;
  });

  // Matrícula do RDO vira nome, usando a MINI MASTER.
  let nomePorMatricula = {};
  try {
    lerTecnicos_().forEach(function (t) { nomePorMatricula[soNumero_(t.matricula)] = t.nome; });
  } catch (e) { /* sem MINI MASTER, mostra a matrícula mesmo */ }
  const quemFez = function (r) {
    if (!r) return '';
    const n = nomePorMatricula[soNumero_(r.autor)];
    return n || r.autor || '';
  };

  const projetos = lerProjetos_()
    .filter(function (pr) { return String(pr.status || 'andamento') === 'andamento'; })
    .map(function (pr) {
      const achado = porParque[chaveTexto_(pr.parque)] || null;
      return {
        id: pr.id, codigo: pr.codigo, parque: pr.parque, cliente: pr.cliente,
        tipoReparo: pr.tipoReparo,
        tecnicos: pr.tecnicos || [],
        estado: achado ? 'ENVIADO' : (fds ? 'NAO_OBRIGATORIO' : 'FALTA'),
        link: achado ? achado.link : '',
        autor: quemFez(achado),
        avanco: achado ? achado.avanco : '',
        finalizado: achado ? achado.fim : ''
      };
    });

  // ---- lista filtrada (para consultar dias passados)
  const fCliente = chaveTexto_(p.cliente || '');
  const fParque  = chaveTexto_(p.parque || '');
  const relatorios = linhas.filter(function (r) {
    if (p.data && r.data !== data) return false;
    if (fCliente && chaveTexto_(r.cliente) !== fCliente) return false;
    if (fParque && chaveTexto_(r.parque) !== fParque) return false;
    return true;
  }).slice(0, 400);

  relatorios.forEach(function (r) { r.autorNome = quemFez(r); });
  relatorios.sort(function (a, b) { return a.parque.localeCompare(b.parque); });

  return {
    ok: true,
    data: data,
    fimDeSemana: fds,
    projetos: projetos,
    relatorios: relatorios,
    total: relatorios.length
  };
}


/** -> listas para os menus de filtro (clientes e parques que existem no RDO) */
function acaoRdoFiltros_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  let linhas;
  try {
    linhas = lerRDO_();
  } catch (e) {
    return { ok: false, motivo: 'RDO_INDISPONIVEL', detalhe: e.message };
  }

  const clientes = {}, parques = {};
  linhas.forEach(function (r) {
    if (r.cliente) clientes[r.cliente] = true;
    if (r.parque) parques[r.parque] = true;
  });

  return {
    ok: true,
    clientes: Object.keys(clientes).sort(),
    parques: Object.keys(parques).sort()
  };
}


/* ===========================================================================
   TÉCNICOS (planilha MINI MASTER)
   Coluna A = matrícula, coluna B = nome. Só leitura.
   =========================================================================== */

function acaoTecnicosLista_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };
  try {
    return { ok: true, tecnicos: lerTecnicos_() };
  } catch (e) {
    return { ok: false, motivo: 'MINIMASTER_INDISPONIVEL', detalhe: e.message };
  }
}

function abaMiniMaster_() {
  const ss = SpreadsheetApp.openById(idPlanilhaMiniMaster_());
  const nome = abaNomeMiniMaster_();
  if (!nome) return ss.getSheets()[0];          // sem aba definida: a primeira
  const aba = ss.getSheetByName(nome);
  if (!aba) throw new Error('A aba "' + nome + '" não existe na MINI MASTER.');
  return aba;
}

function lerTecnicos_() {
  const cache = CacheService.getScriptCache();
  const guardado = cache.get('minimaster_tecnicos');
  if (guardado) {
    try { return JSON.parse(guardado); } catch (e) { /* cache ruim: relê */ }
  }

  const aba = abaMiniMaster_();
  const ultima = aba.getLastRow();
  if (ultima < 2) return [];

  const dados = aba.getRange(1, 1, ultima, 2).getValues();
  const vistos = {};
  const lista = [];

  dados.forEach(function (linha, i) {
    const mat = String(linha[0] == null ? '' : linha[0]).replace(/^'/, '').trim();
    const nome = String(linha[1] == null ? '' : linha[1]).trim();
    if (!nome) return;
    // pula o cabeçalho, se a primeira linha for rótulo em vez de dado
    if (i === 0 && !/\d/.test(mat)) return;
    const chave = chaveTexto_(nome) + '|' + mat;
    if (vistos[chave]) return;
    vistos[chave] = true;
    lista.push({ matricula: mat, nome: nome });
  });

  lista.sort(function (a, b) { return a.nome.localeCompare(b.nome); });
  try {
    cache.put('minimaster_tecnicos', JSON.stringify(lista), 600);
  } catch (e) { /* lista grande demais para o cache: segue sem */ }
  return lista;
}


/* ===========================================================================
   TIPOS DE REPARO (planilha BANCO DE INPUTS, aba ATV POR HR)
   Só leitura. Entram na lista as linhas em que "Atividade obrigatória" está
   como "Não" — o resto é rotina fixa do dia, não define projeto.
   =========================================================================== */

function acaoTiposReparoLista_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };
  try {
    return { ok: true, tipos: lerTiposReparo_() };
  } catch (e) {
    return { ok: false, motivo: 'INPUTS_INDISPONIVEL', detalhe: e.message };
  }
}

function abaInputs_() {
  const ss = SpreadsheetApp.openById(idPlanilhaInputs_());
  const nome = abaNomeInputs_();
  const aba = ss.getSheetByName(nome);
  if (!aba) throw new Error('A aba "' + nome + '" não existe na planilha Banco de inputs.');
  return aba;
}

/* "Não", "NÃO", "nao", "N" — tudo isso é não. Célula vazia NÃO conta como não:
   linha sem preencher costuma ser linha em branco no fim da tabela, não uma
   decisão de quem montou a planilha. */
function ehNao_(valor) {
  const v = String(valor == null ? '' : valor)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().trim();
  return v === 'nao' || v === 'n' || v === 'false' || v === 'no';
}

function lerTiposReparo_() {
  const cache = CacheService.getScriptCache();
  const guardado = cache.get('inputs_tipos_reparo');
  if (guardado) {
    try { return JSON.parse(guardado); } catch (e) { /* cache ruim: relê */ }
  }

  const aba = abaInputs_();
  const ultima = aba.getLastRow();
  const larg = aba.getLastColumn();
  if (ultima < 2 || larg < 1) return [];

  const dados = aba.getRange(1, 1, ultima, larg).getValues();
  const norm = dados[0].map(normalizarCab_);

  const acharCol = function (apelidos) {
    for (let i = 0; i < apelidos.length; i++) {
      const idx = norm.indexOf(apelidos[i]);
      if (idx > -1) return idx;
    }
    for (let i = 0; i < apelidos.length; i++) {
      const idx = norm.findIndex(function (c) { return c && c.indexOf(apelidos[i]) > -1; });
      if (idx > -1) return idx;
    }
    return -1;
  };

  const cTipo = acharCol(COLUNAS_INPUTS.tipo);
  const cObr  = acharCol(COLUNAS_INPUTS.obrigatoria);
  if (cTipo === -1) {
    throw new Error('Não achei a coluna "Tipo de reparo" na aba "' + abaNomeInputs_() +
                    '". Cabeçalho lido: ' + dados[0].filter(String).join(' | '));
  }
  if (cObr === -1) {
    throw new Error('Não achei a coluna "Atividade obrigatória" na aba "' + abaNomeInputs_() +
                    '". Cabeçalho lido: ' + dados[0].filter(String).join(' | '));
  }

  const vistos = {}, lista = [];
  for (let l = 1; l < dados.length; l++) {
    if (!ehNao_(dados[l][cObr])) continue;
    const nome = String(dados[l][cTipo] == null ? '' : dados[l][cTipo]).trim();
    if (!nome) continue;
    const chave = chaveTexto_(nome);
    if (vistos[chave]) continue;
    vistos[chave] = true;
    lista.push(nome);
  }

  lista.sort(function (a, b) { return a.localeCompare(b); });
  try {
    cache.put('inputs_tipos_reparo', JSON.stringify(lista), 600);
  } catch (e) { /* lista grande demais para o cache: segue sem */ }
  return lista;
}


/* ===========================================================================
   PROJETOS EM ANDAMENTO
   =========================================================================== */

function lerProjetos_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_PROJ_CARDS);
  const dados = aba.getDataRange().getValues();
  const lista = [];
  for (let l = 1; l < dados.length; l++) {
    if (!dados[l][0]) continue;
    try {
      const pr = JSON.parse(dados[l][4]);
      pr._rev = Number(dados[l][3]) || 0;
      lista.push(pr);
    } catch (e) { /* linha corrompida: ignora */ }
  }
  return lista;
}

function metaProjetos_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_PROJ_META);
  const dados = aba.getDataRange().getValues();
  const meta = { rev: 1, nextSeq: 1 };
  for (let l = 1; l < dados.length; l++) {
    if (String(dados[l][0]) === 'rev') meta.rev = Number(dados[l][1]) || 1;
    if (String(dados[l][0]) === 'nextSeq') meta.nextSeq = Number(dados[l][1]) || 1;
  }
  return meta;
}

function bumparMetaProjetos_(nextSeq) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_PROJ_META);
  const dados = aba.getDataRange().getValues();
  const linhaDe = {};
  for (let l = 1; l < dados.length; l++) linhaDe[String(dados[l][0])] = l + 1;

  const rev = (Number(dados[(linhaDe['rev'] || 2) - 1][1]) || 1) + 1;
  aba.getRange(linhaDe['rev'] || 2, 2).setValue(rev);
  if (nextSeq && linhaDe['nextSeq']) aba.getRange(linhaDe['nextSeq'], 2).setValue(Number(nextSeq));
  return rev;
}

/** { rev } -> { ok, rev, mudou, projetos?, podeEditar } */
function acaoProjetosCarregar_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  const perfil = perfilDe_(matricula);
  const meta = metaProjetos_();
  if (Number(p.rev) === meta.rev) {
    return { ok: true, rev: meta.rev, mudou: false, podeEditar: podeEditarProjeto_(perfil) };
  }
  return {
    ok: true, rev: meta.rev, mudou: true,
    nextSeq: meta.nextSeq,
    projetos: lerProjetos_(),
    podeEditar: podeEditarProjeto_(perfil),
    perfil: perfil
  };
}

function perfilDe_(matricula) {
  const achado = acharPorMatricula_(abaUsuarios_(), matricula);
  return achado ? String(achado.dados.perfil).toUpperCase() : 'DIRETORIA';
}

/* Quem cria e edita projeto é o SUPERVISOR, que é quem está em campo.
   ADMIN e DIRETORIA acompanham, mas não alteram. */
function podeEditarProjeto_(perfil) {
  return perfil === 'SUPERVISOR';
}

/** { projetos[], removidos[], nextSeq } -> { ok, rev, aceitos[], conflitos[] } */
function acaoProjetosSalvar_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  // A tela já esconde os botões, mas quem manda é aqui: sem esta checagem,
  // qualquer pessoa logada poderia gravar por fora da interface.
  if (!podeEditarProjeto_(perfilDe_(matricula))) {
    registrar_(matricula, 'PROJ_SALVAR', 'NEGADO', 'perfil sem permissão');
    return { ok: false, motivo: 'SEM_PERMISSAO' };
  }

  const trava = LockService.getScriptLock();
  if (!trava.tryLock(20000)) return { ok: false, motivo: 'OCUPADO' };

  try {
    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_PROJ_CARDS);
    const dados = aba.getDataRange().getValues();
    const mapa = {};
    for (let l = 1; l < dados.length; l++) {
      if (dados[l][0]) mapa[String(dados[l][0])] = { linha: l + 1, rev: Number(dados[l][3]) || 0, json: dados[l][4] };
    }

    const agora = new Date();
    const aceitos = [], conflitos = [], novas = [];

    (p.projetos || []).forEach(function (pr) {
      const id = String(pr.id);
      const revEnviada = Number(pr._rev) || 0;
      delete pr._rev;
      const atual = mapa[id];

      if (!atual) {
        novas.push([id, pr.codigo || '', pr.status || 'andamento', 1,
                    JSON.stringify(pr), agora, matricula]);
        aceitos.push({ id: id, rev: 1 });
        return;
      }
      if (atual.rev !== revEnviada) {
        let servidor = null;
        try { servidor = JSON.parse(atual.json); } catch (e) { servidor = null; }
        if (servidor) {
          servidor._rev = atual.rev;
          conflitos.push({ id: id, rev: atual.rev, projeto: servidor });
          return;
        }
      }
      const novaRev = atual.rev + 1;
      aba.getRange(atual.linha, 1, 1, COLUNAS_PROJ.length).setValues([[
        id, pr.codigo || '', pr.status || 'andamento', novaRev,
        JSON.stringify(pr), agora, matricula
      ]]);
      aceitos.push({ id: id, rev: novaRev });
    });

    if (novas.length) {
      aba.getRange(aba.getLastRow() + 1, 1, novas.length, COLUNAS_PROJ.length).setValues(novas);
    }

    const remover = (p.removidos || [])
      .map(function (id) { return mapa[String(id)] ? mapa[String(id)].linha : null; })
      .filter(function (l) { return l; })
      .sort(function (a, b) { return b - a; });
    remover.forEach(function (linha) { aba.deleteRow(linha); });

    const rev = bumparMetaProjetos_(p.nextSeq);
    if (conflitos.length || remover.length) {
      registrar_(matricula, 'PROJ_SALVAR', conflitos.length ? 'CONFLITO' : 'OK',
        aceitos.length + ' aceitos, ' + conflitos.length + ' conflitos, ' + remover.length + ' removidos');
    }
    return { ok: true, rev: rev, aceitos: aceitos, conflitos: conflitos };

  } finally {
    trava.releaseLock();
  }
}
