/* ============================================================================
   SINCRONIA COM A ABA SUPERVISORES (planilha Banco de inputs)

   Antes, a aba era um ESPELHO: o portal apagava da linha 2 para baixo e
   regravava tudo a cada gravação de projeto. Quem digitasse lá perdia o que
   escreveu na próxima gravação — e, com o gatilho diário, perdia de madrugada
   mesmo sem ninguém salvar nada.

   Agora os dois lados escrevem. O mecanismo é comparação com RETRATO:

     - toda vez que o portal grava uma célula na aba, ele guarda o valor que
       gravou, num retrato por projeto (chave 'espelhoInputs' na aba PROJ_META);
     - na próxima passada, uma célula diferente do retrato só pode ter sido
       mexida por uma pessoa na planilha. Essa vira a alteração dele e entra no
       card;
     - célula igual ao retrato ninguém tocou, e o portal pode regravar à
       vontade.

   É o que impede os dois lados de se apagarem: o portal nunca sobrescreve uma
   célula que não reconhece como sua.

   LIMITES, de propósito:

   - Apagar a linha dele NÃO apaga o projeto. A linha volta na passada
     seguinte. Excluir projeto continua sendo só pelo portal — uma tecla a mais
     na planilha não pode sumir com um projeto inteiro.
   - Linha nova vira projeto com o que a linha tem, e só. Data de início,
     previsão de término e níveis dos técnicos não existem na aba; o card nasce
     sem eles e aparece como "Sem data" até alguém completar no portal.
   - Mexer na célula de matrícula registra uma TROCA DE EQUIPE a partir de
     hoje. O histórico anterior fica intacto; os níveis (N1/N2/N3) da formação
     nova ficam em branco, porque a aba não tem onde guardá-los.

   SITUAÇÃO E DATA DE FINALIZAÇÃO (out/2026)

   - DATA DE FINALIZAÇÃO é a data real de término (pr.fim). Preenchida, o
     projeto é Concluído — no portal ou na planilha, tanto faz quem preencheu.
     Apagar a data na planilha reabre o projeto.
   - SITUAÇÃO é calculada pelo portal a cada passada, a partir das datas:
       finalização preenchida -> Concluído
       início depois de hoje  -> A iniciar
       previsão já passou     -> Em atraso
       o resto                -> Em andamento
     O que alguém digitar nela é ignorado e sobrescrito na passada seguinte:
     ela é resultado, não entrada. Quem conclui é a data de finalização.
   - Projeto concluído CONTINUA na aba por DIAS_CONCLUIDO_NA_ABA dias depois da
     finalização (antes saía na hora). O app do RDO lê esta aba e precisa da
     linha com a data de finalização: no próprio dia em que o projeto fecha a
     equipe ainda manda o RDO e o report de EHS.
   ============================================================================ */

const CHAVE_RETRATO = 'espelhoInputs';
const COL_ID_INPUTS = ['id', 'id_projeto', 'idprojeto', 'id_do_projeto'];

/* Colunas novas desta versão. Ficam aqui, e não em COLUNAS_SUP_INPUTS
   (SupervisaoCampo.gs), para a mudança caber num arquivo só.
   Apelidos específicos de propósito: acharColuna_ cai para busca por pedaço
   quando não acha exato, e 'status' solto casaria com outra coisa. */
const COLUNAS_SUP_EXTRAS = {
  finalizacao: ['data_de_finalizacao', 'data_finalizacao', 'finalizacao',
                'data_de_conclusao', 'data_conclusao', 'data_real_de_finalizacao'],
  situacao:    ['situacao', 'situacao_do_projeto', 'status_do_projeto']
};

/* Campos de data desta aba (os de CAMPOS_DATA_INPUTS, no SupervisaoCampo.gs,
   mais a finalização). Escrito por extenso de propósito: montar a partir da
   constante do outro arquivo depende da ordem em que o Apps Script carrega os
   arquivos, e se ela mudar o projeto inteiro para de carregar. */
const CAMPOS_DATA_SYNC = ['inicio', 'fimPrevisto', 'finalizacao'];

/* Quanto tempo o projeto concluído fica na aba depois da data de finalização. */
const DIAS_CONCLUIDO_NA_ABA = 30;

/* Campos que o portal escreve na aba. Qualquer outra coluna é dele e nunca é
   tocada — nem lida, nem limpa. */
const CAMPOS_PORTAL = ['supervisor', 'cliente', 'parque', 'tipo', 'matricula',
                       'turbina', 'blade', 'inicio', 'fimPrevisto',
                       'finalizacao', 'situacao'];

/* Cabeçalho que o portal escreve quando a coluna ainda não existe na aba.
   Colunas novas entram no FIM, nunca no meio: inserir no meio empurraria as
   colunas dele e quebraria qualquer fórmula ou filtro que ele já tenha.
   Os nomes de FINALIZAÇÃO e SITUAÇÃO são os que o app do RDO procura. */
const CABECALHO_PADRAO = {
  inicio:      'INÍCIO',
  fimPrevisto: 'PREVISÃO DE TÉRMINO',
  finalizacao: 'DATA DE FINALIZAÇÃO',
  situacao:    'SITUAÇÃO'
};


/* ---------------------------------------------------------------------------
   SITUAÇÃO (mesma regra da tela de projetos, situacaoDe no HTML)
   --------------------------------------------------------------------------- */

function hojeISOServidor_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function rotuloSituacao_(pr) {
  const hoje = hojeISOServidor_();
  const inicio = dataISO_(pr.inicio || '');
  const prev = dataISO_(pr.fimPrevisto || '');
  if (pr.fim || String(pr.status || 'andamento') === 'concluido') return 'Concluído';
  if (inicio && inicio > hoje) return 'A iniciar';
  if (prev && prev < hoje) return 'Em atraso';
  return 'Em andamento';
}

/* Fica na aba: todo projeto não concluído, e o concluído até
   DIAS_CONCLUIDO_NA_ABA dias depois da finalização. Concluído antigo, sem
   data (de antes desta versão), não tem como ser medido e sai. */
function ficaNaAba_(pr) {
  if (String(pr.status || 'andamento') === 'andamento') return true;
  const fim = dataISO_(pr.fim || '');
  if (!fim) return false;
  const corte = Utilities.formatDate(
    new Date(Date.now() - DIAS_CONCLUIDO_NA_ABA * 24 * 3600 * 1000),
    Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return fim >= corte;
}


/* ---------------------------------------------------------------------------
   RETRATO
   --------------------------------------------------------------------------- */

function lerRetrato_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_PROJ_META);
  const dados = aba.getDataRange().getValues();
  for (let l = 1; l < dados.length; l++) {
    if (String(dados[l][0]) === CHAVE_RETRATO) {
      try { return JSON.parse(dados[l][1]) || {}; } catch (e) { return {}; }
    }
  }
  return {};
}

function gravarRetrato_(retrato) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_PROJ_META);
  const dados = aba.getDataRange().getValues();
  const texto = JSON.stringify(retrato);

  // Uma célula do Sheets aceita 50 000 caracteres. Com ~20 projetos isso dá
  // uns 4 KB, mas se um dia estourar é melhor ficar sem retrato (e voltar a
  // ser espelho) do que derrubar a gravação do projeto.
  if (texto.length > 45000) return false;

  for (let l = 1; l < dados.length; l++) {
    if (String(dados[l][0]) === CHAVE_RETRATO) {
      aba.getRange(l + 1, 2).setValue(texto);
      return true;
    }
  }
  aba.getRange(aba.getLastRow() + 1, 1, 1, 2).setValues([[CHAVE_RETRATO, texto]]);
  return true;
}


/* ---------------------------------------------------------------------------
   ABA: cabeçalho, coluna de id
   --------------------------------------------------------------------------- */

/**
 * Localiza as colunas da aba e garante que exista a coluna de id.
 * Sem o id não há como casar linha com card: casar por parque quebraria na
 * primeira vez que ele corrigisse o nome do parque — viraria projeto novo em
 * vez de correção.
 */
function mapearInputs_() {
  const aba = abaSupervisoresInputs_();
  let larg = Math.max(aba.getLastColumn(), 7);
  let cab = aba.getRange(1, 1, 1, larg).getValues()[0];
  let norm = cab.map(normalizarCab_);

  // Acrescenta uma coluna no fim da aba e devolve o índice dela.
  const acrescentar = function (titulo, esconder) {
    larg = larg + 1;
    aba.getRange(1, larg).setValue(titulo);
    if (esconder) aba.hideColumns(larg);
    cab = aba.getRange(1, 1, 1, larg).getValues()[0];
    norm = cab.map(normalizarCab_);
    return larg - 1;
  };

  let iId = acharColuna_(norm, COL_ID_INPUTS);
  if (iId === -1) iId = acrescentar('id', true);   // controle, ninguém precisa ver

  const col = { id: iId };
  Object.keys(COLUNAS_SUP_INPUTS).forEach(function (campo) {
    col[campo] = acharColuna_(norm, COLUNAS_SUP_INPUTS[campo]);
  });
  Object.keys(COLUNAS_SUP_EXTRAS).forEach(function (campo) {
    col[campo] = acharColuna_(norm, COLUNAS_SUP_EXTRAS[campo]);
  });

  // Início, previsão, finalização e situação o portal cria se não existirem.
  Object.keys(CABECALHO_PADRAO).forEach(function (campo) {
    if (col[campo] === -1) col[campo] = acrescentar(CABECALHO_PADRAO[campo], false);
  });

  const faltando = CAMPOS_PORTAL.filter(function (c) {
    return col[c] === -1 && COLUNAS_SUP_OPCIONAIS.indexOf(c) === -1;
  });
  if (faltando.length) {
    throw new Error('Colunas não encontradas na aba "' + ABA_SUP_INPUTS + '": ' +
                    faltando.join(', ') + '. Cabeçalho lido: ' + cab.filter(String).join(' | '));
  }

  return { aba: aba, larg: larg, col: col };
}


/* ---------------------------------------------------------------------------
   CARD -> CÉLULAS
   Exatamente os mesmos valores que o espelho antigo gravava. É isso que vira
   o retrato, então tem que ser uma função só, usada pelos dois lados.
   --------------------------------------------------------------------------- */

function celulasDoProjeto_(pr) {
  const mats = equipeVigenteDoProjeto_(pr)
    .map(function (t) { return String(t.matricula || '').trim(); })
    .filter(String);

  const turbs = pr.semTurbina ? [] : (pr.turbinas || []);

  let blade;
  if (pr.semTurbina) {
    blade = (pr.bladesSoltas || []).map(function (b) { return String(b || '').trim(); })
      .filter(String).join(', ');
  } else {
    blade = turbs.map(function (t) {
      const bl = (t.blades || []).map(function (b) { return String(b || '').trim(); }).filter(String);
      return bl.length ? t.nome + ': ' + bl.join(', ') : '';
    }).filter(String).join(' | ');
  }

  return {
    supervisor:  (pr.supervisor && pr.supervisor.nome) ? pr.supervisor.nome : '',
    cliente:     pr.cliente || '',
    parque:      pr.parque || '',
    tipo:        pr.tipoReparo || '',
    matricula:   mats.join(', '),
    turbina:     pr.semTurbina ? 'Sem turbina' : turbs.map(function (t) { return t.nome; }).filter(String).join(', '),
    blade:       blade,
    // Sempre em ISO aqui. O retrato é comparado com o que vem da célula, e a
    // célula pode voltar como Date ou como texto dd/mm/aaaa dependendo de como
    // a pessoa digitou — a conversão acontece em lerCampoDaLinha_.
    inicio:      dataISO_(pr.inicio || ''),
    fimPrevisto: dataISO_(pr.fimPrevisto || ''),
    finalizacao: dataISO_(pr.fim || ''),
    situacao:    rotuloSituacao_(pr)
  };
}

/* Como o valor de um campo sai da célula. Data vira ISO; o resto vira texto. */
function lerCampoDaLinha_(campo, valor) {
  return CAMPOS_DATA_SYNC.indexOf(campo) > -1 ? dataISO_(valor) : textoCelula_(valor);
}

/* Como o valor de um campo entra na célula. Data entra como Date de verdade,
   para a pessoa poder usar o seletor do Sheets e para ordenar certo. */
function valorParaCelula_(campo, texto) {
  if (CAMPOS_DATA_SYNC.indexOf(campo) > -1) {
    if (!texto) return '';
    const q = String(texto).split('-');
    return new Date(Number(q[0]), Number(q[1]) - 1, Number(q[2]));
  }
  // Apóstrofo só na matrícula, para o Sheets não comer zero à esquerda.
  return (campo === 'matricula' && texto) ? "'" + texto : texto;
}

/* A célula de matrícula é gravada com apóstrofo na frente para o Sheets não
   comer o zero à esquerda. Na leitura ele não aparece, mas quando a comparação
   é feita contra o que ESTÁ na célula vale normalizar dos dois lados. */
function textoCelula_(v) {
  return String(v == null ? '' : v).replace(/^'/, '').replace(/\s+/g, ' ').trim();
}


/* ---------------------------------------------------------------------------
   CÉLULAS -> CARD
   --------------------------------------------------------------------------- */

/** Nome do técnico a partir da matrícula, pela MINI MASTER. */
function nomePorMatricula_() {
  const mapa = {};
  try {
    lerTecnicos_().forEach(function (t) { mapa[soNumero_(t.matricula)] = t.nome; });
  } catch (e) { /* sem MINI MASTER, fica só a matrícula */ }
  return mapa;
}

/** "T01: A, B | T02: C" -> [{nome:'T01', blades:['A','B']}, ...] */
function lerTurbinasDaCelula_(celTurbina, celBlade) {
  const nomes = textoCelula_(celTurbina).split(',')
    .map(function (s) { return s.trim(); }).filter(String);

  const bladesPor = {};
  textoCelula_(celBlade).split('|').forEach(function (parte) {
    const i = parte.indexOf(':');
    if (i === -1) return;
    const nome = parte.slice(0, i).trim();
    bladesPor[nome] = parte.slice(i + 1).split(',')
      .map(function (b) { return b.trim(); }).filter(String);
  });

  return nomes.map(function (n) { return { nome: n, blades: bladesPor[n] || [] }; });
}

/**
 * Aplica no card uma célula que ele mexeu. Devolve true se mudou algo.
 * Só é chamada para célula que difere do retrato — ou seja, para alteração
 * dele mesmo, nunca para valor que o próprio portal escreveu.
 */
function aplicarCelulaNoProjeto_(pr, campo, valor, nomes) {
  const v = textoCelula_(valor);

  if (campo === 'inicio')      { pr.inicio = dataISO_(valor); return true; }
  if (campo === 'fimPrevisto') { pr.fimPrevisto = dataISO_(valor); return true; }

  // Data de finalização decide o status: preenchida conclui, apagada reabre.
  if (campo === 'finalizacao') {
    pr.fim = dataISO_(valor);
    pr.status = pr.fim ? 'concluido' : 'andamento';
    return true;
  }
  // Situação é calculada; edição manual é ignorada e regravada na passada.
  if (campo === 'situacao') return false;

  if (campo === 'cliente') { pr.cliente = v; return true; }
  if (campo === 'parque')  { pr.parque = v;  return true; }
  if (campo === 'tipo')    { pr.tipoReparo = v; return true; }

  if (campo === 'supervisor') {
    // A matrícula do supervisor fica em branco quando o nome não bate com a
    // aba USUARIOS: melhor ficar sem do que gravar uma matrícula errada.
    let mat = '';
    try {
      const aba = abaUsuarios_();
      const dados = aba.getDataRange().getValues();
      const cab = dados[0];
      const iN = cab.indexOf('nome'), iM = cab.indexOf('matricula');
      for (let l = 1; l < dados.length; l++) {
        if (String(dados[l][iN]).trim().toUpperCase() === v.toUpperCase()) {
          mat = String(dados[l][iM]).trim();
          break;
        }
      }
    } catch (e) { /* segue sem matrícula */ }
    pr.supervisor = v ? { nome: v, matricula: mat } : null;
    return true;
  }

  if (campo === 'matricula') {
    const mats = v.split(',').map(function (s) { return s.trim(); }).filter(String);
    const tecnicos = mats.map(function (m) {
      return { nome: nomes[soNumero_(m)] || m, matricula: m, nivel: '' };
    });
    const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // Troca de equipe a partir de hoje. O histórico anterior fica — foi o que
    // realmente aconteceu em campo e não pode ser reescrito por uma edição de
    // planilha feita hoje.
    if (!Array.isArray(pr.equipes)) {
      pr.equipes = [{ desde: pr.inicio || hoje, tecnicos: pr.tecnicos || [] }];
    }
    pr.equipes = pr.equipes.filter(function (e) { return e.desde !== hoje; });
    pr.equipes.push({ desde: hoje, tecnicos: tecnicos });
    pr.equipes.sort(function (a, b) { return String(a.desde).localeCompare(String(b.desde)); });
    pr.tecnicos = equipeVigenteDoProjeto_(pr);
    return true;
  }

  if (campo === 'turbina' || campo === 'blade') {
    return false;   // tratados juntos, fora daqui
  }
  return false;
}


/* ---------------------------------------------------------------------------
   PASSADA DE SINCRONIA
   --------------------------------------------------------------------------- */

/**
 * Lê a aba, traz para os cards o que ele mexeu, cria card para linha nova e
 * regrava na aba só as células que o portal controla.
 * Devolve { mudou, novos, alterados, linhas }.
 */
function sincronizarInputs_(matricula) {
  const mapa = mapearInputs_();
  const aba = mapa.aba, col = mapa.col, larg = mapa.larg;

  const ultima = aba.getLastRow();
  const grade = ultima > 1 ? aba.getRange(2, 1, ultima - 1, larg).getValues() : [];

  const retrato = lerRetrato_();
  const nomes = nomePorMatricula_();

  // ---- cards de hoje
  const abaCards = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_PROJ_CARDS);
  const linhasCards = abaCards.getDataRange().getValues();
  const porId = {};
  for (let l = 1; l < linhasCards.length; l++) {
    if (!linhasCards[l][0]) continue;
    try {
      porId[String(linhasCards[l][0])] = {
        linha: l + 1,
        rev: Number(linhasCards[l][3]) || 0,
        pr: JSON.parse(linhasCards[l][4])
      };
    } catch (e) { /* linha corrompida: ignora */ }
  }

  const meta = metaProjetos_();
  let proximaSeq = meta.nextSeq;
  const alterados = {};     // id -> true
  const novos = [];         // linhas a gravar em PROJ_CARDS
  const idsEscrever = [];   // [linhaNaAba, id] para carimbar o id da linha nova

  // ---- o que ele mexeu
  grade.forEach(function (linha, i) {
    const id = String(linha[col.id] || '').trim();

    // Linha nova: só vira projeto se tiver parque. Linha com uma célula solta
    // costuma ser rascunho, e virar card vazio só atrapalha.
    if (!id) {
      const parque = textoCelula_(linha[col.parque]);
      if (!parque) return;

      const novoId = Utilities.getUuid();
      const fimReal = col.finalizacao > -1 ? dataISO_(linha[col.finalizacao]) : '';
      const pr = {
        id: novoId,
        codigo: 'PRJ-' + ('000' + proximaSeq).slice(-3) + '-' + new Date().getFullYear(),
        parque: parque,
        cliente: textoCelula_(linha[col.cliente]),
        tipoReparo: textoCelula_(linha[col.tipo]),
        inicio: col.inicio > -1 ? dataISO_(linha[col.inicio]) : '',
        fim: fimReal,
        fimPrevisto: col.fimPrevisto > -1 ? dataISO_(linha[col.fimPrevisto]) : '',
        status: fimReal ? 'concluido' : 'andamento',
        supervisor: null,
        semTurbina: false, turbinas: [], bladesSoltas: [],
        tecnicos: [], equipes: [],
        origem: 'planilha'
      };
      // Com as duas datas preenchidas na linha, o card já nasce completo e não
      // precisa da etiqueta de pendência.
      if (pr.inicio && pr.fimPrevisto) delete pr.origem;
      proximaSeq++;

      aplicarCelulaNoProjeto_(pr, 'supervisor', linha[col.supervisor], nomes);
      if (textoCelula_(linha[col.matricula])) {
        aplicarCelulaNoProjeto_(pr, 'matricula', linha[col.matricula], nomes);
        // Projeto que nasce aqui não teve troca nenhuma: a primeira formação
        // vale desde o início do projeto, não desde hoje.
        if (pr.inicio && pr.equipes.length === 1) pr.equipes[0].desde = pr.inicio;
      }
      aplicarTurbinasNoProjeto_(pr, linha, col);

      porId[novoId] = { linha: null, rev: 0, pr: pr };
      alterados[novoId] = true;
      novos.push(novoId);
      idsEscrever.push([i + 2, novoId]);
      return;
    }

    const alvo = porId[id];
    if (!alvo) return;        // projeto excluído: a linha some depois

    const antes = retrato[id] || {};
    let mexeuTurbina = false;

    CAMPOS_PORTAL.forEach(function (campo) {
      if (col[campo] === -1) return;
      const atual = lerCampoDaLinha_(campo, linha[col[campo]]);
      const gravado = lerCampoDaLinha_(campo, antes[campo]);
      if (atual === gravado) return;              // ninguém tocou

      if (campo === 'turbina' || campo === 'blade') { mexeuTurbina = true; return; }
      if (aplicarCelulaNoProjeto_(alvo.pr, campo, atual, nomes)) alterados[id] = true;
    });

    if (mexeuTurbina) {
      aplicarTurbinasNoProjeto_(alvo.pr, linha, col);
      alterados[id] = true;
    }
  });

  // ---- grava os cards que mudaram
  const agora = new Date();
  const autor = matricula || 'planilha';
  const paraCriar = [];

  Object.keys(alterados).forEach(function (id) {
    const alvo = porId[id];
    if (alvo.linha) {
      abaCards.getRange(alvo.linha, 1, 1, COLUNAS_PROJ.length).setValues([[
        id, alvo.pr.codigo || '', alvo.pr.status || 'andamento', alvo.rev + 1,
        JSON.stringify(alvo.pr), agora, autor
      ]]);
      alvo.rev = alvo.rev + 1;
    } else {
      paraCriar.push([id, alvo.pr.codigo || '', alvo.pr.status || 'andamento', 1,
                      JSON.stringify(alvo.pr), agora, autor]);
      alvo.rev = 1;
    }
  });
  if (paraCriar.length) {
    abaCards.getRange(abaCards.getLastRow() + 1, 1, paraCriar.length, COLUNAS_PROJ.length)
      .setValues(paraCriar);
  }

  // ---- carimba o id nas linhas novas
  idsEscrever.forEach(function (par) {
    aba.getRange(par[0], col.id + 1).setValue(par[1]);
  });

  // ---- regrava na aba só as células do portal, sem encostar nas dele
  const naAba = Object.keys(porId)
    .map(function (id) { return porId[id].pr; })
    .filter(ficaNaAba_)
    .sort(function (a, b) { return String(a.parque || '').localeCompare(String(b.parque || '')); });

  const novoRetrato = escreverNaAba_(aba, col, larg, naAba);
  gravarRetrato_(novoRetrato);

  const mudou = Object.keys(alterados).length > 0;
  if (mudou) bumparMetaProjetos_(proximaSeq);

  return {
    mudou: mudou,
    novos: novos.length,
    alterados: Object.keys(alterados).length - novos.length,
    linhas: naAba.length
  };
}

function aplicarTurbinasNoProjeto_(pr, linha, col) {
  const cT = col.turbina > -1 ? textoCelula_(linha[col.turbina]) : '';
  const cB = col.blade   > -1 ? textoCelula_(linha[col.blade])   : '';

  if (cT.toLowerCase() === 'sem turbina') {
    pr.semTurbina = true;
    pr.turbinas = [];
    pr.bladesSoltas = cB.split(',').map(function (b) { return b.trim(); }).filter(String);
    return;
  }
  pr.semTurbina = false;
  pr.bladesSoltas = [];
  pr.turbinas = lerTurbinasDaCelula_(cT, cB);
}

/**
 * Escreve na aba as células do portal, linha a linha, casando pelo id.
 * Nunca usa clearContent: coluna que o portal não conhece é dele e fica onde
 * está. Linha de projeto que saiu da aba (excluído, ou concluído há mais de
 * DIAS_CONCLUIDO_NA_ABA dias) tem só as células do portal limpas — a linha em
 * si continua, porque pode haver anotação dele nas outras colunas.
 * Devolve o retrato novo.
 */
function escreverNaAba_(aba, col, larg, projetos) {
  const ultima = aba.getLastRow();
  const grade = ultima > 1 ? aba.getRange(2, 1, ultima - 1, larg).getValues() : [];

  const linhaDe = {};
  grade.forEach(function (linha, i) {
    const id = String(linha[col.id] || '').trim();
    if (id) linhaDe[id] = i;           // índice dentro de grade
  });

  const retrato = {};
  const vistos = {};

  projetos.forEach(function (pr) {
    const cels = celulasDoProjeto_(pr);
    retrato[pr.id] = cels;
    vistos[pr.id] = true;

    let i = linhaDe[pr.id];
    if (i === undefined) {
      i = grade.length;
      grade.push(new Array(larg).fill(''));
      grade[i][col.id] = pr.id;
      linhaDe[pr.id] = i;
    }
    CAMPOS_PORTAL.forEach(function (campo) {
      if (col[campo] === -1) return;
      grade[i][col[campo]] = valorParaCelula_(campo, cels[campo]);
    });
  });

  // Projeto que não fica mais na aba: limpa as células do portal e apaga o
  // id, para a linha não voltar a ser casada com nada.
  grade.forEach(function (linha, i) {
    const id = String(linha[col.id] || '').trim();
    if (!id || vistos[id]) return;
    CAMPOS_PORTAL.forEach(function (campo) {
      if (col[campo] > -1) grade[i][col[campo]] = '';
    });
    grade[i][col.id] = '';
  });

  if (grade.length) aba.getRange(2, 1, grade.length, larg).setValues(grade);
  return retrato;
}


/* ---------------------------------------------------------------------------
   PONTOS DE ENTRADA
   --------------------------------------------------------------------------- */

/** Usada pelo salvamento de projeto, no lugar do espelho antigo. */
function sincronizarInputsSeguro_(matricula) {
  try {
    const r = sincronizarInputs_(matricula);
    return { ok: true, linhas: r.linhas };
  } catch (e) {
    registrar_(matricula || '', 'SYNC_INPUTS', 'FALHA', String(e.message));
    return { ok: false, detalhe: e.message };
  }
}

/**
 * Gatilho. Roda com trava porque pode cair junto com uma gravação do portal —
 * e as duas escrevem na mesma aba.
 */
function sincronizarInputsAgora() {
  const trava = LockService.getScriptLock();
  if (!trava.tryLock(20000)) return;
  try {
    const r = sincronizarInputs_('planilha');
    Logger.log('Sincronia: ' + r.novos + ' novos, ' + r.alterados + ' alterados, ' +
               r.linhas + ' linhas na aba.');
  } catch (e) {
    Logger.log('ERRO na sincronia: ' + e.message);
    registrar_('', 'SYNC_INPUTS', 'FALHA', String(e.message));
  } finally {
    trava.releaseLock();
  }
}

/** Gatilho de edição instalado na planilha Banco de inputs. */
function aoEditarInputs(e) {
  try {
    if (!e || !e.range) return;
    if (e.range.getSheet().getName() !== ABA_SUP_INPUTS) return;
    if (e.range.getRow() < 2) return;           // cabeçalho
  } catch (err) { return; }
  sincronizarInputsAgora();
}

/**
 * Roda UMA vez, à mão, para instalar os dois gatilhos na planilha de inputs.
 * Rodar de novo não duplica: os antigos são removidos antes.
 */
function instalarGatilhosInputs() {
  const id = idPlanilhaInputs_();

  ScriptApp.getProjectTriggers().forEach(function (g) {
    const f = g.getHandlerFunction();
    if (f === 'aoEditarInputs' || f === 'sincronizarInputsAgora') ScriptApp.deleteTrigger(g);
  });

  ScriptApp.newTrigger('aoEditarInputs').forSpreadsheet(id).onEdit().create();
  ScriptApp.newTrigger('sincronizarInputsAgora').timeBased().everyMinutes(10).create();

  avisar_('Gatilhos instalados.\n\n' +
          '• aoEditarInputs — dispara quando alguém edita a aba ' + ABA_SUP_INPUTS + '\n' +
          '• sincronizarInputsAgora — rede de segurança a cada 10 minutos');
}

/** Menu: Portal > Sincronizar aba SUPERVISORES. */
function atualizarAbaSupervisores() {
  try {
    const r = sincronizarInputs_('menu');
    avisar_('Aba ' + ABA_SUP_INPUTS + ' sincronizada.\n\n' +
            r.novos + ' projetos novos vindos da planilha\n' +
            r.alterados + ' projetos alterados na planilha\n' +
            r.linhas + ' linhas na aba (ativos + concluídos dos últimos ' +
            DIAS_CONCLUIDO_NA_ABA + ' dias)');
  } catch (e) {
    avisar_('Não deu certo:\n\n' + e.message);
  }
}
