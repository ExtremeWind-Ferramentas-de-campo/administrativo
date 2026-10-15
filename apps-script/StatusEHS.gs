/* ===========================================================================
   STATUS RD EHS — Reporte Diário de EHS
   ---------------------------------------------------------------------------
   Mesma ideia do Status RDO: lê os reportes de EHS (só leitura) e cruza com
   os projetos EM ANDAMENTO para cobrar quem não enviou.

   out/2026: os reportes saíram da planilha do construtor de formulários e
   passaram a ser gravados pelo backend do RDO (ReportEHS.gs) na aba
   "Reports EHS" da MESMA planilha do banco de dados do RDO. Por isso:
     - a planilha é a do ID_RDO (nenhuma propriedade nova para configurar);
     - as propriedades antigas ID_EHS e ABA_EHS são IGNORADAS de propósito:
       elas apontam para a planilha velha e, se valessem, a tela continuaria
       mostrando os reportes antigos sem dar erro nenhum.

   Colunas fixas da aba (gravadas pelo ReportEHS.gs):
     Protocolo | Recebido_em | Matricula_login | Data | Cliente | Parque |
     Supervisor | Link_PDF | Caminho_PDF | (uma coluna por pergunta...)

   Arquivo separado de propósito. Compartilha os utilitários do
   SupervisaoCampo.gs (prop_, idPlanilhaRDO_, dataISO_, chaveTexto_,
   acharColuna_, lerProjetos_, lerTecnicos_, ehFimDeSemana_).

   PROPRIEDADE (opcional, Configurações do projeto > Propriedades do script):
     ABA_EHS_REPORTS   nome da aba (padrão: Reports EHS)
   =========================================================================== */

const ABA_EHS_PADRAO = 'Reports EHS';
const CACHE_EHS_CHAVE = 'ehs_reports_v2';   // chave nova: o cache da planilha antiga não volta

function idPlanilhaEHS_() { return idPlanilhaRDO_(); }
function abaNomeEHS_() { return prop_('ABA_EHS_REPORTS', ABA_EHS_PADRAO); }

/* Apelidos aceitos para cada coluna. A comparação ignora maiúsculas, acentos e
   espaços. Acrescente aqui se o cabeçalho da planilha mudar.

   "data" vem antes de "carimbo_de_data_hora" e a busca tenta nome idêntico em
   todos os apelidos antes de tentar nome que contenha: sem isso, numa planilha
   de Formulários o carimbo de data/hora roubaria a coluna Data. */
const COLUNAS_EHS = {
  data:       ['data', 'data_exp', 'data_relatorio', 'data_do_relatorio'],
  parque:     ['parque', 'parque_eolico', 'nome_parque'],
  cliente:    ['cliente', 'contratante', 'empresa'],
  link:       ['link_pdf', 'link_do_pdf', 'linkpdf', 'url_pdf'],
  autor:      ['matricula_login', 'matricula'],
  supervisor: ['supervisor'],
  recebido:   ['recebido_em']
};

/* Só data, parque e link seguram o funcionamento. Os outros podem faltar
   sem derrubar a tela — o filtro de cliente fica vazio e o card não diz quem
   enviou, mas a cobrança continua de pé. */
const COLUNAS_EHS_ESSENCIAIS = ['data', 'parque', 'link'];

function abaEHS_() {
  const ss = SpreadsheetApp.openById(idPlanilhaEHS_());
  const nome = abaNomeEHS_();
  // tolera maiúscula/acento/espaço: "Reports EHS", "REPORTS EHS", "Reports  EHS "
  const alvo = normalizarCab_(nome);
  const aba = ss.getSheetByName(nome) ||
              ss.getSheets().filter(function (s) { return normalizarCab_(s.getName()) === alvo; })[0];
  if (!aba) {
    const nomes = ss.getSheets().map(function (s) { return s.getName(); }).join(' | ');
    throw new Error('A aba "' + nome + '" não existe na planilha do RDO. Abas encontradas: ' + nomes +
                    '. Ela é criada sozinha no primeiro reporte enviado pelo app.');
  }
  return aba;
}

function detectarColunasEHS_(cabecalho) {
  const norm = cabecalho.map(normalizarCab_);
  const mapa = {};
  Object.keys(COLUNAS_EHS).forEach(function (campo) {
    mapa[campo] = acharColuna_(norm, COLUNAS_EHS[campo]);
  });
  return mapa;
}

/** Lê a planilha do EHS inteira, com cache curto. */
function lerEHS_() {
  const guardado = lerGrande_(CACHE_EHS_CHAVE);
  if (guardado) {
    try { return JSON.parse(guardado); } catch (e) { /* cache ruim: relê */ }
  }

  const aba = abaEHS_();
  const nLin = aba.getLastRow();
  if (nLin < 2) return [];
  const cab = aba.getRange(1, 1, 1, Math.max(1, aba.getLastColumn())).getValues()[0];
  const mapa = detectarColunasEHS_(cab);
  const faltando = COLUNAS_EHS_ESSENCIAIS.filter(function (c) { return mapa[c] === -1; });
  if (faltando.length) {
    throw new Error('Colunas não encontradas na aba "' + aba.getName() + '": ' +
                    faltando.join(', ') + '. Cabeçalho lido: ' + cab.filter(String).join(' | ') +
                    '. Acrescente o nome real em COLUNAS_EHS, no StatusEHS.gs.');
  }

  // Depois das colunas fixas vem uma coluna por pergunta do formulário, com
  // texto longo. Ler só até a última coluna usada aqui evita trazer tudo isso.
  const ultimaCol = Math.max.apply(null, Object.keys(mapa).map(function (k) { return mapa[k]; })) + 1;
  const faixa = aba.getRange(1, 1, nLin, ultimaCol);
  const dados = faixa.getValues();

  // Igual ao RDO: o texto EXIBIDO é o que a pessoa vê na planilha. O valor cru
  // transforma coisas como "5/5" em Date e o card mostraria a data por extrato.
  const exibido = faixa.getDisplayValues();
  const txt = function (l, c) {
    if (c === -1) return '';
    const v = exibido[l][c];
    return String(v == null ? '' : v).trim();
  };

  const linhas = [];
  for (let l = 1; l < dados.length; l++) {
    const linha = dados[l];
    // A data segue pelo valor cru, que é Date de verdade; o resto pelo exibido.
    const data = mapa.data > -1 ? (dataISO_(linha[mapa.data]) || dataISO_(txt(l, mapa.data))) : '';
    const celulaParque = txt(l, mapa.parque);
    if (!data && !celulaParque) continue;

    linhas.push({
      data:       data,
      parque:     celulaParque.replace(/\s+/g, ' '),
      cliente:    txt(l, mapa.cliente),
      link:       txt(l, mapa.link),
      autor:      txt(l, mapa.autor),
      supervisor: txt(l, mapa.supervisor),
      recebido:   txt(l, mapa.recebido)
    });
  }

  // Em pedaços (guardarGrande_, no Codigo.gs): numa chave só o cache morre em
  // silêncio quando a aba passa de ~100 KB — e esta aba cresce ~30 linhas/dia.
  guardarGrande_(CACHE_EHS_CHAVE, JSON.stringify(linhas), CACHE_RDO_SEGUNDOS);
  return linhas;
}


/* ---------------------------------------------------------------------------
   AÇÕES DA API
   --------------------------------------------------------------------------- */

/**
 * -> { data, fimDeSemana, projetos[], relatorios[], total }
 *
 * A cobrança sai dos projetos EM ANDAMENTO: cada um deveria ter um reporte de
 * EHS na data. Sábado e domingo entram como "não obrigatório".
 */
function acaoEhsStatus_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  let linhas;
  try {
    linhas = lerEHS_();
  } catch (e) {
    return { ok: false, motivo: 'EHS_INDISPONIVEL', detalhe: e.message };
  }

  const data = dataISO_(p.data) || dataISO_(new Date());
  const fds = ehFimDeSemana_(data);

  const doDia = linhas.filter(function (r) { return r.data === data; });

  // Casamento pelo nome do parque, mesma regra do RDO. O app grava UM report
  // por matrícula por dia, então o mesmo parque pode ter mais de um (cada
  // técnico que logou e enviou). A cobrança fica satisfeita com qualquer um;
  // o card mostra todos que enviaram e abre o PDF do primeiro.
  const porParque = {};
  doDia.forEach(function (r) {
    const k = chaveTexto_(r.parque);
    (porParque[k] = porParque[k] || []).push(r);
  });

  let nomePorMatricula = {};
  try {
    lerTecnicos_().forEach(function (t) { nomePorMatricula[soNumero_(t.matricula)] = t.nome; });
  } catch (e) { /* sem MINI MASTER, mostra a matrícula mesmo */ }
  const quemFez = function (r) {
    if (!r) return '';
    const n = nomePorMatricula[soNumero_(r.autor)];
    return n || r.autor || '';
  };

  // Autorizado (Sim/Não) é por parque + dia, marcado no portal (aba EHS_AUTORIZACOES).
  let autorizacoes = {};
  try { autorizacoes = lerAutorizacoesEHS_(); } catch (e) { /* sem aba ainda: tudo em branco */ }

  const fCliente = chaveTexto_(p.cliente || '');
  const fParque  = chaveTexto_(p.parque || '');

  const projetos = lerProjetos_()
    .filter(function (pr) { return String(pr.status || 'andamento') === 'andamento'; })
    .filter(function (pr) { return projetoPassaFiltro_(pr, fCliente, fParque); })
    .map(function (pr) {
      const achados = porParque[chaveTexto_(pr.parque)] || [];
      const comLink = achados.filter(function (r) { return r.link; })[0] || achados[0] || null;
      const autores = [];
      achados.forEach(function (r) {
        const n = quemFez(r);
        if (n && autores.indexOf(n) === -1) autores.push(n);
      });
      return {
        id: pr.id, codigo: pr.codigo, parque: pr.parque, cliente: pr.cliente,
        tipoReparo: pr.tipoReparo,
        supervisor: (pr.supervisor && pr.supervisor.nome) ? pr.supervisor.nome
                  : (comLink ? comLink.supervisor : ''),
        estado: achados.length ? 'ENVIADO' : (fds ? 'NAO_OBRIGATORIO' : 'FALTA'),
        link: comLink ? comLink.link : '',
        autor: autores.join(', '),
        qtd: achados.length
      };
    });

  const relatorios = linhas.filter(function (r) {
    if (p.data && r.data !== data) return false;
    if (fCliente && chaveTexto_(r.cliente) !== fCliente) return false;
    if (fParque && chaveTexto_(r.parque) !== fParque) return false;
    return true;
  }).slice(-400);   // as mais recentes ficam no fim da aba

  // Autorizado é por parque + data do PRÓPRIO report (vale também quando a
  // consulta é sem data e a lista traz vários dias).
  relatorios.forEach(function (r) {
    r.autorNome = quemFez(r);
    const a = autorizacoes[chaveAutoriz_(r.data, r.parque)] || {};
    r.autorizado = a.valor || '';
    r.autorizadoPor = a.valor ? (nomePorMatricula[soNumero_(a.matricula)] || a.matricula || '') : '';
  });
  relatorios.sort(function (a, b) { return String(a.parque).localeCompare(String(b.parque)); });

  return {
    ok: true,
    data: data,
    fimDeSemana: fds,
    projetos: projetos,
    relatorios: relatorios,
    total: relatorios.length
  };
}

/** -> listas para os menus de filtro (clientes e parques que existem no EHS) */
function acaoEhsFiltros_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  let linhas;
  try {
    linhas = lerEHS_();
  } catch (e) {
    return { ok: false, motivo: 'EHS_INDISPONIVEL', detalhe: e.message };
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


/* A lista de clientes e parques muda devagar: chave própria com cache longo.
   CACHE_FILTROS_SEG mora no SupervisaoCampo.gs (o mesmo do RDO). Não declarar
   de novo aqui: const repetida entre arquivos derruba o projeto inteiro. */
function acaoEhsFiltrosCache_(p) {
  const chave = CACHE_EHS_CHAVE + '_filtros';
  const guardado = lerGrande_(chave);
  if (guardado) {
    try { return JSON.parse(guardado); } catch (e) { /* cache ruim: refaz */ }
  }
  const res = acaoEhsFiltros_(p);
  const seg = (typeof CACHE_FILTROS_SEG !== 'undefined') ? CACHE_FILTROS_SEG : 900;
  if (res.ok) guardarGrande_(chave, JSON.stringify(res), seg);
  return res;
}


/* ---------------------------------------------------------------------------
   AUTORIZADO (Sim / Não) — por parque + dia
   ---------------------------------------------------------------------------
   Fica na planilha do PORTAL (não na do RDO): o portal só lê a do RDO, e a
   marcação é uma decisão tomada aqui. A aba nasce sozinha na primeira marcação.
   Uma linha por data + parque; marcar de novo sobrescreve; clicar de novo
   na opção marcada desmarca (a linha sai). Quem marcou e
   quando ficam na própria linha, e o LOG guarda o histórico de trocas.
   Qualquer pessoa logada pode marcar (decisão de out/2026).
   --------------------------------------------------------------------------- */

const ABA_EHS_AUTORIZ = 'EHS_AUTORIZACOES';
const COLUNAS_EHS_AUTORIZ = ['chave', 'data', 'parque', 'autorizado', 'matricula', 'atualizado_em'];
const VALORES_AUTORIZ = ['SIM', 'NAO'];

function abaAutorizEHS_(criar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(ABA_EHS_AUTORIZ);
  if (!aba && criar) {
    aba = ss.insertSheet(ABA_EHS_AUTORIZ);
    aba.getRange(1, 1, 1, COLUNAS_EHS_AUTORIZ.length).setValues([COLUNAS_EHS_AUTORIZ]).setFontWeight('bold');
    aba.setFrozenRows(1);
    // data e chave como texto: sem isso o Sheets converte "2026-10-14" em Date
    aba.getRange('A:B').setNumberFormat('@');
  }
  return aba;
}

function chaveAutoriz_(dataIso, parque) { return dataIso + '|' + chaveTexto_(parque); }

/** -> { 'aaaa-mm-dd|parque': { valor, matricula, em } } — a aba inteira (é pequena) */
function lerAutorizacoesEHS_() {
  const aba = abaAutorizEHS_(false);
  const mapa = {};
  if (!aba || aba.getLastRow() < 2) return mapa;
  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, COLUNAS_EHS_AUTORIZ.length).getDisplayValues();
  dados.forEach(function (l) {
    const valor = String(l[3]).toUpperCase();
    if (VALORES_AUTORIZ.indexOf(valor) === -1) return;
    mapa[chaveAutoriz_(String(l[1]), l[2])] = { valor: valor, matricula: String(l[4]), em: String(l[5]) };
  });
  return mapa;
}

/** { data, parque, valor: 'SIM' | 'NAO' | '' } -> { ok, data, parque, autorizado }
 *  valor vazio = desmarcar (clique sem querer): a linha sai da aba. */
function acaoEhsAutorizar_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  const data = dataISO_(p.data);
  const parque = String(p.parque || '').replace(/\s+/g, ' ').trim();
  const valor = String(p.valor || '').toUpperCase();
  if (!data || !parque || (valor && VALORES_AUTORIZ.indexOf(valor) === -1)) {
    return { ok: false, motivo: 'DADOS_INVALIDOS' };
  }

  const trava = LockService.getScriptLock();
  if (!trava.tryLock(20000)) return { ok: false, motivo: 'OCUPADO' };
  try {
    const aba = abaAutorizEHS_(true);
    const chave = chaveAutoriz_(data, parque);
    const agora = new Date();
    const linhaNova = [chave, data, parque, valor, "'" + matricula, agora];

    let linha = -1;
    if (aba.getLastRow() >= 2) {
      const chaves = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getDisplayValues();
      for (let i = 0; i < chaves.length; i++) {
        if (String(chaves[i][0]) === chave) { linha = i + 2; break; }
      }
    }
    if (!valor) {
      if (linha > -1) aba.deleteRow(linha);
    } else if (linha > -1) {
      aba.getRange(linha, 1, 1, linhaNova.length).setValues([linhaNova]);
    } else {
      aba.appendRow(linhaNova);
    }

    registrar_(matricula, 'EHS_AUTORIZAR', 'OK', data + ' | ' + parque + ' | ' + (valor || 'DESMARCADO'));
    return { ok: true, data: data, parque: parque, autorizado: valor };
  } finally {
    trava.releaseLock();
  }
}


/* ---------------------------------------------------------------------------
   MENU E DIAGNÓSTICO
   --------------------------------------------------------------------------- */

/** Menu: Portal > Configurar aba do EHS
 *  A planilha é sempre a do RDO (ID_RDO); aqui só se troca o nome da aba. */
function menuConfigurarEHS() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Aba dos reportes de EHS',
    'Os reportes ficam na planilha do RDO. Nome da aba:\n\nAtual: ' + abaNomeEHS_(),
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  PropertiesService.getScriptProperties()
    .setProperty('ABA_EHS_REPORTS', r.getResponseText().trim() || ABA_EHS_PADRAO);
  CacheService.getScriptCache().removeAll([CACHE_EHS_CHAVE + '_n', CACHE_EHS_CHAVE + '_filtros_n']);
  menuColunasEHS();
}

/** Menu: Portal > Conferir colunas do EHS */
function menuColunasEHS() {
  const props = PropertiesService.getScriptProperties();
  CacheService.getScriptCache().removeAll([CACHE_EHS_CHAVE + '_n', CACHE_EHS_CHAVE + '_filtros_n']);
  let txt = 'REPORTES DE EHS\n\n';
  txt += '  Planilha : a do RDO (ID_RDO = ' + (props.getProperty('ID_RDO') || '(não configurado)') + ')\n';
  txt += '  Aba      : ' + abaNomeEHS_() + '\n';
  if (props.getProperty('ID_EHS') || props.getProperty('ABA_EHS')) {
    txt += '\n  Aviso: ID_EHS / ABA_EHS (planilha antiga) ainda existem nas propriedades.\n' +
           '  Não são mais usadas — pode apagar.\n';
  }
  txt += '\n';

  try {
    const aba = abaEHS_();
    const cab = aba.getRange(1, 1, 1, Math.max(1, aba.getLastColumn())).getValues()[0];
    const mapa = detectarColunasEHS_(cab);

    txt += 'Cabeçalho lido:\n  ' + cab.filter(String).join(' | ') + '\n\nColunas:';
    Object.keys(COLUNAS_EHS).forEach(function (campo) {
      const i = mapa[campo];
      const essencial = COLUNAS_EHS_ESSENCIAIS.indexOf(campo) > -1;
      txt += '\n  ' + campo + ': ' +
             (i === -1
               ? (essencial ? 'NÃO ENCONTRADA (essencial)' : 'não encontrada (opcional)')
               : '"' + cab[i] + '"');
    });

    const linhas = lerEHS_();
    txt += '\n\nLinhas lidas: ' + linhas.length;
    if (linhas.length) {
      const u = linhas[linhas.length - 1];
      txt += '\n  Última: ' + (u.data || 'sem data') + ' — ' + (u.parque || 'sem parque') +
             (u.link ? ' — com link' : ' — SEM LINK');
    }
  } catch (e) {
    txt += 'ERRO — ' + e.message;
  }

  Logger.log(txt);
  avisar_(txt);
}
