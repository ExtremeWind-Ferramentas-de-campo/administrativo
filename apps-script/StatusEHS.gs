/* ===========================================================================
   STATUS RD EHS — Reporte Diário de EHS
   ---------------------------------------------------------------------------
   Mesma ideia do Status RDO, em outra planilha: lê os reportes de EHS (só
   leitura) e cruza com os projetos EM ANDAMENTO para cobrar quem não enviou.

   Arquivo separado de propósito. Compartilha os utilitários do
   SupervisaoCampo.gs (prop_, dataISO_, chaveTexto_, acharColuna_, lerProjetos_,
   lerTecnicos_, ehFimDeSemana_) porque no Apps Script todos os arquivos vivem
   no mesmo escopo — mas nada do RDO precisa ser mexido para o EHS funcionar.

   PROPRIEDADES (Configurações do projeto > Propriedades do script):
     ID_EHS    link ou ID da planilha do reporte diário de EHS
     ABA_EHS   nome da aba (padrão: REPORT DIÁRIO DE EHS - EXTREME WIND)
   =========================================================================== */

const ABA_EHS_PADRAO = 'REPORT DIÁRIO DE EHS - EXTREME WIND';

function idPlanilhaEHS_() {
  const id = extrairId_(prop_('ID_EHS'));
  if (!id) throw new Error('A planilha do EHS ainda não foi configurada. ' +
                           'Crie a propriedade ID_EHS em Configurações do projeto > ' +
                           'Propriedades do script, ou use Portal > Configurar planilha do EHS.');
  return id;
}
function abaNomeEHS_() { return prop_('ABA_EHS', ABA_EHS_PADRAO); }

/* Apelidos aceitos para cada coluna. A comparação ignora maiúsculas, acentos e
   espaços. Acrescente aqui se o cabeçalho da planilha mudar.

   "data" vem antes de "carimbo_de_data_hora" e a busca tenta nome idêntico em
   todos os apelidos antes de tentar nome que contenha: sem isso, numa planilha
   de Formulários o carimbo de data/hora roubaria a coluna Data. */
const COLUNAS_EHS = {
  data:    ['data', 'data_exp', 'data_relatorio', 'data_do_relatorio', 'data_execucao', 'carimbo_de_data_hora'],
  parque:  ['parque', 'parque_eolico', 'nome_parque', 'usina', 'complexo', 'pe'],
  cliente: ['cliente', 'contratante', 'empresa'],
  link:    ['link_do_pdf', 'link_pdf', 'linkpdf', 'link', 'pdf', 'url_pdf', 'arquivo'],
  autor:   ['matricula_login', 'matricula', 'responsavel', 'tecnico', 'supervisor', 'criado_por', 'usuario']
};

/* Só data, parque e link seguram o funcionamento. Cliente e autor podem faltar
   sem derrubar a tela — o filtro de cliente fica vazio e o card não diz quem
   enviou, mas a cobrança continua de pé. */
const COLUNAS_EHS_ESSENCIAIS = ['data', 'parque', 'link'];

function abaEHS_() {
  const ss = SpreadsheetApp.openById(idPlanilhaEHS_());
  const nome = abaNomeEHS_();
  const aba = ss.getSheetByName(nome);
  if (!aba) {
    const nomes = ss.getSheets().map(function (s) { return s.getName(); }).join(' | ');
    throw new Error('A aba "' + nome + '" não existe na planilha do EHS. Abas encontradas: ' + nomes);
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
  const cache = CacheService.getScriptCache();
  const guardado = cache.get('ehs_dados');
  if (guardado) {
    try { return JSON.parse(guardado); } catch (e) { /* cache ruim: relê */ }
  }

  const aba = abaEHS_();
  const faixa = aba.getDataRange();
  const dados = faixa.getValues();
  if (dados.length < 2) return [];

  // Igual ao RDO: o texto EXIBIDO é o que a pessoa vê na planilha. O valor cru
  // transforma coisas como "5/5" em Date e o card mostraria a data por extrato.
  const exibido = faixa.getDisplayValues();
  const txt = function (l, c) {
    if (c === -1) return '';
    const v = exibido[l][c];
    return String(v == null ? '' : v).trim();
  };

  const mapa = detectarColunasEHS_(dados[0]);
  const faltando = COLUNAS_EHS_ESSENCIAIS.filter(function (c) { return mapa[c] === -1; });
  if (faltando.length) {
    throw new Error('Colunas não encontradas na aba "' + abaNomeEHS_() + '": ' +
                    faltando.join(', ') + '. Cabeçalho lido: ' +
                    dados[0].filter(String).join(' | ') +
                    '. Acrescente o nome real em COLUNAS_EHS, no StatusEHS.gs.');
  }

  const linhas = [];
  for (let l = 1; l < dados.length; l++) {
    const linha = dados[l];
    // A data segue pelo valor cru, que é Date de verdade; o resto pelo exibido.
    const data = mapa.data > -1 ? (dataISO_(linha[mapa.data]) || dataISO_(txt(l, mapa.data))) : '';
    const celulaParque = txt(l, mapa.parque);
    if (!data && !celulaParque) continue;

    linhas.push({
      data:    data,
      parque:  celulaParque.replace(/\s+/g, ' '),
      cliente: txt(l, mapa.cliente),
      link:    txt(l, mapa.link),
      autor:   txt(l, mapa.autor)
    });
  }

  try {
    cache.put('ehs_dados', JSON.stringify(linhas), CACHE_RDO_SEGUNDOS);
  } catch (e) {
    // passou do limite do cache: segue sem, só fica mais lento
  }
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

  // Casamento pelo nome do parque, mesma regra do RDO.
  const porParque = {};
  doDia.forEach(function (r) {
    if (!porParque[chaveTexto_(r.parque)]) porParque[chaveTexto_(r.parque)] = r;
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

  const projetos = lerProjetos_()
    .filter(function (pr) { return String(pr.status || 'andamento') === 'andamento'; })
    .map(function (pr) {
      const achado = porParque[chaveTexto_(pr.parque)] || null;
      return {
        id: pr.id, codigo: pr.codigo, parque: pr.parque, cliente: pr.cliente,
        tipoReparo: pr.tipoReparo,
        supervisor: (pr.supervisor && pr.supervisor.nome) ? pr.supervisor.nome : '',
        estado: achado ? 'ENVIADO' : (fds ? 'NAO_OBRIGATORIO' : 'FALTA'),
        link: achado ? achado.link : '',
        autor: quemFez(achado)
      };
    });

  const fCliente = chaveTexto_(p.cliente || '');
  const fParque  = chaveTexto_(p.parque || '');
  const relatorios = linhas.filter(function (r) {
    if (p.data && r.data !== data) return false;
    if (fCliente && chaveTexto_(r.cliente) !== fCliente) return false;
    if (fParque && chaveTexto_(r.parque) !== fParque) return false;
    return true;
  }).slice(0, 400);

  relatorios.forEach(function (r) { r.autorNome = quemFez(r); });
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


/* ---------------------------------------------------------------------------
   MENU E DIAGNÓSTICO
   --------------------------------------------------------------------------- */

/** Menu: Portal > Configurar planilha do EHS */
function menuConfigurarEHS() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const r1 = ui.prompt('Planilha do EHS',
    'Cole o link (ou o ID) da planilha do reporte diário de EHS.\n\n' +
    'Atual: ' + (props.getProperty('ID_EHS') || '(não configurada)'),
    ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const id = extrairId_(r1.getResponseText());
  if (!id) { ui.alert('Não consegui extrair o ID desse texto.'); return; }

  const r2 = ui.prompt('Aba do EHS',
    'Nome da aba onde ficam os reportes.\n\nAtual: ' + abaNomeEHS_(),
    ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;

  props.setProperties({
    ID_EHS: id,
    ABA_EHS: r2.getResponseText().trim() || ABA_EHS_PADRAO
  }, false);
  CacheService.getScriptCache().remove('ehs_dados');

  menuColunasEHS();
}

/** Menu: Portal > Conferir colunas do EHS */
function menuColunasEHS() {
  let txt = 'PLANILHA DO EHS\n\n';
  txt += '  ID_EHS  : ' + (PropertiesService.getScriptProperties().getProperty('ID_EHS') || '(não configurado)') + '\n';
  txt += '  ABA_EHS : ' + abaNomeEHS_() + '\n\n';

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
