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
const CACHE_EHS_CHAVE = 'ehs_reports_v3';   // v3: entrou o caminho do PDF (botão WhatsApp)

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
  recebido:   ['recebido_em'],
  caminho:    ['caminho_pdf']
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
      recebido:   txt(l, mapa.recebido),
      caminho:    txt(l, mapa.caminho)
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

  // Stand by: parque que não precisava enviar no dia, marcado com motivo no portal.
  let standby = {};
  try { standby = lerStandbyEHS_(data); } catch (e) { /* sem aba ainda */ }

  const projetos = lerProjetos_()
    .filter(function (pr) { return cobravelNaData_(pr, data); })   // início depois da data não é cobrado
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
        estado: achados.length ? 'ENVIADO'
              : standby[chaveTexto_(pr.parque)] ? 'STANDBY'
              : (fds ? 'NAO_OBRIGATORIO' : 'FALTA'),
        link: comLink ? comLink.link : '',
        autor: autores.join(', '),
        qtd: achados.length,
        standbyMotivo: achados.length ? '' : ((standby[chaveTexto_(pr.parque)] || {}).motivo || ''),
        standbyPor: achados.length ? '' : (function (sb) {
          return sb ? (nomePorMatricula[soNumero_(sb.matricula)] || sb.matricula || '') : '';
        })(standby[chaveTexto_(pr.parque)])
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
    r.temArquivos = !!r.caminho;
    delete r.caminho;          // o caminho do Dropbox não precisa ir para o navegador
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

  // reports + projetos cadastrados (opcoesDeFiltro_, no SupervisaoCampo.gs):
  // cliente sem nenhum report ainda (ex.: GE) também aparece no menu.
  return opcoesDeFiltro_(linhas);
}


/* A lista de clientes e parques muda devagar: chave própria com cache longo.
   CACHE_FILTROS_SEG mora no SupervisaoCampo.gs (o mesmo do RDO). Não declarar
   de novo aqui: const repetida entre arquivos derruba o projeto inteiro. */
function acaoEhsFiltrosCache_(p) {
  const chave = CACHE_EHS_CHAVE + '_filtros_v2';   // v2: inclui clientes/parques dos projetos
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
   STAND BY — parque que não precisava enviar o report no dia
   ---------------------------------------------------------------------------
   Mesmo esquema do Autorizado: aba EHS_STANDBY na planilha do portal, uma
   linha por data + parque, criada sozinha. Marcar exige motivo; desfazer
   apaga a linha. Se o report chegar depois, ele vence: o parque aparece como
   enviado e a marcação é ignorada.
   --------------------------------------------------------------------------- */

const ABA_EHS_STANDBY = 'EHS_STANDBY';
const COLUNAS_EHS_STANDBY = ['chave', 'data', 'parque', 'motivo', 'matricula', 'atualizado_em'];
const STANDBY_MOTIVO_MAX = 300;

function abaStandbyEHS_(criar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(ABA_EHS_STANDBY);
  if (!aba && criar) {
    aba = ss.insertSheet(ABA_EHS_STANDBY);
    aba.getRange(1, 1, 1, COLUNAS_EHS_STANDBY.length).setValues([COLUNAS_EHS_STANDBY]).setFontWeight('bold');
    aba.setFrozenRows(1);
    aba.getRange('A:B').setNumberFormat('@');
  }
  return aba;
}

/** -> { chaveTexto(parque): { motivo, matricula, em } } da data pedida */
function lerStandbyEHS_(dataIso) {
  const aba = abaStandbyEHS_(false);
  const mapa = {};
  if (!aba || aba.getLastRow() < 2) return mapa;
  aba.getRange(2, 1, aba.getLastRow() - 1, COLUNAS_EHS_STANDBY.length).getDisplayValues()
    .forEach(function (l) {
      if (String(l[1]) !== dataIso || !String(l[3]).trim()) return;
      mapa[chaveTexto_(l[2])] = { motivo: String(l[3]), matricula: String(l[4]), em: String(l[5]) };
    });
  return mapa;
}

/** { data, parque, motivo } -> { ok }   motivo vazio = desfazer o stand by */
function acaoEhsStandby_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  const data = dataISO_(p.data);
  const parque = String(p.parque || '').replace(/\s+/g, ' ').trim();
  const motivo = String(p.motivo || '').replace(/\s+/g, ' ').trim().slice(0, STANDBY_MOTIVO_MAX);
  if (!data || !parque) return { ok: false, motivo: 'DADOS_INVALIDOS' };

  const trava = LockService.getScriptLock();
  if (!trava.tryLock(20000)) return { ok: false, motivo: 'OCUPADO' };
  try {
    const aba = abaStandbyEHS_(true);
    const chave = chaveAutoriz_(data, parque);
    let linha = -1;
    if (aba.getLastRow() >= 2) {
      const chaves = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getDisplayValues();
      for (let i = 0; i < chaves.length; i++) {
        if (String(chaves[i][0]) === chave) { linha = i + 2; break; }
      }
    }
    if (!motivo) {
      if (linha > -1) aba.deleteRow(linha);
    } else {
      const nova = [chave, data, parque, motivo, "'" + matricula, new Date()];
      if (linha > -1) aba.getRange(linha, 1, 1, nova.length).setValues([nova]);
      else aba.appendRow(nova);
    }
    registrar_(matricula, 'EHS_STANDBY', 'OK', data + ' | ' + parque + ' | ' + (motivo || 'DESFEITO'));
    return { ok: true, data: data, parque: parque, motivo: motivo };
  } finally {
    trava.releaseLock();
  }
}


/* ---------------------------------------------------------------------------
   ARQUIVOS DO REPORT (botão "Enviar no WhatsApp")
   ---------------------------------------------------------------------------
   Devolve o PDF e as fotos da pasta do dia (em base64) para o navegador
   montar o compartilhamento do celular — a pessoa escolhe o grupo e envia.
   Nada é enviado daqui: o portal só busca os arquivos.

   O caminho vem da coluna Caminho_PDF da aba Reports EHS (nunca do
   navegador) e precisa estar dentro de EHS_PASTA, para este botão não virar
   leitor do Dropbox inteiro.
     - pasta do dia (…/W42/14-10-2026/): manda tudo o que tem nela;
     - report antigo, direto na pasta da semana: manda só o PDF.

   PROPRIEDADES DO SCRIPT (copie do projeto do RDO, mesmos valores):
     DROPBOX_APP_KEY, DROPBOX_APP_SECRET, DROPBOX_REFRESH_TOKEN
   Opcional:
     EHS_PASTA    padrão igual ao do RDO (16 - CONTROLE DE DOCUMENTAÇÃO DIÁRIA)
   Diagnóstico à mão: testarArquivosEHS()
   --------------------------------------------------------------------------- */

const EHS_PASTA_PADRAO_PORTAL = '/02 - EXTREME WIND/5 - SEGURANCA E MEIO AMBIENTE/16 - CONTROLE DE DOCUMENTAÇÃO DIÁRIA';
const EHS_ARQ_MAX_BYTES = 25 * 1024 * 1024;   // soma de tudo que volta numa resposta
const EHS_ARQ_MAX_QTD = 30;

function ehsPastaPortal_() {
  let pasta = String(prop_('EHS_PASTA', EHS_PASTA_PADRAO_PORTAL)).replace(/\\/g, '/').trim();
  if (pasta.charAt(0) !== '/') pasta = '/' + pasta;
  return pasta.replace(/\/+$/, '');
}

function dbxTokenPortal_() {
  const props = PropertiesService.getScriptProperties();
  const cache = CacheService.getScriptCache();
  const guardado = cache.get('dbx_token_portal');
  if (guardado) return guardado;
  const faltam = ['DROPBOX_APP_KEY', 'DROPBOX_APP_SECRET', 'DROPBOX_REFRESH_TOKEN']
    .filter(function (k) { return !props.getProperty(k); });
  if (faltam.length) throw new Error('Faltam as propriedades do Dropbox no portal: ' + faltam.join(', '));
  const r = UrlFetchApp.fetch('https://api.dropbox.com/oauth2/token', {
    method: 'post', muteHttpExceptions: true,
    payload: {
      grant_type: 'refresh_token',
      refresh_token: props.getProperty('DROPBOX_REFRESH_TOKEN'),
      client_id: props.getProperty('DROPBOX_APP_KEY'),
      client_secret: props.getProperty('DROPBOX_APP_SECRET')
    }
  });
  const j = JSON.parse(r.getContentText());
  if (!j.access_token) throw new Error('Dropbox recusou as credenciais: ' + r.getContentText().slice(0, 200));
  cache.put('dbx_token_portal', j.access_token, Math.min(3000, Number(j.expires_in || 3600) - 300));
  return j.access_token;
}

/* A pasta é do time: na 1ª chamada que der "não achado", tenta pela raiz do
   time e lembra a escolha (mesma lógica do RDO).
   bytes: só no upload (files/upload) — o conteúdo do arquivo a gravar.
   Usado também pelo Avaliacao.gs. */
function dbxPortal_(url, corpo, conteudo, bytes) {
  const token = dbxTokenPortal_();
  const props = PropertiesService.getScriptProperties();
  const modos = props.getProperty('DBX_PORTAL_ROOT') === 'ROOT' ? ['ROOT'] : ['', 'ROOT'];
  let ultimo = null;
  for (let i = 0; i < modos.length; i++) {
    const h = { Authorization: 'Bearer ' + token };
    if (modos[i] === 'ROOT') {
      const ns = dbxRootNsPortal_(token);
      if (!ns) break;
      h['Dropbox-API-Path-Root'] = JSON.stringify({ '.tag': 'root', root: String(ns) });
    }
    const op = { method: 'post', headers: h, muteHttpExceptions: true };
    if (conteudo) {
      h['Dropbox-API-Arg'] = JSON.stringify(corpo).replace(/[\u007f-\uffff]/g, function (c) {
        return '\\u' + ('0000' + c.charCodeAt(0).toString(16)).slice(-4);
      });
      if (bytes) { op.contentType = 'application/octet-stream'; op.payload = bytes; }
    } else {
      op.contentType = 'application/json';
      op.payload = JSON.stringify(corpo);
    }
    const r = UrlFetchApp.fetch(url, op);
    ultimo = r;
    if (r.getResponseCode() === 200) {
      if (modos[i] === 'ROOT') props.setProperty('DBX_PORTAL_ROOT', 'ROOT');
      return r;
    }
    if (r.getResponseCode() !== 409) break;
  }
  throw new Error('Dropbox: ' + (ultimo ? ultimo.getResponseCode() + ' ' + ultimo.getContentText().slice(0, 200) : 'sem resposta'));
}

function dbxRootNsPortal_(token) {
  const c = CacheService.getScriptCache();
  const g = c.get('dbx_root_ns_portal');
  if (g) return g;
  const r = UrlFetchApp.fetch('https://api.dropboxapi.com/2/users/get_current_account', {
    method: 'post', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
  });
  if (r.getResponseCode() !== 200) return null;
  const ns = (JSON.parse(r.getContentText()).root_info || {}).root_namespace_id;
  if (ns) c.put('dbx_root_ns_portal', String(ns), 21600);
  return ns;
}

function tipoArquivo_(nome) {
  const n = String(nome).toLowerCase();
  if (/\.pdf$/.test(n)) return 'application/pdf';
  if (/\.jpe?g$/.test(n)) return 'image/jpeg';
  if (/\.png$/.test(n)) return 'image/png';
  return '';
}

/** { data, parque } -> { ok, arquivos:[{nome, tipo, base64}], texto } */
function acaoEhsArquivos_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  const data = dataISO_(p.data);
  const alvo = chaveTexto_(p.parque || '');
  if (!data || !alvo) return { ok: false, motivo: 'DADOS_INVALIDOS' };

  let linhas;
  try { linhas = lerEHS_(); }
  catch (e) { return { ok: false, motivo: 'EHS_INDISPONIVEL', detalhe: e.message }; }

  // a linha mais recente do parque/dia (o RDO substitui, mas por garantia)
  const achadas = linhas.filter(function (r) { return r.data === data && chaveTexto_(r.parque) === alvo && r.caminho; });
  const r = achadas[achadas.length - 1];
  if (!r) return { ok: false, motivo: 'SEM_ARQUIVO', detalhe: 'Este report não tem o caminho do PDF na planilha.' };

  const raiz = ehsPastaPortal_().toLowerCase();
  const caminho = String(r.caminho).replace(/\\/g, '/');
  if (caminho.toLowerCase().indexOf(raiz + '/') !== 0 || caminho.indexOf('..') > -1) {
    registrar_(matricula, 'EHS_ARQUIVOS', 'NEGADO', caminho);
    return { ok: false, motivo: 'SEM_ARQUIVO', detalhe: 'O PDF está fora da pasta do EHS.' };
  }

  try {
    const pasta = caminho.slice(0, caminho.lastIndexOf('/'));
    let lista;
    if (/\/\d{2}-\d{2}-\d{4}$/.test(pasta)) {
      const lf = JSON.parse(dbxPortal_('https://api.dropboxapi.com/2/files/list_folder',
                                       { path: pasta, recursive: false }).getContentText());
      lista = (lf.entries || []).filter(function (e) { return e['.tag'] === 'file' && tipoArquivo_(e.name); });
    } else {
      lista = [{ name: caminho.slice(caminho.lastIndexOf('/') + 1), path_display: caminho, size: 0 }];
    }
    // PDF primeiro, depois as fotos em ordem de nome (DDS 1, DDS 2, ...)
    lista.sort(function (a, b) {
      const pa = /\.pdf$/i.test(a.name) ? 0 : 1, pb = /\.pdf$/i.test(b.name) ? 0 : 1;
      return (pa - pb) || a.name.localeCompare(b.name, 'pt-BR', { numeric: true });
    });

    const arquivos = [];
    let total = 0;
    lista.slice(0, EHS_ARQ_MAX_QTD).forEach(function (e) {
      if (total > EHS_ARQ_MAX_BYTES) return;
      const resp = dbxPortal_('https://content.dropboxapi.com/2/files/download', { path: e.path_display || e.path_lower }, true);
      const bytes = resp.getBlob().getBytes();
      total += bytes.length;
      if (total > EHS_ARQ_MAX_BYTES) return;
      arquivos.push({ nome: e.name, tipo: tipoArquivo_(e.name), base64: Utilities.base64Encode(bytes) });
    });

    registrar_(matricula, 'EHS_ARQUIVOS', 'OK', data + ' | ' + r.parque + ' | ' + arquivos.length + ' arquivo(s)');
    return {
      ok: true,
      arquivos: arquivos,
      incompleto: arquivos.length < lista.length,
      texto: 'Report diário de EHS — ' + r.parque + ' — ' + data.split('-').reverse().join('/')
    };
  } catch (e) {
    registrar_(matricula, 'EHS_ARQUIVOS', 'FALHA', String(e.message).slice(0, 300));
    return { ok: false, motivo: 'DROPBOX', detalhe: e.message };
  }
}

/** Rode à mão: confere credenciais, pasta e o report mais recente com PDF. */
function testarArquivosEHS() {
  Logger.log('Pasta do EHS: ' + ehsPastaPortal_());
  try { dbxTokenPortal_(); Logger.log('Dropbox: credenciais OK'); }
  catch (e) { Logger.log('Dropbox: ' + e.message); return; }
  const linhas = lerEHS_().filter(function (r) { return r.caminho; });
  if (!linhas.length) { Logger.log('Nenhum report com Caminho_PDF na planilha.'); return; }
  const u = linhas[linhas.length - 1];
  Logger.log('Último report: ' + u.data + ' — ' + u.parque + ' — ' + u.caminho);
  const pasta = u.caminho.slice(0, u.caminho.lastIndexOf('/'));
  try {
    const lf = JSON.parse(dbxPortal_('https://api.dropboxapi.com/2/files/list_folder', { path: pasta }).getContentText());
    Logger.log('Arquivos na pasta: ' + (lf.entries || []).map(function (e) { return e.name; }).join(' | '));
  } catch (e) { Logger.log('Não consegui listar a pasta: ' + e.message); }
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
  CacheService.getScriptCache().removeAll([CACHE_EHS_CHAVE + '_n', CACHE_EHS_CHAVE + '_filtros_v2_n']);
  menuColunasEHS();
}

/** Menu: Portal > Conferir colunas do EHS */
function menuColunasEHS() {
  const props = PropertiesService.getScriptProperties();
  CacheService.getScriptCache().removeAll([CACHE_EHS_CHAVE + '_n', CACHE_EHS_CHAVE + '_filtros_v2_n']);
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