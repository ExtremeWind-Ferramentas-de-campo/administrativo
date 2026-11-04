/**
 * ============================================================================
 * ENVIO AUTOMÁTICO DO REPORT DE EHS PARA GRUPOS DO WHATSAPP (Z-API)
 * Cole como novo arquivo (Arquivo > + > Script, nome "ZapEHS") no mesmo
 * projeto do portal. Usa do StatusEHS.gs: lerEHS_, dbxPortal_,
 * ehsPastaPortal_, tipoArquivo_, chaveAutoriz_; do SupervisaoCampo.gs:
 * chaveTexto_, dataISO_, prop_; do Codigo.gs: registrar_, avisar_.
 *
 * COMO FUNCIONA
 *  1. O analista marca "Autorizado = Sim" no card do report (Status RD EHS).
 *     O portal NÃO envia na hora — mandar PDF e fotos para vários grupos, com
 *     pausa entre as mensagens, leva minutos e travaria o botão. Ele coloca o
 *     report na FILA (aba EHS_ZAP_FILA).
 *  2. Um gatilho a cada minuto roda processarFilaZapEHS(): pega a pasta do dia
 *     do report no Dropbox (Caminho_PDF gravado pelo app de campo), gera um
 *     link temporário de cada arquivo e manda, para cada grupo:
 *        texto de identificação -> PDF -> fotos (em ordem de nome)
 *  3. A fila guarda o que já foi para cada grupo. Se a execução acabar no
 *     meio (limite de tempo do Apps Script) ou um envio falhar, a próxima
 *     rodada continua de onde parou, sem repetir o que já foi.
 *  Desmarcar o Autorizado antes do envio terminar cancela o restante.
 *  Marcar Sim de novo num report já enviado NÃO reenvia (use
 *  reenviarZapEHS('aaaa-mm-dd', 'PARQUE') à mão, se precisar).
 *
 * QUAIS GRUPOS RECEBEM — aba ZAP_GRUPOS (criada por instalarZapEHS):
 *   ativo | grupo_id | nome_grupo | cliente | parque
 *   - ativo: SIM para usar a linha.
 *   - grupo_id: o ID do grupo na Z-API (ex.: 558184237784-1508587077).
 *     Rode listarGruposZap() para ver os IDs (sai no Registro de execução e
 *     na aba ZAP_GRUPOS_DISPONIVEIS).
 *   - cliente / parque: em branco = recebe tudo. Preenchido = só reports
 *     daquele cliente (contém, sem acento) / daquele parque (nome igual).
 *
 * PROPRIEDADES DO SCRIPT
 *   ZAPI_INSTANCE_ID, ZAPI_TOKEN          (obrigatórias)
 *   ZAPI_CLIENT_TOKEN                     (opcional: Token de Segurança da Conta)
 *   Dropbox: as mesmas do Status RD EHS
 *   (DROPBOX_APP_KEY, DROPBOX_APP_SECRET, DROPBOX_REFRESH_TOKEN).
 *
 * INSTALAÇÃO: cole, rode instalarZapEHS() uma vez (cria as abas e o gatilho
 * de 1 em 1 minuto), preencha ZAP_GRUPOS e reimplante o App da Web.
 * Diagnóstico: testarZapEHS().
 * ============================================================================
 */

const ABA_ZAP_FILA = 'EHS_ZAP_FILA';
const COLUNAS_ZAP_FILA = ['chave', 'data', 'parque', 'status', 'criado_em', 'autorizado_por',
                          'enviado_em', 'progresso', 'detalhe'];
const ABA_ZAP_GRUPOS = 'ZAP_GRUPOS';
const COLUNAS_ZAP_GRUPOS = ['ativo', 'grupo_id', 'nome_grupo', 'cliente', 'parque'];

const ZAP_PAUSA_MS = 3000;            // entre mensagens: reduz risco de bloqueio do número
const ZAP_ORCAMENTO_MS = 4.5 * 60000; // tempo de envio por rodada (limite do Apps Script: 6 min)
const ZAP_MAX_TENTATIVAS = 3;         // falhas seguidas no mesmo report antes de virar ERRO
const ZAP_ENVIA_TEXTO = true;         // manda o texto de entrada ("Olá, segue evidências...") antes dos arquivos
const ZAP_EXTENSOES = /\.(pdf|jpe?g|png|webp)$/i;


/* ===========================================================================
   INSTALAÇÃO E DIAGNÓSTICO
   =========================================================================== */

function instalarZapEHS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  abaZapFila_(true);
  let g = ss.getSheetByName(ABA_ZAP_GRUPOS);
  if (!g) {
    g = ss.insertSheet(ABA_ZAP_GRUPOS);
    g.getRange(1, 1, 1, COLUNAS_ZAP_GRUPOS.length).setValues([COLUNAS_ZAP_GRUPOS])
     .setFontWeight('bold').setBackground('#14181b').setFontColor('#ecede8');
    g.setFrozenRows(1);
    g.getRange('B:B').setNumberFormat('@');            // ID do grupo não pode virar número
    g.getRange('A1').setNote('SIM = esta linha está valendo.');
    g.getRange('B1').setNote('ID do grupo na Z-API. Rode listarGruposZap() para ver os IDs.');
    g.getRange('D1').setNote('Em branco = recebe todos os clientes.');
    g.getRange('E1').setNote('Em branco = recebe todos os parques.');
  }

  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'processarFilaZapEHS'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('processarFilaZapEHS').timeBased().everyMinutes(1).create();

  avisar_('WhatsApp do EHS instalado.\n\n' +
          '1. Confira as propriedades ZAPI_INSTANCE_ID e ZAPI_TOKEN (e ZAPI_CLIENT_TOKEN, se usar).\n' +
          '2. Rode listarGruposZap() para ver os IDs dos grupos.\n' +
          '3. Preencha a aba ' + ABA_ZAP_GRUPOS + ' (ativo = SIM, grupo_id, cliente/parque opcionais).\n' +
          '4. Reimplante o App da Web.\n\n' +
          'Gatilho criado: processarFilaZapEHS a cada 1 minuto.');
}

/**
 * Lista os grupos da conta na Z-API (Registro de execução + aba ZAP_GRUPOS_DISPONIVEIS).
 *
 * Consulta duas rotas e junta o resultado:
 *  - /groups  só grupos;
 *  - /chats   todas as conversas (filtra as que são grupo).
 * As duas são PAGINADAS: a primeira versão lia só a 1ª página de /chats e,
 * em conta com muitas conversas, o grupo ficava de fora.
 *
 * Número recém-conectado: o WhatsApp leva alguns minutos sincronizando, e
 * grupo sem mensagem recente pode demorar a aparecer. Se faltar, mande uma
 * mensagem qualquer no grupo pelo celular e rode de novo.
 */
function listarGruposZap() {
  const achados = {};   // id -> nome
  const resumo = [];

  const varrer = function (rota) {
    let total = 0, pagina = 1;
    for (; pagina <= 30; pagina++) {
      const url = zapiBase_() + rota + (rota.indexOf('?') > -1 ? '&' : '?') + 'page=' + pagina + '&pageSize=100';
      const r = UrlFetchApp.fetch(url, { method: 'get', headers: zapiHeaders_(), muteHttpExceptions: true });
      if (r.getResponseCode() !== 200) {
        resumo.push(rota + ': HTTP ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
        return;
      }
      let lista;
      try { lista = JSON.parse(r.getContentText()); } catch (e) { lista = null; }
      if (lista && !Array.isArray(lista)) lista = lista.groups || lista.chats || lista.data || null;
      if (!Array.isArray(lista)) {
        resumo.push(rota + ': resposta fora do formato: ' + r.getContentText().slice(0, 200));
        return;
      }
      if (!lista.length) break;
      lista.forEach(function (c) {
        const id = String(c.phone || c.id || c.groupId || '');
        const ehGrupo = c.isGroup === true || /-group$/.test(id) || (rota === '/groups' && id);
        if (!ehGrupo || !id) return;
        if (!achados[id] || !achados[id].nome) achados[id] = { nome: String(c.name || c.subject || ''), rota: rota };
        total++;
      });
      if (lista.length < 100) break;            // última página
    }
    resumo.push(rota + ': ' + total + ' grupo(s) em ' + pagina + ' página(s)');
  };

  varrer('/groups');
  varrer('/chats');

  const grupos = Object.keys(achados).map(function (id) { return [achados[id].nome, id, achados[id].rota]; })
    .sort(function (a, b) { return a[0].localeCompare(b[0], 'pt-BR'); });
  resumo.forEach(function (s) { Logger.log(s); });
  grupos.forEach(function (g) { Logger.log('Grupo: ' + (g[0] || '(sem nome)') + '  ->  ID: ' + g[1]); });
  Logger.log('Total de grupos diferentes: ' + grupos.length);
  if (!grupos.length) {
    Logger.log('Nenhum grupo. Confira se a instância está CONECTADA (testarZapEHS mostra o /status) e, ' +
               'se o número foi conectado agora, mande uma mensagem no grupo pelo celular e rode de novo.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('ZAP_GRUPOS_DISPONIVEIS') || ss.insertSheet('ZAP_GRUPOS_DISPONIVEIS');
  aba.clear();
  aba.getRange(1, 1, 1, 3).setValues([['nome_grupo', 'grupo_id', 'achado_em']]).setFontWeight('bold');
  aba.getRange('B:B').setNumberFormat('@');
  if (grupos.length) aba.getRange(2, 1, grupos.length, 3).setValues(grupos);
  aba.getRange(grupos.length + 3, 1).setValue('Atualizado em ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + ' — ' + resumo.join(' | '));
}

/** Confere propriedades, Dropbox, Z-API e a configuração dos grupos. */
function testarZapEHS() {
  const props = PropertiesService.getScriptProperties();
  ['ZAPI_INSTANCE_ID', 'ZAPI_TOKEN', 'ZAPI_CLIENT_TOKEN'].forEach(function (k) {
    const v = props.getProperty(k);
    Logger.log(k + ': ' + (v ? v.length + ' caracteres' + (v !== v.trim() ? ' (TEM ESPAÇO SOBRANDO)' : '') : 'não configurada'));
  });
  try { dbxTokenPortal_(); Logger.log('Dropbox: credenciais OK'); }
  catch (e) { Logger.log('Dropbox: ' + e.message); }
  try {
    const r = UrlFetchApp.fetch(zapiBase_() + '/status', { method: 'get', headers: zapiHeaders_(), muteHttpExceptions: true });
    Logger.log('Z-API /status: ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
  } catch (e) { Logger.log('Z-API: ' + e.message); }
  const grupos = lerGruposZap_();
  Logger.log('Grupos ativos em ' + ABA_ZAP_GRUPOS + ': ' + grupos.length);
  grupos.forEach(function (g) {
    Logger.log('  ' + (g.nome || g.id) + ' | cliente: ' + (g.cliente || 'todos') + ' | parque: ' + (g.parque || 'todos'));
  });
  const fila = lerFilaZap_();
  Logger.log('Fila: ' + fila.length + ' report(s); pendentes: ' +
             fila.filter(function (f) { return f.status === 'PENDENTE' || f.status === 'ENVIANDO'; }).length);
}

/* Para reenviar pelo botão Executar do editor: preencha as duas linhas abaixo,
   salve, escolha "reenviarZapEHS" na lista e clique em Executar. */
const ZAP_REENVIO_DATA = '29/10/2026';    // dd/mm/aaaa
const ZAP_REENVIO_PARQUE = 'TUCANO 1';    // igual ao nome do parque no report

/** Manda de novo um report já enviado (todos os grupos). Sem argumentos usa as constantes acima. */
function reenviarZapEHS(dataIso, parque) {
  if (typeof dataIso !== 'string' || !dataIso) dataIso = ZAP_REENVIO_DATA;
  if (typeof parque !== 'string' || !parque) parque = ZAP_REENVIO_PARQUE;
  const br = String(dataIso).trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (br) dataIso = br[3] + '-' + br[2] + '-' + br[1];
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dataIso) || !String(parque).trim()) {
    throw new Error('Preencha ZAP_REENVIO_DATA (dd/mm/aaaa) e ZAP_REENVIO_PARQUE no topo do reenviarZapEHS.');
  }
  parque = String(parque).trim();
  enfileirarZapEHS_(dataIso, parque, 'SIM', '', true);
  processarFilaZapEHS();   // envia já, sem esperar o gatilho
  const fim = lerFilaZap_().filter(function (f) { return f.chave === chaveAutoriz_(dataIso, parque); })[0];
  Logger.log('Reenvio ' + dataIso + ' | ' + parque + ': ' + (fim ? fim.status + (fim.detalhe ? ' — ' + fim.detalhe : '') : '?'));
}


/* ===========================================================================
   FILA (aba EHS_ZAP_FILA)
   =========================================================================== */

function abaZapFila_(criar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(ABA_ZAP_FILA);
  if (!aba && criar) {
    aba = ss.insertSheet(ABA_ZAP_FILA);
    aba.getRange(1, 1, 1, COLUNAS_ZAP_FILA.length).setValues([COLUNAS_ZAP_FILA])
       .setFontWeight('bold').setBackground('#14181b').setFontColor('#ecede8');
    aba.setFrozenRows(1);
    aba.getRange('A:A').setNumberFormat('@');
    aba.getRange('B:B').setNumberFormat('dd/mm/yyyy');
    aba.getRange('F:F').setNumberFormat('@');
    aba.getRange('H1').setNote('Controle interno do envio (o que já foi para cada grupo). Não editar.');
  }
  return aba;
}

function lerFilaZap_() {
  const aba = abaZapFila_(false);
  if (!aba || aba.getLastRow() < 2) return [];
  const n = aba.getLastRow() - 1;
  const v = aba.getRange(2, 1, n, COLUNAS_ZAP_FILA.length).getValues();
  return v.map(function (l, i) {
    let prog = {};
    try { prog = JSON.parse(l[7] || '{}') || {}; } catch (e) { prog = {}; }
    return {
      linha: i + 2, chave: String(l[0]), data: dataISO_(l[1]) || String(l[0]).split('|')[0],
      parque: String(l[2]), status: String(l[3]).toUpperCase(), enviadoEm: l[6],
      progresso: { feitos: prog.feitos || [], falhas: prog.falhas || 0 }, detalhe: String(l[8] || '')
    };
  });
}

function gravarItemFila_(aba, item, mudancas) {
  if (mudancas.status !== undefined) aba.getRange(item.linha, 4).setValue(mudancas.status);
  if (mudancas.enviadoEm !== undefined) aba.getRange(item.linha, 7).setValue(mudancas.enviadoEm);
  if (mudancas.progresso !== undefined) aba.getRange(item.linha, 8).setValue(JSON.stringify(mudancas.progresso));
  if (mudancas.detalhe !== undefined) aba.getRange(item.linha, 9).setValue(String(mudancas.detalhe).slice(0, 500));
}

/**
 * Chamado pelo acaoEhsAutorizar_ (StatusEHS.gs) ao marcar/desmarcar.
 *   SIM            -> entra na fila (se ainda não foi enviado)
 *   NAO ou vazio   -> cancela o que ainda não foi
 * forcar = reenviar mesmo se já enviado (reenviarZapEHS).
 * -> 'NA_FILA' | 'JA_ENVIADO' | 'CANCELADO' | ''
 */
function enfileirarZapEHS_(dataIso, parque, valor, matricula, forcar) {
  const aba = abaZapFila_(true);
  const chave = chaveAutoriz_(dataIso, parque);
  const item = lerFilaZap_().filter(function (f) { return f.chave === chave; })[0];

  if (valor !== 'SIM') {
    if (item && (item.status === 'PENDENTE' || item.status === 'ENVIANDO' || item.status === 'ERRO')) {
      gravarItemFila_(aba, item, { status: 'CANCELADO', detalhe: 'Autorizado desmarcado antes do fim do envio' });
      return 'CANCELADO';
    }
    return '';
  }

  if (item) {
    if (item.status === 'ENVIADO' && !forcar) return 'JA_ENVIADO';
    gravarItemFila_(aba, item, {
      status: 'PENDENTE', detalhe: '',
      progresso: forcar ? { feitos: [], falhas: 0 } : { feitos: item.progresso.feitos, falhas: 0 }
    });
  } else {
    const dia = Utilities.parseDate(dataIso, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    aba.appendRow([chave, dia, parque, 'PENDENTE', new Date(), String(matricula || ''), '',
                   JSON.stringify({ feitos: [], falhas: 0 }), '']);
  }
  PropertiesService.getScriptProperties().setProperty('ZAP_EHS_TEM_FILA', '1');
  return 'NA_FILA';
}

/** Situação do envio por report, para o card: { 'aaaa-mm-dd|parque': status } */
function situacaoZapEHS_() {
  const mapa = {};
  try {
    lerFilaZap_().forEach(function (f) { mapa[f.chave] = { status: f.status, detalhe: f.detalhe }; });
  } catch (e) { /* sem fila */ }
  return mapa;
}


/* ===========================================================================
   GRUPOS (aba ZAP_GRUPOS)
   =========================================================================== */

function lerGruposZap_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ZAP_GRUPOS);
  if (!aba || aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, COLUNAS_ZAP_GRUPOS.length).getDisplayValues()
    .filter(function (l) { return /^(sim|s|x|1|true)$/i.test(String(l[0]).trim()) && String(l[1]).trim(); })
    .map(function (l) {
      return { id: String(l[1]).trim().replace(/^'/, ''), nome: String(l[2]).trim(),
               cliente: chaveTexto_(l[3]), parque: chaveTexto_(l[4]) };
    });
}

/* Grupos que recebem o report deste cliente/parque. */
function gruposDoReport_(cliente, parque) {
  const c = chaveTexto_(cliente), p = chaveTexto_(parque);
  const vistos = {};
  return lerGruposZap_().filter(function (g) {
    if (g.parque && g.parque !== p) return false;
    if (g.cliente && (!c || (c.indexOf(g.cliente) === -1 && g.cliente.indexOf(c) === -1))) return false;
    if (vistos[g.id]) return false;          // mesmo grupo em duas linhas: uma vez só
    vistos[g.id] = true;
    return true;
  });
}


/* ===========================================================================
   PROCESSAMENTO (gatilho de 1 em 1 minuto)
   =========================================================================== */

function processarFilaZapEHS() {
  const props = PropertiesService.getScriptProperties();
  // Sem nada na fila, o gatilho sai sem abrir a planilha.
  if (!props.getProperty('ZAP_EHS_TEM_FILA')) return;

  // Trava PRÓPRIA (do documento), não a do script: o envio leva minutos e a
  // trava do script seguraria os cliques do portal (salvar projeto, marcar
  // autorizado...) esse tempo todo.
  const trava = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!trava.tryLock(1000)) return;           // outra rodada ainda está enviando
  try {
    const inicio = Date.now();
    const aba = abaZapFila_(true);
    let pendentes = lerFilaZap_().filter(function (f) { return f.status === 'PENDENTE' || f.status === 'ENVIANDO'; });
    if (!pendentes.length) { props.deleteProperty('ZAP_EHS_TEM_FILA'); return; }

    let reports = [];
    try { reports = lerEHS_(); }
    catch (e) { Logger.log('Planilha do EHS indisponível: ' + e.message); return; }

    for (let i = 0; i < pendentes.length; i++) {
      if (Date.now() - inicio > ZAP_ORCAMENTO_MS) break;
      // relê a linha: o analista pode ter desmarcado enquanto o anterior enviava
      const atual = lerFilaZap_().filter(function (f) { return f.chave === pendentes[i].chave; })[0];
      if (!atual || (atual.status !== 'PENDENTE' && atual.status !== 'ENVIANDO')) continue;
      enviarItemZap_(aba, atual, reports, inicio);
    }

    const sobrou = lerFilaZap_().some(function (f) { return f.status === 'PENDENTE' || f.status === 'ENVIANDO'; });
    if (!sobrou) props.deleteProperty('ZAP_EHS_TEM_FILA');
  } finally {
    trava.releaseLock();
  }
}

function enviarItemZap_(aba, item, reports, inicio) {
  const prog = item.progresso;
  const falhar = function (motivo, definitivo) {
    prog.falhas = (prog.falhas || 0) + 1;
    const fim = definitivo || prog.falhas >= ZAP_MAX_TENTATIVAS;
    gravarItemFila_(aba, item, { status: fim ? 'ERRO' : 'PENDENTE', progresso: prog, detalhe: motivo });
    registrar_('', 'EHS_ZAP', fim ? 'ERRO' : 'TENTAR_DE_NOVO', item.chave + ' | ' + motivo);
  };

  // o report mais recente do parque/dia (reenvio substitui a linha)
  const doDia = reports.filter(function (r) {
    return r.data === item.data && chaveTexto_(r.parque) === chaveTexto_(item.parque) && r.caminho;
  });
  const rep = doDia[doDia.length - 1];
  if (!rep) return falhar('Report sem caminho do PDF na planilha (Caminho_PDF).', false);

  const grupos = gruposDoReport_(rep.cliente, rep.parque);
  if (!grupos.length) return falhar('Nenhum grupo ativo em ' + ABA_ZAP_GRUPOS + ' para ' + (rep.cliente || 'este cliente') + ' / ' + rep.parque + '.', true);

  let arquivos;
  try { arquivos = arquivosDoReportZap_(rep.caminho); }
  catch (e) { return falhar('Dropbox: ' + e.message, false); }
  if (!arquivos.length) return falhar('Nenhum PDF/foto na pasta do report.', false);

  gravarItemFila_(aba, item, { status: 'ENVIANDO' });
  // a equipe é identificada pelo parque (o report não tem outro nome de equipe)
  const texto = 'Olá, segue evidências de EHS do dia ' + item.data.split('-').reverse().join('/') +
    ' da equipe ' + rep.parque + '. Bom trabalho a todos, com segurança.';
  const links = {};

  for (let g = 0; g < grupos.length; g++) {
    const grupo = grupos[g];
    const etapas = (ZAP_ENVIA_TEXTO ? [{ texto: true, name: '#texto' }] : []).concat(arquivos);
    for (let a = 0; a < etapas.length; a++) {
      const passo = etapas[a];
      const marca = grupo.id + '|' + passo.name;
      if (prog.feitos.indexOf(marca) > -1) continue;
      if (Date.now() - inicio > ZAP_ORCAMENTO_MS) {
        gravarItemFila_(aba, item, { status: 'PENDENTE', progresso: prog, detalhe: 'Continua na próxima rodada' });
        return;
      }
      let ok;
      try {
        if (passo.texto) {
          ok = zapEnviarTexto_(grupo.id, texto);
        } else {
          // link temporário (4 h) gerado uma vez por arquivo e usado em todos os grupos
          if (!links[passo.path]) links[passo.path] = linkTemporarioDbx_(passo.path);
          ok = /\.pdf$/i.test(passo.name)
            ? zapEnviarPdf_(grupo.id, links[passo.path], passo.name)
            : zapEnviarImagem_(grupo.id, links[passo.path]);
        }
      } catch (e) { ok = { ok: false, erro: e.message }; }

      if (!ok.ok) return falhar((grupo.nome || grupo.id) + ' / ' + passo.name + ': ' + ok.erro, false);
      prog.feitos.push(marca);
      gravarItemFila_(aba, item, { progresso: prog });   // grava a cada envio: se cair, não repete
      Utilities.sleep(ZAP_PAUSA_MS);
    }
  }

  prog.falhas = 0;
  gravarItemFila_(aba, item, {
    status: 'ENVIADO', enviadoEm: new Date(), progresso: prog,
    detalhe: arquivos.length + ' arquivo(s) para ' + grupos.length + ' grupo(s)'
  });
  registrar_('', 'EHS_ZAP', 'OK', item.chave + ' | ' + arquivos.length + ' arquivo(s) x ' + grupos.length + ' grupo(s)');
}

/* Arquivos da pasta do dia do report: PDF primeiro, depois as fotos em ordem
   de nome. Mesma regra do botão "Enviar no WhatsApp": o caminho vem da
   planilha e precisa estar dentro de EHS_PASTA; report antigo (sem pasta do
   dia, direto na pasta da semana) manda só o PDF. */
function arquivosDoReportZap_(caminhoPdf) {
  const raiz = ehsPastaPortal_().toLowerCase();
  const caminho = String(caminhoPdf).replace(/\\/g, '/');
  if (caminho.toLowerCase().indexOf(raiz + '/') !== 0 || caminho.indexOf('..') > -1) {
    throw new Error('PDF fora da pasta do EHS: ' + caminho);
  }
  const pasta = caminho.slice(0, caminho.lastIndexOf('/'));
  let lista;
  if (/\/\d{2}-\d{2}-\d{4}$/.test(pasta)) {
    lista = [];
    let r = JSON.parse(dbxPortal_('https://api.dropboxapi.com/2/files/list_folder', { path: pasta, recursive: false }).getContentText());
    lista = lista.concat(r.entries || []);
    while (r.has_more) {
      r = JSON.parse(dbxPortal_('https://api.dropboxapi.com/2/files/list_folder/continue', { cursor: r.cursor }).getContentText());
      lista = lista.concat(r.entries || []);
    }
    lista = lista.filter(function (e) { return e['.tag'] === 'file' && ZAP_EXTENSOES.test(e.name); })
      .map(function (e) { return { name: e.name, path: e.path_display || e.path_lower }; });
  } else {
    lista = [{ name: caminho.slice(caminho.lastIndexOf('/') + 1), path: caminho }];
  }
  lista.sort(function (a, b) {
    const pa = /\.pdf$/i.test(a.name) ? 0 : 1, pb = /\.pdf$/i.test(b.name) ? 0 : 1;
    return (pa - pb) || a.name.localeCompare(b.name, 'pt-BR', { numeric: true });
  });
  return lista;
}

function linkTemporarioDbx_(path) {
  const r = JSON.parse(dbxPortal_('https://api.dropboxapi.com/2/files/get_temporary_link', { path: path }).getContentText());
  if (!r.link) throw new Error('Dropbox não gerou link para ' + path);
  return r.link;
}


/* ===========================================================================
   Z-API
   =========================================================================== */

function zapiBase_() {
  const props = PropertiesService.getScriptProperties();
  const inst = String(props.getProperty('ZAPI_INSTANCE_ID') || '').trim();
  const tok = String(props.getProperty('ZAPI_TOKEN') || '').trim();
  if (!inst || !tok) throw new Error('Faltam as propriedades ZAPI_INSTANCE_ID / ZAPI_TOKEN.');
  return 'https://api.z-api.io/instances/' + inst + '/token/' + tok;
}

function zapiHeaders_() {
  const h = { 'Content-Type': 'application/json' };
  const ct = String(PropertiesService.getScriptProperties().getProperty('ZAPI_CLIENT_TOKEN') || '').trim();
  if (ct) h['Client-Token'] = ct;
  return h;
}

function zapPost_(rota, corpo) {
  const r = UrlFetchApp.fetch(zapiBase_() + rota, {
    method: 'post', headers: zapiHeaders_(), payload: JSON.stringify(corpo), muteHttpExceptions: true
  });
  const st = r.getResponseCode();
  if (st === 200 || st === 201) return { ok: true };
  return { ok: false, erro: 'Z-API ' + st + ' ' + r.getContentText().slice(0, 200) };
}

function zapEnviarTexto_(grupo, texto)        { return zapPost_('/send-text', { phone: grupo, message: texto }); }
// sem o campo caption: foto sem legenda, e o WhatsApp agrupa as fotos seguidas num bloco só
function zapEnviarImagem_(grupo, url)         { return zapPost_('/send-image', { phone: grupo, image: url }); }
function zapEnviarPdf_(grupo, url, nome)      { return zapPost_('/send-document/pdf', { phone: grupo, document: url, fileName: nome.replace(/\.pdf$/i, '') }); }
