/**
 * ============================================================================
 * DOWNLOAD AUTOMÁTICO DAS FOTOS DOS GRUPOS DO WHATSAPP (Z-API) PARA O DROPBOX
 * Cole como novo arquivo (Arquivo > + > Script, nome "ZapDownload") no mesmo
 * projeto do portal. É o caminho inverso do ZapEHS.gs e usa as mesmas
 * credenciais. Usa do ZapEHS.gs: zapiBase_, zapiHeaders_; do StatusEHS.gs:
 * dbxPortal_, dbxTokenPortal_; do SupervisaoCampo.gs: prop_; do Codigo.gs:
 * registrar_, avisar_.
 *
 * COMO FUNCIONA
 *  1. A Z-API chama o App da Web (webhook "Ao receber") a cada mensagem que
 *     chega no número. O doPost do Codigo.gs reconhece a chamada pelo
 *     ?zapdl=<segredo> da URL e passa para webhookZapDownload_().
 *  2. webhookZapDownload_ só GUARDA a foto numa propriedade do script (uma
 *     chave por mensagem) e responde na hora. Não escreve na planilha: com
 *     várias fotos chegando juntas, gravações simultâneas na aba se
 *     atropelavam e fotos se perdiam. Baixar e subir no Dropbox leva segundos
 *     por foto; feito ali, a Z-API esperaria e poderia reenviar.
 *     Ignora: o que não é grupo, o que o próprio número mandou (fromMe: são
 *     os envios do report de EHS), o que não é foto e grupo fora da aba
 *     ZAP_DOWN_GRUPOS.
 *  3. Um gatilho a cada minuto roda processarDownloadsZap(): passa o que o
 *     webhook guardou para a aba ZAP_DOWN_LOG (status PENDENTE), baixa cada
 *     foto pendente da URL da Z-API (a mídia fica lá ~30 dias) e grava em
 *        <ZAP_DOWN_PASTA>/<CLIENTE>/<PARQUE>/W<semana>/<dd-mm-aaaa>/
 *          aaaa-mm-dd_hhmmss_<telefone do autor>_<messageId>.jpg
 *     Semana ISO (segunda a domingo) e dia pela data em que a foto foi
 *     mandada no grupo, no fuso do script — mesmo padrão das pastas do EHS.
 *  4. Mesmo messageId duas vezes (reenvio do webhook) baixa uma vez só.
 *     Falhou 3 vezes seguidas -> ERRO. reprocessarErrosZapDownload() devolve
 *     para a fila os erros com menos de 30 dias.
 *
 * QUAIS GRUPOS SÃO BAIXADOS — aba ZAP_DOWN_GRUPOS (criada por instalarZapDownload):
 *   ativo | grupo_id | nome_grupo | cliente | parque
 *   - ativo: SIM para usar a linha.
 *   - grupo_id: o ID do grupo na Z-API. Os IDs estão na aba
 *     ZAP_GRUPOS_DISPONIVEIS (menu Portal > WhatsApp EHS: listar grupos).
 *   - cliente / parque: OBRIGATÓRIOS — são os nomes das pastas.
 *   São grupos DIFERENTES dos da aba ZAP_GRUPOS (envio). Mudanças na aba
 *   valem em até 5 minutos (a lista fica em cache para o webhook ser rápido).
 *
 * PROPRIEDADES DO SCRIPT
 *   As mesmas do ZapEHS (ZAPI_*) e do Dropbox (DROPBOX_*).
 *   ZAP_DOWN_SEGREDO   criada pelo instalarZapDownload — não apague.
 *   ZAP_DOWN_PASTA     opcional; padrão: ZAP_DOWN_PASTA_PADRAO abaixo.
 *   ZAP_DOWN_URL_APP   opcional; URL /exec do App da Web, no lugar da
 *                      constante ZAP_DOWN_URL_EXEC.
 *
 * INSTALAÇÃO
 *  1. Cole este arquivo e salve. Rode instalarZapDownload() (cria as abas,
 *     o segredo e o gatilho; liga o registro bruto do webhook).
 *  2. Preencha a aba ZAP_DOWN_GRUPOS.
 *  3. Reimplante o App da Web (Implantar > Gerenciar implantações > editar >
 *     Nova versão). Sem isso o doPost novo não vale.
 *  4. Rode configurarWebhookZapDownload() (aponta o "Ao receber" da Z-API
 *     para o App da Web). Ele NÃO substitui um webhook de outra automação
 *     sem você mandar — veja ZAP_DOWN_FORCAR_WEBHOOK.
 *  5. Mande uma foto num grupo cadastrado e confira as abas ZAP_WEBHOOK_RAW
 *     (o JSON que chegou) e ZAP_DOWN_LOG. Diagnóstico: testarZapDownload().
 * ============================================================================
 */

const ABA_ZAP_DOWN_GRUPOS = 'ZAP_DOWN_GRUPOS';
const COLUNAS_ZAP_DOWN_GRUPOS = ['ativo', 'grupo_id', 'nome_grupo', 'cliente', 'parque'];
const ABA_ZAP_DOWN_LOG = 'ZAP_DOWN_LOG';
const COLUNAS_ZAP_DOWN_LOG = ['recebido_em', 'message_id', 'grupo_id', 'nome_grupo', 'cliente', 'parque',
                              'autor', 'legenda', 'url_midia', 'tipo_midia', 'enviado_em', 'status', 'tentativas',
                              'caminho_dropbox', 'detalhe'];
const ABA_ZAP_RAW = 'ZAP_WEBHOOK_RAW';

const ZAP_DOWN_PASTA_PADRAO = '/02 - EXTREME WIND/7 - ENG & CQ/12 - FOTOS WHATS APP';
const ZAP_DOWN_ORCAMENTO_MS = 4.5 * 60000; // tempo por rodada (limite do Apps Script: 6 min)
const ZAP_DOWN_MAX_TENTATIVAS = 3;
const ZAP_DOWN_DIAS_REPROCESSAR = 30;     // a Z-API guarda a mídia ~30 dias
const ZAP_DOWN_RAW_MAX = 200;             // o registro bruto se desliga sozinho depois disso

/* URL /exec do App da Web (a mesma do API_URL em config/config.js do portal).
   Rodando pelo editor, ScriptApp.getService().getUrl() devolve a /dev, que só
   abre para quem está logado no editor — a Z-API não consegue usar.
   A propriedade ZAP_DOWN_URL_APP, se existir, vale no lugar desta. */
const ZAP_DOWN_URL_EXEC = 'https://script.google.com/macros/s/AKfycbyIT6qWi8CXQog8eQ6EHg65_CDVc1JPevbCGUj9wziNuhHDhyWXmOXvk-9shPrNwuxe/exec';

/* Só mude para true se você CONFERIU no painel da Z-API que o webhook
   "Ao receber" atual pode ser substituído (nenhuma outra automação usa). */
const ZAP_DOWN_FORCAR_WEBHOOK = false;


/* ===========================================================================
   INSTALAÇÃO E DIAGNÓSTICO
   =========================================================================== */

function instalarZapDownload() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  let g = ss.getSheetByName(ABA_ZAP_DOWN_GRUPOS);
  if (!g) {
    g = ss.insertSheet(ABA_ZAP_DOWN_GRUPOS);
    g.getRange(1, 1, 1, COLUNAS_ZAP_DOWN_GRUPOS.length).setValues([COLUNAS_ZAP_DOWN_GRUPOS])
     .setFontWeight('bold').setBackground('#14181b').setFontColor('#ecede8');
    g.setFrozenRows(1);
    g.getRange('B:B').setNumberFormat('@');            // ID do grupo não pode virar número
    g.getRange('A1').setNote('SIM = baixar as fotos deste grupo.');
    g.getRange('B1').setNote('ID do grupo na Z-API (aba ZAP_GRUPOS_DISPONIVEIS, coluna grupo_id).');
    g.getRange('D1').setNote('Obrigatório: nome da pasta do cliente.');
    g.getRange('E1').setNote('Obrigatório: nome da pasta do parque.');
    const disp = ss.getSheetByName('ZAP_GRUPOS_DISPONIVEIS');
    if (disp) {
      // lista os IDs conhecidos, mas aceita ID digitado à mão
      g.getRange('B2:B500').setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInRange(disp.getRange('B2:B'), true).setAllowInvalid(true).build());
    }
  }
  abaZapDownLog_(true);

  if (!props.getProperty('ZAP_DOWN_SEGREDO')) props.setProperty('ZAP_DOWN_SEGREDO', Utilities.getUuid().replace(/-/g, ''));
  props.setProperty('ZAP_DOWN_RAW', '1');
  props.setProperty('ZAP_DOWN_RAW_QTD', '0');

  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'processarDownloadsZap'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('processarDownloadsZap').timeBased().everyMinutes(1).create();

  avisar_('Download de fotos do WhatsApp instalado.\n\n' +
          '1. Preencha a aba ' + ABA_ZAP_DOWN_GRUPOS + ' (ativo = SIM, grupo_id, cliente, parque).\n' +
          '2. Reimplante o App da Web (Gerenciar implantações > editar > Nova versão).\n' +
          '3. Rode configurarWebhookZapDownload().\n' +
          '4. Mande uma foto num grupo cadastrado e confira as abas ' + ABA_ZAP_RAW + ' e ' + ABA_ZAP_DOWN_LOG + '.\n\n' +
          'Gatilho criado: processarDownloadsZap a cada 1 minuto.\n' +
          'Registro bruto do webhook LIGADO (desliga sozinho após ' + ZAP_DOWN_RAW_MAX + ' mensagens).');
}

/**
 * Aponta o webhook "Ao receber" da Z-API para este App da Web.
 * Antes, lê o webhook atual: se já houver outro endereço, não troca (a
 * instância só tem UM webhook de recebimento) — a não ser que
 * ZAP_DOWN_FORCAR_WEBHOOK seja true.
 */
function configurarWebhookZapDownload() {
  const props = PropertiesService.getScriptProperties();
  const segredo = props.getProperty('ZAP_DOWN_SEGREDO');
  if (!segredo) throw new Error('Rode instalarZapDownload() primeiro.');
  const base = urlAppZapDown_();
  if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(base)) {
    throw new Error('A URL do App da Web precisa terminar em /exec: "' + base + '". ' +
                    'Confira ZAP_DOWN_URL_EXEC no topo do ZapDownload (igual ao API_URL do config.js do portal).');
  }
  const nova = base + '?zapdl=' + segredo;

  const atual = webhookRecebimentoAtual_();
  Logger.log('Webhook "Ao receber" atual: ' + (atual === null ? '(não consegui ler)' : (atual || '(vazio)')));
  const ehNosso = atual && atual.indexOf('zapdl=') > -1;
  if (!ZAP_DOWN_FORCAR_WEBHOOK && !ehNosso && atual !== '') {
    throw new Error(atual === null
      ? 'Não consegui ler o webhook atual na Z-API. Confira no painel se o "Ao receber" está vazio e, se puder substituir, ' +
        'mude ZAP_DOWN_FORCAR_WEBHOOK para true e rode de novo.'
      : 'O "Ao receber" já aponta para ' + atual + '. Se ninguém mais usa, mude ZAP_DOWN_FORCAR_WEBHOOK para true e rode de novo.');
  }

  const r = UrlFetchApp.fetch(zapiBase_() + '/update-webhook-received', {
    method: 'put', headers: zapiHeaders_(), payload: JSON.stringify({ value: nova }), muteHttpExceptions: true
  });
  Logger.log('Z-API update-webhook-received: ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
  if (r.getResponseCode() !== 200) throw new Error('A Z-API recusou o webhook (veja o registro acima).');
  Logger.log('Pronto. Webhook apontado para ' + base + '?zapdl=***');
}

function urlAppZapDown_() {
  return String(PropertiesService.getScriptProperties().getProperty('ZAP_DOWN_URL_APP') || ZAP_DOWN_URL_EXEC || '').trim();
}

/* URL atual do webhook de recebimento: '' = vazio, null = não deu para ler.
   O /me não tem o nome do campo documentado com certeza: procura qualquer
   chave de callback/webhook de "received" (sem pegar o de status/entrega). */
function webhookRecebimentoAtual_() {
  try {
    const r = UrlFetchApp.fetch(zapiBase_() + '/me', { method: 'get', headers: zapiHeaders_(), muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) {
      Logger.log('Z-API /me: HTTP ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
      return null;
    }
    const j = JSON.parse(r.getContentText());
    const chaves = Object.keys(j || {}).filter(function (k) {
      return /received/i.test(k) && /callback|webhook/i.test(k) && !/delivery|status/i.test(k);
    });
    if (!chaves.length) {
      Logger.log('Z-API /me sem campo de webhook de recebimento. Campos: ' + Object.keys(j || {}).join(', '));
      return null;
    }
    return String(j[chaves[0]] || '').trim();
  } catch (e) {
    Logger.log('Z-API /me: ' + e.message);
    return null;
  }
}

/** Confere propriedades, Dropbox, webhook e a configuração dos grupos. */
function testarZapDownload() {
  const props = PropertiesService.getScriptProperties();
  Logger.log('ZAP_DOWN_SEGREDO: ' + (props.getProperty('ZAP_DOWN_SEGREDO') ? 'ok' : 'FALTA — rode instalarZapDownload()'));
  Logger.log('Registro bruto: ' + (props.getProperty('ZAP_DOWN_RAW') === '1' ? 'ligado (' + (props.getProperty('ZAP_DOWN_RAW_QTD') || 0) + ' gravados)' : 'desligado'));
  Logger.log('Pasta base: ' + pastaZapDown_());
  try {
    dbxPortal_('https://api.dropboxapi.com/2/files/get_metadata', { path: pastaZapDown_() });
    Logger.log('Dropbox: pasta base OK');
  } catch (e) { Logger.log('Dropbox: ' + e.message); }
  Logger.log('Webhook "Ao receber" na Z-API: ' + (function (w) { return w === null ? '(não consegui ler)' : (w || '(vazio)'); })(webhookRecebimentoAtual_()));
  Logger.log('URL do App da Web usada no webhook: ' + (urlAppZapDown_() || '(não configurada)'));

  const grupos = lerGruposZapDown_();
  Logger.log('Grupos ativos em ' + ABA_ZAP_DOWN_GRUPOS + ': ' + Object.keys(grupos).length);
  Object.keys(grupos).forEach(function (id) {
    const g = grupos[id];
    Logger.log('  ' + (g.nome || g.id) + ' -> ' + (g.cliente && g.parque ? g.cliente + '/' + g.parque : 'FALTA cliente/parque, linha ignorada'));
  });
  const log = lerLogZapDown_();
  const conta = {};
  log.forEach(function (l) { conta[l.status] = (conta[l.status] || 0) + 1; });
  Logger.log('Log: ' + JSON.stringify(conta));
  const aguardando = Object.keys(props.getProperties()).filter(function (k) { return k.indexOf('zapdl_q_') === 0; }).length;
  Logger.log('Fotos recebidas aguardando o gatilho passar para a aba: ' + aguardando);
}

/** Liga/desliga à mão o registro bruto do webhook (aba ZAP_WEBHOOK_RAW). */
function ligarRegistroBrutoZap()    { PropertiesService.getScriptProperties().setProperties({ ZAP_DOWN_RAW: '1', ZAP_DOWN_RAW_QTD: '0' }); }
function desligarRegistroBrutoZap() { PropertiesService.getScriptProperties().deleteProperty('ZAP_DOWN_RAW'); }

/** Devolve para a fila as fotos com ERRO recebidas nos últimos 30 dias e processa já. */
function reprocessarErrosZapDownload() {
  const aba = abaZapDownLog_(true);
  const limite = Date.now() - ZAP_DOWN_DIAS_REPROCESSAR * 86400000;
  let n = 0;
  lerLogZapDown_().forEach(function (l) {
    if (l.status !== 'ERRO') return;
    if (!(l.recebidoEm instanceof Date) || l.recebidoEm.getTime() < limite) return;
    gravarLinhaZapDown_(aba, l.linha, { status: 'PENDENTE', tentativas: 0, detalhe: 'Reprocessar' });
    n++;
  });
  Logger.log(n + ' foto(s) devolvida(s) para a fila.');
  if (n) {
    PropertiesService.getScriptProperties().setProperty('ZAP_DOWN_TEM_FILA', '1');
    processarDownloadsZap();
  }
}


/* ===========================================================================
   WEBHOOK (chamado pelo doPost do Codigo.gs)
   =========================================================================== */

function webhookZapDownload_(e) {
  const responder = function (o) {
    return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  };
  const props = PropertiesService.getScriptProperties();
  if (String(e.parameter.zapdl) !== String(props.getProperty('ZAP_DOWN_SEGREDO') || '')) {
    return responder({ ok: false });
  }
  /* NÃO escreve na planilha. Várias fotos de uma vez chegam em chamadas
     simultâneas, e appendRow simultâneo em execuções diferentes cai na mesma
     linha (uma foto apaga a outra: 13 enviadas, 10 anotadas). Cada mensagem
     vira uma propriedade com chave própria; o gatilho, que roda sozinho,
     passa tudo para as abas (descarregarWebhookZap_). */
  try {
    const bruto = (e.postData && e.postData.contents) || '';
    const unico = Date.now() + '_' + Math.random().toString(36).slice(2, 8);
    if (props.getProperty('ZAP_DOWN_RAW') === '1') {
      try { props.setProperty('zapdl_raw_' + unico, JSON.stringify({ t: Date.now(), b: bruto.slice(0, 8000) })); }
      catch (e2) { /* registro bruto nunca derruba o webhook */ }
    }

    const p = JSON.parse(bruto);
    const c = classificarZap_(p, gruposZapDownCache_());
    if (c.ignorado) return responder({ ok: true, ignorado: c.ignorado });
    const img = c.img, url = c.url, grupo = c.grupo;

    const mid = String(p.messageId || p.id || '');
    const cache = CacheService.getScriptCache();
    if (mid && cache.get('zapdl_' + mid)) return responder({ ok: true, ignorado: 'duplicado' });

    const momento = Number(p.momment || p.moment || p.timestamp || 0);
    const enviado = momento ? (momento < 1e12 ? momento * 1000 : momento) : Date.now();
    // mesma ordem de COLUNAS_ZAP_DOWN_LOG; datas em milissegundos
    props.setProperty('zapdl_q_' + (mid.replace(/[^\w-]/g, '') || unico), JSON.stringify([
      Date.now(), mid, grupo.id, p.chatName || grupo.nome, grupo.cliente, grupo.parque,
      String(p.participantPhone || p.senderPhone || ''), String((img.caption || '')).slice(0, 500),
      url, String(img.mimeType || ''), enviado, 'PENDENTE', 0, '', ''
    ]));
    if (mid) cache.put('zapdl_' + mid, '1', 21600);
    return responder({ ok: true });
  } catch (erro) {
    registrar_('', 'ZAP_DOWN', 'ERRO', 'webhook: ' + String(erro));
    return responder({ ok: true });   // 200 sempre: erro nosso não deve virar reenvio infinito
  }
}

/* Regras do webhook, separadas para o diagnóstico usar as mesmas.
   -> { ignorado: motivo } ou { img, url, grupo } */
function classificarZap_(p, grupos) {
  if (p.isGroup !== true) return { ignorado: 'nao_grupo' };
  if (p.fromMe === true) return { ignorado: 'enviada_pelo_proprio_numero' };
  const img = p.image;
  const url = img && (img.imageUrl || img.url);
  if (!url) return { ignorado: 'nao_imagem' };
  const grupo = grupos[idGrupoZap_(p.phone || p.chatId)];
  if (!grupo) return { ignorado: 'grupo_fora_da_lista' };
  return { img: img, url: url, grupo: grupo };
}

/**
 * Relê a aba ZAP_WEBHOOK_RAW e mostra no Registro de execução, para cada
 * mensagem, se ela entraria na fila ou por que foi ignorada. Também conta os
 * erros do webhook gravados na aba LOG.
 */
function diagnosticarRawZap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_ZAP_RAW);
  if (!aba || aba.getLastRow() < 2) { Logger.log('Aba ' + ABA_ZAP_RAW + ' vazia.'); return; }
  const grupos = lerGruposZapDown_();
  const noLog = {};
  lerLogZapDown_().forEach(function (l) { noLog[l.messageId] = l.status; });
  const conta = {}, mids = {};

  aba.getRange(2, 1, aba.getLastRow() - 1, 2).getValues().forEach(function (l, i) {
    let p, motivo;
    try { p = JSON.parse(l[1]); } catch (e) { p = null; }
    if (!p) motivo = 'json_invalido_ou_cortado';
    else motivo = classificarZap_(p, grupos).ignorado || 'ENTRARIA_NA_FILA';
    const mid = p ? String(p.messageId || p.id || '') : '';
    if (motivo === 'ENTRARIA_NA_FILA' && mids[mid]) motivo = 'duplicado_do_webhook';
    mids[mid] = true;
    conta[motivo] = (conta[motivo] || 0) + 1;
    if (!p || p.isGroup === true) {
      Logger.log('Linha ' + (i + 2) + ' | ' + motivo + (mid && noLog[mid] ? ' | no ZAP_DOWN_LOG: ' + noLog[mid] : '') +
                 ' | tipo: ' + (p ? (p.type || '') + ' | campos: ' + Object.keys(p).join(',') : '') +
                 (p && p.isGroup === true ? ' | grupo: ' + (p.phone || p.chatId) + ' (' + (p.chatName || '') + ')' : ''));
    }
  });
  Logger.log('Resumo: ' + JSON.stringify(conta));

  const log = ss.getSheetByName(ABA_LOG);
  if (log && log.getLastRow() > 1) {
    const ini = Math.max(2, log.getLastRow() - 499);
    log.getRange(ini, 1, log.getLastRow() - ini + 1, 5).getValues()
      .filter(function (l) { return l[2] === 'ZAP_DOWN'; })
      .forEach(function (l) { Logger.log('LOG ' + l[0] + ' | ' + l[3] + ' | ' + l[4]); });
  }
}

/* Passa para as abas o que o webhook guardou nas propriedades e apaga as
   chaves. Só o gatilho chama (com a trava), então aqui há um escritor só.
   -> quantas fotos novas entraram no ZAP_DOWN_LOG */
function descarregarWebhookZap_(props, todas) {
  const ordenar = function (pref) {
    return Object.keys(todas).filter(function (k) { return k.indexOf(pref) === 0; }).sort();
  };
  const chavesRaw = ordenar('zapdl_raw_'), chavesFila = ordenar('zapdl_q_');

  if (chavesRaw.length) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let aba = ss.getSheetByName(ABA_ZAP_RAW);
    if (!aba) {
      aba = ss.insertSheet(ABA_ZAP_RAW);
      aba.getRange(1, 1, 1, 2).setValues([['recebido_em', 'json']]).setFontWeight('bold');
      aba.setFrozenRows(1);
    }
    const linhas = chavesRaw.map(function (k) {
      let o; try { o = JSON.parse(todas[k]); } catch (e) { o = { t: Date.now(), b: todas[k] }; }
      return [new Date(o.t), o.b];
    });
    aba.getRange(aba.getLastRow() + 1, 1, linhas.length, 2).setValues(linhas);
    const qtd = Number(props.getProperty('ZAP_DOWN_RAW_QTD') || 0) + linhas.length;
    props.setProperty('ZAP_DOWN_RAW_QTD', String(qtd));
    if (qtd >= ZAP_DOWN_RAW_MAX) props.deleteProperty('ZAP_DOWN_RAW');
    chavesRaw.forEach(function (k) { props.deleteProperty(k); });
  }

  if (chavesFila.length) {
    const aba = abaZapDownLog_(true);
    const linhas = [];
    chavesFila.forEach(function (k) {
      try {
        const l = JSON.parse(todas[k]);
        l[0] = new Date(l[0]); l[10] = new Date(l[10]);
        linhas.push(l);
      } catch (e) { registrar_('', 'ZAP_DOWN', 'ERRO', 'fila ilegível ' + k + ': ' + e.message); }
    });
    linhas.sort(function (a, b) { return a[10] - b[10]; });   // ordem em que foram mandadas no grupo
    if (linhas.length) aba.getRange(aba.getLastRow() + 1, 1, linhas.length, COLUNAS_ZAP_DOWN_LOG.length).setValues(linhas);
    // apaga só depois de gravar: se cair antes, a próxima rodada grava de novo (o messageId repetido vira DUPLICADO)
    chavesFila.forEach(function (k) { props.deleteProperty(k); });
    return linhas.length;
  }
  return 0;
}


/* ===========================================================================
   GRUPOS (aba ZAP_DOWN_GRUPOS)
   =========================================================================== */

/* O mesmo grupo aparece como "558184237784-1508587077" ou
   "120363019502650977-group" conforme a rota da Z-API: compara sem o sufixo. */
function idGrupoZap_(id) {
  return String(id || '').trim().replace(/^'/, '').replace(/@g\.us$|-group$/i, '');
}

/* { idNormalizado: { id, nome, cliente, parque } } — só linhas ativas e com cliente e parque. */
function lerGruposZapDown_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ZAP_DOWN_GRUPOS);
  const mapa = {};
  if (!aba || aba.getLastRow() < 2) return mapa;
  aba.getRange(2, 1, aba.getLastRow() - 1, COLUNAS_ZAP_DOWN_GRUPOS.length).getDisplayValues()
    .filter(function (l) { return /^(sim|s|x|1|true)$/i.test(String(l[0]).trim()) && String(l[1]).trim(); })
    .forEach(function (l) {
      const id = idGrupoZap_(l[1]);
      mapa[id] = { id: String(l[1]).trim().replace(/^'/, ''), nome: String(l[2]).trim(),
                   cliente: String(l[3]).trim(), parque: String(l[4]).trim() };
    });
  return mapa;
}

/* Versão do webhook: em cache por 5 min, para não abrir a planilha a cada
   mensagem que chega no número (a Z-API manda TODAS, de todas as conversas). */
function gruposZapDownCache_() {
  const cache = CacheService.getScriptCache();
  const guardado = cache.get('zapdl_grupos');
  if (guardado) return JSON.parse(guardado);
  const todos = lerGruposZapDown_();
  const validos = {};
  Object.keys(todos).forEach(function (k) { if (todos[k].cliente && todos[k].parque) validos[k] = todos[k]; });
  cache.put('zapdl_grupos', JSON.stringify(validos), 300);
  return validos;
}


/* ===========================================================================
   LOG / FILA (aba ZAP_DOWN_LOG)
   =========================================================================== */

function abaZapDownLog_(criar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(ABA_ZAP_DOWN_LOG);
  if (!aba && criar) {
    aba = ss.insertSheet(ABA_ZAP_DOWN_LOG);
    aba.getRange(1, 1, 1, COLUNAS_ZAP_DOWN_LOG.length).setValues([COLUNAS_ZAP_DOWN_LOG])
       .setFontWeight('bold').setBackground('#14181b').setFontColor('#ecede8');
    aba.setFrozenRows(1);
    aba.getRange('B:C').setNumberFormat('@');
    aba.getRange('G:G').setNumberFormat('@');
    aba.getRange('A:A').setNumberFormat('dd/mm/yyyy HH:mm:ss');
    aba.getRange('K:K').setNumberFormat('dd/mm/yyyy HH:mm:ss');
    aba.getRange('L1').setNote('PENDENTE -> OK | ERRO | DUPLICADO. Para tentar os erros de novo: reprocessarErrosZapDownload().');
  }
  return aba;
}

function lerLogZapDown_() {
  const aba = abaZapDownLog_(false);
  if (!aba || aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, COLUNAS_ZAP_DOWN_LOG.length).getValues().map(function (l, i) {
    return {
      linha: i + 2, recebidoEm: l[0], messageId: String(l[1]), grupoId: String(l[2]), nome: String(l[3]),
      cliente: String(l[4]), parque: String(l[5]), autor: String(l[6]), url: String(l[8]),
      tipo: String(l[9]), enviadoEm: l[10] instanceof Date ? l[10] : new Date(),
      status: String(l[11]).toUpperCase(), tentativas: Number(l[12] || 0), detalhe: String(l[14] || '')
    };
  });
}

function gravarLinhaZapDown_(aba, linha, m) {
  if (m.status !== undefined) aba.getRange(linha, 12).setValue(m.status);
  if (m.tentativas !== undefined) aba.getRange(linha, 13).setValue(m.tentativas);
  if (m.caminho !== undefined) aba.getRange(linha, 14).setValue(m.caminho);
  if (m.detalhe !== undefined) aba.getRange(linha, 15).setValue(String(m.detalhe).slice(0, 500));
}


/* ===========================================================================
   PROCESSAMENTO (gatilho de 1 em 1 minuto)
   =========================================================================== */

function processarDownloadsZap() {
  const props = PropertiesService.getScriptProperties();
  const temNovas = function (todas) {
    return Object.keys(todas).some(function (k) { return k.indexOf('zapdl_q_') === 0 || k.indexOf('zapdl_raw_') === 0; });
  };
  // nada novo do webhook nem pendência antiga: sai sem abrir a planilha
  if (!props.getProperty('ZAP_DOWN_TEM_FILA') && !temNovas(props.getProperties())) return;

  // Trava do usuário: a do documento fica com o envio do EHS (minutos) e a do
  // script com os cliques do portal. Esta não disputa com nenhuma das duas.
  const trava = LockService.getUserLock();
  if (!trava.tryLock(1000)) return;
  try {
    const inicio = Date.now();
    // apaga a marca ANTES de ler: pendência que sobrar nesta rodada marca de novo
    props.deleteProperty('ZAP_DOWN_TEM_FILA');
    descarregarWebhookZap_(props, props.getProperties());
    const aba = abaZapDownLog_(true);
    const log = lerLogZapDown_();
    const feitos = {};
    log.forEach(function (l) { if (l.status === 'OK' && l.messageId) feitos[l.messageId] = l.linha; });

    const pendentes = log.filter(function (l) { return l.status === 'PENDENTE'; });
    for (let i = 0; i < pendentes.length; i++) {
      const item = pendentes[i];
      if (Date.now() - inicio > ZAP_DOWN_ORCAMENTO_MS) {
        props.setProperty('ZAP_DOWN_TEM_FILA', '1');     // continua na próxima rodada
        break;
      }
      if (item.messageId && feitos[item.messageId]) {
        gravarLinhaZapDown_(aba, item.linha, { status: 'DUPLICADO', detalhe: 'Já baixada na linha ' + feitos[item.messageId] });
        continue;
      }
      try {
        const caminho = baixarFotoZap_(item);
        gravarLinhaZapDown_(aba, item.linha, { status: 'OK', caminho: caminho, detalhe: '' });
        if (item.messageId) feitos[item.messageId] = item.linha;
      } catch (e) {
        const n = item.tentativas + 1;
        const fim = n >= ZAP_DOWN_MAX_TENTATIVAS;
        gravarLinhaZapDown_(aba, item.linha, { status: fim ? 'ERRO' : 'PENDENTE', tentativas: n, detalhe: e.message });
        if (fim) registrar_('', 'ZAP_DOWN', 'ERRO', item.messageId + ' | ' + e.message);
        else props.setProperty('ZAP_DOWN_TEM_FILA', '1');
      }
    }
  } finally {
    trava.releaseLock();
  }
}

/* Baixa da Z-API e grava no Dropbox. -> caminho gravado */
function baixarFotoZap_(item) {
  if (!item.cliente || !item.parque) throw new Error('Linha sem cliente/parque.');
  const r = UrlFetchApp.fetch(item.url, { method: 'get', muteHttpExceptions: true, followRedirects: true });
  if (r.getResponseCode() !== 200) throw new Error('Mídia da Z-API: HTTP ' + r.getResponseCode() + ' (passou de 30 dias?)');
  const blob = r.getBlob();
  const bytes = blob.getBytes();
  if (!bytes.length) throw new Error('Mídia da Z-API veio vazia.');

  const tz = Session.getScriptTimeZone();
  const d = item.enviadoEm;
  const tipo = String(item.tipo || blob.getContentType() || '').toLowerCase();
  const ext = /png/.test(tipo) ? 'png' : /webp/.test(tipo) ? 'webp' : /heic/.test(tipo) ? 'heic' : 'jpg';
  const nome = Utilities.formatDate(d, tz, 'yyyy-MM-dd_HHmmss') + '_' +
               (String(item.autor).replace(/\D/g, '') || 'autor') + '_' +
               (String(item.messageId).replace(/[^\w-]/g, '') || Date.now()) + '.' + ext;
  const caminho = [pastaZapDown_(), nomePastaZap_(item.cliente), nomePastaZap_(item.parque),
                   'W' + ('0' + semanaIsoZap_(d, tz)).slice(-2), Utilities.formatDate(d, tz, 'dd-MM-yyyy'), nome].join('/');

  const resp = JSON.parse(dbxPortal_('https://content.dropboxapi.com/2/files/upload',
    { path: caminho, mode: 'add', autorename: true, mute: true }, true, bytes).getContentText());
  return resp.path_display || caminho;
}

function pastaZapDown_() {
  let pasta = String(prop_('ZAP_DOWN_PASTA', ZAP_DOWN_PASTA_PADRAO)).replace(/\\/g, '/').trim();
  if (pasta.charAt(0) !== '/') pasta = '/' + pasta;
  return pasta.replace(/\/+$/, '');
}

/* Nome de pasta seguro: sem barra nem caracteres que o Dropbox recusa. */
function nomePastaZap_(t) {
  return String(t).replace(/[\\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').replace(/[. ]+$/, '').trim() || 'SEM NOME';
}

/* Semana ISO 8601 (começa na segunda; a semana 1 é a que tem a 1ª quinta do ano). */
function semanaIsoZap_(data, tz) {
  const p = Utilities.formatDate(data, tz, 'yyyy-MM-dd').split('-').map(Number);
  const t = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  const dia = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - dia);
  const inicioAno = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  return Math.ceil(((t - inicioAno) / 86400000 + 1) / 7);
}
