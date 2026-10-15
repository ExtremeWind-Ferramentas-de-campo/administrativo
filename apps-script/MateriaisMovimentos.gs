/**
 * MOVIMENTO AUTOMÁTICO DE CARDS (Solicitação de Materiais)
 * Lê o movimentos_portal.json (link em Propriedades do script: URL_MOVIMENTOS)
 * e move cards de "envio" para "concluido" ou "cancelado".
 * Nunca cria nem apaga cards. Só move card que está hoje em "envio".
 * testarMovimentos() = só simula. aplicarMovimentos() = grava.
 */

const MAX_MOVIMENTOS = 30;
const MAX_IDADE_ARQUIVO_H = 36;
const COLUNA_ORIGEM = 'envio';
const DESTINOS_VALIDOS = ['concluido', 'cancelado'];
const USUARIO_AUTOMACAO = 'AUTOMACAO';

function testarMovimentos() { return processarMovimentos_(true); }
function aplicarMovimentos() { return processarMovimentos_(false); }

function processarMovimentos_(simular) {
  const url = PropertiesService.getScriptProperties().getProperty('URL_MOVIMENTOS');
  if (!url) { Logger.log('URL_MOVIMENTOS não configurada.'); return; }

  let arq;
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (resp.getResponseCode() !== 200) {
      registrar_(USUARIO_AUTOMACAO, 'MAT_AUTO', 'FALHA', 'HTTP ' + resp.getResponseCode());
      Logger.log('Falha HTTP ' + resp.getResponseCode());
      return;
    }
    arq = JSON.parse(resp.getContentText());
  } catch (e) {
    registrar_(USUARIO_AUTOMACAO, 'MAT_AUTO', 'FALHA', 'leitura do JSON: ' + e);
    Logger.log('Falha na leitura do JSON: ' + e);
    return;
  }

  const geradoEm = arq.gerado_em ? new Date(arq.gerado_em) : null;
  if (!geradoEm || (Date.now() - geradoEm.getTime()) > MAX_IDADE_ARQUIVO_H * 3600 * 1000) {
    registrar_(USUARIO_AUTOMACAO, 'MAT_AUTO', 'IGNORADO', 'arquivo antigo ou sem gerado_em');
    Logger.log('IGNORADO: arquivo antigo ou sem gerado_em.');
    return;
  }

  const movs = (arq.movimentos || []).filter(function (m) {
    return m && (m.codigo || m.card_id) && DESTINOS_VALIDOS.indexOf(m.destino) > -1;
  });
  if (movs.length > MAX_MOVIMENTOS) {
    registrar_(USUARIO_AUTOMACAO, 'MAT_AUTO', 'IGNORADO', movs.length + ' movimentos > limite ' + MAX_MOVIMENTOS);
    Logger.log('IGNORADO: movimentos acima do limite.');
    return;
  }
  if (!movs.length) { Logger.log('Nenhum movimento no arquivo.'); return; }

  const trava = LockService.getScriptLock();
  if (!trava.tryLock(20000)) { Logger.log('Quadro ocupado; tenta no próximo acionamento.'); return; }

  try {
    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_MAT_CARDS);
    const dados = aba.getDataRange().getValues();
    const agora = new Date();
    let movidos = 0;
    const relatorio = [];

    movs.forEach(function (m) {
      const codigo = normalizar_(m.codigo);
      const projeto = m.card_id ? '' : normalizar_(m.projeto || '');
      const candidatos = [];

      for (let l = 1; l < dados.length; l++) {
        if (!dados[l][0]) continue;
        if (m.card_id) { if (String(dados[l][0]) !== String(m.card_id)) continue; }
        else if (normalizar_(dados[l][1]) !== codigo) continue;
        let card;
        try { card = JSON.parse(dados[l][4]); } catch (e) { continue; }
        if (projeto && normalizar_(JSON.stringify(card)).indexOf(projeto) === -1) continue;
        candidatos.push({ linha: l, card: card });
      }

      const rotulo = (m.codigo || '') + ' ' + (m.projeto || '');
      if (candidatos.length === 0) { relatorio.push(rotulo + ': NAO_ENCONTRADO'); return; }
      if (candidatos.length > 1)   { relatorio.push(rotulo + ': AMBIGUO (' + candidatos.length + ')'); return; }

      const c = candidatos[0];
      const colunaAtual = String(c.card.columnId || dados[c.linha][2]);
      if (colunaAtual !== COLUNA_ORIGEM) {
        relatorio.push(rotulo + ': PULADO (já em ' + colunaAtual + ')');
        return;
      }

      relatorio.push(rotulo + ': ' + COLUNA_ORIGEM + ' -> ' + m.destino + (simular ? ' (simulado)' : ''));
      if (simular) return;

      const novaRev = (Number(dados[c.linha][3]) || 0) + 1;
      c.card.columnId = m.destino;
      if (Array.isArray(c.card.history)) {
        c.card.history.push({ at: agora.toISOString(), by: USUARIO_AUTOMACAO,
                              note: 'movido automaticamente: ' + (m.motivo || m.destino) });
      }
      aba.getRange(c.linha + 1, 1, 1, COLUNAS_MAT_CARDS.length).setValues([[
        dados[c.linha][0], c.card.code || dados[c.linha][1], m.destino, novaRev,
        JSON.stringify(c.card), agora, USUARIO_AUTOMACAO
      ]]);
      movidos++;
    });

    if (!simular && movidos) bumparMeta_({});
    Logger.log(relatorio.join('\n'));
    registrar_(USUARIO_AUTOMACAO, 'MAT_AUTO', simular ? 'SIMULACAO' : 'OK',
               movidos + ' movidos; ' + relatorio.join(' | ').slice(0, 800));
  } finally {
    trava.releaseLock();
  }
}

function normalizar_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}
