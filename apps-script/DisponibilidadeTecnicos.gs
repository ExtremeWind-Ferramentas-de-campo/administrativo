/**
 * ============================================================================
 * MÓDULO: DISPONIBILIDADE DOS TÉCNICOS (painel + lembrete no WhatsApp)
 * Cole como novo arquivo (Arquivo > + > Script, nome "DisponibilidadeTecnicos")
 * no mesmo projeto do portal. Usa do SupervisaoCampo.gs: prop_, extrairId_,
 * abaMiniMaster_, chaveTexto_, soNumero_, perfilDe_; do ZapEHS.gs: zapPost_,
 * zapiBase_, zapiHeaders_; do Codigo.gs: validarSessao_, registrar_, avisar_.
 *
 * DE ONDE VÊM OS DADOS
 *  - Lista de técnicos e telefones: MINI MASTER (a mesma do resto do portal).
 *      coluna A = matrícula, coluna B = nome; a coluna do telefone é achada
 *      pelo CABEÇALHO (telefone / celular / whatsapp / fone / contato).
 *      Se quiser fixar, crie a propriedade MINIMASTER_COL_TELEFONE com a letra
 *      (ex.: F) ou o nome exato do cabeçalho.
 *      Se a MINI MASTER tiver uma coluna "situação"/"status" com DESLIGADO,
 *      INATIVO, DEMITIDO ou AFASTADO, o técnico fica de fora.
 *  - Respostas: a planilha "Disponibilidade" que o Ferramentas de Campo grava
 *      quando o técnico salva em Meus Dados > Disponibilidade
 *      (A = Matrícula | B = Nome | C = Semanas em JSON {"2026-W45":"S"} | D = Atualizado em).
 *
 * O QUE CONTA COMO "FALTA PREENCHER"
 *   Qualquer uma das próximas DISPT_SEMANAS semanas (a atual + 3) sem resposta.
 *   Semanas ISO (segunda a domingo), mesma regra do Ferramentas de Campo.
 *
 * LEMBRETE NO WHATSAPP
 *   Vai pelo mesmo número da Z-API do report de EHS, no privado de cada
 *   técnico — não precisa salvar contato. O botão NÃO envia na hora: coloca
 *   cada técnico na fila (aba DISP_ZAP_FILA) e o gatilho de 1 em 1 minuto
 *   manda, com pausa entre as mensagens para o número não ser bloqueado.
 *   Só ADMIN e SUPERVISOR podem disparar. Todos os perfis veem o painel.
 *
 * PROPRIEDADES DO SCRIPT
 *   ID_DISPONIBILIDADE      link ou ID da planilha de disponibilidade (OBRIGATÓRIA)
 *                           — a mesma da propriedade DISP_SHEET do Ferramentas de Campo.
 *                           A conta que publicou o portal precisa ter ao menos LEITURA nela.
 *   ABA_DISPONIBILIDADE     nome da aba (padrão: Disponibilidade)
 *   MINIMASTER_COL_TELEFONE letra ou cabeçalho da coluna do telefone (opcional)
 *   LINK_FERRAMENTAS_CAMPO  endereço do portal Ferramentas de Campo, que entra
 *                           na mensagem (opcional)
 *   ZAPI_INSTANCE_ID / ZAPI_TOKEN / ZAPI_CLIENT_TOKEN: as mesmas do EHS.
 *
 * INSTALAÇÃO: cole, rode instalarDisponibilidadeTecnicos() uma vez (cria a
 * aba da fila e o gatilho), crie ID_DISPONIBILIDADE e reimplante o App da Web.
 * Diagnóstico: testarDisponibilidadeTecnicos().
 * ============================================================================
 */

const DISPT_SEMANAS = 4;                     // semana atual + 3
const DISPT_TZ = 'America/Recife';
const ABA_DISPT_FILA = 'DISP_ZAP_FILA';
const COLUNAS_DISPT_FILA = ['lote', 'matricula', 'nome', 'telefone', 'mensagem', 'status',
                            'tentativas', 'criado_em', 'criado_por', 'enviado_em', 'detalhe'];
const DISPT_PERFIS_ENVIAM = ['ADMIN', 'SUPERVISOR'];
const DISPT_PAUSA_MIN_MS = 4000;             // pausa aleatória entre mensagens (4 a 8 s)
const DISPT_PAUSA_MAX_MS = 8000;
const DISPT_ORCAMENTO_MS = 4.5 * 60000;      // limite do Apps Script: 6 min por execução
const DISPT_MAX_TENTATIVAS = 3;
const DISPT_MSG_MAX = 1000;

const DISPT_MSG_PADRAO =
  'Olá, {nome}! Tudo bem?\n\n' +
  'Aqui é da Extreme Wind. Precisamos que você atualize a sua *disponibilidade* ' +
  'das próximas semanas no portal Ferramentas de Campo, em *Meus Dados > Disponibilidade*.' +
  '{link}\n\n' +
  'Leva menos de um minuto. Obrigado!';


/* ===========================================================================
   INSTALAÇÃO E DIAGNÓSTICO
   =========================================================================== */

function instalarDisponibilidadeTecnicos() {
  abaDisptFila_(true);
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'processarFilaDisponibilidade'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('processarFilaDisponibilidade').timeBased().everyMinutes(1).create();

  avisar_('Disponibilidade dos técnicos instalada.\n\n' +
          '1. Crie a propriedade ID_DISPONIBILIDADE com o link da planilha de disponibilidade ' +
          '(a mesma do DISP_SHEET do Ferramentas de Campo).\n' +
          '2. (Opcional) LINK_FERRAMENTAS_CAMPO com o endereço do portal de campo, para ir na mensagem.\n' +
          '3. Rode testarDisponibilidadeTecnicos() para conferir.\n' +
          '4. Reimplante o App da Web.\n\n' +
          'Gatilho criado: processarFilaDisponibilidade a cada 1 minuto.');
}

function testarDisponibilidadeTecnicos() {
  try {
    const t = lerTecnicosComTelefone_();
    Logger.log('MINI MASTER: ' + t.lista.length + ' técnicos ativos | coluna do telefone: ' +
               (t.colTelefone || 'NÃO ACHADA') + ' | com telefone válido: ' +
               t.lista.filter(function (x) { return x.telefone; }).length);
    t.lista.filter(function (x) { return !x.telefone; }).slice(0, 20).forEach(function (x) {
      Logger.log('  sem telefone válido: ' + x.nome + ' (' + x.matricula + ') — valor na planilha: "' + x.telefoneCru + '"');
    });
  } catch (e) { Logger.log('MINI MASTER: ERRO — ' + e.message); }
  try {
    const r = lerRespostasDisp_();
    Logger.log('Planilha de disponibilidade: ' + Object.keys(r).length + ' técnico(s) com resposta.');
  } catch (e) { Logger.log('Disponibilidade: ERRO — ' + e.message); }
  Logger.log('Semanas cobradas: ' + disptSemanas_().map(function (s) { return 'W' + s.n + ' (' + s.ini + '–' + s.fim + ')'; }).join(', '));
  try {
    const r = UrlFetchApp.fetch(zapiBase_() + '/status', { method: 'get', headers: zapiHeaders_(), muteHttpExceptions: true });
    Logger.log('Z-API /status: ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
  } catch (e) { Logger.log('Z-API: ' + e.message); }
}


/* ===========================================================================
   SEMANAS (mesma regra ISO do Ferramentas de Campo)
   =========================================================================== */

function disptPad_(n) { return (n < 10 ? '0' : '') + n; }

function disptSemanas_() {
  const p = Utilities.formatDate(new Date(), DISPT_TZ, 'yyyy-MM-dd').split('-');
  const hoje = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]));
  const dow = hoje.getUTCDay() || 7;
  const seg = new Date(hoje.getTime() - (dow - 1) * 86400000);
  const ddmm = function (d) { return disptPad_(d.getUTCDate()) + '/' + disptPad_(d.getUTCMonth() + 1); };
  const lista = [];
  for (let k = 0; k < DISPT_SEMANAS; k++) {
    const s = new Date(seg.getTime() + k * 7 * 86400000);
    const qui = new Date(s.getTime() + 3 * 86400000);
    const dom = new Date(s.getTime() + 6 * 86400000);
    const ano = qui.getUTCFullYear();
    const n = Math.floor((qui.getTime() - Date.UTC(ano, 0, 1)) / 86400000 / 7) + 1;
    lista.push({ chave: ano + '-W' + disptPad_(n), n: n, ini: ddmm(s), fim: ddmm(dom), atual: k === 0 });
  }
  return lista;
}


/* ===========================================================================
   LEITURA
   =========================================================================== */

/* Telefone brasileiro -> 55DDDNÚMERO (só dígitos) ou '' se não der para usar. */
function disptTelefone_(cru) {
  let d = String(cru == null ? '' : cru).replace(/\D/g, '');
  if (!d) return '';
  d = d.replace(/^0+/, '');                                  // 0 de longa distância
  if (d.length === 10 || d.length === 11) d = '55' + d;      // sem DDI
  if (!/^55\d{10,11}$/.test(d)) return '';
  // celular com 8 dígitos (antigo): põe o 9 na frente
  if (d.length === 12 && /[6-9]/.test(d.charAt(4))) d = d.slice(0, 4) + '9' + d.slice(4);
  return d;
}

function disptColuna_(cab, regex) {
  for (let i = 0; i < cab.length; i++) if (regex.test(chaveTexto_(cab[i]))) return i;
  return -1;
}

/** -> { lista:[{ matricula, nome, telefone, telefoneCru }], colTelefone } */
function lerTecnicosComTelefone_() {
  const aba = abaMiniMaster_();
  const ultima = aba.getLastRow(), ultCol = aba.getLastColumn();
  if (ultima < 2) return { lista: [], colTelefone: '' };
  const dados = aba.getRange(1, 1, ultima, ultCol).getDisplayValues();
  const cab = dados[0];

  let iTel = -1;
  const fixo = String(prop_('MINIMASTER_COL_TELEFONE', '')).trim();
  if (fixo) {
    if (/^[A-Za-z]{1,2}$/.test(fixo)) {
      iTel = fixo.toUpperCase().split('').reduce(function (s, c) { return s * 26 + c.charCodeAt(0) - 64; }, 0) - 1;
    } else {
      iTel = disptColuna_(cab, new RegExp('^' + chaveTexto_(fixo).replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '$'));
    }
  }
  if (iTel < 0) iTel = disptColuna_(cab, /whats|celular|telefone|^fone|contato/);
  const iSit = disptColuna_(cab, /^(situacao|status|ativo)/);

  const vistos = {};
  const lista = [];
  dados.slice(1).forEach(function (l) {
    const mat = String(l[0] || '').replace(/^'/, '').trim();
    const nome = String(l[1] || '').replace(/\s+/g, ' ').trim();
    if (!nome || !soNumero_(mat)) return;
    if (iSit > -1 && /deslig|inativ|demit|afast|^nao$/.test(chaveTexto_(l[iSit]))) return;
    const k = soNumero_(mat);
    if (vistos[k]) return;
    vistos[k] = true;
    const cru = iTel > -1 ? String(l[iTel] || '') : '';
    lista.push({ matricula: mat, nome: nome, telefone: disptTelefone_(cru), telefoneCru: cru });
  });
  lista.sort(function (a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
  return { lista: lista, colTelefone: iTel > -1 ? (cab[iTel] || ('coluna ' + (iTel + 1))) : '' };
}

/** -> { soNumero(matricula): { nome, valores:{ '2026-W45':'S'|'N' }, atualizadoEm } } */
function lerRespostasDisp_() {
  const id = extrairId_(prop_('ID_DISPONIBILIDADE'));
  if (!id) throw new Error('Crie a propriedade ID_DISPONIBILIDADE com o link da planilha de disponibilidade.');
  let ss;
  try { ss = SpreadsheetApp.openById(id); }
  catch (e) {
    let conta = '';
    try { conta = Session.getEffectiveUser().getEmail(); } catch (e2) { /* sem e-mail */ }
    throw new Error('Sem acesso à planilha de disponibilidade. Compartilhe-a (Leitor basta) com a conta ' +
                    (conta || 'que publicou o portal') + ' e confira o link em ID_DISPONIBILIDADE. Detalhe: ' + e.message);
  }
  const aba = ss.getSheetByName(prop_('ABA_DISPONIBILIDADE', 'Disponibilidade'));
  if (!aba) throw new Error('A aba "' + prop_('ABA_DISPONIBILIDADE', 'Disponibilidade') + '" não existe na planilha de disponibilidade.');
  const mapa = {};
  if (aba.getLastRow() < 2) return mapa;
  aba.getRange(2, 1, aba.getLastRow() - 1, 4).getValues().forEach(function (r) {
    const k = soNumero_(r[0]);
    if (!k) return;
    let v = {};
    try { v = JSON.parse(r[2] || '{}') || {}; } catch (e) { v = {}; }
    mapa[k] = {
      nome: String(r[1] || ''), valores: v,
      atualizadoEm: r[3] instanceof Date ? Utilities.formatDate(r[3], DISPT_TZ, 'dd/MM/yyyy HH:mm') : String(r[3] || '')
    };
  });
  return mapa;
}


/* ===========================================================================
   AÇÕES DA TELA
   =========================================================================== */

/**
 * -> { ok, semanas, tecnicos:[{ matricula, nome, telefone, valores, atualizadoEm,
 *      situacao:'COMPLETO'|'PARCIAL'|'NUNCA', faltam, lembrete:{status, em}|null }],
 *      envio:{ lote, em, por, total, enviados, pendentes, erros }|null, podeEnviar, colTelefone }
 */
function acaoDisponibilidadeStatus_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  let tec, resp;
  try { tec = lerTecnicosComTelefone_(); }
  catch (e) { return { ok: false, motivo: 'MINIMASTER_INDISPONIVEL', detalhe: e.message }; }
  try { resp = lerRespostasDisp_(); }
  catch (e) { return { ok: false, motivo: 'DISP_INDISPONIVEL', detalhe: e.message }; }

  const semanas = disptSemanas_();
  const fila = lerFilaDispt_();
  const ultimoPorMat = {};
  fila.forEach(function (f) { ultimoPorMat[soNumero_(f.matricula)] = f; });   // a fila está em ordem: fica o mais recente

  const tecnicos = tec.lista.map(function (t) {
    const r = resp[soNumero_(t.matricula)];
    const valores = {};
    let faltam = 0;
    semanas.forEach(function (s) {
      const v = r && r.valores[s.chave];
      if (v === 'S' || v === 'N') valores[s.chave] = v; else faltam++;
    });
    const f = ultimoPorMat[soNumero_(t.matricula)];
    return {
      matricula: t.matricula, nome: t.nome, telefone: t.telefone, telefoneCru: t.telefone ? '' : t.telefoneCru,
      valores: valores, atualizadoEm: r ? r.atualizadoEm : '', faltam: faltam,
      situacao: !r ? 'NUNCA' : (faltam ? 'PARCIAL' : 'COMPLETO'),
      lembrete: f ? { status: f.status, em: f.enviadoEm || f.criadoEm, detalhe: f.status === 'ERRO' ? f.detalhe : '' } : null
    };
  });

  return {
    ok: true, semanas: semanas, tecnicos: tecnicos, colTelefone: tec.colTelefone,
    envio: resumoUltimoLote_(fila),
    podeEnviar: DISPT_PERFIS_ENVIAM.indexOf(perfilDe_(matricula)) > -1,
    mensagemPadrao: DISPT_MSG_PADRAO
  };
}

/**
 * { mensagem, matriculas? } -> { ok, lote, naFila, semTelefone }
 * Sem "matriculas" = todos os técnicos da MINI MASTER com telefone.
 */
function acaoDisponibilidadeLembrar_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };
  if (DISPT_PERFIS_ENVIAM.indexOf(perfilDe_(matricula)) < 0) return { ok: false, motivo: 'SEM_PERMISSAO' };

  let modelo = String(p.mensagem || '').trim().slice(0, DISPT_MSG_MAX);
  if (!modelo) modelo = DISPT_MSG_PADRAO;

  let tec;
  try { tec = lerTecnicosComTelefone_().lista; }
  catch (e) { return { ok: false, motivo: 'MINIMASTER_INDISPONIVEL', detalhe: e.message }; }

  if (Array.isArray(p.matriculas) && p.matriculas.length) {
    const quer = {};
    p.matriculas.forEach(function (m) { quer[soNumero_(m)] = true; });
    tec = tec.filter(function (t) { return quer[soNumero_(t.matricula)]; });
  }
  const comTel = tec.filter(function (t) { return t.telefone; });
  if (!comTel.length) return { ok: false, motivo: 'SEM_DESTINATARIO' };

  const trava = LockService.getScriptLock();
  if (!trava.tryLock(20000)) return { ok: false, motivo: 'OCUPADO' };
  try {
    const fila = lerFilaDispt_();
    const pendentes = {};
    fila.forEach(function (f) { if (f.status === 'PENDENTE') pendentes[soNumero_(f.matricula)] = true; });
    const novos = comTel.filter(function (t) { return !pendentes[soNumero_(t.matricula)]; });
    if (!novos.length) return { ok: false, motivo: 'JA_NA_FILA' };

    const lote = Utilities.formatDate(new Date(), DISPT_TZ, 'yyyyMMdd-HHmmss');
    const agora = new Date();
    const link = prop_('LINK_FERRAMENTAS_CAMPO', '');
    const linhas = novos.map(function (t) {
      const primeiro = t.nome.split(' ')[0];
      const nomeBonito = primeiro.charAt(0).toUpperCase() + primeiro.slice(1).toLowerCase();
      const msg = modelo.replace(/\{nome\}/g, nomeBonito)
                        .replace(/\{link\}/g, link ? '\n\n' + link : '');
      return [lote, "'" + t.matricula, t.nome, "'" + t.telefone, msg, 'PENDENTE', 0, agora, "'" + matricula, '', ''];
    });
    const aba = abaDisptFila_(true);
    aba.getRange(aba.getLastRow() + 1, 1, linhas.length, COLUNAS_DISPT_FILA.length).setValues(linhas);
    registrar_(matricula, 'DISP_LEMBRETE', 'OK', 'lote ' + lote + ' | ' + linhas.length + ' técnico(s)');
    return { ok: true, lote: lote, naFila: linhas.length, semTelefone: tec.length - comTel.length,
             jaNaFila: comTel.length - novos.length };
  } finally {
    trava.releaseLock();
  }
}


/* ===========================================================================
   FILA (aba DISP_ZAP_FILA) E GATILHO
   =========================================================================== */

function abaDisptFila_(criar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(ABA_DISPT_FILA);
  if (!aba && criar) {
    aba = ss.insertSheet(ABA_DISPT_FILA);
    aba.getRange(1, 1, 1, COLUNAS_DISPT_FILA.length).setValues([COLUNAS_DISPT_FILA])
       .setFontWeight('bold').setBackground('#14181b').setFontColor('#ecede8');
    aba.setFrozenRows(1);
    aba.getRange('A:B').setNumberFormat('@');
    aba.getRange('D:D').setNumberFormat('@');
    aba.getRange('H:H').setNumberFormat('dd/mm/yyyy hh:mm');
    aba.getRange('I:I').setNumberFormat('@');
    aba.getRange('J:J').setNumberFormat('dd/mm/yyyy hh:mm');
    aba.getRange('F1').setNote('PENDENTE -> ENVIADO ou ERRO. Para cancelar um envio, troque PENDENTE por CANCELADO.');
  }
  return aba;
}

function lerFilaDispt_() {
  const aba = abaDisptFila_(false);
  if (!aba || aba.getLastRow() < 2) return [];
  const n = aba.getLastRow() - 1;
  const v = aba.getRange(2, 1, n, COLUNAS_DISPT_FILA.length).getValues();
  const fmt = function (d) { return d instanceof Date ? Utilities.formatDate(d, DISPT_TZ, 'dd/MM/yyyy HH:mm') : String(d || ''); };
  return v.map(function (l, i) {
    return {
      linha: i + 2, lote: String(l[0]), matricula: String(l[1]).replace(/^'/, ''), nome: String(l[2]),
      telefone: String(l[3]).replace(/^'/, ''), mensagem: String(l[4]), status: String(l[5]).toUpperCase(),
      tentativas: Number(l[6]) || 0, criadoEm: fmt(l[7]), por: String(l[8]).replace(/^'/, ''),
      enviadoEm: fmt(l[9]), detalhe: String(l[10] || '')
    };
  });
}

function resumoUltimoLote_(fila) {
  if (!fila.length) return null;
  const lote = fila[fila.length - 1].lote;
  const itens = fila.filter(function (f) { return f.lote === lote; });
  const conta = function (st) { return itens.filter(function (f) { return f.status === st; }).length; };
  let por = itens[0].por;
  try {
    const u = acharPorMatricula_(abaUsuarios_(), por);
    if (u && u.dados.nome) por = u.dados.nome;
  } catch (e) { /* fica a matrícula */ }
  return { lote: lote, em: itens[0].criadoEm, por: por, total: itens.length,
           enviados: conta('ENVIADO'), pendentes: conta('PENDENTE'), erros: conta('ERRO'), cancelados: conta('CANCELADO') };
}

/** Gatilho de 1 em 1 minuto: manda o que estiver PENDENTE, com pausa entre as mensagens. */
function processarFilaDisponibilidade() {
  // Não usa o ScriptLock: segurá-lo por minutos travaria os botões das outras
  // telas (todas usam a mesma trava). Uma marca no cache basta para uma
  // rodada não atropelar a outra.
  const cache = CacheService.getScriptCache();
  if (cache.get('dispt_rodando')) return;               // outra rodada ainda enviando
  cache.put('dispt_rodando', '1', 330);
  try {
    const pend = lerFilaDispt_().filter(function (f) { return f.status === 'PENDENTE'; });
    if (!pend.length) return;
    const aba = abaDisptFila_(true);
    const inicio = Date.now();
    for (let i = 0; i < pend.length; i++) {
      if (Date.now() - inicio > DISPT_ORCAMENTO_MS) break;          // continua na próxima rodada
      const f = pend[i];
      // pode ter sido cancelado na aba enquanto esta rodada andava
      if (String(aba.getRange(f.linha, 6).getValue()).toUpperCase() !== 'PENDENTE') continue;
      let r;
      try { r = zapPost_('/send-text', { phone: f.telefone, message: f.mensagem }); }
      catch (e) { r = { ok: false, erro: String(e.message || e) }; }
      const tent = f.tentativas + 1;
      if (r.ok) {
        aba.getRange(f.linha, 6, 1, 2).setValues([['ENVIADO', tent]]);
        aba.getRange(f.linha, 10, 1, 2).setValues([[new Date(), '']]);
      } else {
        aba.getRange(f.linha, 6, 1, 2).setValues([[tent >= DISPT_MAX_TENTATIVAS ? 'ERRO' : 'PENDENTE', tent]]);
        aba.getRange(f.linha, 11).setValue(String(r.erro || 'falha').slice(0, 300));
      }
      SpreadsheetApp.flush();
      if (i < pend.length - 1) {
        Utilities.sleep(DISPT_PAUSA_MIN_MS + Math.floor(Math.random() * (DISPT_PAUSA_MAX_MS - DISPT_PAUSA_MIN_MS)));
      }
    }
  } finally {
    cache.remove('dispt_rodando');
  }
}
