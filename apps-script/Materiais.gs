/**
 * ============================================================================
 * MÓDULO: SOLICITAÇÃO DE MATERIAIS
 * Complemento do Codigo.gs. Cole este arquivo como um segundo script no mesmo
 * projeto do Apps Script (Arquivo > + > Script, nome "Materiais").
 *
 * Depois de colar, rode  configurarMateriais()  uma vez.
 *
 * COMO OS DADOS FICAM
 *  - Aba MAT_CARDS: uma linha por solicitação. A ficha inteira do card vai em
 *    JSON na coluna `json`. Cada linha tem sua própria revisão (`rev`), então
 *    duas pessoas mexendo em cards diferentes nunca se atrapalham.
 *  - Aba MAT_META: nextSeq, colunas do quadro e a revisão global (`rev`).
 *    A revisão global é o que o quadro consulta a cada 8 segundos para saber
 *    se precisa recarregar — sem ela, cada consulta leria a planilha inteira.
 *  - Anexos NÃO ficam na planilha. Vão para uma pasta do Drive e o card guarda
 *    só nome, tamanho, tipo e o id do arquivo. Uma célula do Sheets aceita
 *    50 mil caracteres; um anexo de 3 MB em base64 tem 4 milhões.
 *
 * PRIVACIDADE DOS ANEXOS
 *  Os arquivos ficam privados no Drive, sem link público. Para baixar, o quadro
 *  chama materiaisBaixarAnexo com o token da sessão; o Apps Script confere o
 *  token, lê o arquivo e devolve o conteúdo. Quem não estiver logado no portal
 *  não abre anexo nem com o id na mão.
 * ============================================================================
 */

const ABA_MAT_CARDS = 'MAT_CARDS';
const ABA_MAT_META  = 'MAT_META';

const COLUNAS_MAT_CARDS = ['id', 'code', 'coluna', 'rev', 'json', 'atualizado_em', 'atualizado_por'];

const COLUNAS_PADRAO_QUADRO = [
  { id: 'nova',       name: 'Nova Solicitação' },
  { id: 'separacao',  name: 'Separação (Almoxarifado)' },
  { id: 'envio',      name: 'Envio (Logística)' },
  { id: 'concluido',  name: 'Concluído' },
  { id: 'cancelado',  name: 'Cancelado' }
];

const MAX_ANEXO_BYTES = 4 * 1024 * 1024;


/* ===========================================================================
   INSTALAÇÃO
   =========================================================================== */

function configurarMateriais() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  criarAba_(ss, ABA_MAT_CARDS, COLUNAS_MAT_CARDS);
  criarAba_(ss, ABA_MAT_META,  ['chave', 'valor']);

  const meta = ss.getSheetByName(ABA_MAT_META);
  if (meta.getLastRow() < 2) {
    meta.getRange(2, 1, 3, 2).setValues([
      ['rev', 1],
      ['nextSeq', 1],
      ['columns', JSON.stringify(COLUNAS_PADRAO_QUADRO)]
    ]);
  }

  // A coluna do JSON fica larga demais para leitura manual — encolhe e esconde.
  const cards = ss.getSheetByName(ABA_MAT_CARDS);
  cards.setColumnWidth(5, 120);
  cards.getRange('E1').setNote('Ficha completa do card. Editar à mão aqui quebra o quadro.');

  pastaAnexos_();  // cria a pasta do Drive na primeira execução

  avisar_('Módulo de materiais pronto.\n\n' +
          'Pasta de anexos criada no seu Drive: "Extreme Wind — Anexos de Materiais".');
}

/** Devolve a pasta do Drive dos anexos, criando na primeira vez. */
function pastaAnexos_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PASTA_ANEXOS');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* foi apagada: recria abaixo */ }
  }
  const pasta = DriveApp.createFolder('Extreme Wind — Anexos de Materiais');
  props.setProperty('PASTA_ANEXOS', pasta.getId());
  return pasta;
}


/* ===========================================================================
   AÇÕES CHAMADAS PELO QUADRO
   Ligadas ao doPost em Codigo.gs — veja o switch lá.
   =========================================================================== */

/** { rev } -> { ok, rev, mudou, estado? }  (rev: -1 força carga completa) */
function acaoMateriaisCarregar_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  const meta = lerMeta_();

  // Consulta barata: se nada mudou, não lê a aba de cards.
  if (Number(p.rev) === meta.rev) return { ok: true, rev: meta.rev, mudou: false };

  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_MAT_CARDS);
  const dados = aba.getDataRange().getValues();
  const iId = 0, iRev = 3, iJson = 4;

  const cards = [];
  for (let l = 1; l < dados.length; l++) {
    if (!dados[l][iId]) continue;
    try {
      const card = JSON.parse(dados[l][iJson]);
      card._rev = Number(dados[l][iRev]) || 0;
      cards.push(card);
    } catch (e) {
      registrar_(matricula, 'MAT_CARREGAR', 'FALHA', 'json inválido na linha ' + (l + 1));
    }
  }

  return {
    ok: true,
    rev: meta.rev,
    mudou: true,
    estado: { version: 1, nextSeq: meta.nextSeq, columns: meta.columns, cards: cards }
  };
}


/** { cards[], removidos[], nextSeq, columns } -> { ok, rev, aceitos[], conflitos[] } */
function acaoMateriaisSalvar_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  const trava = LockService.getScriptLock();
  if (!trava.tryLock(20000)) return { ok: false, motivo: 'OCUPADO' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(ABA_MAT_CARDS);
    const dados = aba.getDataRange().getValues();

    // Mapa id -> { linha, rev, json } para não varrer a aba a cada card.
    const mapa = {};
    for (let l = 1; l < dados.length; l++) {
      if (dados[l][0]) mapa[String(dados[l][0])] = { linha: l + 1, rev: Number(dados[l][3]) || 0, json: dados[l][4] };
    }

    const agora = new Date();
    const aceitos = [];
    const conflitos = [];
    const novasLinhas = [];

    (p.cards || []).forEach(function (card) {
      const id = String(card.id);
      const revEnviada = Number(card._rev) || 0;
      delete card._rev;

      const atual = mapa[id];

      // Card novo: só entra se o cliente também acha que é novo.
      if (!atual) {
        novasLinhas.push([id, card.code || '', card.columnId || '', 1,
                          JSON.stringify(card), agora, matricula]);
        aceitos.push({ id: id, rev: 1 });
        return;
      }

      // Alguém salvou este card entre a leitura do cliente e agora.
      if (atual.rev !== revEnviada) {
        let servidor = null;
        try { servidor = JSON.parse(atual.json); } catch (e) { servidor = null; }
        if (servidor) {
          servidor._rev = atual.rev;
          conflitos.push({ id: id, rev: atual.rev, card: servidor });
          return;
        }
        // JSON corrompido na planilha: aceitar a versão do cliente é melhor que travar.
      }

      const novaRev = atual.rev + 1;
      aba.getRange(atual.linha, 1, 1, COLUNAS_MAT_CARDS.length).setValues([[
        id, card.code || '', card.columnId || '', novaRev,
        JSON.stringify(card), agora, matricula
      ]]);
      aceitos.push({ id: id, rev: novaRev });
    });

    if (novasLinhas.length) {
      aba.getRange(aba.getLastRow() + 1, 1, novasLinhas.length, COLUNAS_MAT_CARDS.length)
         .setValues(novasLinhas);
    }

    // Exclusões, de baixo para cima para os índices não se deslocarem.
    const remover = (p.removidos || [])
      .map(function (id) { return mapa[String(id)] ? mapa[String(id)].linha : null; })
      .filter(function (l) { return l; })
      .sort(function (a, b) { return b - a; });
    remover.forEach(function (linha) { aba.deleteRow(linha); });

    const rev = bumparMeta_({
      nextSeq: p.nextSeq,
      columns: p.columns
    });

    // Só registra o que foge do normal. Cada linha de MAT_CARDS já guarda
    // `atualizado_por` e `atualizado_em`, então gravar todo salvamento aqui
    // seria repetir a mesma informação milhares de vezes por mês.
    if (conflitos.length || remover.length) {
      registrar_(matricula, 'MAT_SALVAR', conflitos.length ? 'CONFLITO' : 'OK',
        aceitos.length + ' aceitos, ' + conflitos.length + ' conflitos, ' + remover.length + ' removidos');
    }

    return { ok: true, rev: rev, aceitos: aceitos, conflitos: conflitos };

  } finally {
    trava.releaseLock();
  }
}


/**
 * { nome, tipo, base64, thumbBase64? } -> { ok, driveId, thumbId? }
 *
 * Quando o anexo é imagem, o navegador manda junto uma miniatura já reduzida.
 * Ela vira um segundo arquivo no Drive. É isso que permite o quadro mostrar as
 * fotos nos cards sem baixar o original de cada uma — a miniatura tem uns 15 KB,
 * o original pode ter 4 MB.
 */
function acaoMateriaisAnexar_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };
  if (!p.base64 || !p.nome) return { ok: false, motivo: 'DADOS' };

  let bytes;
  try {
    bytes = Utilities.base64Decode(p.base64);
  } catch (e) {
    return { ok: false, motivo: 'ARQUIVO_INVALIDO' };
  }
  if (bytes.length > MAX_ANEXO_BYTES) return { ok: false, motivo: 'GRANDE_DEMAIS' };

  const blob = Utilities.newBlob(bytes, p.tipo || 'application/octet-stream', p.nome);
  const arquivo = pastaAnexos_().createFile(blob);

  // Sem compartilhamento: o download passa pelo Web App, que confere o token.
  arquivo.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  arquivo.setDescription('Anexo do quadro de materiais. Enviado por ' + matricula + '.');

  let thumbId = null;
  if (p.thumbBase64) {
    try {
      const tb = Utilities.base64Decode(p.thumbBase64);
      const tf = pastaAnexos_().createFile(
        Utilities.newBlob(tb, 'image/jpeg', 'thumb_' + arquivo.getId() + '.jpg'));
      tf.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      thumbId = tf.getId();
    } catch (e) {
      // Sem miniatura o card mostra um ícone no lugar da foto. Não é motivo
      // para perder o anexo, que já subiu.
      registrar_(matricula, 'MAT_ANEXAR', 'AVISO', 'miniatura falhou: ' + p.nome);
    }
  }

  registrar_(matricula, 'MAT_ANEXAR', 'OK', p.nome + ' (' + bytes.length + ' bytes)');
  return { ok: true, driveId: arquivo.getId(), thumbId: thumbId };
}


/**
 * { ids: [driveId da miniatura] } -> { ok, mapa: { id: base64 } }
 *
 * Em lote de propósito: o quadro pede todas as miniaturas de uma vez ao abrir.
 * Uma chamada por imagem estouraria a cota do Apps Script num quadro cheio.
 */
function acaoMateriaisMiniaturas_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  const ids = (p.ids || []).slice(0, 60);
  const mapa = {};
  const alvo = pastaAnexos_().getId();

  ids.forEach(function (id) {
    try {
      const f = DriveApp.getFileById(String(id));
      let dentro = false;
      const pais = f.getParents();
      while (pais.hasNext()) { if (pais.next().getId() === alvo) { dentro = true; break; } }
      if (!dentro) return;
      mapa[id] = Utilities.base64Encode(f.getBlob().getBytes());
    } catch (e) { /* miniatura sumiu: o card cai no ícone */ }
  });

  return { ok: true, mapa: mapa };
}


/** { driveId } -> { ok, base64, nome, tipo } */
function acaoMateriaisBaixarAnexo_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  let arquivo;
  try {
    arquivo = DriveApp.getFileById(String(p.driveId));
  } catch (e) {
    return { ok: false, motivo: 'NAO_ENCONTRADO' };
  }

  // Confere que o arquivo é mesmo da pasta de anexos: sem isso, um driveId
  // qualquer transformaria o Web App em leitor do seu Drive inteiro.
  if (!estaNaPasta_(arquivo)) {
    registrar_(matricula, 'MAT_BAIXAR', 'NEGADO', String(p.driveId));
    return { ok: false, motivo: 'NAO_ENCONTRADO' };
  }

  const blob = arquivo.getBlob();
  registrar_(matricula, 'MAT_BAIXAR', 'OK', arquivo.getName());

  return {
    ok: true,
    nome: arquivo.getName(),
    tipo: blob.getContentType(),
    base64: Utilities.base64Encode(blob.getBytes())
  };
}


/** { driveId, thumbId? } -> { ok } — manda para a lixeira, não apaga de vez. */
function acaoMateriaisRemoverAnexo_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  [p.driveId, p.thumbId].forEach(function (id) {
    if (!id) return;
    try {
      const f = DriveApp.getFileById(String(id));
      if (!estaNaPasta_(f)) return;
      f.setTrashed(true);
      registrar_(matricula, 'MAT_REMOVER_ANEXO', 'OK', f.getName());
    } catch (e) { /* já não existe */ }
  });
  return { ok: true };
}


/* ===========================================================================
   AUXILIARES
   =========================================================================== */

function estaNaPasta_(arquivo) {
  const alvo = pastaAnexos_().getId();
  const pais = arquivo.getParents();
  while (pais.hasNext()) {
    if (pais.next().getId() === alvo) return true;
  }
  return false;
}

function lerMeta_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_MAT_META);
  const dados = aba.getDataRange().getValues();
  const meta = { rev: 1, nextSeq: 1, columns: COLUNAS_PADRAO_QUADRO };

  for (let l = 1; l < dados.length; l++) {
    const chave = String(dados[l][0]);
    if (chave === 'rev')     meta.rev = Number(dados[l][1]) || 1;
    if (chave === 'nextSeq') meta.nextSeq = Number(dados[l][1]) || 1;
    if (chave === 'columns') {
      try { meta.columns = JSON.parse(dados[l][1]); } catch (e) { /* mantém o padrão */ }
    }
  }
  return meta;
}

/** Sobe a revisão global e grava nextSeq/columns. Devolve a revisão nova. */
function bumparMeta_(valores) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_MAT_META);
  const dados = aba.getDataRange().getValues();
  const linhaDe = {};
  for (let l = 1; l < dados.length; l++) linhaDe[String(dados[l][0])] = l + 1;

  const rev = (Number(dados[(linhaDe['rev'] || 2) - 1][1]) || 1) + 1;
  aba.getRange(linhaDe['rev'] || 2, 2).setValue(rev);

  if (valores && valores.nextSeq && linhaDe['nextSeq']) {
    aba.getRange(linhaDe['nextSeq'], 2).setValue(Number(valores.nextSeq));
  }
  if (valores && valores.columns && linhaDe['columns']) {
    aba.getRange(linhaDe['columns'], 2).setValue(JSON.stringify(valores.columns));
  }
  return rev;
}


/* ===========================================================================
   MANUTENÇÃO
   =========================================================================== */

/**
 * Apaga anexos do Drive que nenhum card referencia mais — sobras de remoções
 * que falharam no meio do caminho. Rode de vez em quando, ou num acionador mensal.
 */
function limparAnexosOrfaos() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_MAT_CARDS);
  const dados = aba.getDataRange().getValues();

  const usados = {};
  for (let l = 1; l < dados.length; l++) {
    if (!dados[l][0]) continue;
    try {
      const card = JSON.parse(dados[l][4]);
      (card.attachments || []).forEach(function (a) {
        if (a.driveId) usados[a.driveId] = true;
        if (a.thumbId) usados[a.thumbId] = true;
      });
    } catch (e) { /* linha corrompida: ignora, não apaga nada por causa dela */ }
  }

  let apagados = 0;
  const arquivos = pastaAnexos_().getFiles();
  while (arquivos.hasNext()) {
    const f = arquivos.next();
    if (!usados[f.getId()]) { f.setTrashed(true); apagados++; }
  }
  Logger.log(apagados + ' anexo(s) órfão(s) mandado(s) para a lixeira.');
}
