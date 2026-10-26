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
/* ---------------------------------------------------------------------------
   MINIATURAS

   Era aqui que a tela de materiais perdia a maior parte do tempo. Por
   miniatura eram DUAS idas ao Drive — getFileById e depois percorrer
   getParents() para conferir a pasta — mais o download e o base64. Nada era
   guardado, então um quadro com 20 anexos fazia 40 viagens ao Drive toda vez
   que a tela abria.

   Duas mudanças:

   1. Cada miniatura vai para o cache por 6 horas, numa chave própria. Miniatura
      não muda: o arquivo é gerado uma vez no upload e nunca mais é tocado.
      Da segunda abertura em diante o Drive não é consultado.

   2. A conferência de pasta saiu. No lugar, o id é comparado com os anexos que
      os próprios cards declaram. É uma ida à aba MAT_CARDS (pequena e também
      cacheada) em vez de uma ida ao Drive por arquivo — e é mais restrito que
      antes: agora só passa arquivo que algum card realmente referencia, não
      qualquer arquivo largado na pasta.
   --------------------------------------------------------------------------- */

const CACHE_MINIATURA_SEG = 21600;   // 6 h, o teto do CacheService
const CACHE_IDS_ANEXOS_SEG = 300;

/** Ids de anexo citados pelos cards: { id: 't' (miniatura) | 'o' (original de foto) | 'a' (outro anexo) }.
 *  forcar = relê a aba, ignorando o cache (foto anexada há menos de 5 min). */
function idsDeAnexosConhecidos_(forcar) {
  if (!forcar) {
    const guardado = lerGrande_(CHAVE_IDS_ANEXOS);
    if (guardado) {
      try { return JSON.parse(guardado); } catch (e) { /* cache ruim: refaz */ }
    }
  }

  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_MAT_CARDS);
  const dados = aba.getDataRange().getValues();
  const ids = {};

  for (let l = 1; l < dados.length; l++) {
    if (!dados[l][0]) continue;
    try {
      const card = JSON.parse(dados[l][4]);
      (card.attachments || []).forEach(function (a) {
        if (a.thumbId) ids[String(a.thumbId)] = 't';
        if (a.driveId) ids[String(a.driveId)] = /^image\//.test(a.type || '') ? 'o' : 'a';
      });
    } catch (e) { /* linha corrompida: ignora */ }
  }

  guardarGrande_(CHAVE_IDS_ANEXOS, JSON.stringify(ids), CACHE_IDS_ANEXOS_SEG);
  return ids;
}
/* v2: o valor passou a dizer o tipo do id (antes era só true). */
const CHAVE_IDS_ANEXOS = 'mat_ids_anexos_v2';

/* Tempo máximo gasto baixando miniaturas numa chamada. O que não couber volta
   em "pendentes" e o quadro pede de novo — melhor que estourar o limite de
   execução e perder o lote inteiro. */
const MINIATURAS_MAX_MS = 20000;

/**
 * { ids: [id] } -> { ok, mapa: { id: base64 }, tipos: { id: mime }, pendentes: [id] }
 *
 * O id pode ser:
 *  - a miniatura gerada no navegador (thumbId) — o caso normal;
 *  - o ORIGINAL de uma foto sem miniatura (ex.: HEIC do iPhone, que o
 *    navegador não consegue reduzir). Nesse caso volta a miniatura que o
 *    próprio Drive gera do arquivo; se o Drive ainda não gerou, o id vai para
 *    "pendentes" e o quadro tenta de novo mais tarde.
 */
function acaoMateriaisMiniaturas_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  const ids = (p.ids || []).slice(0, 60).map(String);
  if (!ids.length) return { ok: true, mapa: {}, tipos: {}, pendentes: [] };

  const mapa = {}, tipos = {}, pendentes = [];

  // 1. o que já está no cache sai sem tocar no Drive
  const faltando = [];
  ids.forEach(function (id) {
    const guardado = lerGrande_('thumb_' + id);
    if (guardado) {
      const sep = guardado.indexOf('|');
      // formato novo "mime|base64"; o antigo era só o base64 de um JPEG
      if (sep > 0 && sep < 40) { tipos[id] = guardado.slice(0, sep); mapa[id] = guardado.slice(sep + 1); }
      else { tipos[id] = 'image/jpeg'; mapa[id] = guardado; }
    } else faltando.push(id);
  });
  if (!faltando.length) return { ok: true, mapa: mapa, tipos: tipos, pendentes: pendentes };

  // 2. o resto é baixado, uma ida ao Drive por arquivo. Id que não está na
  //    lista guardada pode ser de foto anexada agora há pouco: relê uma vez.
  let conhecidos = idsDeAnexosConhecidos_(false);
  if (faltando.some(function (id) { return !conhecidos[id]; })) conhecidos = idsDeAnexosConhecidos_(true);

  const inicio = Date.now();
  faltando.forEach(function (id) {
    const tipoId = conhecidos[id];
    if (!tipoId || tipoId === 'a') return;                 // não é foto de nenhum card: não entrega
    if (Date.now() - inicio > MINIATURAS_MAX_MS) { pendentes.push(id); return; }
    try {
      const arquivo = DriveApp.getFileById(id);
      const blob = tipoId === 't' ? arquivo.getBlob() : arquivo.getThumbnail();
      if (!blob) { pendentes.push(id); return; }             // Drive ainda gerando a miniatura
      const mime = blob.getContentType() || 'image/jpeg';
      const b64 = Utilities.base64Encode(blob.getBytes());
      mapa[id] = b64; tipos[id] = mime;
      guardarGrande_('thumb_' + id, mime + '|' + b64, CACHE_MINIATURA_SEG);
    } catch (e) { /* arquivo sumiu do Drive: o card mostra "foto indisponível" */ }
  });

  return { ok: true, mapa: mapa, tipos: tipos, pendentes: pendentes };
}

/** Menu/manutenção: esquece a lista de ids de anexo (use se um card sumir do quadro). */
function limparCacheAnexos() {
  CacheService.getScriptCache().remove(CHAVE_IDS_ANEXOS + '_n');
}

/** Diagnóstico à mão: quantas fotos os cards têm, quantas sem miniatura e
 *  quantas apontam para arquivo que não existe mais no Drive. */
function testarFotosMateriais() {
  const ids = idsDeAnexosConhecidos_(true);
  let fotos = 0, semMini = 0, sumidos = 0;
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_MAT_CARDS);
  aba.getDataRange().getValues().slice(1).forEach(function (l) {
    if (!l[0]) return;
    try {
      (JSON.parse(l[4]).attachments || []).forEach(function (a) {
        if (!/^image\//.test(a.type || '')) return;
        fotos++;
        if (!a.thumbId) semMini++;
        const alvo = a.thumbId || a.driveId;
        try { DriveApp.getFileById(String(alvo)); } catch (e) { sumidos++; }
      });
    } catch (e) { /* linha corrompida */ }
  });
  Logger.log('Ids conhecidos: ' + Object.keys(ids).length + ' | fotos nos cards: ' + fotos +
             ' | sem miniatura: ' + semMini + ' | arquivo sumido do Drive: ' + sumidos);
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
/* ---------------------------------------------------------------------------
   ANEXO FANTASMA

   Sintoma: o card mostra o anexo com nome, tamanho e autor, mas ao baixar dá
   "versão antiga do quadro e não pode ser baixado". O que falta é o driveId —
   o endereço do arquivo no Drive.

   Como nasce: toda chamada do portal é POST, e o /exec do Apps Script responde
   com um redirecionamento. Redirecionamento de POST vira GET e perde o corpo.
   Com arquivo grande, o GET de repetição podia cair no doGet, que antes
   respondia ok:true. O servidor TINHA rodado — o arquivo foi criado no Drive e
   a linha ficou no LOG — mas o navegador leu outra resposta, sem o driveId, e
   gravou a ficha do anexo sem endereço.

   A causa está corrigida (doGet responde ok:false e a tela exige o driveId).
   Esta função é só para consertar o que já aconteceu: procura, entre os
   arquivos da pasta que nenhum card usa, um com o mesmo nome do anexo
   fantasma, e devolve o endereço para o card.

   Religa só quando há UM candidato. Dois arquivos com o mesmo nome viram
   relatório, não chute: ligar o card ao arquivo errado é pior que deixar
   quebrado.
   --------------------------------------------------------------------------- */
function repararAnexosFantasma() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_MAT_CARDS);
  const dados = aba.getDataRange().getValues();

  // ---- 1. endereços que algum card já usa
  const usados = {};
  for (let l = 1; l < dados.length; l++) {
    if (!dados[l][0]) continue;
    try {
      const card = JSON.parse(dados[l][4]);
      (card.attachments || []).forEach(function (a) {
        if (a.driveId) usados[a.driveId] = true;
        if (a.thumbId) usados[a.thumbId] = true;
      });
    } catch (e) { /* linha corrompida: ignora */ }
  }

  // ---- 2. arquivos da pasta que ninguém usa, agrupados por nome
  const porNome = {};
  const arquivos = pastaAnexos_().getFiles();
  while (arquivos.hasNext()) {
    const f = arquivos.next();
    if (usados[f.getId()]) continue;
    if (f.getName().indexOf('thumb_') === 0) continue;
    const n = f.getName();
    if (!porNome[n]) porNome[n] = [];
    porNome[n].push(f);
  }

  // ---- 3. religa o que der
  const religados = [], ambiguos = [], semArquivo = [];

  for (let l = 1; l < dados.length; l++) {
    if (!dados[l][0]) continue;
    let card;
    try { card = JSON.parse(dados[l][4]); } catch (e) { continue; }

    let mudou = false;
    (card.attachments || []).forEach(function (a) {
      if (a.driveId) return;                       // esse está inteiro
      const cand = porNome[a.name] || [];
      const onde = card.code + ' · ' + a.name;

      if (cand.length === 1) {
        a.driveId = cand[0].getId();
        usados[a.driveId] = true;
        porNome[a.name] = [];                      // não serve para outro card
        religados.push(onde);
        mudou = true;
      } else if (cand.length > 1) {
        ambiguos.push(onde + ' (' + cand.length + ' arquivos com esse nome)');
      } else {
        semArquivo.push(onde);
      }
    });

    if (mudou) {
      aba.getRange(l + 1, 5).setValue(JSON.stringify(card));
      aba.getRange(l + 1, 4).setValue((Number(dados[l][3]) || 0) + 1);   // rev
    }
  }

  if (religados.length) bumparMeta_(null);   // a tela recarrega sozinha no próximo ciclo

  const linhas = [];
  linhas.push(religados.length + ' anexo(s) religado(s).');
  religados.forEach(function (t) { linhas.push('  OK   ' + t); });
  if (ambiguos.length) {
    linhas.push('');
    linhas.push(ambiguos.length + ' com mais de um arquivo de mesmo nome — religue à mão:');
    ambiguos.forEach(function (t) { linhas.push('  ?    ' + t); });
  }
  if (semArquivo.length) {
    linhas.push('');
    linhas.push(semArquivo.length + ' sem arquivo correspondente na pasta:');
    semArquivo.forEach(function (t) { linhas.push('  --   ' + t); });
    linhas.push('(confira a lixeira do Drive — limparAnexosOrfaos pode ter levado)');
  }

  const texto = linhas.join('\n');
  Logger.log(texto);
  try { avisar_(texto); } catch (e) { /* rodando sem interface */ }
  return texto;
}


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

  // Nome citado por algum anexo SEM driveId. Um arquivo desses parece órfão mas
  // é o arquivo de um anexo fantasma — mandar para a lixeira destruiria a única
  // chance de religar o card. Veja repararAnexosFantasma, acima.
  const nomesFantasma = {};
  for (let l = 1; l < dados.length; l++) {
    if (!dados[l][0]) continue;
    try {
      const card = JSON.parse(dados[l][4]);
      (card.attachments || []).forEach(function (a) {
        if (!a.driveId && a.name) nomesFantasma[a.name] = true;
      });
    } catch (e) { /* linha corrompida: ignora */ }
  }

  let apagados = 0, poupados = 0;
  const arquivos = pastaAnexos_().getFiles();
  while (arquivos.hasNext()) {
    const f = arquivos.next();
    if (usados[f.getId()]) continue;
    if (nomesFantasma[f.getName()]) { poupados++; continue; }
    f.setTrashed(true); apagados++;
  }
  Logger.log(apagados + ' anexo(s) órfão(s) mandado(s) para a lixeira. ' +
             poupados + ' poupado(s) por serem de anexo fantasma.');
}
