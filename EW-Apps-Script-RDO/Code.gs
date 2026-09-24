/**
 * BACKEND — Relatório de Operação Diária (Extreme Wind) — versão 17
 * O PDF é GERADO AQUI (no servidor), a partir dos dados e fotos enviados
 * pelo formulário. Depois é salvo no Dropbox e as linhas vão para o Sheets.
 *
 * v17:
 *   - UM RDO por técnico por dia: chave = matrícula de quem logou + data do
 *     expediente. Reenvio APAGA as linhas antigas (Relatorios, Atividades,
 *     Funcionarios) e grava as novas, dentro de um LockService.
 *   - Nome do PDF termina com a matrícula em vez do horário, e o upload usa
 *     mode:overwrite -> o PDF do dia é substituído, não duplicado.
 *   - Aba Relatorios ganhou a coluna "Matricula_login".
 *   - Sessão de 48 h (era 12), para o técnico trabalhar offline o dia todo.
 *
 * v16: SEM MUDANÇA no backend. Igual ao 15, só para o par de versões casar.
 *   (a mudança da v16 foi só no html: Próxima atividade virou campo de busca)
 *
 * v15:
 *   - PDF: seção "Resumo da atividade" virou "Atividade realizada" e os campos
 *     ficaram pareados dois por linha.
 *
 * v14: SEM MUDANÇA no backend. Igual ao 13, só para o par de versões casar.
 *   (a correção da v14 foi só no html: rolagem das listas suspensas)
 *
 * v13:
 *   - Aba "ATV POR HR" em formato NOVO (1 linha por atividade):
 *       A = ID (ignorado) | B = Atividade | C = Tipo de reparo
 *       D = Observação obrigatória (Sim) | E = Atividade obrigatória (Sim)
 *       F = Foto obrigatória (Sim)
 *     As colunas são localizadas pelo CABEÇALHO (linha 1); se não achar,
 *     cai para as posições fixas acima.
 *     "Atividade obrigatória = Sim" -> a atividade fica disponível em TODO tipo
 *     de reparo, e o tipo dela deixa de aparecer na lista de tipos.
 *   - Campo de atividade no formulário virou busca com filtro (como o de técnico).
 *
 * v12:
 *   - Cópia do PDF por e-mail para o endereço digitado no formulário.
 *   - Mensagem de login genérica: "Matrícula ou CPF inválido".
 *   - RESUMO DE ATV é lido sempre a partir da linha 2 (linha 1 = cabeçalho).
 *   - PDF: título "Relatório de Operação Diária" e linha "Atividade realizada".
 *   ATENÇÃO: o envio de e-mail adiciona um escopo novo. Depois de colar este
 *   código é obrigatório rodar uma função à mão no editor e ACEITAR as
 *   permissões, senão o envio do RDO passa a falhar.
 *
 * v11:
 *   - BANCO DE INPUTS (planilha separada) alimenta cliente/parque, resumo de
 *     atividade, tipo de reparo e a lista de atividades por hora — sem mexer no
 *     código para mudar essas listas.
 *   - Campos novos: "Tipo de reparo" e "Reparo finalizado".
 *
 * v10:
 *   - LOGIN por matrícula + CPF, conferido AQUI (o CPF nunca sai da planilha).
 *     Login devolve um token assinado (HMAC) válido por 12 h; doPost recusa RDO
 *     sem token válido.
 *   - aba "Funcionarios" virou formato LONGO: 1 linha por técnico.
 *   - PDF: campos pareados 2 por linha; equipe um nome por linha.
 *
 * v9: integração com a MINI MASTER (planilha externa de funcionários).
 *   - doGet?lista=tecnicos  ->  devolve [{nome, mat}] para o autocomplete do form.
 *   - aba "Relatorios": coluna "Tecnicos" virou "Matriculas".
 *
 * Meus Equipamentos (set/2026): ação 'meusEquipamentos' + bloco no fim do
 *   arquivo. Lê o .xlsm do almoxarifado no Dropbox. Ver o cabeçalho do bloco.
 *
 * Propriedades do Script necessárias (Configurações do projeto):
 *   SHEET_ID, DROPBOX_APP_KEY, DROPBOX_APP_SECRET,
 *   DROPBOX_REFRESH_TOKEN, DROPBOX_FOLDER,
 *   MASTER_SHEET_ID        (ID da planilha mini master)
 * Opcionais (só se a mini master mudar de layout — os padrões já são os atuais):
 *   MASTER_ABA        padrão "EXTREME"
 *   MASTER_COL_NOME   padrão "B"  (letra ou o texto do cabeçalho)
 *   MASTER_COL_MAT    padrão "A"  (letra ou o texto do cabeçalho)
 *   MASTER_COL_CPF    padrão "H"  (letra ou o texto do cabeçalho)
 * Obrigatória a partir da v11:
 *   INPUTS_SHEET_ID   ID da planilha "Banco de inputs"
 *     abas: "PARQUE E CLIENTE" (A=parque, B=cliente)
 *           "RESUMO DE ATV"    (A=itens do resumo)
 *           "ATV POR HR"       (linha 1 = cabeçalho; 1 linha por atividade:
 *                               B=Atividade, C=Tipo de reparo,
 *                               D=Observação obrigatória, E=Atividade obrigatória,
 *                               F=Foto obrigatória)
 * Criada sozinha na 1ª execução (não mexer):
 *   LOGIN_SECRET      chave usada para assinar o token de sessão
 */

/* Layout atual da mini master (sobrescrevível por Propriedade do Script) */
var MM_ABA_PADRAO = 'EXTREME';
var MM_COL_NOME_PADRAO = 'B';
var MM_COL_MAT_PADRAO = 'A';
var MM_COL_CPF_PADRAO = 'H';

/* Validade do login (horas) e limite de tentativas por matrícula */
var SESSAO_HORAS = 720;   /* 30 dias: o tecnico nao perde o pre-preenchimento a cada 2 dias */
var LOGIN_MAX_TENTATIVAS = 8;
var LOGIN_JANELA_SEG = 600;


var MM_CACHE_KEY = 'MM_TECNICOS_V2';
var IN_CACHE_KEY = 'INPUTS_V2';

/* Abas do Banco de inputs (sobrescrevíveis por Propriedade do Script) */
var IN_ABA_PC_PADRAO = 'PARQUE E CLIENTE';
var IN_ABA_RESUMO_PADRAO = 'RESUMO DE ATV';
var IN_ABA_ATV_PADRAO = 'ATV POR HR';
var MM_CACHE_SEG = 21600; /* 6 h */

/* ===================== ENTRADAS HTTP ===================== */

function doPost(e) {
  try {
    var dados = JSON.parse(e.postData.contents);

    /* --- login --- */
    if (dados.acao === 'login') return resposta(login(dados));

    /* --- consulta do Pé-de-meia: valida o token lá dentro --- */
    if (dados.acao === 'peDeMeia') return resposta(peDeMeia(dados));

    /* --- Meus Dados: Pé de meia, Cursos e Dívidas (valida o token lá dentro) --- */
    if (dados.acao === 'consultaPessoal') return resposta(consultaPessoal(dados));

    /* --- Meus Dados: Meus Equipamentos (valida o token lá dentro) --- */
    if (dados.acao === 'meusEquipamentos') return resposta(meusEquipamentos(dados));

    /* --- checklist semanal da equipe: valida o token lá dentro --- */
    if (dados.acao === 'checklistStatus') return resposta(checklistStatus(dados));
    if (dados.acao === 'checklistFeito') return resposta(checklistFeito(dados));

    /* --- envio de RDO: exige sessão válida --- */
    var sess = validarToken(dados.token);
    if (!sess.ok) {
      return resposta({
        ok: false, sessao: false,
        erro: sess.expirado ? 'Sessão expirada. Faça login novamente.'
                            : 'Sessão inválida. Faça login novamente.'
      });
    }

    var props = PropertiesService.getScriptProperties();
    var id = gerarId();
    var matLogin = sess.mat;

    /* trava: dois envios do mesmo técnico (ex.: fila offline reenviando) não podem
       apagar/gravar ao mesmo tempo */
    var lock = LockService.getScriptLock();
    var travou = false;
    try { travou = lock.tryLock(30000); } catch (eL) { travou = false; }
    /* sem a trava o código antigo seguia assim mesmo: dois envios simultâneos do
       mesmo técnico podiam apagar as linhas um do outro no meio da gravação */
    if (!travou) {
      return resposta({
        ok: false, retentar: true,
        erro: 'Outro envio deste mesmo RDO ainda está em andamento. Tente de novo em 1 minuto.'
      });
    }

    var subst, linkPdf, pdfBlob, gravou;
    try {
      /* 1) apaga o RDO anterior do mesmo técnico na mesma data */
      subst = apagarRdoAnterior(props, matLogin, dados.data_exp);

      /* 2) gera e sobe o PDF (mesmo nome = sobrescreve o do dia) */
      pdfBlob = gerarPdf(dados, id);
      linkPdf = uploadDropbox(pdfBlob, dados, id, props, matLogin);

      /* 3) se o RDO refeito mudou de cliente/parque, o PDF antigo ficaria órfão */
      if (subst.caminhoAntigo) {
        var base = props.getProperty('DROPBOX_FOLDER') || '/Relatorios';
        var novo = montarCaminho(dados, id, base, matLogin).path;
        if (subst.caminhoAntigo !== novo) apagarDropbox(subst.caminhoAntigo, props);
      }

      /* 4) grava as linhas novas */
      gravou = gravarSheets(dados, id, linkPdf, props, matLogin);
    } finally {
      try { lock.releaseLock(); } catch (eR) {}
    }

    /* cópia por e-mail: falhar aqui NÃO invalida o RDO (já está no Dropbox e no Sheets) */
    var email = enviarCopiaEmail(dados, pdfBlob, id);

    return resposta({
      ok: true, id: id, link: linkPdf,
      gravou: gravou,
      substituiu: subst.apagou > 0, apagadas: subst.apagou,
      emailOk: email.ok, emailErro: email.erro, emailPara: email.para
    });
  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  }
}

function doGet(e) {
  var p = (e && e.parameter) || {};

  /* plano B quando o navegador não deixa ler a resposta do POST */
  if (p.acao === 'login') {
    var rl;
    try { rl = login({ mat: p.mat, cpf: p.cpf, hash: p.hash }); }
    catch (e1) { rl = { ok: false, erro: String(e1) }; }
    return saida(rl, p.callback);
  }

  /* plano B do Pé-de-meia (mesmo motivo do login: POST barrado pelo navegador) */
  /* plano B do Meus Dados (mesmo motivo do login: POST barrado pelo navegador) */
  if (p.acao === 'consultaPessoal') {
    var rc;
    try { rc = consultaPessoal({ token: p.token, tipo: p.tipo }); }
    catch (e3) { rc = { ok: false, erro: String(e3) }; }
    return saida(rc, p.callback);
  }

  if (p.acao === 'meusEquipamentos') {
    var re;
    try { re = meusEquipamentos({ token: p.token }); }
    catch (e5) { re = { ok: false, erro: String(e5) }; }
    return saida(re, p.callback);
  }

  if (p.acao === 'peDeMeia') {
    var rp;
    try { rp = peDeMeia({ token: p.token }); }
    catch (e2) { rp = { ok: false, erro: String(e2) }; }
    return saida(rp, p.callback);
  }

  /* plano B do checklist da equipe. O token vai na URL, igual ao login e ao
     Pé-de-meia — é o caminho que funciona nos aparelhos onde o navegador não
     deixa ler a resposta do POST. */
  if (p.acao === 'checklistStatus') {
    var rc;
    try { rc = checklistStatus({ token: p.token }); }
    catch (e3) { rc = { ok: false, erro: String(e3) }; }
    return saida(rc, p.callback);
  }
  if (p.acao === 'checklistFeito') {
    var rf;
    try { rf = checklistFeito({ token: p.token, checklist: p.checklist }); }
    catch (e4) { rf = { ok: false, erro: String(e4) }; }
    return saida(rf, p.callback);
  }

  if (p.lista === 'tecnicos') {
    var out;
    try {
      out = { ok: true, tecnicos: listaTecnicos() };
    } catch (err) {
      out = { ok: false, erro: String(err), tecnicos: [] };
    }
    return saida(out, p.callback);
  }

  if (p.lista === 'inputs') {
    var oi;
    try { oi = { ok: true, inputs: lerInputs() }; }
    catch (e0) { oi = { ok: false, erro: String(e0), inputs: null }; }
    return saida(oi, p.callback);
  }

  /* uma chamada só: economiza um round-trip no 4G do parque */
  if (p.lista === 'tudo') {
    var t = null, i = null, erros = [];
    try { t = listaTecnicos(); } catch (e1) { erros.push('técnicos: ' + e1); }
    try { i = lerInputs(); } catch (e2) { erros.push('inputs: ' + e2); }
    return saida({
      ok: !!(t && i),
      tecnicos: t || [],
      inputs: i,
      erro: erros.join(' | ')
    }, p.callback);
  }

  return ContentService.createTextOutput('Backend online');
}

/* JSON puro, ou JSONP quando vem ?callback= válido */
function saida(obj, callback) {
  var txt = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][A-Za-z0-9_$]{0,40}$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + txt + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

/* ===================== LOGIN E SESSÃO ===================== */

function soDigitos(v) { return String(v == null ? '' : v).replace(/[^0-9]/g, ''); }

/* Sheets pode guardar CPF/matrícula como número e comer o zero à esquerda. */
function normCpf(v) {
  var d = soDigitos(v);
  if (!d) return '';
  while (d.length < 11) d = '0' + d;
  return d;
}
function normMat(v) {
  var d = soDigitos(v);
  return d.replace(/^0+/, '') || d;
}

function segredoLogin() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('LOGIN_SECRET');
  if (!s) {
    s = Utilities.base64Encode(Utilities.getUuid() + '|' + Utilities.getUuid());
    props.setProperty('LOGIN_SECRET', s);
  }
  return s;
}

function assinar(txt) {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(txt, segredoLogin())
  );
}

function gerarToken(mat) {
  var exp = new Date().getTime() + SESSAO_HORAS * 3600 * 1000;
  var corpo = normMat(mat) + '.' + exp;
  return corpo + '.' + assinar(corpo);
}

function validarToken(tk) {
  if (!tk) return { ok: false };
  var p = String(tk).split('.');
  if (p.length !== 3) return { ok: false };
  if (assinar(p[0] + '.' + p[1]) !== p[2]) return { ok: false };
  if (Number(p[1]) < new Date().getTime()) return { ok: false, expirado: true };
  return { ok: true, mat: p[0] };
}

/* Freio simples de força bruta, por matrícula. */
function tentativasExcedidas(mat) {
  var cache;
  try { cache = CacheService.getScriptCache(); } catch (e) { return false; }
  var k = 'LOGIN_TRY_' + normMat(mat);
  var n = Number(cache.get(k) || 0) + 1;
  cache.put(k, String(n), LOGIN_JANELA_SEG);
  return n > LOGIN_MAX_TENTATIVAS;
}
function limparTentativas(mat) {
  try { CacheService.getScriptCache().remove('LOGIN_TRY_' + normMat(mat)); } catch (e) {}
}

/**
 * Confere matrícula + CPF na mini master.
 * Aceita `cpf` (dígitos) ou `hash` = SHA-256 de "matricula:cpf" em hex —
 * o hash existe para o CPF não viajar em URL no plano B por JSONP.
 * Mensagem de erro é sempre genérica: não revela quais matrículas existem.
 */
function login(dados) {
  var mat = normMat(dados && dados.mat);
  if (!mat) return { ok: false, erro: 'Informe a matrícula.' };
  if (!(dados.cpf || dados.hash)) return { ok: false, erro: 'Informe o CPF.' };

  if (tentativasExcedidas(mat)) {
    return { ok: false, erro: 'Muitas tentativas. Aguarde 10 minutos e tente de novo.' };
  }

  var achado = null;
  lerMiniMasterCompleto().forEach(function (t) {
    if (!achado && normMat(t.mat) === mat) achado = t;
  });
  var GENERICO = { ok: false, erro: 'Matrícula ou CPF inválido.' };
  if (!achado) return GENERICO;

  var cpfPlanilha = normCpf(achado.cpf);
  if (!cpfPlanilha) {
    return { ok: false, erro: 'Sua matrícula não tem CPF cadastrado na mini master. Fale com a administração.' };
  }

  var confere;
  if (dados.hash) {
    confere = (String(dados.hash).toLowerCase() === sha256Hex(mat + ':' + cpfPlanilha));
  } else {
    confere = (normCpf(dados.cpf) === cpfPlanilha);
  }
  if (!confere) return GENERICO;

  limparTentativas(mat);
  return {
    ok: true,
    nome: achado.nome,
    mat: String(achado.mat),
    token: gerarToken(mat),
    horas: SESSAO_HORAS
  };
}

function sha256Hex(txt) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, txt, Utilities.Charset.UTF_8);
  var s = '';
  for (var i = 0; i < b.length; i++) {
    var v = (b[i] < 0 ? b[i] + 256 : b[i]).toString(16);
    s += (v.length === 1 ? '0' : '') + v;
  }
  return s;
}

function resposta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function gerarId() {
  var d = new Date();
  var p = function (n) { return String(n).padStart(2, '0'); };
  return '' + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate())
    + '-' + p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds());
}

/* ===================== MINI MASTER (nome <-> matrícula) ===================== */

/** Lista PÚBLICA (vai para o celular): só nome e matrícula. Nunca CPF. */
function listaTecnicos() {
  return lerMiniMasterCompleto().map(function (t) {
    return { nome: t.nome, mat: t.mat };
  });
}

/** Lista INTERNA (fica no servidor): inclui CPF, usada só pelo login. */
function lerMiniMasterCompleto() {
  var cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) { /* sem cache */ }

  if (cache) {
    var hit = cache.get(MM_CACHE_KEY);
    if (hit) {
      try {
        var j = JSON.parse(hit);
        if (j && j.length) return j;
      } catch (e2) { /* cache corrompido: relê */ }
    }
  }

  var lista = lerMiniMaster();

  if (cache) {
    try { cache.put(MM_CACHE_KEY, JSON.stringify(lista), MM_CACHE_SEG); } catch (e3) { /* > 100 KB: segue sem cache */ }
  }
  return lista;
}

/** Força releitura da mini master (rodar à mão depois de editar a planilha). */
function limparCacheTecnicos() {
  try { CacheService.getScriptCache().remove(MM_CACHE_KEY); } catch (e) {}
  var n = lerMiniMasterCompleto().length;
  Logger.log('Cache limpo. Técnicos carregados: ' + n);
  return n;
}

/** Diagnóstico: roda à mão e mostra o que foi detectado. */
function testarMiniMaster() {
  var l = lerMiniMaster();
  var comCpf = 0;
  l.forEach(function (t) { if (normCpf(t.cpf)) comCpf++; });
  Logger.log('Total de técnicos: ' + l.length);
  Logger.log('Com CPF cadastrado: ' + comCpf + '  |  SEM CPF (não conseguem logar): ' + (l.length - comCpf));
  /* mostra só os 4 últimos dígitos do CPF — não joga CPF inteiro no log */
  Logger.log(JSON.stringify(l.slice(0, 10).map(function (t) {
    var c = normCpf(t.cpf);
    return { nome: t.nome, mat: t.mat, cpf: c ? ('•••.•••.' + c.slice(6, 9) + '-' + c.slice(9)) : 'SEM CPF' };
  }), null, 2));
  return l.length;
}

/** Lista quem está sem CPF na mini master (esses não conseguem entrar). */
function tecnicosSemCpf() {
  var faltam = lerMiniMaster().filter(function (t) { return !normCpf(t.cpf); })
    .map(function (t) { return t.mat + ' - ' + t.nome; });
  Logger.log(faltam.length ? ('Sem CPF (' + faltam.length + '):\n' + faltam.join('\n')) : 'Todos têm CPF.');
  return faltam;
}

function lerMiniMaster() {
  var props = PropertiesService.getScriptProperties();
  var idMaster = props.getProperty('MASTER_SHEET_ID');
  if (!idMaster) throw new Error('Propriedade MASTER_SHEET_ID não configurada.');

  var ss = SpreadsheetApp.openById(idMaster);
  var nomeAba = (props.getProperty('MASTER_ABA') || MM_ABA_PADRAO).trim();
  var sh = ss.getSheetByName(nomeAba) || ss.getSheets()[0];
  if (!sh) throw new Error('Aba "' + nomeAba + '" não encontrada na mini master.');

  var valores = sh.getDataRange().getValues();
  if (!valores.length) return [];

  var map = acharColunas(
    valores,
    props.getProperty('MASTER_COL_NOME') || MM_COL_NOME_PADRAO,
    props.getProperty('MASTER_COL_MAT') || MM_COL_MAT_PADRAO
  );
  var iCpf = letraParaIndice(props.getProperty('MASTER_COL_CPF') || MM_COL_CPF_PADRAO);
  if (map.nome < 0) {
    throw new Error('Não achei a coluna do NOME na aba "' + sh.getName()
      + '". Configure MASTER_COL_NOME (letra ou texto do cabeçalho).');
  }
  if (map.mat < 0) {
    throw new Error('Não achei a coluna da MATRICULA na aba "' + sh.getName()
      + '". Configure MASTER_COL_MAT (letra ou texto do cabeçalho).');
  }

  var out = [], vistos = {};
  for (var r = map.linhaCab + 1; r < valores.length; r++) {
    var nome = limpar(valores[r][map.nome]);
    var mat = limpar(valores[r][map.mat]);
    if (!nome) continue;
    /* chave = matrícula quando existe. Se fosse por nome, um homônimo (ou
       recontratação com matrícula nova) sumiria da lista e não conseguiria logar. */
    var k = mat ? ('M:' + normMat(mat)) : ('N:' + chaveNome(nome));
    if (vistos[k]) continue;          /* duplicado na mini master: fica o 1º */
    vistos[k] = true;
    var cpf = (iCpf >= 0 && iCpf < valores[r].length) ? limpar(valores[r][iCpf]) : '';
    out.push({ nome: nome, mat: mat, cpf: cpf });
  }

  out.sort(function (a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
  return out;
}

function limpar(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return String(v).replace(/\s+/g, ' ').trim();
}

/** chave para comparar nomes: minusculo, sem acento, espaco simples */
function chaveNome(s) {
  var t = limpar(s).toLowerCase();
  try { t = t.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) { /* engine sem normalize */ }
  return t;
}

function letraParaIndice(s) {
  var t = String(s || '').trim().toUpperCase();
  if (!/^[A-Z]{1,2}$/.test(t)) return -1;
  var n = 0;
  for (var i = 0; i < t.length; i++) n = n * 26 + (t.charCodeAt(i) - 64);
  return n - 1;
}

/* Um cabeçalho é uma célula CURTA que COMEÇA com a palavra-chave.
   Evita confundir títulos tipo "RELAÇÃO DE FUNCIONÁRIOS - 2026" com cabeçalho. */
function pareceCabNome(cab) {
  var t = chaveNome(cab);
  return t.length <= 30 && /^(nome|tecnico|funcionario|colaborador|empregado)\b/.test(t);
}
function pareceCabMat(cab) {
  var t = chaveNome(cab);
  return t.length <= 30 && /^(matricula|mat|re|registro)\b/.test(t);
}

/**
 * Resolve os índices das colunas nome/matrícula e a linha de cabeçalho.
 * Prioridade para cada coluna: letra ("B") -> texto exato do cabeçalho -> palavra-chave.
 * linhaCab = -1 significa "a aba não tem cabeçalho, lê desde a 1ª linha".
 */
function acharColunas(valores, cfgNome, cfgMat) {
  var iN = letraParaIndice(cfgNome);
  var iM = letraParaIndice(cfgMat);
  var limite = Math.min(valores.length, 15);
  var linhaCab = -1;

  /* 1) resolve por texto exato do cabeçalho ou por palavra-chave */
  for (var r = 0; r < limite && (iN < 0 || iM < 0); r++) {
    var linha = valores[r];
    for (var c = 0; c < linha.length; c++) {
      var cab = limpar(linha[c]);
      if (!cab) continue;
      if (iN < 0 && ((cfgNome && chaveNome(cab) === chaveNome(cfgNome)) || (!cfgNome && pareceCabNome(cab)))) {
        iN = c; linhaCab = r;
      }
      if (iM < 0 && ((cfgMat && chaveNome(cab) === chaveNome(cfgMat)) || (!cfgMat && pareceCabMat(cab)))) {
        iM = c; linhaCab = r;
      }
    }
  }

  /* 2) colunas vieram por letra: procura a linha de cabeçalho só nessas colunas.
        Prefere a linha em que AS DUAS parecem cabeçalho. */
  if (linhaCab < 0 && iN >= 0 && iM >= 0) {
    var candidata = -1;
    for (var r2 = 0; r2 < limite; r2++) {
      var vN = limpar(valores[r2][iN]), vM = limpar(valores[r2][iM]);
      var okN = pareceCabNome(vN) || pareceCabMat(vN);
      var okM = pareceCabMat(vM) || pareceCabNome(vM);
      if (okN && okM) { candidata = r2; break; }
      if (candidata < 0 && (okN || okM)) candidata = r2;
    }
    linhaCab = candidata;
  }

  return { linhaCab: linhaCab, nome: iN, mat: iM };
}

/* ===================== BANCO DE INPUTS ===================== */

/**
 * Lê a planilha "Banco de inputs" e devolve tudo que o formulário precisa:
 *   { clientes:[], parques:{cliente:[parques]}, resumo:[], reparos:[{tipo, atividades:[{nome, exec}]}] }
 * Cache de 6 h, igual à lista de técnicos.
 */
function lerInputs() {
  var cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) {}
  if (cache) {
    var hit = cache.get(IN_CACHE_KEY);
    if (hit) { try { var j = JSON.parse(hit); if (j && j.reparos && j.comuns) return j; } catch (e2) {} }
  }

  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('INPUTS_SHEET_ID');
  if (!id) throw new Error('Propriedade INPUTS_SHEET_ID não configurada.');
  var ss = SpreadsheetApp.openById(id);

  var atv = lerAbaAtvPorHora(ss, props);
  var out = {
    clientes: [],
    parques: {},
    resumo: lerAbaResumo(ss, props),
    reparos: atv.reparos,
    comuns: atv.comuns
  };
  var pc = lerAbaParqueCliente(ss, props);
  out.clientes = pc.clientes;
  out.parques = pc.parques;

  if (cache) { try { cache.put(IN_CACHE_KEY, JSON.stringify(out), MM_CACHE_SEG); } catch (e3) {} }
  return out;
}

function abaInputs(ss, props, prop, padrao) {
  var nome = (props.getProperty(prop) || padrao).trim();
  var sh = ss.getSheetByName(nome);
  if (!sh) throw new Error('Aba "' + nome + '" não encontrada no Banco de inputs.');
  return sh;
}

/** true se a 1ª linha parece cabeçalho (para pular) */
function ehCabecalho(celulas, palavras) {
  for (var i = 0; i < celulas.length; i++) {
    var t = chaveNome(celulas[i]);
    for (var j = 0; j < palavras.length; j++) {
      if (t === palavras[j]) return true;
    }
  }
  return false;
}

function lerAbaParqueCliente(ss, props) {
  var sh = abaInputs(ss, props, 'INPUTS_ABA_PC', IN_ABA_PC_PADRAO);
  var v = sh.getDataRange().getValues();
  var ini = (v.length && ehCabecalho([v[0][0], v[0][1]], ['parque', 'cliente'])) ? 1 : 0;

  var parques = {}, clientes = [], vistoCli = {};
  for (var r = ini; r < v.length; r++) {
    var parque = limpar(v[r][0]);
    var cliente = limpar(v[r][1]);
    if (!parque || !cliente) continue;
    if (!vistoCli[cliente]) { vistoCli[cliente] = true; clientes.push(cliente); }
    if (!parques[cliente]) parques[cliente] = [];
    if (parques[cliente].indexOf(parque) < 0) parques[cliente].push(parque);
  }
  clientes.sort(function (a, b) { return a.localeCompare(b, 'pt-BR'); });
  Object.keys(parques).forEach(function (c) {
    parques[c].sort(function (a, b) { return a.localeCompare(b, 'pt-BR'); });
  });
  if (!clientes.length) throw new Error('Aba "' + sh.getName() + '" sem nenhum par parque/cliente preenchido.');
  return { clientes: clientes, parques: parques };
}

function lerAbaResumo(ss, props) {
  var sh = abaInputs(ss, props, 'INPUTS_ABA_RESUMO', IN_ABA_RESUMO_PADRAO);
  var v = sh.getDataRange().getValues();
  /* a linha 1 é SEMPRE o cabeçalho da tabela: as opções começam na linha 2 */
  var out = [], visto = {};
  for (var r = 1; r < v.length; r++) {
    var t = limpar(v[r][0]);
    if (!t || visto[chaveNome(t)]) continue;
    visto[chaveNome(t)] = true;
    out.push(t);
  }
  if (!out.length) {
    throw new Error('Aba "' + sh.getName() + '" sem opções a partir da linha 2 '
      + '(a linha 1 é tratada como cabeçalho).');
  }
  return out;
}

/** "Sim" em qualquer caixa/acento. Vazio, "Não", "-" => false. */
function ehSim(v) {
  var t = chaveNome(v);
  return t === 's' || t === 'sim' || t === 'x' || t === 'true' || t === 'verdadeiro';
}

/** Acha uma coluna pelo cabeçalho; se não achar, usa a posição padrão. */
function colPorCabecalho(cab, testa, padrao) {
  for (var c = 0; c < cab.length; c++) {
    if (testa(chaveNome(cab[c]))) return c;
  }
  return padrao;
}

/**
 * ATV POR HR — 1 linha por atividade, cabeçalho na linha 1.
 *   B = Atividade | C = Tipo de reparo | D = Observação obrigatória
 *   E = Atividade obrigatória | F = Foto obrigatória
 * Colunas localizadas pelo cabeçalho, com as posições acima como reserva.
 *
 * Devolve { reparos:[{tipo, atividades:[{nome, exec, obs}]}], comuns:[{nome, exec, obs}] }
 * - "Atividade obrigatória = Sim" joga a atividade em `comuns`: ela aparece em
 *   todo tipo de reparo.
 * - Um tipo cujas atividades sejam TODAS comuns não entra na lista de tipos
 *   (é um agrupador, não um reparo).
 */
function lerAbaAtvPorHora(ss, props) {
  var sh = abaInputs(ss, props, 'INPUTS_ABA_ATV', IN_ABA_ATV_PADRAO);
  var v = sh.getDataRange().getValues();
  if (v.length < 2) {
    throw new Error('Aba "' + sh.getName() + '" sem dados a partir da linha 2.');
  }

  var cab = v[0].map(function (x) { return limpar(x); });
  var iAtv = colPorCabecalho(cab, function (t) {
    return t.indexOf('atividade') === 0 && t.indexOf('obrigat') < 0;
  }, 1);                                                   /* B */
  var iTipo = colPorCabecalho(cab, function (t) {
    return t.indexOf('tipo') >= 0 && t.indexOf('reparo') >= 0;
  }, 2);                                                   /* C */
  var iObs = colPorCabecalho(cab, function (t) {
    return t.indexOf('observ') >= 0;
  }, 3);                                                   /* D */
  var iObrig = colPorCabecalho(cab, function (t) {
    return t.indexOf('atividade') >= 0 && t.indexOf('obrigat') >= 0;
  }, 4);                                                   /* E */
  var iFoto = colPorCabecalho(cab, function (t) {
    return t.indexOf('foto') >= 0;
  }, -1);                                                  /* F — só se existir */

  var grupos = {}, ordem = [], comuns = [], vistoComum = {};

  for (var r = 1; r < v.length; r++) {
    var linha = v[r];
    var nome = limpar(linha[iAtv]);
    var tipo = limpar(linha[iTipo]);
    if (!nome) continue;

    var item = {
      nome: nome,
      exec: (iFoto >= 0) ? ehSim(linha[iFoto]) : false,
      obs: ehSim(linha[iObs])
    };
    var comum = ehSim(linha[iObrig]);

    if (comum) {
      var kc = chaveNome(nome);
      if (!vistoComum[kc]) { vistoComum[kc] = true; comuns.push(item); }
    }

    if (!tipo) continue;            /* atividade comum sem tipo: só em `comuns` */
    if (!grupos[tipo]) { grupos[tipo] = { tipo: tipo, atividades: [], visto: {}, soComuns: true }; ordem.push(tipo); }
    var g = grupos[tipo];
    var k = chaveNome(nome);
    if (g.visto[k]) continue;
    g.visto[k] = true;
    if (!comum) {
      g.soComuns = false;
      g.atividades.push(item);
    }
  }

  var reparos = [];
  ordem.forEach(function (t) {
    var g = grupos[t];
    if (g.soComuns) return;         /* agrupador de atividades comuns: não é tipo de reparo */
    if (!g.atividades.length) return;
    reparos.push({ tipo: g.tipo, atividades: g.atividades });
  });

  if (!reparos.length) {
    throw new Error('Aba "' + sh.getName() + '": nenhum tipo de reparo selecionável. '
      + 'Confira a coluna "Tipo de reparo" e se todas as linhas não estão marcadas '
      + 'como "Atividade obrigatória".');
  }
  return { reparos: reparos, comuns: comuns };
}

/** true se nenhuma atividade exige foto (provável coluna F ausente/vazia) */
function iFotoAusente(i) {
  var achou = i.comuns.some(function (a) { return a.exec; });
  if (achou) return false;
  return !i.reparos.some(function (rp) {
    return rp.atividades.some(function (a) { return a.exec; });
  });
}

/** Força releitura do Banco de inputs (rodar após editar a planilha). */
function limparCacheInputs() {
  try { CacheService.getScriptCache().remove(IN_CACHE_KEY); } catch (e) {}
  var i = lerInputs();
  Logger.log('Cache limpo. Clientes: ' + i.clientes.length
    + ' | Itens de resumo: ' + i.resumo.length
    + ' | Tipos de reparo: ' + i.reparos.length
    + ' | Atividades comuns: ' + i.comuns.length);
  return i;
}

/** Diagnóstico do Banco de inputs — rodar à mão. */
function testarInputs() {
  var i = lerInputs();
  Logger.log('CLIENTES (' + i.clientes.length + '): ' + i.clientes.join(', '));
  i.clientes.forEach(function (c) {
    Logger.log('  ' + c + ' -> ' + (i.parques[c] || []).length + ' parque(s): ' + (i.parques[c] || []).join(', '));
  });
  Logger.log('RESUMO (' + i.resumo.length + '): ' + i.resumo.join(', '));
  Logger.log('TIPOS DE REPARO SELECIONÁVEIS (' + i.reparos.length + '):');
  i.reparos.forEach(function (rp) {
    var ex = rp.atividades.filter(function (a) { return a.exec; }).length;
    var ob = rp.atividades.filter(function (a) { return a.obs; }).length;
    Logger.log('  ' + rp.tipo + ' -> ' + rp.atividades.length + ' atividade(s) próprias | '
      + ex + ' com foto | ' + ob + ' com observação obrigatória');
  });
  Logger.log('ATIVIDADES COMUNS, aparecem em TODO tipo (' + i.comuns.length + '):');
  i.comuns.forEach(function (a) {
    Logger.log('  ' + a.nome + (a.exec ? ' [foto]' : '') + (a.obs ? ' [observação]' : ''));
  });
  if (!i.comuns.length) {
    Logger.log('  (nenhuma) — se o Almoço/Janta não estiver em cada tipo de reparo, '
      + 'os técnicos daquele tipo não conseguirão enviar o RDO.');
  }
  var temAlmoco = i.comuns.some(function (a) { return /almoc|janta/.test(chaveNome(a.nome)); });
  if (!temAlmoco) {
    var faltam = i.reparos.filter(function (rp) {
      return !rp.atividades.some(function (a) { return /almoc|janta/.test(chaveNome(a.nome)); });
    }).map(function (rp) { return rp.tipo; });
    if (faltam.length) {
      Logger.log('ATENÇÃO: sem Almoço/Janta nestes tipos de reparo (o envio ficará bloqueado): '
        + faltam.join(', '));
    }
  }
  if (iFotoAusente(i)) {
    Logger.log('AVISO: nenhuma atividade com "Foto obrigatória = Sim". '
      + 'Confira se a coluna F existe e está preenchida.');
  }
  return i;
}

/* ===================== EQUIPE (normalização) ===================== */

/**
 * Aceita os dois formatos para não quebrar html antigo:
 *   ["João", "Maria"]                          (v8)
 *   [{nome:"João", mat:"123"}, ...]            (v9)
 * Devolve sempre [{nome, mat}]. Sem limite de quantidade.
 */
function normEquipe(tecnicos) {
  var out = [];
  (tecnicos || []).forEach(function (t) {
    if (t === null || t === undefined) return;
    if (typeof t === 'string') {
      if (t.trim()) out.push({ nome: limpar(t), mat: '' });
    } else {
      var nome = limpar(t.nome);
      if (nome) out.push({ nome: nome, mat: limpar(t.mat) });
    }
  });
  return out;
}

function equipeNomes(eq) {
  return eq.map(function (t) { return t.nome; });
}

function equipeMatriculas(eq) {
  return eq.map(function (t) { return t.mat || 'N/A'; });
}

function equipeNomeMat(eq) {
  return eq.map(function (t) { return t.mat ? (t.nome + ' (' + t.mat + ')') : t.nome; });
}

/* ===================== GERAÇÃO DO PDF ===================== */

function gerarPdf(d, id) {
  var html = montarHtml(d);
  var blob = Utilities.newBlob(html, 'text/html', 'r.html').getAs('application/pdf');
  return blob;
}

/* monta pastas {Cliente}/{Parque}/{MM-AAAA} e o nome RDO_..._ddmmaaaa_ID.pdf */
/**
 * Pastas {Cliente}/{Parque}/{MM-AAAA} e nome do arquivo.
 * A partir da v17 o nome termina com a MATRÍCULA de quem logou (não com o
 * horário), para que refazer o RDO do dia sobrescreva o mesmo arquivo.
 */
function montarCaminho(dados, id, base, matLogin) {
  function sanit(s){ return String(s==null?'':s).trim().replace(/\s+/g, '-'); }
  var p = String(dados.data_exp || '').split('-'); // yyyy-mm-dd
  var yyyy = p[0] || '', mm = p[1] || '', dd = p[2] || '';
  var ddmmaaaa = dd + ' ' + mm + ' ' + yyyy;
  var mesFolder = (mm && yyyy) ? (mm + '-' + yyyy) : 'sem-mes';

  var cliente = dados.cliente || 'SEM-CLIENTE';
  var parque = dados.parque || 'SEM-PARQUE';
  var sufixo = matLogin ? ('MAT' + normMat(matLogin)) : String(id);

  var filename = 'RDO_' + sanit(cliente) + '_' + sanit(parque) + '_' + ddmmaaaa + '_' + sufixo + '.pdf';
  var path = base + '/' + cliente + '/' + parque + '/' + mesFolder + '/' + filename;
  return { path: path, filename: filename };
}

function montarHtml(d) {
  function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
  function brData(iso){ if(!iso) return '—'; var p=String(iso).split('-'); return p[2]+'/'+p[1]+'/'+p[0]; }
  function arr(a){ return (a&&a.length)? a.map(esc).join(' &bull; ') : '—'; }
  var CEL_L = 'color:#64748B;font-size:11px;padding:5px 7px;border-bottom:1px solid #E2E8F0;vertical-align:top;';
  var CEL_V = 'font-size:12.5px;font-weight:bold;padding:5px 7px;border-bottom:1px solid #E2E8F0;vertical-align:top;';
  /* linha inteira: rótulo + valor ocupando as 3 colunas restantes */
  function linha(l,v){
    return '<tr><td style="width:20%;' + CEL_L + '">' + l + '</td>'
      + '<td colspan="3" style="' + CEL_V + '">' + (v||'—') + '</td></tr>';
  }
  /* dois pares por linha (economiza espaço no PDF) */
  function linha2(l1,v1,l2,v2){
    return '<tr>'
      + '<td style="width:20%;' + CEL_L + '">' + l1 + '</td>'
      + '<td style="width:30%;' + CEL_V + '">' + (v1||'—') + '</td>'
      + '<td style="width:20%;' + CEL_L + '">' + l2 + '</td>'
      + '<td style="width:30%;' + CEL_V + '">' + (v2||'—') + '</td>'
      + '</tr>';
  }
  function sec(titulo, conteudo){
    return '<div style="page-break-inside:avoid;margin-top:16px;">'
      + '<h3 style="margin:0 0 6px 0;font-size:13px;color:#3B5A8A;border-bottom:2px solid #D6E0EC;padding-bottom:4px;">'
      + titulo + '</h3>' + conteudo + '</div>';
  }

  var equipe = normEquipe(d.tecnicos);

  var atv = (d.atividades||[]).filter(function(a){return a.tipo||a.ini||a.fim||a.obs;});
  var atvRows = atv.length ? atv.map(function(a){
    return '<tr><td style="padding:6px 8px;border-bottom:1px solid #E2E8F0;font-size:12px;">'
      + esc(a.ini||'—') + ' - ' + esc(a.fim||'—')
      + '</td><td style="padding:6px 8px;border-bottom:1px solid #E2E8F0;font-size:12px;font-weight:bold;">'
      + esc(a.tipo||'—')
      + '</td><td style="padding:6px 8px;border-bottom:1px solid #E2E8F0;font-size:12px;">'
      + esc(a.obs||'') + '</td></tr>';
  }).join('') : '<tr><td colspan="3" style="padding:8px;color:#94A3B8;font-size:12px;">Sem registros</td></tr>';

  var imgs = (d.imagens&&d.imagens.length) ? d.imagens.map(function(o){
    var src = (o && o.img) ? o.img : o;
    var nome = (o && o.nome) ? o.nome : '';
    if(!src) return '';
    return '<div style="display:inline-block;width:31%;vertical-align:top;margin:0 1% 8px 0;">'
      + '<img src="' + src + '" style="width:100%;border:1px solid #D6E0EC;">'
      + (nome ? '<div style="font-size:10px;color:#64748B;margin-top:2px;">'+esc(nome)+'</div>' : '')
      + '</div>';
  }).join('') : '<span style="color:#94A3B8;font-size:12px;">Sem fotos</span>';

  var assinatura = d.assinatura
    ? '<img src="' + d.assinatura + '" style="height:90px;border-bottom:1px solid #94A3B8;">'
    : '<div style="color:#94A3B8;font-size:12px;">Não assinado</div>';

  var lider = equipe.length ? equipeNomeMat(equipe)[0] : '';

  return ''
  + '<html><head><meta charset="utf-8"></head>'
  + '<body style="font-family:Arial,Helvetica,sans-serif;color:#1E293B;padding:28px;">'

  + '<table style="width:100%;border-bottom:3px solid #3B5A8A;padding-bottom:8px;"><tr>'
  + '<td style="font-size:20px;font-weight:bold;color:#26374F;">EXTREME WIND '
  + '<span style="font-size:13px;color:#64748B;font-weight:normal;">Blade Services</span><br>'
  + '<span style="font-size:14px;color:#3B5A8A;">Relatório de Operação Diária</span></td>'
  + '<td style="text-align:right;font-size:11px;color:#64748B;">Registrado em<br><b style="color:#26374F;">'
  + esc(d.registrado||'') + '</b></td>'
  + '</tr></table>'

  + sec('Identificação',
      '<table style="width:100%;border-collapse:collapse;table-layout:fixed;">'
      + linha2('Data do expediente', brData(d.data_exp),
               'Horário', esc(d.hora_ini||'—')+' às '+esc(d.hora_fim||'—'))
      + linha2('Cliente', esc(d.cliente), 'Parque', esc(d.parque))
      + linha('Equipe', equipeNomeMat(equipe).map(esc).join('<br>') || '—')
      + linha('E-mail responsável', esc(d.email))
      + '</table>')

  + sec('Local e máquina',
      '<table style="width:100%;border-collapse:collapse;table-layout:fixed;">'
      + linha('Local de atividade', esc(d.local))
      + linha2('Turbina (WTG)', esc(d.turbina), 'Blade', esc(d.blade))
      + linha2('Parada da WTG', esc(d.parada), 'WTG posto em marcha', esc(d.marcha))
      + linha2('Fibra-on', esc(d.fibraon), 'Fibra-off', esc(d.fibraoff))
      + '</table>')

  + sec('Atividade realizada',
      '<table style="width:100%;border-collapse:collapse;table-layout:fixed;">'
      + linha2('Atividade realizada', arr(d.resumo), 'Tipo de reparo', esc(d.tipo_reparo))
      + linha2('Avanço do reparo', ((d.avanco!=null && d.avanco!=='') ? esc(d.avanco)+'%' : 'N/A'),
               'Reparo finalizado', (d.finalizado === true || d.finalizado === 'SIM') ? 'SIM' : 'NÃO')
      + '</table>')

  + sec('Atividades por hora',
      '<table style="width:100%;border-collapse:collapse;">'
      + '<tr style="background:#F4F7FB;">'
      + '<th style="text-align:left;padding:6px 8px;font-size:11px;color:#3B5A8A;">Horário</th>'
      + '<th style="text-align:left;padding:6px 8px;font-size:11px;color:#3B5A8A;">Atividade</th>'
      + '<th style="text-align:left;padding:6px 8px;font-size:11px;color:#3B5A8A;">Observação</th></tr>'
      + atvRows + '</table>')

  + sec('Próxima atividade', '<div style="font-size:13px;">' + arr(d.proxima) + '</div>')

  + sec('Registro fotográfico', '<div>' + imgs + '</div>')

  + sec('Assinatura', assinatura
      + '<div style="font-size:11px;color:#64748B;margin-top:4px;">' + esc(lider) + '</div>')

  + '</body></html>';
}

/* ===================== E-MAIL ===================== */

function dataBr(iso) {
  var p = String(iso || '').split('-');
  return (p.length === 3) ? (p[2] + '/' + p[1] + '/' + p[0]) : String(iso || '');
}

function emailValido(e) {
  return /^[^@\s]+@[^@\s]+\.[^@\s]{2,}$/.test(String(e || '').trim());
}

/**
 * Manda o PDF para o e-mail digitado no formulário.
 * Nunca lança: devolve {ok, erro, para} para o front avisar sem travar o envio.
 * Cota do Gmail: ~100 e-mails/dia em conta comum, 1.500 em Workspace.
 */
function enviarCopiaEmail(dados, pdfBlob, id) {
  var para = String(dados.email || '').trim();
  if (!emailValido(para)) return { ok: false, erro: 'E-mail inválido: ' + para, para: para };

  try {
    var props = PropertiesService.getScriptProperties();
    var base = props.getProperty('DROPBOX_FOLDER') || '/Relatorios';
    var nome = montarCaminho(dados, id, base).filename;

    var equipe = equipeNomeMat(normEquipe(dados.tecnicos));
    var parque = dados.parque || '';
    var assunto = 'RDO ' + parque + ' — ' + dataBr(dados.data_exp);

    var corpo = ''
      + '<div style="font-family:Arial,Helvetica,sans-serif;color:#1E293B;font-size:14px;">'
      + '<p>Segue em anexo o Relatório de Operação Diária.</p>'
      + '<table style="border-collapse:collapse;font-size:14px;">'
      + linhaEmail('Data do expediente', dataBr(dados.data_exp))
      + linhaEmail('Horário', (dados.hora_ini || '—') + ' às ' + (dados.hora_fim || '—'))
      + linhaEmail('Cliente', dados.cliente || '—')
      + linhaEmail('Parque', parque || '—')
      + linhaEmail('Tipo de reparo', dados.tipo_reparo || '—')
      + linhaEmail('Avanço do reparo', (dados.avanco || dados.avanco === 0) ? (dados.avanco + '%') : '—')
      + linhaEmail('Reparo finalizado', (dados.finalizado === true || dados.finalizado === 'SIM') ? 'SIM' : 'NÃO')
      + linhaEmail('Equipe', equipe.join('<br>') || '—')
      + linhaEmail('Nº do relatório', id)
      + '</table>'
      + '<p style="color:#64748B;font-size:12px;margin-top:18px;">'
      + 'Mensagem automática do sistema de RDO — Extreme Wind Blade Services. Não responda a este e-mail.'
      + '</p></div>';

    /* copia o blob para não alterar o nome do que já foi enviado ao Dropbox */
    var anexo = pdfBlob.copyBlob().setName(nome);

    MailApp.sendEmail({
      to: para,
      subject: assunto,
      htmlBody: corpo,
      attachments: [anexo],
      name: 'RDO — Extreme Wind'
    });
    return { ok: true, erro: '', para: para };
  } catch (e) {
    return { ok: false, erro: String(e), para: para };
  }
}

function linhaEmail(rot, val) {
  return '<tr><td style="color:#64748B;padding:3px 12px 3px 0;vertical-align:top;">' + rot + '</td>'
    + '<td style="font-weight:bold;padding:3px 0;">' + (val || '—') + '</td></tr>';
}

/** Teste de e-mail — rodar à mão e ACEITAR as permissões na 1ª vez. */
function testarEmail() {
  var quota = MailApp.getRemainingDailyQuota();
  Logger.log('E-mails restantes hoje: ' + quota);
  var eu = Session.getActiveUser().getEmail();
  MailApp.sendEmail({
    to: eu,
    subject: 'Teste RDO — permissão de e-mail OK',
    htmlBody: '<p>Se você recebeu isto, o envio de e-mail do RDO está autorizado.</p>',
    name: 'RDO — Extreme Wind'
  });
  Logger.log('Enviado para ' + eu);
  return quota;
}

/* ===================== DROPBOX ===================== */

function getDropboxToken(props) {
  var res = UrlFetchApp.fetch('https://api.dropbox.com/oauth2/token', {
    method: 'post',
    payload: {
      grant_type: 'refresh_token',
      refresh_token: props.getProperty('DROPBOX_REFRESH_TOKEN'),
      client_id: props.getProperty('DROPBOX_APP_KEY'),
      client_secret: props.getProperty('DROPBOX_APP_SECRET')
    },
    muteHttpExceptions: true
  });
  var j = JSON.parse(res.getContentText());
  if (!j.access_token) throw new Error('Dropbox auth falhou: ' + res.getContentText());
  return j.access_token;
}

function escaparArg(obj) {
  return JSON.stringify(obj).replace(/[\u007f-\uffff]/g, function (c) {
    return '\\u' + ('0000' + c.charCodeAt(0).toString(16)).slice(-4);
  });
}


function uploadDropbox(blob, dados, id, props, matLogin) {
  var token = getDropboxToken(props);
  var base = props.getProperty('DROPBOX_FOLDER') || '/Relatorios';
  var caminho = montarCaminho(dados, id, base, matLogin);
  var path = caminho.path;

  /* overwrite: refazer o RDO do dia substitui o arquivo, não cria outro */
  var up = UrlFetchApp.fetch('https://content.dropboxapi.com/2/files/upload', {
    method: 'post',
    contentType: 'application/octet-stream',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Dropbox-API-Arg': escaparArg({ path: path, mode: 'overwrite', autorename: false, mute: true })
    },
    payload: blob.getBytes(),
    muteHttpExceptions: true
  });
  if (up.getResponseCode() >= 300) {
    throw new Error('Upload Dropbox falhou: ' + up.getContentText());
  }
  var meta = JSON.parse(up.getContentText());
  var pathReal = meta.path_display || path;

  var link = '';
  try {
    var r = UrlFetchApp.fetch('https://api.dropboxapi.com/2/sharing/create_shared_link_with_settings', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify({ path: pathReal }),
      muteHttpExceptions: true
    });
    var j = JSON.parse(r.getContentText());
    if (j.url) link = j.url;
    else if (j.error && j.error['.tag'] === 'shared_link_already_exists')
      link = j.error.shared_link_already_exists.metadata.url;
  } catch (e2) { /* link é opcional */ }

  return link;
}

/** Apaga um arquivo do Dropbox. Nunca lança. */
function apagarDropbox(path, props) {
  if (!path) return false;
  try {
    var token = getDropboxToken(props);
    var r = UrlFetchApp.fetch('https://api.dropboxapi.com/2/files/delete_v2', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify({ path: path }),
      muteHttpExceptions: true
    });
    return r.getResponseCode() < 300;
  } catch (e) { return false; }
}

/* ===================== SHEETS ===================== */

function gravarSheets(dados, id, linkPdf, props, matLogin) {
  var ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));

  function na(v) {
    if (v === null || v === undefined) return 'N/A';
    if (typeof v === 'string' && v.trim() === '') return 'N/A';
    return v;
  }

  var equipe = normEquipe(dados.tecnicos);

  /* Abas resolvidas ANTES de escrever qualquer coisa: se a aba Atividades
     tivesse sumido/sido renomeada, o código antigo gravava Relatorios e
     Funcionarios e só então estourava em atv.appendRow — sobrava meio RDO. */
  var rel = acharAbaFlex(ss, 'Relatorios');
  if (!rel) throw new Error('A aba "Relatorios" não existe na planilha.');
  var fun = acharAbaFlex(ss, 'Funcionarios') || criarAbaFuncionarios(ss);
  var atv = acharAbaFlex(ss, 'Atividades');
  if (!atv) {
    atv = criarAbaAtividades(ss);
    Logger.log('AVISO: aba "Atividades" nao existia e foi criada agora.');
  }

  /* --- Relatorios: coluna Tecnicos virou Matriculas --- */
  var linhaRel = [
    id,
    na(dados.registrado), na(dados.data_exp), na(dados.hora_ini), na(dados.hora_fim),
    na(dados.cliente), na(dados.parque), na(equipeMatriculas(equipe).join(', ')), na(dados.email),
    na(dados.local), na(dados.turbina), na(dados.blade), na(dados.parada), na(dados.marcha),
    na(dados.fibraon), na(dados.fibraoff), na((dados.resumo || []).join('; ')),
    na((dados.proxima || []).join('; ')), na(dados.avanco),
    na(dados.tipo_reparo), (dados.finalizado === true || dados.finalizado === 'SIM') ? 'SIM' : 'NÃO',
    na(matLogin ? normMat(matLogin) : (equipe[0] ? equipe[0].mat : '')),
    na(linkPdf)
  ];

  /* --- Funcionarios: formato longo, 1 linha por técnico --- */
  var linhasFun = equipe.map(function (t) {
    return [
      id, na(dados.parque), na(dados.data_exp), na(dados.hora_ini), na(dados.hora_fim),
      na(t.nome), na(t.mat)
    ];
  });

  /* --- Atividades --- */
  var linhasAtv = (dados.atividades || []).map(function (a, i) {
    return [
      id, na(dados.parque), na(dados.data_exp), na(dados.turbina), na(dados.blade),
      i + 1, na(a.ini), na(a.fim), na(a.tipo), na(a.obs)
    ];
  });

  /* Uma chamada por aba em vez de uma por linha: um RDO de 12 atividades saía
     em 13 round-trips e era candidato a estourar o tempo no meio. */
  var gravou = { rel: 0, fun: 0, atv: 0 };
  try {
    gravou.rel = escreverBloco(rel, [linhaRel]);
    gravou.fun = escreverBloco(fun, linhasFun);
    gravou.atv = escreverBloco(atv, linhasAtv);
  } catch (e) {
    /* nunca deixar meio RDO na planilha: desfaz o que já entrou */
    try { apagarLinhasPorId(rel, [id]); } catch (e1) {}
    try { apagarLinhasPorId(fun, [id]); } catch (e2) {}
    try { apagarLinhasPorId(atv, [id]); } catch (e3) {}
    throw new Error('Falha ao gravar na planilha: ' + e
      + ' (nada ficou pela metade, mas o RDO NAO foi salvo)');
  }
  return gravou;
}

/**
 * Acha a aba pelo nome. Tenta o nome exato e, se não achar, compara ignorando
 * maiúsculas, acentos e espaços sobrando — uma aba renomeada para "ATIVIDADES"
 * ou "Atividades " (com espaço no fim) faz getSheetByName devolver null e a
 * gravação estourar. Devolve null se realmente não existir.
 */
function acharAbaFlex(ss, nome) {
  var exata = ss.getSheetByName(nome);
  if (exata) return exata;
  var alvo = chaveAba(nome);
  var achada = null;
  ss.getSheets().forEach(function (sh) {
    if (!achada && chaveAba(sh.getName()) === alvo) achada = sh;
  });
  if (achada) {
    Logger.log('AVISO: aba "' + nome + '" foi encontrada como "' + achada.getName()
      + '". Renomeie para "' + nome + '" exatamente.');
  }
  return achada;
}

function chaveAba(t) {
  t = String(t == null ? '' : t).trim().toLowerCase();
  try { t = t.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) {}
  return t.replace(/\s+/g, ' ');
}

/**
 * Escreve várias linhas de uma vez, logo abaixo da última linha preenchida.
 * Cresce a aba se faltar linha ou coluna — setValues estoura se o retângulo
 * não couber, ao contrário de appendRow.
 */
function escreverBloco(sh, linhas) {
  if (!sh || !linhas || !linhas.length) return 0;
  var largura = linhas[0].length;
  if (sh.getMaxColumns() < largura) {
    sh.insertColumnsAfter(sh.getMaxColumns(), largura - sh.getMaxColumns());
  }
  var primeira = sh.getLastRow() + 1;
  var precisa = primeira + linhas.length - 1;
  if (sh.getMaxRows() < precisa) sh.insertRowsAfter(sh.getMaxRows(), precisa - sh.getMaxRows());
  sh.getRange(primeira, 1, linhas.length, largura).setValues(linhas);
  return linhas.length;
}

/** Cria a aba Atividades com o cabeçalho padrão (mesma lógica de Funcionarios). */
function criarAbaAtividades(ss) {
  var sh = ss.insertSheet('Atividades');
  sh.appendRow(['Relatorio_ID', 'Parque', 'Data_exp', 'Turbina', 'Blade',
                'Ordem', 'Hora_ini', 'Hora_fim', 'Atividade', 'Observacao']);
  sh.setFrozenRows(1);
  return sh;
}

/**
 * Diagnóstico manual. Rode no editor do Apps Script e veja o log:
 * mostra se as três abas existem, quantas linhas cada uma tem e se há
 * relatório sem atividade (o sintoma que apareceu).
 */
function diagnosticoAbas() {
  var ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));

  Logger.log('--- ABAS DA PLANILHA (entre colchetes, para ver espaco sobrando) ---');
  ss.getSheets().forEach(function (sh) {
    Logger.log('  [' + sh.getName() + ']  linhas=' + sh.getLastRow()
      + '  colunas=' + sh.getMaxColumns() + (sh.isSheetHidden() ? '  (OCULTA)' : ''));
  });

  Logger.log('--- ABAS QUE O RDO PRECISA ---');
  ['Relatorios', 'Funcionarios', 'Atividades'].forEach(function (n) {
    var exata = ss.getSheetByName(n);
    var flex = acharAbaFlex(ss, n);
    if (exata) { Logger.log(n + ': OK (nome exato)'); return; }
    if (flex) { Logger.log(n + ': *** NOME ERRADO -> esta como [' + flex.getName() + '] ***'); return; }
    Logger.log(n + ': *** NAO EXISTE ***');
  });

  Logger.log('--- PROTECOES (podem bloquear a gravacao) ---');
  ss.getSheets().forEach(function (sh) {
    var ps = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .concat(sh.getProtections(SpreadsheetApp.ProtectionType.RANGE));
    if (ps.length) Logger.log('  [' + sh.getName() + '] tem ' + ps.length + ' protecao(oes)');
  });

  Logger.log('--- TESTE DE ESCRITA REAL NA ABA ATIVIDADES ---');
  var teste = acharAbaFlex(ss, 'Atividades');
  if (!teste) {
    Logger.log('  impossivel testar: aba nao encontrada');
  } else {
    try {
      var linha = teste.getLastRow() + 1;
      teste.getRange(linha, 1, 1, 10).setValues([['TESTE_DIAGNOSTICO', '', '', '', '', '', '', '', '', '']]);
      SpreadsheetApp.flush();
      teste.deleteRow(linha);
      Logger.log('  escrita OK (linha de teste gravada e apagada)');
    } catch (eT) {
      Logger.log('  *** ESCRITA FALHOU: ' + eT + ' ***');
    }
  }

  Logger.log('--- RELATORIOS SEM ATIVIDADES ---');
  var rel = acharAbaFlex(ss, 'Relatorios');
  var atv = acharAbaFlex(ss, 'Atividades');
  if (!rel || !atv || rel.getLastRow() < 2) return;

  var ids = {};
  atv.getRange(1, 1, atv.getLastRow(), 1).getValues().forEach(function (r) { ids[String(r[0])] = true; });

  var v = rel.getDataRange().getValues();
  var iData = idxCabecalho(rel, 'Data_exp');
  var orfaos = 0;
  for (var r = 1; r < v.length; r++) {
    if (!ids[String(v[r][0])]) {
      orfaos++;
      Logger.log('SEM ATIVIDADES -> id ' + v[r][0] + ' | data ' + (iData >= 0 ? v[r][iData] : '?'));
    }
  }
  Logger.log('Relatorios sem nenhuma linha em Atividades: ' + orfaos);
}

/* ===================== UM RDO POR TÉCNICO POR DIA ===================== */

/** Data em texto yyyy-MM-dd, venha como Date, "yyyy-mm-dd" ou "dd/mm/aaaa". */
function normData(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(v == null ? '' : v).trim();
  var m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return m[3] + '-' + m[2] + '-' + m[1];
  return s;
}

function idxCabecalho(sh, nome) {
  var cab = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
  return cab.indexOf(nome);
}

/**
 * Apaga o RDO anterior do mesmo técnico (matrícula de quem logou) na mesma
 * data de expediente, nas três abas. Devolve:
 *   { apagou: n, ids: [...], caminhoAntigo: '/...pdf' | '' }
 * `caminhoAntigo` é reconstruído a partir do cliente/parque/data da linha
 * antiga — serve para apagar o PDF órfão se o RDO refeito trocou de parque.
 */
function apagarRdoAnterior(props, matLogin, dataExp) {
  var vazio = { apagou: 0, ids: [], caminhoAntigo: '' };
  var mat = normMat(matLogin);
  var data = normData(dataExp);
  if (!mat || !data) return vazio;

  var ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  var rel = ss.getSheetByName('Relatorios');
  if (!rel || rel.getLastRow() < 2) return vazio;

  var iMat = idxCabecalho(rel, 'Matricula_login');
  var iData = idxCabecalho(rel, 'Data_exp');
  var iCli = idxCabecalho(rel, 'Cliente');
  var iParq = idxCabecalho(rel, 'Parque');
  if (iMat < 0 || iData < 0) return vazio;   /* planilha não migrada: não apaga nada */

  var v = rel.getDataRange().getValues();
  var linhas = [], ids = [], caminho = '';
  var base = props.getProperty('DROPBOX_FOLDER') || '/Relatorios';

  for (var r = 1; r < v.length; r++) {
    if (normMat(v[r][iMat]) !== mat) continue;
    if (normData(v[r][iData]) !== data) continue;
    linhas.push(r + 1);
    ids.push(String(v[r][0]));
    if (!caminho && iCli >= 0 && iParq >= 0) {
      caminho = montarCaminho(
        { cliente: v[r][iCli], parque: v[r][iParq], data_exp: normData(v[r][iData]) },
        '', base, mat
      ).path;
    }
  }
  if (!linhas.length) return vazio;

  /* de baixo para cima, senão os índices mudam no meio do caminho */
  linhas.sort(function (a, b) { return b - a; });
  linhas.forEach(function (n) { rel.deleteRow(n); });

  apagarLinhasPorId(acharAbaFlex(ss, 'Atividades'), ids);
  apagarLinhasPorId(acharAbaFlex(ss, 'Funcionarios'), ids);

  return { apagou: linhas.length, ids: ids, caminhoAntigo: caminho };
}

/** Apaga de uma aba todas as linhas cuja 1ª coluna esteja na lista de IDs. */
function apagarLinhasPorId(sh, ids) {
  if (!sh || !ids || !ids.length || sh.getLastRow() < 2) return 0;
  var mapa = {};
  ids.forEach(function (i) { mapa[String(i)] = true; });

  var v = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  var linhas = [];
  for (var r = 1; r < v.length; r++) {
    if (mapa[String(v[r][0])]) linhas.push(r + 1);
  }
  linhas.sort(function (a, b) { return b - a; });
  linhas.forEach(function (n) { sh.deleteRow(n); });
  return linhas.length;
}

/**
 * Diagnóstico: procura RDO do MESMO dia/parque/turbina/blade enviados por
 * matrículas DIFERENTES. É o cenário que duplica as atividades sem ninguém
 * ter enviado duas vezes: dois técnicos da mesma equipe mandaram cada um o seu.
 * apagarRdoAnterior não pega isso, porque a chave dele é matrícula + data.
 */
function acharRdoDaMesmaEquipe(diasAtras) {
  var ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  var rel = acharAbaFlex(ss, 'Relatorios');
  if (!rel || rel.getLastRow() < 2) { Logger.log('Relatorios vazia'); return; }

  var iMat = idxCabecalho(rel, 'Matricula_login');
  var iData = idxCabecalho(rel, 'Data_exp');
  var iParq = idxCabecalho(rel, 'Parque');
  var iTurb = idxCabecalho(rel, 'Turbina');
  var iBlade = idxCabecalho(rel, 'Blade');
  if (iMat < 0 || iData < 0) { Logger.log('cabecalhos ausentes'); return; }

  var corte = '';
  if (diasAtras) {
    var dt = new Date(); dt.setDate(dt.getDate() - diasAtras);
    corte = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  var v = rel.getDataRange().getValues();
  var grupos = {};
  for (var r = 1; r < v.length; r++) {
    var data = normData(v[r][iData]);
    if (corte && data < corte) continue;
    var k = [data, v[r][iParq], v[r][iTurb], v[r][iBlade]].join(' | ');
    if (!grupos[k]) grupos[k] = [];
    grupos[k].push({ id: v[r][0], mat: normMat(v[r][iMat]), linha: r + 1 });
  }

  var achou = 0;
  Object.keys(grupos).sort().forEach(function (k) {
    var g = grupos[k];
    if (g.length < 2) return;
    achou++;
    Logger.log('>>> ' + k);
    g.forEach(function (x) {
      Logger.log('      id ' + x.id + '  matricula ' + x.mat + '  (linha ' + x.linha + ')');
    });
  });
  Logger.log(achou ? ('\nGrupos com mais de um RDO: ' + achou)
                   : 'Nenhum caso encontrado no periodo.');
}

/**
 * Limpeza: apaga de Atividades e Funcionarios as linhas cujo Relatorio_ID
 * nao existe mais em Relatorios. Sao as linhas orfas deixadas pela gravacao
 * concorrente da versao antiga. Rode com (true) para apagar de verdade;
 * sem argumento so lista o que seria apagado.
 */
function limparOrfaos(apagarDeVerdade) {
  var ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  var rel = acharAbaFlex(ss, 'Relatorios');
  if (!rel || rel.getLastRow() < 2) { Logger.log('Relatorios vazia'); return; }

  var vivos = {};
  rel.getRange(2, 1, rel.getLastRow() - 1, 1).getValues()
     .forEach(function (r) { vivos[String(r[0])] = true; });

  ['Atividades', 'Funcionarios'].forEach(function (nome) {
    var sh = acharAbaFlex(ss, nome);
    if (!sh || sh.getLastRow() < 2) return;
    var v = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
    var linhas = [], ids = {};
    for (var r = 1; r < v.length; r++) {
      var id = String(v[r][0]);
      if (!id || vivos[id]) continue;
      linhas.push(r + 1);
      ids[id] = (ids[id] || 0) + 1;
    }
    Logger.log(nome + ': ' + linhas.length + ' linha(s) orfa(s) em '
      + Object.keys(ids).length + ' id(s)');
    Object.keys(ids).forEach(function (i) { Logger.log('    ' + i + '  (' + ids[i] + ' linhas)'); });
    if (apagarDeVerdade && linhas.length) {
      linhas.sort(function (a, b) { return b - a; });
      linhas.forEach(function (n) { sh.deleteRow(n); });
      Logger.log('    -> APAGADAS');
    }
  });
  if (!apagarDeVerdade) Logger.log('\nModo lista. Rode limparOrfaos(true) para apagar.');
}

/** Diagnóstico: mostra RDO duplicados (mesma matrícula + data) já existentes. */
function acharRdoDuplicados() {
  var ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  var rel = ss.getSheetByName('Relatorios');
  var iMat = idxCabecalho(rel, 'Matricula_login');
  var iData = idxCabecalho(rel, 'Data_exp');
  if (iMat < 0) { Logger.log('Rode migrarParaV17() primeiro.'); return []; }

  var v = rel.getDataRange().getValues(), conta = {}, dups = [];
  for (var r = 1; r < v.length; r++) {
    var k = normMat(v[r][iMat]) + '|' + normData(v[r][iData]);
    if (!normMat(v[r][iMat])) continue;
    conta[k] = (conta[k] || 0) + 1;
  }
  Object.keys(conta).forEach(function (k) {
    if (conta[k] > 1) dups.push(k + '  -> ' + conta[k] + ' RDO');
  });
  Logger.log(dups.length
    ? ('Matrícula|data com mais de um RDO (linhas antigas, anteriores à v17):\n' + dups.join('\n'))
    : 'Nenhum duplicado de matrícula+data.');
  return dups;
}

/* ===================== UTIL DE PLANILHA ===================== */

function cabecalhoFuncionarios() {
  return ['ID', 'Parque', 'Data_exp', 'Hora_ini', 'Hora_fim', 'Nome', 'Matricula'];
}

function criarAbaFuncionarios(ss) {
  var sh = ss.insertSheet('Funcionarios');
  sh.appendRow(cabecalhoFuncionarios());
  sh.setFrozenRows(1);
  return sh;
}

/**
 * MIGRAÇÃO v8 -> v9. Rodar UMA VEZ, à mão, no editor do Apps Script.
 * NÃO apaga dado nenhum:
 *   - cria a aba "Funcionarios" com cabeçalho (se não existir);
 *   - renomeia o cabeçalho da coluna H de "Tecnicos" para "Matriculas".
 * As linhas antigas de Relatorios continuam com NOMES na coluna H —
 * a partir da 1ª linha nova, essa coluna passa a ter MATRÍCULAS.
 */
function migrarParaV9() {
  var ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  var log = [];

  var rel = ss.getSheetByName('Relatorios');
  if (rel) {
    var cab = rel.getRange(1, 1, 1, rel.getLastColumn()).getValues()[0];
    var iTec = cab.indexOf('Tecnicos');
    if (iTec >= 0) {
      rel.getRange(1, iTec + 1).setValue('Matriculas');
      log.push('Relatorios: cabeçalho "Tecnicos" -> "Matriculas" (coluna ' + (iTec + 1) + ').');
    } else if (cab.indexOf('Matriculas') >= 0) {
      log.push('Relatorios: já estava como "Matriculas".');
    } else {
      log.push('ATENÇÃO: não achei "Tecnicos" nem "Matriculas" no cabeçalho de Relatorios.');
    }
  } else {
    log.push('ATENÇÃO: aba Relatorios não existe.');
  }

  if (!ss.getSheetByName('Funcionarios')) {
    criarAbaFuncionarios(ss);
    log.push('Aba "Funcionarios" criada com cabeçalho.');
  } else {
    log.push('Aba "Funcionarios" já existia (cabeçalho não foi tocado).');
  }

  Logger.log(log.join('\n'));
  return log.join('\n');
}

/**
 * MIGRAÇÃO v9 -> v10. Rodar UMA VEZ, à mão.
 * A aba `Funcionarios` muda de formato (largo -> longo), então ela é
 * RECRIADA VAZIA com o cabeçalho novo. As abas Relatorios e Atividades
 * não são tocadas.
 */
function migrarParaV10() {
  var ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  var log = [];

  var fun = ss.getSheetByName('Funcionarios');
  if (fun) {
    var linhasAntes = Math.max(0, fun.getLastRow() - 1);
    fun.clear();
    fun.appendRow(cabecalhoFuncionarios());
    fun.setFrozenRows(1);
    log.push('Aba "Funcionarios" recriada no formato longo (' + linhasAntes + ' linha(s) do formato antigo apagadas).');
  } else {
    criarAbaFuncionarios(ss);
    log.push('Aba "Funcionarios" criada no formato longo.');
  }

  /* garante o segredo do token já na migração, para o 1º login não gerar corrida */
  segredoLogin();
  log.push('LOGIN_SECRET pronto nas Propriedades do Script.');

  log.push('Relatorios e Atividades: não alteradas.');
  Logger.log(log.join('\n'));
  return log.join('\n');
}

/**
 * MIGRAÇÃO v10 -> v11. Rodar UMA VEZ, à mão.
 * Insere as colunas "Tipo_reparo" e "Reparo_finalizado" na aba Relatorios,
 * logo ANTES de "Link_PDF". Insere colunas de verdade: as linhas antigas
 * continuam alinhadas e ficam com essas duas células em branco.
 * Não apaga nada. Rodar duas vezes é seguro.
 */
function migrarParaV11() {
  var ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  var log = [];

  var rel = ss.getSheetByName('Relatorios');
  if (!rel) {
    Logger.log('ATENÇÃO: aba Relatorios não existe.');
    return 'ATENÇÃO: aba Relatorios não existe.';
  }

  var cab = rel.getRange(1, 1, 1, rel.getLastColumn()).getValues()[0];
  if (cab.indexOf('Tipo_reparo') >= 0) {
    log.push('Relatorios: colunas novas já existiam.');
  } else {
    var iLink = cab.indexOf('Link_PDF');
    if (iLink < 0) {
      log.push('ATENÇÃO: não achei a coluna "Link_PDF"; as colunas novas foram para o fim.');
      var fim = rel.getLastColumn();
      rel.getRange(1, fim + 1, 1, 2).setValues([['Tipo_reparo', 'Reparo_finalizado']]);
    } else {
      rel.insertColumnsBefore(iLink + 1, 2);
      rel.getRange(1, iLink + 1, 1, 2).setValues([['Tipo_reparo', 'Reparo_finalizado']]);
      log.push('Relatorios: "Tipo_reparo" e "Reparo_finalizado" inseridas antes de "Link_PDF" '
        + '(colunas ' + (iLink + 1) + ' e ' + (iLink + 2) + '). Linhas antigas ficam em branco nessas células.');
    }
  }

  /* aquece o Banco de inputs para o erro aparecer aqui, e não no celular */
  try {
    var i = lerInputs();
    log.push('Banco de inputs OK: ' + i.clientes.length + ' cliente(s), '
      + i.resumo.length + ' item(ns) de resumo, ' + i.reparos.length + ' tipo(s) de reparo.');
  } catch (e) {
    log.push('ATENÇÃO — Banco de inputs NÃO carregou: ' + e);
  }

  Logger.log(log.join('\n'));
  return log.join('\n');
}

/**
 * MIGRAÇÃO v16 -> v17. Rodar UMA VEZ, à mão.
 * Insere a coluna "Matricula_login" em Relatorios, antes de "Link_PDF".
 * Insere coluna de verdade: as linhas antigas continuam alinhadas e ficam com
 * essa célula em branco (por isso a regra de substituição só vale para RDO
 * enviados a partir da v17).
 */
function migrarParaV17() {
  var ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  var log = [];
  var rel = ss.getSheetByName('Relatorios');
  if (!rel) { Logger.log('ATENÇÃO: aba Relatorios não existe.'); return 'sem Relatorios'; }

  var cab = rel.getRange(1, 1, 1, rel.getLastColumn()).getValues()[0];
  if (cab.indexOf('Matricula_login') >= 0) {
    log.push('Relatorios: coluna "Matricula_login" já existia.');
  } else {
    var iLink = cab.indexOf('Link_PDF');
    if (iLink < 0) {
      var fim = rel.getLastColumn();
      rel.getRange(1, fim + 1).setValue('Matricula_login');
      log.push('ATENÇÃO: não achei "Link_PDF"; a coluna nova foi para o fim.');
    } else {
      rel.insertColumnsBefore(iLink + 1, 1);
      rel.getRange(1, iLink + 1).setValue('Matricula_login');
      log.push('Relatorios: "Matricula_login" inserida antes de "Link_PDF" (coluna ' + (iLink + 1) + ').');
    }
  }
  log.push('Linhas antigas ficam com Matricula_login em branco: a substituição '
    + 'automática passa a valer só para os RDO enviados de agora em diante.');
  log.push('Sessão do login agora vale ' + SESSAO_HORAS + ' h.');
  Logger.log(log.join('\n'));
  return log.join('\n');
}

/** Setup de planilha NOVA. CUIDADO: apaga o conteúdo das abas. */
function criarCabecalhos() {
  var ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));

  var rel = ss.getSheetByName('Relatorios') || ss.insertSheet('Relatorios');
  rel.clear();
  rel.appendRow([
    'ID', 'Registrado_em', 'Data_exp', 'Hora_ini', 'Hora_fim',
    'Cliente', 'Parque', 'Matriculas', 'Email', 'Local',
    'Turbina', 'Blade', 'Parada_WTG', 'WTG_marcha', 'Fibra_on', 'Fibra_off',
    'Resumo', 'Proxima_atividade', 'Avanco_reparo',
    'Tipo_reparo', 'Reparo_finalizado', 'Matricula_login', 'Link_PDF'
  ]);

  var atv = ss.getSheetByName('Atividades') || ss.insertSheet('Atividades');
  atv.clear();
  atv.appendRow(['Relatorio_ID', 'Parque', 'Data_exp', 'Turbina', 'Blade', 'Ordem', 'Hora_ini', 'Hora_fim', 'Atividade', 'Observacao']);

  var fun = ss.getSheetByName('Funcionarios') || ss.insertSheet('Funcionarios');
  fun.clear();
  fun.appendRow(cabecalhoFuncionarios());
  fun.setFrozenRows(1);
}

/* ===================== CONSULTAS PESSOAIS (Meus Dados) =====================
 * Três abas da MESMA planilha, mesma mecânica:
 *
 *   Pé de meia  → busca por CPF,       1 linha por pessoa
 *   Cursos      → busca por MATRÍCULA, N linhas por pessoa (um curso cada)
 *   Dívidas     → busca por MATRÍCULA, N linhas por pessoa (um motivo cada)
 *
 * Quem faz a ligação é sempre o backend: o token é assinado e carrega a
 * matrícula; a matrícula acha o técnico na mini master; daí sai o CPF (para o
 * Pé de meia) ou a própria matrícula (para Cursos e Dívidas). O CPF nunca vai
 * nem volta pelo navegador, e sem sessão válida ninguém consulta nada.
 *
 * Propriedades do Script (opcionais — os padrões já apontam para o lugar certo):
 *   PEDEMEIA_SHEET_ID   ID da planilha (as três abas moram nela)
 *   PEDEMEIA_ABA        nome da aba do Pé de meia
 *   CURSOS_ABA          nome da aba de Cursos
 *   DIVIDAS_ABA         nome da aba de Dívidas
 *
 * A conta que publicou este Apps Script precisa ter acesso de LEITURA à
 * planilha, senão a consulta devolve erro de permissão.
 * =========================================================================== */

var PDM_SHEET_ID_PADRAO = '1R8CXpCRJUQt39hvWKaTw-ILVmlWdnYF90IttgvdzEpk';
var PDM_CACHE_SEG = 300;            /* 5 min — valor de dinheiro não pode ficar velho */
var PDM_COL_DATA_PADRAO = 'H';      /* onde mora "Valor atualizado em:" */
var PDM_ROTULO_DATA = 'VALOR ATUALIZADO EM';

/* Os três campos de valor de Cursos e Dívidas são os mesmos. */
var CP_CAMPOS_DEVEDOR = [
  { chave: 'valor', rotulo: 'Valor',         alvo: 'VALOR' },
  { chave: 'pago',  rotulo: 'Valor pago',    alvo: 'VALOR PAGO' },
  { chave: 'saldo', rotulo: 'Saldo devedor', alvo: 'SALDO DEVEDOR' }
];

/**
 * O catálogo das consultas. Para acrescentar uma quarta aba amanhã, basta uma
 * entrada aqui e um HTML novo — nada mais no backend precisa mudar.
 *
 *   chaveBusca   'cpf' ou 'mat'
 *   multiplas    true = a pessoa pode ter várias linhas
 *   titulo       coluna que nomeia cada linha (só faz sentido com multiplas)
 *   campos       colunas de valor, na ordem em que aparecem na tela
 *   destaque     chave do campo que sai grande no cartão
 */
var CP_CONSULTAS = {
  peDeMeia: {
    prop: 'PEDEMEIA_ABA', aba: 'Pé de meia',
    chaveBusca: 'cpf', multiplas: false, titulo: null,
    destaque: 'total',
    campos: [
      { chave: 'acumulado', rotulo: 'Acumulado anterior',   alvo: 'ACUMULADO ANTERIOR' },
      { chave: 'semana',    rotulo: 'Pé-de-meia da semana', alvo: 'PE DE MEIA DA SEMANA' },
      { chave: 'ticket',    rotulo: 'Pé-de-meia ticket',    alvo: 'PE DE MEIA TICKET' },
      { chave: 'total',     rotulo: 'Total guardado',       alvo: 'TOTAL GUARDADO' }
    ]
  },
  cursos: {
    prop: 'CURSOS_ABA', aba: 'Cursos',
    chaveBusca: 'mat', multiplas: true,
    titulo: { alvo: 'CURSO', rotulo: 'Curso' },
    destaque: 'saldo',
    campos: CP_CAMPOS_DEVEDOR
  },
  dividas: {
    prop: 'DIVIDAS_ABA', aba: 'Dívidas',
    chaveBusca: 'mat', multiplas: true,
    titulo: { alvo: 'MOTIVO', rotulo: 'Motivo' },
    destaque: 'saldo',
    campos: CP_CAMPOS_DEVEDOR
  }
};

/** Normaliza cabeçalho: sem acento, MAIÚSCULO, qualquer pontuação vira espaço.
 *  É o que faz "PÉ-DE-MEIA DA SEMANA" casar com "PE DE MEIA DA SEMANA". */
function pdmNorm(v) {
  var t = (v === null || v === undefined) ? '' : String(v);
  try { t = t.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) {}
  return t.toUpperCase().replace(/[^A-Z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
}

/** Aceita número puro ou texto tipo "R$ 1.234,56" / "1234.56". */
function pdmNum(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  var s = String(v).replace(/[^0-9,.\-]/g, '');
  if (!s || s === '-') return null;
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  var n = Number(s);
  return isFinite(n) ? n : null;
}

function pdmMoeda(n) {
  if (n === null || n === undefined) return null;
  var neg = n < 0, v = Math.abs(Number(n)).toFixed(2).split('.');
  var inteiro = v[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (neg ? '-R$ ' : 'R$ ') + inteiro + ',' + v[1];
}

function pdmData(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim();
}

/** Acha a aba pelo nome ignorando acento e caixa — "Dívidas", "DÍvidas" e
 *  "dividas" chegam todas na mesma aba. Nome exato tem prioridade. */
function cpAcharAba(ss, nome) {
  var exata = ss.getSheetByName(nome);
  if (exata) return exata;
  var alvo = pdmNorm(nome), abas = ss.getSheets();
  for (var i = 0; i < abas.length; i++) {
    if (pdmNorm(abas[i].getName()) === alvo) return abas[i];
  }
  return null;
}

/**
 * Casa os cabeçalhos em DOIS passes: primeiro igualdade exata, depois
 * "contém" só para o que sobrou, pulando coluna já usada.
 * O pass duplo não é preciosismo: "VALOR" é pedaço de "VALOR PAGO", então um
 * pass só de substring poderia grudar o campo Valor na coluna do Valor pago,
 * dependendo da ordem das colunas.
 */
function cpCasarCabecalho(linha, alvos) {
  var achados = {}, usadas = {}, i, c, h;
  for (c = 0; c < linha.length; c++) {
    h = pdmNorm(linha[c]);
    if (!h) continue;
    for (i = 0; i < alvos.length; i++) {
      if (achados[alvos[i].chave] === undefined && h === alvos[i].alvo) {
        achados[alvos[i].chave] = c; usadas[c] = true;
      }
    }
  }
  for (c = 0; c < linha.length; c++) {
    if (usadas[c]) continue;
    h = pdmNorm(linha[c]);
    if (!h) continue;
    for (i = 0; i < alvos.length; i++) {
      if (achados[alvos[i].chave] === undefined && h.indexOf(alvos[i].alvo) >= 0) {
        achados[alvos[i].chave] = c; usadas[c] = true; break;
      }
    }
  }
  return achados;
}

/** Lê uma aba inteira e devolve já mapeada. Cache curto de 5 min. */
function cpLerTabela(tipo) {
  var cfg = CP_CONSULTAS[tipo];
  if (!cfg) throw new Error('Consulta desconhecida: ' + tipo);

  var chaveCache = 'CP_' + tipo + '_V2';
  var cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) {}
  if (cache) {
    var g = cache.get(chaveCache);
    if (g) { try { return JSON.parse(g); } catch (e2) {} }
  }

  var props = PropertiesService.getScriptProperties();
  var id = (props.getProperty('PEDEMEIA_SHEET_ID') || PDM_SHEET_ID_PADRAO).trim();
  var ss = SpreadsheetApp.openById(id);
  var nomeAba = (props.getProperty(cfg.prop) || cfg.aba).trim();
  var sh = cpAcharAba(ss, nomeAba);
  if (!sh) {
    throw new Error('Aba "' + nomeAba + '" não encontrada. Abas da planilha: '
      + ss.getSheets().map(function (s) { return s.getName(); }).join(', '));
  }

  var valores = sh.getDataRange().getValues();

  /* ---- acha a linha de cabeçalho ----
     É a primeira que tenha a coluna-chave (CPF ou MAT) e pelo menos dois dos
     campos de valor. Nada é por posição fixa: planilha de administração muda
     de lugar, e um cabeçalho que desce uma linha não pode derrubar a tela. */
  var alvoChave = (cfg.chaveBusca === 'cpf')
    ? [{ chave: 'k', alvo: 'CPF' }]
    : [{ chave: 'k', alvo: 'MAT' }, { chave: 'k', alvo: 'MATRICULA' }];
  var alvosValor = cfg.campos.slice();
  if (cfg.titulo) alvosValor = alvosValor.concat([{ chave: '_titulo', alvo: cfg.titulo.alvo }]);

  var linhaCab = -1, iChave = -1, cols = {};
  var limite = Math.min(valores.length, 40);
  for (var r = 0; r < limite && linhaCab < 0; r++) {
    var k = cpCasarCabecalho(valores[r], alvoChave);
    if (k.k === undefined) continue;
    var achados = cpCasarCabecalho(valores[r], alvosValor);
    var quantos = 0;
    for (var q = 0; q < cfg.campos.length; q++) {
      if (achados[cfg.campos[q].chave] !== undefined) quantos++;
    }
    if (quantos >= 2) { linhaCab = r; iChave = k.k; cols = achados; }
  }
  if (linhaCab < 0) {
    throw new Error('Não achei o cabeçalho na aba "' + sh.getName() + '". Preciso de uma linha com a coluna '
      + (cfg.chaveBusca === 'cpf' ? 'CPF' : 'MAT') + ' e os títulos dos valores.');
  }

  /* ---- "Valor atualizado em:" — procura na coluna H, depois no resto ---- */
  var iColData = letraParaIndice(PDM_COL_DATA_PADRAO);
  var rotuloData = '', dataAtualizacao = '';
  var busca = [iColData];
  for (var cc = 0; cc < 30; cc++) if (cc !== iColData) busca.push(cc);
  for (var b = 0; b < busca.length && !rotuloData; b++) {
    var col = busca[b];
    if (col < 0) continue;
    for (var rr = 0; rr < valores.length; rr++) {
      if (col >= valores[rr].length) continue;
      if (pdmNorm(valores[rr][col]).indexOf(PDM_ROTULO_DATA) === 0) {
        rotuloData = String(valores[rr][col]).replace(/\s+/g, ' ').trim();
        /* a data mora logo abaixo; se estiver vazia, olha mais 2 linhas */
        for (var d = 1; d <= 3 && !dataAtualizacao; d++) {
          if (rr + d < valores.length && col < valores[rr + d].length) {
            dataAtualizacao = pdmData(valores[rr + d][col]);
          }
        }
        break;
      }
    }
  }

  /* ---- indexa as linhas pela chave ---- */
  var linhas = {};
  for (var lr = linhaCab + 1; lr < valores.length; lr++) {
    var bruto = valores[lr][iChave];
    var chave = (cfg.chaveBusca === 'cpf') ? normCpf(bruto) : normMat(bruto);
    if (!chave || /^0*$/.test(chave)) continue;

    var reg = {};
    for (var kk = 0; kk < cfg.campos.length; kk++) {
      var cp = cfg.campos[kk], ci = cols[cp.chave];
      reg[cp.chave] = (ci === undefined) ? null : pdmNum(valores[lr][ci]);
    }
    if (cfg.titulo) {
      var it = cols._titulo;
      reg._titulo = (it === undefined) ? '' : limpar(valores[lr][it]);
    }

    if (cfg.multiplas) {
      if (!linhas[chave]) linhas[chave] = [];
      linhas[chave].push(reg);
    } else if (!linhas[chave]) {
      linhas[chave] = reg;            /* duplicado numa aba de 1 linha: fica a 1ª */
    }
  }

  var tabela = {
    tipo: tipo,
    aba: sh.getName(),
    linhaCab: linhaCab + 1,
    colunasAchadas: cols,
    colunaChave: iChave,
    rotuloData: rotuloData || 'Valor atualizado em:',
    dataAtualizacao: dataAtualizacao,
    linhas: linhas
  };
  if (cache) { try { cache.put(chaveCache, JSON.stringify(tabela), PDM_CACHE_SEG); } catch (e3) {} }
  return tabela;
}

function limparCacheConsultas() {
  try {
    var c = CacheService.getScriptCache();
    Object.keys(CP_CONSULTAS).forEach(function (t) { c.remove('CP_' + t + '_V2'); });
  } catch (e) {}
  return 'Cache das consultas pessoais limpo.';
}
/* nome antigo, mantido para não quebrar quem já chamava */
function limparCachePeDeMeia() { return limparCacheConsultas(); }

/** Confere o token e devolve o técnico da mini master. */
function cpTecnicoDaSessao(dados) {
  var sess = validarToken(dados && dados.token);
  if (!sess.ok) {
    return {
      erro: {
        ok: false, sessao: false,
        erro: sess.expirado ? 'Sessão expirada. Faça login novamente.'
                            : 'Sessão inválida. Faça login novamente.'
      }
    };
  }
  var mat = normMat(sess.mat), achado = null;
  lerMiniMasterCompleto().forEach(function (t) {
    if (!achado && normMat(t.mat) === mat) achado = t;
  });
  if (!achado) return { erro: { ok: false, erro: 'Matrícula não encontrada no cadastro.' } };
  return { tecnico: achado, mat: mat };
}

function cpMontarValores(campos, reg) {
  return campos.map(function (c) {
    return {
      chave: c.chave, rotulo: c.rotulo,
      valor: reg ? reg[c.chave] : null,
      texto: reg ? pdmMoeda(reg[c.chave]) : null
    };
  });
}

/** Resposta para a tela. Nunca devolve o CPF de volta pro navegador. */
function consultaPessoal(dados) {
  var tipo = (dados && dados.tipo) || 'peDeMeia';
  var cfg = CP_CONSULTAS[tipo];
  if (!cfg) return { ok: false, erro: 'Consulta desconhecida.' };

  var s = cpTecnicoDaSessao(dados);
  if (s.erro) return s.erro;

  var chave;
  if (cfg.chaveBusca === 'cpf') {
    chave = normCpf(s.tecnico.cpf);
    if (!chave) return { ok: false, erro: 'Sua matrícula não tem CPF cadastrado. Fale com a administração.' };
  } else {
    chave = s.mat;
  }

  var tab = cpLerTabela(tipo);
  var achou = tab.linhas[chave] || null;

  var base = {
    ok: true, tipo: tipo,
    nome: s.tecnico.nome, mat: String(s.tecnico.mat),
    rotuloData: tab.rotuloData, dataAtualizacao: tab.dataAtualizacao,
    destaque: cfg.destaque
  };

  /* --- aba de uma linha só (Pé de meia): mantém o formato antigo --- */
  if (!cfg.multiplas) {
    base.encontrado = !!achou;
    base.valores = cpMontarValores(cfg.campos, achou);
    return base;
  }

  /* --- abas de N linhas (Cursos, Dívidas) --- */
  var regs = achou || [];
  base.encontrado = regs.length > 0;
  base.rotuloTitulo = cfg.titulo ? cfg.titulo.rotulo : '';
  base.itens = regs.map(function (reg) {
    return {
      titulo: reg._titulo || '(sem descrição)',
      valores: cpMontarValores(cfg.campos, reg)
    };
  });

  /* Total só quando há mais de uma linha — com um item só, repetir o mesmo
     número embaixo não informa nada. */
  base.totais = null;
  if (regs.length > 1) {
    base.totais = cfg.campos.map(function (c) {
      var soma = 0, tem = false;
      regs.forEach(function (reg) {
        if (reg[c.chave] !== null && reg[c.chave] !== undefined) { soma += reg[c.chave]; tem = true; }
      });
      return { chave: c.chave, rotulo: c.rotulo, valor: tem ? soma : null, texto: tem ? pdmMoeda(soma) : null };
    });
  }
  return base;
}

/* Nome antigo da ação: a tela pe-de-meia.html já publicada continua chamando
   assim, então ela não precisa ser trocada junto com o resto. */
function peDeMeia(dados) {
  var d = dados || {};
  return consultaPessoal({ token: d.token, tipo: 'peDeMeia' });
}

/** Rode isto no editor do Apps Script para conferir se a leitura pegou tudo. */
function testarConsultas() {
  limparCacheConsultas();
  Object.keys(CP_CONSULTAS).forEach(function (tipo) {
    var cfg = CP_CONSULTAS[tipo];
    Logger.log('───────── ' + tipo + ' ─────────');
    var t;
    try { t = cpLerTabela(tipo); }
    catch (e) { Logger.log('ERRO: ' + e); return; }

    var faltando = cfg.campos.filter(function (c) { return t.colunasAchadas[c.chave] === undefined; })
                             .map(function (c) { return c.rotulo; });
    Logger.log('Aba: ' + t.aba + ' | cabeçalho na linha ' + t.linhaCab
      + ' | coluna-chave: ' + t.colunaChave);
    Logger.log('Colunas achadas: ' + JSON.stringify(t.colunasAchadas));
    Logger.log('Colunas NÃO achadas: ' + (faltando.length ? faltando.join(', ') : 'nenhuma'));
    Logger.log('Data: "' + t.rotuloData + '" -> "' + t.dataAtualizacao + '"');

    var chaves = Object.keys(t.linhas);
    Logger.log('Pessoas na tabela: ' + chaves.length);
    if (chaves.length) {
      var k = chaves[0], v = t.linhas[k];
      var rotulo = (cfg.chaveBusca === 'cpf') ? (k.slice(0, 3) + '.***') : ('mat ' + k);
      Logger.log('Amostra (' + rotulo + '): '
        + (cfg.multiplas ? v.length + ' linha(s) — ' : '') + JSON.stringify(v));
    }
  });
  return 'Veja o log acima.';
}
/* nome antigo, mantido para não quebrar quem já chamava */
function testarPeDeMeia() { return testarConsultas(); }

/* ═══════════════════════════════════════════════════════════════════════
   CHECKLIST SEMANAL COMPARTILHADO PELA EQUIPE DO DIA
   ═══════════════════════════════════════════════════════════════════════

   O PROBLEMA
   Cada aparelho guardava sozinho, no localStorage, quais checklists já tinham
   sido feitos na semana. Três técnicos do mesmo parque viam o mesmo alerta
   vermelho e faziam o mesmo checklist três vezes — o celular de um não sabia
   do outro.

   A REGRA
   Quem trabalha junto divide a baixa. "Junto" sai da coluna H da aba
   Relatorios, que guarda as matrículas da equipe de cada RDO. Se um da equipe
   fizer o checklist, ele some do alerta dos outros até a semana virar.

   Quem NÃO aparece na coluna H não está em parque, e para esse não há alerta
   nenhum — os cartões ficam na cor padrão.

   A SEMANA
   Mesma régua do prazos.js: o ciclo abre na SEXTA e é identificado pela data
   dessa sexta. Baixa dada no ciclo vale até a sexta seguinte, quando a chave
   muda sozinha e o alerta volta. Não existe rotina de limpeza.

   ONDE FICA GUARDADO
   Aba "Checklist_Feitos", na mesma planilha do RDO, criada sozinha no primeiro
   uso. Uma linha por baixa. É registro, não estado: dá para auditar quem deu
   baixa em quê e quando.

   ATENCAO: isto é só o BACKEND. O site ainda não chama estes endpoints — o
   prazos.js continua decidindo tudo pelo localStorage. Enquanto o front não
   for ligado, nada muda para o técnico.
   ═══════════════════════════════════════════════════════════════════════ */

/* Os cinco checklists que a equipe divide. Os ids são os mesmos do prazos.js,
   para o front não precisar traduzir nada.

   Falta de propósito o 'epi' (Equipamentos Individuais): a engenharia listou
   cinco, e ele não estava na lista. Para incluir, basta acrescentar a linha —
   nada mais precisa mudar. */
var CKL_ITENS = [
  { id: 'materiais',   rotulo: 'Materiais' },
  { id: 'veiculo',     rotulo: 'Veiculos' },
  { id: 'loto',        rotulo: 'Kit LOTO' },
  { id: 'cordas',      rotulo: 'Acesso por Cordas' },
  { id: 'ferramentas', rotulo: 'Ferramentas Gerais' }
];

var CKL_ABA = 'Checklist_Feitos';
var CKL_CAB = ['Ciclo', 'Checklist', 'Matricula', 'Nome', 'Parque', 'Data', 'Hora'];
var CKL_ABA_RDO = 'Relatorios';
var CKL_COL_MAT_PADRAO = 'H';    /* coluna das matriculas da equipe */
var CKL_COL_DATA_PADRAO = 'C';   /* data do expediente */
var CKL_COL_PARQUE_PADRAO = 'G';

/* Quantos dias para tras procurar a equipe do tecnico.

   Por que nao e so "hoje": o RDO e enviado no FIM do expediente. Se a busca
   fosse estrita, ninguem apareceria na coluna H durante o dia inteiro de
   trabalho e o alerta ficaria escondido justamente enquanto ha tempo de fazer
   o checklist. Tres dias cobrem a virada de sexta para segunda, que e
   exatamente a janela do prazo.

   Para deixar estrito ("so o RDO de hoje"), ponha 0. */
var CKL_JANELA_DIAS = 3;

var CKL_ABRE_CICLO = 5;   /* 5 = sexta, igual ao prazos.js */

function cklIds() {
  return CKL_ITENS.map(function (i) { return i.id; });
}
function cklRotulo(id) {
  for (var i = 0; i < CKL_ITENS.length; i++) if (CKL_ITENS[i].id === id) return CKL_ITENS[i].rotulo;
  return id;
}
function cklIdValido(id) {
  return cklIds().indexOf(String(id || '').trim()) >= 0;
}

/** yyyy-MM-dd no fuso do script. */
function cklIso(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/** Sexta que abre o ciclo vigente — a mais recente, contando hoje. */
function cklSextaDoCiclo(d) {
  var x = new Date(d || new Date());
  x.setHours(0, 0, 0, 0);
  x.setDate(x.getDate() - ((x.getDay() - CKL_ABRE_CICLO + 7) % 7));
  return x;
}

/**
 * Quebra a celula da coluna H na lista de matriculas.
 *
 * Separa por QUALQUER coisa que nao seja digito, de proposito. O RDO grava
 * "70, 45, 350", mas a celula tambem e editada a mao, e ja apareceu com
 * hifen ("70 - 45 - 350") e so com espaco ("70 45 350"). Uma lista fixa de
 * separadores devolvia "7045350" nesses casos — uma matricula inexistente,
 * e o tecnico simplesmente nao recebia alerta, sem nenhum erro na tela.
 *
 * "N/A" some sozinho: sem digito, vira string vazia e cai no filtro.
 */
function cklMatriculas(celula) {
  return String(celula == null ? '' : celula)
    .split(/\D+/)
    .map(function (s) { return normMat(s); })
    .filter(function (s) { return !!s; });
}

/**
 * Equipe do tecnico nos ultimos CKL_JANELA_DIAS dias.
 * Devolve { ativo, parque, data, equipe:[matriculas] } — equipe ja inclui ele.
 * Nao achou nenhum RDO com essa matricula na janela: ativo = false.
 */
function cklEquipeDoTecnico(mat, hoje) {
  var vazio = { ativo: false, parque: '', data: '', equipe: [] };
  mat = normMat(mat);
  if (!mat) return vazio;

  var props = PropertiesService.getScriptProperties();
  var ss;
  try { ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID')); }
  catch (e) { return vazio; }
  var sh = ss.getSheetByName(CKL_ABA_RDO);
  if (!sh || sh.getLastRow() < 2) return vazio;

  /* Cabecalho manda; a letra fixa e o plano B para planilha sem cabecalho
     nomeado. Assim, mover coluna na planilha nao quebra em silencio. */
  var iMat = idxCabecalho(sh, 'Matriculas');
  if (iMat < 0) iMat = letraParaIndice(CKL_COL_MAT_PADRAO);
  var iData = idxCabecalho(sh, 'Data');
  if (iData < 0) iData = letraParaIndice(CKL_COL_DATA_PADRAO);
  var iParque = idxCabecalho(sh, 'Parque');
  if (iParque < 0) iParque = letraParaIndice(CKL_COL_PARQUE_PADRAO);

  var nCols = Math.max(iMat, iData, iParque) + 1;
  if (sh.getLastColumn() < nCols) return vazio;

  var ref = new Date(hoje || new Date()); ref.setHours(0, 0, 0, 0);
  var limite = new Date(ref); limite.setDate(limite.getDate() - CKL_JANELA_DIAS);
  var isoRef = cklIso(ref), isoLimite = cklIso(limite);

  var dados = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  var melhor = null;
  for (var r = 0; r < dados.length; r++) {
    var equipe = cklMatriculas(dados[r][iMat]);
    if (equipe.indexOf(mat) < 0) continue;
    var dia = normData(dados[r][iData]);
    /* Comparacao de texto yyyy-MM-dd ordena igual a data — e nao depende de
       como o Sheets devolveu a celula (Date ou texto). */
    if (!dia || dia > isoRef || dia < isoLimite) continue;
    if (!melhor || dia > melhor.data) {
      melhor = { ativo: true, parque: String(dados[r][iParque] || ''), data: dia, equipe: equipe };
    }
  }
  return melhor || vazio;
}

/** Aba de registro, criada na primeira baixa. */
function cklAba(ss) {
  var sh = ss.getSheetByName(CKL_ABA);
  if (sh) return sh;
  sh = ss.insertSheet(CKL_ABA);
  sh.getRange(1, 1, 1, CKL_CAB.length).setValues([CKL_CAB]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

/** Baixas do ciclo, indexadas por checklist -> primeira linha que registrou. */
function cklFeitosDoCiclo(ss, ciclo) {
  var sh = ss.getSheetByName(CKL_ABA);
  var out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  var dados = sh.getRange(2, 1, sh.getLastRow() - 1, CKL_CAB.length).getValues();
  dados.forEach(function (l) {
    if (String(l[0]).trim() !== ciclo) return;
    var id = String(l[1] || '').trim();
    if (!id) return;
    /* Fica a PRIMEIRA baixa do ciclo: e ela que conta a historia de quem
       resolveu. Reenvio da fila offline nao sobrescreve o autor. */
    if (!out[id]) {
      out[id] = { por: normMat(l[2]), nome: String(l[3] || ''), parque: String(l[4] || ''),
                  data: normData(l[5]), hora: String(l[6] || '') };
    }
  });
  return out;
}

/** Nome do tecnico pela mini master (so para o registro ficar legivel). */
function cklNomeDe(mat) {
  var alvo = normMat(mat), nome = '';
  try {
    lerMiniMasterCompleto().forEach(function (t) {
      if (!nome && normMat(t.mat) === alvo) nome = t.nome;
    });
  } catch (e) {}
  return nome;
}

/**
 * Estado dos checklists para quem esta logado.
 *
 * ativo = false  -> nao esta em parque; o site nao deve mostrar alerta nenhum.
 * ativo = true   -> devolve, por checklist, se alguem da equipe ja deu baixa.
 */
function checklistStatus(dados) {
  var sess = validarToken(dados && dados.token);
  if (!sess.ok) {
    return {
      ok: false, sessao: false,
      erro: sess.expirado ? 'Sessao expirada. Faca login novamente.'
                          : 'Sessao invalida. Faca login novamente.'
    };
  }

  var mat = normMat(sess.mat);
  var hoje = new Date();
  var ciclo = cklIso(cklSextaDoCiclo(hoje));
  var eq = cklEquipeDoTecnico(mat, hoje);

  var base = {
    ok: true, sessao: true, mat: mat, nome: cklNomeDe(mat),
    ciclo: ciclo, ativo: eq.ativo, parque: eq.parque, dataRdo: eq.data,
    equipe: eq.equipe, janelaDias: CKL_JANELA_DIAS,
    itens: CKL_ITENS.map(function (i) {
      return { id: i.id, rotulo: i.rotulo, feito: false, por: '', porNome: '', quando: '' };
    })
  };

  /* Fora de parque: devolve tudo "nao feito" e ativo=false. O front usa o
     ativo para nao pintar nada — e nao o feito, que aqui nao quer dizer nada. */
  if (!eq.ativo) return base;

  var props = PropertiesService.getScriptProperties();
  var ss;
  try { ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID')); }
  catch (e) { return { ok: false, erro: 'Nao consegui abrir a planilha do RDO: ' + e }; }

  var feitos = cklFeitosDoCiclo(ss, ciclo);
  base.itens.forEach(function (it) {
    var f = feitos[it.id];
    /* So conta se quem deu baixa esta na equipe de hoje deste tecnico. Um
       parque nao apaga o alerta do outro. */
    if (f && eq.equipe.indexOf(f.por) >= 0) {
      it.feito = true;
      it.por = f.por;
      it.porNome = f.nome || cklNomeDe(f.por);
      it.quando = f.data + (f.hora ? ' ' + f.hora : '');
      it.parque = f.parque;
    }
  });
  return base;
}

/**
 * Registra que este tecnico fez um dos checklists no ciclo vigente.
 * Idempotente: repetir no mesmo ciclo nao duplica linha nem troca o autor.
 */
function checklistFeito(dados) {
  var sess = validarToken(dados && dados.token);
  if (!sess.ok) {
    return {
      ok: false, sessao: false,
      erro: sess.expirado ? 'Sessao expirada. Faca login novamente.'
                          : 'Sessao invalida. Faca login novamente.'
    };
  }

  var id = String((dados && dados.checklist) || '').trim();
  if (!cklIdValido(id)) {
    return { ok: false, erro: 'Checklist desconhecido: "' + id + '". Conhecidos: ' + cklIds().join(', ') };
  }

  var mat = normMat(sess.mat);
  var hoje = new Date();
  var ciclo = cklIso(cklSextaDoCiclo(hoje));
  var eq = cklEquipeDoTecnico(mat, hoje);

  var props = PropertiesService.getScriptProperties();
  var ss;
  try { ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID')); }
  catch (e) { return { ok: false, erro: 'Nao consegui abrir a planilha do RDO: ' + e }; }

  /* Duas baixas simultaneas da mesma equipe criariam duas linhas e a segunda
     roubaria a autoria. O lock e curto porque a operacao e uma leitura e um
     append. */
  var lock = LockService.getScriptLock();
  var travou = false;
  try { travou = lock.tryLock(15000); } catch (eL) { travou = false; }

  try {
    var feitos = cklFeitosDoCiclo(ss, ciclo);
    var ja = feitos[id];
    if (ja && (!eq.ativo || eq.equipe.indexOf(ja.por) >= 0)) {
      return { ok: true, ciclo: ciclo, checklist: id, rotulo: cklRotulo(id),
               jaEstava: true, por: ja.por, porNome: ja.nome || cklNomeDe(ja.por),
               quando: ja.data + (ja.hora ? ' ' + ja.hora : '') };
    }
    var agora = new Date();
    cklAba(ss).appendRow([
      ciclo, id, mat, cklNomeDe(mat), eq.parque || '',
      cklIso(agora), Utilities.formatDate(agora, Session.getScriptTimeZone(), 'HH:mm')
    ]);
    return { ok: true, ciclo: ciclo, checklist: id, rotulo: cklRotulo(id),
             jaEstava: false, por: mat, porNome: cklNomeDe(mat),
             quando: cklIso(agora), ativo: eq.ativo, equipe: eq.equipe };
  } finally {
    if (travou) { try { lock.releaseLock(); } catch (eR) {} }
  }
}

/** Rode no editor do Apps Script para conferir a leitura da coluna H. */
function testarChecklistEquipe() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  var sh = ss.getSheetByName(CKL_ABA_RDO);
  if (!sh) { Logger.log('Aba "' + CKL_ABA_RDO + '" nao existe.'); return; }

  var iMat = idxCabecalho(sh, 'Matriculas');
  Logger.log('Coluna das matriculas: ' + (iMat >= 0
    ? 'achada pelo cabecalho "Matriculas", indice ' + iMat + ' (coluna ' + (iMat + 1) + ')'
    : 'cabecalho nao achado — usando o padrao ' + CKL_COL_MAT_PADRAO));
  if (iMat < 0) iMat = letraParaIndice(CKL_COL_MAT_PADRAO);

  var iData = idxCabecalho(sh, 'Data');
  if (iData < 0) iData = letraParaIndice(CKL_COL_DATA_PADRAO);
  Logger.log('Coluna da data: indice ' + iData);

  var hoje = new Date();
  Logger.log('Ciclo vigente (sexta): ' + cklIso(cklSextaDoCiclo(hoje)));
  Logger.log('Janela de atividade: ' + CKL_JANELA_DIAS + ' dia(s)');

  var n = Math.min(sh.getLastRow() - 1, 400);
  if (n < 1) { Logger.log('Sem linhas na aba.'); return; }
  var dados = sh.getRange(2, 1, n, Math.max(iMat, iData) + 1).getValues();

  var ref = new Date(hoje); ref.setHours(0, 0, 0, 0);
  var limite = new Date(ref); limite.setDate(limite.getDate() - CKL_JANELA_DIAS);
  var isoRef = cklIso(ref), isoLimite = cklIso(limite);

  var naJanela = [], todas = {};
  dados.forEach(function (l) {
    var eq = cklMatriculas(l[iMat]);
    eq.forEach(function (m) { todas[m] = true; });
    var dia = normData(l[iData]);
    if (dia && dia <= isoRef && dia >= isoLimite) naJanela.push({ data: dia, equipe: eq });
  });

  Logger.log('Linhas lidas: ' + n);
  Logger.log('Matriculas distintas na coluna H: ' + Object.keys(todas).length);
  Logger.log('RDOs dentro da janela (' + isoLimite + ' a ' + isoRef + '): ' + naJanela.length);
  naJanela.slice(0, 8).forEach(function (x) {
    Logger.log('   ' + x.data + '  equipe: ' + x.equipe.join(', '));
  });

  var amostra = naJanela.length ? naJanela[0].equipe[0] : Object.keys(todas)[0];
  if (amostra) {
    Logger.log('--- simulando a matricula ' + amostra + ' ---');
    Logger.log(JSON.stringify(cklEquipeDoTecnico(amostra, hoje)));
  }
  return { matriculas: Object.keys(todas).length, naJanela: naJanela.length };
}



/* =====================================================================
 * MEUS DADOS — MEUS EQUIPAMENTOS
 * ---------------------------------------------------------------------
 * Fonte: "CONTROLE DE EQUIPAMENTOS - EXTREME.xlsm", no Dropbox do
 * almoxarifado. O almoxarifado continua editando o Excel normalmente;
 * este backend só LÊ o arquivo.
 *
 * Como fica "em tempo real":
 *   um gatilho de tempo roda eqGatilho() a cada 10 min. Ele pergunta ao
 *   Dropbox a revisão do arquivo; se mudou, baixa, lê a aba BASE DE DADOS
 *   direto do XML do .xlsm (sem converter para Google Sheets, sem macros)
 *   e monta um índice por matrícula no cache. A tela do técnico lê só o
 *   índice, então a resposta é rápida.
 *   Atraso real = salvar no Excel + sincronizar o Dropbox + até 10 min.
 *
 * Proteção contra item de uma pessoa aparecer para outra:
 *   o item só entra no índice da matrícula da coluna "MATRÍCULA RESPONSÁVEL"
 *   se o nome da coluna RESPONSÁVEL bater com o nome dessa matrícula na
 *   mini master. Se não bater, o item NÃO aparece e vai para o relatório de
 *   divergências (relatorioDivergenciasEquipamentos).
 *
 * Aba Meus Equipamentos: ID, descrição, qtd (somada), localização, data de saída.
 * Aba Extraviados e Avariados: arquivo "LISTA DE EXTRAVIO E DESCONTOS
 *   Equipamentos.xlsx" (01 - ALMOXARIFADO/12 - Slides), aba Inventário:
 *   descrição, valor, responsável, projeto/parque, desconto (SIM/NÃO), mês;
 *   ocorrência = nota da célula; relatório = hiperlink da linha (se houver).
 * Não sai: Disponível/Descarte e matrícula 0 (EM SEPARAÇÃO, ADM etc.).
 *
 * Propriedades do Script (todas opcionais):
 *   EQUIP_ARQUIVO               caminho ou id do Dropbox (padrão: o id abaixo)
 *   EQUIP_ABA                   padrão "BASE DE DADOS"
 *   EQUIP_EXTRAVIO_ARQUIVO      caminho ou id da lista de extravio (padrão: o id abaixo)
 *   EQUIP_EXTRAVIO_ABA          padrão "Inventário"
 *   EQUIP_EMAIL_DIVERGENCIAS    e-mail que recebe a lista de divergências
 *                               sempre que a planilha muda (vazio = não envia)
 *
 * Rodar à mão no editor:
 *   testarEquipamentos()                  sincroniza e mostra o diagnóstico
 *   relatorioDivergenciasEquipamentos()   lista o que ficou de fora e por quê
 *   instalarGatilhoEquipamentos()         cria o gatilho de 10 min (1 vez só)
 * ===================================================================== */

var EQ_ARQUIVO_PADRAO = 'id:zJii0PvESPAAAAAAAAmXCg';
var EQ_ABA_PADRAO = 'BASE DE DADOS';
var EQ_CACHE_SEG = 21600;          /* 6 h (máximo do CacheService) */
var EQ_TRAVA_SEG = 240;            /* trava "suave" da sincronização */
var EQ_STATUS_FORA = { 'DISPONIVEL': 1, 'DESCARTE': 1 };
/* Extraviado/Quarentena da BASE não entram na tabela: a aba Extraviados e
   Avariados vem da LISTA DE EXTRAVIO E DESCONTOS (abaixo). */
var EQ_STATUS_EXTRAVIO = { 'EXTRAVIADO': 1, 'QUARENTENA': 1 };
var EQ_PFX = 'EQ3_';               /* troca de formato do cache = prefixo novo */

/* ---- LISTA DE EXTRAVIO E DESCONTOS Equipamentos.xlsx (pasta 12 - Slides) ---- */
var EQ_EXT_ARQUIVO_PADRAO = 'id:zJii0PvESPAAAAAAAAjUiw';
var EQ_EXT_ABA_PADRAO = 'Inventário';
var EQ_EXT_COLUNAS = {
  descricao:  ['DESCRICAO'],
  valor:      ['R$ VALOR', 'VALOR (R$)', 'VALOR'],
  resp:       ['RESPONSAVEL', 'TECNICO'],
  projeto:    ['PROJETO/PARQUE', 'PROJETO / PARQUE', 'PROJETO', 'PARQUE'],
  descontado: ['DESCONTO', 'DESCONTADO'],
  mes:        ['MES'],
  mat:        ['MATRICULA', 'MATRICULAS', 'MATRICULA RESPONSAVEL'],   /* opcional */
  obs:        ['OCORRENCIA', 'OBSERVACAO', 'OBSERVACOES'],           /* opcional: senão usa o comentário da célula */
  link:       ['LINK RELATORIO', 'LINK DO RELATORIO', 'RELATORIO', 'LINK'] /* opcional: senão usa o hiperlink da linha */
};
var EQ_EXT_OBRIGATORIAS = ['descricao', 'resp', 'descontado'];
var EQ_EXT_NAO_PESSOA = { 'EQUIPE': 1, 'TOTAL': 1 };

/* cabeçalhos procurados na linha de cabeçalho (comparação sem acento/caixa).
   A 1ª ocorrência vence — as colunas MIRROR têm nome diferente. */
var EQ_COLUNAS = {
  id:         ['ID'],
  codigo:     ['CODIGO'],
  classe:     ['CLASSIFICACAO'],
  categoria:  ['CATEGORIA'],
  serie:      ['N DE SERIE', 'NO DE SERIE', 'NUMERO DE SERIE'],
  marca:      ['MARCA'],
  descricao:  ['DESCRICAO'],
  calib:      ['PROX. CALIBRACAO / INSPECAO', 'PROX CALIBRACAO / INSPECAO'],
  qtd:        ['ESTOQUE'],
  und:        ['UND. MEDIDA', 'UND MEDIDA'],
  status:     ['STATUS', 'PARQUE / FABRICA'],
  local:      ['LOCALIZACAO'],
  saida:      ['DATA SAIDA'],
  resp:       ['RESPONSAVEL'],
  req:        ['N REQUISICAO DE ENTREGA', 'NO REQUISICAO DE ENTREGA'],
  mat:        ['MATRICULA RESPONSAVEL', 'MATRICULA']
};
var EQ_OBRIGATORIAS = ['descricao', 'status', 'resp', 'mat'];

/* ---------- entrada: chamada pela tela meus-dados/equipamentos.html ---------- */

function meusEquipamentos(dados) {
  var s = cpTecnicoDaSessao(dados);
  if (s.erro) return s.erro;

  var cache = CacheService.getScriptCache();
  var gen = cache.get(EQ_PFX + 'GEN');
  if (!gen) {
    /* cache vazio (1ª vez ou 6 h sem mudança): sincroniza agora */
    var r = eqSincronizar(false);
    if (r.ocupado) {
      return { ok: false, retentar: true,
               erro: 'A lista de equipamentos está sendo atualizada. Tente de novo em 1 minuto.' };
    }
    gen = cache.get(EQ_PFX + 'GEN');
    if (!gen) return { ok: false, erro: 'Não foi possível ler a planilha de equipamentos agora.' };
  }

  var meta = {};
  try { meta = JSON.parse(cache.get(EQ_PFX + gen + '_META') || '{}'); } catch (e) {}
  var reg = { itens: [], extravios: [] };
  try { reg = JSON.parse(cache.get(EQ_PFX + gen + '_M' + s.mat) || '{"itens":[],"extravios":[]}'); } catch (e2) {}

  return {
    ok: true, tipo: 'equipamentos',
    nome: s.tecnico.nome, mat: String(s.tecnico.mat),
    encontrado: (reg.itens || []).length + (reg.extravios || []).length > 0,
    itens: reg.itens || [],
    extravios: reg.extravios || [],
    atualizadoEm: meta.atualizadoEm || '',
    extravioAtualizadoEm: meta.extravioAtualizadoEm || '',
    lidoEm: meta.lidoEm || ''
  };
}

/* ---------- gatilho de tempo ---------- */

function eqGatilho() {
  try { eqSincronizar(false); }
  catch (e) { Logger.log('eqGatilho: ' + e); }
}

function instalarGatilhoEquipamentos() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'eqGatilho') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('eqGatilho').timeBased().everyMinutes(10).create();
  Logger.log('Gatilho criado: eqGatilho a cada 10 minutos.');
}

/* ---------- sincronização ---------- */

/**
 * Não usa o LockService do script de propósito: o envio de RDO usa essa
 * mesma trava e ficaria esperando ~20 s enquanto a planilha é lida.
 */
function eqSincronizar(forcar) {
  var cache = CacheService.getScriptCache();
  var props = PropertiesService.getScriptProperties();

  if (cache.get(EQ_PFX + 'SYNC')) return { ok: false, ocupado: true };
  cache.put(EQ_PFX + 'SYNC', '1', EQ_TRAVA_SEG);

  try {
    var token = getDropboxToken(props);
    var arqBase = props.getProperty('EQUIP_ARQUIVO') || EQ_ARQUIVO_PADRAO;
    var arqExt = props.getProperty('EQUIP_EXTRAVIO_ARQUIVO') || EQ_EXT_ARQUIVO_PADRAO;
    var meta = eqDbxMetadata(token, arqBase, props);
    var metaExt = null, erroExt = '';
    try { metaExt = eqDbxMetadata(token, arqExt, props); } catch (eM) { erroExt = String(eM); }
    var rev = meta.rev + '|' + (metaExt ? metaExt.rev : 'sem-lista');

    var genAtual = cache.get(EQ_PFX + 'GEN');
    if (!forcar && genAtual && cache.get(EQ_PFX + 'REV') === rev) {
      return { ok: true, mudou: false, rev: rev };
    }

    var mini = lerMiniMasterCompleto();

    var blob = eqDbxBaixar(token, arqBase, props);
    var aba = props.getProperty('EQUIP_ABA') || EQ_ABA_PADRAO;
    var lido = eqLerAba(Utilities.unzip(blob.setContentType('application/zip')), aba);
    var idx = eqMontarIndice(lido.linhas, mini);

    /* a lista de extravio falhar não derruba a aba de equipamentos */
    var ext = { porMat: {}, divergencias: [], total: 0 }, lidoExt = null;
    if (metaExt) {
      try {
        var blobExt = eqDbxBaixar(token, arqExt, props);
        var abaExt = props.getProperty('EQUIP_EXTRAVIO_ABA') || EQ_EXT_ABA_PADRAO;
        lidoExt = eqLerAba(Utilities.unzip(blobExt.setContentType('application/zip')), abaExt,
                           { colunas: EQ_EXT_COLUNAS, obrigatorias: EQ_EXT_OBRIGATORIAS });
        ext = eqMontarExtravios(lidoExt.linhas, mini);
      } catch (eX) { erroExt = String(eX); }
    }
    if (erroExt) Logger.log('Lista de extravio: ' + erroExt);

    /* geração nova = chaves novas; as antigas expiram sozinhas. Assim um
       técnico que devolveu tudo não continua vendo a lista velha. */
    var gen = Utilities.getUuid().slice(0, 8);
    var lote = {}, mats = {};
    Object.keys(idx.porMat).forEach(function (m) { mats[m] = 1; });
    Object.keys(ext.porMat).forEach(function (m) { mats[m] = 1; });
    Object.keys(mats).forEach(function (m) {
      lote[EQ_PFX + gen + '_M' + m] = JSON.stringify({
        itens: idx.porMat[m] ? idx.porMat[m].itens : [],
        extravios: ext.porMat[m] || []
      });
    });
    lote[EQ_PFX + gen + '_META'] = JSON.stringify({
      atualizadoEm: eqFmtDataHora(meta.server_modified),
      extravioAtualizadoEm: metaExt ? eqFmtDataHora(metaExt.server_modified) : '',
      lidoEm: eqFmtDataHora(new Date().toISOString()),
      rev: rev
    });
    var todasDiv = idx.divergencias.concat(ext.divergencias);
    lote[EQ_PFX + gen + '_DIV'] = JSON.stringify(todasDiv.slice(0, 400));
    eqPutAll(cache, lote);
    cache.putAll(eqObj(EQ_PFX + 'GEN', gen, EQ_PFX + 'REV', rev), EQ_CACHE_SEG);

    var resumo = {
      ok: true, mudou: true, rev: rev,
      base: { linhasLidas: lido.linhas.length, cabecalhoNaLinha: lido.linhaCabecalho, colunas: lido.colunas,
              tecnicos: Object.keys(idx.porMat).length, itensNoApp: idx.totalItens,
              divergencias: idx.divergencias.length, ignorados: idx.ignorados },
      listaExtravio: lidoExt
        ? { linhasLidas: lidoExt.linhas.length, cabecalhoNaLinha: lidoExt.linhaCabecalho, colunas: lidoExt.colunas,
            tecnicos: Object.keys(ext.porMat).length, registrosNoApp: ext.total, divergencias: ext.divergencias.length }
        : { erro: erroExt || 'não lida' }
    };
    eqAvisarDivergencias(props, todasDiv, genAtual);
    return resumo;
  } finally {
    cache.remove(EQ_PFX + 'SYNC');
  }
}

function eqObj() { var o = {}; for (var i = 0; i < arguments.length; i += 2) o[arguments[i]] = arguments[i + 1]; return o; }

function eqPutAll(cache, obj) {
  var chaves = Object.keys(obj), bloco = {};
  for (var i = 0; i < chaves.length; i++) {
    var v = obj[chaves[i]];
    if (v.length > 95000) {
      /* > 100 KB por chave: corta a lista em vez de perder tudo */
      var o = JSON.parse(v);
      var lista = Array.isArray(o) ? o : (o.itens || []);
      while (JSON.stringify(o).length > 95000 && lista.length) lista.pop();
      v = JSON.stringify(o);
    }
    bloco[chaves[i]] = v;
    if (Object.keys(bloco).length >= 100) { cache.putAll(bloco, EQ_CACHE_SEG); bloco = {}; }
  }
  if (Object.keys(bloco).length) cache.putAll(bloco, EQ_CACHE_SEG);
}

/* ---------- Dropbox ---------- */

/* Conta de time: o arquivo pode estar fora do namespace padrão. Se o caminho
   não for achado, tenta de novo a partir da raiz do time e guarda a escolha. */
function eqDbxCall(url, token, props, extra, conteudo) {
  var tentativas = [props.getProperty('EQUIP_PATH_ROOT') || ''];
  if (tentativas[0] === '') tentativas.push('ROOT');

  var ultimo = null;
  for (var i = 0; i < tentativas.length; i++) {
    var headers = { 'Authorization': 'Bearer ' + token };
    if (tentativas[i] === 'ROOT') {
      var ns = eqDbxRootNs(token);
      if (!ns) break;
      headers['Dropbox-API-Path-Root'] = JSON.stringify({ '.tag': 'root', 'root': ns });
    }
    var op = { method: 'post', headers: headers, muteHttpExceptions: true };
    if (conteudo) {
      headers['Dropbox-API-Arg'] = escaparArg(extra);
    } else {
      op.contentType = 'application/json';
      op.payload = JSON.stringify(extra);
    }
    var r = UrlFetchApp.fetch(url, op);
    ultimo = r;
    if (r.getResponseCode() === 200) {
      if (tentativas[i] !== (props.getProperty('EQUIP_PATH_ROOT') || '')) {
        props.setProperty('EQUIP_PATH_ROOT', tentativas[i]);
      }
      return r;
    }
    if (r.getResponseCode() !== 409) break;   /* 409 = não achou: vale tentar a raiz */
  }
  throw new Error('Dropbox não liberou o arquivo de equipamentos (' +
    (ultimo ? ultimo.getResponseCode() + ': ' + ultimo.getContentText().slice(0, 300) : 'sem resposta') +
    '). Confira se a conta do app do RDO enxerga a pasta do almoxarifado.');
}

function eqDbxRootNs(token) {
  var r = UrlFetchApp.fetch('https://api.dropboxapi.com/2/users/get_current_account', {
    method: 'post', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true
  });
  if (r.getResponseCode() !== 200) return null;
  var j = JSON.parse(r.getContentText());
  return j.root_info && j.root_info.root_namespace_id;
}

function eqDbxMetadata(token, arquivo, props) {
  var r = eqDbxCall('https://api.dropboxapi.com/2/files/get_metadata', token, props, { path: arquivo });
  return JSON.parse(r.getContentText());
}

function eqDbxBaixar(token, arquivo, props) {
  return eqDbxCall('https://content.dropboxapi.com/2/files/download', token, props, { path: arquivo }, true).getBlob();
}

/* ---------- leitura do .xlsm (é um zip de XMLs) ---------- */

function eqLerAba(blobs, nomeAba, cfg) {
  var arq = {};
  blobs.forEach(function (b) { arq[b.getName()] = b; });
  var txt = function (n) { return arq[n] ? arq[n].getDataAsString('UTF-8') : ''; };
  return eqLerAbaXml(txt('xl/workbook.xml'), txt('xl/_rels/workbook.xml.rels'),
                     txt('xl/sharedStrings.xml'), txt, nomeAba, cfg);
}

/* Parte pura (sem serviços do Google): dá para testar fora do Apps Script. */
function eqLerAbaXml(wbXml, relsXml, ssXml, lerArquivo, nomeAba, cfg) {
  cfg = cfg || { colunas: EQ_COLUNAS, obrigatorias: EQ_OBRIGATORIAS };
  var alvo = eqNorm(nomeAba), rid = null, m;
  var reSheet = /<sheet\b([^>]*)\/?>/g;
  while ((m = reSheet.exec(wbXml))) {
    var nm = /\bname="([^"]*)"/.exec(m[1]), id = /\br:id="([^"]*)"/.exec(m[1]);
    if (nm && id && eqNorm(eqXmlTexto(nm[1])) === alvo) { rid = id[1]; break; }
  }
  if (!rid) throw new Error('Aba "' + nomeAba + '" não encontrada na planilha.');

  var alvoArq = null, reRel = /<Relationship\b([^>]*)\/?>/g;
  while ((m = reRel.exec(relsXml))) {
    var i2 = /\bId="([^"]*)"/.exec(m[1]), t2 = /\bTarget="([^"]*)"/.exec(m[1]);
    if (i2 && i2[1] === rid && t2) { alvoArq = t2[1]; break; }
  }
  if (!alvoArq) throw new Error('Arquivo da aba "' + nomeAba + '" não encontrado.');
  alvoArq = alvoArq.charAt(0) === '/' ? alvoArq.slice(1) : 'xl/' + alvoArq.replace(/^\.\//, '');

  /* textos compartilhados */
  var ss = [], reSi = /<si>([\s\S]*?)<\/si>/g;
  while ((m = reSi.exec(ssXml))) ss.push(eqJuntarT(m[1]));

  var xml = lerArquivo(alvoArq);
  if (!xml) throw new Error('Não consegui abrir ' + alvoArq + '.');

  /* comentários (notas) e hiperlinks, por número de linha */
  var extras = eqComentariosELinks(xml, alvoArq, lerArquivo);

  var colunas = null, linhaCab = 0, linhas = [];
  var linhaAtual = 0, vals = {};

  function fecharLinha() {
    if (!linhaAtual) return;
    if (!colunas) {
      if (linhaAtual <= 15) {
        var c = eqAcharColunas(vals, cfg);
        if (c) { colunas = c; linhaCab = linhaAtual; }
      }
    } else {
      var reg = { _linha: linhaAtual }, tem = false;
      Object.keys(colunas).forEach(function (k) {
        var v = vals[colunas[k]];
        if (v !== undefined && v !== null && v !== '') { reg[k] = v; tem = true; }
      });
      if (tem) {
        if (extras.coment[linhaAtual]) reg._coment = extras.coment[linhaAtual];
        if (extras.link[linhaAtual]) reg._link = extras.link[linhaAtual];
        linhas.push(reg);
      }
    }
  }

  var reC = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  while ((m = reC.exec(xml))) {
    var ref = /\br="([A-Z]+)(\d+)"/.exec(m[1]);
    if (!ref) continue;
    var ln = Number(ref[2]);
    if (ln !== linhaAtual) { fecharLinha(); linhaAtual = ln; vals = {}; }
    if (!m[2]) continue;
    var fh = /<f\b[^>]*>([\s\S]*?)<\/f>/.exec(m[2]);
    if (fh) {
      var hl = /HYPERLINK\(\s*(?:&quot;|")(https?:[^"&]+)/i.exec(fh[1]);
      if (hl && !extras.link[ln]) extras.link[ln] = eqXmlTexto(hl[1]);
    }
    var tipo = /\bt="([^"]*)"/.exec(m[1]);
    tipo = tipo ? tipo[1] : 'n';
    var v = null;
    if (tipo === 'inlineStr') {
      v = eqJuntarT(m[2]);
    } else {
      var vv = /<v>([\s\S]*?)<\/v>/.exec(m[2]);
      if (vv) {
        if (tipo === 's') v = ss[Number(vv[1])];
        else if (tipo === 'str') v = eqXmlTexto(vv[1]);
        else if (tipo === 'e') v = null;              /* #REF!, #N/D… */
        else if (tipo === 'b') v = vv[1] === '1';
        else v = Number(vv[1]);
      }
    }
    if (typeof v === 'string') v = v.trim();
    vals[ref[1]] = v;
  }
  fecharLinha();

  if (!colunas) {
    throw new Error('Não achei a linha de cabeçalho na aba "' + nomeAba +
      '" (precisa ter as colunas ' + cfg.obrigatorias.join(', ') + ').');
  }
  return { colunas: colunas, linhaCabecalho: linhaCab, linhas: linhas };
}

function eqAcharColunas(vals, cfg) {
  var achadas = {};
  Object.keys(cfg.colunas).forEach(function (k) {
    var alvos = cfg.colunas[k];
    for (var a = 0; a < alvos.length && !achadas[k]; a++) {
      for (var col in vals) {
        if (typeof vals[col] === 'string' && eqNorm(vals[col]) === alvos[a]) {
          if (!achadas[k] || eqColNum(col) < eqColNum(achadas[k])) achadas[k] = col;
        }
      }
    }
  });
  for (var i = 0; i < cfg.obrigatorias.length; i++) if (!achadas[cfg.obrigatorias[i]]) return null;
  return achadas;
}

/* Notas da célula (xl/commentsN.xml) e hiperlinks (<hyperlinks> + rels da aba). */
function eqComentariosELinks(xml, arqAba, lerArquivo) {
  var out = { coment: {}, link: {} }, m;
  var relsPath = arqAba.replace(/([^\/]+)$/, '_rels/$1.rels');
  var rels = {}, relsXml = lerArquivo(relsPath) || '';
  var reRel = /<Relationship\b([^>]*)\/?>/g;
  while ((m = reRel.exec(relsXml))) {
    var id = /\bId="([^"]*)"/.exec(m[1]), tg = /\bTarget="([^"]*)"/.exec(m[1]), tp = /\bType="([^"]*)"/.exec(m[1]);
    if (id && tg) rels[id[1]] = { alvo: eqXmlTexto(tg[1]), tipo: tp ? tp[1] : '' };
  }

  var reH = /<hyperlink\b([^>]*)\/?>/g;
  while ((m = reH.exec(xml))) {
    var ref = /\bref="[A-Z]+(\d+)/.exec(m[1]), rid = /\br:id="([^"]*)"/.exec(m[1]);
    if (ref && rid && rels[rid[1]] && /^https?:/i.test(rels[rid[1]].alvo)) out.link[Number(ref[1])] = rels[rid[1]].alvo;
  }

  Object.keys(rels).forEach(function (k) {
    if (!/\/comments$/.test(rels[k].tipo)) return;
    var p = rels[k].alvo;
    p = p.charAt(0) === '/' ? p.slice(1) : 'xl/' + p.replace(/^(\.\.\/)+/, '').replace(/^\.\//, '');
    var cx = lerArquivo(p) || '';
    var autores = [], ma, reA = /<author>([\s\S]*?)<\/author>/g;
    while ((ma = reA.exec(cx))) autores.push(eqXmlTexto(ma[1]).trim());
    var reCm = /<comment\b[^>]*\bref="[A-Z]+(\d+)"[^>]*>([\s\S]*?)<\/comment>/g, mc;
    while ((mc = reCm.exec(cx))) {
      var t = eqJuntarT(mc[2]);
      if (t.indexOf('Comentário:') >= 0) t = t.slice(t.lastIndexOf('Comentário:') + 11);   /* comentário encadeado */
      var mudou = true;
      while (mudou) {                                   /* tira "extre:" (nome do autor) do começo */
        mudou = false;
        for (var i = 0; i < autores.length; i++) {
          var pre = autores[i] + ':';
          if (autores[i] && t.trim().indexOf(pre) === 0) { t = t.trim().slice(pre.length); mudou = true; }
        }
      }
      t = t.replace(/\s+/g, ' ').trim();
      if (t) { var ln = Number(mc[1]); out.coment[ln] = out.coment[ln] ? out.coment[ln] + ' ' + t : t; }
    }
  });
  return out;
}

function eqColNum(l) { var n = 0; for (var i = 0; i < l.length; i++) n = n * 26 + (l.charCodeAt(i) - 64); return n; }

function eqJuntarT(x) {
  var semFonetica = x.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
  var out = '', m, re = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
  while ((m = re.exec(semFonetica))) out += m[1];
  return eqXmlTexto(out);
}

function eqXmlTexto(s) {
  return String(s)
    .replace(/_x([0-9A-Fa-f]{4})_/g, function (_, h) { return String.fromCharCode(parseInt(h, 16)); })
    .replace(/&#x([0-9A-Fa-f]+);/g, function (_, h) { return String.fromCharCode(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function (_, d) { return String.fromCharCode(Number(d)); })
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

function eqNorm(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[ºª°]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();
}

/* ---------- índice por matrícula, com a checagem de nome ---------- */

function eqMontarIndice(linhas, miniMaster) {
  var nomePorMat = {};
  miniMaster.forEach(function (t) { nomePorMat[normMat(t.mat)] = t.nome; });

  var porMat = {}, agrup = {}, div = [], total = 0;
  var ign = { semMatricula: 0, statusFora: 0 };

  function reg(m) { return (porMat[m] = porMat[m] || { itens: [], extravios: [] }); }

  /* algumas linhas vêm com a DESCRIÇÃO vazia (fórmula sem resultado):
     completa com a descrição de outra linha do mesmo ID */
  var descPorId = {};
  linhas.forEach(function (r) {
    if (r.id != null && r.descricao && !descPorId[r.id]) descPorId[r.id] = String(r.descricao).trim();
  });
  linhas.forEach(function (r) {
    if (!r.descricao && r.id != null && descPorId[r.id]) r.descricao = descPorId[r.id];
  });

  linhas.forEach(function (r) {
    var mat = normMat(r.mat);
    if (!mat || mat === '0') { ign.semMatricula++; return; }
    var st = eqNorm(r.status);
    if (EQ_STATUS_FORA[st]) { ign.statusFora++; return; }
    if (EQ_STATUS_EXTRAVIO[st]) { ign.statusFora++; return; }

    var nomeCad = nomePorMat[mat];
    var problema = null;
    if (!nomeCad) problema = 'matrícula não existe na mini master';
    else if (!eqNomeBate(r.resp, nomeCad)) problema = 'nome não bate com a matrícula (cadastro: ' + nomeCad + ')';
    if (problema) {
      div.push({ origem: 'BASE DE DADOS', linha: r._linha, id: r.id != null ? String(r.id) : '',
                 descricao: String(r.descricao || ''), responsavel: String(r.resp || ''), matricula: mat, motivo: problema });
      return;
    }

    /* uma linha por unidade na planilha -> soma por ID + local + data de saída */
    var saida = r.saida != null ? eqData(r.saida) : '';
    var k = mat + '|' + (r.id != null ? r.id : r.descricao) + '|' + (r.local || '') + '|' + saida;
    var q = typeof r.qtd === 'number' ? r.qtd : 1;
    if (agrup[k]) { agrup[k].qtd += q; }
    else {
      agrup[k] = { id: r.id != null ? String(r.id) : '', descricao: String(r.descricao || '').trim(),
                   qtd: q, localizacao: String(r.local || '').trim(), dataSaida: saida };
      reg(mat).itens.push(agrup[k]);
    }
    total++;
  });

  Object.keys(porMat).forEach(function (m) {
    porMat[m].itens.sort(function (a, b) {
      if (!a.descricao !== !b.descricao) return a.descricao ? -1 : 1;   /* sem descrição vai pro fim */
      return a.descricao.localeCompare(b.descricao);
    });
  });
  return { porMat: porMat, divergencias: div, totalItens: total, ignorados: ign };
}

var EQ_MESES = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO',
                'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'];

/**
 * LISTA DE EXTRAVIO E DESCONTOS: o RESPONSÁVEL é um nome curto e às vezes
 * vários ("LUANN/HÉLIO/TAYLAN", "MARCOS MUNIZ E WAGNER SANTOS"). Se a lista
 * ganhar uma coluna MATRÍCULA, ela manda. Sem ela, cada nome só vale se casar
 * com UM técnico da mini master; ambíguo ou desconhecido vai para divergências
 * (valor de desconto nunca aparece para a pessoa errada).
 */
function eqMontarExtravios(linhas, miniMaster) {
  var cad = miniMaster.map(function (t) {
    return { mat: normMat(t.mat), nome: t.nome, toks: eqNorm(t.nome).split(' ') };
  });
  var nomePorMat = {};
  cad.forEach(function (t) { nomePorMat[t.mat] = t.nome; });

  var porMat = {}, div = [], total = 0;
  function falha(r, nome, motivo, mat) {
    div.push({ origem: 'LISTA DE EXTRAVIO', linha: r._linha, id: '', descricao: String(r.descricao || ''),
               responsavel: nome, matricula: mat || '', motivo: motivo });
  }

  linhas.forEach(function (r) {
    var desc = eqNorm(r.descricao);
    if (!desc || desc === 'TOTAL') return;
    var partes = eqNorm(r.resp).split(/\s*\/\s*|\s+E\s+|\s*,\s*|\s*\+\s*/).filter(Boolean);
    var item = eqExtravioPublico(r);
    var mats = [];

    if (r.mat != null && String(r.mat).trim() !== '') {
      String(r.mat).split(/[^0-9]+/).filter(Boolean).forEach(function (m) {
        m = normMat(m);
        if (!nomePorMat[m]) { falha(r, String(r.resp || ''), 'matrícula não existe na mini master', m); return; }
        var bate = !partes.length || partes.some(function (p) { return eqNomeBate(p, nomePorMat[m]); });
        if (!bate) { falha(r, String(r.resp || ''), 'nome não bate com a matrícula (cadastro: ' + nomePorMat[m] + ')', m); return; }
        mats.push(m);
      });
    } else {
      if (!partes.length) { falha(r, '', 'sem responsável'); return; }
      partes.forEach(function (p) {
        if (EQ_EXT_NAO_PESSOA[p] || p === eqNorm(r.projeto)) { falha(r, p, 'não é o nome de um técnico'); return; }
        var tk = p.split(' ');
        var c = cad.filter(function (t) { return tk.every(function (x) { return t.toks.indexOf(x) >= 0; }); });
        if (c.length === 1) mats.push(c[0].mat);
        else if (!c.length) falha(r, p, 'nome não encontrado na mini master');
        else falha(r, p, 'nome ambíguo: ' + c.slice(0, 4).map(function (t) { return t.mat + ' ' + t.nome; }).join('; '));
      });
    }

    mats.filter(function (m, i) { return mats.indexOf(m) === i; }).forEach(function (m) {
      (porMat[m] = porMat[m] || []).push(item);
      total++;
    });
  });
  return { porMat: porMat, divergencias: div, total: total };
}

function eqExtravioPublico(r) {
  var obs = r.obs != null && String(r.obs).trim() ? String(r.obs).trim() : (r._coment || '');
  var mes = r.mes != null ? (typeof r.mes === 'number' ? eqMesDeSerial(r.mes) : String(r.mes).trim().toUpperCase()) : '';
  if (!mes) {   /* sem coluna MÊS: tenta a data escrita na observação (dd/mm/aaaa) */
    var d = /(\d{1,2})\/(\d{1,2})\/(\d{2,4})/.exec(obs);
    if (d && Number(d[2]) >= 1 && Number(d[2]) <= 12) mes = EQ_MESES[Number(d[2]) - 1];
  }
  var desc = null;
  if (r.descontado != null && r.descontado !== '') {
    var t = eqNorm(r.descontado);
    if (r.descontado === true || /^(SIM|S|X|OK|DESCONTADO|TRUE|1)$/.test(t)) desc = true;
    else if (r.descontado === false || /^(NAO|N|NAO DESCONTADO|FALSE|0)$/.test(t)) desc = false;
  }
  var link = r.link != null && /^https?:/i.test(String(r.link).trim()) ? String(r.link).trim() : (r._link || '');
  if (!/^https:\/\//i.test(link)) link = '';
  return {
    descricao: String(r.descricao || '').trim(),
    valor: eqMoeda(r.valor),
    projeto: r.projeto != null ? String(r.projeto).trim() : '',
    mes: mes,
    descontado: desc,
    ocorrencia: obs,
    linkRelatorio: link
  };
}

function eqMesDeSerial(v) {
  if (v > 20000 && v < 80000) return EQ_MESES[new Date(Math.round((v - 25569) * 86400000)).getUTCMonth()];
  return String(v);
}

function eqMoeda(v) {
  if (typeof v !== 'number') {
    if (v == null || v === '') return '';
    var n = Number(String(v).replace(/[^\d,.-]/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.'));
    if (isNaN(n)) return String(v);
    v = n;
  }
  var p = v.toFixed(2).split('.');
  return 'R$ ' + p[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + p[1];
}

/* Todos os pedaços do nome escrito na planilha precisam existir no nome do
   cadastro (aceita nome abreviado, recusa nome de outra pessoa). */
function eqNomeBate(nomePlanilha, nomeCadastro) {
  var a = eqNorm(nomePlanilha).split(/[\s\/]+/).filter(Boolean);
  if (!a.length) return false;
  var b = {};
  eqNorm(nomeCadastro).split(' ').forEach(function (p) { b[p] = 1; });
  for (var i = 0; i < a.length; i++) if (!b[a[i]]) return false;
  return true;
}

/* número de série do Excel -> dd/mm/aaaa; texto fica como está */
function eqData(v) {
  if (typeof v === 'number' && v > 20000 && v < 80000) {
    var d = new Date(Math.round((v - 25569) * 86400000));
    var p = function (n) { return (n < 10 ? '0' : '') + n; };
    return p(d.getUTCDate()) + '/' + p(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
  }
  return String(v);
}

function eqFmtDataHora(iso) {
  try { return Utilities.formatDate(new Date(iso), 'America/Fortaleza', "dd/MM/yyyy 'às' HH:mm"); }
  catch (e) { return String(iso || ''); }
}

/* ---------- divergências ---------- */

function eqAvisarDivergencias(props, div, genAnterior) {
  var para = props.getProperty('EQUIP_EMAIL_DIVERGENCIAS');
  if (!para || !div.length || !genAnterior) return;   /* 1ª carga não dispara e-mail */
  var linhas = div.slice(0, 200).map(function (d) {
    return '<tr><td>' + eqHtml(d.origem || '') + ' ' + d.linha + '</td><td>' + eqHtml(d.id) + '</td><td>' + eqHtml(d.descricao) +
           '</td><td>' + eqHtml(d.responsavel) + '</td><td>' + d.matricula + '</td><td>' + eqHtml(d.motivo) + '</td></tr>';
  }).join('');
  try {
    MailApp.sendEmail({
      to: para,
      subject: 'Equipamentos: ' + div.length + ' item(ns) fora do app dos técnicos',
      htmlBody: '<p>Estes itens não aparecem no app dos técnicos porque não deu para ligar o ' +
        'nome em RESPONSÁVEL a uma matrícula com segurança (BASE DE DADOS ou LISTA DE EXTRAVIO).</p>' +
        '<table border="1" cellpadding="4" cellspacing="0"><tr><th>Linha</th><th>ID</th><th>Descrição</th>' +
        '<th>Responsável</th><th>Matrícula</th><th>Motivo</th></tr>' + linhas + '</table>',
      name: 'App dos técnicos — Extreme Wind'
    });
  } catch (e) { Logger.log('e-mail de divergências falhou: ' + e); }
}

function eqHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}

/* ---------- rodar à mão ---------- */

function testarEquipamentos() {
  var r = eqSincronizar(true);
  Logger.log(JSON.stringify(r, null, 2));
  relatorioDivergenciasEquipamentos();
  return r;
}

function relatorioDivergenciasEquipamentos() {
  var cache = CacheService.getScriptCache();
  var gen = cache.get(EQ_PFX + 'GEN');
  if (!gen) { Logger.log('Sem dados no cache: rode testarEquipamentos() primeiro.'); return []; }
  var div = JSON.parse(cache.get(EQ_PFX + gen + '_DIV') || '[]');
  var grupos = {};
  div.forEach(function (d) {
    var k = (d.origem || '') + ' | ' + d.responsavel + (d.matricula ? ' → mat. ' + d.matricula : '') + ' | ' + d.motivo;
    grupos[k] = (grupos[k] || 0) + 1;
  });
  Logger.log('Itens fora do app: ' + div.length);
  Object.keys(grupos).sort(function (a, b) { return grupos[b] - grupos[a]; })
    .forEach(function (k) { Logger.log(grupos[k] + ' × ' + k); });
  return div;
}

/** Simula a tela de um técnico sem precisar logar no celular. */
function testarEquipamentosDaMatricula() {
  var mat = '239';   /* troque aqui */
  var cache = CacheService.getScriptCache();
  var gen = cache.get(EQ_PFX + 'GEN');
  var r = gen ? JSON.parse(cache.get(EQ_PFX + gen + '_M' + normMat(mat)) || '{"itens":[],"extravios":[]}') : { itens: [], extravios: [] };
  Logger.log(r.itens.length + ' linhas de equipamento e ' + r.extravios.length + ' extravios/avarias para a matrícula ' + mat);
  Logger.log(JSON.stringify({ itens: r.itens.slice(0, 5), extravios: r.extravios.slice(0, 5) }, null, 2));
}
