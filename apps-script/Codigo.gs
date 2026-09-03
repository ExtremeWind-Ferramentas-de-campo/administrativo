/**
 * ============================================================================
 * PORTAL ADMINISTRATIVO — EXTREME WIND BLADE SERVICES
 * Backend em Google Apps Script para a tela de login (portal-admin.html)
 * ============================================================================
 *
 * ORDEM DE INSTALAÇÃO (faça exatamente nesta sequência):
 *
 *  1. Crie uma planilha nova no Google Sheets.
 *  2. Extensões > Apps Script. Apague o conteúdo e cole este arquivo inteiro.
 *  3. Rode a função  configurarPlanilha()  uma vez. Ela cria as abas e cabeçalhos.
 *  4. Rode a função  definirPepper()  uma vez. Ela gera a chave secreta do hash.
 *  5. Cadastre as pessoas: rode  criarUsuariosExemplo()  para testar, ou use
 *     criarUsuario({...}) com os dados reais.
 *  6. Implantar > Nova implantação > Tipo: App da Web
 *       Executar como:        Eu (seu email)
 *       Quem pode acessar:    Qualquer pessoa
 *     Copie a URL /exec.
 *  7. No portal-admin.html, troque:
 *       const MODO_DEMO = false;
 *       const API_URL   = 'cole a URL /exec aqui';
 *
 * "Qualquer pessoa" libera só a URL do Web App, não a planilha. A planilha
 * continua privada — ninguém além de você precisa de acesso a ela.
 *
 * SEGURANÇA IMPLEMENTADA
 *  - Senha nunca é gravada em texto: SHA-256 de (salt + senha + pepper).
 *  - Salt aleatório por usuário; pepper guardado nas Propriedades do Script.
 *  - Bloqueio após 5 tentativas erradas (15 min).
 *  - Token de sessão de 8 h no CacheService — a planilha não é exposta ao navegador.
 *  - Troca de senha só é aplicada depois do código enviado por email.
 *  - CPF nunca é devolvido inteiro para o navegador, só os 3 últimos dígitos.
 *  - Toda ação fica registrada na aba LOG.
 * ============================================================================
 */

/* ---------------------------------------------------------------------------
   CONSTANTES
   --------------------------------------------------------------------------- */
const ABA_USUARIOS    = 'USUARIOS';
const ABA_RECUPERACAO = 'RECUPERACAO';
const ABA_LOG         = 'LOG';

const MAX_TENTATIVAS   = 5;
const BLOQUEIO_MINUTOS = 15;
const CODIGO_MINUTOS   = 30;   // validade do código de confirmação
const SESSAO_HORAS     = 8;

const NOME_SISTEMA = 'Portal Administrativo — Extreme Wind';

/* Quanto tempo cada tipo de registro fica na aba LOG.

   São dois prazos porque as duas coisas têm valor bem diferente:

   - SEGURANÇA (quem entrou, quem errou a senha, quem trocou, quem foi
     cadastrado). Volume minúsculo — umas 40 linhas por dia numa equipe de 20.
     É o que você vai querer olhar se um dia precisar entender um acesso
     indevido, e esse tipo de coisa costuma aparecer semanas depois.

   - OPERAÇÃO (anexos baixados, arquivos enviados). Volume alto e utilidade
     curta: serve para investigar um problema desta semana, não do semestre
     passado. */
const DIAS_LOG_SEGURANCA = 180;
const DIAS_LOG_OPERACAO  = 15;

const ACOES_SEGURANCA = [
  'LOGIN', 'DEFINIR_SENHA', 'SOLICITAR_TROCA', 'CONFIRMAR_TROCA',
  'RESET_ADMIN', 'CADASTRO', 'SALVAR_PERFIL', 'ERRO'
];

const COLUNAS_USUARIOS = [
  'matricula','nome','cpf','email','setor','cargo','perfil',
  'admissao','nascimento','telefone','emergencia',
  'senha_hash','salt','primeiro_acesso','status',
  'ultimo_acesso','tentativas','bloqueado_ate'
];

const COLUNAS_RECUPERACAO = [
  'id','matricula','codigo_hash','senha_nova_hash','salt_nova',
  'criado_em','expira_em','usado'
];

const COLUNAS_LOG = ['data_hora','matricula','acao','resultado','detalhe'];

/* Campos que o próprio usuário pode alterar pela tela.
   Precisa bater com CAMPOS_PERFIL do HTML — se divergir, o backend manda. */
const EDITAVEL_DONO  = ['email','telefone','nascimento','emergencia'];
const EDITAVEL_ADMIN = ['nome','setor','cargo','admissao','email','telefone','nascimento','emergencia'];

/* Campos que o ADMIN normalmente controla, mas que a própria pessoa pode
   preencher UMA vez, enquanto estiverem em branco. É o que permite cadastrar
   alguém só com matrícula, CPF e email: o resto ela completa no primeiro acesso.
   Depois de preenchidos, voltam a ser exclusivos do administrativo — assim
   ninguém troca o próprio setor ou nome depois. */
const PREENCHIVEL_UMA_VEZ = ['nome','setor','cargo'];


/**
 * Mostra um aviso na tela quando há tela, e no registro de execução quando não há.
 *
 * O Apps Script só oferece interface quando o script roda a partir da planilha
 * aberta. Rodando pelo editor em outra aba, ou por acionador, getUi() estoura
 * "Cannot call SpreadsheetApp.getUi() from this context" — e uma função de
 * instalação não pode falhar por causa da mensagem final.
 */
function avisar_(texto) {
  try {
    SpreadsheetApp.getUi().alert(texto);
  } catch (e) {
    Logger.log('\n' + texto + '\n');
  }
}


/* ===========================================================================
   MENU NA PLANILHA
   Aparece como "Portal" na barra de menus ao abrir a planilha. É o caminho
   fácil para cadastrar gente sem mexer no código.
   =========================================================================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Portal')
    .addItem('Cadastrar pessoa', 'menuCadastrar')
    .addItem('Resetar senha de alguém', 'menuResetar')
    .addSeparator()
    .addItem('Ver usuários cadastrados', 'menuListar')
    .addSeparator()
    .addItem('Configurar planilhas (RDO e MINI MASTER)', 'menuConfigurarPlanilhas')
    .addItem('Conferir colunas do RDO', 'menuColunasRDO')
    .addToUi();
}

function menuCadastrar() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(
    'Cadastrar pessoa',
    'Digite os três dados separados por ponto e vírgula:\n\n' +
    'matrícula ; CPF ; email\n\n' +
    'Exemplo:\n' +
    '10432 ; 523.984.710-23 ; ana@extremewind.com.br\n\n' +
    'Nome, setor e cargo a própria pessoa preenche no primeiro acesso.',
    ui.ButtonSet.OK_CANCEL);

  if (r.getSelectedButton() !== ui.Button.OK) return;

  const p = r.getResponseText().split(';').map(function (s) { return s.trim(); });
  if (p.length < 3) {
    ui.alert('Faltaram campos. São 3: matrícula ; CPF ; email');
    return;
  }

  // O perfil não vai no mesmo campo para ninguém marcar ADMIN por engano
  // ao cadastrar dez pessoas seguidas.
  const rp = ui.prompt('Perfil de acesso',
    'Digite o número do perfil desta pessoa:\n\n' +
    '1 = ADMIN — setor administrativo. Cadastra pessoas e usa os módulos.\n' +
    '       Em Projetos em Andamento, só visualiza.\n' +
    '2 = SUPERVISOR — cria e edita os projetos em andamento.\n' +
    '3 = DIRETORIA — acompanha tudo, sem alterar projetos.',
    ui.ButtonSet.OK_CANCEL);
  if (rp.getSelectedButton() !== ui.Button.OK) return;

  const perfil = { '1': 'ADMIN', '2': 'SUPERVISOR', '3': 'DIRETORIA' }[rp.getResponseText().trim()];
  if (!perfil) { ui.alert('Perfil inválido. Digite 1, 2 ou 3.'); return; }

  try {
    criarUsuario({ matricula: p[0], cpf: p[1], email: p[2], perfil: perfil });
    ui.alert('Pronto!\n\n' +
             'Matrícula: ' + p[0] + '\n' +
             'Perfil: ' + perfil + '\n' +
             'Senha do primeiro acesso: o CPF, só os números (' + soDigitos_(p[1]) + ')\n\n' +
             'No primeiro login o sistema pede uma senha nova e, em seguida, ' +
             'abre "Meus dados" para a pessoa preencher nome, setor e o resto.');
  } catch (erro) {
    ui.alert('Não deu certo:\n\n' + erro.message);
  }
}

function menuResetar() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Resetar senha',
    'Digite a matrícula. A senha volta a ser o CPF e o primeiro acesso é reativado.',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  try {
    resetarSenhaPeloAdmin(r.getResponseText().trim());
    ui.alert('Senha resetada. A pessoa entra com a matrícula e o CPF.');
  } catch (erro) {
    ui.alert('Não deu certo:\n\n' + erro.message);
  }
}

function menuListar() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_USUARIOS);
  const dados = aba.getDataRange().getValues();
  const cab = dados[0];
  const iM = cab.indexOf('matricula'), iN = cab.indexOf('nome');
  const iP = cab.indexOf('perfil'), iS = cab.indexOf('status');
  const iA = cab.indexOf('primeiro_acesso');

  if (dados.length < 2) {
    avisar_('Nenhuma pessoa cadastrada ainda.\n\nUse Portal > Cadastrar pessoa.');
    return;
  }

  let txt = '';
  for (let l = 1; l < dados.length; l++) {
    const primeiro = dados[l][iA] === true || String(dados[l][iA]).toUpperCase() === 'TRUE';
    txt += String(dados[l][iM]).replace(/^'/, '') + ' — ' + (dados[l][iN] || '(nome ainda não preenchido)') +
           ' (' + dados[l][iP] + ', ' + dados[l][iS] + ')' +
           (primeiro ? ' — ainda não trocou a senha, entra com o CPF' : '') + '\n';
  }
  avisar_('Usuários cadastrados (' + (dados.length - 1) + ')\n\n' + txt);
}


/* ===========================================================================
   INSTALAÇÃO
   =========================================================================== */

/** Passo 3 — cria as abas e os cabeçalhos. Pode rodar de novo sem apagar dados. */
function configurarPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  criarAba_(ss, ABA_USUARIOS,    COLUNAS_USUARIOS);
  criarAba_(ss, ABA_RECUPERACAO, COLUNAS_RECUPERACAO);
  criarAba_(ss, ABA_LOG,         COLUNAS_LOG);

  // Protege as colunas sensíveis contra edição manual acidental
  const aba = ss.getSheetByName(ABA_USUARIOS);
  aba.setFrozenRows(1);
  aba.getRange('L:M').setBackground('#f3f3f3').setFontColor('#999999');
  aba.getRange('L1:M1').setNote('Gerado pelo sistema. Não edite à mão.');

  avisar_('Abas criadas.\n\nPróximo passo: rode definirPepper() uma vez.');
}

/** Passo 4 — gera a chave secreta do hash. Rode UMA vez e nunca mais. */
function definirPepper() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('PEPPER')) {
    Logger.log('PEPPER já existe. Não vou sobrescrever — isso invalidaria todas as senhas.');
    return;
  }
  props.setProperty('PEPPER', gerarAleatorio_(32));
  Logger.log('PEPPER criado. Não apague essa propriedade: sem ela, nenhuma senha confere.');
}

function criarAba_(ss, nome, colunas) {
  let aba = ss.getSheetByName(nome);
  if (!aba) aba = ss.insertSheet(nome);
  const atual = aba.getRange(1, 1, 1, Math.max(1, aba.getLastColumn())).getValues()[0];
  if (atual.join('') !== colunas.join('')) {
    aba.getRange(1, 1, 1, colunas.length)
       .setValues([colunas])
       .setFontWeight('bold')
       .setBackground('#14181b')
       .setFontColor('#ecede8');
    aba.setFrozenRows(1);
  }
  return aba;
}


/* ===========================================================================
   CADASTRO DE USUÁRIOS (executar pelo editor do Apps Script)
   =========================================================================== */

/**
 * Cadastra uma pessoa. A senha inicial é o CPF e o primeiro login obriga a troca.
 *
 * criarUsuario({
 *   matricula:'10432', nome:'Ana Paula Ferreira', cpf:'523.984.710-23',
 *   email:'ana.ferreira@extremewind.com.br', setor:'Operações',
 *   cargo:'Supervisora de Campo', perfil:'SUPERVISOR', admissao:'2021-03-15'
 * });
 */
function criarUsuario(dados) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_USUARIOS);
  const matricula = String(dados.matricula).trim();
  const cpf = soDigitos_(dados.cpf);

  if (!matricula)        throw new Error('Matrícula é obrigatória.');
  if (cpf.length !== 11) throw new Error('CPF precisa ter 11 dígitos.');
  if (!dados.email)      throw new Error('Email é obrigatório.');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(dados.email).trim()))
    throw new Error('Email inválido: ' + dados.email);
  if (acharPorMatricula_(aba, matricula)) throw new Error('Matrícula ' + matricula + ' já existe.');

  const salt = gerarAleatorio_(16);
  const linha = COLUNAS_USUARIOS.map(function (c) {
    switch (c) {
      case 'matricula':       return "'" + matricula;   // texto: preserva zero à esquerda
      case 'cpf':             return "'" + cpf;          // apóstrofo preserva zeros à esquerda
      case 'senha_hash':      return hash_(cpf, salt);   // senha inicial = CPF
      case 'salt':            return salt;
      case 'primeiro_acesso': return true;
      case 'status':          return 'ATIVO';
      case 'perfil':          return (dados.perfil || 'DIRETORIA').toUpperCase();
      case 'tentativas':      return 0;
      case 'bloqueado_ate':   return '';
      case 'ultimo_acesso':   return '';
      default:                return dados[c] || '';
    }
  });

  aba.appendRow(linha);
  registrar_(matricula, 'CADASTRO', 'OK', dados.email);
  Logger.log('Usuário ' + matricula + ' criado. Senha inicial: o CPF.');
}

/** Passo 5 (opcional) — dois usuários para testar a tela. */
function criarUsuariosExemplo() {
  criarUsuario({
    matricula: '10432', nome: 'Ana Paula Ferreira', cpf: '523.984.710-23',
    email: Session.getEffectiveUser().getEmail(), setor: 'Operações',
    cargo: 'Supervisora de Campo', perfil: 'SUPERVISOR', admissao: '2021-03-15'
  });
  criarUsuario({
    matricula: '10087', nome: 'Marcos Vieira', cpf: '111.222.333-96',
    email: Session.getEffectiveUser().getEmail(), setor: 'Administrativo',
    cargo: 'Analista Administrativo', perfil: 'ADMIN', admissao: '2019-08-01'
  });
  Logger.log('Os dois usuários usam o SEU email, para você testar a recuperação.\n' +
             'ATENÇÃO: a senha destes dois é o CPF, só números:\n' +
             '  matrícula 10432  ->  52398471023\n' +
             '  matrícula 10087  ->  11122233396\n' +
             'As senhas do modo demonstração do site (ex.: Torre2026) NÃO valem aqui.');
}

/** Reset manual: volta a senha para o CPF e reativa o primeiro acesso. */
function resetarSenhaPeloAdmin(matricula) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_USUARIOS);
  const u = acharPorMatricula_(aba, String(matricula).trim());
  if (!u) throw new Error('Matrícula não encontrada.');

  const salt = gerarAleatorio_(16);
  gravar_(aba, u.linha, {
    salt: salt,
    senha_hash: hash_(u.dados.cpf, salt),
    primeiro_acesso: true,
    tentativas: 0,
    bloqueado_ate: ''
  });
  registrar_(matricula, 'RESET_ADMIN', 'OK', '');
  Logger.log('Senha de ' + matricula + ' voltou a ser o CPF.');
}


/* ===========================================================================
   ENDPOINT WEB
   =========================================================================== */

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, servico: NOME_SISTEMA }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let resposta;
  try {
    const corpo = JSON.parse(e.postData.contents);
    switch (corpo.acao) {
      case 'login':          resposta = acaoLogin_(corpo);          break;
      case 'definirSenha':   resposta = acaoDefinirSenha_(corpo);   break;
      case 'solicitarTroca': resposta = acaoSolicitarTroca_(corpo); break;
      case 'confirmarTroca': resposta = acaoConfirmarTroca_(corpo); break;
      case 'salvarPerfil':   resposta = acaoSalvarPerfil_(corpo);   break;
      case 'sessao':         resposta = acaoSessao_(corpo);         break;

      // Módulo de materiais — implementado em Materiais.gs
      case 'materiaisCarregar':      resposta = acaoMateriaisCarregar_(corpo);      break;
      case 'materiaisSalvar':        resposta = acaoMateriaisSalvar_(corpo);        break;
      case 'materiaisAnexar':        resposta = acaoMateriaisAnexar_(corpo);        break;
      case 'materiaisBaixarAnexo':   resposta = acaoMateriaisBaixarAnexo_(corpo);   break;
      case 'materiaisRemoverAnexo':  resposta = acaoMateriaisRemoverAnexo_(corpo);  break;
      case 'materiaisMiniaturas':    resposta = acaoMateriaisMiniaturas_(corpo);    break;

      // Supervisão de campo — implementado em SupervisaoCampo.gs
      case 'projetosCarregar':       resposta = acaoProjetosCarregar_(corpo);      break;
      case 'projetosSalvar':         resposta = acaoProjetosSalvar_(corpo);        break;
      case 'rdoStatus':              resposta = acaoRdoStatus_(corpo);             break;
      case 'rdoFiltros':             resposta = acaoRdoFiltros_(corpo);            break;
      case 'tecnicosLista':          resposta = acaoTecnicosLista_(corpo);         break;

      default:               resposta = { ok: false, motivo: 'ACAO_DESCONHECIDA' };
    }
  } catch (erro) {
    resposta = { ok: false, motivo: 'ERRO_INTERNO' };
    registrar_('', 'ERRO', 'FALHA', String(erro));
  }
  return ContentService
    .createTextOutput(JSON.stringify(resposta))
    .setMimeType(ContentService.MimeType.JSON);
}


/* ===========================================================================
   AÇÕES
   =========================================================================== */

function acaoLogin_(p) {
  const aba = abaUsuarios_();
  const matricula = String(p.matricula || '').trim();
  const senha = String(p.senha || '');
  if (!matricula || !senha) return { ok: false, motivo: 'CREDENCIAL' };

  const achado = acharPorMatricula_(aba, matricula);

  // Resposta idêntica para matrícula inexistente e senha errada.
  if (!achado) {
    registrar_(matricula, 'LOGIN', 'FALHA', 'matrícula inexistente');
    return { ok: false, motivo: 'CREDENCIAL' };
  }

  const u = achado.dados;

  if (String(u.status).toUpperCase() !== 'ATIVO') {
    registrar_(matricula, 'LOGIN', 'BLOQUEADO', 'status ' + u.status);
    return { ok: false, motivo: 'INATIVO' };
  }

  if (u.bloqueado_ate && new Date(u.bloqueado_ate) > new Date()) {
    registrar_(matricula, 'LOGIN', 'BLOQUEADO', 'tentativas excedidas');
    return { ok: false, motivo: 'BLOQUEADO', minutos: BLOQUEIO_MINUTOS };
  }

  if (hash_(senha, u.salt) !== u.senha_hash) {
    const n = Number(u.tentativas || 0) + 1;
    const mudanca = { tentativas: n };
    if (n >= MAX_TENTATIVAS) {
      mudanca.bloqueado_ate = new Date(Date.now() + BLOQUEIO_MINUTOS * 60000);
      mudanca.tentativas = 0;
      avisarBloqueio_(u);
    }
    gravar_(aba, achado.linha, mudanca);
    registrar_(matricula, 'LOGIN', 'FALHA', 'senha incorreta (' + n + ')');
    return { ok: false, motivo: 'CREDENCIAL' };
  }

  gravar_(aba, achado.linha, { tentativas: 0, bloqueado_ate: '', ultimo_acesso: new Date() });
  registrar_(matricula, 'LOGIN', 'OK', '');

  const primeiro = u.primeiro_acesso === true || String(u.primeiro_acesso).toUpperCase() === 'TRUE';

  if (primeiro) {
    // Ainda não abre sessão: o navegador só recebe o suficiente para a tela de troca.
    return { ok: true, primeiroAcesso: true, usuario: { matricula: matricula, cpf: u.cpf } };
  }

  return {
    ok: true,
    primeiroAcesso: false,
    token: abrirSessao_(matricula),
    usuario: publico_(u)
  };
}


function acaoDefinirSenha_(p) {
  const aba = abaUsuarios_();
  const matricula = String(p.matricula || '').trim();
  const achado = acharPorMatricula_(aba, matricula);
  if (!achado) return { ok: false, motivo: 'CREDENCIAL' };

  const u = achado.dados;

  // Reconfere a senha atual: sem isso qualquer um trocaria a senha de qualquer matrícula.
  if (hash_(String(p.senhaAtual || ''), u.salt) !== u.senha_hash) {
    registrar_(matricula, 'DEFINIR_SENHA', 'FALHA', 'senha atual não confere');
    return { ok: false, motivo: 'CREDENCIAL' };
  }

  const erro = validarSenha_(String(p.senhaNova || ''), u.cpf);
  if (erro) return { ok: false, motivo: erro };

  const salt = gerarAleatorio_(16);
  gravar_(aba, achado.linha, {
    salt: salt,
    senha_hash: hash_(p.senhaNova, salt),
    primeiro_acesso: false,
    tentativas: 0,
    bloqueado_ate: ''
  });
  registrar_(matricula, 'DEFINIR_SENHA', 'OK', 'primeiro acesso');

  const atualizado = acharPorMatricula_(aba, matricula).dados;
  return { ok: true, token: abrirSessao_(matricula), usuario: publico_(atualizado) };
}


function acaoSolicitarTroca_(p) {
  const aba = abaUsuarios_();
  const email = String(p.email || '').trim().toLowerCase();
  const cpf = soDigitos_(p.cpf);
  const senhaNova = String(p.senhaNova || '');

  const achado = acharPorEmailCpf_(aba, email, cpf);

  // Só valida a força da senha se o par email+CPF existir — evita virar validador público.
  if (!achado) {
    registrar_('', 'SOLICITAR_TROCA', 'FALHA', email);
    return { ok: false, motivo: 'NAO_CONFERE' };
  }

  const u = achado.dados;
  const erro = validarSenha_(senhaNova, u.cpf);
  if (erro) return { ok: false, motivo: erro };

  const codigo = String(Math.floor(100000 + Math.random() * 900000));
  const saltNova = gerarAleatorio_(16);
  const agora = new Date();

  const abaRec = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_RECUPERACAO);
  invalidarPendentes_(abaRec, u.matricula);   // um pedido por vez
  abaRec.appendRow([
    Utilities.getUuid(),
    u.matricula,
    hash_(codigo, u.matricula),               // o código também não fica em texto
    hash_(senhaNova, saltNova),
    saltNova,
    agora,
    new Date(agora.getTime() + CODIGO_MINUTOS * 60000),
    false
  ]);

  enviarCodigo_(u, codigo);
  registrar_(u.matricula, 'SOLICITAR_TROCA', 'OK', mascararEmail_(u.email));

  return { ok: true, emailMascarado: mascararEmail_(u.email) };
}


function acaoConfirmarTroca_(p) {
  const codigo = soDigitos_(p.codigo);
  if (codigo.length !== 6) return { ok: false, motivo: 'CODIGO' };

  const trava = LockService.getScriptLock();
  trava.waitLock(10000);
  try {
    const abaRec = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_RECUPERACAO);
    const dados = abaRec.getDataRange().getValues();
    const cab = dados[0];
    const iMat = cab.indexOf('matricula');
    const iCod = cab.indexOf('codigo_hash');
    const iSen = cab.indexOf('senha_nova_hash');
    const iSal = cab.indexOf('salt_nova');
    const iExp = cab.indexOf('expira_em');
    const iUso = cab.indexOf('usado');

    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      if (linha[iUso] === true || String(linha[iUso]).toUpperCase() === 'TRUE') continue;
      if (hash_(codigo, String(linha[iMat])) !== linha[iCod]) continue;

      if (new Date(linha[iExp]) < new Date()) {
        registrar_(linha[iMat], 'CONFIRMAR_TROCA', 'FALHA', 'código expirado');
        return { ok: false, motivo: 'EXPIRADO' };
      }

      // Só aqui a senha nova passa a valer.
      const abaU = abaUsuarios_();
      const achado = acharPorMatricula_(abaU, String(linha[iMat]));
      if (!achado) return { ok: false, motivo: 'CODIGO' };

      gravar_(abaU, achado.linha, {
        senha_hash: linha[iSen],
        salt: linha[iSal],
        primeiro_acesso: false,
        tentativas: 0,
        bloqueado_ate: ''
      });
      abaRec.getRange(i + 1, iUso + 1).setValue(true);
      registrar_(linha[iMat], 'CONFIRMAR_TROCA', 'OK', '');
      return { ok: true };
    }

    registrar_('', 'CONFIRMAR_TROCA', 'FALHA', 'código inválido');
    return { ok: false, motivo: 'CODIGO' };

  } finally {
    trava.releaseLock();
  }
}


/**
 * { token } -> { ok, usuario }
 *
 * Chamada quando o portal recarrega (por exemplo, ao voltar de um módulo).
 * Confirma que o token ainda vale e devolve o cadastro atualizado, para a tela
 * não continuar mostrando dados antigos guardados no navegador.
 */
function acaoSessao_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  const achado = acharPorMatricula_(abaUsuarios_(), matricula);
  if (!achado) return { ok: false, motivo: 'SESSAO' };
  if (String(achado.dados.status).toUpperCase() !== 'ATIVO') return { ok: false, motivo: 'INATIVO' };

  return { ok: true, usuario: publico_(achado.dados) };
}


function acaoSalvarPerfil_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  const aba = abaUsuarios_();
  const achado = acharPorMatricula_(aba, matricula);
  if (!achado) return { ok: false, motivo: 'SESSAO' };

  const u = achado.dados;
  const ehAdmin = String(u.perfil).toUpperCase() === 'ADMIN';
  const permitidos = ehAdmin ? EDITAVEL_ADMIN : EDITAVEL_DONO;

  const mudancas = {};
  Object.keys(p.alteracoes || {}).forEach(function (chave) {
    const vazio = !String(u[chave] == null ? '' : u[chave]).trim();
    const podePreencher = !ehAdmin && vazio && PREENCHIVEL_UMA_VEZ.indexOf(chave) > -1;
    if (permitidos.indexOf(chave) === -1 && !podePreencher) return;  // ignora o resto
    let v = String(p.alteracoes[chave]).trim();
    if (chave === 'email') {
      v = v.toLowerCase();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v)) return;
    }
    if (chave === 'telefone') v = soDigitos_(v).length < 10 ? '' : v;
    mudancas[chave] = v;
  });

  if (!Object.keys(mudancas).length) return { ok: false, motivo: 'NADA' };

  gravar_(aba, achado.linha, mudancas);
  registrar_(matricula, 'SALVAR_PERFIL', 'OK', Object.keys(mudancas).join(','));

  return { ok: true, usuario: publico_(acharPorMatricula_(aba, matricula).dados) };
}


/* ===========================================================================
   SESSÃO
   =========================================================================== */

function abrirSessao_(matricula) {
  const token = Utilities.getUuid() + '.' + gerarAleatorio_(16);
  CacheService.getScriptCache().put('sessao_' + token, matricula, SESSAO_HORAS * 3600);
  return token;
}

/** Renova a validade da sessão a cada uso, para ninguém ser deslogado no meio do trabalho. */
function renovarSessao_(token, matricula) {
  CacheService.getScriptCache().put('sessao_' + String(token), matricula, SESSAO_HORAS * 3600);
}

function validarSessao_(token) {
  if (!token) return null;
  const matricula = CacheService.getScriptCache().get('sessao_' + String(token));
  if (!matricula) return null;
  renovarSessao_(token, matricula);
  return matricula;
}


/* ===========================================================================
   EMAIL
   =========================================================================== */

function enviarCodigo_(u, codigo) {
  const html =
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:520px;color:#2c343a">' +
      '<div style="background:#14181b;color:#ecede8;padding:20px 24px">' +
        '<div style="font-size:12px;letter-spacing:2px;text-transform:uppercase">Extreme Wind Blade Services</div>' +
        '<div style="font-size:19px;font-weight:bold;margin-top:4px">Confirmação de troca de senha</div>' +
      '</div>' +
      '<div style="padding:24px;border:1px solid #d2d5ce;border-top:0">' +
        '<p>Olá, ' + primeiroNome_(u.nome) + '.</p>' +
        '<p>Recebemos um pedido para trocar a senha da matrícula <strong>' + u.matricula + '</strong>. ' +
           'Use o código abaixo na tela de confirmação:</p>' +
        '<div style="font-size:30px;font-weight:bold;letter-spacing:8px;text-align:center;' +
             'padding:16px;background:#ecede8;border-left:4px solid #f0a500;margin:20px 0">' +
             codigo + '</div>' +
        '<p style="font-size:13px;color:#55626b">O código vale por ' + CODIGO_MINUTOS + ' minutos. ' +
           'Sua senha atual continua funcionando até você confirmar.</p>' +
        '<p style="font-size:13px;color:#a8342a"><strong>Não foi você?</strong> Ignore este email e avise o ' +
           'setor administrativo — alguém tentou alterar o acesso da sua matrícula.</p>' +
      '</div>' +
      '<div style="padding:14px 24px;font-size:11px;color:#8c979e">Mensagem automática. Não responda.</div>' +
    '</div>';

  MailApp.sendEmail({
    to: u.email,
    subject: 'Código de confirmação — ' + NOME_SISTEMA,
    htmlBody: html,
    body: 'Código de confirmação: ' + codigo + ' (vale ' + CODIGO_MINUTOS + ' minutos).',
    name: NOME_SISTEMA
  });
}

function avisarBloqueio_(u) {
  try {
    MailApp.sendEmail({
      to: u.email,
      subject: 'Acesso bloqueado temporariamente — ' + NOME_SISTEMA,
      body: 'Olá, ' + primeiroNome_(u.nome) + '.\n\n' +
            'A matrícula ' + u.matricula + ' teve ' + MAX_TENTATIVAS + ' tentativas de senha incorretas ' +
            'e ficou bloqueada por ' + BLOQUEIO_MINUTOS + ' minutos.\n\n' +
            'Se não foi você, avise o setor administrativo.',
      name: NOME_SISTEMA
    });
  } catch (e) { /* cota de email estourada não pode derrubar o login */ }
}


/* ===========================================================================
   ACESSO À PLANILHA
   =========================================================================== */

function abaUsuarios_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_USUARIOS);
}

function acharPorMatricula_(aba, matricula) {
  const dados = aba.getDataRange().getValues();
  const cab = dados[0];
  const i = cab.indexOf('matricula');
  for (let l = 1; l < dados.length; l++) {
    if (String(dados[l][i]).trim() === String(matricula).trim()) {
      return { linha: l + 1, dados: objeto_(cab, dados[l]) };
    }
  }
  return null;
}

function acharPorEmailCpf_(aba, email, cpf) {
  const dados = aba.getDataRange().getValues();
  const cab = dados[0];
  const iE = cab.indexOf('email');
  const iC = cab.indexOf('cpf');
  for (let l = 1; l < dados.length; l++) {
    const mesmoEmail = String(dados[l][iE]).trim().toLowerCase() === email;
    const mesmoCpf = soDigitos_(dados[l][iC]) === cpf;
    if (mesmoEmail && mesmoCpf) return { linha: l + 1, dados: objeto_(cab, dados[l]) };
  }
  return null;
}

function objeto_(cab, linha) {
  const o = {};
  cab.forEach(function (c, i) {
    o[c] = (c === 'cpf' || c === 'matricula') ? String(linha[i]).replace(/^'/, '').trim() : linha[i];
  });
  return o;
}

function gravar_(aba, linha, mudancas) {
  const cab = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  Object.keys(mudancas).forEach(function (chave) {
    const col = cab.indexOf(chave);
    if (col > -1) aba.getRange(linha, col + 1).setValue(mudancas[chave]);
  });
}

function invalidarPendentes_(abaRec, matricula) {
  const dados = abaRec.getDataRange().getValues();
  const cab = dados[0];
  const iMat = cab.indexOf('matricula');
  const iUso = cab.indexOf('usado');
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][iMat]) === String(matricula) && dados[i][iUso] !== true) {
      abaRec.getRange(i + 1, iUso + 1).setValue(true);
    }
  }
}

/** Devolve só o que o navegador precisa. CPF sai mascarado. */
function publico_(u) {
  const cpf = soDigitos_(u.cpf);
  return {
    matricula:  String(u.matricula),
    nome:       u.nome,
    email:      u.email,
    setor:      u.setor,
    cargo:      u.cargo,
    perfil:     String(u.perfil).toUpperCase(),
    admissao:   formatarData_(u.admissao),
    nascimento: formatarData_(u.nascimento),
    telefone:   u.telefone,
    emergencia: u.emergencia,
    cpf:        '•••.•••.' + cpf.slice(6, 9) + '-' + cpf.slice(9)
  };
}

function registrar_(matricula, acao, resultado, detalhe) {
  try {
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(ABA_LOG)
      .appendRow([new Date(), String(matricula), acao, resultado, detalhe || '']);
  } catch (e) { /* log nunca derruba a operação principal */ }
}


/* ===========================================================================
   AUXILIARES
   =========================================================================== */

function hash_(texto, salt) {
  const pepper = PropertiesService.getScriptProperties().getProperty('PEPPER');
  if (!pepper) throw new Error('PEPPER não configurado. Rode definirPepper().');
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt) + String(texto) + pepper,
    Utilities.Charset.UTF_8
  );
  return bytes.map(function (b) {
    return ((b < 0 ? b + 256 : b).toString(16)).padStart(2, '0');
  }).join('');
}

function gerarAleatorio_(tamanho) {
  const letras = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let s = '';
  for (let i = 0; i < tamanho; i++) s += letras.charAt(Math.floor(Math.random() * letras.length));
  return s;
}

function validarSenha_(senha, cpf) {
  if (senha.length < 8) return 'SENHA_CURTA';
  if (!/[A-Za-z]/.test(senha) || !/\d/.test(senha)) return 'SENHA_FRACA';
  if (soDigitos_(senha) === soDigitos_(cpf)) return 'SENHA_IGUAL_CPF';
  return null;
}

function soDigitos_(v) {
  return String(v == null ? '' : v).replace(/\D/g, '');
}

function mascararEmail_(email) {
  const p = String(email).split('@');
  const vis = p[0].slice(0, Math.min(2, p[0].length));
  return vis + Array(Math.max(3, p[0].length - 2) + 1).join('•') + '@' + p[1];
}

function primeiroNome_(nome) {
  return String(nome || '').trim().split(/\s+/)[0] || '';
}

function formatarData_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}


/* ===========================================================================
   MANUTENÇÃO
   Crie um acionador diário para limparExpirados() em Acionadores > Adicionar.
   =========================================================================== */

function limparExpirados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Códigos de recuperação com mais de 7 dias
  const abaRec = ss.getSheetByName(ABA_RECUPERACAO);
  const rec = abaRec.getDataRange().getValues();
  const iExp = rec[0].indexOf('expira_em');
  const limiteRec = new Date(Date.now() - 7 * 24 * 3600 * 1000);
  const linhasRec = [];
  for (let i = 1; i < rec.length; i++) {
    if (rec[i][iExp] && new Date(rec[i][iExp]) < limiteRec) linhasRec.push(i + 1);
  }
  apagarLinhas_(abaRec, linhasRec);

  // LOG: prazo diferente conforme o tipo do registro
  const abaLog = ss.getSheetByName(ABA_LOG);
  const log = abaLog.getDataRange().getValues();
  const corteSeg = new Date(Date.now() - DIAS_LOG_SEGURANCA * 24 * 3600 * 1000);
  const corteOpe = new Date(Date.now() - DIAS_LOG_OPERACAO  * 24 * 3600 * 1000);

  const linhasLog = [];
  for (let i = 1; i < log.length; i++) {
    const quando = log[i][0];
    if (!quando) continue;
    const seguranca = ACOES_SEGURANCA.indexOf(String(log[i][2])) > -1;
    const corte = seguranca ? corteSeg : corteOpe;
    if (new Date(quando) < corte) linhasLog.push(i + 1);
  }
  apagarLinhas_(abaLog, linhasLog);

  Logger.log('Removidas ' + linhasRec.length + ' recuperações e ' +
             linhasLog.length + ' linhas de log.');
}

/**
 * Apaga várias linhas de uma vez, agrupando as que são vizinhas.
 * Apagar uma a uma é o que estoura o limite de 6 minutos do Apps Script quando
 * há dezenas de milhares de linhas acumuladas.
 */
function apagarLinhas_(aba, linhas) {
  if (!linhas.length) return;
  linhas.sort(function (a, b) { return b - a; });   // de baixo para cima

  let fim = linhas[0];
  let ini = fim;
  for (let i = 1; i <= linhas.length; i++) {
    if (i < linhas.length && linhas[i] === ini - 1) { ini = linhas[i]; continue; }
    aba.deleteRows(ini, fim - ini + 1);
    if (i < linhas.length) { fim = linhas[i]; ini = fim; }
  }
}
