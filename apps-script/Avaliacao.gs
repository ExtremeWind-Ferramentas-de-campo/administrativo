/**
 * ============================================================================
 * MÓDULO: AVALIAÇÃO DE PROJETOS
 * Cole como novo arquivo (Arquivo > + > Script, nome "Avaliacao") no mesmo
 * projeto do portal. Usa os utilitários do SupervisaoCampo.gs (prop_,
 * lerProjetos_, chaveTexto_, normalizarCab_, dataISO_) e o acesso ao Dropbox
 * do StatusEHS.gs (dbxPortal_).
 *
 * O QUE FAZ
 *  - Projeto com DATA REAL DE FINALIZAÇÃO (campo "fim" do card em Projetos em
 *    Andamento) é considerado concluído e aparece para ser avaliado.
 *  - Cada setor responde as perguntas que a planilha
 *      Matriz avaliação de projetos.xlsx   (Dropbox)
 *    define para ele. Coluna INDIVIDUAL marcada = uma nota por técnico;
 *    coluna EQUIPE marcada = uma nota para a equipe toda.
 *  - O setor OPERACIONAL informa também o TIPO DE REPARO executado.
 *
 * DE ONDE VÊM AS PERGUNTAS
 *  A matriz é lida direto do Dropbox. A cada 5 minutos o portal confere a
 *  revisão do arquivo: se alguém salvou a planilha, as perguntas novas
 *  aparecem sozinhas, sem mexer em código. A última versão lida fica guardada
 *  e é usada se o Dropbox estiver fora do ar.
 *
 *  Formato esperado (qualquer aba, cabeçalho em qualquer uma das 10 primeiras
 *  linhas): SETOR | AVALIAÇÃO | INDIVIDUAL | EQUIPE. Qualquer texto nas
 *  colunas INDIVIDUAL/EQUIPE ("X", "SIM") conta como marcado. Setor em branco
 *  repete o de cima (célula mesclada). Pergunta sem nenhuma das duas marcas é
 *  ignorada — aparece em Portal > Avaliação: conferir matriz.
 *
 * ONDE FICAM AS RESPOSTAS
 *  BASE:    aba AVAL_RESPOSTAS desta planilha (uma linha por nota).
 *  ESPELHO: Banco de avaliações de projetos.xlsx, na mesma pasta da matriz no
 *           Dropbox, reescrito inteiro a cada avaliação salva. É espelho, não
 *           base: o que for digitado nele à mão se perde na próxima gravação.
 *           Se o espelho falhar, a avaliação fica salva mesmo assim (falha no
 *           LOG como AVAL_ESPELHO) e a próxima gravação corrige.
 *
 * PROPRIEDADES DO SCRIPT (opcionais — os padrões já apontam para a pasta certa)
 *   AVAL_MATRIZ_DROPBOX   caminho da matriz no Dropbox
 *   AVAL_BANCO_DROPBOX    caminho da planilha de avaliações no Dropbox
 *  O acesso ao Dropbox usa as mesmas DROPBOX_APP_KEY / DROPBOX_APP_SECRET /
 *  DROPBOX_REFRESH_TOKEN do Status RD EHS. O app do Dropbox precisa da
 *  permissão files.content.write para gravar o espelho.
 *
 * Diagnóstico à mão: testarAvaliacao()
 * ============================================================================
 */

const AVAL_PASTA_PADRAO  = '/02 - EXTREME WIND/17 - SITE E APP (DASHBOARDS)/Gerenciamento de Equipes';
const AVAL_MATRIZ_PADRAO = AVAL_PASTA_PADRAO + '/Matriz avaliação de projetos.xlsx';
const AVAL_BANCO_PADRAO  = AVAL_PASTA_PADRAO + '/Banco de avaliações de projetos.xlsx';

const ABA_AVAL = 'AVAL_RESPOSTAS';
const COLUNAS_AVAL = [
  'id_avaliacao', 'data_avaliacao', 'projeto_id', 'projeto_codigo', 'parque', 'cliente',
  'tipo_reparo_projeto', 'tipo_reparo_operacoes', 'data_inicio', 'data_fim', 'supervisor',
  'setor', 'avaliador_matricula', 'avaliador_nome', 'tipo_avaliacao',
  'tecnico_matricula', 'tecnico_nome', 'criterio', 'nota', 'observacao'
];
const AVAL_COL = {};
COLUNAS_AVAL.forEach(function (c, i) { AVAL_COL[c] = i; });

/* Escala das notas. 'N/A' é aceito à parte (não entra na média). */
const AVAL_NOTAS = [1, 2, 3, 4, 5];
const AVAL_NA = 'N/A';
const AVAL_OBS_MAX = 500;

/* Setor que informa o tipo de reparo executado. Comparado sem acento. */
const AVAL_SETOR_TIPO_REPARO = 'OPERACIONAL';

/* Perfis que podem avaliar por QUALQUER setor (cobrir quem está de férias,
   corrigir avaliação). Os demais avaliam só pelo próprio setor. */
const AVAL_PERFIS_QUALQUER_SETOR = ['ADMIN'];

const AVAL_CACHE_MATRIZ_SEG = 300;

/* O setor do cadastro (aba USUARIOS) é texto livre. Aqui ficam outras grafias
   aceitas para cada setor da matriz — chave e apelidos sem acento, minúsculos.
   Setor da matriz que não estiver aqui casa só pelo nome igual. */
const AVAL_APELIDOS_SETOR = {
  'operacional':  ['operacoes', 'operacao', 'operacional de campo', 'campo'],
  'rh e dp':      ['rh', 'dp', 'rh/dp', 'rh dp', 'recursos humanos', 'departamento pessoal', 'departamento de pessoal'],
  'ehs':          ['sms', 'hse', 'qsms', 'seguranca', 'seguranca do trabalho', 'seguranca e meio ambiente'],
  'medicao':      ['medicoes', 'faturamento'],
  'engenharia':   ['eng'],
  'frota':        ['frotas', 'transporte', 'logistica'],
  'almoxarifado': ['almox', 'suprimentos', 'estoque']
};


/* ===========================================================================
   INSTALAÇÃO E DIAGNÓSTICO
   =========================================================================== */

function configurarAvaliacao() {
  abaAval_(true);
  avisar_('Aba ' + ABA_AVAL + ' pronta.\n\nRode testarAvaliacao() para conferir a leitura da matriz ' +
          'e a gravação da planilha no Dropbox.');
}

/** Rode à mão: lê a matriz (forçando), mostra o que entendeu e regrava o espelho. */
function testarAvaliacao() {
  Logger.log('Matriz: ' + avalCaminhoMatriz_());
  Logger.log('Banco : ' + avalCaminhoBanco_());
  try {
    const m = lerMatrizAval_(true);
    Logger.log(descreverMatrizAval_(m));
  } catch (e) {
    Logger.log('NÃO consegui ler a matriz: ' + e.message);
    return;
  }
  const r = espelharAvaliacoesDropbox_();
  Logger.log(r.ok ? 'Planilha de avaliações gravada no Dropbox.' : 'Falha ao gravar no Dropbox: ' + r.erro);
}

/** Menu: Portal > Avaliação: conferir matriz */
function menuMatrizAval() {
  try {
    avisar_(descreverMatrizAval_(lerMatrizAval_(true)));
  } catch (e) {
    avisar_('Não consegui ler a matriz de avaliação.\n\n' + avalCaminhoMatriz_() + '\n\n' + e.message);
  }
}

/** Menu: Portal > Avaliação: regravar planilha no Dropbox */
function atualizarBancoAvaliacoesDropbox() {
  const r = espelharAvaliacoesDropbox_();
  avisar_(r.ok ? 'Planilha de avaliações regravada no Dropbox:\n\n' + avalCaminhoBanco_()
               : 'Falha ao gravar no Dropbox:\n\n' + r.erro);
}

function descreverMatrizAval_(m) {
  let txt = 'MATRIZ DE AVALIAÇÃO\n' + avalCaminhoMatriz_() + '\n';
  if (m.modificada) txt += 'Última alteração no Dropbox: ' + m.modificada + '\n';
  if (m.desatualizada) txt += 'ATENÇÃO: Dropbox fora do ar, usando a última versão lida.\n';
  txt += '\n';
  m.setores.forEach(function (s) {
    txt += s.nome + '\n';
    s.criterios.forEach(function (c) {
      txt += '   - ' + c.texto + '  [' + [c.individual ? 'individual' : '', c.equipe ? 'equipe' : '']
        .filter(String).join(' + ') + ']\n';
    });
  });
  if (m.ignoradas && m.ignoradas.length) {
    txt += '\nIgnoradas (sem X em INDIVIDUAL nem em EQUIPE):\n';
    m.ignoradas.forEach(function (i) { txt += '   - ' + i + '\n'; });
  }
  return txt;
}


/* ===========================================================================
   CAMINHOS E ABA
   =========================================================================== */

function avalCaminho_(t) {
  let p = String(t || '').replace(/\\/g, '/').trim();
  if (p.charAt(0) !== '/') p = '/' + p;
  return p;
}
function avalCaminhoMatriz_() { return avalCaminho_(prop_('AVAL_MATRIZ_DROPBOX', AVAL_MATRIZ_PADRAO)); }
function avalCaminhoBanco_()  { return avalCaminho_(prop_('AVAL_BANCO_DROPBOX', AVAL_BANCO_PADRAO)); }

function abaAval_(criar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(ABA_AVAL);
  if (!aba && criar) {
    aba = ss.insertSheet(ABA_AVAL);
    aba.getRange(1, 1, 1, COLUNAS_AVAL.length).setValues([COLUNAS_AVAL])
       .setFontWeight('bold').setBackground('#14181b').setFontColor('#ecede8');
    aba.setFrozenRows(1);
    // Tudo como texto, menos a nota: sem isso o Sheets come o zero à esquerda
    // da matrícula e transforma "2026-10-19" em data.
    aba.getRange(1, 1, aba.getMaxRows(), COLUNAS_AVAL.length).setNumberFormat('@');
    aba.getRange(1, AVAL_COL.nota + 1, aba.getMaxRows(), 1).setNumberFormat('0');
  }
  return aba;
}


/* ===========================================================================
   MATRIZ (Dropbox) — leitura do .xlsx sem converter para Google Sheets
   =========================================================================== */

/**
 * -> { setores:[{nome, criterios:[{id, texto, individual, equipe}]}], ignoradas[],
 *      rev, modificada, desatualizada? }
 * Cache de 5 minutos. Passado o cache, confere a revisão do arquivo no Dropbox
 * e só baixa de novo se mudou.
 */
function lerMatrizAval_(forcar) {
  const cache = CacheService.getScriptCache();
  if (!forcar) {
    const g = cache.get('aval_matriz');
    if (g) { try { return JSON.parse(g); } catch (e) { /* relê */ } }
  }

  const props = PropertiesService.getScriptProperties();
  let guardada = null;
  try { guardada = JSON.parse(props.getProperty('AVAL_MATRIZ_ULTIMA') || 'null'); } catch (e) { guardada = null; }

  let matriz;
  try {
    const caminho = avalCaminhoMatriz_();
    const meta = JSON.parse(dbxPortal_('https://api.dropboxapi.com/2/files/get_metadata',
                                       { path: caminho }).getContentText());
    if (!forcar && guardada && guardada.rev === meta.rev) {
      matriz = guardada;
    } else {
      const blob = dbxPortal_('https://content.dropboxapi.com/2/files/download', { path: caminho }, true).getBlob();
      matriz = interpretarMatrizAval_(lerXlsx_(blob));
      matriz.rev = meta.rev;
      matriz.modificada = Utilities.formatDate(new Date(meta.server_modified), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
      if (!matriz.setores.length) throw new Error('A matriz não tem nenhuma pergunta reconhecida (cabeçalho SETOR | AVALIAÇÃO | INDIVIDUAL | EQUIPE).');
      // última versão boa, para o dia em que o Dropbox não responder
      try { props.setProperty('AVAL_MATRIZ_ULTIMA', JSON.stringify(matriz)); } catch (e) { /* grande demais: segue sem */ }
    }
  } catch (e) {
    if (!guardada) throw e;
    matriz = guardada;
    matriz.desatualizada = true;
    registrar_('', 'AVAL_MATRIZ', 'FALHA', String(e.message).slice(0, 300));
  }

  try { cache.put('aval_matriz', JSON.stringify(matriz), matriz.desatualizada ? 60 : AVAL_CACHE_MATRIZ_SEG); } catch (e) { /* segue */ }
  return matriz;
}

/** planilhas: [{nome, linhas:[[texto...]]}] -> estrutura da matriz */
function interpretarMatrizAval_(planilhas) {
  const setores = [], porSetor = {}, ignoradas = [];
  const marcado = function (v) {
    const t = chaveTexto_(v);
    return !!t && ['nao', 'n', '0', '-', 'false', 'falso'].indexOf(t) === -1;
  };

  planilhas.forEach(function (pl) {
    // acha o cabeçalho nas 10 primeiras linhas
    let cab = -1, cS = -1, cA = -1, cI = -1, cE = -1;
    for (let l = 0; l < Math.min(10, pl.linhas.length) && cab === -1; l++) {
      const norm = pl.linhas[l].map(normalizarCab_);
      const s = acharColuna_(norm, ['setor', 'area', 'departamento']);
      const a = acharColuna_(norm, ['avaliacao', 'criterio', 'pergunta', 'item']);
      if (s > -1 && a > -1) {
        cab = l; cS = s; cA = a;
        cI = acharColuna_(norm, ['individual', 'tecnico']);
        cE = acharColuna_(norm, ['equipe', 'time']);
      }
    }
    if (cab === -1) return;

    let setorAtual = '';
    for (let l = cab + 1; l < pl.linhas.length; l++) {
      const lin = pl.linhas[l];
      const s = String(lin[cS] || '').replace(/\s+/g, ' ').trim();
      if (s) setorAtual = s.toUpperCase();
      const texto = String(lin[cA] || '').replace(/\s+/g, ' ').trim();
      if (!texto || !setorAtual) continue;

      const ind = cI > -1 && marcado(lin[cI]);
      const eq  = cE > -1 && marcado(lin[cE]);
      if (!ind && !eq) { ignoradas.push(setorAtual + ' — ' + texto); continue; }

      const ks = chaveTexto_(setorAtual);
      if (!porSetor[ks]) { porSetor[ks] = { nome: setorAtual, criterios: [] }; setores.push(porSetor[ks]); }
      const id = chaveTexto_(texto);
      const ja = porSetor[ks].criterios.filter(function (c) { return c.id === id; })[0];
      if (ja) { ja.individual = ja.individual || ind; ja.equipe = ja.equipe || eq; continue; }
      porSetor[ks].criterios.push({ id: id, texto: texto, individual: ind, equipe: eq });
    }
  });

  return { setores: setores, ignoradas: ignoradas };
}

/**
 * Lê um .xlsx (blob) e devolve [{nome, linhas:[[texto...]]}], uma por aba.
 * Só o conteúdo das células — sem fórmulas, estilos ou datas formatadas, que a
 * matriz não usa.
 */
function lerXlsx_(blob) {
  blob.setContentType('application/zip');
  const arq = {};
  Utilities.unzip(blob).forEach(function (b) { arq[b.getName()] = b; });
  const txt = function (n) { return arq[n] ? arq[n].getDataAsString('UTF-8') : ''; };

  const compartilhadas = [];
  const ssXml = txt('xl/sharedStrings.xml');
  if (ssXml) {
    const raiz = XmlService.parse(ssXml).getRootElement();
    raiz.getChildren('si', raiz.getNamespace()).forEach(function (si) { compartilhadas.push(textoXml_(si)); });
  }

  const rels = {};
  const relXml = txt('xl/_rels/workbook.xml.rels');
  if (relXml) {
    const rr = XmlService.parse(relXml).getRootElement();
    rr.getChildren('Relationship', rr.getNamespace()).forEach(function (r) {
      rels[r.getAttribute('Id').getValue()] = r.getAttribute('Target').getValue();
    });
  }

  const wb = XmlService.parse(txt('xl/workbook.xml')).getRootElement();
  const nsWb = wb.getNamespace();
  const nsR = XmlService.getNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
  const planilhas = [];

  wb.getChild('sheets', nsWb).getChildren('sheet', nsWb).forEach(function (s, i) {
    const rid = s.getAttribute('id', nsR);
    let alvo = (rid && rels[rid.getValue()]) || ('worksheets/sheet' + (i + 1) + '.xml');
    alvo = alvo.charAt(0) === '/' ? alvo.slice(1) : 'xl/' + alvo;
    const xml = txt(alvo);
    if (!xml) return;

    const raiz = XmlService.parse(xml).getRootElement();
    const ns = raiz.getNamespace();
    const dados = raiz.getChild('sheetData', ns);
    const linhas = [];
    if (dados) {
      dados.getChildren('row', ns).forEach(function (row, iRow) {
        const rAttr = row.getAttribute('r');
        const nLinha = rAttr ? Number(rAttr.getValue()) - 1 : iRow;
        const lin = [];
        row.getChildren('c', ns).forEach(function (c, iCel) {
          const ref = c.getAttribute('r');
          const col = ref ? colunaDaRef_(ref.getValue()) : iCel;
          const t = c.getAttribute('t') ? c.getAttribute('t').getValue() : '';
          const v = c.getChild('v', ns);
          let val = '';
          if (t === 's') val = v ? (compartilhadas[Number(v.getText())] || '') : '';
          else if (t === 'inlineStr') val = c.getChild('is', ns) ? textoXml_(c.getChild('is', ns)) : '';
          else val = v ? v.getText() : '';
          lin[col] = val;
        });
        for (let k = 0; k < lin.length; k++) if (lin[k] == null) lin[k] = '';
        linhas[nLinha] = lin;
      });
    }
    for (let k = 0; k < linhas.length; k++) if (!linhas[k]) linhas[k] = [];
    planilhas.push({ nome: s.getAttribute('name').getValue(), linhas: linhas });
  });
  return planilhas;
}

/** Junta todo <t> dentro do elemento (texto com formatação vem em vários <r><t>). */
function textoXml_(el) {
  let s = '';
  el.getChildren().forEach(function (f) {
    const n = f.getName();
    if (n === 't') s += f.getText();
    else if (n !== 'rPh' && n !== 'phoneticPr') s += textoXml_(f);
  });
  return s;
}

function colunaDaRef_(ref) {
  const letras = String(ref).match(/^[A-Z]+/i);
  if (!letras) return 0;
  let n = 0;
  const l = letras[0].toUpperCase();
  for (let i = 0; i < l.length; i++) n = n * 26 + (l.charCodeAt(i) - 64);
  return n - 1;
}


/* ===========================================================================
   SETOR DE QUEM AVALIA
   =========================================================================== */

function usuarioAval_(matricula) {
  const achado = acharPorMatricula_(abaUsuarios_(), matricula);
  const d = achado ? achado.dados : {};
  return {
    matricula: String(matricula),
    nome: String(d.nome || ''),
    setor: String(d.setor || ''),
    perfil: String(d.perfil || 'DIRETORIA').toUpperCase()
  };
}

/** Nome do setor da matriz que corresponde ao setor do cadastro, ou ''. */
function setorNaMatriz_(setorCadastro, matriz) {
  const u = chaveTexto_(setorCadastro);
  if (!u) return '';
  const nomes = matriz.setores.map(function (s) { return s.nome; });

  for (let i = 0; i < nomes.length; i++) if (chaveTexto_(nomes[i]) === u) return nomes[i];
  for (let i = 0; i < nomes.length; i++) {
    const apel = AVAL_APELIDOS_SETOR[chaveTexto_(nomes[i])] || [];
    if (apel.indexOf(u) > -1) return nomes[i];
  }
  // "Engenharia de campo" -> ENGENHARIA
  for (let i = 0; i < nomes.length; i++) {
    const k = chaveTexto_(nomes[i]);
    if (k.length >= 3 && (u.indexOf(k) > -1 || k.indexOf(u) > -1)) return nomes[i];
  }
  return '';
}

function podeQualquerSetorAval_(perfil) {
  return AVAL_PERFIS_QUALQUER_SETOR.indexOf(String(perfil).toUpperCase()) > -1;
}


/* ===========================================================================
   PROJETOS CONCLUÍDOS
   =========================================================================== */

/* Concluído = tem data real de finalização. O status sozinho não basta: um
   card marcado "concluído" sem data de fim ainda está pela metade. */
function projetoConcluidoAval_(pr) {
  return !!dataISO_(pr && pr.fim);
}

/** Todos os técnicos que passaram pelo projeto (todas as formações), sem repetir. */
function tecnicosDoProjetoAval_(pr) {
  const vistos = {}, lista = [];
  const add = function (t) {
    if (!t) return;
    const nome = String(t.nome || '').trim();
    const mat = String(t.matricula || '').replace(/^'/, '').trim();
    if (!nome && !mat) return;
    const k = chaveTecAval_(mat, nome);
    if (vistos[k]) return;
    vistos[k] = true;
    lista.push({ matricula: mat, nome: nome || ('matrícula ' + mat), nivel: String(t.nivel || '') });
  };
  (pr.tecnicos || []).forEach(add);
  (Array.isArray(pr.equipes) ? pr.equipes : []).forEach(function (e) { (e && e.tecnicos || []).forEach(add); });
  return lista;
}

function chaveTecAval_(matricula, nome) {
  return matricula ? 'm:' + String(matricula).trim() : 'n:' + chaveTexto_(nome);
}

function resumoProjetoAval_(pr) {
  return {
    id: String(pr.id),
    codigo: pr.codigo || '',
    parque: pr.parque || '',
    cliente: pr.cliente || '',
    tipoReparo: pr.tipoReparo || '',
    inicio: dataISO_(pr.inicio),
    fim: dataISO_(pr.fim),
    supervisor: (pr.supervisor && pr.supervisor.nome) || '',
    tecnicos: tecnicosDoProjetoAval_(pr)
  };
}


/* ===========================================================================
   LEITURA DAS RESPOSTAS
   =========================================================================== */

function lerRespostasAval_() {
  const aba = abaAval_(false);
  if (!aba || aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, COLUNAS_AVAL.length).getDisplayValues();
}

/** -> { projetoId: { SETOR: { avaliador, matricula, em, media, notas, tipoReparo } } } */
function resumoAvaliacoesAval_(linhas) {
  const r = {};
  linhas.forEach(function (l) {
    const pid = String(l[AVAL_COL.projeto_id]);
    const setor = String(l[AVAL_COL.setor]);
    if (!pid || !setor) return;
    r[pid] = r[pid] || {};
    const s = r[pid][setor] = r[pid][setor] || {
      avaliador: l[AVAL_COL.avaliador_nome], matricula: l[AVAL_COL.avaliador_matricula],
      em: l[AVAL_COL.data_avaliacao], soma: 0, notas: 0, tipoReparo: ''
    };
    const n = Number(l[AVAL_COL.nota]);
    if (String(l[AVAL_COL.nota]) !== '' && !isNaN(n)) { s.soma += n; s.notas++; }
    if (l[AVAL_COL.tipo_reparo_operacoes]) s.tipoReparo = l[AVAL_COL.tipo_reparo_operacoes];
  });
  Object.keys(r).forEach(function (pid) {
    Object.keys(r[pid]).forEach(function (setor) {
      const s = r[pid][setor];
      s.media = s.notas ? Math.round(s.soma / s.notas * 10) / 10 : null;
      delete s.soma;
    });
  });
  return r;
}


/* ===========================================================================
   AÇÕES DO PORTAL
   =========================================================================== */

/** {} -> { ok, matriz, usuario, projetos[], avaliacoes{}, notas[] } */
function acaoAvaliacaoCarregar_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  let matriz;
  try { matriz = lerMatrizAval_(false); }
  catch (e) { return { ok: false, motivo: 'MATRIZ_INDISPONIVEL', detalhe: e.message }; }

  const u = usuarioAval_(matricula);
  const projetos = lerProjetos_().filter(projetoConcluidoAval_).map(resumoProjetoAval_);
  projetos.sort(function (a, b) { return String(b.fim).localeCompare(String(a.fim)); });

  return {
    ok: true,
    matriz: { setores: matriz.setores, modificada: matriz.modificada || '', desatualizada: !!matriz.desatualizada },
    notas: AVAL_NOTAS,
    setorTipoReparo: AVAL_SETOR_TIPO_REPARO,
    usuario: {
      nome: u.nome, setor: u.setor, perfil: u.perfil,
      setorMatriz: setorNaMatriz_(u.setor, matriz),
      qualquerSetor: podeQualquerSetorAval_(u.perfil)
    },
    projetos: projetos,
    avaliacoes: resumoAvaliacoesAval_(lerRespostasAval_())
  };
}

/** { projetoId, setor } -> { ok, respostas:[{tipo, matricula, nome, criterio, nota, obs}], tipoReparo } */
function acaoAvaliacaoAbrir_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  const id = idAvaliacao_(p.projetoId, p.setor);
  const respostas = [];
  let tipoReparo = '', avaliador = '', em = '';
  lerRespostasAval_().forEach(function (l) {
    if (String(l[AVAL_COL.id_avaliacao]) !== id) return;
    respostas.push({
      tipo: l[AVAL_COL.tipo_avaliacao],
      matricula: l[AVAL_COL.tecnico_matricula],
      nome: l[AVAL_COL.tecnico_nome],
      criterio: chaveTexto_(l[AVAL_COL.criterio]),
      nota: l[AVAL_COL.nota] === '' ? AVAL_NA : l[AVAL_COL.nota],
      obs: l[AVAL_COL.observacao]
    });
    if (l[AVAL_COL.tipo_reparo_operacoes]) tipoReparo = l[AVAL_COL.tipo_reparo_operacoes];
    avaliador = l[AVAL_COL.avaliador_nome]; em = l[AVAL_COL.data_avaliacao];
  });
  return { ok: true, respostas: respostas, tipoReparo: tipoReparo, avaliador: avaliador, em: em };
}

function idAvaliacao_(projetoId, setor) {
  return String(projetoId) + '|' + chaveTexto_(setor).toUpperCase();
}

/**
 * { projetoId, setor, tipoReparo, respostas:[{tipo:'EQUIPE'|'INDIVIDUAL', matricula, nome,
 *   criterio(id), nota, obs}] } -> { ok, espelho }
 * Substitui a avaliação anterior do mesmo setor para o mesmo projeto.
 */
function acaoAvaliacaoSalvar_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  let matriz;
  try { matriz = lerMatrizAval_(false); }
  catch (e) { return { ok: false, motivo: 'MATRIZ_INDISPONIVEL', detalhe: e.message }; }

  const u = usuarioAval_(matricula);
  const setorObj = matriz.setores.filter(function (s) { return chaveTexto_(s.nome) === chaveTexto_(p.setor); })[0];
  if (!setorObj) return { ok: false, motivo: 'SETOR_FORA_DA_MATRIZ' };

  // A tela só oferece o setor da pessoa, mas quem decide é aqui.
  const proprio = setorNaMatriz_(u.setor, matriz);
  if (!podeQualquerSetorAval_(u.perfil) && chaveTexto_(proprio) !== chaveTexto_(setorObj.nome)) {
    registrar_(matricula, 'AVAL_SALVAR', 'NEGADO', 'setor ' + u.setor + ' tentou avaliar como ' + setorObj.nome);
    return { ok: false, motivo: 'SEM_PERMISSAO' };
  }

  const pr = lerProjetos_().filter(function (x) { return String(x.id) === String(p.projetoId); })[0];
  if (!pr) return { ok: false, motivo: 'PROJETO_INEXISTENTE' };
  if (!projetoConcluidoAval_(pr)) return { ok: false, motivo: 'PROJETO_NAO_CONCLUIDO' };

  const ehOperacional = chaveTexto_(setorObj.nome) === chaveTexto_(AVAL_SETOR_TIPO_REPARO);
  const tipoReparoOp = ehOperacional ? String(p.tipoReparo || '').replace(/\s+/g, ' ').trim().slice(0, 200) : '';
  if (ehOperacional && !tipoReparoOp) return { ok: false, motivo: 'FALTA_TIPO_REPARO' };

  // respostas recebidas, indexadas por tipo + técnico + critério
  const recebidas = {};
  (p.respostas || []).forEach(function (r) {
    const tipo = String(r.tipo || '').toUpperCase();
    const k = tipo === 'EQUIPE' ? 'EQUIPE||' + r.criterio
                                : 'INDIVIDUAL|' + chaveTecAval_(r.matricula, r.nome) + '|' + r.criterio;
    recebidas[k] = r;
  });

  const tecnicos = tecnicosDoProjetoAval_(pr);
  const id = idAvaliacao_(pr.id, setorObj.nome);
  const agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const resumo = resumoProjetoAval_(pr);
  const linhas = [];
  let faltam = 0, invalidas = 0;

  const linha = function (tipo, tec, c, r) {
    const nota = notaValidaAval_(r && r.nota);
    if (nota === null) { if (r) invalidas++; else faltam++; return; }
    const l = new Array(COLUNAS_AVAL.length).fill('');
    l[AVAL_COL.id_avaliacao] = id;
    l[AVAL_COL.data_avaliacao] = agora;
    l[AVAL_COL.projeto_id] = String(pr.id);
    l[AVAL_COL.projeto_codigo] = resumo.codigo;
    l[AVAL_COL.parque] = resumo.parque;
    l[AVAL_COL.cliente] = resumo.cliente;
    l[AVAL_COL.tipo_reparo_projeto] = resumo.tipoReparo;
    l[AVAL_COL.tipo_reparo_operacoes] = tipoReparoOp;
    l[AVAL_COL.data_inicio] = resumo.inicio;
    l[AVAL_COL.data_fim] = resumo.fim;
    l[AVAL_COL.supervisor] = resumo.supervisor;
    l[AVAL_COL.setor] = setorObj.nome;
    l[AVAL_COL.avaliador_matricula] = String(matricula);
    l[AVAL_COL.avaliador_nome] = u.nome;
    l[AVAL_COL.tipo_avaliacao] = tipo;
    l[AVAL_COL.tecnico_matricula] = tec ? tec.matricula : '';
    l[AVAL_COL.tecnico_nome] = tec ? tec.nome : '';
    l[AVAL_COL.criterio] = c.texto;
    l[AVAL_COL.nota] = nota === AVAL_NA ? '' : nota;
    l[AVAL_COL.observacao] = String((r && r.obs) || '').replace(/\s+/g, ' ').trim().slice(0, AVAL_OBS_MAX);
    linhas.push(l);
  };

  setorObj.criterios.forEach(function (c) {
    if (c.equipe) linha('EQUIPE', null, c, recebidas['EQUIPE||' + c.id]);
    if (c.individual) {
      tecnicos.forEach(function (t) {
        linha('INDIVIDUAL', t, c, recebidas['INDIVIDUAL|' + chaveTecAval_(t.matricula, t.nome) + '|' + c.id]);
      });
    }
  });

  if (faltam || invalidas) return { ok: false, motivo: 'AVALIACAO_INCOMPLETA', faltam: faltam + invalidas };
  if (!linhas.length) return { ok: false, motivo: 'NADA_A_AVALIAR' };

  const trava = LockService.getScriptLock();
  if (!trava.tryLock(20000)) return { ok: false, motivo: 'OCUPADO' };
  try {
    const aba = abaAval_(true);
    apagarAvaliacaoAval_(aba, id);
    const inicio = aba.getLastRow() + 1;
    aba.getRange(inicio, 1, linhas.length, COLUNAS_AVAL.length).setValues(linhas);

    registrar_(matricula, 'AVAL_SALVAR', 'OK', id + ' | ' + linhas.length + ' notas');
    const espelho = espelharAvaliacoesDropbox_();
    return { ok: true, espelho: espelho.ok };
  } finally {
    trava.releaseLock();
  }
}

/** 1..5 -> número; 'N/A' -> 'N/A'; qualquer outra coisa -> null */
function notaValidaAval_(v) {
  if (v == null || v === '') return null;
  if (String(v).toUpperCase() === AVAL_NA) return AVAL_NA;
  const n = Number(v);
  return AVAL_NOTAS.indexOf(n) > -1 ? n : null;
}

/** Apaga as linhas de uma avaliação, em blocos contíguos, de baixo para cima. */
function apagarAvaliacaoAval_(aba, id) {
  if (aba.getLastRow() < 2) return;
  const ids = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getDisplayValues();
  const blocos = [];
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) !== id) continue;
    const linha = i + 2;
    const ult = blocos[blocos.length - 1];
    if (ult && ult.ini + ult.qtd === linha) ult.qtd++;
    else blocos.push({ ini: linha, qtd: 1 });
  }
  for (let b = blocos.length - 1; b >= 0; b--) aba.deleteRows(blocos[b].ini, blocos[b].qtd);
}


/* ===========================================================================
   ESPELHO NO DROPBOX — Banco de avaliações de projetos.xlsx
   Três abas:
     Respostas  uma linha por nota (igual à AVAL_RESPOSTAS, com o tipo de reparo
                das operações repetido em todas as linhas do projeto)
     Projetos   uma linha por projeto concluído: técnicos, setores que já
                avaliaram, pendentes e média
     Tecnicos   uma linha por técnico: projetos avaliados e média individual
   =========================================================================== */

const AVAL_XLSX_ABAS = {
  Respostas: {
    colunas: COLUNAS_AVAL,
    tipos: { data_avaliacao: 'datahora', data_inicio: 'data', data_fim: 'data', nota: 'num' }
  },
  Projetos: {
    colunas: ['projeto_id', 'projeto_codigo', 'parque', 'cliente', 'tipo_reparo_projeto', 'tipo_reparo_operacoes',
              'data_inicio', 'data_fim', 'supervisor', 'matriculas_tecnicos', 'tecnicos',
              'setores_avaliados', 'setores_pendentes', 'media_geral'],
    tipos: { data_inicio: 'data', data_fim: 'data', media_geral: 'num' }
  },
  Tecnicos: {
    colunas: ['tecnico_matricula', 'tecnico_nome', 'projetos_avaliados', 'notas_individuais', 'media_individual',
              'media_equipe'],
    tipos: { projetos_avaliados: 'num', notas_individuais: 'num', media_individual: 'num', media_equipe: 'num' }
  }
};

/** Monta e grava o .xlsx no Dropbox. Nunca lança: devolve { ok, erro }. */
function espelharAvaliacoesDropbox_() {
  try {
    const bytes = montarBancoAvalXlsx_();
    dbxPortal_('https://content.dropboxapi.com/2/files/upload',
               { path: avalCaminhoBanco_(), mode: 'overwrite', autorename: false, mute: true },
               true, bytes);
    return { ok: true };
  } catch (e) {
    registrar_('', 'AVAL_ESPELHO', 'FALHA', String(e.message).slice(0, 300));
    return { ok: false, erro: e.message };
  }
}

function montarBancoAvalXlsx_() {
  const respostas = lerRespostasAval_();

  // tipo de reparo informado pelas operações, por projeto
  const tipoOp = {};
  respostas.forEach(function (l) {
    if (l[AVAL_COL.tipo_reparo_operacoes]) tipoOp[l[AVAL_COL.projeto_id]] = l[AVAL_COL.tipo_reparo_operacoes];
  });

  const linhasResp = respostas.map(function (l) {
    const c = l.slice();
    if (!c[AVAL_COL.tipo_reparo_operacoes]) c[AVAL_COL.tipo_reparo_operacoes] = tipoOp[c[AVAL_COL.projeto_id]] || '';
    return c;
  });

  // projetos
  let setoresMatriz = [];
  try { setoresMatriz = lerMatrizAval_(false).setores.map(function (s) { return s.nome; }); } catch (e) { /* sem pendentes */ }
  const resumo = resumoAvaliacoesAval_(respostas);
  const linhasProj = lerProjetos_().filter(projetoConcluidoAval_).map(resumoProjetoAval_).map(function (pr) {
    const av = resumo[pr.id] || {};
    const feitos = Object.keys(av);
    const feitosK = feitos.map(chaveTexto_);
    let soma = 0, n = 0;
    feitos.forEach(function (s) { if (av[s].media != null) { soma += av[s].media * av[s].notas; n += av[s].notas; } });
    return [
      pr.id, pr.codigo, pr.parque, pr.cliente, pr.tipoReparo, tipoOp[pr.id] || '',
      pr.inicio, pr.fim, pr.supervisor,
      pr.tecnicos.map(function (t) { return t.matricula; }).filter(String).join(', '),
      pr.tecnicos.map(function (t) { return t.nome; }).join(', '),
      feitos.join(', '),
      setoresMatriz.filter(function (s) { return feitosK.indexOf(chaveTexto_(s)) === -1; }).join(', '),
      n ? Math.round(soma / n * 100) / 100 : ''
    ];
  });

  // técnicos: média das notas individuais e das notas de equipe dos projetos em que estiveram
  const tecs = {}, equipePorProj = {};
  respostas.forEach(function (l) {
    const n = Number(l[AVAL_COL.nota]);
    if (String(l[AVAL_COL.nota]) === '' || isNaN(n)) return;
    if (l[AVAL_COL.tipo_avaliacao] === 'EQUIPE') {
      const e = equipePorProj[l[AVAL_COL.projeto_id]] = equipePorProj[l[AVAL_COL.projeto_id]] || { soma: 0, n: 0 };
      e.soma += n; e.n++;
      return;
    }
    const k = chaveTecAval_(l[AVAL_COL.tecnico_matricula], l[AVAL_COL.tecnico_nome]);
    const t = tecs[k] = tecs[k] || { mat: l[AVAL_COL.tecnico_matricula], nome: l[AVAL_COL.tecnico_nome], proj: {}, soma: 0, n: 0 };
    t.proj[l[AVAL_COL.projeto_id]] = true; t.soma += n; t.n++;
  });
  const linhasTec = Object.keys(tecs).map(function (k) {
    const t = tecs[k];
    let se = 0, ne = 0;
    Object.keys(t.proj).forEach(function (pid) { const e = equipePorProj[pid]; if (e) { se += e.soma; ne += e.n; } });
    return [t.mat, t.nome, Object.keys(t.proj).length, t.n,
            Math.round(t.soma / t.n * 100) / 100, ne ? Math.round(se / ne * 100) / 100 : ''];
  }).sort(function (a, b) { return String(a[1]).localeCompare(String(b[1])); });

  return montarXlsx_([
    { nome: 'Respostas', def: AVAL_XLSX_ABAS.Respostas, linhas: linhasResp },
    { nome: 'Projetos',  def: AVAL_XLSX_ABAS.Projetos,  linhas: linhasProj },
    { nome: 'Tecnicos',  def: AVAL_XLSX_ABAS.Tecnicos,  linhas: linhasTec }
  ]).getBytes();
}


/* ---------------------------------------------------------------------------
   GERADOR DE .XLSX MÍNIMO
   Texto vai como inlineStr (sem tabela de strings compartilhadas); datas
   viram número de série do Excel com formato dd/mm/aaaa, para filtrar e
   ordenar direito no Excel. Cabeçalho congelado e destacado.
   --------------------------------------------------------------------------- */

function montarXlsx_(abas) {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NSR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const arquivos = [];
  const add = function (nome, conteudo) {
    arquivos.push(Utilities.newBlob(conteudo, 'application/xml', nome));
  };

  add('[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    abas.map(function (a, i) {
      return '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    }).join('') +
    '</Types>');

  add('_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>');

  add('xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="' + NS + '" xmlns:r="' + NSR + '"><sheets>' +
    abas.map(function (a, i) {
      return '<sheet name="' + xmlEsc_(a.nome) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>';
    }).join('') +
    '</sheets></workbook>');

  add('xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    abas.map(function (a, i) {
      return '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>';
    }).join('') +
    '<Relationship Id="rId' + (abas.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>');

  // estilos: 0 normal | 1 cabeçalho | 2 data | 3 data e hora
  add('xl/styles.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="' + NS + '">' +
    '<numFmts count="2"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/><numFmt numFmtId="165" formatCode="dd/mm/yyyy hh:mm"/></numFmts>' +
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>' +
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>' +
    '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF0E7C7B"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="4">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '</styleSheet>');

  abas.forEach(function (a, i) {
    const cols = a.def.colunas, tipos = a.def.tipos || {};
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="' + NS + '" xmlns:r="' + NSR + '">' +
      '<sheetViews><sheetView workbookViewId="0"' + (i === 0 ? ' tabSelected="1"' : '') + '>' +
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
      '<cols>' + cols.map(function (c, j) {
        const larg = Math.max(12, Math.min(45, c.length + 4));
        return '<col min="' + (j + 1) + '" max="' + (j + 1) + '" width="' + larg + '" customWidth="1"/>';
      }).join('') + '</cols><sheetData>';

    xml += '<row r="1">' + cols.map(function (c, j) { return celulaXlsx_(j, 1, c, 'texto', 1); }).join('') + '</row>';
    a.linhas.forEach(function (lin, k) {
      const r = k + 2;
      xml += '<row r="' + r + '">' + cols.map(function (c, j) {
        return celulaXlsx_(j, r, lin[j], tipos[c] || 'texto');
      }).join('') + '</row>';
    });
    xml += '</sheetData></worksheet>';
    add('xl/worksheets/sheet' + (i + 1) + '.xml', xml);
  });

  return Utilities.zip(arquivos, 'banco.xlsx');
}

function celulaXlsx_(col, linha, valor, tipo, estilo) {
  const ref = letraColuna_(col) + linha;
  const v = valor == null ? '' : valor;
  if (v === '') return '';
  if (tipo === 'num' && !isNaN(Number(v))) return '<c r="' + ref + '"><v>' + Number(v) + '</v></c>';
  if (tipo === 'data' || tipo === 'datahora') {
    const serie = serieExcel_(String(v));
    if (serie !== null) return '<c r="' + ref + '" s="' + (tipo === 'data' ? 2 : 3) + '"><v>' + serie + '</v></c>';
  }
  return '<c r="' + ref + '" t="inlineStr"' + (estilo ? ' s="' + estilo + '"' : '') +
         '><is><t xml:space="preserve">' + xmlEsc_(v) + '</t></is></c>';
}

/** 'aaaa-mm-dd' ou 'aaaa-mm-dd hh:mm' -> número de série do Excel */
function serieExcel_(t) {
  const m = t.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?/);
  if (!m) return null;
  const dias = (Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000;
  const fracao = m[4] ? (+m[4] * 60 + +m[5]) / 1440 : 0;
  return Math.round((dias + fracao) * 100000) / 100000;
}

function letraColuna_(i) {
  let s = '', n = i + 1;
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function xmlEsc_(t) {
  return String(t)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
