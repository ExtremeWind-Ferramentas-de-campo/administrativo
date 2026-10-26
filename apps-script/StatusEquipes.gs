/**
 * ============================================================================
 * MÓDULO: STATUS DAS EQUIPES (Supervisão de Campo)
 * Cole como novo arquivo (Arquivo > + > Script, nome "StatusEquipes") no
 * mesmo projeto do portal. Usa os utilitários do SupervisaoCampo.gs
 * (lerProjetos_, cobravelNaData_, projetoPassaFiltro_, chaveTexto_, dataISO_,
 * soNumero_, lerTecnicos_) e o registrar_/validarSessao_ do Codigo.gs.
 *
 * Todo dia um responsável diz, para cada parque em andamento, se a equipe
 * está em ATIVIDADE, STAND BY, ABSENTEÍSMO ou GARANTIA. Fora Atividade, a
 * observação é obrigatória.
 *
 * Quais parques aparecem: os mesmos da cobrança do Status RDO / RD EHS —
 * projeto em andamento que já tinha começado na data (cobravelNaData_).
 *
 * ABA EQUIPES_STATUS (planilha do portal, nasce sozinha na 1ª marcação)
 *   cliente | parque | data | supervisor | status | observacao |
 *   registrado_por | atualizado_em
 * Uma linha por parque e dia: marcar de novo substitui. A data é data de
 * verdade (dd/mm/aaaa); a linha é achada por data + parque.
 * ============================================================================
 */

const ABA_EQUIPES_STATUS = 'EQUIPES_STATUS';
const COLUNAS_EQUIPES_STATUS = ['cliente', 'parque', 'data', 'supervisor', 'status', 'observacao',
                                'registrado_por', 'atualizado_em'];

/* Opções da tela. obs = observação obrigatória. */
const STATUS_EQUIPE = {
  ATIVIDADE:   { rotulo: 'Atividade',   obs: false },
  STANDBY:     { rotulo: 'Stand by',    obs: true },
  ABSENTEISMO: { rotulo: 'Absenteísmo', obs: true },
  GARANTIA:    { rotulo: 'Garantia',    obs: true }
};
const OBS_EQUIPE_MAX = 500;

function abaEquipesStatus_(criar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(ABA_EQUIPES_STATUS);
  if (!aba && criar) {
    aba = ss.insertSheet(ABA_EQUIPES_STATUS);
    aba.getRange(1, 1, 1, COLUNAS_EQUIPES_STATUS.length).setValues([COLUNAS_EQUIPES_STATUS])
       .setFontWeight('bold').setBackground('#14181b').setFontColor('#ecede8');
    aba.setFrozenRows(1);
    aba.getRange('C:C').setNumberFormat('dd/mm/yyyy');
    aba.getRange('G:G').setNumberFormat('@');                   // matrícula: preserva zero à esquerda
    aba.getRange('H:H').setNumberFormat('dd/mm/yyyy hh:mm');
  }
  return aba;
}

/* A aba guarda o rótulo legível ("Stand by"); a tela usa a chave (STANDBY).
   Aceita os dois, sem acento/maiúscula, para quem editar a aba à mão. */
function statusDoTexto_(t) {
  const k = chaveTexto_(t).replace(/\s+/g, '');
  const achada = Object.keys(STATUS_EQUIPE).filter(function (s) {
    return chaveTexto_(s) === k || chaveTexto_(STATUS_EQUIPE[s].rotulo).replace(/\s+/g, '') === k;
  })[0];
  return achada || '';
}

/** -> { chaveTexto(parque): { status, observacao, matricula, em } } da data pedida */
function lerStatusEquipes_(dataIso) {
  const aba = abaEquipesStatus_(false);
  const mapa = {};
  if (!aba || aba.getLastRow() < 2) return mapa;
  const n = aba.getLastRow() - 1;
  const vals = aba.getRange(2, 1, n, COLUNAS_EQUIPES_STATUS.length).getValues();
  const exib = aba.getRange(2, 1, n, COLUNAS_EQUIPES_STATUS.length).getDisplayValues();
  vals.forEach(function (l, i) {
    if (dataISO_(l[2]) !== dataIso) return;
    mapa[chaveTexto_(l[1])] = {
      status: statusDoTexto_(l[4]), observacao: String(l[5]),
      matricula: String(l[6]).replace(/^'/, ''), em: exib[i][7]
    };
  });
  return mapa;
}

/**
 * { data, cliente, parque } -> { ok, data, opcoes, projetos:[{ id, parque, cliente,
 *   supervisor, tipoReparo, marcacao:{ status, observacao, por, em } | null }] }
 */
function acaoEquipesStatus_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  const data = dataISO_(p.data) || dataISO_(new Date());
  const fCliente = chaveTexto_(p.cliente || '');
  const fParque  = chaveTexto_(p.parque || '');

  let marcas = {};
  try { marcas = lerStatusEquipes_(data); } catch (e) { /* sem aba ainda */ }

  let nomePorMatricula = {};
  try {
    lerTecnicos_().forEach(function (t) { nomePorMatricula[soNumero_(t.matricula)] = t.nome; });
  } catch (e) { /* sem MINI MASTER, mostra a matrícula */ }
  try {
    lerSupervisores_().forEach(function (s) { nomePorMatricula[soNumero_(s.matricula)] = s.nome; });
  } catch (e) { /* idem */ }

  const projetos = lerProjetos_()
    .filter(function (pr) { return cobravelNaData_(pr, data); })
    .filter(function (pr) { return projetoPassaFiltro_(pr, fCliente, fParque); })
    .map(function (pr) {
      const m = marcas[chaveTexto_(pr.parque)];
      return {
        id: pr.id, parque: pr.parque || '', cliente: pr.cliente || '', tipoReparo: pr.tipoReparo || '',
        supervisor: (pr.supervisor && pr.supervisor.nome) || '',
        marcacao: m ? {
          status: m.status, observacao: m.observacao, em: m.em,
          por: nomePorMatricula[soNumero_(m.matricula)] || m.matricula
        } : null
      };
    })
    .sort(function (a, b) {
      return String(a.cliente).localeCompare(String(b.cliente)) || String(a.parque).localeCompare(String(b.parque));
    });

  const opcoes = Object.keys(STATUS_EQUIPE).map(function (k) {
    return { valor: k, rotulo: STATUS_EQUIPE[k].rotulo, obs: STATUS_EQUIPE[k].obs };
  });
  return { ok: true, data: data, opcoes: opcoes, projetos: projetos };
}

/**
 * { data, parque, status, observacao } -> { ok, marcacao }
 * status vazio = apagar a marcação do dia (clique errado).
 */
function acaoEquipesMarcar_(p) {
  const matricula = validarSessao_(p.token);
  if (!matricula) return { ok: false, motivo: 'SESSAO' };

  const data = dataISO_(p.data);
  const parque = String(p.parque || '').replace(/\s+/g, ' ').trim();
  const status = String(p.status || '').toUpperCase();
  const obs = String(p.observacao || '').replace(/\s+/g, ' ').trim().slice(0, OBS_EQUIPE_MAX);
  if (!data || !parque) return { ok: false, motivo: 'DADOS_INVALIDOS' };
  if (status && !STATUS_EQUIPE[status]) return { ok: false, motivo: 'DADOS_INVALIDOS' };
  if (status && STATUS_EQUIPE[status].obs && !obs) return { ok: false, motivo: 'FALTA_OBSERVACAO' };

  // cliente e supervisor saem do projeto (não do navegador)
  const pr = lerProjetos_().filter(function (x) { return chaveTexto_(x.parque) === chaveTexto_(parque); })[0];
  if (!pr) return { ok: false, motivo: 'PROJETO_INEXISTENTE' };

  const trava = LockService.getScriptLock();
  if (!trava.tryLock(20000)) return { ok: false, motivo: 'OCUPADO' };
  try {
    const aba = abaEquipesStatus_(true);
    let linha = -1;
    if (aba.getLastRow() >= 2) {
      const vals = aba.getRange(2, 1, aba.getLastRow() - 1, 3).getValues();
      for (let i = 0; i < vals.length; i++) {
        if (dataISO_(vals[i][2]) === data && chaveTexto_(vals[i][1]) === chaveTexto_(parque)) { linha = i + 2; break; }
      }
    }
    if (!status) {
      if (linha > -1) aba.deleteRow(linha);
    } else {
      const dia = Utilities.parseDate(data, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const nova = [pr.cliente || '', pr.parque || parque, dia,
                    (pr.supervisor && pr.supervisor.nome) || '', STATUS_EQUIPE[status].rotulo, obs,
                    String(matricula), new Date()];
      if (linha > -1) aba.getRange(linha, 1, 1, nova.length).setValues([nova]);
      else aba.appendRow(nova);
    }
    registrar_(matricula, 'EQUIPE_STATUS', 'OK', data + ' | ' + parque + ' | ' + (status || 'APAGADO'));
    return {
      ok: true,
      marcacao: status ? { status: status, observacao: obs, por: '', em: '' } : null
    };
  } finally {
    trava.releaseLock();
  }
}
