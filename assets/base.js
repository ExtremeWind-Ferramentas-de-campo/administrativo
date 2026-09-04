/* ============================================================================
   BASE DOS MÓDULOS
   Sessão, chamada ao Apps Script, tema e utilidades usadas pelas páginas
   dentro de modulos/. Carregue depois de config/config.js.

   Nenhum módulo tem login próprio: todos usam a sessão aberta no portal,
   guardada em localStorage['ew_sessao'].
   ============================================================================ */

var SESSAO = null;

function lerSessao(){
  try{
    var s = JSON.parse(localStorage.getItem('ew_sessao') || 'null');
    if(!s || !s.token || !s.matricula) return null;
    if(s.expira && Date.now() > s.expira) return null;
    return s;
  }catch(e){ return null; }
}

function encerrarSessao(){
  try{ localStorage.removeItem('ew_sessao'); }catch(e){}
}

function perfilAtual(){
  return SESSAO ? String(SESSAO.perfil || '').toUpperCase() : '';
}

/* Todo perfil enxerga todos os módulos. O que o perfil separa é quem ALTERA:
   projeto é criado e editado por ADMIN e SUPERVISOR; DIRETORIA e USUARIO veem.
   Quem decide de verdade é o Apps Script; isto aqui só monta a tela. */
var PERFIS_EDITAM_PROJETO = ['ADMIN','SUPERVISOR'];
function podeEditarProjetos(){
  return PERFIS_EDITAM_PROJETO.indexOf(perfilAtual()) > -1;
}

/* ---------------------------------------------------------------------------
   COR POR CLIENTE
   Fica aqui, e não em cada módulo, para Status RDO e Projetos em Andamento
   pintarem o mesmo cliente da mesma cor. Nome conhecido tem cor fixa; nome
   novo cai sempre na mesma cor da paleta, sem precisar cadastrar nada.
   --------------------------------------------------------------------------- */
var CORES_CLIENTE = {
  'siemens':  '#00b0a0',
  'ge':       '#6a5acd',
  'nordex':   '#e0702a',
  'vestas':   '#2b7fd0',
  'wobben':   '#c0392b',
  'enercon':  '#c0392b',
  'goldwind': '#1f9d55',
  'weg':      '#0f6fb5'
};
var PALETA_CLIENTE = ['#2b7fd0','#00b0a0','#e0702a','#6a5acd','#c0392b','#1f9d55',
                      '#b5179e','#0f6fb5','#c26a12','#3f7d20'];

function corDoCliente(nome){
  var chave = String(nome || '').trim().toLowerCase();
  if(!chave) return 'var(--faint)';
  if(CORES_CLIENTE[chave]) return CORES_CLIENTE[chave];
  for(var k in CORES_CLIENTE){
    if(chave.indexOf(k) > -1) return CORES_CLIENTE[k];   // "SIEMENS GAMESA"
  }
  var soma = 0;
  for(var i = 0; i < chave.length; i++) soma = (soma * 31 + chave.charCodeAt(i)) % 100000;
  return PALETA_CLIENTE[soma % PALETA_CLIENTE.length];
}

/* ---------- chamada ao backend ---------- */
function chamar(acao, dados){
  if(typeof MODO_DEMO !== 'undefined' && MODO_DEMO){
    return Promise.reject({ motivo:'DEMO' });
  }
  if(typeof API_URL === 'undefined'){
    return Promise.reject({ motivo:'SEM_CONFIG' });
  }
  var corpo = { acao: acao, token: SESSAO ? SESSAO.token : null };
  for(var k in dados){ if(Object.prototype.hasOwnProperty.call(dados,k)) corpo[k] = dados[k]; }

  return fetch(API_URL, {
    method:'POST',
    headers:{ 'Content-Type':'text/plain;charset=utf-8' },  // evita preflight CORS
    redirect:'follow',
    body: JSON.stringify(corpo)
  })
  .then(function(r){
    if(!r.ok) throw { motivo:'REDE' };
    return r.json();
  })
  .then(function(res){
    if(res && res.ok === false && res.motivo === 'SESSAO'){
      encerrarSessao();
      bloquear('Sua sessão expirou. Entre de novo pelo portal para continuar.');
      throw { motivo:'SESSAO' };
    }
    return res;
  });
}

var MENSAGENS_MODULO = {
  SESSAO:          'Sua sessão expirou. Entre de novo pelo portal.',
  SEM_PERMISSAO:   'Seu perfil não permite esta alteração.',
  USUARIOS_INDISPONIVEL:'Não consegui ler a lista de supervisores.',
  INPUTS_INDISPONIVEL:'Não consegui ler a planilha Banco de inputs.',
  MINIMASTER_INDISPONIVEL:'Não consegui ler a planilha MINI MASTER.',
  RDO_INDISPONIVEL:'Não consegui ler a planilha do RDO.',
  OCUPADO:         'Outra pessoa está salvando agora. Tente de novo em instantes.',
  DEMO:            'Modo demonstração: esta tela precisa do servidor configurado.',
  SEM_CONFIG:      'API_URL não está configurada em config/config.js.',
  REDE:            'Sem conexão com o servidor.'
};
function traduzirErro(err){
  var m = err && err.motivo;
  return MENSAGENS_MODULO[m] || 'Não foi possível concluir. Tente novamente.';
}

/* ---------- bloqueio de tela ---------- */
function bloquear(mensagem, comBotao){
  var el = document.getElementById('bloqueio');
  if(!el) return;
  document.getElementById('bloqueioMsg').textContent = mensagem;
  document.getElementById('bloqueioAcao').hidden = (comBotao === false);
  el.hidden = false;
}
function desbloquear(){
  var el = document.getElementById('bloqueio');
  if(el) el.hidden = true;
}

/* ---------- recado rápido ---------- */
var recadoTimer = null;
function recado(msg, tipo){
  var el = document.getElementById('recado');
  if(!el) return;
  el.textContent = msg;
  el.className = 'recado' + (tipo === 'erro' ? ' erro' : '');
  el.hidden = false;
  if(recadoTimer) clearTimeout(recadoTimer);
  recadoTimer = setTimeout(function(){ el.hidden = true; }, 4200);
}

/* ---------- tema ---------- */
(function temaInicial(){
  try{
    var t = localStorage.getItem('ew_theme');
    if(t === 'dark' || t === 'light') document.documentElement.setAttribute('data-theme', t);
  }catch(e){}
})();

/* ---------- utilidades ---------- */
function esc(t){
  return String(t == null ? '' : t)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function uid(pre){
  return (pre || 'id') + '_' + Date.now().toString(36) + Math.random().toString(36).slice(2,8);
}

function hojeISO(){
  var d = new Date();
  return d.getFullYear() + '-' + ('0'+(d.getMonth()+1)).slice(-2) + '-' + ('0'+d.getDate()).slice(-2);
}

function dataBonita(iso){
  var p = String(iso || '').split('-');
  if(p.length !== 3) return iso || '';
  return p[2] + '/' + p[1] + '/' + p[0];
}

var DIAS = ['domingo','segunda-feira','terça-feira','quarta-feira','quinta-feira','sexta-feira','sábado'];
function diaDaSemana(iso){
  var p = String(iso || '').split('-');
  if(p.length !== 3) return '';
  return DIAS[new Date(+p[0], +p[1]-1, +p[2]).getDay()];
}

function soDigitos(t){ return String(t == null ? '' : t).replace(/\D/g,''); }

/* Arranca a página: confere a sessão e chama iniciar() da página. */
function iniciarModulo(){
  var botao = document.getElementById('bloqueioAcao');
  if(botao) botao.addEventListener('click', function(){ location.href = '../index.html'; });

  SESSAO = lerSessao();
  if(!SESSAO){
    bloquear('Esta página usa o login do portal. Entre pelo portal e abra este módulo pelo menu.');
    return false;
  }
  var quem = document.getElementById('quemSou');
  if(quem) quem.textContent = (SESSAO.nome || 'Matrícula ' + SESSAO.matricula) + ' · ' + perfilAtual();
  return true;
}
