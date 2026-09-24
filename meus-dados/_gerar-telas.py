# -*- coding: utf-8 -*-
"""Gera meus-dados/cursos.html e meus-dados/dividas.html a partir de um
template só. As duas telas são idênticas fora do rótulo, do ícone, da cor e
da mensagem de "não tem nada" — manter dois arquivos escritos à mão faria as
duas divergirem na primeira correção."""
import io, os

TEMPLATE = u'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<script src="../guard.js"></script>
<meta http-equiv="Content-Security-Policy" content="default-src 'self'; script-src 'self' 'unsafe-inline' https://script.google.com https://script.googleusercontent.com; style-src 'self' 'unsafe-inline' https://cdnjs.cloudflare.com; font-src 'self' https://cdnjs.cloudflare.com; img-src 'self' data: blob:; connect-src 'self' https://script.google.com https://script.googleusercontent.com; object-src 'none'; base-uri 'self'">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="theme-color" content="#2b7fd0">
<title>__TITULO__ | Extreme Wind</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" integrity="sha384-iw3OoTErCYJJB9mCa8LNS2hbsQ7M3C0EpIsO/H5+EGAkPGc6rk+V8i04oW/K5xq0" crossorigin="anonymous" referrerpolicy="no-referrer">
<link rel="manifest" href="../manifest.json">
<link rel="apple-touch-icon" href="../icon-192.png">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<style>
  *{margin:0;padding:0;box-sizing:border-box}
  :root{
    --ink:#0e2033; --muted:#4f6273;
    --panel:rgba(255,255,255,.82); --panel-border:rgba(255,255,255,.65);
    --card:#ffffff; --card-line:#e2e8f0;
    --bg-1:#6db3ef; --bg-2:#eaf6ff;
    --cor:__COR__; --cor-forte:__COR_FORTE__;
    --ff:'Segoe UI',system-ui,-apple-system,Roboto,Helvetica,Arial,sans-serif;
    --shadow:0 18px 50px rgba(9,34,64,.18);
  }
  :root[data-theme="dark"]{
    --ink:#e8eff8; --muted:#9fb2c6;
    --panel:rgba(15,29,48,.74); --panel-border:rgba(120,150,185,.22);
    --card:#0f1d2e; --card-line:#2a3c52;
    --bg-1:#081525; --bg-2:#173251;
    --cor:__COR_DARK__; --cor-forte:__COR_DARK__;
    --shadow:0 18px 50px rgba(0,0,0,.5);
  }
  @media (prefers-color-scheme:dark){
    :root:not([data-theme="light"]){
      --ink:#e8eff8; --muted:#9fb2c6;
      --panel:rgba(15,29,48,.74); --panel-border:rgba(120,150,185,.22);
      --card:#0f1d2e; --card-line:#2a3c52;
      --bg-1:#081525; --bg-2:#173251;
      --cor:__COR_DARK__; --cor-forte:__COR_DARK__;
      --shadow:0 18px 50px rgba(0,0,0,.5);
    }
  }
  body{font-family:var(--ff);color:var(--ink);min-height:100vh;
    background:linear-gradient(180deg,var(--bg-1),var(--bg-2));
    display:flex;flex-direction:column;-webkit-tap-highlight-color:transparent}
  .topo{display:flex;align-items:center;gap:12px;padding:16px 16px 0;max-width:640px;width:100%;margin:0 auto}
  .nav-back-link{width:40px;height:40px;flex-shrink:0;border-radius:12px;display:grid;place-items:center;
    border:1px solid var(--panel-border);background:var(--panel);color:var(--ink);text-decoration:none;
    box-shadow:var(--shadow)}
  .topo h1{flex:1;text-align:center;font-size:clamp(1.3rem,4vw,1.8rem);font-weight:800;letter-spacing:-.5px}
  .btn-olho{width:40px;height:40px;flex-shrink:0;border-radius:12px;cursor:pointer;
    border:1px solid var(--panel-border);background:var(--panel);color:var(--ink);
    box-shadow:var(--shadow);font-size:.95rem}
  main{flex:1;width:100%;max-width:640px;margin:0 auto;padding:20px 16px 24px;
    display:flex;flex-direction:column;gap:16px}
  .sessao{text-align:center;font-size:.82rem;color:var(--muted);font-weight:600}

  .painel{background:var(--panel);border:1px solid var(--panel-border);border-radius:24px;
    padding:20px 16px 22px;box-shadow:var(--shadow);backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px)}

  .estado{text-align:center;padding:26px 12px;color:var(--muted);font-weight:600;font-size:.95rem;line-height:1.5}
  .estado i{display:block;font-size:2rem;margin-bottom:12px;color:var(--cor)}
  .estado.erro i{color:#b02a1e}
  .estado.vazio i{color:var(--muted)}

  .itens{display:flex;flex-direction:column;gap:14px}
  .item{background:var(--card);border:1px solid var(--card-line);border-top:5px solid var(--cor);
    border-radius:18px;padding:14px 16px 12px}
  .item h3{font-size:1rem;font-weight:800;line-height:1.3;margin-bottom:10px;
    display:flex;align-items:flex-start;gap:8px}
  .item h3 i{color:var(--cor);font-size:.9rem;margin-top:3px;flex-shrink:0}
  .item .num{margin-left:auto;font-size:.72rem;font-weight:700;color:var(--muted);
    background:var(--panel-border);border-radius:8px;padding:2px 8px;flex-shrink:0}
  .linha{display:flex;align-items:center;justify-content:space-between;gap:12px;
    padding:8px 0;border-top:1px solid var(--card-line)}
  .linha .rot{font-size:.84rem;font-weight:700;color:var(--muted)}
  .linha .val{font-size:1rem;font-weight:800;white-space:nowrap;font-variant-numeric:tabular-nums}
  .linha.destaque .rot{color:var(--cor-forte);font-size:.9rem}
  .linha.destaque .val{font-size:1.25rem;color:var(--cor-forte)}
  .linha.quitado .val{color:#15803d}

  .item.total{border-top-color:var(--cor);border:2px solid var(--cor);
    background:color-mix(in srgb,var(--cor) 8%,var(--card))}
  .item.total h3{color:var(--cor-forte)}
  body.oculto .val{filter:blur(9px);user-select:none}

  .atualizado{margin-top:16px;text-align:center;font-size:.82rem;color:var(--muted);
    font-weight:600;line-height:1.6;border-top:1px dashed var(--card-line);padding-top:14px}
  .atualizado strong{display:block;font-size:1rem;color:var(--ink);font-weight:800;margin-top:2px}

  .btn-recarregar{width:100%;min-height:48px;border:none;border-radius:14px;cursor:pointer;
    font-family:inherit;font-size:.95rem;font-weight:800;color:#fff;background:var(--cor);
    box-shadow:0 6px 18px rgba(9,34,64,.22);display:flex;align-items:center;justify-content:center;gap:8px}
  .btn-recarregar:active{filter:brightness(.94)}
  .btn-recarregar:disabled{opacity:.6;cursor:default}

  .nota{font-size:.75rem;color:var(--muted);text-align:center;line-height:1.5}
  footer{text-align:center;color:var(--muted);font-size:.78rem;padding:8px 16px 18px}
  .girando{animation:gira 1s linear infinite}
  @keyframes gira{to{transform:rotate(360deg)}}
  @media(prefers-reduced-motion:reduce){*{animation-duration:.01ms !important;transition-duration:.01ms !important}}
</style>
</head>
<body>

<div class="topo">
  <a class="nav-back-link" href="index.html" title="Voltar"><i class="fas fa-arrow-left"></i></a>
  <h1>__TITULO__</h1>
  <button type="button" class="btn-olho" id="btnOlho" title="Mostrar/ocultar valores" aria-label="Mostrar ou ocultar valores">
    <i class="fas fa-eye"></i>
  </button>
</div>

<main>
  <div class="sessao" id="sessaoNome"></div>

  <section class="painel">
    <div id="conteudo">
      <div class="estado"><i class="fas fa-circle-notch girando"></i>Consultando…</div>
    </div>
    <div class="atualizado" id="blocoData" hidden>
      <span id="rotuloData">Valor atualizado em:</span>
      <strong id="dataAtualizacao">—</strong>
    </div>
  </section>

  <button type="button" class="btn-recarregar" id="btnRecarregar">
    <i class="fas fa-rotate"></i> Atualizar
  </button>

  <p class="nota">Dúvida no valor? Fale com o RH — esta tela é só consulta.</p>
</main>

<footer>Extreme Wind Blade Services &copy; <span id="year"></span> · Uso interno</footer>

<script>
'use strict';
document.getElementById('year').textContent = new Date().getFullYear();

(function(){
  var t = localStorage.getItem('ew_theme');
  if(t === 'dark' || t === 'light') document.documentElement.setAttribute('data-theme', t);
})();

(function(){

/* Mesmo Apps Script do login e do RDO. Se trocar o endereco la, troque aqui. */
const ENDPOINT = "__ENDPOINT__";
const SESSAO_LS_KEY = 'ew_sessao';
const TIPO = '__TIPO__';                 /* casa com CP_CONSULTAS no Code.gs */
const ICONE_ITEM = '__ICONE__';
const VAZIO = '__VAZIO__';

const elConteudo = document.getElementById('conteudo');
const elBlocoData = document.getElementById('blocoData');
const elRotulo = document.getElementById('rotuloData');
const elData = document.getElementById('dataAtualizacao');
const btnRecarregar = document.getElementById('btnRecarregar');

function sessao(){
  try{
    const s = JSON.parse(localStorage.getItem(SESSAO_LS_KEY) || 'null');
    if(!s || !s.token || !s.exp || s.exp <= Date.now()) return null;
    return s;
  }catch(_){ return null; }
}

const s = sessao();
if(s) document.getElementById('sessaoNome').textContent = s.nome + ' \\u00b7 mat. ' + s.mat;

/* olho: esconde os valores quando tem gente olhando por cima */
document.getElementById('btnOlho').addEventListener('click', function(){
  const oculto = document.body.classList.toggle('oculto');
  this.innerHTML = oculto ? '<i class="fas fa-eye-slash"></i>' : '<i class="fas fa-eye"></i>';
});

/* plano B por JSONP, igual ao login: alguns aparelhos nao conseguem ler a
   resposta do POST pro Apps Script */
function pedirJsonp(token){
  return new Promise(resolve=>{
    const cb = '__cpCb' + Date.now();
    let pronto = false;
    window[cb] = j => { pronto = true; resolve(j); try{ delete window[cb]; }catch(_){ window[cb]=null; } };
    const sc = document.createElement('script');
    sc.src = ENDPOINT + (ENDPOINT.indexOf('?')>=0?'&':'?')
           + 'acao=consultaPessoal&tipo=' + encodeURIComponent(TIPO)
           + '&token=' + encodeURIComponent(token) + '&callback=' + cb;
    sc.onerror = ()=>{ if(!pronto) resolve(null); };
    document.head.appendChild(sc);
    setTimeout(()=>{ if(!pronto) resolve(null); }, 20000);
  });
}

function esc(t){
  return String(t==null?'':t).replace(/[&<>"']/g, c =>
    ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function mostrarEstado(icone, texto, classe){
  elConteudo.innerHTML = '<div class="estado ' + (classe||'') + '">'
    + '<i class="fas ' + icone + '"></i>' + texto + '</div>';
}

function mostrarData(r){
  if(r && (r.rotuloData || r.dataAtualizacao)){
    elRotulo.textContent = r.rotuloData || 'Valor atualizado em:';
    elData.textContent = r.dataAtualizacao || '\\u2014';
    elBlocoData.hidden = false;
  }
}

/* uma linha de valor. `destaque` sai grande; saldo zerado sai em verde,
   que e a diferenca que o tecnico procura primeiro. */
function linhaHtml(v, destaque){
  const ehDestaque = (v.chave === destaque);
  const quitado = ehDestaque && v.valor === 0;
  return '<div class="linha' + (ehDestaque?' destaque':'') + (quitado?' quitado':'') + '">'
       + '<span class="rot">' + esc(v.rotulo) + '</span>'
       + '<span class="val">' + esc(v.texto || 'R$ 0,00') + (quitado?' \\u2713':'') + '</span>'
       + '</div>';
}

function mostrarItens(r){
  const itens = r.itens || [];
  const varios = itens.length > 1;
  let html = itens.map(function(it, i){
    return '<div class="item">'
         + '<h3><i class="fas ' + ICONE_ITEM + '"></i>' + esc(it.titulo)
         + (varios ? '<span class="num">' + (i+1) + '/' + itens.length + '</span>' : '')
         + '</h3>'
         + (it.valores||[]).map(v => linhaHtml(v, r.destaque)).join('')
         + '</div>';
  }).join('');

  if(r.totais && r.totais.length){
    html += '<div class="item total">'
          + '<h3><i class="fas fa-equals"></i>Total (' + itens.length + ' itens)</h3>'
          + r.totais.map(v => linhaHtml(v, r.destaque)).join('')
          + '</div>';
  }
  elConteudo.innerHTML = '<div class="itens">' + html + '</div>';
}

function paraOLogin(){
  try{ localStorage.removeItem(SESSAO_LS_KEY); }catch(_){}
  location.replace('../index.html');
}

async function consultar(){
  const ss = sessao();
  if(!ss){ paraOLogin(); return; }

  btnRecarregar.disabled = true;
  btnRecarregar.innerHTML = '<i class="fas fa-circle-notch girando"></i> Consultando\\u2026';
  mostrarEstado('fa-circle-notch girando', 'Consultando\\u2026');
  elBlocoData.hidden = true;

  let r = null;
  try{
    const res = await fetch(ENDPOINT, {
      method:'POST',
      body: JSON.stringify({acao:'consultaPessoal', tipo: TIPO, token: ss.token})
    });
    r = await res.json();
  }catch(_){ r = null; }

  if(!r) r = await pedirJsonp(ss.token);

  btnRecarregar.disabled = false;
  btnRecarregar.innerHTML = '<i class="fas fa-rotate"></i> Atualizar';

  if(!r){
    mostrarEstado('fa-wifi',
      'Sem conex\\u00e3o com o servidor.<br>Esta consulta precisa de internet \\u2014 '
      + 'conecte e toque em <b>Atualizar</b>.', 'erro');
    return;
  }

  if(r.sessao === false){
    mostrarEstado('fa-right-to-bracket', esc(r.erro || 'Sess\\u00e3o expirada.')
      + '<br>Voltando para a tela de entrada\\u2026', 'erro');
    setTimeout(paraOLogin, 2200);
    return;
  }

  if(!r.ok){
    mostrarEstado('fa-triangle-exclamation', esc(r.erro || 'N\\u00e3o foi poss\\u00edvel consultar agora.'), 'erro');
    return;
  }

  mostrarData(r);

  if(!r.encontrado){
    mostrarEstado('__ICONE__', VAZIO, 'vazio');
    return;
  }

  mostrarItens(r);
}

btnRecarregar.addEventListener('click', consultar);
consultar();

})();
</script>
</body>
</html>
'''

ENDPOINT = ("https://script.google.com/macros/s/AKfycbxoOpV339g76UpYa7sO28C6lS99TAz7po2c0dNAk0i1X1HgsyKC_"
            "KXIuuBAS7qbBzLG/exec")

TELAS = [
    {
        'arquivo': 'cursos.html',
        'tipo': 'cursos',
        'titulo': u'Cursos',
        'icone': 'fa-graduation-cap',
        'cor': '#6a4bb8', 'cor_forte': '#4e3392', 'cor_dark': '#9b82e0',
        'vazio': u'N\\u00e3o h\\u00e1 desconto relacionado ao pagamento de cursos.',
    },
    {
        'arquivo': 'dividas.html',
        'tipo': 'dividas',
        'titulo': u'D\u00edvidas',
        'icone': 'fa-file-invoice-dollar',
        'cor': '#c26a12', 'cor_forte': '#95500c', 'cor_dark': '#e3963f',
        'vazio': u'N\\u00e3o h\\u00e1 d\\u00edvidas registradas na sua conta.',
    },
]

DESTINO = os.path.dirname(os.path.abspath(__file__))

for t in TELAS:
    html = (TEMPLATE
            .replace('__ENDPOINT__', ENDPOINT)
            .replace('__TIPO__', t['tipo'])
            .replace('__TITULO__', t['titulo'])
            .replace('__ICONE__', t['icone'])
            .replace('__COR_FORTE__', t['cor_forte'])
            .replace('__COR_DARK__', t['cor_dark'])
            .replace('__COR__', t['cor'])
            .replace('__VAZIO__', t['vazio']))
    assert '__' not in html.replace('__cpCb', ''), 'sobrou marcador em ' + t['arquivo']
    caminho = os.path.join(DESTINO, t['arquivo'])
    io.open(caminho, 'w', encoding='utf-8').write(html)
    print('gerado:', caminho, len(html), 'bytes')
