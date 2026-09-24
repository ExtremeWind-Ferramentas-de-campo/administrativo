/* ─────────────────────────────────────────────────────────────
   SERVICE WORKER — Extreme Wind | Ferramentas de Campo (PWA)
   Um único service worker para o site inteiro: menu, RDO, Fotocard,
   Calculadoras e Checklist.

   Antes eram três (raiz, fotocard/, rdo/). Como todos apagavam
   "todo cache que não é o meu", eles derrubavam o cache um do outro
   e o app do técnico voltava a precisar de internet sem motivo.
   Com um só, o problema deixa de existir.

   Estratégia:
   • HTML  → network-first (online = versão nova; offline = cache)
   • estáticos (png/svg/css/fontes) → cache-first
   • POST e chamadas de API → sempre à rede

   Requer hospedagem em http(s) (ex.: GitHub Pages). Não funciona via file://.
   IMPORTANTE: ao atualizar arquivos do app, incremente o número do CACHE
   abaixo — é o que descarta o cache antigo e força a atualização.
   ───────────────────────────────────────────────────────────── */
const CACHE = 'ew-site-v41';

const CORE = [
  './',
  'index.html',                              // menu + tela de entrada
  'guard.js',                                // porteiro de sessão
  'prazos.js',                               // sinal de prazo semanal nos cartões
  'manifest.json',
  'logo-ew.png',
  'logo-oem.png',
  'logo-oem-ge.png',
  'logo-oem-siemens.png',
  'icon-192.png',
  'icon-512.png',
  'rdo/index.html',                          // Relatório Diário de Operação
  'fotocard/index.html',
  'fotocard/manifest.json',
  'fotocard/icon-192.png',
  'fotocard/icon-512.png',
  'calculadora/index.html',                  // seletor de cliente
  'calculadora/calculadora-nordex.html',
  'calculadora/calculadora-ge.html',
  'calculadora/calculadora-siemens.html',
  'Checklist Almoxarifado/menu.html',                   // submenu: Almoxarifado/Segurança e Frotas
  'Checklist Almoxarifado/almoxarifado-seguranca.html', // materiais + 5 checklists de inspeção
  'Checklist Almoxarifado/index.html',
  'meus-dados/index.html',                   // hub pessoal (fica antes das ferramentas)
  'meus-dados/pe-de-meia.html',              // a tela abre offline; a consulta em si exige rede
  'meus-dados/cursos.html',
  'meus-dados/dividas.html',
  'meus-dados/equipamentos.html'
];
/* Frotas fica fora do CORE de propósito: addAll é tudo-ou-nada, e um 404 aqui
   (pasta ainda não publicada, nome com espaço mal servido) derrubaria o cache
   do site inteiro. Aqui é melhor esforço — se faltar, só essas páginas ficam
   sem offline. */
const FROTAS = [
  'Checklist Frotas/index.html',             // seletor dos 3 checklists de frota
  'Checklist Frotas/gerador-eletrico.html',
  'Checklist Frotas/plataforma.html',
  'Checklist Frotas/veiculo.html'
];
const EXTRA = [
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/webfonts/fa-solid-900.woff2',
  // jsPDF — os checklists de Frotas montam o PDF no próprio aparelho.
  // Sem esta linha, a primeira abertura offline não gera nada.
  'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js'
];

self.addEventListener('install', e => {
  e.waitUntil((async () => {
    const c = await caches.open(CACHE);
    await c.addAll(CORE);                                   // essenciais (locais) — obrigatório
    await Promise.allSettled(FROTAS.map(u => c.add(u)));    // Frotas — melhor esforço
    await Promise.allSettled(EXTRA.map(u => c.add(u)));     // CDN — melhor esforço
    self.skipWaiting();
  })());
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    /* Limpa os caches das versões antigas, inclusive os dos service workers
       que existiam separados (ew-calc-*, ew-fotocard-*, rdo-*). */
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

/* Nada de API entra no cache: cache-first devolveria lista de parques velha,
   e resposta de API não deve ficar guardada em disco no aparelho.
   O Apps Script também precisa FALHAR de verdade quando está offline, para o
   RDO usar a fila e as listas já salvas no aparelho em vez de dado antigo. */
const SEM_CACHE = [
  'https://ew-dropbox-proxy.ew-fotos.workers.dev',
  'https://script.google.com',
  'https://script.googleusercontent.com'
];

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;                        // POST → direto à rede
  if (SEM_CACHE.some(o => req.url.indexOf(o) === 0)) return;

  const isHTML = req.mode === 'navigate' ||
                 (req.headers.get('accept') || '').includes('text/html');

  if (isHTML) {
    // network-first: online pega a versão nova; offline cai no cache
    e.respondWith((async () => {
      try {
        const net = await fetch(req);
        const c = await caches.open(CACHE);
        c.put(req, net.clone());
        return net;
      } catch (_) {
        const guardada = await caches.match(req);
        if (guardada) return guardada;
        /* Sem cópia da página pedida: devolve o index da MESMA pasta, para
           quem pediu /rdo/ não receber a calculadora. */
        const pasta = req.url.slice(0, req.url.lastIndexOf('/') + 1);
        return (await caches.match(pasta + 'index.html')) ||
               (await caches.match('index.html')) ||
               Response.error();
      }
    })());
    return;
  }

  // estáticos → cache-first (funciona offline; guarda fontes/CSS no 1º uso)
  e.respondWith((async () => {
    const hit = await caches.match(req);
    if (hit) return hit;
    try {
      const net = await fetch(req);
      if (net && (net.ok || net.type === 'opaque')) {
        const c = await caches.open(CACHE);
        c.put(req, net.clone());
      }
      return net;
    } catch (_) {
      return hit || Response.error();
    }
  })());
});

// Recebe mensagem para forçar atualização (o Fotocard usa isso no banner)
self.addEventListener('message', e => {
  if (e.data === 'SKIP_WAITING') self.skipWaiting();
});
