/* ─────────────────────────────────────────────────────────────
   SERVICE WORKER — Portal Administrativo (Extreme Wind)
   Existe para o Chrome oferecer "Instalar app". O portal depende do
   Apps Script para tudo (login, projetos, status), então offline ele
   só abre a casca da tela — os dados exigem internet.

   O arquivo anterior com este nome era uma cópia do sw.js do site de
   campo (RDO, Fotocard...). Nunca chegou a ser registrado aqui, e se
   fosse, falhava na instalação: ele pré-carregava rdo/index.html,
   guard.js etc., que não existem nesta pasta.

   Estratégia:
   • HTML, JS e CSS do portal → rede primeiro (deploy novo aparece na
     hora; config.js com API_URL velha em cache seria pior que offline)
   • imagens e fontes → cache primeiro
   • Apps Script e POST → sempre direto à rede, nunca guardados

   CACHE com prefixo próprio: o site de campo fica no mesmo domínio
   (github.io) e o Cache Storage é por domínio, não por pasta. Este SW
   só apaga caches que começam com "ew-portal-".
   Ao publicar mudanças, incremente o número abaixo.
   ───────────────────────────────────────────────────────────── */
const PREFIXO = 'ew-portal-';
const CACHE = PREFIXO + 'v1';

/* Melhor esforço, arquivo por arquivo: um 404 aqui não pode impedir a
   instalação do app inteiro (addAll é tudo-ou-nada). */
const ARQUIVOS = [
  './',
  'index.html',
  'manifest.json',
  'logo-ew.png',
  'icon-192.png',
  'icon-512.png',
  'icon-maskable-512.png',
  'assets/base.css',
  'assets/base.js',
  'config/config.js',
  'config/modulos.js',
  'modulos/solicitacao-materiais.html',
  'modulos/supervisao-campo.html',
  'modulos/projetos-andamento.html',
  'modulos/status-rdo.html',
  'modulos/status-ehs.html',
  'modulos/gantt-campo.html'
];

const NUNCA_CACHE = [
  'https://script.google.com',
  'https://script.googleusercontent.com'
];

self.addEventListener('install', e => {
  e.waitUntil((async () => {
    const c = await caches.open(CACHE);
    await Promise.allSettled(ARQUIVOS.map(u => c.add(u)));
    self.skipWaiting();
  })());
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    const chaves = await caches.keys();
    await Promise.all(chaves
      .filter(k => k.indexOf(PREFIXO) === 0 && k !== CACHE)
      .map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

async function guardar(req, resp){
  if (resp && (resp.ok || resp.type === 'opaque')) {
    try { (await caches.open(CACHE)).put(req, resp.clone()); } catch (_) {}
  }
  return resp;
}

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  if (NUNCA_CACHE.some(o => req.url.indexOf(o) === 0)) return;

  const url = new URL(req.url);
  const mesmoSite = url.origin === self.location.origin;
  const codigo = req.mode === 'navigate' || /\.(html|js|css|json)$/i.test(url.pathname) ||
                 (req.headers.get('accept') || '').includes('text/html');

  if (mesmoSite && codigo) {
    // rede primeiro; sem rede, a cópia guardada
    e.respondWith((async () => {
      try {
        return await guardar(req, await fetch(req));
      } catch (_) {
        return (await caches.match(req, { ignoreSearch: true })) ||
               (req.mode === 'navigate' ? await caches.match('index.html') : null) ||
               Response.error();
      }
    })());
    return;
  }

  // imagens, fontes, Font Awesome do CDN → cache primeiro
  e.respondWith((async () => {
    const hit = await caches.match(req);
    if (hit) return hit;
    try { return await guardar(req, await fetch(req)); }
    catch (_) { return Response.error(); }
  })());
});
