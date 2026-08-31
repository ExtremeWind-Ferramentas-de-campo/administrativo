# Portal Administrativo — Extreme Wind

## O que tem na pasta

```
portal/
├── index.html                          tela de login + menu
├── logo-ew.png / icon-192.png          marca
├── config/
│   ├── config.js                       MODO_DEMO, API_URL e campos do perfil
│   └── modulos.js                      ← as opções do menu ficam aqui
├── modulos/
│   ├── solicitacao-materiais.html      quadro do almoxarifado
│   └── _MODELO.html                    modelo em branco para módulos novos
└── apps-script/
    ├── Codigo.gs                       login, senha, perfil
    └── Materiais.gs                    dados do quadro de materiais
```

---

## Cadastrar quem vai usar o portal

Recarregue a planilha depois de colar o `Codigo.gs`: aparece um menu **Portal**
na barra de cima.

- **Portal > Cadastrar pessoa** — pede tudo num campo só, separado por ponto e
  vírgula. Ao terminar, mostra a senha do primeiro acesso.
- **Portal > Resetar senha de alguém** — devolve a senha para o CPF.
- **Portal > Ver usuários cadastrados** — lista quem existe e quem ainda não
  trocou a senha.

**A senha do primeiro acesso é sempre o CPF, só os números**, sem ponto e sem
traço. No primeiro login o sistema obriga a criar uma senha pessoal.

As senhas do modo demonstração do site (`Torre2026`, por exemplo) só existiam no
navegador. No servidor elas não valem.

---

## Visual

O portal usa a mesma linguagem dos apps de campo: cena de fundo com os
aerogeradores girando, painéis de vidro e botão de tema claro/escuro no canto
superior direito.

O tema é gravado em `localStorage` na chave `ew_theme` — a **mesma** dos outros
apps. Quem escolhe escuro num deles abre o outro já no escuro. Sem escolha
salva, segue o tema do aparelho.

Os ícones vêm do Font Awesome 6 pelo CDN. Sem internet, os cards continuam
funcionando; só os pictogramas somem.

---

## Acrescentar uma opção no menu

1. Coloque o `.html` dentro de `modulos/`.
   Se for começar do zero, copie `modulos/_MODELO.html` — ele já vem com a
   barra e as cores do portal.
2. Abra `config/modulos.js` e copie o bloco de exemplo que está comentado no fim.
3. Ajuste `cod`, `nome`, `desc`, `url`, `perfis`, `icone` e `cor`. Salve.

Para o ícone, procure em fontawesome.com/icons e copie a classe
(ex.: `fas fa-boxes-stacked`). A `cor` pinta a borda de cima do card, o ícone e
o "Abrir".

Não precisa mexer no `index.html`. Se o card não aparecer, é quase sempre
vírgula faltando ou vírgula sobrando entre os blocos `{ }`.

---

## Ligar o backend

Passo a passo completo está comentado no topo de `apps-script/Codigo.gs`.
Resumo:

1. Planilha nova → Extensões > Apps Script → cola o `Codigo.gs`
2. Roda `configurarPlanilha()`
3. Roda `definirPepper()` — **uma vez só**. Se apagar essa propriedade depois,
   nenhuma senha funciona mais.
4. Cola `Materiais.gs` como um segundo script e roda `configurarMateriais()`
5. Cadastra as pessoas: recarregue a planilha e use o menu **Portal >
   Cadastrar pessoa**
6. Implantar > App da Web → *Executar como: eu* / *Acesso: qualquer pessoa* →
   copia a URL `/exec`
7. Em `config/config.js`: `MODO_DEMO = false` e cola a URL em `API_URL`

Enquanto `MODO_DEMO` estiver `true`, isto é um protótipo. Não use com dados reais.

---

## Testar antes de publicar

Abrir o `index.html` com dois cliques (`file://`) faz o navegador bloquear os
arquivos de `config/`. Rode um servidor local, dentro da pasta `portal`:

```
python -m http.server 8000
```

E abra `http://localhost:8000`.

Usuários de teste (modo demo):

| Matrícula | Senha         | O que acontece                          |
|-----------|---------------|-----------------------------------------|
| 10432     | 52398471023   | primeiro acesso, cadastro incompleto    |
| 10087     | Torre2026     | entra como ADMIN, cadastro completo     |

---

## Publicar no GitHub Pages

**Antes de subir, faça estes três:**

1. `config/config.js` → `MODO_DEMO = false` e `API_URL` com a URL `/exec` real
2. **`config/demo.js` já não está nesta pasta** — ele tinha matrícula, CPF e senha
   de exemplo. Para testar sem backend de novo, recrie o arquivo e descomente a
   tag `<script>` dele no `index.html`
3. Confira que o `.nojekyll` foi junto (o Git não mostra arquivos que começam com
   ponto por padrão)

Depois: Settings > Pages, aponte para a branch e a pasta. Os caminhos são todos
relativos, então funciona tanto na raiz quanto em `/portal/`.

`apps-script/` pode subir ou não — não tem senha nem chave dentro (o PEPPER fica
nas Propriedades do Script, não no código). Publicar só expõe a lógica do backend
para quem quiser estudá-la.

### O repositório vai ser público?

O GitHub Pages gratuito só serve site de repositório público. Isso significa que
qualquer pessoa lê o `index.html` e descobre a sua `API_URL`. Não é uma falha:
quem valida senha é o Apps Script, e ele bloqueia a matrícula por 15 minutos
depois de 5 tentativas erradas. Mas é bom você saber que o endereço do endpoint
é público, e que quem quiser tentar entrar precisa de matrícula **e** senha.

Se preferir que nem o endereço apareça, o caminho é repositório privado com
Pages — o que exige plano pago do GitHub.

---

## Quadro de Solicitação de Materiais

O quadro não tem login próprio: usa a sessão aberta no portal. Se alguém abrir
`modulos/solicitacao-materiais.html` direto, aparece uma tela pedindo para entrar
pelo portal.

Como os dados ficam:

- **Cards** → aba `MAT_CARDS`, uma linha por solicitação
- **Configuração do quadro** → aba `MAT_META` (nextSeq, colunas, revisão global)
- **Anexos** → pasta no seu Drive, fora da planilha. O card guarda só a ficha do
  arquivo. Uma célula do Sheets aceita 50 mil caracteres; um anexo de 3 MB em
  base64 tem 4 milhões.

O quadro consulta o servidor a cada 8 segundos. Se ninguém mexeu, a resposta é
só o número da revisão — não lê a planilha inteira à toa.

**Edição simultânea:** cada card tem sua própria revisão. Duas pessoas em cards
diferentes nunca se atrapalham. Na mesma solicitação, quem salva primeiro passa;
o segundo recebe um aviso, vê a versão que ficou e refaz a alteração. Nada é
sobrescrito em silêncio.

Enquanto alguém está com uma solicitação aberta, atualizações que chegam ficam
esperando e entram quando a janela fecha — para o formulário não ser trocado
debaixo de quem está digitando.

### Instalação do módulo

No mesmo projeto do Apps Script, crie um segundo script (Arquivo > + > Script,
nome `Materiais`) e cole `apps-script/Materiais.gs`. Depois rode
`configurarMateriais()` uma vez. Ela cria as duas abas e a pasta de anexos.

Reimplante o App da Web depois de colar (Implantar > Gerenciar implantações >
editar > Nova versão), senão as ações novas não sobem.

### Anexos e privacidade

Os arquivos ficam **privados** no Drive, sem link público. Para baixar, o quadro
chama o Apps Script com o token da sessão; ele confere o token, confere que o
arquivo está mesmo na pasta de anexos e devolve o conteúdo. Quem não estiver
logado no portal não abre anexo nem com o id na mão.

Limite de 4 MB por arquivo — não é da planilha, é o tempo de execução do Apps
Script, que começa a estourar acima disso.

Rode `limparAnexosOrfaos()` de vez em quando: ela manda para a lixeira arquivos
que nenhum card referencia mais.
