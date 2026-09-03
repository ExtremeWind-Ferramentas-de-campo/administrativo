# Portal Administrativo — Extreme Wind

## O que tem na pasta

```
portal/
├── index.html                          tela de login + menu
├── assets/base.css, base.js            estilo e funções comuns dos módulos
├── logo-ew.png / icon-192.png          marca
├── config/
│   ├── config.js                       MODO_DEMO, API_URL e campos do perfil
│   └── modulos.js                      ← as opções do menu ficam aqui
├── modulos/
│   ├── solicitacao-materiais.html      quadro do almoxarifado
│   └── _MODELO.html                    modelo em branco para módulos novos
└── apps-script/
    ├── Codigo.gs                       login, senha, perfil
    ├── Materiais.gs                    dados do quadro de materiais
    └── SupervisaoCampo.gs              projetos e leitura do RDO
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

O cadastro pede só **matrícula, CPF e email**. Nome, setor e cargo a própria
pessoa preenche no primeiro acesso — o portal abre "Meus dados" sozinho e não
deixa passar sem completar. Depois de preenchidos, esses campos travam: só o
setor administrativo altera. Assim ninguém troca o próprio setor depois.

**A senha do primeiro acesso é sempre o CPF, só os números**, sem ponto e sem
traço. No primeiro login o sistema obriga a criar uma senha pessoal.

As senhas do modo demonstração do site (`Torre2026`, por exemplo) só existiam no
navegador. No servidor elas não valem.

---

## Supervisão de Campo

Duas telas dentro do módulo:

**Projetos em Andamento** — parque, cliente, tipo de reparo, início, situação e
a lista de técnicos. Cria e edita: **SUPERVISOR**. ADMIN e DIRETORIA apenas
visualizam — e a recusa é feita no Apps Script, não só escondendo o botão.

**Status RDO** — só leitura da planilha do RDO. Mostra, para a data escolhida,
quais projetos em andamento já têm relatório e quais faltam. Filtra por
`Data_exp`, cliente e parque. Clicar no card abre o `Link_PDF`.

Sábado e domingo aparecem como **"não obrigatório"**: o relatório pode existir,
mas a falta não é cobrada.

### O casamento é pelo nome do parque

A cobrança compara o parque do projeto com a coluna `Parque` do RDO, usando a
célula inteira.

**O parque é digitado pelo supervisor, em texto livre.** O projeto nasce antes
do primeiro RDO, então não dá para escolher de uma lista tirada do RDO: no dia
do cadastro aquele parque ainda não existe lá.

O preço disso é que a grafia passa a depender de quem digita. "SANTO AGOSTINHO
1", "Santo Agostinho 01" e "STO AGOSTINHO 1" são três parques diferentes para o
Status RDO — o projeto vai constar como **sem relatório para sempre**. Escreva
exatamente como aparece na coluna `Parque` do RDO. Para ver a grafia real, use
**Portal > Conferir colunas do RDO**.

**Parque e equipe andam juntos, de propósito.** O número no fim
("SANTO AGOSTINHO 1") é a equipe, mas o casamento usa a célula inteira como um
rótulo só. Separar exigia adivinhar onde termina o nome do parque, e o resultado
saía inconsistente: "OITIS 1" ficava inteiro enquanto "SÃO FERNANDO 1" era
dividido. Mantendo junto, os dois lados usam exatamente o mesmo texto e cada
dupla parque+equipe vira uma linha própria na cobrança.

Por isso **não existe mais campo "número da equipe"** no cadastro: ele já estava
dentro do nome do parque, e ter os dois só criava chance de divergir.

### Tipo de reparo e data de finalização

O **tipo de reparo** vem da planilha **Banco de inputs**, aba `ATV POR HR`. Só
entram na lista as linhas em que a coluna `Atividade obrigatória` está como
`Não` — atividade obrigatória é rotina de todo dia, não é o reparo que define o
projeto. O campo filtra enquanto se digita, mas continua aceitando texto livre:
se a planilha estiver fora do ar, o cadastro não trava.

As colunas são achadas pelo cabeçalho, ignorando maiúsculas e acentos. Se o
cabeçalho mudar, acrescente o nome novo em `COLUNAS_INPUTS`, no
`SupervisaoCampo.gs`.

A situação **Concluído** exige a **data de finalização**, e a data não pode ser
anterior ao início. Enquanto o projeto está em andamento o campo fica escondido
e a data é apagada — assim não sobra data de fim em projeto que ainda roda.

### Endereço das planilhas

Os IDs do RDO e da MINI MASTER ficam nas **Propriedades do Script**, não no
código. O código vai para o GitHub; as propriedades não.

**No editor do Apps Script:** ícone de engrenagem (Configurações do projeto,
na barra da esquerda) > Propriedades do script > Adicionar propriedade.

| Propriedade | Valor |
|---|---|
| `ID_RDO` | link ou ID da planilha do banco de dados do RDO |
| `ABA_RDO` | nome da aba dos relatórios (ex.: `Relatorios`) |
| `ID_MINIMASTER` | link ou ID da planilha MINI MASTER |
| `ABA_MINIMASTER` | nome da aba dos técnicos (deixe vazio para a primeira aba) |
| `ID_INPUTS` | link ou ID da planilha **Banco de inputs** |
| `ABA_INPUTS` | nome da aba dos tipos de reparo (padrão: `ATV POR HR`) |

Pode colar o link inteiro do navegador — o ID é extraído sozinho.

Quem prefere não mexer nas propriedades pode usar **Portal > Configurar
planilhas**, no menu da planilha, que faz o mesmo por perguntas.

Depois, rode **`verConfiguracao()`** pelo botão Executar: mostra o que está
gravado, se consegue ler as três planilhas, quantos técnicos e quantos tipos de
reparo achou, e como parque e equipe estão sendo separados. O resultado sai em Registro de execução.

### "Cannot call SpreadsheetApp.getUi() from this context"

Esse erro aparece quando a função é executada com o editor aberto fora da
planilha. As funções de instalação não dependem mais de tela: quando não há
interface, a mensagem sai em **Registro de execução** e a função termina normal.

Se o menu **Portal** não aparecer na planilha, abra a planilha pelo Google
Sheets e recarregue (F5) — o menu só é criado no momento em que ela abre.

### Card do relatório

Mostra: parque (a célula `Parque` inteira, com o número), cliente, `Data_exp`,
`Turbina`, `Blade`, `Tipo_reparo` e `Avanco_reparo` com barra. "Finalizado"
aparece quando `Reparo_finalizado` começa com SIM.

A borda e a barra seguem a **cor do cliente**. SIEMENS, GE, NORDEX, VESTAS,
WOBBEN/ENERCON, GOLDWIND e WEG têm cor fixa em `CORES_CLIENTE`, no topo do
script de `status-rdo.html`. Cliente novo recebe uma cor estável da paleta
sozinho — o mesmo nome cai sempre na mesma cor. Para fixar, acrescente a linha
lá.

### Colunas da planilha do RDO

Descobertas pelo cabeçalho, comparando com `COLUNAS_RDO` no topo de
`SupervisaoCampo.gs`. Use **Portal > Conferir colunas do RDO** para ver o que
foi reconhecido, o cabeçalho real e o que falta.

A coluna `Parque` traz nome e equipe juntos ("SANTO AGOSTINHO 1"), e é usada
inteira como chave.

Colunas usadas hoje: `Data_exp`, `Parque`, `Cliente`, `Link_PDF`,
`Tipo_reparo`, `Turbina`, `Blade`, `Avanco_reparo`, `Reparo_finalizado` e
`Matricula_login` — esta última vira o nome de quem enviou, cruzando com a
MINI MASTER.

### Técnicos

Os nomes vêm da planilha **MINI MASTER**: coluna A é a matrícula, coluna B é o
nome. No projeto, o supervisor digita o nome, a lista filtra e ele escolhe — a
matrícula vem junto e não é digitada. Assim não nasce técnico com matrícula
trocada. Nome sem escolha na lista não salva.

A lista fica em cache por 10 minutos: técnico novo na MINI MASTER aparece nesse
prazo.

### Sessão

Dura **12 horas** e se renova a cada uso. Fica gravada na aba `SESSOES`, com o
token em hash — quem abrir a planilha não consegue se passar por ninguém.

Antes ficava só no `CacheService`, e era isso que derrubava a sessão pouco depois
de entrar: o Google descarta entradas de cache quando quer, e **toda nova
implantação do App da Web limpa tudo**. O cache continua na frente como atalho,
mas quem manda é a aba.

"Sair" apaga a linha da aba, então o token deixa de valer na hora.

Rode `limparExpirados()` de vez em quando para remover as sessões vencidas.

### Perfis

- **ADMIN** — setor administrativo. Cadastra pessoas e usa os módulos. Em
  Projetos em Andamento, **só visualiza**.
- **SUPERVISOR** — cria e edita os projetos em andamento.
- **DIRETORIA** — acompanha tudo, sem alterar projetos.

Quem enxerga cada módulo é definido por `perfis` em `config/modulos.js`.

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

O quadro abre **na mesma aba** do portal e tem um botão "Portal" no canto
superior direito para voltar — que devolve ao menu, não ao login: o portal
retoma a sessão guardada ao recarregar. A sessão vale 8 horas; passado isso, ou
se o token cair antes, a tela de login volta com o aviso. Não tem login próprio: usa a sessão aberta no
portal. Se alguém abrir
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

### Fotos e PDFs

- **Fotos aparecem no card**, sem precisar abrir. Se houver mais de uma, setas
  passam de uma para a outra ali mesmo. Clicar na foto abre em tela cheia.
- **PDF abre clicando no nome do arquivo**, dentro do próprio quadro, sem baixar.
  No celular, se a página vier em branco, use "Abrir em nova aba" na barra de cima.
- Dentro da solicitação, as fotos aparecem em grade e cada arquivo tem o botão
  de baixar do lado.

O que o card mostra é uma **miniatura** de 340 px gerada no navegador na hora do
envio, guardada como arquivo separado no Drive. O original só é baixado quando
alguém abre ou baixa de fato. Sem isso, um quadro com trinta fotos de 4 MB
levaria mais de 100 MB a cada abertura.

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

---

## Manutenção da planilha

Duas funções, nenhuma urgente. O ideal é criar um acionador mensal para cada
(no editor do Apps Script: ícone do relógio > Adicionar acionador > Timer mensal).

- **`limparExpirados()`** — no `Codigo.gs`. Apaga códigos de recuperação com mais
  de 7 dias e faz a faxina da aba `LOG`.
- **`limparAnexosOrfaos()`** — no `Materiais.gs`. Manda para a lixeira do Drive
  arquivos que nenhum card usa mais.

### Prazos do LOG

São dois, definidos no topo do `Codigo.gs`:

- `DIAS_LOG_SEGURANCA = 180` — login, troca de senha, reset, cadastro, erro.
  Volume baixíssimo (umas 40 linhas por dia numa equipe de 20). É o que você
  consulta se precisar entender um acesso indevido, e esse tipo de coisa
  aparece semanas depois.
- `DIAS_LOG_OPERACAO = 15` — anexos enviados e baixados. Volume alto, utilidade
  curta.

Salvar no quadro **não** gera linha de log: a própria aba `MAT_CARDS` já guarda
`atualizado_por` e `atualizado_em` de cada solicitação. Só conflito e exclusão
são registrados.
