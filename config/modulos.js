/* ============================================================================
   MÓDULOS DO MENU
   Este é o único arquivo que você precisa editar para acrescentar uma opção.

   COMO ADICIONAR UMA OPÇÃO NOVA
   1. Coloque o arquivo .html dentro da pasta  modulos/
      (use modulos/_MODELO.html como ponto de partida, se quiser)
   2. Copie um bloco { ... } abaixo, cole antes do  ];  e ajuste os campos.
   3. Salve. Não precisa mexer no index.html.

   CAMPOS
     cod        3 letras que aparecem no canto do card. Serve como etiqueta curta.
     nome       Título do card.
     desc       Uma linha explicando o que a pessoa faz ali.
     url        Caminho do arquivo. Relativo à pasta do portal:
                  'modulos/nome-do-arquivo.html'
                Também aceita endereço completo (https://...).
     perfis     Quem enxerga o card. Perfis: ADMIN, SUPERVISOR, DIRETORIA.
     novaAba    false = abre na mesma aba (padrão daqui; o módulo tem botão de
                        voltar para o portal)
                true  = abre em outra aba
     icone      Ícone do Font Awesome 6 (grátis). Escolha em fontawesome.com/icons
                e copie a classe, ex.: 'fas fa-boxes-stacked'. Se omitir, vem um cubo.
     cor        Cor do card em hexadecimal. Pinta a borda de cima, o ícone e o "Abrir".
                Se omitir, fica azul.
   ============================================================================ */

const MODULOS = [

  {
    cod:      'SOL',
    nome:     'Solicitação de Materiais',
    desc:     'Solicitar, separar, enviar e confirmar entrega',
    url:      'modulos/solicitacao-materiais.html',
    perfis:   ['ADMIN','SUPERVISOR'],
    novaAba:  false,
    icone:    'fas fa-boxes-stacked',
    cor:      '#c26a12'
  },

  {
    cod:      'SUP',
    nome:     'Supervisão de Campo',
    desc:     'Status dos RDOs e projetos em andamento',
    url:      'modulos/supervisao-campo.html',
    perfis:   ['ADMIN','SUPERVISOR','DIRETORIA'],
    novaAba:  false,
    icone:    'fas fa-helmet-safety',
    cor:      '#0e7c7b'
  }

  /* ---------------------------------------------------------------------------
     MODELO — descomente e ajuste para criar a próxima opção.
     Atenção à vírgula: todo bloco precisa de uma vírgula depois do  }  ,
     menos o último da lista.

  ,{
    cod:      'XXX',
    nome:     'Nome do módulo',
    desc:     'O que a pessoa resolve aqui, em uma linha',
    url:      'modulos/arquivo.html',
    perfis:   ['ADMIN','SUPERVISOR'],
    novaAba:  true,
    icone:    'fas fa-list-check',
    cor:      '#0e7c7b'
  }
  --------------------------------------------------------------------------- */

];
