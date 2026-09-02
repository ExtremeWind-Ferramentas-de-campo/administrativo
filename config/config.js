/* ============================================================================
   CONFIGURAÇÃO DO PORTAL
   Mexa aqui só quando ligar o backend ou mudar os campos do cadastro.
   Para acrescentar uma opção no menu, use config/modulos.js.
   ============================================================================ */

/* MODO_DEMO = false -> chama o Web App do Apps Script em API_URL. É o modo de produção.
   MODO_DEMO = true  -> usaria os usuários de teste de config/demo.js, que não está mais
                        nesta pasta. Se voltar para true sem esse arquivo, o login avisa
                        e trava o botão em vez de quebrar sem explicação. */
const MODO_DEMO = false;

/* Cole aqui a URL /exec da implantação do Apps Script (apps-script/Codigo.gs). */
const API_URL = 'https://script.google.com/macros/s/AKfycbyIT6qWi8CXQog8eQ6EHg65_CDVc1JPevbCGUj9wziNuhHDhyWXmOXvk-9shPrNwuxe/exec';


/* ----------------------------------------------------------------------------
   CAMPOS DO PERFIL
   'edita' define quem altera pela tela:
     'dono'  -> a própria pessoa
     'admin' -> apenas quem tem perfil ADMIN
     'nunca' -> ninguém pela tela (só direto na planilha)

   preenchivel: true  -> campo de admin que a PRÓPRIA pessoa pode preencher uma
                         vez, enquanto estiver em branco. Depois de preenchido,
                         trava e só o administrativo altera. É o que permite
                         cadastrar alguém só com matrícula, CPF e email.

   O backend confere isso de novo. Se mudar aqui, ajuste também as listas
   EDITAVEL_DONO / EDITAVEL_ADMIN no Codigo.gs — senão a alteração é ignorada.
   ---------------------------------------------------------------------------- */
const CAMPOS_PERFIL = [
  { grupo:'Identificação' },
  { chave:'nome',       rotulo:'Nome completo',       edita:'admin', obrigatorio:true, preenchivel:true },
  { chave:'matricula',  rotulo:'Matrícula',           edita:'nunca', obrigatorio:true, mono:true },
  { chave:'cpf',        rotulo:'CPF',                 edita:'nunca', obrigatorio:true, mono:true },
  { chave:'nascimento', rotulo:'Data de nascimento',  edita:'dono',  obrigatorio:true, tipo:'date' },

  { grupo:'Contato' },
  { chave:'email',      rotulo:'Email corporativo',   edita:'dono',  obrigatorio:true, tipo:'email' },
  { chave:'telefone',   rotulo:'Telefone',            edita:'dono',  obrigatorio:true, mono:true, placeholder:'(00) 00000-0000' },
  { chave:'emergencia', rotulo:'Contato de emergência', edita:'dono', obrigatorio:true, placeholder:'Nome e telefone' },

  { grupo:'Vínculo' },
  { chave:'setor',      rotulo:'Setor',               edita:'admin', obrigatorio:true, preenchivel:true },
  { chave:'cargo',      rotulo:'Cargo',               edita:'admin', obrigatorio:false, preenchivel:true },
  { chave:'admissao',   rotulo:'Data de admissão',    edita:'admin', obrigatorio:false, tipo:'date' }
];
