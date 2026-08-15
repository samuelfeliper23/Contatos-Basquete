// ============================================================
//  SISTEMAS WEB I — CRUD DE CONTATOS COM NODE.JS + EJS
//  Turma: 2AI | MTEC Informática para Web
//  Sem banco de dados: os dados ficam na memória do servidor
// ============================================================

// 1. IMPORTAÇÕES
// "require" carrega módulos (pacotes) instalados pelo npm
const express = require('express');  // framework web para Node.js
const path    = require('path');     // módulo nativo do Node para lidar com caminhos de pasta

// 2. CRIAÇÃO DO APP EXPRESS
const app = express(); // cria a aplicação web

// 3. CONFIGURAÇÃO DO EJS COMO MOTOR DE VIEWS
// Diz ao Express: "quando eu mandar renderizar uma view, use EJS"
app.set('view engine', 'ejs');
// Diz onde ficam os arquivos .ejs (pasta "views" dentro do projeto)
app.set('views', path.join(__dirname, 'views'));

// 4. MIDDLEWARES
// Middleware = código que roda ANTES das rotas, processando a requisição

// Permite ler dados enviados por formulários HTML (application/x-www-form-urlencoded)
app.use(express.urlencoded({ extended: false }));

// Serve arquivos estáticos (CSS, imagens) da pasta "public"
app.use(express.static(path.join(__dirname, 'public')));

// ============================================================
// 5. "BANCO DE DADOS" NA MEMÓRIA
// Um simples array de objetos que funciona como nosso banco de dados
// ATENÇÃO: os dados são perdidos quando o servidor reinicia!
// ============================================================
let contatos = [
  { id: 1, nome: 'LeBron James',        data_nasc: '30/12/18984',       email: 'papailebron@email.com', modalidade: 'Basquete' },
  { id: 2, nome: 'Bronny James',    data_nasc: '06/10/2004', email: 'bronnyjr@email.com', modalidade: 'Basquete' },
  { id: 3, nome: 'Luka Doncic', data_nasc: '28/02/1999',  email: 'naoeholukamodric@email.com',   modalidade: 'Basquete' }
];

// Contador para gerar IDs únicos automaticamente
let proximoId = 4;

// ============================================================
// 6. ROTAS CRUD
// CRUD = Create (criar), Read (ler), Update (atualizar), Delete (deletar)
// ============================================================

// ----------------------------------------------------------
// READ — LISTAR TODOS OS CONTATOS
// Rota: GET /
// Quando o usuário acessa http://localhost:3000/
// ----------------------------------------------------------
app.get('/', (req, res) => {
  // res.render('index') → procura o arquivo views/index.ejs
  // O segundo argumento é um objeto com variáveis disponíveis no EJS
  res.render('index', {
    titulo: 'Lista de Contatos',  // variável <%= titulo %> no EJS
    contatos: contatos           // variável <%= contatos %> (array) no EJS
  });
});

// ----------------------------------------------------------
// CREATE (parte 1) — EXIBIR FORMULÁRIO PARA NOVO CONTATO
// Rota: GET /contatos/novo
// ----------------------------------------------------------
app.get('/contatos/novo', (req, res) => {
  res.render('novo', {
    titulo: 'Novo Jogador'
  });
});

// ----------------------------------------------------------
// CREATE (parte 2) — RECEBER E SALVAR O NOVO CONTATO
// Rota: POST /contatos
// O formulário HTML envia os dados para cá com method="POST"
// ----------------------------------------------------------
app.post('/contatos', (req, res) => {
  // req.body contém os campos do formulário
  // Os nomes vêm do atributo "name" dos inputs no HTML
  const { nome, data_nasc, email, modalidade } = req.body;

  // Cria um novo objeto contato
  const novoContato = {
    id: proximoId,  // ID único gerado pelo contador
    nome: nome,
    email: email,
    data_nasc: data_nasc,
    modalidade: modalidade
  };

  // Incrementa o contador para o próximo contato
  proximoId++;

  // Adiciona o novo contato ao array (nosso "banco de dados")
  contatos.push(novoContato);

  // Redireciona o usuário de volta para a lista
  res.redirect('/');
});

// ----------------------------------------------------------
// UPDATE (parte 1) — EXIBIR FORMULÁRIO DE EDIÇÃO
// Rota: GET /contatos/editar/:id
// :id é um parâmetro dinâmico (ex: /contatos/editar/2)
// ----------------------------------------------------------
app.get('/contatos/editar/:id', (req, res) => {
  // req.params.id pega o valor do :id na URL
  // Number(...) converte a string "2" para o número 2
  const id = Number(req.params.id);

  // Procura no array o contato com o ID correspondente
  const contato = contatos.find(c => c.id === id);

  // Se não encontrou, retorna erro 404
  if (!contato) {
    return res.status(404).send('Contato não encontrado.');
  }

  // Renderiza o formulário de edição, passando o contato encontrado
  res.render('editar', {
    titulo: 'Editar Contato',
    contato: contato   // variável <%= contato %> no EJS
  });
});

// ----------------------------------------------------------
// UPDATE (parte 2) — RECEBER E SALVAR AS ALTERAÇÕES
// Rota: POST /contatos/editar/:id
// ----------------------------------------------------------
app.post('/contatos/editar/:id', (req, res) => {
  const id = Number(req.params.id);

  // Procura o índice do contato no array
  const indice = contatos.findIndex(c => c.id === id);

  if (indice === -1) {
    return res.status(404).send('Contato não encontrado.');
  }

  // Atualiza apenas os campos enviados pelo formulário
  // Mantém o mesmo ID
  contatos[indice] = {
    id: id,
    nome:     req.body.nome,
    email:    req.body.email,
    data_nasc: req.body.data_nasc,
    modalidade: req.body.modalidade
  };

  res.redirect('/');
});

// ----------------------------------------------------------
// DELETE — DELETAR UM CONTATO
// Rota: POST /contatos/deletar/:id
// Usamos POST pois formulários HTML só suportam GET e POST
// ----------------------------------------------------------
app.post('/contatos/deletar/:id', (req, res) => {
  const id = Number(req.params.id);

  // filter() cria um NOVO array sem o contato com o ID informado
  // É como dizer: "mantenha todos EXCETO o que tem esse id"
  contatos = contatos.filter(c => c.id !== id);

  res.redirect('/');
});

// ============================================================
// 7. INICIAR O SERVIDOR
// ============================================================
const PORTA = 3000;

app.listen(PORTA, () => {
  // Este callback roda quando o servidor estiver pronto
  console.log(`Servidor rodando em http://localhost:${PORTA}`);
  console.log('Pressione CTRL+C para parar.');
});
