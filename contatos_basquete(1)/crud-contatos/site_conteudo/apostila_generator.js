'use strict';
// =============================================================
//  Gerador da Apostila — CRUD Node.js + Express + EJS
//  ETEC Campo Limpo Paulista | Sistemas Web I | Prof. Marcelo
// =============================================================
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, TableOfContents, LevelFormat
} = require('docx');
const fs = require('fs');

// ─── CONSTANTES ──────────────────────────────────────────────
const F     = 'Arial';
const FM    = 'Courier New';
const CW    = 9026;        // largura do conteúdo em DXA (A4, margem 1")
const BLUE  = '1B3A6B';   // azul escuro
const LBLUE = '2471A3';   // azul médio
const CODEBG = 'F4F4F4';  // fundo cinza claro (código)
const INFOBG = 'EBF5FB';  // fundo azul claro (nota)
const TIPBG  = 'EAFAF1';  // fundo verde claro (dica)
const WARNBG = 'FEF9E7';  // fundo amarelo (atenção)

// ─── HELPERS ─────────────────────────────────────────────────
const bdr  = (c = 'D0D0D0', s = 4) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const lBdr = (c = LBLUE,    s = 16) => ({ style: BorderStyle.SINGLE, size: s, color: c });

function pb()        { return new Paragraph({ children: [new PageBreak()] }); }
function el(h = 100) { return new Paragraph({ children: [new TextRun('')], spacing: { before: 0, after: h } }); }
function bold(t)     { return new TextRun({ text: t, font: F, size: 24, bold: true }); }
function ital(t)     { return new TextRun({ text: t, font: F, size: 24, italic: true }); }
function plain(t)    { return new TextRun({ text: t, font: F, size: 24 }); }
function mono(t)     { return new TextRun({ text: t, font: FM, size: 22, color: 'C0392B' }); }

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 480, after: 240 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLUE, space: 2 } },
    children: [new TextRun({ text, font: F, size: 36, bold: true, color: BLUE })]
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 320, after: 160 },
    children: [new TextRun({ text, font: F, size: 28, bold: true, color: LBLUE })]
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, font: F, size: 24, bold: true, color: '444444' })]
  });
}

function np(runs_or_text, opts = {}) {
  const children = typeof runs_or_text === 'string'
    ? [new TextRun({ text: runs_or_text, font: F, size: 24, color: '1A1A1A', bold: opts.bold, italic: opts.italic })]
    : runs_or_text;
  return new Paragraph({
    children,
    alignment: opts.align !== undefined ? opts.align : AlignmentType.JUSTIFIED,
    spacing: { before: opts.before || 0, after: opts.after || 160 }
  });
}

function bp(text, level = 0) {
  const children = Array.isArray(text) ? text
    : [new TextRun({ text, font: F, size: 24, color: '1A1A1A' })];
  return new Paragraph({ children, numbering: { reference: 'bullets', level }, spacing: { before: 60, after: 80 } });
}

function np_num(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: F, size: 24 })],
    numbering: { reference: 'numbers', level: 0 },
    spacing: { before: 60, after: 80 }
  });
}

// ─── BLOCO DE CÓDIGO ──────────────────────────────────────────
function code(codeStr) {
  const lines = codeStr.split('\n');
  const runs = lines.map((line, i) =>
    new TextRun({ text: line, font: FM, size: 18, color: '1A1A1A', ...(i > 0 ? { break: 1 } : {}) })
  );
  return [
    el(80),
    new Table({
      width: { size: CW, type: WidthType.DXA },
      columnWidths: [CW],
      rows: [new TableRow({ children: [new TableCell({
        borders: { top: bdr('DDDDDD'), bottom: bdr('DDDDDD'), right: bdr('DDDDDD'), left: lBdr() },
        shading: { fill: CODEBG, type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 160, left: 280, right: 160 },
        width: { size: CW, type: WidthType.DXA },
        children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: runs })]
      })] })]
    }),
    el(120)
  ];
}

// ─── CAIXA COLORIDA ───────────────────────────────────────────
function box(icon, title, content, bg = INFOBG, bc = LBLUE) {
  const rows = [];
  if (title) rows.push(new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: `${icon}  ${title}`, font: F, size: 22, bold: true, color: BLUE })]
  }));
  const items = Array.isArray(content) ? content : [content];
  items.forEach((item, i) => {
    const children = typeof item === 'string'
      ? [new TextRun({ text: item, font: F, size: 22, color: '1A1A1A' })]
      : item;
    rows.push(new Paragraph({ spacing: { before: 0, after: i < items.length - 1 ? 80 : 0 }, children }));
  });
  return [
    el(80),
    new Table({
      width: { size: CW, type: WidthType.DXA }, columnWidths: [CW],
      rows: [new TableRow({ children: [new TableCell({
        borders: { top: bdr('AED6F1'), bottom: bdr('AED6F1'), right: bdr('AED6F1'), left: lBdr(bc) },
        shading: { fill: bg, type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 160, left: 220, right: 160 },
        width: { size: CW, type: WidthType.DXA },
        children: rows
      })] })]
    }),
    el(120)
  ];
}

// ─── TABELA GENÉRICA ──────────────────────────────────────────
function tbl(headers, rows, ratios) {
  const r = ratios || headers.map(() => 1 / headers.length);
  let cw = r.map(x => Math.floor(CW * x));
  cw[cw.length - 1] += CW - cw.reduce((a, b) => a + b, 0);

  const mkCell = (text, colI, isHeader = false, rowI = 0) => {
    const children = Array.isArray(text) ? text : [new TextRun({ text, font: F, size: 22, bold: isHeader, color: isHeader ? 'FFFFFF' : '1A1A1A' })];
    return new TableCell({
      borders: { top: bdr(), bottom: bdr(), left: bdr(), right: bdr() },
      shading: { fill: isHeader ? BLUE : (rowI % 2 === 0 ? 'FFFFFF' : 'F2F6FC'), type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 140, right: 100 },
      width: { size: cw[colI], type: WidthType.DXA },
      children: [new Paragraph({ spacing: { before: 0, after: 0 }, children })]
    });
  };

  const tableRows = [
    new TableRow({ tableHeader: true, children: headers.map((h, i) => mkCell(h, i, true)) }),
    ...rows.map((row, ri) => new TableRow({ children: row.map((c, i) => mkCell(c, i, false, ri)) }))
  ];
  return [el(80), new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: cw, rows: tableRows }), el(120)];
}

// =============================================================
//  CONTEÚDO DAS SEÇÕES
// =============================================================

function coverPage() {
  return [
    np('', { after: 1440 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: 'ETEC DE CAMPO LIMPO PAULISTA', font: F, size: 32, bold: true, color: BLUE })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 60 },
      children: [new TextRun({ text: '2º Ensino Médio Técnico — Informática para Internet', font: F, size: 24, color: '666666' })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 720 },
      children: [new TextRun({ text: 'Disciplina: Sistemas Web I', font: F, size: 24, italic: true, color: '666666' })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
      border: { top: { style: BorderStyle.SINGLE, size: 6, color: BLUE } },
      children: [new TextRun({ text: '' })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 100 },
      children: [new TextRun({ text: 'Apostila de Aula', font: F, size: 26, italic: true, color: '888888' })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: 'CRUD com Node.js + Express + EJS', font: F, size: 52, bold: true, color: BLUE })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE } },
      children: [new TextRun({ text: '' })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 1200 },
      children: [new TextRun({ text: 'Cadastro de Contatos sem Banco de Dados', font: F, size: 28, italic: true, color: '555555' })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: 'Professor: Marcelo Macrino dos Santos', font: F, size: 26, color: '333333' })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
      children: [new TextRun({ text: 'Campo Limpo Paulista — 2026', font: F, size: 24, color: '444444' })] }),
    pb()
  ];
}

function tocSection() {
  return [
    h1('Sumário'),
    new TableOfContents('Sumário', { hyperlink: true, headingStyleRange: '1-3' }),
    pb()
  ];
}

// ─── CAP. 1: CONCEITOS ────────────────────────────────────────
function chapter1() {
  return [
    h1('Capítulo 1 — Conceitos Fundamentais'),
    np('Antes de colocar a mão no código, vamos entender os quatro pilares desta aula: Node.js, Express, EJS e CRUD. Compreender esses conceitos vai facilitar muito o desenvolvimento do projeto.'),

    h2('1.1 O que é Node.js?'),
    np('O Node.js é um ambiente de execução JavaScript fora do navegador. Isso significa que você pode usar JavaScript não só para criar botões e animações em páginas HTML, mas também para programar servidores web, APIs e aplicações de linha de comando.'),
    np([bold('Por que usar Node.js?'), plain(' Ele é rápido, leve, usa a mesma linguagem no front-end e no back-end (JavaScript) e possui um ecossistema gigante de pacotes disponíveis pelo npm (Node Package Manager).')]),
    ...box('💡', 'Curiosidade', 'O Node.js foi criado em 2009 por Ryan Dahl e é usado por empresas como Netflix, LinkedIn e PayPal em sistemas de produção.', TIPBG, '27AE60'),

    h2('1.2 O que é Express?'),
    np('Express é um framework (conjunto de ferramentas) para Node.js que simplifica a criação de servidores web. Sem o Express, criar rotas e lidar com requisições HTTP exigiria muito código. Com ele, basta algumas linhas.'),
    np('Pense no Express como a “espinha dorsal” do nosso servidor: ele recebe as requisições do navegador (GET, POST) e define o que deve acontecer em cada endereço (rota).'),

    h2('1.3 O que é EJS?'),
    np('EJS (Embedded JavaScript) é um motor de templates. Ele permite criar arquivos HTML que contêm trechos de JavaScript, tornando as páginas dinâmicas. O servidor processa o arquivo .ejs, substitui as variáveis pelos valores reais, e envia o HTML final para o navegador.'),
    np([bold('Exemplo simples: '), plain('se o servidor envia a variável '), mono('nome = "Ana"'), plain(', o template '), mono('<h1>Olá, <%= nome %>!</h1>'), plain(' gera o HTML '), mono('<h1>Olá, Ana!</h1>'), plain('.')]),
    ...box('⚠️', 'Importante', 'Neste projeto, o HTML é gerado pelo servidor (renderização no servidor). Não usamos React, Angular nem Vue. O navegador recebe HTML puro, pronto para exibir.', WARNBG, 'E67E22'),

    h2('1.4 O que é CRUD?'),
    np('CRUD é o acronismo das quatro operações básicas que um sistema de dados precisa realizar:'),
    ...tbl(
      ['Letra', 'Operação', 'Descrição', 'Método HTTP'],
      [
        ['C', 'Create (Criar)',    'Adicionar um novo registro',             'POST'],
        ['R', 'Read (Ler)',        'Listar ou exibir registros existentes',  'GET'],
        ['U', 'Update (Atualizar)','Alterar um registro existente',          'POST'],
        ['D', 'Delete (Deletar)', 'Remover um registro existente',          'POST'],
      ],
      [0.08, 0.22, 0.45, 0.25]
    ),
    np('Neste projeto, o nosso CRUD gerencia uma lista de contatos (nome, e-mail e telefone) armazenada diretamente na memória do servidor — sem precisar de banco de dados.'),

    h2('1.5 Fluxo de uma Requisição Web'),
    np('Entenda o que acontece quando o aluno digita o endereço no navegador:'),
    np_num('O navegador envia uma requisição HTTP (GET ou POST) para o servidor.'),
    np_num('O Express recebe a requisição e verifica qual rota corresponde ao endereço.'),
    np_num('A rota executa a lógica (buscar, criar, editar ou deletar dados).'),
    np_num('O servidor chama res.render(), que processa o arquivo .ejs com os dados.'),
    np_num('O EJS substitui as variáveis, gera o HTML final e envia de volta ao navegador.'),
    np_num('O navegador exibe a página HTML recebida.'),
    pb()
  ];
}

// ─── CAP. 2: AMBIENTE ─────────────────────────────────────────
function chapter2() {
  return [
    h1('Capítulo 2 — Preparando o Ambiente'),
    np('Antes de começar a programar, precisamos verificar se o Node.js está instalado e criar a estrutura do projeto.'),

    h2('2.1 Verificando a Instalação do Node.js'),
    np('Abra o terminal (Prompt de Comando no Windows ou Terminal no Linux/Mac) e execute os comandos abaixo. Se aparecerem números de versão, tudo está instalado corretamente.'),
    ...code('node --version\nnpm --version'),
    np('Você deve ver algo como:'),
    ...code('v20.11.0\n10.2.4'),
    ...box('📌', 'Node.js não instalado?', 'Acesse https://nodejs.org e baixe a versão LTS (Long Term Support). O npm é instalado automaticamente junto com o Node.js.', INFOBG),

    h2('2.2 Criando a Pasta do Projeto'),
    np('No terminal, navegue até onde deseja criar o projeto e execute:'),
    ...code('mkdir crud-contatos\ncd crud-contatos'),
    np('Em seguida, inicialize o projeto com o npm. O flag '),
    np([mono('-y'), plain(' aceita todas as configurações padrão automaticamente:')], { after: 0 }),
    ...code('npm init -y'),
    np('Isso cria o arquivo '), np([mono('package.json'), plain(', que registra as informações e dependências do projeto.')], { after: 0 }),

    h2('2.3 Instalando as Dependências'),
    np('Precisamos instalar apenas dois pacotes: o Express (servidor) e o EJS (templates).'),
    ...code('npm install express ejs'),
    np('Após a instalação, o arquivo package.json será atualizado e uma pasta '),
    np([mono('node_modules'), plain(' será criada contendo todos os arquivos dos pacotes instalados.')], { after: 0 }),
    ...box('⚠️', 'Atenção', 'A pasta node_modules pode ter milhares de arquivos. Nunca a envie para o GitHub! Adicione-a ao arquivo .gitignore.', WARNBG, 'E67E22'),

    h2('2.4 Estrutura de Pastas do Projeto'),
    np('Crie manualmente as pastas e arquivos conforme a estrutura abaixo:'),
    ...code('crud-contatos/\n├── package.json          ← configurações e dependências\n├── server.js             ← servidor e todas as rotas CRUD\n├── views/                ← pasta dos templates EJS\n│   ├── index.ejs         ← página: lista de contatos\n│   ├── novo.ejs          ← página: formulário de criação\n│   └── editar.ejs        ← página: formulário de edição\n└── public/               ← arquivos estáticos (CSS, imagens)\n    └── style.css         ← folha de estilos'),

    h2('2.5 O arquivo package.json'),
    np('O package.json é o "cartão de visitas" do projeto. Ele informa o nome, versão, scripts e as dependências do projeto.'),
    ...code('{\n  "name": "crud-contatos",\n  "version": "1.0.0",\n  "description": "CRUD de Contatos com Node.js e EJS",\n  "main": "server.js",\n  "scripts": {\n    "start": "node server.js"\n  },\n  "dependencies": {\n    "express": "^4.18.2",\n    "ejs": "^3.1.9"\n  }\n}'),
    bp([bold('"scripts"'), plain(' → define atalhos de terminal. Com este arquivo, você pode rodar o servidor com '), mono('npm start'), plain(' em vez de '), mono('node server.js'), plain('.')]),
    bp([bold('"dependencies"'), plain(' → lista os pacotes necessários. O '), mono('^'), plain(' significa "esta versão ou superior".')]),
    pb()
  ];
}

// ─── CAP. 3: SERVER.JS ────────────────────────────────────────
function chapter3() {
  return [
    h1('Capítulo 3 — O Servidor: server.js'),
    np('O arquivo server.js é o coração do sistema. Ele inicializa o servidor, configura o EJS, define os dados na memória e implementa todas as rotas CRUD. Vamos analisá-lo parte por parte.'),

    h2('3.1 Importando os Módulos'),
    np('Todo arquivo Node.js que precisa de um módulo externo usa a função '), np([mono('require()'), plain('.')], { after: 0 }),
    ...code("// require() carrega os módulos instalados pelo npm\nconst express = require('express');  // framework web\nconst path    = require('path');     // módulo nativo: lida com caminhos de pasta"),
    bp([bold('express'), plain(': baixado via npm. Cria e gerencia o servidor web.')]),
    bp([bold('path'), plain(': módulo nativo do Node.js (já vem instalado). Ajuda a montar caminhos de arquivo de forma segura em qualquer sistema operacional.')]),

    h2('3.2 Configurando o Express e o EJS'),
    ...code("const app = express();  // cria a aplicação web\n\n// Define EJS como motor de views (templates)\napp.set('view engine', 'ejs');\n\n// Define a pasta onde os arquivos .ejs estão guardados\napp.set('views', path.join(__dirname, 'views'));\n\n// Permite ler dados enviados por formulários HTML\napp.use(express.urlencoded({ extended: false }));\n\n// Serve arquivos estáticos (CSS, imagens) da pasta 'public'\napp.use(express.static(path.join(__dirname, 'public')));"),
    bp([bold("app.set('view engine', 'ejs')"), plain(': informa ao Express qual motor usar para renderizar as views.')]),
    bp([bold('express.urlencoded()'), plain(': middleware que lê os dados enviados via formulário (tag '), mono('<form method="POST">'), plain(') e os disponibiliza em '), mono('req.body'), plain('.')]),
    bp([bold('express.static()'), plain(': faz com que os arquivos da pasta '), mono('public'), plain(' sejam acessíveis diretamente pelo navegador (ex: '), mono('/style.css'), plain(').')]),
    ...box('📚', 'O que é um Middleware?', 'Middleware é um código que executa ANTES das rotas, processando a requisição. O app.use() registra um middleware. Pense nele como um porteiro que prepara os dados antes de entregá-los à rota certa.', INFOBG),

    h2('3.3 Os Dados na Memória'),
    np('Como não usamos banco de dados, armazenamos os contatos em um simples array JavaScript. Os dados ficam enquanto o servidor está rodando e são perdidos ao reiniciá-lo.'),
    ...code("// Array de objetos — funciona como nosso 'banco de dados'\nlet contatos = [\n  { id: 1, nome: 'Ana Silva',    email: 'ana@email.com',    telefone: '(11) 91234-5678' },\n  { id: 2, nome: 'Bruno Costa',  email: 'bruno@email.com',  telefone: '(21) 98765-4321' },\n  { id: 3, nome: 'Carla Mendes', email: 'carla@email.com',  telefone: '(31) 99876-5432' },\n];\n\n// Contador para gerar IDs únicos\nlet proximoId = 4;"),
    bp([bold('let'), plain(' (e não '), bold('const'), plain('): usamos '), mono('let'), plain(' porque o array será reatribuído na operação de DELETE.')]),
    bp([bold('proximoId'), plain(': garante que cada novo contato receba um ID único e nunca repetido.')]),

    h2('3.4 READ — Listando os Contatos (GET /)'),
    np('A rota mais simples: quando o usuário acessa a raiz do site, enviamos a lista completa para o template index.ejs.'),
    ...code("app.get('/', (req, res) => {\n  // res.render('index') → procura views/index.ejs\n  // O 2º argumento envia variáveis para o template\n  res.render('index', {\n    titulo: 'Lista de Contatos',  // disponível como <%= titulo %>\n    contatos: contatos            // disponível como <%= contatos %>\n  });\n});"),
    bp([bold('app.get(rota, callback)'), plain(': define uma rota GET. O callback recebe '), mono('req'), plain(' (requisição) e '), mono('res'), plain(' (resposta).')]),
    bp([bold('res.render(view, dados)'), plain(': renderiza o arquivo .ejs substituindo as variáveis pelos valores enviados.')]),

    h2('3.5 CREATE — Criando Contatos'),
    np('A operação CREATE precisa de duas rotas: uma para exibir o formulário (GET) e outra para receber e salvar os dados (POST).'),
    ...code("// GET: exibe o formulário em branco\napp.get('/contatos/novo', (req, res) => {\n  res.render('novo', { titulo: 'Novo Contato' });\n});\n\n// POST: recebe os dados do formulário e salva\napp.post('/contatos', (req, res) => {\n  // req.body contém os campos do formulário HTML\n  const { nome, email, telefone } = req.body;\n\n  const novoContato = {\n    id: proximoId,\n    nome: nome,\n    email: email,\n    telefone: telefone\n  };\n\n  proximoId++;           // incrementa o contador de IDs\n  contatos.push(novoContato); // adiciona ao array\n\n  res.redirect('/');     // redireciona para a lista\n});"),
    bp([bold('req.body'), plain(': objeto que contém os campos do formulário. Os nomes vêm do atributo '), mono('name'), plain(' dos inputs HTML.')]),
    bp([bold('contatos.push()'), plain(': adiciona o novo objeto ao final do array.')]),
    bp([bold('res.redirect(\'/\')'), plain(': após salvar, redireciona o navegador para a página inicial (padrão PRG — Post/Redirect/Get).')]),

    h2('3.6 UPDATE — Editando Contatos'),
    np('Assim como o CREATE, o UPDATE usa duas rotas. O parâmetro '), np([mono(':id'), plain(' na rota é dinâmico — ele representa qualquer número.')], { after: 0 }),
    ...code("// GET: exibe o formulário preenchido com os dados atuais\napp.get('/contatos/editar/:id', (req, res) => {\n  const id = Number(req.params.id);  // pega o :id da URL\n  const contato = contatos.find(c => c.id === id);\n\n  if (!contato) {\n    return res.status(404).send('Contato não encontrado.');\n  }\n\n  res.render('editar', { titulo: 'Editar Contato', contato: contato });\n});\n\n// POST: recebe e salva as alterações\napp.post('/contatos/editar/:id', (req, res) => {\n  const id = Number(req.params.id);\n  const indice = contatos.findIndex(c => c.id === id);\n\n  if (indice === -1) return res.status(404).send('Não encontrado.');\n\n  // Substitui o objeto inteiro, mantendo o mesmo ID\n  contatos[indice] = {\n    id: id,\n    nome:     req.body.nome,\n    email:    req.body.email,\n    telefone: req.body.telefone\n  };\n\n  res.redirect('/');\n});"),
    bp([bold('req.params.id'), plain(': lê o valor dinâmico '), mono(':id'), plain(' da URL. Exemplo: em '), mono('/contatos/editar/2'), plain(', '), mono('req.params.id'), plain(' retorna '), mono('"2"'), plain(' (string).')]),
    bp([bold('Number()'), plain(': converte a string '), mono('"2"'), plain(' para o número '), mono('2'), plain(', pois os IDs do array são números.')]),
    bp([bold('contatos.find()'), plain(': retorna o primeiro elemento que satisfaz a condição (arrow function).')]),
    bp([bold('contatos.findIndex()'), plain(': retorna o índice (posição) do elemento no array.')]),

    h2('3.7 DELETE — Removendo Contatos'),
    np('Como formulários HTML só suportam métodos GET e POST, usamos POST para o DELETE também. A remoção usa o método filter() que cria um novo array sem o elemento indesejado.'),
    ...code("app.post('/contatos/deletar/:id', (req, res) => {\n  const id = Number(req.params.id);\n\n  // filter() cria um NOVO array com todos os contatos\n  // EXCETO o que tem o ID informado\n  contatos = contatos.filter(c => c.id !== id);\n\n  res.redirect('/');\n});"),
    bp([bold('filter()'), plain(': retorna um novo array com apenas os elementos que passam no teste da arrow function. É imutável — não modifica o array original, cria um novo.')]),

    h2('3.8 Iniciando o Servidor'),
    ...code("const PORTA = 3000;\n\napp.listen(PORTA, () => {\n  console.log(`Servidor rodando em http://localhost:${PORTA}`);\n  console.log('Pressione CTRL+C para parar.');\n});"),
    bp([bold('app.listen(porta, callback)'), plain(': coloca o servidor em escuta na porta informada. O callback executa uma vez quando o servidor estiver pronto.')]),
    bp([bold('Template literal'), plain(' (crase '), mono('`...${}...`'), plain('): permite inserir variáveis diretamente dentro de uma string usando '), mono('${}'), plain('.')]),
    pb()
  ];
}

// ─── CAP. 4: VIEWS EJS ────────────────────────────────────────
function chapter4() {
  return [
    h1('Capítulo 4 — As Views com EJS'),
    np('As views são os arquivos .ejs que o servidor usa para gerar as páginas HTML enviadas ao navegador. Elas mesclam HTML estático com trechos de JavaScript usando as tags especiais do EJS.'),

    h2('4.1 Tags do EJS'),
    np('Existem quatro tipos de tags EJS. Veja a tabela:'),
    ...tbl(
      ['Tag EJS', 'Função', 'Exemplo'],
      [
        [mono('<%= ... %>'), 'Imprime o valor (com escape HTML)',           [mono('<%= nome %>'), plain(' → Ana Silva')]],
        [mono('<% ... %>'),  'Executa código JavaScript (sem imprimir)', [mono('<% if (x > 0) { %>')]],
        [mono('<%# ... %>'), 'Comentário (não aparece no HTML)',      [mono('<%# isto é um comentário %>')]]
      ],
      [0.2, 0.42, 0.38]
    ),
    ...box('⚠️', 'Atenção: Comentários EJS multilinha', 'A tag <%# %> fecha ao encontrar o primeiro %>. Para comentários com múiltiplas linhas, use comentários HTML <!-- --> em vez de <%# %> para evitar erros de parsing.', WARNBG, 'E67E22'),

    h2('4.2 index.ejs — Listagem de Contatos'),
    np('Esta view exibe a tabela com todos os contatos. Usa um '), np([mono('if'), plain(' para verificar se há dados, e um '), mono('forEach'), plain(' para percorrer o array.')], { after: 0 }),
    ...code('<!-- Verifica se o array está vazio -->\n<% if (contatos.length === 0) { %>\n  <p>Nenhum contato cadastrado ainda.</p>\n<% } else { %>\n\n  <table>\n    <thead>\n      <tr>\n        <th>#</th><th>Nome</th><th>E-mail</th><th>Telefone</th><th>Ações</th>\n      </tr>\n    </thead>\n    <tbody>\n      <!-- Loop: uma linha por contato -->\n      <% contatos.forEach(function(contato) { %>\n        <tr>\n          <td><%= contato.id %></td>\n          <td><%= contato.nome %></td>\n          <td><%= contato.email %></td>\n          <td><%= contato.telefone %></td>\n          <td>\n            <a href="/contatos/editar/<%= contato.id %>">Editar</a>\n\n            <!-- Formulário de exclusão -->\n            <form action="/contatos/deletar/<%= contato.id %>" method="POST"\n              onsubmit="return confirm(\'Confirmar exclusão?\')">\n              <button type="submit">Excluir</button>\n            </form>\n          </td>\n        </tr>\n      <% }); %>\n    </tbody>\n  </table>\n\n  <p>Total: <strong><%= contatos.length %></strong> contato(s).</p>\n<% } %>'),
    bp([bold('contatos.length'), plain(': propriedade de array que retorna a quantidade de elementos.')]),
    bp([bold('contatos.forEach()'), plain(': percorre cada elemento do array chamando a função para cada um.')]),
    bp([bold('onsubmit="return confirm(...)"'), plain(': exibe uma janela de confirmação antes de enviar o formulário. Se o usuário clicar em "Cancelar", o formulário não é enviado.')]),

    h2('4.3 novo.ejs — Formulário de Criação'),
    np('Esta view exibe um formulário HTML em branco. O atributo '), np([mono('name'), plain(' de cada input é fundamental: é ele que determina o nome da chave em '), mono('req.body'), plain(' no servidor.')], { after: 0 }),
    ...code('<!-- action="/contatos" → envia para POST /contatos -->\n<!-- method="POST"     → método HTTP POST -->\n<form action="/contatos" method="POST" class="formulario">\n\n  <div class="campo">\n    <label for="nome">Nome completo:</label>\n    <!-- name="nome" → req.body.nome no servidor -->\n    <input type="text" id="nome" name="nome"\n           placeholder="Ex: João da Silva" required>\n  </div>\n\n  <div class="campo">\n    <label for="email">E-mail:</label>\n    <!-- type="email" → validação automática pelo navegador -->\n    <input type="email" id="email" name="email"\n           placeholder="Ex: joao@email.com" required>\n  </div>\n\n  <div class="campo">\n    <label for="telefone">Telefone:</label>\n    <input type="text" id="telefone" name="telefone"\n           placeholder="Ex: (11) 91234-5678" required>\n  </div>\n\n  <div class="botoes">\n    <!-- type="submit" → envia o formulário -->\n    <button type="submit">Salvar Contato</button>\n    <!-- Link de cancelar: volta sem salvar -->\n    <a href="/">Cancelar</a>\n  </div>\n\n</form>'),
    bp([bold('action'), plain(': define para qual rota o formulário envia os dados.')]),
    bp([bold('method="POST"'), plain(': define o método HTTP. Formulários só aceitam GET ou POST.')]),
    bp([bold('required'), plain(': atributo HTML5 que impede o envio se o campo estiver vazio.')]),

    h2('4.4 editar.ejs — Formulário de Edição'),
    np('Esta view é similar ao novo.ejs, mas os campos já vêm preenchidos com os dados atuais do contato. Isso é feito com o atributo '), np([mono('value'), plain(' nos inputs.')], { after: 0 }),
    ...code('<!-- A action usa o ID do contato para saber qual atualizar -->\n<form action="/contatos/editar/<%= contato.id %>" method="POST">\n\n  <div class="campo">\n    <label for="nome">Nome completo:</label>\n    <!-- value="..." preenche o campo com o dado atual -->\n    <input type="text" id="nome" name="nome"\n           value="<%= contato.nome %>" required>\n  </div>\n\n  <div class="campo">\n    <label for="email">E-mail:</label>\n    <input type="email" id="email" name="email"\n           value="<%= contato.email %>" required>\n  </div>\n\n  <div class="campo">\n    <label for="telefone">Telefone:</label>\n    <input type="text" id="telefone" name="telefone"\n           value="<%= contato.telefone %>" required>\n  </div>\n\n  <div class="botoes">\n    <button type="submit">Salvar Alterações</button>\n    <a href="/">Cancelar</a>\n  </div>\n\n</form>'),
    bp([bold('value="<%= contato.nome %>"'), plain(': o EJS substitui '), mono('<%= contato.nome %>'), plain(' pelo nome real, preenchendo o campo automaticamente.')]),
    bp([bold('Reutilização de código'), plain(': perceba como editar.ejs e novo.ejs são muito parecidos. Em sistemas maiores, cria-se um template parcial (_form.ejs) para evitar repetição.')]),
    pb()
  ];
}

// ─── CAP. 5: CSS ──────────────────────────────────────────────
function chapter5() {
  return [
    h1('Capítulo 5 — Estilo Visual: style.css'),
    np('O arquivo public/style.css define a aparência visual do sistema. Por ser colocado na pasta '), np([mono('public'), plain(', o Express o serve automaticamente quando o navegador faz uma requisição de '), mono('/style.css'), plain('.')], { after: 0 }),

    h2('5.1 Principais seções do CSS'),
    bp([bold('Reset (*): '), plain('remove margens e paddings padrão do navegador. O '), mono('box-sizing: border-box'), plain(' garante que padding e border não aumentem o tamanho do elemento.')]),
    bp([bold('Container: '), plain('centraliza o conteúdo e limita a largura máxima da página a 860px, tornando-a mais legível em telas largas.')]),
    bp([bold('Tabela: '), plain('o '), mono('border-collapse: collapse'), plain(' remove espaços entre células. O '), mono('tbody tr:hover'), plain(' dá feedback visual ao usuário.')]),
    bp([bold('Botões (.btn): '), plain('classes reutilizáveis com cores diferentes para cada ação (verde = criar, azul = editar, vermelho = excluir, cinza = cancelar).')]),
    bp([bold('Formulários: '), plain('os campos (.campo) são empilhados verticalmente usando '), mono('display: flex; flex-direction: column'), plain('. O '), mono('input:focus'), plain(' destaca o campo selecionado em azul.')]),
    ...box('💡', 'Boa prática', 'Separar o CSS em arquivo próprio (em vez de colocar <style> direto no HTML) facilita a manutenção e permite que o navegador salve o arquivo em cache.', TIPBG, '27AE60'),
    pb()
  ];
}

// ─── CAP. 6: EXECUTANDO ───────────────────────────────────────
function chapter6() {
  return [
    h1('Capítulo 6 — Executando e Testando o Sistema'),

    h2('6.1 Iniciando o Servidor'),
    np('Abra o terminal, entre na pasta do projeto e execute:'),
    ...code('cd crud-contatos\nnode server.js'),
    np('Se tudo estiver correto, o terminal exibirá:'),
    ...code('Servidor rodando em http://localhost:3000\nPressione CTRL+C para parar.'),
    ...box('📌', 'Dica: nodemon', 'Durante o desenvolvimento, use o nodemon para reiniciar o servidor automaticamente ao salvar arquivos:\n  npm install -g nodemon\n  nodemon server.js', TIPBG, '27AE60'),

    h2('6.2 Testando o CRUD no Navegador'),
    np('Abra o navegador e acesse '), np([mono('http://localhost:3000'), plain('. Siga o roteiro abaixo para testar cada operação:')], { after: 0 }),
    ...tbl(
      ['Operação', 'O que fazer', 'Resultado esperado'],
      [
        ['READ',   'Acessar http://localhost:3000',              'Tabela com 3 contatos pré-cadastrados'],
        ['CREATE', 'Clicar em "+ Novo Contato" e preencher',    'Novo contato aparece na lista'],
        ['UPDATE', 'Clicar em "Editar" em qualquer contato',    'Formulário preenchido; após salvar, volta à lista'],
        ['DELETE', 'Clicar em "Excluir" e confirmar',           'Contato removido da lista'],
      ],
      [0.18, 0.42, 0.40]
    ),

    h2('6.3 Mapa de Rotas'),
    np('Resumo de todas as rotas implementadas no server.js:'),
    ...tbl(
      ['Método', 'Rota', 'Ação'],
      [
        ['GET',  '/',                         'Lista todos os contatos'],
        ['GET',  '/contatos/novo',            'Exibe formulário de criação'],
        ['POST', '/contatos',                 'Salva novo contato'],
        ['GET',  '/contatos/editar/:id',      'Exibe formulário de edição'],
        ['POST', '/contatos/editar/:id',      'Salva alterações no contato'],
        ['POST', '/contatos/deletar/:id',     'Remove o contato'],
      ],
      [0.14, 0.36, 0.50]
    ),
    pb()
  ];
}

// ─── CAP. 7: EXERCÍCIOS ───────────────────────────────────────
function chapter7() {
  return [
    h1('Capítulo 7 — Exercícios Propostos'),
    np('Pratique o que aprendeu com os exercícios abaixo. Eles estão ordenados do mais simples ao mais desafiador.'),
    el(80),

    h2('Exercício 1 — Novo Campo'),
    np([bold('Objetivo: '), plain('adicionar o campo "cidade" ao cadastro de contatos.')]),
    bp('No array de dados (server.js), adicione a propriedade cidade a cada contato existente.'),
    bp('No formulário novo.ejs, adicione um novo campo <input> com name="cidade".'),
    bp('No formulário editar.ejs, adicione o mesmo campo com value="<%= contato.cidade %>".'),
    bp('Na rota POST /contatos, adicione cidade: req.body.cidade ao objeto novoContato.'),
    bp('Na rota POST /contatos/editar/:id, inclua cidade: req.body.cidade na atualização.'),
    bp('Na tabela index.ejs, adicione a coluna Cidade.'),
    el(80),

    h2('Exercício 2 — Validação no Servidor'),
    np([bold('Objetivo: '), plain('validar os dados antes de salvar, retornando um erro ao usuário.')]),
    bp('Na rota POST /contatos, verifique se nome, email e telefone foram enviados.'),
    bp('Se algum campo estiver vazio, use res.redirect(\'/contatos/novo\') para voltar ao formulário.'),
    bp([bold('Desafio extra: '), plain('use uma query string para exibir a mensagem de erro no formulário, como '), mono('res.redirect(\'/contatos/novo?erro=campos_obrigatorios\')'), plain(', e leia com '), mono('req.query.erro'), plain(' na rota GET.')]),
    el(80),

    h2('Exercício 3 — Página de Detalhes'),
    np([bold('Objetivo: '), plain('criar uma página que exibe todos os dados de um único contato.')]),
    bp([plain('Crie a rota '), mono('GET /contatos/:id'), plain(' que busca o contato pelo ID.')]),
    bp([plain('Crie o arquivo '), mono('views/detalhe.ejs'), plain(' que exibe nome, e-mail e telefone em cards separados.')]),
    bp([plain('Na tabela index.ejs, transforme o nome do contato em um link: '), mono('<a href="/contatos/<%= contato.id %>">'), plain('.')]),
    el(80),

    h2('Exercício 4 — Contador de Visitas'),
    np([bold('Objetivo: '), plain('contar quantas vezes a página inicial foi acessada.')]),
    bp([plain('Declare uma variável '), mono('let visitas = 0;'), plain(' no início do server.js.')]),
    bp([plain('Na rota GET /, incremente: '), mono('visitas++;'), plain(' antes do res.render.')]),
    bp([plain('Passe visitas para o template: '), mono('res.render(\'index\', { contatos, titulo, visitas })'), plain('.')]),
    bp([plain('Exiba no rodapé do index.ejs: '), mono('Esta página foi acessada <%= visitas %> vez(es).')]),
    el(80),

    h2('Exercício 5 — Busca por Nome'),
    np([bold('Objetivo: '), plain('implementar uma barra de pesquisa que filtra contatos pelo nome.')]),
    bp([plain('No index.ejs, adicione um formulário de busca com '), mono('method="GET"'), plain(' e '), mono('action="/"'), plain('.')]),
    bp([plain('Na rota GET /, leia o parâmetro de busca: '), mono('const busca = req.query.busca || \'\';'), plain('.')]),
    bp([plain('Filtre o array: '), mono('const resultado = contatos.filter(c => c.nome.toLowerCase().includes(busca.toLowerCase()));'), plain('.')]),
    bp([plain('Passe '), mono('resultado'), plain(' em vez de '), mono('contatos'), plain(' para o template.')]),
    el(80),

    h2('Exercício 6 — Persistência em Arquivo (Desafio)'),
    np([bold('Objetivo: '), plain('salvar os contatos em um arquivo JSON para que os dados não se percam ao reiniciar.')]),
    bp([plain('Use o módulo nativo '), mono('fs'), plain(' (File System): '), mono('const fs = require(\'fs\');'), plain('.')]),
    bp([plain('Ao iniciar o servidor, leia o arquivo se ele existir: '), mono('fs.existsSync(\'dados.json\')'), plain('.')]),
    bp([plain('A cada criação, atualização ou exclusão, grave o array com: '), mono('fs.writeFileSync(\'dados.json\', JSON.stringify(contatos))'), plain('.')]),
    ...box('🏆', 'Parabéns!', 'Ao concluir todos os exercícios, você terá um sistema web completo com validação, busca, página de detalhes e persistência de dados. Esses são os mesmos conceitos usados em sistemas profissionais com banco de dados.', TIPBG, '27AE60'),
  ];
}

// =============================================================
//  MONTAGEM DO DOCUMENTO
// =============================================================
const header = new Header({
  children: [new Paragraph({
    alignment: AlignmentType.RIGHT,
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 1 } },
    spacing: { before: 0, after: 100 },
    children: [
      new TextRun({ text: 'ETEC Campo Limpo Paulista  |  Sistemas Web I  |  ', font: F, size: 18, color: '666666' }),
      new TextRun({ text: 'Prof. Marcelo Macrino dos Santos', font: F, size: 18, bold: true, color: BLUE }),
    ]
  })]
});

const footer = new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 1 } },
    spacing: { before: 100, after: 0 },
    children: [
      new TextRun({ text: 'CRUD com Node.js + Express + EJS  —  Página ', font: F, size: 18, color: '666666' }),
      new TextRun({ children: [PageNumber.CURRENT], font: F, size: 18, bold: true, color: BLUE }),
    ]
  })]
});

const emptyHeader = new Header({ children: [new Paragraph({ children: [] })] });
const emptyFooter = new Footer({ children: [new Paragraph({ children: [] })] });

const allChildren = [
  ...coverPage(),
  ...tocSection(),
  ...chapter1(),
  ...chapter2(),
  ...chapter3(),
  ...chapter4(),
  ...chapter5(),
  ...chapter6(),
  ...chapter7(),
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: F, size: 24 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: F, size: 36, bold: true, color: BLUE },
        paragraph: { spacing: { before: 480, after: 240 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: F, size: 28, bold: true, color: LBLUE },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: F, size: 24, bold: true, color: '444444' },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: 'bullets',
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 600, hanging: 300 } } } },
          { level: 1, format: LevelFormat.BULLET, text: '○', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1080, hanging: 300 } } } },
        ]
      },
      { reference: 'numbers',
        levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 600, hanging: 300 } } } }
        ]
      },
    ]
  },
  sections: [{
    properties: {
      titlePage: true,
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: { default: header, first: emptyHeader },
    footers: { default: footer, first: emptyFooter },
    children: allChildren
  }]
});

Packer.toBuffer(doc).then(buffer => {
  const outPath = '/sessions/beautiful-friendly-hypatia/mnt/outputs/Apostila_SistemasWebI_CRUD_NodeJS.docx';
  fs.writeFileSync(outPath, buffer);
  console.log('Apostila gerada com sucesso!');
  console.log('Arquivo: ' + outPath);
}).catch(err => {
  console.error('Erro ao gerar apostila:', err.message);
  process.exit(1);
});
