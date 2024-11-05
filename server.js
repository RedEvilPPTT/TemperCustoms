const express = require('express');
const bodyParser = require('body-parser');
const session = require('express-session');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx'); // Importa o módulo xlsx
const app = express();
const PORT = 3000;

// Configuração do body-parser para ler dados do formulário
app.use(bodyParser.urlencoded({ extended: true }));

// Configuração para servir arquivos estáticos (HTML, CSS, JS)
app.use(express.static(path.join(__dirname, '')));

// Configuração de sessão
app.use(session({
  secret: 'chave-secreta', // Em produção, use uma chave secreta mais segura
  resave: false,
  saveUninitialized: true,
  cookie: { maxAge: 60000 } // Expira em 1 minuto para teste
}));


// Middleware para interpretar JSON
app.use(express.json());
app.use(express.urlencoded({ extended: true })); // Para lidar com formulários se necessário
app.use(express.static('public'));

// Carregar lista de usuários
const usuarios = JSON.parse(fs.readFileSync('usuarios.json', 'utf-8')).usuarios;

// Caminho do arquivo usuarios.json
const usuariosFilePath = path.join(__dirname, 'usuarios.json');

// Função para ler os usuários atuais do arquivo JSON
const lerUsuarios = () => {
  try {
      const data = fs.readFileSync(usuariosFilePath, 'utf8');
      return JSON.parse(data);
  } catch (error) {
      return { usuarios: [] };
  }
};

// Função para salvar os usuários no arquivo JSON
const salvarUsuarios = (dadosUsuarios) => {
  fs.writeFileSync(usuariosFilePath, JSON.stringify(dadosUsuarios, null, 2));
};

// Função para ler carros da planilha Excel
function lerCarrosDoExcel() {
    const workbook = xlsx.readFile('carros.xlsx'); // Nome do arquivo Excel
    const sheetName = workbook.SheetNames[0]; // Nome da primeira aba
    const sheet = workbook.Sheets[sheetName];
  
    // Converte a planilha em JSON
    const data = xlsx.utils.sheet_to_json(sheet);

    console.log("Dados lidos do Excel:", data);
    // Formata os dados para enviar ao frontend
    return data.map((row , index) => ({
      id: index + 1, // Adiciona um ID único para cada carro
      nome: row.Nome,  // Supondo que a coluna de nomes no Excel se chama "Nome"
      autocolantes: row.autocolantes || 0,
      buzina: row.buzina || 0,
      chaparia: row.chaparia || 0,
      matricula: row.matricula || 0,
      pinturas: row.pinturas || 0,
      perolado: row.perolado || 0,
      stance: row.stance || 0,
      pneublindadodrift: row.pneublindadodrift || 0,
      jantes: row.jantes || 0,
      cordasrodas: row.cordasrodas || 0,
      armadura: row.armadura || 0,
      motor: row.motor || 0,
      suspensao: row.suspensao || 0,
      transmissao: row.transmissao || 0,
      travoes: row.travoes || 0,
      turbo: row.turbo || 0,
      kits: row.kits || 0,
      vendacontrato: row.vendacontrato || 0,
      logo: row.logo_url

    }));
  }
  
  // Rota para fornecer dados dos carros ao frontend
app.get('/api/carros', (req, res) => {
    const carros = lerCarrosDoExcel();
    res.json(carros); // Envia os dados em formato JSON
});

// Função para calcular a semana, começando no domingo e terminando no sábado
function getCustomWeek(date) {
  // Obtem o primeiro domingo do ano
  const startOfYear = new Date(date.getFullYear(), 0, 1);
  while (startOfYear.getDay() !== 0) { // 0 representa domingo
      startOfYear.setDate(startOfYear.getDate() + 1);
  }

  // Calcula a diferença em dias entre a data e o primeiro domingo do ano
  const diffInDays = Math.floor((date - startOfYear) / 86400000);

  // Divide o número de dias pela duração de uma semana para obter o número da semana
  const weekNumber = Math.floor(diffInDays / 7) + 2;

  return weekNumber;
}

app.post('/api/salvarFatura', (req, res) => {
  const { usuario, carro, totalKits, totalVendaCarro, texto, valorTotal, valorComIva } = req.body;

  const dateObj = new Date();
  const semana = `Semana ${getCustomWeek(dateObj)}`;
  // Log para verificar os dados recebidos
  console.log("Dados recebidos no servidor:", { usuario, carro, totalKits, totalVendaCarro, texto, valorTotal, valorComIva });

  if (!usuario || !carro || totalKits === undefined || totalVendaCarro === undefined || texto === undefined || valorTotal === undefined || valorComIva === undefined) {
      return res.status(400).json({ message: "Dados inválidos. Verifique os campos." });
  }

  const novaFatura = { usuario, carro, totalKits, totalVendaCarro, texto, valorTotal, valorComIva, semana, data: new Date().toLocaleString() };

  const nomeArquivo = 'faturas.xlsx';
  let workbook;
  let sheet;

  if (fs.existsSync(nomeArquivo)) {
      workbook = xlsx.readFile(nomeArquivo);
      sheet = workbook.Sheets['Faturas'];
  } else {
      workbook = xlsx.utils.book_new();
      sheet = xlsx.utils.json_to_sheet([]);
      xlsx.utils.book_append_sheet(workbook, sheet, 'Faturas');
  }

  const dadosExistentes = xlsx.utils.sheet_to_json(sheet);
  dadosExistentes.push(novaFatura);

  const novoSheet = xlsx.utils.json_to_sheet(dadosExistentes);
  workbook.Sheets['Faturas'] = novoSheet;
  xlsx.writeFile(workbook, nomeArquivo);

  res.status(200).json({ message: 'Fatura salva com sucesso' });
});


// Rota para apresentar a página de registro
app.get('/admin', (req, res) => {
  const htmlForm = `
      <!DOCTYPE html>
      <html lang="pt">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Registro de Utilizador</title>
          <link rel="stylesheet" href="/admin.css">
      </head>
      <body>
          <h2>Adicionar Utilizador</h2>
          <form action="/admin/register" method="POST">
              <label for="username">Nome de Utilizador:</label>
              <input type="text" id="username" name="username" required>
              <br>
              <button type="submit">Adicionar Utilizador</button>
          </form>
      </body>
      </html>
  `;
  res.send(htmlForm);
});


// Rota para processar o registro do usuário
app.post('/admin/register', (req, res) => {
  const { username } = req.body;

  // Validação para garantir que o username foi fornecido
  if (!username) {
      return res.status(400).json({ erro: 'Username é obrigatório.' });
  }

  // Carrega os usuários atuais
  const dadosUsuarios = lerUsuarios();
  const { usuarios } = dadosUsuarios;

  // Checa se o username já está em uso
  const usuarioExistente = usuarios.find(user => user.username === username);
  if (usuarioExistente) {
      return res.status(400).send(`
          <script>
              alert('Este username já está em uso.');
              window.location.href = '/admin';
          </script>
      `);
  }

  // Adiciona o novo username ao array de usuários
  usuarios.push({ username });

  // Salva o novo array de usuários no arquivo
  salvarUsuarios(dadosUsuarios);

  // Retorna um alerta de sucesso e redireciona para a página de registro
  res.send(`
      <script>
          alert('Utilizador "${username}" adicionado com sucesso!');
          window.location.href = '/admin';
      </script>
  `);
});


app.get('/faturas', (req, res) => {
  const file = 'faturas.xlsx';
  const { username, semana } = req.query;


  if (fs.existsSync(file)) {
      const workbook = xlsx.readFile(file);
      const sheet = workbook.Sheets['Faturas'];
      const faturas = xlsx.utils.sheet_to_json(sheet);

        // Adiciona o campo "Semana" a cada fatura com base na data
      faturas.forEach(fatura => {
        if (fatura.data) {
            const date = new Date();
            fatura.semana = `Semana ${getCustomWeek(date)}`;
        }
      });

      // Lista de usuários únicos nas faturas
      const usuariosDisponiveis = [...new Set(faturas.map(fatura => fatura.usuario.toLowerCase()))];
      const semanasDisponiveis = [...new Set(faturas.map(fatura => fatura.semana))];

      // Se um username foi pesquisado, verifica se ele existe na lista de usuários
      const usernameLower = username ? username.toLowerCase() : null;

      if (username && !usuariosDisponiveis.includes(usernameLower)) {
        return res.send(`
                <script>
                    alert("Utilizador '${username}' não encontrado nas faturas.");
                    window.location.href = "/faturas";
                </script>
        `);
    }


        // Filtra as faturas com base no username e na semana
      const faturasFiltradas = faturas.filter(fatura => {
        const matchesUsername = !usernameLower || (fatura.usuario && fatura.usuario.toLowerCase() === usernameLower);
        const matchesSemana = !semana || fatura.semana === semana;
        return matchesUsername && matchesSemana;
      });

      /*const faturasFiltradas = usernameLower 
      ? faturas.filter(fatura => fatura.usuario && fatura.usuario.toLowerCase().includes(username.toLowerCase()))
      : faturas;*/

      const totalFaturas = faturasFiltradas.length;

      // Calcula o total com IVA para todas as faturas filtradas
      const valorTotal = faturasFiltradas.reduce((total, fatura) => {
        const valorTotal = parseFloat(fatura.valorTotal) || 0;
        return total + valorTotal;
      }, 0);
      

        // Calcula o valor total da semana (sem IVA) para as faturas filtradas
        const valorTotalSemana = faturasFiltradas.reduce((total, fatura) => {
          const valorTotal = parseFloat(fatura.valorTotal) || 0;
          return total + valorTotal;
      }, 0);      

      // Geração da tabela HTML para exibir as faturas
      let html = `
      <!DOCTYPE html>
      <html lang="pt">
      <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Faturas</title>
      <link rel="stylesheet" href="/faturas.css">
      <script>
          function confirmarExclusao() {
              const semanaSelecionada = document.querySelector('select[name="semana"]').value;
              if (!semanaSelecionada) {
                  alert("Por favor, selecione uma semana antes de tentar apagar.");
                  return false;
              }
              return confirm("Tem certeza que deseja apagar todas as faturas da " + semanaSelecionada + "?");
          }
      </script>
      </head>
      <body>
      <h2>Faturas</h2>
      <p>Total de Faturas para <strong>${username || 'Todos os utilizadores'}</strong>:<strong>${totalFaturas}</strong></p>
      <p>Total Geral para <strong>${username || 'Todos os utilizadores'}</strong>:<strong>${valorTotal.toFixed(2)} €</strong></p>
      <p>Total da Semana para <strong>"${username || 'Todos os usuários'}" na ${semana || 'todas as semanas'}</strong>: <strong>${valorTotalSemana.toFixed(2)} €</strong></p>
      <form method="GET" action="/faturas">
          <label>Pesquisar por Usuário:</label>
          <input type="text" name="username" placeholder="Digite o username" value="${username || ''}">
          <label>Filtrar por Semana:</label>
          <select name="semana">
            <option value="">Todas as semanas</option>`;
        
        // Adiciona as opções de semana ao select
        semanasDisponiveis.forEach(sem => {
            html += `<option value="${sem}" ${sem === semana ? 'selected' : ''}>${sem}</option>`;
        });

        html += `</select>
          <button type="submit">Pesquisar</button>
      </form>

      <form method="POST" action="/apagarFaturasSemana" onsubmit="return confirmarExclusao()">
          <input type="hidden" name="semana" value="${semana || ''}">
          <button type="submit" ${!semana ? 'disabled' : ''}>Apagar Faturas da ${semana || 'semana selecionada'}</button>
      </form>
      <table border="1">
          <tr>
              <th>Usuário</th>
              <th>Carro</th>
              <th>Total de Kits</th>
              <th>Total de Venda Carro Anuncios</th>
              <th>Valor Total</th>
              <th>Valor com IVA</th>
              <th>Semana</th>
              <th>Data</th>
              <th>Observações</th>
          </tr>`;

      // Itera sobre as faturas filtradas e gera a tabela
      faturasFiltradas.forEach(fatura => {
          html += `
              <tr>
                  <td>${fatura.usuario || 'N/A'}</td>
                  <td>${fatura.carro || 'N/A'}</td>
                  <td>${fatura.totalKits || '0'}</td>
                  <td>${fatura.totalVendaCarro || '0'}</td>
                  <td>${fatura.valorTotal || '0'} €</td>
                  <td>${fatura.valorComIva || '0'} €</td> 
                  <td>${fatura.semana || 'N/A'}</td>
                  <td>${fatura.data || 'N/A'}</td>
                  <td>${fatura.texto || 'N/A'}</td>
              </tr>`;
      });

      html += '</table></body></html>';

      res.send(html);
  } else {
      res.send('Nenhuma fatura encontrada.');
  }
});


// Rota para apagar faturas da semana especificada
app.post('/apagarFaturasSemana', (req, res) => {
  const { semana } = req.body;

  console.log("Semana recebida para exclusão:", semana);

  if (!semana) {
    console.log("Nenhuma semana selecionada.");
    return res.redirect('/faturas');
  }

  const file = 'faturas.xlsx';
  if (fs.existsSync(file)) {
    const workbook = xlsx.readFile(file);
    const sheet = workbook.Sheets['Faturas'];
    const faturas = xlsx.utils.sheet_to_json(sheet);

    // Filtra para manter apenas faturas fora da semana especificada
    const faturasAtualizadas = faturas.filter(fatura => fatura.semana !== semana);

    console.log("Faturas antes da exclusão:", faturas.length);
    console.log("Faturas após exclusão da semana:", faturasAtualizadas.length);

    // Converte as faturas atualizadas de volta para o formato de planilha
    const novoSheet = xlsx.utils.json_to_sheet(faturasAtualizadas);
    workbook.Sheets['Faturas'] = novoSheet;
    xlsx.writeFile(workbook, file);

    console.log("Exclusão concluída e arquivo atualizado.");
  }

  res.redirect('/faturas');
});

// Rota de login
app.post('/login', (req, res) => {
  const { username } = req.body;

  // Verificar se o usuário existe na base de dados
  const usuarioEncontrado = usuarios.find(user => user.username === username);

  if (usuarioEncontrado) {
    // Armazenar o nome de usuário na sessão e redirecionar para a página principal
    req.session.username = username;
    res.send(`
      <script>
        localStorage.setItem("username", "${username}");
        window.location.href = "/home";
      </script>
    `);
  } else {
    res.send(`
      <script>
          alert('Utilizador ${username} não encontrado, porfavor tente novamente ou contacte o administrador');
          window.location.href = '/index.html';
      </script>
  `);
  }

});

// Rota para a página principal (proteção de rota)
app.get('/home', (req, res) => {
  if (req.session.username) {
    res.sendFile(path.join(__dirname, 'faturar.html'));
  } else {
    res.redirect('/');
  }
});

// Rota de logout
/*app.get('/logout', (req, res) => {
    req.session.destroy(err => {
      if (err) {
        return res.send('Erro ao fazer logout');
      }
      res.redirect('/login.html'); // Redireciona para a página de login
    });
});*/
  


// Rota de logout
app.get('/logout', (req, res) => {
  req.session.destroy(err => {
    if (err) {
      return res.send('Erro ao fazer logout');
    }
    res.send(`
      <script>
        localStorage.removeItem("username");
        window.location.href = "/index.html";
      </script>
    `);
  });
});

// Iniciar o servidor
app.listen(PORT, () => {
  console.log(`Servidor rodando em http://localhost:${PORT}`);
});
