/*função que exibe o site*/
function doGet() 
{
  return HtmlService.createHtmlOutputFromFile('index')
  .setTitle('Thesia');
}

//objeto que armazena os IDs de arquivos usados do drive
const propriedades = PropertiesService.getScriptProperties();

const CONFIG = 
{
  PLANILHA_ID: propriedades.getProperty('PLANILHA_ID'),
  MODELO_DOC_ID: propriedades.getProperty('MODELO_DOC_ID')
}

//função que salva a avaliação na planilha do drive
function salvarAvaliacao(dados)
{
  //abre a planilha por id
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);

  //seleciona as abas usadas
  const abaAvaliacoes = planilha.getSheetByName("AVALIACOES");
  const abaArtigo = planilha.getSheetByName("ARTIGO");
  const abaRelatorio = planilha.getSheetByName("RELATORIO");
  const abaOral = planilha.getSheetByName("ORAL");

  //cria um id unico para a avaliação atual
  const idAvaliacao = Utilities.getUuid();

  //data/hora para os registros
  const dataHora = new Date();

  //pesos das notas
  const pesoArtigo = 0.3;
  const pesoRelatorio = 0.3;
  const pesoOral = 0.4;


  //soma dos critérios de Artigo
  const totalArtigo = dados.artigo1 + dados.artigo2 + dados.artigo3
  + dados.artigo4 + dados.artigo5 + dados.artigo6 + dados.artigo7;

  //soma dos critérios de Relatório
  const totalRelatorio = dados.relatorio1 + dados.relatorio2 + dados.relatorio3
  + dados.relatorio4 + dados.relatorio5;

  //soma dos critérios de Oratória aluno 1
  const totalOral1 = dados.oral1aluno1 + dados.oral2aluno1 + dados.oral3aluno1
  + dados.oral4aluno1 + dados.oral5aluno1;

  //soma dos critérios de Oratória aluno 2
  const totalOral2 = dados.oral1aluno2 + dados.oral2aluno2 + dados.oral3aluno2
  + dados.oral4aluno2 + dados.oral5aluno2;

  //soma dos critérios de Oratória aluno 3
  const totalOral3 = dados.oral1aluno3 + dados.oral2aluno3 + dados.oral3aluno3
  + dados.oral4aluno3 + dados.oral5aluno3;

  //soma dos critérios de Oratória aluno 4
  const totalOral4 = dados.oral1aluno4 + dados.oral2aluno4 + dados.oral3aluno4
  + dados.oral4aluno4 + dados.oral5aluno4;


  //soma as notas com seus pesos para gerar nota final de cada aluno
  const final1 = (totalArtigo * pesoArtigo) + (totalRelatorio * pesoRelatorio) + (totalOral1 * pesoOral);
  const final2 = (totalArtigo * pesoArtigo) + (totalRelatorio * pesoRelatorio) + (totalOral2 * pesoOral);
  const final3 = (totalArtigo * pesoArtigo) + (totalRelatorio * pesoRelatorio) + (totalOral3 * pesoOral);
  const final4 = (totalArtigo * pesoArtigo) + (totalRelatorio * pesoRelatorio) + (totalOral4 * pesoOral);

  //salva o resumo na aba AVALIACOES
  abaAvaliacoes.appendRow([
    idAvaliacao, dataHora, dados.avaliador, dados.titulo, dados.aluno1, final1, dados.aluno2, final2, dados.aluno3, final3, dados.aluno4, final4
  ]);

  // salva os critérios do artigo na aba ARTIGO
  abaArtigo.appendRow([
    idAvaliacao, dados.artigo1, dados.artigo2, dados.artigo3, dados.artigo4, dados.artigo5, dados.artigo6, dados.artigo7, totalArtigo
  ]);

  //salva os critérios de relatório na aba RELATORIO
  abaRelatorio.appendRow([
    idAvaliacao, dados.relatorio1, dados.relatorio2, dados.relatorio3, dados.relatorio4, dados.relatorio5, totalRelatorio
  ]);

  //salva os critérios de oratória na aba ORAL - linha do aluno1
  abaOral.appendRow([
    idAvaliacao, 1, dados.aluno1, dados.oral1aluno1, dados.oral2aluno1, dados.oral3aluno1, dados.oral4aluno1, dados.oral5aluno1, totalOral1
  ]);

  //salva os critérios de oratória na aba ORAL - linha do aluno2
  abaOral.appendRow ([
    idAvaliacao, 2, dados.aluno2, dados.oral1aluno2, dados.oral2aluno2, dados.oral3aluno2, dados.oral4aluno2, dados.oral5aluno2, totalOral2
  ]);

  //salva os critérios de oratória na aba ORAL - linha do aluno3
  abaOral.appendRow ([
    idAvaliacao, 3, dados.aluno3, dados.oral1aluno3, dados.oral2aluno3, dados.oral3aluno3, dados.oral4aluno3, dados.oral5aluno3, totalOral3
  ]);

  //salva os critérios de oratória na aba ORAL - linha do aluno4
  abaOral.appendRow ([
    idAvaliacao, 4, dados.aluno4, dados.oral1aluno4, dados.oral2aluno4, dados.oral3aluno4, dados.oral4aluno4, dados.oral5aluno4, totalOral4
  ]);

  //devolve os resultados para o html
  return {
    idAvaliacao: idAvaliacao,
    totalArtigo: totalArtigo,
    totalRelatorio: totalRelatorio,

    totalOral1: totalOral1,
    totalOral2: totalOral2,
    totalOral3: totalOral3,
    totalOral4: totalOral4,

    aluno1: dados.aluno1,
    final1: final1,

    aluno2: dados.aluno2,
    final2: final2,

    aluno3: dados.aluno3,
    final3: final3,

    aluno4: dados.aluno4,
    final4: final4
  };
}