/*função que exibe o site*/

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
  .setTitle('Thesia');
}

//objeto que armazena os IDs de arquivos usados do drive
const propriedades = PropertiesService.getScriptProperties();

const CONFIG = {
  PLANILHA_ID: propriedades.getProperty('PLANILHA_ID'),
  MODELO_DOC_ID: propriedades.getProperty('MODELO_DOC_ID')
}

/*função que salva os dados na planilha
  dados de entrada: data/hora (gerado automaticamente), avaliador, titulo*/

function salvarTeste(avaliador, titulo) {

  //pega planilha por id
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  
  //seleciona a página da planilha
  const aba = planilha.getSheetByName('TESTE');

  //adiciona linha com dados inseridos no banco
  aba.appendRow([
    new Date(),
    avaliador,
    titulo
  ]);

  return 'Dados salvos com sucesso!';
}


//função que gera um arquivo txt com os dados informados
//saída: txt com Nome e Titulo do TCC
function criarArquivoTeste(avaliador, titulo) {
  const conteudo = 'Avaliador: ' + avaliador + '\n' + 'Título do TCC: ' + titulo;

  DriveApp.createFile('Teste_Thesia.txt', conteudo);

  return 'Arquivo criado com sucesso!';
}

//função que gera cópia do modelo de doc
function gerarDocumentoTeste(avaliador, titulo) {

  //id do arquivo de documento modelo a ser preenchido
  const arquivoModelo = DriveApp.getFileById(CONFIG.MODELO_DOC_ID);

  //faz uma cópia e altera o título
  const copia = arquivoModelo.makeCopy(titulo + ' - ' + avaliador);

  //abre o doc por id da cópia
  const documento = DocumentApp.openById(copia.getId());

  //pega o corpo do texto para editar
  const corpo = documento.getBody();

  //substitui os dados digitados no documento
  corpo.replaceText('<<AVALIADOR>>', avaliador);

  corpo.replaceText('<<TITULO>>', titulo);

  //salva e fecha
  documento.saveAndClose();

  //pega a copia criada como arquivo do drive e converte em pdf
  const copiaDocumento = DriveApp.getFileById(copia.getId());

  const pdf = copiaDocumento.getAs(MimeType.PDF);

  pdf.setName(copia.getName() + '.pdf');

  DriveApp.createFile(pdf);

  return 'Documento e PDF criados com sucesso!';
}

function salvarAvaliacaoTeste(dados)
{
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);

  const abaAvaliacoes = planilha.getSheetByName("AVALIACOES");
  const abaArtigo = planilha.getSheetByName("ARTIGO");

  //cria um id unico para a avaliação atual
  const idAvaliacao = Utilities.getUuid();

  //data/hora para os registros
  const dataHora = new Date();

  //pesos das notas
  const pesoArtigo = 0.3;
  const pesoRelatorio = 0.3;
  const pesoOral = 0.4;


  //soma dos critérios de Artigo e
  const totalArtigo = dados.artigo1 + dados.artigo2 + dados.artigo3
  + dados.artigo4 + dados.artigo5 + dados.artigo6 + dados.artigo7;


  //soma as notas para gerar nota final
  const final1 = (totalArtigo * pesoArtigo) + (dados.relatorio * pesoRelatorio) + (dados.oral1 * pesoOral);
  const final2 = (totalArtigo * pesoArtigo) + (dados.relatorio * pesoRelatorio) + (dados.oral2 * pesoOral);
  const final3 = (totalArtigo * pesoArtigo) + (dados.relatorio * pesoRelatorio) + (dados.oral3 * pesoOral);
  const final4 = (totalArtigo * pesoArtigo) + (dados.relatorio * pesoRelatorio) + (dados.oral4 * pesoOral);

  //salva o resumo na aba AVALIACOES
  abaAvaliacoes.appendRow([
    idAvaliacao, dataHora, dados.avaliador, dados.titulo, dados.aluno1, final1, dados.aluno2, final2, dados.aluno3, final3, dados.aluno4, final4
  ]);

  // salva os critérios do artigo na aba ARTIGO
  abaArtigo.appendRow([
    idAvaliacao, dados.artigo1, dados.artigo2, dados.artigo3, dados.artigo4, dados.artigo5, dados.artigo6, dados.artigo7, totalArtigo
  ]);

  //devolve os resultados para o html
  return {
    idAvaliacao: idAvaliacao,
    totalArtigo: totalArtigo,

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