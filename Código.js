/*função que exibe o site*/

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
  .setTitle('Thesia');
}

//objeto que armazena os IDs de arquivos usados do drive
const CONFIG = {
  PLANILHA_ID: 'xxxxxxxxxxx',
  MODELO_DOC_ID: 'xxxxxxxxxx'
}

/*função que salva os dados na planilha
  dados de entrada: data/hora (gerado automaticamente), professor, titulo*/

function salvarTeste(professor, titulo) {

  //pega planilha por id
  const planilha = SpreadsheetApp.openById.CONFIG.PLANILHA_ID;
  
  //seleciona a página da planilha
  const aba = planilha.getSheetByName('TESTE');

  //adiciona linha com dados inseridos no banco
  aba.appendRow([
    new Date(),
    professor,
    titulo
  ]);

  return 'Dados salvos com sucesso!';
}


//função que gera um arquivo txt com os dados informados
//saída: txt com Nome e Titulo do TCC
function criarArquivoTeste(professor, titulo) {
  const conteudo = 'Professor: ' + professor + '\n' + 'Título do TCC: ' + titulo;

  DriveApp.createFile('Teste_Thesia.txt', conteudo);

  return 'Arquivo criado com sucesso!';
}

//função que gera cópia do modelo de doc
function gerarDocumentoTeste(professor, titulo) {

  //id do arquivo de documento modelo a ser preenchido
  const arquivoModelo = DriveApp.getFileById(CONFIG.MODELO_DOC_ID);

  //faz uma cópia e altera o título
  const copia = arquivoModelo.makeCopy(titulo + ' - ' + professor);

  //abre o doc por id da cópia
  const documento = DocumentApp.openById(copia.getId());

  //pega o corpo do texto para editar
  const corpo = documento.getBody();

  //substitui os dados digitados no documento
  corpo.replaceText('<<PROFESSOR>>', professor);

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

  const aba = planilha.getSheetByName('TESTE');

  //calcula novamente para jogar na planilha
  const final1 = (dados.artigo * 0.3) + (dados.relatorio * 0.3) + (dados.oral1 * 0.4);
  const final2 = (dados.artigo * 0.3) + (dados.relatorio * 0.3) + (dados.oral2 * 0.4);
  const final3 = (dados.artigo * 0.3) + (dados.relatorio * 0.3) + (dados.oral3 * 0.4);
  const final4 = (dados.artigo * 0.3) + (dados.relatorio * 0.3) + (dados.oral4 * 0.4);

  //adiciona uma linha com as informações preenchidas na planilha
  aba.appendRow([
    new Date(), dados.professor, dados.titulo, dados.artigo, dados.relatorio, dados.aluno1, 
    dados.oral1, final1, dados.aluno2, dados.oral2, final2, dados.aluno3, dados.oral3, final3, 
    dados.aluno4, dados.oral4, final4
  ]);

  return {
    aluno1: dados.aluno1, final1: final1, aluno2: dados.aluno2, final2: final2, aluno3: dados.aluno3, final3: final3, aluno4: dados.aluno4, final4: final4
  };
}