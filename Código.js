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
  PASTA_FICHAS_ID: propriedades.getProperty('PASTA_FICHAS_ID'),
  PASTA_PDFS_ID: propriedades.getProperty('PASTA_PDFS_ID'),

  MODELO_DOC_1_ID: propriedades.getProperty('MODELO_DOC_1_ID'),
  MODELO_DOC_2_ID: propriedades.getProperty('MODELO_DOC_2_ID'),
  MODELO_DOC_3_ID: propriedades.getProperty('MODELO_DOC_3_ID'),
  MODELO_DOC_4_ID: propriedades.getProperty('MODELO_DOC_4_ID'),
  MODELO_DOC_5_ID: propriedades.getProperty('MODELO_DOC_5_ID')
}


//função para obter data e hora do instante da avaliação
function obterDataHoraAvaliacao() 
{
  //data/hora para os registros
  const dataHora = new Date();
  //define fuso horário da sessão
  const fusoHorario = Session.getScriptTimeZone();
  //separa a data
  const data = Utilities.formatDate(dataHora, fusoHorario, "dd/MM/yyyy");
  //separa a hora
  const hora = Utilities.formatDate(dataHora, fusoHorario, "HH:mm");
  //separa o dia
  const dia = Utilities.formatDate(dataHora, fusoHorario, "dd");
  //separa o ano
  const ano = Utilities.formatDate(dataHora, fusoHorario, "yyyy");

  //escrever meses por extenso
  const meses = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];
  //separa o mes
  const numeroMes = Number(Utilities.formatDate(dataHora, fusoHorario, "MM"));
  //pega o nome do mes de acordo com o numeroMes (-1 pois arrays começam em 0)
  const mes = meses[numeroMes - 1];

  return {
    dataHora: dataHora, data: data, hora: hora, dia: dia, mes: mes, ano: ano
  };
}


//função para exibir as notas com virgula ao invés de ponto na ficha gerada
function formatarNota(valor) 
{
  return Number(valor.toFixed(2)).toString().replace(".", ",");
}


//função que cria uma cópia da ficha modelo para preencher
function gerarFichaAvaliacao(dados, calculos, momento) 
{
  //pega o nome do arquivo modelo baseado na quantidade de alunos
  const chaveModelo = "MODELO_DOC_" + dados.quantidadeAlunos + "_ID";
  const modeloId = CONFIG[chaveModelo];

  Logger.log(chaveModelo);
  Logger.log(modeloId);
  
  //verifica se encontrou o modelo
  if (!modeloId)
  {
    throw new Error("Modelo não encontrado para " + dados.quantidadeAlunos + " aluno(s).");
  }

  //pega o ID do modelo
  const modelo = DriveApp.getFileById(modeloId);


  //pega a pasta onde as fichas serão salvas
  const pastaFichas = DriveApp.getFolderById(CONFIG.PASTA_FICHAS_ID);

  //muda o nome do arquivo copiado
  const nomeArquivo = "Ficha - " + dados.titulo;

  //faz a copia do modelo dentro da pasta de fichas
  const copia = modelo.makeCopy(nomeArquivo, pastaFichas);

  //abre a cópia como Google Docs
  const documento = DocumentApp.openById(copia.getId());

  //pega o corpo do documento (texto)
  const corpo = documento.getBody();


  //PREENCHIMEMTO DA FICHA
  //data e hora
  corpo.replaceText("<<DATA>>", momento.data);
  corpo.replaceText("<<HORA>>", momento.hora);
  corpo.replaceText("<<DIA>>", momento.dia);
  corpo.replaceText("<<MES>>", momento.mes);
  corpo.replaceText("<<ANO>>", momento.ano);

  //dados do curso e tcc
  corpo.replaceText("<<TITULO>>", dados.titulo);
  corpo.replaceText("<<ORIENTADOR>>", dados.orientador);
  corpo.replaceText("<<CURSO>>", dados.curso);

  //dados dos alunos
  for (let i = 1; i <= dados.quantidadeAlunos; i++)
  {
    corpo.replaceText("<<NOME" + i + ">>", dados["aluno" + i]);
    corpo.replaceText("<<RA" + i + ">>", dados["ra" + i]);
  }

  //notas formatadas de artigo cientifico
  corpo.replaceText("<<ART1>>", formatarNota(dados.artigo1));
  corpo.replaceText("<<ART2>>", formatarNota(dados.artigo2));
  corpo.replaceText("<<ART3>>", formatarNota(dados.artigo3));
  corpo.replaceText("<<ART4>>", formatarNota(dados.artigo4));
  corpo.replaceText("<<ART5>>", formatarNota(dados.artigo5));
  corpo.replaceText("<<ART6>>", formatarNota(dados.artigo6));
  corpo.replaceText("<<ART7>>", formatarNota(dados.artigo7));
  corpo.replaceText("<<TOTART>>", formatarNota(calculos.totalArtigo * calculos.pesoArtigo));

  //notas formatadas de relatório técnico
  corpo.replaceText("<<REL1>>", formatarNota(dados.relatorio1));
  corpo.replaceText("<<REL2>>", formatarNota(dados.relatorio2));
  corpo.replaceText("<<REL3>>", formatarNota(dados.relatorio3));
  corpo.replaceText("<<REL4>>", formatarNota(dados.relatorio4));
  corpo.replaceText("<<REL5>>", formatarNota(dados.relatorio5));
  corpo.replaceText("<<TOTREL>>", formatarNota(calculos.totalRelatorio * calculos.pesoRelatorio));

  //notas de apresentação oral
  for (let aluno = 1; aluno <= dados.quantidadeAlunos; aluno++)
  {
    for (let criterio = 1; criterio <= 5; criterio++)
    {
      corpo.replaceText("<<O" + criterio + "A" + aluno + ">>", formatarNota(dados["oral" + criterio + "aluno" + aluno]));
    }

    //total oral com peso 0,4
    corpo.replaceText("<<TOTO" + aluno + ">>", formatarNota(calculos["totalOral" + aluno] * calculos.pesoOral));
  }
  

  //notas finais
  for (let i = 1; i <= dados.quantidadeAlunos; i++)
  {
    corpo.replaceText("<<FIN" + i + ">>", formatarNota(calculos["final" + i]));
  }

  //salva e fecha o documento
  documento.saveAndClose();

  //puxa a pasta de fichas por id
  const pastaPdfs = DriveApp.getFolderById(CONFIG.PASTA_PDFS_ID);

  //puxa a copia que acabou de ser criada
  const arquivoGoogleDocs = DriveApp.getFileById(copia.getId());

  //converte a ficha para pdf
  const blobPdf = arquivoGoogleDocs.getAs(MimeType.PDF).setName(nomeArquivo + ".pdf");

  //salva na pasta de pdfs gerados
  const arquivoPdf = pastaPdfs.createFile(blobPdf);

  //devolve informações sobre a ficha
  return {
    idDocumento: copia.getId(),
    urlDocumento: copia.getUrl(),

    idPdf: arquivoPdf.getId(),
    urlPdf: arquivoPdf.getUrl(),

    nomeArquivo: nomeArquivo
  }
}

//função que salva a avaliação na planilha do drive
function salvarAvaliacao(dados)
{
  //abre a planilha por id
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);

  //cria um id unico para a avaliação atual
  const idAvaliacao = Utilities.getUuid();

  //seleciona as abas usadas
  const abaAvaliacoes = planilha.getSheetByName("AVALIACOES");
  const abaArtigo = planilha.getSheetByName("ARTIGO");
  const abaRelatorio = planilha.getSheetByName("RELATORIO");
  const abaOral = planilha.getSheetByName("ORAL");

  const momento = obterDataHoraAvaliacao();

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

  //quantidade de alunos para calculo das notas orais
  const quantidadeAlunos = Number(dados.quantidadeAlunos);

  //objeto que guarda todos os calculos
  const calculos = {pesoArtigo: pesoArtigo, pesoRelatorio: pesoRelatorio, pesoOral: pesoOral, totalArtigo: totalArtigo, totalRelatorio: totalRelatorio};

  //calcula as notas orais e finais apenas de alunos existentes
  for (let i = 1; i <= quantidadeAlunos; i++) 
  {
    const totalOral = dados["oral1aluno" + i] + dados["oral2aluno" + i] + dados["oral3aluno" +  i] + dados["oral4aluno" + i] + dados["oral5aluno" + i]; 

    const final = (totalArtigo * pesoArtigo) + (totalRelatorio * pesoRelatorio) + (totalOral * pesoOral);

    calculos["totalOral" + i] = totalOral;
    calculos["final" + i] = final;
  }



  //monta a linha para a aba AVALIACOES
  const linhaAvaliacao = [idAvaliacao, momento.dataHora, dados.orientador, dados.titulo, dados.curso];

  //adiciona os dados dos possiveis 5 alunos
  for (let i = 1; i <= 5; i++) 
  {
    if (i <= quantidadeAlunos)
    {
      linhaAvaliacao.push(dados["aluno"+ i], dados["ra" + i], calculos["final" + i]);
    }

    //campos de alunos que não existem ficam vazios
    else 
    {
      linhaAvaliacao.push("", "", "");
    }
  }

  //salva a linha completa
  abaAvaliacoes.appendRow(linhaAvaliacao);
  

  // salva os critérios do artigo na aba ARTIGO
  abaArtigo.appendRow([
    idAvaliacao, dados.artigo1, dados.artigo2, dados.artigo3, dados.artigo4, dados.artigo5, dados.artigo6, dados.artigo7, totalArtigo
  ]);

  //salva os critérios de relatório na aba RELATORIO
  abaRelatorio.appendRow([
    idAvaliacao, dados.relatorio1, dados.relatorio2, dados.relatorio3, dados.relatorio4, dados.relatorio5, totalRelatorio
  ]);


  //salva os critérios de oratória somente dos alunos existentes
  for (let i = 1; i <= quantidadeAlunos; i++)
  {
    abaOral.appendRow([
      idAvaliacao, i, dados["aluno" + i], dados["oral1aluno" + i], dados["oral2aluno" + i], dados["oral3aluno" + i], dados["oral4aluno" + i], dados["oral5aluno" + i], calculos["totalOral" + i]
    ]);
  }

  //chama a função para gerar a ficha com os dados preenchidos
  const ficha = gerarFichaAvaliacao(dados, calculos, momento);

  //chama a função que envia a ficha por email
  enviarFichaEmail(dados, ficha, momento);

  //devolve os resultados para o html
  const resultado = {
    idAvaliacao: idAvaliacao,
    totalArtigo: totalArtigo,
    totalRelatorio: totalRelatorio,
    quantidadeAlunos: quantidadeAlunos
  };

  //retorna conforme a quantidade de alunos
  for (let i = 1; i <= quantidadeAlunos; i++)
  {
    resultado["aluno" + i] = dados["aluno" + i];
    resultado["totalOral" + i] = calculos["totalOral" + i];
    resultado["final" + i] = calculos["final" + i];
  }

  return resultado;
}

//função que envia a ficha em pdf por email
function enviarFichaEmail(dados, ficha, momento) 
{
  //pega o pdf gerado
  const arquivoPdf = DriveApp.getFileById(ficha.idPdf);

  //assunto do email
  const assunto = "Thesia - Ficha de Avaliação - " + dados.titulo;

  //texto do email
  const mensagem = "Olá, \n\n" + "Sua avaliação referente ao TCC \"" + dados.titulo + "\" foi registrada com sucesso.\n\n" +
  "Data: " + momento.data + "\n" + "Horário: " + momento.hora + "\n" + "Curso: " + dados.curso + "\n" + "Orientador(a): " + 
  dados.orientador + "\n\n" + "A ficha de avaliação em PDF segue anexada a este e-mail.\n\n" + "Atenciosamente,\n" + "Thesia";

  GmailApp.sendEmail(dados.email, assunto, mensagem,
    {
      attachments: [arquivoPdf.getBlob()], name: "Thesia"
    }
  );
}