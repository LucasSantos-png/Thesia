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
  MODELO_DOC_ID: propriedades.getProperty('MODELO_DOC_ID'),
  PASTA_FICHAS_ID: propriedades.getProperty('PASTA_FICHAS_ID'),
  //PASTA_PDFS_ID: propriedades.getProperty('PASTA_PDFS_ID')
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
  //pega o arquivo modelo
  const modelo = DriveApp.getFileById(CONFIG.MODELO_DOC_ID);

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
  corpo.replaceText("<<NOME1>>", dados.aluno1);
  corpo.replaceText("<<RA1>>",dados.ra1);

  corpo.replaceText("<<NOME2>>", dados.aluno2);
  corpo.replaceText("<<RA2>>",dados.ra2);

  corpo.replaceText("<<NOME3>>", dados.aluno3);
  corpo.replaceText("<<RA3>>",dados.ra3);

  corpo.replaceText("<<NOME4>>", dados.aluno4);
  corpo.replaceText("<<RA4>>",dados.ra4);

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

  //notas formatadas de oratória
  corpo.replaceText("<<O1A1>>", formatarNota(dados.oral1aluno1));
  corpo.replaceText("<<O2A1>>", formatarNota(dados.oral2aluno1));
  corpo.replaceText("<<O3A1>>", formatarNota(dados.oral3aluno1));
  corpo.replaceText("<<O4A1>>", formatarNota(dados.oral4aluno1));
  corpo.replaceText("<<O5A1>>", formatarNota(dados.oral5aluno1));
  corpo.replaceText("<<TOTO1>>", formatarNota(calculos.totalOral1 * calculos.pesoOral));

  corpo.replaceText("<<O1A2>>", formatarNota(dados.oral1aluno2));
  corpo.replaceText("<<O2A2>>", formatarNota(dados.oral2aluno2));
  corpo.replaceText("<<O3A2>>", formatarNota(dados.oral3aluno2));
  corpo.replaceText("<<O4A2>>", formatarNota(dados.oral4aluno2));
  corpo.replaceText("<<O5A2>>", formatarNota(dados.oral5aluno2));
  corpo.replaceText("<<TOTO2>>", formatarNota(calculos.totalOral2 * calculos.pesoOral));

  corpo.replaceText("<<O1A3>>", formatarNota(dados.oral1aluno3));
  corpo.replaceText("<<O2A3>>", formatarNota(dados.oral2aluno3));
  corpo.replaceText("<<O3A3>>", formatarNota(dados.oral3aluno3));
  corpo.replaceText("<<O4A3>>", formatarNota(dados.oral4aluno3));
  corpo.replaceText("<<O5A3>>", formatarNota(dados.oral5aluno3));
  corpo.replaceText("<<TOTO3>>", formatarNota(calculos.totalOral3 * calculos.pesoOral));

  corpo.replaceText("<<O1A4>>", formatarNota(dados.oral1aluno4));    
  corpo.replaceText("<<O2A4>>", formatarNota(dados.oral2aluno4));    
  corpo.replaceText("<<O3A4>>", formatarNota(dados.oral3aluno4));    
  corpo.replaceText("<<O4A4>>", formatarNota(dados.oral4aluno4));
  corpo.replaceText("<<O5A4>>", formatarNota(dados.oral5aluno4));
  corpo.replaceText("<<TOTO4>>", formatarNota(calculos.totalOral4 * calculos.pesoOral));

  //notas finais
  corpo.replaceText("<<FIN1>>", formatarNota(calculos.final1));
  corpo.replaceText("<<FIN2>>", formatarNota(calculos.final2));
  corpo.replaceText("<<FIN3>>", formatarNota(calculos.final3));
  corpo.replaceText("<<FIN4>>", formatarNota(calculos.final4));

  //salva e fecha o documento
  documento.saveAndClose();

  //devolve informações sobre a ficha
  return {
    idDocumento: copia.getId(),
    urlDocumento: copia.getUrl(),
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

  const calculos = {
    pesoArtigo: pesoArtigo, pesoRelatorio: pesoRelatorio, pesoOral: pesoOral, totalArtigo: totalArtigo, totalRelatorio: totalRelatorio,
    totalOral1: totalOral1, totalOral2: totalOral2, totalOral3: totalOral3, totalOral4: totalOral4,

    final1: final1, final2: final2, final3: final3, final4: final4
  }

  //salva o resumo na aba AVALIACOES
  abaAvaliacoes.appendRow([
    idAvaliacao, momento.dataHora, dados.orientador, dados.titulo, dados.curso, dados.aluno1, dados.ra1, final1, dados.aluno2, dados.ra2, final2, dados.aluno3, dados.ra3, final3, dados.aluno4, dados.ra4, final4
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

  //chama a função para gerar a ficha com os dados preenchidos
  const ficha = gerarFichaAvaliacao(dados, calculos, momento);

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