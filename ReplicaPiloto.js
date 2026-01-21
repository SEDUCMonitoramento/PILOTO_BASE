/**
 * replicarPilotos() main
 * Replica a Piloto pegando os nomes e dados da escola da base idModelo
 */

const config = () => {
  return {
    // https://docs.google.com/spreadsheets/d/1VzFvTCfElAGSEUk8rOWSYwkkhYqsWZzWUoqtAXQHTHQ/edit?gid=1123067777#gid=1123067777
    id_db2026: "1VzFvTCfElAGSEUk8rOWSYwkkhYqsWZzWUoqtAXQHTHQ",

    // https://drive.google.com/drive/folders/1Cb5VNBoxUp33xVcs80D0G_CG67R_kx8S
    /** 2026 > 2026 MONITORAMENTO > 2026 Pilotos_UEs */
    pasta: "1Cb5VNBoxUp33xVcs80D0G_CG67R_kx8S"

  }
}
// console.log(config().id_db2026)

function replicarPilotos() {
  //----> 1. Vai copiar a Base como modelo para as demais 
  let idModelo = obterBASE();
  //----> 2. Lista das escolas com o número, cod escola e o nome completo 
  let escolas = listarESCOLAS();
  let url_pilotos = [];
  let linkPilotos = SpreadsheetApp.openById(config().id_db2026).getSheetByName("LinkPilotos");
  let row = 1;

  //----> 3. Executa a replica e retorna o URL
  escolas.forEach(function (escola) {
    console.log("Iniciando a escola: " + escola[0])
    let nome_da_PILOTO = escola[1] + " - PILOTO: " + escola[2];
    url_pilotos = replicaINDIVIDUAL(nome_da_PILOTO, idModelo, escola[2]);

    //--> 3.1 Escreve os dados de url e id em LinksPilotos
    linkPilotos.getRange("A" + row).setValue(escola[0]); // cod ESCOLA
    linkPilotos.getRange("B" + row).setValue(escola[2]); // NOME
    linkPilotos.getRange("C" + row).setValue(url_pilotos.split("/")[5]); // id
    linkPilotos.getRange("D" + row).setValue(url_pilotos); // URL
    row++; // Incremento a linha para a próxima planilha
    console.log("Finalizou a escola: " + escola[0]);

  });
  return
}

function obterBASE() {
  console.log("Copiando a base...")
  //ID da Piloto Base 2025 link acima
  let idModelo = config().id_db2026;
  console.log("Base copiada com sucesso!")
  return idModelo;
}

function listarESCOLAS() {
  //0: 01
  //1: 2501
  //2: EM SEMENTES DO BEM

  let DADOS_DAS_ESCOLAS = SpreadsheetApp.openById(config().id_db2026)
    .getSheetByName("DADOS DAS ESCOLAS").getRange("A2:C33").getValues();
  escolas = []
  for (let i = 0; i < DADOS_DAS_ESCOLAS.length; i++) {
    if (DADOS_DAS_ESCOLAS[i][1] != "") {
      escolas.push(DADOS_DAS_ESCOLAS[i])
    }
  }
  return escolas
}


function replicaINDIVIDUAL(nome_da_PILOTO, idModelo, nomeEscola) {
  let pilotos_UEs_2026 = DriveApp.getFolderById(config().pasta)
  let arquivo
  let arquivos = pilotos_UEs_2026.getFilesByName(nome_da_PILOTO);
  let arquivoEncontrado = false

  while (arquivos.hasNext()) {
    arquivo = arquivos.next();
    if (arquivo.getName() === nome_da_PILOTO) {
      arquivoEncontrado = true;
      break; // Sai do loop se o arquivo for encontrado
    }
  }

  if (!arquivoEncontrado) {
    let modelo = DriveApp.getFileById(idModelo);
    arquivo = modelo.makeCopy(nome_da_PILOTO, pilotos_UEs_2026);
    mudar_nomeEscola(arquivo.getUrl(), nomeEscola);
  }
  return (arquivo.getUrl());
}

function mudar_nomeEscola(url, nomeEscola) {
  SpreadsheetApp.openByUrl(url).getSheetByName("Piloto").getRange("B1").setValue(nomeEscola);
}


