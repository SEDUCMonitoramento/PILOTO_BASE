// function replicarAbaParaOutraPlanilha() {
//   console.log("Início")


//   // --- Configurações ---
//   let sheetPILOTO = SpreadsheetApp.getActiveSpreadsheet();
//   const pastaANANF = getOuCriaPastaANANF(sheetPILOTO);
//   const linkPastaANANF = pastaANANF.id_pasta_ananf;
//   const pastaANANFLink = "https://drive.google.com/drive/folders/" + linkPastaANANF;
//   console.log(pastaANANFLink)
//   const origemId = "1PX47xrTnfC5jacdSpaQUWPXd0a2HusxiyhcOb6CSrYk";
//   const nomeAbaOrigem = "Doc_Ananf";
//   const destinoId = "18bnKDT7oOs7Qb27CHKxCEWrwmSGVAit2g3JiUWgx3Tg";
//   const nomeModelo = "Doc_Ananf_Modelo";

//   const ssOrigem = SpreadsheetApp.openById(origemId);
//   const abaOrigem = ssOrigem.getSheetByName(nomeAbaOrigem);
//   const ssDestino = SpreadsheetApp.openById(destinoId);

//   return
//   // --- Verifica ou cria modelo ---
//   let modeloAba = ssDestino.getSheetByName(nomeModelo);
//   console.log(!modeloAba)
//   if (!modeloAba) {
//     console.log("Gerando aba")
//     const tempAba = abaOrigem.copyTo(ssOrigem).setName(nomeModelo + "_TEMP");
//     tempAba.getDataRange().setValues(tempAba.getDataRange().getValues());
//     modeloAba = tempAba.copyTo(ssDestino).setName(nomeModelo);
//     ssOrigem.deleteSheet(tempAba);
//     console.log("Criada a aba")
//   }

//   // --- Cria nova aba com nome = data/hora ---
//   const dataHoraFormatada = Utilities.formatDate(new Date(), "America/Sao_Paulo", "yyyy-MM-dd_HHmm");
//   const nomeNovaAba = dataHoraFormatada;
//   console.log("Insere nova aba")
//   const novaAba = ssDestino.insertSheet(nomeNovaAba);
//   console.log("Aba inserida")

//   // --- Copia dados (apenas valores) ---
//   console.log("Tempo 1")
//   const dadosOrigem = abaOrigem.getDataRange().getValues();
//   novaAba.getRange(1, 1, dadosOrigem.length, dadosOrigem[0].length).setValues(dadosOrigem);

//   // --- Garante tamanho da aba destino igual ao modelo ---
//   console.log("Tempo 2")
//   const maxRowsModelo = modeloAba.getMaxRows();
//   const maxColsModelo = modeloAba.getMaxColumns();
//   console.log("Linha " + maxRowsModelo)
//   console.log("Col " + maxColsModelo)


//   console.log("Tempo 3")

//   if (novaAba.getMaxRows() < maxRowsModelo) {
//     console.log("Entrou linha")
//     novaAba.insertRowsAfter(novaAba.getMaxRows(), maxRowsModelo - novaAba.getMaxRows());
//   }
//   if (novaAba.getMaxColumns() < maxColsModelo) {
//     console.log("Entrou coluna")
//     novaAba.insertColumnsAfter(novaAba.getMaxColumns(), maxColsModelo - novaAba.getMaxColumns());
//   }

//   // --- Copia formatação integral ---
//   console.log("Tempo 4")

//   modeloAba.getRange(1, 1, maxRowsModelo, maxColsModelo)
//     .copyTo(novaAba.getRange(1, 1, maxRowsModelo, maxColsModelo),
//       { formatOnly: true });

//   // --- Ajusta colunas e linhas para garantir tamanho idêntico ---
//   for (let c = 1; c <= maxColsModelo; c++) {
//     novaAba.setColumnWidth(c, modeloAba.getColumnWidth(c));
//   }
//   for (let r = 1; r <= maxRowsModelo; r++) {
//     novaAba.setRowHeight(r, modeloAba.getRowHeight(r));
//   }

//   Logger.log(`Aba replicada com sucesso: "${nomeNovaAba}"`);
// }



// function getOuCriaPastaANANF(sheetPILOTO) {

//   // Pega a pasta onde a planilha está
//   const pastaPai = getPastaDaPlanilha(sheetPILOTO);

//   if (!pastaPai) {
//     throw new Error("A planilha não está em nenhuma pasta.");
//   }

//   const idPastaPai = pastaPai.getId();
//   const nomePasta = "ANANF";

//   // Procura a pasta ANANF dentro da pasta pai
//   const pastas = pastaPai.getFoldersByName(nomePasta);

//   let pastaANANF;

//   if (pastas.hasNext()) {
//     // ✅ Pasta já existe
//     pastaANANF = pastas.next();
//     console.log("Pasta ANANF encontrada: " + pastaANANF.getName());
//   } else {
//     // 🆕 Cria a pasta
//     pastaANANF = pastaPai.createFolder(nomePasta);
//     console.log("Pasta ANANF criada: " + pastaANANF.getName());
//   }

//   return {
//     id_pasta_pai: idPastaPai,
//     id_pasta_ananf: pastaANANF.getId()
//   };
// }


// function getPastaDaPlanilha(sheetPILOTO) {
//   const fileId = sheetPILOTO.getId();
//   const file = DriveApp.getFileById(fileId);
//   const parents = file.getParents();

//   if (parents.hasNext()) {
//     return parents.next(); // retorna a pasta
//   }

//   return null;
// }
