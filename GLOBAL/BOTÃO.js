function onOpen(e) {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Atualizar All")
    .addItem("Atualizar All", "escreverFormula_QUERY_das_ALL")
    .addToUi();

  let ui2 = SpreadsheetApp.getUi();
  let o = ui2.createMenu("Piloto")
  o.addItem("Ir para Inicial", "inicial_ir")
  o.addItem("Ir para Piloto", "piloto_ir")

  let salaSelecionada = ""; // Variável global para armazenar a sala selecionada

  // Criando o submenu
  // let salasMenu = ui2.createMenu("Salas");
  // SpreadsheetApp.getUi().alert(salas)

  // salas.forEach((sala) => {
  //   salasMenu.addItem(sala, "irPara")
  // });

  //o.addSubMenu(salasMenu)


  o.addToUi();


  let ui3 = SpreadsheetApp.getUi();
  let submenu3 = ui3.createMenu("Imprimir Listas")
  submenu3.addItem("Lista de presença", "ImpListPRES")
  submenu3.addItem("Lista entrega de uniformes", "ImpListUni")
  submenu3.addItem("Lista entrega de Kit Escola", "ImpListUni")
  submenu3.addItem("Lista contatos dos alunos", "ImpListCont")
  submenu3.addItem("LIMPAR SELEÇÃO", "Retorna")

  submenu3.addToUi();

  let ui4 = SpreadsheetApp.getUi();
  let submenu4 = ui4.createMenu("ANANF")
  submenu4.addItem("Gerar ANANF", "replicarAbaParaOutraPlanilha")
  submenu4.addItem("Ver ANANFs", "abrirSidebarANANF")
  submenu4.addToUi();


}
// const salas = obterSalas()
// var count = 0;

// function irParaSala() {
//   irPara("P2A"); // Chama a função irPara com o nome da sala
//   return ""
// }

// function irPara() {
//   let ss = SpreadsheetApp.getActiveSpreadsheet();
//   let abaPiloto = ss.getSheetByName(salas[count]);
//   console.log("Teste" + salaSelecionada)
//   console.log(count)
//   ss.setActiveSheet(abaPiloto);
// }

function piloto_ir() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let abaPiloto = ss.getSheetByName("Piloto");
  ss.setActiveSheet(abaPiloto);

}
function inicial_ir() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let abaPiloto = ss.getSheetByName("Inicial");
  ss.setActiveSheet(abaPiloto);

}

function obterSalas() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetPiloto = ss.getSheetByName("Piloto");
  let salas = [];
  for (let i = 4; i <= 39; i++) {
    let value = sheetPiloto.getRange("C" + i).getValue();
    if (value) {
      salas.push(value);
    }
  }
  console.log(salas)
  //return ['P2A']
  return salas;
}


