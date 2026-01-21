// // Função para retornar em fórmula o nome da Aba aberta
// function nomeAba() {
//  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
//  }

// // function atualizar_All() {
// //   let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All");
// //   ss.getRange("A2").clearContent();
// //   // Aguarda 2 segundos
// //   Utilities.sleep(5000);
// //   // Insere a fórmula na célula A2
// //   ss.getRange("A2").setFormula("=ListaTodos()");
// // }


// function listaTodos(){
//   function getTurmas(){
//     let qntTurmas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Piloto").getRange("C3").getValue();
//     let turmas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Piloto").getRange("C4:C"+(3+qntTurmas)).getValues();
//     let aux = [];
//     for (let i=0;i<turmas.length;i++){
//       aux.push(turmas[i][0])
//     }
//     return aux
//   }
//   let turmas = getTurmas();
//   let todos = [];

//   for (let i=0;i<turmas.length;i++){
//     let alunos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(turmas[i]).getRange("A7:EB70").getValues();
//     //todos.push(alunos)
//     for(let j=0;j<alunos.length;j++){
//       if (alunos[j][0] != ""){
//         alunos[j][1] = turmas[i]
//         todos.push(alunos[j])
//       }
//     }
//   }
//   return (todos);
// }




