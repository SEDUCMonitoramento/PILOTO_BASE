/**
 * Excluir linkTURMAS após replicar Pilotos
 */

// Variável global para acessar a planilha ativa
let ss = SpreadsheetApp.getActiveSpreadsheet();

function criarTurmas() {
  //1. Limpar e o OcultarLinhas da página Piloto;
  limparEOcultarLinhas();
  // 2. Obtém a lista de turmas da aba "Piloto"
  let turmas = getTurmas();
  // 3. Cria cada turma
  for (let i = 0; i < turmas.length; i++) {
    Logger.log(`Fazendo a turma ${turmas[i]}`)
    criarTurma(turmas[i]);
  }
  // 4. Remove proteções das abas existentes antes de recriá-las
  //tirarProtecao();
  // 5. Fazer a ALL
  escreverFormula_QUERY_das_ALL();
}

function criarTurma(nturma) {
  let base = ss.getSheetByName("Base");
  let turma = ss.getSheetByName(nturma)
  if (!turma){
    // Criar uma cópia da aba "Base"
    turma = base.copyTo(ss);
    turma.setName(nturma).showSheet();
    turma.getRange("A5").setValue(nturma);
  }
  // Proteger a aba
  let protection = turma.protect();
  protection.removeEditors(protection.getEditors()); // Remove permissões padrão
  protection.addEditor(Session.getActiveUser().getEmail()); // Adiciona o usuário ativo

  // Proteger intervalos específicos
  let protectedRanges = [turma.getRange("1:6"), turma.getRange("AB7:AW70"), turma.getRange("AY7:AY70"), turma.getRange("BC7:EB70"), turma.getRange("S7:S70"), turma.getRange("T7:T70"), turma.getRange("U7:U70"), turma.getRange("V7:V70") ];

  protectedRanges.forEach(range => {
    let rangeProtection = range.protect();
    rangeProtection.addEditor(Session.getActiveUser().getEmail());
  });
  
}

function getTurmas() {
  // Obtém os nomes das turmas na coluna "C" da aba "Piloto"
  let sheetPiloto = ss.getSheetByName("Piloto");
  let turmas = [];
  for (let i = 4; i <= 39; i++) {
    let value = sheetPiloto.getRange("C" + i).getValue();
    if (value) {
      turmas.push(value);
    }
  }
  return turmas;
}

function tirarProtecao() {
  // Remove proteções das abas listadas na coluna "C" da aba "Piloto"
  let sheetPiloto = ss.getSheetByName("Piloto");
  let nomes = sheetPiloto.getRange("C4:C40").getValues().flat(); // Lista de turmas
  nomes.forEach(nome => {
    let aba = ss.getSheetByName(nome);
    if (aba) {
      let protecoes = aba.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protecoes.forEach(protecao => protecao.remove());
    }
  });
}
