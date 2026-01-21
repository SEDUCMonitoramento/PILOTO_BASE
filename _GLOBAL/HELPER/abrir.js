function abrir_aba_ativando(aba = "IMP_List") {
    var spreadsheet = SpreadsheetApp.getActive();
    var sheet = spreadsheet.getSheetByName(aba);
    if (sheet) {
        sheet.activate();
        SpreadsheetApp.flush();
    } else {
        SpreadsheetApp.getUi().alert("Aba '" + aba + "' não encontrada!");
    }
}