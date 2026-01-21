function Insere7lin() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('320:326').activate();
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveRange().getLastRow(), 7);
  spreadsheet.getActiveRange().offset(spreadsheet.getActiveRange().getNumRows(), 0, 7, spreadsheet.getActiveRange().getNumColumns()).activate();
};

function Ins6Lin() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(spreadsheet.getCurrentCell().getRow() - 3, 1, 6, sheet.getMaxColumns()).activate();
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveRange().getLastRow(), 6);
  spreadsheet.getActiveRange().offset(spreadsheet.getActiveRange().getNumRows(), 0, 6, spreadsheet.getActiveRange().getNumColumns()).activate();
};

function removerValor() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('K27').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};

function Exportapag() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AJ19').activate();
};

function BaseAnanf() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('Z52:AF52').activate();
};