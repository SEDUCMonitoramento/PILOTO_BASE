//Lista de Presença Reunião de Pais
function ImpListPRES() {
  Retorna()
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('F:S').activate();
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('T:T').activate();
  spreadsheet.getRange('A2:U50').activate();
};


// Lista de Entrega de Uniformes
function ImpListUni() {
  Retorna();
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('T:U').activate();
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('V:V').activate();
  spreadsheet.getRange('A2:X50').activate();
};

// Lista de Kit Material Escolar
function ImpListUni() {
  Retorna()
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('T:X').activate();
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('Y:Y').activate();
  spreadsheet.getRange('A2:AA50').activate();
};

//function ImpListPERSO() {
  //var spreadsheet = SpreadsheetApp.getActive();
  //spreadsheet.getRange('T:AD').activate();
  //spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  //spreadsheet.getRange('AE:AE').activate();
  //spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  //spreadsheet.getRange('AF:AF').activate();
  //spreadsheet.getRange('A7:50').activate();
//};


//Lista de Contatos do Alunos
function ImpListCont() {
  Retorna()
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('T:AA').activate();
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('AB:AB').activate();
  spreadsheet.getRange('A2:AD50').activate();
};

function InterImpress() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A7:AA70').activate();
};

function Retorna() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('T:T').activate();
  spreadsheet.getActiveSheet().showColumns(20, 13);
  spreadsheet.getRange('T1').activate();
  SpreadsheetApp.flush();

};

function proteSALA() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('1:6').activate();
  var protection = spreadsheet.getRange('1:6').protect();
};