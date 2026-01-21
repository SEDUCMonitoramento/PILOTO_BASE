function exportarRangeComoNovaPlanilha() {
  // Obtém a planilha ativa
  const planilhaAtiva = SpreadsheetApp.getActiveSpreadsheet();
  const abaAtiva = planilhaAtiva.getActiveSheet();
  
  // Define o range A1:AG55
  const range = abaAtiva.getRange("A1:AG55");
  
  // Gera nome da nova planilha com data/hora
  const dataHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
  const nomeNovaPlanilha = `Exportação_${dataHora}`;
  
  // Cria uma nova planilha
  const novaPlanilha = SpreadsheetApp.create(nomeNovaPlanilha);
  const novaAba = novaPlanilha.getActiveSheet();
  
  // Copia valores (convertendo tudo para texto)
  copiarValoresComoTexto(range, novaAba);
  
  // Copia formatação
  copiarFormatacao(range, novaAba);
  
  // Ajusta largura das colunas
  ajustarLarguraColunas(range, novaAba);
  
  // Mostra link para a nova planilha
  const url = novaPlanilha.getUrl();
  SpreadsheetApp.getUi().alert(
    'Exportação Concluída!',
    `Nova planilha criada: ${nomeNovaPlanilha}\n\nAcesse em: ${url}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  // Abre a nova planilha em nova aba (opcional)
  mostrarLinkUrl(url);
}

function copiarValoresComoTexto(rangeOrigem, abaDestino) {
  const valores = rangeOrigem.getValues();
  const valoresComoTexto = valores.map(linha => 
    linha.map(valor => {
      if (valor instanceof Date) {
        // Converte datas para string no formato local
        return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      } else if (typeof valor === 'number' || typeof valor === 'boolean') {
        // Converte números e booleanos para string
        return valor.toString();
      } else if (valor === null || valor === undefined) {
        // Mantém células vazias
        return '';
      }
      // Já é string
      return String(valor);
    })
  );
  
  abaDestino.getRange(1, 1, valoresComoTexto.length, valoresComoTexto[0].length)
    .setValues(valoresComoTexto);
}

function copiarFormatacao(rangeOrigem, abaDestino) {
  const rangeDestino = abaDestino.getRange(1, 1, rangeOrigem.getNumRows(), rangeOrigem.getNumColumns());
  
  // Copia formatação básica
  rangeDestino.setNumberFormat("@"); // Define tudo como texto
  rangeDestino.setBackgrounds(rangeOrigem.getBackgrounds());
  rangeDestino.setFontColors(rangeOrigem.getFontColors());
  rangeDestino.setFontFamilies(rangeOrigem.getFontFamilies());
  rangeDestino.setFontSizes(rangeOrigem.getFontSizes());
  rangeDestino.setFontWeights(rangeOrigem.getFontWeights());
  rangeDestino.setFontStyles(rangeOrigem.getFontStyles());
  rangeDestino.setHorizontalAlignments(rangeOrigem.getHorizontalAlignments());
  rangeDestino.setVerticalAlignments(rangeOrigem.getVerticalAlignments());
  rangeDestino.setWraps(rangeOrigem.getWraps());
  rangeDestino.setTextDirections(rangeOrigem.getTextDirections());
  
  // Copia bordas
  rangeDestino.setBorder(
    true,
    true,
    true,
    true,
    true,
    true,
    rangeOrigem.getBorderColor(),
    rangeOrigem.getBorderStyle()
  );
}

function ajustarLarguraColunas(rangeOrigem, abaDestino) {
  const numColunas = rangeOrigem.getNumColumns();
  
  for (let i = 1; i <= numColunas; i++) {
    const largura = rangeOrigem.getSheet().getColumnWidth(i);
    abaDestino.setColumnWidth(i, largura);
  }
}

function mostrarLinkUrl(url) {
  // Cria uma interface HTML para abrir o link
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_blank">
        <script>
          window.onload = function() {
            window.open("${url}", "_blank");
            google.script.host.close();
          };
        </script>
      </head>
      <body>
        <p>Abrindo nova planilha...</p>
      </body>
    </html>
  `;
  
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(300).setHeight(100),
    'Redirecionando'
  );
}

// Função para criar um menu personalizado (opcional)
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Exportação')
    .addItem('Exportar Range como Nova Planilha', 'exportarRangeComoNovaPlanilha')
    .addToUi();
}