/**
 * Script otimizado para criação e proteção de abas de turmas usando Sheets API v4
 * 
 * Otimizações:
 * - Duplicação em lote (batch) de múltiplas abas
 * - Configuração em lote de propriedades
 * - Proteções em lote com unprotectedRanges
 * - Redução drástica de chamadas à API
 * 
 * Áreas liberadas para edição:
 * - A7:R70 (dados principais)
 * - W7:X70
 * - AA7:AA70
 * - AX7:AX70
 * - AZ7:BB70
 */

/**
 * Função principal para criar todas as turmas de forma otimizada
 */
function criarTurmas() {
  const inicioScript = new Date();
  Logger.log('═══════════════════════════════════════════════════════');
  Logger.log('🚀 Iniciando criação de turmas...');

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const planilhaId = planilha.getId();
  const nomeAbaBase = 'Base';

  // 1. Limpar e ocultar linhas da página Piloto
  const inicioLimpeza = new Date();
  limparEOcultarLinhas();
  Logger.log(`✓ Linhas limpas e ocultadas (${new Date() - inicioLimpeza}ms)`);

  // 2. Verificação de segurança
  const mapaAbas = construirMapaAbas(planilha);
  if (!mapaAbas[nomeAbaBase]) {
    Logger.log(`❌ Erro: Aba '${nomeAbaBase}' não encontrada!`);
    SpreadsheetApp.getUi().alert(`Aba '${nomeAbaBase}' não encontrada!`);
    return;
  }

  // 3. Mostrar aba Base temporariamente
  const abaBaseVisivel = mostrarAbaBase(planilha, nomeAbaBase);

  // 4. Identificar turmas a criar
  const todasTurmas = obterListaTurmas(planilha);
  const turmasParaCriar = todasTurmas.filter(nomeTurma => !mapaAbas[nomeTurma]);

  Logger.log(`📊 Total de turmas: ${todasTurmas.length}`);
  Logger.log(`📋 Turmas existentes: ${todasTurmas.length - turmasParaCriar.length}`);
  Logger.log(`➕ Turmas a criar: ${turmasParaCriar.length}`);

  if (!turmasParaCriar.length) {
    Logger.log('ℹ️ Nenhuma turma nova para criar');
    escreverFormula_QUERY_das_ALL();
    ocultarAbaBase(planilha, nomeAbaBase, abaBaseVisivel);
    Logger.log('═══════════════════════════════════════════════════════');
    return;
  }

  Logger.log(`🔄 Criando ${turmasParaCriar.length} turma(s): ${turmasParaCriar.join(', ')}`);

  // 5. Duplicação em lote
  const inicioDuplicacao = new Date();
  const abaBaseId = mapaAbas[nomeAbaBase].getSheetId();
  const indiceInsercao = obterIndiceInsercao(planilha, nomeAbaBase);
  const idsNovasAbas = duplicarAbasEmLote(planilhaId, abaBaseId, turmasParaCriar, indiceInsercao);
  Logger.log(`✓ ${turmasParaCriar.length} aba(s) duplicada(s) (${new Date() - inicioDuplicacao}ms)`);

  // 6. Configuração em lote (propriedades + proteções)
  const inicioConfiguracao = new Date();
  configurarAbasEmLote(planilhaId, idsNovasAbas);
  Logger.log(`✓ Proteções aplicadas em ${turmasParaCriar.length} aba(s) (${new Date() - inicioConfiguracao}ms)`);

  // 7. Escrever nomes das turmas em lote
  const inicioEscrita = new Date();
  escreverNomesEmLote(planilhaId, turmasParaCriar);
  Logger.log(`✓ Nomes escritos em ${turmasParaCriar.length} aba(s) (${new Date() - inicioEscrita}ms)`);

  // 8. Ocultar aba Base novamente
  ocultarAbaBase(planilha, nomeAbaBase, abaBaseVisivel);

  // 9. Fazer a ALL
  const inicioQuery = new Date();
  escreverFormula_QUERY_das_ALL();
  Logger.log(`✓ Fórmula QUERY da aba ALL atualizada (${new Date() - inicioQuery}ms)`);

  const tempoTotal = new Date() - inicioScript;
  Logger.log('═══════════════════════════════════════════════════════');
  Logger.log(`✅ ${turmasParaCriar.length} turma(s) criada(s) com sucesso!`);
  Logger.log(`⏱️ Tempo total: ${(tempoTotal / 1000).toFixed(1)}s (${tempoTotal}ms)`);
  Logger.log('═══════════════════════════════════════════════════════');
}

/**
 * Constrói um mapa de abas para verificação rápida
 * @param {Spreadsheet} planilha - A planilha ativa
 * @returns {Object} Mapa nome -> Sheet
 */
function construirMapaAbas(planilha) {
  const mapa = {};
  planilha.getSheets().forEach(aba => {
    mapa[aba.getName().trim()] = aba;
  });
  return mapa;
}

/**
 * Obtém a lista de nomes de turmas da aba "Piloto"
 * @param {Spreadsheet} planilha - A planilha ativa
 * @returns {Array<string>} Lista com os nomes das turmas
 */
function obterListaTurmas(planilha) {
  const abaPiloto = planilha.getSheetByName("Piloto");
  const valores = abaPiloto.getRange("C4:C39").getValues();

  return valores
    .flat()
    .filter(valor => valor && valor.toString().trim() !== '')
    .map(valor => valor.toString().trim());
}

/**
 * Obtém o índice de inserção das novas abas
 * @param {Spreadsheet} planilha - A planilha ativa
 * @param {string} nomeAbaBase - Nome da aba base
 * @returns {number} Índice onde inserir as novas abas
 */
function obterIndiceInsercao(planilha, nomeAbaBase) {
  const abas = planilha.getSheets();
  const indice = abas.findIndex(aba => aba.getName().trim() === nomeAbaBase);
  return indice < 0 ? abas.length : indice + 1;
}

/**
 * Duplica múltiplas abas em uma única chamada de API
 * @param {string} planilhaId - ID da planilha
 * @param {number} abaBaseId - ID da aba base
 * @param {Array<string>} nomesTurmas - Lista de nomes das turmas
 * @param {number} indiceInsercao - Índice de inserção
 * @returns {Array<number>} IDs das novas abas criadas
 */
function duplicarAbasEmLote(planilhaId, abaBaseId, nomesTurmas, indiceInsercao) {
  const requisicoesDuplicacao = nomesTurmas.map((nomeTurma, indice) => ({
    duplicateSheet: {
      sourceSheetId: abaBaseId,
      newSheetName: nomeTurma,
      insertSheetIndex: indiceInsercao + indice
    }
  }));

  const resposta = Sheets.Spreadsheets.batchUpdate({
    requests: requisicoesDuplicacao
  }, planilhaId);

  return resposta.replies.map(r => r.duplicateSheet.properties.sheetId);
}

/**
 * Configura propriedades e proteções das abas em lote
 * @param {string} planilhaId - ID da planilha
 * @param {Array<number>} idsAbas - IDs das abas a configurar
 */
function configurarAbasEmLote(planilhaId, idsAbas) {
  const emailUsuario = Session.getActiveUser().getEmail();
  const requisicoesConfiguracao = [];

  idsAbas.forEach(idAba => {
    // Visibilidade
    requisicoesConfiguracao.push({
      updateSheetProperties: {
        properties: {
          sheetId: idAba,
          hidden: false
        },
        fields: 'hidden'
      }
    });

    // Proteção com áreas desprotegidas
    requisicoesConfiguracao.push({
      addProtectedRange: {
        protectedRange: {
          range: { sheetId: idAba },
          unprotectedRanges: [
            { sheetId: idAba, startRowIndex: 6, endRowIndex: 70, startColumnIndex: 0, endColumnIndex: 18 },    // A7:R70
            { sheetId: idAba, startRowIndex: 6, endRowIndex: 70, startColumnIndex: 49, endColumnIndex: 50 },   // AX7:AX70
            { sheetId: idAba, startRowIndex: 6, endRowIndex: 70, startColumnIndex: 26, endColumnIndex: 27 },   // AA7:AA70
            { sheetId: idAba, startRowIndex: 6, endRowIndex: 70, startColumnIndex: 22, endColumnIndex: 24 },   // W7:X70
            { sheetId: idAba, startRowIndex: 6, endRowIndex: 70, startColumnIndex: 51, endColumnIndex: 54 },   // AZ7:BB70
          ],
          editors: { users: [emailUsuario] }
        }
      }
    });
  });

  if (requisicoesConfiguracao.length) {
    Sheets.Spreadsheets.batchUpdate({
      requests: requisicoesConfiguracao
    }, planilhaId);
  }
}

/**
 * Escreve os nomes das turmas na célula A5 de cada aba em lote
 * @param {string} planilhaId - ID da planilha
 * @param {Array<string>} nomesTurmas - Lista de nomes das turmas
 */
function escreverNomesEmLote(planilhaId, nomesTurmas) {
  const intervalosValores = nomesTurmas.map(nomeTurma => ({
    range: `${nomeTurma}!A5`,
    values: [[nomeTurma]]
  }));

  Sheets.Spreadsheets.Values.batchUpdate({
    valueInputOption: 'RAW',
    data: intervalosValores
  }, planilhaId);
}

/**
 * Mostra temporariamente a aba Base para duplicação
 * @param {Spreadsheet} planilha - A planilha ativa
 * @param {string} nomeAbaBase - Nome da aba base
 * @returns {boolean} true se a aba estava oculta, false caso contrário
 */
function mostrarAbaBase(planilha, nomeAbaBase) {
  const abaBase = planilha.getSheetByName(nomeAbaBase);

  if (!abaBase) {
    Logger.log(`⚠️ Aba '${nomeAbaBase}' não encontrada`);
    return false;
  }

  if (abaBase.isSheetHidden()) {
    abaBase.showSheet();
    Logger.log(`👁️ Aba '${nomeAbaBase}' exibida temporariamente`);
    return true;
  }

  return false;
}

/**
 * Oculta a aba Base após a duplicação
 * @param {Spreadsheet} planilha - A planilha ativa
 * @param {string} nomeAbaBase - Nome da aba base
 * @param {boolean} estavaoculta - Se a aba estava oculta antes
 */
function ocultarAbaBase(planilha, nomeAbaBase, estavaoculta) {
  if (!estavaoculta) {
    return; // Se não estava oculta, não ocultar agora
  }

  const abaBase = planilha.getSheetByName(nomeAbaBase);
  if (abaBase && !abaBase.isSheetHidden()) {
    abaBase.hideSheet();
    Logger.log(`🔒 Aba '${nomeAbaBase}' ocultada novamente`);
  }
}

/**
 * Remove todas as proteções das abas de turmas listadas na aba "Piloto"
 * @param {Spreadsheet} planilha - A planilha ativa
 */
function removerProtecoes(planilha) {
  Logger.log('🗑️ Removendo proteções...');
  const abaPiloto = planilha.getSheetByName("Piloto");
  const nomesTurmas = abaPiloto.getRange("C4:C40").getValues().flat();
  let contador = 0;

  nomesTurmas.forEach(nomeTurma => {
    const aba = planilha.getSheetByName(nomeTurma);
    if (aba) {
      const protecoes = aba.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protecoes.forEach(protecao => protecao.remove());
      contador++;
      Logger.log(`  ✓ Proteções removidas: ${nomeTurma}`);
    }
  });

  Logger.log(`✅ Total de ${contador} aba(s) desprotegida(s)`);
}
