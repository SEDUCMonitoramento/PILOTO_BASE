// ============================================================================
// CONFIGURAÇÕES
// ============================================================================

const ID_PLANILHA_ORIGEM = "1PX47xrTnfC5jacdSpaQUWPXd0a2HusxiyhcOb6CSrYk";
const NOME_ABA_MODELO = "Doc_Ananf";
const NOME_PASTA_ANANF = "ANANF";
const PREFIXO_PLANILHA = "ANANF_";
const TIMEZONE = "America/Sao_Paulo";
const FORMATO_DATA = "yyyy/MM/dd HH:mm:ss"; // Formato da data São Paulo Brasil
const CELULA_RA = "B13"; // Célula onde está o RA do aluno

const CELULA_NOME = "D12"; // Célula onde está o Nome do aluno

// ============================================================================
// FUNÇÃO PRINCIPAL
// ============================================================================

/**
 * Cria nova planilha ANANF com nome baseado no RA do aluno.
 * Replica a aba modelo mantendo formatação, dados, validações e layout.
 */
function replicarAbaParaOutraPlanilha() {
  console.time("Replicação Total");
  console.log("Iniciando replicação ANANF");

  try {
    const planilhaAtiva = SpreadsheetApp.getActiveSpreadsheet();
    const resultPasta = getOuCriaPastaANANF(planilhaAtiva);
    const idPastaANANF = resultPasta.id_pasta_ananf;

    // Carrega aba modelo (Template)
    const abaModelo = carregarAbaModelo();

    // Identifica planilha ativa (Formulário do Aluno) para pegar dados inicias
    // planilhaAtiva já foi declarada acima
    const abaAtiva = planilhaAtiva.getActiveSheet();

    // Extrai RA e Nome da aba ATIVA (Onde o usuário está)
    // Se não encontrar na aba ativa, tenta ler do mapa global se as células baterem
    const raAluno = abaAtiva.getRange(CELULA_RA).getValue();
    const nomeAluno = abaAtiva.getRange(CELULA_NOME).getValue();

    if (!raAluno) throw new Error("RA do aluno não encontrado na célula " + CELULA_RA + " da planilha ativa.");

    console.log(`RA do aluno: ${raAluno}`);
    console.log(`Nome do aluno: ${nomeAluno}`);

    // Cria nova planilha com nome baseado no RA
    const { novaPlanilha, novaAba } = criarNovaPlanilha(idPastaANANF, raAluno);

    // Replica estrutura, dados e formatação
    replicarConteudo(abaModelo, novaAba);

    const urlPlanilha = novaPlanilha.getUrl();
    console.timeEnd("Replicação Total");
    Logger.log(`Planilha criada com sucesso: ${novaPlanilha.getName()}`);
    Logger.log(`URL: ${urlPlanilha}`);

    // Mostra mensagem de confirmação para o usuário
    mostrarMensagemANANFGerado(nomeAluno, urlPlanilha);

    return urlPlanilha;

  } catch (erro) {
    console.error("Erro na replicação ANANF: " + erro.message);
    SpreadsheetApp.getUi().alert("Erro ao gerar ANANF: " + erro.message);
    throw erro;
  }
}

// ============================================================================
// FUNÇÕES AUXILIARES
// ============================================================================

/**
 * Carrega aba modelo da planilha origem
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} A aba modelo
 */
function carregarAbaModelo() {
  const planilhaOrigem = SpreadsheetApp.openById(ID_PLANILHA_ORIGEM);
  const abaModelo = planilhaOrigem.getSheetByName(NOME_ABA_MODELO);

  if (!abaModelo) {
    throw new Error(`Aba "${NOME_ABA_MODELO}" não encontrada na planilha de origem!`);
  }

  console.log(`Aba modelo carregada`);
  return abaModelo;
}

/**
 * Cria planilha com nome formato: ANANF_[RA] - [data]
 * @param {string} idPastaANANF ID da pasta onde a planilha será salva
 * @param {string|number} raAluno RA do aluno para compor o nome
 * @returns {Object} Objeto contendo { novaPlanilha, novaAba }
 */
function criarNovaPlanilha(idPastaANANF, raAluno) {
  // Gera nome: ANANF_123456789 - 2026/01/19 10:55:00
  const dataFormatada = Utilities.formatDate(new Date(), TIMEZONE, FORMATO_DATA);
  const nomePlanilha = `${PREFIXO_PLANILHA}${raAluno} - ${dataFormatada}`;

  const novaPlanilha = SpreadsheetApp.create(nomePlanilha);

  // Remove aba padrão 'Página1' se existir e renomeia ou usa a primeira
  let novaAba = novaPlanilha.getSheets()[0];
  novaAba.setName(NOME_ABA_MODELO);

  // Move para a pasta correta
  const pastaDestino = DriveApp.getFolderById(idPastaANANF);
  const arquivo = DriveApp.getFileById(novaPlanilha.getId());
  arquivo.moveTo(pastaDestino);

  console.log(`Planilha criada: ${nomePlanilha}`);
  return { novaPlanilha, novaAba };
}

/**
 * Replica dados, formatação, validações e dimensões da aba origem para destino.
 * IMPORTANTE: Converte fórmulas em valores estáticos ("Hard Copy") para não quebrar referências,
 * mas preserva toda a formatação visual (cores, bordas, mesclagens).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} abaOrigem
 * @param {GoogleAppsScript.Spreadsheet.Sheet} abaDestino
 */
function replicarConteudo(abaOrigem, abaDestino) {
  console.log("Replicando conteúdo completo...");

  const numLinhas = abaOrigem.getMaxRows();
  const numColunas = abaOrigem.getMaxColumns();
  console.log(`Dimensões: ${numLinhas} linhas × ${numColunas} colunas`);

  // 1. Ajustar dimensões da planilha destino
  ajustarDimensoes(abaDestino, numLinhas, numColunas);

  // 2. Copiar TUDO (Formatação, Mesclagens, Validações, Larguras)
  // Usamos PASTE_NORMAL para garantir que o visual veridico seja transferido
  const rangeOrigem = abaOrigem.getRange(1, 1, numLinhas, numColunas);
  const rangeDestino = abaDestino.getRange(1, 1);
  rangeOrigem.copyTo(rangeDestino, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  // 3. TRANSFERÊNCIA DE DADOS MAPEADOS (Do info.js)
  // Lê os dados da planilha ATIVA (onde o script roda) e injeta na nova planilha (Template)
  console.log("Transferindo dados mapeados do info.js...");
  const planilhaAtiva = SpreadsheetApp.getActiveSpreadsheet();
  const abaAtiva = planilhaAtiva.getActiveSheet(); // Assume que os dados estão na aba ativa atual

  // Itera sobre o mapa de dados e copia valor por valor
  // Isso garante que apenas os dados úteis sejam transferidos para o formatação limpa
  for (const [campo, celula] of Object.entries(MAPA_DADOS_ANANF)) {
    try {
      const valor = abaAtiva.getRange(celula).getValue();
      abaDestino.getRange(celula).setValue(valor);
      console.log(`  -> Campo '${campo}' (${celula}): copiado.`);
    } catch (e) {
      console.warn(`  [!] Falha ao copiar campo '${campo}' (${celula}): ${e.message}`);
    }
  }

  // 4. Ajustar larguras e alturas (garantia extra)
  copiarDimensoesVisuais(abaOrigem, abaDestino, numLinhas, numColunas);

  console.log("Conteúdo replicado e dados preenchidos com sucesso!");
}

/**
 * Ajusta linhas e colunas da aba destino para casar com a origem
 */
function ajustarDimensoes(aba, linhasNecessarias, colunasNecessarias) {
  const linhasAtuais = aba.getMaxRows();
  const colunasAtuais = aba.getMaxColumns();

  if (linhasNecessarias > linhasAtuais) {
    aba.insertRowsAfter(linhasAtuais, linhasNecessarias - linhasAtuais);
  } else if (linhasNecessarias < linhasAtuais) {
    // Opcional: deletar linhas extras se quiser limpar a planilha nova
    // aba.deleteRows(linhasNecessarias + 1, linhasAtuais - linhasNecessarias);
  }

  if (colunasNecessarias > colunasAtuais) {
    aba.insertColumnsAfter(colunasAtuais, colunasNecessarias - colunasAtuais);
  } else if (colunasNecessarias < colunasAtuais) {
    // aba.deleteColumns(colunasNecessarias + 1, colunasAtuais - colunasNecessarias);
  }
}

/**
 * Copia larguras das colunas e alturas das linhas
 */
function copiarDimensoesVisuais(abaOrigem, abaDestino, numLinhas, numColunas) {
  console.log("  Ajustando larguras e alturas");

  // Copiar largura das colunas
  for (let col = 1; col <= numColunas; col++) {
    const largura = abaOrigem.getColumnWidth(col);
    abaDestino.setColumnWidth(col, largura);
  }

  // Copiar altura das linhas
  for (let linha = 1; linha <= numLinhas; linha++) {
    const altura = abaOrigem.getRowHeight(linha);
    abaDestino.setRowHeight(linha, altura);
  }
}

/**
 * Obtém ou cria pasta ANANF na mesma pasta da planilha atual
 */
function getOuCriaPastaANANF(planilha) {
  const pastaPai = getPastaDaPlanilha(planilha);

  if (!pastaPai) {
    throw new Error("A planilha atual não está salva em nenhuma pasta do Drive.");
  }

  const pastas = pastaPai.getFoldersByName(NOME_PASTA_ANANF);
  let pastaANANF;

  if (pastas.hasNext()) {
    pastaANANF = pastas.next();
    console.log(`Pasta ANANF encontrada: ${pastaANANF.getName()}`);
  } else {
    pastaANANF = pastaPai.createFolder(NOME_PASTA_ANANF);
    console.log(`Pasta ANANF criada: ${pastaANANF.getName()}`);
  }

  return {
    id_pasta_pai: pastaPai.getId(),
    id_pasta_ananf: pastaANANF.getId()
  };
}

/**
 * Obtém pasta pai da planilha
 */
function getPastaDaPlanilha(planilha) {
  const arquivo = DriveApp.getFileById(planilha.getId());
  const parents = arquivo.getParents();

  return parents.hasNext() ? parents.next() : null;
}

/**
 * Mostra mensagem de sucesso (Toast e Alert opcional)
 * @param {string} nomeAluno Nome do aluno
 * @param {string} urlPlanilha URL da planilha gerada
 */
function mostrarMensagemANANFGerado(nomeAluno, urlPlanilha) {
  const msg = `ANANF gerado com sucesso para: ${nomeAluno}`;

  // Tenta mostrar Toast (notificação discreta)
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Sucesso", 10);
  } catch (e) {
    console.log("Não foi possível mostrar toast.");
  }

  // Mostra caixa de diálogo com link copiável (melhor UX)
  const htmlOutput = HtmlService.createHtmlOutput(
    `<p>O ANANF para <strong>${nomeAluno}</strong> foi criado.</p>
     <p><a href="${urlPlanilha}" target="_blank">Clique aqui para abrir a planilha</a></p>`
  )
    .setWidth(300)
    .setHeight(150);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ANANF Criado');
}
