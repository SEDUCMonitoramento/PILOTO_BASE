/**==================================
 * Constantes
 *==================================*/
// Nova Planilha Modelo fornecida pelo usuário
const ID_PLANILHA_ORIGEM = "1ix0ymrAAL1kwYN4Yyc1QiwRW_R-68xyvVqkbPuJQTgg";
const NOME_ABA_MODELO = "Doc_Ananf";
const NOME_PASTA_ANANF = "ANANF";
const PREFIXO_PLANILHA = "ANANF_";
const TIMEZONE = "America/Sao_Paulo";
const FORMATO_DATA = "yyyy/MM/dd HH:mm:ss";

/**==================================
 * Funções
 *==================================*/

/**
 * Cria nova planilha ANANF com nome baseado no RA do aluno.
 * Replica a aba modelo mantendo formatação, dados, validações e layout.
 */
function replicarAbaParaOutraPlanilha() {
  console.time("Replicação Total");
  console.log(">> [INÍCIO] Iniciando replicação ANANF");
  Logger.log(">> [INÍCIO] script replicarAbaParaOutraPlanilha()");

  try {
    const planilhaAtiva = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log(`Planilha ativa ID: ${planilhaAtiva.getId()}`);

    const resultPasta = getOuCriaPastaANANF(planilhaAtiva);
    const idPastaANANF = resultPasta.id_pasta_ananf;
    Logger.log(`Pasta ANANF ID: ${idPastaANANF}`);

    const abaModelo = carregarAbaModelo();
    Logger.log("Aba modelo carregada com sucesso.");

    const abaAtiva = planilhaAtiva.getActiveSheet();
    Logger.log(`Aba ativa: ${abaAtiva.getName()}`);

    // Extrai RA e Nome usando o mapa global (info.js)
    const raAluno = abaAtiva.getRange(MAPA_DADOS_ANANF.RA).getValue();
    const nomeAluno = abaAtiva.getRange(MAPA_DADOS_ANANF.nomeAluno).getValue();

    Logger.log(`Dados Extraídos - RA: ${raAluno} | Nome: ${nomeAluno}`);

    if (!raAluno) {
      Logger.log("ERRO: RA não encontrado.");
      throw new Error(`RA do aluno não encontrado na célula ${MAPA_DADOS_ANANF.RA} da planilha ativa.`);
    }

    console.log(`RA do aluno: ${raAluno}`);
    console.log(`Nome do aluno: ${nomeAluno}`);

    // Cria nova planilha já com a aba copiada
    // Alteração crucial: Passamos a abaModelo para copiar via Sheet.copyTo
    const { novaPlanilha, novaAba } = criarNovaPlanilhaComCopia(idPastaANANF, raAluno, abaModelo);
    Logger.log(`Nova planilha criada: ${novaPlanilha.getName()} (ID: ${novaPlanilha.getId()})`);

    // Preenche os dados
    preencherDados(novaAba);

    const urlPlanilha = novaPlanilha.getUrl();
    console.timeEnd("Replicação Total");

    Logger.log(`>> [SUCESSO] Planilha criada.`);
    Logger.log(`URL: ${urlPlanilha}`);

    mostrarMensagemANANFGerado(nomeAluno, urlPlanilha);

    return urlPlanilha;

  } catch (erro) {
    console.error("Erro na replicação ANANF: " + erro.message);
    Logger.log(">> [ERRO CRÍTICO]: " + erro.message);
    // Stack trace só é útil no console IDE, mas logamos message
    if (erro.stack) Logger.log(erro.stack);
    SpreadsheetApp.getUi().alert("Erro ao gerar ANANF: " + erro.message);
    throw erro;
  }
}

/**==================================
 * Funções auxiliares
 *==================================*/

/**
 * Carrega aba modelo da planilha origem
 */
function carregarAbaModelo() {
  Logger.log(`Tentando abrir planilha origem: ${ID_PLANILHA_ORIGEM}`);
  const planilhaOrigem = SpreadsheetApp.openById(ID_PLANILHA_ORIGEM);
  const abaModelo = planilhaOrigem.getSheetByName(NOME_ABA_MODELO);

  if (!abaModelo) {
    Logger.log(`ERRO: Aba ${NOME_ABA_MODELO} não encontrada.`);
    throw new Error(`Aba "${NOME_ABA_MODELO}" não encontrada na planilha de origem!`);
  }
  return abaModelo;
}

/**
 * Cria planilha nova e já copia a aba modelo para dentro dela.
 * Remove a aba padrão que vem na criação.
 */
function criarNovaPlanilhaComCopia(idPastaANANF, raAluno, abaModelo) {
  Logger.log("Criando nova planilha e copiando aba modelo...");
  const dataFormatada = Utilities.formatDate(new Date(), TIMEZONE, FORMATO_DATA);
  const nomePlanilha = `${PREFIXO_PLANILHA}${raAluno} - ${dataFormatada}`;

  // 1. Cria nova planilha (vem com 'Página1' padrão)
  const novaPlanilha = SpreadsheetApp.create(nomePlanilha);

  // 2. Copia a ABA INTEIRA do modelo para a nova planilha
  // Isso resolve o erro de 'Target range and source range must be on the same spreadsheet'
  Logger.log("Executando abaModelo.copyTo(novaPlanilha)...");
  const novaAba = abaModelo.copyTo(novaPlanilha);

  // 3. Renomeia a aba copiada para o nome correto
  novaAba.setName(NOME_ABA_MODELO);

  // 4. Remove a aba padrão 'Página1' (sempre a primeira, índice 0 se a cópia foi para o fim)
  // Geralmente a cópia vira a última ou segunda. Vamos listar e remover a que não é a nossa.
  const sheets = novaPlanilha.getSheets();
  if (sheets.length > 1) {
    // Remove a primeira aba se ela não for a que acabamos de criar (segurança)
    // A aba padrão 'Página1' geralmente é sheets[0] logo após criação
    const abaPadrao = sheets[0];
    if (abaPadrao.getId() !== novaAba.getId()) {
      Logger.log(`Removendo aba padrão: ${abaPadrao.getName()}`);
      novaPlanilha.deleteSheet(abaPadrao);
    }
  }

  // 5. Move o arquivo para a pasta correta
  const pastaDestino = DriveApp.getFolderById(idPastaANANF);
  const arquivo = DriveApp.getFileById(novaPlanilha.getId());
  arquivo.moveTo(pastaDestino);

  console.log(`Planilha criada: ${nomePlanilha}`);
  Logger.log(`Arquivo movido para pasta de destino.`);

  return { novaPlanilha, novaAba };
}

/**
 * Preenche dados da info.js na nova aba.
 * Não precisa mais copiar formatação/estrutura pois copyTo da aba já fez isso.
 */
function preencherDados(abaDestino) {
  console.log("Transferindo dados mapeados do info.js...");
  Logger.log("Iniciando transferência de dados mapeados...");

  const planilhaAtiva = SpreadsheetApp.getActiveSpreadsheet(); // Onde roda o script
  const abaAtiva = planilhaAtiva.getActiveSheet();

  for (const [campo, celula] of Object.entries(MAPA_DADOS_ANANF)) {
    try {
      const valor = abaAtiva.getRange(celula).getValue();
      abaDestino.getRange(celula).setValue(valor);
      Logger.log(`Campo [${campo}] na celula [${celula}] preenchido com: ${valor}`);
    } catch (e) {
      console.warn(`[!] Falha ao copiar campo '${campo}' (${celula}): ${e.message}`);
      Logger.log(`[AVISO] Falha ao copiar campo '${campo}': ${e.message}`);
    }
  }
}

function getOuCriaPastaANANF(planilha) {
  const pastaPai = getPastaDaPlanilha(planilha);

  // Se não encontrar pasta pai (ex: planilha na raiz), tenta usar o Root do Drive ou lança erro?
  // Código original lançava erro. Vamos manter ou melhorar.
  if (!pastaPai) {
    // throw new Error("A planilha atual não está salva em nenhuma pasta do Drive.");
    Logger.log("Aviso: Planilha na raiz ou sem pasta pai identificável. Buscando na raiz do Drive.");
    // Fallback: buscar na raiz se a planilha não tiver pais (raro, mas possível em "Computadores" ou "Shared with me" dependendo do contexto)
    // Mas vamos manter a lógica original: exigir pasta pai, pois organização é importante.
    throw new Error("A planilha atual não está salva em nenhuma pasta do Drive (Root?). Impossível criar subpasta lá.");
  }

  const pastas = pastaPai.getFoldersByName(NOME_PASTA_ANANF);
  const pastaANANF = pastas.hasNext() ? pastas.next() : pastaPai.createFolder(NOME_PASTA_ANANF);

  return { id_pasta_ananf: pastaANANF.getId() };
}

function getPastaDaPlanilha(planilha) {
  try {
    const parents = DriveApp.getFileById(planilha.getId()).getParents();
    return parents.hasNext() ? parents.next() : null;
  } catch (e) {
    Logger.log("Erro ao buscar pasta pai: " + e.message);
    return null;
  }
}

function mostrarMensagemANANFGerado(nomeAluno, urlPlanilha) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(`ANANF gerado com sucesso para: ${nomeAluno}`, "Sucesso", 10);
  } catch (e) { }

  const htmlOutput = HtmlService.createHtmlOutput(
    `<p>O ANANF para <strong>${nomeAluno}</strong> foi criado.</p>
     <p><a href="${urlPlanilha}" target="_blank">Clique aqui para abrir a planilha</a></p>`
  ).setWidth(300).setHeight(150);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ANANF Criado');
}
