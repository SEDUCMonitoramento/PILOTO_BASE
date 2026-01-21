/**
 * Abre o sidebar com a lista de ANANFs gerados
 */
function abrirSidebarANANF() {
    const html = HtmlService.createHtmlOutputFromFile('Sidebar/SidebarANANF')
        .setTitle('ANANFs Gerados')
        .setWidth(300);

    SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Lista todos os ANANFs gerados na pasta ANANF
 * @returns {Array} Array de objetos com informações dos ANANFs
 */
function listarANANFs() {
    try {
        const planilhaAtiva = SpreadsheetApp.getActiveSpreadsheet();
        const pastaInfo = getOuCriaPastaANANF(planilhaAtiva);
        const pastaANANF = DriveApp.getFolderById(pastaInfo.id_pasta_ananf);

        const arquivos = pastaANANF.getFiles();
        const ananfs = [];

        while (arquivos.hasNext()) {
            const arquivo = arquivos.next();
            const nome = arquivo.getName();

            // Filtra apenas arquivos ANANF (que começam com "ANANF_")
            if (nome.startsWith('ANANF_')) {
                ananfs.push({
                    nome: nome,
                    url: arquivo.getUrl(),
                    data: Utilities.formatDate(arquivo.getLastUpdated(), TIMEZONE, 'dd/MM/yyyy HH:mm'),
                    id: arquivo.getId()
                });
            }
        }

        // Ordena por data de modificação (mais recente primeiro)
        ananfs.sort((a, b) => b.data.localeCompare(a.data));

        return ananfs;

    } catch (erro) {
        console.error('Erro ao listar ANANFs:', erro);
        throw new Error('Não foi possível listar os ANANFs: ' + erro.message);
    }
}

/**
 * Mostra mensagem de confirmação após gerar ANANF
 * @param {string} nomeAluno - Nome do aluno
 * @param {string} urlPlanilha - URL da planilha gerada
 */
function mostrarMensagemANANFGerado(nomeAluno, urlPlanilha) {
    const ui = SpreadsheetApp.getUi();

    const resposta = ui.alert(
        'ANANF Gerado com Sucesso! ✅',
        `ANANF criado para o aluno:\n\n${nomeAluno}\n\nDeseja abrir a planilha agora?`,
        ui.ButtonSet.YES_NO
    );

    if (resposta === ui.Button.YES) {
        const htmlOutput = HtmlService.createHtmlOutput(
            `<script>window.open('${urlPlanilha}', '_blank');google.script.host.close();</script>`
        ).setWidth(1).setHeight(1);

        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Abrindo...');
    }
}
