/**
 * ============================================================================
 * SISTEMA ANANF - DOCUMENTAÇÃO
 * ============================================================================
 * 
 * OBJETIVO:
 * Gerar planilhas ANANF individualizadas para cada aluno, preservando dados e formatação.
 * 
 * ============================================================================
 * ESTRUTURA DE ARQUIVOS DO PROJETO
 * ============================================================================
 * 
 * ANANF/
 * ├── GLOBAL/
 * │   └── BOTÃO.js              # Configuração de menus e botões (onOpen)
 * │
 * ├── HELPER/
 * │   └── mensagem.js           # Funções de UI e mensagens
 * │
 * ├── Sidebar/
 * │   ├── SidebarANANF.html     # Interface HTML do sidebar
 * │   └── FuncoesANANF.js       # [DEPRECADO] Migrado para HELPER/
 * │
 * ├── _ReplicaANANF.js          # Lógica principal de replicação
 * └── _etapa.js                 # Este arquivo (documentação)
 * 
 * ============================================================================
 * FLUXO DE EXECUÇÃO DETALHADO (_ReplicaANANF.js)
 * ============================================================================
 * 
 * 1. INICIALIZAÇÃO
 *    - Identifica planilha ativa e aba modelo "Doc_Ananf".
 *    - Verifica ou cria a pasta "ANANF" no mesmo diretório da planilha mestra.
 *    -> Função: getOuCriaPastaANANF()
 * 
 *                  ↓
 * 
 * 2. EXTRAÇÃO DE DADOS
 *    - Lê o RA do aluno na célula B13.
 *    - Lê o Nome do aluno na célula D12.
 *    -> Função: carregarAbaModelo()
 * 
 *                  ↓
 * 
 * 3. CRIAÇÃO DO ARQUIVO
 *    - Cria nova planilha vazia com nome: "ANANF_[RA] - [DATA]".
 *    - Move este novo arquivo para dentro da pasta "ANANF".
 *    - Renomeia a aba principal para "Doc_Ananf".
 *    -> Função: criarNovaPlanilha()
 * 
 *                  ↓
 * 
 * 4. REPLICAÇÃO INTELIGENTE (Core)
 *    a) Ajusta dimensões (linhas/colunas) da nova aba para igualar à origem.
 *       -> Função: ajustarDimensoes()
 *    b) PASTE NORMAL: Copia TUDO (Visual, Bordas, Cores, Mesclagens, Validações).
 *       -> Função: range.copyTo(PasteType.PASTE_NORMAL)
 *    c) HARD COPY: Lê os valores da origem e COLAR VALORES por cima na destino.
 *       -> Função: range.setValues()
 *    d) Ajusta larguras de colunas e alturas de linhas pixel-a-pixel.
 *       -> Função: copiarDimensoesVisuais()
 *    
 *    -> Função Principal: replicarConteudo()
 * 
 *                  ↓
 * 
 * 5. FINALIZAÇÃO
 *    - Exibe notificação "Toast" de sucesso.
 *    - Abre painel com Link clicável para a nova planilha gerada.
 *    -> Função: mostrarMensagemANANFGerado()
 * 
 * ============================================================================
 * FUNCIONALIDADES PRINCIPAIS
 * ============================================================================
 * 
 * - Cópia Fiel: O visual é idêntico ao original.
 * - Dados Congelados: Fórmulas são convertidas em texto/número fixo.
 * - Tratamento de Erros: Avisa se faltar RA, se não achar a aba, etc.
 * 
 */