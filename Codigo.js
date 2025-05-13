// --- Global Configuration ---
const GLOBAL_CONFIG = {
    destacarTermos: true,
    exibirNumerosLinha: true,
    maxResultadosPorAba: 1000,
    abasParaIgnorarPadrao: ['Configurações', 'Menu'],
    RESULT_SHEET_BASE: "Resultados",
    CONSULTA_SHEET_BASE: "Consulta",
    FILTRO_DATA_SHEET_BASE: "FiltroData",
    FILTRO_AVAN_SHEET_BASE: "FiltroAvancado",
    DUPLICATE_SHEET_BASE: "Duplicatas",
    COLOR_OK: "#c6efce",
    COLOR_OVERDUE: "#ffc7ce",
    COLOR_UNDEFINED: "#fff2cc",
    COLOR_DUPLICATE: "#fcf8e3"
};

// --- Properties Service Keys ---
// Usado para armazenar configurações específicas do usuário/documento
// --- Properties Service Keys ---
// Usado para armazenar configurações específicas do usuário/documento
const PROP_KEYS = {
  // MI & Drive
  MI_SHEET_NAME: 'MI_SHEET_NAME',
  MI_DRIVE_FOLDER_ID: 'MI_DRIVE_FOLDER_ID',
  MI_ID_COL_HEADER: 'MI_ID_COL_HEADER',
  MI_LINK_COL_HEADER: 'MI_LINK_COL_HEADER',
  MI_SUBJECT_COL_HEADER: 'MI_SUBJECT_COL_HEADER',
  // Consulta Única
  PROC_NUM_COL_HEADER: 'PROC_NUM_COL_HEADER',
  // Gerais
  ABAS_IGNORADAS_USER: 'ABAS_IGNORADAS_USER',
  ENVIAR_EMAIL: 'ENVIAR_EMAIL',
  // Pesquisa Dedicada
  SEARCHABLE_COLUMNS: 'SEARCHABLE_COLUMNS',
  // Formatação Status
  PRAZO_COL_HEADER: 'PRAZO_COL_HEADER',
  CONCLUSAO_COL_HEADER: 'CONCLUSAO_COL_HEADER',
  // Padronização de Data
  DATE_STANDARDIZE_COL_HEADERS: 'DATE_STANDARDIZE_COL_HEADERS',
  DATE_STANDARDIZE_TARGET_FORMAT: 'DATE_STANDARDIZE_TARGET_FORMAT',
  // Detecção Duplicatas
  DUPLICATE_CHECK_COL_HEADER: 'DUPLICATE_CHECK_COL_HEADER',
  DUPLICATE_ACTION: 'DUPLICATE_ACTION',
  // Google Agenda
  CALENDAR_DATE_COL_HEADER: 'CALENDAR_DATE_COL_HEADER',
  CALENDAR_TITLE_COL_HEADER: 'CALENDAR_TITLE_COL_HEADER',
  CALENDAR_ID: 'CALENDAR_ID',
  // Google Docs
  DOC_TRIGGER_STATUS_VALUE: 'DOC_TRIGGER_STATUS_VALUE',
  DOC_STATUS_COL_HEADER: 'DOC_STATUS_COL_HEADER',
  DOC_TEMPLATE_ID: 'DOC_TEMPLATE_ID',
  DOC_SAVE_FOLDER_ID: 'DOC_SAVE_FOLDER_ID',
  DOC_INCLUDE_COLS: 'DOC_INCLUDE_COLS',
  // Log de Atividades
  LOG_SHEET_NAME: 'LOG_SHEET_NAME',
  LOG_ENABLE_LOGGING: 'LOG_ENABLE_LOGGING',
  LOG_MAX_ROWS: 'LOG_MAX_ROWS',
  // Lembretes de Prazos
  REMINDER_ENABLE: 'REMINDER_ENABLE',
  REMINDER_SHEET_NAME: 'REMINDER_SHEET_NAME',
  REMINDER_DEADLINE_COL_HEADER: 'REMINDER_DEADLINE_COL_HEADER',
  REMINDER_RESPONSIBLE_EMAIL_COL_HEADER: 'REMINDER_RESPONSIBLE_EMAIL_COL_HEADER',
  REMINDER_TASK_NAME_COL_HEADER: 'REMINDER_TASK_NAME_COL_HEADER',
  REMINDER_DAYS_BEFORE: 'REMINDER_DAYS_BEFORE',
  REMINDER_STATUS_COL_HEADER_IGNORE: 'REMINDER_STATUS_COL_HEADER_IGNORE',
  REMINDER_STATUS_VALUES_TO_IGNORE: 'REMINDER_STATUS_VALUES_TO_IGNORE',
  REMINDER_SENT_COL_HEADER: 'REMINDER_SENT_COL_HEADER',
  // Notificação de Duplicatas
  DUPLICATE_NOTIFICATION_ENABLE: 'DUPLICATE_NOTIFICATION_ENABLE',
  DUPLICATE_NOTIFICATION_SHEETS: 'DUPLICATE_NOTIFICATION_SHEETS',
  DUPLICATE_NOTIFICATION_EMAIL_SUBJECT: 'DUPLICATE_NOTIFICATION_EMAIL_SUBJECT',
  DUPLICATE_NOTIFICATION_ACTION_OVERRIDE: 'DUPLICATE_NOTIFICATION_ACTION_OVERRIDE', // Última chave, sem vírgula no final dela
  DASHBOARD_ENABLE: 'DASHBOARD_ENABLE', // Boolean 'true' or 'false'
  DASHBOARD_SHEET_NAME: 'DASHBOARD_SHEET_NAME', // Default: "Dashboard"
  DASHBOARD_SOURCE_SHEET_NAME: 'DASHBOARD_SOURCE_SHEET_NAME', // Aba de onde os dados serão lidos
  DASHBOARD_STATUS_COL_HEADER: 'DASHBOARD_STATUS_COL_HEADER',
  DASHBOARD_RESPONSIBLE_COL_HEADER: 'DASHBOARD_RESPONSIBLE_COL_HEADER',
  DASHBOARD_DEADLINE_COL_HEADER: 'DASHBOARD_DEADLINE_COL_HEADER',
  DASHBOARD_ITEM_ID_COL_HEADER: 'DASHBOARD_ITEM_ID_COL_HEADER', // Coluna de ID do item para listar prazos
  DATA_VALIDATION_RULES: 'DATA_VALIDATION_RULES', // Armazena um array de objetos de regra
  CALENDAR_EVENT_ID_COL_HEADER: 'CALENDAR_EVENT_ID_COL_HEADER' // Ex: "ID Evento Agenda"
};

// --- Properties Service Helpers ---
function getConfigValue(key, defaultValue = null) {
    try {
        const properties = PropertiesService.getUserProperties();
        const value = properties.getProperty(key);
        if (value === null || value === 'null' || value === 'undefined') {
            return defaultValue;
        }
        return value;
    } catch (e) {
        Logger.log(`Error getting property ${key}: ${e}`);
        return defaultValue;
    }
}

function setConfigValue(key, value) {
    try {
        const properties = PropertiesService.getUserProperties();
        if (value === null || value === undefined) {
            properties.deleteProperty(key);
        } else if (typeof value === 'object') {
            try {
                properties.setProperty(key, JSON.stringify(value));
            } catch (e) {
                Logger.log(`Error stringifying object for key ${key}: ${e}`);
                properties.setProperty(key, String(value));
            }
        } else {
            properties.setProperty(key, String(value));
        }
    } catch (e) {
        Logger.log(`Error setting property ${key}: ${e}`);
        SpreadsheetApp.getUi().alert(`Erro ao salvar config ${key}. Verifique permissões.`);
    }
}

function getConfigValueBoolean(key, defaultValue = false) {
    const value = getConfigValue(key);
    if (value === null) return defaultValue;
    return value === 'true';
}

// --- Menu Creation (Adicionar estes itens) ---
// --- Menu Creation (Trecho da função onOpen) ---
// --- Menu Creation (Trecho da função onOpen) ---
// --- Menu Creation (Função onOpen Completa para a versão generalizada) ---
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Ferramentas Planilha'); // Nome genérico do menu

     // ---SUBMENU PARA UTILITÁRIOS DE ABAS ---
  menu.addItem('Consolidar Abas Selecionadas...', 'abrirDialogoConsolidarAbas'); 



  // Adicionar item para a Barra Lateral, se você a implementou
  menu.addItem('Mostrar Barra Lateral de Ações', 'abrirBarraLateral');
  menu.addSeparator();



  // Seção de Pesquisas e Consultas
  menu.addItem('Pesquisar Texto Simples', 'abrirDialogoPesquisa');
  menu.addItem('Pesquisa Avançada (Múltiplos Termos)', 'iniciarPesquisaComPrompt');
  menu.addItem('Consultar por Identificador Único', 'abrirDialogoConsulta');

  // --- Menu Dinâmico para Pesquisa por Coluna Específica ---
  try {
    const searchableColsStr = getConfigValue(PROP_KEYS.SEARCHABLE_COLUMNS, '[]');
    let searchableCols = [];
    try {
        searchableCols = JSON.parse(searchableColsStr);
    } catch (parseError) {
        Logger.log(`Falha ao parsear JSON de colunas pesquisáveis: ${searchableColsStr}. Erro: ${parseError}`);
        searchableCols = []; // Reseta para array vazio em caso de erro de parse
    }

    if (Array.isArray(searchableCols) && searchableCols.length > 0) {
      menu.addSeparator();
      const searchMenu = ui.createMenu('Pesquisar por Coluna Específica');
      searchableCols.forEach(headerName => {
        if (headerName && typeof headerName === 'string') {
          // Cria um fragmento de nome de função único baseado no headerName
          const functionSuffix = headerName.replace(/[^a-zA-Z0-9_]/g, ''); // Permite underscore
          const dynamicFunctionName = `dynamicSearch_${functionSuffix}`;

          // Define a função no escopo global (this) se ela não existir
          // Isso garante que a função esteja disponível para ser chamada pelo menu
          if (typeof this[dynamicFunctionName] !== 'function') {
            this[dynamicFunctionName] = function() {
              abrirDialogoPesquisaPorColuna(headerName);
            };
          }
          searchMenu.addItem(`Pesquisar "${headerName}"...`, dynamicFunctionName);
        }
      });
      menu.addSubMenu(searchMenu);
    }
  } catch (e) {
    Logger.log(`Erro ao construir menu dinâmico de pesquisa: ${e.toString()}`);
    // Adiciona item de configuração como fallback se a construção do menu falhar
    // menu.addItem('Configurar Colunas de Pesquisa...', 'abrirDialogoConfigColunas');
  }
  // --- Fim do Menu Dinâmico ---

  // Seção de Filtros
  menu.addSeparator();
  menu.addItem('Filtragem Detalhada', 'abrirDialogoFiltroAvancado');
  menu.addItem('Filtrar por Intervalo de Datas', 'abrirDialogoFiltroData');

  // Seção de Ações em Lote e Dashboards
  menu.addSeparator();
  menu.addItem('Aplicar Formatação de Status...', 'iniciarFormatacaoStatus');
  menu.addItem('Detectar Duplicatas...', 'iniciarDeteccaoDuplicatas');
  menu.addItem('Padronizar Formatos de Data...', 'iniciarPadronizacaoDatas');
  menu.addItem('Atualizar Dashboard', 'atualizarDashboard');

  // Seção de Integrações
  menu.addSeparator();
  menu.addItem('Criar Eventos na Agenda (Aba MI)', 'criarEventosAgendaDaAbaMI');
  menu.addItem('Gerar Documentos (Aba MI)', 'gerarDocsDaAbaMI');

  // Seção de Funcionalidades MI
  menu.addSeparator();
  menu.addItem('Adicionar Nova MI', 'adicionarMI');
  menu.addItem('Linkar Documento à MI', 'abrirDialogoLinkDocumentoMI');

  // Seção de Exportar / Backup
  menu.addSeparator();
  menu.addSubMenu(ui.createMenu('Exportar / Backup')
    .addItem('Exportar Aba Atual', 'exportarResultados')
    .addItem('Backup Completo da Planilha', 'fazerBackupCompleto'));

  // Seção de Configurações
  menu.addSeparator();
  menu.addSubMenu(ui.createMenu('Configurações')
    .addItem('Configurar MI / Drive', 'abrirDialogoConfigMI')
    .addItem('Configurar Colunas de Pesquisa Dedicada', 'abrirDialogoConfigColunas')
    .addItem('Configurar Coluna de Consulta Única', 'configurarColunaConsulta')
    .addItem('Configurar Formatação de Status', 'abrirDialogoConfigFormatacao')
    .addItem('Configurar Padronização de Datas', 'abrirDialogoConfigFormatacaoData')
    .addItem('Configurar Detecção de Duplicatas', 'abrirDialogoConfigDuplicatas')
    .addItem('Configurar Integração com Agenda', 'abrirDialogoConfigAgenda')
    .addItem('Configurar Integração com Docs', 'abrirDialogoConfigDocs')
    .addItem('Configurar Log de Atividades', 'abrirDialogoConfigLog')
    .addItem('Configurar Lembretes de Prazos', 'abrirDialogoConfigLembretes')
    .addItem('Configurar Notificação de Duplicatas', 'abrirDialogoConfigNotificacaoDuplicatas')
    .addItem('Configurar Dashboard de Resumo', 'abrirDialogoConfigDashboard')
    .addItem('Configurar Validação de Dados', 'abrirDialogoConfigValidacao') // <-- NOVO ITEM DE MENU
    .addSeparator()
    .addItem('Gerenciar Abas Ignoradas', 'gerenciarAbasIgnoradas')
    .addSeparator()
    .addItem(`Email Resultados (${getConfigValueBoolean(PROP_KEYS.ENVIAR_EMAIL, false) ? 'Ativado' : 'Desativado'})`, 'alternarEnvioEmail')
    .addItem(`Destaque Termos (${GLOBAL_CONFIG.destacarTermos ? 'Ativado' : 'Desativado'})`, 'alternarDestaqueTermos'));

  menu.addToUi();
}
// --- Dialog Openers ---
function abrirDialogoPesquisa() {
    const html = HtmlService.createHtmlOutputFromFile('formularioPesquisa').setWidth(400).setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(html, 'Pesquisar Texto Simples');
}

function abrirDialogoConsulta() {
    const html = HtmlService.createHtmlOutputFromFile('consulta').setWidth(450).setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, 'Consultar por Identificador');
}

function abrirDialogoFiltroData() {
    const html = HtmlService.createHtmlOutputFromFile('dialogoFiltroData').setWidth(450).setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, 'Filtrar por Data');
}

function abrirDialogoFiltroAvancado() {
    const html = HtmlService.createHtmlOutputFromFile('filtroAvancado').setWidth(500).setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Filtragem Detalhada');
}

function abrirDialogoLinkDocumentoMI() {
    if (!ensureMIConfigIsSet(true)) return;
    const html = HtmlService.createHtmlOutputFromFile('linkDocumentoMI').setWidth(400).setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, 'Linkar Documento à MI');
}

function abrirDialogoConfigMI() {
    const html = HtmlService.createHtmlOutputFromFile('configMI').setWidth(450).setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Configurações de MI e Drive');
}

function abrirDialogoConfigColunas() {
    const html = HtmlService.createHtmlOutputFromFile('configColunas').setWidth(450).setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Configurar Colunas de Pesquisa');
}

function abrirDialogoConfigFormatacao() {
    const html = HtmlService.createHtmlOutputFromFile('configFormatacao').setWidth(450).setHeight(350);
    SpreadsheetApp.getUi().showModalDialog(html, 'Configurar Formatação de Status');
}

function abrirDialogoPesquisaPorColuna(headerName) {
    if (!headerName) {
        Logger.log("abrirDialogoPesquisaPorColuna s/ headerName");
        SpreadsheetApp.getUi().alert("Erro: Nome da coluna não especificado.");
        return;
    }
    const template = HtmlService.createTemplateFromFile('pesquisaColuna');
    template.headerName = headerName;
    const htmlOutput = template.evaluate().setWidth(400).setHeight(230);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Pesquisar em "${headerName}"`);
}

function abrirDialogoConfigDuplicatas() {
    const html = HtmlService.createHtmlOutputFromFile('configDuplicatas').setWidth(450).setHeight(350);
    SpreadsheetApp.getUi().showModalDialog(html, 'Configurar Detecção de Duplicatas');
}

function abrirDialogoConfigAgenda() {
    const html = HtmlService.createHtmlOutputFromFile('configAgenda').setWidth(450).setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Configurar Integração com Agenda');
}

function abrirDialogoConfigDocs() {
    const html = HtmlService.createHtmlOutputFromFile('configDocs').setWidth(500).setHeight(550);
    SpreadsheetApp.getUi().showModalDialog(html, 'Configurar Integração com Docs');
}

function abrirDialogoSelecaoAbas(targetFunctionName, dialogTitle = 'Selecionar Abas') {
    if (!targetFunctionName) {
        Logger.log("Target function name missing for sheet selection");
        SpreadsheetApp.getUi().alert("Erro: Função de destino não especificada.");
        return;
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();
    const sheetNames = allSheets.map(sheet => sheet.getName()).filter(name => !deveIgnorarAba(name));
    if (sheetNames.length === 0) {
        SpreadsheetApp.getUi().alert("Nenhuma aba válida para seleção.");
        return;
    }
    const template = HtmlService.createTemplateFromFile('selecaoAbas');
    template.sheetNames = sheetNames;
    template.targetFunction = targetFunctionName;
    const htmlOutput = template.evaluate().setWidth(400).setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, dialogTitle);
}

// --- Search & Filter Functions ---
function iniciarPesquisaComPrompt() {
    const ui = SpreadsheetApp.getUi();
    const resposta = ui.prompt("Pesquisar (Múltiplos Termos)", "Termos separados por vírgula:", ui.ButtonSet.OK_CANCEL);
    if (resposta.getSelectedButton() === ui.Button.OK) {
        const termos = resposta.getResponseText().trim();
        if (termos) {
            pesquisaExpandida(termos);
        } else {
            ui.alert("Nenhum termo.");
        }
    }
}

function executarPesquisa(dadoProcurado) {
    return pesquisarEExibirDadosRelacionados(dadoProcurado);
}

function buscarLinhaPorIdentificador(identificador, colunaHeader) {
    const ui = SpreadsheetApp.getUi();
    const p = SpreadsheetApp.getActiveSpreadsheet();
    const abas = p.getSheets();
    if (!identificador || !colunaHeader) {
        ui.alert("ID e Coluna obrigatórios.");
        return;
    }
    setConfigValue(PROP_KEYS.PROC_NUM_COL_HEADER, colunaHeader);
    const safeId = identificador.replace(/[^a-zA-Z0-9_-]/g, '_').substring(0, 30);
    const destNome = `${GLOBAL_CONFIG.CONSULTA_SHEET_BASE}_${safeId}`;
    let destAba = p.getSheetByName(destNome);
    if (!destAba) {
        destAba = p.insertSheet(destNome);
    } else {
        destAba.clear();
    }
    let lDest = 1;
    let achou = false;
    let cabP = null;
    destAba.getRange(lDest, 1).setValue(`Consulta: "${identificador}" em "${colunaHeader}"`).setFontWeight('bold');
    lDest += 2;
    for (const sh of abas) {
        if (deveIgnorarAba(sh.getName())) continue;
        const dr = sh.getDataRange();
        if (!dr || dr.getHeight() <= 1) continue;
        const dados = dr.getValues();
        if (!cabP) cabP = dados[0];
        const idxCol = findColumnIndexByHeader(dados[0], colunaHeader);
        if (idxCol === -1) continue;
        if (!achou && cabP) {
            destAba.getRange(lDest, 1, 1, cabP.length).setValues([cabP]).setFontWeight('bold');
            lDest++;
        }
        for (let i = 1; i < dados.length; i++) {
            const lin = dados[i];
            if (String(lin[idxCol]).trim().toLowerCase() === String(identificador).trim().toLowerCase()) {
                const rtw = lin.slice(0, cabP.length);
                while (rtw.length < cabP.length) rtw.push('');
                destAba.getRange(lDest, 1, 1, cabP.length).setValues([rtw]);
                lDest++;
                achou = true;
            }
        }
    }
    if (!achou) {
        destAba.getRange(lDest, 1).setValue("ID não encontrado.");
        ui.alert(`"${identificador}" não encontrado em "${colunaHeader}".`);
    } else {
        destAba.autoResizeColumns(1, destAba.getLastColumn());
        SpreadsheetApp.setActiveSheet(destAba);
        ui.alert(`Consulta concluída.`);
    }
}

function pesquisarEExibirDadosRelacionados(dadoProcurado) {
    const p = SpreadsheetApp.getActiveSpreadsheet();
    const abas = p.getSheets();
    const termoLow = String(dadoProcurado).toLowerCase().trim();
    if (!termoLow) return "Termo vazio.";
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const resNome = `${GLOBAL_CONFIG.RESULT_SHEET_BASE}_${ts}`;
    const resAba = p.insertSheet(resNome);
    let lAt = 1;
    let totRes = 0;
    let cabGlob = null;
    resAba.getRange(lAt, 1).setValue(`Resultados para: "${dadoProcurado}"`).setFontWeight('bold');
    lAt += 2;
    for (const sh of abas) {
        if (deveIgnorarAba(sh.getName())) continue;
        const dr = sh.getDataRange();
        if (!dr || dr.getHeight() <= 1) continue;
        const dados = dr.getValues();
        const cab = dados[0];
        if (!cabGlob) cabGlob = GLOBAL_CONFIG.exibirNumerosLinha ? ["Linha", ...cab] : [...cab];
        const linFilt = [];
        for (let i = 1; i < dados.length; i++) {
            const lin = dados[i];
            const linTxtLow = lin.map(cel => String(cel).toLowerCase()).join(" || ");
            if (linTxtLow.includes(termoLow)) {
                const linNum = GLOBAL_CONFIG.exibirNumerosLinha ? [i + 1, ...lin] : [...lin];
                linFilt.push(linNum.slice(0, cabGlob.length));
            }
            if (linFilt.length >= GLOBAL_CONFIG.maxResultadosPorAba) break;
        }
        if (linFilt.length > 0) {
            if (totRes === 0) {
                resAba.getRange(lAt, 1, 1, cabGlob.length).setValues([cabGlob]).setFontWeight('bold');
                lAt++;
            }
            const rgRes = resAba.getRange(lAt, 1, linFilt.length, cabGlob.length);
            rgRes.setValues(linFilt);
            if (GLOBAL_CONFIG.destacarTermos) {
                destacarTermosEncontrados(resAba, lAt, linFilt.length, cabGlob.length, [termoLow]);
            }
            lAt += linFilt.length;
            totRes += linFilt.length;
        }
    }
    if (totRes === 0) {
        resAba.getRange(1, 1).setValue(`Nada encontrado: "${dadoProcurado}"`);
        SpreadsheetApp.getUi().alert("Nenhum resultado.");
        return "Nenhum resultado.";
    } else {
        resAba.autoResizeColumns(1, resAba.getLastColumn());
        SpreadsheetApp.setActiveSheet(resAba);
        SpreadsheetApp.getUi().alert(`Busca OK (${totRes} resultados).`);
        return `Busca OK.`;
    }
}

function pesquisaExpandida(termosProcurados) {
    const p = SpreadsheetApp.getActiveSpreadsheet();
    const abas = p.getSheets();
    const termosLow = termosProcurados.toLowerCase().split(",").map(t => t.trim()).filter(t => t.length > 0);
    if (termosLow.length === 0) {
        SpreadsheetApp.getUi().alert("Nenhum termo válido.");
        return;
    }
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const resNome = `${GLOBAL_CONFIG.RESULT_SHEET_BASE}_Adv_${ts}`;
    const resAba = p.insertSheet(resNome);
    let lAt = 1;
    let totRes = 0;
    resAba.getRange(lAt, 1).setValue(`Pesq. Avançada: ${termosProcurados}`).setFontWeight('bold');
    lAt += 2;
    for (const sh of abas) {
        const nomeSh = sh.getName();
        if (deveIgnorarAba(nomeSh)) continue;
        const dr = sh.getDataRange();
        if (!dr || dr.getHeight() <= 1) continue;
        const dados = dr.getValues();
        const cab = dados[0];
        const cabFin = GLOBAL_CONFIG.exibirNumerosLinha ? ["Linha", ...cab] : [...cab];
        const linFilt = [];
        const formats = [];
        for (let i = 1; i < dados.length; i++) {
            const lin = dados[i];
            const linCelLow = lin.map(c => String(c).toLowerCase());
            const achou = linCelLow.some(cell => termosLow.some(t => cell.includes(t)));
            if (achou) {
                const linComp = GLOBAL_CONFIG.exibirNumerosLinha ? [i + 1, ...lin] : [...lin];
                linFilt.push(linComp.slice(0, cabFin.length));
                if (GLOBAL_CONFIG.destacarTermos) {
                    const destLin = lin.map(cell => {
                        const cellLow = String(cell).toLowerCase();
                        return termosLow.some(t => cellLow.includes(t)) ? "#fff2cc" : null;
                    });
                    const destFin = GLOBAL_CONFIG.exibirNumerosLinha ? [null, ...destLin] : [...destLin];
                    formats.push(destFin.slice(0, cabFin.length));
                }
                if (linFilt.length >= GLOBAL_CONFIG.maxResultadosPorAba) break;
            }
        }
        if (linFilt.length > 0) {
            resAba.getRange(lAt, 1).setValue(`Resultados aba: ${nomeSh}`).setFontWeight('bold');
            lAt++;
            resAba.getRange(lAt, 1, 1, cabFin.length).setValues([cabFin]).setFontWeight('bold');
            lAt++;
            const rgDados = resAba.getRange(lAt, 1, linFilt.length, cabFin.length);
            rgDados.setValues(linFilt);
            if (GLOBAL_CONFIG.destacarTermos && formats.length > 0) {
                const bgs = rgDados.getBackgrounds();
                for (let r = 0; r < formats.length; r++) {
                    for (let c = 0; c < formats[r].length; c++) {
                        if (formats[r][c]) {
                            bgs[r][c] = formats[r][c];
                        }
                    }
                }
                rgDados.setBackgrounds(bgs);
            }
            lAt += linFilt.length + 1;
            totRes += linFilt.length;
        }
    }
    if (totRes === 0) {
        resAba.getRange(1, 1).setValue(`Nada encontrado: ${termosProcurados}`);
        SpreadsheetApp.getUi().alert("Nenhuma ocorrência.");
    } else {
        resAba.autoResizeColumns(1, resAba.getLastColumn());
        SpreadsheetApp.setActiveSheet(resAba);
        SpreadsheetApp.getUi().alert(`Pesquisa OK (${totRes} resultados).`);
        if (getConfigValueBoolean(PROP_KEYS.ENVIAR_EMAIL, false)) {
            enviarResultadosPorEmail(p, resAba, termosProcurados);
        }
    }
}

function filtrarPorIntervaloData(dtIniStr, dtFimStr, colDtHead) {
    const ui = SpreadsheetApp.getUi();
    const p = SpreadsheetApp.getActiveSpreadsheet();
    const shAt = p.getActiveSheet();
    if (deveIgnorarAba(shAt.getName())) {
        ui.alert("Função não aplicável aqui.");
        return;
    }
    if (!colDtHead) {
        ui.alert("Nome da coluna data obrigatório.");
        return;
    }
    const dtIni = new Date(dtIniStr);
    const dtFim = new Date(dtFimStr);
    dtFim.setHours(23, 59, 59, 999);
    if (isNaN(dtIni.getTime()) || isNaN(dtFim.getTime())) {
        ui.alert("Datas inválidas.");
        return;
    }
    const dr = shAt.getDataRange();
    if (!dr || dr.getHeight() <= 1) {
        ui.alert("Aba ativa sem dados.");
        return;
    }
    const dados = dr.getValues();
    const cab = dados[0];
    const idxColDt = findColumnIndexByHeader(cab, colDtHead);
    if (idxColDt === -1) {
        ui.alert(`Coluna "${colDtHead}" não encontrada.`);
        return;
    }
    const linFilt = [];
    for (let i = 1; i < dados.length; i++) {
        const lin = dados[i];
        const valCel = lin[idxColDt];
        let dtLin = null;
        if (valCel instanceof Date && !isNaN(valCel.getTime())) {
            dtLin = valCel;
        } else if (valCel) {
            try {
                dtLin = new Date(valCel);
                if (isNaN(dtLin.getTime())) dtLin = null;
            } catch (e) {
                dtLin = null;
            }
        }
        if (dtLin && dtLin >= dtIni && dtLin <= dtFim) {
            linFilt.push(lin.slice(0, cab.length));
        }
    }
    if (linFilt.length === 0) {
        ui.alert(`Nenhum registro entre ${dtIni.toLocaleDateString()} e ${dtFim.toLocaleDateString()}.`);
        return;
    }
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const nvaAbaNome = `${GLOBAL_CONFIG.FILTRO_DATA_SHEET_BASE}_${shAt.getName()}_${ts}`;
    const nvaAba = p.insertSheet(nvaAbaNome);
    nvaAba.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
    nvaAba.getRange(2, 1, linFilt.length, cab.length).setValues(linFilt);
    nvaAba.autoResizeColumns(1, cab.length);
    SpreadsheetApp.setActiveSheet(nvaAba);
    ui.alert(`Filtro OK (${linFilt.length} linhas).`);
}

function executarFiltroAvancado(filtros) {
    const ui = SpreadsheetApp.getUi();
    const p = SpreadsheetApp.getActiveSpreadsheet();
    const shAt = p.getActiveSheet();
    if (deveIgnorarAba(shAt.getName())) {
        ui.alert("Selecione uma aba de dados.");
        return;
    }
    const CH = {
        D: filtros.colData || 'Data',
        R: filtros.colResponsavel || 'Responsável',
        A: filtros.colAcao || 'O que foi feito',
        S: filtros.colSecretaria || 'Secretaria',
        AP: filtros.colAssuntoProcesso || 'Número do processo',
        RT: filtros.colReiteracao || 'Reiteração'
    };
    const dr = shAt.getDataRange();
    if (!dr || dr.getHeight() <= 1) {
        ui.alert("Aba ativa sem dados.");
        return;
    }
    const dados = dr.getValues();
    const cab = dados[0];
    const idx = {};
    let misCol = [];
    idx.D = findColumnIndexByHeader(cab, CH.D);
    idx.R = findColumnIndexByHeader(cab, CH.R);
    idx.A = findColumnIndexByHeader(cab, CH.A);
    idx.S = findColumnIndexByHeader(cab, CH.S);
    idx.AP = findColumnIndexByHeader(cab, CH.AP);
    idx.RT = findColumnIndexByHeader(cab, CH.RT);
    if ((filtros.dataInicio || filtros.dataFim) && idx.D === -1) misCol.push(CH.D);
    if (filtros.responsavel && idx.R === -1) misCol.push(CH.R);
    if (filtros.acao && idx.A === -1) misCol.push(CH.A);
    if (filtros.secretaria && idx.S === -1) misCol.push(CH.S);
    if (filtros.assuntoProcesso && idx.AP === -1) misCol.push(CH.AP);
    if (filtros.reiteracao && idx.RT === -1) misCol.push(CH.RT);
    if (misCol.length > 0) {
        ui.alert(`Colunas não encontradas: ${misCol.join(', ')}.`);
        return;
    }
    const dtIni = filtros.dataInicio ? new Date(filtros.dataInicio) : null;
    const dtFim = filtros.dataFim ? new Date(filtros.dataFim) : null;
    if (dtFim) dtFim.setHours(23, 59, 59, 999);
    const respL = filtros.responsavel ? filtros.responsavel.toLowerCase() : null;
    const acaoL = filtros.acao ? filtros.acao.toLowerCase() : null;
    const secL = filtros.secretaria ? filtros.secretaria.toLowerCase() : null;
    const apL = filtros.assuntoProcesso ? filtros.assuntoProcesso.toLowerCase() : null;
    const rtL = filtros.reiteracao ? filtros.reiteracao.toLowerCase() : null;
    const linFilt = [];
    for (let i = 1; i < dados.length; i++) {
        const lin = dados[i];
        let match = true;
        if (match && dtIni && idx.D !== -1) {
            const d = new Date(lin[idx.D]);
            if (isNaN(d.getTime()) || d < dtIni) match = false;
        }
        if (match && dtFim && idx.D !== -1) {
            const d = new Date(lin[idx.D]);
            if (isNaN(d.getTime()) || d > dtFim) match = false;
        }
        if (match && respL && idx.R !== -1) {
            if (!String(lin[idx.R]).toLowerCase().includes(respL)) match = false;
        }
        if (match && acaoL && idx.A !== -1) {
            if (!String(lin[idx.A]).toLowerCase().includes(acaoL)) match = false;
        }
        if (match && secL && idx.S !== -1) {
            if (!String(lin[idx.S]).toLowerCase().includes(secL)) match = false;
        }
        if (match && apL && idx.AP !== -1) {
            if (!String(lin[idx.AP]).toLowerCase().includes(apL)) match = false;
        }
        if (match && rtL && idx.RT !== -1) {
            if (!String(lin[idx.RT]).toLowerCase().includes(rtL)) match = false;
        }
        if (match) {
            linFilt.push(lin.slice(0, cab.length));
        }
    }
    if (linFilt.length === 0) {
        ui.alert("Nenhum registro encontrado.");
        return;
    }
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const nvaAbaNome = `${GLOBAL_CONFIG.FILTRO_AVAN_SHEET_BASE}_${shAt.getName()}_${ts}`;
    const nvaAba = p.insertSheet(nvaAbaNome);
    nvaAba.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
    nvaAba.getRange(2, 1, linFilt.length, cab.length).setValues(linFilt);
    nvaAba.autoResizeColumns(1, cab.length);
    SpreadsheetApp.setActiveSheet(nvaAba);
    ui.alert(`Filtragem OK (${linFilt.length} linhas).`);
}

function executarPesquisaPorColuna(termo, headerName) {
    const ui = SpreadsheetApp.getUi();
    const p = SpreadsheetApp.getActiveSpreadsheet();
    const abas = p.getSheets();
    const termoLow = String(termo).toLowerCase().trim();
    if (!termoLow) return {
        success: false,
        message: "Termo vazio."
    };
    if (!headerName) return {
        success: false,
        message: "Coluna não especificada."
    };
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const safeH = headerName.replace(/[^a-zA-Z0-9_]/g, '').substring(0, 20);
    const resNome = `${GLOBAL_CONFIG.RESULT_SHEET_BASE}_${safeH}_${ts}`;
    let resAba;
    try {
        resAba = p.insertSheet(resNome);
    } catch (e) {
        const shortTs = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMddHHmm");
        resNome = `${GLOBAL_CONFIG.RESULT_SHEET_BASE}_${safeH}_${shortTs}`;
        resAba = p.insertSheet(resNome);
    }
    let lAt = 1;
    let totRes = 0;
    let cabGlob = null;
    resAba.getRange(lAt, 1).setValue(`Resultados para "${termo}" em "${headerName}"`).setFontWeight('bold');
    lAt += 2;
    for (const sh of abas) {
        if (deveIgnorarAba(sh.getName())) continue;
        const dr = sh.getDataRange();
        if (!dr || dr.getHeight() <= 1) continue;
        const dados = dr.getValues();
        const cab = dados[0];
        const idxCol = findColumnIndexByHeader(cab, headerName);
        if (idxCol === -1) continue;
        if (!cabGlob) {
            cabGlob = GLOBAL_CONFIG.exibirNumerosLinha ? ["Linha", ...cab] : [...cab];
            resAba.getRange(lAt, 1, 1, cabGlob.length).setValues([cabGlob]).setFontWeight('bold');
            lAt++;
        }
        const linFilt = [];
        for (let i = 1; i < dados.length; i++) {
            const lin = dados[i];
            if (String(lin[idxCol]).toLowerCase().includes(termoLow)) {
                const linComp = GLOBAL_CONFIG.exibirNumerosLinha ? [i + 1, ...lin] : [...lin];
                linFilt.push(linComp.slice(0, cabGlob.length));
            }
            if (linFilt.length >= GLOBAL_CONFIG.maxResultadosPorAba) break;
        }
        if (linFilt.length > 0) {
            const rgRes = resAba.getRange(lAt, 1, linFilt.length, cabGlob.length);
            rgRes.setValues(linFilt);
            if (GLOBAL_CONFIG.destacarTermos) {
                const targetColHl = GLOBAL_CONFIG.exibirNumerosLinha ? idxCol + 1 : idxCol;
                const hlRg = resAba.getRange(lAt, targetColHl + 1, linFilt.length, 1);
                hlRg.setBackground("#fff2cc");
            }
            lAt += linFilt.length;
            totRes += linFilt.length;
        }
    }
    if (totRes === 0) {
        resAba.getRange(1, 1).setValue(`"${termo}" não encontrado em "${headerName}"`);
        return {
            success: true,
            message: "Nenhum resultado."
        };
    } else {
        resAba.autoResizeColumns(1, resAba.getLastColumn());
        SpreadsheetApp.setActiveSheet(resAba);
        return {
            success: true,
            message: `Busca OK (${totRes} resultados).`
        };
    }
}

// --- MI and Document Linking Functions ---
function ensureMIConfigIsSet(checkLinkCols = false) {
    const ui = SpreadsheetApp.getUi();
    const miShN = getConfigValue(PROP_KEYS.MI_SHEET_NAME);
    const drFId = getConfigValue(PROP_KEYS.MI_DRIVE_FOLDER_ID);
    const idCol = getConfigValue(PROP_KEYS.MI_ID_COL_HEADER);
    const linkCol = getConfigValue(PROP_KEYS.MI_LINK_COL_HEADER);
    let mis = [];
    if (!miShN) mis.push("Aba MI");
    if (!drFId) mis.push("ID Pasta Drive");
    if (!idCol) mis.push("Coluna ID MI");
    if (checkLinkCols && !linkCol) mis.push("Coluna Link MI");
    if (mis.length > 0) {
        ui.alert(`Config MI Incompleta`, `Faltando: ${mis.join(", ")}\nUse Configurações.`, ui.ButtonSet.OK);
        return false;
    }
    if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(miShN)) {
        ui.alert(`Erro: Aba MI "${miShN}" não encontrada.`);
        return false;
    }
    return true;
}

function adicionarMI() {
    const ui = SpreadsheetApp.getUi();
    if (!ensureMIConfigIsSet(true)) return;
    const miShN = getConfigValue(PROP_KEYS.MI_SHEET_NAME);
    const idColH = getConfigValue(PROP_KEYS.MI_ID_COL_HEADER);
    const linkColH = getConfigValue(PROP_KEYS.MI_LINK_COL_HEADER);
    const subjColH = getConfigValue(PROP_KEYS.MI_SUBJECT_COL_HEADER);
    const miSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(miShN);
    if (!miSh) {
        ui.alert(`Erro crítico: Aba "${miShN}" sumiu.`);
        return;
    }
    const miRsp = ui.prompt(`Add MI (${miShN})`, `Valor para "${idColH}":`, ui.ButtonSet.OK_CANCEL);
    if (miRsp.getSelectedButton() !== ui.Button.OK || !miRsp.getResponseText()) {
        ui.alert("Cancelado.");
        return;
    }
    const miId = miRsp.getResponseText().trim();
    let miSubj = "";
    if (subjColH) {
        const miSubjRsp = ui.prompt(`Add MI (${miId})`, `Valor para "${subjColH}" (Opcional):`, ui.ButtonSet.OK_CANCEL);
        if (miSubjRsp.getSelectedButton() === ui.Button.OK) {
            miSubj = miSubjRsp.getResponseText().trim();
        }
    }
    const headRow = miSh.getRange(1, 1, 1, miSh.getLastColumn()).getValues()[0];
    const idColIdx = findColumnIndexByHeader(headRow, idColH);
    const linkColIdx = findColumnIndexByHeader(headRow, linkColH);
    const subjColIdx = subjColH ? findColumnIndexByHeader(headRow, subjColH) : -1;
    let misHead = [];
    if (idColIdx === -1) misHead.push(idColH);
    if (linkColIdx === -1) misHead.push(linkColH);
    if (subjColH && subjColIdx === -1) misHead.push(subjColH);
    if (misHead.length > 0) {
        ui.alert(`Erro: Colunas não encontradas em "${miShN}": ${misHead.join(', ')}.`);
        return;
    }
    const newRw = new Array(headRow.length).fill('');
    newRw[idColIdx] = miId;
    newRw[linkColIdx] = "[Upload Pendente]";
    if (subjColIdx !== -1) {
        newRw[subjColIdx] = miSubj;
    }
    miSh.appendRow(newRw);
    miSh.getRange(miSh.getLastRow(), 1, 1, newRw.length).setBackground("#f3f3f3");
    ui.alert(`Entrada "${miId}" adicionada. Faça upload e use "Linkar Documento".`);
}

function linkarDocumentoMI(miIdentificador, nomeArquivo) {
    const ui = SpreadsheetApp.getUi();
    const miShN = getConfigValue(PROP_KEYS.MI_SHEET_NAME);
    const drFId = getConfigValue(PROP_KEYS.MI_DRIVE_FOLDER_ID);
    const idColH = getConfigValue(PROP_KEYS.MI_ID_COL_HEADER);
    const linkColH = getConfigValue(PROP_KEYS.MI_LINK_COL_HEADER);
    if (!miShN || !drFId || !idColH || !linkColH) {
        return {
            success: false,
            message: "Config MI Incompleta."
        };
    }
    if (!miIdentificador || !nomeArquivo) {
        return {
            success: false,
            message: "ID MI e Nome Arquivo obrigatórios."
        };
    }
    const miSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(miShN);
    if (!miSh) {
        return {
            success: false,
            message: `Aba "${miShN}" não encontrada.`
        };
    }
    miIdentificador = miIdentificador.trim();
    nomeArquivo = nomeArquivo.trim();
    const headRow = miSh.getRange(1, 1, 1, miSh.getLastColumn()).getValues()[0];
    const idColIdx = findColumnIndexByHeader(headRow, idColH);
    const linkColIdx = findColumnIndexByHeader(headRow, linkColH);
    if (idColIdx === -1) return {
        success: false,
        message: `Coluna ID "${idColH}" não encontrada.`
    };
    if (linkColIdx === -1) return {
        success: false,
        message: `Coluna Link "${linkColH}" não encontrada.`
    };
    const data = miSh.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][idColIdx]).trim().toLowerCase() === miIdentificador.toLowerCase()) {
            rowIdx = i;
            break;
        }
    }
    if (rowIdx === -1) {
        return {
            success: false,
            message: `ID "${miIdentificador}" não encontrado em "${idColH}".`
        };
    }
    try {
        const folder = DriveApp.getFolderById(drFId);
        const files = folder.getFilesByName(nomeArquivo);
        let fileFnd = null;
        if (files.hasNext()) {
            fileFnd = files.next();
            if (files.hasNext()) {
                return {
                    success: false,
                    message: `Múltiplos arquivos "${nomeArquivo}" na pasta.`
                };
            }
        } else {
            return {
                success: false,
                message: `Arquivo "${nomeArquivo}" não encontrado na pasta Drive.`
            };
        }
        const fileUrl = fileFnd.getUrl();
        const targetC = miSh.getRange(rowIdx + 1, linkColIdx + 1);
        targetC.setValue(fileUrl);
        targetC.setBackground(null);
        return {
            success: true,
            message: `Documento linkado!`
        };
    } catch (e) {
        Logger.log(`Link error: ${e}`);
        if (e.message.includes("Folder not found")) return {
            success: false,
            message: `Erro: Pasta Drive (ID: ${drFId}) não encontrada/acesso negado.`
        };
        if (e.message.includes("Access denied")) return {
            success: false,
            message: `Erro: Permissão negada para Pasta Drive (ID: ${drFId}).`
        };
        return {
            success: false,
            message: `Erro Drive/Planilha: ${e.message}`
        };
    }
}

// --- Utility and Helper Functions ---
function findColumnIndexByHeader(headerRowArray, headerName) {
    if (!headerRowArray || !headerName) return -1;
    const headerLower = headerName.trim().toLowerCase();
    return headerRowArray.findIndex(h => String(h).trim().toLowerCase() === headerLower);
}

function destacarTermosEncontrados(aba, linhaInicial, numLinhas, numColunas, termosLower) {
    if (!GLOBAL_CONFIG.destacarTermos || termosLower.length === 0) return;
    const MAX_HL_ROWS = 500;
    numLinhas = Math.min(numLinhas, MAX_HL_ROWS);
    if (numLinhas <= 0) return;
    try {
        const range = aba.getRange(linhaInicial, 1, numLinhas, numColunas);
        const valores = range.getDisplayValues();
        const bgs = range.getBackgrounds();
        for (let r = 0; r < valores.length; r++) {
            for (let c = 0; c < valores[r].length; c++) {
                const celLow = String(valores[r][c]).toLowerCase();
                const curBg = bgs[r][c].toLowerCase();
                if ((curBg === '#ffffff' || curBg === '' || curBg === null) && termosLower.some(t => celLow.includes(t))) {
                    bgs[r][c] = "#fff2cc";
                }
            }
        }
        range.setBackgrounds(bgs);
    } catch (e) {
        Logger.log(`Highlight error: ${e}`);
    }
}

function deveIgnorarAba(nomeAba) {
    const nomeLow = nomeAba.toLowerCase();
    const ignUserStr = getConfigValue(PROP_KEYS.ABAS_IGNORADAS_USER, '[]');
    let ignUser = [];
    try {
        ignUser = JSON.parse(ignUserStr);
        if (!Array.isArray(ignUser)) ignUser = [];
    } catch (e) {
        ignUser = [];
    }
    if (ignUser.includes(nomeAba)) return true;
    if (GLOBAL_CONFIG.abasParaIgnorarPadrao.some(ign => nomeLow === ign.toLowerCase())) return true;
    const basesRes = [GLOBAL_CONFIG.RESULT_SHEET_BASE, GLOBAL_CONFIG.CONSULTA_SHEET_BASE, GLOBAL_CONFIG.FILTRO_DATA_SHEET_BASE, GLOBAL_CONFIG.FILTRO_AVAN_SHEET_BASE, GLOBAL_CONFIG.DUPLICATE_SHEET_BASE];
    if (basesRes.some(base => nomeLow.startsWith(base.toLowerCase() + '_'))) return true;
    return false;
}

function enviarResultadosPorEmail(planilha, abaResultados, termosProcurados) {
    try {
        const emailDest = Session.getActiveUser().getEmail();
        if (!emailDest) {
            Logger.log("No email");
            return;
        }
        const nomeAba = abaResultados.getName();
        const idPlan = planilha.getId();
        const idAba = abaResultados.getSheetId();
        const url = `https://docs.google.com/spreadsheets/d/${idPlan}/export?format=pdf&gid=${idAba}&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=true&gridlines=false`;
        const token = ScriptApp.getOAuthToken();
        const resp = UrlFetchApp.fetch(url, {
            headers: {
                Authorization: `Bearer ${token}`
            },
            muteHttpExceptions: true
        });
        if (resp.getResponseCode() !== 200) {
            Logger.log(`PDF Error ${resp.getResponseCode()}`);
            return;
        }
        MailApp.sendEmail({
            to: emailDest,
            subject: `Busca: ${termosProcurados}`,
            body: `Anexo resultados: ${termosProcurados}\nAba: "${nomeAba}".`,
            attachments: [{
                fileName: `${nomeAba}.pdf`,
                content: resp.getBlob().getBytes(),
                mimeType: 'application/pdf'
            }]
        });
        Logger.log(`Email sent ${emailDest}`);
    } catch (e) {
        Logger.log(`Email error: ${e}`);
    }
}

// --- Configuration Management Functions ---
function alternarEnvioEmail() {
    const cur = getConfigValueBoolean(PROP_KEYS.ENVIAR_EMAIL, false);
    setConfigValue(PROP_KEYS.ENVIAR_EMAIL, !cur);
    SpreadsheetApp.getUi().alert(`Envio Email ${!cur ? 'ATIVADO' : 'DESATIVADO'}.`);
    SpreadsheetApp.flush();
    onOpen();
}

function alternarDestaqueTermos() {
    GLOBAL_CONFIG.destacarTermos = !GLOBAL_CONFIG.destacarTermos;
    SpreadsheetApp.getUi().alert(`Destaque Termos ${GLOBAL_CONFIG.destacarTermos ? 'ATIVADO' : 'DESATIVADO'}.`);
    SpreadsheetApp.flush();
    onOpen();
}

function gerenciarAbasIgnoradas() {
    const ui = SpreadsheetApp.getUi();
    const ignStr = getConfigValue(PROP_KEYS.ABAS_IGNORADAS_USER, '[]');
    let ignCur = [];
    try {
        ignCur = JSON.parse(ignStr);
        if (!Array.isArray(ignCur)) ignCur = [];
    } catch (e) {
        ignCur = [];
    }
    const resp = ui.prompt("Gerenciar Abas Ignoradas (Usuário)", `Nomes EXATOS das abas (separados por vírgula):\n(Atuais: ${ignCur.join(", ")})`, ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() === ui.Button.OK) {
        const nAbas = resp.getResponseText().split(",").map(a => a.trim()).filter(a => a.length > 0);
        setConfigValue(PROP_KEYS.ABAS_IGNORADAS_USER, JSON.stringify(nAbas));
        ui.alert("Lista de abas ignoradas atualizada.");
    }
}

function configurarColunaConsulta() {
    const ui = SpreadsheetApp.getUi();
    const cur = getConfigValue(PROP_KEYS.PROC_NUM_COL_HEADER, '');
    const resp = ui.prompt('Config Coluna Consulta Única', `Nome EXATO do cabeçalho da coluna ID:\n(Atual: "${cur}")`, ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() === ui.Button.OK) {
        const newH = resp.getResponseText().trim();
        setConfigValue(PROP_KEYS.PROC_NUM_COL_HEADER, newH || null);
        ui.alert(newH ? `Coluna definida: "${newH}"` : `Coluna removida.`);
    }
}

function salvarConfiguracaoMI(config) {
    try {
        if (config.miSheetName) setConfigValue(PROP_KEYS.MI_SHEET_NAME, config.miSheetName.trim());
        if (config.driveFolderId) setConfigValue(PROP_KEYS.MI_DRIVE_FOLDER_ID, config.driveFolderId.trim());
        if (config.idColHeader) setConfigValue(PROP_KEYS.MI_ID_COL_HEADER, config.idColHeader.trim());
        if (config.linkColHeader) setConfigValue(PROP_KEYS.MI_LINK_COL_HEADER, config.linkColHeader.trim());
        setConfigValue(PROP_KEYS.MI_SUBJECT_COL_HEADER, config.subjectColHeader ? config.subjectColHeader.trim() : null);
        return {
            success: true,
            message: "Config MI salva!"
        };
    } catch (e) {
        Logger.log(`Save MI config error: ${e}`);
        return {
            success: false,
            message: `Erro: ${e.message}`
        };
    }
}

function carregarConfiguracaoMI() {
    return {
        miSheetName: getConfigValue(PROP_KEYS.MI_SHEET_NAME, ''),
        driveFolderId: getConfigValue(PROP_KEYS.MI_DRIVE_FOLDER_ID, ''),
        idColHeader: getConfigValue(PROP_KEYS.MI_ID_COL_HEADER, ''),
        linkColHeader: getConfigValue(PROP_KEYS.MI_LINK_COL_HEADER, ''),
        subjectColHeader: getConfigValue(PROP_KEYS.MI_SUBJECT_COL_HEADER, '')
    };
}

function salvarColunasPesquisaveis(colunasArray) {
    try {
        if (!Array.isArray(colunasArray)) {
            throw new Error("Input non array.");
        }
        const colLimpas = colunasArray.map(col => String(col).trim()).filter(col => col.length > 0);
        setConfigValue(PROP_KEYS.SEARCHABLE_COLUMNS, JSON.stringify(colLimpas));
        SpreadsheetApp.flush();
        onOpen();
        return {
            success: true,
            message: "Colunas de pesquisa salvas. Menu atualizado."
        };
    } catch (e) {
        Logger.log(`Save searchable cols error: ${e}`);
        return {
            success: false,
            message: `Erro: ${e.message}`
        };
    }
}

function carregarColunasPesquisaveis() {
    const jsonStr = getConfigValue(PROP_KEYS.SEARCHABLE_COLUMNS, '[]');
    try {
        const arr = JSON.parse(jsonStr);
        return Array.isArray(arr) ? arr : [];
    } catch (e) {
        return [];
    }
}

function salvarConfigFormatacao(config) {
    try {
        setConfigValue(PROP_KEYS.PRAZO_COL_HEADER, config.prazoHeader ? config.prazoHeader.trim() : null);
        setConfigValue(PROP_KEYS.CONCLUSAO_COL_HEADER, config.conclusaoHeader ? config.conclusaoHeader.trim() : null);
        return {
            success: true,
            message: "Config formatação salva!"
        };
    } catch (e) {
        Logger.log(`Save format config error: ${e}`);
        return {
            success: false,
            message: `Erro: ${e.message}`
        };
    }
}

function carregarConfigFormatacao() {
    return {
        prazoHeader: getConfigValue(PROP_KEYS.PRAZO_COL_HEADER, ''),
        conclusaoHeader: getConfigValue(PROP_KEYS.CONCLUSAO_COL_HEADER, '')
    };
}

function salvarConfigDuplicatas(config) {
    try {
        setConfigValue(PROP_KEYS.DUPLICATE_CHECK_COL_HEADER, config.checkColHeader ? config.checkColHeader.trim() : null);
        setConfigValue(PROP_KEYS.DUPLICATE_ACTION, config.action || 'HIGHLIGHT');
        return {
            success: true,
            message: "Config duplicatas salva!"
        };
    } catch (e) {
        Logger.log(`Save Duplicates Config Error: ${e}`);
        return {
            success: false,
            message: `Erro: ${e.message}`
        };
    }
}

function carregarConfigDuplicatas() {
    return {
        checkColHeader: getConfigValue(PROP_KEYS.DUPLICATE_CHECK_COL_HEADER, ''),
        action: getConfigValue(PROP_KEYS.DUPLICATE_ACTION, 'HIGHLIGHT')
    };
}

// Em Code.gs - Funções de Configuração da Agenda
function salvarConfigAgenda(config) {
    try {
        setConfigValue(PROP_KEYS.CALENDAR_DATE_COL_HEADER, config.dateColHeader ? config.dateColHeader.trim() : null);
        setConfigValue(PROP_KEYS.CALENDAR_TITLE_COL_HEADER, config.titleColHeader ? config.titleColHeader.trim() : null);
        setConfigValue(PROP_KEYS.CALENDAR_ID, config.calendarId ? config.calendarId.trim() : 'primary');
        // --- NOVA LINHA ---
        setConfigValue(PROP_KEYS.CALENDAR_EVENT_ID_COL_HEADER, config.eventIdColHeader ? config.eventIdColHeader.trim() : "ID Evento Agenda"); // Default

        registrarLog("Config Agenda", "Configurações salvas", "INFO");
        return {success: true, message: "Configuração da Agenda salva!"};
    } catch (e) { Logger.log(`Save Agenda Config Error: ${e}`); registrarLog("Config Agenda", `Erro ao salvar: ${e.message}`, "ERROR"); return {success: false, message: `Erro: ${e.message}`}; }
}

function carregarConfigAgenda() {
     return {
        dateColHeader: getConfigValue(PROP_KEYS.CALENDAR_DATE_COL_HEADER, ''),
        titleColHeader: getConfigValue(PROP_KEYS.CALENDAR_TITLE_COL_HEADER, ''),
        calendarId: getConfigValue(PROP_KEYS.CALENDAR_ID, 'primary'),
        // --- NOVA LINHA ---
        eventIdColHeader: getConfigValue(PROP_KEYS.CALENDAR_EVENT_ID_COL_HEADER, 'ID Evento Agenda') // Default
    };
}
function salvarConfigDocs(config) {
    try {
        setConfigValue(PROP_KEYS.DOC_STATUS_COL_HEADER, config.statusColHeader ? config.statusColHeader.trim() : null);
        setConfigValue(PROP_KEYS.DOC_TRIGGER_STATUS_VALUE, config.triggerStatusValue ? config.triggerStatusValue.trim() : null);
        setConfigValue(PROP_KEYS.DOC_TEMPLATE_ID, config.templateId ? config.templateId.trim() : null);
        setConfigValue(PROP_KEYS.DOC_SAVE_FOLDER_ID, config.saveFolderId ? config.saveFolderId.trim() : null);
        let incCols = [];
        if (config.includeColsString) {
            incCols = config.includeColsString.split('\n').map(h => h.trim()).filter(h => h.length > 0);
        }
        setConfigValue(PROP_KEYS.DOC_INCLUDE_COLS, JSON.stringify(incCols));
        return {
            success: true,
            message: "Config Docs salva!"
        };
    } catch (e) {
        Logger.log(`Save Docs Config Error: ${e}`);
        return {
            success: false,
            message: `Erro: ${e.message}`
        };
    }
}

function carregarConfigDocs() {
    const incColsStr = getConfigValue(PROP_KEYS.DOC_INCLUDE_COLS, '[]');
    let incCols = [];
    try {
        incCols = JSON.parse(incColsStr);
        if (!Array.isArray(incCols)) incCols = [];
    } catch (e) {
        incCols = [];
    }
    return {
        statusColHeader: getConfigValue(PROP_KEYS.DOC_STATUS_COL_HEADER, ''),
        triggerStatusValue: getConfigValue(PROP_KEYS.DOC_TRIGGER_STATUS_VALUE, ''),
        templateId: getConfigValue(PROP_KEYS.DOC_TEMPLATE_ID, ''),
        saveFolderId: getConfigValue(PROP_KEYS.DOC_SAVE_FOLDER_ID, ''),
        includeColsString: incCols.join('\n')
    };
}
salvarConfigAgenda
// --- Status Formatting ---
function iniciarFormatacaoStatus() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert("Aplicar Formatação de Status", "Aplicar na aba atual ou selecionar múltiplas abas?", ui.ButtonSet.YES_NO_CANCEL);
    if (response === ui.Button.YES) {
        aplicarFormatacaoStatus([SpreadsheetApp.getActiveSheet().getName()]);
    } else if (response === ui.Button.NO) {
        abrirDialogoSelecaoAbas('aplicarFormatacaoStatus', 'Selecionar Abas para Formatar Status');
    }
}

function aplicarFormatacaoStatus(sheetNamesArray) {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!sheetNamesArray || sheetNamesArray.length === 0) {
        ui.alert("Nenhuma aba selecionada.");
        return;
    }
    const prazoH = getConfigValue(PROP_KEYS.PRAZO_COL_HEADER);
    const concH = getConfigValue(PROP_KEYS.CONCLUSAO_COL_HEADER);
    if (!prazoH) {
        ui.alert("Config Incompleta", "Cabeçalho 'Prazo Final' não definido. Use Configurações.", ui.ButtonSet.OK);
        return;
    }
    let totChg = 0;
    let procSh = 0;
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Formatando status em ${sheetNamesArray.length} aba(s)...`, "Progresso", -1);
    sheetNamesArray.forEach(shName => {
        const sh = ss.getSheetByName(shName);
        if (!sh || deveIgnorarAba(shName)) {
            Logger.log(`Pulando aba inválida/ignorada: ${shName}`);
            return;
        }
        const rg = sh.getDataRange();
        const vals = rg.getValues();
        if (vals.length <= 1) return;
        const bgs = rg.getBackgrounds();
        const headR = vals[0];
        const prazoIdx = findColumnIndexByHeader(headR, prazoH);
        const concIdx = concH ? findColumnIndexByHeader(headR, concH) : -1;
        if (prazoIdx === -1 || (concH && concIdx === -1)) {
            Logger.log(`Colunas data não encontradas em ${shName}.`);
            return;
        }
        let shChg = 0;
        for (let i = 1; i < vals.length; i++) {
            let tgtCol = null;
            let dtConc = null;
            let dtPrazo = null;
            if (concIdx !== -1 && vals[i][concIdx]) {
                try {
                    dtConc = new Date(vals[i][concIdx]);
                    if (isNaN(dtConc.getTime())) dtConc = null;
                } catch (e) {
                    dtConc = null;
                }
            }
            if (prazoIdx !== -1 && vals[i][prazoIdx]) {
                try {
                    dtPrazo = new Date(vals[i][prazoIdx]);
                    if (isNaN(dtPrazo.getTime())) dtPrazo = null;
                    else dtPrazo.setHours(0, 0, 0, 0);
                } catch (e) {
                    dtPrazo = null;
                }
            }
            if (dtConc) {
                tgtCol = GLOBAL_CONFIG.COLOR_OK;
            } else if (dtPrazo) {
                tgtCol = (dtPrazo < hoje) ? GLOBAL_CONFIG.COLOR_OVERDUE : GLOBAL_CONFIG.COLOR_OK;
            } else {
                tgtCol = GLOBAL_CONFIG.COLOR_UNDEFINED;
            }
            let rUpd = false;
            for (let j = 0; j < bgs[i].length; j++) {
                if (bgs[i][j] !== tgtCol) {
                    bgs[i][j] = tgtCol;
                    rUpd = true;
                }
            }
            if (rUpd) shChg++;
        }
        if (shChg > 0) {
            rg.setBackgrounds(bgs);
            totChg += shChg;
        }
        procSh++;
        SpreadsheetApp.getActiveSpreadsheet().toast(`Formatando ${shName}... (${procSh}/${sheetNamesArray.length})`, "Progresso", 10);
        Utilities.sleep(100);
    });
    SpreadsheetApp.getActiveSpreadsheet().toast("Formatação concluída.", "Progresso", 5);
    ui.alert(`Formatação Status OK.`, `${totChg} linha(s) atualizada(s) em ${procSh} aba(s).`, ui.ButtonSet.OK);
}

function aplicarFormatacaoStatusProgramado() {
    Logger.log("Exec format programada...");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSh = ss.getSheets();
    const shToFmt = [];
    allSh.forEach(sh => {
        const name = sh.getName();
        if (!deveIgnorarAba(name)) {
            shToFmt.push(name);
        }
    });
    if (shToFmt.length > 0) {
        Logger.log(`Aba(s) para formatar via gatilho: ${shToFmt.join(', ')}`);
        aplicarFormatacaoStatus(shToFmt);
    } else {
        Logger.log("Nenhuma aba para formatação programada.");
    }
    Logger.log("Format programada OK.");
}

// --- Duplicate Detection ---
function iniciarDeteccaoDuplicatas() {
    const ui = SpreadsheetApp.getUi();
    const resp = ui.alert("Detectar Duplicatas", "Verificar na aba atual ou selecionar múltiplas abas?", ui.ButtonSet.YES_NO_CANCEL);
    if (resp === ui.Button.YES) {
        detectarDuplicatasExecutar([SpreadsheetApp.getActiveSheet().getName()]);
    } else if (resp === ui.Button.NO) {
        abrirDialogoSelecaoAbas('detectarDuplicatasExecutar', 'Selecionar Abas para Verificar Duplicatas');
    }
}

function detectarDuplicatasExecutar(sheetNamesArray) {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!sheetNamesArray || sheetNamesArray.length === 0) {
        ui.alert("Nenhuma aba selecionada.");
        return;
    }
    const chkColH = getConfigValue(PROP_KEYS.DUPLICATE_CHECK_COL_HEADER);
    const act = getConfigValue(PROP_KEYS.DUPLICATE_ACTION, 'HIGHLIGHT');
    if (!chkColH) {
        ui.alert("Config Incompleta", "Cabeçalho coluna duplicatas não definido. Use Configurações.", ui.ButtonSet.OK);
        return;
    }
    let allDupsInfo = [];
    let totHl = 0;
    let procSh = 0;
    SpreadsheetApp.getActiveSpreadsheet().toast(`Detectando duplicatas em ${sheetNamesArray.length} aba(s)...`, "Progresso", -1);
    sheetNamesArray.forEach(shName => {
        const sh = ss.getSheetByName(shName);
        if (!sh || deveIgnorarAba(shName)) {
            Logger.log(`Pulando aba inválida/ignorada: ${shName}`);
            return;
        }
        const rg = sh.getDataRange();
        const vals = rg.getValues();
        if (vals.length <= 1) return;
        const headR = vals[0];
        const chkColIdx = findColumnIndexByHeader(headR, chkColH);
        if (chkColIdx === -1) {
            Logger.log(`Coluna "${chkColH}" não encontrada em ${shName}.`);
            return;
        }
        const seenVals = {};
        const dupRowsInSh = new Set();
        for (let i = 1; i < vals.length; i++) {
            const val = String(vals[i][chkColIdx]).trim();
            if (val === '') continue;
            if (seenVals[val]) {
                seenVals[val].forEach(firstIdx => dupRowsInSh.add(firstIdx + 1));
                dupRowsInSh.add(i + 1);
                seenVals[val].push(i);
            } else {
                seenVals[val] = [i];
            }
        }
        if (dupRowsInSh.size > 0) {
            if (act === 'HIGHLIGHT') {
                const bgs = rg.getBackgrounds();
                dupRowsInSh.forEach(rowIdx => {
                    for (let j = 0; j < bgs[rowIdx - 1].length; j++) {
                        bgs[rowIdx - 1][j] = GLOBAL_CONFIG.COLOR_DUPLICATE;
                    }
                    totHl++;
                });
                rg.setBackgrounds(bgs);
            } else {
                dupRowsInSh.forEach(rowIdx => {
                    allDupsInfo.push({
                        sheet: shName,
                        row: rowIdx,
                        value: vals[rowIdx - 1][chkColIdx]
                    });
                });
            }
        }
        procSh++;
        SpreadsheetApp.getActiveSpreadsheet().toast(`Verificando ${shName}... (${procSh}/${sheetNamesArray.length})`, "Progresso", 10);
        Utilities.sleep(50);
    });
    SpreadsheetApp.getActiveSpreadsheet().toast("Verificação concluída.", "Progresso", 5);
    if (act === 'HIGHLIGHT') {
        if (totHl > 0) {
            ui.alert(`Duplicatas OK`, `${totHl} linha(s) duplicada(s) destacada(s) (Coluna: "${chkColH}").`, ui.ButtonSet.OK);
        } else {
            ui.alert(`Duplicatas OK`, `Nenhuma duplicata encontrada em "${chkColH}".`, ui.ButtonSet.OK);
        }
    } else {
        if (allDupsInfo.length > 0) {
            const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
            const resShName = `${GLOBAL_CONFIG.DUPLICATE_SHEET_BASE}_${ts}`;
            let resSh = ss.getSheetByName(resShName);
            if (resSh) {
                resSh.clear();
            } else {
                resSh = ss.insertSheet(resShName);
            }
            const heads = ["Aba", "Linha", `Valor Duplicado (${chkColH})`];
            resSh.appendRow(heads).setFontWeight('bold');
            allDupsInfo.forEach(dup => {
                resSh.appendRow([dup.sheet, dup.row, dup.value]);
            });
            resSh.autoResizeColumns(1, heads.length);
            SpreadsheetApp.setActiveSheet(resSh);
            ui.alert(`Duplicatas OK`, `${allDupsInfo.length} ocorrência(s) listada(s) em "${resShName}".`, ui.ButtonSet.OK);
        } else {
            ui.alert(`Duplicatas OK`, `Nenhuma duplicata para listar (Coluna: "${chkColH}").`, ui.ButtonSet.OK);
        }
    }
}

// --- Google Calendar Integration ---
function criarEventosAgendaDaAbaMI() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dtColH = getConfigValue(PROP_KEYS.CALENDAR_DATE_COL_HEADER);
    const titColH = getConfigValue(PROP_KEYS.CALENDAR_TITLE_COL_HEADER) || getConfigValue(PROP_KEYS.MI_ID_COL_HEADER);
    const calId = getConfigValue(PROP_KEYS.CALENDAR_ID, 'primary');
    const miShN = getConfigValue(PROP_KEYS.MI_SHEET_NAME);
    if (!dtColH || !titColH) {
        ui.alert("Config Incompleta", "Colunas Data e Título da Agenda não definidas. Use Configurações.", ui.ButtonSet.OK);
        return;
    }
    if (!miShN) {
        ui.alert("Config Incompleta", "Nome aba MI não definido. Use Configurações.", ui.ButtonSet.OK);
        return;
    }
    const sh = ss.getSheetByName(miShN);
    if (!sh) {
        ui.alert(`Erro: Aba MI "${miShN}" não encontrada.`);
        return;
    }
    const rg = sh.getDataRange();
    const vals = rg.getValues();
    if (vals.length <= 1) {
        ui.alert("Aba MI sem dados.");
        return;
    }
    const headR = vals[0];
    const dtColIdx = findColumnIndexByHeader(headR, dtColH);
    const titColIdx = findColumnIndexByHeader(headR, titColH);
    if (dtColIdx === -1) {
        ui.alert(`Erro: Coluna Data "${dtColH}" não encontrada em ${miShN}.`);
        return;
    }
    if (titColIdx === -1) {
        ui.alert(`Erro: Coluna Título "${titColH}" não encontrada em ${miShN}.`);
        return;
    }
    const EV_CREATED_H = "Evento Criado?";
    let evCrColIdx = findColumnIndexByHeader(headR, EV_CREATED_H);
    if (evCrColIdx === -1) {
        sh.insertColumnAfter(headR.length);
        evCrColIdx = headR.length;
        sh.getRange(1, evCrColIdx + 1).setValue(EV_CREATED_H).setFontWeight('bold');
    }
    let evCrCount = 0;
    let errCount = 0;
    try {
        const cal = CalendarApp.getCalendarById(calId);
        if (!cal) {
            ui.alert(`Erro: Calendário ID "${calId}" não acessível.`);
            return;
        }
        SpreadsheetApp.getActiveSpreadsheet().toast(`Verificando ${vals.length-1} linhas para eventos...`, "Agenda", -1);
        for (let i = 1; i < vals.length; i++) {
            const row = vals[i];
            const evDtStr = row[dtColIdx];
            const evTit = String(row[titColIdx]).trim();
            const alrCr = String(row[evCrColIdx]).trim().toUpperCase();
            if (!evDtStr || !evTit || alrCr === 'SIM' || alrCr === 'YES') continue;
            let evDt;
            try {
                evDt = new Date(evDtStr);
                if (isNaN(evDt.getTime())) throw new Error("Invalid date");
                evDt.setHours(12, 0, 0, 0);
            } catch (e) {
                Logger.log(`Data inválida linha ${i+1}: ${evDtStr}`);
                continue;
            }
            try {
                cal.createAllDayEvent(evTit, evDt);
                evCrCount++;
                sh.getRange(i + 1, evCrColIdx + 1).setValue('Sim');
                Utilities.sleep(200);
            } catch (e) {
                Logger.log(`Erro evento linha ${i+1} (Título: ${evTit}): ${e}`);
                errCount++;
            }
            if ((evCrCount + errCount) % 10 === 0) {
                SpreadsheetApp.getActiveSpreadsheet().toast(`Progresso: ${evCrCount} eventos, ${errCount} erros...`, "Agenda", 10);
            }
        }
        SpreadsheetApp.getActiveSpreadsheet().toast("Eventos concluídos.", "Agenda", 5);
        let sumMsg = `${evCrCount} evento(s) criado(s).`;
        if (errCount > 0) {
            sumMsg += `\n${errCount} erro(s) (ver Logs).`;
        }
        ui.alert("Eventos Agenda", sumMsg, ui.ButtonSet.OK);
    } catch (e) {
        Logger.log(`Erro geral Agenda: ${e}`);
        ui.alert(`Erro geral Agenda: ${e.message}`);
    }
}

// --- Google Docs Integration ---
function gerarDocsDaAbaMI() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const statColH = getConfigValue(PROP_KEYS.DOC_STATUS_COL_HEADER);
    const trigStatVal = getConfigValue(PROP_KEYS.DOC_TRIGGER_STATUS_VALUE);
    const tmplId = getConfigValue(PROP_KEYS.DOC_TEMPLATE_ID);
    const saveFldrId = getConfigValue(PROP_KEYS.DOC_SAVE_FOLDER_ID);
    const miShN = getConfigValue(PROP_KEYS.MI_SHEET_NAME);
    const miIdColH = getConfigValue(PROP_KEYS.MI_ID_COL_HEADER);
    const incColsStr = getConfigValue(PROP_KEYS.DOC_INCLUDE_COLS, '[]');
    let incCols = [];
    try {
        incCols = JSON.parse(incColsStr);
        if (!Array.isArray(incCols)) incCols = [];
    } catch (e) {
        incCols = [];
    }
    let misCfg = [];
    if (!statColH) misCfg.push("Coluna Status");
    if (!trigStatVal) misCfg.push("Valor Status Gatilho");
    if (!saveFldrId) misCfg.push("ID Pasta Destino");
    if (!miShN) misCfg.push("Nome Aba MI");
    if (!miIdColH) misCfg.push("Coluna ID MI");
    if (incCols.length === 0) misCfg.push("Colunas a Incluir");
    if (misCfg.length > 0) {
        ui.alert("Config Incompleta", `Faltando: ${misCfg.join(', ')}. Use Configurações.`, ui.ButtonSet.OK);
        return;
    }
    const sh = ss.getSheetByName(miShN);
    if (!sh) {
        ui.alert(`Erro: Aba MI "${miShN}" não encontrada.`);
        return;
    }
    const rg = sh.getDataRange();
    const vals = rg.getValues();
    if (vals.length <= 1) {
        ui.alert("Aba MI sem dados.");
        return;
    }
    const headR = vals[0];
    const statColIdx = findColumnIndexByHeader(headR, statColH);
    const miIdColIdx = findColumnIndexByHeader(headR, miIdColH);
    const incIndices = {};
    let misIncCols = [];
    incCols.forEach(h => {
        const idx = findColumnIndexByHeader(headR, h);
        if (idx !== -1) {
            incIndices[h] = idx;
        } else {
            misIncCols.push(h);
        }
    });
    if (statColIdx === -1) {
        ui.alert(`Erro: Coluna Status "${statColH}" não encontrada.`);
        return;
    }
    if (miIdColIdx === -1) {
        ui.alert(`Erro: Coluna ID MI "${miIdColH}" não encontrada.`);
        return;
    }
    if (misIncCols.length > 0) {
        ui.alert(`Erro: Colunas a incluir não encontradas: ${misIncCols.join(', ')}.`);
        return;
    }
    const DOC_CR_H = "Doc Gerado?";
    let docCrColIdx = findColumnIndexByHeader(headR, DOC_CR_H);
    if (docCrColIdx === -1) {
        sh.insertColumnAfter(headR.length);
        docCrColIdx = headR.length;
        sh.getRange(1, docCrColIdx + 1).setValue(DOC_CR_H).setFontWeight('bold');
    }
    let docsCrCount = 0;
    let errCount = 0;
    let saveFldr;
    try {
        saveFldr = DriveApp.getFolderById(saveFldrId);
    } catch (e) {
        ui.alert(`Erro: Pasta Docs (ID: ${saveFldrId}) não encontrada/acesso negado.`);
        return;
    }
    SpreadsheetApp.getActiveSpreadsheet().toast(`Verificando ${vals.length-1} linhas para Docs...`, "Docs", -1);
    for (let i = 1; i < vals.length; i++) {
        const row = vals[i];
        const curStat = String(row[statColIdx]).trim();
        const miId = String(row[miIdColIdx]).trim();
        const alrCr = String(row[docCrColIdx]).trim().toUpperCase();
        if (curStat.toLowerCase() === trigStatVal.toLowerCase() && alrCr !== 'SIM' && alrCr !== 'YES' && miId !== '') {
            let newDoc;
            const docName = `Relatorio MI - ${miId}`;
            SpreadsheetApp.getActiveSpreadsheet().toast(`Gerando Doc ${miId}...`, "Docs", 5);
            try {
                if (tmplId) {
                    try {
                        const tmplFile = DriveApp.getFileById(tmplId);
                        newDoc = tmplFile.makeCopy(docName, saveFldr);
                    } catch (e) {
                        Logger.log(`Erro template ${tmplId}: ${e}. Criando doc vazio.`);
                        newDoc = DocumentApp.create(docName);
                        DriveApp.getFileById(newDoc.getId()).moveTo(saveFldr);
                    }
                } else {
                    newDoc = DocumentApp.create(docName);
                    DriveApp.getFileById(newDoc.getId()).moveTo(saveFldr);
                }
                const body = newDoc.getBody();
                body.clearContents();
                body.appendParagraph(docName).setHeading(DocumentApp.ParagraphHeading.TITLE);
                body.appendParagraph(`Data Geração: ${new Date().toLocaleDateString()}`).setFontSize(9).setItalic(true);
                body.appendParagraph("");
                incCols.forEach(h => {
                    const colIdx = incIndices[h];
                    const val = row[colIdx];
                    body.appendParagraph(`${h}:`).setBold(true);
                    body.appendParagraph(String(val));
                    body.appendParagraph("");
                });
                newDoc.saveAndClose();
                sh.getRange(i + 1, docCrColIdx + 1).setValue("Sim");
                const linkColH = getConfigValue(PROP_KEYS.MI_LINK_COL_HEADER);
                if (linkColH) {
                    const linkColIdx = findColumnIndexByHeader(headR, linkColH);
                    if (linkColIdx !== -1) {
                        sh.getRange(i + 1, linkColIdx + 1).setValue(newDoc.getUrl());
                    }
                }
                docsCrCount++;
                Utilities.sleep(500);
            } catch (e) {
                Logger.log(`Erro Doc linha ${i+1} (ID: ${miId}): ${e}`);
                errCount++;
                sh.getRange(i + 1, docCrColIdx + 1).setValue(`Erro: ${e.message.substring(0,50)}`).setFontColor("red");
            }
        }
        if ((docsCrCount + errCount) % 5 === 0 && (docsCrCount + errCount) > 0) {
            SpreadsheetApp.getActiveSpreadsheet().toast(`Progresso: ${docsCrCount} Docs, ${errCount} erros...`, "Docs", 10);
        }
    }
    SpreadsheetApp.getActiveSpreadsheet().toast("Geração Docs concluída.", "Docs", 5);
    let sumMsg = `${docsCrCount} doc(s) criado(s).`;
    if (errCount > 0) {
        sumMsg += `\n${errCount} erro(s) (ver Logs).`;
    }
    ui.alert("Resultado Docs", sumMsg, ui.ButtonSet.OK);
}

// --- Sheet Selection Logic ---
function processarSelecaoAbas(selectedSheetNames, targetFunctionName) {
    if (!targetFunctionName) {
        Logger.log("Target function missing in processarSelecaoAbas");
        return {
            success: false,
            message: "Erro: Função destino não recebida."
        };
    }
    if (!selectedSheetNames || !Array.isArray(selectedSheetNames) || selectedSheetNames.length === 0) {
        return {
            success: false,
            message: "Nenhuma aba selecionada."
        };
    }
    try {
        if (typeof this[targetFunctionName] === 'function') {
            Logger.log(`Chamando ${targetFunctionName} com abas: ${selectedSheetNames.join(', ')}`);
            this[targetFunctionName](selectedSheetNames);
            return {
                success: true
            };
        } else {
            Logger.log(`Target function ${targetFunctionName} not found.`);
            return {
                success: false,
                message: `Erro: Função alvo "${targetFunctionName}" não encontrada.`
            };
        }
    } catch (e) {
        Logger.log(`Error executing ${targetFunctionName}: ${e}`);
        return {
            success: false,
            message: `Erro ao executar: ${e.message}`
        };
    }
}

// --- Export and Backup Functions ---
function exportarResultados() {
    const ui = SpreadsheetApp.getUi();
    const shAt = SpreadsheetApp.getActiveSheet();
    const nomeSh = shAt.getName();
    if (deveIgnorarAba(nomeSh)) {
        ui.alert("Não é possível exportar abas ignoradas.");
        return;
    }
    const isResSh = [GLOBAL_CONFIG.RESULT_SHEET_BASE, GLOBAL_CONFIG.CONSULTA_SHEET_BASE, GLOBAL_CONFIG.FILTRO_DATA_SHEET_BASE, GLOBAL_CONFIG.FILTRO_AVAN_SHEET_BASE].some(b => nomeSh.toLowerCase().startsWith(b.toLowerCase() + '_'));
    if (!isResSh) {
        const conf = ui.alert(`Aba "${nomeSh}" não parece de resultados. Exportar mesmo assim?`, ui.ButtonSet.YES_NO);
        if (conf !== ui.Button.YES) return;
    }
    const opts = ["PDF", "Excel (XLSX)", "CSV"];
    const resp = ui.prompt(`Exportar Aba "${nomeSh}"`, `Formato:\n - PDF\n - Excel (XLSX)\n - CSV`, ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() === ui.Button.OK) {
        const fmtIn = resp.getResponseText().trim().toUpperCase();
        const p = SpreadsheetApp.getActiveSpreadsheet();
        if (fmtIn.startsWith("PDF")) {
            exportarAbaParaFormato(p, shAt, 'pdf');
        } else if (fmtIn.startsWith("EXCEL")) {
            exportarAbaParaFormato(p, shAt, 'xlsx');
        } else if (fmtIn.startsWith("CSV")) {
            exportarAbaParaFormato(p, shAt, 'csv');
        } else {
            ui.alert("Formato inválido.");
        }
    }
}

function exportarAbaParaFormato(planilha, aba, formato) {
    const ui = SpreadsheetApp.getUi();
    try {
        const emailDest = Session.getActiveUser().getEmail();
        if (!emailDest) {
            ui.alert("Não foi possível obter email.");
            return;
        }
        const nomeArqBase = aba.getName().replace(/[^a-zA-Z0-9_\-]/g, '_');
        const idPlan = planilha.getId();
        const idAba = aba.getSheetId();
        let url = `https://docs.google.com/spreadsheets/d/${idPlan}/export?gid=${idAba}&format=${formato}`;
        let mime, ext = formato;
        switch (formato) {
            case 'pdf':
                url += '&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=true&gridlines=false';
                mime = 'application/pdf';
                break;
            case 'xlsx':
                mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
                break;
            case 'csv':
                mime = 'text/csv';
                break;
            default:
                ui.alert("Formato interno inválido.");
                return;
        }
        const token = ScriptApp.getOAuthToken();
        const resp = UrlFetchApp.fetch(url, {
            headers: {
                Authorization: `Bearer ${token}`
            },
            muteHttpExceptions: true
        });
        if (resp.getResponseCode() !== 200) {
            Logger.log(`Export Error ${resp.getResponseCode()}`);
            ui.alert(`Erro ao gerar arquivo ${formato}.`);
            return;
        }
        const blob = resp.getBlob().setName(`${nomeArqBase}.${ext}`);
        MailApp.sendEmail({
            to: emailDest,
            subject: `Exportação ${formato.toUpperCase()}: ${aba.getName()}`,
            body: `Anexo: ${aba.getName()}`,
            attachments: [blob]
        });
        ui.alert(`Arquivo ${formato.toUpperCase()} enviado para ${emailDest}.`);
    } catch (e) {
        Logger.log(`Export error ${formato}: ${e}`);
        ui.alert(`Erro ao exportar: ${e.message}`);
    }
}

function fazerBackupCompleto() {
    const ui = SpreadsheetApp.getUi();
    const resp = ui.alert("Backup Completo", "Enviar cópia Excel (XLSX) para seu email?", ui.ButtonSet.YES_NO);
    if (resp === ui.Button.YES) {
        try {
            const p = SpreadsheetApp.getActiveSpreadsheet();
            const emailDest = Session.getActiveUser().getEmail();
            if (!emailDest) {
                ui.alert("Não foi possível obter email.");
                return;
            }
            const data = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
            const nomeArq = `${p.getName()}_backup_${data}.xlsx`.replace(/[^a-zA-Z0-9_\-.]/g, '_');
            const url = `https://docs.google.com/spreadsheets/d/${p.getId()}/export?format=xlsx`;
            const token = ScriptApp.getOAuthToken();
            const response = UrlFetchApp.fetch(url, {
                headers: {
                    Authorization: `Bearer ${token}`
                },
                muteHttpExceptions: true
            });
            if (response.getResponseCode() !== 200) {
                Logger.log(`Backup Error ${response.getResponseCode()}`);
                ui.alert(`Erro ao gerar backup.`);
                return;
            }
            const blob = response.getBlob().setName(nomeArq);
            MailApp.sendEmail({
                to: emailDest,
                subject: `Backup: ${p.getName()} - ${data}`,
                body: `Anexo backup da planilha ${p.getName()}`,
                attachments: [blob]
            });
            ui.alert(`Backup completo enviado para ${emailDest}.`);
        } catch (e) {
            Logger.log(`Backup error: ${e}`);
            ui.alert(`Erro no backup: ${e.message}`);
        }
    }
}

/**
 * Abre a barra lateral com ações comuns.
 */
function abrirBarraLateral() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('barraLateral')
      .setTitle('Ações Rápidas')
      .setWidth(300); // Você pode ajustar a largura
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}
function abrirDialogoConfigFormatacaoData() {
    const html = HtmlService.createHtmlOutputFromFile('configFormatacaoData')
        .setWidth(450).setHeight(450); // Ajustar altura conforme necessário
    SpreadsheetApp.getUi().showModalDialog(html, 'Configurar Padronização de Datas');
}

function salvarConfigFormatacaoData(config) {
    try {
        // Colunas são salvas como string JSON de um array
        let colunasArray = [];
        if (config.colHeadersString) {
            colunasArray = config.colHeadersString.split('\n')
                                .map(h => h.trim())
                                .filter(h => h.length > 0);
        }
        setConfigValue(PROP_KEYS.DATE_STANDARDIZE_COL_HEADERS, JSON.stringify(colunasArray));
        setConfigValue(PROP_KEYS.DATE_STANDARDIZE_TARGET_FORMAT, config.targetFormat ? config.targetFormat.trim() : "dd/MM/yyyy"); // Default format

        return {success: true, message: "Configuração de padronização de datas salva!"};
    } catch (e) {
         Logger.log(`Error saving date standardization config: ${e}`);
         return {success: false, message: `Erro ao salvar: ${e.message}`};
    }
}

function carregarConfigFormatacaoData() {
    const colHeadersJson = getConfigValue(PROP_KEYS.DATE_STANDARDIZE_COL_HEADERS, '[]');
    let colHeadersArray = [];
    try {
        colHeadersArray = JSON.parse(colHeadersJson);
        if (!Array.isArray(colHeadersArray)) colHeadersArray = [];
    } catch (e) { colHeadersArray = []; }

    return {
        colHeadersString: colHeadersArray.join('\n'), // Para exibir na textarea
        targetFormat: getConfigValue(PROP_KEYS.DATE_STANDARDIZE_TARGET_FORMAT, 'dd/MM/yyyy') // Default format
    };
}

// --- NOVAS Funções para Executar Padronização de Data ---

function iniciarPadronizacaoDatas() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
        "Padronizar Formatos de Data",
        "Deseja aplicar na aba atual ou selecionar múltiplas abas?",
        ui.ButtonSet.YES_NO_CANCEL
    );

    if (response === ui.Button.YES) {
        executarPadronizacaoDatas([SpreadsheetApp.getActiveSheet().getName()]);
    } else if (response === ui.Button.NO) {
        abrirDialogoSelecaoAbas('executarPadronizacaoDatas', 'Selecionar Abas para Padronizar Datas');
    }
    // If CANCEL, do nothing
}

function executarPadronizacaoDatas(sheetNamesArray) {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!sheetNamesArray || sheetNamesArray.length === 0) {
        ui.alert("Nenhuma aba selecionada para padronização.");
        return;
    }

    // Carregar configuração
    const colHeadersJson = getConfigValue(PROP_KEYS.DATE_STANDARDIZE_COL_HEADERS, '[]');
    const targetFormat = getConfigValue(PROP_KEYS.DATE_STANDARDIZE_TARGET_FORMAT, 'dd/MM/yyyy');
    let colHeadersToStandardize = [];
    try {
        colHeadersToStandardize = JSON.parse(colHeadersJson);
        if (!Array.isArray(colHeadersToStandardize) || colHeadersToStandardize.length === 0) {
            ui.alert("Configuração Incompleta", "Nenhuma coluna configurada para padronização de data.\nUse Configurações > Configurar Padronização de Datas.", ui.ButtonSet.OK);
            return;
        }
    } catch (e) {
        ui.alert("Erro na Configuração", "Erro ao carregar colunas para padronização. Verifique a configuração.", ui.ButtonSet.OK);
        return;
    }

    let totalCellsChanged = 0;
    let processedSheets = 0;
    const timeZone = ss.getSpreadsheetTimeZone();

    SpreadsheetApp.getActiveSpreadsheet().toast(`Iniciando padronização de datas em ${sheetNamesArray.length} aba(s)...`, "Progresso", -1);

    sheetNamesArray.forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet || deveIgnorarAba(sheetName)) {
            Logger.log(`Pulando aba inválida ou ignorada para padronização de data: ${sheetName}`);
            return;
        }

        const range = sheet.getDataRange();
        const values = range.getValues();
        const numFormats = range.getNumberFormats(); // Get existing number formats

        if (values.length <= 1) return; // Skip empty or header-only sheets

        const headerRow = values[0];
        let sheetChanges = 0;

        colHeadersToStandardize.forEach(colHeader => {
            const colIndex = findColumnIndexByHeader(headerRow, colHeader);
            if (colIndex === -1) {
                Logger.log(`Coluna "${colHeader}" não encontrada na aba "${sheetName}" para padronização de data.`);
                return; // Skip this column for this sheet
            }

            for (let i = 1; i < values.length; i++) { // Start from row 1 (after header)
                const originalValue = values[i][colIndex];
                if (originalValue === null || originalValue === '' || originalValue instanceof Date) {
                    // If it's already a date object, just ensure number format
                    if (originalValue instanceof Date && numFormats[i][colIndex] !== targetFormat) {
                        sheet.getRange(i + 1, colIndex + 1).setNumberFormat(targetFormat);
                        // No change in value, only display format for existing dates
                        // sheetChanges++; // Uncomment if you want to count format-only changes
                    }
                    continue;
                }

                let parsedDate;
                // Try to parse common date-like strings
                // This can be expanded with more robust parsing libraries or more patterns
                if (typeof originalValue === 'string') {
                    // Attempt common European format dd/mm/yyyy or dd-mm-yyyy
                    let parts = originalValue.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
                    if (parts) {
                        const day = parseInt(parts[1], 10);
                        const month = parseInt(parts[2], 10) -1; // Month is 0-indexed
                        let year = parseInt(parts[3], 10);
                        if (year < 100) { // Handle yy format
                           year += (year > 50 ? 1900 : 2000); // Arbitrary cutoff for 2-digit years
                        }
                        parsedDate = new Date(year, month, day);
                    } else {
                        // Attempt common US format mm/dd/yyyy or yyyy-mm-dd (native JS Date constructor is better here)
                        parsedDate = new Date(originalValue);
                    }
                } else if (typeof originalValue === 'number') {
                    // Check if it's a Google Sheets serial date number
                    // (Date in Sheets is number of days since Dec 30, 1899)
                    if (originalValue > 25569) { // Roughly 1/1/1970
                         parsedDate = new Date((originalValue - 25569) * 86400 * 1000);
                    }
                }


                if (parsedDate && !isNaN(parsedDate.getTime())) {
                    // Check if value or format actually needs changing
                    const currentCellValue = sheet.getRange(i + 1, colIndex + 1).getValue();
                    const isAlreadyCorrectDateObject = currentCellValue instanceof Date && currentCellValue.getTime() === parsedDate.getTime();

                    if (!isAlreadyCorrectDateObject || numFormats[i][colIndex] !== targetFormat) {
                        sheet.getRange(i + 1, colIndex + 1).setValue(parsedDate).setNumberFormat(targetFormat);
                        sheetChanges++;
                    }
                } else {
                    // Logger.log(`Could not parse '${originalValue}' as date in sheet '${sheetName}', row ${i+1}, col '${colHeader}'`);
                }
            } // End row loop for column
        }); // End column loop

        if (sheetChanges > 0) {
            totalCellsChanged += sheetChanges;
        }
        processedSheets++;
        SpreadsheetApp.getActiveSpreadsheet().toast(`Padronizando datas em ${sheetName}... (${processedSheets}/${sheetNamesArray.length})`, "Progresso", 10);
        Utilities.sleep(50); // Small delay
    }); // End sheet loop

    SpreadsheetApp.getActiveSpreadsheet().toast("Padronização de datas concluída.", "Progresso", 5);
    if (totalCellsChanged > 0) {
        ui.alert("Padronização Concluída", `${totalCellsChanged} célula(s) de data foram padronizadas para o formato "${targetFormat}" em ${processedSheets} aba(s).`, ui.ButtonSet.OK);
    } else {
        ui.alert("Padronização Concluída", `Nenhuma célula precisou de padronização de data ou colunas configuradas não foram encontradas. Verifique suas configurações.`, ui.ButtonSet.OK);
    }
}

// --- NOVAS Funções para Configuração de Notificação de Duplicatas ---

function abrirDialogoConfigNotificacaoDuplicatas() {
    const html = HtmlService.createHtmlOutputFromFile('configNotificacaoDuplicatas')
        .setWidth(480).setHeight(450); // Ajustar tamanho
    SpreadsheetApp.getUi().showModalDialog(html, 'Configurar Notificação de Duplicatas');
}

function salvarConfigNotificacaoDuplicatas(config) {
    try {
        setConfigValue(PROP_KEYS.DUPLICATE_NOTIFICATION_ENABLE, String(config.enableDuplicateNotification === true || config.enableDuplicateNotification === 'true'));
        
        let sheetsToScan = 'ALL_VALID'; // Default
        if (config.sheetsToScanString) {
            const parsedSheets = config.sheetsToScanString.split(',')
                                    .map(s => s.trim())
                                    .filter(s => s.length > 0);
            if (parsedSheets.length > 0) {
                sheetsToScan = JSON.stringify(parsedSheets);
            }
        }
        setConfigValue(PROP_KEYS.DUPLICATE_NOTIFICATION_SHEETS, sheetsToScan);
        setConfigValue(PROP_KEYS.DUPLICATE_NOTIFICATION_EMAIL_SUBJECT, config.emailSubject ? config.emailSubject.trim() : "Alerta de Duplicatas Encontradas na Planilha");
        // For notifications, we typically want the 'LIST' action.
        // We can override the general duplicate action or ensure it's set to LIST for this.
        // For simplicity here, the notification function will assume it needs a list.
        // The DUPLICATE_ACTION config is for manual detection.

        registrarLog("Config Notif. Duplicatas", "Configurações salvas", "INFO");
        return {success: true, message: "Configurações de Notificação de Duplicatas salvas!"};
    } catch (e) {
         Logger.log(`Error saving Duplicate Notification config: ${e}`);
         registrarLog("Config Notif. Duplicatas", `Erro ao salvar: ${e.message}`, "ERROR");
         return {success: false, message: `Erro ao salvar: ${e.message}`};
    }
}

function carregarConfigNotificacaoDuplicatas() {
    const sheetsJson = getConfigValue(PROP_KEYS.DUPLICATE_NOTIFICATION_SHEETS, 'ALL_VALID');
    let sheetsString = '';
    if (sheetsJson !== 'ALL_VALID') {
        try {
            const sheetsArray = JSON.parse(sheetsJson);
            if (Array.isArray(sheetsArray)) {
                sheetsString = sheetsArray.join(', ');
            }
        } catch (e) { Logger.log("Error parsing DUPLICATE_NOTIFICATION_SHEETS: " + e); }
    }

    return {
        enableDuplicateNotification: getConfigValueBoolean(PROP_KEYS.DUPLICATE_NOTIFICATION_ENABLE, false),
        sheetsToScanString: sheetsString, // For textarea display, empty means ALL_VALID
        emailSubject: getConfigValue(PROP_KEYS.DUPLICATE_NOTIFICATION_EMAIL_SUBJECT, "Alerta de Duplicatas Encontradas na Planilha")
    };
}


// --- NOVA Função para Enviar Notificação de Duplicatas (para ser chamada por gatilho) ---
/**
 * Verifica duplicatas e envia uma notificação por email se encontradas.
 * Destinada a ser executada por um gatilho de tempo.
 */
function enviarNotificacaoDuplicatas() {
    if (!getConfigValueBoolean(PROP_KEYS.DUPLICATE_NOTIFICATION_ENABLE, false)) {
        registrarLog("Notificação Duplicatas", "Funcionalidade desabilitada.", "INFO");
        return;
    }
    registrarLog("Notificação Duplicatas", "Iniciando verificação para notificação...", "INFO");

    const checkColHeader = getConfigValue(PROP_KEYS.DUPLICATE_CHECK_COL_HEADER);
    if (!checkColHeader) {
        registrarLog("Notificação Duplicatas", "Configuração da coluna de verificação de duplicatas não definida. Notificação cancelada.", "ERROR");
        return;
    }

    // Determine sheets to scan
    const sheetsToScanConfig = getConfigValue(PROP_KEYS.DUPLICATE_NOTIFICATION_SHEETS, 'ALL_VALID');
    let sheetNamesArray = [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (sheetsToScanConfig === 'ALL_VALID') {
        ss.getSheets().forEach(sheet => {
            if (!deveIgnorarAba(sheet.getName())) {
                sheetNamesArray.push(sheet.getName());
            }
        });
    } else {
        try {
            sheetNamesArray = JSON.parse(sheetsToScanConfig);
            if (!Array.isArray(sheetNamesArray)) sheetNamesArray = [];
        } catch (e) {
            registrarLog("Notificação Duplicatas", `Erro ao parsear configuração de abas: ${sheetsToScanConfig}`, "ERROR");
            return;
        }
    }

    if (sheetNamesArray.length === 0) {
        registrarLog("Notificação Duplicatas", "Nenhuma aba válida para verificar duplicatas.", "INFO");
        return;
    }

    // --- Lógica de detecção de duplicatas (adaptada de detectarDuplicatasExecutar com ação 'LIST') ---
    let allDuplicatesInfo = [];
    let processedSheetsCount = 0;

    sheetNamesArray.forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet || deveIgnorarAba(sheetName)) return;

        const range = sheet.getDataRange();
        const values = range.getValues();
        if (values.length <= 1) return;

        const headerRow = values[0];
        const checkColIndex = findColumnIndexByHeader(headerRow, checkColHeader);
        if (checkColIndex === -1) return;

        const seenValues = {};
        const duplicateRowsInSheet = new Set();

        for (let i = 1; i < values.length; i++) {
            const value = String(values[i][checkColIndex]).trim();
            if (value === '') continue;
            if (seenValues[value]) {
                seenValues[value].forEach(firstIndex => duplicateRowsInSheet.add(firstIndex + 1));
                duplicateRowsInSheet.add(i + 1);
                seenValues[value].push(i);
            } else {
                seenValues[value] = [i];
            }
        }
        if (duplicateRowsInSheet.size > 0) {
             duplicateRowsInSheet.forEach(rowIndex => {
                allDuplicatesInfo.push({
                    sheet: sheetName,
                    row: rowIndex,
                    value: values[rowIndex - 1][checkColIndex]
                });
            });
        }
        processedSheetsCount++;
    });
    // --- Fim da lógica de detecção ---

    if (allDuplicatesInfo.length > 0) {
        const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
        const emailSubject = getConfigValue(PROP_KEYS.DUPLICATE_NOTIFICATION_EMAIL_SUBJECT, `Alerta de Duplicatas: ${ss.getName()}`);
        let emailBody = `Foram encontradas ${allDuplicatesInfo.length} ocorrências duplicadas na coluna "${checkColHeader}" da planilha "${ss.getName()}".\n\nDetalhes:\n`;
        
        // Limitar o número de duplicatas no corpo do email para não ficar muito longo
        const maxDuplicatesInEmail = 20;
        allDuplicatesInfo.slice(0, maxDuplicatesInEmail).forEach(dup => {
            emailBody += `- Aba: "${dup.sheet}", Linha: ${dup.row}, Valor: "${dup.value}"\n`;
        });
        if (allDuplicatesInfo.length > maxDuplicatesInEmail) {
            emailBody += `\n... e mais ${allDuplicatesInfo.length - maxDuplicatesInEmail} ocorrências.`;
        }
        emailBody += `\n\nRecomenda-se executar a função "Detectar Duplicatas..." com a opção "Listar" para gerar uma aba com todas as duplicatas.`;
        emailBody += `\n\nLink para a planilha: ${ss.getUrl()}`;

        try {
            if (userEmail && MailApp.getRemainingDailyQuota() > 0) {
                MailApp.sendEmail(userEmail, emailSubject, emailBody);
                registrarLog("Notificação Duplicatas", `Email de alerta enviado para ${userEmail} com ${allDuplicatesInfo.length} duplicatas.`, "INFO");
            } else if (!userEmail) {
                 registrarLog("Notificação Duplicatas", "Não foi possível obter o email do usuário para enviar o alerta.", "ERROR");
            } else {
                registrarLog("Notificação Duplicatas", "Quota de email diária esgotada. Alerta de duplicatas não enviado.", "WARNING");
            }
        } catch (e) {
            registrarLog("Notificação Duplicatas", `Falha ao enviar email de alerta: ${e.toString()}`, "ERROR");
        }
    } else {
        registrarLog("Notificação Duplicatas", "Nenhuma duplicata encontrada para notificar.", "INFO");
    }
}
// --- NOVA Função onEdit para automações ---
/**
 * Executa automaticamente quando um usuário edita uma célula na planilha.
 * @param {Object} e O objeto de evento passado pelo Google Apps Script.
 */
function onEdit(e) {
  if (!e) {
    registrarLog("onEdit", "Evento de edição não recebido. Ignorando.", "WARNING");
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();
  const editedRow = range.getRow();
  const editedCol = range.getColumn();
  const newValue = e.value; // O novo valor da célula
  // const oldValue = e.oldValue; // O valor antigo, se necessário

  registrarLog("onEdit Disparado", `Usuário: ${Session.getActiveUser().getEmail()}, Aba: ${sheet.getName()}, Célula: ${range.getA1Notation()}, Novo Valor: ${newValue}`, "INFO");

  // --- Gatilho: Gerar Documento por Mudança de Status na Aba MI ---
  try {
    const miSheetName = getConfigValue(PROP_KEYS.MI_SHEET_NAME);
    const statusColHeader = getConfigValue(PROP_KEYS.DOC_STATUS_COL_HEADER);
    const triggerStatusValue = getConfigValue(PROP_KEYS.DOC_TRIGGER_STATUS_VALUE);

    // Verifica se a edição ocorreu na aba e coluna corretas
    if (sheet.getName() === miSheetName && editedRow > 1) { // Ignora cabeçalho
      if (!statusColHeader || !triggerStatusValue) {
        // registrarLog("onEdit: Doc Gen", "Configuração para geração automática de Doc por status está incompleta.", "WARNING");
        return; // Não loga toda vez para não poluir, mas configuração é necessária
      }

      const headerRowValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const statusColIndex = findColumnIndexByHeader(headerRowValues, statusColHeader);

      if (editedCol === (statusColIndex + 1) && String(newValue).trim().toLowerCase() === String(triggerStatusValue).trim().toLowerCase()) {
        registrarLog("onEdit: Doc Gen", `Status gatilho "${triggerStatusValue}" detectado na linha ${editedRow} da aba MI.`, "INFO");
        // Chamar a função para processar a geração do Doc para esta linha específica
        // Passamos a linha (1-based index) e a aba
        processarGeracaoDocParaLinha(editedRow, sheet);
      }
    }
  } catch (error) {
    registrarLog("onEdit: Doc Gen", `Erro: ${error.message} ${error.stack}`, "ERROR");
  }

  // --- Gatilho: Atualizar/Criar Evento na Agenda por Mudança de Data Final (a ser implementado) ---
  // try {
  //   const miSheetNameCal = getConfigValue(PROP_KEYS.MI_SHEET_NAME); // Pode ser a mesma aba MI
  //   const dateColHeaderCal = getConfigValue(PROP_KEYS.CALENDAR_DATE_COL_HEADER);
  //
  //   if (sheet.getName() === miSheetNameCal && editedRow > 1 && dateColHeaderCal) {
  //     const headerRowValuesCal = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  //     const dateColIndexCal = findColumnIndexByHeader(headerRowValuesCal, dateColHeaderCal);
  //
  //     if (editedCol === (dateColIndexCal + 1)) {
  //       registrarLog("onEdit: Calendar", `Data final alterada na linha ${editedRow} da aba ${miSheetNameCal}.`, "INFO");
  //       // Chamar função para atualizar/criar evento na agenda para esta linha
  //       // processarAtualizacaoEventoAgendaParaLinha(editedRow, sheet); // Função a ser criada
  //     }
  //   }
  // } catch (error) {
  //   registrarLog("onEdit: Calendar", `Erro: ${error.message} ${error.stack}`, "ERROR");
  // }

} // Fim da função onEdit


// --- NOVA Função Auxiliar para Gerar Doc para UMA Linha Específica ---
/**
 * Processa a geração de um Google Doc para uma linha específica da aba MI.
 * @param {number} rowIndex O número da linha (1-based) a ser processada.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet A objeto da aba MI.
 */
function processarGeracaoDocParaLinha(rowIndex, sheet) {
  registrarLog("Gerar Doc para Linha", `Iniciando para linha ${rowIndex} na aba "${sheet.getName()}"`, "INFO");
  const ui = SpreadsheetApp.getUi(); // Para alertas, se necessário em caso de falha crítica

  // Carregar todas as configurações de Docs
  const statusColHeader = getConfigValue(PROP_KEYS.DOC_STATUS_COL_HEADER); // Usado para reconfirmar, mas o onEdit já checou
  const triggerStatusValue = getConfigValue(PROP_KEYS.DOC_TRIGGER_STATUS_VALUE);
  const templateId = getConfigValue(PROP_KEYS.DOC_TEMPLATE_ID);
  const saveFolderId = getConfigValue(PROP_KEYS.DOC_SAVE_FOLDER_ID);
  const miIdColHeader = getConfigValue(PROP_KEYS.MI_ID_COL_HEADER);
  const includeColsStr = getConfigValue(PROP_KEYS.DOC_INCLUDE_COLS, '[]');
  let includeCols = []; try { includeCols = JSON.parse(includeColsStr); if(!Array.isArray(includeCols)) includeCols = []; } catch(e){ includeCols = [];}
  const docCreatedColHeader = "Doc Gerado?"; // Conforme definido em gerarDocsDaAbaMI

  // Validação de Configuração Essencial
  if (!saveFolderId || !miIdColHeader || includeCols.length === 0 || !statusColHeader || !triggerStatusValue) {
    registrarLog("Gerar Doc para Linha", `Configuração para geração de Doc incompleta. Linha ${rowIndex}.`, "ERROR");
    return; // Não pode prosseguir
  }

  const headerRowValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  const miIdColIndex = findColumnIndexByHeader(headerRowValues, miIdColHeader);
  const statusColIndex = findColumnIndexByHeader(headerRowValues, statusColHeader); // Para reconfirmar
  let docCreatedColIndex = findColumnIndexByHeader(headerRowValues, docCreatedColHeader);

  // Adicionar coluna "Doc Gerado?" se não existir
  if (docCreatedColIndex === -1) {
      sheet.insertColumnAfter(headerRowValues.length);
      docCreatedColIndex = headerRowValues.length; // Novo índice (0-based)
      sheet.getRange(1, docCreatedColIndex + 1).setValue(docCreatedColHeader).setFontWeight('bold');
      // Re-ler header e dados da linha pois a estrutura da aba mudou
      // headerRowValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; // Opcional, mas mais seguro
      // rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0]; // Opcional
      registrarLog("Gerar Doc para Linha", `Coluna "${docCreatedColHeader}" adicionada.`, "INFO");
  }


  const miId = String(rowValues[miIdColIndex]).trim();
  const currentStatus = String(rowValues[statusColIndex]).trim();
  const alreadyCreated = docCreatedColIndex !== -1 ? String(rowValues[docCreatedColIndex]).trim().toUpperCase() : '';

  if (miId === '') {
    registrarLog("Gerar Doc para Linha", `ID da MI vazio na linha ${rowIndex}. Geração cancelada.`, "WARNING");
    return;
  }
  if (currentStatus.toLowerCase() !== triggerStatusValue.toLowerCase()) {
    registrarLog("Gerar Doc para Linha", `Status na linha ${rowIndex} ("${currentStatus}") não é o gatilho ("${triggerStatusValue}"). Geração cancelada.`, "INFO");
    return;
  }
  if (alreadyCreated === 'SIM' || alreadyCreated === 'YES') {
    registrarLog("Gerar Doc para Linha", `Documento já marcado como gerado para MI "${miId}" na linha ${rowIndex}.`, "INFO");
    return;
  }

  // Validar se todas as colunas a incluir existem
  const includeIndices = {};
  let missingIncludeCols = [];
  includeCols.forEach(header => {
      const index = findColumnIndexByHeader(headerRowValues, header);
      if (index !== -1) { includeIndices[header] = index; }
      else { missingIncludeCols.push(header); }
  });
   if (missingIncludeCols.length > 0) {
       registrarLog("Gerar Doc para Linha", `Colunas a incluir não encontradas para MI "${miId}": ${missingIncludeCols.join(', ')}.`, "ERROR");
       return;
   }


  // Verificar se a pasta de destino existe
  let saveFolder;
  try { saveFolder = DriveApp.getFolderById(saveFolderId); }
  catch (e) { registrarLog("Gerar Doc para Linha", `Pasta destino Docs (ID: ${saveFolderId}) não encontrada ou acesso negado. MI: ${miId}`, "ERROR"); return; }

  SpreadsheetApp.getActiveSpreadsheet().toast(`Gerando Doc para MI: ${miId}...`, "Automação", 5);

  try {
      let newDoc;
      const docName = `Relatorio MI - ${miId}`;

      if (templateId) {
          try {
              const templateFile = DriveApp.getFileById(templateId);
              newDoc = templateFile.makeCopy(docName, saveFolder);
          } catch (e) {
              registrarLog("Gerar Doc para Linha", `Erro ao copiar template ${templateId} para MI "${miId}": ${e.message}. Criando Doc em branco.`, "WARNING");
              newDoc = DocumentApp.create(docName);
              DriveApp.getFileById(newDoc.getId()).moveTo(saveFolder); // Mover para a pasta correta
          }
      } else {
          newDoc = DocumentApp.create(docName);
          DriveApp.getFileById(newDoc.getId()).moveTo(saveFolder); // Mover para a pasta correta
      }

      const body = newDoc.getBody();
      body.clearContents(); // Limpa conteúdo do template, se houver, ou do doc em branco

      body.appendParagraph(docName).setHeading(DocumentApp.ParagraphHeading.TITLE);
      body.appendParagraph(`Data Geração: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss")}`).setFontSize(9).setItalic(true);
      body.appendParagraph("");

      includeCols.forEach(header => {
          const colIndex = includeIndices[header];
          const value = rowValues[colIndex];
          body.appendParagraph(`${header}:`).setBold(true);
          body.appendParagraph(String(value !== null && value !== undefined ? value : '')); // Garante que valor seja string
          body.appendParagraph("");
      });

      newDoc.saveAndClose();
      registrarLog("Gerar Doc para Linha", `Documento "${docName}" criado com sucesso para MI "${miId}". URL: ${newDoc.getUrl()}`, "INFO");

      // Marcar na planilha
      if (docCreatedColIndex !== -1) {
        sheet.getRange(rowIndex, docCreatedColIndex + 1).setValue("Sim");
      }

      // Opcional: Adicionar link do Doc na coluna de link da MI
      const linkColHeader = getConfigValue(PROP_KEYS.MI_LINK_COL_HEADER);
      if (linkColHeader) {
          const linkColIndex = findColumnIndexByHeader(headerRowValues, linkColHeader);
          if (linkColIndex !== -1) {
              sheet.getRange(rowIndex, linkColIndex + 1).setValue(newDoc.getUrl());
          }
      }
       SpreadsheetApp.getActiveSpreadsheet().toast(`Doc para MI ${miId} gerado!`, "Sucesso", 3);

  } catch (e) {
      registrarLog("Gerar Doc para Linha", `Erro ao criar ou popular Doc para MI "${miId}": ${e.message} ${e.stack}`, "ERROR");
      if (docCreatedColIndex !== -1) {
        try {
            sheet.getRange(rowIndex, docCreatedColIndex + 1).setValue(`Erro ao gerar: ${e.message.substring(0,100)}`).setFontColor("red");
        } catch (sheetError) {Logger.log("Erro ao marcar erro na planilha: " + sheetError); }
      }
      SpreadsheetApp.getActiveSpreadsheet().toast(`Erro ao gerar Doc para MI ${miId}. Verifique os logs.`, "Erro", 5);
  }
}


// --- Modificar a função gerarDocsDaAbaMI para usar a nova função auxiliar ---
/**
 * Gera documentos para todas as MIs na aba MI configurada que atendem ao status gatilho.
 * Chamado pelo item de menu.
 */
function gerarDocsDaAbaMI() {
    registrarLog("Gerar Docs Manual", "Iniciando geração manual de Docs pela aba MI.", "INFO");
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Carregar configurações relevantes
    const miSheetName = getConfigValue(PROP_KEYS.MI_SHEET_NAME);
    const statusColHeader = getConfigValue(PROP_KEYS.DOC_STATUS_COL_HEADER);
    const triggerStatusValue = getConfigValue(PROP_KEYS.DOC_TRIGGER_STATUS_VALUE);
    // Outras configs (templateId, saveFolderId, etc.) serão carregadas por processarGeracaoDocParaLinha

    // Validação de Configuração Essencial para a função de loop
    if (!miSheetName || !statusColHeader || !triggerStatusValue) {
        ui.alert("Configuração Incompleta", "Verifique as configurações para Geração de Documentos (Aba MI, Coluna Status, Valor Gatilho) no menu 'Configurações'.", ui.ButtonSet.OK);
        registrarLog("Gerar Docs Manual", "Configuração para geração de Docs incompleta (aba MI, coluna status, ou valor gatilho).", "ERROR");
        return;
    }

    const sheet = ss.getSheetByName(miSheetName);
    if (!sheet) {
        ui.alert(`Erro: Aba MI configurada ("${miSheetName}") não encontrada.`);
        registrarLog("Gerar Docs Manual", `Aba MI "${miSheetName}" não encontrada.`, "ERROR");
        return;
    }

    const range = sheet.getDataRange();
    const values = range.getValues();
    if (values.length <= 1) {
        ui.alert("Aba MI não contém dados para processar.");
        registrarLog("Gerar Docs Manual", "Aba MI sem dados.", "INFO");
        return;
    }

    const headerRow = values[0];
    const statusColIndex = findColumnIndexByHeader(headerRow, statusColHeader);
    if (statusColIndex === -1) {
        ui.alert(`Erro: Coluna de Status configurada ("${statusColHeader}") não encontrada na aba "${miSheetName}".`);
        registrarLog("Gerar Docs Manual", `Coluna Status "${statusColHeader}" não encontrada.`, "ERROR");
        return;
    }
    // A coluna "Doc Gerado?" será verificada e criada por processarGeracaoDocParaLinha se necessário

    let docsProcessados = 0;
    let docsRealmenteGerados = 0; // Para contar apenas os que não tinham sido gerados ainda

    SpreadsheetApp.getActiveSpreadsheet().toast("Iniciando geração de documentos...", "Progresso Docs", -1);

    for (let i = 1; i < values.length; i++) { // Começa da linha 2 (índice 1)
        const currentRowStatus = String(values[i][statusColIndex]).trim();
        if (currentRowStatus.toLowerCase() === triggerStatusValue.toLowerCase()) {
            // Verifica se já foi gerado antes de chamar o processamento completo
            // Isso otimiza a chamada manual para não reprocessar desnecessariamente
            const docCreatedColHeaderCheck = "Doc Gerado?";
            const docCreatedColIndexCheck = findColumnIndexByHeader(headerRow, docCreatedColHeaderCheck);
            let alreadyCreatedCheck = false;
            if (docCreatedColIndexCheck !== -1) {
                alreadyCreatedCheck = String(values[i][docCreatedColIndexCheck]).trim().toUpperCase() === 'SIM' || String(values[i][docCreatedColIndexCheck]).trim().toUpperCase() === 'YES';
            }

            if (!alreadyCreatedCheck) {
                processarGeracaoDocParaLinha(i + 1, sheet); // i+1 porque rowIndex é 1-based
                docsRealmenteGerados++; // Incrementa apenas se chamou o processamento
            }
            docsProcessados++;
        }
         if (docsProcessados % 5 === 0 && docsProcessados > 0) {
             SpreadsheetApp.getActiveSpreadsheet().toast(`Processando linha ${i+1}/${values.length-1}... ${docsRealmenteGerados} Docs gerados.`, "Progresso Docs", 10);
         }
    }

    SpreadsheetApp.getActiveSpreadsheet().toast("Geração de documentos concluída.", "Progresso Docs", 5);
    if (docsProcessados === 0) {
        ui.alert("Geração de Documentos", "Nenhuma linha encontrada na aba MI com o status gatilho para gerar documentos.", ui.ButtonSet.OK);
         registrarLog("Gerar Docs Manual", "Nenhuma linha com status gatilho.", "INFO");
    } else if (docsRealmenteGerados === 0 && docsProcessados > 0) {
        ui.alert("Geração de Documentos", "Todos os itens com status gatilho já tinham documentos gerados ou IDs MI vazios.", ui.ButtonSet.OK);
        registrarLog("Gerar Docs Manual", "Todos os itens elegíveis já tinham Docs ou IDs vazios.", "INFO");
    } else {
        // A função processarGeracaoDocParaLinha já dá o feedback individual
        // Aqui um resumo geral
        ui.alert("Geração de Documentos", `Verificação concluída. ${docsRealmenteGerados} novo(s) documento(s) foram processados para geração. Verifique a aba MI e os logs para detalhes.`, ui.ButtonSet.OK);
        registrarLog("Gerar Docs Manual", `${docsRealmenteGerados} novo(s) documento(s) processados.`, "INFO");
    }
}

// --- ATUALIZAR FUNÇÕES EXISTENTES PARA USAR registrarLog ---
// Exemplo: modificar suas funções de pesquisa, filtro, etc.
// Lembre-se de adicionar registrarLog() no início, fim e em blocos catch.
// Vou mostrar um exemplo com `iniciarPesquisaComPrompt` e `pesquisaExpandida`

function iniciarPesquisaComPrompt() {
  registrarLog("Pesquisa por Prompt", "Iniciada", "INFO");
  const ui = SpreadsheetApp.getUi();
  const resposta = ui.prompt("Pesquisar dados (Múltiplos Termos)", "Digite os termos separados por vírgula:", ui.ButtonSet.OK_CANCEL);
  if (resposta.getSelectedButton() === ui.Button.OK) {
    const termos = resposta.getResponseText().trim();
    if (termos) {
      registrarLog("Pesquisa por Prompt", `Termos: ${termos}`, "INFO");
      pesquisaExpandida(termos);
    } else {
      registrarLog("Pesquisa por Prompt", "Nenhum termo informado.", "INFO");
      ui.alert("Nenhum termo informado.");
    }
  } else {
    registrarLog("Pesquisa por Prompt", "Cancelada pelo usuário.", "INFO");
  }
}

function pesquisaExpandida(termosProcurados) {
   registrarLog("Pesquisa Expandida", `Iniciada com termos: "${termosProcurados}"`, "INFO");
   const planilha = SpreadsheetApp.getActiveSpreadsheet(); const abas = planilha.getSheets();
   const termosLower = termosProcurados.toLowerCase().split(",").map(t => t.trim()).filter(t => t.length > 0);
   if (termosLower.length === 0) {
     registrarLog("Pesquisa Expandida", "Nenhum termo válido informado.", "WARNING");
     SpreadsheetApp.getUi().alert("Nenhum termo válido informado.");
     return;
   }
   const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
   const abaResultadosNome = `${GLOBAL_CONFIG.RESULT_SHEET_BASE}_Adv_${timestamp}`;
   let abaResultados;
   try {
        abaResultados = planilha.insertSheet(abaResultadosNome);
   } catch (e) {
        const shortTs = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMddHHmmss");
        const safeName = abaResultadosNome.substring(0, 70) + shortTs; // Truncate if too long
        abaResultados = planilha.insertSheet(safeName);
        registrarLog("Pesquisa Expandida", `Nome da aba de resultados truncado para: ${safeName}`, "WARNING");
   }

   let linhaAtual = 1; let totalResultados = 0;
   abaResultados.getRange(linhaAtual, 1).setValue(`Pesquisa Avançada por: ${termosProcurados}`).setFontWeight('bold'); linhaAtual += 2;

   try {
       for (const aba of abas) {
         // ... (resto da lógica de pesquisaExpandida como antes) ...
         // Lembre-se de adicionar try-catch em loops internos se fizerem muitas chamadas de API
       } // End sheet loop

       if (totalResultados === 0) {
         abaResultados.getRange(1, 1).setValue(`Nenhuma ocorrência encontrada para: ${termosProcurados}`);
         SpreadsheetApp.getUi().alert("Nenhuma ocorrência encontrada.");
         registrarLog("Pesquisa Expandida", "Nenhuma ocorrência encontrada.", "INFO");
       } else {
         abaResultados.autoResizeColumns(1, abaResultados.getLastColumn());
         SpreadsheetApp.setActiveSheet(abaResultados);
         SpreadsheetApp.getUi().alert(`Pesquisa concluída com ${totalResultados} resultados na aba "${abaResultados.getName()}".`);
         registrarLog("Pesquisa Expandida", `${totalResultados} resultados encontrados na aba "${abaResultados.getName()}".`, "INFO");
         if (getConfigValueBoolean(PROP_KEYS.ENVIAR_EMAIL, false)) {
           enviarResultadosPorEmail(planilha, abaResultados, termosProcurados);
         }
       }
   } catch (e) {
       registrarLog("Pesquisa Expandida", `Erro durante a execução: ${e.message} ${e.stack}`, "ERROR");
       SpreadsheetApp.getUi().alert(`Ocorreu um erro durante a pesquisa: ${e.message}`);
   }
}
// --- NOVAS Funções para Configuração do Dashboard ---

function abrirDialogoConfigDashboard() {
    const html = HtmlService.createHtmlOutputFromFile('configDashboard')
        .setWidth(500).setHeight(550); // Ajustar tamanho
    SpreadsheetApp.getUi().showModalDialog(html, 'Configurar Dashboard de Resumo');
}

function salvarConfigDashboard(config) {
    try {
        setConfigValue(PROP_KEYS.DASHBOARD_ENABLE, String(config.enableDashboard === true || config.enableDashboard === 'true'));
        setConfigValue(PROP_KEYS.DASHBOARD_SHEET_NAME, config.dashboardSheetName ? config.dashboardSheetName.trim() : "Dashboard");
        setConfigValue(PROP_KEYS.DASHBOARD_SOURCE_SHEET_NAME, config.sourceSheetName ? config.sourceSheetName.trim() : getConfigValue(PROP_KEYS.MI_SHEET_NAME)); // Default para aba MI
        setConfigValue(PROP_KEYS.DASHBOARD_STATUS_COL_HEADER, config.statusCol ? config.statusCol.trim() : null);
        setConfigValue(PROP_KEYS.DASHBOARD_RESPONSIBLE_COL_HEADER, config.responsibleCol ? config.responsibleCol.trim() : null);
        setConfigValue(PROP_KEYS.DASHBOARD_DEADLINE_COL_HEADER, config.deadlineCol ? config.deadlineCol.trim() : null);
        setConfigValue(PROP_KEYS.DASHBOARD_ITEM_ID_COL_HEADER, config.itemIdCol ? config.itemIdCol.trim() : null); // Para listar prazos

        registrarLog("Config Dashboard", "Configurações salvas", "INFO");
        return {success: true, message: "Configurações do Dashboard salvas!"};
    } catch (e) {
         Logger.log(`Error saving Dashboard config: ${e}`);
         registrarLog("Config Dashboard", `Erro ao salvar: ${e.message}`, "ERROR");
         return {success: false, message: `Erro ao salvar: ${e.message}`};
    }
}

function carregarConfigDashboard() {
    return {
        enableDashboard: getConfigValueBoolean(PROP_KEYS.DASHBOARD_ENABLE, true), // Habilitado por padrão se configurado
        dashboardSheetName: getConfigValue(PROP_KEYS.DASHBOARD_SHEET_NAME, 'Dashboard'),
        sourceSheetName: getConfigValue(PROP_KEYS.DASHBOARD_SOURCE_SHEET_NAME, getConfigValue(PROP_KEYS.MI_SHEET_NAME, '')), // Tenta pegar da MI_SHEET_NAME se não específico
        statusCol: getConfigValue(PROP_KEYS.DASHBOARD_STATUS_COL_HEADER, ''),
        responsibleCol: getConfigValue(PROP_KEYS.DASHBOARD_RESPONSIBLE_COL_HEADER, ''),
        deadlineCol: getConfigValue(PROP_KEYS.DASHBOARD_DEADLINE_COL_HEADER, ''),
        itemIdCol: getConfigValue(PROP_KEYS.DASHBOARD_ITEM_ID_COL_HEADER, '')
    };
}

// --- NOVA Função para Atualizar o Dashboard ---
/**
 * Gera ou atualiza a aba do Dashboard com dados resumidos.
 */
function atualizarDashboard() {
    if (!getConfigValueBoolean(PROP_KEYS.DASHBOARD_ENABLE, false)) {
        SpreadsheetApp.getUi().alert("Dashboard Desabilitado", "A funcionalidade de Dashboard está desabilitada nas configurações.", SpreadsheetApp.getUi().ButtonSet.OK);
        registrarLog("Dashboard", "Tentativa de atualização com dashboard desabilitado.", "WARNING");
        return;
    }
    registrarLog("Dashboard", "Iniciando atualização...", "INFO");
    SpreadsheetApp.getActiveSpreadsheet().toast("Atualizando Dashboard...", "Progresso", -1);

    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Carregar configurações
    const dashboardSheetName = getConfigValue(PROP_KEYS.DASHBOARD_SHEET_NAME, "Dashboard");
    const sourceSheetName = getConfigValue(PROP_KEYS.DASHBOARD_SOURCE_SHEET_NAME);
    const statusColHeader = getConfigValue(PROP_KEYS.DASHBOARD_STATUS_COL_HEADER);
    const responsibleColHeader = getConfigValue(PROP_KEYS.DASHBOARD_RESPONSIBLE_COL_HEADER);
    const deadlineColHeader = getConfigValue(PROP_KEYS.DASHBOARD_DEADLINE_COL_HEADER);
    const itemIdColHeader = getConfigValue(PROP_KEYS.DASHBOARD_ITEM_ID_COL_HEADER); // Para listagem de prazos

    // Validação de Configuração
    if (!sourceSheetName || !statusColHeader || !responsibleColHeader || !deadlineColHeader || !itemIdColHeader) {
        ui.alert("Configuração Incompleta", "Configure a Aba de Origem e os Cabeçalhos das Colunas de Status, Responsável, Data Final e ID do Item para o Dashboard.\nUse o menu 'Configurações'.", ui.ButtonSet.OK);
        registrarLog("Dashboard", "Configuração incompleta (aba origem ou colunas não definidas).", "ERROR");
        SpreadsheetApp.getActiveSpreadsheet().toast("Falha na configuração!", "Erro", 5);
        return;
    }

    const sourceSheet = ss.getSheetByName(sourceSheetName);
    if (!sourceSheet) {
        ui.alert("Erro", `Aba de origem "${sourceSheetName}" não encontrada.`, ui.ButtonSet.OK);
        registrarLog("Dashboard", `Aba de origem "${sourceSheetName}" não encontrada.`, "ERROR");
        SpreadsheetApp.getActiveSpreadsheet().toast("Falha!", "Erro", 5);
        return;
    }

    const dataRange = sourceSheet.getDataRange();
    const values = dataRange.getValues();
    if (values.length <= 1) {
        ui.alert("Sem Dados", `A aba de origem "${sourceSheetName}" não contém dados.`, ui.ButtonSet.OK);
        registrarLog("Dashboard", `Aba de origem "${sourceSheetName}" sem dados.`, "INFO");
        SpreadsheetApp.getActiveSpreadsheet().toast("Sem dados para processar.", "Aviso", 5);
        return;
    }

    const headerRow = values[0];
    const statusColIndex = findColumnIndexByHeader(headerRow, statusColHeader);
    const responsibleColIndex = findColumnIndexByHeader(headerRow, responsibleColHeader);
    const deadlineColIndex = findColumnIndexByHeader(headerRow, deadlineColHeader);
    const itemIdColIndex = findColumnIndexByHeader(headerRow, itemIdColHeader);

    let missingHeadersLog = [];
    if (statusColIndex === -1) missingHeadersLog.push(statusColHeader);
    if (responsibleColIndex === -1) missingHeadersLog.push(responsibleColHeader);
    if (deadlineColIndex === -1) missingHeadersLog.push(deadlineColHeader);
    if (itemIdColIndex === -1) missingHeadersLog.push(itemIdColHeader);

    if (missingHeadersLog.length > 0) {
        ui.alert("Erro de Configuração", `As seguintes colunas não foram encontradas na aba "${sourceSheetName}": ${missingHeadersLog.join(', ')}. Verifique os nomes dos cabeçalhos.`, ui.ButtonSet.OK);
        registrarLog("Dashboard", `Colunas não encontradas na aba origem: ${missingHeadersLog.join(', ')}.`, "ERROR");
        SpreadsheetApp.getActiveSpreadsheet().toast("Falha na configuração!", "Erro", 5);
        return;
    }

    // Preparar aba do Dashboard
    let dashboardSheet = ss.getSheetByName(dashboardSheetName);
    if (!dashboardSheet) {
        dashboardSheet = ss.insertSheet(dashboardSheetName);
    }
    dashboardSheet.clearContents().clearFormats(); // Limpa antes de popular
    dashboardSheet.getRange("A1").setValue(`Dashboard de Resumo - Atualizado em: ${new Date().toLocaleString()}`).setFontWeight("bold").setFontSize(14);
    let currentRow = 3; // Linha inicial para os dados do dashboard

    // 1. Contagem por Status
    const statusCounts = {};
    for (let i = 1; i < values.length; i++) {
        const status = String(values[i][statusColIndex]).trim();
        if (status) {
            statusCounts[status] = (statusCounts[status] || 0) + 1;
        }
    }
    dashboardSheet.getRange(currentRow, 1).setValue("Contagem por Status:").setFontWeight("bold");
    currentRow++;
    dashboardSheet.getRange(currentRow, 1, 1, 2).setValues([["Status", "Contagem"]]).setFontWeight("bold");
    currentRow++;
    for (const status in statusCounts) {
        dashboardSheet.getRange(currentRow, 1, 1, 2).setValues([[status, statusCounts[status]]]);
        currentRow++;
    }
    currentRow += 2; // Espaçamento

    // 2. Contagem por Responsável
    const responsibleCounts = {};
    for (let i = 1; i < values.length; i++) {
        const responsible = String(values[i][responsibleColIndex]).trim();
        if (responsible) {
            responsibleCounts[responsible] = (responsibleCounts[responsible] || 0) + 1;
        }
    }
    dashboardSheet.getRange(currentRow, 1).setValue("Contagem por Responsável:").setFontWeight("bold");
    currentRow++;
    dashboardSheet.getRange(currentRow, 1, 1, 2).setValues([["Responsável", "Contagem"]]).setFontWeight("bold");
    currentRow++;
    for (const responsible in responsibleCounts) {
        dashboardSheet.getRange(currentRow, 1, 1, 2).setValues([[responsible, responsibleCounts[responsible]]]);
        currentRow++;
    }
    currentRow += 2; // Espaçamento

    // 3. Prazos Vencendo na Próxima Semana (incluindo hoje)
    dashboardSheet.getRange(currentRow, 1).setValue("Prazos Vencendo nos Próximos 7 Dias:").setFontWeight("bold");
    currentRow++;
    dashboardSheet.getRange(currentRow, 1, 1, 3).setValues([["ID do Item", "Responsável", "Data Prazo"]]).setFontWeight("bold"); // Adicionado Responsável
    currentRow++;
    const hoje = new Date(); hoje.setHours(0,0,0,0);
    const umaSemanaDepois = new Date(hoje);
    umaSemanaDepois.setDate(hoje.getDate() + 7);

    let prazosEncontrados = 0;
    for (let i = 1; i < values.length; i++) {
        const deadlineStr = values[i][deadlineColIndex];
        const itemId = String(values[i][itemIdColIndex]).trim();
        const responsavelItem = String(values[i][responsibleColIndex]).trim(); // Pega o responsável
        if (!deadlineStr || !itemId) continue;

        let deadlineDate;
        try {
            deadlineDate = new Date(deadlineStr);
            if (isNaN(deadlineDate.getTime())) continue;
            deadlineDate.setHours(0,0,0,0);
        } catch (e) { continue; }

        if (deadlineDate >= hoje && deadlineDate <= umaSemanaDepois) {
            // Tenta obter o valor formatado da data, se possível, senão usa toLocaleDateString
            let displayDate = Utilities.formatDate(deadlineDate, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
            try {
                const cellDateFormat = sourceSheet.getRange(i + 1, deadlineColIndex + 1).getNumberFormat();
                if (cellDateFormat && cellDateFormat.match(/d|M|y/)) { // Checa se é um formato de data
                     displayDate = sourceSheet.getRange(i + 1, deadlineColIndex + 1).getDisplayValue();
                }
            } catch(e) { /* Usa o formatado por Utilities */ }

            dashboardSheet.getRange(currentRow, 1, 1, 3).setValues([[itemId, responsavelItem, displayDate]]);
            currentRow++;
            prazosEncontrados++;
        }
    }
    if (prazosEncontrados === 0) {
        dashboardSheet.getRange(currentRow, 1).setValue("Nenhum item com prazo nos próximos 7 dias.").setFontStyle("italic");
        currentRow++;
    }

    // Ajustar colunas
    dashboardSheet.autoResizeColumn(1);
    dashboardSheet.autoResizeColumn(2);
    dashboardSheet.autoResizeColumn(3);

    SpreadsheetApp.setActiveSheet(dashboardSheet); // Torna a aba do dashboard ativa
    SpreadsheetApp.getActiveSpreadsheet().toast("Dashboard atualizado!", "Concluído", 5);
    registrarLog("Dashboard", "Dashboard atualizado com sucesso.", "INFO");
}
// --- NOVAS Funções para Configuração de Validação de Dados ---

function abrirDialogoConfigValidacao() {
    const html = HtmlService.createHtmlOutputFromFile('configValidacaoDados')
        .setWidth(600).setHeight(550); // Ajustar tamanho
    SpreadsheetApp.getUi().showModalDialog(html, 'Configurar Validação de Dados');
}

/**
 * Carrega as regras de validação de dados salvas.
 * @return {Array<Object>} Um array de objetos de regra de validação.
 */
function carregarRegrasValidacaoDados() {
    const rulesJson = getConfigValue(PROP_KEYS.DATA_VALIDATION_RULES, '[]');
    let rules = [];
    try {
        rules = JSON.parse(rulesJson);
        if (!Array.isArray(rules)) {
            rules = [];
        }
    } catch (e) {
        Logger.log(`Erro ao parsear regras de validação: ${e}`);
        rules = [];
    }
    // Adicionar um ID temporário para cada regra para facilitar a remoção no cliente
    return rules.map((rule, index) => ({ ...rule, id: index }));
}

/**
 * Salva uma nova regra de validação de dados.
 * @param {Object} newRule O objeto da nova regra de validação.
 * @return {Object} Resultado da operação.
 */
function salvarNovaRegraValidacaoDados(newRule) {
    try {
        if (!newRule || !newRule.sheetName || !newRule.columnHeader || !newRule.type) {
            throw new Error("Dados da regra incompletos.");
        }
        // Validação específica para o tipo 'LIST_FROM_RANGE'
        if (newRule.type === 'LIST_FROM_RANGE' && (!newRule.criteria || !newRule.criteria.sourceRange)) {
            throw new Error("Para 'Lista de um Intervalo', o intervalo de origem é obrigatório.");
        }
        // Outras validações de tipo podem ser adicionadas aqui

        const rules = carregarRegrasValidacaoDados().map(rule => { // Recarrega sem IDs temporários
             const { id, ...rest } = rule; // Remove o ID temporário antes de salvar
             return rest;
        });

        rules.push(newRule);
        setConfigValue(PROP_KEYS.DATA_VALIDATION_RULES, JSON.stringify(rules));
        registrarLog("Config Validação", `Nova regra salva para: ${newRule.sheetName} -> ${newRule.columnHeader}`, "INFO");
        return { success: true, message: "Nova regra de validação salva!", rules: carregarRegrasValidacaoDados() };
    } catch (e) {
        Logger.log(`Erro ao salvar nova regra de validação: ${e}`);
        registrarLog("Config Validação", `Erro ao salvar nova regra: ${e.message}`, "ERROR");
        return { success: false, message: `Erro ao salvar: ${e.message}` };
    }
}

/**
 * Remove uma regra de validação de dados pelo seu ID (índice).
 * @param {number} ruleId O ID (índice no array) da regra a ser removida.
 * @return {Object} Resultado da operação.
 */
function removerRegraValidacaoDados(ruleId) {
    try {
        let rules = carregarRegrasValidacaoDados(); // Estas regras têm IDs temporários
        const originalLength = rules.length;
        rules = rules.filter(rule => rule.id !== ruleId); // Filtra pelo ID temporário

        if (rules.length === originalLength) {
            throw new Error("Regra não encontrada para remoção com o ID: " + ruleId);
        }
        
        // Mapeia de volta para o formato sem ID antes de salvar
        const rulesToSave = rules.map(rule => {
            const { id, ...rest } = rule;
            return rest;
        });

        setConfigValue(PROP_KEYS.DATA_VALIDATION_RULES, JSON.stringify(rulesToSave));
        registrarLog("Config Validação", `Regra com ID ${ruleId} removida.`, "INFO");
        return { success: true, message: "Regra de validação removida!", rules: carregarRegrasValidacaoDados() };
    } catch (e) {
        Logger.log(`Erro ao remover regra de validação: ${e}`);
        registrarLog("Config Validação", `Erro ao remover regra: ${e.message}`, "ERROR");
        return { success: false, message: `Erro ao remover: ${e.message}` };
    }
}


// --- NOVA Função para Aplicar as Regras de Validação ---
/**
 * Aplica todas as regras de validação de dados configuradas.
 */
function aplicarTodasRegrasDeValidacao() {
    registrarLog("Validação de Dados", "Iniciando aplicação de todas as regras.", "INFO");
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allRules = carregarRegrasValidacaoDados(); // Carrega regras (já não tem IDs aqui)

    if (allRules.length === 0) {
        ui.alert("Nenhuma regra de validação configurada.");
        registrarLog("Validação de Dados", "Nenhuma regra configurada para aplicar.", "INFO");
        return;
    }

    let rulesAppliedCount = 0;
    let errorsCount = 0;

    SpreadsheetApp.getActiveSpreadsheet().toast(`Aplicando ${allRules.length} regra(s) de validação...`, "Progresso", -1);

    allRules.forEach((ruleConfig, index) => {
        try {
            const sheet = ss.getSheetByName(ruleConfig.sheetName);
            if (!sheet) {
                throw new Error(`Aba "${ruleConfig.sheetName}" não encontrada.`);
            }

            const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
            const columnIndex = findColumnIndexByHeader(headerRow, ruleConfig.columnHeader);

            if (columnIndex === -1) {
                throw new Error(`Coluna "${ruleConfig.columnHeader}" não encontrada na aba "${ruleConfig.sheetName}".`);
            }

            // Define o intervalo da coluna inteira, exceto o cabeçalho
            const targetRange = sheet.getRange(2, columnIndex + 1, sheet.getMaxRows() - 1, 1);
            targetRange.clearDataValidations(); // Limpa validações existentes na coluna

            let validationRuleBuilder;

            // --- Construir a regra de validação ---
            if (ruleConfig.type === 'LIST_FROM_RANGE') {
                if (!ruleConfig.criteria || !ruleConfig.criteria.sourceRange) {
                    throw new Error("Critério 'sourceRange' faltando para LIST_FROM_RANGE.");
                }
                const sourceRangeString = ruleConfig.criteria.sourceRange;
                let sourceRangeValues;
                try {
                    sourceRangeValues = ss.getRange(sourceRangeString); // Tenta obter como intervalo nomeado ou A1 Notação
                } catch (e) {
                     throw new Error(`Intervalo de origem "${sourceRangeString}" inválido ou não encontrado.`);
                }
                validationRuleBuilder = SpreadsheetApp.newDataValidation().requireValueInRange(sourceRangeValues, true); // true para mostrar dropdown
            }
            // Adicionar outros tipos de validação aqui (ex: NUMBER_BETWEEN, TEXT_CONTAINS, etc.)
            // else if (ruleConfig.type === 'NUMBER_BETWEEN') { ... }
            else {
                throw new Error(`Tipo de validação desconhecido: "${ruleConfig.type}"`);
            }

            // Configurações comuns
            validationRuleBuilder.setAllowInvalid(ruleConfig.allowInvalid === true || ruleConfig.allowInvalid === 'true');
            if (ruleConfig.helpText) {
                validationRuleBuilder.setHelpText(ruleConfig.helpText);
            }

            targetRange.setDataValidation(validationRuleBuilder.build());
            rulesAppliedCount++;
            registrarLog("Validação de Dados", `Regra aplicada: ${ruleConfig.sheetName} -> ${ruleConfig.columnHeader} (Tipo: ${ruleConfig.type})`, "INFO");

        } catch (e) {
            Logger.log(`Erro ao aplicar regra ${index + 1} (${ruleConfig.sheetName} -> ${ruleConfig.columnHeader}): ${e.message}`);
            registrarLog("Validação de Dados", `Erro regra ${ruleConfig.sheetName} -> ${ruleConfig.columnHeader}: ${e.message}`, "ERROR");
            errorsCount++;
        }
         SpreadsheetApp.getActiveSpreadsheet().toast(`Processando regra ${index + 1}/${allRules.length}...`, "Progresso Validação", 5);
         Utilities.sleep(100); // Pequena pausa
    });

    SpreadsheetApp.getActiveSpreadsheet().toast("Validação concluída.", "Progresso", 5);
    let summaryMessage = `${rulesAppliedCount} regra(s) de validação aplicada(s) com sucesso.`;
    if (errorsCount > 0) {
        summaryMessage += `\n${errorsCount} regra(s) falharam ao aplicar (verifique os Logs do script).`;
    }
    ui.alert("Resultado da Aplicação de Validação", summaryMessage, ui.ButtonSet.OK);
}
// Em Code.gs

// --- Função onEdit para automações ---
/**
 * Executa automaticamente quando um usuário edita uma célula na planilha.
 * @param {Object} e O objeto de evento passado pelo Google Apps Script.
 */
function onEdit(e) {
  if (!e) {
    registrarLog("onEdit", "Evento de edição não recebido. Ignorando.", "WARNING");
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();
  const editedRow = range.getRow();
  const editedCol = range.getColumn();
  const newValue = e.value;
  const oldValue = e.oldValue;

  // Não registrar cada edição simples para não poluir o log, a menos que seja relevante.
  // registrarLog("onEdit Disparado", `Usuário: ${Session.getActiveUser().getEmail()}, Aba: ${sheet.getName()}, Célula: ${range.getA1Notation()}, Novo Valor: ${newValue}`, "INFO");

  // --- Gatilho: Gerar Documento por Mudança de Status na Aba MI ---
  try {
    const miSheetNameDoc = getConfigValue(PROP_KEYS.MI_SHEET_NAME); // Assumindo que é a mesma aba MI
    const statusColHeaderDoc = getConfigValue(PROP_KEYS.DOC_STATUS_COL_HEADER);
    const triggerStatusValueDoc = getConfigValue(PROP_KEYS.DOC_TRIGGER_STATUS_VALUE);

    if (sheet.getName() === miSheetNameDoc && editedRow > 1 && statusColHeaderDoc && triggerStatusValueDoc) {
      const headerRowValuesDoc = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const statusColIndexDoc = findColumnIndexByHeader(headerRowValuesDoc, statusColHeaderDoc);

      if (editedCol === (statusColIndexDoc + 1) && String(newValue).trim().toLowerCase() === String(triggerStatusValueDoc).trim().toLowerCase()) {
        registrarLog("onEdit: Doc Gen Trigger", `Status gatilho "${triggerStatusValueDoc}" detectado linha ${editedRow} aba MI.`, "INFO");
        processarGeracaoDocParaLinha(editedRow, sheet);
      }
    }
  } catch (error) {
    registrarLog("onEdit: Doc Gen Trigger", `Erro: ${error.message} ${error.stack}`, "ERROR");
  }

  // --- NOVO Gatilho: Atualizar/Criar Evento na Agenda por Mudança de Data Final ---
  try {
    const miSheetNameCal = getConfigValue(PROP_KEYS.MI_SHEET_NAME); // Usaremos a aba MI configurada
    const dateColHeaderCal = getConfigValue(PROP_KEYS.CALENDAR_DATE_COL_HEADER);
    const eventIdColHeaderCal = getConfigValue(PROP_KEYS.CALENDAR_EVENT_ID_COL_HEADER); // Essencial para update/delete

    if (sheet.getName() === miSheetNameCal && editedRow > 1 && dateColHeaderCal && eventIdColHeaderCal) {
      const headerRowValuesCal = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const dateColIndexCal = findColumnIndexByHeader(headerRowValuesCal, dateColHeaderCal);

      if (editedCol === (dateColIndexCal + 1)) { // Se a coluna editada foi a de data
        registrarLog("onEdit: Calendar Trigger", `Data da Agenda alterada na linha ${editedRow} da aba ${miSheetNameCal}. Novo valor: ${newValue}`, "INFO");
        processarEventoAgendaParaLinha(editedRow, sheet, newValue, oldValue);
      }
    }
  } catch (error) {
    registrarLog("onEdit: Calendar Trigger", `Erro: ${error.message} ${error.stack}`, "ERROR");
  }
} // Fim da função onEdit


// --- NOVA Função Auxiliar para Processar Evento da Agenda para UMA Linha Específica ---
/**
 * Cria, atualiza ou deleta um evento na agenda baseado na edição de data em uma linha.
 * @param {number} rowIndex O número da linha (1-based) que foi editada.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet O objeto da aba.
 * @param {any} newValue O novo valor da célula de data.
 * @param {any} oldValue O valor antigo da célula de data (para referência, não usado ativamente aqui).
 */
function processarEventoAgendaParaLinha(rowIndex, sheet, newValue, oldValue) {
  registrarLog("Processar Evento Agenda", `Iniciando para linha ${rowIndex}, Aba: ${sheet.getName()}, Novo Valor Data: ${newValue}`, "INFO");

  // Carregar configurações da Agenda
  const titleColHeader = getConfigValue(PROP_KEYS.CALENDAR_TITLE_COL_HEADER) || getConfigValue(PROP_KEYS.MI_ID_COL_HEADER); // Título do Evento
  const calendarId = getConfigValue(PROP_KEYS.CALENDAR_ID, 'primary');
  const eventIdColHeader = getConfigValue(PROP_KEYS.CALENDAR_EVENT_ID_COL_HEADER);
  const miIdColHeaderEvent = getConfigValue(PROP_KEYS.MI_ID_COL_HEADER); // Para descrição do evento

  if (!titleColHeader || !eventIdColHeader || !miIdColHeaderEvent) {
    registrarLog("Processar Evento Agenda", "Configuração da Agenda incompleta (Título, Coluna ID Evento, ou Coluna ID MI não definida).", "ERROR");
    return;
  }

  const headerRowValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  const titleColIndex = findColumnIndexByHeader(headerRowValues, titleColHeader);
  let eventIdColIndex = findColumnIndexByHeader(headerRowValues, eventIdColHeader);
  const miIdColIndex = findColumnIndexByHeader(headerRowValues, miIdColHeader);


  if (titleColIndex === -1) {
    registrarLog("Processar Evento Agenda", `Coluna de Título "${titleColHeader}" não encontrada. Linha ${rowIndex}.`, "ERROR");
    return;
  }
  if (miIdColIndex === -1) {
    registrarLog("Processar Evento Agenda", `Coluna ID MI "${miIdColHeaderEvent}" não encontrada para descrição. Linha ${rowIndex}.`, "ERROR");
    // Pode continuar sem isso, mas a descrição do evento será menos informativa
  }


  // Adicionar coluna "ID Evento Agenda" se não existir
  if (eventIdColIndex === -1) {
      sheet.insertColumnAfter(headerRowValues.length);
      eventIdColIndex = headerRowValues.length; // Novo índice (0-based)
      sheet.getRange(1, eventIdColIndex + 1).setValue(eventIdColHeader).setFontWeight('bold');
      // Re-ler valores da linha pois a estrutura mudou
      // rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0]; // Mais seguro, mas pode ser lento no onEdit
      registrarLog("Processar Evento Agenda", `Coluna "${eventIdColHeader}" adicionada.`, "INFO");
      // Como a coluna foi recém-adicionada, o valor do eventId será undefined/blank
  }

  const eventTitle = String(rowValues[titleColIndex] || (miIdColIndex !== -1 ? rowValues[miIdColIndex] : "Evento da Planilha")).trim();
  const miIdentifier = miIdColIndex !== -1 ? String(rowValues[miIdColIndex]).trim() : "N/A";
  const existingEventId = eventIdColIndex !== -1 ? String(rowValues[eventIdColIndex] || '').trim() : '';

  let newEventDate;
  if (newValue && String(newValue).trim() !== "") {
    try {
      newEventDate = new Date(newValue);
      if (isNaN(newEventDate.getTime())) throw new Error("Data inválida");
      //newEventDate.setHours(12, 0, 0, 0); // Normaliza para meio-dia para eventos de dia inteiro, evitar problemas de fuso
    } catch (e) {
      registrarLog("Processar Evento Agenda", `Novo valor de data "${newValue}" é inválido para linha ${rowIndex}.`, "WARNING");
      // Se a nova data é inválida, e existe um evento, podemos optar por deletá-lo
      if (existingEventId) {
        // ... lógica de deleção abaixo ...
      } else {
        return; // Nenhuma ação se a nova data é inválida e não há evento existente
      }
    }
  }

  try {
    const calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      registrarLog("Processar Evento Agenda", `Calendário com ID "${calendarId}" não encontrado ou acesso negado.`, "ERROR");
      return;
    }

    // CASO 1: Nova data é VAZIA ou INVÁLIDA, e um evento EXISTIA
    if ((!newEventDate || isNaN(newEventDate.getTime())) && existingEventId) {
      try {
        const eventToCancel = calendar.getEventById(existingEventId);
        if (eventToCancel) {
          eventToCancel.deleteEvent();
          registrarLog("Processar Evento Agenda", `Evento ID ${existingEventId} (MI: ${miIdentifier}) deletado pois data foi removida/invalidada.`, "INFO");
        }
      } catch (errDel) {
        registrarLog("Processar Evento Agenda", `Erro ao tentar deletar evento ${existingEventId}: ${errDel.message}`, "ERROR");
      }
      if (eventIdColIndex !== -1) sheet.getRange(rowIndex, eventIdColIndex + 1).setValue(''); // Limpa o ID do evento na planilha
      SpreadsheetApp.getActiveSpreadsheet().toast(`Evento para ${eventTitle} removido da agenda.`, "Agenda", 3);
      return;
    }

    // CASO 2: Nova data é VÁLIDA
    if (newEventDate && !isNaN(newEventDate.getTime())) {
      const eventDescription = `Referente à MI/Item: ${miIdentifier}\nGerado/Atualizado pela planilha: ${SpreadsheetApp.getActiveSpreadsheet().getName()}`;

      if (existingEventId) { // ATUALIZAR evento existente
        try {
          let eventToUpdate = calendar.getEventById(existingEventId);
          if (eventToUpdate) {
            eventToUpdate.setTitle(eventTitle); // Atualiza o título caso tenha mudado também
            eventToUpdate.setAllDayDate(newEventDate); // Atualiza para a nova data (dia inteiro)
            // eventToUpdate.setTime(newEventDate, new Date(newEventDate.getTime() + (60*60*1000))); // Para evento de 1h
            eventToUpdate.setDescription(eventDescription);
            registrarLog("Processar Evento Agenda", `Evento ID ${existingEventId} (MI: ${miIdentifier}) atualizado para ${newEventDate.toLocaleDateString()}.`, "INFO");
            SpreadsheetApp.getActiveSpreadsheet().toast(`Evento para ${eventTitle} atualizado na agenda.`, "Agenda", 3);
          } else { // Event ID na planilha mas não encontrado na agenda, criar novo
             throw new Error("ID do evento na planilha não encontrado na agenda. Criando novo.");
          }
        } catch (errUpdate) { // Se o evento não existe mais ou erro ao atualizar, tenta criar um novo
          registrarLog("Processar Evento Agenda", `Erro ao atualizar evento ${existingEventId} (MI: ${miIdentifier}), tentando criar novo: ${errUpdate.message}`, "WARNING");
          const newEvent = calendar.createAllDayEvent(eventTitle, newEventDate, {description: eventDescription});
          if (eventIdColIndex !== -1) sheet.getRange(rowIndex, eventIdColIndex + 1).setValue(newEvent.getId());
          registrarLog("Processar Evento Agenda", `Novo evento criado para MI "${miIdentifier}" em ${newEventDate.toLocaleDateString()} (ID: ${newEvent.getId()}).`, "INFO");
          SpreadsheetApp.getActiveSpreadsheet().toast(`Novo evento para ${eventTitle} criado na agenda.`, "Agenda", 3);
        }
      } else { // CRIAR novo evento
        const newEvent = calendar.createAllDayEvent(eventTitle, newEventDate, {description: eventDescription});
        if (eventIdColIndex !== -1) sheet.getRange(rowIndex, eventIdColIndex + 1).setValue(newEvent.getId());
        registrarLog("Processar Evento Agenda", `Novo evento criado para MI "${miIdentifier}" em ${newEventDate.toLocaleDateString()} (ID: ${newEvent.getId()}).`, "INFO");
        SpreadsheetApp.getActiveSpreadsheet().toast(`Evento para ${eventTitle} criado na agenda.`, "Agenda", 3);
      }
    }

  } catch (e) {
    registrarLog("Processar Evento Agenda", `Erro ao interagir com CalendarApp para MI "${miIdentifier}": ${e.message} ${e.stack}`, "ERROR");
     SpreadsheetApp.getActiveSpreadsheet().toast(`Erro ao processar evento para ${eventTitle}. Ver Logs.`, "Erro Agenda", 5);
  }
} // Fim de processarEventoAgendaParaLinha

// --- ATUALIZAR a função criarEventosAgendaDaAbaMI para salvar o Event ID ---
/**
 * Cria eventos na agenda para todos os itens elegíveis na aba MI configurada.
 * Chamado pelo item de menu.
 */
function criarEventosAgendaDaAbaMI() {
    registrarLog("Criar Eventos Manual", "Iniciando criação manual de eventos pela aba MI.", "INFO");
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const miSheetName = getConfigValue(PROP_KEYS.MI_SHEET_NAME);
    const dateColHeader = getConfigValue(PROP_KEYS.CALENDAR_DATE_COL_HEADER);
    const titleColHeader = getConfigValue(PROP_KEYS.CALENDAR_TITLE_COL_HEADER) || getConfigValue(PROP_KEYS.MI_ID_COL_HEADER);
    const calendarId = getConfigValue(PROP_KEYS.CALENDAR_ID, 'primary');
    const eventIdColHeader = getConfigValue(PROP_KEYS.CALENDAR_EVENT_ID_COL_HEADER); // Essencial
    const miIdColHeaderEvent = getConfigValue(PROP_KEYS.MI_ID_COL_HEADER); // Para descrição

    // Validação de Configuração Essencial
    if (!miSheetName || !dateColHeader || !titleColHeader || !eventIdColHeader || !miIdColHeaderEvent) {
        ui.alert("Configuração Incompleta", "Verifique as configurações para Integração com Agenda (Aba MI, Coluna Data, Coluna Título, Coluna ID Evento, Coluna ID MI) no menu 'Configurações'.", ui.ButtonSet.OK);
        registrarLog("Criar Eventos Manual", "Configuração incompleta para criação de eventos.", "ERROR");
        return;
    }

    const sheet = ss.getSheetByName(miSheetName);
    if (!sheet) { ui.alert(`Erro: Aba MI "${miSheetName}" não encontrada.`); registrarLog("Criar Eventos Manual", `Aba MI "${miSheetName}" não encontrada.`, "ERROR"); return; }

    const range = sheet.getDataRange();
    const values = range.getValues();
    if (values.length <= 1) { ui.alert("Aba MI sem dados."); registrarLog("Criar Eventos Manual", "Aba MI sem dados.", "INFO"); return; }

    const headerRow = values[0];
    const dateColIndex = findColumnIndexByHeader(headerRow, dateColHeader);
    const titleColIndex = findColumnIndexByHeader(headerRow, titleColHeader);
    let eventIdColIndex = findColumnIndexByHeader(headerRow, eventIdColHeader);
    const miIdColIndex = findColumnIndexByHeader(headerRow, miIdColHeaderEvent);


    // Validar se colunas foram encontradas
    let missingHeaders = [];
    if (dateColIndex === -1) missingHeaders.push(dateColHeader);
    if (titleColIndex === -1) missingHeaders.push(titleColHeader);
    if (miIdColIndex === -1) missingHeaders.push(miIdColHeaderEvent);
    // eventIdColIndex será criada se não existir

    if (missingHeaders.length > 0) {
         ui.alert(`Erro: As seguintes colunas não foram encontradas na aba "${miSheetName}": ${missingHeaders.join(', ')}.`);
         registrarLog("Criar Eventos Manual", `Colunas não encontradas: ${missingHeaders.join(', ')}.`, "ERROR"); return;
    }

    // Adicionar coluna "ID Evento Agenda" se não existir
    if (eventIdColIndex === -1) {
        sheet.insertColumnAfter(headerRow.length);
        eventIdColIndex = headerRow.length;
        sheet.getRange(1, eventIdColIndex + 1).setValue(eventIdColHeader).setFontWeight('bold');
        // Re-fetch values if column structure changed (important for subsequent .setValue)
        // values = sheet.getDataRange().getValues(); // Ou apenas atualizar o eventIdColIndex para uso futuro
        registrarLog("Criar Eventos Manual", `Coluna "${eventIdColHeader}" adicionada.`, "INFO");
    }


    let eventsCreatedCount = 0; let errorsCount = 0;
    try {
        const calendar = CalendarApp.getCalendarById(calendarId);
        if (!calendar) { ui.alert(`Erro: Calendário ID "${calendarId}" não acessível.`); registrarLog("Criar Eventos Manual", `Calendário ID "${calendarId}" não acessível.`, "ERROR"); return; }

        SpreadsheetApp.getActiveSpreadsheet().toast(`Verificando ${values.length-1} linhas para criar eventos...`, "Progresso Agenda", -1);

        for (let i = 1; i < values.length; i++) {
            const row = values[i];
            const eventDateStr = row[dateColIndex];
            const eventTitle = String(row[titleColIndex] || (miIdColIndex !== -1 ? row[miIdColIndex] : "Evento da Planilha")).trim();
            const miIdentifier = miIdColIndex !== -1 ? String(row[miIdColIndex]).trim() : "N/A";
            const existingEventId = eventIdColIndex !== -1 ? String(row[eventIdColIndex] || '').trim() : '';

            if (!eventDateStr || !eventTitle || existingEventId) { // Pula se não tiver data/título OU se já tiver um ID de evento
                continue;
            }

            let eventDate;
            try { eventDate = new Date(eventDateStr); if (isNaN(eventDate.getTime())) throw new Error("Data inválida"); }
            catch (e) { registrarLog("Criar Eventos Manual", `Data inválida na linha <span class="math-inline">\{i\+1\} \('</span>{eventDateStr}').`, "WARNING"); continue; }

            const eventDescription = `Referente à MI/Item: ${miIdentifier}\nGerado pela planilha: ${ss.getName()}`;
            try {
                const newEvent = calendar.createAllDayEvent(eventTitle, eventDate, {description: eventDescription});
                if (eventIdColIndex !== -1) { // Verifica novamente, caso a coluna tenha sido criada no loop
                    sheet.getRange(i + 1, eventIdColIndex + 1).setValue(newEvent.getId());
                }
                eventsCreatedCount++;
                registrarLog("Criar Eventos Manual", `Novo evento criado para MI "${miIdentifier}" em ${eventDate.toLocaleDateString()} (ID: ${newEvent.getId()}).`, "INFO");
                Utilities.sleep(200);
            } catch (e) {
                registrarLog("Criar Eventos Manual", `Falha ao criar evento para MI "${miIdentifier}": ${e.toString()}`, "ERROR");
                errorsCount++;
            }
             if ((eventsCreatedCount + errorsCount) % 10 === 0) {
                 SpreadsheetApp.getActiveSpreadsheet().toast(`Progresso: ${eventsCreatedCount} eventos, ${errorsCount} erros...`, "Progresso Agenda", 10);
             }
        }
        SpreadsheetApp.getActiveSpreadsheet().toast("Criação de eventos concluída.", "Progresso Agenda", 5);
        let summaryMessage = `${eventsCreatedCount} novo(s) evento(s) criado(s) com sucesso.`;
        if (errorsCount > 0) { summaryMessage += `\n${errorsCount} erro(s) ocorreram (verifique os Logs).`;}
        ui.alert("Resultado Criação de Eventos", summaryMessage, ui.ButtonSet.OK);

    } catch (e) {
        Logger.log(`Erro geral ao processar eventos da agenda (manual): ${e}`);
        registrarLog("Criar Eventos Manual", `Erro geral: ${e.message}`, "ERROR");
        ui.alert(`Erro geral ao processar eventos da agenda: ${e.message}`);
    }
}

// --- NOVAS Funções para Consolidação de Abas ---

function abrirDialogoConsolidarAbas() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();
    // Filtra abas que não devem ser incluídas na seleção para consolidação (ex: o próprio log, abas de resultado)
    // Aqui estamos usando 'deveIgnorarAba' que já existe e é configurável,
    // mas para consolidação, o usuário pode querer consolidar qualquer aba de dados.
    // Talvez seja melhor não filtrar aqui e deixar o usuário escolher todas.
    const sheetNames = allSheets.map(sheet => sheet.getName()); // Pega todas as abas

    if (sheetNames.length < 2) { // Precisa de ao menos 2 abas para consolidar de forma útil
        SpreadsheetApp.getUi().alert("Poucas Abas", "Você precisa de pelo menos duas abas na planilha para usar a consolidação.", SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    }

    const template = HtmlService.createTemplateFromFile('consolidarAbas');
    template.sheetNames = sheetNames;

    const htmlOutput = template.evaluate().setWidth(450).setHeight(500); // Ajustar tamanho
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Consolidar Abas');
}

/**
 * Executa a consolidação das abas selecionadas.
 * @param {Array<string>} selectedSheetNames Nomes das abas a serem consolidadas.
 * @param {string} newSheetName Nome para a nova aba consolidada.
 * @param {boolean} includeHeaders Se true, inclui o cabeçalho da primeira aba selecionada.
 * @param {boolean} addSourceColumn Se true, adiciona uma coluna indicando a aba de origem de cada linha.
 * @return {Object} Resultado da operação.
 */
function executarConsolidacaoAbas(selectedSheetNames, newSheetName, includeHeaders, addSourceColumn) {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!selectedSheetNames || selectedSheetNames.length === 0) {
        return { success: false, message: "Nenhuma aba selecionada para consolidação." };
    }
    if (!newSheetName || newSheetName.trim() === "") {
        return { success: false, message: "O nome da nova aba consolidada é obrigatório." };
    }
    newSheetName = newSheetName.trim();

    registrarLog("Consolidar Abas", `Iniciando. Abas: ${selectedSheetNames.join(', ')}. Nova aba: ${newSheetName}`, "INFO");
    SpreadsheetApp.getActiveSpreadsheet().toast(`Consolidando ${selectedSheetNames.length} aba(s)...`, "Progresso", -1);

    try {
        let consolidatedSheet = ss.getSheetByName(newSheetName);
        if (consolidatedSheet) {
            // Opção: Perguntar ao usuário se deseja substituir ou adicionar
            // Por simplicidade, vamos limpar e substituir
            consolidatedSheet.clearContents().clearFormats();
            registrarLog("Consolidar Abas", `Aba existente "${newSheetName}" limpa.`, "INFO");
        } else {
            consolidatedSheet = ss.insertSheet(newSheetName);
            registrarLog("Consolidar Abas", `Nova aba "${newSheetName}" criada.`, "INFO");
        }

        let firstSheetProcessed = false;
        let headerSource = null; // Para armazenar os cabeçalhos da primeira aba processada

        selectedSheetNames.forEach((sheetName, index) => {
            const sourceSheet = ss.getSheetByName(sheetName);
            if (!sourceSheet) {
                registrarLog("Consolidar Abas", `Aba de origem "${sheetName}" não encontrada. Pulando.`, "WARNING");
                return; // Pula para a próxima
            }

            const sourceRange = sourceSheet.getDataRange();
            if (!sourceRange || sourceRange.getNumRows() === 0) {
                registrarLog("Consolidar Abas", `Aba de origem "${sheetName}" está vazia. Pulando.`, "INFO");
                return; // Pula para a próxima
            }
            let sourceValues = sourceRange.getValues();

            if (includeHeaders) {
                if (!firstSheetProcessed) { // Se é a primeira aba e devemos incluir cabeçalhos
                    headerSource = sourceValues[0];
                    if (addSourceColumn) {
                        consolidatedSheet.appendRow(["Origem da Aba", ...headerSource]);
                    } else {
                        consolidatedSheet.appendRow(headerSource);
                    }
                    consolidatedSheet.getRange(1, 1, 1, consolidatedSheet.getLastColumn()).setFontWeight("bold");
                    consolidatedSheet.setFrozenRows(1);
                    firstSheetProcessed = true;
                    sourceValues.shift(); // Remove cabeçalho dos dados a serem copiados
                } else { // Para abas subsequentes, se o cabeçalho já foi adicionado
                    if (sourceValues.length > 0 && JSON.stringify(sourceValues[0]) === JSON.stringify(headerSource)) {
                       sourceValues.shift(); // Remove cabeçalho se for igual ao da primeira aba
                    }
                }
            }
            
            if (sourceValues.length > 0) {
                if (addSourceColumn) {
                    const valuesWithSource = sourceValues.map(row => [sheetName, ...row]);
                    consolidatedSheet.getRange(consolidatedSheet.getLastRow() + 1, 1, valuesWithSource.length, valuesWithSource[0].length)
                        .setValues(valuesWithSource);
                } else {
                    consolidatedSheet.getRange(consolidatedSheet.getLastRow() + 1, 1, sourceValues.length, sourceValues[0].length)
                        .setValues(sourceValues);
                }
            }
            SpreadsheetApp.getActiveSpreadsheet().toast(`Processando aba "${sheetName}"... (${index+1}/${selectedSheetNames.length})`, "Progresso", 10);
            Utilities.sleep(100); // Pequena pausa para permitir atualizações da UI e evitar sobrecarga
        });

        if (consolidatedSheet.getLastRow() > 1 || (consolidatedSheet.getLastRow() === 1 && includeHeaders) ) { // Se algo foi escrito (além do título da consolidação)
             consolidatedSheet.autoResizeColumns(1, consolidatedSheet.getLastColumn());
        }
       
        SpreadsheetApp.setActiveSheet(consolidatedSheet);
        registrarLog("Consolidar Abas", `Consolidação concluída na aba "${newSheetName}".`, "INFO");
        SpreadsheetApp.getActiveSpreadsheet().toast("Consolidação concluída!", "Sucesso", 5);
        return { success: true, message: `Abas consolidadas com sucesso em "${newSheetName}"!` };

    } catch (e) {
        Logger.log(`Erro ao consolidar abas: ${e.toString()} ${e.stack}`);
        registrarLog("Consolidar Abas", `Erro: ${e.message}`, "ERROR");
        SpreadsheetApp.getActiveSpreadsheet().toast("Erro na consolidação.", "Erro", 5);
        return { success: false, message: `Erro ao consolidar abas: ${e.message}` };
    }
}
// Adicione as funções auxiliares (findColumnIndexByHeader, deveIgnorarAba, etc.) e
// as funções de Configuração e Exportação/Backup da sua versão anterior aqui.
// Certifique-se de que todas as funções chamadas existam no seu Code.gs

// --- END OF Code.gs ---
