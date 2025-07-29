const { app, BrowserWindow, ipcMain, dialog, shell } = require("electron");
const { autoUpdater } = require("electron-updater");
const path = require("path");
const fs = require("fs");
const fsp = require("fs").promises; // Usaremos a versão de promessas do FS
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");
const axios = require("axios");
const os = require('os');

const { initializeApp, cert } = require("firebase-admin/app");
const { getFirestore } = require("firebase-admin/firestore");

autoUpdater.logger = require("electron-log");
autoUpdater.logger.transports.file.level = "info";

let mainWindow;

function sendUpdateStatusToWindow(text) {
    if (mainWindow && mainWindow.webContents) {
        mainWindow.webContents.send("update-message", text);
    }
}

autoUpdater.on("checking-for-update", () => sendUpdateStatusToWindow("Verificando por atualizações..."));
autoUpdater.on("update-available", (info) => sendUpdateStatusToWindow(`Atualização disponível (v${
    info.version
}). Baixando...`));
autoUpdater.on("update-not-available", () => sendUpdateStatusToWindow(""));
autoUpdater.on("error", (err) => sendUpdateStatusToWindow(`Erro na atualização: ${
    err.toString()
}`));
autoUpdater.on("download-progress", (p) => sendUpdateStatusToWindow(`Baixando atualização: ${
    Math.round(p.percent)
}%`));
autoUpdater.on("update-downloaded", (info) => {
    sendUpdateStatusToWindow(`Atualização v${
        info.version
    } baixada. Reinicie para instalar.`);
    if (mainWindow && mainWindow.webContents) {
        mainWindow.webContents.executeJavaScript(`
            const um = document.getElementById("update-message");
            if(um){ um.style.cursor="pointer"; um.style.textDecoration="underline"; um.onclick = () => require("electron").ipcRenderer.send("restart-app-for-update"); }
        `);
    }
});

ipcMain.on("restart-app-for-update", () => autoUpdater.quitAndInstall());
ipcMain.on('open-path', (event, filePath) => {
    shell.openPath(filePath).catch(err => {
        const msg = `ERRO: Não foi possível abrir o arquivo em ${filePath}`;
        console.error("Falha ao abrir o caminho:", err);
        event.sender.send("log", msg);
        event.sender.send("automation-log", msg);
    });
});

async function runPhoneAdjustment(filePath, event, backup) {
    const log = (msg) => event.sender.send("log", msg);
    if (!filePath || !fs.existsSync(filePath)) {
        log(`❌ Erro: Arquivo para ajuste de fones não encontrado em: ${filePath}`);
        return;
    }
    log(`\n--- Iniciando Ajuste de Fones para: ${
        path.basename(filePath)
    } ---`);
    try {
        if (backup) {
            const p = path.parse(filePath);
            const backupPath = path.join(p.dir, `${
                p.name
            }.backup_fones_${
                Date.now()
            }${
                p.ext
            }`);
            fs.copyFileSync(filePath, backupPath);
            log(`Backup do arquivo criado em: ${backupPath}`);
        }
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const worksheet = workbook.worksheets[0];
        const phoneColumns = [];
        worksheet.getRow(1).eachCell({
            includeEmpty: true
        }, (cell, colNumber) => {
            if (cell.value && typeof cell.value === "string" && cell.value.trim().toLowerCase().startsWith("fone")) {
                phoneColumns.push(colNumber);
            }
        });
        phoneColumns.sort((a, b) => a - b);
        if (phoneColumns.length === 0) {
            log("⚠️ Nenhuma coluna \"fone\" encontrada. Ajuste pulado.");
            return;
        }
        log(`Ajustando ${
            phoneColumns.length
        } colunas de telefone...`);
        let processedRows = 0;
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) 
                return;
            
            const phoneValuesInRow = phoneColumns.map(colNumber => row.getCell(colNumber).value).filter(v => v !== null && v !== undefined && String(v).trim() !== "");
            phoneColumns.forEach((colNumber, index) => {
                row.getCell(colNumber).value = index < phoneValuesInRow.length ? phoneValuesInRow[index] : null;
            });
            processedRows++;
        });
        await workbook.xlsx.writeFile(filePath);
        log(`✅ Ajuste de fones concluído. ${processedRows} linhas processadas.`);
    } catch (err) {
        log(`❌ Erro catastrófico durante o ajuste de fones: ${
            err.message
        }`);
        console.error(err);
    }
}

const getCredentialsPath = () => app.isPackaged ? path.join(process.resourcesPath, "firebase-credentials.json") : path.join(__dirname, "firebase-credentials.json");

try {
    initializeApp({
        credential: cert(require(getCredentialsPath()))
    });
} catch (error) {
    console.error("ERRO FATAL: Não foi possível carregar 'firebase-credentials.json'.", error);
    dialog.showErrorBox("Erro Crítico", "As credenciais do Firebase não foram encontradas ou são inválidas. A aplicação será encerrada.");
    app.quit();
}

const db = getFirestore();
const cnpjsCollection = db.collection("cnpjs_armazenados");
const enrichmentCollection = db.collection("cnpjs_enriquecidos");
const rootCollection = db.collection("Raiz");
let storedCnpjs = new Set();

async function loadStoredCnpjs() {
    try {
        const snapshot = await cnpjsCollection.get();
        storedCnpjs = new Set(snapshot.docs.map(doc => doc.id));
        console.log(`${
            storedCnpjs.size
        } CNPJs carregados do Firestore.`);
        if (mainWindow) {
            mainWindow.webContents.send("log", `✅ Sincronização com o BD concluída. ${
                storedCnpjs.size
            } CNPJs carregados.`);
        }
    } catch (err) {
        console.error("Falha ao carregar CNPJs do Firestore:", err);
        if (mainWindow) {
            mainWindow.webContents.send("log", `❌ ERRO ao carregar histórico do Firebase: ${
                err.message
            }`);
        }
    }
}

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1200,
        height: 900,
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            preload: path.join(__dirname, "preload.js")
        }
    });
    mainWindow.loadFile("index.html");
    mainWindow.webContents.on("did-finish-load", () => {
        loadStoredCnpjs();
        autoUpdater.checkForUpdatesAndNotify();
    });
}

ipcMain.on("prepare-enrichment-files", async (event, filePaths) => {
    const log = (msg) => event.sender.send("enrichment-log", msg);
    for (const filePath of filePaths) {
        log(`Preparando arquivo: ${
            path.basename(filePath)
        }...`);
        try {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(filePath);
            const worksheet = workbook.worksheets[0];
            if (!worksheet) {
                log(`⚠️ Planilha não encontrada em ${
                    path.basename(filePath)
                }. Pulando.`);
                continue;
            }
            const existingHeaders = new Set();
            worksheet.getRow(1).eachCell({
                includeEmpty: true
            }, (cell) => {
                if (cell.value) 
                    existingHeaders.add(String(cell.value).trim().toLowerCase());
                
            });
            let addedHeaders = false;
            let nextAvailableColumn = worksheet.columnCount + 1;
            for (let i = 1; i <= 14; i++) {
                const headerName = `fone${i}`;
                if (! existingHeaders.has(headerName)) {
                    worksheet.getCell(1, nextAvailableColumn).value = headerName;
                    nextAvailableColumn++;
                    addedHeaders = true;
                }
            }
            if (addedHeaders) {
                await workbook.xlsx.writeFile(filePath);
                log(`✅ Cabeçalhos de fone adicionados para: ${
                    path.basename(filePath)
                }`);
            } else {
                log(`Todos os cabeçalhos necessários já existem em: ${
                    path.basename(filePath)
                }`);
            }
        } catch (err) {
            log(`❌ Erro ao preparar ${
                path.basename(filePath)
            }: ${
                err.message
            }`);
        }
    }
});

app.whenReady().then(createWindow);
app.on("window-all-closed", () => {
    if (process.platform !== "darwin") 
        app.quit();
    
});

ipcMain.handle("select-file", async (event, {
    title,
    multi
}) => {
    const {canceled, filePaths} = await dialog.showOpenDialog(mainWindow, {
        title: title,
        properties: [
            multi ? "multiSelections" : "openFile",
            "openFile"
        ],
        filters: [
            {
                name: "Planilhas",
                extensions: ["xlsx", "xls", "csv"]
            }
        ]
    });
    return canceled ? null : filePaths;
});

function letterToIndex(letter) {
    return letter.toUpperCase().charCodeAt(0) - 65;
}

async function readSpreadsheet(filePath) {
    try {
        if (path.extname(filePath).toLowerCase() === ".csv") {
            const data = await fsp.readFile(filePath, "utf8");
            return XLSX.read(data, {
                type: "string",
                cellDates: true
            });
        } else {
            const buffer = await fsp.readFile(filePath);
            return XLSX.read(buffer, {
                type: 'buffer',
                cellDates: true
            });
        }
    } catch (e) {
        console.error(`Erro ao ler planilha: ${filePath}`, e);
        throw new Error(`Não foi possível ler o arquivo ${
            path.basename(filePath)
        }. Verifique se o caminho está correto e se você tem permissão.`);
    }
}
function writeSpreadsheet(workbook, filePath) {
    XLSX.writeFile(workbook, filePath);
}

ipcMain.handle("get-enriched-cnpj-count", async () => (await enrichmentCollection.get()).size);
ipcMain.handle("download-enriched-data", async () => {
    try {
        const {canceled, filePath} = await dialog.showSaveDialog(mainWindow, {
            title: "Salvar Dados Enriquecidos",
            defaultPath: `dados_enriquecidos_${
                Date.now()
            }.xlsx`,
            filters: [
                {
                    name: "Excel Files",
                    extensions: ["xlsx"]
                }
            ]
        });
        if (canceled || !filePath) 
            return {success: false, message: "Download cancelado."};
        
        const snapshot = await enrichmentCollection.get();
        if (snapshot.empty) 
            return {success: false, message: "Nenhum dado encontrado."};
        
        const headers = [
            "cpf",
            ... Array.from({
                length: 14
            }, (_, i) => `fone${
                i + 1
            }`)
        ];
        const data = snapshot.docs.map(doc => {
            const phones = doc.data().telefones || [];
            return [
                doc.id,
                ... Array.from({
                    length: 14
                }, (_, i) => phones[i] || "")
            ];
        });

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Dados Enriquecidos");
        worksheet.addRow(headers);
        worksheet.addRows(data);
        await workbook.xlsx.writeFile(filePath);
        return {success: true, message: `Arquivo salvo com sucesso: ${filePath}`};
    } catch (error) {
        return {
            success: false,
            message: `Erro ao gerar arquivo: ${
                error.message
            }`
        };
    }
});

ipcMain.on("start-db-load", async (event, {masterFiles}) => {
    const log = (msg) => event.sender.send("enrichment-log", msg);
    const progress = (current, total, fileName, cnpjsProcessed) => event.sender.send("db-load-progress", {
        current,
        total,
        fileName,
        cnpjsProcessed
    });
    log(`--- Iniciando Carga para o Banco de Dados de Enriquecimento ---`);
    let totalCnpjsProcessed = 0;
    const saveChunkToFirestore = async (dataMap, filePath) => {
        if (dataMap.size === 0) 
            return;
        
        const batchSize = 400;
        const cnpjsArray = Array.from(dataMap.entries());
        for (let i = 0; i < cnpjsArray.length; i += batchSize) {
            const batch = db.batch();
            const chunk = cnpjsArray.slice(i, i + batchSize);
            chunk.forEach(([cnpj, phones]) => {
                const docRef = enrichmentCollection.doc(cnpj);
                const uniqueValidPhones = [... new Set(phones)].filter(p => String(p).replace(/\D/g, '').length >= 8);
                if (uniqueValidPhones.length > 0) {
                    batch.set(docRef, {
                        telefones: uniqueValidPhones,
                        ultima_atualizacao: new Date(),
                        fonte_dados: path.basename(filePath)
                    }, {merge: true});
                }
            });
            await batch.commit();
        }
        totalCnpjsProcessed += dataMap.size;
    };
    try {
        for (let fileIndex = 0; fileIndex < masterFiles.length; fileIndex++) {
            const filePath = masterFiles[fileIndex];
            const fileName = path.basename(filePath);
            progress(fileIndex + 1, masterFiles.length, fileName, totalCnpjsProcessed);
            log(`\nProcessando arquivo mestre: ${fileName}`);
            try {
                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.readFile(filePath);
                const worksheet = workbook.worksheets[0];
                if (!worksheet || worksheet.rowCount === 0) {
                    log(`⚠️ Arquivo ${fileName} vazio ou inválido. Pulando.`);
                    continue;
                }
                const headerMap = new Map();
                worksheet.getRow(1).eachCell({
                    includeEmpty: true
                }, (cell, colNum) => headerMap.set(colNum, String(cell.value || "").trim().toLowerCase()));
                let cnpjColIdx = [... headerMap.entries()].find(([_, h]) => h === "cpf" || h === "cnpj")?.[0] ?? -1;
                const phoneColIdxs = [... headerMap.entries()].filter(([_, h]) => /^(fone|telefone|celular)/.test(h)).map(([colNum]) => colNum);

                if (cnpjColIdx === -1 || phoneColIdxs.length === 0) {
                    log(`❌ ERRO: Colunas de documento ou telefone não encontradas. Pulando.`);
                    continue;
                }

                let cnpjsToUpdate = new Map();
                for (let i = 2; i <= worksheet.rowCount; i++) {
                    const row = worksheet.getRow(i);
                    const cnpj = String(row.getCell(cnpjColIdx).value || "").replace(/\D/g, "").trim();
                    if (cnpj.length < 8) 
                        continue;
                    
                    const phones = phoneColIdxs.map(idx => String(row.getCell(idx).value || "").trim()).filter(Boolean);
                    if (phones.length > 0) 
                        cnpjsToUpdate.set(cnpj, [
                            ...(
                                cnpjsToUpdate.get(cnpj) || []
                            ),
                            ... phones
                        ]);
                    
                    if (i % 5000 === 0) {
                        await saveChunkToFirestore(cnpjsToUpdate, filePath);
                        cnpjsToUpdate.clear();
                        progress(fileIndex + 1, masterFiles.length, fileName, totalCnpjsProcessed);
                    }
                }
                if (cnpjsToUpdate.size > 0) 
                    await saveChunkToFirestore(cnpjsToUpdate, filePath);
                
            } catch (err) {
                log(`❌ ERRO ao processar ${fileName}: ${
                    err.message
                }`);
            }
        }
    } catch (err) {
        log(`❌ Um erro crítico ocorreu: ${
            err.message
        }`);
    } finally {
        log(`\n✅ Carga finalizada. Total de ${totalCnpjsProcessed} CNPJs únicos processados.`);
        event.sender.send("db-load-finished");
    }
});

function formatEta(totalSeconds) {
    if (! isFinite(totalSeconds) || totalSeconds < 0) 
        return "Calculando...";
    
    const m = Math.floor(totalSeconds / 60);
    const s = Math.floor(totalSeconds % 60);
    return `${
        String(m).padStart(2, '0')
    }:${
        String(s).padStart(2, '0')
    }`;
}

async function runEnrichmentProcess({
    filesToEnrich,
    strategy,
    backup
}, log, progress, onFinish) {
    log("--- Iniciando Processo de Enriquecimento por Lotes ---");
    let totalEnrichedRowsOverall = 0, totalProcessedRowsOverall = 0, totalNotFoundInDbOverall = 0;
    const BATCH_SIZE = 2000;
    try {
        for (const fileObj of filesToEnrich) {
            const {path: filePath, id} = fileObj;
            const startTime = Date.now();
            log(`\nProcessando arquivo: ${
                path.basename(filePath)
            }`);
            progress(id, 0, null);
            if (backup) {
                const p = path.parse(filePath);
                fs.copyFileSync(filePath, path.join(p.dir, `${
                    p.name
                }.backup_enrich_${
                    Date.now()
                }${
                    p.ext
                }`));
                log(`Backup criado.`);
            }
            try {
                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.readFile(filePath);
                const worksheet = workbook.worksheets[0];
                let cnpjCol = -1, statusCol = -1;
                const phoneCols = [];
                worksheet.getRow(1).eachCell((cell, colNum) => {
                    const h = String(cell.value || "").trim().toLowerCase();
                    if (h === "cpf" || h === "cnpj") 
                        cnpjCol = colNum;
                     else if (h.startsWith("fone")) 
                        phoneCols.push(colNum);
                     else if (h === "status") 
                        statusCol = colNum;
                    
                });
                phoneCols.sort((a, b) => a - b);
                if (cnpjCol === -1) {
                    log(`❌ ERRO: Coluna 'cpf'/'cnpj' não encontrada. Pulando.`);
                    continue;
                }
                if (statusCol === -1) {
                    statusCol = worksheet.columnCount + 1;
                    worksheet.getCell(1, statusCol).value = "status";
                }
                const totalRows = worksheet.rowCount - 1;
                let enrichedInFile = 0, notFoundInFile = 0;

                const totalBatches = Math.ceil((worksheet.rowCount - 1) / BATCH_SIZE);
                log(`Arquivo possui ${totalRows} linhas, divididas em ${totalBatches} lotes.`);

                for (let i = 2; i <= worksheet.rowCount; i += BATCH_SIZE) {
                    const currentBatchNum = Math.floor((i - 2) / BATCH_SIZE) + 1;
                    const cnpjsInBatch = new Map();
                    const endIndex = Math.min(i + BATCH_SIZE - 1, worksheet.rowCount);

                    for (let j = i; j <= endIndex; j++) {
                        const row = worksheet.getRow(j);
                        const cnpj = String(row.getCell(cnpjCol).text || "").replace(/\D/g, "").trim();
                        if (cnpj) 
                            cnpjsInBatch.set(cnpj, {rowNum: j, row: row});
                        
                    }
                    if (cnpjsInBatch.size === 0) 
                        continue;
                    

                    log(`Lote ${currentBatchNum}/${totalBatches}: Processando ${
                        cnpjsInBatch.size
                    } CNPJs...`);

                    const enrichmentDataForBatch = new Map();
                    const cnpjKeys = Array.from(cnpjsInBatch.keys());
                    const queryChunks = [];
                    for (let k = 0; k < cnpjKeys.length; k += 30) 
                        queryChunks.push(cnpjKeys.slice(k, k + 30));
                    
                    for (const chunk of queryChunks) {
                        const snapshot = await enrichmentCollection.where('__name__', 'in', chunk).get();
                        snapshot.forEach(doc => enrichmentDataForBatch.set(doc.id, doc.data().telefones || []));
                    }

                    log(`Lote ${currentBatchNum}/${totalBatches}: ${
                        enrichmentDataForBatch.size
                    } CNPJs encontrados no BD. Atualizando planilha...`);

                    for (const [cnpj, {row}] of cnpjsInBatch.entries()) {
                        let rowWasEnriched = false;
                        if (enrichmentDataForBatch.has(cnpj)) {
                            const phonesFromDb = enrichmentDataForBatch.get(cnpj);
                            const existingPhones = phoneCols.map(idx => row.getCell(idx).value).filter(Boolean);
                            const shouldProcess = (strategy === "overwrite") || (strategy === "append" && existingPhones.length < phoneCols.length) || (strategy === "ignore" && existingPhones.length === 0);
                            if (shouldProcess) {
                                rowWasEnriched = true;
                                if (strategy === "overwrite") 
                                    phoneCols.forEach(idx => row.getCell(idx).value = null);
                                
                                let phonesToWrite = [... phonesFromDb];
                                phoneCols.forEach(idx => {
                                    if (strategy === "append" && row.getCell(idx).value) 
                                        return;
                                    
                                    if (phonesToWrite.length > 0) 
                                        row.getCell(idx).value = phonesToWrite.shift();
                                    
                                });
                            }
                        } else {
                            if (cnpj) 
                                notFoundInFile++;
                            
                        }
                        row.getCell(statusCol).value = rowWasEnriched ? "Enriquecido" : "Pobre";
                        if (rowWasEnriched) 
                            enrichedInFile++;
                        
                    }
                    const processedRowsInFile = endIndex - 1;
                    const eta = formatEta((totalRows - processedRowsInFile) / (processedRowsInFile / (Date.now() - startTime)));
                    progress(id, Math.round((processedRowsInFile / totalRows) * 100), eta);
                }
                await workbook.xlsx.writeFile(filePath);
                progress(id, 100, "00:00");
                log(`✅ Arquivo ${
                    path.basename(filePath)
                } concluído! Total de enriquecidos: ${enrichedInFile}. Não encontrados: ${notFoundInFile}.`);
                totalEnrichedRowsOverall += enrichedInFile;
                totalNotFoundInDbOverall += notFoundInFile;
                totalProcessedRowsOverall += totalRows;
            } catch (err) {
                log(`❌ ERRO catastrófico em ${
                    path.basename(filePath)
                }: ${
                    err.message
                }`);
            }
        }
    } finally {
        log(`\n--- ✅ Processo de Enriquecimento Finalizado ---`);
        log(`Resumo Geral: Total Linhas Processadas: ${totalProcessedRowsOverall}, Enriquecidas: ${totalEnrichedRowsOverall}, Não Encontradas: ${totalNotFoundInDbOverall}`);
        if (onFinish) 
            onFinish();
        
    }
}

ipcMain.on("start-enrichment", async (event, options) => {
    await runEnrichmentProcess(options, (msg) => event.sender.send("enrichment-log", msg), (id, pct, eta) => event.sender.send("enrichment-progress", {
        id,
        progress: pct,
        eta
    }), () => event.sender.send("enrichment-finished"));
});


// #################################################################
// #           NOVO HANDLER PARA A ABA DE MONITORAMENTO            #
// #################################################################
ipcMain.handle('fetch-monitoring-report', async (event, url) => {
    console.log(`Buscando relatório do endpoint: ${url}`);
    try {
        const response = await axios.get(url, {
            timeout: 90000
        }); // Timeout de 90 segundos

        if (response.status === 200) {
            // A API pode retornar uma string "Nenhum registro encontrado" em vez de um JSON vazio
            if (typeof response.data === 'string' && response.data.includes("Nenhum registro encontrado")) {
                return {
                    success: true,
                    data: []
                }; // Retorna array vazio para consistência
            }
            return {
                success: true,
                data: response.data
            };
        } else {
            return {
                success: false,
                message: `A API retornou um status inesperado: ${
                    response.status
                }`
            };
        }
    } catch (error) {
        console.error("Erro ao buscar relatório de monitoramento:", error);
        return {
            success: false,
            message: `Falha na comunicação com a API: ${
                error.message
            }`
        };
    }
});


// #################################################################
// #           FUNÇÃO PARA ALIMENTAR A BASE RAIZ                   #
// #################################################################

ipcMain.on("feed-root-database", async (event, filePaths) => {
    const log = (msg) => event.sender.send("log", msg);

    log(`--- Iniciando Alimentação da Base Raiz ---`);

    const BATCH_SIZE = 5000;
    const FIRESTORE_QUERY_LIMIT = 30;
    const FIRESTORE_WRITE_LIMIT = 499;

    let totalNewCnpjsAdded = 0;

    const findCnpjColumn = (headerRow) => {
        let cnpjColIdx = -1;
        headerRow.eachCell((cell, colNumber) => {
            const header = String(cell.value || "").trim().toLowerCase();
            if (header === 'cpf' || header === 'cnpj') {
                cnpjColIdx = colNumber;
            }
        });
        return cnpjColIdx;
    };

    const processChunk = async (cnpjChunk, sourceFile) => {
        if (cnpjChunk.size === 0) 
            return;
        
        log(`Verificando ${
            cnpjChunk.size
        } CNPJs únicos contra o BD...`);

        const cnpjsToCheck = Array.from(cnpjChunk);
        const trulyNewCnpjs = new Set();

        for (let i = 0; i < cnpjsToCheck.length; i += FIRESTORE_QUERY_LIMIT) {
            const queryChunk = cnpjsToCheck.slice(i, i + FIRESTORE_QUERY_LIMIT);
            if (queryChunk.length === 0) 
                continue;
            

            const existingCnpjs = new Set();
            try {
                const snapshot = await rootCollection.where('__name__', 'in', queryChunk).get();
                snapshot.forEach(doc => existingCnpjs.add(doc.id));
            } catch (e) {
                log(`❌ Erro ao verificar existência de CNPJs: ${
                    e.message
                }`);
                continue;
            }

            queryChunk.forEach(cnpj => {
                if (! existingCnpjs.has(cnpj)) {
                    trulyNewCnpjs.add(cnpj);
                }
            });
        }

        const newCnpjsCount = trulyNewCnpjs.size;
        if (newCnpjsCount === 0) {
            log(`Nenhum CNPJ novo para adicionar deste lote.`);
            return;
        }

        log(`Encontrados ${newCnpjsCount} CNPJs novos. Enviando para o Firestore...`);

        const cnpjsToWrite = Array.from(trulyNewCnpjs);
        for (let i = 0; i < cnpjsToWrite.length; i += FIRESTORE_WRITE_LIMIT) {
            const writeChunk = cnpjsToWrite.slice(i, i + FIRESTORE_WRITE_LIMIT);
            const batch = db.batch();
            const batchId = `raiz-feed-${
                Date.now()
            }`;

            writeChunk.forEach(cnpj => {
                const docRef = rootCollection.doc(cnpj);
                batch.set(docRef, {
                    adicionado_em: new Date(),
                    fonte: sourceFile,
                    lote_id: batchId
                });
            });

            try {
                await batch.commit();
                log(`✅ Lote de ${
                    writeChunk.length
                } CNPJs salvo na coleção Raiz com sucesso.`);
                totalNewCnpjsAdded += writeChunk.length;
            } catch (e) {
                log(`❌ Erro ao salvar lote na coleção Raiz: ${
                    e.message
                }`);
            }
        }
    };


    for (const filePath of filePaths) {
        const fileName = path.basename(filePath);
        log(`\nIniciando processamento do arquivo: ${fileName}`);

        try {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(filePath);
            const worksheet = workbook.worksheets[0];

            if (!worksheet || worksheet.rowCount <= 1) {
                log(`⚠️ Arquivo ${fileName} está vazio ou não possui dados. Pulando.`);
                continue;
            }

            const headerRow = worksheet.getRow(1);
            const cnpjColIdx = findCnpjColumn(headerRow);

            if (cnpjColIdx === -1) {
                log(`❌ ERRO: Coluna 'cpf' ou 'cnpj' não encontrada em ${fileName}. Pulando.`);
                continue;
            }
            log(`Coluna de CNPJ/CPF encontrada na posição: ${cnpjColIdx}`);

            const totalDataRows = worksheet.rowCount - 1;
            const estimatedTotalBatches = Math.ceil(totalDataRows / BATCH_SIZE);
            let currentBatchNumber = 0;
            log(`Arquivo possui ${totalDataRows} linhas de dados. Serão processados em aproximadamente ${estimatedTotalBatches} lote(s) de leitura de ${BATCH_SIZE} registros.`);

            let cnpjsFromFile = new Set();

            for (let i = 2; i <= worksheet.rowCount; i++) {
                const row = worksheet.getRow(i);
                const cellValue = row.getCell(cnpjColIdx).value;
                const cnpj = cellValue ? String(cellValue).replace(/\D/g, "").trim() : null;

                if (cnpj && (cnpj.length === 11 || cnpj.length === 14)) {
                    cnpjsFromFile.add(cnpj);
                }

                if (cnpjsFromFile.size >= BATCH_SIZE) {
                    currentBatchNumber++;
                    log(`\n--- Processando Lote de Leitura ${currentBatchNumber} de ~${estimatedTotalBatches} ---`);
                    await processChunk(cnpjsFromFile, fileName);
                    cnpjsFromFile.clear();
                }
            }

            if (cnpjsFromFile.size > 0) {
                log(`\n--- Processando Lote Final (${
                    cnpjsFromFile.size
                } registros) ---`);
                await processChunk(cnpjsFromFile, fileName);
                cnpjsFromFile.clear();
            }

            log(`\n✅ Finalizado o processamento do arquivo ${fileName}.`);

        } catch (err) {
            log(`❌ Erro catastrófico ao processar o arquivo ${fileName}: ${
                err.message
            }`);
            console.error(err);
        }
    }

    log(`\n--- Alimentação da Base Raiz Concluída ---`);
    log(`Total de CNPJs novos adicionados à Raiz: ${totalNewCnpjsAdded}`);
    event.sender.send("root-feed-finished");
});


// #################################################################
// #           LÓGICA DE LIMPEZA, MERGE E BANCO DE DADOS (ORIGINAL)  #
// #################################################################

ipcMain.handle("save-stored-cnpjs-to-excel", async (event) => {
    if (storedCnpjs.size === 0) {
        dialog.showMessageBox(mainWindow, {
            type: "info",
            title: "Aviso",
            message: "Nenhum CNPJ armazenado para salvar."
        });
        return {
            success: false,
            message: "Nenhum CNPJ armazenado para salvar."
        };
    }
    const {canceled, filePath} = await dialog.showSaveDialog(mainWindow, {
        title: "Salvar CNPJs Armazenados",
        defaultPath: `cnpjs_armazenados_${
            Date.now()
        }.xlsx`,
        filters: [
            {
                name: "Excel Files",
                extensions: ["xlsx"]
            }
        ]
    });
    if (canceled || !filePath) {
        return {
            success: false,
            message: "Operação de salvar cancelada."
        };
    }
    try {
        const data = Array.from(storedCnpjs).map(cnpj => [cnpj]);
        const worksheet = XLSX.utils.aoa_to_sheet([
            ["cpf"],
            ... data
        ]);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "CNPJs");
        XLSX.writeFile(workbook, filePath);
        dialog.showMessageBox(mainWindow, {
            type: "info",
            title: "Sucesso",
            message: `Arquivo salvo com sucesso em: ${filePath}`
        });
        return {
            success: true,
            message: `Arquivo salvo com sucesso em: ${filePath}`
        };
    } catch (err) {
        console.error("Erro ao salvar Excel:", err);
        dialog.showMessageBox(mainWindow, {
            type: "error",
            title: "Erro",
            message: `Erro ao salvar arquivo: ${
                err.message
            }`
        });
        return {
            success: false,
            message: `Erro ao salvar arquivo: ${
                err.message
            }`
        };
    }
});

ipcMain.handle("delete-batch", async (event, batchId) => {
    const log = (msg) => event.sender.send("log", msg);
    if (!batchId) {
        return {
            success: false,
            message: "ID do lote inválido."
        };
    }
    log(`Buscando documentos do lote "${batchId}" no Firestore...`);
    try {
        const querySnapshot = await cnpjsCollection.where("batchId", "==", batchId).get();
        if (querySnapshot.empty) {
            return {
                success: false,
                message: `Nenhum CNPJ encontrado para o lote "${batchId}". Verifique o ID.`
            };
        }
        const count = querySnapshot.size;
        log(`Encontrados ${count} CNPJs para exclusão. Removendo em lotes...`);
        const docRefs = [];
        querySnapshot.forEach(doc => {
            docRefs.push(doc.ref);
            storedCnpjs.delete(doc.id);
        });
        const batchSize = 500;
        for (let i = 0; i < docRefs.length; i += batchSize) {
            const chunk = docRefs.slice(i, i + batchSize);
            const firestoreBatch = db.batch();
            chunk.forEach(docRef => {
                firestoreBatch.delete(docRef);
            });
            await firestoreBatch.commit();
            log(`Lote parcial de ${
                chunk.length
            } registros excluído...`);
        }
        log(`Total de CNPJs no cache local agora: ${
            storedCnpjs.size
        }`);
        return {
            success: true,
            message: `✅ ${count} CNPJs do lote "${batchId}" foram excluídos com sucesso!`
        };
    } catch (err) {
        console.error("Erro ao excluir lote do Firestore:", err);
        return {
            success: false,
            message: `❌ Erro ao excluir lote: ${
                err.message
            }`
        };
    }
});

ipcMain.handle("update-blocklist", async (event, backup) => {
    const log = (msg) => event.sender.send("log", msg);
    try {
        const blocklistPath = "G:\\Meu Drive\\Marketing\\!Campanhas\\URA - Automatica\\Limpeza de base\\bases para a raiz\\Blocklist.xlsx";
        const rootPath = "G:\\Meu Drive\\Marketing\\!Campanhas\\URA - Automatica\\Limpeza de base\\raiz_att.xlsx";
        if (backup) {
            const timestamp = Date.now();
            const bkp = path.join(path.dirname(rootPath), `${
                path.basename(rootPath, path.extname(rootPath))
            }.backup_${timestamp}${
                path.extname(rootPath)
            }`);
            fs.copyFileSync(rootPath, bkp);
            log(`Backup da raiz criado em: ${bkp}`);
        }
        const wbBlock = await readSpreadsheet(blocklistPath);
        const dataBlock = XLSX.utils.sheet_to_json(wbBlock.Sheets[wbBlock.SheetNames[0]], {header: 1}).flat().filter(v => v);
        const wbRoot = await readSpreadsheet(rootPath);
        const dataRoot = XLSX.utils.sheet_to_json(wbRoot.Sheets[wbRoot.SheetNames[0]], {header: 1}).flat().filter(v => v);
        const merged = Array.from(new Set([
            ... dataRoot,
            ... dataBlock
        ])).map(v => [v]);
        const newWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWB, XLSX.utils.aoa_to_sheet(merged), wbRoot.SheetNames[0]);
        writeSpreadsheet(newWB, rootPath);
        return {
            success: true,
            message: "Raiz atualizada com blocklist com sucesso."
        };
    } catch (err) {
        return {
            success: false,
            message: err.message
        };
    }
});

ipcMain.on("start-cleaning", async (event, {
    isAutoRoot,
    rootFile,
    cleanFiles,
    rootCol,
    destCol,
    backup,
    checkDb,
    saveToDb,
    autoAdjust
}) => {
    const log = (msg) => event.sender.send("log", msg);
    try {
        const batchId = `batch-${
            Date.now()
        }`;
        if (saveToDb) 
            log(`Este lote de salvamento terá o ID: ${batchId}`);
        

        const rootSet = new Set();
        if (isAutoRoot) {
            log("Auto Raiz ATIVADO. Carregando lista raiz do Banco de Dados...");
            const snapshot = await rootCollection.get();
            snapshot.forEach(doc => rootSet.add(doc.id));
            log(`✅ Raiz do BD carregada. Total de CNPJs na raiz: ${
                snapshot.size
            }.`);
        } else {
            if (!rootFile || !fs.existsSync(rootFile)) {
                return log(`❌ Arquivo raiz não encontrado: ${rootFile}`);
            }
            const rootIdx = letterToIndex(rootCol);
            const wbRoot = await readSpreadsheet(rootFile);
            const sheetRoot = wbRoot.Sheets[wbRoot.SheetNames[0]];
            const rowsRoot = XLSX.utils.sheet_to_json(sheetRoot, {header: 1}).map(r => r[rootIdx]).filter(v => v).map(v => String(v).trim());
            rowsRoot.forEach(item => rootSet.add(item));
            log(`Lista raiz do arquivo carregada com ${
                rootSet.size
            } valores.`);
        }

        log(`Histórico de CNPJs em memória com ${
            storedCnpjs.size
        } registros.`);
        if (checkDb) 
            log("Opção \"Consultar Banco de Dados\" está ATIVADA.");
        
        if (saveToDb) 
            log("Opção \"Salvar no Banco de Dados\" está ATIVADA.");
        
        if (autoAdjust) 
            log("Opção \"Ajustar Fones Pós-Limpeza\" está ATIVADA.");
        

        const allNewCnpjs = new Set();

        for (const fileObj of cleanFiles) {
            const newlyFoundInFile = await processFile(fileObj, rootSet, destCol, event, backup, checkDb, saveToDb, storedCnpjs);
            if (saveToDb && newlyFoundInFile.size > 0) {
                newlyFoundInFile.forEach(cnpj => allNewCnpjs.add(cnpj));
            }

            if (autoAdjust) {
                await runPhoneAdjustment(fileObj.path, event, false);
            }
        }

        if (saveToDb && allNewCnpjs.size > 0) {
            log(`\nEnviando ${
                allNewCnpjs.size
            } novos CNPJs para o banco de dados em lotes...`);
            const cnpjsArray = Array.from(allNewCnpjs);
            const batchSize = 499;
            const totalBatches = Math.ceil(cnpjsArray.length / batchSize);
            for (let i = 0; i < cnpjsArray.length; i += batchSize) {
                const chunk = cnpjsArray.slice(i, i + batchSize);
                const batch = db.batch();
                chunk.forEach(cnpj => {
                    const docRef = cnpjsCollection.doc(cnpj);
                    batch.set(docRef, {
                        numero: cnpj,
                        adicionado_em: new Date(),
                        batchId: batchId
                    });
                    storedCnpjs.add(cnpj);
                });
                await batch.commit();
                const currentBatchNum = Math.floor(i / batchSize) + 1;
                event.sender.send("upload-progress", {
                    current: currentBatchNum,
                    total: totalBatches
                });
            }
            log(`✅ Banco de dados atualizado. Total agora: ${
                storedCnpjs.size
            } registros.`);
            log(`✅ ID do Lote salvo: ${batchId} (use este ID para futuras exclusões)`);
        }

        log(`\n✅ Processo concluído para todos os arquivos.`);
    } catch (err) {
        log(`❌ Erro inesperado no processo de limpeza: ${
            err.message
        }`);
        console.error(err);
    }
});

async function processFile(fileObj, rootSet, destCol, event, backup, checkDb, saveToDb, cnpjsHistory) {
    const file = fileObj.path;
    const id = fileObj.id;
    const log = (msg) => event.sender.send("log", msg);
    const progress = (pct) => event.sender.send("progress", {id, progress: pct});
    log(`\nProcessando arquivo de limpeza: ${
        path.basename(file)
    }...`);
    if (! fs.existsSync(file)) 
        return new Set();
    
    if (backup) {
        const p = path.parse(file);
        const bkp = path.join(p.dir, `${
            p.name
        }.backup_${
            Date.now()
        }${
            p.ext
        }`);
        fs.copyFileSync(file, bkp);
        log(`Backup criado: ${bkp}`);
    }
    const wb = await readSpreadsheet(file);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, {header: 1});
    if (data.length <= 1) {
        log(`⚠️ Arquivo vazio ou sem dados: ${file}`);
        return new Set();
    }
    const header = data[0];

    const destColIdx = letterToIndex(destCol);
    const cpfColIdx = header.findIndex(h => String(h).trim().toLowerCase() === "cpf");

    if (cpfColIdx === -1) {
        log(`❌ ERRO: A coluna \"cpf\" não foi encontrada no arquivo ${
            path.basename(file)
        }. Pulando este arquivo.`);
        return new Set();
    }
    const foneIdxs = header.reduce((acc, cell, i) => {
        if (typeof cell === "string" && cell.trim().toLowerCase().startsWith("fone")) 
            acc.push(i);
        
        return acc;
    }, []);
    const cleaned = [header];
    let removedByRoot = 0;
    let removedDuplicates = 0;
    let cleanedPhones = 0;
    const totalRows = data.length - 1;
    const newCnpjsInThisFile = new Set();
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const key = row[destColIdx] ? String(row[destColIdx]).trim() : "";
        const cnpj = row[cpfColIdx] ? String(row[cpfColIdx]).trim().replace(/\D/g, "") : "";
        if (checkDb && cnpj && cnpjsHistory.has(cnpj)) {
            removedDuplicates++;
            continue;
        }
        if (key && rootSet.has(key)) {
            removedByRoot++;
            continue;
        }
        foneIdxs.forEach(idx => {
            const v = row[idx] ? String(row[idx]).trim() : "";
            if (/^\d{10}$/.test(v)) {
                row[idx] = "";
                cleanedPhones++;
            }
        });
        cleaned.push(row);
        if (saveToDb && cnpj && ! cnpjsHistory.has(cnpj)) {
            newCnpjsInThisFile.add(cnpj);
        }
        if (i % 2000 === 0) {
            progress(Math.floor((i / totalRows) * 100));
            await new Promise(resolve => setImmediate(resolve));
        }
    }
    const newWB = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWB, XLSX.utils.aoa_to_sheet(cleaned), wb.SheetNames[0]);
    writeSpreadsheet(newWB, file);
    progress(100);
    log(`Arquivo: ${
        path.basename(file)
    }\n • Clientes repetidos (BD): ${removedDuplicates}\n • Removidos pela Raiz: ${removedByRoot}\n • Fones limpos: ${cleanedPhones}\n • Total final: ${
        cleaned.length - 1
    }`);
    return newCnpjsInThisFile;
}

ipcMain.on("start-db-only-cleaning", async (event, {filesToClean, saveToDb}) => {
    const log = (msg) => event.sender.send("log", msg);
    const batchId = `batch-${
        Date.now()
    }`;
    log(`--- Iniciando Limpeza Apenas pelo Banco de Dados para ${
        filesToClean.length
    } arquivo(s) ---`);
    if (saveToDb) 
        log(`Opção \"Salvar no Banco de Dados\" ATIVADA. ID do Lote: ${batchId}`);
    

    log(`Usando ${
        storedCnpjs.size
    } CNPJs do histórico em memória.`);
    const allNewCnpjs = new Set();

    for (const filePath of filesToClean) {
        log(`\nProcessando: ${
            path.basename(filePath)
        }`);
        try {
            const wb = await readSpreadsheet(filePath);
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(sheet, {header: 1});
            if (data.length <= 1) {
                log(`⚠️ Arquivo vazio ou sem dados: ${filePath}`);
                continue;
            }
            const header = data[0];
            const cpfColIdx = header.findIndex(h => String(h).trim().toLowerCase() === "cpf");
            if (cpfColIdx === -1) {
                log(`❌ ERRO: A coluna \"cpf\" não foi encontrada em ${
                    path.basename(filePath)
                }. Pulando.`);
                continue;
            }
            let removedCount = 0;
            const cleaned = [header];
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                const cnpj = row[cpfColIdx] ? String(row[cpfColIdx]).trim().replace(/\D/g, "") : "";
                if (cnpj && storedCnpjs.has(cnpj)) {
                    removedCount++;
                    continue;
                }
                cleaned.push(row);
                if (saveToDb && cnpj) {
                    allNewCnpjs.add(cnpj);
                }
            }
            const newWB = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWB, XLSX.utils.aoa_to_sheet(cleaned), wb.SheetNames[0]);
            writeSpreadsheet(newWB, filePath);
            log(`✅ Arquivo ${
                path.basename(filePath)
            } concluído. Removidos: ${removedCount}. Total final: ${
                cleaned.length - 1
            }`);
        } catch (err) {
            log(`❌ Erro ao processar ${
                path.basename(filePath)
            }: ${
                err.message
            }`);
            console.error(err);
        }
    }

    if (saveToDb && allNewCnpjs.size > 0) {
        log(`\nEnviando ${
            allNewCnpjs.size
        } novos CNPJs (não encontrados no BD) para o banco de dados...`);
        const cnpjsArray = Array.from(allNewCnpjs);
        const batchSize = 499;
        const totalBatches = Math.ceil(cnpjsArray.length / batchSize);
        for (let i = 0; i < cnpjsArray.length; i += batchSize) {
            const chunk = cnpjsArray.slice(i, i + batchSize);
            const batch = db.batch();
            chunk.forEach(cnpj => {
                const docRef = cnpjsCollection.doc(cnpj);
                batch.set(docRef, {
                    numero: cnpj,
                    adicionado_em: new Date(),
                    batchId: batchId
                });
                storedCnpjs.add(cnpj);
            });
            await batch.commit();
            const currentBatchNum = Math.floor(i / batchSize) + 1;
            event.sender.send("upload-progress", {
                current: currentBatchNum,
                total: totalBatches
            });
        }
        log(`✅ Banco de dados atualizado. Total agora: ${
            storedCnpjs.size
        } registros.`);
        log(`✅ ID do Lote salvo: ${batchId} (use este ID para futuras exclusões)`);
    }

    log("\n--- Limpeza Apenas pelo Banco de Dados finalizada. ---");
});

ipcMain.on("start-merge", async (event, files) => {
    const log = (msg) => event.sender.send("log", msg);
    if (!files || files.length < 2) {
        log("❌ Erro: Por favor, selecione pelo menos dois arquivos para mesclar.");
        dialog.showErrorBox("Erro de Mesclagem", "Você precisa selecionar no mínimo dois arquivos para a mesclagem.");
        return;
    }
    log(`\n--- Iniciando Mesclagem de ${
        files.length
    } arquivos ---`);
    try {
        const {canceled, filePath: savePath} = await dialog.showSaveDialog(mainWindow, {
            title: "Salvar Arquivo Mesclado",
            defaultPath: `mesclado_${
                Date.now()
            }.xlsx`,
            filters: [
                {
                    name: "Planilhas Excel",
                    extensions: ["xlsx"]
                }
            ]
        });
        if (canceled || !savePath) {
            log("Operação de mesclagem cancelada pelo usuário.");
            return;
        }
        log(`Arquivo de destino: ${savePath}`);
        let allDataRows = [];
        let totalRows = 0;
        const firstFilePath = files[0];
        log(`Lendo arquivo base: ${
            path.basename(firstFilePath)
        }`);
        const firstWb = await readSpreadsheet(firstFilePath);
        const firstWs = firstWb.Sheets[firstWb.SheetNames[0]];
        const firstFileData = XLSX.utils.sheet_to_json(firstWs, {
            header: 1,
            defval: ""
        });
        allDataRows.push(... firstFileData);
        totalRows += firstFileData.length;
        log(`Adicionadas ${
            firstFileData.length
        } linhas (com cabeçalho) do arquivo base.`);
        for (let i = 1; i < files.length; i++) {
            const filePath = files[i];
            log(`Lendo arquivo para anexar: ${
                path.basename(filePath)
            }`);
            const wb = await readSpreadsheet(filePath);
            const ws = wb.Sheets[wb.SheetNames[0]];
            const fileData = XLSX.utils.sheet_to_json(ws, {
                header: 1,
                defval: ""
            }).slice(1);
            if (fileData.length > 0) {
                allDataRows.push(... fileData);
                totalRows += fileData.length;
                log(`Adicionadas ${
                    fileData.length
                } linhas de dados de ${
                    path.basename(filePath)
                }.`);
            } else {
                log(`⚠️ Arquivo ${
                    path.basename(filePath)
                } não continha dados além do cabeçalho.`);
            }
        }
        log(`\nTotal de linhas a serem escritas: ${totalRows}. Criando o arquivo final...`);
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.aoa_to_sheet(allDataRows);
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Mesclado");
        writeSpreadsheet(newWorkbook, savePath);
        log(`✅ Mesclagem concluída com sucesso! O arquivo foi salvo em: ${savePath}`);
        dialog.showMessageBox(mainWindow, {
            type: "info",
            title: "Sucesso",
            message: `Arquivos mesclados com sucesso!\n\nO resultado foi salvo em:\n${savePath}`
        });
    } catch (err) {
        log(`❌ Erro catastrófico durante a mesclagem: ${
            err.message
        }`);
        console.error(err);
        dialog.showErrorBox("Erro de Mesclagem", `Ocorreu um erro inesperado: ${
            err.message
        }`);
    }
});

ipcMain.on("start-adjust-phones", async (event, {filePath, backup}) => {
    const log = (msg) => event.sender.send("log", msg);
    log(`\n--- Iniciando Ajuste de Fones para ${
        path.basename(filePath)
    } ---`);
    await runPhoneAdjustment(filePath, event, backup);
    log(`\n✅ Ajuste de fones concluído para o arquivo.`);
});

let apiQueue = {
    pending: [],
    processing: null,
    completed: []
};
let isApiQueueRunning = false;
ipcMain.on("add-files-to-api-queue", (event, filePaths) => {
    apiQueue.pending.push(... filePaths);
    apiQueue.pending = [... new Set(apiQueue.pending)];
    event.sender.send("api-queue-update", apiQueue);
});
ipcMain.on("start-api-queue", (event, {keyMode}) => {
    if (isApiQueueRunning) 
        return;
    
    isApiQueueRunning = true;
    processNextInApiQueue(event, keyMode);
});
ipcMain.on("reset-api-queue", (event) => {
    apiQueue = {
        pending: [],
        processing: null,
        completed: []
    };
    isApiQueueRunning = false;
    event.sender.send("api-queue-update", apiQueue);
    event.sender.send("api-log", "Fila e status reiniciados.");
});
async function processNextInApiQueue(event, keyMode) {
    if (apiQueue.pending.length === 0) {
        event.sender.send("api-log", "\n✅ Fila de processamento concluída.");
        apiQueue.processing = null;
        isApiQueueRunning = false;
        event.sender.send("api-queue-update", apiQueue);
        return;
    }
    apiQueue.processing = apiQueue.pending.shift();
    event.sender.send("api-queue-update", apiQueue);

    event.sender.send("api-log", `--- Iniciando processamento de: ${
        path.basename(apiQueue.processing)
    } ---`);
    await runApiConsultation(apiQueue.processing, keyMode, (msg) => event.sender.send("api-log", msg), (current, total) => event.sender.send("api-progress", {current, total}));

    apiQueue.completed.push(apiQueue.processing);
    apiQueue.processing = null;
    event.sender.send("api-queue-update", apiQueue);
    processNextInApiQueue(event, keyMode);
}

async function runApiConsultation(filePath, keyMode, log, progress) {
    const credentials = {
        c6: {
            CLIENT_ID: "EA8ZUFeZVSeqMGr49XJSsZKFuxSZub3i",
            CLIENT_SECRET: "EUomxjGf6BvBZ1HO",
            name: "Chave 1 (Padrão)"
        },
        im: {
            CLIENT_ID: "imWzrW41HcnoJgvZqHCaLvziUGlhAJAH",
            CLIENT_SECRET: "A0lAqZO73uW3wryU",
            name: "Chave 2 (Alternativa)"
        }
    };
    const TOKEN_URL = "https://crm-leads-p.c6bank.info/querie-partner/token";
    const CONSULTA_URL = "https://crm-leads-p.c6bank.info/querie-partner/client/avaliable";
    const BATCH_SIZE = 20000;
    const RETRY_MS = 6 * 60 * 1000;
    const DELAY_SUCESSO_MS = 3 * 60 * 1000;
    const MAX_RETRIES = 3;
    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));
    const normalizeCnpj = (cnpj) => (String(cnpj).replace(/\D/g, "")).padStart(14, "0");
    try {
        log(`Iniciando processo com o modo de chave: '${keyMode}'.`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const worksheet = workbook.worksheets[0];
        const COLUNA_CNPJ = "cpf";
        let cnpjColNumber = -1;
        worksheet.getRow(1).eachCell({
            includeEmpty: true
        }, (cell, colNumber) => {
            if (cell.value && String(cell.value).trim().toLowerCase() === COLUNA_CNPJ) {
                cnpjColNumber = colNumber;
            }
        });
        if (cnpjColNumber === -1) 
            throw new Error(`A coluna "${COLUNA_CNPJ}" não foi encontrada.`);
        
        const COLUNA_RESPOSTA_LETTER = "C";
        const registros = [];
        worksheet.eachRow({
            includeEmpty: false
        }, (row, rowNum) => {
            if (rowNum > 1) {
                const cnpjCell = row.getCell(cnpjColNumber);
                const cnpjValue = cnpjCell.value;
                const respostaCell = row.getCell(COLUNA_RESPOSTA_LETTER);
                if (! respostaCell.value) {
                    registros.push({
                        cnpj: normalizeCnpj(cnpjValue || ""),
                        rowNum
                    });
                }
            }
        });
        if (registros.length === 0) {
            log("✅ Nenhum registro novo para consultar neste arquivo.");
            return;
        }
        log(`Encontrados ${
            registros.length
        } registros novos para processar.`);
        const lotes = [];
        for (let i = 0; i < registros.length; i += BATCH_SIZE) {
            lotes.push(registros.slice(i, i + BATCH_SIZE));
        }
        for (let i = 0; i < lotes.length; i++) {
            const lote = lotes[i];
            log(`\n=== Processando Lote ${
                i + 1
            }/${
                lotes.length
            } (${
                lote.length
            } registros) ===`);
            progress(i + 1, lotes.length);
            let currentCreds;
            if (keyMode === "intercalar") {
                currentCreds = i % 2 === 0 ? credentials.c6 : credentials.im;
                log(`Usando credenciais intercaladas: ${
                    currentCreds.name
                }`);
            } else if (keyMode === "chave2") {
                currentCreds = credentials.im;
            } else {
                currentCreds = credentials.c6;
            }
            if (keyMode !== "intercalar" && i === 0) {
                log(`Usando credenciais fixas: ${
                    currentCreds.name
                }`);
            }
            let sucesso = false;
            let retries = 0;
            while (!sucesso && retries < MAX_RETRIES) {
                try {
                    log("Gerando token de acesso...");
                    const tokenParams = new URLSearchParams({
                        grant_type: "client_credentials",
                        client_id: currentCreds.CLIENT_ID,
                        client_secret: currentCreds.CLIENT_SECRET
                    });
                    const tokenResp = await axios.post(TOKEN_URL, tokenParams.toString(), {
                        headers: {
                            "Content-Type": "application/x-www-form-urlencoded"
                        },
                        timeout: 30000
                    });
                    const token = tokenResp.data.access_token;
                    log("Consultando API...");
                    const cnpjArray = lote.map(r => r.cnpj);
                    const consultaResp = await axios.post(CONSULTA_URL, {
                        CNPJ: cnpjArray
                    }, {
                        headers: {
                            Authorization: `Bearer ${token}`,
                            "Content-Type": "application/json"
                        },
                        timeout: 30000
                    });
                    const key = Object.keys(consultaResp.data).find(k => k.toLowerCase().includes("cnpj") && Array.isArray(consultaResp.data[k]));
                    const encontrados = key ? new Set(consultaResp.data[key].map(normalizeCnpj)) : new Set();
                    log(`Atualizando planilha em memória...`);
                    let countDisponivel = 0;
                    lote.forEach(({cnpj, rowNum}) => {
                        if (encontrados.has(cnpj)) {
                            worksheet.getCell(`${COLUNA_RESPOSTA_LETTER}${rowNum}`).value = "disponível";
                            countDisponivel++;
                        } else {
                            worksheet.getCell(`${COLUNA_RESPOSTA_LETTER}${rowNum}`).value = "cliente";
                        }
                    });
                    const countCliente = lote.length - countDisponivel;
                    log(`Resultados do Lote: ${countDisponivel} disponível(is), ${countCliente} cliente(s).`);
                    log(`💾 Salvando progresso do lote ${
                        i + 1
                    } na planilha...`);
                    const tempFilePath = path.join(path.dirname(filePath), `${
                        path.basename(filePath, ".xlsx")
                    }_temp.xlsx`);
                    await workbook.xlsx.writeFile(tempFilePath);
                    fs.unlinkSync(filePath);
                    fs.renameSync(tempFilePath, filePath);
                    log(`✅ Progresso salvo com sucesso.`);
                    sucesso = true;
                } catch (err) {
                    retries++;
                    log(`❌ Erro no processamento do lote (tentativa ${retries}/${MAX_RETRIES}): ${
                        err.message
                    }.`);
                    if (retries < MAX_RETRIES) {
                        log(`Tentando novamente em ${
                            RETRY_MS / 60000
                        } minutos...`);
                        await sleep(RETRY_MS);
                    } else {
                        log(`Máximo de tentativas atingido para este lote. Pulando para o próximo.`);
                    }
                }
            }
            if (i < lotes.length - 1) {
                log(`Aguardando ${
                    DELAY_SUCESSO_MS / 60000
                } minutos antes do próximo lote...`);
                await sleep(DELAY_SUCESSO_MS);
            }
        }
        log(`\n🎉 Arquivo ${
            path.basename(filePath)
        } processado e salvo.`);
    } catch (error) {
        log(`❌ Erro fatal ao processar o arquivo ${
            path.basename(filePath)
        }: ${
            error.message
        }`);
        console.error(error);
    }
}