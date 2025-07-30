document.addEventListener('DOMContentLoaded', () => {
    const getBasename = (p) => p.split(/[\\/]/).pop();

    const tabButtons = document.querySelectorAll('.tab-button');
    const tabContents = document.querySelectorAll('.tab-content');
    const currentTabTitle = document.getElementById('current-tab-title');
    const currentTabDescription = document.getElementById('current-tab-description');
    const mainTitle = document.getElementById("main-app-title");

    window.electronAPI.onUpdateMessage((message) => {
        const updateMessageElement = document.getElementById('update-message');
        if (updateMessageElement) {
            updateMessageElement.innerText = message;
        }
    });

    const gridConfigs = {
        'localGrid': 'local-sections',
        'apiGrid': 'api-sections',
        'enrichmentGrid': 'enrichment-sections',
    };

    let sortableInstances = {};

    function saveSectionState(tabId, sections) {
        const state = sections.map(section => ({
            id: section.dataset.sectionId,
            visible: !section.classList.contains('hidden')
        }));
        localStorage.setItem(`sections-${tabId}`, JSON.stringify(state));
    }

    function loadSectionState(tabId) {
        const saved = localStorage.getItem(`sections-${tabId}`);
        return saved ? JSON.parse(saved) : null;
    }

    function applySectionState(grid, state) {
        if (!state) return;

        const sections = Array.from(grid.querySelectorAll('.section'));

        state.forEach((savedSection, index) => {
            const section = sections.find(s => s.dataset.sectionId === savedSection.id);
            if (section) {
                grid.appendChild(section);
                if (!savedSection.visible) {
                    section.classList.add('hidden');
                    const hideBtn = section.querySelector('.hide-section');
                    if (hideBtn) {
                        hideBtn.textContent = '👁‍🗨';
                        hideBtn.title = 'Exibir';
                    }
                } else {
                    const hideBtn = section.querySelector('.hide-section');
                    if (hideBtn) {
                        hideBtn.textContent = '👁';
                        hideBtn.title = 'Ocultar';
                    }
                }
            }
        });
    }

    function initializeSortable(gridId, storageKey) {
        const grid = document.getElementById(gridId);
        if (!grid) return;

        const savedState = loadSectionState(storageKey);
        applySectionState(grid, savedState);

        const sortable = Sortable.create(grid, {
            animation: 300,
            ghostClass: 'sortable-ghost',
            chosenClass: 'sortable-chosen',
            dragClass: 'sortable-drag',
            handle: '.drag-handle',
            onEnd: function (evt) {
                const sections = Array.from(grid.querySelectorAll('.section'));
                saveSectionState(storageKey, sections);
            }
        });

        sortableInstances[gridId] = sortable;

        grid.addEventListener('click', function (e) {
            if (e.target.classList.contains('hide-section')) {
                const section = e.target.closest('.section');
                const isHidden = section.classList.contains('hidden');

                if (isHidden) {
                    section.classList.remove('hidden');
                    e.target.textContent = '👁';
                    e.target.title = 'Ocultar';
                } else {
                    section.classList.add('hidden');
                    e.target.textContent = '👁‍🗨';
                    e.target.title = 'Exibir';
                }
                const sections = Array.from(grid.querySelectorAll('.section'));
                saveSectionState(storageKey, sections);
            }
        });
    }

    Object.entries(gridConfigs).forEach(([gridId, storageKey]) => {
        initializeSortable(gridId, storageKey);
    });

    const tabInfo = {
        'Limpeza Local': {
            title: 'Limpeza Local de Bases',
            description: 'Otimize suas bases de dados localmente, removendo duplicidades e ajustando informações com precisão. Ideal para manter seus registros impecáveis.'
        },
        'Consulta CNPJ (API)': {
            title: 'Consulta CNPJ via API',
            description: 'Realize consultas de CNPJ diretamente via API para obter a situação da empresa.'
        },
        'Enriquecimento': {
            title: 'Enriquecimento de Dados',
            description: 'Amplie suas bases com informações valiosas, como telefones e outros dados de contato, a partir de fontes confiáveis. Maximize o potencial de suas campanhas.'
        },
        'Monitoramento': {
            title: 'Monitoramento de Relatórios',
            description: 'Acompanhe os dados de chamadas em tempo real. Filtre e visualize informações para análise de performance e tomada de decisão.'
        }
    };

    function openTab(evt, tabNameId) {
        tabContents.forEach(content => {
            content.classList.remove('active');
            content.style.display = 'none';
        });

        tabButtons.forEach(button => {
            button.classList.remove('active');
        });

        const currentTabContent = document.getElementById(tabNameId);
        if (currentTabContent) {
            currentTabContent.style.display = 'block';
            currentTabContent.classList.add('active');
        }
        if (evt && evt.currentTarget) {
            evt.currentTarget.classList.add('active');
        }

        const body = document.body;
        body.classList.remove('c6-theme', 'enrichment-theme', 'monitoring-theme');
        if (tabNameId === 'api') {
            body.classList.add('c6-theme');
        } else if (tabNameId === 'enriquecimento') {
            body.classList.add('enrichment-theme');
        } else if (tabNameId === 'monitoramento') {
            body.classList.add('monitoring-theme');
        }

        const tabButtonText = evt ? evt.currentTarget.textContent.trim() : tabButtons[0].textContent.trim();
        if (tabInfo[tabButtonText]) {
            mainTitle.classList.add("hidden");
            currentTabTitle.textContent = tabInfo[tabButtonText].title;
            currentTabDescription.textContent = tabInfo[tabButtonText].description;
        } else {
            mainTitle.classList.remove("hidden");
            currentTabTitle.textContent = "";
            currentTabDescription.textContent = "";
        }
    }

    tabButtons.forEach(button => {
        button.addEventListener('click', (event) => {
            const buttonText = event.currentTarget.textContent.trim();
            let tabNameId = '';

            if (buttonText.includes('Local')) {
                tabNameId = 'local';
            } else if (buttonText.includes('API')) {
                tabNameId = 'api';
            } else if (buttonText.includes('Enriquecimento')) {
                tabNameId = 'enriquecimento';
            } else if (buttonText.includes('Monitoramento')) {
                tabNameId = 'monitoramento';
            }

            if (tabNameId) {
                openTab(event, tabNameId);
            }
        });
    });

    if (tabButtons.length > 0) {
        tabButtons[0].click();
    }

    // #################################################################
    // #           LÓGICA DA ABA DE LIMPEZA LOCAL E OUTRAS             #
    // #################################################################

    let rootFile = null;
    let cleanFiles = [];
    let mergeFiles = [];
    let backupEnabled = false;
    let autoAdjustPhones = false;
    let checkDbEnabled = false;
    let saveToDbEnabled = false;

    const selectRootBtn = document.getElementById('selectRootBtn');
    const autoRootBtn = document.getElementById('autoRootBtn');
    const feedRootBtn = document.getElementById('feedRootBtn');
    const updateBlocklistBtn = document.getElementById('updateBlocklistBtn');
    const addCleanFileBtn = document.getElementById('addCleanFileBtn');
    const startCleaningBtn = document.getElementById('startCleaningBtn');
    const resetLocalBtn = document.getElementById('resetLocalBtn');
    const adjustPhonesBtn = document.getElementById('adjustPhonesBtn');
    const backupCheckbox = document.getElementById('backupCheckbox').parentElement;
    const autoAdjustPhonesCheckbox = document.getElementById('autoAdjustPhonesCheckbox');
    const rootFilePathSpan = document.getElementById('rootFilePath');
    const selectedCleanFilesDiv = document.getElementById('selectedCleanFiles');
    const progressContainer = document.getElementById('progressContainer');
    const logDiv = document.getElementById('log');
    const rootColSelect = document.getElementById('rootCol');
    const destColSelect = document.getElementById('destCol');
    const selectMergeFilesBtn = document.getElementById('selectMergeFilesBtn');
    const startMergeBtn = document.getElementById('startMergeBtn');
    const selectedMergeFilesDiv = document.getElementById('selectedMergeFiles');
    const saveStoredCnpjsBtn = document.getElementById('saveStoredCnpjsBtn');
    const checkDbCheckbox = document.getElementById('checkDbCheckbox');
    const saveToDbCheckbox = document.getElementById('saveToDbCheckbox');
    const consultDbBtn = document.getElementById('consultDbBtn');
    const uploadProgressContainer = document.getElementById('uploadProgressContainer');
    const uploadProgressTitle = document.getElementById('uploadProgressTitle');
    const uploadProgressBarFill = document.getElementById('uploadProgressBarFill');
    const uploadProgressText = document.getElementById('uploadProgressText');
    const batchIdInput = document.getElementById('batchIdInput');
    const deleteBatchBtn = document.getElementById('deleteBatchBtn');

    function addFileToUI(container, filePath, isSingle) {
        if (isSingle) {
            container.innerHTML = '';
        }
        const fileDiv = document.createElement('div');
        fileDiv.className = 'file-item new-item';
        fileDiv.textContent = getBasename(filePath);
        container.appendChild(fileDiv);
        setTimeout(() => {
            fileDiv.classList.remove('new-item');
        }, 500);
    }

    function resetUploadProgress() {
        if (uploadProgressContainer) uploadProgressContainer.style.display = 'none';
        if (uploadProgressBarFill) uploadProgressBarFill.style.width = '0%';
        if (uploadProgressText) uploadProgressText.textContent = '';
    }

    if (backupCheckbox) backupCheckbox.addEventListener('change', (e) => { backupEnabled = e.target.querySelector('input').checked; });
    if (autoAdjustPhonesCheckbox) autoAdjustPhonesCheckbox.addEventListener('change', () => { autoAdjustPhones = autoAdjustPhonesCheckbox.checked; });

    if (checkDbCheckbox) checkDbCheckbox.addEventListener('change', () => {
        checkDbEnabled = checkDbCheckbox.checked;
        appendLog(`Consulta ao Banco de Dados: ${checkDbEnabled ? 'ATIVADA' : 'DESATIVADA'}`);
    });

    if (saveToDbCheckbox) saveToDbCheckbox.addEventListener('change', () => {
        saveToDbEnabled = saveToDbCheckbox.checked;
        appendLog(`Salvar no Banco de Dados: ${saveToDbEnabled ? 'ATIVADO' : 'DESATIVADO'}`);
    });

    if (saveStoredCnpjsBtn) saveStoredCnpjsBtn.addEventListener('click', async () => {
        appendLog('Solicitando salvamento do histórico de CNPJs em Excel...');
        const result = await window.electronAPI.saveStoredCnpjsToExcel();
        appendLog(result.message);
    });

    if (deleteBatchBtn) deleteBatchBtn.addEventListener('click', async () => {
        const batchId = batchIdInput.value.trim();
        if (!batchId) {
            appendLog('❌ ERRO: Por favor, insira um ID de Lote para excluir.');
            return;
        }
        const confirmation = confirm(`ATENÇÃO!\n\nVocê tem certeza que deseja excluir PERMANENTEMENTE todos os CNPJs do lote "${batchId}" do banco de dados?\n\nEsta ação não pode ser desfeita.`);
        if (confirmation) {
            appendLog(`Enviando solicitação para excluir o lote: ${batchId}...`);
            const result = await window.electronAPI.deleteBatch(batchId);
            appendLog(result.message);
            if (result.success) {
                batchIdInput.value = '';
            }
        } else {
            appendLog('Operação de exclusão cancelada pelo usuário.');
        }
    });

    if (consultDbBtn) consultDbBtn.addEventListener('click', async () => {
        appendLog('Selecionando arquivos para consulta apenas pelo BD...');
        const files = await window.electronAPI.selectFile({ title: 'Selecione arquivos para limpar apenas pelo BD', multi: true });
        if (!files || files.length === 0) {
            appendLog('Nenhum arquivo selecionado.');
            return;
        }
        window.electronAPI.startDbOnlyCleaning({
            filesToClean: files,
            saveToDb: saveToDbEnabled
        });
    });

    if (selectRootBtn) selectRootBtn.addEventListener('click', async () => {
        const files = await window.electronAPI.selectFile({ title: 'Selecione a Lista Raiz', multi: false });
        if (files && files.length > 0) {
            rootFile = files[0];
            addFileToUI(rootFilePathSpan, rootFile, true);
            appendLog(`Arquivo raiz selecionado: ${rootFile}`);
        }
    });

    if (autoRootBtn) autoRootBtn.addEventListener('click', () => {
        if (autoRootBtn.dataset.on) {
            delete autoRootBtn.dataset.on;
            autoRootBtn.textContent = "Auto Raiz: OFF";
            rootFile = null;
            rootFilePathSpan.innerHTML = '<span style="color:var(--text-muted); font-style:italic;">Usará arquivo local selecionado</span>';
            selectRootBtn.disabled = false;
        } else {
            autoRootBtn.dataset.on = 'true';
            autoRootBtn.textContent = "Auto Raiz: ON";
            rootFile = null; // Limpa o arquivo, pois usará o BD
            rootFilePathSpan.innerHTML = '<span style="color:var(--accent-light); font-weight: 600;">Usará a base de dados Raiz</span>';
            selectRootBtn.disabled = true;
        }
        appendLog(`Auto Raiz: ${autoRootBtn.dataset.on ? 'ON (usando Banco de Dados)' : 'OFF'}`);
    });

    if (updateBlocklistBtn) updateBlocklistBtn.addEventListener('click', async () => {
        const result = await window.electronAPI.updateBlocklist(backupEnabled);
        appendLog(result.success ? result.message : `Erro: ${result.message}`);
    });

    if (addCleanFileBtn) addCleanFileBtn.addEventListener('click', async () => {
        const files = await window.electronAPI.selectFile({ title: 'Selecione arquivos para limpar', multi: true });
        if (!files?.length) return;
        cleanFiles = [];
        selectedCleanFilesDiv.innerHTML = '';
        progressContainer.innerHTML = '';
        files.forEach(file => {
            const id = `clean-${cleanFiles.length}`;
            cleanFiles.push({ path: file, id });
            appendLog(`Adicionado para limpeza: ${file}`);
            addFileToUI(selectedCleanFilesDiv, file, false);
            progressContainer.innerHTML += `<div class="file-progress" style="margin-bottom: 15px;"><strong>${getBasename(file)}</strong><div class="progress-bar-container"><div class="progress-bar-fill" id="${id}"></div></div></div>`;
        });
    });

    if (startCleaningBtn) startCleaningBtn.addEventListener('click', () => {
        const isAutoRoot = autoRootBtn.dataset.on === 'true';
        if (!isAutoRoot && !rootFile) {
            return appendLog('ERRO: Selecione o arquivo raiz ou ative o Auto Raiz.');
        }
        if (!cleanFiles.length) {
            return appendLog('ERRO: Adicione ao menos um arquivo para limpar.');
        }

        resetUploadProgress();
        appendLog('Iniciando limpeza...');
        window.electronAPI.startCleaning({
            isAutoRoot,
            rootFile: isAutoRoot ? null : rootFile,
            cleanFiles,
            rootCol: rootColSelect.value,
            destCol: destColSelect.value,
            backup: backupEnabled,
            checkDb: checkDbEnabled,
            saveToDb: saveToDbEnabled,
            autoAdjust: autoAdjustPhones
        });
    });

    if (resetLocalBtn) resetLocalBtn.addEventListener('click', () => {
        rootFile = null;
        cleanFiles = [];
        mergeFiles = [];
        backupEnabled = false;
        autoAdjustPhones = false;
        checkDbEnabled = false;
        saveToDbEnabled = false;
        if (rootFilePathSpan) rootFilePathSpan.innerHTML = '';
        if (selectedCleanFilesDiv) selectedCleanFilesDiv.innerHTML = '';
        if (progressContainer) progressContainer.innerHTML = '';
        if (logDiv) logDiv.textContent = '';
        if (selectedMergeFilesDiv) selectedMergeFilesDiv.innerHTML = '';
        if (batchIdInput) batchIdInput.value = '';
        if (backupCheckbox) backupCheckbox.querySelector('input').checked = false;
        if (autoAdjustPhonesCheckbox) autoAdjustPhonesCheckbox.checked = false;
        if (checkDbCheckbox) checkDbCheckbox.checked = false;
        if (saveToDbCheckbox) saveToDbCheckbox.checked = false;
        if (autoRootBtn) {
            delete autoRootBtn.dataset.on;
            autoRootBtn.textContent = 'Auto Raiz: OFF';
            selectRootBtn.disabled = false;
        }
        resetUploadProgress();
        appendLog('Módulo de Limpeza Local reiniciado.');
    });

    if (adjustPhonesBtn) adjustPhonesBtn.addEventListener('click', async () => {
        const files = await window.electronAPI.selectFile({ title: 'Selecione arquivo para ajustar fones', multi: false });
        if (!files?.length) return appendLog('Nenhum arquivo selecionado.');
        window.electronAPI.startAdjustPhones({ filePath: files[0], backup: backupEnabled });
    });

    if (selectMergeFilesBtn) selectMergeFilesBtn.addEventListener('click', async () => {
        const files = await window.electronAPI.selectFile({ title: 'Selecione arquivos para mesclar', multi: true });
        if (!files?.length) return;
        mergeFiles = files;
        selectedMergeFilesDiv.innerHTML = '';
        files.forEach(f => {
            addFileToUI(selectedMergeFilesDiv, f, false);
        });
    });

    if (startMergeBtn) startMergeBtn.addEventListener('click', () => {
        if (!mergeFiles.length) return appendLog('ERRO: Selecione arquivos para mesclar.');
        appendLog('Iniciando mesclagem...');
        window.electronAPI.startMerge(mergeFiles);
    });

    if (feedRootBtn) feedRootBtn.addEventListener('click', async () => {
        appendLog('Selecionando arquivos para alimentar a base Raiz...');
        const files = await window.electronAPI.selectFile({ title: 'Selecione planilhas com CNPJs para a Raiz', multi: true });
        if (!files || files.length === 0) {
            appendLog('Nenhum arquivo selecionado. Operação cancelada.');
            return;
        }
        feedRootBtn.disabled = true;
        appendLog(`Iniciando o processo de alimentação da Raiz com ${files.length} arquivo(s).`);
        window.electronAPI.feedRootDatabase(files);
    });

    window.electronAPI.onRootFeedFinished(() => {
        if (feedRootBtn) feedRootBtn.disabled = false;
        appendLog('✅ Processo de alimentação da Raiz finalizado.');
    });


    window.electronAPI.onLog((msg) => appendLog(msg));
    window.electronAPI.onProgress(({ id, progress }) => {
        const bar = document.getElementById(id);
        if (bar) bar.style.width = `${progress}%`;
    });
    window.electronAPI.onUploadProgress(({ current, total }) => {
        uploadProgressContainer.style.display = 'block';
        uploadProgressTitle.textContent = 'Enviando para o Banco de Dados Compartilhado:';
        const percent = Math.round((current / total) * 100);
        uploadProgressBarFill.style.width = `${percent}%`;
        uploadProgressText.textContent = `Enviando lote ${current} de ${total}...`;
        if (current === total) {
            uploadProgressTitle.textContent = 'Envio para o Banco de Dados Concluído!';
        }
    });

    function appendLog(msg) {
        if (!logDiv) return;
        if (logDiv.textContent === 'Aguardando início do sistema...') {
            logDiv.innerHTML = '';
        }
        logDiv.innerHTML += `> ${msg.replace(/\n/g, '<br>> ')}\n`;
        logDiv.scrollTop = logDiv.scrollHeight;
    }

    const apiDropzone = document.getElementById('apiDropzone');
    const apiProcessingDiv = document.getElementById('apiProcessing');
    const apiPendingDiv = document.getElementById('apiPending');
    const apiCompletedDiv = document.getElementById('apiCompleted');
    const apiKeySelection = document.getElementById('apiKeySelection');
    const startApiBtn = document.getElementById('startApiBtn');
    const resetApiBtn = document.getElementById('resetApiBtn');
    const selectApiFileBtn = document.getElementById('selectApiFileBtn');
    const apiStatusSpan = document.getElementById('apiStatus');
    const apiProgressBarFill = document.getElementById('apiProgressBarFill');
    const apiLogDiv = document.getElementById('apiLog');

    if (apiDropzone) {
        apiDropzone.addEventListener('dragover', (event) => { event.preventDefault(); event.stopPropagation(); apiDropzone.style.borderColor = 'var(--accent-color)'; apiDropzone.style.backgroundColor = 'var(--bg-lighter)'; });
        apiDropzone.addEventListener('dragleave', (event) => { event.preventDefault(); event.stopPropagation(); apiDropzone.style.borderColor = 'var(--border-color)'; apiDropzone.style.backgroundColor = 'transparent'; });
        apiDropzone.addEventListener('drop', (event) => { event.preventDefault(); event.stopPropagation(); apiDropzone.style.borderColor = 'var(--border-color)'; apiDropzone.style.backgroundColor = 'transparent'; const files = Array.from(event.dataTransfer.files).filter(file => file.path.endsWith('.xlsx') || file.path.endsWith('.xls') || file.path.endsWith('.csv')).map(file => file.path); if (files.length > 0) { window.electronAPI.addFilesToApiQueue(files); } });
    }

    if (selectApiFileBtn) selectApiFileBtn.addEventListener('click', async () => { const files = await window.electronAPI.selectFile({ title: 'Selecione as planilhas de CNPJs', multi: true }); if (files && files.length > 0) { window.electronAPI.addFilesToApiQueue(files); } });
    if (startApiBtn) startApiBtn.addEventListener('click', () => { startApiBtn.disabled = true; resetApiBtn.disabled = true; apiStatusSpan.textContent = 'Iniciando processamento da fila...'; window.electronAPI.startApiQueue({ keyMode: apiKeySelection.value }); });
    if (resetApiBtn) resetApiBtn.addEventListener('click', () => { window.electronAPI.resetApiQueue(); });

    function updateApiQueueUI(queue) {
        const { pending, processing, completed } = queue;
        if (!apiProcessingDiv || !apiPendingDiv || !apiCompletedDiv || !startApiBtn) return;

        apiProcessingDiv.innerHTML = processing ? '' : `<span style="color:var(--text-secondary)">Nenhum</span>`;
        if (processing) addFileToUI(apiProcessingDiv, processing, true);

        apiPendingDiv.innerHTML = pending.length > 0 ? '' : `<span style="color:var(--text-secondary)">Nenhum arquivo na fila</span>`;
        if (pending.length > 0) pending.forEach(file => addFileToUI(apiPendingDiv, file, false));

        apiCompletedDiv.innerHTML = completed.length > 0 ? '' : `<span style="color:var(--text-secondary)">Nenhum arquivo concluído</span>`;
        if (completed.length > 0) completed.forEach(file => addFileToUI(apiCompletedDiv, file, false));

        startApiBtn.disabled = pending.length === 0 || !!processing;
    }

    window.electronAPI.onApiQueueUpdate((queue) => { updateApiQueueUI(queue); if (!queue.processing && queue.pending.length === 0 && queue.completed.length > 0) { apiStatusSpan.textContent = 'Fila concluída!'; resetApiBtn.disabled = false; } });
    window.electronAPI.onApiLog((message) => { appendApiLog(message); });
    window.electronAPI.onApiProgress(({ current, total }) => { const percent = Math.round((current / total) * 100); apiProgressBarFill.style.width = `${percent}%`; apiStatusSpan.textContent = `Processando Lote ${current} de ${total}`; });

    function appendApiLog(msg) {
        if (apiLogDiv) {
            apiLogDiv.innerHTML += `> ${msg.replace(/\n/g, '<br>> ')}\n`;
            apiLogDiv.scrollTop = apiLogDiv.scrollHeight;
        }
    }

    const selectMasterFilesBtn = document.getElementById('selectMasterFilesBtn');
    const selectedMasterFilesDiv = document.getElementById('selectedMasterFiles');
    const startLoadToDbBtn = document.getElementById('startLoadToDbBtn');
    const selectEnrichFilesBtn = document.getElementById('selectEnrichFilesBtn');
    const selectedEnrichFilesDiv = document.getElementById('selectedEnrichFiles');
    const startEnrichmentBtn = document.getElementById('startEnrichmentBtn');
    const enrichmentLogDiv = document.getElementById('enrichmentLog');
    const enrichmentProgressContainer = document.getElementById('enrichmentProgressContainer');
    const enrichedCnpjCountSpan = document.getElementById('enrichedCnpjCount');
    const refreshCountBtn = document.getElementById('refreshCountBtn');
    const downloadEnrichedDataBtn = document.getElementById('downloadEnrichedDataBtn');

    const dbLoadProgressContainer = document.getElementById('dbLoadProgressContainer');
    const dbLoadProgressTitle = document.getElementById('dbLoadProgressTitle');
    const dbLoadProgressPercent = document.getElementById('dbLoadProgressPercent');
    const dbLoadProgressBarFill = document.getElementById('dbLoadProgressBarFill');
    const dbLoadProgressText = document.getElementById('dbLoadProgressText');
    const dbLoadProgressStats = document.getElementById('dbLoadProgressStats');

    let enrichmentMasterFiles = [];
    let enrichmentEnrichFiles = [];

    function appendEnrichmentLog(msg) {
        if (!enrichmentLogDiv) return;
        if (enrichmentLogDiv.textContent === 'Aguardando início...') {
            enrichmentLogDiv.innerHTML = '';
        }
        enrichmentLogDiv.innerHTML += `> ${msg.replace(/\n/g, '<br>> ')}\n`;
        enrichmentLogDiv.scrollTop = enrichmentLogDiv.scrollHeight;
    }

    async function updateEnrichedCnpjCount() {
        if (!enrichedCnpjCountSpan) return;
        try {
            enrichedCnpjCountSpan.textContent = 'Carregando...';
            const count = await window.electronAPI.getEnrichedCnpjCount();
            enrichedCnpjCountSpan.textContent = count.toLocaleString('pt-BR');
        } catch (error) {
            enrichedCnpjCountSpan.textContent = 'Erro';
            appendEnrichmentLog(`❌ Erro ao carregar contador: ${error.message}`);
        }
    }

    if (enrichedCnpjCountSpan) updateEnrichedCnpjCount();

    if (refreshCountBtn) refreshCountBtn.addEventListener('click', updateEnrichedCnpjCount);

    if (downloadEnrichedDataBtn) downloadEnrichedDataBtn.addEventListener('click', async () => {
        downloadEnrichedDataBtn.disabled = true;
        downloadEnrichedDataBtn.textContent = 'Preparando download...';
        try {
            const result = await window.electronAPI.downloadEnrichedData();
            if (result.success) {
                appendEnrichmentLog(`✅ ${result.message}`);
            } else {
                appendEnrichmentLog(`❌ ${result.message}`);
            }
        } catch (error) {
            appendEnrichmentLog(`❌ Erro no download: ${error.message}`);
        } finally {
            downloadEnrichedDataBtn.disabled = false;
            downloadEnrichedDataBtn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16">
                <path d="M.5 9.9a.5.5 0 0 1 .5.5v2.5a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-2.5a.5.5 0 0 1 1 0v2.5a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2v-2.5a.5.5 0 0 1 .5-.5z"/>
                <path d="M7.646 1.146a.5.5 0 0 1 .708 0l3 3a.5.5 0 0 1-.708.708L8.5 2.707V11.5a.5.5 0 0 1-1 0V2.707L5.354 4.854a.5.5 0 1 1-.708-.708l3-3z"/>
            </svg>Baixar Dados Enriquecidos`;
        }
    });

    if (selectMasterFilesBtn) selectMasterFilesBtn.addEventListener('click', async () => {
        const files = await window.electronAPI.selectFile({ title: 'Selecione as Planilhas Mestras', multi: true });
        if (!files?.length) return;
        enrichmentMasterFiles = files;
        selectedMasterFilesDiv.innerHTML = '';
        files.forEach(file => {
            addFileToUI(selectedMasterFilesDiv, file, false);
        });
    });

    if (startLoadToDbBtn) startLoadToDbBtn.addEventListener('click', () => {
        if (enrichmentMasterFiles.length === 0) return appendEnrichmentLog('❌ ERRO: Selecione pelo menos uma planilha mestra.');
        startLoadToDbBtn.disabled = true;

        dbLoadProgressContainer.style.display = 'block';
        dbLoadProgressBarFill.style.width = '0%';
        dbLoadProgressPercent.textContent = '0%';
        dbLoadProgressText.textContent = 'Iniciando...';
        dbLoadProgressStats.textContent = '';

        appendEnrichmentLog('Iniciando carga para o banco de dados...');
        window.electronAPI.startDbLoad({ masterFiles: enrichmentMasterFiles });
    });

    if (selectEnrichFilesBtn) selectEnrichFilesBtn.addEventListener('click', async () => {
        const files = await window.electronAPI.selectFile({ title: 'Selecione Arquivos para Enriquecer', multi: true });
        if (!files?.length) return;

        window.electronAPI.prepareEnrichmentFiles(files);

        enrichmentEnrichFiles = [];
        selectedEnrichFilesDiv.innerHTML = '';
        enrichmentProgressContainer.innerHTML = '';
        files.forEach(file => {
            const id = `enrich-${enrichmentEnrichFiles.length}`;
            enrichmentEnrichFiles.push({ path: file, id });
            appendEnrichmentLog(`Adicionado para enriquecimento: ${file}`);
            addFileToUI(selectedEnrichFilesDiv, file, false);
            enrichmentProgressContainer.innerHTML += `
                <div class="file-progress" style="margin-bottom: 15px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
                        <strong>${getBasename(file)}</strong>
                        <span id="eta-${id}" style="font-size: 12px; color: var(--text-secondary);"></span>
                    </div>
                    <div class="progress-bar-container">
                        <div class="progress-bar-fill" id="${id}"></div>
                    </div>
                </div>`;
        });
    });

    if (startEnrichmentBtn) startEnrichmentBtn.addEventListener('click', () => {
        if (enrichmentEnrichFiles.length === 0) return appendEnrichmentLog('❌ ERRO: Selecione pelo menos um arquivo para enriquecer.');
        startEnrichmentBtn.disabled = true;
        const strategy = document.querySelector('input[name="enrichStrategy"]:checked').value;
        const backup = document.getElementById('backupCheckbox').checked;
        appendEnrichmentLog(`Iniciando enriquecimento com a estratégia: ${strategy.toUpperCase()}`);
        window.electronAPI.startEnrichment({ filesToEnrich: enrichmentEnrichFiles, strategy, backup });
    });

    window.electronAPI.onEnrichmentLog((msg) => appendEnrichmentLog(msg));

    window.electronAPI.onEnrichmentProgress(({ id, progress, eta }) => {
        const bar = document.getElementById(id);
        if (bar) bar.style.width = `${progress}%`;

        const etaElement = document.getElementById(`eta-${id}`);
        if (etaElement) {
            etaElement.textContent = eta ? `ETA: ${eta}` : '';
            if (progress === 100) {
                etaElement.textContent = 'Concluído!';
            }
        }
    });

    window.electronAPI.onDbLoadProgress(({ current, total, fileName, cnpjsProcessed }) => {
        const percent = Math.round((current / total) * 100);
        dbLoadProgressBarFill.style.width = `${percent}%`;
        dbLoadProgressPercent.textContent = `${percent}%`;
        dbLoadProgressText.textContent = `Processando: ${fileName}`;
        dbLoadProgressStats.textContent = `${cnpjsProcessed} CNPJs processados`;
    });

    window.electronAPI.onDbLoadFinished(() => {
        if (startLoadToDbBtn) startLoadToDbBtn.disabled = false;
        updateEnrichedCnpjCount();

        setTimeout(() => {
            dbLoadProgressContainer.style.display = 'none';
        }, 3000);

        dbLoadProgressTitle.textContent = 'Carga Concluída!';
        dbLoadProgressBarFill.style.width = '100%';
        dbLoadProgressPercent.textContent = '100%';
        dbLoadProgressText.textContent = 'Finalizado com sucesso';
    });

    window.electronAPI.onEnrichmentFinished(() => {
        if (startEnrichmentBtn) startEnrichmentBtn.disabled = false;
    });

    // #################################################################
    // #           LÓGICA PARA A ABA DE MONITORAMENTO (ATUALIZADA)     #
    // #################################################################

    // --- Variáveis e Constantes Globais para Monitoramento ---
    let lastSuspiciousCalls = [];
    const SUSPICIOUS_TABULATIONS = ['MUDO/ENCERRAR [43]', 'MUDO [33]'];
    const SUSPICIOUS_DURATION_SECONDS = 180; // 3 minutos

    function getDurationInSeconds(timeString) {
        if (!timeString || typeof timeString !== 'string') return 0;
        const parts = timeString.split(':');
        if (parts.length === 3) {
            return parseInt(parts[0], 10) * 3600 + parseInt(parts[1], 10) * 60 + parseInt(parts[2], 10);
        }
        return 0;
    }

    // --- Lógica para o Modal de Seleção ---
    const selectionModal = document.getElementById('selection-modal');
    const modalTitle = document.getElementById('modal-title');
    const modalCloseBtn = selectionModal.querySelector('.modal-close-btn');
    const modalSearchInput = document.getElementById('modal-search-input');
    const modalListContainer = document.getElementById('modal-list-container');
    let currentModalContext = null;

    const operadores = [
        { id: '143', name: 'Adriene Rodrigues' }, { id: '58', name: 'Ana Carolina' }, { id: '46', name: 'Ana Clara Lopes' },
        { id: '105', name: 'Ana Maia' }, { id: '117', name: 'Ana Rovere' }, { id: '40', name: 'Anna Barbosa' },
        { id: '156', name: 'Arthur Medeiros' }, { id: '112', name: 'Beatriz Martins' }, { id: '92', name: 'Bianca Antunes' },
        { id: '179', name: 'Brenda Ewald' }, { id: '202', name: 'Bruna Lobato' }, { id: '144', name: 'Cairo Motta' },
        { id: '126', name: 'Camila Nogueira' }, { id: '205', name: 'Daiany Porto' }, { id: '115', name: 'Daniel Neves' },
        { id: '174', name: 'Diana Viana' }, { id: '138', name: 'Douglas Reis' }, { id: '77', name: 'Erik Freitas' },
        { id: '64', name: 'Felipe Martins' }, { id: '150', name: 'Fernanda Novaes' }, { id: '94', name: 'Gabriela Pitzer' },
        { id: '189', name: 'Giselle Mota' }, { id: '108', name: 'Giselly Salles' }, { id: '192', name: 'Graça Vitória' },
        { id: '139', name: 'Guilherme Maudonet' }, { id: '104', name: 'Heloisa Bispo' }, { id: '146', name: 'Ian Branco' },
        { id: '55', name: 'Jennyfer Vieira' }, { id: '111', name: 'Jessica Oliveira' }, { id: '109', name: 'João Honorato' },
        { id: '195', name: 'Joao Soares' }, { id: '78', name: 'Joyce Menezes' }, { id: '132', name: 'Juliana Oliveira' },
        { id: '122', name: 'Karolina Silva' }, { id: '147', name: 'Kauã Oliveira' }, { id: '194', name: 'Kawan Gabriel' },
        { id: '57', name: 'Larissa Moroni' }, { id: '71', name: 'Larissa Oliveira' }, { id: '135', name: 'Lohana Soares' },
        { id: '68', name: 'Luana Alves' }, { id: '136', name: 'Luana Ribeiro' }, { id: '178', name: 'Manuela Giraldes' },
        { id: '126', name: 'Marcos Vinicius' }, { id: '133', name: 'Maria Cristina' }, { id: '182', name: 'Maria Luna' },
        { id: '175', name: 'Maria Martins' }, { id: '103', name: 'Mariana Oliveira' }, { id: '104', name: 'Maria Seixas' },
        { id: '171', name: 'Maria Sotelino' }, { id: '127', name: 'Matheus Ribeiro' }, { id: '180', name: 'Mauricio Freitas' },
        { id: '173', name: 'Mirella Lira' }, { id: '116', name: 'Nicolle Santos' }, { id: '129', name: 'Paula Santos' },
        { id: '61', name: 'Ramon Gonçalves' }, { id: '34', name: 'Raphael Machado' }, { id: '37', name: 'Renata Souza' },
        { id: '193', name: 'Ricardo França' }, { id: '30', name: 'Rodrigo Santana' }, { id: '128', name: 'Samara Gomes' },
        { id: '96', name: 'Samella Figueira' }, { id: '98', name: 'Sarah Leite' }, { id: '157', name: 'Thais Maciel' },
        { id: '74', name: 'Thays Florencio' }, { id: '93', name: 'Vanessa Barros' }, { id: '42', name: 'Vanessa dos Santos' },
        { id: '177', name: 'Victor Alves' }, { id: '204', name: 'Vitor Faria' }, { id: '149', name: 'Vivian Ferreira' },
        { id: '190', name: 'Vivian Simplicio' }, { id: '134', name: 'Wanessa Fernandes' }
    ];
    const servicos = [
        { id: '159', name: '[C6 BANK] - EQUIPE BRUNA', category: 'Abertura' }, { id: '235', name: '[C6 BANK] - EQUIPE CAMILA', category: 'Abertura' },
        { id: '160', name: '[C6 BANK] - EQUIPE LAIANE', category: 'Abertura' }, { id: '233', name: '[C6 BANK] - EQUIPE TEF', category: 'Abertura' },
        { id: '161', name: '[C6 BANK] - EQUIPE WALESKA', category: 'Abertura' }, { id: '194', name: '[C6 BANK] -pt2 NOVO TRANSBORDO', category: 'Abertura' },
        { id: '124', name: '[C6 BANK] - CADÊNCIA GERAL', category: 'Abertura' }, { id: '181', name: '[C6/MB] C6 Pay Relacionamento ANTONIO', category: 'Relacionamento' },
        { id: '232', name: '[C6/MB] C6 Pay Relacionamento JOAO AVILA', category: 'Relacionamento' }, { id: '180', name: '[C6/MB] C6 Pay Relacionamento RAPHAELA CALDERON', category: 'Relacionamento' },
        { id: '168', name: 'Relacionamento Ana Clara', category: 'Relacionamento' }, { id: '227', name: 'Relacionamento Anna Barbosa', category: 'Relacionamento' },
        { id: '154', name: 'Relacionamento Antonio Costa', category: 'Relacionamento' }, { id: '169', name: 'Relacionamento Cairo Motta', category: 'Relacionamento' },
        { id: '203', name: 'Relacionamento Diana Viana', category: 'Relacionamento' }, { id: '186', name: 'Relacionamento digite1', category: 'Relacionamento' },
        { id: '202', name: 'Relacionamento Douglas Reis', category: 'Relacionamento' }, { id: '176', name: 'Relacionamento Fernanda Novaes', category: 'Relacionamento' },
        { id: '171', name: 'Relacionamento Guilherme Maudonet', category: 'Relacionamento' }, { id: '155', name: 'Relacionamento Higor Campos', category: 'Relacionamento' },
        { id: '229', name: 'Relacionamento Jennyfer Vieira', category: 'Relacionamento' }, { id: '172', name: 'Relacionamento Jessica Oliveira', category: 'Relacionamento' },
        { id: '148', name: 'Relacionamento João Avila', category: 'Relacionamento' }, { id: '201', name: 'Relacionamento João Honorato', category: 'Relacionamento' },
        { id: '150', name: 'Relacionamento Juliana Oliveira', category: 'Relacionamento' }, { id: '146', name: 'Relacionamento Karolina', category: 'Relacionamento' },
        { id: '228', name: 'Relacionamento Larissa Oliveira', category: 'Relacionamento' }, { id: '158', name: 'Relacionamento Luana Ribeiro', category: 'Relacionamento' },
        { id: '174', name: 'Relacionamento Marcos Vinicius', category: 'Relacionamento' }, { id: '188', name: 'Relacionamento Maria Cristina', category: 'Relacionamento' },
        { id: '225', name: 'Relacionamento Maria Seixas', category: 'Relacionamento' }, { id: '151', name: 'Relacionamento Matheus Ribeiro', category: 'Relacionamento' },
        { id: '153', name: 'Relacionamento Paula Santos', category: 'Relacionamento' }, { id: '193', name: 'Relacionamento Raphaela Calderon', category: 'Relacionamento' },
        { id: '177', name: 'Relacionamento Raphael Machado', category: 'Relacionamento' }, { id: '173', name: 'Relacionamento Renata Souza', category: 'Relacionamento' },
        { id: '170', name: 'Relacionamento Ricardo França', category: 'Relacionamento' }, { id: '226', name: 'Relacionamento Roberto Bianna', category: 'Relacionamento' },
        { id: '231', name: 'Relacionamento Rodrigo Santana', category: 'Relacionamento' }
    ];
    const gruposOperador = [
        { id: '85', name: 'Equipe Bruna' }, { id: '120', name: 'Equipe Camila' }, { id: '123', name: 'Equipe Laiane' },
        { id: '133', name: 'Equipe TEF' }, { id: '87', name: 'Equipe Waleska' }
    ];

    function renderModalList(items, searchTerm) {
        const lowerSearchTerm = searchTerm.toLowerCase();
        const filteredItems = items.filter(item => item.name.toLowerCase().includes(lowerSearchTerm));
        let html = '<ul class="custom-select-list">';
        if (currentModalContext.type === 'servico') {
            const grouped = filteredItems.reduce((acc, servico) => {
                (acc[servico.category] = acc[servico.category] || []).push(servico);
                return acc;
            }, {});
            const categoryOrder = ['Abertura', 'Relacionamento'];
            for (const category of categoryOrder) {
                if (grouped[category] && grouped[category].length > 0) {
                    html += `<li class="group-header">${category}</li>`;
                    html += grouped[category].map(s => `<li data-id="${s.id}" data-name="${s.name}">${s.name}</li>`).join('');
                }
            }
        } else {
            html += filteredItems.map(item => `<li data-id="${item.id}" data-name="${item.name}">${item.name}</li>`).join('');
        }
        html += '</ul>';
        modalListContainer.innerHTML = html;
    }

    function openSelectionModal(context) {
        currentModalContext = context;
        modalTitle.textContent = context.title;
        modalSearchInput.value = '';
        renderModalList(context.data, '');
        selectionModal.classList.remove('hidden');
        modalSearchInput.focus();
    }

    function closeSelectionModal() {
        selectionModal.classList.add('hidden');
        currentModalContext = null;
    }

    if (selectionModal) {
        modalCloseBtn.addEventListener('click', closeSelectionModal);
        selectionModal.addEventListener('click', (e) => { if (e.target === selectionModal) { closeSelectionModal(); } });
        modalSearchInput.addEventListener('input', () => { if (currentModalContext) { renderModalList(currentModalContext.data, modalSearchInput.value); } });
        modalListContainer.addEventListener('click', (e) => {
            const target = e.target;
            if (target && target.tagName === 'LI' && !target.classList.contains('group-header')) {
                const { id, name } = target.dataset;
                currentModalContext.searchEl.value = name;
                currentModalContext.hiddenInputEl.value = id;
                closeSelectionModal();
            }
        });
    }

    // --- Lógica para o Modal de Chamadas Suspeitas ---
    const suspiciousCallsModal = document.getElementById('suspicious-calls-modal');
    const suspiciousCallsList = document.getElementById('suspicious-calls-list');
    const suspiciousCallsCloseBtn = suspiciousCallsModal.querySelector('.modal-close-btn');

    function showSuspiciousCallsModal() {
        let tableHTML = `
            <table class="modal-table">
                <thead>
                    <tr>
                        <th>Operador</th>
                        <th>CNPJ Cliente</th>
                        <th>Duração</th>
                        <th>Tabulação</th>
                    </tr>
                </thead>
                <tbody>
        `;
        if (lastSuspiciousCalls.length > 0) {
            lastSuspiciousCalls.forEach(call => {
                tableHTML += `
                    <tr>
                        <td>${call.nome_operador || 'N/A'}</td>
                        <td>${call.cpf || 'N/A'}</td>
                        <td>${call.tempo_ligacao || '00:00:00'}</td>
                        <td>${call.tabulacao || 'N/A'}</td>
                    </tr>
                `;
            });
        } else {
            tableHTML += '<tr><td colspan="4" style="text-align:center;">Nenhuma chamada suspeita encontrada.</td></tr>';
        }
        tableHTML += '</tbody></table>';
        suspiciousCallsList.innerHTML = tableHTML;
        suspiciousCallsModal.classList.remove('hidden');
    }

    function closeSuspiciousCallsModal() {
        suspiciousCallsModal.classList.add('hidden');
    }

    if (suspiciousCallsCloseBtn) suspiciousCallsCloseBtn.addEventListener('click', closeSuspiciousCallsModal);
    if (suspiciousCallsModal) suspiciousCallsModal.addEventListener('click', (e) => { if (e.target === suspiciousCallsModal) { closeSuspiciousCallsModal(); } });

    // --- Lógica Principal da Aba de Monitoramento ---
    const apiParametersContainer = document.getElementById('api-parameters');
    const generateReportBtn = document.getElementById('generateReportBtn');
    const monitoringLog = document.getElementById('monitoring-log');
    const dashboardSummary = document.getElementById('dashboard-summary');
    const dashboardDetails = document.getElementById('dashboard-details');
    const dataInicioInput = document.getElementById('data_inicio_monitor');
    const dataFimInput = document.getElementById('data_fim_monitor');
    const monitoringSearchInput = document.getElementById('monitoringSearchInput');
    const dateFilterMenu = document.getElementById('date-filter-menu');

    const apiParams = [
        { name: 'id', label: 'Call ID' }, { name: 'nome', label: 'Nome Cliente' }, { name: 'chave', label: 'Chave' },
        { name: 'cpf', label: 'CPF' }, { name: 'operador_id', label: 'Operador' }, { name: 'fone_origem', label: 'Fone Origem' },
        { name: 'fone_destino', label: 'Fone Destino' }, { name: 'sentido', label: 'Sentido' }, { name: 'tronco_id', label: 'ID Tronco' },
        { name: 'digitado', label: 'Digitado' }, { name: 'resultado', label: 'Resultado' }, { name: 'tabulacao_id', label: 'ID Tabulação' },
        { name: 'operacao_id', label: 'ID Operação' }, { name: 'tipoServico', label: 'Tipo Serviço' }, { name: 'servico_id', label: 'Desempenho de campanha' },
        { name: 'grupo_operador_id', label: 'Desempenho de equipes' },
    ];

    if (apiParametersContainer) {
        apiParametersContainer.innerHTML = apiParams.map(param => {
            const isSelectable = ['operador_id', 'servico_id', 'grupo_operador_id'].includes(param.name);
            const inputHtml = isSelectable ?
                `<div id="${param.name}-select-container" class="custom-select-container hidden">
                    <input type="text" id="${param.name}-search" readonly placeholder="Clique para selecionar..." style="cursor: pointer;">
                    <input type="hidden" id="input-${param.name}">
                </div>` :
                `<input type="text" id="input-${param.name}" class="hidden" placeholder="Valor...">`;

            return `
                <div class="param-item">
                    <div class="toggle-switch">
                        <label class="switch">
                            <input type="checkbox" id="check-${param.name}" data-param-name="${param.name}">
                            <span class="slider"></span>
                        </label>
                        <span class="toggle-label">${param.label}</span>
                    </div>
                    ${inputHtml}
                </div>
            `;
        }).join('');

        apiParams.forEach(param => {
            const checkbox = document.getElementById(`check-${param.name}`);
            if (!checkbox) return;
            const isSelectable = ['operador_id', 'servico_id', 'grupo_operador_id'].includes(param.name);
            const inputElement = isSelectable ? document.getElementById(`${param.name}-select-container`) : document.getElementById(`input-${param.name}`);

            checkbox.addEventListener('change', () => {
                inputElement.classList.toggle('hidden', !checkbox.checked);
                if (!checkbox.checked) {
                    if (isSelectable) {
                        inputElement.querySelector('input[type="text"]').value = '';
                        inputElement.querySelector('input[type="hidden"]').value = '';
                    } else {
                        inputElement.value = '';
                    }
                }
            });

            if (isSelectable) {
                const searchEl = document.getElementById(`${param.name}-search`);
                searchEl.addEventListener('click', () => {
                    if (!checkbox.checked) return;
                    let data, type;
                    if (param.name === 'operador_id') { data = operadores; type = 'operador'; } 
                    else if (param.name === 'servico_id') { data = servicos; type = 'servico'; } 
                    else { data = gruposOperador; type = 'grupo_operador'; }
                    openSelectionModal({
                        type: type, title: `Selecionar ${param.label}`, data: data,
                        searchEl: searchEl, hiddenInputEl: document.getElementById(`input-${param.name}`)
                    });
                });
            }
        });
    }

    if (monitoringSearchInput) {
        const foneDestinoCheckbox = document.getElementById('check-fone_destino');
        const foneDestinoInput = document.getElementById('input-fone_destino');

        monitoringSearchInput.addEventListener('input', () => {
            const searchTerm = monitoringSearchInput.value.trim();
            foneDestinoInput.value = searchTerm;
            const hasSearchTerm = searchTerm !== '';
            foneDestinoCheckbox.checked = hasSearchTerm;
            foneDestinoInput.classList.toggle('hidden', !hasSearchTerm);
        });
    }

    const getHtmlDate = (date) => {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    const getApiDate = (dateString) => {
        if (!dateString) return '';
        const [year, month, day] = dateString.split('-');
        return `${day}/${month}/${year}`;
    }

    if (dateFilterMenu) {
        dateFilterMenu.addEventListener('click', (e) => {
            if (e.target.tagName === 'LI') {
                const period = e.target.dataset.period;
                const today = new Date();
                let startDate, endDate = new Date();
                switch (period) {
                    case 'today': startDate = today; endDate = today; break;
                    case 'yesterday': startDate = new Date(today); startDate.setDate(today.getDate() - 1); endDate = startDate; break;
                    case 'this_week': startDate = new Date(today); const dayOfWeek = today.getDay(); startDate.setDate(today.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1)); endDate = today; break;
                    case 'last_week': startDate = new Date(today); startDate.setDate(today.getDate() - today.getDay() - 6); endDate = new Date(startDate); endDate.setDate(startDate.getDate() + 6); break;
                    case 'this_month': startDate = new Date(today.getFullYear(), today.getMonth(), 1); endDate = today; break;
                }
                if (dataInicioInput) dataInicioInput.value = getHtmlDate(startDate);
                if (dataFimInput) dataFimInput.value = getHtmlDate(endDate);
                e.target.closest('details').removeAttribute('open');
            }
        });
    }


    if (generateReportBtn) generateReportBtn.addEventListener('click', async () => {
        generateReportBtn.disabled = true;
        monitoringLog.innerHTML = '> 🌀 Gerando relatório... Por favor, aguarde.';
        dashboardSummary.innerHTML = '';
        dashboardDetails.innerHTML = '';
        let baseUrl = 'https://mbfinance.fastssl.com.br/api/relatorio/captura_valores_analitico.php?';
        let params = [];
        apiParams.forEach(param => {
            const checkbox = document.getElementById(`check-${param.name}`);
            if (checkbox && checkbox.checked) {
                const input = document.getElementById(`input-${param.name}`);
                params.push(`${param.name}=${encodeURIComponent(input.value)}`);
            } else {
                params.push(`${param.name}=`);
            }
        });
        params.push(`data_inicio=${getApiDate(dataInicioInput.value)}`);
        params.push(`data_fim=${getApiDate(dataFimInput.value)}`);
        params.push('formato=json');
        const finalUrl = baseUrl + params.join('&');
        const result = await window.electronAPI.fetchMonitoringReport(finalUrl);
        if (result.success && result.data) {
            updateDashboard(result.data);
            monitoringLog.innerHTML = `> ✅ Relatório gerado com sucesso. ${result.data.length} registros encontrados.`;
        } else {
            monitoringLog.innerHTML = `> ❌ ERRO: ${result.message || 'Falha ao buscar dados da API.'}`;
        }
        generateReportBtn.disabled = false;
    });

    function updateDashboard(data) {
        if (!data || !Array.isArray(data)) {
            dashboardSummary.innerHTML = '<p style="color: var(--text-muted); text-align: center; grid-column: 1 / -1;">Ocorreu um erro ou nenhum dado foi retornado.</p>';
            dashboardDetails.innerHTML = '';
            return;
        }
         if (data.length === 0) {
            dashboardSummary.innerHTML = '<p style="color: var(--text-muted); text-align: center; grid-column: 1 / -1;">Nenhum dado retornado para os filtros selecionados.</p>';
            dashboardDetails.innerHTML = '';
            return;
        }
    
        const totalCalls = data.length;
        const aggregators = {
            tabulacao: {}, resultado: {}, nome_operador: {}, nome_campanha: {},
        };
        const detailedTabulations = {};
        let totalDurationSeconds = 0;
        let suspiciousCalls = [];
    
        data.forEach(item => {
            const tabulacao = item.tabulacao || 'Não Preenchido';
            const duration = getDurationInSeconds(item.tempo_ligacao);
            for (const key in aggregators) {
                const value = item[key] || 'Não Preenchido';
                aggregators[key][value] = (aggregators[key][value] || 0) + 1;
            }
            if (!detailedTabulations[tabulacao]) { detailedTabulations[tabulacao] = []; }
            detailedTabulations[tabulacao].push({
                cpf: item.cpf,
                duration: item.tempo_ligacao || '00:00:00'
            });
            if (SUSPICIOUS_TABULATIONS.includes(tabulacao) && duration >= SUSPICIOUS_DURATION_SECONDS) {
                suspiciousCalls.push(item);
            }
            totalDurationSeconds += duration;
        });
        lastSuspiciousCalls = suspiciousCalls; 
    
        const avgDurationSeconds = totalCalls > 0 ? totalDurationSeconds / totalCalls : 0;
        const roundedAvgSeconds = Math.round(avgDurationSeconds);
        const avgMinutes = Math.floor(roundedAvgSeconds / 60);
        const avgSeconds = roundedAvgSeconds % 60;
        const tma = `${String(avgMinutes).padStart(2, '0')}:${String(avgSeconds).padStart(2, '0')}`;
    
        dashboardSummary.innerHTML = `
            <div class="summary-card">
                <div class="summary-card-title">Total de Chamadas</div>
                <div class="summary-card-value">${totalCalls.toLocaleString('pt-BR')}</div>
            </div>
            <div class="summary-card">
                <div class="summary-card-title">TMA</div>
                <div class="summary-card-value">${tma}</div>
            </div>
            <button class="summary-card summary-card-button" id="suspicious-card-btn" ${lastSuspiciousCalls.length === 0 ? 'disabled' : ''}>
                <div class="summary-card-title">Tabulações Suspeitas</div>
                <div class="summary-card-value warning">${lastSuspiciousCalls.length}</div>
            </button>
            <div class="summary-card">
                <div class="summary-card-title">Operadores Envolvidos</div>
                <div class="summary-card-value">${Object.keys(aggregators.nome_operador).length}</div>
            </div>
        `;
        document.getElementById('suspicious-card-btn').addEventListener('click', showSuspiciousCallsModal);
    
        dashboardDetails.innerHTML = '';
        const isOperatorFiltered = document.getElementById('check-operador_id')?.checked;
    
        const createDetailCard = (title, dataObject) => {
            if (title === 'Top Tabulações' && isOperatorFiltered) {
                return createInteractiveTabulationCard('Top Tabulações', detailedTabulations);
            }
            const sortedData = Object.entries(dataObject).sort(([, a], [, b]) => b - a);
            let listItems = sortedData.map(([name, count]) => `
                <li><span class="name" title="${name}">${name}</span><span class="count">${count.toLocaleString('pt-BR')}</span></li>
            `).join('');
            if (!listItems) listItems = '<li>Nenhum dado.</li>';
            return `<div class="detail-card"><h3>${title}</h3><ul class="detail-list custom-scrollbar">${listItems}</ul></div>`;
        }
    
        const createInteractiveTabulationCard = (title, detailedDataObject) => {
            const sortedTabulations = Object.keys(detailedDataObject).sort((a, b) => detailedDataObject[b].length - detailedDataObject[a].length);
            let detailsHtml = sortedTabulations.map(tabulationName => {
                const calls = detailedDataObject[tabulationName];
                const isSuspicious = SUSPICIOUS_TABULATIONS.includes(tabulationName);
                const callListHtml = calls.map(call => `
                    <li>
                        <span class="call-cnpj">CNPJ: ${call.cpf || 'N/A'}</span>
                        <span class="call-duration">Duração: ${call.duration}</span>
                    </li>
                `).join('');
                return `
                    <details>
                        <summary class="${isSuspicious ? 'suspicious-summary' : ''}">
                            <span>${tabulationName}</span>
                            <span>${calls.length} chamadas</span>
                        </summary>
                        <ul class="tabulation-call-list custom-scrollbar">${callListHtml}</ul>
                    </details>
                `;
            }).join('');
            return `
                <div class="detail-card interactive-tabulation">
                    <h3>${title} (Detalhado)</h3>
                    <div class="custom-scrollbar" style="max-height: 400px; overflow-y: auto; padding-right: 5px;">${detailsHtml}</div>
                </div>
            `;
        };
        
        dashboardDetails.innerHTML += createDetailCard('Top Tabulações', aggregators.tabulacao);
        dashboardDetails.innerHTML += createDetailCard('Resultados por Chamada', aggregators.resultado);
        dashboardDetails.innerHTML += createDetailCard('Top Operadores', aggregators.nome_operador);
        dashboardDetails.innerHTML += createDetailCard('Top Campanhas', aggregators.nome_campanha);
    }
});