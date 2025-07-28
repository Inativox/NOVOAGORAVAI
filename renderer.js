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
        'enrichmentGrid': 'enrichment-sections'
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
            onEnd: function(evt) {
                const sections = Array.from(grid.querySelectorAll('.section'));
                saveSectionState(storageKey, sections);
            }
        });

        sortableInstances[gridId] = sortable;

        grid.addEventListener('click', function(e) {
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
        'Automação': {
            title: 'Fluxo de Automação "Executar Tudo"',
            description: 'Combine todas as etapas de tratamento de bases em um único processo. Limpe, consulte e enriqueça seus dados com um único clique.'
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
        body.classList.remove('c6-theme', 'enrichment-theme', 'automation-theme');
        if (tabNameId === 'api') {
            body.classList.add('c6-theme');
        } else if (tabNameId === 'enriquecimento') {
            body.classList.add('enrichment-theme');
        } else if (tabNameId === 'automacao') {
            body.classList.add('automation-theme');
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
            } else if (buttonText.includes('Automação')) {
                tabNameId = 'automacao';
            }

            if (tabNameId) {
                openTab(event, tabNameId);
            }
        });
    });

    if (tabButtons.length > 0) {
        tabButtons[0].click();
    }

    let rootFile             = null;
    let cleanFiles           = [];
    let mergeFiles           = [];
    let backupEnabled        = false;
    let autoAdjustPhones     = false;
    let checkDbEnabled       = false;
    let saveToDbEnabled      = false;

    const selectRootBtn           = document.getElementById('selectRootBtn');
    const autoRootBtn             = document.getElementById('autoRootBtn');
    const feedRootBtn             = document.getElementById('feedRootBtn');
    const updateBlocklistBtn      = document.getElementById('updateBlocklistBtn');
    const addCleanFileBtn         = document.getElementById('addCleanFileBtn');
    const startCleaningBtn        = document.getElementById('startCleaningBtn');
    const resetLocalBtn           = document.getElementById('resetLocalBtn');
    const adjustPhonesBtn         = document.getElementById('adjustPhonesBtn');
    const backupCheckbox          = document.getElementById('backupCheckbox');
    const autoAdjustPhonesCheckbox= document.getElementById('autoAdjustPhonesCheckbox');
    const rootFilePathSpan        = document.getElementById('rootFilePath');
    const selectedCleanFilesDiv   = document.getElementById('selectedCleanFiles');
    const progressContainer       = document.getElementById('progressContainer');
    const logDiv                  = document.getElementById('log');
    const rootColSelect           = document.getElementById('rootCol');
    const destColSelect           = document.getElementById('destCol');
    const selectMergeFilesBtn     = document.getElementById('selectMergeFilesBtn');
    const startMergeBtn           = document.getElementById('startMergeBtn');
    const selectedMergeFilesDiv   = document.getElementById('selectedMergeFiles');
    const saveStoredCnpjsBtn      = document.getElementById('saveStoredCnpjsBtn');
    const checkDbCheckbox         = document.getElementById('checkDbCheckbox');
    const saveToDbCheckbox        = document.getElementById('saveToDbCheckbox');
    const consultDbBtn            = document.getElementById('consultDbBtn');
    const uploadProgressContainer = document.getElementById('uploadProgressContainer');
    const uploadProgressTitle     = document.getElementById('uploadProgressTitle');
    const uploadProgressBarFill   = document.getElementById('uploadProgressBarFill');
    const uploadProgressText      = document.getElementById('uploadProgressText');
    const batchIdInput            = document.getElementById('batchIdInput');
    const deleteBatchBtn          = document.getElementById('deleteBatchBtn');
    
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
        uploadProgressContainer.style.display = 'none';
        uploadProgressBarFill.style.width = '0%';
        uploadProgressText.textContent = '';
    }

    backupCheckbox.addEventListener('change', () => { backupEnabled = backupCheckbox.checked; });
    autoAdjustPhonesCheckbox.addEventListener('change', () => { autoAdjustPhones = autoAdjustPhonesCheckbox.checked; });

    checkDbCheckbox.addEventListener('change', () => {
        checkDbEnabled = checkDbCheckbox.checked;
        appendLog(`Consulta ao Banco de Dados: ${checkDbEnabled ? 'ATIVADA' : 'DESATIVADA'}`);
    });

    saveToDbCheckbox.addEventListener('change', () => {
        saveToDbEnabled = saveToDbCheckbox.checked;
        appendLog(`Salvar no Banco de Dados: ${saveToDbEnabled ? 'ATIVADO' : 'DESATIVADO'}`);
    });

    saveStoredCnpjsBtn.addEventListener('click', async () => {
        appendLog('Solicitando salvamento do histórico de CNPJs em Excel...');
        const result = await window.electronAPI.saveStoredCnpjsToExcel();
        appendLog(result.message);
    });

    deleteBatchBtn.addEventListener('click', async () => {
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
            if(result.success) {
                batchIdInput.value = '';
            }
        } else {
            appendLog('Operação de exclusão cancelada pelo usuário.');
        }
    });

    consultDbBtn.addEventListener('click', async () => {
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

    selectRootBtn.addEventListener('click', async () => {
      const files = await window.electronAPI.selectFile({ title: 'Selecione a Lista Raiz', multi: false });
      if (files && files.length > 0) {
        rootFile = files[0];
        addFileToUI(rootFilePathSpan, rootFile, true);
        appendLog(`Arquivo raiz selecionado: ${rootFile}`);
      }
    });

    autoRootBtn.addEventListener('click', () => {
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

    updateBlocklistBtn.addEventListener('click', async () => {
      const result = await window.electronAPI.updateBlocklist(backupEnabled);
      appendLog(result.success ? result.message : `Erro: ${result.message}`);
    });

    addCleanFileBtn.addEventListener('click', async () => {
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

    startCleaningBtn.addEventListener('click', () => {
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

    resetLocalBtn.addEventListener('click', () => {
      rootFile = null;
      cleanFiles = [];
      mergeFiles = [];
      backupEnabled = false;
      autoAdjustPhones = false;
      checkDbEnabled = false;
      saveToDbEnabled = false;
      rootFilePathSpan.innerHTML = '';
      selectedCleanFilesDiv.innerHTML = '';
      progressContainer.innerHTML = '';
      logDiv.textContent = '';
      selectedMergeFilesDiv.innerHTML = '';
      batchIdInput.value = '';
      backupCheckbox.checked = false;
      autoAdjustPhonesCheckbox.checked = false;
      checkDbCheckbox.checked = false;
      saveToDbCheckbox.checked = false;
      delete autoRootBtn.dataset.on;
      autoRootBtn.textContent = 'Auto Raiz: OFF';
      selectRootBtn.disabled = false;
      resetUploadProgress();
      appendLog('Módulo de Limpeza Local reiniciado.');
    });

    adjustPhonesBtn.addEventListener('click', async () => {
      const files = await window.electronAPI.selectFile({ title: 'Selecione arquivo para ajustar fones', multi: false });
      if (!files?.length) return appendLog('Nenhum arquivo selecionado.');
      window.electronAPI.startAdjustPhones({ filePath: files[0], backup: backupEnabled });
    });

    selectMergeFilesBtn.addEventListener('click', async () => {
      const files = await window.electronAPI.selectFile({ title: 'Selecione arquivos para mesclar', multi: true });
      if (!files?.length) return;
      mergeFiles = files;
      selectedMergeFilesDiv.innerHTML = '';
      files.forEach(f => {
        addFileToUI(selectedMergeFilesDiv, f, false);
      });
    });

    startMergeBtn.addEventListener('click', () => {
      if (!mergeFiles.length) return appendLog('ERRO: Selecione arquivos para mesclar.');
      appendLog('Iniciando mesclagem...');
      window.electronAPI.startMerge(mergeFiles);
    });

    feedRootBtn.addEventListener('click', async () => {
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
        feedRootBtn.disabled = false;
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
    
    apiDropzone.addEventListener('dragover', (event) => { event.preventDefault(); event.stopPropagation(); apiDropzone.style.borderColor = 'var(--accent-color)'; apiDropzone.style.backgroundColor = 'var(--bg-lighter)'; });
    apiDropzone.addEventListener('dragleave', (event) => { event.preventDefault(); event.stopPropagation(); apiDropzone.style.borderColor = 'var(--border-color)'; apiDropzone.style.backgroundColor = 'transparent'; });
    apiDropzone.addEventListener('drop', (event) => { event.preventDefault(); event.stopPropagation(); apiDropzone.style.borderColor = 'var(--border-color)'; apiDropzone.style.backgroundColor = 'transparent'; const files = Array.from(event.dataTransfer.files).filter(file => file.path.endsWith('.xlsx') || file.path.endsWith('.xls') || file.path.endsWith('.csv')).map(file => file.path); if (files.length > 0) { window.electronAPI.addFilesToApiQueue(files); } });
    
    selectApiFileBtn.addEventListener('click', async () => { const files = await window.electronAPI.selectFile({ title: 'Selecione as planilhas de CNPJs', multi: true }); if (files && files.length > 0) { window.electronAPI.addFilesToApiQueue(files); } });
    startApiBtn.addEventListener('click', () => { startApiBtn.disabled = true; resetApiBtn.disabled = true; apiStatusSpan.textContent = 'Iniciando processamento da fila...'; window.electronAPI.startApiQueue({ keyMode: apiKeySelection.value }); });
    resetApiBtn.addEventListener('click', () => { window.electronAPI.resetApiQueue(); });
    
    function updateApiQueueUI(queue) { 
        const { pending, processing, completed } = queue; 
        
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
    
    function appendApiLog(msg) { apiLogDiv.innerHTML += `> ${msg.replace(/\n/g, '<br>> ')}\n`; apiLogDiv.scrollTop = apiLogDiv.scrollHeight; }

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
        if (enrichmentLogDiv.textContent === 'Aguardando início...') {
            enrichmentLogDiv.innerHTML = '';
        }
        enrichmentLogDiv.innerHTML += `> ${msg.replace(/\n/g, '<br>> ')}\n`;
        enrichmentLogDiv.scrollTop = enrichmentLogDiv.scrollHeight;
    }

    async function updateEnrichedCnpjCount() {
        try {
            enrichedCnpjCountSpan.textContent = 'Carregando...';
            const count = await window.electronAPI.getEnrichedCnpjCount();
            enrichedCnpjCountSpan.textContent = count.toLocaleString('pt-BR');
        } catch (error) {
            enrichedCnpjCountSpan.textContent = 'Erro';
            appendEnrichmentLog(`❌ Erro ao carregar contador: ${error.message}`);
        }
    }

    updateEnrichedCnpjCount();

    refreshCountBtn.addEventListener('click', updateEnrichedCnpjCount);

    downloadEnrichedDataBtn.addEventListener('click', async () => {
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

    selectMasterFilesBtn.addEventListener('click', async () => {
        const files = await window.electronAPI.selectFile({ title: 'Selecione as Planilhas Mestras', multi: true });
        if (!files?.length) return;
        enrichmentMasterFiles = files;
        selectedMasterFilesDiv.innerHTML = '';
        files.forEach(file => {
            addFileToUI(selectedMasterFilesDiv, file, false);
        });
    });

    startLoadToDbBtn.addEventListener('click', () => {
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

    selectEnrichFilesBtn.addEventListener('click', async () => {
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

    startEnrichmentBtn.addEventListener('click', () => {
        if (enrichmentEnrichFiles.length === 0) return appendEnrichmentLog('❌ ERRO: Selecione pelo menos um arquivo para enriquecer.');
        startEnrichmentBtn.disabled = true;
        const strategy = document.querySelector('input[name="enrichStrategy"]:checked').value;
        appendEnrichmentLog(`Iniciando enriquecimento com a estratégia: ${strategy.toUpperCase()}`);
        window.electronAPI.startEnrichment({ filesToEnrich: enrichmentEnrichFiles, strategy, backup: backupCheckbox.checked });
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
        startLoadToDbBtn.disabled = false;
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
        startEnrichmentBtn.disabled = false;
    });

    // #################################################################
    // #           NOVA LÓGICA PARA A ABA DE AUTOMAÇÃO COMPLETA        #
    // #################################################################
    
    gridConfigs['automationGrid'] = 'automation-sections';
    initializeSortable('automationGrid', 'automation-sections');

    const selectAutomationFileBtn = document.getElementById('selectAutomationFileBtn');
    const selectedAutomationFileDiv = document.getElementById('selectedAutomationFile');
    const startAutomationBtn = document.getElementById('startAutomationBtn');
    const automationLog = document.getElementById('automationLog');
    const automationStepTitle = document.getElementById('automationStepTitle');
    const automationStepMessage = document.getElementById('automationStepMessage');
    const automationProgressBarFill = document.getElementById('automationProgressBarFill');

    let automationFilePath = null;

    function resetAutomationUI() {
        startAutomationBtn.disabled = true;
        startAutomationBtn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path d="M11.596 8.697l-6.363 3.692c-.54.313-1.233-.066-1.233-.697V4.308c0-.63.692-1.01 1.233-.696l6.363 3.692a.802.802 0 0 1 0 1.393z"/></svg>2. Executar Tudo`;
        automationStepTitle.textContent = 'Aguardando Início';
        automationStepMessage.textContent = '';
        automationProgressBarFill.style.width = '0%';
    }

    selectAutomationFileBtn.addEventListener('click', async () => {
        const files = await window.electronAPI.selectFile({ title: 'Selecione a planilha para o fluxo completo', multi: false });
        if (files && files.length > 0) {
            automationFilePath = files[0];
            addFileToUI(selectedAutomationFileDiv, automationFilePath, true);
            
            automationStepTitle.textContent = 'Aguardando Início';
            automationStepMessage.textContent = 'Pronto para executar.';
            automationLog.innerHTML = `> Arquivo selecionado: ${getBasename(automationFilePath)}\n> Pronto para iniciar a automação.`;

            startAutomationBtn.disabled = false;
        }
    });

    startAutomationBtn.addEventListener('click', () => {
        if (!automationFilePath) {
            return alert('Por favor, selecione um arquivo para automatizar.');
        }

        const isAutoRoot = autoRootBtn.dataset.on === 'true';
        if (!isAutoRoot && !rootFile) {
            return alert("Por favor, selecione ou configure um 'Arquivo Raiz' na aba 'Limpeza Local' antes de executar a automação.");
        }
        
        startAutomationBtn.disabled = true;
        automationLog.innerHTML = '';
        automationStepTitle.textContent = 'Iniciando Automação';
        automationStepMessage.textContent = 'Preparando para executar o fluxo...';
        
        const steps = {
            limpeza: document.getElementById('automation-step-limpeza').checked,
            api: document.getElementById('automation-step-api').checked,
            enriquecimento: document.getElementById('automation-step-enriquecimento').checked
        };

        if (!steps.limpeza && !steps.api && !steps.enriquecimento) {
            alert("Por favor, selecione pelo menos uma etapa para a automação.");
            startAutomationBtn.disabled = false;
            return;
        }

        const settings = {
            checkDb: document.getElementById('checkDbCheckbox').checked,
            saveToDb: document.getElementById('saveToDbCheckbox').checked,
            backup: document.getElementById('backupCheckbox').checked,
            autoAdjust: document.getElementById('autoAdjustPhonesCheckbox').checked,
            destCol: document.getElementById('destCol').value,
            rootCol: document.getElementById('rootCol').value,
            apiKeyMode: document.getElementById('apiKeySelection').value,
            enrichStrategy: document.querySelector('input[name="enrichStrategy"]:checked').value,
            rootFilePath: isAutoRoot ? null : rootFile
        };

        window.electronAPI.startFullAutomation({ 
            filePath: automationFilePath, 
            settings: settings, 
            steps: steps,
            isAutoRoot: isAutoRoot
        });
    });


    function appendAutomationLog(msg) {
        if (automationLog.textContent.startsWith('Aguardando início')) {
            automationLog.innerHTML = '';
        }
        automationLog.innerHTML += `> ${msg.replace(/\n/g, '<br>> ')}\n`;
        automationLog.scrollTop = automationLog.scrollHeight;
    }

    window.electronAPI.onAutomationLog((msg) => {
        appendAutomationLog(msg);
    });

    window.electronAPI.onAutomationStepUpdate(({ step, title, message, progress }) => {
        automationStepTitle.textContent = `Etapa: ${title}`;
        automationStepMessage.textContent = message;
        automationProgressBarFill.style.width = `${progress}%`;
    });

    window.electronAPI.onAutomationFinished(({ success, message, finalPath }) => {
        startAutomationBtn.disabled = false;
        appendAutomationLog(message);

        if (success) {
            automationStepTitle.textContent = 'Automação Concluída!';
            automationStepMessage.textContent = 'Processo finalizado com sucesso.';
            automationLog.innerHTML += `<br><button id="openFinalFileBtn" data-path="${finalPath}">Abrir Arquivo Final</button>`;
            document.getElementById('openFinalFileBtn').addEventListener('click', (e) => {
                window.electronAPI.openPath(e.target.dataset.path);
            });
        } else {
            automationStepTitle.textContent = 'Erro na Automação';
            automationStepMessage.textContent = 'Ocorreu um erro. Verifique os logs.';
        }
    });
});