let currentTab = null;
let extractionCount = 0;
let selectedCount = 0;
let previewSource = 'selection'; // 'selection' | 'storage'
let currentLoadedSession = null; // { pageId, groupId }
let lastPreview = [];
let currentLoadedAggregated = [];

let currentOverrides = {};  // [{ pageId, groupId, key, type }]

// Compose a unique override key from context + selector to avoid collisions across iframes
function overrideKey(contextDocument, selector) {
    const ctx = (contextDocument && String(contextDocument).trim()) || 'document';
    const sel = (selector && String(selector).trim()) || '';
    return `${ctx} >>> ${sel}`;
}

// Pending edit queues for tree view Apply Changes
let pendingGroupAdds = [];
let pendingGroupRemoves = [];
let pendingGroupNames = {};
let pendingTypeChanges = [];
let pendingMoves = [];
let pendingDeletes = [];
let pendingInserts = [];
let pendingTempGroupIdCounter = -1;

// Initialize
document.addEventListener('DOMContentLoaded', async () => {
    await loadSavedSettings();
    setupEventListeners();
    updateStats();
});

async function loadSavedSettings() {
    const result = await chrome.storage.local.get([
        'deepseek_api_key', 'extraction_count', 'ai_provider', 'ai_model', 'label_mode',
        'azure_endpoint', 'azure_deployment', 'azure_api_version', 'theme'
    ]);

    if (result.deepseek_api_key) {
        document.getElementById('api-key').value = result.deepseek_api_key;
    }

    if (result.extraction_count) {
        extractionCount = result.extraction_count;
        document.getElementById('extract-count').textContent = extractionCount;
    }

    const provider = result.ai_provider || 'deepseek';
    if (document.getElementById('ai-provider')) {
        document.getElementById('ai-provider').value = provider;
        // Show/hide Azure settings based on provider
        toggleAzureSettings(provider === 'azure');
    }
    if (result.ai_model && document.getElementById('ai-model')) {
        document.getElementById('ai-model').value = result.ai_model;
    }

    if (document.getElementById('label-mode')) {
        document.getElementById('label-mode').value = result.label_mode || 'original';
    }

    // Load Azure settings
    if (result.azure_endpoint) document.getElementById('azure-endpoint').value = result.azure_endpoint;
    if (result.azure_deployment) document.getElementById('azure-deployment').value = result.azure_deployment;
    if (result.azure_api_version) document.getElementById('azure-api-version').value = result.azure_api_version;

    // Load theme
    if (result.theme === 'dark' || (!result.theme && window.matchMedia('(prefers-color-scheme: dark)').matches)) {
        document.body.classList.add('dark');
    }
}

function setupEventListeners() {
    document.getElementById('btn-auto').addEventListener('click', async () => {
        await sendBroadcastToFrames('AUTO_SELECT');
        await sendMessageToContent('AUTO_SELECT'); // direct fallback to top frame
        updateStatus('Auto-select requested', 'info');
    });
    document.getElementById('btn-manual').addEventListener('click', toggleManualMode);
    document.getElementById('btn-extract').addEventListener('click', extractElements);
    document.getElementById('btn-download').addEventListener('click', downloadData);
    document.getElementById('btn-convert').addEventListener('click', convertAndDownload);
    document.getElementById('btn-generate').addEventListener('click', generateSampleData);
    const btnEntry = document.getElementById('btn-entry');
    if (btnEntry) btnEntry.addEventListener('click', runEntryTest);
    document.getElementById('btn-clear').addEventListener('click', clearAllAndExit);
    // Removed inline Refresh/Validate buttons; handled via Load Latest
    const btnLoadAll = document.getElementById('btn-load-all');
    if (btnLoadAll) btnLoadAll.addEventListener('click', loadAllExtractionsToPreview);
    const btnSaveSession = document.getElementById('btn-save-session');
    if (btnSaveSession) btnSaveSession.addEventListener('click', savePreviewEditsToSession);
    // Removed Editing & Shortcuts section
    const themeBtn = document.getElementById('btn-theme');
    if (themeBtn) themeBtn.addEventListener('click', toggleTheme);
    const applyColors = document.getElementById('btn-apply-colors');
    if (applyColors) applyColors.addEventListener('click', applyHighlightColors);
    // Removed minimap toggle

    // Settings toggle
    const settingsToggle = document.getElementById('settings-toggle');
    if (settingsToggle) {
        settingsToggle.addEventListener('click', toggleSettings);
        // Start collapsed by default
        collapseSettings();
    }

    document.getElementById('api-key').addEventListener('change', async (e) => {
        const apiKey = e.target.value;
        await chrome.storage.local.set({ deepseek_api_key: apiKey });
        updateStatus('API key saved', 'success');
    });

    const providerEl = document.getElementById('ai-provider');
    const modelEl = document.getElementById('ai-model');
    if (providerEl) {
        providerEl.addEventListener('change', (e) => {
            toggleAzureSettings(e.target.value === 'azure');
            saveAISettings();
        });
    }
    if (modelEl) modelEl.addEventListener('change', saveAISettings);

    // API key visibility toggle
    const toggleApiKeyBtn = document.getElementById('toggle-api-key');
    if (toggleApiKeyBtn) {
        toggleApiKeyBtn.addEventListener('click', () => {
            const apiKeyInput = document.getElementById('api-key');
            apiKeyInput.type = apiKeyInput.type === 'password' ? 'text' : 'password';
            toggleApiKeyBtn.textContent = apiKeyInput.type === 'password' ? '👁️' : '🙈';
        });
    }

    const applyLabelModeBtn = document.getElementById('btn-apply-label-mode');
    if (applyLabelModeBtn) applyLabelModeBtn.addEventListener('click', applyLabelMode);

    const applyApiBtn = document.getElementById('btn-apply-api');
    if (applyApiBtn) applyApiBtn.addEventListener('click', applyApiSettings);
}

async function saveAISettings() {
    const provider = document.getElementById('ai-provider')?.value || 'deepseek';
    const model = document.getElementById('ai-model')?.value || '';
    await chrome.storage.local.set({ ai_provider: provider, ai_model: model });
    updateStatus('AI settings saved', 'success');
}

function toggleAzureSettings(show) {
    const azureSettings = document.getElementById('azure-settings');
    if (azureSettings) {
        azureSettings.style.display = show ? 'block' : 'none';
    }
}

async function applyApiSettings() {
    const apiKey = document.getElementById('api-key')?.value || '';
    const provider = document.getElementById('ai-provider')?.value || 'deepseek';

    if (apiKey) {
        await chrome.storage.local.set({ deepseek_api_key: apiKey });
    }

    // Save Azure-specific settings if Azure is selected
    if (provider === 'azure') {
        const azureEndpoint = document.getElementById('azure-endpoint')?.value || '';
        const azureDeployment = document.getElementById('azure-deployment')?.value || '';
        const azureApiVersion = document.getElementById('azure-api-version')?.value || '2024-02-01';

        if (!azureEndpoint) {
            updateStatus('Azure endpoint is required for Azure OpenAI', 'error');
            return;
        }

        await chrome.storage.local.set({
            azure_endpoint: azureEndpoint,
            azure_deployment: azureDeployment,
            azure_api_version: azureApiVersion
        });
    }

    await saveAISettings();
    updateStatus('API settings applied successfully', 'success');
}

async function getCurrentTab() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    return tab;
}

async function sendMessageToContent(action, data = {}) {
    const tab = await getCurrentTab();
    if (!tab) {
        updateStatus('No active tab found', 'error');
        return null;
    }
    if (!isInjectableUrl(tab.url)) {
        updateStatus('This page does not allow content scripts. Open a regular website.', 'error');
        return null;
    }

    try {
        // Content scripts auto-inject via manifest, so just send message directly
        const response = await chrome.tabs.sendMessage(tab.id, { action, ...data });
        return response;
    } catch (error) {
        const msg = String(error?.message || error);
        // Suppress noisy error on restricted pages or when content scripts aren't injected yet
        if (msg.includes('Could not establish connection') || msg.includes('Receiving end does not exist')) {
            updateStatus('Open a regular website tab to use actions.', 'warning');
            return null;
        }
        console.error('Error sending message:', error);
        updateStatus('Error communicating with page. Try refreshing.', 'error');
        return null;
    }
}

async function sendBroadcastToFrames(targetAction, payload = {}) {
    const tab = await getCurrentTab();
    if (!tab) {
        updateStatus('No active tab found', 'error');
        return null;
    }
    if (!isInjectableUrl(tab.url)) {
        updateStatus('This page does not allow content scripts. Open a regular website.', 'error');
        return null;
    }
    try {
        // Content scripts auto-inject via manifest
        return await chrome.runtime.sendMessage({ action: 'BROADCAST', tabId: tab.id, targetAction, payload });
    } catch (e) {
        console.error('Broadcast error:', e);
        updateStatus('Broadcast to frames failed', 'error');
        return null;
    }
}

function isInjectableUrl(url) {
    // If URL is not available (no tabs permission/activeTab grant), do not block.
    if (!url) return true;
    return !(
        url.startsWith('chrome://') ||
        url.startsWith('edge://') ||
        url.startsWith('about:') ||
        url.startsWith('chrome-extension://') ||
        url.startsWith('https://chrome.google.com/webstore')
    );
}

async function toggleManualMode() {
    const btn = document.getElementById('btn-manual');
    const isActive = btn.classList.contains('active');

    if (isActive) {
        btn.classList.remove('active');
        btn.textContent = 'Manual Mode';
        await sendBroadcastToFrames('MANUAL_MODE_OFF');
        await sendMessageToContent('MANUAL_MODE_OFF'); // direct fallback
        updateStatus('Manual mode off', 'info');
    } else {
        btn.classList.add('active');
        btn.textContent = 'Exit Manual';
        await sendBroadcastToFrames('MANUAL_MODE_ON');
        await sendMessageToContent('MANUAL_MODE_ON'); // direct fallback
        updateStatus('Manual mode on', 'info');
    }
}

async function extractElements() {
    try {
        const pageName = document.getElementById('page-name').value;
        const tab = await getCurrentTab();

        if (!tab) {
            updateStatus('No active tab found', 'error');
            return;
        }

        // First check if any elements are selected
        updateStatus('Checking selected elements...', 'warning');
        const statsResponse = await chrome.runtime.sendMessage({
            action: 'GET_STATS_ALL',
            tabId: tab.id
        });

        console.log('Stats response:', statsResponse);

        if (!statsResponse || statsResponse.selectedCount === 0) {
            updateStatus('No elements selected. Click "Auto Select" or "Manual Mode" first.', 'error');
            return;
        }

        updateStatus(`Extracting ${statsResponse.selectedCount} elements...`, 'warning');

        const store = await chrome.storage.local.get('label_overrides');
        const overrides = store.label_overrides || {};

        // Add timeout wrapper
        const timeoutPromise = new Promise((_, reject) =>
            setTimeout(() => reject(new Error('Extraction timed out after 10 seconds')), 10000)
        );

        const extractPromise = chrome.runtime.sendMessage({
            action: 'EXTRACT_ALL',
            tabId: tab.id,
            pageName,
            overrides
        });

        const response = await Promise.race([extractPromise, timeoutPromise]);

        console.log('Extract response:', response);

        if (response && response.success) {
            // Increment extraction count
            extractionCount++;
            await chrome.storage.local.set({ extraction_count: extractionCount });

            // Update both display locations
            document.getElementById('extract-count').textContent = extractionCount;
            document.getElementById('stat-extractions').textContent = extractionCount;

            // Update other stats
            await updateStats();
            updateStatus(`✅ Extracted ${response.count} elements (Total extractions: ${extractionCount})`, 'success');
        } else {
            const errorMsg = response?.error || 'Extraction failed. Unknown error.';
            updateStatus(`❌ ${errorMsg}`, 'error');
        }
    } catch (error) {
        console.error('Extract error:', error);
        updateStatus(`❌ Error: ${error.message}`, 'error');
    }
}

async function downloadData() {
    updateStatus('Preparing download...', 'warning');

    const allData = await chrome.storage.local.get('extractedData');
    if (!allData.extractedData || Object.keys(allData.extractedData).length === 0) {
        updateStatus('No data to download. Extract elements first.', 'error');
        return;
    }

    const timestamp = new Date().toISOString().slice(0, 10);

    // Download JSON
    const jsonBlob = new Blob([JSON.stringify(allData.extractedData, null, 2)], { type: 'application/json' });
    downloadBlob(jsonBlob, `element_extraction_${timestamp}.json`);

    // Download CSV
    const ovStore = await chrome.storage.local.get('label_overrides');
    const overrides = ovStore.label_overrides || {};
    const csvData = convertJsonToCsv(allData.extractedData, overrides);
    const csvBlob = new Blob(['\uFEFF' + csvData], { type: 'text/csv;charset=utf-8' });
    downloadBlob(csvBlob, `element_extraction_${timestamp}_mapping.csv`);

    updateStatus('Files downloaded successfully', 'success');
}

async function convertAndDownload() {
    updateStatus('Converting to JavaScript format...', 'warning');

    const allData = await chrome.storage.local.get('extractedData');
    if (!allData.extractedData || Object.keys(allData.extractedData).length === 0) {
        updateStatus('No data to convert. Extract elements first.', 'error');
        return;
    }

    const ovStore = await chrome.storage.local.get('label_overrides');
    const overrides = ovStore.label_overrides || {};
    const converted = convertJsonToJsFormat(allData.extractedData, overrides);
    const timestamp = new Date().toISOString().slice(0, 10);

    // Download JS function
    const jsBlob = new Blob([converted.jsFunction], { type: 'application/javascript' });
    downloadBlob(jsBlob, `element_extraction_${timestamp}_converted.js`);

    // Download CSV mapping
    const csvBlob = new Blob([converted.csvMapping], { type: 'text/csv' });
    downloadBlob(csvBlob, `element_extraction_${timestamp}_mapping.csv`);

    // Download carriers JSON
    const carriersBlob = new Blob([JSON.stringify(converted.carriers, null, 4)], { type: 'application/json' });
    downloadBlob(carriersBlob, `element_extraction_${timestamp}_carriers.json`);

    updateStatus('Conversion complete - files downloaded', 'success');
}

// removed Custom Export Template feature per request

async function generateSampleData() {
    const apiKeyResult = await chrome.storage.local.get(['deepseek_api_key', 'ai_provider', 'ai_model']);
    const apiKey = apiKeyResult.deepseek_api_key;
    const provider = apiKeyResult.ai_provider || 'deepseek';
    const model = apiKeyResult.ai_model || '';

    if (!apiKey) {
        updateStatus('Please enter your API key first', 'error');
        return;
    }

    try {
        // Phase 1: Preparation (0-10%)
        showProgress(5, 'Fetching extracted data...');
        const allData = await chrome.storage.local.get('extractedData');

        if (!allData.extractedData || Object.keys(allData.extractedData).length === 0) {
            hideProgress();
            updateStatus('No data found. Extract elements first.', 'error');
            return;
        }

        // Phase 2: Conversion (10-30%)
        showProgress(15, 'Converting data to JS format...');
        await new Promise(resolve => setTimeout(resolve, 200)); // Brief pause for UI update
        const converted = convertJsonToJsFormat(allData.extractedData);

        showProgress(25, 'Building actions list...');
        const allActions = buildActionsFromCarriers(converted.carriers, allData.extractedData);

        showProgress(30, `Preparing ${allActions.length} actions...`);
        const prompt = buildPromptForActions(allActions);

        // Phase 3: API Call (30-80%)
        showProgress(35, `Calling ${provider.toUpperCase()} API...`);
        updateStatus(`Generating data for ${allActions.length} fields...`, 'warning');

        // Simulate progress during API call (estimate 5-10 seconds)
        const startTime = Date.now();
        const estimatedTime = Math.min(allActions.length * 100, 10000); // 100ms per field, max 10s
        const progressInterval = setInterval(() => {
            const elapsed = Date.now() - startTime;
            const percent = Math.min(35 + (elapsed / estimatedTime) * 45, 79);
            showProgress(percent, `Generating ${allActions.length} fields...`);
        }, 500);

        const response = await chrome.runtime.sendMessage({
            action: 'CALL_AI_API',
            provider,
            model,
            apiKey: apiKey,
            prompt: prompt
        });

        clearInterval(progressInterval);

        if (response.error) {
            throw new Error(response.error);
        }

        // Phase 4: Enrichment (80-95%)
        showProgress(85, 'Processing AI response...');
// Fill sample values into overrides and UI instead of downloading
const ovStore = await chrome.storage.local.get('label_overrides');
const overrides = ovStore.label_overrides || {};
const actionsList = buildActionsFromCarriers(converted.carriers, allData.extractedData);
const keyByActionId = {};
actionsList.forEach(a => { keyByActionId[a.actionId] = overrideKey(a.contextDocument || 'document', a.selector); });
Object.keys(response.data || {}).forEach(actionId => {
    const k = keyByActionId[actionId];
    if (k) {
        overrides[k] = { ...(overrides[k]||{}), sample: response.data[actionId] };
    }
});
await chrome.storage.local.set({ label_overrides: overrides });
showProgress(100, 'Sample values ready');
updateStatus('Sample values populated. Review in Preview & Edit.', 'success');// Hide progress after 2 seconds
        setTimeout(() => hideProgress(), 2000);

    } catch (error) {
        console.error('AI generation error:', error);
        hideProgress();
        updateStatus(`AI Error: ${error.message}`, 'error');
    }
}

function showProgress(percent, text) {
    const progressContainer = document.getElementById('ai-progress');
    const progressFill = document.getElementById('progress-fill');
    const progressText = document.getElementById('progress-text');
    const progressPercent = document.getElementById('progress-percent');

    if (progressContainer) progressContainer.style.display = 'block';
    if (progressFill) progressFill.style.width = `${Math.min(percent, 100)}%`;
    if (progressText) progressText.textContent = text || 'Processing...';
    if (progressPercent) progressPercent.textContent = `${Math.round(percent)}%`;
}

function hideProgress() {
    const progressContainer = document.getElementById('ai-progress');
    if (progressContainer) progressContainer.style.display = 'none';

    // Reset
    const progressFill = document.getElementById('progress-fill');
    if (progressFill) progressFill.style.width = '0%';
}

async function clearAllAndExit() {
    const confirmClear = confirm('This will clear all extracted data and exit selection mode. Continue?');
    if (!confirmClear) return;

    // Clear storage
    await chrome.storage.local.remove('extractedData');

    // Reset extraction count
    extractionCount = 0;
    await chrome.storage.local.set({ extraction_count: 0 });
    document.getElementById('extract-count').textContent = '0';

    // Exit manual mode if active
    const manualBtn = document.getElementById('btn-manual');
    if (manualBtn.classList.contains('active')) {
        manualBtn.classList.remove('active');
        manualBtn.textContent = 'Manual Mode';
    }

    // Send message to content script to clear everything
    await sendBroadcastToFrames('CLEAR_ALL');

    // Reset stats
    selectedCount = 0;
    updateStats();

    updateStatus('All data cleared and exited selection mode', 'success');
}

function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}

function updateStatus(message, type = 'info') {
    const statusBox = document.getElementById('status');
    statusBox.textContent = message;
    statusBox.className = 'status-box';

    if (type === 'success') {
        statusBox.classList.add('success');
    } else if (type === 'error') {
        statusBox.classList.add('error');
    } else if (type === 'warning') {
        statusBox.classList.add('warning');
    }
}

async function updateStats() {
    const tab = await getCurrentTab();
    const [stats, summary] = await Promise.all([
        chrome.runtime.sendMessage({ action: 'GET_STATS_ALL', tabId: tab.id }),
        chrome.runtime.sendMessage({ action: 'GET_SELECTED_SUMMARY_ALL', tabId: tab.id })
    ]);
    if (stats) {
        selectedCount = stats.selectedCount || 0;
        document.getElementById('stat-selected').textContent = selectedCount;
    }
    document.getElementById('stat-extractions').textContent = extractionCount;
    if (summary && summary.success) {
        const bt = summary.byType || {};
        document.getElementById('stat-inputs').textContent = totalByType(bt, ['input', 'textarea']);
        document.getElementById('stat-buttons').textContent = totalByType(bt, ['button']);
        document.getElementById('stat-selects').textContent = totalByType(bt, ['select']);
    }
}

function totalByType(byType, keys) {
    return keys.reduce((sum, k) => sum + (byType[k] || 0), 0);
}

async function renderPreview() {
    const container = document.getElementById('preview-list');
    if (!container) return;
    container.textContent = 'Loading...';
    const tab = await getCurrentTab();
    let items = [];
    let summary = null;
    if (previewSource === 'storage' && currentLoadedSession) {
        const all = await chrome.storage.local.get('extractedData');
        const sessions = all.extractedData || {};
        const page = sessions[currentLoadedSession.pageId];
        const group = page?.extractions?.find(g => g.groupId === currentLoadedSession.groupId);
        if (!group) { container.textContent = 'No saved extraction found'; return; }
        items = group.elements || [];
        summary = { success: true, items: items.map(el => ({ label: el.label, type: el.type, selector: el.selector, contextDocument: el.contextDocument || 'document', rawSelector: el.selector.replace(/^.*>>>\s*/,'') , matches: 1, attributes: el.attributes, accessibility: el.accessibility })) };
    } else if (previewSource === 'storageAllTree' && currentLoadedTree) {
        // Render tree view
        const ovStore = await chrome.storage.local.get('label_overrides'); currentOverrides = ovStore.label_overrides || {}; renderTree(container);
        return;
    } else {
        summary = await chrome.runtime.sendMessage({ action: 'GET_SELECTED_SUMMARY_ALL', tabId: tab.id });
        if (!summary || !summary.success) { container.textContent = 'No selection'; return; }
    }
    container.textContent = '';
    lastPreview = [];

    // Load overrides
    const store = await chrome.storage.local.get('label_overrides');
    const overrides = store.label_overrides || {};

    summary.items.forEach((item, idx) => {
        const row = document.createElement('div');
        row.className = 'preview-item';

        const labelInput = document.createElement('input');
        labelInput.className = 'edit-input preview-label';
        const displayLabel = computeDisplayLabel(item.label, item);
        labelInput.value = overrides[oKey]?.label || displayLabel || '';
        labelInput.placeholder = 'Label';
        labelInput.addEventListener('change', async () => {
            overrides[oKey] = { ...(overrides[oKey]||{}), label: labelInput.value };
            await chrome.storage.local.set({ label_overrides: overrides });
        });

        // Sample Value input (under label)
        const sampleInput = document.createElement('input');
        sampleInput.className = 'edit-input';
        sampleInput.placeholder = 'Sample Value';
        sampleInput.value = overrides[oKey]?.sample || '';
        sampleInput.addEventListener('change', async () => {
            overrides[oKey] = { ...(overrides[oKey]||{}), sample: sampleInput.value };
            await chrome.storage.local.set({ label_overrides: overrides });
        });

        const groupInput = document.createElement('input');
        groupInput.className = 'edit-input preview-type';
        groupInput.value = overrides[oKey]?.group || '';
        groupInput.placeholder = 'Group (optional)';
        groupInput.addEventListener('change', async () => {
            overrides[oKey] = { ...(overrides[oKey]||{}), group: groupInput.value };
            await chrome.storage.local.set({ label_overrides: overrides });
        });

        const selectorWrap = document.createElement('div');
        selectorWrap.className = 'preview-selector';
        const contextInput = document.createElement('input');
        contextInput.className = 'edit-selector';
        contextInput.placeholder = 'ContextDocument (iframe or document)';
        contextInput.style.marginBottom = '4px';
        const originalSelector = item.selector;
        const originalContext = item.contextDocument || 'document';
        contextInput.value = overrides[oKey]?.contextDocument || originalContext;
        const selectorInput = document.createElement('input');
        selectorInput.className = 'edit-selector';
        selectorInput.placeholder = 'Selector';
        const currentOverride = overrides[oKey]?.newSelector;
        selectorInput.value = currentOverride || item.selector;

        const validateBoth = async () => {
            // Extract raw iframe selector from formatted context for validation
            const rawIframeSelector = extractRawIframeSelector(contextInput.value);
            const fullSel = rawIframeSelector ? `${rawIframeSelector} >>> ${selectorInput.value}` : selectorInput.value;
            await validateSingleRawSelector(computeRawSelector(fullSel), status);
        };

        selectorInput.addEventListener('change', async () => {
            const val = selectorInput.value.trim();
            overrides[oKey] = { ...(overrides[oKey]||{}), newSelector: val };
            await chrome.storage.local.set({ label_overrides: overrides });
            // update lastPreview raw selector for validation
            const lp = lastPreview.find(p => p.rawSelector === computeRawSelector(originalSelector));
            if (lp) lp.rawSelector = computeRawSelector(val);
            // instant revalidate
            await validateBoth();
        });

        contextInput.addEventListener('change', async () => {
            overrides[oKey] = { ...(overrides[oKey]||{}), contextDocument: contextInput.value };
            await chrome.storage.local.set({ label_overrides: overrides });
            await validateBoth();
        });

        const status = document.createElement('span');
        status.textContent = item.matches === 1 ? 'unique' : `${item.matches} matches`;
        status.className = item.matches === 1 ? 'selector-unique' : 'selector-multi';
        selectorWrap.appendChild(contextInput);
        selectorWrap.appendChild(selectorInput);

        const actions = document.createElement('div');
        actions.className = 'preview-actions';
        const btnView = document.createElement('button');
        btnView.textContent = 'View';
        btnView.addEventListener('click', async () => {
            // Extract raw iframe selector from formatted context for highlighting
            const rawIframeSelector = extractRawIframeSelector(contextInput.value);
            const fullSel = rawIframeSelector ? `${rawIframeSelector} >>> ${selectorInput.value}` : selectorInput.value;
            await sendBroadcastToFrames('PREVIEW_HIGHLIGHT', { rawSelector: computeRawSelector(fullSel), contextDocument: contextInput.value || 'document' });
        });

        actions.appendChild(btnView);
        row.appendChild(labelInput);
        row.appendChild(sampleInput);
        // Add type dropdown next to label
        const typeSelect = document.createElement('select');
        typeSelect.className = 'edit-input';
        ;['Input','Button','Select','RadioButton','Checkbox'].forEach(t => { const o=document.createElement('option'); o.value=t; o.textContent=t; typeSelect.appendChild(o); });
        typeSelect.value = overrides[oKey]?.type || mapTypeToOption(item.type);
        typeSelect.addEventListener('change', async () => {
            overrides[oKey] = { ...(overrides[oKey]||{}), type: typeSelect.value };
            await chrome.storage.local.set({ label_overrides: overrides });
        });
        row.appendChild(typeSelect);
        row.appendChild(groupInput);
        row.appendChild(selectorWrap);
        row.appendChild(status);
        row.appendChild(actions);
        container.appendChild(row);
        lastPreview.push({ rawSelector: item.rawSelector || computeRawSelector(selectorInput.value), statusEl: status, rowEl: row });
    });
}

function computeRawSelector(full) {
    if (!full) return '';
    const parts = full.split('>>>');
    return (parts[parts.length - 1] || '').trim();
}

function extractRawIframeSelector(formattedContext) {
    // Extract raw selector from "document.querySelector('...').contentWindow.document"
    if (!formattedContext || formattedContext === 'document') {
        return '';
    }
    const match = formattedContext.match(/document\.querySelector\('(.+?)'\)\.contentWindow\.document/);
    if (match && match[1]) {
        // Unescape single quotes
        return match[1].replace(/\\'/g, "'");
    }
    // If it doesn't match the pattern, assume it's already a raw selector
    return formattedContext;
}

function computeDisplayLabel(label, item) {
    let l = (label || '').trim();
    const generic = !l || /^(input|select|button|textarea)(_|\[)/i.test(l) || l.toLowerCase() === 'unknown';
    if (!generic) return l.replace(/\s*:\s*$/, '').trim();
    const attrs = item.attributes || {};
    if (attrs.placeholder) return attrs.placeholder;
    if (attrs.name) return attrs.name;
    if (attrs.id) return attrs.id;
    return l;
}

function elementKey(el) {
    return `${(el.selector||'').trim()}__${(el.label||'').trim()}__${(el.type||'').trim()}`;
}

function mapTypeToOption(t) {
    const s = (t || '').toString().toLowerCase();
    // Order matters: check 'radio' before 'button' to avoid matching 'radiobutton' as 'button'
    if (s.includes('radio')) return 'RadioButton';
    if (s.includes('checkbox')) return 'Checkbox';
    if (s.includes('select')) return 'Select';
    if (s.includes('button') || s.includes('submit')) return 'Button';
    return 'Input';
}

function renderTree(container) {
    container.textContent = '';
    lastPreview = [];
    const pages = (currentLoadedTree && currentLoadedTree.pages) || [];
    pages.forEach(page => {
        const dPage = document.createElement('details');
        dPage.className = 'tree-page'; dPage.open = true;
        const sPage = document.createElement('summary'); sPage.textContent = page.pageName || page.pageId;
        const metaP = document.createElement('span'); metaP.className = 'meta';
        const groups = page.groups || []; const elemCount = groups.reduce((s,g)=> s + (g.elements||[]).length, 0);
        metaP.textContent = ` • groups: ${groups.length} • elements: ${elemCount}`;
        // Normalize to ASCII separators
        metaP.textContent = ` - groups: ${groups.length} - elements: ${elemCount}`;
        sPage.appendChild(metaP); dPage.appendChild(sPage);
        // Page-level actions
        const pageActions = document.createElement('div');
        pageActions.style.margin = '6px 0 6px 0';
        const btnNewGroup = document.createElement('button'); btnNewGroup.textContent = 'New Group';
        btnNewGroup.addEventListener('click', async () => {
            const tempId = pendingTempGroupIdCounter--;
            if (!page.groups) page.groups = [];
            page.groups.push({ groupId: tempId, timestamp: new Date().toISOString(), groupName: '', elements: [] });
            pendingGroupAdds.push({ pageId: page.pageId, tempId });
            const ovStore = await chrome.storage.local.get('label_overrides'); currentOverrides = ovStore.label_overrides || {}; renderTree(container);
        });
        pageActions.appendChild(btnNewGroup);
        dPage.appendChild(pageActions);

        groups.forEach(group => {
            const dGroup = document.createElement('details'); dGroup.className = 'tree-group'; dGroup.open = true;
            const sGroup = document.createElement('summary');
            sGroup.textContent = `Group ${group.groupId}`;
            const metaG = document.createElement('span'); metaG.className='meta'; metaG.textContent = group.timestamp?` • ${group.timestamp}`:''; metaG.textContent = group.timestamp ? ` - ${group.timestamp}` : ``; sGroup.appendChild(metaG);
            dGroup.appendChild(sGroup);

            const groupNameWrap = document.createElement('div'); groupNameWrap.style.margin = '6px 0 0 0';
            const groupNameInput = document.createElement('input'); groupNameInput.className='edit-input'; groupNameInput.placeholder='Group name (optional)'; groupNameInput.style.minWidth='220px';
            groupNameInput.value = group.groupName || '';
            groupNameInput.addEventListener('change', () => {
                if (!pendingGroupNames[page.pageId]) pendingGroupNames[page.pageId] = {};
                pendingGroupNames[page.pageId][group.groupId] = groupNameInput.value.trim();
            });
            // Delete empty group
            if ((group.elements||[]).length === 0) {
                const delBtn = document.createElement('button'); delBtn.textContent = 'Delete Group'; delBtn.style.marginLeft = '8px';
                delBtn.addEventListener('click', async () => {
                    pendingGroupRemoves.push({ pageId: page.pageId, groupId: group.groupId });
                    // Remove from in-memory tree view
                    const idx = page.groups.findIndex(g => g.groupId === group.groupId);
                    if (idx >= 0) page.groups.splice(idx,1);
                    const ovStore = await chrome.storage.local.get('label_overrides'); currentOverrides = ovStore.label_overrides || {}; renderTree(container);
                });
                groupNameWrap.appendChild(delBtn);
            }
            groupNameWrap.appendChild(groupNameInput); dGroup.appendChild(groupNameWrap);

            const list = document.createElement('div'); list.className='tree-elements';
            const groupsForSelect = groups.map(g => ({ id: g.groupId, name: g.groupName||`Group ${g.groupId}` }));

            (group.elements||[]).forEach((el, elIndex) => {
                const row = document.createElement('div'); row.className='preview-item card';
                const labelInput = document.createElement('input'); labelInput.className='edit-input preview-label'; labelInput.placeholder='Label'; labelInput.value=computeDisplayLabel(el.label,{attributes:el.attributes})||'';
                const sampleKey = overrideKey(el.contextDocument||'document', el.selector);
                const sampleInput = document.createElement('input'); sampleInput.className='edit-input'; sampleInput.placeholder='Sample Value'; sampleInput.value = ((currentOverrides[sampleKey]?.sample) || el.sample || '');
                const typeSelect = document.createElement('select'); typeSelect.className='edit-input';
                ;['Input','Button','Select','RadioButton','Checkbox'].forEach(t => { const o=document.createElement('option'); o.value=t; o.textContent=t; typeSelect.appendChild(o); });
                typeSelect.value = mapTypeToOption(el.type);
                typeSelect.addEventListener('change', () => {
                    const key = elementKey(el);
                    const existing = pendingTypeChanges.find(p => p.pageId===page.pageId && p.groupId===group.groupId && p.key===key);
                    if (existing) existing.type = typeSelect.value; else pendingTypeChanges.push({ pageId: page.pageId, groupId: group.groupId, key, type: typeSelect.value });
                });
                const selectorWrap = document.createElement('div'); selectorWrap.className='preview-selector';
                const contextInput = document.createElement('input'); contextInput.className='edit-selector'; contextInput.placeholder='ContextDocument (iframe or document)'; contextInput.value=el.contextDocument||'document'; contextInput.style.marginBottom='4px';
                const selectorInput = document.createElement('input'); selectorInput.className='edit-selector'; selectorInput.placeholder='Selector'; selectorInput.value=el.selector||'';
                const groupSelect = document.createElement('select'); groupSelect.className='edit-input'; groupSelect.style.minWidth='160px';
                groupsForSelect.forEach(opt=>{ const o=document.createElement('option'); o.value=String(opt.id); o.textContent=opt.name; if(opt.id===group.groupId)o.selected=true; groupSelect.appendChild(o); });
                const orderInput = document.createElement('input'); orderInput.type='number'; orderInput.min='0'; orderInput.value=String(elIndex); orderInput.className='edit-input'; orderInput.style.width='80px';
                const status = document.createElement('span'); status.textContent='...'; status.className='selector-multi';
                const actions = document.createElement('div'); actions.className='preview-actions';
                const btnView = document.createElement('button'); btnView.textContent='View'; btnView.addEventListener('click', async()=>{
                    const rawIframeSelector = extractRawIframeSelector(contextInput.value);
                    const fullSel = rawIframeSelector ? `${rawIframeSelector} >>> ${selectorInput.value}` : selectorInput.value;
                    await sendBroadcastToFrames('PREVIEW_HIGHLIGHT',{
                        rawSelector: computeRawSelector(fullSel),
                        contextDocument: contextInput.value || 'document'
                    });
                });
                const btnRemove = document.createElement('button'); btnRemove.textContent='Remove'; btnRemove.addEventListener('click', ()=>{ pendingDeletes.push({ pageId: page.pageId, groupId: group.groupId, key: elementKey(el)}); row.style.opacity='0.5'; });
                const btnInsert = document.createElement('button'); btnInsert.textContent='Insert'; btnInsert.addEventListener('click',()=>{
                    const newRow=document.createElement('div'); newRow.className='preview-item card'; newRow.style.border='1px dashed #cbd5e1';
                    const nl=document.createElement('input'); nl.className='edit-input preview-label'; nl.placeholder='Label';
                    const ns=document.createElement('input'); ns.className='edit-selector'; ns.placeholder='Selector';
                    const nt=document.createElement('input'); nt.className='edit-input'; nt.placeholder='Type (e.g., input[text])';
                    const add=document.createElement('button'); add.textContent='Add'; add.addEventListener('click',()=>{ pendingInserts.push({ pageId: page.pageId, groupId: group.groupId, element:{ label:nl.value, selector:ns.value, type:nt.value }}); newRow.remove(); });
                    newRow.appendChild(nl); newRow.appendChild(nt); newRow.appendChild(ns); newRow.appendChild(add); list.insertBefore(newRow, row.nextSibling);
                });
                actions.appendChild(btnView); actions.appendChild(btnRemove); actions.appendChild(btnInsert);

                // Persist label/selector/context overrides
                labelInput.addEventListener('change', async()=>{ const store=await chrome.storage.local.get('label_overrides'); const ov=store.label_overrides||{}; const k=overrideKey(contextInput.value||'document', selectorInput.value||el.selector); ov[k]={ ...(ov[k]||{}), label: labelInput.value }; await chrome.storage.local.set({label_overrides:ov}); });
                sampleInput.addEventListener('change', async()=>{ const store=await chrome.storage.local.get('label_overrides'); const ov=store.label_overrides||{}; const k=overrideKey(contextInput.value||'document', selectorInput.value||el.selector); ov[k]={ ...(ov[k]||{}), sample: sampleInput.value }; await chrome.storage.local.set({label_overrides:ov}); });
                selectorInput.addEventListener('change', async()=>{ const store=await chrome.storage.local.get('label_overrides'); const ov=store.label_overrides||{}; const k=overrideKey(contextInput.value||'document', el.selector); ov[k]={ ...(ov[k]||{}), newSelector: selectorInput.value }; await chrome.storage.local.set({label_overrides:ov}); });
                contextInput.addEventListener('change', async()=>{ const store=await chrome.storage.local.get('label_overrides'); const ov=store.label_overrides||{}; const k=overrideKey(contextInput.value||'document', selectorInput.value||el.selector); ov[k]={ ...(ov[k]||{}), contextDocument: contextInput.value }; await chrome.storage.local.set({label_overrides:ov}); });
                // Validate using combined full selector for validation
                const validateBoth = async()=>{ const rawIframeSelector = extractRawIframeSelector(contextInput.value); const fullSel = rawIframeSelector ? `${rawIframeSelector} >>> ${selectorInput.value}` : selectorInput.value; await validateSingleRawSelector(computeRawSelector(fullSel), status); };
                selectorInput.addEventListener('change', validateBoth);
                contextInput.addEventListener('change', validateBoth);
                groupSelect.addEventListener('change', ()=>{ pendingMoves.push({ pageId: page.pageId, fromGroupId: group.groupId, key: elementKey(el), toGroupId: Number(groupSelect.value), toIndex: Number(orderInput.value)}); });
                orderInput.addEventListener('change', ()=>{ pendingMoves.push({ pageId: page.pageId, fromGroupId: group.groupId, key: elementKey(el), toGroupId: Number(groupSelect.value), toIndex: Number(orderInput.value)}); });

                selectorWrap.appendChild(contextInput);
                selectorWrap.appendChild(selectorInput);
                row.appendChild(labelInput);
                row.appendChild(sampleInput);
                row.appendChild(typeSelect);
                row.appendChild(groupSelect);
                row.appendChild(orderInput);
                row.appendChild(selectorWrap);
                row.appendChild(status);
                row.appendChild(actions);
                list.appendChild(row);
                lastPreview.push({ rawSelector: computeRawSelector(selectorInput.value), statusEl: status, rowEl: row });
            });

            dGroup.appendChild(list); dPage.appendChild(dGroup);
        });
        container.appendChild(dPage);
    });
}

async function validateSingleRawSelector(raw, statusEl) {
    if (!raw || !statusEl) return;
    try {
        const tab = await getCurrentTab();
        const result = await chrome.runtime.sendMessage({ action: 'VALIDATE_RAW_SELECTORS', tabId: tab.id, rawSelectors: [raw] });
        if (result && result.success) {
            const c = result.counts ? (result.counts[raw] || 0) : 0;
            statusEl.textContent = (c === 1) ? 'unique' : `${c} matches`;
            statusEl.className = (c === 1) ? 'selector-unique' : 'selector-multi';
        }
    } catch (e) {
        // ignore
    }
}

async function validateCurrentPreview() {
    if (!lastPreview || lastPreview.length === 0) {
        await renderPreview();
        if (!lastPreview || lastPreview.length === 0) { updateStatus('Nothing to validate', 'error'); return; }
    }
    const tab = await getCurrentTab();
    const rawSelectors = lastPreview.map(i => i.rawSelector).filter(Boolean);
    if (rawSelectors.length === 0) { updateStatus('No selectors found', 'error'); return; }
    const result = await chrome.runtime.sendMessage({ action: 'VALIDATE_RAW_SELECTORS', tabId: tab.id, rawSelectors });
    if (!result || !result.success) { updateStatus('Validation failed', 'error'); return; }
    const counts = result.counts || {};
    lastPreview.forEach(item => {
        const c = counts[item.rawSelector] || 0;
        item.statusEl.textContent = (c === 1) ? 'unique' : `${c} matches`;
        item.statusEl.className = (c === 1) ? 'selector-unique' : 'selector-multi';
    });
    updateStatus('Selectors validated', 'success');
}

async function loadAllExtractionsToPreview() {
    const all = await chrome.storage.local.get('extractedData');
    const sessions = all.extractedData || {};
    const pages = Object.values(sessions);
    if (pages.length === 0) { updateStatus('No saved extractions', 'error'); return; }
    // Build tree model
    currentLoadedTree = { pages: [] };
    pages.forEach(p => {
        currentLoadedTree.pages.push({
            pageId: p.pageId,
            pageName: p.pageName,
            url: p.url,
            groups: (p.extractions || []).map(g => ({ groupId: g.groupId, timestamp: g.timestamp, groupName: g.groupName, elements: g.elements || [] }))
        });
    });
    previewSource = 'storageAllTree';
    await renderPreview();
    await validateCurrentPreview();
    let total = 0; currentLoadedTree.pages.forEach(p => (p.groups||[]).forEach(g => total += (g.elements||[]).length));
    updateStatus(`Loaded all extractions (${total} elements)`, 'success');
}

async function savePreviewEditsToSession() {
    const store = await chrome.storage.local.get(['extractedData', 'label_overrides']);
    const sessions = store.extractedData || {};
    const overrides = store.label_overrides || {};
    let updated = 0;
    let elementsChanged = 0;
    const pagesImpacted = new Set();

    if (previewSource === 'storage' && currentLoadedSession) {
        const page = sessions[currentLoadedSession.pageId];
        if (!page) { updateStatus('Session not found', 'error'); return; }
        const groupIndex = page.extractions.findIndex(g => g.groupId === currentLoadedSession.groupId);
        if (groupIndex === -1) { updateStatus('Group not found', 'error'); return; }
        const group = page.extractions[groupIndex];
        (group.elements || []).forEach(el => {
            const o = overrides[overrideKey(el.contextDocument || 'document', el.selector)] || overrides[el.selector];
            if (!o) return;
            let changed = false;
            if (o.label && o.label !== el.label) { el.label = o.label; updated++; changed = true; }
            if (o.group && o.group !== el.group) { el.group = o.group; updated++; changed = true; }
            if (o.newSelector && o.newSelector !== el.selector) { el.selector = o.newSelector; updated++; changed = true; }
            if (o.sample !== undefined && String(o.sample) !== String(el.sample ?? '')) { el.sample = o.sample; updated++; changed = true; }
            if (o.contextDocument && o.contextDocument !== el.contextDocument) { el.contextDocument = o.contextDocument; updated++; changed = true; }
            if (o.type && o.type !== el.type) { el.type = o.type; updated++; changed = true; }
            if (changed) { elementsChanged++; pagesImpacted.add(currentLoadedSession.pageId); }
        });
        page.lastUpdated = new Date().toISOString();
        sessions[currentLoadedSession.pageId] = page;
    } else if (previewSource === 'storageAll') {
        // Apply overrides across all saved pages/groups
        Object.keys(sessions).forEach(pageId => {
            const page = sessions[pageId];
            (page.extractions || []).forEach(group => {
                (group.elements || []).forEach(el => {
                    const o = overrides[overrideKey(el.contextDocument || 'document', el.selector)] || overrides[el.selector];
                    if (!o) return;
                    let changed = false;
                    if (o.label && o.label !== el.label) { el.label = o.label; updated++; changed = true; }
                    if (o.group && o.group !== el.group) { el.group = o.group; updated++; changed = true; }
                    if (o.newSelector && o.newSelector !== el.selector) { el.selector = o.newSelector; updated++; changed = true; }
                    if (o.sample !== undefined && String(o.sample) !== String(el.sample ?? '')) { el.sample = o.sample; updated++; changed = true; }
                    if (o.contextDocument && o.contextDocument !== el.contextDocument) { el.contextDocument = o.contextDocument; updated++; changed = true; }
                    if (o.type && o.type !== el.type) { el.type = o.type; updated++; changed = true; }
                    if (changed) { elementsChanged++; pagesImpacted.add(pageId); }
                });
            });
            page.lastUpdated = new Date().toISOString();
            sessions[pageId] = page;
        });
    } else if (previewSource === 'storageAllTree') {
        // Apply overrides across all saved pages/groups
        Object.keys(sessions).forEach(pageId => {
            const page = sessions[pageId];
            (page.extractions || []).forEach(group => {
                (group.elements || []).forEach(el => {
                    const o = overrides[overrideKey(el.contextDocument || 'document', el.selector)] || overrides[el.selector];
                    if (!o) return;
                    let changed = false;
                    if (o.label && o.label !== el.label) { el.label = o.label; updated++; changed = true; }
                    if (o.group && o.group !== el.group) { el.group = o.group; updated++; changed = true; }
                    if (o.newSelector && o.newSelector !== el.selector) { el.selector = o.newSelector; updated++; changed = true; }
                    if (o.sample !== undefined && String(o.sample) !== String(el.sample ?? '')) { el.sample = o.sample; updated++; changed = true; }
                    if (o.contextDocument && o.contextDocument !== el.contextDocument) { el.contextDocument = o.contextDocument; updated++; changed = true; }
                    if (o.type && o.type !== el.type) { el.type = o.type; updated++; changed = true; }
                    if (changed) { elementsChanged++; pagesImpacted.add(pageId); }
                });
            });
            page.lastUpdated = new Date().toISOString();
            sessions[pageId] = page;
        });
    } else {
        updateStatus('Load a saved extraction first', 'error');
        return;
    }

    // Create newly added groups, assign real IDs
    const newGroupIdMap = {}; // { [pageId]: { [tempId]: newId } }
    (pendingGroupAdds||[]).forEach(add => {
        const page = sessions[add.pageId]; if (!page) return;
        const list = page.extractions || (page.extractions = []);
        let maxId = 0; list.forEach(g => { if (typeof g.groupId === 'number') maxId = Math.max(maxId, g.groupId); });
        const newId = maxId + 1;
        if (!newGroupIdMap[add.pageId]) newGroupIdMap[add.pageId] = {};
        newGroupIdMap[add.pageId][add.tempId] = newId;
        list.push({ groupId: newId, timestamp: new Date().toISOString(), elements: [] });
        page.lastUpdated = new Date().toISOString();
        sessions[add.pageId] = page;
    });

    // Apply group name overrides
    Object.keys(pendingGroupNames || {}).forEach(pageId => {
        const page = sessions[pageId];
        if (!page) return;
        const gmap = pendingGroupNames[pageId];
        (page.extractions||[]).forEach(g => {
            let key = g.groupId;
            // Support temp IDs by resolving mapping
            Object.keys(newGroupIdMap[pageId]||{}).forEach(tid => {
                if (Number(tid) === g.groupId) key = newGroupIdMap[pageId][tid];
            });
            const newName = gmap[key] ?? gmap[g.groupId];
            if (typeof newName === 'string') { g.groupName = newName; updated++; }
        });
        page.lastUpdated = new Date().toISOString();
        sessions[pageId] = page;
    });

    // Apply deletions
    (pendingDeletes||[]).forEach(del => {
        const page = sessions[del.pageId]; if (!page) return;
        const grp = (page.extractions||[]).find(g => g.groupId === del.groupId);
        if (!grp || !grp.elements) return;
        const idx = grp.elements.findIndex(e => elementKey(e) === del.key);
        if (idx >= 0) { grp.elements.splice(idx,1); elementsChanged++; updated++; }
        page.lastUpdated = new Date().toISOString();
        sessions[del.pageId] = page;
    });

    // Apply inserts
    (pendingInserts||[]).forEach(ins => {
        const page = sessions[ins.pageId]; if (!page) return;
        let gid = ins.groupId;
        if (gid < 0 && newGroupIdMap[ins.pageId] && newGroupIdMap[ins.pageId][gid] != null) {
            gid = newGroupIdMap[ins.pageId][gid];
        }
        const grp = (page.extractions||[]).find(g => g.groupId === gid);
        if (!grp) return; if (!grp.elements) grp.elements = [];
        const el = ins.element || {};
        grp.elements.push({
            label: (el.label||'').trim(),
            selector: (el.selector||'').trim(),
            type: (el.type||'').trim(),
            html: '', position: {}, attributes: {}, validation: {}, accessibility: {}, frame: {}
        });
        elementsChanged++; updated++;
        page.lastUpdated = new Date().toISOString();
        sessions[ins.pageId] = page;
    });

    // Apply moves (change group and/or order)
    (pendingMoves||[]).forEach(mv => {
        const page = sessions[mv.pageId]; if (!page) return;
        const from = (page.extractions||[]).find(g => g.groupId === mv.fromGroupId);
        let toId = mv.toGroupId;
        if (toId < 0 && newGroupIdMap[mv.pageId] && newGroupIdMap[mv.pageId][toId] != null) {
            toId = newGroupIdMap[mv.pageId][toId];
        }
        const to = (page.extractions||[]).find(g => g.groupId === toId);
        if (!from || !to) return;
        const idx = from.elements ? from.elements.findIndex(e => elementKey(e) === mv.key) : -1;
        if (idx < 0) return;
        const [elem] = from.elements.splice(idx,1);
        const insertAt = Math.max(0, Math.min((to.elements||[]).length, Number.isFinite(mv.toIndex)? mv.toIndex : to.elements.length));
        if (!to.elements) to.elements = [];
        to.elements.splice(insertAt, 0, elem);
        elementsChanged++; updated++;
        page.lastUpdated = new Date().toISOString();
        sessions[mv.pageId] = page;
    });

    // Remove empty groups
    (pendingGroupRemoves||[]).forEach(rem => {
        const page = sessions[rem.pageId]; if (!page) return;
        const list = page.extractions || [];
        const idx = list.findIndex(g => g.groupId === rem.groupId);
        if (idx >= 0 && (list[idx].elements||[]).length === 0) {
            list.splice(idx,1); updated++;
        }
        page.extractions = list; page.lastUpdated = new Date().toISOString();
        sessions[rem.pageId] = page;
    });

    // Apply explicit type changes captured from tree (fallback if overrides missed)
    (pendingTypeChanges||[]).forEach(tc => {
        const page = sessions[tc.pageId]; if (!page) return;
        const grp = (page.extractions||[]).find(g => g.groupId === tc.groupId);
        if (!grp || !grp.elements) return;
        const idx = grp.elements.findIndex(e => elementKey(e) === tc.key);
        if (idx >= 0) {
            grp.elements[idx].type = tc.type;
            updated++; elementsChanged++;
        }
        page.lastUpdated = new Date().toISOString();
        sessions[tc.pageId] = page;
    });

    await chrome.storage.local.set({ extractedData: sessions });
    updateStatus(`Changes applied (${updated} field updates)`, 'success');

    // Show compact summary instead of full list
    try {
        const container = document.getElementById('preview-list');
        if (container) {
            const pagesCount = pagesImpacted.size;
            container.innerHTML = `<div style="padding:10px; font-size:13px; color:#334155;">
  <div><strong>Applied:</strong> ${updated} field updates across ${elementsChanged} elements</div>
  <div><strong>Pages impacted:</strong> ${pagesCount}</div>
  <div style="margin-top:6px;">Use <em>Load All Extractions</em> to review again.</div>
</div>`;
        }
    } catch {}
    lastPreview = [];
    currentLoadedAggregated = [];
    currentLoadedTree = null;
    previewSource = null;
    pendingGroupNames = {}; pendingDeletes = []; pendingInserts = []; pendingMoves = []; pendingGroupAdds = []; pendingGroupRemoves = []; pendingTempGroupIdCounter = -1; pendingTypeChanges = [];
}


async function toggleTheme() {
    const isDark = document.body.classList.toggle('dark');
    await chrome.storage.local.set({ theme: isDark ? 'dark' : 'light' });
    updateStatus(`${isDark ? 'Dark' : 'Light'} mode enabled`, 'success');
}

async function applyHighlightColors() {
    const selected = document.getElementById('color-selected').value || '#ff4444';
    const hover = document.getElementById('color-hover').value || '#2196F3';
    await sendBroadcastToFrames('SET_HIGHLIGHT_COLORS', { selected, hover });
    updateStatus('Highlight colors applied', 'success');
}

async function applyLabelMode() {
    const mode = document.getElementById('label-mode')?.value || 'original';
    await chrome.storage.local.set({ label_mode: mode });
    await sendBroadcastToFrames('SET_LABEL_MODE', { mode });
    updateStatus(`Label mode set to ${mode}`, 'success');
}

function toggleSettings() {
    const header = document.getElementById('settings-toggle');
    const content = document.getElementById('settings-content');

    if (header && content) {
        header.classList.toggle('collapsed');
        content.classList.toggle('collapsed');
    }
}

function collapseSettings() {
    const header = document.getElementById('settings-toggle');
    const content = document.getElementById('settings-content');

    if (header && content) {
        header.classList.add('collapsed');
        content.classList.add('collapsed');
    }
}

// Listen for messages from content script
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.action === 'UPDATE_STATUS') {
        updateStatus(message.message, message.type);
        updateStats();
    } else if (message.action === 'UPDATE_STATS') {
        updateStats();
    } else if (message.action === 'REQUEST_DOWNLOAD') {
        downloadData();
    }
});

// Import utility functions from the original code
function convertJsonToCsv(inputData, overrides = {}) {
    const byNewSelector = {};
    Object.keys(overrides).forEach(k => {
        const o = overrides[k] || {}; if (o.newSelector) byNewSelector[o.newSelector] = o;
    });
    const elementMapping = [];
    let globalActionCounter = 1;

    for (const pageId in inputData) {
        const pageData = inputData[pageId];
        const domain = extractDomain(pageData.url || '');

        if (pageData.extractions) {
            pageData.extractions.forEach(extraction => {
                const pageName = pageData.pageName || `Extracted_Page_${pageId}`;
                const groupName = extraction.groupName || `Extracted_Group_${extraction.groupId}`;

                extraction.elements.forEach(element => {
                    const ov = overrides[element.selector] || byNewSelector[element.selector] || {};
                    const effType = ov.type || element.type;
                    const actionType = determineActionType(effType);
                    const actionId = `A_EXTRACTED_${globalActionCounter}`;
                    globalActionCounter++;

                    let label = element.label.trim();
                    if (!label || ['input_text', 'input_button', 'select_select-one', 'input_checkbox'].includes(label)) {
                        if (element.selector.startsWith('#')) {
                            label = element.selector.substring(1);
                        } else {
                            label = `element_${globalActionCounter - 1}`;
                        }
                    }

                    const htmlSnippet = element.html || '';

                    elementMapping.push({
                        ActionID: actionId,
                        TargetElement: element.selector || '',
                        ContextDocument: element.contextDocument || 'document',
                        Label: label,
                        ActionType: actionType,
                        ElementType: element.type || '',
                        Position: `x:${element.position?.x || ''}, y:${element.position?.y || ''}`,
                        Domain: domain,
                        PageName: pageName,
                        GroupName: groupName,
                        HTML: htmlSnippet
                    });
                });
            });
        }
    }

    return convertToCsvFormat(elementMapping);
}

function convertToCsvFormat(data) {
    if (!data || data.length === 0) {
        return 'ActionID,TargetElement,Label,ActionType,ElementType,Position,Domain,PageName,GroupName,HTML\n';
    }

    const headers = ['ActionID', 'TargetElement', 'Label', 'ActionType', 'ElementType',
                    'Position', 'Domain', 'PageName', 'GroupName', 'HTML'];
    let csv = headers.join(',') + '\n';

    data.forEach(row => {
        const values = headers.map(header => {
            let value = row[header] || '';
            if (typeof value === 'string' && (value.includes(',') || value.includes('"') || value.includes('\n'))) {
                value = '"' + value.replace(/"/g, '""') + '"';
            }
            return value;
        });
        csv += values.join(',') + '\n';
    });

    return csv;
}

function determineActionType(elementType) {
    const type = (elementType || '').toLowerCase();
    // Order matters: check radio before button to avoid matching 'radiobutton' as 'button'
    if (type.includes('radio')) return 'RadioButton';
    if (type.includes('checkbox')) return 'Checkbox';
    if (type.includes('select')) return 'Select';
    if (type.includes('button') || type.includes('submit')) return 'Button';
    return 'Input';
}

function extractDomain(url) {
    if (!url) return 'unknown.domain.com';
    try {
        const urlObj = new URL(url);
        return urlObj.hostname;
    } catch (e) {
        const match = url.match(/https?:\/\/([^\/]+)/);
        return match ? match[1] : 'unknown.domain.com';
    }
}

function convertJsonToJsFormat(inputData, overrides = {}) {
    const carriersByDomain = {};
    const elementMapping = [];
    const globalActionCounter = [1];
    const byNewSelector = {};
    Object.keys(overrides).forEach(k => { const o = overrides[k]||{}; if (o.newSelector) byNewSelector[o.newSelector]=o; });

    for (const pageId in inputData) {
        const pageData = inputData[pageId];
        const url = pageData.url || '';
        const domain = extractDomain(url);

        if (!carriersByDomain[domain]) {
            carriersByDomain[domain] = {
                carrier: "EXTRACTED",
                Domain: domain,
                Functions: {},
                JSON: []
            };
        }

        if (pageData.extractions) {
            pageData.extractions.forEach(extraction => {
                // Apply overrides to a shallow copy of elements
                const exCopy = { ...extraction, elements: (extraction.elements||[]).map(el => {
            const key = overrideKey(el.contextDocument || 'document', el.selector);
            const ov = overrides[key] || byNewSelector[el.selector] || {};
                    return {
                        ...el,
                        label: ov.label || el.label,
                        selector: ov.newSelector || el.selector,
                        type: ov.type || el.type,
                        // Ensure sample and contextDocument from overrides are used for Entry
                        sample: (ov.sample !== undefined) ? ov.sample : el.sample,
                        contextDocument: ov.contextDocument || el.contextDocument || 'document'
                    };
                }) };
                const { page, pageMapping } = convertExtractionToPage(
                    pageData, pageId, exCopy, globalActionCounter
                );
                // Merge DataGroups under the same page name
                const pagesArr = carriersByDomain[domain].JSON;
                const existing = pagesArr.find(p => p.PageName === page.PageName);
                if (existing) {
                    // Merge indicators (dedupe by TargetElement)
                    const existingIndicators = new Set((existing.PageIndicator || []).map(pi => pi.TargetElement));
                    (page.PageIndicator || []).forEach(pi => {
                        if (!existingIndicators.has(pi.TargetElement)) {
                            existing.PageIndicator.push(pi);
                            existingIndicators.add(pi.TargetElement);
                        }
                    });
                    // Append group
                    existing.DataGroups.push(...page.DataGroups);
                } else {
                    pagesArr.push(page);
                }
                elementMapping.push(...pageMapping);
            });
        }
    }

    const carriers = Object.values(carriersByDomain);
    const jsonTemplate = [{ rawnote: "reminderWording" }, ...carriers];
    const jsFunction = generateJsFunction(jsonTemplate);
    const csvMapping = convertMappingToCsv(elementMapping);

    return { jsFunction, csvMapping, carriers };
}

function convertExtractionToPage(pageData, pageId, extraction, globalActionCounter) {
    const elements = extraction.elements || [];
    const pageName = pageData.pageName || `Extracted_Page_${pageId}`;
    const groupName = `Extracted_Group_${extraction.groupId}`;
    const domain = extractDomain(pageData.url || '');

    const pageIndicators = [];
    if (elements.length > 0) {
        for (let i = 0; i < Math.min(3, elements.length); i++) {
            pageIndicators.push({
                TargetElement: elements[i].selector || '',
                ExpectedValue: "",
                MatchType: "Exists",
                ContextDocument: elements[i].contextDocument || "document"
            });
        }
    }

    const actions = [];
    const pageMapping = [];

    elements.forEach(element => {
        const { action, mapping } = convertElementToAction(element, globalActionCounter);
        actions.push(action);

        mapping.PageName = pageName;
        mapping.GroupName = groupName;
        mapping.Domain = domain;
        pageMapping.push(mapping);
    });

    const dataGroup = {
        GroupName: groupName,
        IsEntered: false,
        Actions: actions
    };

    const page = {
        PageName: pageName,
        PageIndicator: pageIndicators,
        DataGroups: [dataGroup]
    };

    return { page, pageMapping };
}

function convertElementToAction(element, actionCounter) {
    const elementType = (element.type || '').toLowerCase();
    let actionType = 'Input';

    // Order matters: check radio before button
    if (elementType.includes('radio')) {
        actionType = 'RadioButton';
    } else if (elementType.includes('checkbox')) {
        actionType = 'Checkbox';
    } else if (elementType.includes('select')) {
        actionType = 'Select';
    } else if (elementType.includes('button') || elementType.includes('submit')) {
        actionType = 'Button';
    }

    const actionId = `A_EXTRACTED_${actionCounter[0]}`;
    actionCounter[0]++;

    let label = (element.label || '').trim();
    if (!label || ['input_text', 'input_button', 'select_select-one', 'input_checkbox'].includes(label)) {
        const selector = element.selector || '';
        if (selector.startsWith('#')) {
            label = selector.substring(1);
        } else {
            label = `element_${actionCounter[0] - 1}`;
        }
    }

    const inputValue = (element.sample && String(element.sample).length > 0) ? String(element.sample) : `|setValue('${actionId}')|`;

    const action = {
        TargetElement: element.selector || '',
        InputValue: inputValue,
        ActionType: actionType,
        CustomCodeName: "",
        EventsConfig: "NULL",
        ContextDocument: element.contextDocument || "document",
        IsEntered: false
    };

    let htmlSnippet = element.html || '';
    if (htmlSnippet.length > 100) {
        htmlSnippet = htmlSnippet.substring(0, 100) + '...';
    }

    const mapping = {
        ActionID: actionId,
        TargetElement: element.selector || '',
        ContextDocument: element.contextDocument || "document",
        Label: label,
        ActionType: actionType,
        ElementType: element.type || '',
        Position: `x:${element.position?.x || ''}, y:${element.position?.y || ''}`,
        HTML: htmlSnippet
    };

    return { action, mapping };
}

function generateJsFunction(jsonTemplate) {
    const jsonStr = JSON.stringify(jsonTemplate, null, 0);

    return `function VBA_GENERATE_JSON(outputkeyDict, ReminderDict) {
    Object.keys(outputkeyDict).forEach(key => {
        const value = outputkeyDict[key];
        console.log(\`key: \${key}, value: \${value}, type: \${typeof value}\`);
    });
    const setValue = (key) => outputkeyDict.hasOwnProperty(key) ? outputkeyDict[key] : null;
    const reminderWording = Object.values(ReminderDict).join('|');
    const jsontemplate = ${jsonStr};
    return jsontemplate;
};`;
}

function convertMappingToCsv(mappingData) {
    if (!mappingData || mappingData.length === 0) {
        return 'ActionID,TargetElement,Label,ActionType,ElementType,Position,Domain,PageName,GroupName,HTML\n';
    }

    const headers = ['ActionID', 'TargetElement', 'Label', 'ActionType', 'ElementType',
                    'Position', 'Domain', 'PageName', 'GroupName', 'HTML'];
    let csv = headers.join(',') + '\n';

    mappingData.forEach(row => {
        const values = headers.map(header => {
            let value = row[header] || '';
            if (typeof value === 'string' && (value.includes(',') || value.includes('"') || value.includes('\n'))) {
                value = '"' + value.replace(/"/g, '""') + '"';
            }
            return value;
        });
        csv += values.join(',') + '\n';
    });

    return csv;
}

function buildActionsFromCarriers(carriers, extractedData) {
    const allActions = [];
    const selectorToElementMap = {};

    for (const pageId in extractedData) {
        const pageData = extractedData[pageId];
        if (pageData.extractions) {
            pageData.extractions.forEach(extraction => {
                extraction.elements.forEach(element => {
                    selectorToElementMap[element.selector] = element;
                });
            });
        }
    }

    const carriersArray = Array.isArray(carriers) ? carriers : [carriers];

    carriersArray.forEach(carrier => {
        if (carrier && carrier.JSON && Array.isArray(carrier.JSON)) {
            carrier.JSON.forEach(page => {
                if (page.DataGroups && Array.isArray(page.DataGroups)) {
                    page.DataGroups.forEach(group => {
                        if (group.Actions && Array.isArray(group.Actions)) {
                            group.Actions.forEach(action => {
                                const match = action.InputValue.match(/setValue\('([^']+)'\)/);
                                const actionId = match ? match[1] : null;

                                if (actionId) {
                                    const originalElement = selectorToElementMap[action.TargetElement];

                                    allActions.push({
                                            actionId: actionId,
                                            actionType: action.ActionType,
                                            selector: action.TargetElement,
                                            contextDocument: action.ContextDocument || 'document',
                                            label: originalElement ? originalElement.label : '',
                                            elementType: originalElement ? originalElement.type : '',
                                            placeholder: originalElement?.attributes?.placeholder || '',
                                            name: originalElement?.attributes?.name || '',
                                            validation: originalElement?.validation || {},
                                        htmlSnippet: originalElement?.html ? originalElement.html.slice(0, 150) + '...' : ''
                                    });
                                }
                            });
                        }
                    });
                }
            });
        }
    });

    return allActions;
}

function buildPromptForActions(actions) {
    let actionsDescription = '';

    actions.forEach((actionInfo, index) => {
        actionsDescription += `Action ${index + 1}:\n`;
        actionsDescription += `  - ActionID: "${actionInfo.actionId}"\n`;
        actionsDescription += `  - ActionType: ${actionInfo.actionType}\n`;
        actionsDescription += `  - Selector: ${actionInfo.selector}\n`;
        actionsDescription += `  - Label: "${actionInfo.label}"\n`;
        actionsDescription += `  - ElementType: ${actionInfo.elementType}\n`;

        if (actionInfo.placeholder) {
            actionsDescription += `  - Placeholder: "${actionInfo.placeholder}"\n`;
        }
        if (actionInfo.name) {
            actionsDescription += `  - Name attribute: "${actionInfo.name}"\n`;
        }
        if (actionInfo.validation) {
            actionsDescription += `  - Validation: ${JSON.stringify(actionInfo.validation)}\n`;
        }
        if (actionInfo.htmlSnippet) {
            actionsDescription += `  - HTML: ${actionInfo.htmlSnippet}\n`;
        }

        actionsDescription += '\n';
    });

    return `You are a helpful assistant that generates realistic sample data for web form automation.

I have extracted ${actions.length} form actions from a webpage with detailed context about each field. Please generate realistic and appropriate sample values for each action based on all the context provided (label, placeholder, validation rules, HTML attributes, etc.).

EXTRACTED FORM FIELDS WITH FULL CONTEXT:
${actionsDescription}

INSTRUCTIONS:
1. Analyze the Label, Placeholder, Name, and HTML context to understand what data each field expects
2. For Input actions: Generate realistic values matching the field purpose:
   - Email fields → realistic email (e.g., "john.smith@example.com")
   - Name fields → full name or first/last name as appropriate
   - Phone fields → valid phone format (e.g., "555-0123" or "(555) 123-4567")
   - Address fields → complete address
   - Date fields → date in appropriate format (YYYY-MM-DD)
   - Number fields → appropriate numbers
   - Password fields → strong password (e.g., "SecurePass123!")
   - Topic/Subject → descriptive topic (e.g., "Monthly Team Meeting")
   - Duration → time duration (e.g., "60")
3. For Select/Dropdown actions: Generate realistic option text (e.g., "Pacific Time (US & Canada)", "Option 1")
4. For Checkbox/RadioButton actions: Use "true" or ""
5. For Button actions: Use "true" or "" (empty string)
6. Consider validation rules if provided (e.g., maxlength, pattern, min/max)
7. Look at ALL context clues (label, placeholder, name, html) to infer the correct data type

Return ONLY valid JSON with ActionID as keys and sample values as values.

CRITICAL: Return ONLY the JSON object, absolutely no other text, explanations, or markdown formatting.

Expected format:
{
  "A_EXTRACTED_1": "sample value 1",
  "A_EXTRACTED_2": "sample value 2",
  "A_EXTRACTED_3": "sample value 3"
}`;
}

function enrichCarriersWithAIData(carriers, aiSampleData) {
    const carriersArray = Array.isArray(carriers) ? carriers : [carriers];
    const enriched = JSON.parse(JSON.stringify(carriersArray));

    enriched.forEach(carrier => {
        if (!carrier.JSON || !Array.isArray(carrier.JSON)) return;

        carrier.JSON.forEach(page => {
            if (!page.DataGroups || !Array.isArray(page.DataGroups)) return;

            page.DataGroups.forEach(group => {
                if (!group.Actions || !Array.isArray(group.Actions)) return;

                group.Actions.forEach(action => {
                    const match = action.InputValue.match(/setValue\('([^']+)'\)/);
                    const actionId = match ? match[1] : null;

                    if (actionId && aiSampleData[actionId]) {
                        action.InputValue = aiSampleData[actionId];
                    } else if (actionId) {
                        action.InputValue = generateFallbackByActionType(action);
                    }
                });
            });
        });
    });

    return enriched;
}

function generateFallbackByActionType(action) {
    const selector = (action.TargetElement || '').toLowerCase();
    const type = action.ActionType;

    if (selector.includes('email')) return 'user@example.com';
    if (selector.includes('phone') || selector.includes('tel')) return '555-0123';
    if (selector.includes('name')) return 'John Doe';
    if (selector.includes('firstname') || selector.includes('fname')) return 'John';
    if (selector.includes('lastname') || selector.includes('lname')) return 'Doe';
    if (selector.includes('address')) return '123 Main Street';
    if (selector.includes('city')) return 'New York';
    if (selector.includes('state')) return 'NY';
    if (selector.includes('zip') || selector.includes('postal')) return '10001';
    if (selector.includes('country')) return 'United States';
    if (selector.includes('age')) return '30';
    if (selector.includes('date') || selector.includes('dob')) return '1990-01-01';
    if (selector.includes('password')) return 'Password123!';
    if (selector.includes('username')) return 'johndoe';
    if (selector.includes('company')) return 'Acme Corp';
    if (selector.includes('url') || selector.includes('website')) return 'https://example.com';
    if (selector.includes('topic')) return 'Sample Topic';
    if (selector.includes('duration')) return '60';

    if (type === 'Checkbox' || type === 'RadioButton') return 'true';
    if (type === 'Select') return 'Option 1';
    if (type === 'Button') return 'true';

    return 'Sample Text';
}



async function runEntryTest() {
    try {
        updateStatus('Preparing entry test...', 'warning');
        const allData = await chrome.storage.local.get('extractedData');
        if (!allData.extractedData || Object.keys(allData.extractedData).length === 0) {
            updateStatus('No data found. Extract elements first.', 'error');
            return;
        }
        const ovStore = await chrome.storage.local.get('label_overrides');
        const overrides = ovStore.label_overrides || {};
        const converted = convertJsonToJsFormat(allData.extractedData, overrides);
        const tab = await getCurrentTab();
        const domain = extractDomain(tab.url || '');
        const carriersArr = Array.isArray(converted.carriers) ? converted.carriers : [];
        let domainCarriers = carriersArr.filter(c => c.Domain === domain);
        if (domainCarriers.length === 0) domainCarriers = carriersArr; // fallback to all
        let dataGroups = [];
        console.log('ENTRY: carriers for domain', domain, domainCarriers.length);
        domainCarriers.forEach(c => {
            (c.JSON || []).forEach(p => {
                if (Array.isArray(p.DataGroups)) dataGroups.push(...p.DataGroups);
            });
        });
        console.log('ENTRY: dataGroups count', dataGroups.length, dataGroups.map(g => (g.Actions||[]).length));
        const resp = await chrome.runtime.sendMessage({ action: 'RUN_ENTRY', tabId: tab.id, dataGroups, functions: converted.functions || {} });
        if (resp && resp.success) {
            const totals = `applied ${resp.appliedActions||0}/${resp.totalActions||0}`;
            const blocked = resp.blockedContexts ? `, blocked: ${resp.blockedContexts}` : '';
            const missing = resp.missingElements ? `, missing: ${resp.missingElements}` : '';
            const skipped = resp.skippedFrame ? `, skipped (wrong frame): ${resp.skippedFrame}` : '';

            console.log('ENTRY: Full results:', resp);

            if (resp.appliedActions === resp.totalActions && resp.totalActions > 0) {
                updateStatus(`Entry applied successfully! (${totals})`, 'success');
            } else if (resp.appliedActions > 0) {
                updateStatus(`Entry partially applied (${totals}${blocked}${missing}${skipped}). Check console for details.`, 'warning');
            } else {
                updateStatus(`Entry failed to apply any actions (${totals}${blocked}${missing}${skipped}). Check console for details.`, 'error');
            }
        } else {
            updateStatus('Entry finished with errors.', 'error');
        }
    } catch (e) {
        console.error(e);
        updateStatus('Entry error: ' + (e.message||e), 'error');
    }
}








