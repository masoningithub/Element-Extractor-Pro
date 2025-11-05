// Background service worker
// Handles side panel opening and DeepSeek API calls to avoid CORS issues

// Open side panel when extension icon is clicked
chrome.action.onClicked.addListener((tab) => {
    chrome.sidePanel.open({ windowId: tab.windowId });
});

// Track frames per tab for multi-frame coordination
const tabFrames = new Map(); // tabId -> Set(frameId)

// Helper to register a frame when content script loads
function registerFrame(sender) {
    const tabId = sender?.tab?.id;
    const frameId = sender?.frameId;
    if (tabId == null || frameId == null) return;
    if (!tabFrames.has(tabId)) tabFrames.set(tabId, new Set());
    tabFrames.get(tabId).add(frameId);
}

// Clean up when tab is removed
chrome.tabs.onRemoved.addListener((tabId) => {
    tabFrames.delete(tabId);
});

// Convenience: send message to a particular frame
function sendToFrame(tabId, frameId, msg) {
    return new Promise((resolve) => {
        try {
            chrome.tabs.sendMessage(tabId, msg, { frameId }, (resp) => {
                // Consume lastError to avoid Unchecked runtime.lastError warnings
                if (chrome.runtime && chrome.runtime.lastError) {
                    // no-op: intentionally ignore errors for non-injectable pages/frames
                }
                resolve(resp);
            });
        } catch (e) {
            resolve(null);
        }
    });
}

// Content scripts are automatically injected via manifest.json content_scripts
// with all_frames: true, so no dynamic injection needed

// Listen for messages from side panel and content scripts
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.action === 'FRAME_READY') {
        registerFrame(sender);
        sendResponse && sendResponse({ ok: true });
        return; // no async
    }

    if (message.action === 'REQUEST_DOWNLOAD') {
        // Forward to side panel for handling, then ack immediately
        try { chrome.runtime.sendMessage({ action: 'REQUEST_DOWNLOAD' }); } catch (e) {}
        sendResponse({ ok: true });
        return; // sync
    }

    if (message.action === 'CALL_AI_API') {
        callAI(message.provider || 'deepseek', message.apiKey, message.model, message.prompt)
            .then(data => sendResponse({ success: true, data }))
            .catch(error => {
                console.error('AI API error:', error);
                sendResponse({ success: false, error: error.message });
            });
        return true; // async
    }

    // ENSURE_CONTENT removed - content scripts auto-inject via manifest

    // Broadcast a content-script action to all known frames of a tab
    if (message.action === 'BROADCAST') {
        const tabId = message.tabId || sender?.tab?.id;
        const frames = tabFrames.get(tabId) || new Set([0]);
        const tasks = Array.from(frames).map(fid => sendToFrame(tabId, fid, { action: message.targetAction, ...message.payload }));
        Promise.all(tasks).then(() => sendResponse({ ok: true })).catch(() => sendResponse({ ok: false }));
        return true;
    }

    // Gather selected counts from all frames
    if (message.action === 'GET_STATS_ALL') {
        const tabId = message.tabId || sender?.tab?.id;
        const frames = tabFrames.get(tabId) || new Set([0]);
        const tasks = Array.from(frames).map(fid => sendToFrame(tabId, fid, { action: 'GET_STATS' }));
        Promise.all(tasks).then(responses => {
            const total = (responses || []).reduce((sum, r) => sum + (r?.selectedCount || 0), 0);
            sendResponse({ selectedCount: total });
        });
        return true;
    }

    if (message.action === 'GET_SELECTED_SUMMARY_ALL') {
        const tabId = message.tabId || sender?.tab?.id;
        const frames = tabFrames.get(tabId) || new Set([0]);
        const tasks = Array.from(frames).map(fid => sendToFrame(tabId, fid, { action: 'GET_SELECTED_SUMMARY' }));
        Promise.all(tasks).then(responses => {
            const items = [];
            const byType = {};
            let selectedCount = 0;
            for (const r of responses) {
                if (r && r.success) {
                    selectedCount += r.selectedCount || 0;
                    (r.items||[]).forEach(i => items.push(i));
                    const bt = r.byType || {};
                    Object.keys(bt).forEach(k => { byType[k] = (byType[k]||0) + bt[k]; });
                }
            }
            sendResponse({ success: true, selectedCount, byType, items });
        });
        return true;
    }

    // Extract across all frames and save via top frame
    if (message.action === 'EXTRACT_ALL') {
        const tabId = message.tabId || sender?.tab?.id;
        const frames = tabFrames.get(tabId) || new Set([0]);
        const tasks = Array.from(frames).map(fid => sendToFrame(tabId, fid, { action: 'EXTRACT_PART' }));
        Promise.all(tasks).then(async (responses) => {
            const allElements = [];
            for (const r of responses) {
                if (r && r.success && Array.isArray(r.elements)) {
                    allElements.push(...r.elements);
                }
            }
            // Deduplicate by selector+label+type
            const seen = new Set();
            const unique = [];
            for (const el of allElements) {
                const key = `${el.selector}__${el.label}__${el.type}`;
                if (!seen.has(key)) { seen.add(key); unique.push(el); }
            }
            // Apply overrides if provided
            const overrides = message.overrides || {};
            unique.forEach(el => {
                const o = overrides[el.selector];
                if (o) {
                    if (o.label) el.label = o.label;
                    if (o.group) el.group = o.group;
                    if (o.newSelector) el.selector = o.newSelector;
                }
            });
            // Save via top frame (frameId 0)
            const resp = await sendToFrame(tabId, 0, { action: 'SAVE_EXTRACTION', pageName: message.pageName, elements: unique });
            sendResponse(resp || { success: true, count: unique.length });
        });
        return true;
    }

    if (message.action === 'VALIDATE_RAW_SELECTORS') {
        const tabId = message.tabId || sender?.tab?.id;
        const frames = tabFrames.get(tabId) || new Set([0]);
        const selectors = Array.isArray(message.rawSelectors) ? message.rawSelectors : [];
        const tasks = Array.from(frames).map(fid => sendToFrame(tabId, fid, { action: 'VALIDATE_RAW_SELECTORS_FRAME', rawSelectors: selectors }));
        Promise.all(tasks).then(responses => {
            const totals = {};
            selectors.forEach(s => totals[s] = 0);
            (responses || []).forEach(r => {
                if (r && r.success && r.counts) {
                    Object.keys(r.counts).forEach(k => { totals[k] += (r.counts[k] || 0); });
                }
            });
            sendResponse({ success: true, counts: totals });
        });
        return true;
    }

    // Run Entry actions on the active tab (all frames)
    if (message.action === 'RUN_ENTRY') {
        const tabId = message.tabId || sender?.tab?.id;
        const frames = tabFrames.get(tabId) || new Set([0]);
        const payload = { action: 'RUN_ENTRY', dataGroups: message.dataGroups || [], functions: message.functions || {} };
        const tasks = Array.from(frames).map(fid => sendToFrame(tabId, fid, payload));
        Promise.all(tasks).then(resps => {
            const agg = (resps || []).reduce((acc, r) => {
                if (!r || !r.success) return acc;
                acc.totalActions += (r.totalActions || 0);
                acc.appliedActions += (r.appliedActions || 0);
                acc.missingElements += (r.missingElements || 0);
                acc.blockedContexts += (r.blockedContexts || 0);
                acc.skippedFrame += (r.skippedFrame || 0);
                return acc;
            }, { totalActions: 0, appliedActions: 0, missingElements: 0, blockedContexts: 0, skippedFrame: 0 });
            console.log('ENTRY: Aggregated results from all frames:', agg);
            sendResponse({ success: true, ...agg });
        }).catch(() => sendResponse({ success: false }));
        return true; // async response
    }
});

async function callAI(provider, apiKey, model, prompt) {
    const sysMsg = 'You are a helpful assistant that generates realistic sample data for web forms. Always return valid JSON only.';
    let url, body;
    switch (provider) {
        case 'azure': {
            // Get Azure settings from storage
            const azureSettings = await chrome.storage.local.get(['azure_endpoint', 'azure_deployment', 'azure_api_version']);
            const endpoint = azureSettings.azure_endpoint;

            if (!endpoint) {
                throw new Error('Azure endpoint not configured. Please set it in Settings â†’ Azure Configuration.');
            }

            const deployment = azureSettings.azure_deployment || model || 'gpt-4';
            const apiVersion = azureSettings.azure_api_version || '2024-02-01';

            // Ensure endpoint has trailing slash
            const normalizedEndpoint = endpoint.endsWith('/') ? endpoint : endpoint + '/';
            url = `${normalizedEndpoint}openai/deployments/${deployment}/chat/completions?api-version=${apiVersion}`;

            body = {
                messages: [
                    { role: 'system', content: sysMsg },
                    { role: 'user', content: prompt }
                ],
                temperature: 0.7,
                top_p: 0.95,
                frequency_penalty: 0.0,
                presence_penalty: 0.0
            };
            const headers = { 'Content-Type': 'application/json', 'api-key': apiKey };
            const response = await fetch(url, { method: 'POST', headers, body: JSON.stringify(body) });
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Azure OpenAI API error (${response.status}): ${errorText}`);
            }
            const data = await response.json();
            const content = data.choices?.[0]?.message?.content;
            const jsonMatch = content && content.match(/\{[\s\S]*\}/);
            if (!jsonMatch) throw new Error('Invalid JSON response from Azure AI. Response: ' + content);
            return JSON.parse(jsonMatch[0]);
        }
        case 'openai':
            url = 'https://api.openai.com/v1/chat/completions';
            body = { model: model || 'gpt-4o-mini', messages: [{ role:'system', content: sysMsg }, { role:'user', content: prompt }], temperature: 0.7 };
            break;
        case 'anthropic':
            url = 'https://api.anthropic.com/v1/messages';
            body = { model: model || 'claude-3-5-sonnet-20241022', max_tokens: 1024, system: sysMsg, messages: [{ role: 'user', content: prompt }] };
            break;
        case 'google':
            url = 'https://generativelanguage.googleapis.com/v1beta/models/' + (model || 'gemini-1.5-flash') + ':generateContent?key=' + encodeURIComponent(apiKey);
            body = { contents: [{ parts: [{ text: sysMsg }, { text: prompt }] }] };
            break;
        default:
            url = 'https://api.deepseek.com/chat/completions';
            body = { model: model || 'deepseek-chat', messages: [{ role:'system', content: sysMsg }, { role:'user', content: prompt }], stream: false, temperature: 0.7 };
    }
    const headers = { 'Content-Type': 'application/json' };
    if (provider !== 'google') headers['Authorization'] = `Bearer ${apiKey}`;

    try {
        const response = await fetch(url, { method: 'POST', headers, body: JSON.stringify(body) });

        if (!response.ok) {
            const status = response.status;
            let errorText = '';
            try {
                errorText = await response.text();
            } catch (e) {
                errorText = 'Unable to read error response';
            }

            // Provide specific error messages based on status code
            if (status === 401) {
                throw new Error(`âŒ Invalid API key for ${provider}.\n\nâœ… Fix: Check your API key in Settings and make sure it's correct.`);
            } else if (status === 429) {
                throw new Error(`â±ï¸ Rate limit exceeded for ${provider}.\n\nâœ… Fix: Wait a few moments and try again. Consider upgrading your API plan if this persists.`);
            } else if (status === 402 || status === 403) {
                throw new Error(`ðŸ’³ Insufficient credits for ${provider}.\n\nâœ… Fix: Add credits to your ${provider} account at their website.`);
            } else if (status === 400) {
                throw new Error(`âš ï¸ Bad request to ${provider} API.\n\nâœ… Fix: Check your model name in Settings. Error details: ${errorText.substring(0, 200)}`);
            } else if (status === 404) {
                throw new Error(`ðŸ” API endpoint not found for ${provider}.\n\nâœ… Fix: Verify your endpoint URL (for Azure) or model name in Settings.`);
            } else if (status >= 500) {
                throw new Error(`ðŸ”§ ${provider} server error (${status}).\n\nâœ… Fix: The AI provider is experiencing issues. Try again in a few minutes.`);
            } else {
                throw new Error(`âŒ ${provider} API error (${status}).\n\nDetails: ${errorText.substring(0, 300)}`);
            }
        }

        const data = await response.json();
        let content;
        if (provider === 'openai' || provider === 'deepseek') content = data.choices?.[0]?.message?.content;
        else if (provider === 'anthropic') content = data.content?.[0]?.text;
        else if (provider === 'google') content = data.candidates?.[0]?.content?.parts?.map(p=>p.text).join('\n');

        if (!content) {
            throw new Error(`ðŸ“­ Empty response from ${provider}.\n\nâœ… Fix: Try again or use a different AI provider.`);
        }

        const jsonMatch = content.match(/\{[\s\S]*\}/);
        if (!jsonMatch) {
            throw new Error(`ðŸ“‹ ${provider} returned invalid JSON.\n\nâœ… Fix: Try again. If this persists, the AI model may not support this task well. Try a different provider.\n\nResponse preview: ${content.substring(0, 200)}...`);
        }

        return JSON.parse(jsonMatch[0]);

    } catch (error) {
        // Handle network errors
        if (error instanceof TypeError && error.message.includes('fetch')) {
            throw new Error(`ðŸŒ Network error: Cannot reach ${provider} API.\n\nâœ… Fix: Check your internet connection and try again.`);
        }
        // Re-throw our custom errors
        throw error;
    }
}

// Initialize extension
chrome.runtime.onInstalled.addListener(() => {
    console.log('Element Extractor Pro extension installed');
});

