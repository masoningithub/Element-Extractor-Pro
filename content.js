// Content script for element extraction
let selectedElements = new Set();
let isManualMode = false;
let extractionCount = 0;
let pageId = generatePageId();
let labelMode = 'original';
// TODO: Implement undo/redo in future version (v1.4.0+)
let undoStack = [];
let redoStack = [];
// TODO: Implement minimap in future version (v1.4.0+)
let minimapEnabled = false;
let minimapCanvas = null;

// Message handler
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    switch (message.action) {
        case 'AUTO_SELECT':
            autoSelectElements();
            sendResponse({ success: true });
            break;
        case 'MANUAL_MODE_ON':
            enableManualMode();
            sendResponse({ success: true });
            break;
        case 'MANUAL_MODE_OFF':
            disableManualMode();
            sendResponse({ success: true });
            break;
        case 'EXTRACT':
            extractElements(message.pageName).then(result => {
                sendResponse(result);
            });
            return true; // Keep channel open for async response
        case 'PING':
            sendResponse({ ok: true });
            break;
        case 'EXTRACT_PART':
            try {
                const elementsData = Array.from(selectedElements).map(el => extractElementInfo(el));
                sendResponse({ success: true, frameUrl: window.location.href, elements: elementsData });
            } catch (e) {
                sendResponse({ success: false, error: e?.message || 'extract part failed' });
            }
            break;
        case 'SAVE_EXTRACTION':
            try {
                const finalPageName = message.pageName || generateDefaultPageName();
                extractionCount++;
                const uniqueElements = removeDuplicateElements(message.elements || []);
                const extractionData = {
                    pageName: finalPageName,
                    pageId: pageId,
                    url: window.location.href,
                    timestamp: new Date().toISOString(),
                    domSignature: getDOMSignature(),
                    extractionGroup: extractionCount,
                    elementsCount: uniqueElements.length,
                    elements: uniqueElements
                };
                saveToStorage(pageId, extractionData).then(() => {
                    clearHighlights();
                    sendResponse({ success: true, count: uniqueElements.length });
                });
            } catch (e) {
                sendResponse({ success: false, error: e?.message || 'save failed' });
            }
            return true;
        case 'GET_SELECTED_SUMMARY':
            try {
                const items = Array.from(selectedElements).map(el => {
                    const info = extractElementInfo(el);
                    const rawSelector = buildRawSelector(el);
                    const matches = validateSelectorUniqueness(el, rawSelector);
                    return { label: info.label, type: info.type, selector: info.selector, rawSelector, matches };
                });
                const byType = {};
                items.forEach(i => { const t = (i.type||'').toLowerCase(); byType[t] = (byType[t]||0)+1; });
                sendResponse({ success: true, selectedCount: selectedElements.size, byType, items });
            } catch (e) {
                sendResponse({ success: false, error: e?.message || 'summary failed' });
            }
            break;
        case 'VALIDATE_RAW_SELECTORS_FRAME':
            try {
                const selectors = Array.isArray(message.rawSelectors) ? message.rawSelectors : [];
                const result = {};
                selectors.forEach(sel => {
                    try { result[sel] = document.querySelectorAll(sel).length; } catch { result[sel] = 0; }
                });
                sendResponse({ success: true, counts: result });
            } catch (e) {
                sendResponse({ success: false });
            }
            break;
        case 'PREVIEW_HIGHLIGHT':
            try {
                const ctx = message.contextDocument || 'document';
                const isTop = (window.top === window.self);
                if (!shouldApplyInThisFrame(ctx, isTop)) { sendResponse({ success: false, skipped: true }); break; }
                const el = findElementByRawSelector(message.rawSelector);
                if (el) {
                    el.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    el.classList.add('ee-hovered');
                    setTimeout(() => el.classList.remove('ee-hovered'), 1000);
                    sendResponse({ success: true });
                } else {
                    sendResponse({ success: false, error: 'element not found' });
                }
            } catch (e) {
                sendResponse({ success: false });
            }
            break;
        case 'GET_STATS':
            sendResponse({ selectedCount: selectedElements.size });
            break;
        case 'CLEAR_ALL':
            clearAll();
            sendResponse({ success: true });
            break;
        case 'SET_HIGHLIGHT_COLORS':
            try {
                if (message.selected) document.documentElement.style.setProperty('--ee-selected', message.selected);
                if (message.hover) document.documentElement.style.setProperty('--ee-hover', message.hover);
                sendResponse({ success: true });
            } catch (e) { sendResponse({ success: false }); }
            break;
        case 'SET_LABEL_MODE':
            labelMode = message.mode === 'enhanced' ? 'enhanced' : 'original';
            sendResponse({ success: true, mode: labelMode });
            break;
        case 'RUN_ENTRY':
            try {
                const result = runEntryActions(message.dataGroups || [], message.functions || {});
                sendResponse({ success: true, ...result });
            } catch (e) {
                sendResponse({ success: false, error: e?.message });
            }
            break;
        default:
            sendResponse({ success: false, error: 'Unknown action' });
    }
});

function runEntryActions(dataGroups, functionsJSON) {
    try { registerDynamicFunctions(functionsJSON); } catch {}
    const result = { totalActions: 0, appliedActions: 0, missingElements: 0, blockedContexts: 0, skippedFrame: 0 };
    if (!Array.isArray(dataGroups)) return result;
    const isTop = (window.top === window.self);
    const frameContext = isTop ? 'TOP_FRAME' : `IFRAME(${window.location.href})`;
    try { console.log(`ENTRY [${frameContext}]: Starting with ${dataGroups.length} groups`); } catch {}

    dataGroups.forEach(group => {
        const actions = group.Actions || [];
        actions.forEach(action => {
            if (action && action.InputValue !== undefined && action.InputValue !== null && action.InputValue !== 'null') {
                // Filter actions to the correct frame
                const ctx = action.ContextDocument || 'document';
                const shouldApply = shouldApplyInThisFrame(ctx, isTop);

                if (!shouldApply) {
                    result.skippedFrame++;
                    return;
                }

                result.totalActions++;
                try {
                    console.log(`ENTRY [${frameContext}]: Applying action to ${action.TargetElement}, type=${action.ActionType}, value=${action.InputValue}`);
                } catch {}

                const applied = handleEntryAction(action);
                if (applied === true) {
                    result.appliedActions++;
                } else if (applied === 'blocked') {
                    result.blockedContexts++;
                    try { console.warn(`ENTRY [${frameContext}]: Blocked context for ${action.TargetElement}`); } catch {}
                } else {
                    result.missingElements++;
                    try { console.warn(`ENTRY [${frameContext}]: Element not found: ${action.TargetElement}`); } catch {}
                }
            }
        });
    });
    try { console.log(`ENTRY [${frameContext}]: summary`, result); } catch {}
    return result;
}

function shouldApplyInThisFrame(ctx, isTop) {
    if (!ctx || ctx === 'document') return isTop;
    try {
        // General form: document.querySelector('...').contentWindow.document
        const m = String(ctx).match(/document\.querySelector\((['"])(.+?)\1\)\.contentWindow\.document/);
        if (!m) return !isTop; // if we cannot parse but context is not 'document', assume child
        const sel = m[2];

        // If we're in the top frame, don't apply actions meant for iframes
        if (isTop) return false;

        // We're in an iframe - check if this is the target iframe
        try {
            const fe = window.frameElement;
            if (!fe) {
                // Can't access frameElement (cross-origin) - try URL matching as fallback
                return matchFrameByUrl(sel);
            }

            // Check if selector matches this iframe element
            let match;

            // Try matching by ID: iframe#myid
            if ((match = sel.match(/^iframe#([a-zA-Z0-9_-]+)/))) {
                const id = match[1];
                if (fe.id === id) return true;
            }

            // Try matching by ID: #myid (without iframe prefix)
            if (sel.startsWith('#')) {
                const id = sel.substring(1).split(/[^\w-]/)[0]; // Extract just the ID part
                if (fe.id === id) return true;
            }

            // Try matching by name attribute: iframe[name="..."]
            if ((match = sel.match(/\[name\s*=\s*['"]([^'"]+)['"]\]/))) {
                const name = match[1];
                const feName = fe.getAttribute && fe.getAttribute('name');
                if (feName === name) return true;
            }

            // Try matching by title attribute: iframe[title="..."]
            if ((match = sel.match(/\[title\s*=\s*['"]([^'"]+)['"]\]/))) {
                const title = match[1];
                const feTitle = fe.getAttribute && fe.getAttribute('title');
                if (feTitle === title) return true;
            }

            // Try matching by class: iframe.classname
            if ((match = sel.match(/^iframe\.([a-zA-Z0-9_-]+)/))) {
                const className = match[1];
                if (fe.classList && fe.classList.contains(className)) return true;
            }

            // Try matching by nth-of-type: iframe:nth-of-type(n)
            if ((match = sel.match(/^iframe:nth-of-type\((\d+)\)/))) {
                const targetIndex = parseInt(match[1], 10) - 1; // Convert to 0-based
                try {
                    const parent = fe.parentElement;
                    if (parent) {
                        const iframes = Array.from(parent.querySelectorAll('iframe'));
                        const actualIndex = iframes.indexOf(fe);
                        if (actualIndex === targetIndex) return true;
                    }
                } catch {}
            }

            // Try matching by src attribute patterns
            const feSrc = fe.getAttribute && fe.getAttribute('src');
            if (feSrc && matchFrameBySrcAttribute(sel, feSrc)) {
                return true;
            }

            // Fallback to URL matching
            return matchFrameByUrl(sel);

        } catch (e) {
            // Cross-origin error - fall back to URL matching
            return matchFrameByUrl(sel);
        }
    } catch {}
    return !isTop;
}

function matchFrameBySrcAttribute(selector, frameSrc) {
    let match;
    // Check various src attribute patterns
    if ((match = selector.match(/src\s*\*\=\s*['"]([^'"]+)['"]/))) {
        return frameSrc.includes(match[1]);
    }
    if ((match = selector.match(/src\s*\^\=\s*['"]([^'"]+)['"]/))) {
        return frameSrc.startsWith(match[1]);
    }
    if ((match = selector.match(/src\s*\$\=\s*['"]([^'"]+)['"]/))) {
        return frameSrc.endsWith(match[1]);
    }
    if ((match = selector.match(/src\s*\=\s*['"]([^'"]+)['"]/))) {
        return frameSrc === match[1];
    }
    return false;
}

function matchFrameByUrl(selector) {
    // Try to match current frame's URL against patterns in the selector
    const currentUrl = window.location.href;
    let match;

    if ((match = selector.match(/src\s*\*\=\s*['"]([^'"]+)['"]/))) {
        return currentUrl.includes(match[1]);
    }
    if ((match = selector.match(/src\s*\^\=\s*['"]([^'"]+)['"]/))) {
        return currentUrl.startsWith(match[1]);
    }
    if ((match = selector.match(/src\s*\$\=\s*['"]([^'"]+)['"]/))) {
        return currentUrl.endsWith(match[1]);
    }
    if ((match = selector.match(/src\s*\=\s*['"]([^'"]+)['"]/))) {
        return currentUrl === match[1];
    }

    // Fallback: any quoted token in URL
    if ((match = selector.match(/['"]([^'"]+)['"]/))) {
        return currentUrl.includes(match[1]);
    }

    // If no pattern matched, assume this iframe should handle it
    // (better to try and fail than to skip entirely)
    return true;
}

function resolveLayerEntry(layer) {
    if (!layer || layer === 'document') return document;
    // Try to parse a formatted contextDocument: document.querySelector('...').contentWindow.document
    try {
        const m = String(layer).match(/document\.querySelector\('(.+?)'\)\.contentWindow\.document/);
        if (m && m[1]) {
            const sel = m[1];
            const iframe = document.querySelector(sel);
            if (!iframe) { console.warn('ENTRY: iframe not found for selector', sel); return null; }
            try {
                const doc = iframe.contentDocument || iframe.contentWindow?.document;
                if (!doc) { console.warn('ENTRY: iframe document unavailable', sel); return null; }
                return doc;
            } catch (e) {
                console.warn('ENTRY: cross-origin iframe access blocked', sel, e);
                return 'BLOCKED_IFRAME';
            }
        }
    } catch {}
    try { const v = eval(layer); return v || document; } catch { return document; }
}

function resolveElementEntry(doc, selector) {
    if (!selector) return null;
    try {
        if (selector.endsWith('()')) {
            const fname = selector.slice(0,-2);
            if (window.dynamicFunctions && typeof window.dynamicFunctions[fname] === 'function') return window.dynamicFunctions[fname]();
            if (typeof window[fname] === 'function') return window[fname]();
            return null;
        }
        return doc.querySelector(selector);
    } catch { return null; }
}

function handleEntryAction(action) {
    const ctx = action.ContextDocument || 'document';
    const doc = resolveLayerEntry(ctx);
    if (doc === 'BLOCKED_IFRAME') return 'blocked';
    const el = resolveElementEntry(doc, action.TargetElement);
    if (!el && action.ActionType !== 'Button') { console.warn('ENTRY: element not found', { ctx, sel: action.TargetElement }); return false; }
    const val = action.InputValue;
    try {
        switch (action.ActionType) {
            case 'Input': setInputValue(el, String(val)); break;
            case 'Select': setSelectValue(el, String(val)); break;
            case 'Checkbox': setCheckboxValue(el, val); break;
            case 'RadioButton': setRadioValue(doc, action.TargetElement, val); break;
            case 'Button': if (el) el.click(); break;
            default: setInputValue(el, String(val));
        }
        return true;
    } catch (e) {
        console.warn('ENTRY: error applying action', action, e);
        return false;
    }
}

function setInputValue(el, value) {
    if (!el) return; el.focus(); el.value = value;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
}
function setSelectValue(el, value) {
    if (!el) return; const opt = Array.from(el.options||[]).find(o => o.value===value || o.text===value);
    if (opt) el.value = opt.value; else el.value = value;
    el.dispatchEvent(new Event('change', { bubbles: true }));
}
function setCheckboxValue(el, value) { if (!el) return; el.checked = !!(String(value).toLowerCase() === 'true' || value === true); el.dispatchEvent(new Event('change', { bubbles: true })); }
function setRadioValue(doc, selector, value) {
    try {
        const last = selector.split('>>>').pop().trim();
        const nameMatch = last.match(/\[name\s*=\s*"([^"]+)"\]/);
        if (nameMatch) {
            const name = nameMatch[1];
            const radios = doc.querySelectorAll(`input[type="radio"][name="${name}"]`);
            radios.forEach(r => { r.checked = (r.value==value || r.id==value); r.dispatchEvent(new Event('change',{bubbles:true})); });
        } else {
            const el = resolveElementEntry(doc, selector);
            if (el) { el.checked = true; el.dispatchEvent(new Event('change',{bubbles:true})); }
        }
    } catch {}
}

function registerDynamicFunctions(functionsJSON) {
    if (!functionsJSON || typeof functionsJSON !== 'object') return;
    if (!window.dynamicFunctions) window.dynamicFunctions = {};
    Object.entries(functionsJSON).forEach(([sig, body]) => {
        try {
            const m = sig.match(/^([^(]+)\(([^)]*)\)$/);
            if (!m) return; const name = m[1]; const params = m[2].split(',').map(p=>p.trim()).filter(Boolean);
            const isAsync = body.includes('async') || body.includes('await');
            const wrapped = isAsync ? `return (async () => { ${body} })()` : body;
            window.dynamicFunctions[name] = new Function(...params, wrapped);
        } catch {}
    });
}

// Register this frame with background for multi-frame coordination
try { chrome.runtime.sendMessage({ action: 'FRAME_READY' }); } catch (e) {}

function generatePageId() {
    const url = window.location.href;
    const title = document.title;
    const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    const domSignature = getDOMSignature();
    const hash = safeBase64Encode(url + title + domSignature).slice(0, 10);
    return `${title}_${timestamp}_${hash}`.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 50);
}

function safeBase64Encode(str) {
    try {
        return btoa(encodeURIComponent(str).replace(/%([0-9A-F]{2})/g, (match, p1) => {
            return String.fromCharCode('0x' + p1);
        })).replace(/[^a-zA-Z0-9]/g, '');
    } catch (e) {
        let hash = 0;
        for (let i = 0; i < str.length; i++) {
            const char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash;
        }
        return Math.abs(hash).toString(36);
    }
}

function getDOMSignature() {
    const selectors = ['form', '[role="main"]', 'main', '#content', '.content'];
    let signature = '';
    for (const selector of selectors) {
        const element = document.querySelector(selector);
        if (element) {
            const tagName = element.tagName.toLowerCase();
            const id = element.id ? `#${element.id}` : '';
            const className = element.className ? `.${element.className.split(' ')[0]}` : '';
            signature = `${tagName}${id}${className}`;
            break;
        }
    }
    return signature || `body_${document.querySelectorAll('*').length}`;
}

function getAllInteractiveElements() {
    const selectors = [
        'input:not([type="hidden"])',
        'select',
        'button',
        'textarea',
        '[role="button"]',
        '[onclick]',
        'a[href]',
        '[contenteditable="true"]',
        '[tabindex]:not([tabindex="-1"])'
    ];

    const elements = [];
    selectors.forEach(selector => {
        const found = document.querySelectorAll(selector);
        found.forEach(element => {
            if (isElementVisible(element) && !elements.includes(element)) {
                elements.push(element);
            }
        });
    });

    return elements;
}

function isElementVisible(element) {
    const rect = element.getBoundingClientRect();
    const computedStyle = window.getComputedStyle(element);
    return rect.width > 0 &&
           rect.height > 0 &&
           computedStyle.visibility !== 'hidden' &&
           computedStyle.display !== 'none' &&
           computedStyle.opacity !== '0';
}

function autoSelectElements() {
    clearHighlights();
    const elements = getAllInteractiveElements();

    elements.forEach(element => {
        element.classList.add('ee-available');
        selectedElements.add(element);
    });

    notifyExtension('UPDATE_STATUS', {
        message: `Auto-selected ${elements.length} elements. Click "Manual Mode" to adjust.`,
        type: 'success'
    });

    notifyExtension('UPDATE_STATS');
    if (minimapEnabled) updateMinimap();

    if (elements.length > 0) {
        elements[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
}

function enableManualMode() {
    isManualMode = true;
    document.addEventListener('contextmenu', handleRightClick, true);
    document.addEventListener('mouseover', handleMouseOver, true);
    document.addEventListener('mouseout', handleMouseOut, true);
    document.addEventListener('keydown', handleKeyDown, true);

    notifyExtension('UPDATE_STATUS', {
        message: 'Manual mode: Right-click elements to add/remove. Hover to see info.',
        type: 'info'
    });
}

function disableManualMode() {
    isManualMode = false;
    document.removeEventListener('contextmenu', handleRightClick, true);
    document.removeEventListener('mouseover', handleMouseOver, true);
    document.removeEventListener('mouseout', handleMouseOut, true);
    document.removeEventListener('keydown', handleKeyDown, true);
    clearHoverEffects();

    notifyExtension('UPDATE_STATUS', {
        message: `Manual mode disabled. Selected ${selectedElements.size} elements.`,
        type: 'info'
    });
}

function handleMouseOver(event) {
    if (!isManualMode) return;
    const element = event.target;
    clearHoverEffects();
    element.classList.add('ee-hovered');
    createTooltip(element, event);
}

function handleMouseOut(event) {
    if (!isManualMode) return;
    setTimeout(() => {
        if (!document.querySelector(':hover.ee-hovered')) {
            clearHoverEffects();
        }
    }, 100);
}

function createTooltip(element, event) {
    const existingTooltip = document.querySelector('.ee-tooltip');
    if (existingTooltip) existingTooltip.remove();

    const tooltip = document.createElement('div');
    tooltip.className = 'ee-tooltip';

    const selector = generateRobustSelector(element);
    const label = getEnhancedLabel(element);
    const elementType = getElementType(element);
    const isSelected = selectedElements.has(element);

    tooltip.innerHTML = `
        <div style="color: ${isSelected ? '#ff6b6b' : '#4CAF50'}; font-weight: bold;">
            ${isSelected ? 'âœ“ SELECTED' : '+ AVAILABLE'}
        </div>
        <div><strong>Label:</strong> ${label || 'No label'}</div>
        <div><strong>Type:</strong> ${elementType}</div>
        <div><strong>Selector:</strong> ${selector}</div>
        <div style="font-size: 10px; color: #ccc; margin-top: 4px;">
            Right-click to ${isSelected ? 'remove' : 'add'}
        </div>
    `;

    const rect = element.getBoundingClientRect();
    tooltip.style.left = Math.min(rect.right + 10, window.innerWidth - 360) + 'px';
    tooltip.style.top = Math.max(rect.top, 10) + 'px';

    document.body.appendChild(tooltip);
}

function clearHoverEffects() {
    document.querySelectorAll('.ee-hovered').forEach(el => {
        el.classList.remove('ee-hovered');
    });
    const tooltip = document.querySelector('.ee-tooltip');
    if (tooltip) tooltip.remove();
}

function performUndo() {
    const last = undoStack.pop();
    if (!last) return;
    if (last.action === 'add') {
        if (selectedElements.has(last.element)) {
            selectedElements.delete(last.element);
            last.element.classList.remove('ee-selected', 'ee-available');
            removeNumberBadge(last.element);
        }
        redoStack.push({ action: 'remove', element: last.element });
    } else if (last.action === 'remove') {
        selectedElements.add(last.element);
        last.element.classList.add('ee-selected');
        addNumberBadge(last.element);
        redoStack.push({ action: 'add', element: last.element });
    }
    notifyExtension('UPDATE_STATS');
    if (minimapEnabled) updateMinimap();
}

function performRedo() {
    const next = redoStack.pop();
    if (!next) return;
    if (next.action === 'add') {
        selectedElements.add(next.element);
        next.element.classList.add('ee-selected');
        addNumberBadge(next.element);
        undoStack.push({ action: 'add', element: next.element });
    } else if (next.action === 'remove') {
        if (selectedElements.has(next.element)) {
            selectedElements.delete(next.element);
            next.element.classList.remove('ee-selected', 'ee-available');
            removeNumberBadge(next.element);
        }
        undoStack.push({ action: 'remove', element: next.element });
    }
    notifyExtension('UPDATE_STATS');
    if (minimapEnabled) updateMinimap();
}

function addNumberBadge(el) {
    if (!el) return;
    const existing = el.querySelector('.ee-number-badge');
    if (existing) return;
    const badge = document.createElement('span');
    badge.className = 'ee-number-badge';
    badge.textContent = String(selectedElements.size);
    if (window.getComputedStyle(el).position === 'static') {
        el.style.position = 'relative';
    }
    el.appendChild(badge);
}

function removeNumberBadge(el) {
    const b = el && el.querySelector('.ee-number-badge');
    if (b) b.remove();
}

function handleRightClick(event) {
    if (!isManualMode) return;

    event.preventDefault();
    event.stopPropagation();

    const element = event.target;
    let targetElement = element;

    const selectedChildren = Array.from(targetElement.querySelectorAll('*')).filter(child =>
        selectedElements.has(child)
    );

    if (!selectedElements.has(targetElement) && selectedChildren.length === 1) {
        targetElement = selectedChildren[0];
    }

    if (selectedElements.has(targetElement)) {
        selectedElements.delete(targetElement);
        targetElement.classList.remove('ee-selected', 'ee-available');
        undoStack.push({ action: 'remove', element: targetElement });
        redoStack = [];
        notifyExtension('UPDATE_STATUS', {
            message: `Removed: ${getElementDescription(targetElement)}`,
            type: 'info'
        });
    } else {
        selectedElements.add(targetElement);
        targetElement.classList.add('ee-selected');
        addNumberBadge(targetElement);
        undoStack.push({ action: 'add', element: targetElement });
        redoStack = [];
        notifyExtension('UPDATE_STATUS', {
            message: `Added: ${getElementDescription(targetElement)}`,
            type: 'success'
        });
    }

    notifyExtension('UPDATE_STATS');
    clearHoverEffects();
    createTooltip(targetElement, event);
    if (minimapEnabled) updateMinimap();
}

function handleKeyDown(event) {
    if (event.key === 'Escape') {
        disableManualMode();
    } else if (event.ctrlKey && event.key.toLowerCase() === 'z') {
        performUndo();
    } else if (event.ctrlKey && event.key.toLowerCase() === 'y') {
        performRedo();
    } else if (event.ctrlKey && !event.shiftKey && event.key.toLowerCase() === 'e') {
        autoSelectElements();
    } else if (event.ctrlKey && !event.shiftKey && event.key.toLowerCase() === 'm') {
        if (isManualMode) disableManualMode(); else enableManualMode();
    } else if (event.ctrlKey && event.shiftKey && event.key.toLowerCase() === 'e') {
        try { chrome.runtime.sendMessage({ action: 'EXTRACT_ALL' }); } catch {}
    } else if (event.ctrlKey && event.shiftKey && event.key.toLowerCase() === 'd') {
        try { chrome.runtime.sendMessage({ action: 'REQUEST_DOWNLOAD' }); } catch {}
    }
}

function getElementDescription(element) {
    const tagName = element.tagName.toLowerCase();
    const id = element.id ? `#${element.id}` : '';
    const className = element.className ? `.${element.className.split(' ')[0]}` : '';
    const text = element.textContent ? element.textContent.slice(0, 20) : '';
    return `${tagName}${id}${className} "${text}"`;
}

async function extractElements(pageName) {
    if (selectedElements.size === 0) {
        notifyExtension('UPDATE_STATUS', {
            message: 'No elements selected! Run Auto Select first.',
            type: 'error'
        });
        return { success: false, count: 0 };
    }

    const finalPageName = pageName || generateDefaultPageName();
    extractionCount++;

    const elementsData = Array.from(selectedElements).map(element => {
        return extractElementInfo(element);
    });

    const uniqueElements = removeDuplicateElements(elementsData);

    const extractionData = {
        pageName: finalPageName,
        pageId: pageId,
        url: window.location.href,
        timestamp: new Date().toISOString(),
        domSignature: getDOMSignature(),
        extractionGroup: extractionCount,
        elementsCount: uniqueElements.length,
        elements: uniqueElements
    };

    await saveToStorage(pageId, extractionData);
    clearHighlights();

    return { success: true, count: uniqueElements.length };
}

function generateDefaultPageName() {
    const title = document.title.slice(0, 30);
    const url = new URL(window.location.href);
    const path = url.pathname.split('/').filter(p => p).slice(-1)[0] || 'home';
    return `${title}_${path}`.replace(/[^a-zA-Z0-9_-]/g, '_');
}

function extractElementInfo(element) {
    const label = (labelMode === 'enhanced') ? getLabelEnhanced(element) : getEnhancedLabel(element);
    const fullSelector = generateRobustSelector(element);
    const elementType = getElementType(element);
    const htmlString = element.outerHTML;

    // Parse iframe selector: split "iframe[src*='...'] >>> #element" into separate fields
    const { contextDocument, targetSelector } = parseIframeSelector(fullSelector);

    return {
        label: label.trim(),
        selector: targetSelector,  // Element selector without iframe prefix
        contextDocument: contextDocument,  // Iframe selector or 'document' for main frame
        type: elementType,
        html: htmlString,
        position: getElementPosition(element),
        attributes: getKeyAttributes(element),
        validation: getValidationInfo(element),
        accessibility: getAccessibilityInfo(element),
        frame: getFrameContext()
    };
}

function parseIframeSelector(fullSelector) {
    // Check if selector contains iframe separator ">>>"
    if (fullSelector.includes(' >>> ')) {
        const parts = fullSelector.split(' >>> ');
        const iframeSelector = parts[0].trim();
        return {
            contextDocument: formatContextDocument(iframeSelector),
            targetSelector: parts[1].trim()
        };
    }
    // No iframe, element is in main document
    return {
        contextDocument: 'document',
        targetSelector: fullSelector
    };
}

function formatContextDocument(iframeSelector) {
    // Convert "iframe[src*='...']" to "document.querySelector('iframe[src*=\'...\']').contentWindow.document"
    if (!iframeSelector || iframeSelector === 'document') {
        return 'document';
    }
    // Escape single quotes in the selector for the JavaScript string
    const escapedSelector = iframeSelector.replace(/'/g, "\\'");
    return `document.querySelector('${escapedSelector}').contentWindow.document`;
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

function getEnhancedLabel(element) {
    let label = '';

    if (element.id) {
        const associatedLabel = document.querySelector(`label[for="${element.id}"]`);
        if (associatedLabel) label = associatedLabel.textContent.trim();
    }

    if (!label) {
        const parentLabel = element.closest('label');
        if (parentLabel) {
            const labelClone = parentLabel.cloneNode(true);
            const inputInClone = labelClone.querySelector('input, select, button, textarea');
            if (inputInClone) inputInClone.remove();
            label = labelClone.textContent.trim();
        }
    }

    if (!label) {
        label = element.getAttribute('aria-label') || '';
        if (!label && element.getAttribute('aria-labelledby')) {
            const labelledBy = document.getElementById(element.getAttribute('aria-labelledby'));
            if (labelledBy) label = labelledBy.textContent.trim();
        }
    }

    if (!label) label = element.getAttribute('placeholder') || '';
    if (!label) label = element.getAttribute('title') || '';

    if (!label) {
        let current = element.previousElementSibling;
        while (current && !label) {
            if (current.textContent && current.textContent.trim() &&
                !current.querySelector('input, select, button, textarea')) {
                label = current.textContent.trim();
                break;
            }
            current = current.previousElementSibling;
        }
    }

    if (!label) label = getTableHeaderLabel(element);

    if (!label && (element.tagName === 'BUTTON' || element.tagName === 'A')) {
        label = element.textContent.trim();
    }

    return (label && label.trim()) || `${element.tagName.toLowerCase()}_${element.type || 'unknown'}`;
}

function getLabelEnhanced(element) {
    // Enhanced: original strategy plus fieldset legend and nearby text before fallback
    let label = '';
    // for=[id]
    if (element.id) {
        const associatedLabel = document.querySelector(`label[for="${element.id}"]`);
        if (associatedLabel) label = associatedLabel.textContent.trim();
    }
    // wrapping label
    if (!label) {
        const parentLabel = element.closest('label');
        if (parentLabel) {
            const labelClone = parentLabel.cloneNode(true);
            const inputInClone = labelClone.querySelector('input, select, button, textarea');
            if (inputInClone) inputInClone.remove();
            label = labelClone.textContent.trim();
        }
    }
    // aria
    if (!label) {
        label = element.getAttribute('aria-label') || '';
        if (!label && element.getAttribute('aria-labelledby')) {
            const labelledBy = document.getElementById(element.getAttribute('aria-labelledby'));
            if (labelledBy) label = labelledBy.textContent.trim();
        }
    }
    // placeholders/titles
    if (!label) label = element.getAttribute('placeholder') || '';
    if (!label) label = element.getAttribute('title') || '';
    // previous sibling text
    if (!label) {
        let current = element.previousElementSibling;
        while (current && !label) {
            if (current.textContent && current.textContent.trim() &&
                !current.querySelector('input, select, button, textarea')) {
                label = current.textContent.trim();
                break;
            }
            current = current.previousElementSibling;
        }
    }
    // table header
    if (!label) label = getTableHeaderLabel(element);
    // enhanced extras
    if (!label) label = getFieldsetLegendLabel(element);
    if (!label) label = getNearbyTextLabel(element);
    // buttons/links
    if (!label && (element.tagName === 'BUTTON' || element.tagName === 'A')) {
        label = element.textContent.trim();
    }
    return (label && label.trim()) || `${element.tagName.toLowerCase()}_${element.type || 'unknown'}`;
}

function getTableHeaderLabel(element) {
    const td = element.closest('td');
    if (!td) return '';
    const tr = td.parentElement;
    if (!tr) return '';
    const cellIndex = Array.from(tr.children).indexOf(td);
    const table = tr.closest('table');
    if (!table) return '';
    const thead = table.querySelector('thead');
    if (thead) {
        const headerRows = thead.querySelectorAll('tr');
        for (const headerRow of headerRows) {
            const headerCell = headerRow.children[cellIndex];
            if (headerCell) return headerCell.textContent.trim();
        }
    }
    return '';
}

function getFieldsetLegendLabel(element) {
    const fieldset = element.closest('fieldset');
    if (!fieldset) return '';
    const legend = fieldset.querySelector('legend');
    return legend ? legend.textContent.trim() : '';
}

function getNearbyTextLabel(element) {
    // Look for common label containers near the element
    const container = element.closest('[class*="label"], [class*="field"], [class*="form"]');
    let text = '';
    if (container) {
        // prefer text nodes not containing input/select/textarea
        const clone = container.cloneNode(true);
        const kill = clone.querySelectorAll('input,select,textarea,button');
        kill.forEach(n => n.remove());
        text = clone.textContent.trim();
        if (text) return text;
    }
    // Check previous siblings up to 2 steps
    let prev = element.previousElementSibling;
    for (let i=0; i<2 && prev && !text; i++) {
        if (prev && prev.textContent && prev.textContent.trim()) {
            const hasInteractive = prev.querySelector('input,select,textarea,button');
            if (!hasInteractive) { text = prev.textContent.trim(); break; }
        }
        prev = prev.previousElementSibling;
    }
    return text;
}

function generateRobustSelector(element) {
    if (element.id && document.querySelectorAll(`#${element.id}`).length === 1) {
        return prefixWithFrame(`#${element.id}`);
    }

    if (element.name && element.name.trim()) {
        const nameSelector = `[name="${element.name}"]`;
        if (document.querySelectorAll(nameSelector).length === 1) {
            return prefixWithFrame(nameSelector);
        }
    }

    const tagName = element.tagName.toLowerCase();
    let selector = tagName;

    if (element.type) selector += `[type="${element.type}"]`;

    if (element.className) {
        const classes = element.className.split(' ').filter(c => c.trim());
        for (const cls of classes) {
            const classSelector = `${selector}.${cls}`;
            if (document.querySelectorAll(classSelector).length <= 3) {
                selector = classSelector;
                break;
            }
        }
    }

    if (document.querySelectorAll(selector).length > 1) {
        const parent = element.parentElement;
        if (parent && parent !== document.body) {
            const parentSelector = generateParentContext(parent);
            selector = `${parentSelector} > ${selector}`;
        }
    }

    if (document.querySelectorAll(selector).length > 1) {
        const parent = element.parentElement;
        if (parent) {
            const siblings = Array.from(parent.children).filter(child =>
                child.tagName === element.tagName && child.type === element.type);
            if (siblings.length > 1) {
                const index = siblings.indexOf(element) + 1;
                selector += `:nth-child(${index})`;
            }
        }
    }

    return prefixWithFrame(selector);
}

function generateParentContext(parent) {
    if (parent.id) return `#${parent.id}`;
    if (parent.className) {
        const firstClass = parent.className.split(' ')[0];
        return `.${firstClass}`;
    }
    return parent.tagName.toLowerCase();
}

function getElementType(element) {
    const tagName = element.tagName.toLowerCase();
    const type = element.type ? `[${element.type}]` : '';
    return `${tagName}${type}`;
}

function getElementPosition(element) {
    const rect = element.getBoundingClientRect();
    return {
        x: Math.round(rect.left + window.scrollX),
        y: Math.round(rect.top + window.scrollY),
        width: Math.round(rect.width),
        height: Math.round(rect.height)
    };
}

function getFrameContext() {
    const inFrame = (window.top !== window.self);
    const prefix = getFrameSelectorPrefix();
    return {
        inFrame,
        url: window.location.href,
        selectorPrefix: prefix
    };
}

function getFrameSelectorPrefix() {
    if (window.top === window.self) return '';
    try {
        const fe = window.frameElement;
        if (fe) {
            // Priority 1: ID (most specific and reliable)
            if (fe.id && fe.id.trim()) {
                return `iframe#${fe.id}`;
            }

            // Priority 2: Name attribute
            const name = fe.getAttribute && fe.getAttribute('name');
            if (name && name.trim()) {
                return `iframe[name="${name}"]`;
            }

            // Priority 3: Title attribute (sometimes used for accessibility)
            const title = fe.getAttribute && fe.getAttribute('title');
            if (title && title.trim()) {
                return `iframe[title="${title}"]`;
            }

            // Priority 4: Class (if it has a unique or descriptive class)
            if (fe.className && fe.className.trim()) {
                const classes = fe.className.trim().split(/\s+/);
                if (classes.length > 0) {
                    // Use first class as identifier
                    return `iframe.${classes[0]}`;
                }
            }

            // Priority 5: Src attribute patterns
            const src = fe.getAttribute && fe.getAttribute('src');
            if (src) {
                try {
                    const u = new URL(src, document.baseURI);
                    // Use hostname + path for more specificity
                    const path = u.pathname.split('/').filter(p => p).slice(-1)[0] || '';
                    if (path) {
                        return `iframe[src*="${u.hostname}"][src*="${path}"]`;
                    }
                    return `iframe[src*="${u.hostname}"]`;
                } catch {
                    // Fallback to partial src match
                    const shortSrc = src.length > 30 ? src.substring(0, 30) : src;
                    return `iframe[src*="${shortSrc}"]`;
                }
            }

            // Priority 6: Use index among siblings as last resort
            try {
                const parent = fe.parentElement;
                if (parent) {
                    const iframes = Array.from(parent.querySelectorAll('iframe'));
                    const index = iframes.indexOf(fe);
                    if (index >= 0) {
                        return `iframe:nth-of-type(${index + 1})`;
                    }
                }
            } catch {}

            // Absolute fallback
            return 'iframe';
        }
    } catch (e) {
        // Cross-origin: try to use current URL as identifier
        try {
            const currentUrl = new URL(window.location.href);
            return `iframe[src*="${currentUrl.hostname}"]`;
        } catch {}
    }
    return 'iframe';
}

function prefixWithFrame(sel) {
    const prefix = getFrameSelectorPrefix();
    return prefix ? `${prefix} >>> ${sel}` : sel;
}

function getKeyAttributes(element) {
    const attributes = {};
    const keyAttrs = ['id', 'name', 'class', 'placeholder', 'value', 'maxlength', 'required', 'readonly', 'disabled'];
    keyAttrs.forEach(attr => {
        if (element.hasAttribute(attr)) {
            attributes[attr] = element.getAttribute(attr);
        }
    });
    return attributes;
}

function getValidationInfo(element) {
    const validation = {};
    if (element.required) validation.required = true;
    if (element.pattern) validation.pattern = element.pattern;
    if (element.minLength) validation.minLength = element.minLength;
    if (element.maxLength) validation.maxLength = element.maxLength;
    if (element.min) validation.min = element.min;
    if (element.max) validation.max = element.max;
    if (element.step) validation.step = element.step;
    return validation;
}

function getAccessibilityInfo(element) {
    const accessibility = {};
    if (element.getAttribute('aria-label')) {
        accessibility.ariaLabel = element.getAttribute('aria-label');
    }
    if (element.getAttribute('aria-describedby')) {
        accessibility.ariaDescribedby = element.getAttribute('aria-describedby');
    }
    if (element.getAttribute('role')) {
        accessibility.role = element.getAttribute('role');
    }
    if (element.tabIndex !== undefined) {
        accessibility.tabIndex = element.tabIndex;
    }
    return accessibility;
}

function removeDuplicateElements(elements) {
    const seen = new Set();
    return elements.filter(element => {
        const key = `${element.selector}_${element.label}_${element.type}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });
}

async function saveToStorage(pageId, data) {
    return new Promise((resolve) => {
        chrome.storage.local.get('extractedData', (result) => {
            let extractedData = result.extractedData || {};
            let pageData = extractedData[pageId];

            if (pageData) {
                pageData.extractions = pageData.extractions || [];
                const existingIndex = pageData.extractions.findIndex(
                    ext => ext.groupId === data.extractionGroup
                );

                if (existingIndex !== -1) {
                    pageData.extractions[existingIndex] = {
                        groupId: data.extractionGroup,
                        timestamp: data.timestamp,
                        elements: data.elements
                    };
                } else {
                    pageData.extractions.push({
                        groupId: data.extractionGroup,
                        timestamp: data.timestamp,
                        elements: data.elements
                    });
                }
                pageData.lastUpdated = data.timestamp;
            } else {
                pageData = {
                    pageId: pageId,
                    pageName: data.pageName,
                    url: data.url,
                    domSignature: data.domSignature,
                    created: data.timestamp,
                    lastUpdated: data.timestamp,
                    extractions: [{
                        groupId: data.extractionGroup,
                        timestamp: data.timestamp,
                        elements: data.elements
                    }]
                };
            }

            extractedData[pageId] = pageData;
            chrome.storage.local.set({ extractedData }, resolve);
        });
    });
}

function clearHighlights() {
    document.querySelectorAll('.ee-selected, .ee-available').forEach(element => {
        element.classList.remove('ee-selected', 'ee-available');
    });
    selectedElements.clear();
    if (minimapEnabled) updateMinimap();
}

function clearAll() {
    // Disable manual mode if active
    if (isManualMode) {
        disableManualMode();
    }

    // Clear all highlights and selections
    clearHighlights();
    clearHoverEffects();

    // Reset extraction count
    extractionCount = 0;

    // Regenerate page ID for fresh start
    pageId = generatePageId();

    notifyExtension('UPDATE_STATUS', {
        message: 'All cleared - Ready to start fresh',
        type: 'success'
    });
    if (minimapEnabled) updateMinimap();
}

function notifyExtension(action, data = {}) {
    chrome.runtime.sendMessage({ action, ...data });
}

// Cleanup on unload
window.addEventListener('beforeunload', () => {
    clearHighlights();
    clearHoverEffects();
    disableManualMode();
});

console.log('âœ… Element Extractor Pro content script loaded');

// Minimap helpers
function ensureMinimap() {
    if (minimapCanvas) return;
    const c = document.createElement('canvas');
    c.style.position = 'fixed';
    c.style.right = '10px';
    c.style.bottom = '10px';
    c.style.width = '140px';
    c.style.height = '100px';
    c.width = 140; c.height = 100;
    c.style.zIndex = '2147483647';
    c.style.background = 'rgba(0,0,0,0.5)';
    c.style.borderRadius = '6px';
    document.body.appendChild(c);
    minimapCanvas = c;
    updateMinimap();
}

function removeMinimap() {
    if (minimapCanvas) { minimapCanvas.remove(); minimapCanvas = null; }
}

function updateMinimap() {
    if (!minimapCanvas) ensureMinimap();
    const ctx = minimapCanvas.getContext('2d');
    const w = minimapCanvas.width, h = minimapCanvas.height;
    ctx.clearRect(0,0,w,h);
    ctx.fillStyle = 'rgba(255,255,255,0.1)';
    ctx.fillRect(0,0,w,h);
    const docW = Math.max(document.documentElement.scrollWidth, window.innerWidth);
    const docH = Math.max(document.documentElement.scrollHeight, window.innerHeight);
    ctx.fillStyle = '#4CAF50';
    selectedElements.forEach(el => {
        const r = el.getBoundingClientRect();
        const absX = r.left + window.scrollX;
        const absY = r.top + window.scrollY;
        const x = Math.floor((absX / docW) * w);
        const y = Math.floor((absY / docH) * h);
        ctx.beginPath(); ctx.arc(x, y, 2, 0, Math.PI*2); ctx.fill();
    });
}
function buildRawSelector(element) {
    // Similar to generateRobustSelector but without frame prefix
    if (element.id && document.querySelectorAll(`#${element.id}`).length === 1) {
        return `#${element.id}`;
    }
    if (element.name && element.name.trim()) {
        const nameSelector = `[name="${element.name}"]`;
        if (document.querySelectorAll(nameSelector).length === 1) {
            return nameSelector;
        }
    }
    const tagName = element.tagName.toLowerCase();
    let selector = tagName;
    if (element.type) selector += `[type="${element.type}"]`;
    if (element.className) {
        const classes = element.className.split(' ').filter(c => c.trim());
        for (const cls of classes) {
            const classSelector = `${selector}.${cls}`;
            if (document.querySelectorAll(classSelector).length <= 3) {
                selector = classSelector;
                break;
            }
        }
    }
    if (document.querySelectorAll(selector).length > 1) {
        const parent = element.parentElement;
        if (parent && parent !== document.body) {
            const parentSelector = generateParentContext(parent);
            selector = `${parentSelector} > ${selector}`;
        }
    }
    if (document.querySelectorAll(selector).length > 1) {
        const parent = element.parentElement;
        if (parent) {
            const siblings = Array.from(parent.children).filter(child =>
                child.tagName === element.tagName && child.type === element.type);
            if (siblings.length > 1) {
                const index = siblings.indexOf(element) + 1;
                selector += `:nth-child(${index})`;
            }
        }
    }
    return selector;
}

function validateSelectorUniqueness(element, rawSelector) {
    try { return document.querySelectorAll(rawSelector).length; } catch { return 0; }
}

function findElementByRawSelector(raw) {
    try { return document.querySelector(raw); } catch { return null; }
}
