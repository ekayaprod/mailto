/**
 * MailTo Link Generator - Standalone
 * Version: 3.0.0
 * 
 * A self-contained mailto link generator with template library management.
 * Supports .msg, .eml, and .oft file imports, folder organization, and CSV import/export.
 */

'use strict';

/* =============================================================================
   CONFIGURATION
   ============================================================================= */

const CONFIG = {
    STORAGE_KEY: 'mailto_generator_v3',
    VERSION: '3.0.0',
    CSV_HEADERS: ['name', 'path', 'to', 'cc', 'bcc', 'subject', 'body'],
    MAILTO_PARAMS: ['cc', 'bcc', 'subject']
};

/* =============================================================================
   UTILITY FUNCTIONS
   ============================================================================= */

const Utils = {
    /**
     * Escapes HTML special characters to prevent XSS
     */
    escapeHTML: (str) => {
        const div = document.createElement('div');
        div.textContent = str ?? '';
        return div.innerHTML;
    },

    /**
     * Generates a unique ID using crypto.randomUUID or fallback
     */
    generateId: () => {
        if (crypto?.randomUUID) return crypto.randomUUID();
        return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
    },

    /**
     * Debounces a function call
     */
    debounce: (func, delay) => {
        let timeout;
        return (...args) => {
            clearTimeout(timeout);
            timeout = setTimeout(() => func(...args), delay);
        };
    },

    /**
     * Copies text to clipboard
     */
    copyToClipboard: async (text) => {
        try {
            await navigator.clipboard.writeText(text);
            return true;
        } catch (err) {
            console.error('Failed to copy:', err);
            return false;
        }
    },

    /**
     * Triggers a file download
     */
    downloadFile: (content, filename, mimeType) => {
        const blob = new Blob([content], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    },

    /**
     * Opens file picker dialog
     */
    openFilePicker: (callback, accept = '*') => {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = accept;
        input.onchange = (e) => {
            if (e.target.files[0]) callback(e.target.files[0]);
        };
        input.click();
    },

    /**
     * Reads a file as text
     */
    readTextFile: (file) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = () => reject(new Error('Failed to read file'));
            reader.readAsText(file);
        });
    },

    /**
     * Parses CSV line handling quoted values
     */
    parseCSVLine: (line) => {
        const values = [];
        let currentVal = '';
        let inQuotes = false;

        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (inQuotes) {
                if (char === '"' && line[i + 1] === '"') {
                    currentVal += '"';
                    i++;
                } else if (char === '"') {
                    inQuotes = false;
                } else {
                    currentVal += char;
                }
            } else {
                if (char === '"') inQuotes = true;
                else if (char === ',') {
                    values.push(currentVal);
                    currentVal = '';
                } else {
                    currentVal += char;
                }
            }
        }
        values.push(currentVal);
        return values;
    },

    /**
     * Converts data to CSV format
     */
    toCSV: (data, headers) => {
        const escapeCell = (cell) => {
            const str = String(cell ?? '');
            if (str.includes('"') || str.includes(',') || str.includes('\n')) {
                return `"${str.replace(/"/g, '""')}"`;
            }
            return str;
        };

        const rows = data.map(obj => 
            headers.map(h => escapeCell(obj[h])).join(',')
        );
        return [headers.join(','), ...rows].join('\n');
    },

    /**
     * Parses CSV file content
     */
    parseCSV: (text, requiredHeaders) => {
        const lines = text.split('\n').filter(l => l.trim());
        if (lines.length === 0) return { data: [], errors: [] };

        const headers = Utils.parseCSVLine(lines[0]).map(h => h.trim());
        const errors = [];

        // Validate headers
        for (const reqHeader of requiredHeaders) {
            if (!headers.includes(reqHeader)) {
                errors.push(`Missing required header: "${reqHeader}"`);
            }
        }
        if (errors.length > 0) return { data: [], errors };

        // Parse rows
        const data = lines.slice(1).map((line, idx) => {
            const values = Utils.parseCSVLine(line);
            const obj = {};
            headers.forEach((header, i) => {
                if (requiredHeaders.includes(header)) {
                    obj[header] = values[i] || '';
                }
            });
            return obj;
        });

        return { data, errors };
    }
};

/* =============================================================================
   UI COMPONENTS
   ============================================================================= */

const UI = {
    /**
     * Shows a modal dialog
     */
    showModal: (title, content, buttons = []) => {
        const overlay = document.getElementById('modal-overlay');
        const body = document.getElementById('modal-body');
        
        body.innerHTML = `
            <h3>${Utils.escapeHTML(title)}</h3>
            <div>${content}</div>
            <div class="modal-actions"></div>
        `;

        const actions = body.querySelector('.modal-actions');
        buttons.forEach(btn => {
            const button = document.createElement('button');
            button.className = btn.class || 'btn-secondary';
            button.textContent = btn.label;
            button.onclick = () => {
                if (!btn.callback || btn.callback() !== false) {
                    UI.hideModal();
                }
            };
            actions.appendChild(button);
        });

        overlay.classList.add('show');
    },

    /**
     * Hides the modal dialog
     */
    hideModal: () => {
        document.getElementById('modal-overlay').classList.remove('show');
    },

    /**
     * Shows a toast notification
     */
    showToast: (() => {
        let timeout;
        return (message) => {
            const toast = document.getElementById('toast');
            clearTimeout(timeout);
            toast.textContent = message;
            toast.classList.add('show');
            timeout = setTimeout(() => toast.classList.remove('show'), 3000);
        };
    })(),

    /**
     * Renders the file tree list
     */
    renderList: (container, items, emptyMessage, createItemFn) => {
        container.innerHTML = '';
        
        if (!items || items.length === 0) {
            container.innerHTML = `<div class="empty-state-message">${emptyMessage}</div>`;
            return;
        }

        const fragment = document.createDocumentFragment();
        items.forEach(item => {
            const element = createItemFn(item);
            if (element) fragment.appendChild(element);
        });
        container.appendChild(fragment);
    }
};

/* =============================================================================
   SVG ICONS
   ============================================================================= */

const Icons = {
    folder: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path d="M.54 3.87.5 3.5a.5.5 0 0 1 .5-.5h5a.5.5 0 0 1 .5.5v.07L6.2 7H1.12zM0 4.25a.5.5 0 0 1 .5-.5h6.19l.74 1.85a.5.5 0 0 1 .44.25h4.13a.5.5 0 0 1 .5.5v.5a.5.5 0 0 1-.5.5H.5a.5.5 0 0 1-.5-.5zM.5 7a.5.5 0 0 0-.5.5v5a.5.5 0 0 0 .5.5h15a.5.5 0 0 0 .5-.5v-5a.5.5 0 0 0-.5-.5z"/></svg>',
    template: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path d="M0 4a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V4Zm2-1a1 1 0 0 0-1 1v.217l7 4.2 7-4.2V4a1 1 0 0 0-1-1H2Zm13 2.383-4.708 2.825L15 11.105V5.383Zm-.034 6.876-5.64-3.471L8 9.583l-1.326-.795-5.64 3.47A1 1 0 0 0 2 13h12a1 1 0 0 0 .966-.741ZM1 11.105l4.708-2.897L1 5.383v5.722Z"/></svg>',
    trash: '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16"><path d="M5.5 5.5A.5.5 0 0 1 6 6v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5Zm2.5 0a.5.5 0 0 1 .5.5v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5Zm3 .5a.5.5 0 0 0-1 0v6a.5.5 0 0 0 1 0V6Z"/><path d="M14.5 3a1 1 0 0 1-1 1H13v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V4h-.5a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1H6a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1h3.5a1 1 0 0 1 1 1v1ZM4.118 4 4 4.059V13a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V4.059L11.882 4H4.118ZM2.5 3h11V2h-11v1Z"/></svg>',
    move: '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16"><path fill-rule="evenodd" d="M15 2a1 1 0 0 0-1-1H2a1 1 0 0 0-1 1v12a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V2zM0 2a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V2zm5.854 8.854a.5.5 0 1 0-.708-.708L4 11.293V1.5a.5.5 0 0 0-1 0v9.793l-1.146-1.147a.5.5 0 0 0-.708.708l2 2a.5.5 0 0 0 .708 0l2-2z"/></svg>',
    edit: '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16"><path d="M12.146.146a.5.5 0 0 1 .708 0l3 3a.5.5 0 0 1 0 .708l-10 10a.5.5 0 0 1-.168.11l-5 2a.5.5 0 0 1-.65-.65l2-5a.5.5 0 0 1 .11-.168l10-10zM11.207 2.5 13.5 4.793 14.793 3.5 12.5 1.207 11.207 2.5zm1.586 3L10.5 3.207 4 9.707V10h.5a.5.5 0 0 1 .5.5v.5h.5a.5.5 0 0 1 .5.5v.5h.293l6.5-6.5zm-9.761 5.175-.106.106-1.528 3.821 3.821-1.528.106-.106A.5.5 0 0 1 5 12.5V12h-.5a.5.5 0 0 1-.5-.5V11h-.5a.5.5 0 0 1-.468-.325z"/></svg>'
};

/* =============================================================================
   STATE MANAGEMENT
   ============================================================================= */

const State = {
    data: null,
    currentFolderId: 'root',

    /**
     * Loads state from localStorage
     */
    load: () => {
        try {
            const stored = localStorage.getItem(CONFIG.STORAGE_KEY);
            if (!stored) {
                State.data = { library: [], version: CONFIG.VERSION };
                return;
            }

            const parsed = JSON.parse(stored);
            // Handle legacy array format
            if (Array.isArray(parsed)) {
                State.data = { library: parsed, version: CONFIG.VERSION };
            } else {
                State.data = parsed;
            }
        } catch (err) {
            console.error('Failed to load state:', err);
            State.data = { library: [], version: CONFIG.VERSION };
        }
    },

    /**
     * Saves state to localStorage
     */
    save: () => {
        try {
            State.data.version = CONFIG.VERSION;
            localStorage.setItem(CONFIG.STORAGE_KEY, JSON.stringify(State.data));
        } catch (err) {
            console.error('Failed to save state:', err);
            UI.showToast('Failed to save changes');
        }
    },

    /**
     * Recursively finds an item in the tree
     */
    findItem: (id, items = State.data.library, parent = null) => {
        if (id === 'root') {
            return { id: 'root', name: 'Root', type: 'folder', children: State.data.library };
        }

        for (const item of items) {
            if (item.id === id) return { item, parent };
            if (item.type === 'folder' && item.children) {
                const result = State.findItem(id, item.children, item);
                if (result) return result;
            }
        }
        return null;
    },

    /**
     * Gets all folders for dropdown population
     */
    getAllFolders: (items = State.data.library, level = 0) => {
        let folders = [];
        if (level === 0) folders.push({ id: 'root', name: 'Root', level: 0 });

        items.forEach(item => {
            if (item.type === 'folder') {
                folders.push({ id: item.id, name: item.name, level: level + 1 });
                if (item.children) {
                    folders = folders.concat(State.getAllFolders(item.children, level + 1));
                }
            }
        });
        return folders;
    },

    /**
     * Gets breadcrumb path for current folder
     */
    getBreadcrumb: (folderId) => {
        if (folderId === 'root') return [{ id: 'root', name: 'Root' }];

        const path = [];
        const stack = State.data.library.map(item => [item, []]);
        const visited = new Set();

        while (stack.length > 0) {
            const [curr, parentPath] = stack.pop();
            if (visited.has(curr.id)) continue;
            visited.add(curr.id);

            const currentPath = [...parentPath, { id: curr.id, name: curr.name }];
            
            if (curr.id === folderId) {
                path.push(...currentPath);
                break;
            }

            if (curr.type === 'folder' && curr.children) {
                for (let i = curr.children.length - 1; i >= 0; i--) {
                    stack.push([curr.children[i], currentPath]);
                }
            }
        }

        return [{ id: 'root', name: 'Root' }, ...path];
    },

    /**
     * Flattens tree into array with path for CSV export
     */
    flattenLibrary: (items = State.data.library, parentPath = '') => {
        let flattened = [];
        
        items.forEach(item => {
            const currentPath = parentPath ? `${parentPath}/${item.name}` : item.name;
            
            if (item.type === 'template') {
                const parsed = MailTo.parse(item.mailto);
                flattened.push({
                    name: item.name,
                    path: parentPath || '/',
                    to: parsed.to,
                    cc: parsed.cc,
                    bcc: parsed.bcc,
                    subject: parsed.subject,
                    body: parsed.body
                });
            }
            
            if (item.type === 'folder' && item.children) {
                flattened = flattened.concat(State.flattenLibrary(item.children, currentPath));
            }
        });
        
        return flattened;
    },

    /**
     * Imports CSV data into library structure
     */
    importFromCSV: (records) => {
        const folderMap = new Map([['/', State.data.library]]);

        records.forEach(record => {
            let path = (record.path || '/').trim();
            if (!path.startsWith('/')) path = '/' + path;

            // Ensure folder path exists
            if (!folderMap.has(path)) {
                const parts = path.split('/').filter(p => p);
                let currentPath = '/';
                let currentArray = State.data.library;

                parts.forEach(part => {
                    const nextPath = currentPath === '/' ? `/${part}` : `${currentPath}/${part}`;
                    
                    if (!folderMap.has(nextPath)) {
                        let folder = currentArray.find(item => item.type === 'folder' && item.name === part);
                        if (!folder) {
                            folder = {
                                id: Utils.generateId(),
                                type: 'folder',
                                name: part,
                                children: []
                            };
                            currentArray.push(folder);
                        }
                        folderMap.set(nextPath, folder.children);
                    }

                    currentArray = folderMap.get(nextPath);
                    currentPath = nextPath;
                });
            }

            // Add template
            const targetArray = folderMap.get(path);
            const mailto = MailTo.build({
                to: record.to,
                cc: record.cc,
                bcc: record.bcc,
                subject: record.subject,
                body: record.body
            });

            targetArray.push({
                id: Utils.generateId(),
                type: 'template',
                name: record.name,
                mailto: mailto
            });
        });
    }
};

/* =============================================================================
   MAILTO OPERATIONS
   ============================================================================= */

const MailTo = {
    /**
     * Parses a mailto link into components
     */
    parse: (str) => {
        const data = { to: '', cc: '', bcc: '', subject: '', body: '' };
        if (!str || !str.startsWith('mailto:')) return data;

        try {
            const qIndex = str.indexOf('?');
            if (qIndex === -1) {
                data.to = decodeURIComponent(str.substring(7));
                return data;
            }

            data.to = decodeURIComponent(str.substring(7, qIndex));
            const params = new URLSearchParams(str.substring(qIndex + 1));
            
            ['subject', 'body', 'cc', 'bcc'].forEach(key => {
                if (params.has(key)) data[key] = params.get(key);
            });
        } catch (err) {
            console.error('Failed to parse mailto:', err);
        }

        return data;
    },

    /**
     * Builds a mailto link from components
     */
    build: (data) => {
        try {
            const params = [];
            CONFIG.MAILTO_PARAMS.forEach(key => {
                if (data[key]) params.push(`${key}=${encodeURIComponent(data[key])}`);
            });
            if (data.body) {
                params.push(`body=${encodeURIComponent(data.body).replace(/%0A/g, '%0D%0A')}`);
            }
            return `mailto:${encodeURIComponent(data.to || '')}?${params.join('&')}`;
        } catch (err) {
            console.error('Failed to build mailto:', err);
            return '';
        }
    }
};

/* =============================================================================
   APPLICATION LOGIC
   ============================================================================= */

const App = {
    elements: {},

    /**
     * Initializes the application
     */
    init: async () => {
        // Cache DOM elements
        App.elements = {
            treeContainer: document.getElementById('tree-list-container'),
            breadcrumb: document.getElementById('breadcrumb-container'),
            uploadWrapper: document.getElementById('upload-wrapper'),
            fileInput: document.getElementById('msg-upload'),
            resultTo: document.getElementById('result-to'),
            resultCc: document.getElementById('result-cc'),
            resultBcc: document.getElementById('result-bcc'),
            resultSubject: document.getElementById('result-subject'),
            resultBody: document.getElementById('result-body'),
            resultMailto: document.getElementById('result-mailto'),
            resultLink: document.getElementById('result-link'),
            outputWrapper: document.getElementById('output-wrapper'),
            saveTemplateName: document.getElementById('save-template-name'),
            saveTargetFolder: document.getElementById('save-target-folder'),
            btnNewFolder: document.getElementById('btn-new-folder'),
            btnGenerate: document.getElementById('btn-generate'),
            btnSave: document.getElementById('btn-save-to-library'),
            btnClear: document.getElementById('btn-clear-all'),
            btnCopy: document.getElementById('copy-mailto-btn'),
            btnImportCSV: document.getElementById('btn-import-csv'),
            btnExportCSV: document.getElementById('btn-export-csv')
        };

        // Load MsgReader module
        try {
            const module = await import('./msgreader.js');
            window.MsgReader = module.MsgReader;
        } catch (err) {
            console.error('Failed to load msgreader.js:', err);
            UI.showToast('Warning: Email file import unavailable');
        }

        State.load();
        App.attachEventListeners();
        App.renderLibrary();
        App.refreshFolderDropdown();
    },

    /**
     * Attaches all event listeners
     */
    attachEventListeners: () => {
        // Upload zone
        App.elements.uploadWrapper.addEventListener('click', () => {
            App.elements.fileInput.click();
        });

        App.elements.fileInput.addEventListener('change', (e) => {
            if (e.target.files[0]) App.handleFileUpload(e.target.files[0]);
        });

        // Drag and drop
        ['dragenter', 'dragover'].forEach(evt => {
            App.elements.uploadWrapper.addEventListener(evt, (e) => {
                e.preventDefault();
                e.stopPropagation();
            });
        });

        App.elements.uploadWrapper.addEventListener('drop', (e) => {
            e.preventDefault();
            e.stopPropagation();
            if (e.dataTransfer.files[0]) App.handleFileUpload(e.dataTransfer.files[0]);
        });

        // Buttons
        App.elements.btnNewFolder.addEventListener('click', App.createFolder);
        App.elements.btnGenerate.addEventListener('click', App.generateLink);
        App.elements.btnSave.addEventListener('click', App.saveTemplate);
        App.elements.btnClear.addEventListener('click', App.clearForm);
        App.elements.btnCopy.addEventListener('click', App.copyLink);
        App.elements.btnImportCSV.addEventListener('click', App.importCSV);
        App.elements.btnExportCSV.addEventListener('click', App.exportCSV);

        // Tree navigation
        App.elements.treeContainer.addEventListener('click', App.handleTreeClick);
        App.elements.breadcrumb.addEventListener('click', App.handleBreadcrumbClick);

        // Close modal on overlay click
        document.getElementById('modal-overlay').addEventListener('click', (e) => {
            if (e.target.id === 'modal-overlay') UI.hideModal();
        });
    },

    /**
     * Handles uploaded email file
     */
    handleFileUpload: (file) => {
        if (!window.MsgReader) {
            UI.showToast('Email parser not available');
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const parsed = window.MsgReader.read(e.target.result);
                
                App.elements.resultSubject.value = parsed.subject || '';
                App.elements.resultBody.value = parsed.body || '';

                const recipientMap = { 1: [], 2: [], 3: [] };
                parsed.recipients.forEach(r => {
                    const addr = r.email || (r.name?.includes('@') ? r.name : '');
                    if (addr) recipientMap[r.recipientType || 1].push(addr);
                });

                App.elements.resultTo.value = recipientMap[1].join(', ');
                App.elements.resultCc.value = recipientMap[2].join(', ');
                App.elements.resultBcc.value = recipientMap[3].join(', ');

                UI.showToast('File loaded successfully');
            } catch (err) {
                console.error('File parsing error:', err);
                UI.showModal('Import Error', `<p>Failed to parse file: ${Utils.escapeHTML(err.message)}</p>`, [
                    { label: 'OK' }
                ]);
            }
        };
        reader.readAsArrayBuffer(file);
    },

    /**
     * Generates mailto link
     */
    generateLink: () => {
        const data = {
            to: App.elements.resultTo.value,
            cc: App.elements.resultCc.value,
            bcc: App.elements.resultBcc.value,
            subject: App.elements.resultSubject.value,
            body: App.elements.resultBody.value
        };

        const mailto = MailTo.build(data);

        if (mailto.length > 2000) {
            UI.showToast('Warning: Link exceeds 2000 characters');
        }

        App.elements.resultMailto.value = mailto;
        App.elements.resultLink.href = mailto;
        App.elements.outputWrapper.classList.remove('hidden');
        App.refreshFolderDropdown();
    },

    /**
     * Saves template to library
     */
    saveTemplate: () => {
        if (!App.elements.resultMailto.value) {
            App.generateLink();
        }

        const name = App.elements.saveTemplateName.value.trim() || 
                     App.elements.resultSubject.value.trim() || 
                     'Untitled Template';
        const targetId = App.elements.saveTargetFolder.value || State.currentFolderId;
        
        const result = State.findItem(targetId);
        if (!result) return;

        const folder = result.item || result;
        if (!folder.children) folder.children = State.data.library;

        folder.children.push({
            id: Utils.generateId(),
            type: 'template',
            name: name,
            mailto: App.elements.resultMailto.value
        });

        State.save();
        App.renderLibrary();
        UI.showToast('Template saved');
    },

    /**
     * Creates a new folder
     */
    createFolder: () => {
        UI.showModal('New Folder', `
            <div class="form-group">
                <label for="folder-name">Folder Name</label>
                <input type="text" id="folder-name" class="form-input" placeholder="Enter folder name">
            </div>
        `, [
            { label: 'Cancel' },
            { label: 'Create', class: 'btn-primary', callback: () => {
                const input = document.getElementById('folder-name');
                const name = input.value.trim();
                
                if (!name) {
                    UI.showToast('Folder name required');
                    return false;
                }

                const result = State.findItem(State.currentFolderId);
                const folder = result?.item || result || { children: State.data.library };

                folder.children.push({
                    id: Utils.generateId(),
                    type: 'folder',
                    name: name,
                    children: []
                });

                State.save();
                App.renderLibrary();
                App.refreshFolderDropdown();
            }}
        ]);
    },

    /**
     * Clears the form
     */
    clearForm: () =>
