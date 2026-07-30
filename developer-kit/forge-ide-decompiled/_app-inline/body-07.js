        // --- EVENT BINDINGS ---
        $(document).ready(function () {
            async function copyLinkAndPrompt(url, label) {
                const targetUrl = String(url || '').trim();
                if (!targetUrl) return;
                let copied = false;
                try {
                    if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
                        await navigator.clipboard.writeText(targetUrl);
                        copied = true;
                    }
                } catch (_) { }
                if (!copied) {
                    try {
                        const ta = document.createElement('textarea');
                        ta.value = targetUrl;
                        ta.setAttribute('readonly', '');
                        ta.style.position = 'fixed';
                        ta.style.left = '-9999px';
                        document.body.appendChild(ta);
                        ta.select();
                        copied = document.execCommand('copy');
                        document.body.removeChild(ta);
                    } catch (_) { }
                }
                const appName = label || 'External app';
                if (copied) {
                    alert(appName + ' link copied. Open a different tab and paste it in the address bar.');
                } else {
                    alert('Could not copy automatically. Open a different tab and paste this link:\n\n' + targetUrl);
                }
            }

            window.copyLinkAndPrompt = copyLinkAndPrompt;

            $('#campfire-btn').on('click', async function (event) {
                event.preventDefault();
                const url = this.getAttribute('data-link-url') || '';
                await copyLinkAndPrompt(url, 'Campfire');
            });

            $('#resources-spark-item').on('click', async function (event) {
                event.preventDefault();
                const url = this.getAttribute('data-link-url') || '';
                await copyLinkAndPrompt(url, 'Spark');
            });

            // Hover-open for dropdowns (desktop & pointer devices)
            function enableHoverDropdowns() {
                const hoverCapable = window.matchMedia('(hover: hover)').matches;
                if (!hoverCapable) return; // keep click behavior on touch devices

                $('#myTab .dropdown').each(function () {
                    const $dd = $(this);
                    $dd.on('mouseenter', function () {
                        $dd.addClass('show');
                        $dd.find('> .dropdown-menu').addClass('show');
                        $dd.find('> .dropdown-toggle').attr('aria-expanded', 'true');
                    });
                    $dd.on('mouseleave', function () {
                        $dd.removeClass('show');
                        $dd.find('> .dropdown-menu').removeClass('show');
                        $dd.find('> .dropdown-toggle').attr('aria-expanded', 'false');
                    });
                });
            }
            enableHoverDropdowns();

            // Keep only one highlighted tool at a time
            function syncActiveToTarget(targetLink) {
                // Clear all
                $('#myTab .dropdown-item, #myTab > .nav-item > .nav-link').removeClass('active');
                $('#myTab .dropdown-toggle').removeClass('active');

                // Activate the clicked/target item
                $(targetLink).addClass('active');

                // Also highlight the parent category (Plan/Build/Test/Ship) if applicable
                const $dd = $(targetLink).closest('.dropdown');
                if ($dd.length) {
                    $dd.find('> .dropdown-toggle').addClass('active');
                }
            }

            // On tab shown, fix active classes
            $('a[data-bs-toggle="tab"]').on('shown.bs.tab', function (e) {
                syncActiveToTarget(e.target);
            });

            // Initial highlight for whatever pane is active on load (editor)
            (function setInitialActive() {
                const $initialPane = $('.tab-pane.show.active');
                if ($initialPane.length) {
                    const id = $initialPane.attr('id');
                    const $link = $('#myTab').find('a[data-bs-toggle="tab"][href="#' + id + '"]');
                    if ($link.length) syncActiveToTarget($link.get(0));
                }
            })();

            // Quick-jump buttons for the Tour
            $('.js-open-tab').on('click', function () {
                const target = $(this).data('target');
                if (!target) return;
                const $link = $('a[data-bs-toggle="tab"][href="' + target + '"]');
                if ($link.length) {
                    bootstrap.Tab.getOrCreateInstance($link[0]).show(); // 'shown.bs.tab' will sync classes
                }
            });

            // Floating Tour button opens the hidden Tour tab
            $('#tour-fab').on('click', function () {
                const tourTab = document.querySelector('a#tour-tab');
                if (tourTab) {
                    bootstrap.Tab.getOrCreateInstance(tourTab).show();
                }
            });

            const quickPromptTemplates = {
                'new-build': {
                    title: 'Quick Prompt: New Build',
                    description: 'Use this when creating a brand-new offline HTML app.',
                    placeholder: 'Example: track maintenance requests with add/edit/delete, status filters, and CSV export.'
                },
                'edit-code': {
                    title: 'Quick Prompt: Edit Code',
                    description: 'Use this to request a feature or code change to an existing app.',
                    placeholder: 'Example: Add a filter panel and keep current behavior unchanged.'
                },
                'debug-code': {
                    title: 'Quick Prompt: Debug Code',
                    description: 'Use this to fix a bug in an existing app.',
                    placeholder: 'Example: Save button does nothing after adding a row. Console says ...'
                }
            };
            let activeQuickPromptTemplate = 'new-build';

            const getLoadedFolderRef = () => {
                if (typeof loadFolder !== 'undefined' && loadFolder) return loadFolder;
                if (typeof window !== 'undefined' && window.loadFolder) return window.loadFolder;
                return null;
            };

            const isLoadedProjectFolder = () => {
                const folder = getLoadedFolderRef();
                return !!(folder && folder.fileHandle);
            };

            const LARGE_CODE_FILE_TOKEN_LIMIT = 100000;
            const PROMPT_ALWAYS_EXCLUDED_FILENAMES = new Set([
                'devconsole.js',
                'testrecorder.js',
                'test-recorder.js',
                'advanceddebug.js'
            ]);

            const getQuickPromptEntryPointOptions = () => {
                const folder = getLoadedFolderRef();
                if (!folder || !Array.isArray(folder.fileStructure)) {
                    return [];
                }
                const htmlFiles = folder.fileStructure
                    .filter(f =>
                        f &&
                        f.kind === 'file' &&
                        /\.html?$/i.test(String(f.relativePath || f.name || '')) &&
                        !isPromptExcludedPath(f.relativePath || f.name || '')
                    )
                    .map(f => String(f.relativePath || f.name || '').replace(/\\/g, '/'))
                    .filter(Boolean);
                htmlFiles.sort((a, b) => {
                    const aIndex = a.toLowerCase() === 'index.html';
                    const bIndex = b.toLowerCase() === 'index.html';
                    if (aIndex && !bIndex) return -1;
                    if (!aIndex && bIndex) return 1;
                    return a.localeCompare(b);
                });
                return htmlFiles;
            };

            const normalizeProjectPath = value => String(value || '')
                .replace(/\\/g, '/')
                .replace(/^\.\//, '')
                .replace(/^\/+/, '')
                .split('/')
                .filter(Boolean)
                .join('/');

            const isPromptExcludedPath = value => {
                const normalized = normalizeProjectPath(value).toLowerCase();
                if (!normalized) return true;
                const baseName = normalized.split('/').pop();
                if (PROMPT_ALWAYS_EXCLUDED_FILENAMES.has(baseName)) return true;
                const parts = normalized.split('/');
                if (parts.some(part => ['.git', '.vscode', '.checkpoints', 'node_modules', 'shipped app files', 'shipped apps'].includes(part))) {
                    return true;
                }
                return /\.(?:png|jpe?g|gif|webp|ico|bmp|pdf|zip|gz|tgz|7z|rar|exe|dll|wasm|pptx?|xlsx?|docx?)$/i.test(normalized);
            };

            const estimatePromptTokenCount = text => {
                const normalizedText = String(text || '').replace(/\s+/g, ' ').trim();
                return Math.ceil(normalizedText.length / 4);
            };

            const isJsOrCssPromptFile = value => /\.(?:css|js|mjs|cjs)$/i.test(normalizeProjectPath(value));

            const shouldExcludeLargeCodeFile = (relativePath, tokenCount) => {
                return isJsOrCssPromptFile(relativePath) && Number(tokenCount || 0) > LARGE_CODE_FILE_TOKEN_LIMIT;
            };

            const escapePromptModalHtml = value => String(value || '')
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#39;');

            const formatPromptTokenCount = tokenCount => `~${Number(tokenCount || 0).toLocaleString()} tokens`;

            window.forgePromptLargeFileConfirm = async candidates => {
                const files = Array.isArray(candidates)
                    ? candidates.filter(item => item && item.path)
                    : [];
                if (!files.length) return new Set();

                if (!(window.bootstrap && bootstrap.Modal)) {
                    const names = files.map(item => `${item.path} (${formatPromptTokenCount(item.tokenCount)})`).join('\n');
                    const ok = window.confirm(`Forge found large JS/CSS files:\n\n${names}\n\nExclude these from the AI prompt?`);
                    return new Set(ok ? files.map(item => item.path) : []);
                }

                return await new Promise(resolve => {
                    const existing = document.getElementById('prompt-large-file-modal');
                    if (existing) existing.remove();

                    const modalEl = document.createElement('div');
                    modalEl.className = 'modal fade';
                    modalEl.id = 'prompt-large-file-modal';
                    modalEl.tabIndex = -1;
                    modalEl.setAttribute('aria-labelledby', 'prompt-large-file-modal-label');
                    modalEl.setAttribute('aria-hidden', 'true');

                    const rows = files.map((item, index) => {
                        const id = `prompt-large-file-${index}`;
                        return `
                            <label class="list-group-item d-flex gap-2 align-items-start bg-dark text-light border-secondary" for="${id}">
                                <input class="form-check-input mt-1 prompt-large-file-choice" type="checkbox" id="${id}" value="${escapePromptModalHtml(item.path)}" checked>
                                <span class="flex-grow-1">
                                    <span class="d-block text-break">${escapePromptModalHtml(item.path)}</span>
                                    <small class="text-muted">${escapePromptModalHtml(formatPromptTokenCount(item.tokenCount))}</small>
                                </span>
                            </label>
                        `;
                    }).join('');

                    modalEl.innerHTML = `
                        <div class="modal-dialog modal-lg modal-dialog-centered">
                            <div class="modal-content bg-dark text-light border-secondary">
                                <div class="modal-header border-secondary">
                                    <h5 class="modal-title" id="prompt-large-file-modal-label">Large code files found</h5>
                                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
                                </div>
                                <div class="modal-body">
                                    <p class="mb-2">These referenced JS/CSS files are over ~100k tokens. Checked files will be excluded from the AI prompt.</p>
                                    <div class="d-flex gap-2 mb-2 flex-wrap">
                                        <button type="button" class="btn btn-sm btn-outline-light" id="prompt-large-file-select-all">Select all</button>
                                        <button type="button" class="btn btn-sm btn-outline-light" id="prompt-large-file-clear">Clear selection</button>
                                    </div>
                                    <div class="list-group" style="max-height: 45vh; overflow:auto;">
                                        ${rows}
                                    </div>
                                </div>
                                <div class="modal-footer border-secondary">
                                    <button type="button" class="btn btn-outline-light" id="prompt-large-file-include-all">Include all</button>
                                    <button type="button" class="btn btn-primary" id="prompt-large-file-confirm">Exclude selected</button>
                                </div>
                            </div>
                        </div>
                    `;

                    document.body.appendChild(modalEl);
                    const modal = bootstrap.Modal.getOrCreateInstance(modalEl, { backdrop: 'static' });
                    let settled = false;
                    const getChecked = () => new Set(Array.from(modalEl.querySelectorAll('.prompt-large-file-choice:checked')).map(input => input.value));
                    const settle = paths => {
                        settled = true;
                        modal.hide();
                        resolve(paths);
                    };

                    modalEl.querySelector('#prompt-large-file-select-all')?.addEventListener('click', () => {
                        modalEl.querySelectorAll('.prompt-large-file-choice').forEach(input => { input.checked = true; });
                    });
                    modalEl.querySelector('#prompt-large-file-clear')?.addEventListener('click', () => {
                        modalEl.querySelectorAll('.prompt-large-file-choice').forEach(input => { input.checked = false; });
                    });
                    modalEl.querySelector('#prompt-large-file-include-all')?.addEventListener('click', () => settle(new Set()));
                    modalEl.querySelector('#prompt-large-file-confirm')?.addEventListener('click', () => settle(getChecked()));
                    modalEl.addEventListener('hidden.bs.modal', () => {
                        const defaultPaths = new Set(files.map(item => item.path));
                        modalEl.remove();
                        if (!settled) resolve(defaultPaths);
                    }, { once: true });

                    modal.show();
                });
            };

            const resolveLargePromptFileExclusions = async candidates => {
                const files = Array.isArray(candidates) ? candidates.filter(item => item && item.path) : [];
                if (!files.length) return new Set();
                if (typeof window.forgePromptLargeFileConfirm === 'function') {
                    return await window.forgePromptLargeFileConfirm(files);
                }
                return new Set(files.map(item => item.path));
            };

            const getProjectFileByRelativePath = relativePath => {
                const folder = getLoadedFolderRef();
                const normalized = normalizeProjectPath(relativePath);
                if (!folder || !Array.isArray(folder.fileStructure) || !normalized || isPromptExcludedPath(normalized)) return null;
                return folder.fileStructure.find(file =>
                    file &&
                    file.kind === 'file' &&
                    normalizeProjectPath(file.relativePath || file.name || '') === normalized
                ) || null;
            };

            const readProjectFileText = async file => {
                if (!file || isPromptExcludedPath(file.relativePath || file.name || '')) return '';
                const folder = getLoadedFolderRef();
                if (!folder) return '';

                if (typeof folder.getFileContent === 'function') {
                    try {
                        return await folder.getFileContent(file);
                    } catch (_) {}
                }

                if (file.entry && typeof file.entry.getFile === 'function') {
                    try {
                        const fileObj = await file.entry.getFile();
                        return await fileObj.text();
                    } catch (_) {}
                }

                return '';
            };

            const resolveHtmlRefToProjectPath = (refValue, basePath = []) => {
                const raw = String(refValue || '').trim();
                if (!raw || raw.startsWith('#') || /^(?:[a-z][a-z0-9+.-]*:|\/\/)/i.test(raw)) {
                    return null;
                }

                const clean = raw.split('#')[0].split('?')[0].trim();
                if (!clean) return null;

                const folder = getLoadedFolderRef();
                if (folder && typeof folder._resolveHtmlRefToProjectPath === 'function') {
                    try {
                        return folder._resolveHtmlRefToProjectPath(clean, basePath);
                    } catch (_) {}
                }

                if (clean.startsWith('/')) {
                    return normalizeProjectPath(clean.replace(/^\/+/, ''));
                }
                return normalizeProjectPath([...basePath, clean].join('/'));
            };

            const gatherCodebaseTextForEntryPoint = async entryRelativePath => {
                const folder = getLoadedFolderRef();
                if (!folder || !Array.isArray(folder.fileStructure)) return '';

                const normalizedEntryPath = normalizeProjectPath(entryRelativePath);
                const fallbackEntryPath = getQuickPromptEntryPointOptions()[0] || '';
                const entryFile =
                    getProjectFileByRelativePath(normalizedEntryPath) ||
                    getProjectFileByRelativePath(fallbackEntryPath);
                if (!entryFile) return '';

                const entryText = await readProjectFileText(entryFile);
                if (!String(entryText || '').trim()) return '';

                const filesToRead = new Set([
                    normalizeProjectPath(entryFile.relativePath || entryFile.name || normalizedEntryPath)
                ]);
                const basePath = Array.isArray(entryFile.path) ? entryFile.path.slice() : [];

                try {
                    if (typeof DOMParser === 'function') {
                        const doc = new DOMParser().parseFromString(entryText, 'text/html');
                        Array.from(doc.querySelectorAll('script[src]')).forEach(node => {
                            const resolved = resolveHtmlRefToProjectPath(node.getAttribute('src'), basePath);
                            if (resolved && /\.(?:js|mjs|cjs)$/i.test(resolved) && !isPromptExcludedPath(resolved)) {
                                filesToRead.add(resolved);
                            }
                        });
                        Array.from(doc.querySelectorAll('link[href]')).forEach(node => {
                            const href = node.getAttribute('href');
                            const rel = String(node.getAttribute('rel') || '').toLowerCase();
                            const isStylesheet = rel.split(/\s+/).includes('stylesheet') || /\.css$/i.test(String(href || ''));
                            if (!isStylesheet) return;
                            const resolved = resolveHtmlRefToProjectPath(href, basePath);
                            if (resolved && /\.css$/i.test(resolved) && !isPromptExcludedPath(resolved)) {
                                filesToRead.add(resolved);
                            }
                        });
                    }
                } catch (_) {}

                const fileEntries = [];
                for (const relativePath of filesToRead) {
                    if (isPromptExcludedPath(relativePath)) continue;
                    const file = getProjectFileByRelativePath(relativePath);
                    if (!file) continue;
                    const text = await readProjectFileText(file);
                    if (!String(text || '').trim()) continue;
                    const tokenCount = estimatePromptTokenCount(text);
                    fileEntries.push({ path: relativePath, text, tokenCount });
                }

                const excludedLargeFiles = await resolveLargePromptFileExclusions(
                    fileEntries
                        .filter(item => shouldExcludeLargeCodeFile(item.path, item.tokenCount))
                        .map(item => ({ path: item.path, tokenCount: item.tokenCount }))
                );

                const parts = [];
                for (const entry of fileEntries) {
                    if (excludedLargeFiles.has(entry.path)) continue;
                    parts.push(`--- ${entry.path} ---`);
                    parts.push(entry.text);
                    parts.push('');
                }

                return parts.join('\n').trim();
            };

            const hasDevConsoleInLoadedProject = () => {
                const folder = getLoadedFolderRef();
                if (!folder || !Array.isArray(folder.fileStructure)) return false;
                return folder.fileStructure.some(f =>
                    f &&
                    f.kind === 'file' &&
                    String(f.name || '').toLowerCase() === 'devconsole.js'
                );
            };
            window.hasDevConsoleInLoadedProject = hasDevConsoleInLoadedProject;

            const setQuickPromptStatus = (message, level = '') => {
                const el = document.getElementById('quick-prompt-status');
                if (!el) return;
                el.textContent = message || '';
                el.classList.remove('text-success', 'text-warning', 'text-danger');
                if (level) el.classList.add(level);
            };

            const refreshQuickPromptControls = () => {
                const newChatWrap = document.getElementById('quick-prompt-new-chat-wrap');
                const newChatCheckbox = document.getElementById('quick-prompt-new-chat');
                const entryWrap = document.getElementById('quick-prompt-entrypoint-wrap');
                const entrySelect = document.getElementById('quick-prompt-entrypoint');
                const noteEl = document.getElementById('quick-prompt-note');
                const devConsoleCtaEl = document.getElementById('quick-prompt-devconsole-cta');
                if (!newChatWrap || !newChatCheckbox || !entryWrap || !entrySelect || !noteEl) return;

                const needsCodeContextOption = activeQuickPromptTemplate === 'edit-code' || activeQuickPromptTemplate === 'debug-code';
                const hasLoadedFolder = isLoadedProjectFolder();
                const entryOptions = getQuickPromptEntryPointOptions();
                const hasEntryOptions = entryOptions.length > 0;
                const canAppend = needsCodeContextOption && hasLoadedFolder && hasEntryOptions;

                newChatWrap.style.display = needsCodeContextOption ? '' : 'none';
                newChatCheckbox.disabled = !canAppend;
                if (!needsCodeContextOption || !canAppend) {
                    newChatCheckbox.checked = false;
                }

                const showEntry = needsCodeContextOption && canAppend && newChatCheckbox.checked;
                entryWrap.style.display = showEntry ? '' : 'none';
                entrySelect.innerHTML = '';

                if (showEntry) {
                    entryOptions.forEach(rel => {
                        const opt = document.createElement('option');
                        opt.value = rel;
                        opt.textContent = rel;
                        entrySelect.appendChild(opt);
                    });
                    const indexOption = entryOptions.find(v => String(v || '').toLowerCase() === 'index.html');
                    entrySelect.value = indexOption || entryOptions[0];
                }

                if (!needsCodeContextOption) {
                    noteEl.textContent = 'This creates a fresh single-file HTML app prompt.';
                } else if (!hasLoadedFolder) {
                    noteEl.textContent = 'Load a project folder first if you want to append codebase context.';
                } else if (!hasEntryOptions) {
                    noteEl.textContent = 'No HTML entry point found yet. Add an HTML file to append codebase context.';
                } else {
                    noteEl.textContent = 'If "New conversation" is checked, Forge appends code from the selected entry point.';
                }

                if (devConsoleCtaEl) {
                    const showDevConsoleCta = activeQuickPromptTemplate === 'debug-code'
                        && hasLoadedFolder
                        && !hasDevConsoleInLoadedProject();
                    devConsoleCtaEl.style.display = showDevConsoleCta ? '' : 'none';
                }
            };

            const applyQuickPromptTemplate = templateKey => {
                const key = quickPromptTemplates[templateKey] ? templateKey : 'new-build';
                const template = quickPromptTemplates[key];
                activeQuickPromptTemplate = key;

                const titleEl = document.getElementById('quick-prompts-modal-label');
                const descEl = document.getElementById('quick-prompt-description');
                const requestWrapEl = document.getElementById('quick-prompt-request-wrap');
                const requestEl = document.getElementById('quick-prompt-request');
                const debugFieldsEl = document.getElementById('quick-prompt-debug-fields');
                const debugDescriptionEl = document.getElementById('quick-prompt-debug-description');
                const debugErrorsEl = document.getElementById('quick-prompt-debug-errors');
                const debugBuildStatusEl = document.getElementById('quick-prompt-debug-build-status');
                const isDebug = key === 'debug-code';

                if (titleEl) titleEl.textContent = template.title;
                if (descEl) descEl.textContent = template.description;
                if (requestWrapEl) {
                    requestWrapEl.style.display = isDebug ? 'none' : '';
                }
                if (debugFieldsEl) {
                    debugFieldsEl.style.display = isDebug ? '' : 'none';
                }
                if (requestEl) {
                    requestEl.placeholder = template.placeholder;
                    requestEl.value = '';
                }
                if (debugDescriptionEl) {
                    debugDescriptionEl.value = '';
                }
                if (debugErrorsEl) {
                    debugErrorsEl.value = '';
                }
                if (debugBuildStatusEl) {
                    debugBuildStatusEl.checked = false;
                }
                const newChatCheckbox = document.getElementById('quick-prompt-new-chat');
                if (newChatCheckbox) {
                    newChatCheckbox.checked = isDebug || key === 'edit-code';
                }
                setQuickPromptStatus('');
                refreshQuickPromptControls();
            };

            const openQuickPromptModal = (templateKey = 'new-build') => {
                const modalEl = document.getElementById('quick-prompts-modal');
                if (!modalEl || !(window.bootstrap && bootstrap.Modal)) return;
                applyQuickPromptTemplate(templateKey);
                bootstrap.Modal.getOrCreateInstance(modalEl).show();
            };

            const advancedFeatureSpecs = {
                6: {
                    title: 'JSON Save/Load',
                    description: 'Generate a prompt to add JSON backup and restore to your existing app.',
                    contextLabel: 'What data should be saved and restored?',
                    placeholder: 'Example: Save all tasks and settings, but skip temporary filter state.',
                    defaultContext: '[all app data]'
                },
                7: {
                    title: 'CSV/Excel Import',
                    description: 'Generate a prompt to add spreadsheet import and preview.',
                    contextLabel: 'Describe the spreadsheet format users will upload',
                    placeholder: 'Example columns: Name, Rank, Email, Unit (one person per row).',
                    defaultContext: '[spreadsheet format details]'
                },
                8: {
                    title: 'Document Export',
                    description: 'Generate a prompt to add Word/PPT/PDF style exports.',
                    contextLabel: 'Describe the exported document layout',
                    placeholder: 'Example: Title, summary paragraph, then a table of events.',
                    defaultContext: '[document format details]'
                },
                9: {
                    title: 'OCR (Image to Text)',
                    description: 'Generate a prompt to add image upload and OCR extraction.',
                    contextLabel: 'What types of images will users upload? (optional)',
                    placeholder: 'Example: Photos of handwritten forms and scanned checklists.',
                    defaultContext: '[optional image types]'
                },
                10: {
                    title: 'LLM JSON Pipeline',
                    description: 'Generate a prompt for a boomerang flow: app data -> AI -> JSON -> app.',
                    contextLabel: 'What data is analyzed and what output should AI return?',
                    placeholder: 'Example: Analyze mission event logs and return anomaly summary JSON.',
                    defaultContext: '[analysis goals and JSON schema details]'
                },
                11: {
                    title: 'SharePoint List Integration',
                    description: 'Generate a prompt for SharePoint list-backed persistence.',
                    contextLabel: 'Describe your app data model, what should be stored, and any hosting details',
                    placeholder: 'Example: Store task records with status, assignee, due date, notes, and linked files. App will run in Firepit or as a legacy IntelShare SitePages .aspx page.',
                    defaultContext: '[SharePoint site URL, deploy target, app data model, what should be stored, whether lists already exist, and any exact list/library/folder details if already created]'
                },
                18: {
                    title: 'SharePoint Live Polling (After List Integration)',
                    description: 'Generate a follow-up prompt to add simple polling-based live updates after SharePoint list integration already exists.',
                    contextLabel: 'Describe the existing code paths, state flow, and where live updates should appear',
                    placeholder: 'Example: The app already reads and writes tasks from a SharePoint list. Add lightweight polling so status changes from other users appear in the task board and detail pane without reload.',
                    defaultContext: '[existing SharePoint-backed code flow, current state/update behavior, and where remote updates should appear in the UI]'
                },
                19: {
                    title: 'Fusion Wiki Integration',
                    description: 'Generate a prompt to add Confluence page content-property persistence to a single-file HTML app deployed in a Confluence HTML macro.',
                    contextLabel: 'Describe the app state to persist and any Confluence hosting details',
                    placeholder: 'Example: Persist team notes, filters, and selected doc id on the current Fusion Confluence page. Add basic polling so other page viewers see saved changes.',
                    defaultContext: '[app state to persist, desired sync behavior, and Fusion wiki hosting details]'
                },
                12: {
                    title: 'Ask Sage API Integration',
                    description: 'Generate a prompt to add Ask Sage API calls using the CAPRA or Army tenant endpoint patterns.',
                    contextLabel: 'Describe what Ask Sage should do in your app',
                    placeholder: 'Example: Send selected mission notes to Ask Sage and render response with citations.',
                    defaultContext: '[asksage endpoint usage and response handling details]'
                },
                13: {
                    title: 'GenAI.mil API Integration',
                    description: 'Generate a prompt to add GenAI.mil API calls using the OpenAI-compatible endpoint pattern.',
                    contextLabel: 'Describe what GenAI.mil should do in your app',
                    placeholder: 'Example: Send selected reports to GenAI.mil and render the model response in a review pane.',
                    defaultContext: '[genai.mil endpoint usage and response handling details]'
                },
                14: {
                    title: 'Leaflet Map',
                    description: 'Generate a prompt to add an offline-friendly Leaflet map backed by a local GeoJSON bundle.',
                    contextLabel: 'Describe the map view, overlays, and how users should interact with it',
                    placeholder: 'Example: Show a full-width world map with country hover labels and highlight selected mission areas in a side panel.',
                    defaultContext: '[map layout, overlays, and user interactions]'
                },
                15: {
                    title: 'Shared JSON Live Database',
                    description: 'Open the built-in Shared JSON Live Database wizard and generate a copy/paste-ready prompt for wiring the included library into the app.',
                    contextLabel: '',
                    placeholder: '',
                    defaultContext: ''
                },
                16: {
                    title: 'Split Into Multiple Files',
                    description: 'Generate a prompt to refactor an existing app into one HTML file plus multiple JavaScript and CSS files.',
                    contextLabel: 'Describe how you want the app split or what parts should stay grouped together',
                    placeholder: 'Example: Keep index.html as the only entry point. Split JS into app bootstrap, state, table rendering, and file helpers. Split CSS into base, layout, and components.',
                    defaultContext: '[keep one html entry point; split js/css into sensible modules while preserving behavior]'
                }
            };

            const getAdvancedFeatureSpec = step => advancedFeatureSpecs[Number(step)] || null;

            const setAdvancedIncorporateStatus = (message, level = '') => {
                const el = document.getElementById('advanced-incorporate-status');
                if (!el) return;
                el.textContent = message || '';
                el.classList.remove('text-success', 'text-warning', 'text-danger');
                if (level) el.classList.add(level);
            };

            const setAdvancedIncorporateNote = message => {
                const el = document.getElementById('advanced-incorporate-note');
                if (!el) return;
                el.textContent = message || '';
            };

            let advancedIncorporateCopyFeedbackFadeTimer = null;
            let advancedIncorporateCopyFeedbackHideTimer = null;

            const hideAdvancedIncorporateCopyFeedback = () => {
                const el = document.getElementById('advanced-incorporate-copy-feedback');
                if (!el) return;
                if (advancedIncorporateCopyFeedbackFadeTimer) {
                    clearTimeout(advancedIncorporateCopyFeedbackFadeTimer);
                    advancedIncorporateCopyFeedbackFadeTimer = null;
                }
                if (advancedIncorporateCopyFeedbackHideTimer) {
                    clearTimeout(advancedIncorporateCopyFeedbackHideTimer);
                    advancedIncorporateCopyFeedbackHideTimer = null;
                }
                el.classList.remove('is-visible', 'is-fading');
                el.style.display = 'none';
            };

            const showAdvancedIncorporateCopyFeedback = message => {
                const el = document.getElementById('advanced-incorporate-copy-feedback');
                const textEl = document.getElementById('advanced-incorporate-copy-feedback-text');
                if (!el || !textEl) return;

                hideAdvancedIncorporateCopyFeedback();
                textEl.textContent = message || 'Prompt copied to clipboard';
                el.style.display = 'flex';
                void el.offsetWidth;
                el.classList.add('is-visible');

                advancedIncorporateCopyFeedbackFadeTimer = setTimeout(() => {
                    el.classList.add('is-fading');
                }, 1600);

                advancedIncorporateCopyFeedbackHideTimer = setTimeout(() => {
                    hideAdvancedIncorporateCopyFeedback();
                }, 2400);
            };

            const isAdvancedSharePointStep = step => {
                const n = Number(step);
                return n === 11 || n === 18;
            };
            const isAdvancedSharePointSetupStep = step => Number(step) === 11;
            const isAdvancedLeafletStep = step => Number(step) === 14;
            const normalizeAdvancedSharePointSiteUrl = value => {
                const raw = String(value || '').trim();
                if (!raw) return '';
                return raw
                    .replace('https://flankspeed.sharepoint-mil.us.mcas-gov.us/', 'https://flankspeed.sharepoint-mil.us/')
                    .replace('http://flankspeed.sharepoint-mil.us.mcas-gov.us/', 'https://flankspeed.sharepoint-mil.us/');
            };

            const refreshAdvancedSharePointFields = step => {
                const wrapEl = document.getElementById('advanced-incorporate-sharepoint-fields');
                const siteUrlEl = document.getElementById('advanced-incorporate-sharepoint-site-url');
                const deployTargetEl = document.getElementById('advanced-incorporate-sharepoint-deploy-target');
                const hasListEl = document.getElementById('advanced-incorporate-sharepoint-has-list');
                const detailsWrapEl = document.getElementById('advanced-incorporate-sharepoint-details-wrap');
                const detailsEl = document.getElementById('advanced-incorporate-sharepoint-details');
                const noteEl = document.getElementById('advanced-incorporate-sharepoint-note');
                if (!wrapEl || !siteUrlEl || !deployTargetEl || !hasListEl || !detailsWrapEl || !detailsEl || !noteEl) return;

                const numericStep = Number(step);
                const isSharePoint = isAdvancedSharePointSetupStep(numericStep);
                wrapEl.style.display = isSharePoint ? '' : 'none';
                if (!isSharePoint) return;

                detailsWrapEl.style.display = hasListEl.checked ? '' : 'none';
                const isLegacy = deployTargetEl.value === 'legacy-intelshare';
                noteEl.textContent = hasListEl.checked
                    ? 'Since the SharePoint objects already exist, provide the exact list title, library title, folder path, and column details.'
                    : (isLegacy
                        ? 'Legacy IntelShare deploys the app directly as a SitePages .aspx file. The generated prompt will avoid Firepit assumptions.'
                        : 'Leave this unchecked if the list or library does not exist yet. For Prometheus, use https://flankspeed.sharepoint-mil.us/sites/Prometheus/. The generated prompt will ask AI to create the list and required columns automatically.');
            };

            const refreshAdvancedLeafletFields = step => {
                const wrapEl = document.getElementById('advanced-incorporate-leaflet-fields');
                if (!wrapEl) return;
                const isLeaflet = isAdvancedLeafletStep(step);
                wrapEl.style.display = isLeaflet ? '' : 'none';
                if (isLeaflet && typeof leafletMapTab !== 'undefined' && leafletMapTab && typeof leafletMapTab.syncFormFromState === 'function') {
                    leafletMapTab.syncFormFromState();
                }
            };

            const getAdvancedIncorporateContext = step => {
                const ctxEl = document.getElementById('advanced-incorporate-context');
                if (!ctxEl) return '';
                if (!isAdvancedSharePointStep(step) || Number(step) === 18) {
                    return String(ctxEl.value || '').trim();
                }

                const siteUrlEl = document.getElementById('advanced-incorporate-sharepoint-site-url');
                const deployTargetEl = document.getElementById('advanced-incorporate-sharepoint-deploy-target');
                const hasListEl = document.getElementById('advanced-incorporate-sharepoint-has-list');
                const detailsEl = document.getElementById('advanced-incorporate-sharepoint-details');
                const siteUrl = normalizeAdvancedSharePointSiteUrl(siteUrlEl && 'value' in siteUrlEl ? siteUrlEl.value : '');
                const deployTarget = String(deployTargetEl && 'value' in deployTargetEl ? deployTargetEl.value : 'firepit').trim();
                const appDetails = String(ctxEl.value || '').trim();
                const hasExistingObjects = !!(hasListEl && hasListEl.checked);
                const existingDetails = String(detailsEl && 'value' in detailsEl ? detailsEl.value : '').trim();

                if (!siteUrl) {
                    return '';
                }

                const parts = [
                    `SharePoint site URL: ${siteUrl}`,
                    deployTarget === 'legacy-intelshare'
                        ? 'Deployment target: Legacy IntelShare. The app will be uploaded directly to the SharePoint SitePages folder as a .aspx page. Firepit/SPFx will not be used. Supported hostnames are intelshare.intelink.sgov.gov and intelshare.intelink.gov.'
                        : 'Deployment target: Flank Speed SharePoint Firepit/SPFx page. The app will run through Firepit on a SharePoint site page.',
                    hasExistingObjects
                        ? 'The SharePoint list/library/folder already exists.'
                        : 'The SharePoint list/library/folder does not exist yet. Tell me exactly what to create.',
                    `App data model / storage needs / hosting details: ${appDetails || '[describe what the app stores and how it is hosted]'}`
                ];

                if (Number(step) === 18) {
                    parts.push('This is the follow-up SharePoint live polling prompt. The base SharePoint list integration is already implemented and should be extended, not redesigned.');
                }

                if (hasExistingObjects) {
                    parts.push(`Existing SharePoint object details: ${existingDetails || '[provide exact list title, library title, folder path, and columns/types]'}`);
                }

                return parts.join('\n');
            };

            const buildAdvancedPromptFromTemplate = (step, contextText) => {
                const s = Number(step);
                const spec = getAdvancedFeatureSpec(s);
                if (isAdvancedLeafletStep(s) && typeof leafletMapTab !== 'undefined' && leafletMapTab && typeof leafletMapTab.buildAdvancedPrompt === 'function') {
                    return leafletMapTab.buildAdvancedPrompt(contextText);
                }
                const template = (typeof aiHelper !== 'undefined' && aiHelper && aiHelper.templates)
                    ? aiHelper.templates[s]
                    : null;
                if (!template) return '';
                const cleanContext = String(contextText || '').trim();
                const fallback = spec && spec.defaultContext ? spec.defaultContext : '[details]';
                return String(template).replace('{context}', cleanContext || fallback);
            };

            const copyToClipboardBestEffort = async text => {
                try {
                    await navigator.clipboard.writeText(text);
                    return true;
                } catch (_) {
                    return false;
                }
            };

            const refreshAdvancedIncorporateControls = () => {
                const modalEl = document.getElementById('advanced-incorporate-modal');
                const newChatWrap = document.getElementById('advanced-incorporate-new-chat-wrap');
                const newChatCheckbox = document.getElementById('advanced-incorporate-new-chat');
                const entryWrap = document.getElementById('advanced-incorporate-entrypoint-wrap');
                const entrySelect = document.getElementById('advanced-incorporate-entrypoint');
                if (!modalEl || !newChatWrap || !newChatCheckbox || !entryWrap || !entrySelect) return;

                const step = Number(modalEl.getAttribute('data-advanced-step'));
                refreshAdvancedSharePointFields(step);
                refreshAdvancedLeafletFields(step);

                const hasLoadedFolder = isLoadedProjectFolder();
                const entryOptions = getQuickPromptEntryPointOptions();
                const canAppend = hasLoadedFolder && entryOptions.length > 0;

                newChatWrap.style.display = '';
                newChatCheckbox.disabled = !canAppend;
                if (!canAppend) {
                    newChatCheckbox.checked = false;
                }

                const showEntry = canAppend && newChatCheckbox.checked;
                entryWrap.style.display = showEntry ? '' : 'none';
                entrySelect.innerHTML = '';

                if (showEntry) {
                    entryOptions.forEach(rel => {
                        const opt = document.createElement('option');
                        opt.value = rel;
                        opt.textContent = rel;
                        entrySelect.appendChild(opt);
                    });
                    const indexOption = entryOptions.find(v => String(v || '').toLowerCase() === 'index.html');
                    entrySelect.value = indexOption || entryOptions[0];
                }

                if (!hasLoadedFolder) {
                    setAdvancedIncorporateNote('Load a project folder first to append codebase context.');
                } else if (!entryOptions.length) {
                    setAdvancedIncorporateNote('No HTML entry point found. Add an HTML file to append codebase context.');
                } else {
                    setAdvancedIncorporateNote('If checked, your selected entry point and referenced code are appended for a new conversation.');
                }
            };

            const openAdvancedIncorporateModal = step => {
                const spec = getAdvancedFeatureSpec(step);
                const modalEl = document.getElementById('advanced-incorporate-modal');
                if (!spec || !(window.bootstrap && bootstrap.Modal)) return;
                if (Number(step) === 15) {
                    const nosqlModalEl = document.getElementById('sharedrive-nosql-modal');
                    if (!nosqlModalEl) return;
                    bootstrap.Modal.getOrCreateInstance(nosqlModalEl).show();
                    return;
                }
                if (!modalEl) return;

                modalEl.setAttribute('data-advanced-step', String(step));
                const titleEl = document.getElementById('advanced-incorporate-modal-label');
                const descEl = document.getElementById('advanced-incorporate-description');
                const labelEl = document.getElementById('advanced-incorporate-context-label');
                const ctxEl = document.getElementById('advanced-incorporate-context');
                const outEl = document.getElementById('advanced-incorporate-output');
                const sharePointSiteUrlEl = document.getElementById('advanced-incorporate-sharepoint-site-url');
                const sharePointDeployTargetEl = document.getElementById('advanced-incorporate-sharepoint-deploy-target');
                const sharePointHasListEl = document.getElementById('advanced-incorporate-sharepoint-has-list');
                const sharePointDetailsEl = document.getElementById('advanced-incorporate-sharepoint-details');

                if (titleEl) titleEl.textContent = `Incorporate: ${spec.title}`;
                if (descEl) descEl.textContent = spec.description;
                if (labelEl) labelEl.textContent = spec.contextLabel;
                if (ctxEl) {
                    ctxEl.value = '';
                    ctxEl.placeholder = spec.placeholder;
                }
                if (sharePointSiteUrlEl && 'value' in sharePointSiteUrlEl) sharePointSiteUrlEl.value = '';
                if (sharePointDeployTargetEl && 'value' in sharePointDeployTargetEl) sharePointDeployTargetEl.value = 'firepit';
                if (sharePointHasListEl) sharePointHasListEl.checked = false;
                if (sharePointDetailsEl && 'value' in sharePointDetailsEl) sharePointDetailsEl.value = '';
                if (isAdvancedLeafletStep(step) && typeof leafletMapTab !== 'undefined' && leafletMapTab && typeof leafletMapTab.resetUi === 'function') {
                    leafletMapTab.resetUi();
                }
                if (outEl) outEl.textContent = '';
                hideAdvancedIncorporateCopyFeedback();
                const newChatCheckbox = document.getElementById('advanced-incorporate-new-chat');
                if (newChatCheckbox) newChatCheckbox.checked = true;
                setAdvancedIncorporateStatus('');
                refreshAdvancedIncorporateControls();

                bootstrap.Modal.getOrCreateInstance(modalEl).show();
            };
            window.openAdvancedIncorporateModal = openAdvancedIncorporateModal; // expose for promptLab.js

            const addDevConsoleFromQuickPrompt = async () => {
                const button = document.getElementById('quick-prompt-add-devconsole');
                if (!isLoadedProjectFolder()) {
                    setQuickPromptStatus('Load a project folder first.', 'text-warning');
                    return;
                }
                if (typeof devConsoleTab === 'undefined' || !devConsoleTab || typeof devConsoleTab.addDevConsoleToProject !== 'function') {
                    setQuickPromptStatus('DevConsole integration tool is not available.', 'text-danger');
                    return;
                }

                const originalLabel = button ? button.textContent : '';
                if (button) {
                    button.disabled = true;
                    button.textContent = 'Adding DevConsole.js...';
                }
                setQuickPromptStatus('Adding DevConsole.js to the current project...', '');

                try {
                    await devConsoleTab.addDevConsoleToProject();
                    refreshQuickPromptControls();

                    if (hasDevConsoleInLoadedProject()) {
                        setQuickPromptStatus('DevConsole.js added to the project. Capture the errors, then paste them here.', 'text-success');
                    } else {
                        setQuickPromptStatus('DevConsole.js could not be confirmed in the project. Check the alert or console for details.', 'text-warning');
                    }
                } catch (error) {
                    const message = error && error.message ? error.message : 'Failed to add DevConsole.js.';
                    setQuickPromptStatus(message, 'text-danger');
                } finally {
                    if (button) {
                        button.disabled = false;
                        button.textContent = originalLabel || 'Add DevConsole.js to Project';
                    }
                }
            };

            const getDebugBuildStatusPromptDetails = (shippedOnlyBug) => {
                if (shippedOnlyBug) {
                    return {
                        label: 'Bug only reproduces in the shipped/compiled build.',
                        shippingContext: 'Forge shipping removes the unshipped development banner, packages the app as a compiled single HTML artifact with an embedded manifest, may minify the HTML, and may wrap the app in Forge\'s runtime/security shell. If the build targeted SharePoint compatibility mode, shipping can also rewrite some inline event handlers and change runtime/network behavior for SharePoint hosting.'
                    };
                }
                return {
                    label: 'Bug reproduces in the editable source app before shipping.',
                    shippingContext: ''
                };
            };

            const buildQuickPromptText = (templateKey, requestText, debugDescription = '', debugErrors = '', debugBuildStatus = false) => {
                const clean = String(requestText || '').trim();
                const diffOutputRules = [
                    'Output requirements:',
                    '- Return one copy/pasteable unified diff only, not complete files.',
                    '- Do not use markdown fences around the diff.',
                    '- Do not add custom wrapper marker lines.',
                    '- Include every changed file in the same diff.',
                    '- Use git-style file headers: diff --git a/<path> b/<path>, --- a/<path>, +++ b/<path>, and @@ hunks.',
                    '- Use relative paths exactly as shown in the codebase context.',
                    '- Prefer small independent hunks; do not group unrelated replacements into one large hunk.',
                    '- Keep enough unchanged context around each hunk to identify the location uniquely.',
                    '- For new files, use --- /dev/null and +++ b/<relative/path>.',
                    '- Do not include explanation text before or after the diff.',
                    '- Do not use snippets, ellipses, or step-by-step instructions.'
                ].join('\n');
                if (templateKey === 'edit-code') {
                    return `Edit my existing offline HTML app:\n${clean}\n\nKeep existing behavior unless I explicitly ask to change it.\nDo not output files you are not changing.\n\n${diffOutputRules}`;
                }
                if (templateKey === 'debug-code') {
                    const cleanDescription = String(debugDescription || '').trim();
                    const cleanErrors = String(debugErrors || '').trim() || 'No console errors shown.';
                    const buildStatusDetails = getDebugBuildStatusPromptDetails(debugBuildStatus);
                    const shippingSection = buildStatusDetails.shippingContext
                        ? `\n\nCompiled/shipping context:\n${buildStatusDetails.shippingContext}`
                        : '';
                    return `Debug and fix this issue in my offline HTML app:\n\nBug description:\n${cleanDescription}\n\nConsole errors/logs:\n${cleanErrors}\n\nBuild/shipping status:\n${buildStatusDetails.label}${shippingSection}\n\nKeep unrelated behavior intact.\nDo not output files you are not changing.\n\n${diffOutputRules}`;
                }
                return `Create a single-file, vanilla, offline html file application that:\n${clean}\n\nReturn complete file output, not snippets.`;
            };

            const showQuickPromptCopiedMessage = (appendedCode) => {
                const statusEl = document.getElementById('global-save-status');
                if (!statusEl) return;
                const msg = appendedCode
                    ? 'Prompt + codebase copied. Next: go to AI Services and paste it.'
                    : 'Prompt copied. Next: go to AI Services and paste it.';
                statusEl.textContent = msg;
                setTimeout(() => {
                    if (statusEl.textContent === msg) {
                        statusEl.textContent = '';
                    }
                }, 4500);
            };

            const generateQuickPromptOutput = async () => {
                const requestEl = document.getElementById('quick-prompt-request');
                const debugDescriptionEl = document.getElementById('quick-prompt-debug-description');
                const debugErrorsEl = document.getElementById('quick-prompt-debug-errors');
                const debugBuildStatusEl = document.getElementById('quick-prompt-debug-build-status');
                const newChatCheckbox = document.getElementById('quick-prompt-new-chat');
                const entrySelect = document.getElementById('quick-prompt-entrypoint');
                const requestText = requestEl ? String(requestEl.value || '').trim() : '';
                const debugDescriptionText = debugDescriptionEl ? String(debugDescriptionEl.value || '').trim() : '';
                const debugErrorsText = debugErrorsEl ? String(debugErrorsEl.value || '').trim() : '';
                const debugBuildStatus = !!(debugBuildStatusEl && debugBuildStatusEl.checked);
                const isDebugTemplate = activeQuickPromptTemplate === 'debug-code';
                const requiresMainRequest = !isDebugTemplate;
                if (requiresMainRequest && !requestText) {
                    setQuickPromptStatus('Enter your request first.', 'text-warning');
                    return;
                }
                if (isDebugTemplate && !debugDescriptionText) {
                    setQuickPromptStatus('Enter a plain-language bug description first.', 'text-warning');
                    return;
                }

                try {
                    const isEditOrDebug = activeQuickPromptTemplate === 'edit-code' || activeQuickPromptTemplate === 'debug-code';
                    const wantsNewChat = !!(newChatCheckbox && newChatCheckbox.checked);
                    const availableEntryPoints = getQuickPromptEntryPointOptions();

                    let output = buildQuickPromptText(
                        activeQuickPromptTemplate,
                        requestText,
                        debugDescriptionText,
                        debugErrorsText,
                        debugBuildStatus
                    );
                    let appendedCode = false;

                    if (isEditOrDebug && wantsNewChat && availableEntryPoints.length) {
                        const selectedEntryPoint = entrySelect && entrySelect.value
                            ? entrySelect.value
                            : availableEntryPoints[0];
                        const codeContext = await gatherCodebaseTextForEntryPoint(selectedEntryPoint);
                        if (codeContext) {
                            output += `\n\n--- CODEBASE ---\n${codeContext}`;
                            appendedCode = true;
                        }
                    }

                    const copied = await copyToClipboardBestEffort(output);
                    if (!copied) {
                        setQuickPromptStatus('Copy failed. Try again.', 'text-danger');
                        document.dispatchEvent(new CustomEvent('forge:quick-prompt-generated', {
                            detail: { copied: false, template: activeQuickPromptTemplate, appendedCode }
                        }));
                        return;
                    }

                    setQuickPromptStatus(
                        appendedCode
                            ? 'Prompt + codebase copied. Next: go to AI Services and paste it.'
                            : 'Prompt copied. Next: go to AI Services and paste it.',
                        'text-success'
                    );
                    if (activeQuickPromptTemplate === 'edit-code'
                        && typeof checkpointManager !== 'undefined'
                        && checkpointManager
                        && typeof checkpointManager.armPendingPasteCheckpoint === 'function'
                    ) {
                        checkpointManager.armPendingPasteCheckpoint(requestText);
                    }
                    showQuickPromptCopiedMessage(appendedCode);

                    const modalEl = document.getElementById('quick-prompts-modal');
                    if (modalEl && window.bootstrap && bootstrap.Modal) {
                        const instance = bootstrap.Modal.getOrCreateInstance(modalEl);
                        instance.hide();
                    }

                    setTimeout(() => {
                        document.dispatchEvent(new CustomEvent('forge:quick-prompt-generated', {
                            detail: { copied: true, template: activeQuickPromptTemplate, appendedCode }
                        }));
                    }, 180);
                } catch (error) {
                    const message = error && error.message ? error.message : 'Could not generate quick prompt.';
                    setQuickPromptStatus(message, 'text-danger');
                    console.error('Quick prompt generation error:', error);
                }
            };

            $(document).on('click', '[data-quick-prompt-template]', function (e) {
                e.preventDefault();
                const templateKey = this.getAttribute('data-quick-prompt-template') || 'new-build';
                // Map old template keys to Prompt Lab types
                const typeMap = { 'new-build': 'new-build', 'edit-code': 'edit', 'debug-code': 'debug', 'feature-add': 'feature-add' };
                const plType = typeMap[templateKey] || 'new-build';
                if (typeof toggleRightPanel === 'function') toggleRightPanel(true);
                if (typeof promptLab !== 'undefined') {
                    promptLab.switchRpTab('prompt-lab');
                    promptLab.switchType(plType);
                }
            });
            $('#quick-prompt-generate-btn').on('click', () => generateQuickPromptOutput());
            $('#quick-prompt-new-chat').on('change', () => refreshQuickPromptControls());
            $(document).on('click', '#quick-prompt-add-devconsole', function (e) {
                e.preventDefault();
                addDevConsoleFromQuickPrompt();
            });
            $('#quick-prompts-modal').on('shown.bs.modal', () => {
                const requestEl = document.getElementById('quick-prompt-request');
                const debugDescriptionEl = document.getElementById('quick-prompt-debug-description');
                refreshQuickPromptControls();
                if (activeQuickPromptTemplate === 'debug-code' && debugDescriptionEl) {
                    debugDescriptionEl.focus();
                    return;
                }
                if (requestEl) requestEl.focus();
            });

            $(document).on('click', '[data-advanced-incorporate-step]', function (e) {
                e.preventDefault();
                const step = Number(this.getAttribute('data-advanced-incorporate-step'));
                if (!step) return;
                openAdvancedIncorporateModal(step);
            });

            $('#advanced-incorporate-modal').on('shown.bs.modal', () => {
                const modalEl = document.getElementById('advanced-incorporate-modal');
                const step = Number(modalEl && modalEl.getAttribute('data-advanced-step'));
                if (isAdvancedSharePointStep(step)) {
                    const siteUrlEl = document.getElementById('advanced-incorporate-sharepoint-site-url');
                    if (siteUrlEl) siteUrlEl.focus();
                    return;
                }
                const ctxEl = document.getElementById('advanced-incorporate-context');
                if (ctxEl) ctxEl.focus();
            });
            $('#advanced-incorporate-new-chat').on('change', () => refreshAdvancedIncorporateControls());
            $('#advanced-incorporate-sharepoint-has-list').on('change', () => refreshAdvancedIncorporateControls());
            $('#advanced-incorporate-sharepoint-deploy-target').on('change', () => refreshAdvancedIncorporateControls());

            $('#advanced-incorporate-generate-btn').on('click', async () => {
                const modalEl = document.getElementById('advanced-incorporate-modal');
                const ctxEl = document.getElementById('advanced-incorporate-context');
                const outEl = document.getElementById('advanced-incorporate-output');
                const newChatCheckbox = document.getElementById('advanced-incorporate-new-chat');
                const entrySelect = document.getElementById('advanced-incorporate-entrypoint');
                if (!modalEl || !ctxEl || !outEl || !newChatCheckbox) return;

                const step = Number(modalEl.getAttribute('data-advanced-step'));
                if (!step) {
                    setAdvancedIncorporateStatus('No feature selected.', 'text-warning');
                    return;
                }

                hideAdvancedIncorporateCopyFeedback();

                const contextText = getAdvancedIncorporateContext(step);
                if (isAdvancedSharePointSetupStep(step)) {
                    const hasListEl = document.getElementById('advanced-incorporate-sharepoint-has-list');
                    const detailsEl = document.getElementById('advanced-incorporate-sharepoint-details');
                    if (!contextText) {
                        setAdvancedIncorporateStatus('Enter the SharePoint site URL first.', 'text-warning');
                        return;
                    }
                    if (hasListEl && hasListEl.checked && !String(detailsEl && 'value' in detailsEl ? detailsEl.value : '').trim()) {
                        setAdvancedIncorporateStatus('If the SharePoint objects already exist, enter the exact list/library details.', 'text-warning');
                        return;
                    }
                }

                const prompt = buildAdvancedPromptFromTemplate(step, contextText);
                if (!prompt) {
                    setAdvancedIncorporateStatus('Prompt template not found for this feature.', 'text-danger');
                    return;
                }

                const wantsNewChat = !!newChatCheckbox.checked;
                let output = prompt;
                let appendedCode = false;

                if (wantsNewChat) {
                    const availableEntryPoints = getQuickPromptEntryPointOptions();
                    const selectedEntryPoint = entrySelect && entrySelect.value
                        ? entrySelect.value
                        : availableEntryPoints[0];
                    const codeContext = await gatherCodebaseTextForEntryPoint(selectedEntryPoint);
                    if (codeContext) {
                        output += `\n\nCurrent codebase:\n--- CODEBASE ---\n${codeContext}`;
                        appendedCode = true;
                    }
                }

                outEl.textContent = output;
                const copied = await copyToClipboardBestEffort(output);
                if (copied) {
                    setAdvancedIncorporateStatus(
                        appendedCode
                            ? 'Prompt + codebase copied. Paste it into a new AI conversation.'
                            : 'Prompt copied. Paste it into your AI chat.',
                        'text-success'
                    );
                    showAdvancedIncorporateCopyFeedback(
                        appendedCode
                            ? 'Prompt + codebase copied to clipboard'
                            : 'Prompt copied to clipboard'
                    );
                } else {
                    setAdvancedIncorporateStatus('Prompt generated, but copy failed. Copy manually from the output box.', 'text-warning');
                }
            });

            $('#add-security-headers').on('change', function () {
                if (this && document.querySelector('#sharepoint-compat-mode')?.checked) {
                    this.dataset.disabledBySharePointCompat = 'false';
                }
                if (this && this.checked === false) {
                    alert(
                        "Security headers were disabled.\n\n" +
                        "Forge can no longer provide strong assurance that app data cannot be exfiltrated without closer review.\n\n" +
                        "Coordinate with your local ISSM/ISSO or the Prometheus group before operational use."
                    );
                }
            });

            $('#allow-asksage-api').on('change', function () {
                if (this && this.checked === true) {
                    alert(
                        "Ask Sage API access was enabled.\n\n" +
                        "Review the API data-handling details and verify what information types are authorized for your Ask Sage tenant, including CUI, before use."
                    );
                }
            });

            $('#allow-genaimil-api').on('change', function () {
                if (this && this.checked === true) {
                    alert(
                        "GenAI.mil API access was enabled.\n\n" +
                        "Review the API data-handling details and verify what information types are authorized for GenAI.mil, including CUI, before use."
                    );
                }
            });

	            $('#allow-cdn-pulldowns').on('change', function () {
	                if (this && this.checked === true) {
	                    alert(
	                        "CDN pull-down allowlist was enabled.\n\n" +
	                        "Only allow approved CDN origins. These entries will be added to connect-src in the compiled output.\n\n" +
	                        "CDN requests with query params (?) and URL fragments (#) are blocked. Use path-based CDN URLs only."
	                    );
	                }
	            });
	
	            $('#sharepoint-compat-mode').on('change', function () {
	                const isEnabled = !!(this && this.checked);
	                const securityCheckbox = document.querySelector('#add-security-headers');
	                if (!securityCheckbox) return;
	                if (isEnabled) {
	                    securityCheckbox.dataset.disabledBySharePointCompat = securityCheckbox.checked ? 'true' : 'false';
	                    securityCheckbox.checked = false;
	                    return;
	                }
	                if (securityCheckbox.dataset.disabledBySharePointCompat === 'true') {
	                    securityCheckbox.checked = true;
	                }
	                delete securityCheckbox.dataset.disabledBySharePointCompat;
	            });
	
	            $('#compileButton').on('click', () => compiler.startCompilation({
                    saveToShippedApps: !!document.getElementById('compiler-save-to-shipped-apps')?.checked,
                    shipTarget: document.querySelector('#sharepoint-compat-mode')?.checked ? 'sharepoint' : 'offline',
                    shipReleaseType: compiler.getSelectedCompilerReleaseType()
                }));
                $('#compiler-live-preview-refresh').on('click', () => compiler.refreshLivePreview());
                $('#live-preview-quick-btn').on('click', () => compiler.refreshLivePreview({ switchToCompiler: true }));
                $('#compiler-live-preview-auto').on('change', function () {
                    try {
                        localStorage.setItem('forge:live-preview:auto', this.checked ? '1' : '0');
                    } catch (_) { }
                    compiler.setLivePreviewStatus(this.checked ? 'Live preview will refresh after saves.' : 'Auto-refresh off.', 'info');
                });
                document.addEventListener('forge:file-saved', () => {
                    const auto = document.getElementById('compiler-live-preview-auto');
                    const frame = document.getElementById('compiler-live-preview-frame');
                    if (auto && auto.checked && frame && frame.getAttribute('srcdoc')) {
                        compiler.refreshLivePreview({ skipSave: true, quiet: true });
                    }
                });
                const resetShipChoiceSubOptions = () => {
                    const offlineOptions = document.getElementById('ship-offline-options');
                    const sharePointOptions = document.getElementById('ship-sharepoint-options');
                    const fusionOptions = document.getElementById('ship-fusion-options');
                    const offlineBtn = document.getElementById('ship-offline-options-btn');
                    const sharePointBtn = document.getElementById('ship-sharepoint-btn');
                    const fusionBtn = document.getElementById('ship-fusion-options-btn');
                    if (offlineOptions) offlineOptions.classList.add('d-none');
                    if (sharePointOptions) sharePointOptions.classList.add('d-none');
                    if (fusionOptions) fusionOptions.classList.add('d-none');
                    if (offlineBtn) offlineBtn.setAttribute('aria-expanded', 'false');
                    if (sharePointBtn) sharePointBtn.setAttribute('aria-expanded', 'false');
                    if (fusionBtn) fusionBtn.setAttribute('aria-expanded', 'false');
                };
                const showShipChoiceSubOptions = (target) => {
                    const offlineOptions = document.getElementById('ship-offline-options');
                    const sharePointOptions = document.getElementById('ship-sharepoint-options');
                    const fusionOptions = document.getElementById('ship-fusion-options');
                    const offlineBtn = document.getElementById('ship-offline-options-btn');
                    const sharePointBtn = document.getElementById('ship-sharepoint-btn');
                    const fusionBtn = document.getElementById('ship-fusion-options-btn');
                    const showOffline = target === 'offline';
                    const showSharePoint = target === 'sharepoint';
                    const showFusion = target === 'fusion';
                    if (offlineOptions) offlineOptions.classList.toggle('d-none', !showOffline);
                    if (sharePointOptions) sharePointOptions.classList.toggle('d-none', !showSharePoint);
                    if (fusionOptions) fusionOptions.classList.toggle('d-none', !showFusion);
                    if (offlineBtn) offlineBtn.setAttribute('aria-expanded', showOffline ? 'true' : 'false');
                    if (sharePointBtn) sharePointBtn.setAttribute('aria-expanded', showSharePoint ? 'true' : 'false');
                    if (fusionBtn) fusionBtn.setAttribute('aria-expanded', showFusion ? 'true' : 'false');
                };
                const getSelectedFusionFullscreen = () => {
                    return document.getElementById('ship-fusion-fullscreen-toggle')?.checked !== false;
                };
	            $('#ship-now-btn').on('click', () => {
                    resetShipChoiceSubOptions();
	                const modal = new bootstrap.Modal(document.getElementById('ship-choice-modal'));
	                modal.show();
	            });
                document.getElementById('ship-choice-modal')?.addEventListener('show.bs.modal', () => {
                    resetShipChoiceSubOptions();
                    compiler.refreshShipVersionPreview().catch((err) => {
                        console.warn('Could not refresh ship version preview:', err);
                    });
                });
                document.querySelectorAll('input[name="ship-release-type"]').forEach((input) => {
                    input.addEventListener('change', () => {
                        compiler.refreshShipVersionPreview().catch((err) => {
                            console.warn('Could not refresh ship version preview:', err);
                        });
                    });
                });
                document.getElementById('compiler-tab')?.addEventListener('shown.bs.tab', () => {
                    compiler.refreshCompilerShipVersionPreview().catch((err) => {
                        console.warn('Could not refresh compiler ship version preview:', err);
                    });
                });
                document.querySelectorAll('input[name="compiler-release-type"]').forEach((input) => {
                    input.addEventListener('change', () => {
                        compiler.refreshCompilerShipVersionPreview().catch((err) => {
                            console.warn('Could not refresh compiler ship version preview:', err);
                        });
                    });
                });
                document.getElementById('compiler-save-to-shipped-apps')?.addEventListener('change', () => {
                    compiler.refreshCompilerShipVersionPreview().catch((err) => {
                        console.warn('Could not refresh compiler ship version preview:', err);
                    });
                });
	            $('#ship-offline-options-btn').on('click', () => {
                    showShipChoiceSubOptions('offline');
	            });
	            $('#ship-offline-btn').on('click', () => {
	                bootstrap.Modal.getInstance(document.getElementById('ship-choice-modal'))?.hide();
	                compiler.startCompilation({
	                    forceInlineCdnChecked: true,
	                    forceSecurityHeadersChecked: true,
                        forceAllowAskSageApi: false,
                        forceAllowGenAiMilApi: false,
                        forceApiAllowlistRaw: '',
                        forceSharePointCompatMode: false,
                        forceAllowCdnPulldowns: false,
                        forceCdnAllowlistRaw: '',
	                    navigateToCompilerOnWarning: false,
	                    navigateToCompilerOnMissing: false,
	                    saveToShippedApps: true,
                        shipTarget: 'offline',
                        shipReleaseType: compiler.getSelectedShipReleaseType()
	                });
	            });
	            $('#ship-offline-ai-btn').on('click', () => {
                    alert(
                        "Offline + AI APIs was selected.\n\n" +
                        "The shipped app remains a standalone offline package, but Forge will allow outbound connections to CAPRA/Ask Sage and GenAI.mil API domains. Verify the target tenant authorizes the data types your app sends before operational use."
                    );
	                bootstrap.Modal.getInstance(document.getElementById('ship-choice-modal'))?.hide();
	                compiler.startCompilation({
	                    forceInlineCdnChecked: true,
	                    forceSecurityHeadersChecked: true,
                        forceAllowAskSageApi: true,
                        forceAllowGenAiMilApi: true,
                        forceApiAllowlistRaw: '',
                        forceSharePointCompatMode: false,
                        forceAllowCdnPulldowns: false,
                        forceCdnAllowlistRaw: '',
	                    navigateToCompilerOnWarning: false,
	                    navigateToCompilerOnMissing: false,
	                    saveToShippedApps: true,
                        shipTarget: 'offline-ai',
                        shipReleaseType: compiler.getSelectedShipReleaseType()
	                });
	            });
	            $('#ship-sharepoint-btn').on('click', () => {
                    showShipChoiceSubOptions('sharepoint');
	            });
	            $('#ship-flankspeed-sharepoint-btn').on('click', () => {
	                bootstrap.Modal.getInstance(document.getElementById('ship-choice-modal'))?.hide();
	                compiler.startCompilation({
	                    forceInlineCdnChecked: true,
	                    forceNoSecurityHeaders: true,
	                    navigateToCompilerOnWarning: false,
	                    navigateToCompilerOnMissing: false,
	                    saveToShippedApps: true,
                        shipTarget: 'sharepoint',
                        shipReleaseType: compiler.getSelectedShipReleaseType()
	                });
	            });
	            $('#ship-fusion-options-btn').on('click', () => {
                    showShipChoiceSubOptions('fusion');
	            });
	            $('#ship-fusion-copy-paste-btn').on('click', () => {
	                bootstrap.Modal.getInstance(document.getElementById('ship-choice-modal'))?.hide();
                    const useFullscreen = getSelectedFusionFullscreen();
	                compiler.startCompilation({
	                    forceInlineCdnChecked: true,
	                    forceSecurityHeadersChecked: true,
	                    navigateToCompilerOnWarning: false,
	                    navigateToCompilerOnMissing: false,
	                    saveToShippedApps: true,
                        shipTarget: useFullscreen ? 'fusion-wiki-fullscreen' : 'fusion-wiki',
                        shipReleaseType: compiler.getSelectedShipReleaseType(),
                        copyToClipboardAfterSave: true,
                        fusionBridgeMode: !useFullscreen,
                        wrapFusionFullscreenIframe: useFullscreen,
                        shipSavedModalTitle: useFullscreen ? 'Fusion Copy/Paste Fullscreen Code Copied' : 'Fusion Copy/Paste Code Copied',
                        clipboardSuccessMessage: useFullscreen
                            ? 'The secured fullscreen shipped code has been copied to your clipboard. Paste it into the Fusion HTML macro.'
                            : 'The secured shipped code has been copied to your clipboard. Paste it into the Fusion HTML macro.',
                        clipboardFailureMessage: 'The shipped file was saved, but Forge could not copy it automatically. Open the saved file in Shipped Apps and copy the full HTML before pasting into Fusion.',
                        deploymentInstructionsHtml: `
                            <div class="p-3 mb-3" style="background:#102032; border:1px solid #2563eb; border-radius:10px;">
                                <div class="small text-uppercase text-info mb-2" style="letter-spacing:0.04em;">Copy/Paste Fusion Deployment</div>
                                <p class="mb-2">This deployment includes Forge's Fusion security wrapper.</p>
                                <p class="mb-2">Open Fusion at <a href="https://wiki.fusion.navy.mil/" target="_blank" rel="noopener noreferrer">https://wiki.fusion.navy.mil/</a>.</p>
                                <ol class="mb-0">
                                    <li>Create a space.</li>
                                    <li>Press <strong>Create</strong> at the top to create a page.</li>
                                    <li>Press the <strong>+</strong> sign.</li>
                                    <li>Select <strong>Other macros</strong>.</li>
                                    <li>Search for <strong>html</strong>.</li>
                                    <li>Select <strong>HTML</strong>.</li>
                                    <li>Press <strong>Insert</strong>.</li>
                                    <li>Paste the copied code into the HTML area.</li>
                                </ol>
                            </div>
                        `
	                });
	            });
	            $('#ship-fusion-firepit-btn').on('click', () => {
	                bootstrap.Modal.getInstance(document.getElementById('ship-choice-modal'))?.hide();
                    const useFullscreen = getSelectedFusionFullscreen();
	                compiler.startCompilation({
	                    forceInlineCdnChecked: true,
	                    forceNoSecurityHeaders: true,
	                    navigateToCompilerOnWarning: false,
	                    navigateToCompilerOnMissing: false,
	                    saveToShippedApps: true,
                        shipTarget: 'fusion-wiki-file-list',
                        shipReleaseType: compiler.getSelectedShipReleaseType(),
                        copyToClipboardAfterSave: true,
                        copyFusionFirepitRendererToClipboard: true,
                        fusionFirepitFullscreen: useFullscreen,
                        shipSavedModalTitle: useFullscreen ? 'Fusion Firepit Fullscreen Renderer Copied' : 'Fusion Firepit Renderer Copied',
                        clipboardSuccessMessage: 'The fusion-firepit.html renderer has been copied to your clipboard. Paste it into the Fusion HTML macro, then point it at the uploaded app HTML file.',
                        clipboardFailureMessage: 'The shipped app file was saved, and Forge generated the Firepit renderer, but the browser blocked clipboard copy. Try Firepit Deployment again or use a browser context that permits clipboard access.',
                        deploymentInstructionsHtml: `
                            <div class="p-3 mb-3" style="background:#102032; border:1px solid #2563eb; border-radius:10px;">
                                <div class="small text-uppercase text-info mb-2" style="letter-spacing:0.04em;">Fusion Firepit Deployment</div>
                                <p class="mb-2">This deployment saves the app without Forge's copy/paste security wrapper because Firepit provides the host security wrapper.</p>
                                <p class="mb-2">Open Fusion at <a href="https://wiki.fusion.navy.mil/" target="_blank" rel="noopener noreferrer">https://wiki.fusion.navy.mil/</a>.</p>
                                <ol class="mb-0">
                                    <li>Upload the saved <code>-current.html</code> app file from <code>Shipped Apps</code> to the Fusion page file list or attachments area.</li>
                                    <li>Copy the uploaded HTML file link.</li>
                                    <li>Add an <strong>HTML</strong> macro to the Fusion page.</li>
                                    <li>Paste the copied <code>fusion-firepit.html</code> renderer code into the HTML macro and publish the page.</li>
                                    <li>Paste the uploaded app file link into <strong>HTML file link</strong> and press <strong>Render App</strong>. The renderer saves the link in the page properties.</li>
                                </ol>
                            </div>
                        `
	                });
	            });
	            $('#ship-legacy-sharepoint-btn').on('click', () => {
	                bootstrap.Modal.getInstance(document.getElementById('ship-choice-modal'))?.hide();
	                compiler.startCompilation({
	                    forceInlineCdnChecked: true,
	                    forceNoSecurityHeaders: true,
	                    navigateToCompilerOnWarning: false,
	                    navigateToCompilerOnMissing: false,
	                    saveToShippedApps: true,
                        shipTarget: 'legacy-sharepoint',
                        shipReleaseType: compiler.getSelectedShipReleaseType()
	                });
	            });
                document.querySelectorAll('input[name="sharepoint-deploy-mode"]').forEach((input) => {
                    input.addEventListener('change', () => {
                        if (typeof compiler !== 'undefined' && compiler && typeof compiler.refreshSharePointDeployModeUi === 'function') {
                            compiler.refreshSharePointDeployModeUi();
                        }
                    });
                });
                $('#deployToSharePointButton').on('click', () => compiler.startSharePointDeployment());
                $('#deploy-sharepoint-btn').on('click', () => compiler.startSharePointDeployment({
                    forceInlineCdnChecked: true,
                    forceNoSecurityHeaders: true,
                    navigateToCompilerOnWarning: false,
                    navigateToCompilerOnMissing: false
                }));
            $(document).on('click', '[data-action="start-product-tour"]', () => {
                try {
                    localStorage.setItem('wc:productTourClicked:v1', '1');
                } catch (_) {
                    // no-op
                }
                if (window.forgeBeginnerOnboardingState) {
                    window.forgeBeginnerOnboardingState.productTourClicked = true;
                }
                $('#resources-menu-btn').removeClass('forge-product-tour-flash');
                if (typeof newProjectWalkthrough !== 'undefined') {
                    newProjectWalkthrough.start();
                }
            });
            $('#hash-verify-btn').on('click', () => compiler.verifyUploadedHash());
            $(document).on('click', '[data-action="load-directory"]', e => {
                e.preventDefault();
                loadFolder.getFile();
            });
            $(document).on('click', '[data-action="create-root-file"]', e => {
                e.preventDefault();
                loadFolder.showCreateFileDialog();
            });
            $(document).on('click', '[data-ai-action]', function (e) {
                e.preventDefault();
                const step = Number(this.getAttribute('data-ai-step'));
                if (!step) return;
                const action = this.getAttribute('data-ai-action');
                if (action === 'generate' && typeof aiHelper !== 'undefined') {
                    aiHelper.generate(step);
                } else if (action === 'copy' && typeof aiHelper !== 'undefined') {
                    aiHelper.copyPrompt(step, this);
                } else if (action === 'generate-and-copy' && typeof aiHelper !== 'undefined') {
                    aiHelper.generateAndCopy(step, this);
                }
            });

            if (typeof aiHelper !== 'undefined' && aiHelper && typeof aiHelper.init === 'function') {
                aiHelper.init();
            }

            // Search bindings
            $(document).on('click', '[data-sidebar-mode]', function () {
                const mode = this.getAttribute('data-sidebar-mode');
                search.setSidebarMode(mode, { focus: mode === 'search' });
            });
            $(document).on('click', '#global-search-btn', () => search.run());
            $(document).on('keydown', '#global-search-input', e => { if (e.key === 'Enter') { e.preventDefault(); search.run(); } });
            $(document).on('click', '#global-replace-btn', () => search.replaceAll());
            $(document).on('keydown', '#global-replace-input', e => { if (e.key === 'Enter') { e.preventDefault(); search.replaceAll(); } });
            $(document).on('click', '#search-clear', () => search.clear());
            $(document).on('keydown', e => {
                if ((e.ctrlKey || e.metaKey) && e.shiftKey && String(e.key || '').toLowerCase() === 'f') {
                    if (!loadFolder.fileHandle) return;
                    e.preventDefault();
                    search.setSidebarMode('search', { focus: true });
                }
            });
            search.init();

            // Tab modules
            if (typeof testRecorderTab !== 'undefined') testRecorderTab.init();
            if (typeof advancedDebugTab !== 'undefined') advancedDebugTab.init();
            if (typeof mathLogicTesterTab !== 'undefined') mathLogicTesterTab.init();
            sharedriveNosqlTab.init();
            if (typeof leafletMapTab !== 'undefined') leafletMapTab.init();
            if (typeof sastTab !== 'undefined') sastTab.init();
            if (typeof securityReviewer !== 'undefined') securityReviewer.init();
            if (typeof newProjectWalkthrough !== 'undefined') newProjectWalkthrough.init();

            // Enable bootstrap tooltip for the Tour button
            (function enableTourTooltip() {
                const tourFab = document.getElementById('tour-fab');
                if (tourFab && window.bootstrap && bootstrap.Tooltip) {
                    bootstrap.Tooltip.getOrCreateInstance(tourFab, { placement: 'left' });
                }
            })();

            // Clear heavy tab UIs when navigating away
            $('a[data-bs-toggle="tab"]').on('hide.bs.tab', function (e) {
                if (e.target.id === 'ai-helper-tab') {
                    aiHelper.clear();
                } else if (e.target.id === 'sast-tab') {
                    sastTab.clearTransient();
                } else if (e.target.id === 'security-reviewer-tab') {
                    securityReviewer.clearTransient();
                }
            });

            // Save shortcuts
            document.addEventListener('keydown', e => {
                if (e.ctrlKey && e.key === 's') {
                    e.preventDefault();
                    if (loadFolder.fileHandle) editor.saveAll();
                }
                if (e.ctrlKey && e.shiftKey && (e.key === 'S' || e.key === 's')) {
                    e.preventDefault();
                    if (loadFolder.fileHandle) editor.saveCurrent();
                }
            });

            // Hide right-click context menus when clicking elsewhere
            document.addEventListener('click', () => {
                loadFolder.hideContextMenus();
            });

            // Warn about unsaved changes on browser close
            window.addEventListener('beforeunload', (e) => {
                if (editor.dirtyFiles && editor.dirtyFiles.size > 0) {
                    e.preventDefault();
                    e.returnValue = 'You have unsaved changes. Are you sure you want to leave?';
                    return e.returnValue;
                }
            });

            if (typeof compiler !== 'undefined' && compiler && typeof compiler.syncSharePointDeployUi === 'function') {
                compiler.syncSharePointDeployUi();
            }
            if (typeof compiler !== 'undefined' && compiler && typeof compiler.initLivePreviewUi === 'function') {
                compiler.initLivePreviewUi();
            }
        });

        // Show search UI and enable tabs once a directory is loaded
        (function () {
            const $startupGate = $('#startup-gate');
            const $startupLoadDirectoryBtn = $('#startup-load-directory-btn');
            const $startupInstructionsPanel = $('#startup-instructions-panel');
            const $loadDirectoryBtn = $('#load-directory-btn');
            const $quickActions = $('#editor-quick-actions');
            const $helpActions = $('#help-actions');
            const $productTourTopBtn = $('#resources-menu-btn');
            const PRODUCT_TOUR_CLICKED_KEY = 'wc:productTourClicked:v1';
            const productTourClicked = (() => {
                try {
                    return localStorage.getItem(PRODUCT_TOUR_CLICKED_KEY) === '1';
                } catch (_) {
                    return false;
                }
            })();
            const onboardingState = window.forgeBeginnerOnboardingState || (window.forgeBeginnerOnboardingState = {
                beginnerMode: false,
                awaitingAiReturn: false,
                leftAfterAi: false,
                aiReturnShownForLaunch: false,
                awaitingBeginnerPaste: false,
                awaitingProductTourAfterSave: false,
                awaitingProductTourReturnAfterSave: false,
                leftAfterSaveForRunTest: false,
                shipPromptShown: false,
                quickPromptsHintShown: false,
                productTourClicked: productTourClicked
            });
            if (typeof onboardingState.awaitingBeginnerPaste !== 'boolean') {
                onboardingState.awaitingBeginnerPaste = false;
            }
            if (typeof onboardingState.shipPromptShown !== 'boolean') {
                onboardingState.shipPromptShown = !!onboardingState.productTourPromptShown;
            }
            if (typeof onboardingState.quickPromptsHintShown !== 'boolean') {
                onboardingState.quickPromptsHintShown = false;
            }
            let productTourNudgeTimer = null;
            let beginnerPasteHighlightTimer = null;
            let beginnerPastePopoverEl = null;
            let beginnerInlineCueEl = null;
            let beginnerInlineCueActive = false;
            let beginnerInlineCueAnchorResolver = null;
            let beginnerInlineCueHighlightedEls = [];
            let beginnerInlineCueFlashEl = null;
            let beginnerInlineCueAutoHideTimer = null;

            const clearBeginnerPasteHighlight = () => {
                if (beginnerPasteHighlightTimer) {
                    clearTimeout(beginnerPasteHighlightTimer);
                    beginnerPasteHighlightTimer = null;
                }
                document.querySelectorAll('.forge-beginner-paste-highlight').forEach(el => {
                    el.classList.remove('forge-beginner-paste-highlight');
                });
            };

            const getBeginnerPasteTargets = () => {
                const targets = [];
                const activeEditorPane = document.querySelector('#editor .editor[style*="display: block"]')
                    || document.querySelector('#editor .editor:not([style*="display: none"])');
                if (activeEditorPane) targets.push(activeEditorPane);
                const cmInActivePane = activeEditorPane ? activeEditorPane.querySelector('.cm-editor') : null;
                if (cmInActivePane) targets.push(cmInActivePane);
                const anyCmEditor = document.querySelector('#editor .cm-editor');
                if (anyCmEditor) targets.push(anyCmEditor);
                if (!targets.length) {
                    const editorHost = document.getElementById('editor');
                    if (editorHost) targets.push(editorHost);
                }
                return Array.from(new Set(targets.filter(Boolean)));
            };

            const getBeginnerPasteAnchor = () => {
                const activeCmEditor = document.querySelector('#editor .editor[style*="display: block"] .cm-editor')
                    || document.querySelector('#editor .editor:not([style*="display: none"]) .cm-editor')
                    || document.querySelector('#editor .cm-editor');
                if (activeCmEditor) return activeCmEditor;
                return document.getElementById('editor') || null;
            };

            const ensureBeginnerPastePopover = () => {
                if (beginnerPastePopoverEl && document.body.contains(beginnerPastePopoverEl)) {
                    return beginnerPastePopoverEl;
                }
                beginnerPastePopoverEl = document.createElement('div');
                beginnerPastePopoverEl.id = 'forge-beginner-paste-popover';
                beginnerPastePopoverEl.className = 'forge-beginner-paste-popover';
                beginnerPastePopoverEl.textContent = 'Paste AI code into index.html here.';
                document.body.appendChild(beginnerPastePopoverEl);
                return beginnerPastePopoverEl;
            };

            const positionBeginnerPastePopover = () => {
                const popoverEl = ensureBeginnerPastePopover();
                if (!popoverEl) return;

                const anchor = getBeginnerPasteAnchor();
                if (!anchor) {
                    popoverEl.classList.remove('is-visible', 'is-below');
                    return;
                }

                const rect = anchor.getBoundingClientRect();
                if ((!rect.width && !rect.height) || rect.bottom < 0 || rect.top > window.innerHeight) {
                    popoverEl.classList.remove('is-visible', 'is-below');
                    return;
                }

                popoverEl.classList.add('is-visible');
                const popW = popoverEl.offsetWidth || 280;
                const popH = popoverEl.offsetHeight || 36;
                const gutter = 10;
                const viewportW = window.innerWidth || document.documentElement.clientWidth || 1280;

                const idealLeft = rect.left + (rect.width / 2) - (popW / 2);
                const minLeft = 8;
                const maxLeft = Math.max(8, viewportW - popW - 8);
                const left = Math.min(Math.max(idealLeft, minLeft), maxLeft);

                let top = rect.top - popH - gutter;
                let useBelow = false;
                if (top < 8) {
                    useBelow = true;
                    top = rect.bottom + gutter;
                }

                popoverEl.style.left = `${Math.round(left)}px`;
                popoverEl.style.top = `${Math.round(top)}px`;
                popoverEl.classList.toggle('is-below', useBelow);
            };

            const ensureBeginnerInlineCue = () => {
                if (beginnerInlineCueEl && document.body.contains(beginnerInlineCueEl)) {
                    return beginnerInlineCueEl;
                }
                beginnerInlineCueEl = document.createElement('div');
                beginnerInlineCueEl.id = 'forge-beginner-inline-cue';
                beginnerInlineCueEl.className = 'forge-beginner-inline-cue';
                beginnerInlineCueEl.innerHTML = `
                    <div class="forge-beginner-inline-cue-body"></div>
                    <button type="button" class="forge-beginner-inline-cue-close" aria-label="Close tip" title="Close">x</button>
                `;
                const closeBtn = beginnerInlineCueEl.querySelector('.forge-beginner-inline-cue-close');
                if (closeBtn) {
                    closeBtn.addEventListener('click', (event) => {
                        event.preventDefault();
                        event.stopPropagation();
                        hideBeginnerInlineCue();
                    });
                }
                document.body.appendChild(beginnerInlineCueEl);
                return beginnerInlineCueEl;
            };

            const clearBeginnerInlineCueHighlights = () => {
                beginnerInlineCueHighlightedEls.forEach(el => {
                    if (!el) return;
                    el.classList.remove('forge-walkthrough-highlight');
                });
                beginnerInlineCueHighlightedEls = [];
                if (beginnerInlineCueFlashEl) {
                    beginnerInlineCueFlashEl.classList.remove('forge-product-tour-flash');
                }
                beginnerInlineCueFlashEl = null;
            };

            const hideBeginnerInlineCue = () => {
                beginnerInlineCueActive = false;
                beginnerInlineCueAnchorResolver = null;
                if (beginnerInlineCueAutoHideTimer) {
                    clearTimeout(beginnerInlineCueAutoHideTimer);
                    beginnerInlineCueAutoHideTimer = null;
                }
                clearBeginnerInlineCueHighlights();
                if (beginnerInlineCueEl) {
                    beginnerInlineCueEl.classList.remove('is-visible', 'is-below');
                }
            };

            const _resolveBeginnerInlineCueAnchor = () => {
                if (!beginnerInlineCueAnchorResolver) return null;
                try {
                    const anchor = beginnerInlineCueAnchorResolver();
                    return anchor || null;
                } catch (_) {
                    return null;
                }
            };

            const positionBeginnerInlineCue = () => {
                if (!beginnerInlineCueActive) {
                    if (beginnerInlineCueEl) {
                        beginnerInlineCueEl.classList.remove('is-visible', 'is-below');
                    }
                    return;
                }
                const cueEl = ensureBeginnerInlineCue();
                if (!cueEl) return;

                const anchor = _resolveBeginnerInlineCueAnchor();
                if (!anchor) {
                    cueEl.classList.remove('is-visible', 'is-below');
                    return;
                }

                const rect = anchor.getBoundingClientRect();
                if ((!rect.width && !rect.height) || rect.bottom < 0 || rect.top > window.innerHeight) {
                    cueEl.classList.remove('is-visible', 'is-below');
                    return;
                }

                cueEl.classList.add('is-visible');
                const cueW = cueEl.offsetWidth || 340;
                const cueH = cueEl.offsetHeight || 46;
                const gutter = 10;
                const viewportW = window.innerWidth || document.documentElement.clientWidth || 1280;

                const idealLeft = rect.left + (rect.width / 2) - (cueW / 2);
                const minLeft = 8;
                const maxLeft = Math.max(8, viewportW - cueW - 8);
                const left = Math.min(Math.max(idealLeft, minLeft), maxLeft);

                let top = rect.top - cueH - gutter;
                let useBelow = false;
                if (top < 8) {
                    useBelow = true;
                    top = rect.bottom + gutter;
                }

                cueEl.style.left = `${Math.round(left)}px`;
                cueEl.style.top = `${Math.round(top)}px`;
                cueEl.classList.toggle('is-below', useBelow);
            };

            const showBeginnerInlineCue = ({ html, anchor, highlight = [], flash = null, hideAfterMs = 0 } = {}) => {
                const cueEl = ensureBeginnerInlineCue();
                if (!cueEl) return;

                beginnerInlineCueActive = true;
                if (beginnerInlineCueAutoHideTimer) {
                    clearTimeout(beginnerInlineCueAutoHideTimer);
                    beginnerInlineCueAutoHideTimer = null;
                }
                const cueBodyEl = cueEl.querySelector('.forge-beginner-inline-cue-body');
                if (cueBodyEl) {
                    cueBodyEl.innerHTML = html || '';
                } else {
                    cueEl.innerHTML = html || '';
                }

                if (typeof anchor === 'function') {
                    beginnerInlineCueAnchorResolver = anchor;
                } else if (anchor && typeof anchor === 'string') {
                    beginnerInlineCueAnchorResolver = () => document.querySelector(anchor);
                } else if (anchor && anchor.nodeType === 1) {
                    beginnerInlineCueAnchorResolver = () => anchor;
                } else {
                    beginnerInlineCueAnchorResolver = null;
                }

                clearBeginnerInlineCueHighlights();
                const highlightList = Array.isArray(highlight) ? highlight : [highlight];
                beginnerInlineCueHighlightedEls = Array.from(new Set(highlightList.filter(Boolean)));
                beginnerInlineCueHighlightedEls.forEach(el => el.classList.add('forge-walkthrough-highlight'));

                if (flash) {
                    beginnerInlineCueFlashEl = flash;
                    beginnerInlineCueFlashEl.classList.add('forge-product-tour-flash');
                }

                positionBeginnerInlineCue();

                const delay = Number(hideAfterMs || 0);
                if (delay > 0) {
                    beginnerInlineCueAutoHideTimer = setTimeout(() => {
                        hideBeginnerInlineCue();
                    }, delay);
                }
            };

            window.forgeBeginnerInlineHint = {
                show: ({
                    html = '',
                    anchor = null,
                    targetSelector = '',
                    targetEl = null,
                    highlight = [],
                    highlightSelectors = [],
                    flash = null,
                    flashSelector = '',
                    hideAfterMs = 0
                } = {}) => {
                    let nextAnchor = anchor;
                    if (!nextAnchor) {
                        if (targetEl && targetEl.nodeType === 1) {
                            nextAnchor = targetEl;
                        } else if (targetSelector) {
                            nextAnchor = () => document.querySelector(targetSelector);
                        }
                    }

                    const highlightEls = [];
                    const rawHighlightList = Array.isArray(highlight) ? highlight : [highlight];
                    rawHighlightList.forEach(item => {
                        if (item && item.nodeType === 1) highlightEls.push(item);
                    });
                    const selectorList = Array.isArray(highlightSelectors) ? highlightSelectors : [highlightSelectors];
                    selectorList
                        .map(sel => String(sel || '').trim())
                        .filter(Boolean)
                        .forEach(sel => {
                            const el = document.querySelector(sel);
                            if (el) highlightEls.push(el);
                        });

                    let flashEl = flash;
                    if (!(flashEl && flashEl.nodeType === 1) && flashSelector) {
                        flashEl = document.querySelector(flashSelector);
                    }

                    showBeginnerInlineCue({
                        html,
                        anchor: nextAnchor,
                        highlight: highlightEls,
                        flash: flashEl && flashEl.nodeType === 1 ? flashEl : null,
                        hideAfterMs
                    });
                },
                hide: () => {
                    hideBeginnerInlineCue();
                },
                reposition: () => {
                    positionBeginnerInlineCue();
                }
            };

            const hideBeginnerPasteCue = () => {
                clearBeginnerPasteHighlight();
                if (beginnerPastePopoverEl) {
                    beginnerPastePopoverEl.classList.remove('is-visible', 'is-below');
                }
            };

            const cueBeginnerPasteTarget = ({ sticky = false } = {}) => {
                clearBeginnerPasteHighlight();
                const targets = getBeginnerPasteTargets();
                if (!targets.length) {
                    if (sticky) {
                        positionBeginnerPastePopover();
                    }
                    return;
                }

                try {
                    targets[0].scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'nearest' });
                } catch (_) {
                    // no-op
                }

                targets.forEach(el => el.classList.add('forge-beginner-paste-highlight'));
                positionBeginnerPastePopover();

                if (!sticky) {
                    beginnerPasteHighlightTimer = setTimeout(() => {
                        targets.forEach(el => el.classList.remove('forge-beginner-paste-highlight'));
                        beginnerPasteHighlightTimer = null;
                    }, 9000);
                }
            };

            const focusIndexHtmlEditorForBeginner = async () => {
                activateCodeEditorTab();
                if (!(window.editor && typeof editor.openFile === 'function')) {
                    return false;
                }

                const isIndexName = (name) => String(name || '').trim().toLowerCase() === 'index.html';
                const focusActiveView = () => {
                    if (typeof editor.getActiveUuid !== 'function') return;
                    const activeUuid = editor.getActiveUuid();
                    const activeView = activeUuid && editor.instance ? editor.instance[activeUuid] : null;
                    if (activeView && typeof activeView.focus === 'function') {
                        activeView.focus();
                    }
                };

                try {
                    if (typeof editor.getActiveUuid === 'function') {
                        const activeUuid = editor.getActiveUuid();
                        const activeMeta = activeUuid && editor._meta ? editor._meta[activeUuid] : null;
                        if (isIndexName(activeMeta && activeMeta.name)) {
                            focusActiveView();
                            return true;
                        }
                    }

                    const openIndexUuid = Object.keys(editor._meta || {}).find(uuid =>
                        isIndexName(editor._meta[uuid] && editor._meta[uuid].name)
                    );
                    if (openIndexUuid) {
                        await editor.openFile(openIndexUuid);
                        focusActiveView();
                        return true;
                    }

                    const files = Array.isArray(loadFolder && loadFolder.fileStructure) ? loadFolder.fileStructure : [];
                    const rootIndex = files.find(file =>
                        file &&
                        isIndexName(file.name) &&
                        Array.isArray(file.path) &&
                        file.path.length === 0
                    );
                    const anyIndex = rootIndex || files.find(file => file && isIndexName(file.name));
                    if (anyIndex && anyIndex.uuid) {
                        await editor.openFile(anyIndex.uuid);
                        focusActiveView();
                        return true;
                    }

                    const treeIndexItem = Array.from(document.querySelectorAll('#file-tree li.file[data-uuid]'))
                        .find(el => isIndexName(el.querySelector('.file-label') && el.querySelector('.file-label').textContent));
                    if (treeIndexItem) {
                        treeIndexItem.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
                        await new Promise(resolve => setTimeout(resolve, 80));
                        focusActiveView();
                        return true;
                    }
                } catch (_) {
                    return false;
                }

                return false;
            };

            const prepareBeginnerPasteTarget = ({ sticky = false } = {}) => {
                Promise.resolve()
                    .then(async () => {
                        const shouldShowCue = await shouldShowBeginnerPasteCue();
                        if (!shouldShowCue) {
                            return false;
                        }
                        await focusIndexHtmlEditorForBeginner();
                        return true;
                    })
                    .then((shouldCue) => {
                        if (!shouldCue) return;
                        setTimeout(() => {
                            cueBeginnerPasteTarget({ sticky });
                            if (beginnerInlineCueActive) {
                                positionBeginnerInlineCue();
                            }
                        }, 120);
                    })
                    .catch(() => false);
            };

            const clearProductTourNudge = () => {
                if (productTourNudgeTimer) {
                    clearTimeout(productTourNudgeTimer);
                    productTourNudgeTimer = null;
                }
                $productTourTopBtn.removeClass('forge-product-tour-flash');
            };

            const markProductTourClicked = () => {
                onboardingState.productTourClicked = true;
                try {
                    localStorage.setItem(PRODUCT_TOUR_CLICKED_KEY, '1');
                } catch (_) {
                    // no-op
                }
                clearProductTourNudge();
            };

            const showProductTourNudgeForNonBeginners = () => {
                if (onboardingState.beginnerMode || onboardingState.productTourClicked) return;
                if (!$productTourTopBtn.length) return;
                clearProductTourNudge();
                $productTourTopBtn.addClass('forge-product-tour-flash');
                productTourNudgeTimer = setTimeout(() => {
                    $productTourTopBtn.removeClass('forge-product-tour-flash');
                    productTourNudgeTimer = null;
                }, 12000);
            };

            const showBeginnerReturnCue = () => {
                if (!onboardingState.beginnerMode || !onboardingState.awaitingAiReturn || !onboardingState.leftAfterAi) {
                    return;
                }
                if (onboardingState.aiReturnShownForLaunch) {
                    return;
                }
                onboardingState.aiReturnShownForLaunch = true;
                Promise.resolve()
                    .then(async () => {
                        const shouldShowPasteCue = await shouldShowBeginnerPasteCue();
                        onboardingState.awaitingAiReturn = false;
                        onboardingState.leftAfterAi = false;
                        if (!shouldShowPasteCue) {
                            onboardingState.awaitingBeginnerPaste = false;
                            onboardingState.awaitingProductTourAfterSave = false;
                            hideBeginnerPasteCue();
                            hideBeginnerInlineCue();
                            return;
                        }

                        onboardingState.awaitingBeginnerPaste = true;
                        onboardingState.awaitingProductTourAfterSave = true;
                        prepareBeginnerPasteTarget({ sticky: true });
                        showBeginnerInlineCue({
                            html: '<strong>Next:</strong> paste AI output into <code>index.html</code>, then save in Forge (<code>Ctrl+S</code>).',
                            anchor: () => getBeginnerPasteAnchor() || document.getElementById('editor'),
                            highlight: getBeginnerPasteTargets()
                        });
                    })
                    .catch(() => {
                        onboardingState.awaitingAiReturn = false;
                        onboardingState.leftAfterAi = false;
                    });
            };

            const showShipPromptCue = () => {
                if (onboardingState.shipPromptShown) return;
                const shipNowBtn = document.getElementById('ship-now-btn');
                if (!shipNowBtn) {
                    return;
                }
                onboardingState.shipPromptShown = true;
                activateCodeEditorTab();
                showBeginnerInlineCue({
                    html: '<strong>Next:</strong> click <strong>Ship</strong> here before CUI entry or sharing. Ship adds security headers like <code>CSP</code>, <code>X-Content-Type-Options</code>, <code>Referrer-Policy</code>, and <code>Permissions-Policy</code>.',
                    anchor: () => document.getElementById('ship-now-btn'),
                    highlight: [shipNowBtn],
                    flash: shipNowBtn
                });
            };

            const showQuickPromptsFeatureCue = () => {
                const quickPromptsBtn = document.getElementById('quick-prompts-btn');
                if (!quickPromptsBtn) return;
                onboardingState.quickPromptsHintShown = true;
                activateCodeEditorTab();
                setTimeout(() => {
                    const btn = document.getElementById('quick-prompts-btn');
                    if (!btn) return;
                    showBeginnerInlineCue({
                        html: '<strong>Next:</strong> to add a new feature, use <strong>Quick Prompts</strong> -> <strong>Edit Code</strong>, describe one feature, then paste that prompt into AI.',
                        anchor: () => document.getElementById('quick-prompts-btn') || document.getElementById('editor-header'),
                        highlight: [btn],
                        flash: btn,
                        hideAfterMs: 12000
                    });
                }, 140);
            };

            const showRunTestCue = () => {
                const saveBtn = document.getElementById('saveButton');
                showBeginnerInlineCue({
                    html: '<strong>Saved.</strong> Now run outside Forge: open <strong>File Explorer</strong> on your computer, go to your folder, then double-click <code>index.html</code>.',
                    anchor: () => saveBtn || document.getElementById('editor-header') || document.getElementById('editor'),
                    highlight: saveBtn ? [saveBtn] : []
                });
            };

            const markBeginnerPasteCompleted = () => {
                if (!onboardingState.awaitingBeginnerPaste) return;
                onboardingState.awaitingBeginnerPaste = false;
                hideBeginnerPasteCue();
                hideBeginnerInlineCue();
                if (typeof window.hideForgeBeginnerProviderTip === 'function') {
                    window.hideForgeBeginnerProviderTip();
                }
            };

            const _looksPopulatedHtml = (text) => {
                const t = String(text || '').trim();
                if (!t) return false;
                const tagMatches = t.match(/<\s*\/?\s*[a-z!][^>]*>/gi) || [];
                if (tagMatches.length < 2) return false;
                return /<\s*(!doctype\s+html|html|head|body|main|section|article|div|script|style|p|h1|h2|table|form|ul|ol|li)\b/i.test(t);
            };

            const _readIndexHtmlTextForBeginnerCue = async () => {
                const isIndexName = (name) => String(name || '').trim().toLowerCase() === 'index.html';
                const metaByUuid = window.editor && editor._meta ? editor._meta : {};
                const addCandidate = (list, seen, uuid) => {
                    const id = String(uuid || '').trim();
                    if (!id || seen.has(id)) return;
                    seen.add(id);
                    list.push(id);
                };

                const readEditorCandidate = async (uuid) => {
                    const meta = metaByUuid && metaByUuid[uuid] ? metaByUuid[uuid] : null;
                    if (!meta || !isIndexName(meta.name)) return '';

                    if (typeof editor._getValue === 'function') {
                        const liveText = String(editor._getValue(uuid) || '');
                        if (liveText.trim()) return liveText;
                    } else {
                        const view = editor.instance ? editor.instance[uuid] : null;
                        const viewText = view && view.state && view.state.doc ? String(view.state.doc.toString() || '') : '';
                        if (viewText.trim()) return viewText;
                    }

                    const cachedText = String(meta.text || '');
                    if (cachedText.trim()) return cachedText;

                    try {
                        if (meta.entry && typeof meta.entry.getFile === 'function') {
                            const fh = await meta.entry.getFile();
                            const diskText = String(await fh.text() || '');
                            if (diskText.trim()) return diskText;
                        }
                    } catch (_) {
                        // no-op
                    }

                    return '';
                };

                if (window.editor) {
                    const candidateUuids = [];
                    const seen = new Set();

                    if (typeof editor.getActiveUuid === 'function') {
                        const activeUuid = editor.getActiveUuid();
                        const activeMeta = activeUuid ? metaByUuid[activeUuid] : null;
                        if (activeMeta && isIndexName(activeMeta.name)) {
                            addCandidate(candidateUuids, seen, activeUuid);
                        }
                    }

                    Object.keys(metaByUuid).forEach(uuid => {
                        const meta = metaByUuid[uuid];
                        if (!meta || !isIndexName(meta.name)) return;
                        if (Array.isArray(meta.path) && meta.path.length === 0) {
                            addCandidate(candidateUuids, seen, uuid);
                        }
                    });

                    Object.keys(metaByUuid).forEach(uuid => {
                        const meta = metaByUuid[uuid];
                        if (!meta || !isIndexName(meta.name)) return;
                        addCandidate(candidateUuids, seen, uuid);
                    });

                    for (const uuid of candidateUuids) {
                        const text = await readEditorCandidate(uuid);
                        if (text.trim()) {
                            return text;
                        }
                    }
                }

                if (!(window.loadFolder && Array.isArray(loadFolder.fileStructure))) {
                    return '';
                }

                const files = loadFolder.fileStructure || [];
                const rootIndex = files.find(file =>
                    file &&
                    file.kind === 'file' &&
                    isIndexName(file.name) &&
                    Array.isArray(file.path) &&
                    file.path.length === 0
                );
                const anyIndex = rootIndex || files.find(file => file && file.kind === 'file' && isIndexName(file.name));
                if (!anyIndex) return '';

                if (typeof loadFolder.getFileContent === 'function') {
                    try {
                        return await loadFolder.getFileContent(anyIndex);
                    } catch (_) {
                        return '';
                    }
                }

                return '';
            };

            const shouldShowBeginnerPasteCue = async () => {
                const htmlText = await _readIndexHtmlTextForBeginnerCue();
                if (_looksPopulatedHtml(htmlText)) {
                    onboardingState.awaitingBeginnerPaste = false;
                    hideBeginnerPasteCue();
                    hideBeginnerInlineCue();
                    return false;
                }
                return true;
            };

            const showPreloadActions = () => {
                $startupGate.addClass('is-visible');
                $startupLoadDirectoryBtn.addClass('forge-load-directory-flash');
                $loadDirectoryBtn.hide().removeClass('forge-load-directory-flash');
                $quickActions.hide();
                $helpActions.hide();
                hideBeginnerInlineCue();
                onboardingState.quickPromptsHintShown = false;
                clearProductTourNudge();
            };

            const showLoadedActions = () => {
                $startupGate.removeClass('is-visible');
                $startupLoadDirectoryBtn.removeClass('forge-load-directory-flash');
                $startupInstructionsPanel.removeClass('is-visible');
                $loadDirectoryBtn.hide().removeClass('forge-load-directory-flash');
                $quickActions.show();
                $helpActions.css('display', 'flex');
                showProductTourNudgeForNonBeginners();
                // Auto-open Prompt Lab when entering the editor
                if (typeof window.toggleRightPanel === 'function') window.toggleRightPanel(true);
                if (typeof promptLab !== 'undefined' && promptLab.switchRpTab) promptLab.switchRpTab('prompt-lab');
            };

            const activateCodeEditorTab = () => {
                const editorTab = document.getElementById('editor-tab');
                if (!editorTab) return;
                if (window.bootstrap && bootstrap.Tab) {
                    bootstrap.Tab.getOrCreateInstance(editorTab).show();
                    return;
                }
                editorTab.click();
            };

            $startupLoadDirectoryBtn.on('click', () => {
                onboardingState.beginnerMode = true;
                onboardingState.shipPromptShown = false;
                onboardingState.quickPromptsHintShown = false;
            });

            document.addEventListener('visibilitychange', () => {
                if (!onboardingState.beginnerMode) return;

                if (onboardingState.awaitingProductTourReturnAfterSave) {
                    if (document.hidden) {
                        onboardingState.leftAfterSaveForRunTest = true;
                        hideBeginnerInlineCue();
                        return;
                    }
                    if (onboardingState.leftAfterSaveForRunTest) {
                        onboardingState.awaitingProductTourReturnAfterSave = false;
                        onboardingState.leftAfterSaveForRunTest = false;
                        showShipPromptCue();
                    }
                    return;
                }

                if (onboardingState.awaitingBeginnerPaste) {
                    if (document.hidden) return;
                    prepareBeginnerPasteTarget({ sticky: true });
                    return;
                }

                if (!onboardingState.awaitingAiReturn) return;
                if (document.hidden) {
                    onboardingState.leftAfterAi = true;
                    hideBeginnerPasteCue();
                    hideBeginnerInlineCue();
                    return;
                }
                showBeginnerReturnCue();
            });

            window.addEventListener('focus', () => {
                if (!onboardingState.beginnerMode) return;
                if (onboardingState.awaitingProductTourReturnAfterSave && onboardingState.leftAfterSaveForRunTest) {
                    onboardingState.awaitingProductTourReturnAfterSave = false;
                    onboardingState.leftAfterSaveForRunTest = false;
                    showShipPromptCue();
                    return;
                }
                if (onboardingState.awaitingBeginnerPaste) {
                    prepareBeginnerPasteTarget({ sticky: true });
                    return;
                }
                if (!onboardingState.awaitingAiReturn || !onboardingState.leftAfterAi) return;
                showBeginnerReturnCue();
            });

            window.addEventListener('resize', () => {
                if (!onboardingState.beginnerMode) return;
                if (onboardingState.awaitingBeginnerPaste) {
                    positionBeginnerPastePopover();
                }
                if (beginnerInlineCueActive) {
                    positionBeginnerInlineCue();
                }
            });
            window.addEventListener('scroll', () => {
                if (!onboardingState.beginnerMode) return;
                if (onboardingState.awaitingBeginnerPaste) {
                    positionBeginnerPastePopover();
                }
                if (beginnerInlineCueActive) {
                    positionBeginnerInlineCue();
                }
            }, true);

            document.addEventListener('paste', (event) => {
                if (!onboardingState.beginnerMode) return;
                const editorHost = document.getElementById('editor');
                if (!editorHost || !editorHost.contains(event.target)) return;
                if (!(window.editor && typeof editor.getActiveUuid === 'function')) return;

                const activeUuid = editor.getActiveUuid();
                if (!activeUuid) return;
                const activeMeta = editor._meta ? editor._meta[activeUuid] : null;
                const activeName = String(activeMeta && activeMeta.name ? activeMeta.name : '').trim().toLowerCase();
                if (activeName !== 'index.html') return;

                const pastedText = (event.clipboardData || window.clipboardData)?.getData?.('text') || '';
                if (!String(pastedText).trim()) return;
                hideBeginnerInlineCue();
                if (typeof window.hideForgeBeginnerProviderTip === 'function') {
                    window.hideForgeBeginnerProviderTip();
                }
                if (onboardingState.awaitingBeginnerPaste) {
                    markBeginnerPasteCompleted();
                }
            }, true);

            const shipNowBtn = document.getElementById('ship-now-btn');
            if (shipNowBtn) {
                shipNowBtn.addEventListener('click', () => {
                    const shouldShowFeatureCue = !!(
                        onboardingState.beginnerMode &&
                        onboardingState.shipPromptShown &&
                        !onboardingState.quickPromptsHintShown
                    );
                    hideBeginnerInlineCue();
                    if (shouldShowFeatureCue) {
                        setTimeout(() => {
                            showQuickPromptsFeatureCue();
                        }, 900);
                    }
                });
            }

            document.addEventListener('click', (event) => {
                if (!onboardingState.beginnerMode || !beginnerInlineCueActive) return;
                const clickTarget = event.target;
                if (!(clickTarget instanceof Element)) return;
                if (
                    clickTarget.closest('#quick-prompts-btn') ||
                    clickTarget.closest('[data-quick-prompt-template]')
                ) {
                    hideBeginnerInlineCue();
                }
            });

            document.addEventListener('forge:index-html-saved', () => {
                if (!onboardingState.beginnerMode || !onboardingState.awaitingProductTourAfterSave) return;
                markBeginnerPasteCompleted();
                onboardingState.awaitingProductTourAfterSave = false;
                onboardingState.awaitingProductTourReturnAfterSave = true;
                onboardingState.leftAfterSaveForRunTest = false;
                showRunTestCue();
            });

            showPreloadActions();

            const originalGetFile = loadFolder.getFile.bind(loadFolder);
            loadFolder.getFile = async function () {
                const loaded = await originalGetFile();
                if (!loaded || !loadFolder.fileHandle) {
                    showPreloadActions();
                    return;
                }
                search.onProjectLoaded();

                // Enable tool links once project is loaded
                $('#compiler-tab, #decompiler-tab, #devconsole-tab, #test-recorder-tab, #advanced-debug-tab, #math-logic-tester-tab, #ai-helper-tab, #sast-tab, #security-reviewer-tab').removeClass('disabled');
                if (typeof securityReviewer !== 'undefined') securityReviewer.onProjectLoaded();
                if (typeof compiler !== 'undefined' && compiler && typeof compiler.setLivePreviewStatus === 'function') {
                    compiler.setLivePreviewStatus('Ready to build a runtime preview.', 'info');
                }

                // Switch "Open App Folder" to quick actions (AI Services + Ship)
                showLoadedActions();
                activateCodeEditorTab();
            };
        })();

        let forgeBeginnerProviderTipEl = null;
        let forgeBeginnerProviderLaunchTimer = null;
        let forgeBeginnerProviderCountdownTimer = null;
        let forgeBeginnerProviderSelectedUrl = '';
        let forgeBeginnerProviderSelectedLabel = '';
        let forgeBeginnerProviderTipLeftTab = false;

        const clearBeginnerProviderTipTimers = () => {
            if (forgeBeginnerProviderLaunchTimer) {
                clearTimeout(forgeBeginnerProviderLaunchTimer);
                forgeBeginnerProviderLaunchTimer = null;
            }
            if (forgeBeginnerProviderCountdownTimer) {
                clearInterval(forgeBeginnerProviderCountdownTimer);
                forgeBeginnerProviderCountdownTimer = null;
            }
        };

        const ensureBeginnerProviderTip = () => {
            if (forgeBeginnerProviderTipEl && document.body.contains(forgeBeginnerProviderTipEl)) {
                return forgeBeginnerProviderTipEl;
            }
            const el = document.createElement('div');
            el.id = 'forge-beginner-provider-tip';
            el.className = 'forge-beginner-provider-tip';
            el.innerHTML = `
                <div class="tip-head">
                    <p class="tip-title" id="forge-beginner-provider-tip-title"></p>
                    <button type="button" class="tip-close" id="forge-beginner-provider-tip-close" aria-label="Close AI Services setup tip" title="Close">x</button>
                </div>
                <div class="tip-body" id="forge-beginner-provider-tip-body"></div>
                <div class="tip-countdown" id="forge-beginner-provider-tip-countdown"></div>
                <div class="tip-actions">
                    <button type="button" class="btn btn-primary btn-sm tip-open-btn" id="forge-beginner-provider-tip-open-btn" disabled>AI Services</button>
                </div>
            `;
            document.body.appendChild(el);
            const closeBtn = el.querySelector('#forge-beginner-provider-tip-close');
            const openBtn = el.querySelector('#forge-beginner-provider-tip-open-btn');
            if (closeBtn) {
                closeBtn.addEventListener('click', () => {
                    hideBeginnerProviderTip();
                });
            }
            if (openBtn) {
                openBtn.addEventListener('click', () => {
                    if (!forgeBeginnerProviderSelectedUrl) return;
                    const openedWin = window.open(forgeBeginnerProviderSelectedUrl, '_blank', 'noopener,noreferrer');
                    const countdownEl = el.querySelector('#forge-beginner-provider-tip-countdown');
                    if (countdownEl) {
                        countdownEl.textContent = openedWin
                            ? 'Opened in a new tab. Keep this tip visible while you configure the model.'
                            : 'Could not open a new tab. Try again.';
                    }
                });
            }
            forgeBeginnerProviderTipEl = el;
            return el;
        };

        const hideBeginnerProviderTip = () => {
            clearBeginnerProviderTipTimers();
            forgeBeginnerProviderTipLeftTab = false;
            if (forgeBeginnerProviderTipEl) {
                forgeBeginnerProviderTipEl.classList.remove('is-visible');
            }
        };
        window.hideForgeBeginnerProviderTip = hideBeginnerProviderTip;

        const showBeginnerProviderTipAndLaunch = ({ providerKey, url }) => {
            const tipEl = ensureBeginnerProviderTip();
            if (!tipEl) return;
            const titleEl = tipEl.querySelector('#forge-beginner-provider-tip-title');
            const bodyEl = tipEl.querySelector('#forge-beginner-provider-tip-body');
            const countdownEl = tipEl.querySelector('#forge-beginner-provider-tip-countdown');
            const openBtn = tipEl.querySelector('#forge-beginner-provider-tip-open-btn');

            const configMap = {
                capra: {
                    title: 'Capra Setup',
                    body: 'In Capra, switch the model in the bottom right to <strong>GPT 5.x</strong> before you start.',
                    openLabel: 'Open Capra'
                },
                google: {
                    title: 'Google AI Studio Setup',
                    body: 'In Google AI Studio, pick <strong>Playground</strong> on the left (not Build), then select <strong>Gemini 3.1 Pro Preview</strong> on the top right.',
                    openLabel: 'Open Google AI Studio'
                }
            };
            const cfg = configMap[providerKey];
            if (!cfg) {
                if (url) window.open(url, '_blank', 'noopener,noreferrer');
                return;
            }

            clearBeginnerProviderTipTimers();
            forgeBeginnerProviderSelectedUrl = String(url || '');
            forgeBeginnerProviderSelectedLabel = String(cfg.openLabel || 'AI Services');

            if (titleEl) titleEl.textContent = cfg.title;
            if (bodyEl) bodyEl.innerHTML = cfg.body;
            tipEl.classList.add('is-visible');
            forgeBeginnerProviderTipLeftTab = false;

            let secondsLeft = 6;
            const setCountdownText = (text) => {
                if (countdownEl) countdownEl.textContent = text;
            };
            if (openBtn) {
                openBtn.disabled = true;
                openBtn.textContent = `${forgeBeginnerProviderSelectedLabel} (${secondsLeft}s)`;
            }
            setCountdownText('Review these settings first, then AI Services when the timer ends.');

            forgeBeginnerProviderCountdownTimer = setInterval(() => {
                secondsLeft = Math.max(0, secondsLeft - 1);
                if (openBtn && secondsLeft > 0) {
                    openBtn.textContent = `${forgeBeginnerProviderSelectedLabel} (${secondsLeft}s)`;
                }
            }, 1000);

            forgeBeginnerProviderLaunchTimer = setTimeout(() => {
                clearBeginnerProviderTipTimers();
                if (openBtn) {
                    openBtn.disabled = false;
                    openBtn.textContent = forgeBeginnerProviderSelectedLabel || 'AI Services';
                }
                setCountdownText('Ready. Click the button to AI Services.');
            }, 6000);
        };

        const getBeginnerAiProviderKey = ({ url = '', label = '' } = {}) => {
            const normalizedUrl = String(url || '').toLowerCase();
            const normalizedLabel = String(label || '').toLowerCase();
            if (normalizedUrl.includes('capra') || normalizedLabel.includes('capra')) {
                return 'capra';
            }
            if (normalizedUrl.includes('aistudio.google.com') || normalizedLabel.includes('google ai studio')) {
                return 'google';
            }
            return '';
        };

        document.addEventListener('visibilitychange', () => {
            if (!(forgeBeginnerProviderTipEl && forgeBeginnerProviderTipEl.classList.contains('is-visible'))) {
                return;
            }
            if (document.hidden) {
                forgeBeginnerProviderTipLeftTab = true;
                return;
            }
            if (forgeBeginnerProviderTipLeftTab) {
                hideBeginnerProviderTip();
            }
        });

        // Handle AI Services dropdown item clicks
        $(document).on('click', '#open-ai-dropdown .dropdown-item[data-ai-action="athena"]', function (e) {
            e.preventDefault();
            if (typeof window.toggleRightPanel === 'function') {
                window.toggleRightPanel(true);
            }
            if (typeof promptLab !== 'undefined' && promptLab.switchRpTab) {
                promptLab.switchRpTab('prometheus', { allowSetupModal: true });
            }
            const trigger = document.getElementById('open-ai-btn');
            if (trigger && window.bootstrap && bootstrap.Dropdown) {
                const dd = bootstrap.Dropdown.getOrCreateInstance(trigger);
                dd.hide();
            }
        });

        $(document).on('click', '#open-ai-dropdown .dropdown-item[data-ai-url]', function (e) {
            e.preventDefault();
            const url = $(this).data('ai-url');
            const label = $(this).text();
            const providerKey = getBeginnerAiProviderKey({ url, label });
            const onboardingState = window.forgeBeginnerOnboardingState;
            if (onboardingState && onboardingState.beginnerMode) {
                clearBeginnerProviderTipTimers();
                onboardingState.awaitingAiReturn = true;
                onboardingState.leftAfterAi = false;
                onboardingState.aiReturnShownForLaunch = false;
                onboardingState.awaitingBeginnerPaste = false;
                onboardingState.awaitingProductTourAfterSave = false;
                onboardingState.awaitingProductTourReturnAfterSave = false;
                onboardingState.leftAfterSaveForRunTest = false;
                document.querySelectorAll('.forge-beginner-paste-highlight').forEach(el => {
                    el.classList.remove('forge-beginner-paste-highlight');
                });
                const beginnerPastePopover = document.getElementById('forge-beginner-paste-popover');
                if (beginnerPastePopover) {
                    beginnerPastePopover.classList.remove('is-visible', 'is-below');
                }
                if (window.forgeBeginnerInlineHint && typeof window.forgeBeginnerInlineHint.hide === 'function') {
                    window.forgeBeginnerInlineHint.hide();
                }
                if (providerKey === 'capra' || providerKey === 'google') {
                    showBeginnerProviderTipAndLaunch({
                        providerKey,
                        url
                    });
                    return;
                }
                hideBeginnerProviderTip();
            }
            if (url) {
                window.open(url, '_blank', 'noopener,noreferrer');
            }
        });

        // =====================================================
        // AI HELP PROMPT COPY HANDLER
        // =====================================================
        $('#copy-ai-help-prompt').on('click', function () {
            const prompt = document.getElementById('ai-help-prompt').value;
            navigator.clipboard.writeText(prompt).then(() => {
                const btn = this;
                const original = btn.textContent;
                btn.textContent = 'Copied!';
                btn.classList.remove('btn-primary');
                btn.classList.add('btn-success');
                setTimeout(() => {
                    btn.textContent = original;
                    btn.classList.remove('btn-success');
                    btn.classList.add('btn-primary');
                }, 1500);
            }).catch(err => {
                alert('Failed to copy: ' + err);
            });
        });
    
