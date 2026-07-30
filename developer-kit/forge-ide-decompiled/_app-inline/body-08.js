        (function () {
            const ATHENA_OPEN_KEY = 'forge:athena-open';
            const ATHENA_WIDTH_KEY = 'forge:athena-width';
            const ATHENA_MIN_WIDTH = 320;
            const ATHENA_MAX_WIDTH = 680;
            const ATHENA_DEFAULT_WIDTH = 480;
            const EDITOR_SIDEBAR_WIDTH_KEY = 'wct:sidebar-width';
            const EDITOR_SIDEBAR_MIN_WIDTH = 180;
            const EDITOR_SIDEBAR_DEFAULT_WIDTH = 312;

            function clamp(v, min, max) {
                return Math.max(min, Math.min(max, v));
            }

            function isCompactAthenaLayout() {
                return window.matchMedia('(max-width: 991.98px)').matches;
            }

            function getAthenaPanel() {
                return document.getElementById('right-panel');
            }

            function setAthenaWidth(px) {
                const width = clamp(parseInt(px, 10) || ATHENA_DEFAULT_WIDTH, ATHENA_MIN_WIDTH, ATHENA_MAX_WIDTH);
                document.documentElement.style.setProperty('--right-panel-w', width + 'px');
                // Position the resize handle at the panel's left edge
                const handle = document.getElementById('agent-resize-handle');
                if (handle) handle.style.right = width + 'px';
            }

            function getStoredAthenaWidth() {
                const stored = parseInt(localStorage.getItem(ATHENA_WIDTH_KEY) || '', 10);
                if (!isNaN(stored)) return clamp(stored, ATHENA_MIN_WIDTH, ATHENA_MAX_WIDTH);
                return ATHENA_DEFAULT_WIDTH;
            }

            function syncAthenaButtons(isOpen) {
                const reopenBtn = document.getElementById('right-panel-reopen-btn');
                if (reopenBtn) {
                    const shouldShowReopen = !isOpen && isCodeEditorActive() && !isStartupGateVisible();
                    reopenBtn.hidden = !shouldShowReopen;
                    reopenBtn.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
                }
            }

            function isCodeEditorActive() {
                const activePane = document.querySelector('.tab-pane.active.show');
                return !!(activePane && activePane.id === 'editor-container');
            }

            function isStartupGateVisible() {
                const startupGate = document.getElementById('startup-gate');
                return !!(startupGate && startupGate.classList.contains('is-visible'));
            }

            function getEditorSidebar() {
                return document.getElementById('sidebar');
            }

            function getEditorSidebarMainPane() {
                const editorContainer = document.getElementById('editor-container');
                return editorContainer ? editorContainer.querySelector('.row > .col-9') : null;
            }

            function isCompactEditorSidebarLayout() {
                return window.matchMedia('(max-width: 820px)').matches;
            }

            function getStoredEditorSidebarWidth() {
                const stored = parseInt(localStorage.getItem(EDITOR_SIDEBAR_WIDTH_KEY) || '', 10);
                if (!isNaN(stored) && stored >= EDITOR_SIDEBAR_MIN_WIDTH) return stored;
                return EDITOR_SIDEBAR_DEFAULT_WIDTH;
            }

            function syncEditorSidebarButtons() {
                const sidebar = getEditorSidebar();
                const toggleBtn = document.getElementById('editor-sidebar-toggle-btn');
                const reopenBtn = document.getElementById('editor-sidebar-reopen-btn');
                if (!sidebar) return;

                const collapsed = sidebar.classList.contains('collapsed');
                if (toggleBtn) {
                    toggleBtn.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
                    toggleBtn.setAttribute('aria-label', collapsed ? 'Show file and search sidebar' : 'Hide file and search sidebar');
                    toggleBtn.setAttribute('title', (collapsed ? 'Show' : 'Hide') + ' sidebar (Ctrl+B)');
                }
                if (reopenBtn) {
                    reopenBtn.hidden = !collapsed;
                    reopenBtn.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
                }
            }

            function setEditorSidebarCollapsed(forceCollapsed) {
                const sidebar = getEditorSidebar();
                const mainPane = getEditorSidebarMainPane();
                if (!sidebar || !mainPane) return;

                const collapsed = typeof forceCollapsed === 'boolean'
                    ? forceCollapsed
                    : !sidebar.classList.contains('collapsed');
                const compact = isCompactEditorSidebarLayout();

                if (collapsed) {
                    if (!compact) {
                        const currentWidth = Math.round(sidebar.getBoundingClientRect().width);
                        if (currentWidth >= EDITOR_SIDEBAR_MIN_WIDTH) {
                            localStorage.setItem(EDITOR_SIDEBAR_WIDTH_KEY, String(currentWidth));
                        }
                        sidebar.style.width = '0px';
                    } else {
                        sidebar.style.width = '';
                    }
                    sidebar.classList.add('collapsed');
                } else {
                    sidebar.classList.remove('collapsed');
                    if (compact) {
                        sidebar.style.width = '';
                    } else {
                        sidebar.style.width = getStoredEditorSidebarWidth() + 'px';
                    }
                }

                mainPane.style.width = collapsed ? '100%' : '';
                mainPane.style.maxWidth = collapsed ? '100%' : '';
                syncEditorSidebarButtons();
            }

            function toggleEditorSidebar() {
                const sidebar = getEditorSidebar();
                if (!sidebar) return;
                setEditorSidebarCollapsed(!sidebar.classList.contains('collapsed'));
            }

            function ensureAthenaReady() {
                if (typeof window.aiAgent !== 'undefined' && window.aiAgent && typeof window.aiAgent.loadProfiles === 'function') {
                    window.aiAgent.loadProfiles({ allowSetupModal: false });
                }
            }

            window.restartAthenaConversation = function () {
                if (typeof window.aiAgent === 'undefined' || !window.aiAgent) return;
                const runFreshChat = function () {
                    if (typeof window.aiAgent.newChat === 'function') {
                        window.aiAgent.newChat();
                    } else if (typeof window.aiAgent.clearChat === 'function') {
                        window.aiAgent.clearChat();
                    }
                    const input = document.getElementById('ai-chat-input');
                    if (input) {
                        input.value = '';
                        input.focus();
                        input.placeholder = 'Ask Prometheus... (Enter to send, Shift+Enter for newline)';
                    }
                };

                const wasBusy = !!window.aiAgent._busy;
                if (wasBusy && typeof window.aiAgent.abort === 'function') {
                    try { window.aiAgent.abort(); } catch (_) { }
                }
                if (wasBusy) {
                    setTimeout(runFreshChat, 80);
                } else {
                    runFreshChat();
                }
            };

            window.toggleAthenaConfig = function (forceOpen) {
                if (!window.aiAgent) return;
                if (typeof forceOpen === 'boolean' && typeof window.aiAgent._setConfigOpen === 'function') {
                    window.aiAgent._setConfigOpen(forceOpen, { focus: forceOpen });
                    return;
                }
                if (typeof window.aiAgent._toggleConfig === 'function') {
                    window.aiAgent._toggleConfig();
                }
            };

            window.toggleRightPanel = function (forceOpen) {
                const panel = getAthenaPanel();
                if (!panel) return;
                const open = typeof forceOpen === 'boolean'
                    ? forceOpen
                    : panel.classList.contains('collapsed');

                const handle = document.getElementById('agent-resize-handle');

                if (open) {
                    panel.classList.remove('collapsed');
                    document.body.classList.add('athena-open');
                    ensureAthenaReady();
                    if (!isCompactAthenaLayout()) {
                        setAthenaWidth(getStoredAthenaWidth());
                        if (handle) handle.classList.add('visible');
                    }
                } else {
                    panel.classList.add('collapsed');
                    document.body.classList.remove('athena-open');
                    if (handle) handle.classList.remove('visible');
                }

                syncAthenaButtons(open);
                localStorage.setItem(ATHENA_OPEN_KEY, open ? '1' : '0');
            };

            function initAthenaResizer() {
                const handle = document.getElementById('agent-resize-handle');
                const panel = getAthenaPanel();
                if (!handle || !panel) return;

                let dragging = false;
                let startX = 0;
                let startWidth = 0;

                handle.addEventListener('mousedown', function (e) {
                    if (panel.classList.contains('collapsed') || isCompactAthenaLayout()) return;
                    e.preventDefault();
                    dragging = true;
                    startX = e.clientX;
                    startWidth = panel.getBoundingClientRect().width;
                    handle.classList.add('dragging');
                    document.body.style.userSelect = 'none';
                    document.body.style.cursor = 'col-resize';
                });

                document.addEventListener('mousemove', function (e) {
                    if (!dragging) return;
                    const nextWidth = clamp(startWidth - (e.clientX - startX), ATHENA_MIN_WIDTH, ATHENA_MAX_WIDTH);
                    setAthenaWidth(nextWidth);
                });

                document.addEventListener('mouseup', function () {
                    if (!dragging) return;
                    dragging = false;
                    handle.classList.remove('dragging');
                    document.body.style.userSelect = '';
                    document.body.style.cursor = '';
                    const width = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--right-panel-w'), 10);
                    if (!isNaN(width)) {
                        localStorage.setItem(ATHENA_WIDTH_KEY, String(width));
                    }
                });
            }

            document.addEventListener('DOMContentLoaded', function () {
                const reopenBtn = document.getElementById('right-panel-reopen-btn');
                if (reopenBtn) {
                    reopenBtn.addEventListener('click', function () {
                        window.toggleRightPanel(true);
                    });
                }

                const editorSidebarToggleBtn = document.getElementById('editor-sidebar-toggle-btn');
                const editorSidebarReopenBtn = document.getElementById('editor-sidebar-reopen-btn');
                if (editorSidebarToggleBtn) {
                    editorSidebarToggleBtn.addEventListener('click', function () {
                        toggleEditorSidebar();
                    });
                }
                if (editorSidebarReopenBtn) {
                    editorSidebarReopenBtn.addEventListener('click', function () {
                        setEditorSidebarCollapsed(false);
                    });
                }

                setAthenaWidth(getStoredAthenaWidth());
                initAthenaResizer();
                setEditorSidebarCollapsed(false);

                const isStartup = isStartupGateVisible();
                const shouldOpen = !isStartup && isCodeEditorActive();
                window.toggleRightPanel(shouldOpen);
                syncEditorSidebarButtons();

                if (window.jQuery) {
                    $(document).on('shown.bs.tab', 'a[data-bs-toggle="tab"]', function () {
                        const target = (this.getAttribute('href') || this.getAttribute('data-bs-target') || '').trim();
                        if (target === '#editor-container' && !document.body.classList.contains('athena-open')) {
                            window.toggleRightPanel(true);
                        } else if (target && target !== '#editor-container' && document.body.classList.contains('athena-open')) {
                            window.toggleRightPanel(false);
                        }
                        const open = document.body.classList.contains('athena-open');
                        syncAthenaButtons(open);
                        syncEditorSidebarButtons();
                    });
                }

                window.addEventListener('resize', function () {
                    if (isCompactAthenaLayout()) {
                        document.documentElement.style.setProperty('--right-panel-w', 'min(92vw, 420px)');
                    } else {
                        setAthenaWidth(getStoredAthenaWidth());
                    }
                    const sidebar = getEditorSidebar();
                    if (sidebar) {
                        if (sidebar.classList.contains('collapsed')) {
                            sidebar.style.width = isCompactEditorSidebarLayout() ? '' : '0px';
                        } else {
                            sidebar.style.width = isCompactEditorSidebarLayout() ? '' : (getStoredEditorSidebarWidth() + 'px');
                        }
                    }
                    syncEditorSidebarButtons();
                });

                document.addEventListener('keydown', function (e) {
                    const key = String(e.key || '').toLowerCase();
                    if ((e.ctrlKey || e.metaKey) && !e.shiftKey && key === 'b') {
                        if (!isCodeEditorActive()) return;
                        e.preventDefault();
                        toggleEditorSidebar();
                    }
                });
            });
        })();
    
