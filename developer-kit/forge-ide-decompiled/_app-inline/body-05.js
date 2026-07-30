        const cm6Globals = window.cm6;
        if (!cm6Globals) {
            throw new Error('CodeMirror bundle failed to load.');
        }

        const {
            EditorState,
            EditorView,
            keymap,
            basicSetup,
            searchKeymap,
            foldCode,
            foldKeymap,
            syntaxTree,
            lintGutter,
            lintKeymap,
            linter,
            javascript,
            javascriptLanguage,
            html,
            css,
            python,
            oneDark
        } = cm6Globals;

        const modules = {
            EditorState,
            EditorView,
            keymap,
            basicSetup,
            searchKeymap: Array.isArray(searchKeymap) ? searchKeymap : [],
            foldCode: typeof foldCode === 'function' ? foldCode : null,
            foldKeymap: Array.isArray(foldKeymap) ? foldKeymap : [],
            syntaxTree: typeof syntaxTree === 'function' ? syntaxTree : null,
            lintGutter: typeof lintGutter === 'function' ? lintGutter : null,
            lintKeymap: Array.isArray(lintKeymap) ? lintKeymap : [],
            linter: typeof linter === 'function' ? linter : null,
            javascript,
            javascriptLanguage,
            html,
            css,
            python,
            oneDark
        };

        modules.materialLikeTheme = modules.EditorView.theme({
            '&': { backgroundColor: '#111820', color: '#dce3eb' },
            '.cm-content': {
                fontFamily: "'JetBrains Mono', 'Fira Code', 'Consolas', monospace",
                fontSize: '14px'
            },
            '.cm-gutters': { backgroundColor: '#0e1419', color: '#5a6b7e', border: 'none' },
            '&.cm-focused .cm-selectionBackground, .cm-selectionBackground, & ::selection': {
                backgroundColor: 'rgba(59, 130, 246, 0.18)'
            }
        }, { dark: true });

        modules.searchKeymapExtensions = modules.searchKeymap.length
            ? [modules.keymap.of(modules.searchKeymap)]
            : [];

        modules.foldKeymapExtensions = modules.foldKeymap.length
            ? [modules.keymap.of(modules.foldKeymap)]
            : [];

        modules.ctrlQFoldKeymap = modules.foldCode
            ? [modules.keymap.of([{ key: 'Ctrl-q', run: modules.foldCode, preventDefault: true }])]
            : [];

        modules._jsLintExtensions = null;
        modules.getJsLinterExtensions = () => {
            if (!modules.linter || !modules.lintGutter) {
                return [];
            }
            if (!modules._jsLintExtensions) {
                const lintExtension = modules.linter(view => {
                    const doc = view.state.doc;
                    const text = doc.toString();
                    const diagnostics = [];

                    let tree = null;
                    try {
                        tree = modules.syntaxTree ? modules.syntaxTree(view.state) : null;
                    } catch (_) {
                        tree = null;
                    }

                    if (!tree && modules.javascriptLanguage) {
                        try {
                            tree = modules.javascriptLanguage.parser.parse(text);
                        } catch (_) {
                            tree = null;
                        }
                    }

                    if (tree) {
                        tree.iterate({
                            enter(type, from, to) {
                                if (type.isError) {
                                    const end = to > from ? to : Math.min(doc.length, from + 1);
                                    diagnostics.push({
                                        from,
                                        to: end,
                                        severity: 'error',
                                        message: 'Syntax error'
                                    });
                                }
                            }
                        });
                    }

                    return diagnostics;
                });

                const keymapExt = modules.lintKeymap.length
                    ? modules.keymap.of(modules.lintKeymap)
                    : null;

                modules._jsLintExtensions = [modules.lintGutter(), lintExtension];
                if (keymapExt) modules._jsLintExtensions.push(keymapExt);
            }
            return modules._jsLintExtensions;
        };

        window.cmModules = modules;
        window.cmModulesReady = Promise.resolve(modules);
    
