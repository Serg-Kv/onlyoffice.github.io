((window) => {
    
    let originalState = null, hasCleanedDoc = false, undoCount = 0;

    const CONFIG = {
        removeBold: { method: 'SetBold', value: false },
        removeItalic: { method: 'SetItalic', value: false },
        removeUnderline: { method: 'SetUnderline', value: false },
        removeStrikeout: { method: 'SetStrikeout', value: false },
        clearTextColor: { method: 'SetColor', value: [0, 0, 0, true] },
        removeHighlight: { method: 'SetHighlight', value: 'none' },
        resetLetterSpacing: { method: 'SetSpacing', value: 0 },
        resetVertOffset: { method: 'SetPosition', value: 0 },
        disableAllCaps: { method: 'SetCaps', value: false },
        disableSmallCaps: { method: 'SetSmallCaps', value: false },
        resetBaseline: { method: 'SetVertAlign', value: 'baseline' }
    };

    const CASE_OPTIONS = {
        upper: t => t.toUpperCase(),
        lower: t => t.toLowerCase(),
        sentence: t => t.charAt(0).toUpperCase() + t.slice(1).toLowerCase(),
        capitalize: t => t.replace(/\b\w/g, l => l.toUpperCase()),
        toggle: t => t.split('').map(c => c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase()).join('')
    };

   
    const getContextItems = () => [
        { 
            id: 'textCleaner', 
            text: 'TextCleanerMenuTitle', 
            icons: getThemeIcon(), 
            items: [
                { id: 'clearFormattingCtx', text: 'ClearFormatting', items: [
                    { id: 'removeBoldCtx', text: 'RemoveBold' },
                    { id: 'removeItalicCtx', text: 'RemoveItalic' },
                    { id: 'removeUnderlineCtx', text: 'RemoveUnderline' },
                    { id: 'removeStrikeoutCtx', text: 'RemoveStrikeout' },
                    { id: 'clearTextColorCtx', text: 'ClearTextColor' },
                    { id: 'removeHighlightCtx', text: 'RemoveHighlight' },
                    { id: 'removeBgOutlineCtx', text: 'RemoveBgOutline' }
                ]},
                { id: 'fontStandardizationCtx', text: 'FontStandardization', items: [
                    { id: 'resetLetterSpacingCtx', text: 'ResetLetterSpacing' },
                    { id: 'resetVertOffsetCtx', text: 'ResetVertOffset' }
                ]},
                { id: 'textCaseConversionCtx', text: 'TextCaseConversion', items: [
                    { id: 'doNotChangeCaseCtx', text: 'CaseNone' },
                    { id: 'sentenceCaseCtx', text: 'SentenceCase' },
                    { id: 'lowerCaseCtx', text: 'LowerCase' },
                    { id: 'upperCaseCtx', text: 'UpperCase' },
                    { id: 'capitalizeEachWordCtx', text: 'CapitalizeEach' },
                    { id: 'toggleCaseCtx', text: 'ToggleCase' }
                ]},
                { id: 'specialFormattingCtx', text: 'SpecialFormatting', items: [
                    { id: 'disableAllCapsCtx', text: 'DisableAllCaps' },
                    { id: 'disableSmallCapsCtx', text: 'DisableSmallCaps' },
                    { id: 'resetBaselineCtx', text: 'ResetBaseline' }
                ]}
            ]
        }
    ];

    function getThemeIcon() {
        try {
            if (window.Asc && window.Asc.plugin && window.Asc.plugin.info && window.Asc.plugin.info.theme) {
                const theme = window.Asc.plugin.info.theme;
                return theme.type === 'dark' ? 'resources/dark/icon.svg' : 'resources/light/icon.svg';
            }
            
            if (document && document.body) {
                const isDark = document.body.classList.contains('theme-dark') || 
                              getComputedStyle(document.documentElement).getPropertyValue('--theme-type') === 'dark';
                
                return isDark ? 'resources/dark/icon.svg' : 'resources/light/icon.svg';
            }
        } catch (error) {
            console.log('Theme detection failed, using light theme icon');
        }
        
        
        return 'resources/light/icon.svg';
    }

    // Utility functions
    const $ = id => document.getElementById(id);
    
    // Improved translation function that falls back to English for unsupported languages
    const tr = key => {
        try {
            if (window.Asc && window.Asc.plugin && typeof window.Asc.plugin.tr === 'function') {
                const translation = window.Asc.plugin.tr(key);
                
                
                if (translation && translation !== key) {
                    return translation;
                }
                
                
                return getEnglishFallback(key);
            }
        } catch (e) {
            console.log('Translation function failed, using fallback:', e);
        }
        
        return getEnglishFallback(key);
    };
    
    // English fallback translations for unsupported languages
    const getEnglishFallback = key => {
        const englishTranslations = {
            "TextCleaner": "Text Cleaner",
            "AllParameters": "All parameters",
            "ClearFormatting": "Clear Formatting",
            "clean-button": "Clean",
            "PluginInstructions": "Select a section of text to clear its formatting, or click <b>Clean</b> to clear formatting in the entire document.",
            "RemoveBold": "Remove bold",
            "RemoveItalic": "Remove italic",
            "RemoveUnderline": "Remove underline",
            "RemoveStrikeout": "Remove strikethrough",
            "ClearTextColor": "Clear text color",
            "RemoveHighlight": "Remove highlight",
            "RemoveBgOutline": "Remove background & outline",
            "FontStandardization": "Font Standardization",
            "ApplyFontStandardization": "Apply font standardization",
            "ResetLetterSpacing": "Reset letter spacing",
            "ResetVertOffset": "Reset vertical offset",
            "TextCaseConversion": "Text Case Conversion",
            "CaseNone": "Do Not Change",
            "SentenceCase": "Sentence case.",
            "LowerCase": "lowercase",
            "UpperCase": "UPPERCASE",
            "CapitalizeEach": "Capitalize Each Word",
            "ToggleCase": "tOGGLE cASE",
            "SpecialFormatting": "Special Formatting",
            "DisableAllCaps": "Disable ALL CAPS",
            "DisableSmallCaps": "Disable Small Caps",
            "ResetBaseline": "Reset to baseline",
            "TextCleanerMenuTitle": "Text Cleaner",
            "CleaningCompleted": "Text cleaning completed successfully!",
            "OperationsApplied": "operations applied",
            "RevertToOriginal": "Revert to Original",
            "NewClean": "New Clean",
            "DoNotClosePanel": "Please do not close the plugin panel.",
            "Loading": "Loading..."
        };
        
        return englishTranslations[key] || key;
    };
    
    const callCommand = (func, callback) => window.Asc.plugin.callCommand(func, false, true, callback);

    // Generic text property applier
    const applyTextProp = (method, value) => {
        // Store in global scope for callCommand access
        Asc.scope.currentMethod = method;
        Asc.scope.currentValue = value;
        
        callCommand(() => {
            const doc = Api.GetDocument();
            const range = doc.GetRangeBySelect();
            const textPr = Api.CreateTextPr();
            
            if (Array.isArray(Asc.scope.currentValue)) {
                textPr[Asc.scope.currentMethod](...Asc.scope.currentValue);
            } else {
                textPr[Asc.scope.currentMethod](Asc.scope.currentValue);
            }

            if (range && range.GetText && range.GetText() !== "") {
                range.SetTextPr(textPr);
            } else {
                const paragraphs = doc.GetAllParagraphs();
                for (let i = 0; i < paragraphs.length; i++) {
                    paragraphs[i].SetTextPr(textPr);
                }
            }
        });
        undoCount++;
    };

    // Special handlers
    const specialHandlers = {
        removeBgOutline: () => {
            callCommand(() => {
                const doc = Api.GetDocument();
                const range = doc.GetRangeBySelect();
                const noStroke = Api.CreateStroke(0, Api.CreateSolidFill(Api.CreateRGBColor(0, 0, 0)));
                
                const processItems = items => {
                    for (let i = 0; i < items.length; i++) {
                        const item = items[i];
                        item.SetShd("clear", 255, 255, 255);
                        const textPr = Api.CreateTextPr();
                        textPr.SetOutLine(noStroke);
                        item.SetTextPr(textPr);
                        
                        const paraPr = item.GetParaPr && item.GetParaPr();
                        if (paraPr) {
                            paraPr.SetLeftBorder("none", 0, 0, 0, 0, 0);
                            paraPr.SetRightBorder("none", 0, 0, 0, 0, 0);
                            paraPr.SetTopBorder("none", 0, 0, 0, 0, 0);
                            paraPr.SetBottomBorder("none", 0, 0, 0, 0, 0);
                            if (paraPr.SetBetweenBorder) paraPr.SetBetweenBorder("none", 0, 0, 0, 0, 0);
                        }
                    }
                };

                if (range && range.GetText && range.GetText() !== "") {
                    processItems([range]);
                } else {
                    processItems(doc.GetAllParagraphs());
                }
            });
            undoCount++;
        },

        applyFontStandardization: settings => {
            if (!settings.applyFontStandardization || (!settings.targetFontFamily && !settings.targetFontSize)) return;
            
            Asc.scope.targetFontFamily = settings.targetFontFamily;
            Asc.scope.targetFontSize = settings.targetFontSize;
            
            callCommand(() => {
                const doc = Api.GetDocument();
                const range = doc.GetRangeBySelect();
                const textPr = Api.CreateTextPr();
                
                if (Asc.scope.targetFontFamily) textPr.SetFontFamily(Asc.scope.targetFontFamily);
                if (Asc.scope.targetFontSize) textPr.SetFontSize(Asc.scope.targetFontSize * 2);

                if (range && range.GetText && range.GetText() !== "") {
                    range.SetTextPr(textPr);
                } else {
                    const paragraphs = doc.GetAllParagraphs();
                    for (let i = 0; i < paragraphs.length; i++) {
                        paragraphs[i].SetTextPr(textPr);
                    }
                }
            });
            undoCount++;
        },

        textCaseConversion: caseOption => {
            if (caseOption === "none") return;
        
            Asc.scope.textCaseOption = caseOption;
        
            callCommand(() => {
                const doc = Api.GetDocument();
                const range = doc.GetRangeBySelect();
                
                let convertCase;
                switch (Asc.scope.textCaseOption) {
                    case "upper":
                        convertCase = t => t.toUpperCase();
                        break;
                    case "lower":
                        convertCase = t => t.toLowerCase();
                        break;
                    case "sentence":
                        convertCase = t => t.charAt(0).toUpperCase() + t.slice(1).toLowerCase();
                        break;
                    case "capitalize":
                        convertCase = t => t.replace(/\b\w/g, l => l.toUpperCase());
                        break;
                    case "toggle":
                        convertCase = t => t.split('').map(c => c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase()).join('');
                        break;
                    default:
                        convertCase = t => t;
                }
        
                const processParagraphs = paragraphs => {
                    for (let i = 0; i < paragraphs.length; i++) {
                        const para = paragraphs[i];
                        
                        if (!para.GetElementsCount) continue;
        
                        const elementsCount = para.GetElementsCount();
                        let fullText = "";
                        let runs = [];
        
                        for (let j = 0; j < elementsCount; j++) {
                            const elem = para.GetElement(j);
                            if (elem.GetText) {
                                const text = elem.GetText();
                                if (text) {
                                    fullText += text;
                                    runs.push({ element: elem, text: text, length: text.length });
                                }
                            }
                        }
                        
                        if (fullText.trim() === "") continue;
        
                        const newFullText = convertCase(fullText);
        
                        if (newFullText !== fullText) {
                            para.RemoveAllElements();
                            let currentPos = 0;
                            for(let k = 0; k < runs.length; k++) {
                                const run = runs[k];
                                const newRunText = newFullText.substring(currentPos, currentPos + run.length);
                                const newRun = Api.CreateRun();
                                
                                const oldPr = run.element.GetTextPr();
                                newRun.SetTextPr(oldPr);
                                newRun.AddText(newRunText);
                                
                                para.AddElement(newRun);
                                currentPos += run.length;
                            }
                        }
                    }
                };
        
                if (range && range.GetText && range.GetText().trim() !== "") {
                    processParagraphs(range.GetAllParagraphs());
                } else {
                    processParagraphs(doc.GetAllParagraphs());
                }
            });
            undoCount++;
        }
    };

    // Main functions
    const getSettings = preset => preset || {
        removeBold: $("remove-bold")?.checked || false,
        removeItalic: $("remove-italic")?.checked || false,
        removeUnderline: $("remove-underline")?.checked || false,
        removeStrikeout: $("remove-strikeout")?.checked || false,
        clearTextColor: $("clear-text-color")?.checked || false,
        removeHighlight: $("remove-highlight")?.checked || false,
        removeBgOutline: $("remove-bg-outline")?.checked || false,
        resetLetterSpacing: $("reset-letter-spacing")?.checked || false,
        resetVertOffset: $("reset-vert-offset")?.checked || false,
        applyFontStandardization: $("apply-font-standardization")?.checked || false,
        targetFontFamily: $("font-family-select")?.value || "",
        targetFontSize: parseInt($("font-size-select")?.value || "0"),
        textCaseOption: document.querySelector('input[name="text-case-option"]:checked')?.value || "none",
        disableAllCaps: $("disable-all-caps")?.checked || false,
        disableSmallCaps: $("disable-small-caps")?.checked || false,
        resetBaseline: $("reset-baseline")?.checked || false
    };

    const runCleanCommand = preset => {
        const settings = getSettings(preset);
        Asc.scope.settings = settings;

        if (!originalState) {
            originalState = "saved";
        }
        undoCount = 0;

        // Auto-adjust case conversion for caps options
        if ((settings.disableAllCaps || settings.disableSmallCaps) && settings.textCaseOption === "none") {
            settings.textCaseOption = "lower";
        }

        // Apply standard text properties
        Object.entries(CONFIG).forEach(([key, config]) => {
            if (settings[key]) applyTextProp(config.method, config.value);
        });

        // Apply special handlers
        if (settings.removeBgOutline) specialHandlers.removeBgOutline();
        specialHandlers.applyFontStandardization(settings);
        specialHandlers.textCaseConversion(settings.textCaseOption);
        
        console.log("All text cleaning operations completed");
    };

    const showLoadingOverlay = () => {
        const loading = $('loading-view');
        const main = document.querySelector('.main-container');
        if (loading && main) {
            main.style.display = 'none';
            loading.style.display = 'block';
            setTimeout(() => {
                loading.style.display = 'none';
                showActionButtons();
            }, 1000);
        }
    };

    const showActionButtons = () => {
        const actionView = $('action-buttons-view');
        const undoInfo = $('undo-steps-info');
        if (actionView) {
            actionView.style.display = 'block';
            hasCleanedDoc = true;
            if (undoInfo) undoInfo.textContent = `${undoCount} ${tr("OperationsApplied")}`;
            setupActionHandlers();
        }
    };

    const setupActionHandlers = () => {
        const revert = $('revert-button');
        const newClean = $('new-clean-button');
        
        if (revert) revert.onclick = revertToOriginal;
        if (newClean) newClean.onclick = resetToMainView;
    };

    const revertToOriginal = () => {
        if (!originalState || undoCount === 0) return;
        
        const performUndo = stepsRemaining => {
            if (stepsRemaining <= 0) {
                originalState = null;
                hasCleanedDoc = false;
                undoCount = 0;
                resetToMainView();
                return;
            }
            window.Asc.plugin.executeMethod("Undo", null, () => {
                setTimeout(() => performUndo(stepsRemaining - 1), 100);
            });
        };
        performUndo(undoCount);
    };

    const resetToMainView = () => {
        const actionView = $('action-buttons-view');
        const main = document.querySelector('.main-container');
        if (actionView) actionView.style.display = 'none';
        if (main) main.style.display = 'flex';
    };

    const refreshButtonState = async () => {
        const hasText = await new Promise(resolve => 
            callCommand(() => {
                const doc = Api.GetDocument();
                return doc.GetText({ Numbering: false }).trim().length > 0;
            }, resolve)
        );
        const btn = $('clean-button');
        if (btn) btn.disabled = !hasText;
    };

    const onDomReady = () => {
        const selectAll = $('select-all-options');
        const checkboxes = document.querySelectorAll('input[type="checkbox"]:not(#select-all-options):not(#apply-font-standardization)');

        checkboxes.forEach(cb => cb.checked = true);
        const fontStd = $('apply-font-standardization');
        if (fontStd) fontStd.checked = false;
        if (selectAll) {
            selectAll.checked = true;
            selectAll.addEventListener('change', function() {
                checkboxes.forEach(cb => cb.checked = this.checked);
            });
        }

        // Clean button
        const cleanBtn = $('clean-button');
        if (cleanBtn) {
            cleanBtn.addEventListener('click', () => {
                showLoadingOverlay();
                if (window.Asc?.plugin?.button) window.Asc.plugin.button(0);
            });
        }

        // Accordion toggle
        document.querySelectorAll('.acc-head').forEach(btn => {
            btn.addEventListener('click', () => {
                const target = document.querySelector(btn.dataset.target);
                if (!target) return;
                const isOpen = target.style.display !== 'none';
                target.style.display = isOpen ? 'none' : 'block';
                const chevron = btn.querySelector('.chevron');
                if (chevron) chevron.style.transform = `rotate(${isOpen ? '0' : '180'}deg)`;
            });
            
            // Set initial chevron state
            const target = document.querySelector(btn.dataset.target);
            const chevron = btn.querySelector('.chevron');
            if (target && chevron) {
                chevron.style.transform = target.style.display !== 'none' ? 'rotate(0deg)' : 'rotate(180deg)';
            }
        });
    };

    // Context menu functionality
    const contextMenuActions = {
        removeBoldCtx: () => runCleanCommand({ removeBold: true }),
        removeItalicCtx: () => runCleanCommand({ removeItalic: true }),
        removeUnderlineCtx: () => runCleanCommand({ removeUnderline: true }),
        removeStrikeoutCtx: () => runCleanCommand({ removeStrikeout: true }),
        clearTextColorCtx: () => runCleanCommand({ clearTextColor: true }),
        removeHighlightCtx: () => runCleanCommand({ removeHighlight: true }),
        removeBgOutlineCtx: () => runCleanCommand({ removeBgOutline: true }),
        resetLetterSpacingCtx: () => runCleanCommand({ resetLetterSpacing: true }),
        resetVertOffsetCtx: () => runCleanCommand({ resetVertOffset: true }),
        doNotChangeCaseCtx: () => runCleanCommand({ textCaseOption: "none" }),
        sentenceCaseCtx: () => runCleanCommand({ textCaseOption: "sentence" }),
        lowerCaseCtx: () => runCleanCommand({ textCaseOption: "lower" }),
        upperCaseCtx: () => runCleanCommand({ textCaseOption: "upper" }),
        capitalizeEachWordCtx: () => runCleanCommand({ textCaseOption: "capitalize" }),
        toggleCaseCtx: () => runCleanCommand({ textCaseOption: "toggle" }),
        disableAllCapsCtx: () => runCleanCommand({ disableAllCaps: true, textCaseOption: "lower" }),
        disableSmallCapsCtx: () => runCleanCommand({ disableSmallCaps: true, textCaseOption: "lower" }),
        resetBaselineCtx: () => runCleanCommand({ resetBaseline: true })
    };

    window.Asc.plugin.init = function() {
        console.log("TextCleaner plugin initialized");
        setTimeout(() => {
            refreshButtonState();
            setInterval(refreshButtonState, 1500);
        }, 100);
    };

    window.Asc.plugin.button = id => {
        if (id === 0) runCleanCommand();
        else window.Asc.plugin.executeCommand("close", "");
    };

    window.Asc.plugin.onTranslate = () => {
        if (!$("PluginInstructions")) return;
        
        const setTr = idKey => {
            const el = $(idKey);
            if (el) el.innerHTML = tr(idKey);
        };

        const addChevron = id => {
            const head = $(id);
            if (!head) return;
            head.innerHTML = tr(id);
            const img = document.createElement('img');
            img.src = 'resources/light/chevron-down.svg';
            img.className = 'chevron';
            img.style.cssText = 'width:6px; float:right; transition:transform 0.2s';
            head.appendChild(img);
        };

        ['PluginInstructions', 'AllParameters', 'RemoveBold', 'RemoveItalic', 'RemoveUnderline', 
         'RemoveStrikeout', 'ClearTextColor', 'RemoveHighlight', 'RemoveBgOutline', 'ApplyFontStandardization',
         'NormalizeSpaces', 'NormalizeNumbers', 'ResetLetterSpacing', 'ResetVertOffset', 'FixCasing',
         'DisableAllCaps', 'DisableSmallCaps', 'ResetBaseline', 'clean-button', 'CaseNone', 'SentenceCase',
         'LowerCase', 'UpperCase', 'CapitalizeEach', 'ToggleCase', 'CleaningCompleted', 'OperationsApplied',
         'RevertToOriginal', 'NewClean', 'DoNotClosePanel', 'Loading'].forEach(setTr);

        ['ClearFormatting', 'FontStandardization', 'TextCaseConversion', 'SpecialFormatting'].forEach(addChevron);
    };

    window.Asc.plugin.event_onContextMenuShow = options => {
        if (!options) return;
        
        const contextItems = getContextItems().map(item => ({
            ...item,
            text: tr(item.text),
            icons: getThemeIcon(), 
            items: item.items ? translateContextItems(item.items) : undefined
        }));
        
        window.Asc.plugin.executeMethod("AddContextMenuItem", [{
            guid: window.Asc.plugin.guid,
            items: contextItems
        }]);
    };

    const translateContextItems = items => items.map(item => ({
        ...item,
        text: tr(item.text),
        items: item.items ? translateContextItems(item.items) : undefined
    }));

    window.Asc.plugin.event_onContextMenuClick = id => {
        const itemId = id.split("_oo_sep_")[0];
        if (contextMenuActions[itemId]) contextMenuActions[itemId]();
    };

    window.Asc.plugin.event_onDocumentContentReady = refreshButtonState;
    window.Asc.plugin.event_onTargetChanged = refreshButtonState;

    document.addEventListener('DOMContentLoaded', onDomReady);

})(window);