
var PREVIEW_LIMIT = 28; // Max characters shown in preview snippet

// Text separators used by ONLYOFFICE API
var PARA_SEP = '\r';
var LINE_SEP = '\n';
var TAB_SYMBOL = '\t';

// Options for GetSelectedText API method
var SELECTION_PROPS = {
    Numbering: false,
    Math: false,
    ParaSeparator: PARA_SEP,
    TableRowSeparator: PARA_SEP,
    TableCellSeparator: PARA_SEP,
    NewLineSeparator: LINE_SEP,
    TabSymbol: TAB_SYMBOL
};

// Actions for conversion
var ACTIONS = {
    cyrToLat: {
        label: 'Cyrillic → Latin',
        convert: function (text) {
            return cyrlatconverter.toLatin(text);
        }
    },
    latToCyr: {
        label: 'Latin → Cyrillic',
        convert: function (text, options) {
            return cyrlatconverter.toCyrillic(text, options);
        }
    },
    autoSwap: {
        label: 'Automatic swap',
        convert: function (text, options) {
            return cyrlatconverter.autoConvert(text, options);
        }
    }
};

// Context menu order
var MENU_ORDER = ['cyrToLat', 'latToCyr', 'autoSwap'];  
var lastSelectedText = '';
var dictionaryLoaded = false;

// Load Serbian dictionary (~264k words) for foreign word detection.
// Words not in dictionary are considered foreign and kept in Latin.
(function loadDict() {
    if (typeof cyrlatconverter !== 'undefined' && cyrlatconverter.loadDictionary) {
        cyrlatconverter.loadDictionary('resources/dict/sr-Latn.dic', function(err, success) {
            dictionaryLoaded = success;
            if (success) {
                refreshDetection();
            }
        });
    }
})();

// Called when plugin initializes or selection changes (initOnSelectionChanged)
window.Asc.plugin.init = function (text) {
    if (typeof text === 'string' && text.trim()) {
        updateDetection(text);
        return;
    }
    refreshDetection();
};

// Handle theme changes (light/dark mode)
window.Asc.plugin.onThemeChanged = function (theme) {
    if (window.Asc.plugin.onThemeChangedBase) {
        window.Asc.plugin.onThemeChangedBase(theme);
    }

    var isDark = !!(theme && (theme.type === 'dark' || theme.type === 'contrast-dark'));
    document.documentElement.classList.toggle('theme-dark', isDark);
    document.body.classList.toggle('theme-dark', isDark);
    document.body.style.backgroundColor = '';
    document.body.style.color = '';
};

// Build context menu with conversion options
Asc.plugin.attachEvent('onContextMenuShow', function () {
    var items = MENU_ORDER.map(function (id) {
        return { id: id, text: tr(ACTIONS[id].label) };
    });

    window.Asc.plugin.executeMethod('UpdateContextMenuItem', [{
        guid: window.Asc.plugin.guid,
        items: [{
            id: 'srConverter',
            text: tr('Serbian converter'),
            items: items
        }]
    }]);
});

// Handle context menu item click
window.Asc.plugin.event_onContextMenuClick = function (id) {
    action(id);
};

// Refresh detection when cursor moves
window.Asc.plugin.event_onTargetPositionChanged = function () {
    refreshDetection();
};

// Fetch current selection from document and update UI
function refreshDetection() {
    if (!window.Asc.plugin.executeMethod) {
        updateDetection(lastSelectedText);
        return;
    }

    window.Asc.plugin.executeMethod('GetSelectedText', [SELECTION_PROPS], function (selectedText) {
        updateDetection(selectedText || '');
    });
}

// Truncate text for preview display
function previewSnippet(text) {
    var compact = String(text || '').replace(/\s+/g, ' ').trim();
    if (compact.length <= PREVIEW_LIMIT) return compact;
    return compact.slice(0, PREVIEW_LIMIT - 1) + '…';
}

// Detect script type (Cyrillic/Latin) and update panel UI
function updateDetection(text) {
    lastSelectedText = text || '';

    var scriptEl = document.getElementById('detectScript');
    var preview = document.getElementById('detectPreview');
    var fromEl = document.getElementById('previewFrom');
    var toEl = document.getElementById('previewTo');
    var primary = document.getElementById('btnPrimary');
    if (!scriptEl || !primary) return;

    var trimmed = lastSelectedText.trim();

    // No selection - reset UI
    if (!trimmed) {
        scriptEl.textContent = tr('No text selected');
        if (preview) preview.hidden = true;
        if (fromEl) fromEl.textContent = '';
        if (toEl) toEl.textContent = '';
        primary.disabled = true;
        primary.textContent = tr('Convert');
        primary.setAttribute('data-action', 'autoSwap');
        return;
    }

    // Detect script and prepare conversion preview
    var script = cyrlatconverter.detectScript(trimmed);
    var options = currentOptions();
    var converted = trimmed;
    var actionId = 'autoSwap';
    var label = tr('Convert');

    if (script === 'cyrillic') {
        scriptEl.textContent = tr('Cyrillic');
        converted = cyrlatconverter.toLatin(trimmed);
        actionId = 'cyrToLat';
        label = tr('Convert to Latin');
    } else if (script === 'latin') {
        scriptEl.textContent = tr('Latin');
        converted = cyrlatconverter.toCyrillic(trimmed, options);
        actionId = 'latToCyr';
        label = tr('Convert to Cyrillic');
    } else {
        scriptEl.textContent = tr('Unknown');
    }

    // Update button and preview
    primary.disabled = false;
    primary.textContent = label;
    primary.setAttribute('data-action', actionId);

    if (preview && fromEl && toEl) {
        fromEl.textContent = previewSnippet(trimmed);
        toEl.textContent = previewSnippet(converted);
        preview.hidden = false;
    }
}

// Called when primary button is clicked
function runPrimary() {
    var primary = document.getElementById('btnPrimary');
    action((primary && primary.getAttribute('data-action')) || 'autoSwap');
}

// Execute conversion action by ID
function action(id) {
    var handler = ACTIONS[id];
    if (!handler) return;

    window.Asc.plugin.executeMethod('GetSelectedText', [SELECTION_PROPS], function (selectedText) {
        if (!selectedText || selectedText.trim() === '') {
            showToast(tr('Please select some text first.'), true, 3000);
            return;
        }

        var options = currentOptions();

        // Convert each paragraph separately to preserve structure
        var paragraphs = selectedText.split(PARA_SEP).map(function (paragraph) {
            return handler.convert(paragraph, options);
        });

        // Check if any conversion happened
        if (paragraphs.join(PARA_SEP) === selectedText) {
            showToast(tr('Nothing to convert here.'), true, 3000);
            return;
        }

        replaceSelection(paragraphs);
    });
}

// Replace selected text in document with converted text
function replaceSelection(paragraphs) {
    Asc.scope.paragraphs = paragraphs;

    // Try ReplaceTextSmart first (preserves formatting)
    window.Asc.plugin.executeMethod('ReplaceTextSmart', [paragraphs, TAB_SYMBOL, LINE_SEP], function (isDone) {
        if (isDone) {
            showToast(tr('Conversion finished.'));
            refreshDetection();
            return;
        }

        // Fallback to callCommand for older API versions
        window.Asc.plugin.callCommand(function () {
            Api.ReplaceTextSmart(Asc.scope.paragraphs);
        }, false, true, function () {
            showToast(tr('Conversion finished.'));
            refreshDetection();
        });
    });
}

// Get current conversion options from UI
function currentOptions() {
    return {
        keepForeign: isChecked('optForeign')
    };
}

// Check if checkbox is checked (defaults to true if not found)
function isChecked(id) {
    var checkbox = document.getElementById(id);
    if (!checkbox) return true;
    return !!checkbox.checked;
}

// Translate text using ONLYOFFICE i18n system
function tr(text) {
    return window.Asc.plugin.tr ? window.Asc.plugin.tr(text) : text;
}

// Show toast notification
function showToast(message, isWarning, duration) {
    var toast = document.getElementById('toast');
    toast.classList.toggle('toast-warn', !!isWarning);
    toast.textContent = message;
    toast.classList.add('show');

    setTimeout(function () {
        toast.classList.remove('show');
    }, duration || 1500);
}

// Called when translations are loaded - update all UI text
window.Asc.plugin.onTranslate = function () {
    setText('hintTop', 'Select text in the document to convert it.');
    setText('detectLabel', 'Selected text is recognized as');
    setText('optForeignLabel', 'Keep foreign words and brands unchanged (Windows, Google)');
    setText('hintBottom', 'Right-click selected text for a quick conversion');
    updateDetection(lastSelectedText);
};

// Set translated text for element by ID
function setText(id, text) {
    var element = document.getElementById(id);
    if (element) element.textContent = tr(text);
}
