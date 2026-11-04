const tr = (s) =>
    (window.Asc && window.Asc.plugin && typeof window.Asc.plugin.tr === "function")
        ? window.Asc.plugin.tr(s) : s;

// Initialize the plugin
window.Asc.plugin.init = function() {};

window.Asc.plugin.button = function() {
    this.executeCommand("close", "");
};

// Context menu display event
Asc.plugin.attachEvent("onContextMenuShow", (options) => {
    const items = {
        guid: window.Asc.plugin.guid,
        items: [{
            id: "onClickItem1",
            text: { zh: "繁简转换器", en: "convert" },
            items: [{
                    id: "convertChineseToFan",
                    text: { zh: "简体字转繁体字", en: "Convert Select to Traditional" },
                },
                {
                    id: "convertFanToChinese",
                    text: { zh: "繁体字转简体字", en: "Convert Select to Simplified" },
                },
                {
                    id: "addPinyin",
                    text: { zh: "添加拼音（括号）", en: "Add pinyin (inline)" },
                },
                {
                    id: "addPinyinDefault",
                    text: { zh: "添加拼音（默认）", en: "Add pinyin (default)" },
                },
                {
                    id: "removePinyin",
                    text: { zh: "移除拼音", en: "Remove pinyin" },
                },
            ],
        }]
    };
    window.Asc.plugin.executeMethod("UpdateContextMenuItem", [items]);
});

function handleTextOperation(id) {
    switch (window.Asc.plugin.info.editorType) {
        case 'word': {
            getSelectedText(id);
            break;
        }
        case 'cell': {
            window.Asc.plugin.callCommand(function() {
                var oWorksheet = Api.GetActiveSheet();
                var oRange = oWorksheet.GetSelection();
                var selectedText = oRange.GetValue();
                return selectedText;
            }, false, false, (selectedText) => {
                if (typeof selectedText !== 'string') {
                    core(id, selectedText);
                }else getSelectedText(id, selectedText);
            });
            break;
        }
    }
}

function getSelectedText(id, preselectedText = null) {
    Asc.scope.preselectedText = preselectedText;
    const props = {
        Numbering: false,
        Math: false,
        TableCellSeparator: "\n",
        TableRowSeparator: "\n",
        ParaSeparator: "\n",
        TabSymbol: "\t",
        NewLineSeparator: "\r",
    };
    window.Asc.plugin.executeMethod("GetSelectedText", [props], function (data) {
        data = data || Asc.scope.preselectedText;
        if (data && data.trim() !== '') {
            core(id, data);
        } else {
            showToast(tr("Please select some text first"), "#FFAA00", 3000);
        }
    });
}

window.Asc.plugin.event_onContextMenuClick = (id) => {
    handleTextOperation(id);
};

function selectedPinyinMode() {
    const pinyinDefault = document.getElementById("pinyinDefault");
    return pinyinDefault && pinyinDefault.checked ? "default" : "inline";
}

function action(id) {
    if (id === "addPinyin" && selectedPinyinMode() === "default") {
        id = "addPinyinDefault";
    }
    handleTextOperation(id);
}

function core(id, selectedText) {
    const fileType = window.Asc.plugin.info.editorType;
    if (id == "convertChineseToFan") {
        replaceTranslatedText(selectedText, fileType, 'cn', 'tw');
        return;
    } else if (id == "convertFanToChinese") {
        replaceTranslatedText(selectedText, fileType, 'tw', 'cn');
        return;
    } else if (id == "addPinyin" || id == "addPinyinDefault") {
        const mode = id == "addPinyinDefault" ? "default" : "inline";
        // Word text is replaced with Api.ReplaceTextSmart so runs, links and
        // drawings stay in place. Paragraph insert/delete is still done directly:
        // ReplaceTextSmart can add paragraphs but cannot remove one, and the
        // Default layout needs both.
        if (fileType === 'word' && typeof selectedText === 'string') {
            applyWordPinyinOperation(selectedText, mode === "default" ? "add-default" : "add-inline");
            return;
        }
        replaceTextSmart(addPinyinAnnotations(selectedText, mode));
        return;
    } else if (id == "removePinyin") {
        if (fileType === 'word' && typeof selectedText === 'string') {
            applyWordPinyinOperation(selectedText, "remove");
            return;
        }
        if (typeof selectedText === "string" && selectionIsOrphanPinyin(selectedText)) {
            reportWordPinyinResult("pair");
            return;
        }
        replaceTextSmart(removePinyinAnnotations(selectedText));
        return;
    }
}

function replaceTranslatedText(selectedText, fileType, curLang, targetLang){
    const converter = OpenCC.Converter({ from: curLang, to: targetLang });
    let translatedArray = [];
    if(fileType === 'cell'){
        if(typeof selectedText === 'string') {
            selectedText = [[selectedText]];
        }
        for(let i = 0; i < selectedText.length; i++) {
            selectedText[i].forEach(element => {
                const translatedText = converter(element);
                translatedArray.push(translatedText);
            });
        }
    }else {
        const translatedText = converter(selectedText);
        translatedArray = textSeperator(translatedText);
    }
    replaceTextSmart(translatedArray);
}

function replaceTextSmart(replacementText){
    Asc.scope.newText = replacementText;
    window.Asc.plugin.executeMethod ("ReplaceTextSmart", [Asc.scope.newText, "\t", "\n"],
        function (isDone) {
        if (isDone) {
            reportWordPinyinResult("ok");
            return;
        }
        window.Asc.plugin.callCommand(function () {
            if (typeof Api.ReplaceTextSmart !== "function") {
                return "failed";
            }
            return Api.ReplaceTextSmart(Asc.scope.newText) ? "ok" : "failed";
        }, false, false, reportWordPinyinResult);
    });
}

// Word-mode pinyin add/remove. The replacement text is computed here. callCommand
// applies most edits with ReplaceTextSmart, which changes the selected characters
// inside the existing runs, then inserts or deletes whole paragraphs when the
// layout needs a different count. A paragraph is deleted only when this plugin
// wrote it and the selection covers that whole paragraph. Inline parentheses are
// removed from the runs when that run text is exactly the selection, so bold,
// italic, color, underline and highlight stay on the characters that remain. An
// emptied note run is removed with it, so its color does not spread onto the next
// character. A run that also holds a drawing is not rewritten: only the note
// characters are taken out, and the drawing stays.
function applyWordPinyinOperation(selectedText, mode) {
    // GetSelectedText ends with a trailing separator, so drop the resulting empty
    // entry (like textSeperator/paragraphCount do) to match GetAllParagraphs() below.
    const paragraphsText = String(selectedText).split("\n");
    if (paragraphsText.length > 0 && paragraphsText[paragraphsText.length - 1] === "") {
        paragraphsText.pop();
    }
    // Bookmarks name the paragraphs this plugin generated. The text is read first,
    // then the edit runs, so a pasted U+200B is not treated as one of ours.
    window.Asc.plugin.callCommand(function () {
        var oRange = Api.GetDocument().GetRangeBySelect();
        if (!oRange) {
            return "";
        }
        var paragraphs = oRange.GetAllParagraphs();
        if (!paragraphs) {
            return "";
        }
        var ownedById = {};
        var nameById = {};
        var doc = Api.GetDocument();
        var names = (doc.GetAllBookmarksNames && doc.GetAllBookmarksNames()) || [];
        for (var i = 0; i < names.length; i++) {
            var name = names[i];
            if (typeof name !== "string" || name.indexOf("ccPinyin") !== 0 || !doc.GetBookmarkRange) {
                continue;
            }
            var bRange = doc.GetBookmarkRange(name);
            var bParas = bRange && bRange.GetAllParagraphs ? bRange.GetAllParagraphs() : null;
            if (!bParas) {
                continue;
            }
            for (var p = 0; p < bParas.length; p++) {
                if (!bParas[p] || !bParas[p].GetInternalId) {
                    continue;
                }
                var bid = String(bParas[p].GetInternalId());
                ownedById[bid] = true;
                if (!nameById[bid]) {
                    nameById[bid] = name;
                }
            }
        }
        var flags = "";
        var found = [];
        var full = [];
        var textOptions = { Numbering: false, Math: false, NewLineSeparator: "\r", TabSymbol: "\t" };
        for (var j = 0; j < paragraphs.length; j++) {
            var id = paragraphs[j].GetInternalId ? String(paragraphs[j].GetInternalId()) : "";
            if (id && ownedById[id]) {
                flags += "1";
                found.push(nameById[id]);
            } else {
                flags += "0";
                found.push("");
            }
            var fullText = paragraphs[j].GetText ? paragraphs[j].GetText(textOptions) : "";
            full.push(typeof fullText === "string" ? fullText : "");
        }
        return JSON.stringify({ flags: flags, names: found, full: full });
    }, false, false, function (raw) {
        if (raw == null || raw === "") {
            reportWordPinyinResult("unsupported");
            return;
        }
        var parsed = null;
        if (typeof raw === "string") {
            try {
                parsed = JSON.parse(raw);
            } catch (e) {
                parsed = null;
            }
        } else if (raw && typeof raw === "object") {
            parsed = raw;
        }
        var flagString = parsed && typeof parsed.flags === "string" ? parsed.flags : "";
        var ownedFlags = [];
        if (flagString.length === paragraphsText.length) {
            for (var f = 0; f < flagString.length; f++) {
                ownedFlags.push(flagString.charAt(f) === "1");
            }
        }
        var fullTexts = parsed && Array.isArray(parsed.full) && parsed.full.length === paragraphsText.length ? parsed.full : null;
        const instructions = buildPinyinInstructions(paragraphsText, mode, ownedFlags, fullTexts);
        if (mode === "remove" && instructions.every(function (step) { return !step || step.action === "none"; })) {
            reportWordPinyinResult("unchanged");
            return;
        }
        Asc.scope.pinyinInstructions = instructions;
        // The strings the instructions were built from. callCommand needs them to find
        // the selection inside each full paragraph and keep the unselected ends.
        Asc.scope.pinyinSourceParagraphs = paragraphsText;
        Asc.scope.pinyinBookmarkNames = parsed && parsed.names ? parsed.names : [];
        window.Asc.plugin.callCommand(function () {
        var oRange = Api.GetDocument().GetRangeBySelect();
        if (!oRange) {
            return "unsupported";
        }
        var paragraphs = oRange.GetAllParagraphs();
        var instructions = Asc.scope.pinyinInstructions;
        var source = Asc.scope.pinyinSourceParagraphs;
        // Safety check: only proceed if the live paragraph count matches what was
        // computed from the plain selected text. If they disagree (e.g. an unusual
        // selection shape), do nothing rather than risk misaligned edits.
        if (!paragraphs || !instructions || !source || paragraphs.length !== instructions.length || paragraphs.length !== source.length) {
            return "unsupported";
        }

        var docForMarks = Api.GetDocument();
        var bookmarkNames = Asc.scope.pinyinBookmarkNames || [];
        var usedMarks = {};
        var markSerial = 1;
        var markList = (docForMarks.GetAllBookmarksNames && docForMarks.GetAllBookmarksNames()) || [];
        for (var b = 0; b < markList.length; b++) {
            var markName = markList[b];
            if (typeof markName !== "string") {
                continue;
            }
            usedMarks[markName] = true;
            if (markName.indexOf("ccPinyin") === 0) {
                var markNum = parseInt(markName.substring(8), 10);
                if (markNum >= markSerial) {
                    markSerial = markNum + 1;
                }
            }
        }
        function freshPinyinBookmark() {
            var created = "ccPinyin" + markSerial;
            markSerial++;
            usedMarks[created] = true;
            return created;
        }

        // Same width classes as charWidth() outside this command. callCommand cannot
        // see the plugin's functions, so the rule is repeated here.
        var WIDE = /[\u2E80-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFE30-\uFE6F\uFF00-\uFF60\uFFE0-\uFFE6]/;
        function padForPrefix(prefix) {
            var pad = "";
            for (var c = 0; c < prefix.length; c++) {
                pad += WIDE.test(prefix.charAt(c)) ? "  " : " ";
            }
            return pad;
        }

        // Where the selected text sits inside the full paragraph.
        // A selection that continues into the next paragraph owns the end of this one.
        // A selection that arrived from the previous paragraph owns the start of this one.
        // A selection inside a single paragraph must match exactly one place.
        // Returns null when the place cannot be known. The text edit still follows
        // the live selection; only the indent of a generated pinyin line is skipped.
        function locate(full, selected, index, count) {
            if (selected === "" && full !== "") {
                return null;
            }
            if (full === selected) {
                return { prefix: "", suffix: "" };
            }
            if (count === 1) {
                var at = full.indexOf(selected);
                if (at < 0) {
                    return null;
                }
                if (full.indexOf(selected, at + 1) !== -1) {
                    return null;
                }
                return { prefix: full.slice(0, at), suffix: full.slice(at + selected.length) };
            }
            if (index === 0) {
                if (selected.length > full.length || full.slice(full.length - selected.length) !== selected) {
                    return null;
                }
                return { prefix: full.slice(0, full.length - selected.length), suffix: "" };
            }
            if (index === count - 1) {
                if (full.slice(0, selected.length) !== selected) {
                    return null;
                }
                return { prefix: "", suffix: full.slice(selected.length) };
            }
            return null;
        }

        var textOptions = { Numbering: false, Math: false, NewLineSeparator: "\r", TabSymbol: "\t" };
        // One replacement per selected paragraph. Unchanged and soon-deleted paragraphs
        // keep their own selected text, so the diff does not touch them. An inserted
        // pinyin line is not part of this array: it becomes its own paragraph below.
        var replacements = [];
        var prefixByIndex = new Array(paragraphs.length);
        var directCuts = [];
        var needsSmartReplace = false;
        for (var i = 0; i < paragraphs.length; i++) {
            var instr = instructions[i];
            var selected = String(source[i]);
            if (!instr || instr.action === "none" || instr.action === "delete" || typeof instr.whole === "string") {
                replacements.push(selected);
                if (instr && typeof instr.whole === "string") {
                    needsSmartReplace = true;
                }
                continue;
            }
            if (!instr.lines || instr.lines.length === 0) {
                return "unsupported";
            }
            if (instr.action === "settext" && instr.cuts && instr.cuts.length) {
                replacements.push(selected);
                directCuts.push({ index: i, cuts: instr.cuts, text: selected });
                continue;
            }
            needsSmartReplace = true;
            replacements.push(String(instr.lines[instr.lines.length - 1]));
            if (instr.action === "insert") {
                var full = paragraphs[i].GetText(textOptions);
                var loc = typeof full === "string" ? locate(full, selected, i, paragraphs.length) : null;
                prefixByIndex[i] = loc ? loc.prefix : "";
            }
        }

        function collectApiRuns(container, runs) {
            if (!container || typeof container.GetElementsCount !== "function" || typeof container.GetElement !== "function") {
                return;
            }
            var count = container.GetElementsCount();
            for (var i = 0; i < count; i++) {
                var el = container.GetElement(i);
                if (!el) {
                    continue;
                }
                if (typeof el.GetText === "function" && (typeof el.SetText === "function" || typeof el.ClearContent === "function")) {
                    runs.push(el);
                } else if (typeof el.GetElementsCount === "function") {
                    collectApiRuns(el, runs);
                }
            }
        }

        function contentHost(run) {
            if (run && run.Run && run.Run.Content && typeof run.Run.Content.length === "number") {
                return run.Run;
            }
            if (run && run.Content && typeof run.Content.length === "number") {
                return run;
            }
            return null;
        }

        function isDrawingItem(item) {
            if (!item || typeof item !== "object") {
                return false;
            }
            if (item.Drawing || item.drawing || item.ParaDrawing || item.GraphicObj || item.GraphicObject) {
                return true;
            }
            if (item.IsDrawing === true || item.isDrawing === true) {
                return true;
            }
            if (typeof item.IsDrawing === "function" && item.IsDrawing()) {
                return true;
            }
            if (typeof item.IsPicture === "function" && item.IsPicture()) {
                return true;
            }
            if (typeof item.Get_DrawingType === "function") {
                return true;
            }
            var kind = "";
            if (typeof item.GetClassType === "function") {
                kind = String(item.GetClassType());
            } else if (typeof item.Type === "string") {
                kind = item.Type;
            }
            kind = kind.toLowerCase();
            return kind.indexOf("draw") !== -1 || kind.indexOf("image") !== -1 || kind.indexOf("picture") !== -1;
        }

        function textOfItem(item) {
            if (!item || typeof item !== "object" || isDrawingItem(item)) {
                return null;
            }
            if (typeof item.GetCodePoint === "function") {
                var code = item.GetCodePoint();
                if (typeof code === "number" && code > 0) {
                    return String.fromCodePoint(code);
                }
            }
            if (typeof item.Value === "number" && item.Value > 0) {
                return String.fromCodePoint(item.Value);
            }
            if (typeof item.Value === "string") {
                return item.Value;
            }
            if (typeof item.text === "string") {
                return item.text;
            }
            if (typeof item.Text === "string" && typeof item.GetText !== "function") {
                return item.Text;
            }
            return null;
        }

        function hostHasDrawing(run) {
            var host = contentHost(run);
            var probes = [];
            if (run && run.Run) {
                probes.push(run.Run);
            }
            if (host) {
                probes.push(host);
            }
            if (run) {
                probes.push(run);
            }
            for (var p = 0; p < probes.length; p++) {
                if (!probes[p] || typeof probes[p].GetAllDrawingObjects !== "function") {
                    continue;
                }
                try {
                    var drawings = probes[p].GetAllDrawingObjects();
                    if (drawings && drawings.length) {
                        return true;
                    }
                } catch (e) {
                }
            }
            if (!host || !host.Content) {
                return false;
            }
            for (var i = 0; i < host.Content.length; i++) {
                if (isDrawingItem(host.Content[i])) {
                    return true;
                }
            }
            return false;
        }

        function removeContentItem(host, index) {
            if (typeof host.Remove_FromContent === "function") {
                host.Remove_FromContent(index, 1);
                return true;
            }
            if (typeof host.RemoveFromContent === "function") {
                host.RemoveFromContent(index, 1);
                return true;
            }
            if (host.Content && typeof host.Content.splice === "function") {
                host.Content.splice(index, 1);
                return true;
            }
            return false;
        }

        function writeItemText(item, text) {
            if (typeof item.Value === "string" || typeof item.Value === "number") {
                item.Value = text;
                return true;
            }
            if (typeof item.text === "string") {
                item.text = text;
                return true;
            }
            if (typeof item.Text === "string" && typeof item.GetText !== "function") {
                item.Text = text;
                return true;
            }
            return false;
        }

        // SetText and ClearContent drop every item in the run, including a drawing
        // that shares it with the note. Those runs lose only the note characters.
        function cutDrawingRun(run, source, dropAt) {
            var host = contentHost(run);
            if (!host || !host.Content) {
                return false;
            }
            var content = host.Content;
            var plan = [];
            var cursor = 0;
            for (var i = 0; i < content.length; i++) {
                var piece = textOfItem(content[i]);
                if (piece == null) {
                    plan.push({ index: i, keep: true });
                    continue;
                }
                var kept = "";
                for (var c = 0; c < piece.length; c++) {
                    if (!dropAt[cursor + c]) {
                        kept += piece.charAt(c);
                    }
                }
                plan.push({ index: i, piece: piece, kept: kept });
                cursor += piece.length;
            }
            if (cursor !== source.length) {
                return false;
            }
            for (var check = 0; check < plan.length; check++) {
                var ahead = plan[check];
                if (ahead.keep || ahead.kept === ahead.piece || ahead.kept === "") {
                    continue;
                }
                if (typeof content[ahead.index].Value !== "string" && typeof content[ahead.index].Value !== "number" && typeof content[ahead.index].text !== "string" && !(typeof content[ahead.index].Text === "string" && typeof content[ahead.index].GetText !== "function")) {
                    return false;
                }
            }
            for (var step = plan.length - 1; step >= 0; step--) {
                var edit = plan[step];
                if (edit.keep || edit.kept === edit.piece) {
                    continue;
                }
                if (edit.kept === "") {
                    if (!removeContentItem(host, edit.index)) {
                        return false;
                    }
                    continue;
                }
                if (!writeItemText(content[edit.index], edit.kept)) {
                    return false;
                }
            }
            return true;
        }

        function writeRunText(run, text, source, dropAt) {
            if (hostHasDrawing(run)) {
                return cutDrawingRun(run, source, dropAt || {});
            }
            if (text === "") {
                if (typeof run.Delete === "function" && run.Delete()) {
                    return true;
                }
                if (typeof run.ClearContent === "function") {
                    run.ClearContent();
                    return true;
                }
                return false;
            }
            if (typeof run.SetText === "function") {
                run.SetText(text);
                return true;
            }
            if (typeof run.ClearContent === "function" && typeof run.AddText === "function") {
                run.ClearContent();
                return run.AddText(text) !== false;
            }
            return false;
        }

        function applyDirectCuts(jobs) {
            var plans = [];
            for (var j = 0; j < jobs.length; j++) {
                var runs = [];
                collectApiRuns(paragraphs[jobs[j].index], runs);
                if (!runs.length) {
                    return false;
                }
                var texts = [];
                var joined = "";
                var runOptions = { NewLineSeparator: "\r", TabSymbol: "\t" };
                for (var r = 0; r < runs.length; r++) {
                    var part = runs[r].GetText(runOptions);
                    if (typeof part !== "string") {
                        part = "";
                    }
                    texts.push(part);
                    joined += part;
                }
                // Each run is edited on its own, so its bold, color, underline and
                // highlight stay. A run that holds only a note is deleted.
                var offset = 0;
                if (joined !== jobs[j].text) {
                    offset = joined.indexOf(jobs[j].text);
                    if (offset < 0 || joined.indexOf(jobs[j].text, offset + 1) !== -1) {
                        return false;
                    }
                }
                var drop = {};
                var cuts = jobs[j].cuts;
                for (var c = 0; c < cuts.length; c++) {
                    for (var n = cuts[c][0]; n < cuts[c][1]; n++) {
                        drop[n + offset] = true;
                    }
                }
                var selEnd = offset + jobs[j].text.length;
                var pos = 0;
                var edits = [];
                for (var t = 0; t < runs.length; t++) {
                    var source = texts[t];
                    var kept = "";
                    var localDrop = {};
                    for (var ch = 0; ch < source.length; ch++) {
                        var at = pos + ch;
                        if (at < offset || at >= selEnd || !drop[at]) {
                            kept += source.charAt(ch);
                        } else {
                            localDrop[ch] = true;
                        }
                    }
                    pos += source.length;
                    if (kept !== source) {
                        edits.push({ run: runs[t], text: kept, source: source, drop: localDrop });
                    }
                }
                if (pos !== joined.length) {
                    return false;
                }
                plans.push(edits);
            }
            for (var plan = 0; plan < plans.length; plan++) {
                var edits = plans[plan];
                for (var e = edits.length - 1; e >= 0; e--) {
                    if (!writeRunText(edits[e].run, edits[e].text, edits[e].source, edits[e].drop)) {
                        return false;
                    }
                }
            }
            return true;
        }

        var cutsApplied = false;
        if (!needsSmartReplace && directCuts.length) {
            cutsApplied = applyDirectCuts(directCuts);
        }
        // Rebuilding the paragraph would replace bold, color, underline and highlight
        // with one run. If the notes cannot be lifted out of the existing runs, stop.
        if (!cutsApplied && !needsSmartReplace && directCuts.length) {
            return "failed";
        }
        if (!cutsApplied) {
            for (var job = 0; job < directCuts.length; job++) {
                var cutInstr = instructions[directCuts[job].index];
                replacements[directCuts[job].index] = String(cutInstr.lines[cutInstr.lines.length - 1]);
            }
            var smartDone = false;
            try {
                smartDone = typeof Api.ReplaceTextSmart === "function" && Api.ReplaceTextSmart(replacements, "\t", "\r");
            } catch (e) {
                smartDone = false;
            }
            if (!smartDone && directCuts.length) {
                cutsApplied = applyDirectCuts(directCuts);
            }
            if (!smartDone && !cutsApplied) {
                return "failed";
            }
        }

        // A partial selection of a Default character line must be rewritten from the
        // whole paragraph. ReplaceTextSmart only edits the selection, so the generated
        // spaces outside that selection would otherwise stay and no longer match.
        for (var w = paragraphs.length - 1; w >= 0; w--) {
            var wholeStep = instructions[w];
            if (!wholeStep || typeof wholeStep.whole !== "string") {
                continue;
            }
            if (!paragraphs[w] || typeof paragraphs[w].Select !== "function") {
                return "failed";
            }
            paragraphs[w].Select();
            if (!Api.ReplaceTextSmart([String(wholeStep.whole)], "\t", "\r")) {
                return "failed";
            }
        }

        // Process from the last paragraph to the first: deleting or inserting next to
        // a paragraph never changes the position of paragraphs before it that have
        // not been processed yet. settext is already applied above.
        for (var n = paragraphs.length - 1; n >= 0; n--) {
            var oPara = paragraphs[n];
            var step = instructions[n];
            if (!step || step.action === "none" || step.action === "settext") {
                continue;
            }
            if (step.action === "delete") {
                // Delete removes the whole paragraph. Skip it when the selection
                // is only part of that paragraph, so the unselected ends stay.
                // A missing U+200B or an alignment space stored as a normal space
                // is still the whole generated paragraph.
                var liveText = typeof oPara.GetText === "function" ? oPara.GetText(textOptions) : null;
                var piece = String(source[n]);
                if (typeof liveText === "string" && liveText !== piece) {
                    var at = piece ? liveText.indexOf(piece) : -1;
                    var covered = false;
                    if (at >= 0 && liveText.indexOf(piece, at + 1) === -1) {
                        var edge = liveText.slice(0, at) + liveText.slice(at + piece.length);
                        covered = edge.replace(/[\s\u200B\uFEFF]/g, "") === "";
                    } else {
                        var lettersOf = function (value) {
                            var text = String(value);
                            if (text.normalize) {
                                text = text.normalize("NFC");
                            }
                            return text.replace(/[\s\u200B\uFEFF]/g, "");
                        };
                        covered = lettersOf(liveText) === lettersOf(piece) && lettersOf(piece) !== "";
                    }
                    if (!covered) {
                        continue;
                    }
                }
                if (bookmarkNames[n] && docForMarks.DeleteBookmark) {
                    docForMarks.DeleteBookmark(bookmarkNames[n]);
                }
                oPara.Delete();
                continue;
            }
            if (step.action === "insert") {
                var prefixPad = padForPrefix(prefixByIndex[n] || "");
                var anchor = oPara;
                for (var k = step.lines.length - 2; k >= 0; k--) {
                    var line = step.lines[k];
                    if (prefixPad) {
                        // Keep the ownership marker at index 0 so a later Remove still finds it.
                        if (line.charAt(0) === "\u200B") {
                            line = "\u200B" + prefixPad + line.substring(1);
                        } else {
                            line = prefixPad + line;
                        }
                    }
                    anchor = anchor.InsertParagraph(line, "before", true);
                    // Bookmark only the generated pinyin line. The character line has no marker.
                    if (anchor && line.charAt(0) === "\u200B" && anchor.GetRange) {
                        var createdRange = anchor.GetRange();
                        if (createdRange && createdRange.AddBookmark) {
                            createdRange.AddBookmark(freshPinyinBookmark());
                        }
                    }
                }
            }
        }
        return "ok";
    }, false, false, reportWordPinyinResult);
    });
}

// Builds a per-paragraph instruction array for a pinyin operation (passed via
// Asc.scope into callCommand). Processed as a whole array, not paragraph-by-paragraph,
// because deleting a marker paragraph also requires stripping the alignment padding
// (U+2002/U+2007) from the character paragraph right after it.
// The editor often drops the leading U+200B, and it can turn the alignment
// spaces into ordinary spaces, while the paragraph itself still has them.
// That is still the whole generated paragraph. A selection inside other words,
// such as "can you" inside "BEFORE can you AFTER", is not.
function coversGeneratedParagraph(full, selected) {
    const paragraph = String(full);
    const piece = String(selected);
    if (paragraph === piece) {
        return true;
    }
    const at = piece ? paragraph.indexOf(piece) : -1;
    if (at >= 0 && paragraph.indexOf(piece, at + 1) === -1) {
        const edge = paragraph.slice(0, at) + paragraph.slice(at + piece.length);
        if (edge.replace(/[\s\u200B\uFEFF]/g, "") === "") {
            return true;
        }
        return false;
    }
    const letters = function (value) {
        return value.normalize("NFC").replace(/[\s\u200B\uFEFF]/g, "");
    };
    return letters(paragraph) === letters(piece) && letters(piece) !== "";
}

function entireParagraphSelected(full, index, selected) {
    if (!full || full[index] == null) {
        return true;
    }
    return coversGeneratedParagraph(full[index], selected);
}

function paragraphFullText(full, index, selected) {
    if (!full || full[index] == null || selected == null) {
        return null;
    }
    const fullString = String(full[index]);
    if (fullString === selected) {
        return fullString;
    }
    if (selected && fullString.indexOf(selected) !== -1) {
        return fullString;
    }
    return null;
}

function finishTextEdit(original, nextText, wholeParagraph) {
    if (wholeParagraph != null) {
        if (nextText === wholeParagraph) {
            return { action: "none" };
        }
        // The batch replace only sees the selection. `whole` is applied to the
        // entire paragraph afterwards, so the unselected ends keep the same edit.
        return { action: "settext", lines: [original], whole: nextText };
    }
    return nextText !== original ? { action: "settext", lines: [nextText] } : { action: "none" };
}

function buildPinyinInstructions(paragraphsText, mode, ownedFlags, fullTexts) {
    const flags = ownedFlags || [];
    const full = Array.isArray(fullTexts) && fullTexts.length === paragraphsText.length ? fullTexts : null;
    const standalonePinyin = paragraphsText.map((t, i) => {
        // A partial selection still belongs to the whole paragraph. Deleting that
        // paragraph would also remove the unselected text around the selection.
        if (!entireParagraphSelected(full, i, t)) {
            return false;
        }
        const next = paragraphsText[i + 1];
        if (next !== undefined && /[\u4e00-\u9fa5]/.test(String(next))) {
            return false;
        }
        const judged = full && full[i] != null ? String(full[i]) : String(t);
        return isRemovableStandalonePinyin(judged, !!flags[i]);
    });
    const isMarkerLine = paragraphsText.map((t, i) => {
        if (!entireParagraphSelected(full, i, t)) {
            return false;
        }
        const nextSelected = paragraphsText[i + 1];
        const nextFull = nextSelected === undefined ? null : paragraphFullText(full, i + 1, nextSelected);
        const next = nextFull != null ? nextFull : nextSelected;
        const judged = full && full[i] != null ? String(full[i]) : String(t);
        return isPinyinMarkerLine(judged, !!flags[i], next);
    });
    const cleanedTexts = paragraphsText.slice();
    const rewriteWhole = new Array(paragraphsText.length).fill(false);
    for (let i = 0; i < cleanedTexts.length; i++) {
        if (!isMarkerLine[i] || cleanedTexts[i + 1] === undefined) {
            continue;
        }
        // ReplaceTextSmart edits the selected text, so clean that string.
        // rewriteWhole stays false: rewriting the whole paragraph leaves these spaces.
        const pinyinLine = paragraphFullText(full, i, paragraphsText[i]) || paragraphsText[i];
        const selectedChar = paragraphsText[i + 1];
        let stripped = stripPartialAnnotation(pinyinLine, selectedChar);
        if (stripped === selectedChar || /[\u2002\u2007]/.test(stripped)) {
            const recovered = recoverUnpairedDefaultLine(selectedChar);
            if (recovered != null) stripped = recovered;
        }
        cleanedTexts[i + 1] = stripped;
    }
    const steps = cleanedTexts.map((text, i) => {
        // A generated Default line is deleted when its syllables are the readings of
        // the character paragraph under it. A pasted U+200B in ordinary text is not.
        if (isMarkerLine[i] || (mode === "remove" && standalonePinyin[i])) {
            return { action: "delete" };
        }
        const original = paragraphsText[i];
        const context = paragraphFullText(full, i, text) || text;
        const cleaned = stripInlinePinyin(text, context);
        const recovered = mode === "remove" ? recoverUnpairedDefaultLine(cleaned) : null;
        const restored = recovered != null ? recovered : cleaned;
        const wholeParagraph = rewriteWhole[i] && full ? String(full[i]) : null;
        if (mode === "remove" || !/[\u4e00-\u9fa5]/.test(restored)) {
            return finishTextEdit(original, restored, wholeParagraph);
        }
        if (mode === "add-inline") {
            return finishTextEdit(original, annotateInlineForParagraph(restored), wholeParagraph);
        }
        // "add-default" (Default layout)
        const lines = annotateDefault(restored);
        if (wholeParagraph != null) {
            const characterLine = lines[lines.length - 1];
            if (lines.length > 1) {
                return { action: "insert", lines: lines, whole: characterLine };
            }
            return finishTextEdit(original, characterLine, wholeParagraph);
        }
        return lines.length > 1 ? { action: "insert", lines: lines } : { action: "settext", lines: lines };
    }).map(function (step, index) {
        if (!step || step.action !== "settext" || !step.lines || step.lines.length !== 1 || typeof step.whole === "string") {
            return step;
        }
        // Only inline parentheses are cut out of the runs. Alignment spaces stay on
        // the normal replace path: cutting them by character index drops the wrong han.
        const cuts = pureDeletionCuts(paragraphsText[index], step.lines[0]);
        if (!cuts) {
            return step;
        }
        return { action: "settext", lines: step.lines, cuts: cuts };
    });
    // Cleaning the character line and leaving its pinyin line in place is the bug.
    // A partial Default line has one syllable per padded character, not per character
    // of the plain tail. "can you" does not match that span, so it stays.
    if (mode === "remove") {
        for (let i = 0; i < steps.length - 1; i++) {
            if (steps[i] && steps[i].action === "delete") {
                continue;
            }
            if (!entireParagraphSelected(full, i, paragraphsText[i])) {
                continue;
            }
            const nextStep = steps[i + 1];
            const nextText = nextStep && nextStep.action === "settext" && nextStep.lines ? String(nextStep.lines[nextStep.lines.length - 1]) : "";
            if (!nextText || nextText === String(paragraphsText[i + 1]) || /[\u2002\u2007]/.test(nextText) || !/[\u4e00-\u9fa5]/.test(nextText)) {
                continue;
            }
            const pinyinLine = full && full[i] != null ? String(full[i]) : String(paragraphsText[i]);
            const syllables = pinyinLineSyllables(pinyinLine) || pinyinLineSyllables(paragraphsText[i]);
            const nextFull = paragraphFullText(full, i + 1, paragraphsText[i + 1]);
            const spanHans = Math.max(
                paddedHanCount(paragraphsText[i + 1]),
                nextFull != null ? paddedHanCount(nextFull) : 0
            );
            if (syllables && spanHans && syllables.length === spanHans) {
                steps[i] = { action: "delete" };
            }
        }
    }
    return steps;
}

// Indexes of characters removed from original to produce next. Null when next is
// not what remains after deletions alone.
function pureDeletionCuts(original, next) {
    const source = String(original);
    const cleaned = String(next);
    if (cleaned === source) {
        return null;
    }
    const cuts = [];
    let read = 0;
    let start = -1;
    for (let i = 0; i < source.length; i++) {
        if (read < cleaned.length && source.charAt(i) === cleaned.charAt(read)) {
            if (start !== -1) {
                cuts.push([start, i]);
                start = -1;
            }
            read++;
        } else if (start === -1) {
            start = i;
        }
    }
    if (read !== cleaned.length) {
        return null;
    }
    if (start !== -1) {
        cuts.push([start, source.length]);
    }
    if (!cuts.length) {
        return null;
    }
    for (let c = 0; c < cuts.length; c++) {
        for (let i = cuts[c][0]; i < cuts[c][1]; i++) {
            const ch = source.charAt(i);
            if (ch !== "(" && ch !== ")" && ch !== "（" && ch !== "）" && ch !== " " && ch !== "\t" && !isPinyinLetter(ch)) {
                return null;
            }
        }
    }
    return cuts;
}

// Same per-character inline annotation as addInlinePinyin's string branch, but for a
// single already-isolated paragraph's text.
function withReading(ch, py) {
    return py ? `${ch}(${py})` : ch;
}

function annotateInlineForParagraph(text) {
    const filteredPinyins = addPinyin(text);
    let idx = 0, result = '';
    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (/[\u4e00-\u9fa5]/.test(ch)) {
            result += withReading(ch, filteredPinyins[idx]);
            idx++;
        } else {
            result += ch;
        }
    }
    return result;
}

function textSeperator(text) {
    const result = text.split(/\n/).map(line => line.replace('\r', '\n'));
    if (result.length > 0 && result[result.length - 1] === '')result.pop();
    return result;
}

function addPinyinAnnotations(text, mode) {
    const source = text;
    text = stripExistingPinyin(text);
    const annotated = mode === "default" ? addPinyinDefault(text) : addInlinePinyin(text);
    if (typeof source === "string") {
        // Pad against the paragraph count AFTER removing existing annotations, not the
        // original (pre-strip) source. Otherwise a removed annotation paragraph keeps
        // getting replaced by a blank paragraph every time the layout is switched.
        const count = paragraphCount(text);
        while (annotated.length < count) {
            annotated.push("");
        }
    }
    return annotated;
}

function stripExistingPinyin(text) {
    if (typeof text !== "string") {
        return text.map(row => row.map(cell => removePinyinFromPlain(cell).join("\n")));
    }
    return removePinyinFromPlain(text).join("\n");
}

function splitParagraphs(text) {
    const normalized = String(text).replace(/\r\n/g, "\n");
    const lines = normalized.split("\n");
    if (lines.length > 0 && lines[lines.length - 1] === "") {
        lines.pop();
    }
    return lines;
}

function paragraphCount(text) {
    return Math.max(splitParagraphs(text).length, 1);
}

function addInlinePinyin(text) {
    let result = [], idx = 0;
    if(typeof text !== 'string') {
        for(let i = 0; i < text.length; i++) {
            text[i].forEach(element => {
                const filteredPinyins = addPinyin(element);
                let tempIdx = 0, tempResult = '';
                for(let i=0; i<element.length; i++) {
                    const ch = element[i];
                    if (/[\u4e00-\u9fa5]/.test(ch)) {
                        tempResult += withReading(ch, filteredPinyins[tempIdx]);
                        tempIdx++;
                    }else {
                        tempResult += ch;
                    }
                }
                result.push(tempResult);
            });
        }
    }else {
        const filteredPinyins = addPinyin(text);
        const textArray = textSeperator(text);   
        for (let i = 0; i < textArray.length; i++) {
            const char = textArray[i];
            let lineResult = '';
            for(let j = 0; j < char.length; j++) {
                const ch = char[j];
                if (/[\u4e00-\u9fa5]/.test(ch)) {
                    lineResult += withReading(ch, filteredPinyins[idx]);
                    idx++;
                } else {
                    lineResult += ch;
                }
            }
            result.push(lineResult);
        }
    }
    return result;
}

// pinyin-pro's built-in dictionary doesn't cover every 行(háng) compound word
// (e.g. it knows 换行/行首/行尾 but not 断行); patch in the ones found to be missing manually here.
if (typeof pinyinPro.addDict === "function") {
    pinyinPro.addDict({ "断行": "duàn háng" });
}

function addPinyin(text){
    // One pinyin-pro call for the whole string.
    const source = String(text);
    const pinyins = pinyinPro.pinyin(source, {
        toneType: 'symbol', // 带音调符号
        multiple: false, // 只取一个读音（库会根据上下文判断）
        type: 'array' // 返回数组形式，和原文字一一对应
    });
    // A character outside the basic multilingual plane is two UTF-16 units.
    // pinyin-pro's array follows code points, so index into that when the lengths differ.
    const chars = Array.from(source);
    const byCodePoint = pinyins.length === chars.length;
    const units = byCodePoint ? chars : source.split("");
    const filteredPinyins = [];
    const count = Math.min(units.length, pinyins.length);
    for (let index = 0; index < count; index++) {
        const char = units[index];
        if (char && /[\u4e00-\u9fa5]/.test(char)) {
            filteredPinyins.push(pinyins[index] || "");
        }
    }
    return filteredPinyins;
}

function charWidth(ch) {
    return /[\u2E80-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFE30-\uFE6F\uFF00-\uFF60\uFFE0-\uFFE6]/.test(ch) ? 2 : 1;
}

function stringWidth(text) {
    let width = 0;
    for (let i = 0; i < text.length; i++) {
        width += charWidth(text[i]);
    }
    return width;
}

// Undo only the U+2002 / U+2007 that buildDefaultLine inserted. A figure space or
// en space that was already in the source stays: on the pinyin line it is mirrored
// by a normal space, while generated padding is a U+2002 on both lines or U+2007
// wrapped around a character whose syllable is wider.
function isPinyinLetter(ch) {
    return /[A-Za-zāáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜü]/.test(ch);
}

function syllableAhead(py, pi) {
    let i = pi;
    while (py.charAt(i) === "\u2007") {
        i++;
    }
    return isPinyinLetter(py.charAt(i));
}

function readPaddedSyllable(py, pi) {
    let i = pi;
    while (py.charAt(i) === "\u2007") {
        i++;
    }
    const start = i;
    while (isPinyinLetter(py.charAt(i))) {
        i++;
    }
    if (i === start) {
        return null;
    }
    const text = py.slice(start, i);
    while (py.charAt(i) === "\u2007") {
        i++;
    }
    return { text: text, end: i };
}

function stripGeneratedPadding(pinyinRaw, characterLine) {
    let py = String(pinyinRaw);
    if (py.charAt(0) === "\u200B") {
        py = py.slice(1);
    }
    const ch = String(characterLine);
    if (!py) {
        return ch;
    }
    let pi = 0;
    let ci = 0;
    let out = "";
    while (ci < ch.length) {
        if (py.charAt(pi) === "\u2002" && ch.charAt(ci) === "\u2002") {
            pi++;
            ci++;
            continue;
        }
        const current = ch.charAt(ci);
        const hanHere = /[\u4e00-\u9fa5]/.test(current);
        if (hanHere || (current === "\u2007" && syllableAhead(py, pi))) {
            const syl = readPaddedSyllable(py, pi);
            if (!syl) {
                out += ch.slice(ci);
                break;
            }
            const hanAt = (() => {
                let j = ci;
                while (ch.charAt(j) === "\u2007") {
                    j++;
                }
                return /[\u4e00-\u9fa5]/.test(ch.charAt(j)) ? j : -1;
            })();
            if (hanAt < 0) {
                out += ch.slice(ci);
                break;
            }
            const han = ch.charAt(hanAt);
            const gap = Math.max(stringWidth(syl.text), charWidth(han)) - charWidth(han);
            const left = Math.floor(gap / 2);
            const right = gap - left;
            if (hanAt - ci !== left) {
                out += ch.slice(ci);
                break;
            }
            let after = hanAt + 1;
            let sawRight = 0;
            while (sawRight < right && ch.charAt(after) === "\u2007") {
                after++;
                sawRight++;
            }
            if (sawRight !== right) {
                out += ch.slice(ci);
                break;
            }
            out += han;
            pi = syl.end;
            ci = after;
            continue;
        }
        out += current;
        ci++;
        const mirrored = charWidth(current);
        for (let s = 0; s < mirrored; s++) {
            if (py.charAt(pi) === " ") {
                pi++;
            } else {
                break;
            }
        }
    }
    // A character fragment that does not cover the whole pinyin line is not the
    // character line. Returning it unchanged avoids rewriting only the selected part.
    let end = pi;
    while (end < py.length) {
        const tail = py.charAt(end);
        if (tail !== " " && tail !== "\u2002" && tail !== "\u2007") {
            break;
        }
        end++;
    }
    if (end < py.length) {
        return ch;
    }
    return out;
}

// A partial Default annotation indents the pinyin line with normal spaces, one per
// column of the unselected prefix. Wide characters take two spaces, same as charWidth.
function splitPinyinIndent(pinyinRaw) {
    let py = String(pinyinRaw);
    if (py.charAt(0) === "\u200B") {
        py = py.slice(1);
    }
    let indent = 0;
    while (py.charAt(indent) === " ") {
        indent++;
    }
    return { indent: indent, body: py.slice(indent) };
}

function consumeIndent(characterLine, indent) {
    if (indent === 0) {
        return { index: 0 };
    }
    const ch = String(characterLine);
    let ci = 0;
    let width = 0;
    while (ci < ch.length && width < indent) {
        width += charWidth(ch.charAt(ci));
        ci++;
    }
    if (width !== indent) {
        return null;
    }
    return { index: ci };
}

// The pinyin line may cover only the selected half of the paragraph.
// The unselected tail stays, and only the padded span is restored.
function stripPartialAnnotation(pinyinRaw, characterLine) {
    const source = String(characterLine);
    const split = splitPinyinIndent(pinyinRaw);
    const located = consumeIndent(source, split.indent);
    if (!located || !split.body) {
        return stripGeneratedPadding(pinyinRaw, source);
    }
    const tail = source.slice(located.index);
    const strippedTail = stripGeneratedPadding(split.body, tail);
    if (strippedTail === tail) {
        return source;
    }
    return source.slice(0, located.index) + strippedTail;
}

// A character paragraph left behind after its pinyin line is gone. The padded span
// may be the whole line or only the half that was selected. 你好 X does not match
// a rebuild, so a figure space from the source stays.
function recoverUnpairedDefaultLine(text) {
    const source = String(text);
    const plain = source.replace(/[\u2002\u2007]/g, "");
    if (!plain || plain === source || !/[\u4e00-\u9fa5]/.test(plain)) {
        return null;
    }
    const whole = buildDefaultLine(plain, addPinyin(plain), 0);
    if (whole.lines && whole.lines.length === 2 && whole.lines[1] === source) {
        return plain;
    }
    // The padded span is one slice. Pads sit inside it, so the slice borders are the
    // plain characters just around the first and last alignment space. Readings come
    // from that slice alone: the rest of the paragraph would change a boundary 行.
    // One character is skipped here. A single figure space, as in 你好 X, can look
    // like that character's own padding, and the source space has to stay.
    let first = -1;
    let last = -1;
    for (let i = 0; i < source.length; i++) {
        const ch = source.charAt(i);
        if (ch === "\u2002" || ch === "\u2007") {
            if (first < 0) {
                first = i;
            }
            last = i;
        }
    }
    if (first < 0) {
        return null;
    }
    let lead = 0;
    for (let i = 0; i < first; i++) {
        if (source.charAt(i) !== "\u2002" && source.charAt(i) !== "\u2007") {
            lead++;
        }
    }
    let throughLast = 0;
    for (let i = 0; i < last; i++) {
        if (source.charAt(i) !== "\u2002" && source.charAt(i) !== "\u2007") {
            throughLast++;
        }
    }
    const starts = lead > 0 ? [lead, lead - 1] : [lead];
    const ends = throughLast < plain.length ? [throughLast, throughLast + 1] : [throughLast];
    for (let s = 0; s < starts.length; s++) {
        for (let e = 0; e < ends.length; e++) {
            const start = starts[s];
            const end = ends[e];
            if (end <= start) {
                continue;
            }
            const slice = plain.slice(start, end);
            const hanInSlice = slice.match(/[\u4e00-\u9fa5]/g);
            if (!hanInSlice || hanInSlice.length < 2) {
                continue;
            }
            const built = buildDefaultLine(slice, addPinyin(slice), 0);
            if (built.lines && built.lines.length === 2 && plain.slice(0, start) + built.lines[1] + plain.slice(end) === source) {
                return plain;
            }
        }
    }
    return null;
}

// How many Chinese characters sit in the padded span. A partial Default line
// covers only that span, not the plain tail of the same paragraph.
function paddedHanCount(source) {
    const text = String(source);
    if (!/[\u2002\u2007]/.test(text)) {
        return 0;
    }
    const plain = text.replace(/[\u2002\u2007]/g, "");
    let first = -1;
    let last = -1;
    for (let i = 0; i < text.length; i++) {
        const ch = text.charAt(i);
        if (ch === "\u2002" || ch === "\u2007") {
            if (first < 0) {
                first = i;
            }
            last = i;
        }
    }
    let lead = 0;
    for (let i = 0; i < first; i++) {
        if (text.charAt(i) !== "\u2002" && text.charAt(i) !== "\u2007") {
            lead++;
        }
    }
    let throughLast = 0;
    for (let i = 0; i < last; i++) {
        if (text.charAt(i) !== "\u2002" && text.charAt(i) !== "\u2007") {
            throughLast++;
        }
    }
    const starts = lead > 0 ? [lead, lead - 1] : [lead];
    const ends = throughLast < plain.length ? [throughLast, throughLast + 1] : [throughLast];
    for (let s = 0; s < starts.length; s++) {
        for (let e = 0; e < ends.length; e++) {
            const start = starts[s];
            const end = ends[e];
            if (end <= start) {
                continue;
            }
            const slice = plain.slice(start, end);
            const hans = slice.match(/[\u4e00-\u9fa5]/g);
            if (!hans || hans.length < 2) {
                continue;
            }
            const built = buildDefaultLine(slice, addPinyin(slice), 0);
            if (built.lines && built.lines.length === 2 && plain.slice(0, start) + built.lines[1] + plain.slice(end) === text) {
                return hans.length;
            }
        }
    }
    const whole = buildDefaultLine(plain, addPinyin(plain), 0);
    if (whole.lines && whole.lines.length === 2 && whole.lines[1] === text) {
        const hans = plain.match(/[\u4e00-\u9fa5]/g);
        return hans ? hans.length : 0;
    }
    return 0;
}

function padCenter(text, width) {
    const gap = width - stringWidth(text);
    if (gap <= 0) {
        return text;
    }
    const left = Math.floor(gap / 2);
    return "\u2007".repeat(left) + text + "\u2007".repeat(gap - left);
}

// U+200B marks a generated Default pinyin line, including in cells, where a
// bookmark cannot be stored. The marker alone is not enough: the line must be
// syllables. Removing it also needs the characters on the next line, or a
// ccPinyin bookmark in Word.
const PINYIN_DEFAULT_MARKER = "\u200B";

function isGeneratedPinyinLine(text) {
    let body = String(text);
    if (body.charAt(0) === PINYIN_DEFAULT_MARKER) {
        body = body.slice(1);
    }
    if (!body || /[\u4e00-\u9fa5]/.test(body)) {
        return false;
    }
    const tokens = body.split(/[ \u2002]+/);
    let syllables = 0;
    for (let i = 0; i < tokens.length; i++) {
        const core = tokens[i].replace(/\u2007/g, "");
        if (!core) {
            continue;
        }
        if (!isPinyinSyllable(core)) {
            return false;
        }
        syllables++;
    }
    return syllables > 0;
}

function pinyinLineSyllables(text) {
    let body = String(text);
    if (body.charAt(0) === PINYIN_DEFAULT_MARKER) {
        body = body.slice(1);
    }
    const tokens = body.split(/[ \u2002]+/);
    const syllables = [];
    for (let i = 0; i < tokens.length; i++) {
        const core = tokens[i].replace(/\u2007/g, "");
        if (!core) {
            continue;
        }
        if (!isPinyinSyllable(core)) {
            return null;
        }
        syllables.push(normalizeSyllable(core));
    }
    return syllables;
}

function pinyinLineMatchesReadings(text, nextText) {
    const syllables = pinyinLineSyllables(text);
    if (!syllables || syllables.length === 0 || !hasGeneratedMark(text)) {
        return false;
    }
    const split = splitPinyinIndent(text);
    const located = consumeIndent(String(nextText), split.indent);
    if (!located) {
        return false;
    }
    // Alignment spaces sit between the characters. Left in place they break words,
    // so 换行 is read as xíng instead of háng. A short line was built from the
    // selection alone, and the same character can get another reading once the rest
    // of the paragraph is included, so the padding itself identifies that line.
    const readable = String(nextText).slice(located.index).replace(/[\s\u200B\uFEFF]+/g, "");
    const readings = addPinyin(hanTextForReadings(readable)).map(normalizeSyllable);
    if (syllables.length > readings.length) {
        return false;
    }
    if (syllables.length < readings.length) {
        if (stripPartialAnnotation(text, nextText) !== String(nextText)) {
            return true;
        }
        // A narrow syllable adds no alignment space. The line is still ours when it
        // carries the generated marker or alignment spaces and its syllables are the
        // readings of that slice alone. "can" above 中文 does not match zhōng.
        const raw = String(text);
        if (raw.charAt(0) !== "\u200B" && !/[\u2002\u2007]/.test(raw)) {
            return false;
        }
        let head = "";
        let left = syllables.length;
        for (let c = 0; c < readable.length && left > 0; c++) {
            const ch = readable.charAt(c);
            if (/[\u4e00-\u9fa5]/.test(ch)) {
                head += ch;
                left--;
            }
        }
        if (left !== 0) {
            return false;
        }
        const headReadings = addPinyin(head).map(normalizeSyllable);
        if (headReadings.length !== syllables.length) {
            return false;
        }
        for (let i = 0; i < syllables.length; i++) {
            if (!headReadings[i] || syllables[i] !== headReadings[i]) {
                return false;
            }
        }
        return true;
    }
    for (let i = 0; i < syllables.length; i++) {
        if (!readings[i] || syllables[i] !== readings[i]) {
            return false;
        }
    }
    return true;
}

// A leading U+200B or the alignment spaces U+2002/U+2007. Ordinary spaces and
// the fact that the words are syllables do not show that this plugin wrote the line.
function hasGeneratedMark(text) {
    const raw = String(text);
    return raw.charAt(0) === PINYIN_DEFAULT_MARKER || /[\u2002\u2007]/.test(raw);
}

// A pinyin paragraph with no Chinese line under it. A ccPinyin bookmark or a
// generated mark is required. "can you" has neither, so it stays.
function isRemovableStandalonePinyin(text, owned) {
    if (!isGeneratedPinyinLine(text)) {
        return false;
    }
    return !!owned || hasGeneratedMark(text);
}

function isPinyinMarkerLine(text, owned, nextText) {
    if (!isGeneratedPinyinLine(text)) {
        return false;
    }
    if (nextText === undefined || !/[\u4e00-\u9fa5]/.test(String(nextText))) {
        return false;
    }
    if (owned) {
        return true;
    }
    return pinyinLineMatchesReadings(text, nextText);
}

function isOrphanPinyinSelection(paragraphsText, ownedFlags, fullTexts) {
    let found = false;
    for (let i = 0; i < paragraphsText.length; i++) {
        const fullLine = paragraphFullText(fullTexts, i, paragraphsText[i]);
        const line = fullLine != null ? fullLine : String(paragraphsText[i]);
        if (!isGeneratedPinyinLine(line) || !hasGeneratedMark(line)) {
            continue;
        }
        if (paragraphsText[i + 1] !== undefined) {
            continue;
        }
        found = true;
    }
    return found;
}

function selectionIsOrphanPinyin(text) {
    const lines = splitParagraphs(text);
    let found = false;
    for (let i = 0; i < lines.length; i++) {
        if (!isGeneratedPinyinLine(lines[i]) || !hasGeneratedMark(lines[i])) {
            continue;
        }
        if (lines[i + 1] !== undefined) {
            continue;
        }
        found = true;
    }
    return found;
}

// One syllable above each Chinese character. Other characters stay on the lower line.
function buildDefaultLine(line, pinyins, idx) {
    let pinyinLine = "";
    let characterLine = "";
    let hanCount = 0;
    for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (/[\u4e00-\u9fa5]/.test(ch)) {
            const syllable = pinyins[idx] || "";
            idx++;
            if (!syllable) {
                pinyinLine += " ".repeat(charWidth(ch));
                characterLine += ch;
                continue;
            }
            hanCount++;
            const width = Math.max(stringWidth(syllable), stringWidth(ch));
            if (pinyinLine && !/\s$/.test(pinyinLine)) {
                pinyinLine += "\u2002";
                characterLine += "\u2002";
            }
            pinyinLine += padCenter(syllable, width);
            characterLine += padCenter(ch, width);
        } else {
            pinyinLine += " ".repeat(charWidth(ch));
            characterLine += ch;
        }
    }
    if (!hanCount) {
        return { lines: [line], idx: idx };
    }
    // Each string becomes its own paragraph. A soft break stays in the character
    // line and is mirrored by a space on the pinyin line, so it is not a new paragraph.
    // Only the generated pinyin line is trimmed/marked; the character line keeps the
    // source text (including any trailing whitespace) untouched.
    return {
        lines: [PINYIN_DEFAULT_MARKER + pinyinLine.replace(/\s+$/, ""), characterLine],
        idx: idx
    };
}

function addPinyinDefault(text) {
    if (typeof text !== "string") {
        const result = [];
        for (let i = 0; i < text.length; i++) {
            text[i].forEach(element => {
                result.push(annotateDefault(String(element)).join("\n"));
            });
        }
        return result;
    }
    return annotateDefault(text);
}

function annotateDefault(text) {
    const lines = [];
    splitParagraphs(text).forEach(paragraph => {
        const pinyins = addPinyin(paragraph);
        const built = buildDefaultLine(paragraph, pinyins, 0);
        built.lines.forEach(part => lines.push(part));
    });
    return lines;
}

// Drop parenthetical notes so the reading is chosen from the characters themselves.
// 中(can)国 is read as 中国, not as a string that contains the letters c-a-n.
function hanTextForReadings(text) {
    let prev;
    let cur = String(text);
    do {
        prev = cur;
        cur = cur.replace(/([\u4e00-\u9fa5])[ \t\u3000]*[（(][^）)]*[）)]/g, "$1");
    } while (cur !== prev);
    return cur;
}

function followingNote(source, hanIndex) {
    let j = hanIndex + 1;
    while (j < source.length && /[ \t\u3000]/.test(source.charAt(j))) {
        j++;
    }
    const open = source.charAt(j);
    if (open !== "(" && open !== "（") {
        return null;
    }
    for (let k = j + 1; k < source.length; k++) {
        const ch = source.charAt(k);
        if (ch === ")" || ch === "）") {
            return { note: source.slice(j + 1, k), end: k };
        }
    }
    return null;
}

function sameGeneratedReading(note, reading) {
    const left = normalizeSyllable(String(note).trim());
    const right = normalizeSyllable(String(reading).trim());
    return left !== "" && left === right;
}

function stripInlinePinyin(text, context) {
    const source = String(text);
    // A parenthesis is removed only when it is the reading this plugin would
    // write for that character. (can) stays on 中, because can is not zhōng.
    // (important) stays too. (zhōng), （zhōng） and a space before the parenthesis
    // are removed when the note is that reading.
    // A note written for a selection is judged on that stretch, not on the whole
    // paragraph: 行 in 换行 is háng in the sentence and xíng when the selection
    // ends before the character that decides it.
    if (!/[\u4e00-\u9fa5]/.test(source)) {
        return source;
    }
    const hans = [];
    for (let i = 0; i < source.length; i++) {
        if (!/[\u4e00-\u9fa5]/.test(source.charAt(i))) {
            continue;
        }
        hans.push({ index: i, note: followingNote(source, i) });
    }
    const fullReadings = addPinyin(hanTextForReadings(source));
    // The selected half can sit inside a paragraph whose notes were chosen from
    // the whole paragraph. 行 is háng there and xíng in the half alone.
    let contextReadings = null;
    let hanOffset = 0;
    if (context && String(context) !== source) {
        const around = String(context);
        const at = around.indexOf(source);
        if (at >= 0 && around.indexOf(source, at + 1) === -1) {
            hanOffset = (hanTextForReadings(around.slice(0, at)).match(/[\u4e00-\u9fa5]/g) || []).length;
            contextReadings = addPinyin(hanTextForReadings(around));
        }
    }
    const stripAt = new Set();
    let span = [];
    function finishSpan() {
        if (span.length === 0) {
            return;
        }
        const last = span[span.length - 1];
        const slice = source.slice(span[0].index, last.note.end + 1);
        const readings = addPinyin(hanTextForReadings(slice));
        if (readings.length === span.length) {
            for (let n = 0; n < span.length; n++) {
                if (sameGeneratedReading(span[n].note.note, readings[n])) {
                    stripAt.add(span[n].index);
                }
            }
        }
        span = [];
    }
    for (let h = 0; h < hans.length; h++) {
        const item = hans[h];
        const syllableNote = item.note && isPinyinSyllable(item.note.note.trim());
        if (!syllableNote) {
            finishSpan();
            continue;
        }
        if (fullReadings[h] && sameGeneratedReading(item.note.note, fullReadings[h])) {
            stripAt.add(item.index);
        }
        if (contextReadings && contextReadings[hanOffset + h] && sameGeneratedReading(item.note.note, contextReadings[hanOffset + h])) {
            stripAt.add(item.index);
        }
        span.push(item);
    }
    finishSpan();
    let hanIndex = 0;
    let result = "";
    for (let i = 0; i < source.length; i++) {
        const ch = source.charAt(i);
        if (/[\u4e00-\u9fa5]/.test(ch)) {
            const item = hans[hanIndex];
            hanIndex++;
            if (item && stripAt.has(item.index) && item.note) {
                result += ch;
                i = item.note.end;
                continue;
            }
        }
        result += ch;
    }
    return result;
}

const PINYIN_SYLLABLES = new Set("a ai an ang ao ba bai ban bang bao bei ben beng bi bian biao bie bin bing bo bu ca cai can cang cao ce cen ceng cha chai chan chang chao che chen cheng chi chong chou chu chua chuai chuan chuang chui chun chuo ci cong cou cu cuan cui cun cuo da dai dan dang dao de dei den deng di dia dian diao die ding diu dong dou du duan dui dun duo e ei en eng er fa fan fang fei fen feng fo fou fu ga gai gan gang gao ge gei gen geng gong gou gu gua guai guan guang gui gun guo ha hai han hang hao he hei hen heng hm hng hong hou hu hua huai huan huang hui hun huo ji jia jian jiang jiao jie jin jing jiong jiu ju juan jue jun ka kai kan kang kao ke kei ken keng kong kou ku kua kuai kuan kuang kui kun kuo la lai lan lang lao le lei leng li lia lian liang liao lie lin ling liu lo long lou lu luan lun luo lv lve m ma mai man mang mao me mei men meng mi mian miao mie min ming miu mo mou mu n na nai nan nang nao ne nei nen neng ng ni nian niang niao nie nin ning niu nong nou nu nuan nun nuo nv nve o ou pa pai pan pang pao pei pen peng pi pian piao pie pin ping po pou pu qi qia qian qiang qiao qie qin qing qiong qiu qu quan que qun ran rang rao re ren reng ri rong rou ru rua ruan rui run ruo sa sai san sang sao se sen seng sha shai shan shang shao she shei shen sheng shi shou shu shua shuai shuan shuang shui shun shuo si song sou su suan sui sun suo ta tai tan tang tao te tei teng ti tian tiao tie ting tong tou tu tuan tui tun tuo wa wai wan wang wei wen weng wo wu xi xia xian xiang xiao xie xin xing xiong xiu xu xuan xue xun ya yan yang yao ye yi yin ying yo yong you yu yuan yue yun za zai zan zang zao ze zei zen zeng zha zhai zhan zhang zhao zhe zhei zhen zheng zhi zhong zhou zhu zhua zhuai zhuan zhuang zhui zhun zhuo zi zong zou zu zuan zui zun zuo".split(" "));

function normalizeSyllable(token) {
    return String(token).normalize("NFC").toLowerCase()
        .replace(/[āáǎà]/g, "a")
        .replace(/[ēéěè]/g, "e")
        .replace(/[īíǐì]/g, "i")
        .replace(/[ōóǒò]/g, "o")
        .replace(/[ūúǔù]/g, "u")
        .replace(/[ǖǘǚǜü]/g, "v");
}

function isPinyinSyllable(token) {
    const syllable = normalizeSyllable(token);
    if (PINYIN_SYLLABLES.has(syllable)) {
        return true;
    }
    return syllable.length > 1 && syllable.endsWith("r") && PINYIN_SYLLABLES.has(syllable.slice(0, -1));
}

function removePinyinFromPlain(text) {
    const lines = splitParagraphs(text);
    const kept = [];
    for (let i = 0; i < lines.length; i++) {
        // Same rule as buildPinyinInstructions: a leading U+200B is only a generated
        // pinyin line when the rest is syllables and the next line has Chinese.
        if (isPinyinMarkerLine(lines[i], false, lines[i + 1])) {
            const next = lines[i + 1];
            if (next !== undefined) {
                lines[i + 1] = stripPartialAnnotation(lines[i], next);
            }
            continue;
        }
        const cleaned = stripInlinePinyin(lines[i]);
        const recovered = recoverUnpairedDefaultLine(cleaned);
        kept.push(recovered != null ? recovered : cleaned);
    }
    // No padding back to the original paragraph count: ReplaceTextSmart replaces the
    // whole selection with this array, so a removed "default" pinyin paragraph should
    // simply disappear rather than leave a blank paragraph behind.
    return kept;
}

function removePinyinAnnotations(textArray) {
    if (typeof textArray !== "string") {
        const ans = [];
        for (let i = 0; i < textArray.length; i++) {
            textArray[i].forEach(element => {
                ans.push(removePinyinFromPlain(element).join("\n"));
            });
        }
        return ans;
    }
    return removePinyinFromPlain(textArray);
}

function reportWordPinyinResult(result) {
    if (result === "ok" || result === true) {
        showToast(tr("Success!"));
        return;
    }
    if (result === "failed") {
        showToast(tr("The text could not be updated."), "#FFAA00", 4000);
        return;
    }
    if (result === "pair") {
        showToast(tr("Select the pinyin line and the Chinese text under it together."), "#FFAA00", 4000);
        return;
    }
    if (result === "unchanged") {
        showToast(tr("No pinyin was found in the selection."), "#FFAA00", 4000);
        return;
    }
    showToast(tr("This selection can't be changed. Select whole paragraphs or one continuous piece of text."), "#FFAA00", 4000);
}

function showToast(message, backgroundColor = "#4BB543", duration = 1000) {
    const toast = document.getElementById('toast');
    toast.style.backgroundColor = backgroundColor
    toast.textContent = message; // 更新文字
    toast.classList.add('show'); // 显示

    setTimeout(() => {
        toast.classList.remove('show');
    }, duration);
}

window.Asc.plugin.onTranslate = function() {
    document.getElementById("operatorHints").innerHTML = tr("Please select the text you want to operate and click the corresponding operation button.");
    document.getElementById("button1").innerHTML = tr("Simplified → Traditional");
    document.getElementById("button2").innerHTML = tr("Traditional → Simplified");
    document.getElementById("button3").innerHTML = tr("Add pinyin");
    document.getElementById("button4").innerHTML = tr("Remove pinyin");
    document.getElementById("pinyinLayoutLabel").innerHTML = tr("Pinyin layout");
    document.getElementById("pinyinInlineLabel").innerHTML = tr("Inline");
    document.getElementById("pinyinDefaultLabel").innerHTML = tr("Default");
    document.getElementById("operatorHints2").innerHTML = tr("Tip: Right-click after selecting text to use the function directly.");
}
