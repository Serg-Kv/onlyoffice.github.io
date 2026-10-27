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
                    id: "addPinyinAbove",
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
        id = "addPinyinAbove";
    }
    handleTextOperation(id);
}

function core(id, selectedText) {
    const fileType = window.Asc.plugin.info.editorType;
    if (id == "convertChineseToFan") {
        replaceTranslatedText(selectedText, fileType, 'cn', 'tw');
    } else if (id == "convertFanToChinese") {
        replaceTranslatedText(selectedText, fileType, 'tw', 'cn');
    } else if (id == "addPinyin" || id == "addPinyinAbove") {
        const mode = id == "addPinyinAbove" ? "default" : "inline";
        replaceTextSmart(addPinyinAnnotations(selectedText, mode));
    } else if (id == "removePinyin") {
        replaceTextSmart(removePinyinAnnotations(selectedText));
    }
    showToast(tr("Success!"));
}

function replaceTranslatedText(selectedText, fileType, curLang, targetLang){
    const converter = OpenCC.Converter({ from: curLang, to: targetLang });   // what is openCC.converters
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
        if (!isDone)
            window.Asc.plugin.callCommand (function () {
                Api.ReplaceTextSmart (Asc.scope.newText);
            });
    });
}

function textSeperator(text) {
    const result = text.split(/\n/).map(line => line.replace('\r', '\n'));
    if (result.length > 0 && result[result.length - 1] === '')result.pop();
    return result;
}

function addPinyinAnnotations(text, mode) {
    const source = text;
    text = stripExistingPinyin(text);
    const annotated = mode === "default" ? addPinyinAbove(text) : addInlinePinyin(text);
    if (typeof source === "string") {
        const count = paragraphCount(source);
        while (annotated.length < count) {
            annotated.push("");
        }
    }
    return annotated;
}

function stripExistingPinyin(text) {
    if (typeof text !== "string") {
        return text.map(row => row.map(cell => removePinyinFromPlain(cell, false).join("\n")));
    }
    return removePinyinFromPlain(text, false).join("\n");
}

function paragraphCount(text) {
    const normalized = String(text).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    const lines = normalized.split("\n");
    if (lines.length > 0 && lines[lines.length - 1] === "") {
        lines.pop();
    }
    return Math.max(lines.length, 1);
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
                        tempResult += `${ch}(${filteredPinyins[tempIdx]})`;
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
                    const py = filteredPinyins[idx];
                    lineResult += `${ch}(${py})`;
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

function addPinyin(text){
    // 1. 使用 pinyin-pro 一次性获取整段文字的拼音数组
    const pinyins = pinyinPro.pinyin(text, {
        toneType: 'symbol', // 带音调符号
        multiple: false, // 只取一个读音（库会根据上下文判断）
        type: 'array' // 返回数组形式，和原文字一一对应
    });
    const filteredPinyins = pinyins.filter((py, index) => {
        // Get the corresponding character from the original text
        const char = text[index];
        // Only keep pinyin for Chinese characters
        return char && /[\u4e00-\u9fa5]/.test(char);
    });
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

function padCenter(text, width) {
    const gap = width - stringWidth(text);
    if (gap <= 0) {
        return text;
    }
    const left = Math.floor(gap / 2);
    return "\u2007".repeat(left) + text + "\u2007".repeat(gap - left);
}

// One syllable above each Chinese character. Other characters stay on the lower line.
function buildAboveLine(line, pinyins, idx) {
    let pinyinLine = "";
    let characterLine = "";
    let hanCount = 0;
    for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (/[\u4e00-\u9fa5]/.test(ch)) {
            const syllable = pinyins[idx] || "";
            idx++;
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
    // Each string becomes its own paragraph. A break inside one string is inserted as a space.
    return {
        lines: [pinyinLine.replace(/\s+$/, ""), characterLine.replace(/\s+$/, "")],
        idx: idx
    };
}

function addPinyinAbove(text) {
    if (typeof text !== "string") {
        const result = [];
        for (let i = 0; i < text.length; i++) {
            text[i].forEach(element => {
                result.push(annotateAbove(String(element)).join("\n"));
            });
        }
        return result;
    }
    return annotateAbove(text);
}

function annotateAbove(text) {
    const paragraphs = textSeperator(text);
    const lines = [];
    paragraphs.forEach(paragraph => {
        const pinyins = addPinyin(paragraph);
        let idx = 0;
        paragraph.split(/\r|\n/).forEach(line => {
            const built = buildAboveLine(line, pinyins, idx);
            idx = built.idx;
            built.lines.forEach(part => lines.push(part));
        });
    });
    return lines;
}

function stripInlinePinyin(text) {
    return text.replace(/([\u4e00-\u9fa5])\([^)]*\)/g, "$1");
}

function countHan(text) {
    const matches = text.match(/[\u4e00-\u9fa5]/g);
    return matches ? matches.length : 0;
}

const PINYIN_SYLLABLES = new Set("a ai an ang ao ba bai ban bang bao bei ben beng bi bian biao bie bin bing bo bu ca cai can cang cao ce cen ceng cha chai chan chang chao che chen cheng chi chong chou chu chua chuai chuan chuang chui chun chuo ci cong cou cu cuan cui cun cuo da dai dan dang dao de dei den deng di dia dian diao die ding diu dong dou du duan dui dun duo e ei en eng er fa fan fang fei fen feng fo fou fu ga gai gan gang gao ge gei gen geng gong gou gu gua guai guan guang gui gun guo ha hai han hang hao he hei hen heng hm hng hong hou hu hua huai huan huang hui hun huo ji jia jian jiang jiao jie jin jing jiong jiu ju juan jue jun ka kai kan kang kao ke kei ken keng kong kou ku kua kuai kuan kuang kui kun kuo la lai lan lang lao le lei leng li lia lian liang liao lie lin ling liu lo long lou lu luan lun luo lv lve m ma mai man mang mao me mei men meng mi mian miao mie min ming miu mo mou mu n na nai nan nang nao ne nei nen neng ng ni nian niang niao nie nin ning niu nong nou nu nuan nun nuo nv nve o ou pa pai pan pang pao pei pen peng pi pian piao pie pin ping po pou pu qi qia qian qiang qiao qie qin qing qiong qiu qu quan que qun ran rang rao re ren reng ri rong rou ru rua ruan rui run ruo sa sai san sang sao se sen seng sha shai shan shang shao she shei shen sheng shi shou shu shua shuai shuan shuang shui shun shuo si song sou su suan sui sun suo ta tai tan tang tao te tei teng ti tian tiao tie ting tong tou tu tuan tui tun tuo wa wai wan wang wei wen weng wo wu xi xia xian xiang xiao xie xin xing xiong xiu xu xuan xue xun ya yan yang yao ye yi yin ying yo yong you yu yuan yue yun za zai zan zang zao ze zei zen zeng zha zhai zhan zhang zhao zhe zhei zhen zheng zhi zhong zhou zhu zhua zhuai zhuan zhuang zhui zhun zhuo zi zong zou zu zuan zui zun zuo".split(" "));

function normalizeSyllable(token) {
    return token.toLowerCase()
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

function isPinyinOnlyLine(line) {
    const trimmed = line.trim();
    if (!trimmed || /[\u4e00-\u9fa5]/.test(line)) {
        return false;
    }
    const tokens = trimmed.split(/\s+/);
    return tokens.length > 0 && tokens.every(isPinyinSyllable);
}

function removePinyinFromPlain(text, clearRemovedParagraphs) {
    const normalized = String(text).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    const lines = normalized.split("\n");
    if (lines.length > 0 && lines[lines.length - 1] === "") {
        lines.pop();
    }
    const cleaned = lines.map(stripInlinePinyin);
    const kept = [];
    let removed = 0;
    for (let i = 0; i < cleaned.length; i++) {
        const next = cleaned[i + 1];
        const tokens = cleaned[i].trim().split(/\s+/).filter(Boolean);
        if (next !== undefined && isPinyinOnlyLine(cleaned[i]) && tokens.length === countHan(next)) {
            cleaned[i + 1] = next.replace(/[\u2002\u2007]/g, "");
            removed++;
            continue;
        }
        kept.push(cleaned[i]);
    }
    if (clearRemovedParagraphs) {
        for (let i = 0; i < removed; i++) {
            kept.push("");
        }
    }
    return kept;
}

function removePinyinAnnotations(textArray) {
    if (typeof textArray !== "string") {
        const ans = [];
        for (let i = 0; i < textArray.length; i++) {
            textArray[i].forEach(element => {
                ans.push(removePinyinFromPlain(element, false).join("\n"));
            });
        }
        return ans;
    }
    return removePinyinFromPlain(textArray, true);
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
