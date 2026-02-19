// scripts/code.js
(function (window, undefined) {
    // ---------------- 初始化 ----------------
    window.Asc.plugin.init = function () {
        // this.executeMethod("AddToolbarMenuItem", [getToolbarItems()]);
        bindToolbarEvents.call(this);
    };

    // 子窗口句柄
    let __toolbarAdded = false;
    let winSetting = null;
    let winOptions = null;
    let winReport = null;
    let winInfo = null;
    let reportApplyTimer = null; // ⬅️ 新增

    function readJSON(key, fallback) {
        try {
            const raw = localStorage.getItem(key);
            if (raw === null || raw === undefined || raw === "") return fallback;
            return JSON.parse(raw);
        } catch (e) {
            return fallback;
        }
    }

    // ======= PPT 调试探针（可选）=======
    // 1) 在控制台执行 __pptProbe()：列出当前页形状及段落数
    window.__pptProbe = function () {
        const plugin = window.Asc.plugin;
        plugin.callCommand(
            function () {
                function getContent(d) {
                    try {
                        if (d && typeof d.GetContent === "function") return d.GetContent();
                        if (d && typeof d.GetDocContent === "function")
                            return d.GetDocContent();
                    } catch (e) { }
                    return null;
                }

                var sel =
                    typeof Api.GetSelection === "function" ? Api.GetSelection() : null;
                var shapes =
                    sel && typeof sel.GetShapes === "function" ? sel.GetShapes() : null;
                if (!Array.isArray(shapes)) shapes = shapes ? [shapes] : [];
                if (shapes.length === 0) {
                    var pres =
                        typeof Api.GetPresentation === "function"
                            ? Api.GetPresentation()
                            : null;
                    // var slide = pres?.GetCurrentSlide?.();
                    // var all = slide?.GetAllObjects?.();
                    var slide =
                        pres && typeof pres.GetCurrentSlide === "function"
                            ? pres.GetCurrentSlide()
                            : null;
                    var all =
                        slide && typeof slide.GetAllObjects === "function"
                            ? slide.GetAllObjects()
                            : null;
                    if (Array.isArray(all)) shapes = all;
                }
                var stats = [];
                for (var i = 0; i < shapes.length; i++) {
                    var dc = getContent(shapes[i]);
                    var n = dc?.GetAllParagraphs?.().length || 0;
                    stats.push({ shape: i, paras: n });
                }
                Asc.scope.__stats = stats;
            },
            false,
            true,
            function () {
                getInfoModal("PPT Probe: shapes=" + ((Asc.scope.__stats || []).length || 0));
            }
        );
    };

    // 2) Execute __pptDryApply() in the console: attempting to write a test string into the first paragraph of the first shape.
    window.__pptDryApply = function () {
        const plugin = window.Asc.plugin;
        plugin.callCommand(
            function () {
                function getContent(d) {
                    try {
                        if (d && typeof d.GetContent === "function") return d.GetContent();
                        if (d && typeof d.GetDocContent === "function")
                            return d.GetDocContent();
                    } catch (e) { }
                    return null;
                }
                var pres =
                    typeof Api.GetPresentation === "function"
                        ? Api.GetPresentation()
                        : null;
                var slide = pres?.GetCurrentSlide?.();
                var all = slide?.GetAllObjects?.();
                var shapes = Array.isArray(all) ? all : [];
                var dc = shapes[0] ? getContent(shapes[0]) : null;
                var p0 = dc?.GetAllParagraphs?.()[0];
                if (p0?.Select && typeof Api.ReplaceTextSmart === "function") {
                    p0.Select();
                    Api.ReplaceTextSmart(
                        ["[SMART-TEST]\n(能看到这一行说明 PPT 回填链路 OK)"],
                        "\t",
                        "\n"
                    );
                    Asc.scope.__ok = true;
                } else Asc.scope.__ok = false;
            },
            false,
            true,
            function () {
                getInfoModal(Asc.scope.__ok ? "Dry apply OK" : "Dry apply failed");
            }
        );
    };

    const REPORT_CLOSE_MS = 1000;
    // 放在 (function (window, undefined) { 之后的任意顶层位置
    const tr = (s) =>
        window.Asc && window.Asc.plugin ? window.Asc.plugin.tr(s) : s;
    // —— 本地化回调：词典就绪后，刷新工具栏文本与提示 ——
    window.Asc.plugin.onTranslate = function () {
        getInfoModal(tr("The plugin is ready, the toolbar menu has been updated. Please go to the Chinese-Auto Format Tool tab above to use the formatting features."));
        // ……你原来的 setText / 提示等本地化代码（如果有）……
        const items = getToolbarItems(); // 这里的 tabs[0].text 要用 tr("Chinese Auto-format Tool")
        if (!__toolbarAdded) {
            // ✅ First time: Create the tab with current language, tab title will use the translated text
            window.Asc.plugin.executeMethod("AddToolbarMenuItem", [items]);
            __toolbarAdded = true;
        } else {
            // ✅ Subsequent language switches: Only update button text/tooltips
            window.Asc.plugin.executeMethod("UpdateToolbarMenuItem", [items]);
        }
    };

    // report.html → 主窗口：接收 Asc.PluginWindow 发来的消息
    // 仅替换 onMessage 这一段
    window.Asc.plugin.onMessage = function (data) {
        if (!data) return;

        if (data.type === "apply-now" && Array.isArray(data.lines)) {
            const plugin = window.Asc.plugin;
            const isPPT = plugin.info && plugin.info.editorType === "slide";

            if (isPPT) {
                // —— PPT 专用：把“行”折成按形状的块，然后一形状一块回填 —— //
                try {
                    const raw = JSON.stringify(data.lines);
                    const blocks =
                        typeof parseBlocks === "function"
                            ? parseBlocks(raw)
                            : (function foldByBlank(lines) {
                                const out = [];
                                let buf = [];
                                for (const ln of lines) {
                                    if (ln === "") {
                                        if (buf.length) {
                                            out.push(buf.join("\n"));
                                            buf = [];
                                        }
                                    } else buf.push(ln);
                                }
                                if (buf.length) out.push(buf.join("\n"));
                                return out.map((b) => b.replace(/\n+$/, "").trim());
                            })(data.lines);

                    if (typeof applyPptBlocks === "function" && blocks.length) {
                        applyPptBlocks(blocks); // 内部自带 callCommand 与段落选中
                    } else {
                        plugin.executeMethod("ShowError", [ tr("Nothing to apply for shapes.")]);
                    }

                    // 迟一点关掉报告窗，避免与命令体回调冲突
                    setTimeout(() => {
                        try {
                            if (winReport) {
                                winReport.close();
                                winReport = null;
                            }
                        } catch (e) { }
                    }, 800);
                } catch (e) {
                    plugin.executeMethod("ShowError", ["PPT apply failed: " + e.message]);
                }
                return; // ✅ 已处理 PPT
            }

            // —— Word / Excel 仍走原逻辑 —— //
            Asc.scope.convertedLines = data.lines;
            window.Asc.plugin.callCommand(
                function () {
                    if (Asc.scope.convertedLines) {
                        Api.ReplaceTextSmart(Asc.scope.convertedLines);
                    }
                },
                false,
                true,
                function () {
                    try {
                        localStorage.setItem(
                            "currentSelectionLines",
                            JSON.stringify(data.lines)
                        );
                    } catch (e) { }
                    setTimeout(() => {
                        if (winReport) {
                            try {
                                winReport.close();
                            } catch (e) { }
                            winReport = null;
                        }
                    }, 800);
                }
            );
        }
    };

    // 生成绝对 URL
    function resolveUrl(path) {
        try {
            return new URL(path, window.location.href).toString();
        } catch (e) {
            return path;
        }
    }

    // ---------------- 统一进入报告流程 ----------------
    function proceedToReport(selectedText) {
        const editorType = window.Asc.plugin.info.editorType || "word";
        let results;
        try {
            // Depends on runFormatCheck(text) in scripts/formatChecker.js
            results = runFormatCheck(selectedText, editorType);
        } catch (e) {
            window.Asc.plugin.executeMethod("ShowError", [
                tr("Detection failed: formatChecker.js is missing or has an error"),
            ]);
            return;
        }

        const lines =
            editorType === "cell"
                ? selectedText.split(/\t|\n/)
                : selectedText.split(/\n/); // 不过滤空行，保持行数一致
        const fixed = results.map((r) => r.fixed);
        const report = results.filter((r) => r.errors && r.errors.length > 0);

        if(lines[lines.length - 1] === ""){
            lines.pop();
            fixed.pop();
        }
        if (lines.length !== fixed.length) {
            window.Asc.plugin.executeMethod("ShowError", [
                tr("Conversion failed: paragraph count mismatch"),
            ]);
            return;
        }
        if (report.length === 0) {
            window.Asc.plugin.executeMethod("ShowError", [
                tr("No fixable issues found"),
            ]);
            return;
        }

        // 供报告页与撤销/连续修改使用的共享数据
        localStorage.setItem("zhlintReport", JSON.stringify(report, null, 2));
        localStorage.setItem("originalLines", JSON.stringify(lines));
        localStorage.setItem("fixedLines", JSON.stringify(fixed));
        localStorage.setItem("currentSelectionLines", JSON.stringify(lines)); // 基线：首次进入报告即为原文
        localStorage.removeItem("batchReplaceResult");
        localStorage.removeItem("zhlintUndoLines");

        if (winReport) {
            try {
                winReport.close();
            } catch (e) { }
            winReport = null;
        }
        winReport = new window.Asc.PluginWindow();
        winReport.show({
            url: resolveUrl("panels/report.html"),
            description: tr("Formatting report"),
            isModal: false,
            isVisual: true,
            size: [720, 480],
            EditorsSupport: ["word", "cell", "slide"],
            buttons: [
                { text: tr("Apply & Save"), primary: true },
                { text: tr("Cancel"), primary: false },
            ],
        });
    }

    // 打开选项页后轮询等待结果
    function watchSpaceOptions(selectedText) {
        localStorage.removeItem("spaceOptReady");

        const timer = setInterval(() => {
            const ready = localStorage.getItem("spaceOptReady") === "1";
            const open = localStorage.getItem("spaceOptOpen") === "1";

            if (ready) {
                clearInterval(timer);
                try {
                    if (winOptions) {
                        winOptions.close();
                        winOptions = null;
                    }
                } catch (e) { }
                localStorage.removeItem("spaceOptReady");
                localStorage.removeItem("spaceOptOpen");
                proceedToReport(selectedText);
            } else if (!open) {
                clearInterval(timer);
            }
        }, 80);
    }

    // ---------------- 绑定工具栏按钮 ----------------
    function bindToolbarEvents() {
        // A. 强制转全角
        this.attachToolbarMenuClickEvent("quanjiao", function () {
            const plugin = window.Asc.plugin;

            // 供命令体读取的配置
            Asc.scope.__punct__ = {
                map: {
                    ",": "，",
                    ";": "；",
                    ":": "：",
                    ".": "。",
                    '"': "”",
                    "'": "’",
                    "-": "—",
                    "–": "—",
                    $: "＄",
                    "¥": "￥",
                    "£": "￡",
                    "¢": "￠",
                    "<": "《",
                    ">": "》",
                    "(": "（",
                    ")": "）",
                    "/": "／",
                    "?": "？",
                    "!": "！",
                },
                // settings: { punctuation: JSON.parse(localStorage.getItem("selectedPunctuation") || "[]") }
                settings: { punctuation: readJSON("selectedPunctuation", []) },
            };

            // 选中文本：优先用 ReplaceTextSmart（保留样式）
            const props = {
                Numbering: false,
                Math: false,
                TableCellSeparator: "\t",
                TableRowSeparator: "\n",
                ParaSeparator: "\n",
                TabSymbol: "\t",
                NewLineSeparator: "\r",
            };
            plugin.executeMethod("GetSelectedText", [props], function (t) {
                // const picked = (t || "").replace(/\r\n?/g, "\n");
                const picked = t || "";

                // 公用转换（编辑器侧，非命令体）
                const esc = (x) => x.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
                const convertLine = (line) => {
                    if (!line) return line;
                    const map = Asc.scope.__punct__.map;
                    const on = Asc.scope.__punct__.settings.punctuation || [];
                    let v = line.replace(/(?:\.{3,}|…+)/g, "……");
                    for (const [half, full] of Object.entries(map)) {
                        if (on.length === 0 || on.includes(full))
                            v = v.replace(new RegExp(esc(half), "g"), full);
                    }
                    return v;
                };

                if (picked.trim()) {
                    const out = picked.split(/\t|\n/).map(convertLine);
                    if(out[out.length - 1] === "") out.pop();
                    Asc.scope._lines = out.map(l => l.replace(/\r/g, "\r\n"));
                    plugin.callCommand(
                        function () {
                            if ( Asc.scope._lines && typeof Api.ReplaceTextSmart === "function") {
                                Api.ReplaceTextSmart(Asc.scope._lines); // 保留原样式地替换
                            }
                        },
                        false,
                        true,
                        function () {
                            getInfoModal(tr("Converted selection: ") + out.length + tr(" line(s)."));
                        }
                    );
                    return; // 已处理，退出
                }

                // =============== Excel 分支（命令体内重建转换函数！） ===============
                plugin.callCommand(
                    function () {
                        function escIn(x) {
                            return x.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
                        }
                        function convertIn(line) {
                            if (!line) return line;
                            var m = Asc.scope.__punct__.map;
                            var on = Asc.scope.__punct__.settings.punctuation || [];
                            var v = line.replace(/(?:\.{3,}|…+)/g, "……");
                            for (var k in m) {
                                if (!m.hasOwnProperty(k)) continue;
                                var full = m[k];
                                if (on.length === 0 || on.indexOf(full) !== -1)
                                    v = v.replace(new RegExp(escIn(k), "g"), full);
                            }
                            return v;
                        }

                        try {
                            if (typeof Api.GetActiveSheet === "function") {
                                var ws = Api.GetActiveSheet();
                                var rng = ws && ws.GetSelection && ws.GetSelection();
                                if (!rng && typeof Api.GetSelection === "function")
                                    rng = Api.GetSelection();
                                if (
                                    rng &&
                                    typeof rng.GetValue === "function" &&
                                    typeof rng.SetValue === "function"
                                ) {
                                    var val = rng.GetValue(); // string 或 二维数组 :contentReference[oaicite:4]{index=4}
                                    var changed = false, out;

                                    function convCell(v) {
                                        if (typeof v !== "string") return v;
                                        var nv = convertIn(v);
                                        if (nv !== v) changed = true;
                                        return nv;
                                    }

                                    if (Array.isArray(val)) {
                                        out = [];
                                        for (var r = 0; r < val.length; r++) {
                                            var row = val[r];
                                            if (Array.isArray(row)) {
                                                var nr = new Array(row.length);
                                                for (var c = 0; c < row.length; c++)
                                                    nr[c] = convCell(row[c]);
                                                out.push(nr);
                                            } else out.push(convCell(row));
                                        }
                                    } else out = convCell(val);
                                    if (changed) {
                                        rng.SetValue(out);
                                        return true;
                                    } // 写回 :contentReference[oaicite:5]{index=5}
                                    return false;
                                }
                            }
                        } catch (e) { }
                    },
                    false,
                    true,
                    function (returnValue) {
                        if (returnValue) {
                            getInfoModal(tr("Converted punctuation in selected cells."));
                            return; // Excel 成功，结束
                        }

                        // =============== PPT 分支（命令体内重建转换函数！） ===============
                        plugin.callCommand(
                            function () {
                                function escIn(x) {
                                    return x.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
                                }
                                function convertIn(line) {
                                    if (!line) return line;
                                    var m = Asc.scope.__punct__.map;
                                    var on = Asc.scope.__punct__.settings.punctuation || [];
                                    var v = line.replace(/(?:\.{3,}|…+)/g, "……");
                                    for (var k in m) {
                                        if (!m.hasOwnProperty(k)) continue;
                                        var full = m[k];
                                        if (on.length === 0 || on.indexOf(full) !== -1)
                                            v = v.replace(new RegExp(escIn(k), "g"), full);
                                    }
                                    return v;
                                }

                                var sel =
                                    typeof Api.GetSelection === "function"
                                        ? Api.GetSelection()
                                        : null; // 选区（演示）:contentReference[oaicite:6]{index=6}
                                var shapes =
                                    sel && typeof sel.GetShapes === "function"
                                        ? sel.GetShapes()
                                        : null; // 被选中的图形
                                if (!Array.isArray(shapes)) shapes = shapes ? [shapes] : [];

                                if (shapes.length === 0) {
                                    // 兜底：取当前页所有对象
                                    var pres =
                                        typeof Api.GetPresentation === "function"
                                            ? Api.GetPresentation()
                                            : null;
                                    var slide =
                                        pres && typeof pres.GetCurrentSlide === "function"
                                            ? pres.GetCurrentSlide()
                                            : null;
                                    if (slide && typeof slide.GetAllObjects === "function") {
                                        var all = slide.GetAllObjects();
                                        if (Array.isArray(all)) {
                                            var chosen = [];
                                            for (var i = 0; i < all.length; i++) {
                                                var o = all[i];
                                                try {
                                                    if (
                                                        o &&
                                                        typeof o.IsSelected === "function" &&
                                                        o.IsSelected()
                                                    )
                                                        chosen.push(o);
                                                } catch (e) { }
                                            }
                                            shapes = chosen.length ? chosen : all;
                                        }
                                    }
                                }

                                var hit = 0;
                                function getContent(draw) {
                                    try {
                                        if (draw && typeof draw.GetContent === "function")
                                            return draw.GetContent(); // 新版 :contentReference[oaicite:7]{index=7}
                                        if (draw && typeof draw.GetDocContent === "function")
                                            return draw.GetDocContent(); // 旧版 :contentReference[oaicite:8]{index=8}
                                    } catch (e) { }
                                    return null;
                                }

                                for (var sIdx = 0; sIdx < shapes.length; sIdx++) {
                                    var dc = getContent(shapes[sIdx]); // ApiDocumentContent
                                    if (!dc || typeof dc.GetAllParagraphs !== "function")
                                        continue;
                                    var paras = dc.GetAllParagraphs(); // 段落数组 :contentReference[oaicite:9]{index=9}
                                    if (!paras || !paras.length) continue;

                                    var changed = false;
                                    for (var pIdx = 0; pIdx < paras.length; pIdx++) {
                                        var p = paras[pIdx];
                                        var old =
                                            p && typeof p.GetText === "function"
                                                ? p.GetText({
                                                    Numbering: false,
                                                    Math: false,
                                                    NewLineSeparator: "\n",
                                                    TabSymbol: "\t",
                                                })
                                                : "";
                                        if (!old) continue;

                                        if (/[,\.\-:;'"$¥£¢<>()[\]/?!]|…/.test(old)) {
                                            var neo = convertIn(old);
                                            if (
                                                neo !== old &&
                                                typeof p.Select === "function" &&
                                                typeof Api.ReplaceTextSmart === "function"
                                            ) {
                                                p.Select(); // 选中段落
                                                Api.ReplaceTextSmart([neo], "\t", "\n"); // 保留样式地替换 :contentReference[oaicite:10]{index=10}
                                                changed = true;
                                            }
                                        }
                                    }
                                    if (changed) hit++;
                                }
                                return hit > 0;
                            },
                            false,
                            true,
                            function (returnValue) {
                                if (returnValue)
                                    getInfoModal(tr("Converted punctuation for text in shape(s)."));
                                // 否则静默
                            }
                        );
                    }
                ); // ← Excel 分支回调
            }); // ← GetSelectedText 回调结束
        });

        // B. 强制转半角
        this.attachToolbarMenuClickEvent("banjiao", function () {
            const plugin = window.Asc.plugin;

            // 供命令体读取的配置
            Asc.scope.__punct__ = {
                map: {
                    "，": ",",
                    "；": ";",
                    "：": ":",
                    "。": ".",
                    "“": '"',
                    "”": '"',
                    "‘": "'",
                    "’": "'",
                    // 破折号：中文 EM DASH 与全角连字符都收敛成半角连字符
                    "—": "-",
                    "－": "-",
                    // 货币
                    "＄": "$",
                    "￥": "¥",
                    "￡": "£",
                    "￠": "¢",
                    "《": "<",
                    "》": ">",
                    "（": "(",
                    "）": ")",
                    "？": "?",
                    "！": "!",
                    "／": "/",
                },
                // 开关：为空表示全部处理；否则只处理设置中选中的“全角目标字符”
                // settings: { punctuation: JSON.parse(localStorage.getItem("selectedPunctuation") || "[]") }
                settings: { punctuation: readJSON("selectedPunctuation", []) },
            };

            // —— 选中文本：优先用 ReplaceTextSmart（保留样式）
            const props = {
                Numbering: false,
                Math: false,
                TableCellSeparator: "\t",
                TableRowSeparator: "\n",
                ParaSeparator: "\n",
                TabSymbol: "\t",
                NewLineSeparator: "\r",
            };

            plugin.executeMethod("GetSelectedText", [props], function (t) {
                // const picked = (t || "").replace(/\r\n?/g, "\n");
                const picked = t || "";

                // 公用转换（编辑器侧，非命令体）
                const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
                const convertLine = (line) => {
                    if (!line) return line;
                    const map = Asc.scope.__punct__.map;
                    const on = Asc.scope.__punct__.settings.punctuation || [];
                    // ✅ 省略号归一到半角：任何“……/…/连续三个及以上点” → "..."
                    let v = line.replace(/(?:…+|\.{3,})/g, "……");
                    for (const [full, half] of Object.entries(map)) {
                        if (on.length === 0 || on.includes(full))
                            v = v.replace(new RegExp(esc(full), "g"), half);
                    }
                    return v;
                };

                if (picked.trim()) {
                    const out = picked.split(/\t|\n/).map(convertLine);
                    if(out[out.length - 1] === "") out.pop();
                    Asc.scope._lines = out.map(l => l.replace(/\r/g, "\r\n"));
                    plugin.callCommand(
                        function () {
                            if (
                                Asc.scope._lines &&
                                typeof Api.ReplaceTextSmart === "function"
                            ) {
                                Api.ReplaceTextSmart(Asc.scope._lines); // 保留原样式地替换
                            }
                        },
                        false,
                        true,
                        function () {
                            getInfoModal(tr("Converted selection: ") + out.length + tr(" line(s)."));
                        }
                    );
                    return; // 已处理，退出
                }

                // =============== Excel 分支（命令体内重建转换函数！） ===============
                plugin.callCommand(
                    function () {
                        function escIn(x) {
                            return x.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
                        }
                        function convertIn(line) {
                            if (!line) return line;
                            var m = Asc.scope.__punct__.map;
                            var on = Asc.scope.__punct__.settings.punctuation || [];
                            var v = line.replace(/(?:…+|\.{3,})/g, "……"); // 半角省略号
                            for (var k in m) {
                                if (!m.hasOwnProperty(k)) continue;
                                var half = m[k];
                                // on 里存放“全角目标”，即 k
                                if (on.length === 0 || on.indexOf(k) !== -1)
                                    v = v.replace(new RegExp(escIn(k), "g"), half);
                            }
                            return v;
                        }

                        try {
                            if (typeof Api.GetActiveSheet === "function") {
                                var ws = Api.GetActiveSheet();
                                var rng = ws && ws.GetSelection && ws.GetSelection();
                                if (!rng && typeof Api.GetSelection === "function")
                                    rng = Api.GetSelection();
                                if (
                                    rng &&
                                    typeof rng.GetValue === "function" &&
                                    typeof rng.SetValue === "function"
                                ) {
                                    var val = rng.GetValue(); // 可能是 string 或 二维数组
                                    var changed = false,
                                        out;

                                    function convCell(v) {
                                        if (typeof v !== "string") return v;
                                        var nv = convertIn(v);
                                        if (nv !== v) changed = true;
                                        return nv;
                                    }

                                    if (Array.isArray(val)) {
                                        out = [];
                                        for (var r = 0; r < val.length; r++) {
                                            var row = val[r];
                                            if (Array.isArray(row)) {
                                                var nr = new Array(row.length);
                                                for (var c = 0; c < row.length; c++)
                                                    nr[c] = convCell(row[c]);
                                                out.push(nr);
                                            } else out.push(convCell(row));
                                        }
                                    } else out = convCell(val);

                                    if (changed) {
                                        rng.SetValue(out);
                                        return true;
                                    } // 写回
                                    return false;
                                }
                            }
                        } catch (e) { }
                    },
                    false,
                    true,
                    function (returnValue) {
                        if (returnValue) {
                            getInfoModal(tr("Converted punctuation in selected cells."));
                            return; // Excel 成功，结束
                        }

                        // =============== PPT 分支（命令体内重建转换函数！） ===============
                        plugin.callCommand(
                            function () {
                                function escIn(x) {
                                    return x.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
                                }
                                function convertIn(line) {
                                    if (!line) return line;
                                    var m = Asc.scope.__punct__.map;
                                    var on = Asc.scope.__punct__.settings.punctuation || [];
                                    var v = line.replace(/(?:…+|\.{3,})/g, "……");
                                    for (var k in m) {
                                        if (!m.hasOwnProperty(k)) continue;
                                        var half = m[k];
                                        if (on.length === 0 || on.indexOf(k) !== -1)
                                            v = v.replace(new RegExp(escIn(k), "g"), half);
                                    }
                                    return v;
                                }

                                var sel =
                                    typeof Api.GetSelection === "function"
                                        ? Api.GetSelection()
                                        : null; // 选区（演示）
                                var shapes =
                                    sel && typeof sel.GetShapes === "function"
                                        ? sel.GetShapes()
                                        : null; // 被选中的图形
                                if (!Array.isArray(shapes)) shapes = shapes ? [shapes] : [];

                                if (shapes.length === 0) {
                                    // 兜底：取当前页所有对象
                                    var pres =
                                        typeof Api.GetPresentation === "function"
                                            ? Api.GetPresentation()
                                            : null;
                                    var slide =
                                        pres && typeof pres.GetCurrentSlide === "function"
                                            ? pres.GetCurrentSlide()
                                            : null;
                                    if (slide && typeof slide.GetAllObjects === "function") {
                                        var all = slide.GetAllObjects();
                                        if (Array.isArray(all)) {
                                            var chosen = [];
                                            for (var i = 0; i < all.length; i++) {
                                                var o = all[i];
                                                try {
                                                    if (
                                                        o &&
                                                        typeof o.IsSelected === "function" &&
                                                        o.IsSelected()
                                                    )
                                                        chosen.push(o);
                                                } catch (e) { }
                                            }
                                            shapes = chosen.length ? chosen : all;
                                        }
                                    }
                                }

                                var hit = 0;
                                function getContent(draw) {
                                    try {
                                        if (draw && typeof draw.GetContent === "function")
                                            return draw.GetContent(); // 新版
                                        if (draw && typeof draw.GetDocContent === "function")
                                            return draw.GetDocContent(); // 旧版
                                    } catch (e) { }
                                    return null;
                                }

                                for (var sIdx = 0; sIdx < shapes.length; sIdx++) {
                                    var dc = getContent(shapes[sIdx]); // ApiDocumentContent
                                    if (!dc || typeof dc.GetAllParagraphs !== "function")
                                        continue;
                                    var paras = dc.GetAllParagraphs(); // 段落数组
                                    if (!paras || !paras.length) continue;

                                    var changed = false;
                                    for (var pIdx = 0; pIdx < paras.length; pIdx++) {
                                        var p = paras[pIdx];
                                        var old =
                                            p && typeof p.GetText === "function"
                                                ? p.GetText({
                                                    Numbering: false,
                                                    Math: false,
                                                    NewLineSeparator: "\n",
                                                    TabSymbol: "\t",
                                                })
                                                : "";
                                        if (!old) continue;

                                        // 命中全角/省略号才处理
                                        if (
                                            /[，。；：‘’“”《》（）？！／—－…]/.test(old) ||
                                            /…|\.{3,}/.test(old)
                                        ) {
                                            var neo = convertIn(old);
                                            if (
                                                neo !== old &&
                                                typeof p.Select === "function" &&
                                                typeof Api.ReplaceTextSmart === "function"
                                            ) {
                                                p.Select(); // 选中段落
                                                Api.ReplaceTextSmart([neo], "\t", "\n"); // 保留样式地替换
                                                changed = true;
                                            }
                                        }
                                    }
                                    if (changed) hit++;
                                }
                                return hit > 0;
                            },
                            false,
                            true,
                            function (returnValue) {
                                if (returnValue)
                                    getInfoModal(tr("Converted punctuation for text in shape(s)."));
                                // Otherwise, remain silent.
                            }
                        );
                    }
                ); // ← Excel 分支回调
            }); // ← GetSelectedText 回调结束
        });

        // C. 智能转换 → 空格策略 → 进入报告
        this.attachToolbarMenuClickEvent("zhineng", function () {
            // have to explore this ---------------------------------------
            const plugin = window.Asc.plugin;
            const resolveUrl = window.resolveUrl || ((p) => p);

            // —— 读取选区/内容（Word/Excel/PPT 通用）——
            const props = {
                Numbering: false,
                Math: false,
                TableCellSeparator: "\n",
                TableRowSeparator: "\n",
                ParaSeparator: "\n",
                TabSymbol: "\t",
                NewLineSeparator: "\r",
            };

            // —— 首选：直接取“选中文本”（Word/可选中对象的场景）
            plugin.executeMethod("GetSelectedText", [props], function (s) {
                if (s && s.trim()) {
                    localStorage.setItem("_smart_source_type", "word");
                    // openPanel(s.replace(/\r\n?/g, "\n"));
                    openPanel(s);
                    return;
                }

                // 兜底：获取“选中内容的纯文本”
                plugin.executeMethod(
                    "GetSelectedContent",
                    [{ type: "text" }],
                    function (s2) {
                        if (s2 && s2.trim()) {
                            localStorage.setItem("_smart_source_type", "word");
                            openPanel(s2.replace(/\r\n?/g, "\n"));
                            return;
                        }

                        // —— 进入命令体：尝试 Excel / PPT ——
                        plugin.callCommand(
                            function () {
                                var resultText = "";
                                var sourceType = "";

                                // Excel：按选区取值（单元格 or 二维数组），序列化为行列文本
                                try {
                                    if (typeof Api.GetActiveSheet === "function") {
                                        var ws = Api.GetActiveSheet();
                                        var rng = ws && ws.GetSelection && ws.GetSelection();
                                        if (!rng && typeof Api.GetSelection === "function")
                                            rng = Api.GetSelection();

                                        if (rng && typeof rng.GetValue === "function") {
                                            var val = rng.GetValue();
                                            if (val) {
                                                sourceType = "excel";
                                                if (typeof val === "string") {
                                                    resultText = val;
                                                } else if (Array.isArray(val)) {
                                                    var lines = [];
                                                    for (var r = 0; r < val.length; r++) {
                                                        lines.push(val[r].join("\t"));
                                                    }
                                                    resultText = lines.join("\n");
                                                }
                                            }
                                        }
                                    }
                                } catch (e) {
                                    console.error(">>> Excel 异常:", e);
                                }

                                // PPT：收集被选中/当前页对象 → 逐形状聚合段落文本（形状间用空行分隔）
                                if (!resultText) {
                                    try {
                                        function getContent(draw) {
                                            try {
                                                if (draw && typeof draw.GetContent === "function")
                                                    return draw.GetContent();
                                                if (draw && typeof draw.GetDocContent === "function")
                                                    return draw.GetDocContent();
                                            } catch (e) { }
                                            return null;
                                        }

                                        var sel =
                                            typeof Api.GetSelection === "function"
                                                ? Api.GetSelection()
                                                : null;
                                        var shapes =
                                            sel && typeof sel.GetShapes === "function"
                                                ? sel.GetShapes()
                                                : null;
                                        if (!Array.isArray(shapes)) shapes = shapes ? [shapes] : [];

                                        if (shapes.length === 0) {
                                            var pres =
                                                typeof Api.GetPresentation === "function"
                                                    ? Api.GetPresentation()
                                                    : null;
                                            var slide =
                                                pres && typeof pres.GetCurrentSlide === "function"
                                                    ? pres.GetCurrentSlide()
                                                    : null;
                                            if (slide && typeof slide.GetAllObjects === "function") {
                                                var all = slide.GetAllObjects();
                                                if (Array.isArray(all)) {
                                                    var chosen = [];
                                                    for (var i = 0; i < all.length; i++) {
                                                        try {
                                                            if (
                                                                all[i] &&
                                                                typeof all[i].IsSelected === "function" &&
                                                                all[i].IsSelected()
                                                            )
                                                                chosen.push(all[i]);
                                                        } catch (e) { }
                                                    }
                                                    shapes = chosen.length ? chosen : all;
                                                }
                                            }
                                        }

                                        // 保存形状信息与“每个形状的段落数”（供其它模式使用；本方案只用 shapeIndices）
                                        var shapeTexts = [];
                                        var shapeIndices = [];
                                        var paraCounts = [];

                                        for (var sIdx = 0; sIdx < shapes.length; sIdx++) {
                                            var dc = getContent(shapes[sIdx]);
                                            if (!dc || typeof dc.GetAllParagraphs !== "function")
                                                continue;
                                            var paras = dc.GetAllParagraphs();
                                            if (!paras || !paras.length) continue;

                                            var shapeText = "";
                                            for (var pIdx = 0; pIdx < paras.length; pIdx++) {
                                                var p = paras[pIdx];
                                                var txt =
                                                    p && typeof p.GetText === "function"
                                                        ? p.GetText({
                                                            Numbering: false,
                                                            Math: false,
                                                            NewLineSeparator: "\n",
                                                            TabSymbol: "\t",
                                                        })
                                                        : "";
                                                // 统一去掉段尾的各种换行符：\r \n U+2028/U+2029 U+0085 垂直制表 \x0B
                                                txt = String(txt).replace(
                                                    /[\r\t\n\u2028\u2029\u0085\x0B]+$/g,
                                                    ""
                                                );

                                                if (txt) {
                                                    if (shapeText) shapeText += "\n";
                                                    shapeText += txt;
                                                }
                                            }

                                            if (shapeText.trim()) {
                                                shapeTexts.push(shapeText);
                                                shapeIndices.push(sIdx);
                                                paraCounts.push(paras.length);
                                            }
                                        }

                                        if (shapeTexts.length > 0) {
                                            sourceType = "ppt";
                                            resultText = shapeTexts.join("\n\n"); // 形状间空行分隔（用于回填时分块）
                                            Asc.scope._pptShapeCount = shapeTexts.length;
                                            Asc.scope._pptShapeIndices = shapeIndices;
                                            Asc.scope._pptParaCounts = paraCounts;
                                        }
                                    } catch (e) {
                                        console.error(">>> PPT 异常:", e);
                                    }
                                }

                                // 写入 localStorage，让外层读取
                                try {
                                    if (resultText) {
                                        localStorage.setItem("_smart_temp_text", resultText);
                                        localStorage.setItem("_smart_ready", "1");
                                        localStorage.setItem("_smart_source_type", sourceType);

                                        if (sourceType === "ppt") {
                                            localStorage.setItem(
                                                "_smart_ppt_shape_count",
                                                Asc.scope._pptShapeCount.toString()
                                            );
                                            localStorage.setItem(
                                                "_smart_ppt_shape_indices",
                                                JSON.stringify(Asc.scope._pptShapeIndices)
                                            );
                                            localStorage.setItem(
                                                "_smart_ppt_para_counts",
                                                JSON.stringify(Asc.scope._pptParaCounts)
                                            );
                                        }
                                    } else {
                                        localStorage.setItem("_smart_ready", "0");
                                    }
                                } catch (e) {
                                    console.error(">>> localStorage 写入失败:", e);
                                    localStorage.setItem("_smart_ready", "0");
                                }
                            },
                            false,
                            true,
                            function () {
                                var ready = localStorage.getItem("_smart_ready");

                                if (ready === "1") {
                                    var text = localStorage.getItem("_smart_temp_text");

                                    localStorage.removeItem("_smart_temp_text");
                                    localStorage.removeItem("_smart_ready");

                                    if (text && text.trim()) {
                                        openPanel(text);
                                    } else {
                                        plugin.executeMethod("ShowError", [
                                            tr("Please select the text to diagnose!"),
                                        ]);
                                    }
                                } else {
                                    plugin.executeMethod("ShowError", [
                                        tr("Please select the text to diagnose!"),
                                    ]);
                                }
                            }
                        );
                    }
                );
            });

            // —— 打开空格策略选项面板 ——
            function openPanel(text) {
                localStorage.setItem("smart_source_text", text);
                localStorage.removeItem("spaceOptReady");

                if (winOptions) {
                    try {
                        winOptions.close();
                    } catch (e) { }
                    winOptions = null;
                }

                winOptions = new window.Asc.PluginWindow();
                winOptions.show({
                    url: resolveUrl("panels/space-options.html"),
                    description: tr("Spacing options"),
                    isModal: true,
                    isVisual: true,
                    size: [560, 340],
                    EditorsSupport: ["word", "cell", "slide"],
                    buttons: [
                        { text: tr("Confirm"), primary: true },
                        { text: tr("Cancel"), primary: false },
                    ],
                });

                // 你的既有逻辑
                watchSpaceOptions(text);

                // ★ 如果这次来源是 PPT：启动轮询，等 report.html 把处理结果放进 localStorage
                if (localStorage.getItem("_smart_source_type") === "ppt") {
                    startPptApplyWatcher();
                }
            }

            // ========== 以下为 PPT 回填的轻量实现：localStorage + 轮询 + 一次性写回 ==========

            // 轮询：等 report.html 把处理结果塞进 localStorage（_smart_apply_lines）
            function startPptApplyWatcher() {
                // 先清掉旧的
                if (window.__pptApplyTimer) {
                    try {
                        clearInterval(window.__pptApplyTimer);
                    } catch (e) { }
                }
                localStorage.removeItem("_smart_apply_lines");

                window.__pptApplyTimer = setInterval(() => {
                    const raw = localStorage.getItem("_smart_apply_lines");
                    if (!raw) return; // 还没准备好

                    // 拿到了就消费掉 & 停表
                    localStorage.removeItem("_smart_apply_lines");
                    try {
                        clearInterval(window.__pptApplyTimer);
                    } catch (e) { }
                    window.__pptApplyTimer = null;

                    const blocks = parseBlocks(raw);

                    // 真正回填
                    applyPptBlocks(blocks);
                }, 120);
            }

            // 把行或字符串折叠成“形状块”：
            // - 传入 JSON.stringify 的数组（3块）→ 直接返回
            // - 传入 JSON.stringify 的数组（多行+空行）→ 按空行折叠
            // - 传入字符串 → 按换行拆，再按空行折叠
            function parseBlocks(raw) {
                try {
                    const v = JSON.parse(raw);
                    if (Array.isArray(v)) {
                        if (v.length && v.every((s) => typeof s === "string" && s !== ""))
                            return v.slice(); // 已是块
                        return foldByBlank(v); // 行数组（含空行）→ 块
                    }
                    if (typeof v === "string")
                        return foldByBlank(v.replace(/\r/g, "").split("\n"));
                } catch (e) {
                    // raw 不是 JSON，就当纯文本
                    return foldByBlank(String(raw).replace(/\r/g, "").split("\n"));
                }
                return [];
            }

            function foldByBlank(lines) {
                const out = [];
                let buf = [];
                for (const ln of lines) {
                    if (ln === "") {
                        if (buf.length) {
                            out.push(buf.join("\n"));
                            buf = [];
                        }
                    } else buf.push(ln);
                }
                if (buf.length) out.push(buf.join("\n"));
                return out.map((b) => b.replace(/\n+$/, "").trim());
            }

            // 在命令体中：每块文本 ↔ 一个形状（首段容器），一次性 ReplaceTextSmart
            // 在命令体中：每块文本 ↔ 一个形状（首段容器），一次性 ReplaceTextSmart
            function applyPptBlocks(blocks) {
                const plugin = window.Asc.plugin;

                // ✅ 把外层数据放到 Asc.scope，供命令体沙箱读取
                try {
                    Asc.scope._pptBlocks = Array.isArray(blocks) ? blocks.slice() : [];
                } catch (e) {
                    Asc.scope._pptBlocks = [];
                }

                const nBlocks = Asc.scope._pptBlocks.length; // 仅用于外层日志

                plugin.callCommand(
                    function () {
                        function getContent(draw) {
                            try {
                                if (draw && typeof draw.GetContent === "function")
                                    return draw.GetContent();
                                if (draw && typeof draw.GetDocContent === "function")
                                    return draw.GetDocContent();
                            } catch (e) { }
                            return null;
                        }

                        // ⬇️ 在沙箱里用 Asc.scope 取回 blocks
                        var blocks = Asc.scope._pptBlocks || [];

                        // 取本页形状（优先已选，否则整页）
                        var sel =
                            typeof Api.GetSelection === "function"
                                ? Api.GetSelection()
                                : null;
                        var shapes =
                            sel && typeof sel.GetShapes === "function"
                                ? sel.GetShapes()
                                : null;
                        if (!Array.isArray(shapes)) shapes = shapes ? [shapes] : [];
                        if (shapes.length === 0) {
                            var pres =
                                typeof Api.GetPresentation === "function"
                                    ? Api.GetPresentation()
                                    : null;
                            var slide =
                                pres && typeof pres.GetCurrentSlide === "function"
                                    ? pres.GetCurrentSlide()
                                    : null;
                            if (slide && typeof slide.GetAllObjects === "function") {
                                var all = slide.GetAllObjects();
                                if (Array.isArray(all)) {
                                    var chosen = [];
                                    for (var i = 0; i < all.length; i++) {
                                        try {
                                            if (all[i]?.IsSelected?.()) chosen.push(all[i]);
                                        } catch (e) { }
                                    }
                                    shapes = chosen.length ? chosen : all;
                                }
                            }
                        }

                        // 取提取时记录的顺序；没有就按当前顺序
                        var idxs = [];
                        try {
                            idxs = JSON.parse(
                                localStorage.getItem("_smart_ppt_shape_indices") || "[]"
                            );
                        } catch (e) { }
                        if (!Array.isArray(idxs) || !idxs.length)
                            idxs = shapes.map(function (_, i) {
                                return i;
                            });

                        var applied = 0,
                            n = Math.min(blocks.length, idxs.length);
                        for (var k = 0; k < n; k++) {
                            var sIndex = idxs[k];
                            var dc = getContent(shapes[sIndex]);
                            if (!dc || typeof dc.GetAllParagraphs !== "function") continue;

                            var paras = dc.GetAllParagraphs() || [];
                            if (!paras.length) continue;

                            var p0 = paras[0];
                            var block = String(blocks[k] || "");
                            if (p0?.Select && typeof Api.ReplaceTextSmart === "function") {
                                p0.Select();
                                Api.ReplaceTextSmart([block], "\t", "\n"); // 块内 \n 会自动分段
                                applied++;
                            }
                        }

                        Asc.scope._pptAppliedShapes = applied;
                    },
                    false,
                    true,
                    function () {
                        try {
                            getInfoModal("Applied to " + (Asc.scope._pptAppliedShapes || 0) + " shape(s).");
                        } catch (e) { }
                        try {
                            if (typeof winReport !== "undefined" && winReport) {
                                winReport.close();
                                winReport = null;
                            }
                        } catch (e) { }
                        // 清缓存（避免下次错位）
                        localStorage.removeItem("_smart_ppt_shape_count");
                        localStorage.removeItem("_smart_ppt_shape_indices");
                        localStorage.removeItem("_smart_ppt_para_counts");
                        // 清掉 blocks
                        Asc.scope._pptBlocks = null;
                    }
                );
            }
        });

        // D. 设置
        this.attachToolbarMenuClickEvent("setting", function () {
            if (winSetting) {
                try {
                    winSetting.close();
                } catch (e) { }
                winSetting = null;
            }
            winSetting = new window.Asc.PluginWindow();
            winSetting.show({
                url: resolveUrl("panels/setting.html"),
                description: tr("Settings"),
                isModal: true,
                isVisual: true,
                size: [360, 360],
                buttons: [{ text: tr("Save"), primary: false }],
                EditorsSupport: ["word", "slide", "cell"],
            });
        });
    }

    // ---------------- Unify window close callback ----------------
    // Helper function to close window if it matches
    function closeWindowIfMatch(win, windowId) {
        if (win && windowId === win.id) {
            try {
                win.close();
            } catch (e) { }
            return true;
        }
        return false;
    }

    window.Asc.plugin.button = function (id, windowId) {
        if (closeWindowIfMatch(winInfo, windowId)) {
            localStorage.removeItem("info_message");
            winInfo = null;
            return;
        }
        if (closeWindowIfMatch(winSetting, windowId)) {
            winSetting = null;
            return;
        }

        if (winOptions && windowId === winOptions.id) {
            // ⬇️ Added: id === 0 represents Confirm; id === 1 represents Cancel.
            if (id === 0) {
                localStorage.setItem("spaceOptReady", "1"); // 交给后续逻辑判断/推进
            } else {
                localStorage.removeItem("spaceOptReady");
            }

            try {
                winOptions.close();
            } catch (e) { }
            winOptions = null;

            const ok = localStorage.getItem("spaceOptReady") === "1";
            const selectedText = localStorage.getItem("smart_source_text") || "";
            localStorage.removeItem("spaceOptReady");

            if (ok && selectedText.trim()) {
                proceedToReport(selectedText);
            }
            return;
        }

        if (winReport && windowId === winReport.id) {
            if (id === 0) {
                localStorage.setItem("reportApplyNow", "1"); // Trigger child page execution

                // ✅ 新增：兜底轮询等待子页产出结果（避免消息通道被拦截时卡住）
                if (reportApplyTimer) {
                    clearInterval(reportApplyTimer);
                    reportApplyTimer = null;
                }
                reportApplyTimer = setInterval(() => {
                    // 子窗已经被关就停止
                    if (!winReport) {
                        clearInterval(reportApplyTimer);
                        reportApplyTimer = null;
                        return;
                    }

                    const ready = localStorage.getItem("reportApplyReady") === "1";
                    if (!ready) return;

                    clearInterval(reportApplyTimer);
                    reportApplyTimer = null;
                    localStorage.removeItem("reportApplyReady");

                    const replaceResult = localStorage.getItem("batchReplaceResult");
                    if (!replaceResult) return;

                    const isPPT =
                        window.Asc.plugin.info &&
                        window.Asc.plugin.info.editorType === "slide";
                    if (isPPT) {
                        // —— PPT：桥接给按形状写回 —— //
                        localStorage.setItem("_smart_apply_lines", replaceResult); // 触发 startPptApplyWatcher → applyPptBlocks
                        try {
                            if (winReport) {
                                winReport.close();
                                winReport = null;
                            }
                        } catch (e) { }
                        // 防御式：可以重复调用，不会有问题
                        try {
                            startPptApplyWatcher();
                        } catch (e) { }
                        return;
                    }

                    // —— Word/Excel: Keep the original logic. Smart converter—— //
                    try {
                        Asc.scope.convertedLines = JSON.parse(replaceResult).map(l => l.replace(/\r/g, "\r\n"));;
                    } catch (e) {
                        Asc.scope.convertedLines = null;
                    }
                    window.Asc.plugin.callCommand(
                        function () {
                            if (Asc.scope.convertedLines)
                                Api.ReplaceTextSmart(Asc.scope.convertedLines);
                        },
                        false,
                        true,
                        function () {
                            try {
                                localStorage.setItem("currentSelectionLines", replaceResult);
                            } catch (e) { }
                            if (winReport) {
                                try {
                                    winReport.close();
                                } catch (e) { }
                                winReport = null;
                            }
                        },
                        REPORT_CLOSE_MS
                    );
                }, 1000);

                return; // 交给上面的轮询/消息回调来收尾
            }

            try {
                winReport.close();
            } catch (e) { }
            winReport = null;

            const closeReason = localStorage.getItem("reportCloseReason") || "open";
            localStorage.removeItem("reportCloseReason");

            const replaceResult = localStorage.getItem("batchReplaceResult");

            // New feature: If the user closes the dialog by clicking × without clicking “Confirm and Save”, an error message is displayed.
            if (closeReason === "closed_without_save") {
                localStorage.removeItem("batchReplaceResult"); // 清理潜在脏数据
                if (typeof showNotice === "function") {
                    // The second parameter true indicates the ‘error/danger’ style (as per your project’s convention).
                    showNotice("你没有保存任何需要修改的内容", true);
                } else {
                    alert("你没有保存任何需要修改的内容");
                }
                return;
            }

            // 正常保存关闭：应用替换（安全兜底：缺结果就静默返回）
            if (!replaceResult) return;

            const isPPT2 =
                window.Asc.plugin.info && window.Asc.plugin.info.editorType === "slide";
            if (isPPT2) {
                // —— PPT：同样桥接到形状回填 —— //
                localStorage.setItem("_smart_apply_lines", replaceResult);
                localStorage.removeItem("batchReplaceResult");
                try {
                    startPptApplyWatcher();
                } catch (e) { }
                return;
            }

            // —— Word/Excel：保持原逻辑 —— //
            Asc.scope.convertedLines = JSON.parse(replaceResult).map(l => l.replace(/\r/g, "\r\n"));;
            localStorage.removeItem("batchReplaceResult");
            window.Asc.plugin.callCommand(function () {
                if (Asc.scope.convertedLines)
                    Api.ReplaceTextSmart(Asc.scope.convertedLines);
            });
            return;
        }
    };

    // ---------------- Toolbar Definition ----------------
    function getToolbarItems() {
        return {
            guid: window.Asc.plugin.guid,
            tabs: [
                {
                    id: "tab_2",
                    text: window.Asc.plugin.tr("Chinese Auto-format Tool"),
                    items: [
                        {
                            id: "zhineng",
                            type: "button",
                            text: tr("Smart Convert"),
                            hint: tr("Validate Chinese typography automatically"),
                            icons: "resources/buttons/icon_zhineng.png",
                            lockInViewMode: true,
                        },
                        {
                            id: "quanjiao",
                            type: "button",
                            text: tr("Force Full-width"),
                            hint: tr("Force all symbols in selection to full-width"),
                            icons: "resources/buttons/icon_quanjiao.png",
                            lockInViewMode: true,
                            separator: true,
                        },
                        {
                            id: "banjiao",
                            type: "button",
                            text: tr("Force Half-width"),
                            hint: tr("Force all symbols in selection to half-width"),
                            icons: "resources/buttons/icon_banjiao.png",
                            lockInViewMode: true,
                        },
                        {
                            id: "setting",
                            type: "button",
                            text: tr("Settings"),
                            hint: tr("Configure conversion rules"),
                            icons: "resources/buttons/icon_setting.png",
                            lockInViewMode: true,
                        },
                    ],
                },
            ],
        };
    }

    function getInfoModal(message){
        localStorage.setItem("info_message", message || "");
        winInfo = new window.Asc.PluginWindow();
        winInfo.show({
            url: resolveUrl("panels/info.html"),
            description: tr("Info"),
            isModal: true,
            isVisual: true,
            size: [400, 100],
            EditorsSupport: ["word", "cell", "slide"],
            buttons: [
                { text: tr("OK"), primary: true }
            ],
        });
    }
})(window);