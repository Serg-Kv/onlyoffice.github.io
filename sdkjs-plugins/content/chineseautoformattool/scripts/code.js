// scripts/code.js
(function (window, undefined) {
  // ---------------- 初始化 ----------------
  window.Asc.plugin.init = function () { 
	bindToolbarEvents.call(this);
  };

  // 子窗口句柄
  let __toolbarAdded = false;
  let winSetting = null;
  let winOptions = null;
  let winReport = null;
  let winInfo = null;
  let selectedTextToFormat = "";
  let smart_ppt_para_counts = null;

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
          } catch (e) {}
          return null;
        }

        var sel =
          typeof Api.GetSelection === "function" ? Api.GetSelection() : null;
        var shapes =
          sel && typeof sel.GetShapes === "function" ? sel.GetShapes() : null;
        if (!Array.isArray(shapes)) shapes = shapes ? [shapes] : [];
        if (shapes.length === 0) {
          var pres = typeof Api.GetPresentation === "function" ? Api.GetPresentation() : null;
          var slide = pres && typeof pres.GetCurrentSlide === "function" ? pres.GetCurrentSlide() : null;
          var all = slide && typeof slide.GetAllObjects === "function" ? slide.GetAllObjects() : null;
          if (Array.isArray(all)) shapes = all;
        }
        var stats = [];
        for (var i = 0; i < shapes.length; i++) {
          var dc = getContent(shapes[i]);
          var n = dc?.GetAllParagraphs?.().length || 0;
          stats.push({ shape: i, paras: n });
        }
        Asc.scope.__stats = stats;
      }, false, true, function () {
        getInfoModal(
          "PPT Probe: shapes=" + ((Asc.scope.__stats || []).length || 0),
        );
      },
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
          } catch (e) {}
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
          Api.ReplaceTextSmart(["[SMART-TEST]\n(能看到这一行说明 PPT 回填链路 OK)"], "\t", "\n");
          Asc.scope.__ok = true;
        } else Asc.scope.__ok = false;
      }, false, true, function () {
        getInfoModal(Asc.scope.__ok ? "Dry apply OK" : "Dry apply failed");
      },
    );
  };

  // 放在 (function (window, undefined) { 之后的任意顶层位置
  const tr = (s) => window.Asc && window.Asc.plugin ? window.Asc.plugin.tr(s) : s;
  // —— 本地化回调：词典就绪后，刷新工具栏文本与提示 ——
  window.Asc.plugin.onTranslate = function () {
    getInfoModal(
      tr("The plugin is ready, the toolbar menu has been updated. Please go to the Chinese-Auto Format Tool tab above to use the formatting features.")
    );
    // ……你原来的 setText / 提示等本地化代码（如果有）……
    const items = getToolbarItems(); // 这里的 tabs[0].text 要用 tr("Chinese Formatter")
    if (!__toolbarAdded) {
      // ✅ First time: Create the tab with current language, tab title will use the translated text
      window.Asc.plugin.executeMethod("AddToolbarMenuItem", [items]);
      __toolbarAdded = true;
    } else {
      // ✅ Subsequent language switches: Only update button text/tooltips
      window.Asc.plugin.executeMethod("UpdateToolbarMenuItem", [items]);
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
  function proceedToReport() {
    const editorType = window.Asc.plugin.info.editorType || "word";
    let results;
    try {
      // Depends on runFormatCheck(text) in scripts/formatChecker.js
      results = runFormatCheck(selectedTextToFormat, editorType);
    } catch (e) {
      window.Asc.plugin.executeMethod("ShowError", [
        tr("Detection failed: formatChecker.js is missing or has an error"),
      ]);
      return;
    }

    const lines = editorType === "cell"
        ? selectedTextToFormat.split(/\t|\n/)
        : selectedTextToFormat.split(/\n/); // Do not filter empty lines; preserve line count consistency
    const fixed = results.map((r) => r.fixed);
    const report = results.filter((r) => r.errors && r.errors.length > 0);

    if (lines[lines.length - 1] === "") {
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

	closeWindowIfMatch(winReport);
    winReport = new window.Asc.PluginWindow();
    winReport.attachEvent("onWindowReportReady", function () {
      winReport.command("onReportPageData", {
        zhlintReport: report,
        originalLines: lines,
        fixedLines: fixed,
        type: "report-data",
        base: lines, // 供 PPT 用的基线数据（按行分割的字符串数组），report.html 里不直接用，传给 applyPptBlocks 处理成 shape-based 块后再用
      });
    });
    winReport.attachEvent("onWindowReportMessage", batchReplaceResultHandler);
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

  function batchReplaceResultHandler(data) {
    if (!data) return;
    const isPPT = window.Asc.plugin.info && window.Asc.plugin.info.editorType === "slide";
    if (isPPT) {
      // 防御式：可以重复调用，不会有问题
      if (winReport) {
        closeWindowIfMatch(winReport);
        winReport = null;
      }
      try {
        startPptApplyWatcher(data);
      } catch (e) {}
      return;
    }
    try {
        Asc.scope.convertedLines = JSON.parse(data);
    } catch (e) {
      	Asc.scope.convertedLines = null;
    }
    window.Asc.plugin.callCommand(
      function () {
        if (Asc.scope.convertedLines)
          Api.ReplaceTextSmart(Asc.scope.convertedLines, "\t", "\r");
      }, false, true, function () {
        if (winReport) {
          closeWindowIfMatch(winReport);
          winReport = null;
        }
      },
    );
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
          if (out[out.length - 1] === "") out.pop();
          Asc.scope._lines = out;
          plugin.callCommand(
            function () {
              if (Asc.scope._lines && typeof Api.ReplaceTextSmart === "function") {
                Api.ReplaceTextSmart(Asc.scope._lines, "\t", "\r"); // 保留原样式地替换
              }
            }, false, true, function () {
              getInfoModal(
                tr("Converted selection: ") + out.length + tr(" line(s)."),
              );
            },
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
                if ( rng && typeof rng.GetValue === "function" && typeof rng.SetValue === "function" ) {
                  var val = rng.GetValue(); // string 或 二维数组 :contentReference[oaicite:4]{index=4}
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
                  } // 写回 :contentReference[oaicite:5]{index=5}
                  return false;
                }
              }
            } catch (e) {}
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

                var sel = typeof Api.GetSelection === "function"
					? Api.GetSelection()
                    : null; // 选区（演示）:contentReference[oaicite:6]{index=6}
                var shapes = sel && typeof sel.GetShapes === "function"
                    ? sel.GetShapes()
                    : null; // 被选中的图形
                if (!Array.isArray(shapes)) shapes = shapes ? [shapes] : [];

                if (shapes.length === 0) {
                  // 兜底：取当前页所有对象
                  var pres = typeof Api.GetPresentation === "function" ? Api.GetPresentation() : null;
                  var slide = pres && typeof pres.GetCurrentSlide === "function" ? pres.GetCurrentSlide() : null;
                  if (slide && typeof slide.GetAllObjects === "function") {
                    var all = slide.GetAllObjects();
                    if (Array.isArray(all)) {
                      var chosen = [];
                      for (var i = 0; i < all.length; i++) {
                        var o = all[i];
                        try {
                          	if (o && typeof o.IsSelected === "function" && o.IsSelected())
                            	chosen.push(o);
                        } catch (e) {}
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
                  } catch (e) {}
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
                    var old = p && typeof p.GetText === "function"
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
						if ( neo !== old && typeof p.Select === "function" &&
							typeof Api.ReplaceTextSmart === "function" ) {
							p.Select(); // 选中段落
							Api.ReplaceTextSmart([neo], "\t", "\n"); // 保留样式地替换 :contentReference[oaicite:10]{index=10}
							changed = true;
						}
                    }
                  }
                  if (changed) hit++;
                }
                return hit > 0;
              }, false, true, function (returnValue) {
                if (returnValue)
                  getInfoModal(
                    tr("Converted punctuation for text in shape(s)."),
                  );
                // 否则静默
              },
            );
          },
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
          if (out[out.length - 1] === "") out.pop();
          Asc.scope._lines = out;
          plugin.callCommand(
            function () {
              if ( Asc.scope._lines && typeof Api.ReplaceTextSmart === "function" ) {
                Api.ReplaceTextSmart(Asc.scope._lines, "\t", "\r"); // 保留原样式地替换
              }
            }, false, true, function () {
              getInfoModal(
                tr("Converted selection: ") + out.length + tr(" line(s)."),
              );
            },
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
                if ( rng && typeof rng.GetValue === "function" && typeof rng.SetValue === "function") {
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
            } catch (e) {}
          }, false, true, function (returnValue) {
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

                var sel = typeof Api.GetSelection === "function"
                    ? Api.GetSelection()
                    : null; // 选区（演示）
                var shapes = sel && typeof sel.GetShapes === "function"
                    ? sel.GetShapes()
                    : null; // 被选中的图形
                if (!Array.isArray(shapes)) shapes = shapes ? [shapes] : [];

                if (shapes.length === 0) {
                  // 兜底：取当前页所有对象
                  var pres = typeof Api.GetPresentation === "function" ? Api.GetPresentation() : null;
                  var slide = pres && typeof pres.GetCurrentSlide === "function" ? pres.GetCurrentSlide() : null;
                  if (slide && typeof slide.GetAllObjects === "function") {
                    var all = slide.GetAllObjects();
                    if (Array.isArray(all)) {
                      var chosen = [];
                      for (var i = 0; i < all.length; i++) {
                        var o = all[i];
                        try {
                          if ( o && typeof o.IsSelected === "function" && o.IsSelected())
                            chosen.push(o);
                        } catch (e) {}
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
                  } catch (e) {}
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
                    var old = p && typeof p.GetText === "function"
                        ? p.GetText({
                            Numbering: false,
                            Math: false,
                            NewLineSeparator: "\n",
                            TabSymbol: "\t",
                          })
                        : "";
                    if (!old) continue;

                    // 命中全角/省略号才处理
                    if ( /[，。；：‘’“”《》（）？！／—－…]/.test(old) || /…|\.{3,}/.test(old) ) {
                      var neo = convertIn(old);
                      if ( neo !== old && typeof p.Select === "function" && 
						typeof Api.ReplaceTextSmart === "function" ) {
                        p.Select(); // 选中段落
                        Api.ReplaceTextSmart([neo], "\t", "\n"); // 保留样式地替换
                        changed = true;
                      }
                    }
                  }
                  if (changed) hit++;
                }
                return hit > 0;
              }, false, true, function (returnValue) {
                if (returnValue)
                  getInfoModal(tr("Converted punctuation for text in shape(s)."));
                // Otherwise, remain silent.
              },
            );
          },
        ); // ← Excel 分支回调
      }); // ← GetSelectedText 回调结束
    });

    // C. 智能转换 → 空格策略 → 进入报告
    this.attachToolbarMenuClickEvent("zhineng", function () {
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
          openPanel(s);
          return;
        }

        // 兜底：获取“选中内容的纯文本”
        plugin.executeMethod( "GetSelectedContent", [{ type: "text" }], function (s2) {
            if (s2 && s2.trim()) {
              openPanel(s2.replace(/\r\n?/g, "\n"));
              return;
            }

            // —— 进入命令体：尝试 Excel / PPT ——
            plugin.callCommand(
              function () {
                var resultText = "";
                var sourceType = "";
                var shapeTexts = [];
                var shapeIndices = [];

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
                      } catch (e) {}
                      return null;
                    }

                    var sel = typeof Api.GetSelection === "function" ? Api.GetSelection() : null;
                    var shapes = sel && typeof sel.GetShapes === "function" ? sel.GetShapes() : null;
                    if (!Array.isArray(shapes)) shapes = shapes ? [shapes] : [];

                    if (shapes.length === 0) {
                      var pres = typeof Api.GetPresentation === "function" ? Api.GetPresentation() : null;
                      var slide = pres && typeof pres.GetCurrentSlide === "function"
                          ? pres.GetCurrentSlide()
                          : null;
                      if (slide && typeof slide.GetAllObjects === "function") {
                        var all = slide.GetAllObjects();
                        if (Array.isArray(all)) {
                          var chosen = [];
                          for (var i = 0; i < all.length; i++) {
                            try {
                              if (all[i] && typeof all[i].IsSelected === "function" && all[i].IsSelected())
                                chosen.push(all[i]);
                            } catch (e) {}
                          }
                          shapes = chosen.length ? chosen : all;
                        }
                      }
                    }

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
                        txt = String(txt).replace(/[\r\t\n\u2028\u2029\u0085\x0B]+$/g, "");

                        if (txt) {
                          if (shapeText) shapeText += "\n";
                          shapeText += txt;
                        }
                      }

                      if (shapeText.trim()) {
                        shapeTexts.push(shapeText);
                        shapeIndices.push(sIdx);
                      }
                    }

                    if (shapeTexts.length > 0) {
                      sourceType = "ppt";
                      resultText = shapeTexts.join("\n\n"); // 形状间空行分隔（用于回填时分块）
                    }
                  } catch (e) {
                    console.error(">>> PPT 异常:", e);
                  }
                }

                try {
                  if (resultText) {
                    if (sourceType === "ppt") {
                      smart_ppt_para_counts = shapeIndices;
                    }
                    return { ready: true, text: resultText };
                  } else {
                    return { ready: false };
                  }
                } catch (e) {
                  return { ready: false };
                }
              }, false, true, function (returnValue) {
                if (returnValue && returnValue.ready) {
                  const text = returnValue.text;
                  if (text && text.trim()) {
                    openPanel(text);
                  } else {
                    plugin.executeMethod("ShowError", [tr("Please select the text to diagnose!")]);
                  }
                } else {
                  plugin.executeMethod("ShowError", [tr("Please select the text to diagnose!")]);
                }
              },
            );
          },
        );
      });

      // —— 打开空格策略选项面板 ——
      function openPanel(text) {
        selectedTextToFormat = text;
		closeWindowIfMatch(winOptions);

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
      }

      // ========== Below is a lightweight implementation for PPT write-back: localStorage + polling + one-time write-back ==========

      // Polling: wait for report.html to store the processed result into localStorage (_smart_apply_lines)
      function startPptApplyWatcher(raw) {
        if (!raw) return;
        applyPptBlocks(parseBlocks(raw));
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

        plugin.callCommand(function () {
            function getContent(draw) {
              try {
                if (draw && typeof draw.GetContent === "function")
                  return draw.GetContent();
                if (draw && typeof draw.GetDocContent === "function")
                  return draw.GetDocContent();
              } catch (e) {}
              return null;
            }

            // ⬇️ 在沙箱里用 Asc.scope 取回 blocks
            var blocks = Asc.scope._pptBlocks || [];

            // 取本页形状（优先已选，否则整页）
            var sel = typeof Api.GetSelection === "function" ? Api.GetSelection() : null;
            var shapes = sel && typeof sel.GetShapes === "function" ? sel.GetShapes() : null;
            if (!Array.isArray(shapes)) shapes = shapes ? [shapes] : [];
            if (shapes.length === 0) {
              var pres = typeof Api.GetPresentation === "function" ? Api.GetPresentation() : null;
              var slide = pres && typeof pres.GetCurrentSlide === "function" ? pres.GetCurrentSlide() : null;
              if (slide && typeof slide.GetAllObjects === "function") {
                var all = slide.GetAllObjects();
                if (Array.isArray(all)) {
                  var chosen = [];
                  for (var i = 0; i < all.length; i++) {
                    try {
                      if (all[i]?.IsSelected?.()) chosen.push(all[i]);
                    } catch (e) {}
                  }
                  shapes = chosen.length ? chosen : all;
                }
              }
            }

            // 取提取时记录的顺序；没有就按当前顺序
            var idxs = [];
            try {
              idxs = smart_ppt_para_counts || [];
            } catch (e) {}
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
          }, false, true, function () {
            try {
              getInfoModal("Applied to " + (Asc.scope._pptAppliedShapes || 0) + " shape(s).");
            } catch (e) {}
            try {
              if (typeof winReport !== "undefined" && winReport) {
                winReport.close();
                winReport = null;
              }
            } catch (e) {}
            smart_ppt_para_counts = null;
            // 清掉 blocks
            Asc.scope._pptBlocks = null;
          },
        );
      }
    });

    // D. 设置
    this.attachToolbarMenuClickEvent("setting", function () {
      closeWindowIfMatch(winSetting);
      winSetting = new window.Asc.PluginWindow();
      winSetting.show({
        url: resolveUrl("panels/setting.html"),
        description: tr("Settings"),
        isModal: true,
        isVisual: true,
        size: [450, 380],
        buttons: [{ text: tr("Save"), primary: false }],
        EditorsSupport: ["word", "slide", "cell"],
      });
    });
  }

  // ---------------- Unify window close callback ----------------
  // Helper function to close window if it matches
  function closeWindowIfMatch(win) {
    try {
      win.close();
    } catch (e) {}
  }

  window.Asc.plugin.button = function (id, windowId) {
    if (winInfo && winInfo.id == windowId) {
      closeWindowIfMatch(winInfo);
      winInfo = null;
      return;
    }
    if (winSetting && winSetting.id == windowId) {
      closeWindowIfMatch(winSetting);
      winSetting = null;
      return;
    }

    if (winOptions && windowId === winOptions.id) {
      // ⬇️ Added: id === 0 represents Confirm; id === 1 represents Cancel
      if (id === 0) {
        if (selectedTextToFormat.trim()) {
          proceedToReport();
        }
      }
      closeWindowIfMatch(winOptions);
      winOptions = null;
    }

    if (winReport && windowId === winReport.id) {
      if (id === 0) {
        winReport.command("onReportWindowClosed");
      } else {
        closeWindowIfMatch(winReport);
        winReport = null;
      }
    }
  };

  // ---------------- Toolbar Definition ----------------
  function getToolbarItems() {
    return {
      guid: window.Asc.plugin.guid,
      tabs: [
        {
          id: "tab_2",
          text: window.Asc.plugin.tr("Chinese Formatter"),
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

  function getInfoModal(message) {
	closeWindowIfMatch(winInfo);
    winInfo = new window.Asc.PluginWindow();
    winInfo.attachEvent("onWindowReady", function () {
      winInfo.command("onWindowMessage", {
        message: message || "",
        type: "info",
      });
    });
    winInfo.show({
      url: resolveUrl("panels/info.html"),
      description: tr("Info"),
      isModal: true,
      isVisual: true,
      size: [400, 100],
      EditorsSupport: ["word", "cell", "slide"],
      buttons: [{ text: tr("OK"), primary: true }],
    });
  }
})(window);
