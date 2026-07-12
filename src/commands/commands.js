/* global console, Excel, Office */

// ExcelKit — 公式助手功能命令

// ---- 工具函数 ----

function isPureNumberExpr(expr) {
  return /^[-+]?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][-+]?\d+)?$/.test(expr.trim());
}

// 在括号嵌套/字符串场景下找匹配的右括号
function findMatchingRightParen(text, leftParenIndex) {
  var depth = 0, inString = false;
  for (var i = leftParenIndex + 1; i < text.length; i++) {
    var ch = text[i];
    if (inString) {
      if (ch === '"') {
        if (text[i + 1] === '"') i++; else inString = false;
      }
      continue;
    }
    if (ch === '"') { inString = true; continue; }
    if (ch === "(") { depth++; continue; }
    if (ch === ")") { if (depth === 0) return i; depth--; }
  }
  return -1;
}

// 找 ROUND 参数分隔符（兼容 , / ，/ ; / ；）
function findLastTopLevelArgSeparator(text) {
  var depth = 0, inString = false, lastSep = -1;
  for (var i = 0; i < text.length; i++) {
    var ch = text[i];
    if (inString) {
      if (ch === '"') { if (text[i + 1] === '"') i++; else inString = false; }
      continue;
    }
    if (ch === '"') { inString = true; continue; }
    if (ch === "(") { depth++; continue; }
    if (ch === ")") { if (depth > 0) depth--; continue; }
    if (depth === 0 && (ch === "," || ch === "，" || ch === ";" || ch === "；")) lastSep = i;
  }
  return lastSep;
}

// 解析 ROUND(expr, N) → { expr, digits }
function extractOuterRoundArgs(formula) {
  var body = String(formula).replace(/^=\s*/, "");
  var m = /^ROUND\b/i.exec(body);
  if (!m) return null;
  var i = m[0].length;
  while (i < body.length && /\s/.test(body[i])) i++;
  if (body[i] !== "(") return null;
  var endParen = findMatchingRightParen(body, i);
  if (endParen < 0) return null;
  var inside = body.slice(i + 1, endParen);
  var sepIndex = findLastTopLevelArgSeparator(inside);
  if (sepIndex < 0) return null;
  var exprPart = inside.slice(0, sepIndex).trim();
  var nPart = inside.slice(sepIndex + 1).trim();
  if (!exprPart || !/^[-+]?\d+$/.test(nPart)) return null;
  if (body.slice(endParen + 1).trim()) return null;
  return { expr: exprPart, digits: parseInt(nPart, 10) };
}

// removeRound 专用：纯数字结果是否需要放宽格式
function shouldRelaxFormatForConstant(beforeFormat, digits) {
  var fmt = (beforeFormat || "").trim();
  if (!fmt || /^general$/i.test(fmt) || !/^0(?:\.0+)?$/.test(fmt)) return false;
  var dot = fmt.indexOf(".");
  return (dot >= 0 ? fmt.length - dot - 1 : 0) === digits;
}

// 判断表达式是否被完整括号包裹（且括号前没有 -）
function isWrappedInParens(expr) {
  var s = expr.trim();
  if (!s.startsWith("(")) return false;
  var end = findMatchingRightParen(s, 0);
  return end === s.length - 1 && end >= 2 && s[end - 1] !== "-";
}

function normalizeNumberText(value) {
  if (typeof value === "number") return isFinite(value) ? String(value) : null;
  var s = String(value).replace(/ /g, " ").trim();
  if (!s) return null;
  if (isPureNumberExpr(s)) return s;
  var n = Number(s);
  return !isNaN(n) && isFinite(n) ? String(n) : null;
}

// ---- 格式转换（文本 ↔ 数字）工具函数 ----

// 增强版文本→数字解析：处理会计导出格式
function parseNumberFromText(text) {
  if (typeof text !== "string") return null;
  var cleaned = text.trim();
  if (!cleaned) return null;

  // 去除货币符号、千位分隔符和多余空格
  cleaned = cleaned.replace(/[¥$€£,，\s]/g, "");

  // 处理会计格式中的负数括号: "(1234.56)" → "-1234.56"
  if (cleaned.startsWith("(") && cleaned.endsWith(")")) {
    cleaned = "-" + cleaned.slice(1, -1);
  }

  var n = Number(cleaned);
  return (!isNaN(n) && isFinite(n)) ? n : null;
}

// ---- 日期识别工具函数 ----

// Excel 日期序列号：以 1899-12-30 为 day 0
function dateToExcelSerial(year, month, day) {
  // month is 1-12, day is 1-31
  var jsDate = new Date(Date.UTC(year, month - 1, day));
  var epoch = Date.UTC(1899, 11, 30); // Dec 30, 1899 = Excel day 0
  return Math.floor((jsDate.getTime() - epoch) / 86400000);
}

// 验证是否为合法日期（month 1-12, day 在当月范围内，含闰年）
function isValidDate(year, month, day) {
  if (year < 1900 || year > 9999) return false;
  if (month < 1 || month > 12) return false;
  if (day < 1 || day > 31) return false;
  var d = new Date(Date.UTC(year, month - 1, day));
  return d.getUTCFullYear() === year
      && d.getUTCMonth() === month - 1
      && d.getUTCDate() === day;
}

// 尝试将文本解析为日期，返回 { serial, format } 或 null
function parseDateFromText(text) {
  if (typeof text !== "string") return null;
  var s = text.trim();
  if (!s) return null;

  // 模式 1: 中文日期 — 2025年6月11日 / 2025年06月11日 / 2025年6月11
  var m1 = /^(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*(?:日)?\s*$/.exec(s);
  if (m1) {
    var y = parseInt(m1[1], 10), mo = parseInt(m1[2], 10), d = parseInt(m1[3], 10);
    if (isValidDate(y, mo, d)) {
      return { serial: dateToExcelSerial(y, mo, d), format: 'yyyy年m月d日' };
    }
    return null;
  }

  // 模式 2: 连字符/点号分隔 — 2025-06-11 / 2025/06/11 / 2025.06.11
  var m2 = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(s);
  if (m2) {
    var sep = s.charAt(4); // 分隔符位置在年份之后
    var y2 = parseInt(m2[1], 10), mo2 = parseInt(m2[2], 10), d2 = parseInt(m2[3], 10);
    if (isValidDate(y2, mo2, d2)) {
      var fmt;
      if (sep === "-") fmt = "yyyy-mm-dd";
      else if (sep === "/") fmt = "yyyy/mm/dd";
      else fmt = "yyyy.mm.dd";
      return { serial: dateToExcelSerial(y2, mo2, d2), format: fmt };
    }
    return null;
  }

  return null;
}

// ---- 文本 → 数字 ----

function convertTextToNumber(event) {
  Excel.run(async function (context) {
    var range = context.workbook.getSelectedRange();
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();

    // 只处理选中区域与已用区域的交集，避免全列选中时加载百万空行
    var processRange = range.getIntersection(usedRange);
    processRange.load(["formulas", "values", "numberFormat", "rowCount", "columnCount"]);

    try {
      await context.sync();
    } catch (e) {
      // 选中区域全部为空（无交集），直接结束
      if (e instanceof OfficeExtension.Error && e.code === "ItemNotFound") {
        event.completed();
        return;
      }
      throw e;
    }

    var formulas = processRange.formulas;
    var values = processRange.values;
    var numberFormats = processRange.numberFormat;
    var modified = false;

    for (var r = 0; r < processRange.rowCount; r++) {
      for (var c = 0; c < processRange.columnCount; c++) {
        var f = formulas[r][c];

        // 策略1：跳过公式单元格（保护底稿计算逻辑）
        if (typeof f === "string" && f.startsWith("=")) continue;
        // 已是数字的单元格：如果格式是 "@"（文本），修正为 "General"
        if (typeof f === "number") {
          if (numberFormats[r] && numberFormats[r][c] === "@") {
            if (!numberFormats[r]) { numberFormats[r] = []; }
            numberFormats[r][c] = "General";
            modified = true;
          }
          continue;
        }
        // 跳过空白
        if (f === null || f === undefined || f === "") continue;

        // 尝试文本→数字转换（日期优先）
        if (typeof f === "string") {
          // 优先尝试日期解析
          var dateResult = parseDateFromText(f);
          if (dateResult !== null) {
            values[r][c] = dateResult.serial;
            if (!numberFormats[r]) {
              numberFormats[r] = [];
            }
            numberFormats[r][c] = dateResult.format;
            modified = true;
            continue;
          }

          var num = parseNumberFromText(f);
          if (num !== null) {
            values[r][c] = num;
            // 确保 numberFormats 行存在
            if (!numberFormats[r]) {
              numberFormats[r] = [];
            }
            // 将文本格式 "@" 改为 "General"
            if (numberFormats[r][c] === "@") {
              numberFormats[r][c] = "General";
            }
            modified = true;
          }
          // 无法转换的文本（如公司名称、项目摘要）直接跳过
        }
      }
    }

    if (modified) {
      processRange.values = values;
      await context.sync();
      processRange.numberFormat = numberFormats;
      await context.sync();
    }
    range.select();
    await context.sync();
  }).catch(function (e) { console.error(e); })
    .finally(function () { event.completed(); });
}

// ---- 数字 → 文本 ----

function convertNumberToText(event) {
  Excel.run(async function (context) {
    var range = context.workbook.getSelectedRange();
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();

    var processRange = range.getIntersection(usedRange);
    // 防线1：额外加载 numberFormatLocal，Mac WKWebView 下比 numberFormat 更可靠
    processRange.load(["values", "text", "formulas", "numberFormat", "numberFormatLocal", "rowCount", "columnCount"]);

    try {
      await context.sync();
    } catch (e) {
      if (e instanceof OfficeExtension.Error && e.code === "ItemNotFound") {
        event.completed(); return;
      }
      throw e;
    }

    var values = processRange.values;
    var texts = processRange.text;
    var formulas = processRange.formulas;
    var formats = processRange.numberFormat;
    var localFormats = processRange.numberFormatLocal;
    var modified = false;

    for (var r = 0; r < processRange.rowCount; r++) {
      for (var c = 0; c < processRange.columnCount; c++) {
        var val = values[r][c];
        var txt = (texts[r] && texts[r][c] !== undefined && texts[r][c] !== null) ? String(texts[r][c]).trim() : "";
        var formula = formulas[r][c];
        var formatStr = (formats[r] && formats[r][c]) ? String(formats[r][c]) : "";
        var localFormatStr = (localFormats[r] && localFormats[r][c]) ? String(localFormats[r][c]) : "";

        // 策略1：跳过公式单元格
        if (typeof formula === "string" && formula.startsWith("=")) continue;

        // 只处理数字单元格
        if (typeof val !== "number" || !isFinite(val)) continue;

        // 防线2：增强正则 — 同时校验 numberFormat + numberFormatLocal
        // 注意：e/g 会误匹配 "General" / "G/通用格式"，已移除
        var dateFmtRegex = /[ymdhs年月日時分秒]/i;
        var isDateByFormat = dateFmtRegex.test(formatStr) || dateFmtRegex.test(localFormatStr);

        // 防线2补充：显示文本正则 — 兼容 - / . 全角破折号 等分隔符
        var isDateByText = /^\d{2,4}[-/\.－—–年]\d{1,2}[-/\.－—–]\d{1,2}/.test(txt);

        // 防线3：值域兜底 — 日期序列号范围 35000-75000（≈1996-2105），必须为整数
        var isDateByRange = val >= 35000 && val <= 75000 && val === Math.floor(val);

        if (isDateByFormat || isDateByText || isDateByRange) {
          // 如果 text 罢工返回了裸序列号（如 "46184"）或列宽不足返回了 "#"
          var txtIsSerial = (txt === String(val)) || /^#+$/.test(txt);

          if (txtIsSerial || !txt) {
            // 序列号反推日期：25569 = 1900-01-01 到 1970-01-01 的天数差
            var jsDate = new Date(Math.round((val - 25569) * 86400 * 1000));
            var y = jsDate.getUTCFullYear();
            var m = jsDate.getUTCMonth() + 1;
            var d = jsDate.getUTCDate();
            var mm = m < 10 ? "0" + m : String(m);
            var dd = d < 10 ? "0" + d : String(d);

            // 根据格式串猜测用户偏好的分隔符
            if (localFormatStr.indexOf("-") >= 0 || formatStr.indexOf("-") >= 0) {
              values[r][c] = y + "-" + mm + "-" + dd;
            } else {
              values[r][c] = y + "/" + mm + "/" + dd;
            }
          } else {
            // text 属性正常返回了日期字符串 → 直接使用
            values[r][c] = txt;
          }
        } else {
          // 普通数字：用 String(val) 保留全量精度
          values[r][c] = String(val);
        }

        if (!formats[r]) { formats[r] = []; }
        formats[r][c] = "@";
        modified = true;
      }
    }

    if (modified) {
      processRange.values = values;
      processRange.numberFormat = formats;
      await context.sync();
    }
    range.select();
    await context.sync();
  }).catch(function (e) { console.error(e); })
    .finally(function () { event.completed(); });
}

// ---- 添加 ROUND ----

function runAddRound(event, digits) {
  Excel.run(async function (context) {
    var range = context.workbook.getSelectedRange();
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();

    var processRange = range.getIntersection(usedRange);
    processRange.load(["formulas", "values", "rowCount", "columnCount"]);

    try {
      await context.sync();
    } catch (e) {
      if (e instanceof OfficeExtension.Error && e.code === "ItemNotFound") {
        event.completed(); return;
      }
      throw e;
    }

    var formulas = processRange.formulas, values = processRange.values, modified = false;
    for (var r = 0; r < formulas.length; r++) {
      for (var c = 0; c < formulas[r].length; c++) {
        var formula = formulas[r][c], value = values[r][c];
        if (typeof formula === "string" && /^=\s*ROUND\s*\(/i.test(formula)) continue;
        if (typeof formula === "string" && formula.startsWith("=")) {
          var inner = formula.replace(/^=\s*/, "");
          if (!inner) continue;
          formulas[r][c] = "=ROUND(" + inner + ", " + digits + ")";
          modified = true; continue;
        }
        if (value === null || value === "" || value === undefined) continue;
        var numText = normalizeNumberText(value);
        if (!numText) continue;
        formulas[r][c] = "=ROUND(" + numText + ", " + digits + ")";
        modified = true;
      }
    }
    if (modified) { processRange.formulas = formulas; await context.sync(); }
    range.select(); await context.sync();
  }).catch(function (e) { console.error(e); })
    .finally(function () { event.completed(); });
}

function addRound0(event) { runAddRound(event, 0); }
function addRound1(event) { runAddRound(event, 1); }
function addRound2(event) { runAddRound(event, 2); }
function addRound3(event) { runAddRound(event, 3); }
function addRound4(event) { runAddRound(event, 4); }

function addRoundCustom(event) {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/dialog-precision.html",
    { height: 30, width: 30, displayInIframe: true },
    function (result) {
      result.value.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        var n = parseInt(arg.message, 10);
        if (!isNaN(n) && n >= 0 && n <= 15) {
          result.value.close();
          runAddRound(event, n);
        } else {
          event.completed();
        }
      });
      result.value.addEventHandler(Office.EventType.DialogEventReceived, function () {
        event.completed();
      });
    }
  );
}

// ---- 移除 ROUND ----

function removeRound(event) {
  Excel.run(async function (context) {
    var range = context.workbook.getSelectedRange();
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();

    var processRange = range.getIntersection(usedRange);
    processRange.load(["formulas", "numberFormat", "rowCount", "columnCount"]);

    try {
      await context.sync();
    } catch (e) {
      if (e instanceof OfficeExtension.Error && e.code === "ItemNotFound") {
        event.completed(); return;
      }
      throw e;
    }

    var formulas = processRange.formulas, numberFormats = processRange.numberFormat, modified = false;
    for (var r = 0; r < formulas.length; r++) {
      for (var c = 0; c < formulas[r].length; c++) {
        var f = formulas[r][c];
        if (typeof f !== "string" || !f.startsWith("=")) continue;
        if (!/^=\s*ROUND\s*\(/i.test(f)) continue;
        var parsed = extractOuterRoundArgs(f);
        if (!parsed) continue;
        var inner = parsed.expr.replace(/^=\s*/, "").trim();
        if (!inner) continue;
        if (isPureNumberExpr(inner)) {
          var bf = String((numberFormats[r] && numberFormats[r][c]) || "");
          formulas[r][c] = String(parseFloat(inner));
          if (shouldRelaxFormatForConstant(bf, parsed.digits)) {
            if (!numberFormats[r]) { numberFormats[r] = []; for (var ci = 0; ci < processRange.columnCount; ci++) numberFormats[r].push("General"); }
            numberFormats[r][c] = "General";
          }
          modified = true; continue;
        }
        formulas[r][c] = "=" + inner; modified = true;
      }
    }
    if (modified) { processRange.formulas = formulas; if (numberFormats) processRange.numberFormat = numberFormats; await context.sync(); }
    range.select(); await context.sync();
  }).catch(function (e) { console.error(e); })
    .finally(function () { event.completed(); });
}

// ---- 正负转换 ----

// -(完整括号表达式) → 解包；其余 → -(body) 包裹
function negateFormula(formula) {
  var src = String(formula || "");
  if (!src.startsWith("=")) return null;
  var body = src.slice(1).trim();
  if (!body) return null;
  if (body.startsWith("-(") && isWrappedInParens(body.slice(1)))
    return "=" + body.slice(2, body.length - 1).trim();
  return "=-(" + body + ")";
}

function toggleSign(event) {
  Excel.run(async function (context) {
    var range = context.workbook.getSelectedRange();
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();

    var processRange = range.getIntersection(usedRange);
    processRange.load(["formulas", "rowCount", "columnCount"]);

    try {
      await context.sync();
    } catch (e) {
      if (e instanceof OfficeExtension.Error && e.code === "ItemNotFound") {
        event.completed(); return;
      }
      throw e;
    }

    var formulas = processRange.formulas, modified = false;
    for (var r = 0; r < range.rowCount; r++) {
      for (var c = 0; c < range.columnCount; c++) {
        var f = formulas[r][c];
        if (typeof f === "string" && f.startsWith("=")) {
          if (/^=\s*$/i.test(f)) continue;
          var n = negateFormula(f);
          if (n !== null) { formulas[r][c] = n; modified = true; }
          continue;
        }
        if (typeof f === "number" && isFinite(f)) {
          formulas[r][c] = -f; modified = true; continue;
        }
        var t = (typeof f === "string" && !f.startsWith("=")) ? f.trim() : "";
        if (!t) continue;
        var num = Number(t);
        if (isFinite(num)) { formulas[r][c] = String(-num); modified = true; }
      }
    }
    if (modified) { processRange.formulas = formulas; await context.sync(); }
    range.select(); await context.sync();
  }).catch(function (e) { console.error(e); })
    .finally(function () { event.completed(); });
}

// ---- 注册 ----
Office.onReady(function () {
  Office.actions.associate("addRound2", addRound2);
  Office.actions.associate("addRound0", addRound0);
  Office.actions.associate("addRound1", addRound1);
  Office.actions.associate("addRound3", addRound3);
  Office.actions.associate("addRound4", addRound4);
  Office.actions.associate("addRoundCustom", addRoundCustom);
  Office.actions.associate("removeRound", removeRound);
  Office.actions.associate("toggleSign", toggleSign);
  Office.actions.associate("convertTextToNumber", convertTextToNumber);
  Office.actions.associate("convertNumberToText", convertNumberToText);
});
