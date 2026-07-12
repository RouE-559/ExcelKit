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

// ---- 添加 ROUND ----

function runAddRound(event, digits) {
  Excel.run(async function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(["formulas", "values", "rowCount", "columnCount"]);
    await context.sync();
    var formulas = range.formulas, values = range.values, modified = false;
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
    if (modified) { range.formulas = formulas; await context.sync(); }
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
    range.load(["formulas", "numberFormat", "rowCount", "columnCount"]);
    await context.sync();
    var formulas = range.formulas, numberFormats = range.numberFormat, modified = false;
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
            if (!numberFormats[r]) { numberFormats[r] = []; for (var ci = 0; ci < range.columnCount; ci++) numberFormats[r].push("General"); }
            numberFormats[r][c] = "General";
          }
          modified = true; continue;
        }
        formulas[r][c] = "=" + inner; modified = true;
      }
    }
    if (modified) { range.formulas = formulas; if (numberFormats) range.numberFormat = numberFormats; await context.sync(); }
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
    range.load(["formulas", "rowCount", "columnCount"]);
    await context.sync();
    var formulas = range.formulas, modified = false;
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
    if (modified) { range.formulas = formulas; await context.sync(); }
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
});
