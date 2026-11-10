/*
 * app.js — Univer 版 MD5 转换表的主逻辑。
 *
 * 功能：
 *  - 用 Univer（@univerjs/presets + preset-sheets-core 0.10.1）渲染真正的电子表格；
 *  - A 列输入/粘贴原文，监听 SheetValueChanged 事件，自动在 B 列填 32 位小写 MD5
 *    （UTF-8，标准 MD5，与 MaxCompute MD5() 一致，见 js/md5.js）；
 *  - A 列设为文本格式（'@'），保住 00123 这类前导零；
 *  - 预置 1000 行，粘贴更多行时自动加行；
 *  - 工具栏：导入 .xlsx/.csv、批量粘贴、复制全部 MD5、复制 A+B、导出 CSV、
 *    大小写切换、去首尾空格、清空；
 *  - Univer 加载失败时自动切换到备用文本转换区，页面不会白屏。
 * 全部依赖均为本地相对路径文件，运行时零外网请求。
 */
(function () {
  'use strict';

  var state = {
    api: null,
    sheet: null,
    ready: false,
    timer: null,
    upper: false,
    trim: false
  };

  function el(id) { return document.getElementById(id); }

  function setStatus(msg) {
    var s = el('status');
    if (s) s.textContent = msg;
  }

  /* ---------------- MD5 计算 ---------------- */

  function hashFor(raw) {
    if (raw === null || raw === undefined) return '';
    var s = String(raw);
    if (state.trim) s = s.replace(/^\s+|\s+$/g, '');
    if (s === '') return '';
    var h = window.md5(s);
    return state.upper ? h.toUpperCase() : h;
  }

  /* ---------------- 表格读写 ---------------- */

  function sheetMaxRows() {
    try { return state.sheet.getMaxRows(); } catch (e) { return 1000; }
  }

  function lastUsedRow() {
    try {
      var dr = state.sheet.getDataRange();
      if (!dr) return 0;
      // Univer 0.10.1 的 FRange 没有 getRowCount()，用 getLastRow()（或 getRow()+getHeight()-1）。
      if (typeof dr.getLastRow === 'function') return Math.max(0, dr.getLastRow());
      return Math.max(0, dr.getRow() + dr.getHeight() - 1);
    } catch (e) {
      return 0;
    }
  }

  function applyTextFormat(fromRow, toRow) {
    if (toRow < fromRow) return;
    try {
      state.sheet.getRange(fromRow, 0, toRow - fromRow + 1, 1).setNumberFormat('@');
    } catch (e) { /* 数字格式失败不影响 MD5 计算 */ }
  }

  function ensureRows(neededLastRow) {
    var max = sheetMaxRows();
    if (neededLastRow >= max - 1) {
      var add = Math.max(500, neededLastRow - max + 50);
      try {
        state.sheet.insertRowsAfter(max - 1, add);
        applyTextFormat(max, max + add - 1);
      } catch (e) { /* 加行失败时已有行仍可正常使用 */ }
    }
  }

  function readAB(last) {
    return state.sheet.getRange(0, 0, last + 1, 2).getValues();
  }

  // 重新计算全部 B 列：只写入与期望值不一致的连续区间，避免无谓刷新与事件循环。
  function recompute() {
    if (!state.ready) return;
    var last = lastUsedRow();
    if (last < 1) { setCount(0); return; }
    ensureRows(last + 1);
    var vals;
    try { vals = readAB(last); } catch (e) { return; }

    var segments = [];
    var segStart = -1;
    var segVals = [];
    var count = 0;
    for (var r = 1; r <= last; r++) {
      var a = vals[r] ? vals[r][0] : null;
      var b = vals[r] ? vals[r][1] : null;
      var expect = hashFor(a);
      if (expect) count++;
      var cur = (b === null || b === undefined) ? '' : String(b);
      if (cur !== expect) {
        if (segStart < 0) { segStart = r; segVals = []; }
        segVals.push([expect]);
      } else if (segStart >= 0) {
        segments.push({ start: segStart, vals: segVals });
        segStart = -1; segVals = [];
      }
    }
    if (segStart >= 0) segments.push({ start: segStart, vals: segVals });

    for (var i = 0; i < segments.length; i++) {
      try {
        state.sheet.getRange(segments[i].start, 1, segments[i].vals.length, 1)
          .setValues(segments[i].vals);
      } catch (e) { /* 忽略单段写失败 */ }
    }
    setCount(count);
  }

  function setCount(n) {
    var s = el('count');
    if (s) s.textContent = '已生成 MD5：' + n + ' 行';
  }

  function scheduleRecompute() {
    if (state.timer) clearTimeout(state.timer);
    state.timer = setTimeout(recompute, 200);
  }

  function clearData() {
    if (!state.ready) return;
    var last = lastUsedRow();
    if (last < 1) return;
    var blank = [];
    for (var r = 1; r <= last; r++) blank.push(['', '']);
    for (var i = 0; i < blank.length; i += 1000) {
      var chunk = blank.slice(i, i + 1000);
      try { state.sheet.getRange(1 + i, 0, chunk.length, 2).setValues(chunk); } catch (e) {}
    }
    setCount(0);
  }

  function firstEmptyRow() {
    var last = lastUsedRow();
    if (last < 1) return 1;
    var vals;
    try { vals = state.sheet.getRange(0, 0, last + 1, 1).getValues(); } catch (e) { return last + 1; }
    for (var r = 1; r <= last; r++) {
      var v = vals[r] ? vals[r][0] : null;
      if (v === null || v === undefined || String(v) === '') return r;
    }
    return last + 1;
  }

  // 把一批字符串写入 A 列（从第 1 行数据行开始覆盖，或追加到末尾）。
  function fillLines(lines, append) {
    if (!state.ready) { setStatus('表格尚未就绪，无法填入。'); return 0; }
    if (!append) clearData();
    var start = append ? firstEmptyRow() : 1;
    ensureRows(start + lines.length - 1);
    for (var i = 0; i < lines.length; i += 1000) {
      var chunk = [];
      for (var j = i; j < Math.min(i + 1000, lines.length); j++) chunk.push([lines[j]]);
      try { state.sheet.getRange(start + i, 0, chunk.length, 1).setValues(chunk); } catch (e) {}
    }
    recompute();
    setStatus('已填入 ' + lines.length + ' 行，B 列 MD5 已生成。');
    return lines.length;
  }

  /* ---------------- 文本解析（Excel TSV / CSV / 多行文本） ---------------- */

  // 支持带引号的 CSV/TSV：引号内可含分隔符、换行与双引号转义。
  function parseDelimited(text) {
    var rows = [];
    var row = [];
    var cur = '';
    var inQuotes = false;
    var s = String(text);
    for (var n = 0; n < s.length; n++) {
      var ch = s.charAt(n);
      if (inQuotes) {
        if (ch === '"') {
          if (s.charAt(n + 1) === '"') { cur += '"'; n++; }
          else inQuotes = false;
        } else {
          cur += ch;
        }
      } else if (ch === '"') {
        inQuotes = true;
      } else if (ch === '\t' || ch === ',') {
        row.push(cur); cur = '';
      } else if (ch === '\n') {
        row.push(cur); cur = '';
        rows.push(row); row = [];
      } else if (ch === '\r') {
        // \r\n 或单独 \r 作为行分隔：交给随后的 \n 处理或在此断行
        if (s.charAt(n + 1) !== '\n') {
          row.push(cur); cur = '';
          rows.push(row); row = [];
        }
      } else {
        cur += ch;
      }
    }
    row.push(cur);
    rows.push(row);
    return rows;
  }

  // 按工具栏选项（取第几列、跳过表头、跳过空行）从二维数组中提炼出要转换的字符串列。
  function pickLines(rows) {
    var col = parseInt(el('opt-col').value, 10) || 0;
    var skipHeader = el('opt-skip-header').checked;
    var skipEmpty = el('opt-skip-empty').checked;
    var out = [];
    for (var r = 0; r < rows.length; r++) {
      if (r === 0 && skipHeader && rows.length > 1) continue;
      var v = rows[r] && rows[r][col] !== undefined && rows[r][col] !== null
        ? String(rows[r][col]) : '';
      if (skipEmpty && v.replace(/^\s+|\s+$/g, '') === '') continue;
      out.push(v);
    }
    // 去掉末尾纯空行（Excel 复制常带尾随换行）
    while (out.length && out[out.length - 1] === '') out.pop();
    return out;
  }

  function linesFromText(text) {
    return pickLines(parseDelimited(text));
  }

  /* ---------------- 剪贴板 ---------------- */

  function fallbackCopy(text, doneMsg) {
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    try {
      document.execCommand('copy');
      setStatus(doneMsg);
    } catch (e) {
      setStatus('复制失败，请在表格中手动选择复制。');
    }
    document.body.removeChild(ta);
  }

  function copyText(text, doneMsg) {
    if (!text) { setStatus('没有可复制的内容。'); return; }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(
        function () { setStatus(doneMsg); },
        function () { fallbackCopy(text, doneMsg); }
      );
    } else {
      fallbackCopy(text, doneMsg);
    }
  }

  function collectRows() {
    var last = lastUsedRow();
    var rows = [];
    if (last < 1) return rows;
    var vals = readAB(last);
    for (var r = 1; r <= last; r++) {
      var a = vals[r] ? vals[r][0] : null;
      var aStr = (a === null || a === undefined) ? '' : String(a);
      rows.push({ a: aStr, b: hashFor(a) });
    }
    return rows;
  }

  /* ---------------- 导入文件 ---------------- */

  function importFile(file) {
    if (!file) return;
    var isCsv = /\.csv$/i.test(file.name);
    if (typeof XLSX === 'undefined' && !isCsv) {
      setStatus('SheetJS 未加载，无法解析 Excel 文件；可改用复制粘贴。');
      return;
    }
    var reader = new FileReader();
    reader.onload = function (ev) {
      try {
        var rows;
        if (typeof XLSX !== 'undefined') {
          var wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
          var ws = wb.Sheets[wb.SheetNames[0]];
          rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
        } else {
          rows = parseDelimited(new TextDecoder('utf-8').decode(ev.target.result));
        }
        var lines = pickLines(rows.map(function (row) {
          return row.map(function (c) {
            return c === null || c === undefined ? '' : String(c);
          });
        }));
        if (!lines.length) { setStatus('文件中没有读到可转换的内容。'); return; }
        fillLines(lines, el('opt-mode').value === 'append');
      } catch (err) {
        setStatus('导入失败：' + err.message);
      }
    };
    reader.onerror = function () { setStatus('读取文件失败。'); };
    reader.readAsArrayBuffer(file);
  }

  /* ---------------- CSV 导出 ---------------- */

  function csvEscape(v) {
    var s = String(v === null || v === undefined ? '' : v);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  function exportCsv() {
    var rows = collectRows();
    var lines = ['原文,MD5'];
    rows.forEach(function (row) {
      if (row.a === '' && row.b === '') return;
      lines.push(csvEscape(row.a) + ',' + csvEscape(row.b));
    });
    var blob = new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'md5.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
    setStatus('已导出 CSV（' + (lines.length - 1) + ' 行）。');
  }

  /* ---------------- Univer 初始化 ---------------- */

  function initUniver() {
    if (!window.UniverPresets || !window.UniverPresets.createUniver) {
      throw new Error('UniverPresets 未加载（vendor/univer-presets.umd.js）');
    }
    if (!window.UniverPresetSheetsCore ||
        !window.UniverPresetSheetsCore.UniverSheetsCorePreset) {
      throw new Error('UniverPresetSheetsCore 未加载（vendor/univer-preset-sheets-core.umd.js）');
    }
    var LocaleType = window.UniverCore.LocaleType;
    var locales = {};
    var zhCN = window.UniverPresetSheetsCoreZhCN;
    if (zhCN) {
      locales[LocaleType.ZH_CN] = window.UniverCore.mergeLocales
        ? window.UniverCore.mergeLocales(zhCN)
        : zhCN;
    }

    var created = window.UniverPresets.createUniver({
      locale: LocaleType.ZH_CN,
      locales: locales,
      presets: [
        window.UniverPresetSheetsCore.UniverSheetsCorePreset({
          container: 'univer-app'
        })
      ]
    });
    state.api = created.univerAPI;

    var workbook = state.api.createWorkbook({
      id: 'md5-workbook',
      name: 'MD5转换',
      sheetOrder: ['sheet1'],
      sheets: {
        sheet1: {
          id: 'sheet1',
          name: 'MD5转换',
          rowCount: 1000,
          columnCount: 26,
          cellData: {
            0: {
              0: { v: '原文 / 字符串', s: { bl: 1 } },
              1: { v: 'MD5（32 位小写）', s: { bl: 1 } }
            }
          }
        }
      }
    });
    state.sheet = workbook.getActiveSheet();

    try { state.sheet.setColumnWidth(0, 380); } catch (e) {}
    try { state.sheet.setColumnWidth(1, 330); } catch (e) {}
    try { state.sheet.setFrozenRows(1); } catch (e) {}
    applyTextFormat(0, 999);

    state.api.addEvent(state.api.Event.SheetValueChanged, function () {
      scheduleRecompute();
    });

    state.ready = true;
    setStatus('就绪：在 A 列输入或粘贴字符串，B 列会自动生成 MD5。');
  }

  /* ---------------- 备用转换区（Univer 加载失败时） ---------------- */

  function showFallback(reason) {
    var fb = el('fallback');
    var app = el('univer-app');
    if (app) app.style.display = 'none';
    if (fb) {
      fb.hidden = false;
      var msg = el('fallback-msg');
      if (msg) msg.textContent = reason;
    }
    setStatus('已切换到备用转换区（每行一个字符串，实时转换）。');
  }

  function convertFallback() {
    var input = el('fb-in').value;
    var lines = input.split('\n');
    var out = lines.map(function (line) { return hashFor(line); });
    el('fb-out').value = out.join('\n');
  }

  /* ---------------- 工具栏与全局事件 ---------------- */

  function wireUi() {
    el('opt-upper').addEventListener('change', function () {
      state.upper = this.checked;
      recompute();
      if (!el('fallback').hidden) convertFallback();
    });
    el('opt-trim').addEventListener('change', function () {
      state.trim = this.checked;
      recompute();
      if (!el('fallback').hidden) convertFallback();
    });

    el('file-input').addEventListener('change', function () {
      if (this.files && this.files[0]) importFile(this.files[0]);
      this.value = '';
    });

    el('btn-batch-toggle').addEventListener('click', function () {
      var box = el('batch-box');
      box.hidden = !box.hidden;
      if (!box.hidden) el('batch-text').focus();
    });

    el('btn-batch-apply').addEventListener('click', function () {
      var text = el('batch-text').value;
      if (!text) { setStatus('批量粘贴区是空的。'); return; }
      var lines = linesFromText(text);
      if (!lines.length) { setStatus('没有解析到可填入的行。'); return; }
      fillLines(lines, el('opt-mode').value === 'append');
    });

    el('btn-copy-md5').addEventListener('click', function () {
      var rows = collectRows().filter(function (r) { return r.b !== ''; });
      copyText(rows.map(function (r) { return r.b; }).join('\n'),
        '已复制 ' + rows.length + ' 条 MD5。');
    });

    el('btn-copy-both').addEventListener('click', function () {
      var rows = collectRows().filter(function (r) { return r.a !== '' || r.b !== ''; });
      copyText(rows.map(function (r) { return r.a + '\t' + r.b; }).join('\n'),
        '已复制 A+B 共 ' + rows.length + ' 行（制表符分隔，可直接粘回 Excel）。');
    });

    el('btn-export').addEventListener('click', exportCsv);

    el('btn-clear').addEventListener('click', function () {
      clearData();
      setStatus('已清空（表头保留）。');
    });

    el('fb-in').addEventListener('input', convertFallback);

    // 页面任意位置粘贴（焦点不在表格、不在输入框时）：整表填入 A 列。
    document.addEventListener('paste', function (e) {
      if (!state.ready) return;
      var t = e.target;
      if (t && (t.tagName === 'TEXTAREA' || t.tagName === 'INPUT' || t.tagName === 'SELECT')) return;
      if (t && t.closest && t.closest('#univer-app')) return; // 表格内粘贴交给 Univer 原生处理
      var cd = e.clipboardData || window.clipboardData;
      if (!cd) return;
      var text = cd.getData('text');
      if (!text) return;
      e.preventDefault();
      var lines = linesFromText(text);
      if (lines.length) {
        fillLines(lines, el('opt-mode').value === 'append');
      }
    });

    // 拖放导入 .xlsx/.csv
    document.addEventListener('dragover', function (e) { e.preventDefault(); });
    document.addEventListener('drop', function (e) {
      e.preventDefault();
      if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]) {
        importFile(e.dataTransfer.files[0]);
      }
    });
  }

  function boot() {
    wireUi();
    try {
      initUniver();
    } catch (err) {
      showFallback('Univer 表格加载失败：' + (err && err.message ? err.message : err));
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
})();
