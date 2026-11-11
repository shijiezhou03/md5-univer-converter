/*
 * app.js — Univer 版 MD5 转换表的主逻辑。
 *
 * 功能：
 *  - 用 Univer（@univerjs/presets + preset-sheets-core 0.10.1）渲染真正的电子表格；
 *  - A 列输入/粘贴原文，监听 SheetValueChanged 事件，自动在 B 列填 32 位小写 MD5
 *    （UTF-8，标准 MD5，与 MaxCompute MD5() 一致，见 js/md5.js）；
 *  - 写入 A 列时用显式字符串类型（ICellData t=STRING），粘贴/导入的 00123、1e5
 *    等「数字样」字符串不会被 Univer 转成数字（其 setValues/编辑器默认会转）；
 *  - 预置 1048576 行（Excel 行数上限）：Univer 的 insertRowsAfter 成本随表变大
 *    线性上升（实测 4.6 万行时插 5000 行要 25 秒），预置满行数后永不加行；
 *  - 数据量上限：单次粘贴/导入超过 100 万行、或导入文件超过 25 MB 时，
 *    直接拒绝并在状态栏提示分批次处理，不硬算到页面卡死；
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
    trim: false,
    busySeq: 0,
    recomputeSeq: 0,
    pasteSeq: 0,
    passDirty: false
  };

  // 让出主线程一帧，保证大批量处理时页面不假死、进度能刷出来。
  function tick() {
    return new Promise(function (resolve) { setTimeout(resolve, 0); });
  }

  // UniverCore.CellValueType.STRING：写入单元格时显式声明字符串类型，
  // 阻止 Univer 把 '00123'、'1e5' 这类字符串解析成数字。
  function stringType() {
    try {
      if (window.UniverCore && window.UniverCore.CellValueType) {
        return window.UniverCore.CellValueType.STRING;
      }
    } catch (e) { /* 忽略 */ }
    return 1;
  }

  function el(id) { return document.getElementById(id); }

  // 数据量上限：单次超过 100 万行直接拒绝并提示分批（浏览器内存与表格
  // 行数上限 1048576 都兜不住更大的量）；导入文件超过 25 MB 同样先拦下，
  // 避免解析超大 Excel 时页面卡死/崩溃，用户拆小后再导。
  var MAX_BATCH_ROWS = 1000000;
  var MAX_FILE_BYTES = 25 * 1024 * 1024;

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

  // 分块加行：单次 insertRowsAfter 太大（上万行）时 Univer 可能失败或长时间卡住，
  // 改成每块至多 5000 行逐块插入；返回是否已扩到目标行。
  function ensureRows(neededLastRow) {
    var max = sheetMaxRows();
    if (neededLastRow < max - 1) return true;
    var target = neededLastRow + 50;
    var guard = 0;
    while (sheetMaxRows() <= target && guard < 200) {
      guard++;
      var cur = sheetMaxRows();
      var add = Math.min(5000, target - cur + 1);
      try {
        state.sheet.insertRowsAfter(cur - 1, add);
      } catch (e) {
        return false;
      }
      if (sheetMaxRows() <= cur) return false; // 插入没有生效，避免死循环
    }
    applyTextFormat(max, sheetMaxRows() - 1);
    return sheetMaxRows() > neededLastRow;
  }

  function readAB(last) {
    return state.sheet.getRange(0, 0, last + 1, 2).getValues();
  }

  // 计算一段行（1 起、含两端）的 B 列期望值并回写不一致处，返回其中非空 MD5 的行数。
  // 分段读写，避免一次性构造超大数组；大批量时按段调用、段间让出主线程。
  function computeRows(fromRow, toRow) {
    var vals;
    try {
      vals = state.sheet.getRange(fromRow, 0, toRow - fromRow + 1, 2).getValues();
    } catch (e) {
      return 0;
    }
    var count = 0;
    var segStart = -1;
    var segVals = [];
    function flush(endRow) {
      if (segStart < 0) return;
      try {
        state.sheet.getRange(segStart, 1, segVals.length, 1).setValues(segVals);
      } catch (e) { /* 忽略单段写失败 */ }
      segStart = -1;
      segVals = [];
    }
    for (var i = 0; i < vals.length; i++) {
      var r = fromRow + i;
      var a = vals[i] ? vals[i][0] : null;
      var b = vals[i] ? vals[i][1] : null;
      var expect = hashFor(a);
      if (expect) count++;
      var cur = (b === null || b === undefined) ? '' : String(b);
      if (cur !== expect) {
        if (segStart < 0) segStart = r;
        // B 列同样显式写字符串类型：纯数字组成的哈希不能被转成数字。
        segVals.push([{ v: expect, t: stringType() }]);
      } else {
        flush(r - 1);
      }
    }
    flush(toRow);
    return count;
  }

  // 重新计算全部 B 列。行数少时同步完成；行数多时分段异步算，
  // 期间有新编辑则放弃本轮（新一轮会接上），保证大表也不假死、不漏算。
  function recompute() {
    if (!state.ready || state.busySeq !== 0) return;
    var last = lastUsedRow();
    if (last < 1) { setCount(0); return; }
    ensureRows(last + 1);
    if (last <= 3000) {
      setCount(computeRows(1, last));
      return;
    }
    var seq = ++state.recomputeSeq;
    state.busySeq = seq;
    state.passDirty = false;
    (async function () {
      var total = 0;
      var curLast = last;
      try {
        for (var start = 1; start <= curLast; start += 2000) {
          total += computeRows(start, Math.min(start + 1999, curLast));
          setStatus('正在生成 MD5… ' + Math.min(start + 1999, curLast) + ' / ' + curLast + ' 行');
          await tick();
          curLast = Math.max(curLast, lastUsedRow());
        }
        setCount(total);
        setStatus('就绪：在 A 列输入或粘贴字符串，B 列会自动生成 MD5。');
      } finally {
        if (state.busySeq === seq) state.busySeq = 0;
        // 本轮期间有过变更（包括用户的新编辑）：再补算一轮，保证不漏。
        if (state.passDirty) {
          state.passDirty = false;
          recompute();
        }
      }
    })();
  }

  function setCount(n) {
    var s = el('count');
    if (s) s.textContent = '已生成 MD5：' + n + ' 行';
  }

  function scheduleRecompute() {
    // 大批量计算进行中：它自己写 B 列也会触发变更事件，不能据此作废重跑
    // （会反复从头开始）；只记一笔「期间有变更」，本轮结束后再补算一轮。
    if (state.busySeq > 0) { state.passDirty = true; return; }
    if (state.busySeq !== 0) return; // 批量填入中：填完会统一重算
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

  // 把一批字符串从 start 行起写入 A 列。
  // 分块异步写入：每块至多 1000 行，块间让出主线程并刷新进度，几万行也不会假死；
  // 写完逐块回读校验，状态栏报的是实际写入行数，写失败会明确报错而不是假装成功。
  async function doFill(lines, start) {
    if (!state.ready) { setStatus('表格尚未就绪，无法填入。'); return 0; }
    if (lines.length > MAX_BATCH_ROWS) {
      setStatus('数据量太大：本次 ' + lines.length + ' 行，超过 100 万行上限，无法一次加密，请分批次处理（每批不超过 100 万行）。');
      return 0;
    }
    if (start + lines.length - 1 > 1048575) {
      setStatus('数据量太大：加上已有数据将超过表格 1048576 行上限，无法一次加密，请清空表格或分批次处理。');
      return 0;
    }
    if (state.busySeq !== 0) { setStatus('正在处理上一批数据，请稍候再试。'); return 0; }
    state.busySeq = -1;
    state.recomputeSeq++;
    try {
      if (!ensureRows(start + lines.length - 1)) {
        setStatus('表格行数扩展失败，只填入了前 ' + (sheetMaxRows() - start) + ' 行。');
        return 0;
      }
      var written = 0;
      var failed = false;
      for (var i = 0; i < lines.length; i += 1000) {
        var chunk = [];
        for (var j = i; j < Math.min(i + 1000, lines.length); j++) {
          chunk.push([{ v: lines[j], t: stringType() }]);
        }
        try {
          state.sheet.getRange(start + i, 0, chunk.length, 1).setValues(chunk);
          var check = state.sheet.getRange(start + i, 0, 1, 1).getValues();
          if (!check || String(check[0][0] == null ? '' : check[0][0]) !== String(lines[i])) {
            failed = true;
            break;
          }
          written += chunk.length;
        } catch (e) {
          failed = true;
          break;
        }
        if (lines.length > 3000) {
          setStatus('正在填入… ' + written + ' / ' + lines.length + ' 行');
          await tick();
        }
      }
      if (failed) {
        setStatus('填入中断：只成功写入 ' + written + ' / ' + lines.length + ' 行，请减小批量后重试。');
        return written;
      }
    } finally {
      if (state.busySeq === -1) state.busySeq = 0;
    }
    recompute();
    setStatus('已填入 ' + lines.length + ' 行，B 列 MD5 已生成。');
    return lines.length;
  }

  // 整表填入（批量粘贴/导入/页面粘贴用）：替换模式先清空，追加模式接在末尾。
  async function fillLines(lines, append) {
    if (!append) clearData();
    var start = append ? firstEmptyRow() : 1;
    return doFill(lines, start);
  }

  // 粘贴巡检：把被 Univer 异步粘贴改成数字的格改回字符串原文（判定见粘贴处理处）。
  function settlePaste(lines, startRow, token) {
    if (token !== state.pasteSeq || !state.ready || state.busySeq !== 0) return;
    var vals;
    try {
      vals = state.sheet.getRange(startRow, 0, lines.length, 1).getValues();
    } catch (e) {
      return;
    }
    var fixed = 0;
    for (var i = 0; i < lines.length; i++) {
      var want = lines[i];
      var cur = vals[i] ? vals[i][0] : null;
      var curStr = (cur === null || cur === undefined) ? '' : String(cur);
      if (curStr === want) continue;
      var wantNum = Number(want);
      var curNum = Number(curStr);
      if (want !== '' && isFinite(wantNum) && isFinite(curNum) && wantNum === curNum) {
        try {
          state.sheet.getRange(startRow + i, 0, 1, 1)
            .setValues([[{ v: want, t: stringType() }]]);
          fixed++;
        } catch (e) { /* 忽略单格失败 */ }
      }
    }
    if (fixed) scheduleRecompute();
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
    if (file.size > MAX_FILE_BYTES) {
      setStatus('文件太大：' + file.name + ' 约 ' + (file.size / 1048576).toFixed(1) + ' MB，超过 25 MB 导入上限，无法导入，请拆分成小文件（每批不超过 100 万行）分批导入。');
      return;
    }
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
          rowCount: 1048576,
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
    applyTextFormat(0, 1048575);

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

    // 粘贴统一由本页接管（捕获阶段，先于 Univer 处理）：直接读剪贴板原文，
    // 按选定的列提炼后写入 A 列。走 Univer 原生粘贴会有两个问题：
    // 1) '00123'、'1e5' 等字符串会被它解析成数字，原文丢失；
    // 2) 大批量粘贴时它内部处理极慢。接管后两种问题都不存在。
    // 注意粘贴事件的 target 不一定在表格容器内（焦点元素可能是 body 或
    // Univer 的隐藏剪贴板元素），不能用 target 位置判断是否表格粘贴：
    // 焦点在表格/页面上时按当前选区行填入，在工具栏等区域时按替换/追加模式整表填。
    document.addEventListener('paste', function (e) {
      if (!state.ready || state.busySeq !== 0) return;
      var t = e.target;
      // 注意：不能按 isContentEditable 排除 —— Univer 的剪贴板代理元素就是
      // 一个 contentEditable 的 DIV，表格粘贴的 target 恰恰是它。
      if (t && (t.tagName === 'TEXTAREA' || t.tagName === 'INPUT' ||
                t.tagName === 'SELECT')) return;
      var cd = e.clipboardData || window.clipboardData;
      if (!cd) return;
      var text = '';
      try { text = cd.getData('text'); } catch (err) { return; }
      if (!text) return; // 无文本（如图片）时交回默认处理
      var lines = linesFromText(text);
      if (!lines.length) return;
      e.preventDefault();
      e.stopImmediatePropagation();
      var inGrid = (t && t.closest && t.closest('#univer-app')) || t === document.body;
      if (!inGrid) {
        fillLines(lines, el('opt-mode').value === 'append');
        return;
      }
      var startRow = 1;
      try {
        var sel = state.sheet.getSelection();
        var cell = sel && typeof sel.getCurrentCell === 'function' ? sel.getCurrentCell() : null;
        if (cell && typeof cell.startRow === 'number') startRow = Math.max(1, cell.startRow);
        else if (sel && typeof sel.getRow === 'function') startRow = Math.max(1, sel.getRow());
      } catch (err) { /* 取不到选区时从第 1 行数据行开始 */ }
      doFill(lines, startRow);
      // Univer 还会通过 Ctrl+V 按键路径异步读剪贴板再粘一次（DOM 事件拦不住），
      // 可能把刚写好的字符串又改成数字。延迟巡检几轮：发现某格是对应原文的
      // 「数字变形」（数值相等但文本不同）就改回字符串原文，其余不动。
      var token = ++state.pasteSeq;
      [400, 1500, 4000].forEach(function (delay) {
        setTimeout(function () { settlePaste(lines, startRow, token); }, delay);
      });
    }, true);

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
