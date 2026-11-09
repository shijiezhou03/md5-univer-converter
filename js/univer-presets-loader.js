/*
 * univer-presets-loader.js — 加载被拆分的 @univerjs/presets 0.10.1 UMD。
 *
 * 背景：vendor/univer-presets.umd.js 原文件约 7.5 MB，超过 GitHub 单次推送的
 * 请求大小上限（约 3.5 MB 可用、7.5 MB 单文件返回 HTTP 413），因此把原文件
 * 按字符串拆成 3 片，分别存放在：
 *   vendor/univer-presets.chunk1.js / chunk2.js / chunk3.js
 * 每片只做一件事：把自己那段字符串 push 进 window.__UNIVER_PRESETS_CHUNKS__。
 * 三片拼接后与原文件逐字节一致（SHA-256 校验值见 README）。
 *
 * 本加载器在三片之后同步执行：拼接字符串 → 以全局作用域执行（动态创建
 * <script> 元素注入文本，与直接加载原文件等价），使 UMD 按原方式注册
 * window.UniverPresets / window.UniverCore 等全局。执行是同步的，发生在
 * 后续 preset-sheets-core 脚本与 js/app.js 之前，因此 app.js 原有的判断
 * 方式（启动时检查 window.UniverPresets.createUniver 是否存在，缺失则
 * 抛出错误并切换备用转换区）无需任何改动即可继续工作。
 *
 * 失败时：把错误记录到 window.__UNIVER_PRESETS_LOAD_ERROR__ 并打印到控制台，
 * window.UniverPresets 保持未定义，app.js 会走原有的备用区逻辑，页面不白屏。
 */
(function () {
  'use strict';

  function fail(err) {
    window.__UNIVER_PRESETS_READY__ = false;
    window.__UNIVER_PRESETS_LOAD_ERROR__ = err;
    if (window.console && typeof console.error === 'function') {
      console.error('[univer-presets-loader] Univer presets 加载失败：', err);
    }
  }

  var chunks = window.__UNIVER_PRESETS_CHUNKS__;
  if (!chunks || chunks.length !== 3) {
    fail(new Error('presets 分片缺失（应有 3 片，实际 ' +
      (chunks ? chunks.length : 0) + ' 片）：请确认 vendor/univer-presets.chunk1-3.js 都已加载'));
    return;
  }

  var code = chunks.join('');
  window.__UNIVER_PRESETS_CHUNKS__ = null; // 释放分片字符串，执行后不再需要

  try {
    // 动态内联 <script>：插入即在全局作用域同步执行，语义与直接
    // <script src="vendor/univer-presets.umd.js"> 加载原文件一致。
    var script = document.createElement('script');
    script.textContent = code + '\n//# sourceURL=vendor/univer-presets.umd.js';
    document.head.appendChild(script);
    if (script.parentNode) script.parentNode.removeChild(script);
  } catch (e) {
    // 极少数环境下动态脚本注入不可用时，退回全局间接 eval。
    try {
      (0, eval)(code + '\n//# sourceURL=vendor/univer-presets.umd.js');
    } catch (e2) {
      fail(e2);
      return;
    }
  }

  if (window.UniverPresets && window.UniverPresets.createUniver) {
    window.__UNIVER_PRESETS_READY__ = true;
  } else {
    fail(new Error('presets 代码已执行，但 window.UniverPresets 未注册（执行过程中可能出错，见控制台）'));
  }
})();
