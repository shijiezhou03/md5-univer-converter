/*
 * md5.js — 标准 MD5 实现（RFC 1321），纯 JavaScript，无依赖。
 * 输入按 UTF-8 编码后计算，输出 32 位小写十六进制。
 * 与 MaxCompute 的 MD5() 函数结果一致（MaxCompute 对字符串按 UTF-8 计算 MD5 并输出小写 hex）。
 *
 * 浏览器中挂载为 window.md5；Node.js 中可通过 module.exports 导出以便校验。
 */
(function (global) {
  'use strict';

  function rotateLeft(lValue, iShiftBits) {
    return (lValue << iShiftBits) | (lValue >>> (32 - iShiftBits));
  }

  function addUnsigned(x, y) {
    var x4 = x & 0x40000000;
    var y4 = y & 0x40000000;
    var x8 = x & 0x80000000;
    var y8 = y & 0x80000000;
    var result = (x & 0x3FFFFFFF) + (y & 0x3FFFFFFF);
    if (x4 & y4) {
      return result ^ 0x80000000 ^ x8 ^ y8;
    }
    if (x4 | y4) {
      if (result & 0x40000000) {
        return result ^ 0xC0000000 ^ x8 ^ y8;
      }
      return result ^ 0x40000000 ^ x8 ^ y8;
    }
    return result ^ x8 ^ y8;
  }

  function f(x, y, z) { return (x & y) | (~x & z); }
  function g(x, y, z) { return (x & z) | (y & ~z); }
  function h(x, y, z) { return x ^ y ^ z; }
  function iFn(x, y, z) { return y ^ (x | ~z); }

  function ff(a, b, c, d, x, s, ac) {
    a = addUnsigned(a, addUnsigned(addUnsigned(f(b, c, d), x), ac));
    return addUnsigned(rotateLeft(a, s), b);
  }
  function gg(a, b, c, d, x, s, ac) {
    a = addUnsigned(a, addUnsigned(addUnsigned(g(b, c, d), x), ac));
    return addUnsigned(rotateLeft(a, s), b);
  }
  function hh(a, b, c, d, x, s, ac) {
    a = addUnsigned(a, addUnsigned(addUnsigned(h(b, c, d), x), ac));
    return addUnsigned(rotateLeft(a, s), b);
  }
  function ii(a, b, c, d, x, s, ac) {
    a = addUnsigned(a, addUnsigned(addUnsigned(iFn(b, c, d), x), ac));
    return addUnsigned(rotateLeft(a, s), b);
  }

  // 将字符串按 UTF-8 编码为“字节字符串”（每个 charCode 即一个字节）。
  // 注意：不做任何换行符替换，保证与 MaxCompute 对原始字符串逐字节一致。
  function utf8Encode(string) {
    var utftext = '';
    for (var n = 0; n < string.length; n++) {
      var c = string.charCodeAt(n);
      if (c < 128) {
        utftext += String.fromCharCode(c);
      } else if (c < 2048) {
        utftext += String.fromCharCode((c >> 6) | 192);
        utftext += String.fromCharCode((c & 63) | 128);
      } else if (c >= 0xD800 && c <= 0xDBFF && n + 1 < string.length) {
        // 代理对（surrogate pair）→ 4 字节 UTF-8
        var c2 = string.charCodeAt(++n);
        var cp = 0x10000 + ((c - 0xD800) * 0x400) + (c2 - 0xDC00);
        utftext += String.fromCharCode((cp >> 18) | 240);
        utftext += String.fromCharCode(((cp >> 12) & 63) | 128);
        utftext += String.fromCharCode(((cp >> 6) & 63) | 128);
        utftext += String.fromCharCode((cp & 63) | 128);
      } else {
        utftext += String.fromCharCode((c >> 12) | 224);
        utftext += String.fromCharCode(((c >> 6) & 63) | 128);
        utftext += String.fromCharCode((c & 63) | 128);
      }
    }
    return utftext;
  }

  function convertToWordArray(byteString) {
    var lWordCount;
    var lMessageLength = byteString.length;
    var lNumberOfWordsTemp1 = lMessageLength + 8;
    var lNumberOfWordsTemp2 = (lNumberOfWordsTemp1 - (lNumberOfWordsTemp1 % 64)) / 64;
    var lNumberOfWords = (lNumberOfWordsTemp2 + 1) * 16;
    var lWordArray = new Array(lNumberOfWords - 1);
    var lBytePosition = 0;
    var lByteCount = 0;
    while (lByteCount < lMessageLength) {
      lWordCount = (lByteCount - (lByteCount % 4)) / 4;
      lBytePosition = (lByteCount % 4) * 8;
      lWordArray[lWordCount] = (lWordArray[lWordCount] | 0) |
        (byteString.charCodeAt(lByteCount) << lBytePosition);
      lByteCount++;
    }
    lWordCount = (lByteCount - (lByteCount % 4)) / 4;
    lBytePosition = (lByteCount % 4) * 8;
    lWordArray[lWordCount] = lWordArray[lWordCount] | (0x80 << lBytePosition);
    lWordArray[lNumberOfWords - 2] = lMessageLength << 3;
    lWordArray[lNumberOfWords - 1] = lMessageLength >>> 29;
    return lWordArray;
  }

  function wordToHex(value) {
    var hex = '';
    for (var count = 0; count <= 3; count++) {
      var byte = (value >>> (count * 8)) & 255;
      var hexByte = '0' + byte.toString(16);
      hex += hexByte.substr(hexByte.length - 2, 2);
    }
    return hex;
  }

  function md5(string) {
    var x = convertToWordArray(utf8Encode(String(string)));
    var a = 0x67452301;
    var b = 0xEFCDAB89;
    var c = 0x98BADCFE;
    var d = 0x10325476;

    for (var k = 0; k < x.length; k += 16) {
      var AA = a, BB = b, CC = c, DD = d;

      a = ff(a, b, c, d, x[k + 0], 7, 0xD76AA478); d = ff(d, a, b, c, x[k + 1], 12, 0xE8C7B756);
      c = ff(c, d, a, b, x[k + 2], 17, 0x242070DB); b = ff(b, c, d, a, x[k + 3], 22, 0xC1BDCEEE);
      a = ff(a, b, c, d, x[k + 4], 7, 0xF57C0FAF); d = ff(d, a, b, c, x[k + 5], 12, 0x4787C62A);
      c = ff(c, d, a, b, x[k + 6], 17, 0xA8304613); b = ff(b, c, d, a, x[k + 7], 22, 0xFD469501);
      a = ff(a, b, c, d, x[k + 8], 7, 0x698098D8); d = ff(d, a, b, c, x[k + 9], 12, 0x8B44F7AF);
      c = ff(c, d, a, b, x[k + 10], 17, 0xFFFF5BB1); b = ff(b, c, d, a, x[k + 11], 22, 0x895CD7BE);
      a = ff(a, b, c, d, x[k + 12], 7, 0x6B901122); d = ff(d, a, b, c, x[k + 13], 12, 0xFD987193);
      c = ff(c, d, a, b, x[k + 14], 17, 0xA679438E); b = ff(b, c, d, a, x[k + 15], 22, 0x49B40821);

      a = gg(a, b, c, d, x[k + 1], 5, 0xF61E2562); d = gg(d, a, b, c, x[k + 6], 9, 0xC040B340);
      c = gg(c, d, a, b, x[k + 11], 14, 0x265E5A51); b = gg(b, c, d, a, x[k + 0], 20, 0xE9B6C7AA);
      a = gg(a, b, c, d, x[k + 5], 5, 0xD62F105D); d = gg(d, a, b, c, x[k + 10], 9, 0x02441453);
      c = gg(c, d, a, b, x[k + 15], 14, 0xD8A1E681); b = gg(b, c, d, a, x[k + 4], 20, 0xE7D3FBC8);
      a = gg(a, b, c, d, x[k + 9], 5, 0x21E1CDE6); d = gg(d, a, b, c, x[k + 14], 9, 0xC33707D6);
      c = gg(c, d, a, b, x[k + 3], 14, 0xF4D50D87); b = gg(b, c, d, a, x[k + 8], 20, 0x455A14ED);
      a = gg(a, b, c, d, x[k + 13], 5, 0xA9E3E905); d = gg(d, a, b, c, x[k + 2], 9, 0xFCEFA3F8);
      c = gg(c, d, a, b, x[k + 7], 14, 0x676F02D9); b = gg(b, c, d, a, x[k + 12], 20, 0x8D2A4C8A);

      a = hh(a, b, c, d, x[k + 5], 4, 0xFFFA3942); d = hh(d, a, b, c, x[k + 8], 11, 0x8771F681);
      c = hh(c, d, a, b, x[k + 11], 16, 0x6D9D6122); b = hh(b, c, d, a, x[k + 14], 23, 0xFDE5380C);
      a = hh(a, b, c, d, x[k + 1], 4, 0xA4BEEA44); d = hh(d, a, b, c, x[k + 4], 11, 0x4BDECFA9);
      c = hh(c, d, a, b, x[k + 7], 16, 0xF6BB4B60); b = hh(b, c, d, a, x[k + 10], 23, 0xBEBFBC70);
      a = hh(a, b, c, d, x[k + 13], 4, 0x289B7EC6); d = hh(d, a, b, c, x[k + 0], 11, 0xEAA127FA);
      c = hh(c, d, a, b, x[k + 3], 16, 0xD4EF3085); b = hh(b, c, d, a, x[k + 6], 23, 0x04881D05);
      a = hh(a, b, c, d, x[k + 9], 4, 0xD9D4D039); d = hh(d, a, b, c, x[k + 12], 11, 0xE6DB99E5);
      c = hh(c, d, a, b, x[k + 15], 16, 0x1FA27CF8); b = hh(b, c, d, a, x[k + 2], 23, 0xC4AC5665);

      a = ii(a, b, c, d, x[k + 0], 6, 0xF4292244); d = ii(d, a, b, c, x[k + 7], 10, 0x432AFF97);
      c = ii(c, d, a, b, x[k + 14], 15, 0xAB9423A7); b = ii(b, c, d, a, x[k + 5], 21, 0xFC93A039);
      a = ii(a, b, c, d, x[k + 12], 6, 0x655B59C3); d = ii(d, a, b, c, x[k + 3], 10, 0x8F0CCC92);
      c = ii(c, d, a, b, x[k + 10], 15, 0xFFEFF47D); b = ii(b, c, d, a, x[k + 1], 21, 0x85845DD1);
      a = ii(a, b, c, d, x[k + 8], 6, 0x6FA87E4F); d = ii(d, a, b, c, x[k + 15], 10, 0xFE2CE6E0);
      c = ii(c, d, a, b, x[k + 6], 15, 0xA3014314); b = ii(b, c, d, a, x[k + 13], 21, 0x4E0811A1);
      a = ii(a, b, c, d, x[k + 4], 6, 0xF7537E82); d = ii(d, a, b, c, x[k + 11], 10, 0xBD3AF235);
      c = ii(c, d, a, b, x[k + 2], 15, 0x2AD7D2BB); b = ii(b, c, d, a, x[k + 9], 21, 0xEB86D391);

      a = addUnsigned(a, AA); b = addUnsigned(b, BB);
      c = addUnsigned(c, CC); d = addUnsigned(d, DD);
    }

    return (wordToHex(a) + wordToHex(b) + wordToHex(c) + wordToHex(d)).toLowerCase();
  }

  global.md5 = md5;
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = md5;
  }
})(typeof window !== 'undefined' ? window : globalThis);
