# MD5 转换 Excel（Univer 版）

一个完全离线运行的网页版「Excel」：把字符串填进 **A 列**，**B 列**自动生成对应的
**32 位小写 MD5**，可单个输入，也可批量粘贴 / 导入，与阿里云 MaxCompute 的
`MD5()` 函数结果一致。

表格由 [Univer](https://univer.ai) 渲染——它就是原 Luckysheet 团队推出的继任
前端电子表格框架（工具栏、公式栏、框选、右键菜单、撤销/重做、冻结表头一应俱全）。

## 功能

- A 列输入或粘贴原文，B 列实时自动生成 MD5（监听表格值变更事件计算）
- **从别的 Excel 直接粘贴**：在表格里点中 A 列单元格 `Ctrl+V`，走 Univer 原生
  粘贴，整列自动填入并逐行计算；也可以在页面任意位置 `Ctrl+V` 整表填入
- 预置 1048576 行（Excel 行数上限），粘贴/导入几万行也没问题：分块写入带进度，不会假死；粘贴时 `00123`、`1e5` 这类字符串会原样保留，不会被转成数字
- 数据量上限：单次粘贴/导入超过 100 万行、或导入文件超过 25 MB 时，会直接提示数据量太大、请分批次处理，不会硬算到卡死
- A 列为文本格式，`00123` 这类前导零不会丢
- 工具栏：导入 `.xlsx` / `.xls` / `.csv`（可选按钮、拖放文件到页面均可）、
  批量粘贴区、复制全部 MD5、复制 A+B（制表符分隔，可直接粘回 Excel）、
  导出 CSV、清空
- 选项：MD5 大写 ABC 切换、计算前去首尾空格、导入取第几列（第 1–8 列）、
  覆盖 / 追加、跳过表头行、跳过空行
- Univer 万一加载失败，会自动切换到备用文本转换区（逐行转换），页面不会白屏

## 离线 / 公司内网使用

**本包零外网依赖**：所有 JS / CSS 依赖都已下载好放在 `vendor/` 里，页面只用
相对路径加载，运行时不发出任何外网请求，断网可用。

用法（二选一）：

1. **单机打开**：把整个文件夹（或解压后的 zip）拷到目标电脑，双击
   `index.html` 即可使用。
2. **内网共用**：把整个文件夹放到内网服务器的静态目录即可，例如
   IIS 的站点目录、Nginx 的 `root` 目录，其他同事用浏览器打开
   `http://内网地址/文件夹名/index.html` 使用。

注意：请保持文件夹结构不变（`index.html`、`js/`、`css/`、`vendor/`
的相对位置不要改动、不要只拷 `index.html` 一个文件）。

## MD5 与 MaxCompute 的一致性

- 算法：标准 MD5（RFC 1321），纯 JavaScript 实现，见 `js/md5.js`（无第三方依赖）
- 编码：字符串按 **UTF-8** 编码后计算 —— 与 MaxCompute `MD5(string)` 的口径一致
- 输出：32 位**小写**十六进制；工具栏可切换为大写显示
- 空单元格不计算（B 列留空）；注意：空字符串本身的 MD5 是
  `d41d8cd98f00b204e9800998ecf8427e`，但表格中 A 列为空表示「无数据」，故不填
- 校验向量（可自行核对）：

  | 原文 | MD5 |
  | --- | --- |
  | `hello` | `5d41402abc4b2a76b9719d911017c592` |
  | `abc` | `900150983cd24fb0d6963f7d28e17f72` |
  | `你好` | `7eca689f0d3389d9dea66ae112e5cfd7` |

  `js/md5.js` 亦可在 Node.js 中 `require` 后与 `crypto` 模块对照校验。

## 目录结构

```
.
├── index.html                  页面入口（只引用本地相对路径文件）
├── css/
│   └── style.css               页面自有样式
├── js/
│   ├── md5.js                  MD5 实现（UTF-8，RFC 1321）
│   ├── univer-presets-loader.js  presets 分片加载器（拼接 3 片并全局执行，见下）
│   └── app.js                  Univer 初始化、自动计算、导入/导出/复制逻辑
└── vendor/                     全部第三方依赖（已下载到本地）
    ├── react.production.min.js
    ├── react-dom.production.min.js
    ├── rxjs.umd.min.js
    ├── univer-presets.chunk1.js   ┐ @univerjs/presets UMD 拆成的 3 片
    ├── univer-presets.chunk2.js   │ （替代原 univer-presets.umd.js，
    ├── univer-presets.chunk3.js   ┘ 原因与校验见下方说明）
    ├── univer-preset-sheets-core.umd.js
    ├── univer-preset-sheets-core.css
    ├── univer-design-locale-zh-CN.js
    ├── univer-preset-sheets-core-locale-zh-CN.js
    └── xlsx.full.min.js
```

## 依赖清单与许可证

| 依赖 | 版本 | 用途 | 许可证 |
| --- | --- | --- | --- |
| `@univerjs/presets`（UMD，拆为 3 片存储，见下方说明） | 0.10.1 | Univer 核心包（含 core / design / engine-render / engine-formula 等与 `createUniver`） | Apache-2.0 |
| `@univerjs/preset-sheets-core`（UMD） | 0.10.1 | 电子表格预设（表格 UI、粘贴、公式栏等） | Apache-2.0 |
| `@univerjs/preset-sheets-core` 中文语言包（UMD） | 0.10.1 | 界面中文 | Apache-2.0 |
| `@univerjs/design` 中文语言包（UMD） | 0.10.1 | 设计组件中文（语言包合并需要） | Apache-2.0 |
| `react` / `react-dom`（UMD） | 18.3.1 | Univer UMD 的运行时同伴依赖（全局 `React` / `ReactDOM`） | MIT |
| `rxjs`（UMD） | 7.8.1 | Univer UMD 的运行时同伴依赖（全局 `rxjs`，含 `rxjs.operators`） | Apache-2.0 |
| `xlsx`（SheetJS，UMD） | 0.18.5 | 解析导入的 `.xlsx` / `.xls` / `.csv` 文件 | Apache-2.0 |

说明：Univer 0.10.1 的 UMD 构建把 React、ReactDOM、RxJS 作为外部全局依赖，
因此这三者也一并下载到了本地；`@wendellhu/redi` 等其余依赖已包含在 Univer 的
UMD 包内部，无需另行下载。各依赖的许可证全文以其官方仓库 / npm 包内文本为准，
内网使用请自行确认符合所在单位的开源合规要求。

### 关于 presets 的拆分（univer-presets.chunk1–3.js）

`@univerjs/presets` 0.10.1 的 UMD 原文件 `univer-presets.umd.js` 约 7.5 MB，
超过 GitHub 接口单次推送的请求大小上限（实测约 3.5 MB 可推、7.5 MB 单文件
返回 HTTP 413），无法整文件推送到 GitHub。因此在本包中把它**按字符串拆成
3 片**存储，每片本身是一个合法的经典脚本，只把自己那段原文字符串 push 进
`window.__UNIVER_PRESETS_CHUNKS__`；页面加载时由 `js/univer-presets-loader.js`
在三片之后同步把三段拼接、并以全局作用域执行（等价于直接加载原文件），
UMD 照常注册 `window.UniverPresets` 等全局，**拆分没有修改原文件的任何内容**：

- 三片拼接结果与原文件的 SHA-256 完全一致：
  `0c98c819521cbf879b61448602ff3817de01d054adf0a770d0082e4cc4a5cf0a`
  （原文件 7,516,129 字节；已在拆分时逐字节校验通过）
- 若日后改用其他分发方式（如直接拷文件、内网文件服务器），也可以用原版
  单文件替换这 3 片 + 加载器：在 `index.html` 中把 4 行分片/加载器引用换回
  一行 `<script src="vendor/univer-presets.umd.js"></script>` 即可，其余代码
  无需改动。

## 浏览器要求

现代桌面浏览器即可（Chrome / Edge / Firefox 近期版本）。表格数据只在本地
浏览器内存中处理，不会上传到任何服务器；刷新页面后数据不保留，需要留存请
使用「导出 CSV」。
