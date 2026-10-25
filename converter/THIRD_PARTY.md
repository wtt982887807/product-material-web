# 本地 Word 转换组件

转换使用真实 LibreOffice WebAssembly 引擎处理文档结构与版式，不经过纯文本或 HTML 重建。输入为 DOC、RTF，以及原有修订需要兼容处理的 DOCX。Word 2003 XML、任意 XML、PDF 和图片不在本轮支持范围内。

## 来源与版本

- 桥接库：allotropia 发布的 `zetajs` **1.2.0**，MIT 许可。
- 官方项目：<https://github.com/allotropia/zetajs>
- 获取地址：<https://registry.npmjs.org/zetajs/-/zetajs-1.2.0.tgz>
- 官方 npm integrity：`sha512-lAY5Q7qvvGYlmsMyuAByL/re9uIbDmJhHLXvF2FNhOdnv8V1/jUTS5EwIB6dAohFtzrd9HhP//5eoAvNGc9hPA==`
- 本地保留未修改的 `zeta.js`、`zetaHelper.js` 和 `LICENSE`，总计 53,251 字节。
- 引擎：官方示例推荐的 <https://cdn.zetaoffice.net/zetaoffice_latest/>。这是**动态更新地址**，不是固定版本地址。引擎未复制进仓库。
- LibreOffice 引擎及其第三方组件沿用各自上游许可证；引擎源代码入口见 [ZetaJS 官方说明](https://github.com/allotropia/zetajs#using-with-an-own-build)。ZetaJS 的 MIT 许可不替代引擎许可。

桥接文件 SHA-256：

| 文件 | SHA-256 |
| --- | --- |
| `zeta.js` | `b70ee7022908822c512535bbf43c4e44c297acc364e7542abba0d3071f5f9bfd` |
| `zetaHelper.js` | `c07abdbc0d210ac803823d002520fac20d34ff18f20cbd5a78d203b5c9901a96` |
| `LICENSE` | `a8c9b7d037ed112c3b9f85d5a1122aedd218947b21cf3f2c4200d13d7fe1a7d9` |

## 浏览器与部署

调用方负责先建立跨域隔离环境。需 HTTPS 或 localhost，顶层页面和同源转换 iframe 均须满足 COOP `same-origin` 与 COEP `require-corp`（或浏览器支持的 `credentialless`）。需要 SharedArrayBuffer、WebAssembly 和 Web Workers。GitHub Pages 可使用同源 Service Worker 处理隔离，首次重载应在用户选择文件之前完成。

转换 iframe 路径为 `converter/frame.html`。其中 `frame.js` 按需加载官方引擎，`office-thread.js` 在 LibreOffice 自己的 worker 内执行。官方桥接使用 blob worker 和 data URL 模块；如果另设 CSP，须兼容这些资源以及官方 CDN。不要允许文档去使用外部代理。

2026-10-25 对官方 CDN 进行只读 HEAD 检查：WASM 与 data 使用 Brotli 压缩，四个引擎文件的传输长度合计 53,243,567 字节，约 50.8 MiB。其 `soffice.js` 默认申请 1 GiB 共享内存。以上是当日动态版本观测值，不是永久固定规格。优先使用最新版桌面 Chrome 或 Edge；内存不足、缓存受限或网络中断时会明确失败。字体缺失与 LibreOffice / Word 布局差异仍需对实际物料复核。

## 文件处理边界

`convertLegacyWord({name, bytes}, {onProgress})` 返回 DOCX 的 `Uint8Array`。进度回调接收 `{stage, message}`。调用函数才加载引擎，普通 DOCX 流程无需加载引擎。

输入文件先复制到内存，再写入 WASM 虚拟 `/tmp/`。临时名称由内部编号及白名单后缀生成，不使用上传文件名作为路径。原始文件字节不被修改，源文件不发送到 CDN 或转换服务器。

加载选项为 `Hidden=true`、`ReadOnly=false`（可编辑的内存副本）、`Silent=true`、`MacroExecutionMode=NEVER_EXECUTE`、`UpdateDocMode=NO_UPDATE`，并拒绝密码或其他交互要求。输出前执行 `.uno:AcceptAllTrackedChanges`，确认 `getRedlines().getCount()` 为零，再以 `Office Open XML Text` 保存。无法确认旧修订全部接受时停止输出。

每次转换后关闭文档并移除输入、输出临时文件。单一队列顺序执行，正常情况下复用引擎；错误或 240 秒超时会移除 iframe 以释放其线程，下次操作创建新实例。源文件上限为 60 MB。

## 二进制 DOC 底纹兼容

2026-10-25 的真实浏览器验证发现，LibreOffice 在读取部分 DOC 时忽略 `sprmTDefTableShdRaw`，蓝色单元格在加载时就成为透明。两个 DOCX 输出过滤器均不能修复这一导入缺口。`src/doc-shading.js` 仅在交给引擎的内存副本中，将明确颜色的 Raw 底纹属性映射到此导入器支持的对应属性：`D670→D612`、`D671→D616`、`D672→D60C`。各操作数的原始颜色、长度、正文和文件偏移保持不变。这是针对 LibreOffice 的兼容桥，生成的中间 DOC 不提供给用户，也不声称是通用 DOC 规范化。

解析顺序为 CFB 目录及 FAT/MiniFAT、FIB、PlcBtePapx、PapxFkp、Prl，并跟随 Data 中的 PrcData 指针。只访问根目录的 WordDocument/0Table/1Table/Data，不扫描正文、媒体或嵌入对象中的相似字节。对长度、版本、循环、扇区重叠和属性边界进行检查。ShdNil 的样式继承、复杂底纹图案或检测到的新旧底纹冲突不能仅靠 opcode 替换可靠表达，因此停止并提示先由 Word 另存为 DOCX。

依据：[Microsoft Table Properties](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/b39a6648-501c-4361-8366-4f042f579469)、[PapxFkp](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/34aaeaf3-9578-41af-a3f5-c12f6f66bf1b)、[PapxInFkp](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/580510b8-df7a-467e-a51c-0d71eb15c7cd)、[PrcData](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/473fd992-c824-4655-8880-3186bd432f80)、[LibreOffice WW8 importer](https://github.com/LibreOffice/core/blob/master/sw/source/filter/ww8/ww8par2.cxx)。实现根据公开文件格式规范编写，未复制 LibreOffice C++ 源码。

私有 DOC 样本经浏览器转换并由原生 Word 打开/渲染后，三种蓝色单元格的数量与 Word 原生另存为结果一致（14、36、14），六页版面及修订归零通过复核。私有样本、诊断页和中间文件均不进入仓库。此验证不代表所有 Word/LibreOffice 排版差异均已消除。

## 官方 API 依据

- [官方转换示例](https://github.com/allotropia/zetajs/tree/main/examples/convertpdf)
- [LibreOffice 转换过滤器](https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html)
- [MacroExecMode](https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1document_1_1MacroExecMode.html)
- [UpdateDocMode](https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1document_1_1UpdateDocMode.html)
- [Emscripten pthreads 与隔离要求](https://emscripten.org/docs/porting/pthreads.html)

## 测试范围

`tests/legacy-word-converter.test.mjs` 检查惰性初始化、串行复用、输入快照、超时销毁与重试、输出 DOCX 结构、UNO 安全选项及修订归零门槛。它使用边界替身，不替代真实浏览器的引擎启动、样本文档转换和 Office 排版验证。

`tests/doc-shading.test.mjs` 使用完全合成的 CFB/PAPX 文档检查三段底纹映射、输入不变、伪造字节不被扫描、扩展属性、坏链及不支持的底纹语义。测试不包含产品或公司介绍原文。
