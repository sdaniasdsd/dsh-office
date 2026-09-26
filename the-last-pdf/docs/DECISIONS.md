# 技术路线记录

状态：用户已选全部推荐项（D1=A、D2=A、D3=A、D4=A）。决定在 2026-09-26 确认，代码实现和限制见上层 README。版本按 NPM lock 和 `requirements-docling.txt` 固定；模型和二进制以 inventory 清单记录哈希。

| 决策 | 用户选择 | 本区适配 |
| --- | --- | --- |
| D1 原生解析与渲染 | A | PDF.js 6.3.289、`@napi-rs/canvas` 1.0.9、qpdf 12.4.1 |
| D2 OCR、布局和表格 | A | Docling 2.130.0 按需执行；RapidOCR 中文模型，Docling 内显式 PDFium 后端 |
| D3 创建 | A | PDFKit 0.20.2，以受约束的结构化输入创建新 PDF |
| D4 编辑 | A | pdf-lib 1.17.1，限页面、表单、注释及盖章；不作任意正文重写或删隐 |

选择采纳的是整体主路线。Docling PDF backend 是这个已选路线内部的适配细节：此 Windows 工作区里 Docling-Parse 的 glyph 资源定位无法处理当前中文安装路径，故代码显式传入官方支持的 `PyPdfiumDocumentBackend`，并将 backend 版本写进每个模型证据，未做隐式失败降级。

上游项目参考及边界见 [COMMUNITY-SOURCES.md](COMMUNITY-SOURCES.md)。尤其 qpdf 结构检查不等于视觉验收、OCR 结果不等于可写源定位、PDF 页面坐标不等于段落编辑锚点。真实业务语料中的字符错误率、复杂表格完整率、跨页阅读顺序仍需另建标注样本测量。
