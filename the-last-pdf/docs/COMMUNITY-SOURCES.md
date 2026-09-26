# 开源社区技术依据

查阅日期：2026-09-26。技术判断来自项目官方仓库/API/文档；选用包通过 npm lock 和 Python requirement 固定版本。第三方上游是活动项目，后续升级需重新跑回归。

| 项目与一手资料 | 后续适配器具体采用/核验的细节 | 不能由它直接推出的能力 |
| --- | --- | --- |
| [PDF.js API](https://mozilla.github.io/pdf.js/api/draft/api.js.html)、[Node 渲染例子](https://github.com/mozilla/pdf.js/blob/master/examples/node/pdf2png/pdf2png.mjs) | PDF.js 6.3.289：独立 TypedArray 传给 getDocument，getTextContent/getViewport/render，标准字体/CMap/WASM，按文档任务 cancel/cleanup；与 @napi-rs/canvas 1.0.9 跑通 Windows Node | text item 不天然等于段落或稳定可写锚点 |
| [qpdf 检查](https://qpdf.readthedocs.io/en/stable/cli.html#option-check)、[12.4.1 release](https://github.com/qpdf/qpdf/releases/tag/v12.4.1) | qpdf 12.4.1 `--check` 由参数化子进程运行；退出状态映射为 pass/warn/fail；安装包 SHA-256 经运行脚本固定 | qpdf 不渲染；结构检查通过不保证显示或语义正确 |
| [pdfplumber README](https://github.com/jsvine/pdfplumber) | chars/lines/rects、裁切坐标、表格边缘与交点参数；针对机器生成 PDF 使用 | 不自带扫描 OCR，表格推断需要对照样本 |
| [pypdfium2 文档](https://pypdfium2.readthedocs.io/en/stable/readme.html) | PdfDocument、页面渲染、原生资源 close；Python 桥接清理；随二进制交付依赖许可 | Python 包许可不替代整个 PDFium 构建的依赖许可 |
| [Docling advanced options](https://docling-project.github.io/docling/usage/advanced_options/)、[custom backend](https://docling-project.github.io/docling/_generated/examples/custom_convert/) | Docling 2.130.0 离线 `artifacts_path`、OCR/table 开关；`PyPdfiumDocumentBackend` 作为显式 backend；RapidOCR chinese ONNX；子进程回传单一 JSON | 默认原生 DoclingParse 在本区中文 Windows 路径无法定位字体资源；layout/table 模型预测也不保证表格一定分类正确 |
| [PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR) | 按选定流水线及模型提取识别、版面、结构结果；封装 Python 运行环境 | 中文效果必须在目标语料上实测，不能直接继承宣传基准 |
| [PDFKit](https://pdfkit.org/) | 字体嵌入、文字布局、分页、表格与图像生成；产物流写入新文件 | 不是现有 PDF 的通用编辑器 |
| [Playwright PDF API](https://playwright.dev/docs/api/class-page#page-pdf) | Chromium 打印 CSS、page size、字体就绪、资源加载约束 | 页面截图与可检索 PDF 不是同一交付 |
| [pdf-lib README](https://github.com/Hopding/pdf-lib#limitations) | 页面复制、表单、字体、appearance 更新；用成熟 API 实现明确支持的操作 | 不提供普通正文删除/替换 API；ignoreEncryption 不会解密 |
| [pypdf 文本提取说明](https://pypdf.readthedocs.io/en/stable/user/extract-text.html)、[仓库](https://github.com/py-pdf/pypdf) | 页面操作、表单、元数据；原生文字和扫描图片分别处理 | 不是 OCR；PDF 文字定位不等于可靠阅读顺序 |
| [PyMuPDF Page API](https://pymupdf.readthedocs.io/en/latest/page.html)、[仓库](https://github.com/pymupdf/PyMuPDF) | 原生字典提取、页图、批注与 redact；明确坐标及保存规则 | 画白色矩形不等于删除底层内容；许可路线需选择 |
| [Apache PDFBox](https://pdfbox.apache.org/)、[FAQ](https://pdfbox.apache.org/3.0/faq.html) | 内容流、字体、表单等的成熟实现参照；Java 作为可选后端 | 无法假定字形流自带段落、表格和可重排文档模型 |
| [Poppler](https://poppler.freedesktop.org/) | 可继承现有 DOCX 的 pdftoppm 工具经验，作为渲染对照 | Windows 构建与分发是额外选择，许可证不是 MIT |

## 许可证记录

官方仓库声明：PDF.js / qpdf / PaddleOCR / PDFBox 为 Apache-2.0 路线；pdf-lib / PDFKit / pdfplumber 为 MIT 路线；pypdf 为 BSD-3-Clause。Docling 的代码许可和所选模型/依赖许可应分别记录。pypdfium2 提供 Apache-2.0/BSD-3-Clause，PDFium 和其依赖许可需随构建核验。PyMuPDF 提供 AGPL 与商业许可选择，Poppler 为 GPL。以上用于技术选项比较，不据此宣称整个组合天然满足某种分发模式。

## 当前实现中已经采用的设计

本地 DOCX 源码提供的是模块边界与来源验证实践。PDF 语法、页面对象与字形处理委托给 PDF.js/qpdf/Docling/PDFKit/pdf-lib 上游，而非本区重写 PDF 字节协议。

主要锁定版本：PDF.js 6.3.289、qpdf 12.4.1、PDFKit 0.20.2、pdf-lib 1.17.1、@napi-rs/canvas 1.0.9、Docling 2.130.0、docling-core 2.99.0、docling-parse 7.22.0、pypdfium2 5.8.0、RapidOCR 3.9.2、onnxruntime 1.20.1。
