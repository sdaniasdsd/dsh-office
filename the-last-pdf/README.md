# PDF 分区

这是从 `../docx分区` 的实战迁出的 PDF 工作区。当前已按确认的 D1-A / D2-A / D3-A / D4-A 接通可运行实现：PDF.js + qpdf、按需 Docling、PDFKit，以及 pdf-lib 的受限编辑和不可变交付。

## 运行

需要 Windows 与 Node.js >=22.13。基础引擎、真实 PDF 示例与测试：

```powershell
cd 'D:\开源团队作品\pdf分区'
npm ci
npm run typecheck
npm test
./scripts/setup-runtime.ps1
npm run demo
```

Docling 是可选且首次配置较重，需要 Python 3.11；显式安装运行时和模型后执行完整的结构/OCR 样例：

```powershell
./scripts/setup-runtime.ps1 -WithDocling
npm run demo:docling
```

qpdf、Python 和模型路径可经 `ModuleOptions.backend` 配置，也能用 `QPDF_PATH`、`DOCLING_PYTHON`、`PDF_FONT_PATH` 环境变量覆盖示例默认值。Docling 执行设为离线模式。模型文件、qpdf 包、Python 虚拟环境和样例输出不纳入源代码提交；`scripts/inventory.py` 记录本机依赖清单及模型文件 SHA-256。

## 模块与真实边界

| 模块 | 已接能力 | 当前边界 |
| --- | --- | --- |
| pdf-inspect | PDF.js 元数据、qpdf 结构检查、PDF 对象模型主动内容扫描 | `qpdf --check` 不说明页面视觉正确 |
| pdf-easy-parse | 按页原生文字 | 不推断 OCR、阅读顺序、段落或格式样式 |
| pdf-parse | PDF.js 文字片段、页几何、裁切后坐标、源摘要和双 IR | 文本片段不天然等于段落；没有可写的正文锚点 |
| pdf-complex-parse | Docling 布局、OCR、表格；本机预置模型并返回上游原始观察档案 | 依赖 Python/模型；类型化投影不包含 Docling 全部字段，原始结果可单独检查 |
| pdf-create | PDFKit 结构化段落、标题、分页表格、字体嵌入 | 非 ASCII 必须传入字体；不声称 DOCX 版式兼容 |
| pdf-edit | 页面挑选/旋转/追加、盖章、Text 批注、AcroForm 字段 | 不支持任意正文替换、删隐、XFA、签名、带控件的页面重排 |
| pdf-render | PDF.js 页面 PNG | 需要显式登记逐页目视复核，初始状态始终 pending |
| pdf-artifact | 摘要寻址产物与追加式 manifest 版本链 | 视觉状态来自绑定源摘要的审阅记录 |
| pdf-contracts / pdf-engines | 共享契约、预算、来源检查与默认引擎 | 接宿主时仍需把本地存储替换/连接到实际 office-files |

示例的真实 PDF 和页面图写入 `output/pdf/`。它会创建中文双页样张、加章和批注，检查可填写表单及生成来源绑定的 JSON 证据与 manifest。`npm run demo:docling` 还创建没有文字层的栅格页，以检查 OCR 和表格观察。示例不会声称未经人眼检查的内容已通过视觉审阅。

## 安全和技术边界

PDF.js/Node/C++ PDFium 等组件同步处理可能超过通用 AbortSignal 的即时抢占能力。这里由子进程执行 qpdf/Docling 的路径可杀进程；同线程解析/渲染靠预算约束和 PDF.js cancel，不是操作系统级资源隔离。服务端部署应给 Worker/进程容器增加内存及 CPU 限制。

加密文件只按空密码尝试读取；不询问或持久化口令。签名文档、XFA 和需要保留字段层级的裁页操作明确停止。输出通过新摘要引用，不覆盖源文档。

详细的迁移依据、开源上游、依赖版本、验证结果及尚未覆盖的 PDF 类别： [迁移说明](docs/MIGRATION.md)、[开源社区记录](docs/COMMUNITY-SOURCES.md)、[验证记录](VALIDATION.md)、[验收语料计划](docs/ACCEPTANCE.md)。
