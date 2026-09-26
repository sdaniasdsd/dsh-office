# PDF 分区验证记录

验证日期：2026-09-26。环境：Windows / PowerShell、Node v24.19.0、npm 11.17.0、TypeScript、Vitest v2.1.9。qpdf 12.4.1；PDF.js 6.3.289；PDFKit 0.20.2；pdf-lib 1.17.1；@napi-rs/canvas 1.0.9。

| 检查 | 结果 |
| --- | --- |
| `npm run typecheck` | 通过，退出码 0 |
| `npm test` | 3 个文件，64 项通过，1 项可选 Docling integration skipped |
| `npm run demo` | 真实中文双页 PDF/表单创建、编辑、解析、渲染、验证、manifest 示例；通过 |
| PDF.js 图像 vs Poppler | 双页 A4 样张已独立渲染检查；文字、页码、表格、盖章完整，没有缺字/裁切/错位 |
| AcroForm 两端核对 | 文本值/复选状态、page `/Widget` 归属、有效 `/V`、appearance `/AP /N` 经重开检查；错误和脱离字段树的 widget 被拒绝 |
| `npm run demo:docling` | 通过；Docling 识别主样张 2 张表/25 格，cell geometry/source pointers 与 raw observation 一致；纯栅格页原生文字 0 项，OCR 恢复预期中文标题 |
| Docling demo manifest 检查 | 主 PDF r1→r2 与独立 scan PDF r1 两条链均重读完整文件/预览/证据并通过；视觉 review 都如预期保留 pending |

运行时的中文安装路径使默认 Docling-Parse backend 在加载 `glyphs/standard/additional.dat` 时失败。已参考 Docling 支持的 backend 配置显式选择 `PyPdfiumDocumentBackend`，并在 JSON observation 和 evidence engine 字段标明 backend。复杂样张返回两张 TableItem（5×3、5×2）；25 个单元格的 text、span、role、bbox 和 source pointers 进入 IR，自动化断言对照 raw observation 与 typed IR。不同的小型 fixture 可能由 Docling 分类成普通文字，不能宣称所有表格都稳定检出。

Docling 本地基线：docling 2.130.0、docling-core 2.99.0、docling-parse 7.22.0、RapidOCR 3.9.2、onnxruntime 1.20.1、pypdfium2 5.8.0、torch 2.14.0、transformers 5.13.0。`runtime/inventory.json` 记录运行包、模型文件尺寸及 SHA-256。仓库忽略模型本身，其他机器按锁定依赖重新下载。

真实测试目前主要使用本机生成、可重复的少页 fixtures，且 OCR 检查对象是标准字体/大小清晰排版。此数据只能证明接口的输入输出与此种简单 fixture 行为，不能推出法律合同、扫描件、扭曲图片、复杂跨页表格或大量页面准确率。`docs/ACCEPTANCE.md` 指明下一轮真实样本矩阵与待测项。

## 命令

```powershell
npm ci
npm run typecheck
npm test
npm run demo
./scripts/setup-runtime.ps1 -WithDocling # 首次配置时下载模型，需要网络
npm run demo:docling
```
