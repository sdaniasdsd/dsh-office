# DOCX 实战迁移记录

检查日期：2026-09-26。源目录 `D:/开源团队作品/docx分区`，目标目录 `D:/开源团队作品/pdf分区` 最初为空。两目录没有发现适用的 AGENTS.md，分区根目录不属于 Git checkout。本轮只在 PDF 目录新增文件。

## 框架依据

| 实际读取的 DOCX 源码/文档 | 可复用实践 | PDF 落点 |
| --- | --- | --- |
| docx-inspect/src/contract.ts | JSON 信封、三接口、分类错误、Profile 配置 | pdf-contracts/contract、schema、runtime |
| docx-create/package.json、README.md | TS/ESM/Vitest/Zod、无源创建、不变输出 | 本分区工具链；create.execute 可无 artifactRef |
| docx-parse/README.md | 双 IR、三向来源索引 | pdf-dual-ir/v1 与索引校验 |
| docx-complex-parse/README.md | 原始观察与归一化分离、coverage、置信度来源 | evidence.method、nullable confidence、coverage |
| docx-render/src/index.ts、contract.ts | 注入引擎、预览来源检查、视觉 pending | pdf-render 结果契约与来源 SHA 绑定 |
| docx-artifact/src/contract.ts | 版本化 manifest、不可变提交、CAS 与幂等 | 写入端口及交付路线规范；未复制实现 |
| docx-edit/README.md | 从来源定位编辑，写入后重开验证 | PDF 的字段/批注对象与页面计划，待引擎接入 |
| docx-styles/README.md | 格式专属部件应有明确所有者 | PDF 不伪造 styles.xml；主题归生成计划 |
| docx-parse/types/office-deps.d.ts | 公共类型只做边界声明 | 仅保留 ArtifactRef/Warning/报告形状兼容桩 |

没有逐文件替换 docx 字符串来宣称迁移完成。DOCX 的 ZIP、OPC relationship、paraId、XML 路径、Word 修订与 styles.xml 均不能直接解释 PDF。

## 保留和调整

保留模块公开三接口、工厂、register、配置注入、独立注册描述。根目录 workspace 统一安装和测试；各模块仍具有独立包名和工厂。公共实现仅在 pdf-contracts 内共享，避免八份相同的信封与超时逻辑。模块业务域与引擎文件待实现时再拆分。

PDF 页码固定为一基索引。页面宽高、标准化 viewport 和 bbox 采用 CropBox + 页面旋转之后的左上角坐标，单位 pt；页面 CropBox 在 IR 中表示这个已投影视口 `[0,0,width,height]`。PDF.js 视口矩阵负责原始裁切偏移、页面旋转及 `UserUnit`。物理观察可保留 objectRef，获取不到时为 null。区域 bbox 表示定位证据，不表示内容流的可写选择器。

DOCX 的稳定 paraId 不迁移为 PDF 永久 ID。PDF 对象号、文字片段序号和布局区域均只在当前文件摘要下使用；重写、压缩或 OCR 后应重建证据。`source.id + source.sha256` 是边界。文字指纹只检索候选，重复字符串不自动选择第一项。一个语义段落可以关联多个片段，多个语义单元也可能共享物理区域，因此索引是多对多。

保留 observed/partial/unavailable，不把空文字等同于空白页；未知加密、脚本、附件状态必须为 null。OCR 输出明确标注 method，不虚构模型未提供的置信度。

## 共享宿主的待接入点

本轮源单包类型桩的 `FormatKind`/`container` 只覆盖 docx/docm/dotx/dotm、zip/ole/rtf 等值，`FormatIR` 是 OPC 包模型。不能复用为 PDF 对象/页面模型。正式接入时，宿主需要以格式判别联合承载 `PdfFormatProfile` 和 PDF IR，并检查所有 switch 分支；或将它们作为独立格式载荷。此处只建立局部 PDF 契约，未修改任何真实共享包。

原 `SafetyPolicy` 的 allowMacros 不能表示 PDF JavaScript，allowExternalLinks 也不等同于 Launch action。当前局部策略覆盖加密、脚本、附件和外链；动作枚举和签名修改策略须由实际 detector/编辑路线补齐。宿主接入时做显式映射，不能使用类型断言跳过。

ArtifactRef 与 Warning 保持逐字段形状兼容；模块输入新增的 PDF 业务 payload 是草案，不宣布与任何现有宿主 RPC 自动兼容。module.json 当前声明 framework-only；只有实际引擎、输入 schema 和宿主注入都验收后才能改为生产能力。

## 实施落点与未覆盖内容

共享模块保留 `inspect / execute / verify`、工厂、`register()`、配置覆盖和 JSON 边界。社区适配器集中到 `pdf-engines/`：PDF.js 把页面 text item 与 geometry 作为只读 observations；Docling 由限输出/可杀进程的 Python 桥接器启动；PDFKit 直接从结构计划创建新文件；pdf-lib 只应用计划中的对象操作。

来源摘要和一基页码绑定 IR 与预览；OCR 区域没有 object reference，也不能凭 bbox 变成正文编辑器。Docling 原始观察随结构 IR 一并保存为独立 source-bound JSON 证据；table cells project text、span、role，以及几何相容时的 cell bbox/pointer。其余上游字段继续保留在 raw observation 中；置信度仍为 null。已生成样张中，两张表（5×3 与 5×2）的 25 个单元格全部映射且带几何；这不能证明所有复杂/跨页表都会被识别。

输出写到本地测试存储的内容摘要文件，再以 hardlink 创建不覆盖的目标；版本 manifest 递增提交并校验前版摘要。正式 Profile 仍需把 `ArtifactReader/Writer` 接到宿主 office-files 的安全目录、CAS、并发及清理政策。

真实页面图片已通过 PDF.js 与独立 Poppler 对照并检查；这属于本轮验收样张复核，不代表每个用户上传 PDF 的审阅状态已自动完成。合同/扫描件语料、AcroForm 第三方模板兼容、PDF/A、签名增量更新和长期吞吐/内存压力尚需专项样本与部署环境实测。

通用 runtime timeout 会向引擎广播 AbortSignal 并终止支持它的 qpdf/Docling 子进程。PDF.js 和 PDFKit 的单线程 CPU 工作不能被 JS AbortSignal 立刻抢占；服务端应给这些操作分配受资源限额的 Worker/进程。
