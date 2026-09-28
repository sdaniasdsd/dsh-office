# DSH Office 插件

`@deepseek-ai/dsh-docx` bundle，包含 DOCX、PPTX、XLSX 和原生 PDF 模块：检查、轻量/双 IR/复杂解析、创建、样式、编辑、DOCX 渲染、版本交付、幻灯片提取与文本编辑、工作簿读取/编辑，以及 PDF.js 元数据与原生文字抽取。使用 DSH 原生 MCP bridge 注册工具、审批及管理生命周期，不修改 DSH 主程序。

**运行时是单独的包**：核心包 `@deepseek-ai/dsh-docx`（约 15 MB，纯 JS/npm）＋ 运行时包 `@deepseek-ai/dsh-docx-runtime`（Windows x64：私有 Python 3.13、LibreOffice、Poppler）。元包 `@deepseek-ai/dsh-docx-full` 只写依赖，一步装好两个。

DSH 已运行的 Node 作为子进程运行时，不要求另外安装 Node。运行时不配置 PATH、不执行安装脚本、不在运行时下载依赖；但它**不是硬依赖**——见下面的解析顺序与降级行为。Docling 深度模型和 rdocx 页坐标后端仍属于额外能力，不在运行时包里；原接口保留，不冒充可用。

## 运行时怎么找（解析顺序）

`dsh/runtime-config.mjs` 按以下顺序取运行时根目录，命中即止：

1. 插件配置里的 `runtimeRoot`（也可分别给 `pythonPath` / `sofficePath` / `pdftoppmPath`）；
2. 环境变量 `DSH_OFFICE_RUNTIME_ROOT`（也可用 `DOCX_PYTHON` / `DOCX_SOFFICE` / `DOCX_PDFTOPPM` 单独指定）；
3. **兄弟运行时包**：从核心包目录向上逐层找 `node_modules/@deepseek-ai/dsh-docx-runtime`，读它的 `runtime.json`；
4. 核心包自带的 `runtime/win32-x64/`（旧布局，兼容用）。

**缺运行时不再是启动错误**：只有核心包自己的 `lib/server.mjs` 缺失才算包不完整。运行时不全时，插件照常启动，只把存在的路径注入子进程，并在 stderr 明确列出缺了什么（同时写入 `DSH_DOCX_RUNTIME_MISSING`）；依赖运行时的能力（解析、PPTX、渲染等）会在调用时返回 `ENGINE_UNAVAILABLE`，`docx_doctor` 会给出逐项结论。创建、样式、编辑、交付清单等纯 Node 能力不受影响。

通过 DSH 正常安装插件（不是手工配置引擎）：

```powershell
dsh plugin --profile web add 'D:\交付目录\deepseek-ai-dsh-docx-full-0.11.0.tgz'
```

维护者从源码构建安装包：先在工程根目录运行 `npm run build:dsh`，再运行 `pwsh -NoProfile -File scripts/fetch-dsh-runtime.ps1` 获取并校验固定版本的离线运行时，最后运行 `npm run pack:dsh`。产物为 `dist/deepseek-ai-dsh-docx-0.11.0.tgz`。公开发布二进制前，先完成 `THIRD_PARTY.md` 要求的许可证与再分发义务审核。

重新加载 Profile 后使用 `mcp__docx__*` 工具访问 DOCX、PPTX、XLSX 和原生 PDF 能力。`pdf_call` 读取现有 PDF 的元数据及原生文字；它与 `docx-render` 的 DOCX→PDF 输出不同，不做 OCR、编辑或视觉合格声明。默认只导入 DSH 启动工作目录内的文件，产物放在该目录 `.dsh-docx` 中，不覆盖原件。不能默认给插件整个硬盘权限。宿主部署者可通过 `docx` 行的 config 设置 workspaceRoot、dataRoot、toolCallTimeoutMs、serverName；一般用户无需填写引擎路径。

带批注或脚注的文档要改版式时，用 `docx-edit` 的 `plan.edits[].kind='formatParagraph'`（原地格式化：styleId / outlineLevel / alignment / 间距 / 字体直接格式），它只改目标段落的 `w:pPr` 与该段各 run 的直接格式，`comments.xml`、`footnotes.xml` 等未被触碰的部件保持原始字节。`docx-create` 是按 plan 造新包，包里没有源文档的批注与脚注可带过去，因此不能用它给带批注的文档改版式。执行与校验是两步：`execute` 记录请求到 `expectedFormats`，`verify` 回读保存后的包逐项断言格式确实落上（`format.paragraph.N`），落不上的判失败而不是静默通过。

`docx-edit` 写的是 `w:pStyle` 这类**引用**，**定义**由 `docx-styles` 负责——它按成熟 OOXML 工具链的部件划分独立成模块（对照 docx4j 的 `StyleDefinitionsPart`、Apache POI 的 `XWPFStyles`、Open XML SDK 的 `StyleDefinitionsPart`），只拥有 `styles.xml`，不碰内容。用 `plan.styles` 写入具名样式定义，**已存在的其它样式一字节不改**（重写整个部件会静默丢掉文档原有的每一条样式）。`verify` 把引用分成三档：`defined` / 内置但本文档未定义（Word 会自行合成，属渲染器依赖，不是断链）/ 真正的悬空引用。**顺序：先 `docx-styles` 定义、后 `docx-edit` 引用**；中间态只含尚未被引用的样式，合法且无悬空。另外 `styles.order.valid` 会校验子元素顺序——`w:pPr` 里 `w:pBdr` 必须早于 `w:spacing`、`w:outlineLvl` 必须晚于它，顺序不对 Word 会提议"修复"文件。

## 宿主兼容契约

本插件**不交付** `@deepseek-ai/dsh-mcp-client`、`dsh-tools`、`dsh-llm`、`dsh-scope` 等 DSH 宿主包；它们是 Profile 所在 Desktop 进程的单例运行时。把这些包随插件装入 Profile 会让同一 ESM 进程同时解析两份 DSH 核心：旧桥接代码可能向新宿主的 `dsh-llm` 请求已删除的导出，或因 `dsh-scope` 的模块身份不同让多 Agent 工具注册冲突。

`@deepseek-ai/dsh-docx@0.11.0` 将这组包声明为 peer 依赖，兼容范围为 DSH `>=0.2.0-rc.1 <0.3.0`；加载器因此只使用 Desktop 自带的一份实现。`host-compatibility.json` 是可检查的产品契约，`npm run test:dsh-package-closure` 在打包前拒绝任何重新捆绑 DSH 核心包的产物。Cordis 同样复用宿主的 `@deepseek-ai/cordis`。不宣称兼容该范围之外的历史或未来 DSH 版本。

`./core` 导出模块命名空间，包括双 IR 修订记录、治理规则与摘要，以及 PPTX / XLSX 模块接口。PPTX run 替换和 XLSX 单元格写入都生成新 artifact。ExcelJS 不计算公式，且含图表、外链、透视表等可能无法保真的工作簿会被拒绝修改。MCP 工具只是交互入口，不替代这些代码接口。视觉 findings 由实际页面检查后提供，渲染成功不自动等于视觉 reviewed。

当前目录是构建产物；维护代码位于 **dsh-office 外壳**（即本仓库的 `dsh-office/`）的 `dsh/`、`src/` 与 `scripts/build-dsh.mjs`；各家族模块源码在仓库根的同级目录 `the-last-docx/`、`the-last-pdf/`、`the-last-pptx/`、`the-last-xlsx/`，由 `scripts/assemble.mjs` 装配进 `dsh-office/modules/`，并由 `scripts/provenance.mjs` 逐文件校验一致。构建会内置锁定的 python-pptx 与 Pillow / XlsxWriter / typing-extensions 运行轮子。安装不需要 npm install、tsx、TypeScript 或 Python pip。
