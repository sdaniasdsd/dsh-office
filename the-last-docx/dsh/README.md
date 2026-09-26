# DSH DOCX 插件

单个 `@deepseek-ai/dsh-docx` bundle，包含检查、轻量解析、双 IR、复杂结构解析、创建、样式、编辑、渲染和版本交付。九个模块不拆成九个用户插件。使用 DSH 原生 MCP bridge 注册工具、审批及管理生命周期，不修改 DSH 主程序。

目标平台：Windows x64。DSH 已运行的 Node 作为子进程运行时，不要求另外安装 Node。私有 Python、LibreOffice、Poppler 随包携带；无需配置 PATH，不执行安装脚本，不在运行时下载依赖。Docling 深度模型和 rdocx 页坐标后端仍属于额外能力，不在本基础离线包中；原接口保留，不冒充可用。

通过 DSH 正常安装插件（不是手工配置引擎）：

```powershell
dsh plugin --profile web add 'D:\交付目录\deepseek-ai-dsh-docx-0.3.0.tgz'
```

重新加载 Profile 后使用六个 `mcp__docx__*` 工具。默认只导入 DSH 启动工作目录内的文件，产物放在该目录 `.dsh-docx` 中，不覆盖原件。不能默认给插件整个硬盘权限。宿主部署者可通过 `docx` 行的 config 设置 workspaceRoot、dataRoot、toolCallTimeoutMs、serverName；一般用户无需填写引擎路径。

带批注或脚注的文档要改版式时，用 `docx-edit` 的 `plan.edits[].kind='formatParagraph'`（原地格式化：styleId / outlineLevel / alignment / 间距 / 字体直接格式），它只改目标段落的 `w:pPr` 与该段各 run 的直接格式，`comments.xml`、`footnotes.xml` 等未被触碰的部件保持原始字节。`docx-create` 是按 plan 造新包，包里没有源文档的批注与脚注可带过去，因此不能用它给带批注的文档改版式。执行与校验是两步：`execute` 记录请求到 `expectedFormats`，`verify` 回读保存后的包逐项断言格式确实落上（`format.paragraph.N`），落不上的判失败而不是静默通过。

`docx-edit` 写的是 `w:pStyle` 这类**引用**，**定义**由 `docx-styles` 负责——它按成熟 OOXML 工具链的部件划分独立成模块（对照 docx4j 的 `StyleDefinitionsPart`、Apache POI 的 `XWPFStyles`、Open XML SDK 的 `StyleDefinitionsPart`），只拥有 `styles.xml`，不碰内容。用 `plan.styles` 写入具名样式定义，**已存在的其它样式一字节不改**（重写整个部件会静默丢掉文档原有的每一条样式）。`verify` 把引用分成三档：`defined` / 内置但本文档未定义（Word 会自行合成，属渲染器依赖，不是断链）/ 真正的悬空引用。**顺序：先 `docx-styles` 定义、后 `docx-edit` 引用**；中间态只含尚未被引用的样式，合法且无悬空。另外 `styles.order.valid` 会校验子元素顺序——`w:pPr` 里 `w:pBdr` 必须早于 `w:spacing`、`w:outlineLvl` 必须晚于它，顺序不对 Word 会提议"修复"文件。

`@deepseek-ai/dsh-mcp-client` 会随 DSH 插件自动安装；Cordis 复用 DSH 宿主的 `@deepseek-ai/cordis`，不私带第二套宿主核心。适配基于 DSH 0.1.0-rc.8 源码，并兼容当前 DSH Desktop 0.1.1-rc.2 中的桥接接口；不宣称兼容所有历史 DSH 版本。

`./core` 导出全部原模块命名空间，包括双 IR 修订记录、治理规则与摘要。MCP 工具只是交互入口，不替代这些代码接口。视觉 findings 由实际页面检查后提供，渲染成功不自动等于视觉 reviewed。

当前目录是构建产物；维护代码位于 the-last-docx 根工程的 dsh/ 与 scripts/build-dsh.mjs。安装不需要 npm install、tsx、TypeScript 或 Python pip。
