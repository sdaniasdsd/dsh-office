# The Last DOCX

DOCX 模块组的本地整合交付，版本 0.1.0。目标是保留已有模块内核，在外层补齐公共类型、文件管理、注册适配和 Agent 调用入口。不是新的 Word 排版程序，也不把结构检查当作视觉验收。

## 当前能用到什么程度

已经跑通真实的创建 → 检查 → 轻量/双 IR/复杂结构解析 → 节点桥接 → 编辑 → 交付清单 → 版本复核。原有八个模块都保留独立测试；双 IR 的修订记录、治理判定和 Agent 摘要导出保留。

真实逐页渲染尚未验收：当前配置未发现可启动的 LibreOffice、Poppler。复杂页坐标需要 rdocx，深度轻量解析需要 Docling；当前 Python 环境也未发现这两个包。不会自动安装程序、下载模型或用替身结果冒充真实结果。详见 [验收记录](VALIDATION.md)。

## 快速启动

需要 Node.js 20+。检查/双 IR/复杂结构以及整条交付链还需要 Python；本机实际验证版本见验收记录。

```powershell
Set-Location 'D:\开源团队作品\the-last-docx'
npm ci
npm run doctor
npm run typecheck
npm test
npm run test:modules
npm run demo
```

`demo` 将新的 DOCX、解析证据和版本清单写入 `examples/generated/<随机编号>/`，不覆盖旧文件。它验证结构和引用，不生成 PDF，也不承诺视觉合格。`build` 当前是类型检查；运行入口使用随项目安装的 tsx，不需要全局安装 TypeScript。

## 插件和项目怎么接

交付目录包含 `.codex-plugin/plugin.json`、`.mcp.json` 和真实 stdio MCP 服务。已经使用官方 SDK 客户端测试启动及工具调用，但**没有安装到应用、注册到市场或改动全局设置**。

配置允许导入的文档目录：

```powershell
npm run configure -- 'D:\你的文档工作目录'
node scripts/start.mjs
```

第二条是 MCP 进程入口，正常启动后等待协议输入，不是聊天式命令行。宿主读取 `.mcp.json` 启动它。复制或移动插件后重新运行 `configure`，因为配置中的启动脚本是绝对路径。该命令只更新当前插件的 `.mcp.json`。

也可直接使用 CLI：

```powershell
$env:THE_LAST_DOCX_WORKSPACE = 'D:\你的文档工作目录'
npm run cli -- list
npm run cli -- import 'D:\你的文档工作目录\报告.docx'
npm run cli -- call docx-parse execute 'D:\请求文件\parse.json'
```

请求 JSON 使用 import 返回的完整 `artifactRef`，并包含 `requestId`。不要直接把磁盘路径塞进伪造引用。默认产物放在工作目录下 `.docx-data`；`THE_LAST_DOCX_DATA` 可指定其他存储目录。`DOCX_PYTHON`、`DOCX_SOFFICE`、`DOCX_PDFTOPPM` 可指定程序路径。环境变量只在宿主层读取；能力模块仍只接受 Profile 注入配置。

MCP 公开六个工具：`docx_modules`（契约与配置）、`docx_doctor`（依赖检查）、`docx_import`（导入快照）、`docx_call`（原模块接口）、`docx_analyze`（显式组合解析）、`docx_read_artifact`（读取证据/页面图片）。大结果保存成 JSON 引用，不截断 IR。

已有 TypeScript 项目可以从 `src/index.ts` 导入：

```ts
import { DocxProfile, LocalArtifactFiles, docxParse } from './src/index';

const files = await LocalArtifactFiles.create('D:/项目/产物', ['D:/项目/输入']);
const profile = new DocxProfile({ files, pythonPath: 'python' });
// 可选：接入现有 registry 的最小 registerModule 端口。
await profile.registerWith(registry);
const source = await files.importFile('D:/项目/输入/报告.docx');
const analysis = await profile.analyze(source, 'analysis-001', 'structure', true);
const output = await profile.call('docx-artifact', 'execute', {
  requestId: 'delivery-001', artifactRef: source,
  delivery: { documentId: 'report', revision: 1, bridges: analysis.bridges },
});
// 原治理接口仍可用；是否记录意图、如何持久化由上层决定。
const ledger = docxParse.createRevisionLedger();
await profile.dispose();
```

此代码中的 `registry` 由集成项目提供；插件不会猜测你另一个项目的注册表实现。`registerWith` 应调用一次，重复注册和中途失败回滚由目标 registry 管理；整个 Profile 的生命周期由宿主统一关闭。

## 八个模块的分工

| 注册名 | 保留的职责 | 接口注意事项 |
| --- | --- | --- |
| docx-inspect | 格式、安全和包结构检查 | 安全基线在写入、渲染、交付前复用 |
| docx-easy-parse | 原轻量解析及可选 Docling | 原输出 moduleId 仍是 docx-parse；用注册名区分 |
| docx-parse | 双 IR、稳定锚点、修订记录与治理辅助 | 编辑的定位依据，不用版面节点代替 |
| docx-complex-parse | 复杂结构及可选页面观察 | structure 不要求页面坐标；layout 需要 rdocx |
| docx-create | 创建、样式、母版填充 | execute 使用 plan；空白创建不需要源文件 |
| docx-edit | 节点编辑、批注、修订策略、关系保护 | execute 使用 plan.edits 和双 IR 目标 |
| docx-render | DOCX → PDF → 逐页 PNG、缩略图、复核承接 | 需要 LibreOffice + Poppler；图片存在不代表视觉通过 |
| docx-artifact | 引用核验、预览关联、manifest 与版本链 | execute 接收 DOCX；verify 接收 manifest 引用 |

具体请求以各模块 `src/contract.ts`、schema 和 README 为准。公共输出继续保留 result/artifacts/warnings/verification，不暴露引擎对象。轻量 IR 和双 IR 是不同契约，不能按同一种结果解释。

## 必须知道的边界

- 编辑会生成新产物。编辑后必须重新解析/渲染，再关联到新版本；旧来源证据会被拒绝。
- 未提供预览或视觉复核，交付可成功但状态为 pending/partial。需要强制验收时，在可信 Profile 配置中为 docx-artifact 设置 `requirePreview`、`requireVisualReview` feature flags。
- 创建模块当前是 A4 单节、三套原创样式、纯文字槽位填充；不具备循环模板、多节分栏等完整 Word 功能。目录/页码字段仍需排版程序更新，详见 [创建模块边界](modules/docx-create/README.md)。
- 统一宿主默认导入/读取上限 64 MiB；各模块还有自己的更细预算。Node 进程内同步工作不是硬资源沙箱。具体限制见 [架构和安全说明](ARCHITECTURE.md)。
- 默认不强制安装 Docling、rdocx、LibreOffice，不要求 GPU。启用高开销能力前先评估实际文档和部署硬件；本版没有完成大批量性能基准，不能给出保证性的内存/速度数字。

依赖与再分发见 [THIRD_PARTY.md](THIRD_PARTY.md)。`assemble`、`provenance` 是开发维护命令，会访问原始模块目录；普通运行不依赖原始目录。不要用 assemble 覆盖尚未同步回源目录的模块改动。
