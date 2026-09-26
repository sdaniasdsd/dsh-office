# docx-create

独立的结构化 DOCX 创建模块，主引擎固定为 `@docx4j/core-ts@0.1.5`。公开能力是 `inspect / execute / verify`，不承担全局编排，不产出第二套解析 IR。

## 已实现

- 结构化创建：三套样式母版、三级标题、格式化文字、表格、PNG/JPEG 图片、页眉页脚、页码、目录字段、显式分页。
- DOCX 母版填充：`{{name}}` 单行文字槽位；支持跨文字片段、正文、表格、页眉页脚；原件不修改。未修改的部件不重新构建。
- 产物具有 SHA-256、大小、媒体类型和 `.docx` 标签。新增段落有稳定 `w14:paraId`，对接现有双 IR 和按节点编辑。
- ZIP 展开预算、XML 安全检查、内部关系目标检查、输入/输出限制、Profile 安全钩子、可序列化错误、非阻塞遥测。

## 启动与测试

```powershell
npm ci
npm run typecheck
npm test
npm run fixtures:build
```

创建/填充仅需要 Node.js 20+ 和 npm 依赖，不需要 Python、Java、Pandoc 或 LibreOffice。跨模块集成测试使用本地 `docx-parse` 的 Python+lxml 环境；没有该环境时只运行 `test:contract`，不能声称完整集成已通过。`docx-render` 的 LibreOffice/Poppler 是渲染阶段依赖，不随创建模块强制安装。

## 调用方式

```ts
import { createDocxCreateModule } from './src/index';

const module = createDocxCreateModule({
  artifactStore: officeFilesAdapter, // 由宿主注入，不读取全局路径或配置
  config: { limits: { maxOutputBytes: 16 * 1024 * 1024 } },
});
const created = await module.handlers.execute({
  requestId: 'report-001', operation: 'execute',
  plan: { kind: 'create', document: {
    preset: 'report', header: '项目报告', pageNumbers: true,
    blocks: [
      { id: 'title', kind: 'paragraph', style: 'Title', runs: [{ text: '项目周报' }] },
      { id: 'summary', kind: 'paragraph', runs: [{ text: '本周完成模块联调。' }] },
    ],
  } },
});
const artifactRef = created.result.artifactRef;
// 由 Profile 接着调用 parser.handlers.execute({ artifactRef, ... })
// 或 renderer.handlers.execute({ artifactRef, ... })，本模块不自动调用它们。
await module.dispose();
```

母版填充使用 `plan: { kind: 'fillTemplate', templateRef, values: { title: '新标题' } }`。创建空白文档不要求虚构源文件。图片通过 `ArtifactRef` 读取，尺寸单位为毫米，支持 PNG/JPEG，不接受任意路径或网络地址来绕过 `office-files`。

## 与 Profile 的接口适配

`module.json` 和 `DOCX_CREATE_DEFINITION` 提供注册元数据；注册时绑定 `createDocxCreateModule` 返回的 `handlers`。宿主需要注入 `artifactStore`、可选 `safetyGuard` 和 `telemetry`。当前目录未提供真实 Profile registry / office-files 包，所以这里提供注册入口和适配端口，不声称已修改上层注册表。

`ArtifactStore.read(ref, maxBytes)` 和兄弟模块一致。**write 的创建输入是 `{ sources: ArtifactRef[], requestId, bytes, suggestedName }`**，因为无源文档创建没有 `source`。宿主必须创建新的、不可变的输出，且负责持久化/回滚/并发冲突；不得覆盖任何输入。编辑/渲染的 `{ source }` 端口不能未经适配直接当成创建端口。测试中的同时适配器只消费共有的 `bytes`。

公共类型由 `office-core / office-safety / office-test-kit` 导入，没有在本模块复制公共契约。为便于单包开发，tsconfig 暂时引用已有 `docx-parse/types/office-deps.d.ts` 类型桩；正式 Profile 接入后应改用真实公共包，移除该 include。`docx-parse/edit/render` 都只作开发和集成测试依赖，不是创建运行时依赖。

运行时输入 schema 与配置 schema 可从 `DOCX_CREATE_SCHEMAS` 获取，使用 JSON Schema Draft-07。`src/domain/docx-create.ts` 是创建计划的单一类型/schema来源。拒绝未知字段和有损 JSON 值；请求级配置只能收紧 Profile 的限制。

## 边界与尚未完成的事

1. 这是第一版生成内核，不是 Word 排版引擎。目录/页码字段返回 `pending-update`，视觉验证恒为 `pending`；`verify.ok=true, partial=true` 只表示已执行的局部结构检查通过，绝不代表截断/重叠/分页通过。没有完整 OOXML XSD 验证或图片解码检查。
2. 模板只支持单行纯文字槽位，不支持循环表格、条件表达式、复杂内容控件、槽位中的字段或修订；受保护或开启修订记录的母版拒绝填充。缺失值/多余值/不支持结构明确报错，不执行模板表达式。含外部关系、宏、嵌入程序、数字签名或主动取数域的文档基线拒绝，即便调用策略允许也不会绕过。文本填充只替换有变化的 XML 部件，关系、媒体和未引用的供应商附属部件也保留原始字节（ZIP 压缩容器本身会重建）。
3. 页面固定 A4、四边 25 mm，内容宽度 160 mm；不支持多节/横向页面/分栏/脚注生成。表格宽度和图片边界有结构限制，实际溢出交给渲染检查。三套原创样式不是商业级母版库。
4. `timeoutMs` 是引擎阶段的软期限：超时后不提交输出，但进程内同步工作不能被强杀。宿主文件读写期限由宿主负责；需要硬 CPU/内存隔离时，把同一个 `CreateEngine` 端口放入宿主 Worker/子进程。本版本不宣称提供沙箱。
5. 最终字体、TOC 更新和布局取决于部署端排版程序。当前开发环境未发现可执行 LibreOffice，已测试缺失依赖会明确失败；尚未完成真实逐页视觉验收。

## 源码位置

`contract.ts` 公共边界；`domain/` 计划及校验；`engine/adapter.ts` 创建引擎；`engine/template.ts` 保真文字填充；`mapper.ts` 归一化结果；`verifier.ts` 局部验证；`schema.ts` 导出机器可读 schema；`fixtures/build.ts` 生成可检查样例。
