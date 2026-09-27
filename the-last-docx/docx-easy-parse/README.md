# docx-parse

`dsh-office-profile` 的 DOCX 结构解析模块（`profileGroup: DOCX`，版本 `0.2.0`）。

把一个 `.docx` / `.docm` 产物解析成**结构化的归一化 IR**：段落、标题层级、表格、
样式、OPC 关系、批注、脚注、元数据。只读、确定性、不产出新文件。

---

## 它是什么，不是什么

**是**：一个只读的结构解析器。输入一个 artifact 引用，输出可 JSON 往返的结构化结果，
并对结果给出稳定分类的错误与告警。

**不是**：

- 不是渲染器。不解析视觉排版，不产出 HTML/PDF 像素结果。
- 不是编辑器。不写回文档，因此**不物化 DOM**——这是相比 Aspose.Words / Apache POI /
  Open XML SDK 这类编辑库的关键差异，也是它能对恶意文档保持低攻击面的原因。
- 不是格式转换器。OLE/CFB（老 `.doc`）、RTF、加密 OOXML、非 Word 的 OOXML 包都在
  识别容器后明确拒绝，而不是猜测处理。

## 环境要求

| 依赖 | 说明 |
|---|---|
| Node | `>=18` |
| Python | `3.10+`（实测 3.11.4）。解析发生在子进程中 |
| `lxml` | **可选**。缺失时自动回退标准库 `xml.etree`，并在结果里上报 `WARN: XML_BACKEND_FALLBACK` |
| `docling` | **可选**。仅深度引擎需要，见下文 |

Python 侧只用标准库 + 可选 `lxml`，不依赖 `python-docx`。`zipfile`、`hashlib`、
`posixpath`、`re` 是全部必需导入。

## 快速开始

```ts
import { createDocxParseModule, resolveConfig } from '@dsh-office-profile/docx-parse';

const module = createDocxParseModule({ config: resolveConfig({}) });

const output = await module.handlers.execute({
  artifactRef: { id: 'doc-1', uri: 'D:/docs/report.docx' },
  operation: 'execute',
  requestId: 'req-1',
});

console.log(output.result.ir.outline);   // 由标题段落派生的大纲
console.log(output.result.ir.counts);    // 块/表格/样式/标题计数
console.log(output.warnings);            // 非致命诊断

await module.dispose();
```

注册到 Profile 时用 `register`，它会调用 `registry.registerModule(module)`。

## 三个接口

| 接口 | 回答什么 | 使用的引擎 | 输出 |
|---|---|---|---|
| `inspect` | 「这是什么产物？」——格式身份与能力画像 | **恒为轻量引擎** | `FormatProfile` |
| `execute` | 「把结构给我」——完整 IR + 补全的 artifact 引用 | 深度引擎（若配置） | `ExecuteResult { ir, artifact }` |
| `verify` | 「这份文档满足我给的要求吗？」 | 深度引擎（若配置） | `VerificationReport` |

`inspect` 恒走轻量引擎是一条**路由不变量**而非优化：它的语义是「快速看一眼」，
让它付出加载版面模型的代价（冷启动以秒计）会直接违背该接口的承诺。这条不变量由
测试用可计数的假引擎强制守护。

三个接口的输入都携带 `requestId` 与可选的 `options`（单次调用覆盖配置）；
`verify` 额外要求 `policy`，`inspect` 接受可选 `policy`（命中 error 级检查即
快速失败 `SAFETY_POLICY_DENIED`）。

### 输出信封

```ts
{
  moduleId, requestId, operation,
  result,             // 因接口而异
  artifacts: [],      // 解析不产出新文件，恒为空
  warnings,           // 非致命诊断
  telemetry,          // 性能与规模摘要（不含文档内容）
  verification?,      // 仅 verify（或 inspect 开启 verifyAfterInspect）
}
```

## IR 形状

```ts
interface DocxParseIR {
  profile: FormatProfile;          // 格式身份与能力画像
  metadata: DocumentMetadata;      // docProps/core.xml
  outline: OutlineEntry[];         // 由标题段落派生
  blocks: Block[];                 // 正文块（paragraph | table），按文档顺序
  styles: StyleRecord[];           // word/styles.xml
  relationships: RelationshipRecord[];  // 全部 .rels
  comments: CommentRecord[];       // word/comments.xml
  footnotes: FootnoteRecord[];     // 脚注 + 尾注
  counts: DocumentCounts;          // 派生计数
}
```

`outline` 与 `counts` 是**派生**的，不是独立数据源——它们由 `blocks` 计算得出，
因此永远不会与正文漂移。

## 配置

```ts
resolveConfig({
  engine: { ... },
  limits: { ... },
  timeoutMs: 20_000,
  featureFlags: { ... },
});
```

配置在 `resolveConfig` 内被**深冻结**（含嵌套的 `engine.deep`），防止拿到
`ModuleConfig` 后误改导致同一次请求内行为不一致。非法配置立即抛
`INVALID_INPUT`，把问题拦在模块启动之前。

### `engine`

| 字段 | 默认 | 说明 |
|---|---|---|
| `driver` | `'python'` | 目前仅支持 `python` |
| `pythonPath` | `'python'` | 解释器路径 |
| `scriptPath` | 内置 `src/engine/docx_parse.py` | 覆盖解析脚本 |
| `env` | — | 追加到子进程环境变量 |
| `deep` | — | 可选深度引擎，见下文 |

### `limits`

对**解压后展开规模**设限（zip 炸弹防护）。默认值：

```
maxArchiveEntries          4096
maxEntryUncompressedBytes  32 MiB
maxTotalUncompressedBytes  256 MiB
maxRelationships           4096
maxBlocks                  20000
maxTableCells              20000
maxStyles                  2000
maxComments                2000
maxFootnotes               2000
```

超限时：`enforceLimits` 为 `true`（默认）抛 `LIMIT_EXCEEDED`；为 `false` 时
尽最大努力产出结果并上报 `LIMIT_APPLIED` 告警。

### `featureFlags`

`parseParagraphs` / `parseHeadings` / `parseTables` / `parseStyles` /
`parseRelationships` / `parseComments` / `parseFootnotes` / `extractMetadata`
默认全开；`enforceLimits` 见上。关闭某项会同时把对应验证检查标记为 `skip`。

### `timeoutMs`

默认 `20_000`。深度引擎有独立的 `engine.deep.timeoutMs`（默认 `120_000`），
因为模型冷启动远超轻量解析。

## 错误码

抛出 `DocxParseError`，带稳定的 `code` / `path` / `details`，并可 JSON 往返。

| 错误码 | 含义 |
|---|---|
| `INVALID_INPUT` | 入参非法、字段缺失、不可序列化，或配置非法 |
| `ARTIFACT_NOT_FOUND` | 引用的文件不存在、不可读，或不是普通文件 |
| `UNSUPPORTED_ARTIFACT_URI` | URI scheme 非本地文件（http、memory 等） |
| `FORMAT_MISMATCH` | 声明的扩展名/MIME 与字节内容不符，或包不是 WordprocessingML |
| `UNSUPPORTED_CONTAINER` | 容器可识别但超出范围（OLE/CFB、RTF、加密 OOXML） |
| `SAFETY_POLICY_DENIED` | 安全策略拒绝（`inspect` 快速失败） |
| `ENGINE_UNAVAILABLE` | 无法启动解释器或解析脚本 |
| `ENGINE_FAILED` | 引擎进程非零退出 |
| `ENGINE_TIMEOUT` | 超出超时预算 |
| `ENGINE_PROTOCOL_ERROR` | 引擎输出不符合协议（JSON 非法或结构不符） |
| `PARSE_FAILED` | 包被截断或结构损坏 |
| `LIMIT_EXCEEDED` | 超出资源预算 |
| `VERIFICATION_FAILED` | `verify` 无法产出报告，或强制检查硬失败 |

关键设计：**底层引擎无权选择错误码**。引擎只上报「问题（issue）」，
「问题 → 错误码」的映射由 `mapper.ts` 决定。因此更换引擎实现时，
上层看到的错误语义完全一致。

## 告警码

非致命诊断，经 `warnings` 返回，每条带 `severity`（`info` / `warn` / `error`）
与 `details.issueCode`（保留引擎原始问题码，便于排查「引擎说了什么」）。

| 告警码 | 严重级 | 触发条件 |
|---|---|---|
| `EMPTY_DOCUMENT` | info | 结构合法但无可见内容 |
| `NO_HEADINGS` | info | 开启了 `parseHeadings` 且有可见内容，但没有任何标题 |
| `STYLES_MISSING` | warn | 缺少 `word/styles.xml` |
| `BROKEN_STYLE_REFERENCE` | warn | 段落引用了 styles.xml 中不存在的样式 id |
| `DANGLING_RELATIONSHIP` | warn | 内部关系目标在包内不存在 |
| `COMMENTS_PRESENT` | info | 包含批注 |
| `FOOTNOTES_PRESENT` | info | 包含脚注/尾注 |
| `UNKNOWN_BLOCK_SKIPPED` | info | 正文里存在本模块不解析的元素，其内容被跳过 |
| `DECLARED_HINT_MISMATCH` | warn | 调用方声明的扩展名/MIME 与探测结果冲突（软冲突） |
| `LIMIT_APPLIED` | warn | 触及 `limits` 预算，结果可能不完整 |
| `PARTIALLY_PARSED` | warn | 仅完成部分解析（XML 损坏、部件不可读，或深度引擎降级） |
| `XML_BACKEND_FALLBACK` | info | 未使用 lxml，回退到标准库 |

另有引擎问题码透传：`details.issueCode` 保留引擎上报的原始码
（如 `DOCLING_UNAVAILABLE` / `DOCLING_FAILED`），其 `code` 统一落到
`PARTIALLY_PARSED`。这是刻意留的前向兼容——引擎是开放词汇，可以独立演进
或由调用方替换，发出未登记的码不会丢信息。

**表里每个码都必须有真实产出路径。** 这条不变量由
[`contract.test.ts`](tests/contract.test.ts) 两条跨语言断言守护：mapper 的
「引擎问题码 → 告警码」映射中，每个键都必须能在引擎脚本里找到对应的
`issue(...)` 调用；同时每个声明的告警码都必须能追溯到映射值或 TS 侧直出。
保留「预留但无人发出」的码会让契约承诺大于实际能力，且不会有任何测试失败来提醒。

### `UNKNOWN_BLOCK_SKIPPED` 的判定

正文块只解析 `w:p`、`w:tbl`，并透明下钻 `w:sdt`。其余块级元素分两类：

- **标记类**（`w:sectPr`、`w:bookmarkStart/End`、`w:proofErr`、
  `w:commentRangeStart/End`、`w:permStart/End`、修订范围标记等）不含文本，
  一律静默忽略。这一点很重要：**真实 Word 文档的 `w:body` 末尾恒定有
  `w:sectPr`**，若不豁免，每一份真实文档都会被误报。
- **其余**（`m:oMathPara` 公式、`w:altChunk` 嵌入内容、`w:customXml`、
  块级 `w:ins`/`w:del` 等）可能承载内容但本模块不解析，按元素名去重后上报
  （上限 20 种，其余聚合为一条）。

设计意图：调用方必须能知道「抽取到的文本 ≠ 文档全部文本」。静默丢弃是最糟的
选项——它让缺失不可观测。

## 验证检查项

`verify`（以及 `inspect` + `verifyAfterInspect`）产出如下稳定检查 id：

```
structure.document           结构可解析
output.serializable          输出可 JSON 往返
policy.encryption            加密要求（enforce 时）
policy.headings              要求包含标题
policy.minParagraphs         要求最少段落数
policy.styles                要求存在样式表
policy.resolvedStyles        要求样式引用均可解析
policy.danglingRelationships 要求无悬空关系
```

未在策略中声明的项返回 `status: 'skip'`（而非 `pass`）——「没要求」与「满足了」
是两件不同的事，不应混为一谈。

## 深度引擎（可选 Docling）

在 `engine.deep` 配置后，`execute` 与 `verify` 改走 Docling 后端：

```ts
resolveConfig({
  engine: {
    deep: { driver: 'docling', pythonPath: 'python', timeoutMs: 120_000 },
  },
});
```

采用「**OPC 层复用 + 正文单向增强**」的分工，而不是让 Docling 全量接管：
Docling 把任何输入归一化为统一的 `DoclingDocument`，该模型**不携带** OPC 关系、
批注、脚注，也不保留 `w:pStyle` / `w:outlineLvl` / 字符级 run 格式。若让它接管
全流程，契约中的这些字段会静默消失——那不是换引擎，是能力回退。

因此正文先由标准库产出**完整**记录，再由 Docling **补充**标准库未判定为标题的
段落。合并方向是单向的，Docling 无权覆盖 OOXML 的权威结论，保证「开启深度引擎」
在任何情况下都不会让结果比关闭时更差。

失败语义：

- Docling 不可导入，或解析抛错 → **降级**为轻量结果，并产出告警：
  `code` 为 `PARTIALLY_PARSED`，`details.issueCode` 保留引擎原始原因
  （`DOCLING_UNAVAILABLE` 或 `DOCLING_FAILED`）。
  降级必须可见，不能静默。
- 脚本路径配错 → 硬失败 `ENGINE_FAILED`。配置错误不应被伪装成「正常降级」。

> **注意**：Docling 面向的是 PDF 版面问题（OCR、多栏阅读顺序、表格结构）。
> DOCX 已是结构化 XML，标准库本就能拿到权威样式信息，而 Docling 反而丢失样式粒度。
> 这里的实际增益主要是**自定义样式名文档的标题识别**。若期望整体保真度提升，
> 收益会低于预期。

## 代码结构

```
src/
  index.ts        composition root：编排 + 错误归一化 + 信封组装；【公开契约唯一的出口】
  internal.ts     内部接缝的统一出口（不构成契约，可自由重构）
  contract.ts     类型化契约：接口/输入输出/错误码/告警码/能力表
  config.ts       默认值、校验、覆盖合并、深冻结、JSON-Schema
  mapper.ts       裁决 + 派生 + 归一化（引擎问题 → 错误/告警，blocks → IR）
  verifier.ts     局部质量检查
  errors.ts       DocxParseError 与分类翻译
  telemetry.ts    遥测端口与实现
  domain/
    docx-parse.ts 领域模型（= 引擎载荷形状）+ 信任边界 + 派生纯函数
  engine/
    adapter.ts              引擎端口、子进程运行器、DocxEngine 实现
    docx_parse.py           轻量引擎（纯标准库 + 可选 lxml）
    docx_parse_docling.py   深度引擎桥接（复用 OPC 层，Docling 只补标题）
tests/
  contract.test.ts    契约测试（不需要 Python，用假引擎）
  edge-cases.test.ts  真实样本的边缘场景（需要 Python）
  regression.test.ts  确定性与形状稳定性
  deep-engine.test.ts 路由隔离、配置校验、深度引擎合并逻辑
  support.ts          共享夹具与假引擎
  stubs/docling/      docling 桩模块（见下）
fixtures/
  generate_fixtures.py  19 个可审计样本的生成器
module.json            模块清单（与代码内常量逐字段比对，回归测试守护）
types/office-deps.d.ts 未安装的 peer 依赖的类型垫片
```

### 分层职责

- **引擎**（`engine/`）：拿字节，产出中立的 `ParseResult`。唯一了解 `zipfile` /
  `lxml` 的地方。
- **映射**（`mapper.ts`）：裁决 + 派生 + 归一化。
- **验证**（`verifier.ts`）：对 `ParseResult` 做局部质量检查。
- `index.ts` 只做「顺序编排 + 错误归一化 + 信封组装」，不含解析逻辑。

### 公开与内部的边界

`package.json` 的 `exports` 只映射了 `"."` → `src/index.ts`，因此**该文件导出了
什么，就等于调用方能依赖什么**。据此它只导出四类东西：

1. **身份** —— 模块清单、能力表、错误码/告警码表；
2. **装配** —— `createDocxParseModule` / `register` 及其选项类型；
3. **配置** —— `resolveConfig` / `configSchema` / `ConfigOverrides`；
4. **扩展** —— `DocxEngine` 端口、错误类型，以及全部输入输出类型。

映射、领域派生、验证、遥测与具体引擎实现都在 `src/internal.ts` 后面，
不构成契约，可以自由重构而不属于破坏性变更。这意味着你可以**替换引擎后端**
（实现 `DocxEngine` 并通过 `createDocxParseModule` 注入）而不触及公开契约。

### 单一数据源（不做双 IR）

引擎输出的形状 = 领域模型 = IR 的数据部分，`mapper` 只做裁决/派生/排序，
不存在「引擎中间模型 → 模块 IR」的二次搬运。`adapter.ts` 是唯一了解子进程协议的文件。

子进程协议：stdout 恰好一个 JSON 对象、stderr 供人阅读的诊断、退出码 0（已产出载荷，
即便产物不受支持）或 2（彻底无法产出）。强制 UTF-8（`PYTHONIOENCODING` /
`PYTHONUTF8`），避免中文路径在 Windows 控制台变成乱码。

## 测试

```bash
npm run typecheck
npm test              # 全量
npm run test:contract # 仅契约测试（无需 Python）
```

**契约测试完全不依赖 Python**：它用假引擎驱动完整流水线，因此可以在没装解释器的
机器上验证对外承诺。真实的解析正确性由 edge-cases / regression 覆盖。

样本不在仓库里提交二进制，而是由 [`fixtures/generate_fixtures.py`](fixtures/generate_fixtures.py)
在测试前重建：生成器本身就是文档（读一遍就知道每个样本覆盖什么场景），
也避免把恶意样本特征提交进仓库触发杀软告警。

环境变量：

| 变量 | 作用 |
|---|---|
| `DOCX_PARSE_PYTHON` | 指定测试与 fixture 生成所用的解释器（默认 PATH 上的 `python`） |

**依赖失败的用例会被 skip，而不是误报为失败**：没有 Python 时跳过真实引擎用例，
没有 Docling 时跳过 Docling 集成用例。

### `tests/stubs/docling/`

一个行为由环境变量驱动的 `docling` 桩模块。它让「标题合并逻辑」可以在不安装
约 1.5GB 真实依赖（torch 等）的前提下被精确验证：层级映射、只补充不覆盖、
`iterate_items()` 两种形状兼容、失败降级。

桩只证明**合并逻辑本身正确**，不证明真实 Docling 的 API 形状与假设一致——
后者由 `deep-engine.test.ts` 中的 gated 集成用例兜底，在装了 `docling` 的机器上
设 `DOCX_PARSE_PYTHON` 指向该环境即可自动启用。

## 安全说明

- 解析过程**从不把部件解压到磁盘**，也从不解析（更不访问）关系目标的内容，
  因此恶意文档无法借此产生写文件、联网或路径穿越行为。
- XML 解析已加固，抵御 XXE 与实体膨胀攻击。
- ZIP 中央目录里的 `file_size` 仅作参考，真实读取按硬字节上限流式截断。
- 引擎载荷在 `parseParseResult` 处逐字段校验，是**信任边界**——畸形结构会被拒绝
  为 `ENGINE_PROTOCOL_ERROR`，换引擎不降低校验强度。

## 设计取舍与已知限制

- **轻量优先，不做双 IR**。DOCX 已是结构化 XML，重型版面模型（Docling / Marker /
  MinerU）解决的是 PDF 的像素/OCR 问题，对 DOCX 不产生收益；引入它们会牺牲
  确定性与可审计性。
- **`counts` 与 `outline` 是派生的**，不独立存储，因此不可能与正文漂移。
- **深度引擎的收益有限**（见上文「深度引擎」一节），主要针对自定义样式名的标题识别。
- **`blocks` 顺序即文档顺序**；阅读顺序不做重排（DOCX 的 XML 顺序就是权威顺序）。
- 引擎每次调用启动一个短命子进程：换来的是零状态泄漏与强隔离，代价是几十毫秒的
  进程启动开销。对超大文档批量处理时值得注意。
