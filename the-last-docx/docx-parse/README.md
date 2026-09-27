# docx-parse

> `dsh-office-profile` 中负责 **DOCX 结构解析** 的独立模块。它把一份复杂的 Word 文档拆成两个互相印证、且能精确互指的视图 —— **双 IR**。

---

## 这个模块解决什么问题

办公场景里的 DOCX 从来不是「一串文字」。一份真实的合同、研报或公文，结构上至少混着这些：

- 多级标题，级别可能来自 `outlineLvl`、也可能只藏在样式名里、还得靠启发式兜底
- 表格，带横向合并（`gridSpan`）与纵向合并（`vMerge`）
- 批注、脚注、尾注，各自挂在正文中某个具体位置
- 样式表（`styles.xml`），且样式之间通过 `basedOn` 形成继承链
- 列表编号、宏、模板、加密容器、损坏包……

**「文字提取」对这些结构一律视而不见，而办公 Agent 恰恰要跟这些结构打交道。**

更麻烦的是下一层需求：Agent 改完一处之后，怎么知道改的**就是**用户说的那一处？怎么知道没把批注、书签、样式带坏？——「第 3 段」这种说法在文档一动之后就失效了。

`docx-parse` 就是为了让上层能**精确指代文档里的某个节点，并在文档变动后仍然认得它**。

---

## 双 IR：一份文档，两个视图，一张对应表

| 视图 | 给谁看 | 内容 |
| --- | --- | --- |
| **语义视图** `semantic` | 人与 LLM 的「地图」 | 标题、段落、表格行列、单元格段落、批注/脚注/尾注，都是归一化后的意义单位 |
| **物理视图** `physical` | 写入端的「施工图纸」 | OOXML 里真实存在的元素：结构路径、原生 `paraId`、内容指纹 |
| **对应表** `sourceMap` | 连接两者的「账本」 | 语义节点 ↔ 物理节点，外加三向索引 |

对应表是三张索引，缺一不可：

| 索引 | 方向 | 回答的问题 |
| --- | --- | --- |
| `bySemanticId` | 语义 id → 物理指针 | 「要改这个节点，动 XML 哪里」 |
| `byPointer` | 物理指针 → 语义 id | 「这个 XML 元素是什么」 |
| `byFingerprint` | 内容指纹 → 语义 id | 「用户说的那句话在哪」 |

三张索引全部是普通对象而非 `Map` —— 因为输出必须**可 JSON 序列化**，这是跨越 RPC 边界的硬性要求。

---

## 三个能力

模块向 Profile 注册三个 handler，输入输出都可序列化。

### `inspect` —— 这是什么产物，结构长什么样

返回 `FormatProfile`：文档种类（docx/docm/dotx/dotm）、包结构、是否含宏、是否加密、资源占用。**不产出完整 IR**，用于快速摸底。

可选传入 `policy`，命中即快速拒绝（`SAFETY_POLICY_DENIED`）。

### `execute` —— 产出双 IR

```ts
import { createDocxParseModule } from './src/index';

const module = createDocxParseModule();
const output = await module.handlers.execute({
  artifactRef: { id: 'doc-1', uri: 'file:///reports/quarterly.docx' },
  operation: 'execute',
  requestId: 'req-1',
});

const content = output.result.ir.content;   // 双 IR
content.semantic.blocks;                    // 标题 / 段落 / 表格
content.physical.nodes;                     // 真实 OOXML 元素
content.sourceMap.bySemanticId;             // 三向索引
```

### `verify` —— 一致性验证报告

产出满足 `office-test-kit` 形状的报告，逐项带稳定 id：

| 检查项 | 断言的是 |
| --- | --- |
| `structure.package` | 包结构有效 |
| `content.semanticView` | 语义视图自洽 |
| `sourcemap.integrity` | **每个语义节点都能解析到锚点，且指针不被两个节点争抢** |
| `identity.deterministic` | 重算 id 能复现既有 id |
| `output.serializable` | 输出可 JSON 序列化 |
| `policy.*` | 加密、块预算、单元格预算 |

每项状态为 `pass` / `fail` / `skip` —— `skip` 表示策略没提出该要求，而非通过。

---

## 身份方案：composite-anchor-v1

节点身份是本模块相对「只读探测文档」最核心的增量。

**核心思路是冗余取证** —— 不赌任何单一选择器，同时记三种彼此独立的定位依据：

| 选择器 | 独立性来源 | 失效场景 | 由谁兜底 |
| --- | --- | --- | --- |
| `paraId` | OOXML 原生 id | 可选，常常缺失 | quote / path |
| `quote` | 内容 | 内容被编辑 | paraId / path |
| `structuralPath` | 位置 | 增删导致漂移 | paraId / quote |

只要还有一条有效，节点就能被重新找回；三条全失效才算真正「失联」。

**id 只由 `paraId` 或结构路径派生，不含内容。** 这是刻意的：内容一改 id 就变，挂在节点上的批注与书签就会失联。内容选择器的职责是「重新找回」，不是「定义身份」——两者必须分开。

一个实测例子（`tests/blackbox.test.ts`）：文档中插入一段，由于没有 `paraId`，`Gamma` 的 id 必然改变，但用旧版的内容指纹去新版查询，照样把它从 `w:p[3]` 追到了 `w:p[4]`。

> 局限要说清楚：缺 `paraId` 时身份退化为位置，插入会让后续 id 集体错位。这正是 Word 原生文档带 `paraId`、而脚本生成的常常不带所带来的差异。

---

## 修订层与门槛判定（内部模块）

除三个能力外，模块内还有两个**内部目录**：`src/revision/` 与 `src/governance/`。

### `revision` —— 修订账本

- **按稳定 id 做差异，而非文本 diff**：能区分「这句被改写了」与「这句没了、旁边多了一句」。中途插入一段时，后续未动节点不会被误报。
- **变化维度细分**：`text` / `structure` / `style` / `level`。把 H2 改成 H3 而文字一字未动，通用 diff 完全看不见，这里能看见。
- **只追加账本**：修订历史一旦写入不可改写。
- **信任边界**：`record()` 先校验意图再落账，畸形输入直接抛 `INVALID_INPUT`。

```ts
import { createRevisionLedger } from './src/index';

const ledger = createRevisionLedger();
await ledger.record({
  intent: {
    id: 'i1', requestId: 'req-1', source: 'user',
    stage: 'revision', kind: 'content',
    text: '把营收改成 12%',      // 用户原话，原样保留
    targetIds: [],
  },
  after: newDocIR,               // 改动之后重新解析出的双 IR
});
```

`after` 省略时表示「只表达意图、文档还没变」，不会伪造差异。

### `governance` —— 硬编码的门槛条件

判断一次任务该用多严的治理强度，输出 `strict` / `lite` / `exempt` 三档，并说明**命中了哪条信号**。

规则自上而下，命中即返回：

1. 无任何严格信号 → `lite`
2. 单次生成、且用户没先给格式要求 → `exempt`
3. 用户先给了格式要求 → `strict`
4. 场景对文本提交严格 → `strict`
5. 用户格式修改意愿明显（≥2 次） → `strict`

四个信号：`explicitFormatRequirements` / `strictTextSubmission` / `strongFormatEditIntention` / `singleShotGeneration`。证据**只累加不否决**：账本说「有」、宿主说「没有」，仍然算有——沉默不代表否认，默认导向宽松。

### `brief` —— 交给 Agent 的判断依据

把双 IR、账本、门槛判定缝成一份结构固定的简报（`renderAgentBrief`）：治理档位 → 执行纪律 → 文档大纲 → 可点名目标 → 修订账本。可点名目标一律带 `id=`，并明确要求不得用「第 N 段」指代。

---

## 边界与非目标

**本模块只读，不写。** 解析与验证是它的全部职责；把改动落到 OOXML 上属于另一个模块。

它物化给下游写入端的，恰好是写入所需的三样东西：

- `semanticId` —— 改哪个
- 物理指针 —— 改哪里
- `anchor` —— 改完如何自检没改错

---

## 配置

配置由 Profile 注入，模块**从不读全局配置**。

```ts
const module = createDocxParseModule({
  config: {
    limits: { maxBlocks: 20000 },
    featureFlags: { parseComments: false },
    timeoutMs: 15000,
  },
});
```

- **`limits`** —— 资源预算，抵御 zip 炸弹与超大包（条目数、解压体积、关系数、块数、单元格数、注释数）
- **`featureFlags`** —— `parseHeadings` / `parseTables` / `parseStyles` / `parseComments` / `parseFootnotes` / `resolveAnchors` / `enforceLimits`
- **`engine`** —— 默认 Python 桥接（`src/engine/docx_parse.py`，`lxml` 缺失时回退 stdlib）

单次调用可在 `options` 里覆盖上述任意一项。

---

## 错误码与告警码

**错误码**表示调用失败，共 14 个，稳定不变。两类值得一提：

- `SOURCEMAP_INCOMPLETE` —— 本模块独有：不是「文档坏了」，而是「两份视图对不上号」。出现即显式失败，绝不返回一份看起来完整、实则无法定位的 IR。
- `UNSUPPORTED_CONTAINER` —— OLE/CFB（旧式 `.doc`、加密 OOXML）、RTF 明确超出范围。

**告警码**表示调用成功但结果需要被注意，共 11 个。遇到未建模的 OOXML 元素一律降级为告警而非失败——保证「部分解析」仍然可用，同时如实告知哪些内容没被理解：

`EMPTY_DOCUMENT` / `PARTIALLY_PARSED` / `UNSUPPORTED_CONTENT` / `LIMIT_APPLIED` / `STYLE_FALLBACK` / `HEADING_LEVEL_INFERRED` / `ANCHOR_INCOMPLETE` / `COMMENT_ORPHANED` / `NOTE_ORPHANED` / `ENCRYPTED_ARTIFACT` / `VERIFICATION_FAILED`

---

## 目录结构

```
src/
  contract.ts        冻结的公开表面（类型、错误码、告警码、配置形状）
  index.ts           Profile 注册入口：三个 handler + 模块工厂
  config.ts          配置解析与默认值
  errors.ts          错误分类
  mapper.ts          引擎中立观察 → 双 IR / FormatIR 的唯一翻译层
  sourcemap.ts       身份层：复合锚点、语义 id、三向对应表
  verifier.ts        结构 / 可序列化 / 对应表完整性 / 身份确定性
  telemetry.ts       遥测（不含文档内容与绝对路径）
  engine/            Python 桥接（薄封装，不泄漏引擎类型）
  domain/            引擎中立领域模型 + 信任边界校验
  revision/          修订账本（内部模块，将来可整体搬出）
  governance/        门槛判定（同上）
  brief.ts           给 Agent 的判断依据渲染
```

`revision/` 与 `governance/` 各自有 `index.ts` 作为接缝，搬成独立模块时整目录移走即可。

---

## 开发

```bash
npm install
npm run typecheck          # tsc --noEmit
npm test                   # 129 个测试
npm run fixtures:build     # 重新生成 fixtures（需要 python）
```

测试分六层：

| 文件 | 层次 | 特点 |
| --- | --- | --- |
| `contract.test.ts` | 契约 | 三能力形状、错误码/告警码逐项对齐 |
| `edge-cases.test.ts` | 边界 | 畸形输入、截断包、zip 炸弹、非 Word 产物 |
| `regression.test.ts` | 回归 | 对应表自洽、确定性、批注锚定 |
| `revision.test.ts` | 修订层 | 身份感知差异、意图校验、账本语义 |
| `governance.test.ts` | 判定层 | 硬编码规则逐条钉住 |
| `blackbox.test.ts` | **黑盒** | 只走公开接口，拿真实 `.docx` 跑通全链路 |

黑盒测试会把结果写成 `fixtures/_generated/blackbox-report.md`，可直接阅读真实输出。

### 纪律

- **确定性**：不生成时间戳或随机数，节点 id 由宿主提供或从文档派生；重复解析同一文件产出逐字节一致的 IR。
- **可序列化**：任何输入输出都必须能 JSON 序列化，且不含引擎对象（Python 句柄、`lxml` 实例、`Buffer`）。
- **信任边界**：引擎只上报「观察到的原始事实」，无权决定错误码、IR 形状或节点身份。
- **纯函数工厂**：`createDocxParseModule` 不做 I/O、不注册副作用，可随意创建互不干扰的实例。
