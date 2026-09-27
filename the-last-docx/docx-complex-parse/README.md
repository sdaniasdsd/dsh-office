# docx-complex-parse

**复杂 DOCX 解析模块。** 面向「普通解析器搞不定」的文档：带合并的复杂表格、跨页关系、浮动对象、需要页内坐标的来源定位，以及每项结论的置信度。

常规 DOCX 的结构解析由兄弟模块 `docx-parse` 负责。本模块是它的**上位替代品**，但代价高得多——需要 Python 运行时与 rdocx 版面分析——因此设计成按需加载：Profile 只在真的遇到复杂文档时才把它接进来。

---

## 它解决什么问题

| 处理范围 | 本模块的回答 |
| --- | --- |
| **复杂表格** | 逻辑网格展开：把 `tblGrid` / `gridSpan` / `vMerge` / `gridBefore` / `gridAfter` 还原成矩形网格 + 合并区域清单 |
| **跨页关系** | 表格跨页范围、全部块级版面片段、页到节点索引，以及 XML 中声明的分页点和分节 |
| **浮动对象** | 正文 DrawingML 锚定/内联对象的种类、尺寸、相对偏移、环绕与媒体关系；绝对页内矩形保留未知 |
| **来源坐标** | 每个结构同时带**结构路径**（永远可得到）与**页内矩形**（需要版面模型，可能为 null） |
| **置信度** | 每项结论带 `score` 与 `source`，说明这个结论是怎么来的 |

---

## 两条证据线，一个会合点

复杂文档里，「结构」和「位置」来自两个互不相干的来源，本模块不把它们混为一谈：

```
结构线（纯 XML，永远可得）          版面线（rdocx，可能拿不到）
  w:body 的直接子节点                页尺寸 + 每个正文块的页内矩形
  w:tbl → w:tblGrid/gridSpan/vMerge        │
  w:p                                      │
        └──────── body_index ──────────────┘
          （两条线按文档顺序给正文块编号，序号即会合点）
```

- 引擎 `src/engine/docx_complex_parse.py` 用 XML 读结构、用 [rdocx](https://pypi.org/project/rdocx/) 读版面，输出一份**原始观察**。
- 版面拿不到时不会假装成功：`layout.available = false`，坐标降级为纯结构路径，并广播 `PAGE_GEOMETRY_UNAVAILABLE`。调用方可以用 `featureFlags.requirePageGeometry` 把这种降级改成硬失败。

### 为什么结构不走 rdocx

rdocx 的 Python 绑定是**纯高层 API**：它不暴露 `gridSpan` / `vMerge` / `tblGrid`。而分页与页内几何在 XML 里根本不存在。所以这是两条必须分开的线，不是一个可以在两者之间挑一个的选择题。

---

## 逻辑网格展开（核心算法）

OOXML 表格**先有网格**，合并是相对指令，不是一个可以直接读出来的矩形。参照 [python-docx](https://python-docx.readthedocs.io/) 的 `Table._cells` 模型实现，规则是：

1. `w:tblGrid/w:gridCol` 决定逻辑列数；
2. 逐行游标遍历，`w:gridBefore` 让本行起点右移，`w:gridAfter` 让终点左移；
3. `w:gridSpan` 让一个单元格横占多列，其余列位补占位格；
4. `w:vMerge` 的 `restart` 是纵向起点，`continue`（或**省略 `w:val`**，这是最常见写法）向下延伸；
5. **落在合并区内的网格地址，返回该区最左上角的主格。**

在第 5 条上，本模块与 python-docx 完全一致——已用 8 个 fixture 与 python-docx 的 `_cells` **逐格比对**通过。

### 同时修掉了 python-docx 的三处不完整

| 情形 | python-docx 1.2 | 本模块 |
| --- | --- | --- |
| `w:gridBefore` / `w:gridAfter` | 未建模，整行格子左移错位 | 按声明右移/左移 |
| 找不到起点的 `vMerge=continue` | `IndexError: list index out of range` | 降级为独立格 + `MERGE_INCONSISTENT` |
| 行比 `tblGrid` 更宽 | 索引整体错位 | 扩列并在 `merges` 里保持自洽 |

以上三条已登记为「参考不完整」而非「实现有误」——它们是本模块**有意**与参考不同的地方。

---

## 三个接口

```ts
import { createDocxComplexParseModule } from '@dsh-office-profile/docx-complex-parse';

const module = createDocxComplexParseModule({
  config: { engine: { driver: 'rdocx', pythonPath: '/usr/bin/python3' } },
});

await module.handlers.inspect({ artifactRef, requestId, operation: 'inspect' });
await module.handlers.execute({ artifactRef, requestId, operation: 'execute' });
await module.handlers.verify({ artifactRef, requestId, operation: 'verify', policy });
```

- **`inspect`** → `FormatProfile`（格式/容器/加密/宏/外部引用/嵌入对象/部件与关系）。传 `policy` 时顺带评估，硬失败即 `SAFETY_POLICY_DENIED`。
- **`execute`** → `ExecuteResult{ ir, artifact }`。`ir` 是 `FormatIR`（包级事实）+ `ComplexContent`（归一化后的复杂结构）。
- **`verify`** → `VerificationReport`。三态 `pass` / `fail` / **`skip`**；`skip` 表示策略**没有声明**该要求，与「通过」不是一回事。

三个 handler 共用同一段「跑引擎 + 校验 + 可用性判定」，因此同一份 artifact 在三个接口下的结论完全一致。

### 包级 IR 与内容 IR 的关联

`DocxComplexParseIR` 仍是 `FormatIR + { content: ComplexContent }`，不引入第二个独立返回值。
包级的 `parts / relationships / indicators` 保留容器事实；内容侧通过
`part + structuralPath` 定位原文，通过 `FloatingObject.part + relationshipId` 关联媒体。
`verify` 的 `content.references` 检查节点、分页、分节、页索引以及图形关系的关联；
内部媒体目标必须在包内存在，外部关系只校验引用、不访问网络。

- `blocks[].fragments` 保留全部模型片段；`coordinate` 仍代表起始片段。
  跨页块在每一张实际出现的页的 `nodeIds` 中登记一次，保持文档顺序。
- `breaks` 读取 `w:br type="page"`、`pageBreakBefore`、`lastRenderedPageBreak`
  及分页型分节。段落内的分页点用精确 XML 路径表达，`beforeNodeId=null`；
  保存时的渲染标记不用于推算当前页号。分页声明的前后页号保持未知。
- `sections` 将段尾/正文末尾的 `sectPr` 映射到相应节的起点，读取页尺寸、页边距、
  分栏和开始方式；单位为 `twip`。尺寸或边距声明不完整时，`geometry=null`。
- `floats` 保留 DrawingML 的锚定与内联对象。尺寸与偏移用 `emu`，允许负偏移；
  未提供对象级版面时，页号、矩形及坐标单位保持 `null`，同时发出告警。
- `content.coverage` 区分 `observed / partial / disabled / unavailable`，范围为当前正文扫描。
  空数组不再承担“已确认不存在”的隐含断言。旧 v1 引擎可以省略新增观察集合；
  旧消费方可以忽略新增的 `fragments`、`coverage` 字段。

三个观察开关分别生效。关闭分节输出仍可读取分节导致的分页声明；关闭分页输出仍可产出分节。
`maxFloatingObjects` 已接入引擎和 TS 预算检查；关闭限额强制执行后仍报告 `LIMIT_APPLIED`。

声明语义参照 Microsoft 的 [SectionProperties](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.sectionproperties)
与 [DrawingML Anchor](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.anchor) 文档。

---

## 置信度：`source` 比 `score` 重要

每个结论都带 `Confidence{ score, source, notes }`：

| `source` | 含义 |
| --- | --- |
| `native` | OOXML 直接给出的（格子位置、段落文本） |
| `structural` | 由结构推导的（被合并覆盖的占位格、列数靠最长行推导） |
| `heuristic` | 猜的（续接但找不到起点的格子） |
| `model` | 来自版面模型的 |

`score` 低于 `config.confidenceFloor`（默认 0.5）的结论会附带 `LOW_CONFIDENCE` 告警——**只是告警，不阻断**。

---

## 安全相关：空数组是断言，不是沉默

`indicators.macros` 为空数组，等于本模块在说「这份文档没有宏」。这是**安全结论**，所以：

- 它只能来自引擎对容器的实际读取（`[Content_Types].xml`、全部 `.rels`、`vbaProject.bin`），不能猜；
- 协议上它是**必填**的——缺失按 `ENGINE_PROTOCOL_ERROR` 拒绝，而不是回落到空数组；
- `encrypted: null` 表示「无法判断」，**不折算成 `false`**。策略说「不允许加密」而状态未知时，验证按**不满足**处理——把决定权交回策略作者，而不是替他把关。

---

## 配置

由 Profile 注入，模块**从不**读取全局配置。全部字段与默认值见 `module.json` 的 `configSchema`。

```ts
{
  engine: { driver: 'rdocx', pythonPath: 'python' },
  limits: { maxBlocks: 50_000, maxTableCells: 200_000, /* …8 项 */ },
  timeoutMs: 60_000,
  confidenceFloor: 0.5,
  featureFlags: { /* 9 个开关，仅 requirePageGeometry 默认 false */ },
}
```

- `limits` 是**解压炸弹防护**：条目数、单条目解压上限、总解压上限、关系数、块数、单元格数、页数。命中且 `enforceLimits` 打开时以 `LIMIT_EXCEEDED` 失败；关闭时降级为 `LIMIT_APPLIED` 告警——不静默放过。
- `driver` 可配 `rdocx` / `docling` / `mammoth` / `markitdown`，但**目前只有 `rdocx` 有内置实现**。选了其余三个会在调用时以 `ENGINE_UNAVAILABLE` 快速失败，而不是悄悄换用别的引擎。

---

## 运行要求

- **Node** ≥ 18（ESM）。
- **Python** ≥ 3.10，并安装：

  ```bash
  pip install rdocx
  ```

  默认采集版面时，缺少 `rdocx` 会以 `ENGINE_UNAVAILABLE` 失败；已安装但排版失败时按
  `requirePageGeometry` 决定降级或失败。显式设置 `parsePageGeometry=false` 可只读取结构与声明，无需 rdocx。

---

## 测试

```bash
npm run typecheck     # tsc --noEmit
npm test              # 全部单测
npm run test:grid     # 逻辑网格展开（含边界与自相矛盾文档）
npm run test:domain   # 信任边界：引擎违约一律 ENGINE_PROTOCOL_ERROR
npm run test:mapper   # 归一化：坐标、页几何、置信度、纯函数性
npm run test:assembly # 装配层：三个 handler、策略、预算、dispose
npm run test:observations # Python 进程 → 原始协议 → 双 IR 关联（需 Python，无需 rdocx）
```

原有单元测试使用假引擎。`observations.test.ts` 会在临时目录生成 OOXML ZIP 夹具，
通过真实 Python 子进程覆盖声明解析、功能开关、协议校验、预算与跨 IR 引用；
`tests/test_observations.py` 补充 XML 边界用例。版面桥接测试用确定性 rdocx 桩，
不代表真实 rdocx 排版准确性。解释器可通过 `DOCX_PARSE_PYTHON` 指定。

---

## 已知边界

- 声明扫描覆盖直接正文段落/表格中的分页与 DrawingML，以及正文分节；页眉页脚、
  内容控件包裹的正文块、文本框内部的独立文本流尚未建模。
- 旧式 VML 图形通过 `UNSUPPORTED_CONTENT` 明确报告，浮动对象覆盖状态为 `partial`；
  对象的绝对坐标、跨页表内单元格的具体页号无法由块级版面可靠恢复，保留 `null`。
- `lastRenderedPageBreak` 只表示上次排版保存的断点位置，不含可信的当前页号。
- 页面汇总中的边距仍是内容片段范围的推估；分节中的边距才是 XML 声明值。
- 本模块**只读**：不修改文档，不执行宏，也不分析宏的内容——它只回答「有没有」。

---

## 目录

```
src/
  contract.ts                    对外契约（类型、错误码、告警码、配置）—— 冻结表面
  index.ts                       装配层：模块定义 + 三个 handler + 注册
  config.ts                      配置解析、合并、校验、冻结
  telemetry.ts                   遥测与日志桥接
  verifier.ts                    局部验证（三态检查 + 汇总）
  mapper.ts                      归一化：原始观察 → ComplexContent / FormatIR
  errors.ts                      错误分类体系
  domain/docx-complex-parse.ts   信任边界：校验引擎输出
  engine/adapter.ts              引擎适配（唯一知道「要 spawn Python」的文件）
  engine/docx_complex_parse.py   引擎本体：XML 读结构 + rdocx 读版面
  engine/observations.py         分页、分节、DrawingML 声明读取
tests/                           单测 + OOXML ZIP/真实进程集成测试
```

## 依赖

| 模块 | 必需 | 用途 |
| --- | --- | --- |
| `office-core` | 是 | `FormatIR` / `ArtifactRef` / `Warning` 等共享类型 |
| `office-safety` | 是 | `SafetyPolicy` |
| `office-files` | 否 | 非本地 URI 的物化器 |
| `office-test-kit` | 否 | 共享验证报告链路 |
| `python` + `rdocx` | 是（运行时） | 版面分析 |
