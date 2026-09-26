# WorkBuddy vs DSH + DOCX 插件：对比测试包

这是一套**测试用例与执行规程**，不是已经跑完的成绩报告。它用相同输入、相同冻结提示词和同一个渲染环境，对比 WorkBuddy 自带的 DOCX 能力与 DSH 接入 `@deepseek-ai/dsh-docx` 后的完整表现。

## 测什么

| 用例 | 主要看点 | 起始文件 |
| --- | --- | --- |
| `cases/01-format-preservation.yaml` | 版式整理时能否保留正文、表格、批注、脚注 | `modules/docx-parse/fixtures/_generated/rich.docx` |
| `cases/02-targeted-edit-review.yaml` | 多页节点编辑、相同状态的消歧、摘要联动、修订痕迹与关系保护 | `fixtures/case-02-project-status.docx` |
| `cases/03-create-from-brief.yaml` | 从空白 DOCX 创建结构化报告、表格与交付完整性 | `modules/docx-inspect/fixtures/_generated/empty.docx` |
| `cases/04-labor-contract-amendment.yaml` | 多条款与附件联动、工资比例推导、相同金额/相近期限消歧、修订/批注/关系保护 | `fixtures/case-04-synthetic-labor-contract.docx` |

四个样本均由仓库内的 fixture 生成器产生，不含真实客户资料。若文件尚未生成，在仓库根目录运行：

```powershell
python modules/docx-parse/fixtures/generate_fixtures.py --out modules/docx-parse/fixtures/_generated
python modules/docx-inspect/fixtures/generate_fixtures.py --out modules/docx-inspect/fixtures/_generated
python benchmarks/workbuddy-vs-dsh/build_case_02.py
python benchmarks/workbuddy-vs-dsh/build_case_04_labor_contract.py
```

第二个用例是专用多页样本，包含重复状态、合并标题行和重复表头、既有批注、真实脚注、页码字段及超链接。提示词要求只更新指定工作流、摘要数字和一处控制说明，同时新增批注与真实修订；特别检查模型是否误改同文案的 Acceptance testing 行。这比单句/单格小样本更能检验复杂文档中的精确定位和最小改动能力。

第四个用例是虚构的 8 页劳动合同，重点测试正文与附件字段同步、基本工资和 80% 试用期工资的推导、相同金额的不同薪酬性质、常规办公地址与应急集合点消歧，以及在保留原批注/脚注/链接的同时新增修订和批注。它是文档编辑能力的压力测试，不是可签署合同或法律意见。

## 公平对比规程

1. **先记平台与模型。** 尽量让 WorkBuddy 和 DSH 使用同一模型版本、相近推理档位和相同联网/文件权限。如果无法统一，仍可比较产品整体体验，但结论不能归因于“插件本身更强/更弱”。在 `runner.json` 记录模型、版本、配置与时间。
2. **每个候选从干净副本开始。** 不把 WorkBuddy 产物交给 DSH 继续改，也不反过来。两个候选分别得到同一份源文件（先核对 SHA-256）和同一份用例提示词。源文件只读保留。
3. **首稿优先。** 每个用例先各跑一次，保存第一次输出为 `round-00` 并评分。若另行测修复能力，最多两轮；修复轮必须重复原提示词并只附上一轮实测到的差异，不能趁修复追加新需求。首稿分与修复后最终分分别报告。
4. **统一渲染再评视觉。** 对基线和两边每一轮 DOCX，都使用同一版 LibreOffice 26.8.0 + Poppler 26.09.0、同一字体环境、同一 DPI 生成逐页 PNG。可以用本插件包内的运行时作公共渲染器，但 WorkBuddy 和 DSH 的候选都必须经过它；不能拿 WorkBuddy 的 Word 预览对比 DSH 的 LibreOffice 页面图。
5. **每页都检查。** 保存每页 PNG，不只留拼图；查看所有页面，记录页数、缺页/空白页、裁切、重叠、表格宽度、标题与分页。视觉图是版式证据，DOCX 结构验证是内容/关系证据，两者不能互相替代。
6. **不要手工修候选。** 可以把候选文件复制到证据目录，但不能先人工修改再评分。保留每轮原件、提示词、SHA-256、验证输出与截图。

## 建议的执行流程

1. 复制用例源文件到两个独立运行目录，核对哈希一致；渲染并保存基线每页图片。
2. 把 YAML 中 `primary_prompt` 原样发给 WorkBuddy。另开干净 DSH 会话，把同一提示词发给 DSH，并按 DSH 插件流程导入同一用例文件。除必要的导入/导出操作外，不给其中一边额外解释或操作指导。
3. 保存两边首次生成的 DOCX、对话/工具记录和运行耗时；计算候选文件 SHA-256。
4. 使用公共 LibreOffice/Poppler 渲染候选所有页面。用插件的 `docx-render` 做结构化渲染可以；评审结论必须基于同一公共渲染配置生成的图。
5. 执行每个 YAML 指定的内容与结构核对项，再按 `scoring.md` 评分。不能只看 DSH 的 `verify` 返回值，也不能只看文字抽取。
6. 把记录填入 `report-template.md`，逐页图片和候选 DOCX 放入对应运行目录。没有实际执行的 runner 标记 `not_run`，不要预填分数。

### DSH 插件的导入/渲染

插件只允许读取配置工作区内的文件。可将每个输入/候选文件复制到专用测试工作区（不要授予整个磁盘），再按以下顺序调用：

1. `docx_import({ path: "case-input.docx" })`，保存返回的完整 `artifactRef`。
2. `docx_call({ moduleId: "docx-render", operation: "execute", input: { operation: "execute", requestId: "render-case-round-00", artifactRef } })`。
3. 保存 PDF、每页 `image` 和 `thumbnail` 产物。对 WorkBuddy 和 DSH 的基线/候选都用同一插件版本与同一配置渲染；该步骤只负责转换，不让 Agent 帮另一个 runner 评审或修复。

在 WorkBuddy 与 DSH 的工作目录里分别保留独立副本。将 WorkBuddy 结果交给公共渲染器时，只复制它交付的 DOCX 到测试工作区；不要在 DSH Agent 会话中让模型重新处理该文件。

## 两类结论要分开

- **首稿质量**：同一提示词下第一次交付是否正确、好看、完整。这是主比较指标。
- **完整产品体验**：完成同一任务耗时、额外配置、往返次数、是否能看页面并修正、是否需要用户搬运文件。这里测的是 WorkBuddy 与“DSH + 模型 + 插件”的整套体验，不是单独 DOCX 引擎的纯性能。

本测试包不预设哪边获胜。结构定位、批注/修订与可重复渲染是 DSH 插件特意覆盖的能力；版式观感仍以相同渲染器的实际页面为准。四个用例是小规模能力探针，不等同于大样本性能基准；要比较稳定性，建议日后固定环境各重复三次，并把每次运行分开存档。
