# DSH Office 高要求视觉编辑基准 — 50 例

## 执行结论

完成 50/50 个冻结的办公室日常用户式请求：DOCX 20、PPTX 15、XLSX 15。首轮 28 PASS / 22 FAIL，最终 28 PASS / 22 FAIL；首轮平均 62/100，最终平均 62/100。证据不把引擎拒绝或验证失败记成通过。

所有 50 条中文原始提示都原样冻结；但本仓库 Profile 没有自然语言规划模型，提示由测试脚本人工映射为 DSH typed calls，因此结果衡量模块/后端能力，不衡量模型自主理解。PDF 编辑未测：当前 pdf-office 仅只读。

## 主要工程问题

- DOCX 表格格式：6 例首次写表格报 ENGINE_FAILED（Center 对 JcEnumeration 非法）；省略对齐重试仍失败，说明修复还需定位到表格引擎/校验边界，而非单纯映射字段。
- DOCX 段落写入后校验：5 例报通用 VERIFICATION_FAILED；缩减到粗体/颜色后仍失败且无属性级诊断，缺少可操作校验差异。
- DOCX 分页回归：D03 首轮从 9 页增至 10 页，多出仅含表格末条的近空白末页；将同四段行距/段后距收紧后回到 9 页（round 01），但首轮仍判 FAIL。
- PPTX：10 个指定文本替换请求通过；5 个主题级字体/字号/色彩请求被 pptx-office 明确拒绝（当前仅支持 extract / replaceText）。
- XLSX：10 个打印宽度布局请求通过；5 个单元格样式请求被 xlsx-office 明确拒绝（当前没有 formatCells）。打印版式仍须在真实 Excel 验证；LibreOffice PDF 仅作一致的审阅渲染。

## 分例首轮和最终分数

| 案例 | 格式 | 首轮 | 最终 | 修复轮 | 未通过硬门 |
|---|---|---:|---:|---:|---|
| D01 | DOCX | PASS 93 | PASS 93 | 0 | — |
| D02 | DOCX | FAIL 20 | FAIL 20 | 1 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| D03 | DOCX | FAIL 66 | FAIL 65 | 1 | one_step_completion, score_threshold |
| D04 | DOCX | FAIL 20 | FAIL 20 | 1 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| D05 | DOCX | PASS 93 | PASS 93 | 0 | — |
| D06 | DOCX | FAIL 20 | FAIL 20 | 1 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| D07 | DOCX | PASS 93 | PASS 93 | 0 | — |
| D08 | DOCX | FAIL 20 | FAIL 20 | 1 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| D09 | DOCX | PASS 93 | PASS 93 | 0 | — |
| D10 | DOCX | FAIL 20 | FAIL 20 | 1 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| D11 | DOCX | PASS 93 | PASS 93 | 0 | — |
| D12 | DOCX | FAIL 20 | FAIL 20 | 1 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| D13 | DOCX | FAIL 20 | FAIL 20 | 1 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| D14 | DOCX | PASS 93 | PASS 93 | 0 | — |
| D15 | DOCX | FAIL 20 | FAIL 20 | 1 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| D16 | DOCX | FAIL 20 | FAIL 20 | 1 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| D17 | DOCX | PASS 93 | PASS 93 | 0 | — |
| D18 | DOCX | FAIL 20 | FAIL 20 | 1 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| D19 | DOCX | PASS 93 | PASS 93 | 0 | — |
| D20 | DOCX | FAIL 20 | FAIL 20 | 1 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| P01 | PPTX | PASS 93 | PASS 93 | 0 | — |
| P02 | PPTX | PASS 93 | PASS 93 | 0 | — |
| P03 | PPTX | FAIL 20 | FAIL 20 | 0 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| P04 | PPTX | PASS 93 | PASS 93 | 0 | — |
| P05 | PPTX | PASS 93 | PASS 93 | 0 | — |
| P06 | PPTX | FAIL 20 | FAIL 20 | 0 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| P07 | PPTX | PASS 93 | PASS 93 | 0 | — |
| P08 | PPTX | PASS 93 | PASS 93 | 0 | — |
| P09 | PPTX | FAIL 20 | FAIL 20 | 0 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| P10 | PPTX | PASS 93 | PASS 93 | 0 | — |
| P11 | PPTX | PASS 93 | PASS 93 | 0 | — |
| P12 | PPTX | FAIL 20 | FAIL 20 | 0 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| P13 | PPTX | PASS 93 | PASS 93 | 0 | — |
| P14 | PPTX | PASS 93 | PASS 93 | 0 | — |
| P15 | PPTX | FAIL 20 | FAIL 20 | 0 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| X01 | XLSX | PASS 93 | PASS 93 | 0 | — |
| X02 | XLSX | PASS 93 | PASS 93 | 0 | — |
| X03 | XLSX | FAIL 20 | FAIL 20 | 0 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| X04 | XLSX | PASS 93 | PASS 93 | 0 | — |
| X05 | XLSX | PASS 93 | PASS 93 | 0 | — |
| X06 | XLSX | FAIL 20 | FAIL 20 | 0 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| X07 | XLSX | PASS 93 | PASS 93 | 0 | — |
| X08 | XLSX | PASS 93 | PASS 93 | 0 | — |
| X09 | XLSX | FAIL 20 | FAIL 20 | 0 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| X10 | XLSX | PASS 93 | PASS 93 | 0 | — |
| X11 | XLSX | PASS 93 | PASS 93 | 0 | — |
| X12 | XLSX | FAIL 20 | FAIL 20 | 0 | fileIntegrity, contentPreservation, intentComplete, score_threshold |
| X13 | XLSX | PASS 93 | PASS 93 | 0 | — |
| X14 | XLSX | PASS 93 | PASS 93 | 0 | — |
| X15 | XLSX | FAIL 20 | FAIL 20 | 0 | fileIntegrity, contentPreservation, intentComplete, score_threshold |

## 证据和复现

逐例的冻结提示、操作计划、源/候选 SHA-256、DSH 验证输出、轮次提示、分数、门禁、缺陷及每页图像路径均在 cases/{ID}/。首轮与修复轮截图按页保存在 dsh/round-00/、dsh/round-01/。所有材料源自既有冻结的官方来源语料，源文件哈希运行前后不变。图像 120 DPI，PPTX/XLSX 由同一捆绑 LibreOffice + Poppler 审阅，DOCX 使用 DSH docx-render。

摘要 JSON：final-report.json。测试计划和 runner：cases.json、scripts/office-aesthetic-50.mjs、scripts/office-aesthetic-50-repair-docx.mjs。
