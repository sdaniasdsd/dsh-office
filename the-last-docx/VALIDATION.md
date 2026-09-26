# 验收记录

日期：2026-09-26，Windows 本地环境。结论：**结构能力和交付链整合测试通过；真实页面渲染、深度解析与生产规模验收未完成。** 不能把下列总数解释为所有 Word 功能已经覆盖。

## 可重复执行的检查

| 检查 | 本轮结果 |
| --- | --- |
| `npm run typecheck` | 根工程通过 |
| `npm test` | 3 个文件，10 项通过 |
| `npm run test:modules` | 八个模块全部类型检查通过，429 项测试通过、1 项跳过，进程退出码 0 |
| `npm run provenance` | 92 个模块 src 文件与原始目录逐字节摘要一致，changedSources 为空 |
| plugin-creator 的 validate_plugin.py | 插件清单和文件结构通过 |
| `npm audit --omit=dev --registry=https://registry.npmjs.org --json` | 本次查询生产依赖已知漏洞 0 |
| `npm install --package-lock-only --ignore-scripts --no-audit --no-fund` | 锁文件同步完成 |

合计 **439 项通过、1 项跳过**。默认 npm 镜像不提供 audit 接口，首次查询返回 404；只对审计命令显式使用官方 registry 后成功，没有修改全局 registry。

## 模块测试明细

| 模块 | 通过 | 跳过 | 证据范围 |
| --- | ---: | ---: | --- |
| docx-inspect | 65 | 0 | 真实 ZIP/格式、安全拒绝、契约与确定性 |
| docx-easy-parse | 94 | 1 | 真实轻量解析；Docling 合并逻辑含替身，真实 Docling 项跳过 |
| docx-parse | 129 | 0 | 双 IR、锚点映射、真实差异、修订记录、治理及简报 |
| docx-complex-parse | 89 | 0 | 真实 Python 结构观察；页面引擎使用确定性替身，不是实际 rdocx 精度证明 |
| docx-create | 28 | 0 | 创建、模板、边界及真实 parser/editor 联调 |
| docx-edit | 5 | 0 | 节点替换、批注、锚点失配拒绝、原生修订 |
| docx-render | 8 | 0 | 契约、视觉发现引用、替身引擎及缺依赖处理；不含实际排版 |
| docx-artifact | 11 | 0 | 来源一致性、预览与版本引用、JSON、错误及拒绝场景 |

docx-edit 的原有最小样例仍输出 `No default paragraph style!!` 警告，测试通过但不据此声称样例视觉完整；没有为了消除警告改动原模块内核。

## 插件层实际验证

1. 八个独立注册身份；轻量解析原身份保留；统一 registry 适配调用可用。
2. 根入口完整导出双 IR 修订、差异、门槛判定和摘要 API；原模块相应行为测试保留。
3. 同一真实生成文档经过 inspect、easy、dual、complex structure，产生精确节点桥接，再通过双 IR 目标编辑并重新解析。
4. 提交第一版 manifest、verify、提交第二版；第二版携带第一版旧解析证据会被拒绝。
5. 导入快照不随原文件变化；拒绝工作目录外导入、外部 URI、篡改字节及 Windows 目录联接逃逸。
6. 同版本相同内容幂等；不同内容冲突；并发提交仅一个成功；Windows 保留名文档编号不污染文件路径。
7. 调用方不能通过工具参数改程序路径或放宽受控预算；关闭后的 Profile 拒绝调用。
8. 官方 MCP SDK 内存客户端和实际 stdio 子进程客户端均成功连接，列出六个工具，并调用创建、验证错误返回。

这里测的是本插件的最小 registry 端口及 MCP 服务，不是未提供源码的企业 Profile 主工程，也不是桌面应用安装验收。

## 运行环境实测

最新 `npm run doctor` 在默认配置下返回：Node v24.19.0、Python 3.11.4、lxml 6.1.0 可用；rdocx、Docling 未发现；LibreOffice、pdftoppm 未找到或无法启动。此前会话曾检测到 Poppler，但当前进程 PATH/配置无法启动它，因此以最新探测为准，不把历史发现算作当前可用。

doctor 是可用性探测，不是准确性或负载基准。已有外部程序时，可用 DOCX_SOFFICE、DOCX_PDFTOPPM 配置完整可执行路径后重新检查，不需要重复安装。

真实创建和交付示例已生成于：

`examples/generated/ef67d390-38d5-4625-9234-544f3051a77e/`

该目录的 `delivery-result.json` 含 DOCX、manifest 引用及校验报告。报告为 `ok: true, partial: true`；没有 PDF/逐页 PNG，视觉项未验证。示例只用于集成测试，不作为排版成品展示。

## 尚不能签收的内容

- LibreOffice + Poppler 的真实逐页渲染，以及不同字体环境的截断、重叠、分页与目录更新验证。
- 真正 rdocx 页坐标精度、Docling 模型效果及高开销能力的资源占用。
- 大规模复杂客户文档、生产并发/崩溃恢复、网络文件系统、跨操作系统验证。
- 硬资源沙箱、发布签名、完整软件物料清单和正式再分发许可证审核。
- 外部 Profile 主工程集成，以及应用市场安装/重装验收。

这些边界不会因结构测试通过而自动变成“已完成”。当前交付可以继续接入和试用；正式宣称完整视觉闭环前应补齐对应真实环境与样本文档验收。
