# 整合设计与兼容边界

## 保留内核，只在接口之间做适配

九个模块保留自己的 contract、engine、mapper、verifier、errors、telemetry 和测试。插件宿主负责装配、共享文件服务和显式调用顺序，不把整条处理链塞进某个能力模块。

`docx-styles` 按成熟 OOXML 工具链的部件划分独立成模块（对照 docx4j 的 `StyleDefinitionsPart`、Apache POI 的 `XWPFStyles`、Open XML SDK 的 `StyleDefinitionsPart`）：它只拥有 `styles.xml`，不碰文档内容。`docx-edit` 负责写 `w:pStyle` 这类**引用**，`docx-styles` 负责**定义**。两者拆开后，顺序由 Profile 显式编排——**先定义、后引用**，中间态只含"尚未被引用的样式"，合法且无悬空引用。这不是把同一次写包拆成两次，而是把「定义」与「引用」两类事实分开授权。

`modules.lock.json` 记录源文件摘要及比对结果。模块源码不因整合改写；包名、工作区依赖和 tsconfig 的变化用于解决安装及类型声明冲突。原始模块目录保留，新的 docx-artifact 同时提供独立源码和插件副本。

借鉴的是成熟社区的边界设计，不是照搬其运行时：依赖注入参考 [Backstage Services](https://backstage.io/docs/backend-system/architecture/services/)，扩展模块参考 [Backstage Modules](https://backstage.io/docs/next/backend-system/building-plugins-and-modules/index/)。交付描述中的摘要、大小、媒体类型借鉴 [OCI Descriptor](https://github.com/opencontainers/image-spec/blob/main/descriptor.md)，但 manifest 是自己的 `docx-delivery/v1`，不声称符合 OCI。Agent 传输使用 [MCP 官方 TypeScript SDK](https://ts.sdk.modelcontextprotocol.io/server)，没有另造一套协议。

## 三处实际冲突及处理

1. **两个解析包重名。** 轻量解析包的工作区名改成 `@dsh-office-profile/docx-easy-parse`，外层注册名同样区分；原 definition/output 的身份信息保留。`list()` 同时展示 registryId 和原 definition，`registerWith()` 仅在注册适配对象中改 definition.id。不重写原能力名或结果。若外部 registry 按 capability 名全局去重，还须由该 registry 按模块命名空间适配，不能假定只按 id 去重。
2. **复杂解析的节点不等于编辑目标。** `source-bridge.ts` 先校验同一源文件 id + sha256，再用 part + structuralPath 对齐双 IR sourceMap。仅接受精确命中；缺失项明确记录，不用文字相似、顺序或页坐标猜节点。复杂解析的页面编号与渲染编号也不偷偷互换。
3. **写入端口 source/sources 不同。** 宿主文件服务兼容两种既有形状，创建允许无源文件，编辑/渲染保留单来源。产物统一不可变；不要求所有模块为此改公开契约。

## 双 IR 中间件没有丢失

`src/index.ts` 的 `docxParse` 命名空间完整导出原包，包括 `createRevisionLedger`、`diffDualIR`、`irFingerprint`、`decideGovernance`、`renderAgentBrief`。原修订及治理测试一起执行。

这些 API 仍由上层显式使用，不添加第四个模块 capability，不修改原门槛规则。插件不自动为每次 MCP 调用写入治理历史，不把 manifest 的文件版本号当成用户意图的修订记录。需要跨会话记忆时，调用方应注入原 ledger 的持久化 store，并显式保存意图与前后 IR。

## 公共包的真实定位

`packages/office-core`、`office-safety`、`office-files`、`office-test-kit` 从既有共享声明提升为单一类型来源，避免各模块各自声明同名类型；`office-preview` 引用原渲染契约。它们是当前交付的契约包，不是虚构一套已上线的企业 Office 服务。

真实本地文件实现是 `src/files.ts`。如果整合进另一个已经拥有 office-* 包的工程，应由该工程选择唯一公共契约版本，并适配文件、registry 与生命周期端口，不能同时安装两套不兼容定义。本地交付已提供最小 registerModule 对接测试，但没有访问或修改未指定的上层工程。

## 文件和交付规则

- 导入先限制工作目录并解析 realpath，防止目录联接逃逸；复制快照，不修改原件。
- 内容标识是 `sha256:<摘要>`，每次读取复核内容、大小和摘要。路径只允许托管的 file: URI；不自动访问网络。
- 新产物按摘要保存；版本清单在 `versions/doc-<documentId>/<revision>.json`。先写临时文件并同步，再通过同目录独占硬链接发布。已有相同内容为幂等成功，不同内容报 VERSION_CONFLICT。
- 后续版本必须引用已提交的相邻上一版摘要，没有可被争抢覆盖的可变 HEAD。要求文件系统支持硬链接；不支持时明确失败，不退回不安全覆盖写入。已测试本机 Windows 文件系统，不宣称跨网络文件系统一致性。
- 超时不等于回滚：提交完成但确认丢失时，以相同内容重试核实。临时文件通常在 finally 清理；进程被强制终止可能留下 `.pending-*`，不会被当作正式产物，当前不自动做垃圾回收。
- manifest 的 DOCX、预览及证据全部绑定实际字节。证据信封是 `{schema:'docx-evidence/v1',kind,source:{id,sha256},payload}`；原 payload 不被改写成另一种 IR。
- parent manifest 逐级校验文档标识、revision、摘要和引用预算。没有预览或视觉发现时保留部分验证。图像/PDF 只做格式签名与引用检查，不声称深度解码或自动看懂页面。

## 安全与验证不是一回事

默认写入/渲染/交付预检拒绝宏、外部关系、嵌入对象、altChunk、加密；模块自身可以有更严格的限制。工具输入不能改程序路径和预算，不能关闭 enforceLimits。可信宿主配置仍须遵循每个模块的 schema。

文件完整性只能证明“引用指向这些字节”，不是数字签名、作者身份认证或不可抵赖。视觉 reviewed 只表示承接了调用方明确提供的 findings；空 findings 可以表示调用方已检查未发现问题，但不是插件自己证明了视觉完美。

本地程序运行在当前用户权限下，不是恶意文档的强隔离沙箱；同权限进程仍可能修改文件或制造竞争。哈希复核、realpath 与预算是防错和防护层，不取代 OS 沙箱。上线处理不可信文档需宿主配置独立低权限进程、CPU/内存/时间限制和字体环境。

宿主默认单文件读取 64 MiB；docx-artifact 默认 manifest 2 MiB、总引用数据 512 MiB、引用数 1000、页面数 300、超时 30 秒。模块调整预算不能绕过宿主的读取上限；需要更大文件时应同时评估并调整宿主。进程内编辑/创建的同步计算是软期限，不是强杀保障。

## 渲染问题如何回流

解析负责内容、结构和稳定锚点；渲染负责让排版程序产生页面并承接视觉观察。看到问题后由 Profile/Agent 将发现与结构证据对照：文字或结构错误交给编辑，页边距/字体/分页问题交给相应布局参数。修改后重新解析、重新渲染、重新关联交付证据。

不能把渲染中的溢出直接归结为解析失败，也不能因为解析通过就跳过视觉检查。本版不自动修文档、不自动更换字体，也不新增模糊定位修复逻辑。
