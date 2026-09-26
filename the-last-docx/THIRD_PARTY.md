# 依赖和再分发说明

本文件记录当前锁文件及安装包元数据，不代替许可证正文或法律审核。项目自身尚未选择对外发布许可证；package.json 为 private，不应擅自标成 MIT 或 Apache 开源发布。

## npm 运行依赖

| 组件 | 验证版本 | 包声明许可证 | 用途 |
| --- | --- | --- | --- |
| @docx4j/core-ts | 0.1.5 | Apache-2.0 | 创建、编辑 OOXML |
| @xmldom/xmldom | 0.9.12 | MIT | XML 处理 |
| fflate | 0.8.3 | MIT | ZIP 数据 |
| @modelcontextprotocol/sdk | 1.30.1 | MIT | MCP 服务与协议 |
| tsx | 4.23.15 | MIT | 本地 TypeScript 启动 |
| zod | 3.25.76 | MIT | 输入与配置校验 |
| zod-to-json-schema | 3.25.0 | ISC | 公共 schema |

精确传递依赖以 package-lock.json 为准。构建脚本会从实际打入 JS bundle 的 Node 包中收集顶层 LICENSE、LICENCE、NOTICE 与 COPYING 文件至 `third-party-licenses/`；仍需结合本表核对传递依赖和许可证。类型检查与测试另外使用 TypeScript、Vitest 和 Node 类型声明。npm audit 的零漏洞结果只代表查询时数据库未报告已知项，不是供应链或运行安全保证。

## 随包运行时

本 Windows x64 DSH 交付包计划随包提供以下运行时；用户无需单独安装、配置 PATH 或在首次运行时联网下载：

| 组件 | 固定版本 | 用途 | 许可/分发说明 |
| --- | --- | --- | --- |
| Python embeddable | 3.13.15 | 隔离运行 DOCX Python 引擎 | PSF-2.0；随包保留原始许可文件 |
| lxml | 6.1.0 | OOXML/XML 解析与桥接 | BSD-3-Clause；wheel 元数据与许可随包保留 |
| Poppler for Windows | 26.09.0-0 | PDF 页图转换 | 上游 GPL 条款；保留发行包中的 COPYING/NOTICE |
| LibreOffice (The Document Foundation MSI) | 26.8.0 | DOCX 转 PDF 排版 | 通过 Windows Installer 管理式抽取到插件目录；随包保留 LGPL 与其他许可/NOTICE |
| Microsoft Visual C++ Runtime | 14.x (由上述 MSI 附带) | LibreOffice 原生 DLL 运行依赖 | 原 DLL 不修改；分发受 Microsoft Visual Studio 许可条款约束，正式对外分发前单独核对 |

实际运行包由 `scripts/fetch-dsh-runtime.ps1` 下载并校验 SHA-256 后，抽取至 `dist/dsh-docx/runtime/win32-x64`。LibreOffice 仅执行 `msiexec /a` 管理式抽取，目标是插件自己的目录，不执行安装、注册表写入或全局配置。二进制按上游 MSI 原样分发，不修改其程序文件；对外公开再分发前仍须由发布方复核完整许可文本、所需源代码提供方式及各上游的再分发条件。构建期使用的 7-Zip 命令行程序不进入最终插件包。

Docling 和 rdocx 是可选深度解析/版面后端，原模块适配接口保留，但不在这个基础离线包中安装或声称可用；它们通常会额外引入模型或较重依赖。

本交付没有复制商业 Word 母版或商业字体；创建模块的三个基础样式为项目内原创配置。字体由用户系统提供；不同字体环境可能令分页略有差异。

不自动安装重型运行时，也不自动下载模型。检测可用性用 `npm run doctor`；真实能力验收以对应模块的实际文档测试为准。
