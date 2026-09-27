# DSH Office 0.8.1

一个小补丁：**让插件的 MCP 握手报出它真实的版本号**。

## 修的是什么

安装好的 0.8.0 在握手时自称 `0.5.0`，而它的 `package.json` 是 `0.8.0`：

```js
new McpServer({ name: "the-last-docx", version: "0.5.0" })   // 0.8.0 包里的实际内容
```

原因是两处不一致：`scripts/build-dsh.mjs` 已经用 `define` 把版本注入为 `__DSH_DOCX_VERSION__`，
但 `src/mcp.ts` 那一行仍写死字符串，注入的常量没人读。这不是显示问题——任何按握手版本判断
"该不该升级"的自动化都会被它骗过去。

现在 `src/mcp.ts` 读注入的常量；从 TypeScript 源码直接跑（没有任何东西定义它）时，报
`0.0.0-dev`，而不是假装成一个已发布版本。

## 变更

- `src/mcp.ts`：握手版本改读 `__DSH_DOCX_VERSION__`，附上为什么不能硬编码的说明
- `scripts/build-dsh.mjs`：版本号 → `0.8.1`
- 提交：`e2c1f32`（在这条提交的独立 worktree 里复跑：`tsc --noEmit` 退出码 0、根套件 32 passed）

## 包内容

与 0.8.0 相同的能力面：九个 DOCX 模块 + `pptx-office` + `xlsx-office`，随包运行时
（Python 3.11、LibreOffice、poppler）。注册的工具：
`docx_modules, docx_doctor, docx_import, docx_call, pptx_call, xlsx_call, docx_analyze, docx_read_artifact, docx_from_reference`。

包由 `e2c1f32` 的**干净 worktree** 构建，而不是工作区——工作区当时带着另一条产品线未提交的改动，
从那里构建会发出仓库里并不存在的代码。运行时取自已安装的 0.8.0 包，与该版本发布的内容一致。

## 安装

```bash
pnpm add "@deepseek-ai/dsh-docx@file:/绝对路径/deepseek-ai-dsh-docx-0.8.1.tgz"
```

安装后重启 DSH，MCP 客户端会拉起新进程；握手应显示 `serverInfo.version = 0.8.1`。
