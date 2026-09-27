# dsh-office

包装并提供办公场景的一系列插件：DOCX / PDF / PPTX / XLSX 四个家族，加一个把它们装配成单一 DSH 插件的外壳。

## 仓库布局

```
the-last-docx/     DOCX 家族工作区：9 个模块（inspect / easy-parse / parse / complex-parse /
                   create / styles / edit / render / artifact）
the-last-pdf/      PDF 家族工作区：pdf-* 实现与文档，以及被插件注册的 pdf-office 模块
the-last-pptx/     PPTX 家族工作区：pptx-office（检查/抽取、按坐标替换文本、标题-正文层级排版）
the-last-xlsx/     XLSX（表格）家族工作区：xlsx-office（结构抽取、单元格读写、打印版式安全修复）
dsh-office/        插件外壳：src/ profile 与 MCP 入口、dsh/ 插件清单与 cordis patch、
                   packages/ 五个共享包（office-core/files/safety/preview/test-kit）、
                   tests/、scripts/、benchmarks/、reports/、runtime/，
                   以及 modules/ —— 由四个家族装配出来的产物
```

### 谁是谁的源码

- **家族目录是源码的家**：日常改代码、加测试都在 `the-last-{docx,pdf,pptx,xlsx}/` 里。
- **`dsh-office/modules/` 是装配产物**：`pnpm assemble`（或 `npm run assemble`）从四个家族复制过去，
  并补上打包需要的改名与 peer 依赖；`npm run provenance` 逐文件比对两边是否一致，
  结果写进 `dsh-office/modules.lock.json`。
- 外壳内部的布局保持 `modules/<name>/src/...` 不变，所以构建脚本、Python 引擎脚本拷贝、
  workspace 链接都不受家族目录位置影响。

## 怎么用

```bash
cd dsh-office
npm install            # 或 pnpm install（外壳的 workspace 是 modules/* 与 packages/*）
npm run assemble       # 从四个家族装配 modules/
npm run provenance     # 校验 modules/ 与家族源码逐文件一致
npm run typecheck
npm test               # 外壳的 profile / MCP / 文件层测试
npm run test:modules   # 12 个模块逐个 typecheck + 测试
npm run fetch:toolchain # 按 toolchain.lock.json 取运行时（下载 → 校验 sha256 → 解包 → 自检三件二进制）
npm run build:dsh      # 一次产出三个包到 dist/：核心 / 运行时 / 元包
npm run pack:dsh       # 打成三个 tgz（运行时不存在的构建只打核心包，并打印原因）
```

## 三个包（运行时是单独的包，产地是 dsh-toolchain）

| 包 | 内容 | 体积（0.9.1 实测） |
| --- | --- | --- |
| `@deepseek-ai/dsh-docx` | 核心：JS bundle + dsh 清单 + 文档，**不含运行时** | tarball 2.2 MB |
| `@deepseek-ai/dsh-docx-runtime` | Windows x64 运行时：私有 Python 3.13 + LibreOffice + Poppler，带 `runtime.json` 清单 | tarball 546 MB |
| `@deepseek-ai/dsh-docx-full` | 元包：依赖指向本仓库 release 的两个子包直链，想一步装好就用它 | tarball ~1 KB |

**运行时的产地不是本仓库**，而是 [`sdaniasdsd/dsh-toolchain`](https://github.com/sdaniasdsd/dsh-toolchain)：
`toolchain.lock.json` 钉住那一版工件的 tag / 资产名 / sha256 / 字节数，`npm run fetch:toolchain` 按它下载、
校验并解到 `runtime/win32-x64/`（`--check-only` 只自检），之后 `build:dsh` 把它按 `<pkg>/runtime/<platform>`
布局重新装进运行时包。换运行时版本 = 改 `toolchain.lock.json` 一个文件。

核心包按「插件配置 → 环境变量 → 兄弟运行时包 → 自带 runtime/」的顺序找运行时；**缺运行时不致命**：
插件照常启动，只有依赖运行时的能力会在调用时明确报不可用，`docx_doctor` 逐项给结论。
想复用宿主已装的 LibreOffice/Python，用 `runtimeRoot` 或 `DSH_OFFICE_RUNTIME_ROOT` 指过去即可。
契约（三条路径 + `runtime.json`）写在工具链仓库的 `docs/CONTRACT.md`。

只在某个家族里干活时，也可以先装配、再从 `dsh-office/modules/<name>` 跑那个模块的测试；
家族目录自身不放依赖（依赖由外壳统一安装）。

## 版本与发布

- 插件包名仍是 `@deepseek-ai/dsh-docx`（历史名，未随目录改名；改它会动到已安装 profile 的依赖键），
  版本号在 `dsh-office/scripts/build-dsh.mjs` 里由 `VERSION` 单一来源决定，并注入 MCP 握手版本。
- 发布：`node scripts/build-dsh.mjs` → `npm pack ./dist/dsh-docx --pack-destination ./dist`
  → 以 `v<版本>` 打成 GitHub Release 并附上 tgz。
- 安装到 DSH：把 profile 里 `@deepseek-ai/dsh-docx` 的依赖指向该 tgz 后 `pnpm install`，再重启 DSH。

## 文档

- `dsh-office/README.md`：插件外壳能用到什么程度、怎么启动。
- `dsh-office/ARCHITECTURE.md`：边界设计与借鉴来源。
- `dsh-office/OFFICE_ENGINE_DECISIONS.md`：引擎选型（PDF.js、python-pptx、ExcelJS 等）。
- `dsh-office/VALIDATION.md`：验收记录与未验收项。
- 各家族目录下的 `README.md`：该家族有哪些模块、怎么单独验证。
