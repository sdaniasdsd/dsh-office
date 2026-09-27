# 依赖清单

这个仓库的依赖分三层：**npm 包**（模块与外壳）、**Python 包**（引擎脚本）、**外部运行时**（随包分发，不用 pip/npm 装）。
下面每条都写明**声明位置**；没被声明却在用的，单列在「缺口」里。

---

## 一、npm 包

### 1.1 外壳 `dsh-office/package.json`

| 包 | 版本 | 用途 |
| --- | --- | --- |
| `@modelcontextprotocol/sdk` | 1.30.1 | MCP 服务端（`src/server.ts`、`src/mcp.ts`） |
| `zod` | 3.25.76 | 各模块契约校验 |
| `zod-to-json-schema` | 3.25.0 | 把契约转成 MCP 工具可见的 JSON Schema |
| `@docx4j/core-ts` | 0.1.5 | OOXML 读写（create / styles / edit） |
| `@xmldom/xmldom` | 0.9.12 | 无浏览器环境的 DOM 实现 |
| `fflate` | 0.8.3 | zip 读写 |
| `pdfjs-dist` | 6.3.289 | 原生 PDF 只读检查与文本抽取 |
| `tsx` | ^4.20.5 | 直接跑 TypeScript 源码（`npm start` / `npm run demo`） |
| dev：`esbuild` | ^0.28.2 | 打包脚本 `scripts/build-dsh.mjs` 用它 bundle 出 `lib/*.mjs` |
| dev：`typescript` / `vitest` / `@types/node` | ^5.7.2 / ^2.1.9 / ^22.10.2 | 类型检查与测试 |

打包（`scripts/build-dsh.mjs`）时只有 **`pdfjs-dist` 保持 external** 并写进发布清单的 `dependencies`；
其余第三方被 esbuild 打进 `lib/server.mjs`、`lib/index.mjs`，所以不必出现在发布清单里。

esbuild 的平台二进制由可选依赖 `@esbuild/win32-x64` 提供（实测：`node_modules/@esbuild/win32-x64`
存在，lock 里记为 0.28.2）。npm 会提示 postinstall 未被 allow-scripts 放行，但二进制来自那个可选包，
打包照常可用（实测 `npm run build:dsh` 退出码 0）；pnpm 侧则需要 `pnpm-workspace.yaml` 里的
`onlyBuiltDependencies: [esbuild]`（已随仓库提交）。

### 1.2 各家族模块额外用到的

| 模块 | 额外依赖 | 声明位置 |
| --- | --- | --- |
| `docx-create` / `docx-styles` / `docx-edit` | `@docx4j/core-ts`、`@xmldom/xmldom` | 各模块 package.json |
| `docx-create` / `docx-artifact` | `zod`、`zod-to-json-schema`、`fflate` | 各模块 package.json |
| `xlsx-office` | `exceljs` 4.4.0、`@xmldom/xmldom`、`fflate` | `the-last-xlsx/xlsx-office/package.json` |
| `pdf-office` | `pdfjs-dist` 6.3.289、`zod` | `the-last-pdf/pdf-office/package.json` |
| `pptx-office` | **无 npm 依赖**（只走 Python） | `the-last-pptx/pptx-office/package.json` |
| `docx-*` 全部 | 模块之间用 `@dsh-office-profile/docx-*@*` 互相引用（workspace 内） | 各模块 package.json |

### 1.3 PDF 家族工作区 `the-last-pdf/`（不属于插件，需要时才装）

`pdf-engines` 额外声明：`pdfjs-dist` 6.3.289、`pdfkit` 0.20.2、`pdf-lib` 1.17.1、`@pdf-lib/fontkit` 1.1.1、
`@napi-rs/canvas` 1.0.9（原生二进制，平台相关）、`zod`；devDeps 里还有 `@types/pdfkit`。
其余 `pdf-*` 互相依赖仓库内的 `@dsh-office-profile/pdf-contracts` 与 `pdf-engines`。

### 1.4 仓库内共享包（不需要下载）

`dsh-office/packages/` 下五个：`office-core`、`office-files`、`office-safety`、`office-preview`、`office-test-kit`。
各模块把它们声明为 peerDependencies（`*`，optional），由外壳的 workspace 链接提供。

### 1.5 曾经的缺口（已修）

- **`esbuild` 曾被发现"在用但没声明"**：`scripts/build-dsh.mjs` 直接 `import { build } from 'esbuild'`，
  而当时所有 package.json 里都搜不到 esbuild；它能跑只是因为被 vite/vitest 提升到了 `node_modules`
  （版本 0.28.2）。清空 `node_modules` 或换机器后 `npm run build:dsh` 有可能直接失败。
  现在已在外壳 devDependencies 里声明 `"esbuild": "^0.28.2"`，锁文件同步更新，重建验证通过。

---

## 二、Python 包

### 2.1 声明的

| 文件 | 内容 | 谁用 |
| --- | --- | --- |
| `the-last-pptx/pptx-office/requirements.txt` | `python-pptx==1.0.2` | pptx-office 引擎 `src/engine/pptx_bridge.py` |
| `the-last-pdf/requirements-docling.txt` | `docling==2.130.0`、`docling-core`、`docling-parse`、`rapidocr`、`onnxruntime`、`pypdfium2`、`torch==2.14.0`、`torchvision`、`transformers`、`huggingface-hub` | PDF 家族的深度解析路线（很重，按需装） |

装配后 `dsh-office/modules/pptx-office/requirements.txt` 是同一份内容。

### 2.2 随包运行时里实际装了什么（实测回执）

用已安装插件自带的解释器逐个 `import`（`runtime/win32-x64/python/python.exe`，Python **3.13.15**）：

| 包 | 状态 |
| --- | --- |
| `lxml` | present 6.1.0（docx-* 引擎的 XML 后端） |
| `pptx`（python-pptx） | present 1.0.2 |
| `PIL`（Pillow） | present 12.3.0 |
| `XlsxWriter` | present 3.2.9 |
| `typing_extensions` | present |
| `oletools` | **MISSING** —— 宏语义分析（`analyzeMacroCode`）的可选增强，代码里 try/except 降级 |
| `docling` | **MISSING** —— `docx-easy-parse` 的深度路线（按需） |
| `rdocx` | **MISSING** —— `docx-complex-parse` 的页面坐标（按需） |

也就是说：**DOCX/PPTX/XLSX 的默认路径开箱可用**；宏语义分析、深度解析、复杂页坐标属于可选增强，
缺包时对应的能力会走降级路径，而不是报错崩溃。

### 2.3 引擎脚本里 import 的第三方模块（扫描结果）

`lxml`（9 个文件）、`oletools`（4 个，docx-inspect）、`rdocx`（2 个，docx-complex-parse，测试里有 stub）、
`docling`（8 个，docx-easy-parse 深度路线）、`pptx`（8 个，pptx-office）、`docx`（12 个，基准/脚本用 python-docx 造样例）、
`PIL`（2 个，维护脚本 `scripts/prepare-stress-visuals.py`）。

---

## 三、外部运行时（单独一个包，不再是硬依赖）

由 `dsh-office/scripts/fetch-dsh-runtime.ps1` 下载并校验，装到 `runtime/win32-x64/`：

| 组件 | 版本 | 来源 |
| --- | --- | --- |
| Python | 3.13.15（embed amd64） | python.org 官方 zip |
| LibreOffice | 26.8.0 | 官方 MSI，管理式解包取 `program/soffice.com` |
| poppler | 26.09.0 | `oschwartz10612/poppler-windows` release（取 `pdftoppm`） |

**发布形态已经从"一个包"拆成三个**（`npm run build:dsh` + `npm run pack:dsh` 一次产出）：

| 包 | 内容 | 体积（实测 0.9.0） |
| --- | --- | --- |
| `@deepseek-ai/dsh-docx` | 核心：`dsh/`、`lib/`（bundle）、第三方许可、文档 | tarball **2.2 MB**（解压 14 MB） |
| `@deepseek-ai/dsh-docx-runtime` | `runtime/win32-x64/**` + `runtime.json` 清单 | tarball **546 MB**（解压 1.67 GB） |
| `@deepseek-ai/dsh-docx-full` | 只写依赖的元包（核心 + 运行时，同版本） | tarball **651 B** |

核心包找运行时的顺序（`dsh/runtime-config.mjs`）：插件配置 `runtimeRoot` → 环境变量 `DSH_OFFICE_RUNTIME_ROOT`（或 `DOCX_PYTHON/DOCX_SOFFICE/DOCX_PDFTOPPM`）→ **兄弟运行时包**（向上找 `node_modules/@deepseek-ai/dsh-docx-runtime`，读 `runtime.json`）→ 自带 `runtime/win32-x64`。
缺运行时**不再是启动失败**：只把存在的路径注入子进程，缺什么写进 stderr 与 `DSH_DOCX_RUNTIME_MISSING`，相关能力在调用时返回 `ENGINE_UNAVAILABLE`，`docx_doctor` 逐项给结论。

- 发布包解压后约 1.7 GB，其中 LibreOffice 占 1504 MB、poppler 121 MB、python 47 MB（实测）。
- **仓库里的 `dsh-office/runtime/win32-x64` 是空的**（实测只有 1 个条目）——新克隆必须先跑一次 fetch 脚本，否则 `build:dsh` 会**跳过**运行时包并打印原因（不再像旧版那样在最后一步 `ENOENT` 崩掉）。

---

## 四、DSH 安装侧的依赖

| 项 | 值 | 位置 |
| --- | --- | --- |
| 插件包 | `@deepseek-ai/dsh-docx`（file: 指向 tgz） | DSH profile 的 `package.json` |
| MCP 客户端 | `@deepseek-ai/dsh-mcp-client` `>=0.1.0-rc.8 <1` | 发布清单 dependencies |
| 宿主核心 | `@deepseek-ai/cordis` `^4.0.1`（peer，复用宿主） | 发布清单 peerDependencies |
| Node | `^22.19.0 \|\| >=24.0.0`（发布清单）；开发用 20+ | 发布清单 engines |

安装：把 profile 里 `@deepseek-ai/dsh-docx` 指向 `deepseek-ai-dsh-docx-<版本>.tgz` → 装依赖 → 重启 DSH。

---

## 五、新机器上的完整步骤

```bash
cd dsh-office
npm install                                        # 或 pnpm install
powershell -File scripts/fetch-dsh-runtime.ps1     # 需要渲染 / 宏检测时；约 1.8 GB
npm run assemble                                   # 从四个家族装配 modules/
npm run provenance                                 # 校验 modules/ 与家族源码一致
npm run typecheck && npm test && npm run test:modules
npm run build:dsh && npm run pack:dsh              # 产出 dist/deepseek-ai-dsh-docx-<版本>.tgz
```

## 六、版本从哪来

- 各模块与外壳：各自的 `package.json`。
- **插件版本号单一来源**：`dsh-office/scripts/build-dsh.mjs` 的 `VERSION`，它写进发布清单并注入 MCP 握手版本。
- 发布清单的运行时依赖只有两条：`@deepseek-ai/dsh-mcp-client`、`pdfjs-dist`（唯一 external）。
- 发布的 tgz **不含模块源码**，只含打包结果 `lib/`、`dsh/`、`runtime/`、`third-party-licenses/` 与几份文档。
