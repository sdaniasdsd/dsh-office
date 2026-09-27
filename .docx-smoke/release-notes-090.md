# DSH Office 0.9.0 —— 运行时独立成包

源码：`main@f71d522`。这个版本把原来一个 548 MB 的包**拆成三个**：核心 2.2 MB、运行时 546 MB、元包 651 B。

## 三个资产（按需选一个或两个）

| 资产 | 是什么 | 字节 | sha256 |
| --- | --- | --- | --- |
| `deepseek-ai-dsh-docx-0.9.0.tgz` | **核心包**：JS bundle + dsh 清单 + 文档，不含运行时（纯 Node/npm） | 2,345,103 | `487a4a070ea51cf304e37ae0c0e28afdf672d9c36fc030996e9af4520b201d54` |
| `deepseek-ai-dsh-docx-runtime-0.9.0.tgz` | **运行时包**：Windows x64 的私有 Python 3.13 + LibreOffice + Poppler，附 `runtime.json` 清单 | 572,864,213 | `5e128b92c9e8a391dddf2e519d63f4c1497b1c7b5538ce095798e8a186df2b36` |
| `deepseek-ai-dsh-docx-full-0.9.0.tgz` | **元包**：只写依赖（核心 + 运行时，同版本），想一步装好就用它 | 651 | `66601c23ff2b23ce6a9f824921c38fa5cff0178e39413a4ea4c9ad69132f66d2` |

运行时包的组件实测（来自它自己的 `runtime.json`）：

| 组件 | 入口 | 体积 | 文件数 |
| --- | --- | --- | --- |
| python | `runtime/win32-x64/python/python.exe` | 50,013,265 B | 665 |
| libreoffice | `runtime/win32-x64/libreoffice/program/soffice.com` | 1,577,413,569 B | 19,456 |
| poppler | `runtime/win32-x64/poppler/poppler-26.09.0/Library/bin/pdftoppm.exe` | 126,544,679 B | 561 |

## 装法

```powershell
# 一步装好（核心 + 运行时）
dsh plugin --profile web add 'D:\交付目录\deepseek-ai-dsh-docx-full-0.9.0.tgz'

# 或者分开装：只装核心（省 546 MB），要用渲染/解析再补运行时
dsh plugin --profile web add 'D:\交付目录\deepseek-ai-dsh-docx-0.9.0.tgz'
dsh plugin --profile web add 'D:\交付目录\deepseek-ai-dsh-docx-runtime-0.9.0.tgz'
```

> **注意**：元包的依赖键是 `@deepseek-ai/dsh-docx` 与 `@deepseek-ai/dsh-docx-runtime`，这两个包**没有发布到公共 npm**。
> 所以从 tarball 安装时，元包只有在两个子包能被解析到时才装得上——拿不准就分别装核心与运行时（如上第二段），
> 或用一条 profile 依赖同时给出两个 `file:` 路径。

## 没有运行时也不会挂

核心包按「插件配置 `runtimeRoot` → 环境变量 `DSH_OFFICE_RUNTIME_ROOT`（或 `DOCX_PYTHON`/`DOCX_SOFFICE`/`DOCX_PDFTOPPM`）
→ 兄弟运行时包（向上找 `node_modules/@deepseek-ai/dsh-docx-runtime` 读 `runtime.json`）→ 自带 `runtime/`」的顺序找运行时。
全都找不到时**照常启动**，把缺什么写进 stderr 与 `DSH_DOCX_RUNTIME_MISSING`，只有依赖运行时的能力会在调用时返回
`ENGINE_UNAVAILABLE`，`docx_doctor` 会逐项给结论。创建、样式、编辑、交付清单等纯 Node 能力不受影响。

实测（本机）：

| 配置 | create | parse | pptx | render |
| --- | --- | --- | --- | --- |
| 只有核心包，保留宿主 PATH | ✓ | ✓（用了系统 Python） | ✓ | ✗ `spawn soffice ENOENT` |
| 只有核心包，PATH 收成 System32（干净机器） | ✓ | ✗ `Unable to launch the parse interpreter "python"` | ✗ | ✗ |
| 核心 + 运行时包 | ✓ | ✓ | ✓ | ✓ 产出 PDF |

## 用自己的工具链（接手的人看这段）

- **想复用宿主已有的 LibreOffice / Python**：设 `DSH_OFFICE_RUNTIME_ROOT` 指向一个含 `win32-x64/` 的目录，
  或用插件配置 `runtimeRoot`（也可分别给 `pythonPath` / `sofficePath` / `pdftoppmPath` 绝对路径）。这样就不用装 546 MB 的运行时包。
- **想自己构建运行时包**：在 `dsh-office/` 下跑 `pwsh -File scripts/fetch-dsh-runtime.ps1`（下载并校验
  python-3.13.15 embed、LibreOffice 26.8.0 MSI、poppler 26.09.0，装到 `runtime/win32-x64/`），
  再 `npm run build:dsh`（会产出核心 + 运行时 + 元包，并统计 `runtime.json`）、`npm run pack:dsh`（打成三个 tgz）。
  也可以 `node scripts/build-dsh.mjs --runtime=<已有一份运行时的目录>` 直接吃现成的。
- 运行时包只声明 `os: ['win32'] / cpu: ['x64']`；核心包不再限定平台。

## 已知问题

- `docx_doctor` 的 `xlsx` 会报 **false**，是假阴性：exceljs 被打进 bundle、没有单独安装，探测不到它。
  实际 `xlsx-office` 模块可用。修它需要一次依赖模型取舍（声明为运行时依赖并 external，或改由打包事实推导能力）。
- `docling`（deepEasyParse）与 `rdocx`（complexLayout）**不在**运行时包里，按需另装，缺包时明确报不可用、不冒充。
