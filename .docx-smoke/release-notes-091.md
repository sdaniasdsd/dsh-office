# DSH Office 0.9.1 —— 适配工具链交接

工具链已经有家了：[`sdaniasdsd/dsh-toolchain`](https://github.com/sdaniasdsd/dsh-toolchain)。
这个版本把**运行时的产地**从本仓库挪过去，本仓库只负责消费它。

源码：`main@9e41e44`（本说明里"元包"的修正见下方"更正"一节，在后续提交里）。

## 装法（两步；请按这个装）

```powershell
dsh plugin --profile web add '<核心>.tgz'      # deepseek-ai-dsh-docx-0.9.1.tgz          2.2 MB
dsh plugin --profile web add '<运行时>.tgz'    # deepseek-ai-dsh-docx-runtime-0.9.1.tgz  546 MB
```

也可以只装核心，用 `DSH_OFFICE_RUNTIME_ROOT` 或插件配置 `runtimeRoot` / `pythonPath` / `sofficePath` /
`pdftoppmPath` 指向你已有的 LibreOffice/Python；缺运行时不会让插件起不来，只有依赖运行时的能力会返回
`ENGINE_UNAVAILABLE`，`docx_doctor` 逐项给结论。契约见工具链仓库的 `docs/CONTRACT.md`。

## 资产

| 资产 | 是什么 | 字节 | sha256 |
| --- | --- | --- | --- |
| `deepseek-ai-dsh-docx-0.9.1.tgz` | **核心**：JS bundle + dsh 清单 + 文档，不含运行时 | 2,346,401 | `c27d8477…997a` |
| `deepseek-ai-dsh-docx-runtime-0.9.1.tgz` | **运行时**：Windows x64 的 Python 3.13 + LibreOffice 26.8.0 + Poppler 26.09.0 | 572,865,006 | `35023910…364d` |

## 更正：元包（`-full`）的直链依赖不成立，已从本 release 撤下

原本给元包写的依赖是这两个资产的**直链 URL**，想做到"一步装好"。实测**装不上**：

```
$ pnpm install                # 只依赖 -full.tgz
[ERR_PNPM_EXOTIC_SUBDEP] Exotic dependency "@deepseek-ai/dsh-docx" (resolved via url)
is not allowed in subdependencies when blockExoticSubdeps is enabled
```

pnpm 默认开启 `blockExoticSubdeps`，**子依赖里的 URL 依赖一律拒绝**，所以这条路对 DSH 的安装器不成立；
换成版本号依赖同样解析不到（两个子包没有发到 registry）。因此这个 release **不再附带 `-full` 资产**，
源码侧也把元包的依赖改回版本号，并在它的 README 里写明"需要 registry 才装得上"。
等两个子包发到 registry 时，元包就能直接当"一步装好"用。

## 这个版本的改动

- **运行时的产地**：`dsh-office/toolchain.lock.json` 钉住工具链工件（repo / tag `v0.9.0` / 资产名 / sha256 / 字节数）。
  `npm run fetch:toolchain` 按它下载 → 校验 sha256 与字节数（不一致拒绝解包）→ 解到 `runtime/win32-x64/` → 自检三件二进制；
  `--check-only` 只自检。
- **运行时包里的产地记录**：`runtime.json` 新增 `toolchain` 字段（repo/tag/asset/sha256），装完一眼能看出这份运行时从哪来。
- 本仓库的 `scripts/fetch-dsh-runtime.ps1` 加抬头说明：这份配方的家已搬到工具链仓库，本仓库保留它只用于"从上游重建"。

## 验收（本机实跑）

- `npm run fetch:toolchain`：从 `dsh-toolchain v0.9.0` 下载 572,864,213 B、sha256 与 pin 一致 → 解包 →
  自检 `Python 3.13.15` / `LibreOffice 26.8.0.3` / `pdftoppm 26.09.0` 全通过，退出码 0。
- 构建与打包：core 2,346,401 B / runtime 572,865,006 B。
- **内容同一性**：本包 `runtime.json` 的 components 与工具链那份逐项相同
  （python 50013265 B/665 files、libreoffice 1577413569 B/19456 files、poppler 126544679 B/561 files）。
  tgz 本身不是逐字节可复现的，所以两份 tgz 的 sha256 不同——比的是解出来的内容。
- `tsc --noEmit` 0；根测试 32 passed；12 个模块 check-modules 退出码 0（510 passed / 2 skipped）；
  provenance 4 家族 / 110 文件 / 0 处不一致。
- 三组真实调用（拆包行为回归）：

  | 配置 | create | parse | pptx | render | doctor |
  | --- | --- | --- | --- | --- | --- |
  | 只有核心包（保留宿主 PATH） | ✓ | ✓ | ✓ | ✗ `spawn soffice ENOENT` | render:false、nativePdf:true |
  | 只有核心包（PATH 收成 System32） | ✓ | ✗ | ✗ | ✗ | parse:false、pptx:false |
  | 核心 + 运行时包 | ✓ | ✓ | ✓ | ✓ 产出 PDF | render:true、nativePdf:true |
- **装进 DSH profile 后复验**：核心 0.9.1 + 运行时 0.9.1；握手 `{"name":"the-last-docx","version":"0.9.1"}`；
  运行时解析 `source=runtime-package`、无缺失；`docx_doctor` 报 `create/parse/pptx/render/nativePdf` 全 true；
  真跑 create（5958 B 产物）/ parse / render（产出 PDF）全部成功。

## 已知问题

- `docx_doctor` 的 `xlsx` 报 false 是假阴性：exceljs 被打进 bundle、没有单独安装，探测不到；实际 `xlsx-office` 可用。
- `docling`（deepEasyParse）与 `rdocx`（complexLayout）不在运行时包里，按需另装，缺包时明确报不可用。
- 换 poppler 版本要注意：目录名 `poppler-26.09.0` 被插件侧默认路径写死，需同时改插件侧或显式给 `pdftoppmPath`。
