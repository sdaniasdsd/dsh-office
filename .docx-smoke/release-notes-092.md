# DSH Office 0.9.2 —— doctor 不再无端把可用报成不可用

源码：`main@02c36e4`。本版只有诊断修正，**核心包之外的东西没变**：

- **运行时资产不重复上传**：内容与 0.9.1 那份逐字节相同，请用
  [`v0.9.1` 的 `deepseek-ai-dsh-docx-runtime-0.9.1.tgz`](https://github.com/sdaniasdsd/dsh-office/releases/tag/v0.9.1)
  （sha256 `35023910…364d`），或直接取工具链仓库
  [`dsh-toolchain v0.9.0`](https://github.com/sdaniasdsd/dsh-toolchain/releases/tag/v0.9.0) 的资产。
- 本 release 只附核心包 `deepseek-ai-dsh-docx-0.9.2.tgz`。

## 修了什么

`docx_doctor` 有两处把"可用"报成"不可用"，原因都在**探测方法**，不在能力本身：

1. **`xlsx` 恒为 false**：exceljs 是 esbuild **inline 进 `lib/server.mjs`** 的，磁盘上没有单独安装，
   子进程 `import("exceljs")` 必然失败。构建现在把"哪些 npm 包被打进 bundle"写成 `lib/bundled.json`
   （本次 85 个包，含 `exceljs 4.4.0`），doctor 读它并把 detail 写成来路：

   ```
   exceljs: { available: true,
              detail: "bundled into the server bundle (exceljs 4.4.0);
                       not installed separately, so a child-process import cannot see it" }
   ```

   从源码直接跑（没有该文件）时表为空，行为退回原样。

2. **`pptx` 偶发 false**：八条探测并发起，`import pptx`（连带 PIL/lxml）冷启动时会超过原来的 5 秒预算。
   python 侧五条探测的预算 5s → 15s（soffice 早前已改成 20s）。

顺带修 `scripts/fetch-toolchain.mjs`：解包时除了 `package/runtime/win32-x64/**`，也把包根的
`package/runtime.json` 取到 `<repo>/runtime.json`——之前漏了它，解出来的运行时不自描述，
用工具链的 `scripts/verify-runtime.mjs` 复核会少一项（其他 11 项本来就过）。补上后该脚本 12/12、退出码 0。
`runtime.json` 与 `runtime/` 一样不入库。

## 验收（本机实跑）

- `tsc --noEmit` 0；根测试 32 passed；12 模块 check-modules 退出码 0（510 passed / 2 skipped）；
  provenance 4 家族 / 110 文件 / 0 处不一致。
- 构建产物 doctor：`artifact / parse / pptx / xlsx / render / nativePdf / create / edit / complexStructure` 全 true，
  只剩真缺的 `complexLayout`(rdocx) 与 `deepEasyParse`(docling) 为 false。
- bundled 事实那条路径在**真实安装**里生效：装进 DSH profile 后 doctor 的 `xlsx` 为 true，detail 写明
  "bundled into the server bundle (exceljs 4.4.0)"。
- 用工具链验收脚本复核新取回的运行时（含补取的 runtime.json）：12 项全过。

## 装法（同 0.9.1，两步）

```powershell
dsh plugin --profile web add 'deepseek-ai-dsh-docx-0.9.2.tgz'          # 核心，2.2 MB
dsh plugin --profile web add 'deepseek-ai-dsh-docx-runtime-0.9.1.tgz'  # 运行时，546 MB（内容未变）
```

核心与运行时的版本不需要一致：核心只按 `runtime.json` 是否存在来找运行时，所以 0.9.2 的核心配 0.9.1 的运行时
（也就是 0.9.0 工具链那份内容）是正常用法。
