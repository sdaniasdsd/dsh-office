# DSH Office 0.9.3 — 修好随包运行时下必然崩溃的 docx-complex-parse

## 这次修的是什么

装好的插件里，`docx-complex-parse` 和 `docx_analyze(complex="structure")` 以前只回一句
`ENGINE_FAILED: Parse process exited with an error`，没有任何可诊断信息。

**实测根因**：随包发布的解释器是 Python embeddable 版，目录里带 `python313._pth`，
那等价于 isolated 模式 —— 脚本所在目录不会自动进入 `sys.path`，`PYTHONPATH` 也被忽略。
`docx_complex_parse.py` 里 `from observations import read_observations` 因此抛
`ModuleNotFoundError`，进程以非零码退出，连失败信封都发不出来，上层就只剩一句
「Parse process exited with an error」。

`._pth` 里原本有一条 `../../../lib/engines/docx-complex-parse`，想绕开这件事，但它从来
指不中：那条相对路径是相对 `python.exe` 所在目录解析的，而引擎随**核心包**走
（`lib/engines/<模块>/`），运行时是**另一棵树**。插件安装布局下 `python.exe` 在
`<pkg>/runtime/win32-x64/python/`，往上三级是运行时包自己 —— 不是插件包。

## 改了什么

加性、向后兼容，且不再依赖任何写死的相对路径：

- `docx-complex-parse` 引擎：`import observations` 之前显式把自己的目录插进 `sys.path`。
- `docx-easy-parse` 的 docling 驱动：同样处理 `import docx_parse`（同一个坑的另一个入口）。
- `scripts/fetch-dsh-runtime.ps1`（本仓库）与 `dsh-toolchain/scripts/fetch-runtime.ps1`：
  `._pth` 只留解释器自身路径，并写明「引擎目录不进 `._pth`，需要时由引擎自己插」。

## 验证回执

跑的都是**装好的** Python 3.13 embeddable 解释器 + 装好的引擎目录：

| 项目 | 结果 |
| --- | --- |
| 修复前（git HEAD 版本，无 BOM） | exit 1，`ModuleNotFoundError: No module named 'observations'` |
| 修复后 structure 模式 | exit 0，stdout 32192 字节合法 JSON |
| 修复后带版面模式 | exit 0，stdout `{"failure":{"reason":"runtime_missing", ... rdocx ...}}` |
| `docx_parse_docling` import | exit 0（兄弟模块已解析） |
| `test_observations.py` | 5 passed，exit 0 |
| `scripts/check-modules.mjs` | exit 0 |
| assemble + provenance | 4 families / 110 files / 0 changedSources |
| 装进 profile 后 `docx_analyze(complex="structure")` | ✓ `bridges=4`，`complexMode=structure` |
| 装进 profile 后 `docx_analyze(complex="layout")` | `ENGINE_UNAVAILABLE`：`The 'rdocx' package is required but not importable.` |

关键是最后两行：同一个调用以前是「崩溃 + 无信息」，现在是「structure 能用；layout 给出
明确的可诊断错误码」。`rdocx` 确实不在运行时里，这一条是**如实报告缺能力**，不是本次要修的东西。

## 装了哪个包

只装**核心包**。运行时内容没变（`rdocx` 缺失、三件二进制都在），继续用 0.9.1 那份运行时即可；
`._pth` 的修正只影响未来重新构建的运行时，对已装的这份无害（引擎现在自己管 `sys.path`）。

- 核心 `@deepseek-ai/dsh-docx@0.9.3`：本 Release 的 `deepseek-ai-dsh-docx-0.9.3.tgz`
- 运行时 `@deepseek-ai/dsh-docx-runtime@0.9.1`：用 [v0.9.1 的资产](https://github.com/sdaniasdsd/dsh-office/releases/tag/v0.9.1)，
  或按 [dsh-toolchain](https://github.com/sdaniasdsd/dsh-toolchain) 自己构建

## 已知未验证 / 未做

- 运行时的 `._pth` 修正**没有重新发布运行时包**，所以 v0.9.1 那份运行时资产里仍带那条死路径。
  它现在是死代码（引擎自己插 `sys.path`），不影响使用；下次重建运行时才会消失。
- `docx-complex-parse` 的**版面**路径（`rdocx`）依然不可用，本版未引入 `rdocx`。
- Python 引擎单测不在 `check-modules.mjs` 里（那只跑 `tsc` + `vitest`），这次是手工执行的。
  把 python 侧纳入自动检查是下一件事，不是这一版做的。
