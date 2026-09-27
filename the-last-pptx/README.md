# the-last-pptx — PPTX 家族工作区

PPTX 相关模块的**源码家目录**。

| 模块 | 做什么 |
| --- | --- |
| `pptx-office` | 幻灯片检查与抽取、按坐标替换文本、按标题/正文层级做受控排版；**不新建 deck**（建 deck 用 python-pptx，本模块随后做排版与校验） |

## 怎么用

```bash
cd ../dsh-office
pnpm install && pnpm assemble && pnpm provenance && pnpm test:modules
```

运行时不另装：随插件分发的 Python 里带 `python-pptx`（`dsh-office/runtime/win32-x64/python`），
引擎脚本是 `pptx-office/src/engine/pptx_bridge.py`，由 `scripts/build-dsh.mjs` 拷进包内。
