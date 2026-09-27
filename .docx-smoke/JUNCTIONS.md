# 这个目录是什么

`D:\认真版agent\.docx-smoke\` 的原样归档：本轮 dsh-office 拆包 / 工具链适配 / 0.9.0–0.9.3
发布过程中用到的全部验证驱动脚本、提交信息草稿、发布说明与夹具数据。

归档时**只做了一处删除**，见下。其余按原样（含 `data*/` 夹具、`tmp/`、`.docx-data/`、
以及脚本里写死的本机绝对路径 `C:\Users\AA\...` / `D:\开源团队作品\...`）。

## 唯一没有进来的东西：3 个 junction

`split-check/node_modules/` 下有 3 个 NTFS 目录联接（junction），指向本机其它目录。
它们**没有被复制**（robocopy `/XJ`），因为其中一条指向构建出来的运行时整棵树
（约 1.67 GB / 约 2 万文件），跟进仓库会把归档从 42.85 MB 撑到 GB 级：

| 联接（相对本目录） | 目标 |
| --- | --- |
| `split-check/node_modules/pdfjs-dist` | `D:\开源团队作品\dsh-office-clone\dsh-office\node_modules\pdfjs-dist` |
| `split-check/node_modules/@deepseek-ai/dsh-docx-runtime` | `D:\开源团队作品\dsh-office-clone\dsh-office\dist\dsh-docx-runtime` |
| `split-check/node_modules/@deepseek-ai/dsh-mcp-client` | `D:\开源团队作品\dsh-office-clone\dsh-office\node_modules\@deepseek-ai\dsh-mcp-client` |

要复现 `split-check` 的夹具布局，把上面三条联接按原样重建即可（目标都在本仓库内，
`dsh-docx-runtime` 那份需要先 `npm run build:dsh` 或 `npm run fetch:toolchain`）。

## 规模回执

- 归档文件数：2160 个文件 / 459 个目录 / 42.85 MiB（= 源目录去掉上述 3 个联接）
- 源目录实测（robocopy `/L /XJ`）：2160 文件 / 462 目录 / 42.85 MiB
- 两侧数字一致，除联接外无遗漏
