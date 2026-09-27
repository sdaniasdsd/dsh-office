# the-last-xlsx — XLSX（表格）家族工作区

表格相关模块的**源码家目录**。

| 模块 | 做什么 |
| --- | --- |
| `xlsx-office` | 工作簿检查与结构抽取、读写单元格、打印版式安全修复；宏文件不支持，公式不重算 |

## 怎么用

```bash
cd ../dsh-office
pnpm install && pnpm assemble && pnpm provenance && pnpm test:modules
```

依赖 `exceljs`、`@xmldom/xmldom`、`fflate`，由外壳的 workspace 统一安装。
