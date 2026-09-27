# the-last-docx — DOCX 家族工作区

九个 DOCX 模块的**源码家目录**（每个模块一个包，包名 `@dsh-office-profile/docx-*`）：

| 模块 | 做什么 |
| --- | --- |
| `docx-inspect` | 输入探测与安全读取（OOXML 结构、宏、外链、嵌入对象） |
| `docx-easy-parse` | 保留旧版输出形态的简化解析（legacy IR） |
| `docx-parse` | 结构化解析：语义视图 + 物理视图 + 源映射（dual IR） |
| `docx-complex-parse` | 复杂版面：逻辑表格栅格、分页关系、浮动对象、置信度 |
| `docx-create` | 按计划新建文档（场景 → 装帧寄存器、表格、编号、页面几何） |
| `docx-styles` | styles.xml 的命名样式编排与解析 |
| `docx-edit` | 就地编辑（修订、批注、段落/表格格式），未触及部分字节不变 |
| `docx-render` | 渲染成 PDF 与逐页图片，供视觉复核 |
| `docx-artifact` | 版本化交付与不可变清单 |

## 怎么用

这个目录只放源码；依赖与链接由外壳统一管理：

```bash
cd ../dsh-office
pnpm install          # 把 @dsh-office-profile/* 链接进 node_modules
pnpm assemble         # 把四个家族复制进 dsh-office/modules/（装配产物）
pnpm provenance       # 逐文件校验 modules/ 与家族源码一致
pnpm test:modules     # 逐模块 typecheck + 测试
```

> 历史：这九个模块原在本机仓库外的 `docx分区` 里开发，再由 `assemble` 复制进仓库。
> 现在仓库内的本目录就是它们的家；仍想在外部目录开发时用
> `pnpm assemble --source docx=<外部路径>`（外壳里也留了 `assemble:docx-partition` 脚本）。
