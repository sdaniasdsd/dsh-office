# pdf-easy-parse

用 PDF.js 按页提取原生文字片段并投影成轻量结果。

```ts
createPdfEasyParseModule({ files })
```

必须连接 `ArtifactReader`，输入以实际字节重新计算 SHA-256。空文本页保留页码，覆盖标记为 partial；本模块不运行 OCR，不猜段落或阅读顺序。图片扫描件、格式样式和表格网格使用其他模块。
