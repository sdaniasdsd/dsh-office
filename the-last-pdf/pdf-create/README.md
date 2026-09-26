# pdf-create

PDFKit 直接从严格 schema 的结构化计划创建新 PDF：标题、段落、分页表格和页码。

```ts
createPdfCreateModule({ writer, backend?: { fontPath, qpdfPath } })
```

payload 见 `pdf-engines/src/plans.ts` 的 `createPlanSchema`。非 ASCII 字符须以 `fontRef` 或可信 `backend.fontPath` 提供有对应 glyph 的 TTF/OTF。输出单独写入；引擎以 pdf-lib 重开页数、用 qpdf 检结构。生成图像仍是 pending 状态，等待视觉复核。
