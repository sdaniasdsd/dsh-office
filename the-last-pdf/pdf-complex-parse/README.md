# pdf-complex-parse

由本地 Docling Python 进程生成布局、OCR 和表格观察。执行前显式提供 `backend.pythonPath` 与已下载的 `backend.modelPath`。

```ts
createPdfComplexParseModule({ files, writer, backend: { pythonPath, modelPath } })
```

Python 进程使用 Docling 2.130.0、中文 RapidOCR 和显式配置的 PDFium backend；设置 Hugging Face/Transformers 离线标记，绝不在转换时下载。完整模型 observations 作为来源绑定 JSON 证据保存，IR 是有类型的下游投影。`confidence` 留空，模型区域不可写。小型/跨页表格识别率还需用标注语料测定。

首次配置步骤见根 README 的 `npm run demo:docling`。
