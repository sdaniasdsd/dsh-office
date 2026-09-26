# pdf-parse

用 PDF.js 产生 PDF 页面、原生 text item 与位置的双 IR；每条 observation 带源 SHA、页面、bbox 和证据。

```ts
createPdfParseModule({ files })
```

IR 文本节点保留原始 text item 粒度，并不是推断出的段落。物理文字没有可写锚点或 object reference；source map 不能用重复文字指纹代替选择器。bbox 约定左上原点、应用 CropBox 和旋转后的 pt。阅读顺序及表格覆盖在该模块中标 unavailable；复杂布局交由 `pdf-complex-parse`。
