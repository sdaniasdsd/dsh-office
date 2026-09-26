# pdf-edit

使用 pdf-lib 执行 `pdf-engines/src/plans.ts` 定义的页面旋转/选择/追加、文字盖章、Text 批注、AcroForm 文本和复选字段更新。

```ts
createPdfEditModule({ files, writer, backend?: { fontPath, qpdfPath } })
```

原件保持原样，新 PDF 作为新内容摘要产物保存。签名、加密、XFA、带表单页面筛选及断开/歧义 widget 字段拒绝写入。字段写入后重新检查 canonical `/Fields` 值、page `/Widget` 祖先和有效值、每个正常 appearance stream。文本盖章仅限单行、未旋转、`UserUnit=1` 页面；不能替代任意正文编辑或安全删隐。
