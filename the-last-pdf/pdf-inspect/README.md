# pdf-inspect

PDF.js 元数据及 PDF 对象扫描，报告页数、加密、JavaScript、附件、外链、AcroForm、XFA、签名和标签结构。

```ts
createPdfInspectModule({ files, backend?: { qpdfPath, requireQpdf } })
```

通用方式是注入 `ArtifactReader`、调用 `handlers.inspect({ requestId, operation: 'inspect', artifactRef })`。默认 engine 在 reader 存在时绑定；否则返回 `ENGINE_UNAVAILABLE`。引擎不执行文档中的动作。对象扫描对未知项返回 null；qpdf 输出不能证明视觉正确。

版本与范围见根目录 README 和 `docs/COMMUNITY-SOURCES.md`。
