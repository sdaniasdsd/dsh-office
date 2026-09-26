# pdf-artifact

为一份 PDF 注册递增版本 manifest。预览和 JSON 证据都携带文档的 `id + sha256`，来源错配时拒绝提交。

```ts
createPdfArtifactModule({ files, writer })
```

`LocalArtifactStore` 作为可运行的本机 reference adapter，以内容摘要保存文件，并用不覆盖 hardlink 追加 revision manifest；后续 revision 需校验父清单摘要。最终 verify 遍历整条版本链、重读并重 hash 文档、预览及证据。正式宿主需要把此端口连接到受控 `office-files` durable store。
