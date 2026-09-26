# pdf-render

PDF.js 以受预算约束的 DPI 逐页生成 PNG，并将图像写入 `ArtifactWriter`。

```ts
createPdfRenderModule({ files, writer })
```

输入页数顺序与来源摘要绑定。最大像素面积在分配 Canvas 之前核验；每图和总输出字节也有界限。结果 `visualReview: pending`，须审阅者明确提交全部页复核证据后才可交付为 reviewed。
