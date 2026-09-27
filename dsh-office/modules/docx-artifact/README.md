# docx-artifact

交付模块，不是新的解析器或渲染器。`inspect` 检查 DOCX，`execute` 校验产物引用并原子提交版本化 manifest，`verify` 接收 manifest 的 ArtifactRef，重新核验所有引用与版本链。

依赖由 Profile 注入：`files.read / commitManifest` 负责读取与不可变提交，`inspector.inspect` 复用现有检查能力。模块从不读取全局路径，不打开网络，不生成预览，不改写 DOCX。

```ts
const module = createDocxArtifactModule({ files, inspector });
const delivered = await module.handlers.execute({
  artifactRef, requestId: 'delivery-1', operation: 'execute',
  delivery: { documentId: 'report-001', revision: 1, preview: renderResult, bridges },
});
```

没有预览也可交付，但报告为部分验证；Profile 可强制预览和视觉复核。视觉状态只承接调用方提供的 `docx-render` findings，不把图片存在等同于视觉通过。每张图/PDF/桥接 JSON 校验真实大小与 SHA-256，并绑定源 DOCX 的 id + sha256，旧预览不能套用到新版本。

桥接证据文件为 `{schema:'docx-evidence/v1',kind,source:{id,sha256},payload}`；payload 原样保存各模块公开输出。不会转换或覆盖双 IR，也不会把复杂解析节点冒充编辑锚点。引擎只检查文件与来源绑定，不宣称重新验证全部 IR 语义。

版本 1 无父版本；后续必须指向上一 revision 的 manifest。原子冲突/幂等由 `commitManifest` 契约负责，不允许覆盖。同版本相同内容可重试，不同内容必须冲突。timeout 使用 AbortSignal，适配器须在提交前检查信号；如果提交确认丢失，用同内容重试确认，不能以超时推断已回滚。

`npm ci; npm run typecheck; npm test` 可单包执行。公共类型临时引用已有 docx-parse 类型桩，交付插件中统一替换为共享 office-* 契约包。输入/manifest/config schema 从 `DOCX_ARTIFACT_SCHEMAS` 导出。
