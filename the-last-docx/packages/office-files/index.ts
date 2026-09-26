// Promoted from docx-inspect/types/office-deps.d.ts; single shared owner.

  /** 一个本地文件路径，附带可选的清理钩子。 */
  export interface MaterializedArtifact {
    path: string;
    cleanup?: () => Promise<void>;
  }

  /**
   * 把 `ArtifactRef` 物化为可读本地路径的端口。
   *
   * docx-inspect 的默认实现只理解本地路径与 `file:` URI；当 Profile 需要支持
   * 远程/压缩包内的 artifact 时，注入真实实现即可，本模块代码无需改动。
   */
  export interface ArtifactMaterializer {
    materialize(ref: { id: string; uri: string }): Promise<MaterializedArtifact>;
  }
