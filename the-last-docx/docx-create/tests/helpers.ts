import { createDocxCreateModule } from '../src/index';
import type { ArtifactStore, ExecuteInput } from '../src/contract';
import type { ArtifactRef } from 'office-core';
import { unzipSync, strFromU8, strToU8, zipSync } from 'fflate';
export function memoryStore() {
  const files = new Map<string, Uint8Array>(); let sequence = 0;
  const store: ArtifactStore = {
    async read(ref) { const bytes = files.get(ref.id); if (!bytes) throw new Error('Missing artifact'); return bytes.slice(); },
    async write({ bytes }) {
      const id = `output-${++sequence}`; const ref = { id, uri: `memory://${id}` };
      files.set(id, bytes.slice()); return ref;
    },
  };
  const put = (id: string, bytes: Uint8Array): ArtifactRef => { files.set(id, bytes); return { id, uri: `memory://${id}` }; };
  return { store, files, put };
}
export const simple: ExecuteInput = {
  requestId: 'test', operation: 'execute', plan: { kind: 'create', document: {
    blocks: [{ kind: 'paragraph', id: 'intro', runs: [{ text: '初始内容' }] }],
  } },
};
export async function generated() {
  const memory = memoryStore(); const module = createDocxCreateModule({ artifactStore: memory.store });
  const out = await module.handlers.execute(simple);
  return { ...memory, module, out, bytes: memory.files.get(out.result.artifactRef.id)! };
}
export function part(bytes: Uint8Array, name = 'word/document.xml'): string { return strFromU8(unzipSync(bytes)[name]!); }
export function rewrite(bytes: Uint8Array, change: (entries: Record<string, Uint8Array>) => void): Uint8Array {
  const entries = unzipSync(bytes); change(entries); return zipSync(entries);
}
export const asBytes = strToU8;
