import { mkdtemp, readFile, writeFile, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { randomUUID } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { createDocxParseModule } from '@dsh-office-profile/docx-parse';
import { createDocxEditModule, targetFromDualIR } from '@dsh-office-profile/docx-edit';
import { createDocxRenderModule } from '@dsh-office-profile/docx-render';
import { createDocxCreateModule } from '../src/index';
import { simple } from './helpers';
import type { ArtifactRef } from 'office-core';

describe('local sibling module integration (real parser and editor)', () => {
  it('create -> dual IR -> anchored edit -> parse retains the same paragraph identity', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'docx-create-integration-'));
    const store = {
      async read(ref: ArtifactRef) { return new Uint8Array(await readFile(ref.uri)); },
      async write({ bytes }: { bytes: Uint8Array }) {
        const id = randomUUID(); const uri = join(directory, `${id}.docx`);
        await writeFile(uri, bytes, { flag: 'wx' }); return { id, uri };
      },
    };
    const creator = createDocxCreateModule({ artifactStore: store });
    const parser = createDocxParseModule(); const editor = createDocxEditModule({ artifactStore: store });
    try {
      const created = await creator.handlers.execute(simple);
      const artifactRef = created.result.artifactRef;
      const parsed = await parser.handlers.execute({ requestId: 'parse-created', operation: 'execute', artifactRef });
      const paragraph = parsed.result.ir.content.semantic.blocks.find(b => b.kind === 'paragraph');
      if (!paragraph || paragraph.kind !== 'paragraph') throw new Error('Missing created paragraph');
      expect(created.result.paragraphIds).toContain(paragraph.anchor.paraId);
      expect(parsed.result.ir.content.sourceMap.bySemanticId[paragraph.id]?.length).toBeGreaterThan(0);
      const edited = await editor.handlers.execute({ requestId: 'edit-created', operation: 'execute', artifactRef,
        plan: { edits: [{ kind: 'replaceText', target: targetFromDualIR(parsed.result.ir.content, paragraph.id), find: '初始内容', replace: '修改完成' }] },
      });
      const editedRef = edited.artifacts[0]!;
      const reparsed = await parser.handlers.execute({ requestId: 'parse-edited', operation: 'execute', artifactRef: editedRef });
      const same = reparsed.result.ir.content.semantic.blocks.find(b => b.id === paragraph.id);
      expect(same?.anchor.quote).toBe('修改完成');
      const verification = await creator.handlers.verify({ requestId: 'verify-edited', operation: 'verify', artifactRef: editedRef });
      expect(verification.result.ok).toBe(true);
      expect(new Uint8Array(await readFile(artifactRef.uri))).not.toEqual(new Uint8Array(await readFile(editedRef.uri)));
    } finally {
      await creator.dispose(); await parser.dispose(); await editor.dispose();
      // Exact task-owned directory returned by mkdtemp, never a user/workspace root.
      await rm(directory, { recursive: true, force: true });
    }
  }, 30_000);
  it('passes output refs to renderer; missing layout runtime is not visual success', async () => {
    const { generated } = await import('./helpers'); const result = await generated();
    const renderer = createDocxRenderModule({
      artifactStore: { read: result.store.read, async write() { throw new Error('Should not write without renderer'); } },
      config: { engine: { sofficePath: 'nonexistent-docx-create-test-soffice' }, timeoutMs: 5000 },
    });
    try {
      await expect(renderer.handlers.execute({ operation: 'execute', requestId: 'render', artifactRef: result.out.result.artifactRef })).rejects.toMatchObject({ code: 'ENGINE_UNAVAILABLE' });
    } finally { await renderer.dispose(); await result.module.dispose(); }
  });
});
