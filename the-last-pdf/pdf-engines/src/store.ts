import { mkdir, readFile, stat, writeFile, link, unlink } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { randomUUID } from 'node:crypto';
import { sha256, PdfModuleError } from '@dsh-office-profile/pdf-contracts';
import type { ArtifactRef, ArtifactReader, ArtifactWriter } from '@dsh-office-profile/pdf-contracts';

export class LocalArtifactStore implements ArtifactReader, ArtifactWriter {
  readonly root: string;
  constructor(root: string) { this.root = resolve(root); }
  async read(ref: ArtifactRef, maxBytes: number, signal: AbortSignal): Promise<Uint8Array> {
    signal.throwIfAborted();
    if (/^[a-z]+:/i.test(ref.uri) && !/^[a-z]:[\\/]/i.test(ref.uri) && !ref.uri.startsWith('file:'))
      throw new PdfModuleError('ARTIFACT_NOT_FOUND', 'Only local/file artifact references are supported by LocalArtifactStore.');
    const path = ref.uri.startsWith('file:') ? fileURLToPath(ref.uri) : resolve(ref.uri);
    if ((await stat(path)).size > maxBytes) throw new PdfModuleError('LIMIT_EXCEEDED', 'File exceeds read budget.');
    const bytes = await readFile(path, { signal });
    if (bytes.length > maxBytes) throw new PdfModuleError('LIMIT_EXCEEDED', 'File grew beyond read budget.');
    if ((ref.sha256 && ref.sha256 !== sha256(bytes)) || (ref.sizeBytes !== undefined && ref.sizeBytes !== bytes.length))
      throw new PdfModuleError('ARTIFACT_MISMATCH', 'Stored bytes do not match reference.');
    return bytes;
  }
  private async publish(path: string, bytes: Uint8Array, signal: AbortSignal): Promise<void> {
    await mkdir(this.root, { recursive: true });
    const temporary = join(this.root, `.pending-${randomUUID()}`);
    try {
      await writeFile(temporary, bytes, { flag: 'wx', signal });
      signal.throwIfAborted();
      // A hard-link publish is atomic and never replaces an existing destination.
      try { await link(temporary, path); }
      catch (error) {
        if ((error as NodeJS.ErrnoException).code !== 'EEXIST') throw error;
        if (sha256(await readFile(path)) !== sha256(bytes)) throw new PdfModuleError('ARTIFACT_CONFLICT', 'Immutable destination contains different bytes.');
      }
    } finally { await unlink(temporary).catch(() => {}); }
  }
  async write(input: Parameters<ArtifactWriter['write']>[0]): Promise<ArtifactRef> {
    const digest = sha256(input.bytes);
    const name = input.suggestedName.replace(/[^A-Za-z0-9._-]/g, '_').slice(-100) || 'artifact.bin';
    const path = join(this.root, `${digest}-${name}`);
    await this.publish(path, input.bytes, input.signal);
    return { id: `sha256:${digest}`, uri: pathToFileURL(path).href, sha256: digest, sizeBytes: input.bytes.length,
      mediaType: name.endsWith('.pdf') ? 'application/pdf' : name.endsWith('.png') ? 'image/png' : 'application/json', label: name };
  }
  async commitManifest(input: Parameters<ArtifactWriter['commitManifest']>[0]): Promise<ArtifactRef> {
    if (!/^[A-Za-z0-9_-]{1,80}$/.test(input.documentId) || !Number.isSafeInteger(input.revision) || input.revision < 1)
      throw new PdfModuleError('INVALID_INPUT', 'Invalid manifest identity.');
    if (input.revision === 1 && input.parentSha256) throw new PdfModuleError('ARTIFACT_CONFLICT', 'First revision cannot have parent.');
    if (input.revision > 1) {
      const parent = join(this.root, `${input.documentId}.r${input.revision - 1}.manifest.json`);
      let previous: Uint8Array;
      try { previous = await readFile(parent); } catch { throw new PdfModuleError('ARTIFACT_CONFLICT', 'Previous committed revision is missing.'); }
      if (!input.parentSha256 || sha256(previous) !== input.parentSha256)
        throw new PdfModuleError('ARTIFACT_CONFLICT', 'Previous revision digest mismatch.');
    }
    const path = join(this.root, `${input.documentId}.r${input.revision}.manifest.json`);
    await this.publish(path, input.bytes, input.signal);
    const digest = sha256(input.bytes);
    return { id: `manifest:${input.documentId}:${input.revision}`, uri: pathToFileURL(path).href,
      sha256: digest, sizeBytes: input.bytes.length, mediaType: 'application/json' };
  }
}
