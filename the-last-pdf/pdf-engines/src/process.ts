import { spawn } from 'node:child_process';
import { StringDecoder } from 'node:string_decoder';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { PdfModuleError } from '@dsh-office-profile/pdf-contracts';
import type { EngineContext, ModuleOptions, VerificationCheck } from '@dsh-office-profile/pdf-contracts';
import { sha256 } from '@dsh-office-profile/pdf-contracts';

export function runProcess(command: string, args: string[], signal: AbortSignal, maxBytes = 16 * 1024 * 1024) {
  return new Promise<{ code: number; stdout: string; stderr: string }>((resolve, reject) => {
    signal.throwIfAborted();
    const child = spawn(command, args, { windowsHide: true, shell: false, stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '', stderr = '', size = 0;
    const outDecoder = new StringDecoder('utf8'), errDecoder = new StringDecoder('utf8');
    let failure: unknown;
    const abort = () => { failure = signal.reason; child.kill(); };
    signal.addEventListener('abort', abort, { once: true });
    const receive = (kind: 'stdout' | 'stderr', chunk: Buffer) => {
      size += chunk.length;
      if (size > maxBytes) { failure = new PdfModuleError('LIMIT_EXCEEDED', 'Subprocess output budget exceeded.'); child.kill(); return; }
      if (kind === 'stdout') stdout += outDecoder.write(chunk); else stderr += errDecoder.write(chunk);
    };
    child.stdout.on('data', c => receive('stdout', c)); child.stderr.on('data', c => receive('stderr', c));
    child.on('error', () => { failure = new PdfModuleError('ENGINE_UNAVAILABLE', `Cannot start ${command}.`); });
    child.on('close', code => {
      stdout += outDecoder.end(); stderr += errDecoder.end();
      signal.removeEventListener('abort', abort);
      if (failure) reject(failure); else resolve({ code: code ?? -1, stdout, stderr });
    });
  });
}
export async function withPdfFile<T>(bytes: Uint8Array, fn: (path: string) => Promise<T>): Promise<T> {
  const directory = await mkdtemp(join(tmpdir(), 'pdf-profile-'));
  try { const path = join(directory, 'input.pdf'); await writeFile(path, bytes); return await fn(path); }
  finally { await rm(directory, { recursive: true, force: true }); }
}
export async function readArtifact(ref: import('@dsh-office-profile/pdf-contracts').ArtifactRef,
  context: EngineContext, options: ModuleOptions, budget = context.config.limits.maxInputBytes): Promise<Uint8Array> {
  if (!options.files) throw new PdfModuleError('ARTIFACT_NOT_FOUND', 'Artifact reader is not connected.');
  if (ref.sizeBytes !== undefined && ref.sizeBytes > budget) throw new PdfModuleError('LIMIT_EXCEEDED', 'Artifact exceeds declared read budget.');
  const bytes = await options.files.read(ref, budget, context.signal);
  if (!(bytes instanceof Uint8Array) || bytes.byteLength > budget) throw new PdfModuleError('LIMIT_EXCEEDED', 'Artifact exceeds read budget.');
  if ((ref.sha256 && sha256(bytes) !== ref.sha256) || (ref.sizeBytes !== undefined && bytes.length !== ref.sizeBytes))
    throw new PdfModuleError('ARTIFACT_MISMATCH', 'Artifact digest or length mismatch.');
  return bytes;
}
export async function checkQpdf(context: EngineContext, options: ModuleOptions): Promise<VerificationCheck> {
  try {
    const result = await withPdfFile(context.source!.bytes, path => runProcess(options.backend?.qpdfPath ?? 'qpdf', ['--check', path], context.signal, 1024 * 1024));
    return { id: 'structure.qpdf', severity: result.code === 3 ? 'warn' : 'error',
      status: result.code === 0 ? 'pass' : 'fail',
      message: result.code === 0 ? 'qpdf structural check passed.' : result.code === 3 ? 'qpdf reported recoverable warnings.' : 'qpdf structural check failed.' };
  } catch (error) {
    if (error instanceof PdfModuleError && error.code === 'ENGINE_UNAVAILABLE' && options.backend?.requireQpdf === false)
      return { id: 'structure.qpdf', severity: 'warn', status: 'skip', message: 'qpdf unavailable; explicitly optional in this Profile.' };
    throw error;
  }
}
