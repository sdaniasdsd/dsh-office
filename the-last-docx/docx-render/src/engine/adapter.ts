import { mkdir, mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, extname, join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { spawn } from 'node:child_process';
import type { ArtifactRef } from 'office-core';
import type { EnginePage, EngineRenderResult, ModuleConfig, RenderEngine } from '../contract';
import { DocxRenderError } from '../errors';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const DOCM_MIME = 'application/vnd.ms-word.document.macroEnabled.12';
const DOTM_MIME = 'application/vnd.ms-word.template.macroEnabled.12';
const SUPPORTED_EXTENSIONS = new Set(['.docx', '.docm', '.dotx', '.dotm']);

function extensionFor(ref: ArtifactRef): string | null {
  const name = ref.label || ref.uri.split(/[?#]/, 1)[0] || '';
  const extension = extname(name).toLowerCase();
  return extension || null;
}

function run(command: string, args: string[], timeoutMs: number, cwd: string): Promise<void> {
  return new Promise((resolve, reject) => {
    let settled = false;
    let timedOut = false;
    let killGrace: ReturnType<typeof setTimeout> | undefined;
    const child = spawn(command, args, { cwd, windowsHide: true, shell: false, stdio: ['ignore', 'ignore', 'pipe'] });
    let stderr = '';
    child.stderr.setEncoding('utf8');
    child.stderr.on('data', (chunk: string) => { stderr = `${stderr}${chunk}`.slice(-4000); });
    const finish = (error?: Error) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (killGrace) clearTimeout(killGrace);
      if (error) reject(error); else resolve();
    };
    const timer = setTimeout(() => {
      timedOut = true;
      child.kill();
      killGrace = setTimeout(() => finish(new DocxRenderError('ENGINE_TIMEOUT', `Renderer process exceeded timeoutMs (${timeoutMs}).`)), 2000);
    }, timeoutMs);
    child.once('error', (error) => finish(new DocxRenderError('ENGINE_UNAVAILABLE', `Could not start renderer executable: ${error.message}`)));
    child.once('close', (code) => {
      if (timedOut) return finish(new DocxRenderError('ENGINE_TIMEOUT', `Renderer process exceeded timeoutMs (${timeoutMs}).`));
      if (code === 0) return finish();
      const detail = stderr.trim();
      finish(new DocxRenderError('ENGINE_FAILED', `Renderer exited with code ${code}${detail ? `: ${detail}` : '.'}`));
    });
  });
}

async function listPageImages(directory: string, prefix: string): Promise<string[]> {
  const { readdir } = await import('node:fs/promises');
  const names = await readdir(directory);
  return names
    .filter((name) => new RegExp(`^${prefix}-(\\d+)\\.png$`, 'i').test(name))
    .sort((left, right) => Number(left.match(/-(\d+)\.png$/i)?.[1]) - Number(right.match(/-(\d+)\.png$/i)?.[1]))
    .map((name) => join(directory, name));
}

export class LibreOfficePopplerEngine implements RenderEngine {
  readonly name = 'libreoffice-poppler';
  private profilePromise: Promise<{ root: string; directory: string }> | undefined;
  private readonly profileRoots = new Set<string>();
  private conversionQueue: Promise<void> = Promise.resolve();
  private disposed = false;
  constructor(private readonly baseConfig: ModuleConfig['engine']) {}

  private async profile(): Promise<{ root: string; directory: string }> {
    if (!this.profilePromise) {
      this.profilePromise = (async () => {
        const base = this.baseConfig.tempRoot ?? tmpdir();
        await mkdir(base, { recursive: true });
        const root = await mkdtemp(join(base, 'docx-render-profile-'));
        const directory = join(root, 'lo-profile');
        await mkdir(directory, { recursive: false });
        this.profileRoots.add(root);
        return { root, directory };
      })();
    }
    return this.profilePromise;
  }

  private async serializeConversion<T>(work: () => Promise<T>): Promise<T> {
    const previous = this.conversionQueue;
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    this.conversionQueue = previous.then(() => gate);
    await previous;
    try { return await work(); } finally { release(); }
  }

  async inspect(bytes: Uint8Array, artifactRef: ArtifactRef): Promise<{ extension: string | null; mediaType: string | null }> {
    const extension = extensionFor(artifactRef);
    if (!SUPPORTED_EXTENSIONS.has(extension ?? '')) {
      throw new DocxRenderError('FORMAT_MISMATCH', 'docx-render accepts DOCX-family files (.docx, .docm, .dotx, .dotm).');
    }
    if (bytes.byteLength < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
      throw new DocxRenderError('FORMAT_MISMATCH', 'The artifact is not a ZIP-based Office Open XML document.');
    }
    const expectedMimes = [DOCX_MIME, DOCM_MIME, DOTM_MIME, 'application/vnd.openxmlformats-officedocument.wordprocessingml.template'];
    if (artifactRef.mediaType && !expectedMimes.includes(artifactRef.mediaType)) {
      throw new DocxRenderError('FORMAT_MISMATCH', 'The artifact media type is not a supported WordprocessingML type.');
    }
    return { extension, mediaType: artifactRef.mediaType ?? (extension === '.docx' ? DOCX_MIME : null) };
  }

  async render(bytes: Uint8Array, config: ModuleConfig, artifactRef: ArtifactRef): Promise<EngineRenderResult> {
    const deadline = Date.now() + config.timeoutMs;
    const runWithinBudget = (command: string, args: string[], cwd: string) =>
      run(command, args, Math.max(1, deadline - Date.now()), cwd);
    const tempBase = config.engine.tempRoot ?? this.baseConfig.tempRoot ?? tmpdir();
    await mkdir(tempBase, { recursive: true });
    const workDir = await mkdtemp(join(tempBase, 'docx-render-'));
    try {
      const inputExtension = extensionFor(artifactRef) ?? '.docx';
      const inputName = `source${inputExtension}`;
      const inputPath = join(workDir, inputName);
      const pdfPath = join(workDir, 'source.pdf');
      const imagePrefix = 'page';
      await writeFile(inputPath, bytes);
      const soffice = config.engine.sofficePath || this.baseConfig.sofficePath;
      const pdftoppm = config.engine.pdftoppmPath || this.baseConfig.pdftoppmPath;
      if (this.disposed) throw new DocxRenderError('ENGINE_UNAVAILABLE', 'The renderer has been disposed.');
      await this.serializeConversion(async () => {
        const profile = await this.profile();
        const profileUri = pathToFileURL(profile.directory).href;
        try {
          await runWithinBudget(soffice, [
            '--headless', '--nologo', '--nodefault', '--norestore',
            `-env:UserInstallation=${profileUri}`,
            '--convert-to', 'pdf', '--outdir', workDir, inputPath,
          ], workDir);
        } catch (error) {
          // A crashed or timed-out office process may leave a locked/corrupt profile.
          // Retire it; the next serialized conversion receives a fresh isolated one.
          this.profilePromise = undefined;
          throw error;
        }
      });
      const pdfInfo = await stat(pdfPath).catch(() => undefined);
      if (!pdfInfo || pdfInfo.size === 0) throw new DocxRenderError('ENGINE_PROTOCOL_ERROR', 'LibreOffice reported success but did not produce a PDF.');
      if (pdfInfo.size > config.limits.maxPdfBytes) throw new DocxRenderError('LIMIT_EXCEEDED', 'The rendered PDF exceeds maxPdfBytes.');

      await runWithinBudget(pdftoppm, [
        '-png', '-r', String(config.limits.dpi), '-f', '1', '-l', String(config.limits.maxPages + 1),
        pdfPath, join(workDir, imagePrefix),
      ], workDir);
      const imagePaths = await listPageImages(workDir, imagePrefix);
      if (imagePaths.length === 0) throw new DocxRenderError('ENGINE_PROTOCOL_ERROR', 'PDF conversion succeeded but no page images were produced.');
      if (imagePaths.length > config.limits.maxPages) throw new DocxRenderError('LIMIT_EXCEEDED', 'The document exceeds maxPages.');

      const thumbnails = new Map<number, string>();
      if (config.featureFlags.emitThumbnails) {
        await runWithinBudget(pdftoppm, [
          '-png', '-scale-to', String(config.limits.thumbnailMaxEdge), '-f', '1', '-l', String(config.limits.maxPages + 1),
          pdfPath, join(workDir, 'thumb'),
        ], workDir);
        for (const path of await listPageImages(workDir, 'thumb')) {
          const pageNumber = Number(path.match(/-(\d+)\.png$/i)?.[1]);
          thumbnails.set(pageNumber, path);
        }
      }

      let totalImageBytes = 0;
      const pages: EnginePage[] = [];
      for (let index = 0; index < imagePaths.length; index += 1) {
        const imagePath = imagePaths[index];
        if (!imagePath) continue;
        const pageNumber = Number(imagePath.match(/-(\d+)\.png$/i)?.[1]);
        const image = new Uint8Array(await readFile(imagePath));
        totalImageBytes += image.byteLength;
        const thumbnailPath = thumbnails.get(pageNumber);
        const thumbnail = thumbnailPath ? new Uint8Array(await readFile(thumbnailPath)) : undefined;
        totalImageBytes += thumbnail?.byteLength ?? 0;
        if (totalImageBytes > config.limits.maxTotalImageBytes) throw new DocxRenderError('LIMIT_EXCEEDED', 'Rendered images exceed maxTotalImageBytes.');
        pages.push({ pageNumber, image, ...(thumbnail ? { thumbnail } : {}) });
      }
      return { pdf: new Uint8Array(await readFile(pdfPath)), pages };
    } finally {
      await rm(workDir, { recursive: true, force: true });
    }
  }

  async dispose(): Promise<void> {
    this.disposed = true;
    await this.serializeConversion(async () => {
      await Promise.all([...this.profileRoots].map(root => rm(root, { recursive: true, force: true }).catch(() => undefined)));
      this.profileRoots.clear();
      this.profilePromise = undefined;
    });
  }
}

export function suggestedStem(ref: ArtifactRef): string {
  const label = ref.label ?? basename(ref.uri);
  const stem = label.replace(/\.[^.]*$/, '').replace(/[^a-zA-Z0-9._-]+/g, '_');
  return stem || 'document';
}
