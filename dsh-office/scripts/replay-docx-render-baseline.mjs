import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { createHash, randomUUID } from 'node:crypto';
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const baselineRoot = resolve(process.argv[2] ?? '');
const replayRoot = resolve(process.argv[3] ?? '');
if (!process.argv[2] || !process.argv[3]) throw new Error('Usage: node scripts/replay-docx-render-baseline.mjs <frozen-run-root> <new-replay-root>');
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
await mkdir(replayRoot, { recursive: false });
const isolatedData = join(replayRoot, 'data');
await mkdir(isolatedData, { recursive: true });
const evidencePath = join(replayRoot, 'docx-render-50.json');
const manifest = JSON.parse(await readFile(join(baselineRoot, 'manifest.json'), 'utf8'));
const report = JSON.parse(await readFile(join(baselineRoot, 'report.json'), 'utf8'));
const specimens = manifest.specimens.filter(item => item.format === 'docx');
const imported = new Map(report.imports.filter(item => item.format === 'docx').map(item => [item.file, item.artifactRef]));
if (specimens.length !== 5 || imported.size !== 5) throw new Error('Expected five frozen DOCX samples and five imported artifactRefs.');
const refs = new Map();
for (const specimen of specimens) {
  const ref = imported.get(specimen.file);
  if (!ref?.sha256 || !ref.uri) throw new Error(`Missing DSH DOCX artifactRef for ${specimen.file}.`);
  const original = new URL(ref.uri), bytes = await readFile(original);
  if (hash(bytes) !== ref.sha256) throw new Error(`Imported artifact hash mismatch: ${specimen.file}.`);
  const cloned = join(isolatedData, 'objects', ref.sha256, 'artifact.docx');
  await mkdir(join(isolatedData, 'objects', ref.sha256), { recursive: true });
  await copyFile(original, cloned);
  refs.set(specimen.file, { ...ref, uri: pathToFileURL(cloned).href });
}

async function captureHashes() {
  return Promise.all(manifest.specimens.map(async item => {
    const [source, workspace] = await Promise.all([
      readFile(join(baselineRoot, 'sources', item.format, item.file)),
      readFile(join(baselineRoot, 'workspace', item.workspacePath ?? `${item.format}/${item.file}`)),
    ]);
    return { format: item.format, file: item.file, sourceSha256: hash(source), workspaceSha256: hash(workspace) };
  }));
}
const hashesBefore = await captureHashes();
for (const item of manifest.specimens) {
  const actual = hashesBefore.find(value => value.format === item.format && value.file === item.file);
  if (actual.sourceSha256 !== item.sourceSha256 || actual.workspaceSha256 !== item.workspaceSha256)
    throw new Error(`Frozen corpus hash differs from manifest: ${item.format}/${item.file}`);
}

const profileRoot = process.env.DSH_DOCX_PROFILE_ROOT ?? 'C:/Users/AA/AppData/Roaming/com.yeagoo.dsh-desktop/harness/profiles/web/node_modules/@deepseek-ai/dsh-docx';
const nodeRuntime = process.env.DSH_NODE_RUNTIME ?? 'C:/Users/AA/AppData/Local/DSH Desktop/runtime/node.exe';
const bundleRuntime = `${profileRoot}/runtime/win32-x64`;
const env = Object.fromEntries(Object.entries(process.env).filter(([, value]) => typeof value === 'string'));
delete env.PYTHONIOENCODING; delete env.PYTHONUTF8;
Object.assign(env, { THE_LAST_DOCX_WORKSPACE: join(baselineRoot, 'workspace'), THE_LAST_DOCX_DATA: isolatedData,
  DOCX_PYTHON: `${bundleRuntime}/python/python.exe`, DOCX_SOFFICE: `${bundleRuntime}/libreoffice/program/soffice.com`,
  DOCX_PDFTOPPM: `${bundleRuntime}/poppler/poppler-26.09.0/Library/bin/pdftoppm.exe` });
const client = new Client({ name: 'dsh-office-b3-render-replay', version: '1.0.0' });
const transport = new StdioClientTransport({ command: nodeRuntime, args: [`${profileRoot}/lib/server.mjs`], cwd: baselineRoot, env, stderr: 'pipe' });
const records = [], serverStderr = [];
let failedFast = false;
transport.stderr?.on('data', chunk => serverStderr.push(Buffer.isBuffer(chunk) ? chunk.toString('utf8') : String(chunk)));
try {
  await client.connect(transport);
  for (const specimen of specimens) {
    if (failedFast) break;
    const artifactRef = refs.get(specimen.file);
    const baselineRecord = report.records.find(record => record.format === 'pdf' && record.sourceDocx === specimen.file && record.round === 1 && record.ok);
    const expected = { pageCount: baselineRecord.renderResult.pageCount,
      pages: baselineRecord.renderResult.pages.map(page => ({ image: page.image.sha256, thumbnail: page.thumbnail.sha256 })) };
    for (let round = 1; round <= 10; round++) {
      const requestId = `b3-render-${specimen.file}-${String(round).padStart(2, '0')}-${randomUUID()}`;
      const start = performance.now();
      let error, pagePixelHashes, pageCount, pixelsMatch;
      try {
        const response = await client.callTool({ name: 'docx_call', arguments: { moduleId: 'docx-render', operation: 'execute',
          input: { operation: 'execute', requestId, artifactRef } } });
        if (response.isError) error = response.content?.filter(item => item.type === 'text').map(item => item.text).join('\n') ?? 'DSH tool returned an error without text.';
        else {
          const envelope = JSON.parse(response.content?.filter(item => item.type === 'text').map(item => item.text).join('\n') ?? '{}');
          const rendered = envelope.result ?? envelope;
          if (!rendered?.pages || !Array.isArray(rendered.pages)) throw new Error(`DSH response omitted render pages; envelope=${JSON.stringify(envelope).slice(0, 3000)}`);
          pageCount = rendered.pageCount;
          pagePixelHashes = rendered.pages.map(page => ({ image: page.image.sha256, thumbnail: page.thumbnail?.sha256 ?? null }));
          pixelsMatch = pageCount === expected.pageCount && JSON.stringify(pagePixelHashes) === JSON.stringify(expected.pages);
          if (!pixelsMatch) error = `Render differs from B0: ${JSON.stringify({ gotPageCount: pageCount, expectedPageCount: expected.pageCount,
            got: pagePixelHashes, expected: expected.pages })}`;
        }
      } catch (cause) { error = String(cause?.stack ?? cause); }
      records.push({ file: specimen.file, round, requestId, durationMs: Math.round(performance.now() - start),
        ok: !error, pageCount, pixelsMatch, pagePixelHashes, error });
      process.stdout.write(`${error ? 'FAIL' : 'PASS'} ${specimen.file} round=${round} ${records.at(-1).durationMs}ms\n`);
      if (error) { process.stdout.write(`FIRST FAILURE DETAILS: ${error}\n`); failedFast = true; break; }
    }
  }
} finally { await client.close().catch(() => {}); }

const hashesAfter = await captureHashes();
const hashStable = JSON.stringify(hashesBefore) === JSON.stringify(hashesAfter);
const groups = specimens.map(specimen => {
  const items = records.filter(record => record.file === specimen.file), failed = items.filter(record => !record.ok).length;
  const durations = items.map(record => record.durationMs).sort((a, b) => a - b);
  return { file: specimen.file, rounds: items.length, passed: items.length - failed, failed,
    meanDurationMs: Math.round(durations.reduce((sum, value) => sum + value, 0) / Math.max(1, durations.length)),
    p95DurationMs: durations[Math.min(durations.length - 1, Math.ceil(durations.length * 0.95) - 1)] ?? null };
});
const errors = [];
if (!hashStable) errors.push('Frozen source/workspace hashes changed during render replay.');
if (records.length !== 50) errors.push(`Expected 50 DOCX render calls; got ${records.length}.`);
for (const group of groups) if (group.rounds !== 10 || group.failed !== 0) errors.push(`${group.file}: expected 10/10 pixel-stable renders.`);
const result = { createdAt: new Date().toISOString(), baselineRoot, isolatedData, profile: `source-built DSH profile at ${profileRoot}`,
  backend: 'DSH profile MCP stdio server', moduleId: 'docx-render', settings: { dpi: 120, emitPdf: true, emitThumbnails: true },
  groups, records, hashesBefore, hashesAfter, hashStable, serverStderr: serverStderr.join(''), status: errors.length ? 'fail' : 'pass', errors };
await writeFile(evidencePath, JSON.stringify(result, null, 2), { encoding: 'utf8', flag: 'wx' });
process.stdout.write(JSON.stringify({ evidencePath, status: result.status, calls: records.length,
  failures: records.filter(record => !record.ok).length, groups, hashStable, errors }, null, 2) + '\n');
if (errors.length) process.exitCode = 2;
