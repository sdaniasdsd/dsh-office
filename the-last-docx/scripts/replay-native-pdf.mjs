import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { createHash, randomUUID } from 'node:crypto';
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const baselineRoot = resolve(process.argv[2] ?? '');
const replayRoot = resolve(process.argv[3] ?? '');
if (!process.argv[2] || !process.argv[3]) throw new Error('Usage: node scripts/replay-native-pdf.mjs <frozen-run-root> <new-replay-root>');
const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
await mkdir(replayRoot, { recursive: false });
const evidencePath = join(replayRoot, 'native-pdf-50.json');
const isolatedData = join(replayRoot, 'data');
await mkdir(isolatedData, { recursive: true });

const manifest = JSON.parse(await readFile(join(baselineRoot, 'manifest.json'), 'utf8'));
const report = JSON.parse(await readFile(join(baselineRoot, 'report.json'), 'utf8'));
const specimens = manifest.pdfSamples;
if (!Array.isArray(specimens) || specimens.length !== 5) throw new Error(`Expected five frozen PDF specimens; got ${specimens?.length ?? 0}.`);
const refs = new Map();
for (const specimen of specimens) {
  const match = report.records.find(record => record.format === 'pdf' && record.sourceDocx === specimen.sourceDocx && record.round === 10 && record.ok);
  const ref = match?.renderResult?.pdf;
  if (!ref?.sha256 || !ref.uri || !ref.id) throw new Error(`Missing final PDF artifactRef for ${specimen.sourceDocx}.`);
  const original = new URL(ref.uri);
  const originalBytes = await readFile(original);
  if (sha256(originalBytes) !== ref.sha256) throw new Error(`PDF artifact hash mismatch: ${specimen.sourceDocx}.`);
  const cloned = join(isolatedData, 'objects', ref.sha256, 'artifact.pdf');
  await mkdir(join(isolatedData, 'objects', ref.sha256), { recursive: true });
  await copyFile(original, cloned);
  refs.set(specimen.sourceDocx, { ...ref, uri: pathToFileURL(cloned).href });
}

async function captureBaselineHashes() {
  const values = [];
  for (const item of manifest.specimens) {
    const [source, workspace] = await Promise.all([
      readFile(join(baselineRoot, 'sources', item.format, item.file)),
      readFile(join(baselineRoot, 'workspace', item.workspacePath ?? `${item.format}/${item.file}`)),
    ]);
    values.push({ format: item.format, file: item.file, sourceSha256: sha256(source), workspaceSha256: sha256(workspace) });
  }
  for (const [file, ref] of refs) {
    const originalBytes = await readFile(new URL(report.records.find(record => record.format === 'pdf' && record.sourceDocx === file && record.round === 10).renderResult.pdf.uri));
    values.push({ format: 'pdf-derived', file, sha256: sha256(originalBytes) });
  }
  return values;
}
const hashesBefore = await captureBaselineHashes();
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
const client = new Client({ name: 'dsh-office-b2-native-pdf-replay', version: '1.0.0' });
const transport = new StdioClientTransport({ command: nodeRuntime, args: [`${profileRoot}/lib/server.mjs`], cwd: baselineRoot, env, stderr: 'pipe' });
const records = [], serverStderr = [];
transport.stderr?.on('data', chunk => serverStderr.push(Buffer.isBuffer(chunk) ? chunk.toString('utf8') : String(chunk)));
try {
  await client.connect(transport);
  for (const specimen of specimens) {
    const file = specimen.sourceDocx.replace(/\.docx$/i, '.pdf');
    const artifactRef = refs.get(specimen.sourceDocx);
    for (let round = 1; round <= 10; round++) {
      const requestId = `b2-pdf-${file}-${String(round).padStart(2, '0')}-${randomUUID()}`;
      const start = performance.now();
      let error;
      try {
        const response = await client.callTool({ name: 'pdf_call', arguments: { operation: 'execute',
          input: { operation: 'execute', requestId, artifactRef } } });
        if (response.isError) error = response.content?.filter(item => item.type === 'text').map(item => item.text).join('\n') ?? 'DSH tool returned an error without text.';
      } catch (cause) { error = String(cause?.stack ?? cause); }
      const record = { file, round, requestId, durationMs: Math.round(performance.now() - start), ok: !error, error };
      records.push(record);
      process.stdout.write(`${record.ok ? 'PASS' : 'FAIL'} ${file} round=${round} requestId=${requestId}\n`);
    }
  }
} finally { await client.close().catch(() => {}); }

const hashesAfter = await captureBaselineHashes();
const hashStable = JSON.stringify(hashesBefore) === JSON.stringify(hashesAfter);
const groups = specimens.map(specimen => {
  const file = specimen.sourceDocx.replace(/\.docx$/i, '.pdf');
  const items = records.filter(record => record.file === file), failed = items.filter(record => !record.ok).length;
  return { file, rounds: items.length, passed: items.length - failed, failed };
});
const errors = [];
if (!hashStable) errors.push('Frozen source, workspace, or derived PDF hashes changed during read-only replay.');
if (records.length !== 50) errors.push(`Expected 50 PDF parse calls; got ${records.length}.`);
for (const group of groups) if (group.rounds !== 10 || group.failed !== 0) errors.push(`${group.file}: expected 10/10, got ${group.passed}/${group.rounds}.`);
const result = { createdAt: new Date().toISOString(), baselineRoot, isolatedData, profile: `source-built DSH profile at ${profileRoot}`,
  backend: 'DSH profile MCP stdio server', moduleId: 'pdf-office', engine: 'PDF.js', specimens: specimens.map(item => item.sourceDocx),
  groups, records, hashesBefore, hashesAfter, hashStable, serverStderr: serverStderr.join(''), status: errors.length ? 'fail' : 'pass', errors };
await writeFile(evidencePath, JSON.stringify(result, null, 2), { encoding: 'utf8', flag: 'wx' });
process.stdout.write(JSON.stringify({ evidencePath, status: result.status, calls: records.length,
  failures: records.filter(record => !record.ok).length, groups, hashStable, errors }, null, 2) + '\n');
if (errors.length) process.exitCode = 2;
