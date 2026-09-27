import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { createHash, randomUUID } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';

const root = resolve(process.argv[2] ?? '');
if (!process.argv[2]) throw new Error('Usage: node scripts/replay-pptx-baseline.mjs <run-root> [output-json] [baseline|fixed]');
const outputPath = resolve(process.argv[3] ?? join(root, 'b0-replay', `pptx-default-env-${Date.now()}.json`));
const mode = process.argv[4] ?? 'baseline';
if (!['baseline', 'fixed'].includes(mode)) throw new Error(`Unknown replay mode: ${mode}`);
const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
const manifest = JSON.parse(await readFile(join(root, 'manifest.json'), 'utf8'));
const report = JSON.parse(await readFile(join(root, 'report.json'), 'utf8'));
const pptxSpecs = manifest.specimens.filter(item => item.format === 'pptx');
if (pptxSpecs.length !== 5) throw new Error(`Expected exactly five PPTX specimens; got ${pptxSpecs.length}.`);
const artifactRefs = new Map(report.imports.filter(item => item.format === 'pptx').map(item => [item.file, item.artifactRef]));
if (artifactRefs.size !== 5) throw new Error(`Expected five stored DSH PPTX artifactRefs; got ${artifactRefs.size}.`);

async function captureHashes() {
  const values = [];
  for (const specimen of manifest.specimens) {
    const [source, workspace] = await Promise.all([
      readFile(join(root, 'sources', specimen.format, specimen.file)),
      readFile(join(root, 'workspace', specimen.workspacePath ?? `${specimen.format}/${specimen.file}`)),
    ]);
    values.push({ file: specimen.file, format: specimen.format,
      sourceSha256: sha256(source), workspaceSha256: sha256(workspace) });
  }
  return values;
}
const hashesBefore = await captureHashes();
for (const specimen of manifest.specimens) {
  const actual = hashesBefore.find(item => item.file === specimen.file && item.format === specimen.format);
  if (actual.sourceSha256 !== specimen.sourceSha256 || actual.workspaceSha256 !== specimen.workspaceSha256)
    throw new Error(`Refusing replay because corpus hash differs from manifest: ${specimen.format}/${specimen.file}`);
}

const profileRoot = process.env.DSH_DOCX_PROFILE_ROOT ?? 'C:/Users/AA/AppData/Roaming/com.yeagoo.dsh-desktop/harness/profiles/web/node_modules/@deepseek-ai/dsh-docx';
const nodeRuntime = process.env.DSH_NODE_RUNTIME ?? 'C:/Users/AA/AppData/Local/DSH Desktop/runtime/node.exe';
const bundleRuntime = `${profileRoot}/runtime/win32-x64`;
const env = Object.fromEntries(Object.entries(process.env).filter(([, value]) => typeof value === 'string'));
delete env.PYTHONIOENCODING;
delete env.PYTHONUTF8;
Object.assign(env, {
  THE_LAST_DOCX_WORKSPACE: join(root, 'workspace'),
  THE_LAST_DOCX_DATA: join(root, 'data'),
  DOCX_PYTHON: `${bundleRuntime}/python/python.exe`,
  DOCX_SOFFICE: `${bundleRuntime}/libreoffice/program/soffice.com`,
  DOCX_PDFTOPPM: `${bundleRuntime}/poppler/poppler-26.09.0/Library/bin/pdftoppm.exe`,
});
const expectedFailed = mode === 'baseline'
  ? new Set(['civil-service-line-management.pptx', 'timms-workshop.pptx', 'civil-society-covenant.pptx'])
  : new Set();
const client = new Client({ name: 'dsh-office-b0-pptx-replay', version: '1.0.0' });
const transport = new StdioClientTransport({ command: nodeRuntime, args: [`${profileRoot}/lib/server.mjs`],
  cwd: join(root, 'workspace'), env, stderr: 'pipe' });
const records = [];
const serverStderr = [];
transport.stderr?.on('data', chunk => serverStderr.push(Buffer.isBuffer(chunk) ? chunk.toString('utf8') : String(chunk)));
try {
  await client.connect(transport);
  for (const specimen of pptxSpecs) {
    const artifactRef = artifactRefs.get(specimen.file);
    for (let round = 1; round <= 10; round++) {
      const requestId = `b0-pptx-${specimen.file}-${String(round).padStart(2, '0')}-${randomUUID()}`;
      const startedAt = new Date().toISOString();
      const start = performance.now();
      let response;
      let error;
      try {
        response = await client.callTool({ name: 'pptx_call', arguments: {
          operation: 'execute',
          input: { operation: 'execute', requestId, artifactRef, payload: { action: 'extract' } },
        } });
        if (response.isError) error = response.content?.filter(item => item.type === 'text').map(item => item.text).join('\n') ?? 'DSH tool returned isError without text.';
      } catch (e) {
        error = String(e?.stack ?? e);
      }
      records.push({ file: specimen.file, round, requestId, startedAt,
        durationMs: Math.round(performance.now() - start), ok: !error, error });
      process.stdout.write(`${records.at(-1).ok ? 'PASS' : 'FAIL'} ${specimen.file} round=${round} requestId=${requestId}\n`);
    }
  }
} finally {
  await client.close().catch(() => {});
}
const hashesAfter = await captureHashes();
const hashStable = JSON.stringify(hashesBefore) === JSON.stringify(hashesAfter);
const groups = pptxSpecs.map(specimen => {
  const items = records.filter(record => record.file === specimen.file);
  const failed = items.filter(record => !record.ok).length;
  return { file: specimen.file, rounds: items.length, passed: items.length - failed, failed,
    expectedFailed: expectedFailed.has(specimen.file) };
});
const errors = [];
if (!hashStable) errors.push('Source/workspace hashes changed during read-only replay.');
if (records.length !== 50) errors.push(`Expected 50 PPTX replay records; got ${records.length}.`);
for (const group of groups) {
  const expectedFailures = group.expectedFailed ? 10 : 0;
  if (group.failed !== expectedFailures) errors.push(`${group.file}: expected ${expectedFailures} failures, got ${group.failed}.`);
}
for (const record of records.filter(item => !item.ok)) {
  if (!/UnicodeEncodeError|codec can't encode|ENGINE_FAILED/i.test(record.error ?? ''))
    errors.push(`${record.file} round ${record.round} failed with an unexpected error.`);
}
const result = { createdAt: new Date().toISOString(), baselineRoot: root,
  profile: mode === 'baseline' ? manifest.profile : `source-built DSH profile at ${profileRoot}`,
  mode, backend: 'DSH profile MCP stdio server',
  environment: { mode: 'default Windows locale; PYTHONIOENCODING and PYTHONUTF8 unset' },
  expectedFailureFiles: [...expectedFailed], groups, records, hashesBefore, hashesAfter, hashStable,
  serverStderr: serverStderr.join(''), status: errors.length === 0 ? 'pass' : 'fail', errors };
await mkdir(dirname(outputPath), { recursive: true });
await writeFile(outputPath, JSON.stringify(result, null, 2), { encoding: 'utf8', flag: 'wx' });
process.stdout.write(JSON.stringify({ outputPath, mode, status: result.status, total: records.length,
  failures: records.filter(item => !item.ok).length, groups, hashStable, errors }, null, 2) + '\n');
if (errors.length) process.exitCode = 2;
