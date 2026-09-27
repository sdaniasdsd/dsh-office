import { createHash } from 'node:crypto';
import { mkdir, readFile, stat, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';

const root = resolve(process.argv[2] ?? '');
if (!process.argv[2]) throw new Error('Usage: node scripts/verify-office-stress-baseline.mjs <run-root> [output-json]');
const outputPath = resolve(process.argv[3] ?? join(root, 'b0-baseline-verification.json'));
const manifest = JSON.parse(await readFile(join(root, 'manifest.json'), 'utf8'));
const report = JSON.parse(await readFile(join(root, 'report.json'), 'utf8'));
const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
const specimens = manifest.specimens ?? [];
const errors = [];
const observed = [];

if (specimens.length !== 15) errors.push(`Expected 15 specimens; found ${specimens.length}.`);
for (const format of ['docx', 'pptx', 'xlsx']) {
  const count = specimens.filter(specimen => specimen.format === format).length;
  if (count !== 5) errors.push(`Expected 5 ${format.toUpperCase()} specimens; found ${count}.`);
}

for (const specimen of specimens) {
  const sourcePath = join(root, 'sources', specimen.format, specimen.file);
  const workspacePath = join(root, 'workspace', specimen.workspacePath ?? `${specimen.format}/${specimen.file}`);
  const [sourceBytes, workspaceBytes] = await Promise.all([readFile(sourcePath), readFile(workspacePath)]);
  const sourceSha256 = sha256(sourceBytes);
  const workspaceSha256 = sha256(workspaceBytes);
  if (sourceSha256 !== specimen.sourceSha256) errors.push(`Source SHA-256 mismatch: ${specimen.format}/${specimen.file}`);
  if (workspaceSha256 !== specimen.workspaceSha256) errors.push(`Workspace SHA-256 mismatch: ${specimen.format}/${specimen.file}`);
  observed.push({ format: specimen.format, file: specimen.file, sourceSha256, workspaceSha256,
    sourceSizeBytes: sourceBytes.length, workspaceSizeBytes: workspaceBytes.length,
    sourceMatchesManifest: sourceSha256 === specimen.sourceSha256,
    workspaceMatchesManifest: workspaceSha256 === specimen.workspaceSha256,
    workspaceSanitization: specimen.sanitization ?? null });
}

const expectedByFormat = { docx: 50, pdf: 50, pptx: 50, xlsx: 50 };
const roundsByFormat = Object.fromEntries(Object.entries(expectedByFormat).map(([format, expected]) => {
  const actual = report.records.filter(record => record.format === format).length;
  if (actual !== expected) errors.push(`Expected ${expected} ${format.toUpperCase()} operation records; found ${actual}.`);
  return [format, { expected, actual }];
}));
if (report.records.length !== 200) errors.push(`Expected 200 total operation records; found ${report.records.length}.`);
if (report.actualCalls !== 231) errors.push(`Expected 231 DSH calls in the main run; found ${report.actualCalls}.`);
if (report.imports.length !== 15) errors.push(`Expected 15 DSH imports; found ${report.imports.length}.`);
if (report.preflights.length !== 15) errors.push(`Expected 15 DSH preflights; found ${report.preflights.length}.`);

const pptxFailures = report.failures.filter(item => item.format === 'pptx');
const failedByFile = Object.fromEntries(specimens.filter(s => s.format === 'pptx').map(s => [s.file,
  pptxFailures.filter(item => item.file === s.file).length]));
if (pptxFailures.length !== 30) errors.push(`Expected the 30 recorded PPTX baseline failures; found ${pptxFailures.length}.`);
for (const file of ['civil-service-line-management.pptx', 'timms-workshop.pptx', 'civil-society-covenant.pptx']) {
  if (failedByFile[file] !== 10) errors.push(`Expected 10 PPTX failures for ${file}; found ${failedByFile[file] ?? 0}.`);
}
for (const file of ['qualifications-reform.pptx', 'prevent-duty-leadership.pptx']) {
  if (failedByFile[file] !== 0) errors.push(`Expected zero PPTX failures for ${file}; found ${failedByFile[file] ?? 0}.`);
}
for (const failure of pptxFailures) {
  const record = report.records.find(item => item.format === 'pptx' && item.file === failure.file &&
    item.round === failure.round && !item.ok);
  if (!record?.requestId) errors.push(`PPTX failure has no requestId in its operation record: ${failure.file} round ${failure.round}.`);
  if (record?.error !== failure.error) errors.push(`Failure summary differs from operation record: ${failure.file} round ${failure.round}.`);
  if (!/UnicodeEncodeError|codec can't encode|ENGINE_FAILED/i.test(failure.error ?? '')) {
    errors.push(`PPTX failure is not the expected encoding failure: ${failure.file} round ${failure.round}.`);
  }
}

const verification = {
  createdAt: new Date().toISOString(),
  baselineRoot: root,
  profile: manifest.profile,
  roundsPerMaterial: manifest.rounds,
  dshCalls: report.actualCalls,
  totalOperationRecords: report.records.length,
  operationRecordsByFormat: roundsByFormat,
  pptxFailureCount: pptxFailures.length,
  pptxFailureDetailsLinkedToRequestIds: pptxFailures.every(item => report.records.some(record => record.format === 'pptx' &&
    record.file === item.file && record.round === item.round && !record.ok && Boolean(record.requestId) && record.error === item.error)),
  pptxFailuresByFile: failedByFile,
  specimens: observed,
  status: errors.length === 0 ? 'pass' : 'fail',
  errors,
};
await mkdir(dirname(outputPath), { recursive: true });
await writeFile(outputPath, JSON.stringify(verification, null, 2), { encoding: 'utf8', flag: 'wx' });
process.stdout.write(JSON.stringify({ outputPath, status: verification.status, specimens: observed.length,
  roundsByFormat, calls: verification.dshCalls, pptxFailures: verification.pptxFailureCount, errors }, null, 2) + '\n');
if (errors.length) process.exitCode = 2;
