import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtemp, mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const baselineRoot = resolve(process.argv[2] ?? '');
const outputPath = resolve(process.argv[3] ?? '');
if (!process.argv[2] || !process.argv[3]) throw new Error('Usage: node scripts/profile-render-phases.mjs <frozen-run-root> <new-output-json>');
const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
const manifest = JSON.parse(await readFile(join(baselineRoot, 'manifest.json'), 'utf8'));
const report = JSON.parse(await readFile(join(baselineRoot, 'report.json'), 'utf8'));
const selected = [
  { file: 'flexible-working-form.docx', expectedPages: 3 },
  { file: 'asbestos-management-plan.docx', expectedPages: 9 },
];
const runtimeRoot = process.env.DSH_DOCX_RUNTIME ?? resolve('dist/dsh-docx/runtime/win32-x64');
const soffice = process.env.DOCX_SOFFICE ?? join(runtimeRoot, 'libreoffice', 'program', 'soffice.com');
const pdftoppm = process.env.DOCX_PDFTOPPM ?? join(runtimeRoot, 'poppler', 'poppler-26.09.0', 'Library', 'bin', 'pdftoppm.exe');
const tempRoot = await mkdtemp(join(tmpdir(), 'dsh-b3-render-'));
const results = [];

function run(command, args, cwd) {
  return new Promise((resolveRun, reject) => {
    const started = performance.now();
    const child = spawn(command, args, { cwd, windowsHide: true, shell: false, stdio: ['ignore', 'ignore', 'pipe'] });
    let stderr = '';
    child.stderr.setEncoding('utf8');
    child.stderr.on('data', chunk => { stderr = `${stderr}${chunk}`.slice(-4000); });
    child.once('error', reject);
    child.once('close', code => {
      if (code !== 0) reject(new Error(`${command} exited ${code}: ${stderr.trim()}`));
      else resolveRun({ durationMs: Math.round(performance.now() - started), stderr: stderr.trim() });
    });
  });
}

try {
  for (const sample of selected) {
    const spec = manifest.specimens.find(item => item.format === 'docx' && item.file === sample.file);
    if (!spec) throw new Error(`Missing frozen specimen ${sample.file}.`);
    const sourcePath = join(baselineRoot, 'workspace', spec.workspacePath ?? `docx/${sample.file}`);
    const sourceBytes = await readFile(sourcePath), sourceSha256 = sha256(sourceBytes);
    if (sourceSha256 !== spec.workspaceSha256) throw new Error(`Refusing render: frozen DOCX hash mismatch for ${sample.file}.`);
    const caseDir = join(tempRoot, sample.file.replace(/\.docx$/i, ''));
    await mkdir(caseDir, { recursive: true });
    const inputPath = join(caseDir, 'source.docx'), pdfPath = join(caseDir, 'source.pdf');
    await writeFile(inputPath, sourceBytes, { flag: 'wx' });
    const profileDir = join(caseDir, 'lo-profile');
    await mkdir(profileDir, { recursive: false });
    const conversion = await run(soffice, ['--headless', '--nologo', '--nodefault', '--norestore',
      `-env:UserInstallation=${pathToFileURL(profileDir).href}`, '--convert-to', 'pdf', '--outdir', caseDir, inputPath], caseDir);
    const pdfBytes = await readFile(pdfPath);
    const full = await run(pdftoppm, ['-png', '-r', '120', '-f', '1', '-l', '201', pdfPath, join(caseDir, 'page')], caseDir);
    const thumbnails = await run(pdftoppm, ['-png', '-scale-to', '480', '-f', '1', '-l', '201', pdfPath, join(caseDir, 'thumb')], caseDir);
    const names = await readdir(caseDir);
    const pageFiles = names.filter(name => /^page-\d+\.png$/i.test(name)).sort((a, b) => Number(a.match(/-(\d+)\.png$/i)?.[1]) - Number(b.match(/-(\d+)\.png$/i)?.[1]));
    const thumbFiles = names.filter(name => /^thumb-\d+\.png$/i.test(name)).sort((a, b) => Number(a.match(/-(\d+)\.png$/i)?.[1]) - Number(b.match(/-(\d+)\.png$/i)?.[1]));
    const baselineRecord = report.records.find(record => record.format === 'pdf' && record.sourceDocx === sample.file && record.round === 1 && record.ok);
    const baselinePageHashes = baselineRecord.renderResult.pages.map(page => page.image.sha256);
    const baselineThumbHashes = baselineRecord.renderResult.pages.map(page => page.thumbnail.sha256);
    const pageHashes = await Promise.all(pageFiles.map(async name => sha256(await readFile(join(caseDir, name)))));
    const thumbHashes = await Promise.all(thumbFiles.map(async name => sha256(await readFile(join(caseDir, name)))));
    const afterSha256 = sha256(await readFile(sourcePath));
    results.push({ file: sample.file, expectedPages: sample.expectedPages, pageCount: pageFiles.length, thumbnailCount: thumbFiles.length,
      sourceSha256, sourceUnchanged: sourceSha256 === afterSha256, pdfSha256: sha256(pdfBytes), pdfBytes: pdfBytes.length,
      phases: { libreOfficeDocxToPdfMs: conversion.durationMs, poppler120DpiPagesMs: full.durationMs,
        poppler480EdgeThumbnailsMs: thumbnails.durationMs, totalMs: conversion.durationMs + full.durationMs + thumbnails.durationMs },
      baselinePagePixelHashesMatch: JSON.stringify(pageHashes) === JSON.stringify(baselinePageHashes),
      baselineThumbnailPixelHashesMatch: JSON.stringify(thumbHashes) === JSON.stringify(baselineThumbHashes),
      pageHashes, thumbnailHashes: thumbHashes });
  }
} finally {
  await rm(tempRoot, { recursive: true, force: true });
}

const errors = [];
for (const result of results) {
  if (result.pageCount !== result.expectedPages || result.thumbnailCount !== result.expectedPages) errors.push(`${result.file}: output page count differs from baseline.`);
  if (!result.sourceUnchanged) errors.push(`${result.file}: source changed during phase profiling.`);
  if (!result.baselinePagePixelHashesMatch) errors.push(`${result.file}: 120-DPI page pixels differ from baseline.`);
  if (!result.baselineThumbnailPixelHashesMatch) errors.push(`${result.file}: thumbnail pixels differ from baseline.`);
}
const output = { createdAt: new Date().toISOString(), baselineRoot, runtime: { soffice, pdftoppm },
  method: 'one DOCX-to-PDF render and one 120-DPI full-page + 480-edge thumbnail render per selected specimen; timings are subprocess wall times',
  results, status: errors.length ? 'fail' : 'pass', errors };
await writeFile(outputPath, JSON.stringify(output, null, 2), { encoding: 'utf8', flag: 'wx' });
process.stdout.write(JSON.stringify({ outputPath, status: output.status, results: results.map(({ file, pageCount, phases,
  baselinePagePixelHashesMatch, baselineThumbnailPixelHashesMatch, sourceUnchanged }) => ({ file, pageCount, phases,
  baselinePagePixelHashesMatch, baselineThumbnailPixelHashesMatch, sourceUnchanged })), errors }, null, 2) + '\n');
if (errors.length) process.exitCode = 2;
