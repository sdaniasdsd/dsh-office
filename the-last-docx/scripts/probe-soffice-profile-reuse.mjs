import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtemp, mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const baselineRoot = resolve(process.argv[2] ?? '');
const outputPath = resolve(process.argv[3] ?? '');
if (!process.argv[2] || !process.argv[3]) throw new Error('Usage: node scripts/probe-soffice-profile-reuse.mjs <frozen-run-root> <new-output-json>');
const source = join(baselineRoot, 'workspace', 'docx', 'flexible-working-form.docx');
const sourceBytes = await readFile(source), sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
const sourceSha256 = sha256(sourceBytes);
const runtime = process.env.DSH_DOCX_RUNTIME ?? resolve('dist/dsh-docx/runtime/win32-x64');
const soffice = process.env.DOCX_SOFFICE ?? join(runtime, 'libreoffice', 'program', 'soffice.com');
const pdftoppm = process.env.DOCX_PDFTOPPM ?? join(runtime, 'poppler', 'poppler-26.09.0', 'Library', 'bin', 'pdftoppm.exe');
const tempRoot = await mkdtemp(join(tmpdir(), 'dsh-lo-profile-reuse-'));

function run(command, args, cwd) {
  return new Promise((resolveRun, reject) => {
    const started = performance.now();
    const child = spawn(command, args, { cwd, windowsHide: true, shell: false, stdio: ['ignore', 'ignore', 'pipe'] });
    let stderr = '';
    child.stderr.setEncoding('utf8'); child.stderr.on('data', chunk => { stderr = `${stderr}${chunk}`.slice(-4000); });
    child.once('error', reject);
    child.once('close', code => code === 0 ? resolveRun({ durationMs: Math.round(performance.now() - started), stderr })
      : reject(new Error(`${command} exited ${code}: ${stderr.trim()}`)));
  });
}

let result;
try {
  const profile = join(tempRoot, 'shared-user-profile');
  await mkdir(profile);
  const observations = [];
  for (let round = 1; round <= 2; round++) {
    const caseDir = join(tempRoot, `case-${round}`);
    await mkdir(caseDir);
    const input = join(caseDir, 'source.docx');
    await writeFile(input, sourceBytes, { flag: 'wx' });
    const conversion = await run(soffice, ['--headless', '--nologo', '--nodefault', '--norestore',
      `-env:UserInstallation=${pathToFileURL(profile).href}`, '--convert-to', 'pdf', '--outdir', caseDir, input], caseDir);
    const pdf = join(caseDir, 'source.pdf');
    const raster = await run(pdftoppm, ['-png', '-r', '120', '-f', '1', '-l', '201', pdf, join(caseDir, 'page')], caseDir);
    const images = (await readdir(caseDir)).filter(name => /^page-\d+\.png$/i.test(name)).sort();
    const pixelHashes = await Promise.all(images.map(async name => sha256(await readFile(join(caseDir, name)))));
    observations.push({ round, conversionMs: conversion.durationMs, rasterMs: raster.durationMs,
      pdfSha256: sha256(await readFile(pdf)), pageCount: images.length, pixelHashes });
  }
  result = { baselineRoot, sourceSha256, sourceUnchanged: sourceSha256 === sha256(await readFile(source)),
    method: 'two sequential isolated output directories using one reused LibreOffice UserInstallation profile; compare 120-DPI PNG hashes',
    observations, profileReusePixelStable: JSON.stringify(observations[0].pixelHashes) === JSON.stringify(observations[1].pixelHashes),
    warmConversionImprovementMs: observations[0].conversionMs - observations[1].conversionMs };
} finally { await rm(tempRoot, { recursive: true, force: true }); }

await writeFile(outputPath, JSON.stringify(result, null, 2), { encoding: 'utf8', flag: 'wx' });
process.stdout.write(JSON.stringify({ outputPath, sourceUnchanged: result.sourceUnchanged, observations: result.observations.map(({ round,
  conversionMs, rasterMs, pageCount }) => ({ round, conversionMs, rasterMs, pageCount })),
  profileReusePixelStable: result.profileReusePixelStable, warmConversionImprovementMs: result.warmConversionImprovementMs }, null, 2) + '\n');
