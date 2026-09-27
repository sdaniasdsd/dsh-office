import { createHash, randomUUID } from 'node:crypto';
import { spawn } from 'node:child_process';
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { basename, join, resolve } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { DocxProfile } from '../src/profile.ts';
import { LocalArtifactFiles } from '../src/files.ts';

const repo = resolve(import.meta.dirname, '..');
const runRoot = join(repo, 'dist', 'office-stress-20260927-visual-confirm-200');
const sourceRoot = join(runRoot, 'workspace', 'xlsx');
const evidenceRoot = join(repo, 'dist', 'repair-batch-02-xlsx');
const runtimeRoot = process.env.DSH_OFFICE_RUNTIME
  ?? join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web', 'node_modules', '@deepseek-ai', 'dsh-docx', 'runtime', 'win32-x64');
const runtime = {
  pythonPath: join(runtimeRoot, 'python', 'python.exe'),
  sofficePath: join(runtimeRoot, 'libreoffice', 'program', 'soffice.com'),
  pdftoppmPath: join(runtimeRoot, 'poppler', 'poppler-26.09.0', 'Library', 'bin', 'pdftoppm.exe'),
};
const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');
const fileSha = async (path) => sha256(await readFile(path));
const artifactPath = (ref) => fileURLToPath(new URL(ref.uri));
const sheets = {
  'condition-survey-template.xlsx': [
    'Full Elemental List', 'Site Level', 'Block Level', 'Room Level',
  ].map((sheet) => ({ sheet, orientation: 'landscape', fitToWidth: 1, fitToHeight: 0 })),
  'qa-modelling-template.xlsx': [
    'Logs', 'Units', 'Scenarios >>', 'Inputs>>', 'Calculations >>', 'Outputs >>', 'Lookups',
  ].map((sheet) => ({ sheet, orientation: 'landscape', fitToWidth: 1, fitToHeight: sheet === 'Lookups' ? 1 : 0 })),
};

function run(command, args, timeoutMs = 180_000) {
  return new Promise((resolveRun, reject) => {
    const child = spawn(command, args, { windowsHide: true, shell: false, stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '', stderr = '';
    const timer = setTimeout(() => { child.kill(); reject(new Error(`${basename(command)} exceeded ${timeoutMs}ms.`)); }, timeoutMs);
    child.stdout.on('data', (chunk) => { stdout += chunk.toString('utf8'); });
    child.stderr.on('data', (chunk) => { stderr += chunk.toString('utf8'); });
    child.once('error', (error) => { clearTimeout(timer); reject(error); });
    child.once('close', (code) => {
      clearTimeout(timer);
      if (code !== 0) reject(new Error(`${basename(command)} exited ${code}: ${stderr.slice(-1200)}`));
      else resolveRun({ stdout, stderr });
    });
  });
}

async function exportPreview(xlsxPath, outputDir, id) {
  const pdfDir = join(outputDir, 'pdf');
  const pageDir = join(outputDir, 'pages');
  const profile = join(tmpdir(), `dsh-xlsx-preview-${id}-${process.pid}-${randomUUID()}`);
  await mkdir(profile, { recursive: false });
  await mkdir(pdfDir, { recursive: true });
  await mkdir(pageDir, { recursive: true });
  const uri = pathToFileURL(profile).href;
  try {
    await run(runtime.sofficePath, [`-env:UserInstallation=${uri}`, '--headless', '--nologo', '--nodefault', '--norestore', '--convert-to', 'pdf', '--outdir', pdfDir, xlsxPath]);
    const pdfPath = join(pdfDir, `${basename(xlsxPath, '.xlsx')}.pdf`);
    await run(runtime.pdftoppmPath, ['-r', '120', '-png', pdfPath, join(pageDir, 'page')]);
    const pageNames = (await import('node:fs/promises')).readdir(pageDir);
    const pages = (await pageNames).filter((name) => /^page-\d+\.png$/.test(name)).sort((a, b) => Number(a.match(/\d+/)?.[0]) - Number(b.match(/\d+/)?.[0]));
    if (!pages.length) throw new Error(`No page images were produced for ${xlsxPath}.`);
    const rendered = { pdfPath, pdfSha256: await fileSha(pdfPath), pageCount: pages.length, pages: await Promise.all(pages.map(async (name, index) => ({
      pageNumber: index + 1, path: join(pageDir, name), sha256: await fileSha(join(pageDir, name)),
    }))) };
    return rendered;
  } finally {
    await rm(profile, { recursive: true, force: true });
  }
}

async function makeContactSheets(pageDir) {
  const code = [
    'from pathlib import Path', 'from PIL import Image, ImageDraw', 'import sys',
    'root=Path(sys.argv[1]); files=sorted(root.glob("page-*.png"), key=lambda p:int(p.stem.split("-")[-1]))',
    'out=[]; chunk_size=12',
    'for chunk_no,start in enumerate(range(0,len(files),chunk_size),1):',
    ' chunk=files[start:start+chunk_size]; cols=3; rows=(len(chunk)+2)//3; w,h,label,gap,margin=340,480,28,12,12',
    ' canvas=Image.new("RGB",(margin*2+cols*w+gap*(cols-1),margin*2+rows*(h+label)+gap*(rows-1)),"white"); draw=ImageDraw.Draw(canvas)',
    ' for pos,path in enumerate(chunk):',
    '  im=Image.open(path).convert("RGB"); im.thumbnail((w,h)); x=margin+(pos%cols)*(w+gap); y=margin+(pos//cols)*(h+label+gap); canvas.paste(im,(x,y+label)); draw.text((x,y),f"Page {int(path.stem.split(chr(45))[-1])}",fill="black")',
    ' target=root.parent/f"contact-{chunk_no:02d}.png"; canvas.save(target); out.append(str(target))',
    'print("\\n".join(out))',
  ].join('\n');
  const { stdout } = await run(runtime.pythonPath, ['-c', code, pageDir]);
  return stdout.trim().split(/\r?\n/).filter(Boolean);
}

await mkdir(evidenceRoot, { recursive: true });
const store = await LocalArtifactFiles.create(join(evidenceRoot, 'store'), [sourceRoot]);
const profile = new DocxProfile({ files: store, pythonPath: runtime.pythonPath });
const results = [];
try {
  for (const [filename, config] of Object.entries(sheets)) {
    const sourcePath = join(sourceRoot, filename);
    const sourceBytes = await readFile(sourcePath);
    const sourceHash = sha256(sourceBytes);
    const sourceRef = await store.importFile(sourcePath);
    const edited = await profile.call('xlsx-office', 'execute', {
      artifactRef: sourceRef,
      requestId: `repair-b2-xlsx-print-layout-${filename.replaceAll(/[^a-z0-9]+/gi, '-').toLowerCase()}`,
      payload: { action: 'setPrintLayout', sheets: config },
    });
    const result = edited.result;
    const candidateRef = result.artifactRef;
    const candidateBytes = await store.read(candidateRef);
    const caseRoot = join(evidenceRoot, basename(filename, '.xlsx'));
    await mkdir(caseRoot, { recursive: true });
    const candidatePath = join(caseRoot, 'candidate.xlsx');
    await writeFile(candidatePath, candidateBytes);
    const rendered = await exportPreview(candidatePath, join(caseRoot, 'render'), 'candidate');
    rendered.contactSheets = await makeContactSheets(join(caseRoot, 'render', 'pages'));
    const sourceUnchanged = await fileSha(sourcePath) === sourceHash;
    if (!sourceUnchanged) throw new Error(`Frozen source changed during repair: ${filename}`);
    results.push({
      filename, source: { path: sourcePath, sha256: sourceHash, unchangedAfterRun: sourceUnchanged },
      candidate: { path: candidatePath, sha256: sha256(candidateBytes), bytes: candidateBytes.length },
      dsh: { requestId: edited.requestId, moduleId: edited.moduleId, changedPackageParts: result.changedPackageParts,
        packageScopedWrite: result.packageScopedWrite, untouchedPackagePartsPreservedByteForByte: result.untouchedPackagePartsPreservedByteForByte,
        sheetContentPreserved: result.sheetContentPreserved,
        unsupportedFeatures: result.unsupportedFeatures, worksheetCount: result.worksheetCount, populatedCellCount: result.populatedCellCount },
      appliedLayout: config,
      renderedCandidate: rendered,
    });
    console.log(JSON.stringify({ filename, sourceHash, candidateHash: sha256(candidateBytes), pages: rendered.pageCount,
      partsChanged: result.changedPackageParts.map((item) => item.part), unsupportedFeatures: result.unsupportedFeatures }));
  }
} finally {
  await profile.dispose();
}

const manifest = {
  batch: 'B2 — XLSX print preview pagination',
  createdAt: new Date().toISOString(),
  attentionLock: 'Only explicit print settings on the reviewed sheets may change: landscape, one page wide, unlimited height; preserve cell data and every non-target package part.',
  primaryRepairInstruction: 'Reduce horizontal print fragments without scaling a sheet to one page tall; retain vertical pages and keep the frozen input workbooks unchanged.',
  renderer: { engine: 'LibreOffice + Poppler review-only (not DSH rendering)', ...runtime, pageImageDpi: 120 },
  cases: results,
};
const manifestPath = join(evidenceRoot, 'manifest.json');
await writeFile(manifestPath, JSON.stringify(manifest, null, 2) + '\n', 'utf8');
console.log(JSON.stringify({ manifestPath, cases: results.length, pages: results.map((item) => ({ file: item.filename, pages: item.renderedCandidate.pageCount })) }, null, 2));
