import { createHash, randomUUID } from 'node:crypto';
import { spawn } from 'node:child_process';
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { basename, join, resolve } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { DocxProfile } from '../src/profile.ts';
import { LocalArtifactFiles } from '../src/files.ts';

const repo = resolve(import.meta.dirname, '..');
const sourceRoot = join(repo, 'dist', 'office-stress-20260927-visual-confirm-200', 'workspace', 'xlsx');
const sourcePath = join(sourceRoot, 'qa-assumptions-log.xlsx');
const outputRoot = join(repo, 'dist', 'repair-batch-03-xlsx-assumptions');
const runtimeRoot = process.env.DSH_OFFICE_RUNTIME
  ?? join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web', 'node_modules', '@deepseek-ai', 'dsh-docx', 'runtime', 'win32-x64');
const runtime = {
  pythonPath: join(runtimeRoot, 'python', 'python.exe'),
  sofficePath: join(runtimeRoot, 'libreoffice', 'program', 'soffice.com'),
  pdftoppmPath: join(runtimeRoot, 'poppler', 'poppler-26.09.0', 'Library', 'bin', 'pdftoppm.exe'),
};
const config = [{
  sheet: 'Model inputs and assumptions', orientation: 'landscape', fitToWidth: 2, fitToHeight: 0,
}];
const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');
const fileSha = async (path) => sha256(await readFile(path));
const artifactPath = (ref) => fileURLToPath(new URL(ref.uri));

function run(command, args, timeoutMs = 180_000) {
  return new Promise((resolveRun, reject) => {
    const child = spawn(command, args, { windowsHide: true, shell: false, stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '', stderr = '';
    const timer = setTimeout(() => { child.kill(); reject(new Error(`${basename(command)} exceeded ${timeoutMs}ms.`)); }, timeoutMs);
    child.stdout.on('data', (chunk) => { stdout += chunk.toString('utf8'); });
    child.stderr.on('data', (chunk) => { stderr += chunk.toString('utf8'); });
    child.once('error', (error) => { clearTimeout(timer); reject(error); });
    child.once('close', (code) => { clearTimeout(timer); code === 0 ? resolveRun({ stdout, stderr }) : reject(new Error(`${basename(command)} exited ${code}: ${stderr.slice(-1200)}`)); });
  });
}

async function render(candidatePath) {
  const renderRoot = join(outputRoot, 'render'), pageDir = join(renderRoot, 'pages'), pdfDir = join(renderRoot, 'pdf');
  const profileDir = join(tmpdir(), `dsh-xlsx-b3-${process.pid}-${randomUUID()}`);
  await Promise.all([mkdir(pageDir, { recursive: true }), mkdir(pdfDir, { recursive: true }), mkdir(profileDir)]);
  try {
    await run(runtime.sofficePath, [`-env:UserInstallation=${pathToFileURL(profileDir).href}`, '--headless', '--nologo', '--nodefault', '--norestore', '--convert-to', 'pdf', '--outdir', pdfDir, candidatePath]);
    const pdfPath = join(pdfDir, 'candidate.pdf');
    await run(runtime.pdftoppmPath, ['-r', '120', '-png', pdfPath, join(pageDir, 'page')]);
  } finally {
    await rm(profileDir, { recursive: true, force: true });
  }
  const { stdout } = await run(runtime.pythonPath, ['-c', [
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
  ].join('\n'), pageDir]);
  const pageNames = (await import('node:fs/promises')).readdir(pageDir);
  const pages = (await pageNames).filter((name) => /^page-\d+\.png$/.test(name)).sort((a, b) => Number(a.match(/\d+/)?.[0]) - Number(b.match(/\d+/)?.[0]));
  return {
    pdfPath: join(pdfDir, 'candidate.pdf'), pdfSha256: await fileSha(join(pdfDir, 'candidate.pdf')),
    pageCount: pages.length,
    pages: await Promise.all(pages.map(async (name, index) => ({ pageNumber: index + 1, path: join(pageDir, name), sha256: await fileSha(join(pageDir, name)) }))),
    contactSheets: stdout.trim().split(/\r?\n/).filter(Boolean),
  };
}

await mkdir(outputRoot, { recursive: true });
const sourceBytes = await readFile(sourcePath), sourceHash = sha256(sourceBytes);
const store = await LocalArtifactFiles.create(join(outputRoot, 'store'), [sourceRoot]);
const profile = new DocxProfile({ files: store, pythonPath: runtime.pythonPath });
try {
  const source = await store.importFile(sourcePath);
  const edited = await profile.call('xlsx-office', 'execute', {
    artifactRef: source,
    requestId: 'repair-b3-xlsx-assumptions-print-layout',
    payload: { action: 'setPrintLayout', sheets: config },
  });
  const candidateBytes = await store.read(edited.result.artifactRef);
  const candidatePath = join(outputRoot, 'candidate.xlsx');
  await writeFile(candidatePath, candidateBytes);
  const rendered = await render(candidatePath);
  const sourceUnchanged = await fileSha(sourcePath) === sourceHash;
  if (!sourceUnchanged) throw new Error('Frozen source workbook changed during Batch 3.');
  const manifest = {
    batch: 'B3 — XLSX assumptions log pagination',
    createdAt: new Date().toISOString(),
    attentionLock: 'Only the Model inputs and assumptions sheet print settings may change; preserve all cells, names, links, and untouched OOXML parts. Leave Welcome unchanged.',
    primaryRepairInstruction: 'Reduce horizontal fragmentation while keeping the wide form table at a readable scale; fit to at most two landscape pages across and do not force content onto one page vertically.',
    renderer: { engine: 'LibreOffice + Poppler review-only (not DSH rendering)', ...runtime, pageImageDpi: 120 },
    source: { path: sourcePath, sha256: sourceHash, unchangedAfterRun: sourceUnchanged },
    candidate: { path: candidatePath, sha256: sha256(candidateBytes), bytes: candidateBytes.length },
    dsh: {
      requestId: edited.requestId, moduleId: edited.moduleId, changedPackageParts: edited.result.changedPackageParts,
      packageScopedWrite: edited.result.packageScopedWrite,
      untouchedPackagePartsPreservedByteForByte: edited.result.untouchedPackagePartsPreservedByteForByte,
      sheetContentPreserved: edited.result.sheetContentPreserved,
      unsupportedFeatures: edited.result.unsupportedFeatures,
      worksheetCount: edited.result.worksheetCount, populatedCellCount: edited.result.populatedCellCount,
    },
    appliedLayout: config,
    renderedCandidate: rendered,
  };
  const manifestPath = join(outputRoot, 'manifest.json');
  await writeFile(manifestPath, JSON.stringify(manifest, null, 2) + '\n', 'utf8');
  console.log(JSON.stringify({ manifestPath, sourceHash, candidateHash: manifest.candidate.sha256,
    pages: rendered.pageCount, partsChanged: edited.result.changedPackageParts.map((item) => item.part),
    unsupportedFeatures: edited.result.unsupportedFeatures }, null, 2));
} finally {
  await profile.dispose();
}
