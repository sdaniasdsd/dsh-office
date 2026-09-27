import { createHash, randomUUID } from 'node:crypto';
import { copyFile, mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { basename, dirname, extname, join, resolve } from 'node:path';
import { spawn } from 'node:child_process';
import { tmpdir } from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { tableTargetFromDualIR, targetFromDualIR } from '@dsh-office-profile/docx-edit';
import { DocxProfile } from '../src/profile.ts';
import { LocalArtifactFiles } from '../src/files.ts';

const repo = resolve(import.meta.dirname, '..');
const runRoot = join(repo, 'dist', 'office-aesthetic-50');
const workspace = join(repo, 'dist', 'office-stress-20260927-visual-confirm-200', 'workspace');
const probePath = join(runRoot, 'probe.json');
const batch = process.argv[2] ?? 'all';
if (!['all', 'docx', 'pptx', 'xlsx'].includes(batch)) throw new Error('Usage: npx tsx scripts/office-aesthetic-50.mjs [all|docx|pptx|xlsx]');
const idFilter = process.argv[3] ?? null;
const runtimeRoot = process.env.DSH_OFFICE_RUNTIME
  ?? join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web', 'node_modules', '@deepseek-ai', 'dsh-docx', 'runtime', 'win32-x64');
const runtime = {
  pythonPath: join(runtimeRoot, 'python', 'python.exe'),
  sofficePath: join(runtimeRoot, 'libreoffice', 'program', 'soffice.com'),
  pdftoppmPath: join(runtimeRoot, 'poppler', 'poppler-26.09.0', 'Library', 'bin', 'pdftoppm.exe'),
};
const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');
const fileSha = async (path) => sha256(await readFile(path));
const filesByFormat = {
  docx: ['asbestos-management-plan.docx', 'flexible-working-form.docx', 'mhra-applicant-response.docx', 'pipl-law.docx', 'qa-evidence-report-template.docx'],
  pptx: ['civil-service-line-management.pptx', 'civil-society-covenant.pptx', 'prevent-duty-leadership.pptx', 'qualifications-reform.pptx', 'timms-workshop.pptx'],
  xlsx: ['condition-survey-template.xlsx', 'green-book-appraisal-tables.xlsx', 'green-book-discount-factors.xlsx', 'qa-assumptions-log.xlsx', 'qa-modelling-template.xlsx'],
};
const titles = {
  'asbestos-management-plan.docx': 'asbestos management plan',
  'flexible-working-form.docx': 'flexible-working application form',
  'mhra-applicant-response.docx': 'MHRA applicant response',
  'pipl-law.docx': 'Personal Information Protection Law',
  'qa-evidence-report-template.docx': 'QA evidence report',
  'civil-service-line-management.pptx': 'line management standards deck',
  'civil-society-covenant.pptx': 'Civil Society Covenant deck',
  'prevent-duty-leadership.pptx': 'Prevent Duty leadership deck',
  'qualifications-reform.pptx': 'qualifications reform deck',
  'timms-workshop.pptx': 'Timms Review workshop deck',
  'condition-survey-template.xlsx': 'condition survey workbook',
  'green-book-appraisal-tables.xlsx': 'Green Book appraisal tables',
  'green-book-discount-factors.xlsx': 'discount factors workbook',
  'qa-assumptions-log.xlsx': 'assumptions log',
  'qa-modelling-template.xlsx': 'modelling QA workbook',
};

const docxPrompts = [
  (name) => `这份${name}下周要给主管过目，现在标题和章节看起来像一整块。帮我把主标题、主要小节做出清楚但克制的层级，留白舒服一点，用稳重的蓝灰色；正文、条文、页码和表格里的字一个都别改。`,
  (name) => `同事说这份${name}打印后不好扫读。请把已有表格的边线和单元格留白整理得轻一点、整齐一点；确实没有表格时就把字段标签和小标题层级拉开。空白栏位、原文和数字全部保留。`,
  (name) => `我想让这份${name}读起来松一点，不要每段都挤在一起。请调整开头几段的行距和段后距，并把章节题头做得更醒目；不要重写内容，也别把整份文件弄成花哨宣传册。`,
  (name) => `准备把${name}放进正式评审包，麻烦做一轮统一的轻量美化：标题、章节、小表格的视觉语言协调，正文仍然严肃易读；所有法律/安全/审批措辞、字段、数字和顺序保持不变。`,
];
const pptxPrompts = [
  (name) => `这份${name}首页文字有点压画面，开场和封面能不能收紧一点？保留核心名称和必要信息，别碰政策/法律名称、数字、后面页面、配色或版式。`,
  (name) => `这份${name}目录页扫起来有点费劲，帮我把目录标题和一两个导航词改得短、直白一点；内容顺序、意思、字号和现有主题都保持原样。`,
  (name) => `这份${name}要在会议室投屏，整体看着偏旧。请统一主题字体，拉开标题和正文的字号层次，增加重点的视觉对比；不要改任何文字、图表和政策信息。`,
];
const xlsxPrompts = [
  (name) => `这份${name}要拿去会议上打印，宽表现在横向散成好几页。请把指定输入/汇总表设成横向、一页宽，纵向可多页；数据、公式、空白录入区和其余工作表都别动。`,
  (name) => `同事打印${name}时嫌表格太挤。优先让关键工作表横向阅读、宽度最多两页，不要为了页数把纵向内容压成一页；保留原有数据和计算，封面和说明页不调整。`,
  (name) => `能不能把${name}的表头做成更清楚的深色底白字，加细边框和统一字体？只动工作表外观，任何数值、公式、链接、名称和分页内容都要原封不动。`,
];

const pptxReplacements = {
  'civil-service-line-management.pptx': [
    [
      { slideNumber: 1, text: 'The Civil Service', replaceWith: '' },
      { slideNumber: 1, text: 'Core stage', replaceWith: 'Core' },
    ],
    [
      { slideNumber: 3, text: 'Contents', replaceWith: 'At a glance' },
      { slideNumber: 3, text: 'What’s next?', replaceWith: 'Next steps' },
    ],
  ],
  'civil-society-covenant.pptx': [
    [
      { slideNumber: 1, text: 'This pack contains information about the Civil Society Covenant, including:', replaceWith: 'Introducing the Civil Society Covenant:' },
      { slideNumber: 1, text: 'Background', replaceWith: 'Context' },
    ],
    [
      { slideNumber: 2, text: 'Slide pack content', replaceWith: 'Contents' },
      { slideNumber: 2, text: 'What is the Civil Society Covenant', replaceWith: 'The Civil Society Covenant' },
    ],
  ],
  'prevent-duty-leadership.pptx': [
    [
      { slideNumber: 1, text: 'Policy Into Practice: ', replaceWith: 'Prevent Duty: ' },
      { slideNumber: 1, text: 'The Leadership Challenge', replaceWith: 'Leadership in practice' },
    ],
    [
      { slideNumber: 2, text: 'Outline', replaceWith: 'At a glance' },
      { slideNumber: 3, text: 'Sector-specific guidance', replaceWith: 'Sector guidance' },
    ],
  ],
  'qualifications-reform.pptx': [
    [
      { slideNumber: 1, text: 'GCSE, AS and A level reforms in England', replaceWith: 'GCSE, AS and A level reforms — England' },
      { slideNumber: 1, text: 'Updated May 2019 ', replaceWith: 'Updated: May 2019' },
    ],
    [
      { slideNumber: 2, text: 'Contents', replaceWith: 'At a glance' },
      { slideNumber: 2, text: 'What is happening, when?', replaceWith: 'Timeline' },
    ],
  ],
  'timms-workshop.pptx': [
    [
      { slideNumber: 1, text: 'Facilitator note: Before you begin', replaceWith: 'Facilitator checklist' },
      { slideNumber: 1, text: 'Before you begin, please make sure you have:', replaceWith: 'Before the workshop:' },
    ],
    [
      { slideNumber: 2, text: 'Welcome to our Timms Review workshop', replaceWith: 'Timms Review workshop' },
      { slideNumber: 2, text: '[Add details, for example time and date]', replaceWith: '[Add date and time]' },
    ],
  ],
};

const xlsxSheets = {
  'condition-survey-template.xlsx': [
    ['Full Elemental List', 'Site Level', 'Block Level', 'Room Level'],
    ['Site Level', 'Block Level'],
  ],
  'green-book-appraisal-tables.xlsx': [
    ['Shortlist AST', 'Detailed option summary table', 'Value for money summary'],
    ['Shortlist AST', 'Value for money summary'],
  ],
  'green-book-discount-factors.xlsx': [
    ['Standard Discount Factors', 'Health Discount Factors'],
    ['Standard Discount Factors'],
  ],
  'qa-assumptions-log.xlsx': [
    ['Model inputs and assumptions'],
    ['Model inputs and assumptions'],
  ],
  'qa-modelling-template.xlsx': [
    ['Logs', 'Units', 'Inputs>>', 'Calculations >>', 'Outputs >>'],
    ['Units', 'Lookups'],
  ],
};

const cases = [];
for (const file of filesByFormat.docx) for (let variant = 0; variant < 4; variant++) {
  cases.push({ id: `D${String(cases.length + 1).padStart(2, '0')}`, format: 'docx', file, variant,
    prompt: docxPrompts[variant](titles[file]), operationFamily: ['heading-hierarchy', 'table-or-label-polish', 'opening-readability', 'combined-review-pack'][variant],
    expectedEffects: variant === 1 ? ['Table/field structure is easier to scan through restrained spacing and borders.', 'Blank input cells and labels remain present.'] : ['Title and section hierarchy is visibly clearer.', 'Text, values, labels, legal/safety wording, and order remain unchanged.'],
    invariants: ['Preserve source wording, facts, fields, values, and order.', 'Keep all pages and page numbering; no clipping or new blank pages.', 'Use DSH in-place formatting only; do not rebuild the source from a plan.'] });
}
for (const file of filesByFormat.pptx) for (let variant = 0; variant < 3; variant++) {
  cases.push({ id: `P${String(cases.filter((item) => item.format === 'pptx').length + 1).padStart(2, '0')}`, format: 'pptx', file, variant,
    prompt: pptxPrompts[variant](titles[file]), operationFamily: ['cover-copy-trim', 'navigation-copy-trim', 'visual-theme-restyle'][variant],
    expectedEffects: variant === 2 ? ['Theme font and title/body hierarchy visibly improve at presentation scale.', 'All slide text, chart/table contents, and policy facts remain unchanged.'] : ['Requested cover/navigation labels become shorter and scan more cleanly.', 'Existing slide theme, layout, shapes, and unrequested wording are preserved.'],
    invariants: ['Do not alter legal/policy references, figures, charts, or slide order.', 'Keep all slides; no clipped or missing content.', 'Every replacement must use an exact source-revision run address.'] });
}
for (const file of filesByFormat.xlsx) for (let variant = 0; variant < 3; variant++) {
  cases.push({ id: `X${String(cases.filter((item) => item.format === 'xlsx').length + 1).padStart(2, '0')}`, format: 'xlsx', file, variant,
    prompt: xlsxPrompts[variant](titles[file]), operationFamily: ['print-width-layout', 'multi-sheet-readability', 'cell-visual-style'][variant],
    expectedEffects: variant === 2 ? ['Headers gain a coherent high-contrast font/fill/border treatment.', 'All values, formulas, links, names, and print content remain unchanged.'] : ['Requested worksheets print in the requested landscape width with unlimited vertical pagination unless a two-page width is specified.', 'All cell content and unrelated worksheets remain unchanged.'],
    invariants: ['Preserve formulas, cached values, styles, names, external relationships, and populated cells.', 'Do not invent a print area, hide cells, or remove blank form-entry regions.', 'Produce a separate immutable candidate and visually inspect every printed page.'] });
}
if (cases.length !== 50) throw new Error(`Case plan has ${cases.length} cases, expected exactly 50.`);

function run(command, args, timeoutMs = 240_000) {
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
async function atomicJson(path, value) {
  await mkdir(dirname(path), { recursive: true });
  const temp = `${path}.${process.pid}.${randomUUID()}.tmp`;
  await writeFile(temp, JSON.stringify(value, null, 2) + '\n', 'utf8');
  const { rename } = await import('node:fs/promises');
  await rename(temp, path);
}
const artifactPath = (ref) => fileURLToPath(new URL(ref.uri));
async function copyArtifact(store, ref, outputPath) {
  await mkdir(dirname(outputPath), { recursive: true });
  const bytes = await store.read(ref);
  await writeFile(outputPath, bytes);
  return { path: outputPath, sha256: sha256(bytes), bytes: bytes.length };
}
async function createContact(pageDir, contactPath) {
  const { stdout } = await run(runtime.pythonPath, ['-c', [
    'from pathlib import Path', 'from PIL import Image, ImageDraw', 'import sys',
    'root=Path(sys.argv[1]); out=Path(sys.argv[2]); files=sorted(root.glob("page-*.png"), key=lambda p:int(p.stem.split("-")[-1]))',
    'cols=4; w,h,label,gap,margin=300,230,25,10,12; rows=(len(files)+cols-1)//cols',
    'canvas=Image.new("RGB",(margin*2+cols*w+gap*(cols-1),margin*2+rows*(h+label+gap)),"white"); draw=ImageDraw.Draw(canvas)',
    'for pos,path in enumerate(files):',
    ' im=Image.open(path).convert("RGB"); im.thumbnail((w,h)); x=margin+(pos%cols)*(w+gap); y=margin+(pos//cols)*(h+label+gap); canvas.paste(im,(x,y+label)); draw.text((x,y),f"Page {int(path.stem.split(chr(45))[-1])}",fill="black")',
    'out.parent.mkdir(parents=True,exist_ok=True); canvas.save(out); print(len(files))',
  ].join('\n'), pageDir, contactPath]);
  return Number(stdout.trim());
}
async function renderWithLibreOffice(inputPath, renderRoot) {
  const pdfDir = join(renderRoot, 'pdf'), pageDir = join(renderRoot, 'pages');
  const profileDir = join(tmpdir(), `dsh-office-aesthetic-${process.pid}-${randomUUID()}`);
  await Promise.all([mkdir(pdfDir, { recursive: true }), mkdir(pageDir, { recursive: true }), mkdir(profileDir)]);
  try {
    await run(runtime.sofficePath, [`-env:UserInstallation=${pathToFileURL(profileDir).href}`, '--headless', '--nologo', '--nodefault', '--norestore', '--convert-to', 'pdf', '--outdir', pdfDir, inputPath]);
    const pdfPath = join(pdfDir, `${basename(inputPath, extname(inputPath))}.pdf`);
    await run(runtime.pdftoppmPath, ['-r', '120', '-png', pdfPath, join(pageDir, 'page')]);
    const pages = (await readdir(pageDir)).filter((name) => /^page-\d+\.png$/.test(name)).sort((a, b) => Number(a.match(/\d+/)?.[0]) - Number(b.match(/\d+/)?.[0]));
    return { pdfPath, pdfSha256: await fileSha(pdfPath), pageDir, pageCount: pages.length,
      pages: await Promise.all(pages.map(async (name, index) => ({ pageNumber: index + 1, path: join(pageDir, name), sha256: await fileSha(join(pageDir, name)) }))) };
  } finally { await rm(profileDir, { recursive: true, force: true }); }
}
async function materializeExternalRender(rendered, targetRoot) {
  const pagesDir = join(targetRoot, 'pages');
  await mkdir(pagesDir, { recursive: true });
  const pages = [];
  for (const page of rendered.pages) {
    const path = join(pagesDir, `page-${String(page.pageNumber).padStart(3, '0')}.png`);
    await copyFile(page.path, path);
    pages.push({ pageNumber: page.pageNumber, path, sha256: await fileSha(path) });
  }
  const pdfPath = join(targetRoot, 'artifact.pdf');
  await copyFile(rendered.pdfPath, pdfPath);
  const contactPath = join(targetRoot, 'contact-01.png');
  const pageCount = await createContact(pagesDir, contactPath);
  if (pageCount !== pages.length || pageCount !== rendered.pageCount) throw new Error('Rendered page inventory mismatch.');
  return { pdf: { path: pdfPath, sha256: await fileSha(pdfPath) }, pageCount, pages, contactSheet: contactPath };
}
async function renderDocx(profile, ref, store, targetRoot, requestId) {
  const rendered = await profile.call('docx-render', 'execute', { artifactRef: ref, requestId });
  const pagesDir = join(targetRoot, 'pages');
  await mkdir(pagesDir, { recursive: true });
  const pages = [];
  for (const page of rendered.result.pages) pages.push(await copyArtifact(store, page.image, join(pagesDir, `page-${String(page.pageNumber).padStart(3, '0')}.png`)).then((file) => ({ pageNumber: page.pageNumber, ...file })));
  const pdf = rendered.result.pdf ? await copyArtifact(store, rendered.result.pdf, join(targetRoot, 'artifact.pdf')) : null;
  const contactSheet = join(targetRoot, 'contact-01.png');
  const pageCount = await createContact(pagesDir, contactSheet);
  if (pageCount !== pages.length || pageCount !== rendered.result.pageCount) throw new Error('DSH DOCX render page inventory mismatch.');
  return { engine: rendered.result.engine, pageCount, pages, pdf, contactSheet };
}

const probe = JSON.parse(await readFile(probePath, 'utf8'));
const probeByKey = new Map(probe.map((entry) => [`${entry.format}/${entry.file}`, entry]));
const runningCases = cases.filter((item) => (batch === 'all' || item.format === batch) && (!idFilter || item.id === idFilter));
const outputRoot = join(runRoot, 'cases');
await mkdir(runRoot, { recursive: true });
await mkdir(outputRoot, { recursive: true });
const store = await LocalArtifactFiles.create(join(runRoot, 'store'), [workspace, outputRoot]);
const profile = new DocxProfile({ files: store, ...runtime });
const parsedDocx = new Map();
const baselineCache = new Map();
const summary = { runId: 'office-aesthetic-50', createdAt: new Date().toISOString(), runner: 'DSH Office Profile local modules (direct Profile calls)', runnerVersion: 'source-built workspace at the current Git revision', totalPlanned: 50,
  mode: 'first-pass with evidence-led repair enabled (0..5); only repairs supported by the selected DSH module are permitted', naturalLanguageNote: 'Primary prompts are frozen verbatim but manually mapped to typed DSH module operations; this Profile does not include a natural-language planning model. Findings therefore measure module/backend effects and capability limits, not autonomous prompt interpretation.', excludedFormat: { format: 'PDF', status: 'not_run', reason: 'pdf-office is a read-only input module; this run targets visual editing outcomes.' }, cases: [] };

function docxTargets(ir) {
  const top = ir.semantic.blocks;
  const topParas = top.filter((block) => block.kind === 'paragraph' && block.text.trim());
  const nestedParas = top.flatMap((block) => block.kind === 'table' ? block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs)) : [block]).filter((p) => p.kind === 'paragraph' && p.text.trim());
  const title = nestedParas.find((p) => /title/i.test(p.styleName ?? '')) ?? topParas[0] ?? nestedParas[0];
  let headings = topParas.filter((p) => p.id !== title?.id && /heading/i.test(p.styleName ?? ''));
  if (!headings.length) headings = topParas.filter((p) => p.id !== title?.id && /^第[一二三四五六七八九十百零0-9]+[章节条]/.test(p.text.trim()));
  const bodies = topParas.filter((p) => p.id !== title?.id && !headings.some((h) => h.id === p.id) && p.text.trim().length >= 70);
  const tables = top.filter((block) => block.kind === 'table');
  const labels = nestedParas.filter((p) => p.id !== title?.id && p.text.trim().length >= 3 && p.text.trim().length <= 60 && /[:：\?？]$/.test(p.text.trim()));
  return { title, headings, bodies, tables, labels, allParagraphs: nestedParas };
}
function pEdit(ir, p, format) { return { kind: 'formatParagraph', target: targetFromDualIR(ir, p.id), ...format }; }
function tableEdit(ir, table, index) { return { kind: 'formatTable', target: tableTargetFromDualIR(ir, table.id), layout: 'autofit', alignment: 'Center',
  cellVerticalAlignment: 'center', cellMargins: { top: 3, left: 4, bottom: 3, right: 4 },
  borders: { top: { style: 'single', size: 0.5, color: '#C8D2DC' }, bottom: { style: 'single', size: 0.5, color: '#C8D2DC' },
    left: { style: 'single', size: 0.5, color: '#D6DEE5' }, right: { style: 'single', size: 0.5, color: '#D6DEE5' },
    insideH: { style: 'single', size: 0.35, color: '#D6DEE5' }, insideV: { style: 'single', size: 0.35, color: '#D6DEE5' } },
  ...(index === 0 && table.rows?.[0]?.cells?.length > 1 && table.rows[0].cells.every((cell) => cell.text.trim() && cell.text.trim().length <= 42) ? { headerRow: true } : {}) };
}
function docxPlan(testCase, ir) {
  const targets = docxTargets(ir); const color = ['#244A67', '#176B67', '#38506A', '#23415E'][testCase.variant]; const edits = [];
  const titleFormat = { alignment: 'Centered', spaceAfter: 12, font: { bold: true, size: 21, color, name: 'Arial' } };
  const headingFormat = { spaceBefore: 10, spaceAfter: 5, outlineLevel: 2, font: { bold: true, size: 14, color, name: 'Arial' } };
  if (testCase.variant === 0 || testCase.variant === 3) {
    if (targets.title) edits.push(pEdit(ir, targets.title, titleFormat));
    for (const heading of targets.headings.slice(0, 4)) edits.push(pEdit(ir, heading, headingFormat));
  }
  if (testCase.variant === 1 || testCase.variant === 3) {
    if (targets.tables.length) for (const [index, table] of targets.tables.entries()) edits.push(tableEdit(ir, table, index));
    else for (const label of targets.labels.slice(0, 5)) edits.push(pEdit(ir, label, { spaceBefore: 6, spaceAfter: 3, font: { bold: true, size: 11, color, name: 'Arial' } }));
  }
  if (testCase.variant === 2) {
    for (const body of targets.bodies.slice(0, 4)) edits.push(pEdit(ir, body, { lineSpacing: 15, spaceAfter: 6, font: { size: 10.5, name: 'Arial' } }));
    for (const heading of targets.headings.slice(0, 3)) edits.push(pEdit(ir, heading, headingFormat));
    if (!targets.bodies.length && targets.title) edits.push(pEdit(ir, targets.title, titleFormat));
  }
  if (!edits.length) throw new Error(`No safe target found for ${testCase.id}.`);
  return { edits };
}
function findPptxChange(probeEntry, slideNumber, text, replaceWith) {
  const matches = [];
  for (const slide of probeEntry.result.slides) if (slide.slideNumber === slideNumber) for (const shape of slide.shapes) for (const run of shape.paragraphs ?? []) {
    if (run.text === text) matches.push({ slideNumber, shapeId: shape.shapeId, paragraphIndex: run.paragraphIndex, runIndex: run.runIndex, expectedText: text, replaceWith });
  }
  if (matches.length !== 1) throw new Error(`PPTX target '${text}' expected once on slide ${slideNumber}; found ${matches.length}.`);
  return matches[0];
}
function planPptx(testCase, probeEntry) {
  if (testCase.variant === 2) return { action: 'formatText', changes: [{ scope: 'allSlides', titleFontSize: 30, bodyFontSize: 18, accentColor: '#244A67' }] };
  return { action: 'replaceText', changes: pptxReplacements[testCase.file][testCase.variant].map((change) =>
    findPptxChange(probeEntry, change.slideNumber, change.text, change.replaceWith)) };
}
function planXlsx(testCase) {
  if (testCase.variant === 2) return { action: 'formatCells', changes: [{ sheet: '*', range: 'A1:Z1', font: { bold: true, color: '#FFFFFF' }, fill: '#244A67', border: 'thin' }] };
  const sheets = xlsxSheets[testCase.file][testCase.variant];
  return { action: 'setPrintLayout', sheets: sheets.map((sheet) => ({ sheet, orientation: 'landscape', fitToWidth: testCase.variant === 0 ? 1 : 2, fitToHeight: 0 })) };
}
function yamlContract(testCase, sourceHash) {
  return JSON.stringify({ case_id: testCase.id, format: testCase.format, mode: 'first-pass-and-repair', source_artifact: `${testCase.format}/${testCase.file}`,
    source_sha256: sourceHash, initial_content: `${titles[testCase.file]} from the frozen official-source office corpus.`, primary_prompt: testCase.prompt,
    expected_effects: testCase.expectedEffects, invariants: testCase.invariants,
    acceptance: { score_minimum: 85, required_gates: ['file_integrity', 'complete_render', 'content_preservation', 'no_critical_visual_defect', 'intent_complete'], repair_limit: 5 } }, null, 2) + '\n';
}

await mkdir(runRoot, { recursive: true });
await writeFile(join(runRoot, 'cases.json'), JSON.stringify(cases, null, 2) + '\n', 'utf8');
try {
  for (const testCase of runningCases) {
    const sourcePath = join(workspace, testCase.format, testCase.file);
    const sourceBytes = await readFile(sourcePath); const sourceHash = sha256(sourceBytes);
    const probeEntry = probeByKey.get(`${testCase.format}/${testCase.file}`);
    if (!probeEntry || probeEntry.sourceSha256 !== sourceHash) throw new Error(`Frozen-source hash mismatch for ${testCase.id}.`);
    const sourceRef = await store.importFile(sourcePath);
    const caseRoot = join(outputRoot, testCase.id);
    const roundRoot = join(caseRoot, 'dsh', 'round-00');
    await mkdir(roundRoot, { recursive: true });
    await writeFile(join(caseRoot, 'case.yaml'), yamlContract(testCase, sourceHash), 'utf8');
    const baselineKey = `${testCase.format}/${testCase.file}`;
    let baseline = baselineCache.get(baselineKey);
    if (!baseline) {
      const baselineRoot = join(runRoot, 'shared-baselines', testCase.format, basename(testCase.file, extname(testCase.file)));
      await mkdir(baselineRoot, { recursive: true });
      const artifactPath = join(baselineRoot, `source${extname(testCase.file)}`);
      await writeFile(artifactPath, sourceBytes);
      if (testCase.format === 'docx') {
        baseline = { sourceArtifact: { path: artifactPath, sha256: sourceHash }, render: await renderDocx(profile, sourceRef, store, join(baselineRoot, 'render'), `aesthetic50-baseline-${basename(testCase.file)}`) };
      } else {
        const rendered = await renderWithLibreOffice(artifactPath, join(baselineRoot, 'render'));
        baseline = { sourceArtifact: { path: artifactPath, sha256: sourceHash }, render: await materializeExternalRender(rendered, join(baselineRoot, 'rendered')) };
      }
      baselineCache.set(baselineKey, baseline);
    }
    let operationPlan, dshCall, candidateRef = sourceRef, verification = null, operationOk = false;
    try {
      if (testCase.format === 'docx') {
        let ir = parsedDocx.get(testCase.file);
        if (!ir) {
          const parsed = await profile.call('docx-parse', 'execute', { artifactRef: sourceRef, requestId: `aesthetic50-${testCase.id}-parse` });
          ir = parsed.result.ir.content; parsedDocx.set(testCase.file, ir);
        }
        operationPlan = docxPlan(testCase, ir);
        const result = await profile.call('docx-edit', 'execute', { artifactRef: sourceRef, requestId: `aesthetic50-${testCase.id}-execute`, plan: operationPlan });
        candidateRef = result.result.artifact;
        verification = result.result.verification;
        dshCall = { moduleId: 'docx-edit', operation: 'execute', requestId: result.requestId, edits: result.result.edits, verification };
        operationOk = true;
      } else if (testCase.format === 'pptx') {
        operationPlan = planPptx(testCase, probeEntry);
        const result = await profile.call('pptx-office', 'execute', { artifactRef: sourceRef, requestId: `aesthetic50-${testCase.id}-execute`, payload: operationPlan });
        candidateRef = result.result.artifactRef;
        verification = await profile.call('pptx-office', 'verify', { artifactRef: candidateRef, requestId: `aesthetic50-${testCase.id}-verify`, payload: {
          expectedSlideCount: probeEntry.result.slideCount,
          textIncludes: operationPlan.changes.map((change) => change.replaceWith).filter(Boolean),
        } });
        dshCall = { moduleId: 'pptx-office', operation: 'execute', requestId: result.requestId, result: result.result, verification };
        operationOk = true;
      } else {
        operationPlan = planXlsx(testCase);
        const result = await profile.call('xlsx-office', 'execute', { artifactRef: sourceRef, requestId: `aesthetic50-${testCase.id}-execute`, payload: operationPlan });
        candidateRef = result.result.artifactRef;
        verification = await profile.call('xlsx-office', 'verify', { artifactRef: candidateRef, requestId: `aesthetic50-${testCase.id}-verify`, payload: {
          expectedWorksheetCount: probeEntry.result.worksheetCount, sheetNamesInclude: probeEntry.result.worksheets.map((sheet) => sheet.name),
        } });
        dshCall = { moduleId: 'xlsx-office', operation: 'execute', requestId: result.requestId, result: result.result, verification };
        operationOk = true;
      }
    } catch (error) {
      dshCall = { moduleId: testCase.format === 'docx' ? 'docx-edit' : `${testCase.format}-office`, operation: 'execute',
        requestId: `aesthetic50-${testCase.id}-execute`, error: { code: error.code ?? error.name, message: String(error.message ?? error) }, expectedCapabilityGap: testCase.variant === 2 };
      candidateRef = sourceRef;
    }
    const candidateFile = await copyArtifact(store, candidateRef, join(roundRoot, `artifact${extname(testCase.file)}`));
    let render;
    if (testCase.format === 'docx') render = await renderDocx(profile, candidateRef, store, join(roundRoot, 'render'), `aesthetic50-${testCase.id}-render`);
    else if (operationOk) {
      const candidatePath = candidateFile.path;
      const rendered = await renderWithLibreOffice(candidatePath, join(roundRoot, 'render'));
      render = await materializeExternalRender(rendered, join(roundRoot, 'rendered'));
    } else {
      const baselinePageDir = join(runRoot, 'shared-baselines', testCase.format, basename(testCase.file, extname(testCase.file)), 'rendered', 'pages');
      const casePageDir = join(roundRoot, 'rendered', 'pages');
      await mkdir(casePageDir, { recursive: true });
      const pageNames = (await readdir(baselinePageDir)).filter((name) => /^page-\d+\.png$/.test(name));
      const pages = [];
      for (const name of pageNames) { const path = join(casePageDir, name); await copyFile(join(baselinePageDir, name), path); pages.push({ pageNumber: Number(name.match(/\d+/)?.[0]), path, sha256: await fileSha(path) }); }
      const contactSheet = join(roundRoot, 'rendered', 'contact-01.png');
      const pageCount = await createContact(casePageDir, contactSheet);
      render = { pageCount, pages, contactSheet, identicalToBaselineRender: true };
    }
    const sourceUnchanged = await fileSha(sourcePath) === sourceHash;
    const manifest = { caseId: testCase.id, format: testCase.format, file: testCase.file, primaryPrompt: testCase.prompt,
      operationFamily: testCase.operationFamily, operationPlan, dshCall, source: { path: sourcePath, sha256: sourceHash, unchanged: sourceUnchanged },
      baseline: baseline.render, candidate: { ...candidateFile, pageCount: render.pageCount }, renderProfile: { engine: testCase.format === 'docx' ? 'DSH docx-render / bundled LibreOffice' : 'bundled LibreOffice + Poppler review-only', dpi: 120, sameRendererAsBaseline: true },
      round: 0, status: operationOk ? 'first-pass-rendered-pending-visual-score' : 'first-pass-capability-gap', verifier: verification,
      visualReview: { status: 'pending', candidatePages: render.pages.map((p) => p.path), contactSheet: render.contactSheet }, repairsUsed: 0, repairLimit: 5 };
    await atomicJson(join(roundRoot, 'manifest.json'), manifest);
    await atomicJson(join(caseRoot, 'manifest.json'), manifest);
    const brief = `# ${testCase.id} — DSH Profile\n\nVisual inspection and scores are pending. See dsh/round-00/manifest.json for the frozen contract, prompt, operations, hashes, verifier output, and every page image.\n`;
    await writeFile(join(caseRoot, 'report.md'), brief, 'utf8');
    summary.cases.push({ id: testCase.id, format: testCase.format, file: testCase.file, operationOk, sourceUnchanged, pageCount: render.pageCount, candidateSha256: candidateFile.sha256, contactSheet: render.contactSheet, capabilityGap: !operationOk });
    await atomicJson(join(runRoot, `run-${batch}.json`), { ...summary, finishedAt: new Date().toISOString() });
    console.log(`${testCase.id} ${testCase.format} ${testCase.file}: ${operationOk ? 'DSH action pass' : `capability gap ${dshCall.error?.code}`} ; ${render.pageCount} pages`);
  }
} finally { await profile.dispose(); }
