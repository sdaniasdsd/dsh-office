import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { copyFile, mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { createHash, randomUUID } from 'node:crypto';
import { basename, join, resolve } from 'node:path';
import { strToU8, unzipSync, zipSync } from 'fflate';

const root = resolve(process.argv[2]);
const rounds = Number(process.argv[3] ?? 10);
if (!Number.isInteger(rounds) || rounds < 1 || rounds > 10) throw new Error('rounds must be 1..10');
const sourceRoot = join(root, 'sources');
const workspace = join(root, 'workspace');
const data = join(root, 'data');
const reportPath = join(root, 'report.json');
const manifestPath = join(root, 'manifest.json');
const pkg = process.env.DSH_DOCX_PROFILE_ROOT ?? 'C:/Users/AA/AppData/Roaming/com.yeagoo.dsh-desktop/harness/profiles/web/node_modules/@deepseek-ai/dsh-docx';
const runtime = process.env.DSH_RUNTIME_ROOT ?? 'C:/Users/AA/AppData/Local/DSH Desktop/runtime';
const profileLabel = process.env.DSH_PROFILE_LABEL ?? `@deepseek-ai/dsh-docx at ${pkg}`;
const bundleRuntime = `${pkg}/runtime/win32-x64`;

const specimens = [
  { format: 'docx', file: 'pipl-law.docx', source: 'official-site-derived', url: 'https://www.cac.gov.cn/2021-08/20/c_1631050028355286.htm', category: ['法律条文'], note: 'Official CAC HTML rendered into a minimal DOCX fixture; source text provenance retained.' },
  { format: 'docx', file: 'flexible-working-form.docx', source: 'official-file', url: 'https://www.gov.uk/government/publications/the-right-to-request-flexible-working-form', category: ['标准表格'], note: 'Official flexible working request form.' },
  { format: 'docx', file: 'mhra-applicant-response.docx', source: 'official-file', url: 'https://www.gov.uk/government/publications/response-template-for-applicants', category: ['公文', '标准表格'], note: 'Official applicant response template for regulatory correspondence.' },
  { format: 'docx', file: 'asbestos-management-plan.docx', source: 'official-file', url: 'https://www.hse.gov.uk/asbestos/assets/docs/blank-management-plan.docx', category: ['公文', '标准表格'], note: 'Official workplace asbestos management plan.' },
  { format: 'docx', file: 'qa-evidence-report-template.docx', source: 'official-file', url: 'https://www.gov.uk/government/publications/energy-security-and-net-zero-modelling-quality-assurance-qa-tools-and-guidance', category: ['测试报告'], note: 'Official analytical quality-assurance report template.' },
  { format: 'pptx', file: 'civil-service-line-management.pptx', source: 'official-file', url: 'https://www.gov.uk/government/publications/civil-service-line-management-standards', category: ['公文', '标准'], note: 'Official Civil Service line management standard deck.' },
  { format: 'pptx', file: 'timms-workshop.pptx', source: 'official-file', url: 'https://www.gov.uk/government/publications/timms-review-of-personal-independence-payment-run-a-workshop', category: ['公文', '会议材料'], note: 'Official workshop slide deck.' },
  { format: 'pptx', file: 'civil-society-covenant.pptx', source: 'official-file', url: 'https://www.gov.uk/government/publications/local-civil-society-covenant-resources', category: ['公文', '政策'], note: 'Official government/civil-society policy deck.' },
  { format: 'pptx', file: 'qualifications-reform.pptx', source: 'official-file', url: 'https://www.gov.uk/government/publications/qualifications-reform-resources-for-teachers', category: ['法律与政策'], note: 'Official qualifications policy deck.' },
  { format: 'pptx', file: 'prevent-duty-leadership.pptx', source: 'official-file', url: 'https://www.gov.uk/government/publications/the-leadership-challenge-the-prevent-duty-for-governing-bodies-and-senior-leaders-in-higher-education-he', category: ['法律条文', '合规培训'], note: 'Official Prevent Duty compliance training deck.' },
  { format: 'xlsx', file: 'qa-modelling-template.xlsx', source: 'official-file', url: 'https://www.gov.uk/government/publications/energy-security-and-net-zero-modelling-quality-assurance-qa-tools-and-guidance', category: ['测试报告', '标准表格'], note: 'Official model quality assurance workbook.' },
  { format: 'xlsx', file: 'qa-assumptions-log.xlsx', source: 'official-file', url: 'https://www.gov.uk/government/publications/energy-security-and-net-zero-modelling-quality-assurance-qa-tools-and-guidance', category: ['测试报告', '标准表格'], note: 'Official assumptions-log workbook.' },
  { format: 'xlsx', file: 'green-book-appraisal-tables.xlsx', source: 'official-file', url: 'https://www.gov.uk/government/publications/the-green-book-templates-and-support-material', category: ['公文', '标准表格'], note: 'Official business-case appraisal table template.' },
  { format: 'xlsx', file: 'green-book-discount-factors.xlsx', source: 'official-file', url: 'https://www.gov.uk/government/publications/the-green-book-templates-and-support-material', category: ['标准表格'], note: 'Official Green Book discount factor workbook.' },
  { format: 'xlsx', file: 'condition-survey-template.xlsx', source: 'official-file', url: 'https://www.gov.uk/government/publications/commissioning-a-condition-survey-for-school-and-college-buildings', category: ['测试报告', '标准表格'], note: 'Official condition survey workbook.' },
];

function decodeHtml(text) {
  return text.replace(/&nbsp;|&#160;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"').replace(/&#39;|&apos;/gi, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)));
}
function xml(text) { return text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;'); }
async function makeLawFixture() {
  const html = await readFile(join(sourceRoot, 'legal', 'pipl.html'), 'utf8');
  const title = '中华人民共和国个人信息保护法';
  const cleaned = html.replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>/gi, '')
    .replace(/<br\s*\/?\s*>|<\/(?:p|div|section|h[1-6]|li|tr|td|th|article)>/gi, '\n')
    .replace(/<[^>]+>/g, ' ');
  const paragraphs = decodeHtml(cleaned).split(/[\r\n\t]+/).map(s => s.replace(/\s+/g, ' ').trim())
    .filter(s => s.length >= 2 && !/^(首页|网站地图|联系我们|关闭|打印|分享|收藏)$/.test(s));
  if (!paragraphs.some(p => p.includes('个人信息保护法')) || paragraphs.length < 25) throw new Error('Official law HTML extraction did not produce a credible body.');
  const body = [title, ...paragraphs].map(text => `<w:p><w:r><w:t xml:space="preserve">${xml(text)}</w:t></w:r></w:p>`).join('');
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`;
  const contentTypes = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';
  const rels = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
  await writeFile(join(sourceRoot, 'docx', 'pipl-law.docx'), zipSync({
    '[Content_Types].xml': strToU8(contentTypes), '_rels/.rels': strToU8(rels), 'word/document.xml': strToU8(document),
  }, { level: 6 }));
}
async function sanitizeExternalRelationships(path) {
  const archive = unzipSync(new Uint8Array(await readFile(path)));
  const externalIds = new Set();
  let removedRelationships = 0;
  for (const [name, bytes] of Object.entries(archive)) {
    if (!name.endsWith('.rels')) continue;
    const text = new TextDecoder().decode(bytes);
    const cleaned = text.replace(/<Relationship\b[^>]*\/>|<Relationship\b[^>]*>[\s\S]*?<\/Relationship>/g, tag => {
      if (!/\bTargetMode\s*=\s*["']External["']/i.test(tag)) return tag;
      const id = /\bId\s*=\s*["']([^"']+)["']/i.exec(tag)?.[1];
      if (id) externalIds.add(id);
      removedRelationships++;
      return '';
    });
    archive[name] = strToU8(cleaned);
  }
  if (externalIds.size) {
    for (const [name, bytes] of Object.entries(archive)) {
      if (!name.endsWith('.xml')) continue;
      let text = new TextDecoder().decode(bytes);
      for (const id of externalIds) {
        const escaped = id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        text = text.replace(new RegExp(`<w:hyperlink\\b(?=[^>]*\\br:id=["']${escaped}["'])[^>]*>([\\s\\S]*?)<\\/w:hyperlink>`, 'g'), '$1');
      }
      archive[name] = strToU8(text);
    }
  }
  if (removedRelationships) await writeFile(path, zipSync(archive, { level: 6 }));
  return { removedRelationships, externalRelationshipIds: externalIds.size };
}
function parseResult(response) {
  if (response.isError) throw new Error(response.content?.map(item => item.text ?? '').join('\n') || 'MCP tool returned an error.');
  const item = response.content?.find(block => block.type === 'text');
  if (!item) throw new Error('DSH tool returned no text result.');
  return JSON.parse(item.text);
}
function summary(result) {
  const r = result?.result ?? result;
  if (!r || typeof r !== 'object') return { type: typeof r };
  const s = { keys: Object.keys(r).slice(0, 20) };
  for (const key of ['format', 'pageCount', 'slideCount', 'sheets', 'worksheets', 'paragraphCount', 'tableCount', 'headingCount', 'warnings', 'engine']) {
    if (r[key] !== undefined) s[key] = key === 'warnings' ? r[key].length : r[key];
  }
  if (result.resultRef) s.resultRef = result.resultRef;
  if (result.pdf) s.pdf = result.pdf;
  if (result.pages) s.pages = result.pages.map(p => ({ pageNumber: p.pageNumber, image: p.image, thumbnail: p.thumbnail }));
  if (result.artifacts) s.artifacts = result.artifacts;
  return s;
}
function findSheet(inspected) {
  const r = inspected?.result ?? {};
  const candidates = r.worksheets ?? r.sheets ?? r.worksheetSummaries ?? [];
  const first = Array.isArray(candidates) ? candidates[0] : undefined;
  if (typeof first === 'string') return first;
  if (first && typeof first.name === 'string') return first.name;
  if (first && typeof first.sheetName === 'string') return first.sheetName;
  return typeof r.firstSheet === 'string' ? r.firstSheet : undefined;
}
async function atomicJson(path, value) {
  const temp = `${path}.${process.pid}.tmp`;
  await writeFile(temp, JSON.stringify(value, null, 2), 'utf8');
  await rename(temp, path);
}

await mkdir(sourceRoot, { recursive: true });
await mkdir(workspace, { recursive: true });
await mkdir(data, { recursive: true });
await makeLawFixture();
for (const specimen of specimens) {
  const source = join(sourceRoot, specimen.format, specimen.file);
  const dest = join(workspace, specimen.format, specimen.file);
  await mkdir(join(workspace, specimen.format), { recursive: true });
  await copyFile(source, dest);
  if (specimen.format === 'docx') specimen.sanitization = await sanitizeExternalRelationships(dest);
  specimen.workspacePath = `${specimen.format}/${specimen.file}`;
  specimen.sourceSha256 = createHash('sha256').update(await readFile(source)).digest('hex');
  specimen.workspaceSha256 = createHash('sha256').update(await readFile(dest)).digest('hex');
  specimen.sizeBytes = (await readFile(dest)).byteLength;
}
await atomicJson(manifestPath, { createdAt: new Date().toISOString(), profile: profileLabel, profileRoot: pkg, formats: ['docx', 'pdf', 'pptx', 'xlsx'], perFormat: 5, rounds, totalMaterialRounds: 20 * rounds, specimens,
  pdfSamples: specimens.filter(s => s.format === 'docx').map(s => ({ sourceDocx: s.file, kind: 'DSH docx-render PDF output', derived: true })) });

const env = Object.fromEntries(Object.entries(process.env).filter(([, value]) => typeof value === 'string'));
Object.assign(env, {
  THE_LAST_DOCX_WORKSPACE: workspace,
  THE_LAST_DOCX_DATA: data,
  DOCX_PYTHON: `${bundleRuntime}/python/python.exe`,
  DOCX_SOFFICE: `${bundleRuntime}/libreoffice/program/soffice.com`,
  DOCX_PDFTOPPM: `${bundleRuntime}/poppler/poppler-26.09.0/Library/bin/pdftoppm.exe`,
});
const client = new Client({ name: 'dsh-office-stress', version: '1.0.0' });
const transport = new StdioClientTransport({ command: `${runtime}/node.exe`, args: [`${pkg}/lib/server.mjs`], cwd: workspace, env, stderr: 'pipe' });
const report = { startedAt: new Date().toISOString(), finishedAt: null, profile: profileLabel, backend: 'DSH profile MCP stdio server', roundsPerMaterial: rounds, plannedMaterialRounds: 20 * rounds, actualCalls: 0, imports: [], preflights: [], records: [], failures: [], pdfCoverageNote: 'This run exercises DOCX-to-PDF output rendering; native PDF input parsing is covered by the separate pdf-office replay. PDF stress artifacts are emitted by the DSH render module from the five official DOCX office materials.', visualReview: 'pending; page PNG artifacts are recorded for follow-up inspection.' };
const serverStderr = [];
transport.stderr?.on('data', chunk => serverStderr.push(Buffer.isBuffer(chunk) ? chunk.toString('utf8') : String(chunk)));
const refs = new Map();
let done = 0;
async function callTool(name, args) {
  const startedAt = new Date().toISOString();
  const start = performance.now();
  report.actualCalls++;
  try {
    const parsed = parseResult(await client.callTool({ name, arguments: args }));
    return { ok: true, startedAt, durationMs: Math.round(performance.now() - start), parsed };
  } catch (error) {
    return { ok: false, startedAt, durationMs: Math.round(performance.now() - start), error: String(error?.message ?? error) };
  }
}
function noteFailure(op, specimen, round, output, requestId) {
  if (!output.ok) report.failures.push({ format: specimen.format, file: specimen.file, round, operation: op,
    requestId, startedAt: output.startedAt, durationMs: output.durationMs, error: output.error });
}
try {
  await client.connect(transport);
  const tools = await client.listTools();
  report.toolNames = tools.tools.map(tool => tool.name);
  const doctor = await callTool('docx_doctor', {});
  report.doctor = doctor.ok ? doctor.parsed : { error: doctor.error };
  if (!doctor.ok) report.failures.push({ operation: 'docx_doctor', error: doctor.error });

  for (const specimen of specimens) {
    const imported = await callTool('docx_import', { path: specimen.workspacePath });
    if (!imported.ok) {
      report.failures.push({ operation: 'docx_import', format: specimen.format, file: specimen.file, error: imported.error });
      continue;
    }
    const ref = imported.parsed.artifactRef ?? imported.parsed.result?.artifactRef ?? imported.parsed;
    refs.set(specimen.file, ref);
    report.imports.push({ format: specimen.format, file: specimen.file, sourceSha256: specimen.sourceSha256,
      workspaceSha256: specimen.workspaceSha256, artifactRef: ref });
    await atomicJson(reportPath, report);
  }

  // One contract/structure preflight per source, called through DSH profile tools.
  for (const specimen of specimens) {
    const artifactRef = refs.get(specimen.file);
    if (!artifactRef) continue;
    let name, args;
    if (specimen.format === 'docx') { name = 'docx_call'; args = { moduleId: 'docx-inspect', operation: 'inspect', input: { artifactRef, requestId: `stress-preflight-${randomUUID()}` } }; }
    else if (specimen.format === 'pptx') { name = 'pptx_call'; args = { operation: 'inspect', input: { artifactRef, operation: 'inspect', requestId: `stress-preflight-${randomUUID()}` } }; }
    else { name = 'xlsx_call'; args = { operation: 'inspect', input: { artifactRef, operation: 'inspect', requestId: `stress-preflight-${randomUUID()}` } }; }
    const output = await callTool(name, args);
    report.preflights.push({ format: specimen.format, file: specimen.file, ok: output.ok, durationMs: output.durationMs,
      summary: output.ok ? summary(output.parsed) : undefined, rawResult: output.ok ? output.parsed : undefined, error: output.error });
    noteFailure('inspect-preflight', specimen, 0, output);
    if (specimen.format === 'xlsx' && output.ok) specimen.sheetName = findSheet(output.parsed);
    await atomicJson(reportPath, report);
  }

  // Each of the 20 material cases receives the requested number of complete DSH calls.
  for (const specimen of specimens) {
    const artifactRef = refs.get(specimen.file);
    if (!artifactRef) continue;
    for (let round = 1; round <= rounds; round++) {
      const requestId = `stress-${specimen.format}-${basename(specimen.file)}-${String(round).padStart(2, '0')}-${randomUUID()}`;
      let name, args, operation;
      if (specimen.format === 'docx') {
        operation = 'docx-parse.execute'; name = 'docx_call';
        args = { moduleId: 'docx-parse', operation: 'execute', input: { artifactRef, requestId } };
      } else if (specimen.format === 'pptx') {
        operation = 'pptx.extract'; name = 'pptx_call';
        args = { operation: 'execute', input: { artifactRef, operation: 'execute', requestId, payload: { action: 'extract' } } };
      } else {
        operation = 'xlsx.readRange'; name = 'xlsx_call';
        const sheet = specimen.sheetName;
        if (!sheet) {
          report.failures.push({ format: 'xlsx', file: specimen.file, round, operation, error: 'No worksheet name returned by DSH inspect; readRange cannot safely guess.' });
          continue;
        }
        args = { operation: 'execute', input: { artifactRef, operation: 'execute', requestId, payload: { action: 'readRange', sheet, range: 'A1:J20' } } };
      }
      const output = await callTool(name, args);
      const record = { format: specimen.format, file: specimen.file, round, operation, requestId, ok: output.ok, startedAt: output.startedAt, durationMs: output.durationMs,
        summary: output.ok ? summary(output.parsed) : undefined, resultRef: output.ok ? output.parsed.resultRef : undefined, error: output.error };
      report.records.push(record); noteFailure(operation, specimen, round, output, requestId);
      if (specimen.format === 'docx' && output.ok) record.parseSummary = summary(output.parsed);
      done++;
      if (done % 5 === 0 || done === report.plannedMaterialRounds) {
        await atomicJson(reportPath, report);
        process.stdout.write(`completed ${done}/${report.plannedMaterialRounds} material rounds; calls=${report.actualCalls}; failures=${report.failures.length}\n`);
      }
    }
  }

  // Five render-material cases are the actual PDF artifacts emitted by docx-render.
  const docxSamples = specimens.filter(s => s.format === 'docx');
  for (const specimen of docxSamples) {
    const artifactRef = refs.get(specimen.file);
    if (!artifactRef) continue;
    for (let round = 1; round <= rounds; round++) {
      const requestId = `stress-pdf-${basename(specimen.file)}-${String(round).padStart(2, '0')}-${randomUUID()}`;
      const output = await callTool('docx_call', { moduleId: 'docx-render', operation: 'execute', input: { artifactRef, requestId } });
      const record = { format: 'pdf', sourceDocx: specimen.file, round, operation: 'docx-render.execute', requestId, ok: output.ok, startedAt: output.startedAt, durationMs: output.durationMs,
        summary: output.ok ? summary(output.parsed) : undefined, renderResult: output.ok ? output.parsed.result : undefined, error: output.error };
      report.records.push(record); noteFailure('docx-render.execute', { ...specimen, format: 'pdf' }, round, output, requestId);
      done++;
      if (done % 5 === 0 || done === report.plannedMaterialRounds) {
        await atomicJson(reportPath, report);
        process.stdout.write(`completed ${done}/${report.plannedMaterialRounds} total rounds including PDF renders; calls=${report.actualCalls}; failures=${report.failures.length}\n`);
      }
    }
  }
} finally {
  report.finishedAt = new Date().toISOString();
  report.serverStderr = serverStderr.join('');
  await atomicJson(reportPath, report);
  await client.close().catch(() => {});
}

process.stdout.write(JSON.stringify({ reportPath, manifestPath, roundsPerMaterial: rounds, records: report.records.length, calls: report.actualCalls, failures: report.failures.length, durationSec: Math.round((Date.parse(report.finishedAt) - Date.parse(report.startedAt)) / 1000) }, null, 2));
