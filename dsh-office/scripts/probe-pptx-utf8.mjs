import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';

const root = resolve(process.argv[2]);
if (!process.argv[2]) throw new Error('Usage: node scripts/probe-pptx-utf8.mjs <run-root> [output-json]');
const outputPath = resolve(process.argv[3] ?? join(root, `utf8-contrast-${Date.now()}.json`));
const report = JSON.parse(await readFile(`${root}/report.json`, 'utf8'));
const failures = new Set(report.failures.filter(f => f.format === 'pptx').map(f => f.file));
const refs = new Map(report.imports.filter(i => i.format === 'pptx').map(i => [i.file, i.artifactRef]));
const pkg = 'C:/Users/AA/AppData/Roaming/com.yeagoo.dsh-desktop/harness/profiles/web/node_modules/@deepseek-ai/dsh-docx';
const runtime = 'C:/Users/AA/AppData/Local/DSH Desktop/runtime';
const bundleRuntime = `${pkg}/runtime/win32-x64`;
const env = Object.fromEntries(Object.entries(process.env).filter(([, value]) => typeof value === 'string'));
Object.assign(env, {
  THE_LAST_DOCX_WORKSPACE: `${root}/workspace`, THE_LAST_DOCX_DATA: `${root}/data`,
  DOCX_PYTHON: `${bundleRuntime}/python/python.exe`, DOCX_SOFFICE: `${bundleRuntime}/libreoffice/program/soffice.com`,
  DOCX_PDFTOPPM: `${bundleRuntime}/poppler/poppler-26.09.0/Library/bin/pdftoppm.exe`,
  PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1',
});
const client = new Client({ name: 'dsh-pptx-utf8-contrast', version: '1.0.0' });
const transport = new StdioClientTransport({ command: `${runtime}/node.exe`, args: [`${pkg}/lib/server.mjs`], cwd: `${root}/workspace`, env, stderr: 'pipe' });
const output = { env: { PYTHONIOENCODING: env.PYTHONIOENCODING, PYTHONUTF8: env.PYTHONUTF8 }, cases: [] };
try {
  await client.connect(transport);
  for (const file of failures) {
    const start = performance.now();
    const response = await client.callTool({ name: 'pptx_call', arguments: {
      operation: 'execute', input: { operation: 'execute', requestId: `utf8-contrast-${Date.now()}`, artifactRef: refs.get(file), payload: { action: 'extract' } },
    } });
    let result;
    try { result = JSON.parse(response.content?.find(c => c.type === 'text')?.text ?? '{}'); }
    catch { result = { parseError: true }; }
    const data = result?.result ?? {};
    const slides = Array.isArray(data.slides) ? data.slides : [];
    output.cases.push({
      file, ok: response.isError !== true, durationMs: Math.round(performance.now() - start),
      engine: data.engine, slideCount: data.slideCount ?? slides.length,
      shapeCount: slides.reduce((sum, slide) => sum + (Array.isArray(slide.shapes) ? slide.shapes.length : 0), 0),
      extractedTextChars: slides.reduce((sum, slide) => sum + (slide.shapes ?? []).flatMap(shape => shape.paragraphs ?? []).reduce((n, paragraph) => n + String(paragraph.text ?? '').length, 0), 0),
      error: response.isError ? result : undefined,
    });
  }
} finally {
  await client.close().catch(() => {});
}
await mkdir(dirname(outputPath), { recursive: true });
await writeFile(outputPath, JSON.stringify(output, null, 2), { encoding: 'utf8', flag: 'wx' });
process.stdout.write(JSON.stringify({ path: outputPath, cases: output.cases }, null, 2));
