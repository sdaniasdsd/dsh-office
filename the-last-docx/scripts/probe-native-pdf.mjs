import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { access, copyFile, mkdir, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';

const root = resolve(process.argv[2]);
if (!process.argv[2]) throw new Error('Usage: node scripts/probe-native-pdf.mjs <fresh-run-root> [output-json]');
const outputPath = resolve(process.argv[3] ?? join(root, `native-pdf-boundary-${Date.now()}.json`));
const workspace = `${root}/workspace`;
const source = `${root}/sources/pdf/nist-test-report.pdf`;
const relative = 'pdf/nist-test-report.pdf';
await mkdir(`${workspace}/pdf`, { recursive: true });
await access(`${workspace}/${relative}`).then(() => { throw new Error(`Refusing to overwrite an existing workspace fixture: ${workspace}/${relative}`); }, error => {
  if (error.code !== 'ENOENT') throw error;
});
await copyFile(source, `${workspace}/${relative}`);
const pkg = 'C:/Users/AA/AppData/Roaming/com.yeagoo.dsh-desktop/harness/profiles/web/node_modules/@deepseek-ai/dsh-docx';
const runtime = 'C:/Users/AA/AppData/Local/DSH Desktop/runtime';
const bundleRuntime = `${pkg}/runtime/win32-x64`;
const env = Object.fromEntries(Object.entries(process.env).filter(([, value]) => typeof value === 'string'));
Object.assign(env, {
  THE_LAST_DOCX_WORKSPACE: workspace, THE_LAST_DOCX_DATA: `${root}/data`,
  DOCX_PYTHON: `${bundleRuntime}/python/python.exe`, DOCX_SOFFICE: `${bundleRuntime}/libreoffice/program/soffice.com`,
  DOCX_PDFTOPPM: `${bundleRuntime}/poppler/poppler-26.09.0/Library/bin/pdftoppm.exe`,
});
const client = new Client({ name: 'dsh-native-pdf-boundary', version: '1.0.0' });
const transport = new StdioClientTransport({ command: `${runtime}/node.exe`, args: [`${pkg}/lib/server.mjs`], cwd: workspace, env, stderr: 'pipe' });
let data;
try {
  await client.connect(transport);
  const imported = await client.callTool({ name: 'docx_import', arguments: { path: relative } });
  const importedValue = JSON.parse(imported.content.find(c => c.type === 'text').text);
  const artifactRef = importedValue.artifactRef ?? importedValue.result?.artifactRef ?? importedValue;
  const response = await client.callTool({ name: 'docx_call', arguments: {
    moduleId: 'docx-render', operation: 'inspect', input: { artifactRef, requestId: `native-pdf-boundary-${Date.now()}` },
  } });
  let result;
  try { result = JSON.parse(response.content?.find(c => c.type === 'text')?.text ?? '{}'); } catch { result = { parseError: true }; }
  data = { source: 'https://www.nist.gov/system/files/documents/el/isd/ms/DRAFTTest-Report-ForkTines.pdf',
    artifactRef, importSucceeded: imported.isError !== true, renderInspectAccepted: response.isError !== true,
    renderInspectError: response.isError ? result : undefined, result: response.isError ? undefined : result };
} finally { await client.close().catch(() => {}); }
await mkdir(dirname(outputPath), { recursive: true });
await writeFile(outputPath, JSON.stringify(data, null, 2), { encoding: 'utf8', flag: 'wx' });
process.stdout.write(JSON.stringify({ path: outputPath, importSucceeded: data.importSucceeded,
  renderInspectAccepted: data.renderInspectAccepted, errorCode: data.renderInspectError?.code,
  errorMessage: data.renderInspectError?.message }, null, 2));
