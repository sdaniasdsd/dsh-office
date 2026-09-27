// Read back what is actually installed in the DSH web profile: version, the tools
// it registers, whether the DOCX formatting work is in the bundle, and the
// version it reports in the MCP handshake.
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { join } from 'node:path';

const PROFILE = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web');
const installed = join(PROFILE, 'node_modules', '@deepseek-ai', 'dsh-docx');
const manifest = JSON.parse(readFileSync(join(installed, 'package.json'), 'utf8'));
const server = readFileSync(join(installed, 'lib', 'server.mjs'), 'utf8');
const tools = [...server.matchAll(/registerTool\('([^']+)'/gu)].map((match) => match[1]);

console.log(`installed version  : ${manifest.version}`);
console.log(`main               : ${manifest.main}`);
console.log(`registered tools   : ${tools.length} → ${tools.join(', ')}`);
for (const marker of ['cellFormats', 'firstLineIndentPt', 'parseFormatting', 'firstLineChars', 'carryFormatting', '__DSH_DOCX_VERSION__']) {
  console.log(`  ${marker.padEnd(22)}: ${(server.match(new RegExp(marker, 'gu')) ?? []).length}`);
}
const modules = ['docx-inspect', 'docx-easy-parse', 'docx-parse', 'docx-complex-parse', 'docx-create', 'docx-styles', 'docx-edit', 'docx-render', 'docx-artifact', 'pptx-office', 'xlsx-office'];
console.log(`module ids named   : ${modules.filter((id) => server.includes(id)).join(', ')}`);
const engines = join(installed, 'lib', 'engines');
if (existsSync(engines)) console.log(`bundled engines    : ${readdirSync(engines).join(', ')}`);
for (const name of ['python', 'libreoffice', 'poppler']) {
  const path = join(installed, 'runtime', 'win32-x64', name);
  console.log(`runtime ${name.padEnd(11)}: ${existsSync(path) ? 'present' : 'MISSING'}`);
}
const pythonExe = join(installed, 'runtime', 'win32-x64', 'python', 'python.exe');
if (existsSync(pythonExe)) console.log(`python.exe         : ${statSync(pythonExe).size} bytes`);

const initialize = JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'verify-installed', version: '1' } } });
const handshake = spawnSync(process.execPath, [join(installed, 'lib', 'server.mjs')], { input: `${initialize}\n`, encoding: 'utf8', windowsHide: true, timeout: 20000 });
const firstLine = (handshake.stdout ?? '').split('\n').find((line) => line.trim().length > 0) ?? '';
try {
  const reply = JSON.parse(firstLine);
  console.log(`handshake          : ${JSON.stringify(reply.result?.serverInfo ?? reply)}`);
} catch {
  console.log(`handshake          : no JSON reply (stdout ${JSON.stringify((handshake.stdout ?? '').slice(0, 120))}, stderr ${JSON.stringify((handshake.stderr ?? '').slice(0, 200))})`);
}
