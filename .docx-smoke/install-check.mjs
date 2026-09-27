// 真实安装验证：在一个全新的消费方目录里，用**两个 file: 依赖**装上核心包与运行时包，
// 然后从装好的核心包里调用 resolveDshConfig，看它能不能自己找到兄弟运行时包。
// 用 pnpm（DSH profile 就是 pnpm 布局），因为隔离布局下"能不能找到"这件事最不确定。
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

const DIST = 'D:\\开源团队作品\\dsh-office-clone\\dsh-office\\dist';
const ROOT = 'D:\\认真版agent\\.docx-smoke\\install-check';
const CORE_TGZ = join(DIST, 'deepseek-ai-dsh-docx-0.9.0.tgz');
const RUNTIME_TGZ = join(DIST, 'deepseek-ai-dsh-docx-runtime-0.9.0.tgz');

rmSync(ROOT, { recursive: true, force: true });
mkdirSync(ROOT, { recursive: true });
writeFileSync(join(ROOT, 'package.json'), `${JSON.stringify({
  name: 'dsh-profile-like-consumer',
  private: true,
  version: '0.0.0',
  dependencies: {
    '@deepseek-ai/dsh-docx': `file:${CORE_TGZ.replaceAll('\\', '/')}`,
    '@deepseek-ai/dsh-docx-runtime': `file:${RUNTIME_TGZ.replaceAll('\\', '/')}`,
  },
}, null, 2)}\n`);

console.log('$ pnpm install  （两个 file: 依赖，运行时包 1.67 GB 解压）');
const install = spawnSync('pnpm', ['install', '--prefer-offline'], { cwd: ROOT, stdio: 'inherit', windowsHide: true, shell: process.platform === 'win32' });
console.log(`pnpm exit: ${install.status}`);
if (install.status !== 0) process.exit(install.status ?? 1);

const corePath = join(ROOT, 'node_modules', '@deepseek-ai', 'dsh-docx');
const runtimePath = join(ROOT, 'node_modules', '@deepseek-ai', 'dsh-docx-runtime');
console.log(`\n核心包装上了: ${existsSync(join(corePath, 'lib', 'server.mjs'))}`);
console.log(`运行时包装上了: ${existsSync(join(runtimePath, 'runtime.json'))}`);

const { resolveDshConfig } = await import(pathToFileURL(join(corePath, 'dsh', 'runtime-config.mjs')).href);
const config = await resolveDshConfig({ workspaceRoot: ROOT });
console.log(`\nruntime source = ${config.env.DSH_DOCX_RUNTIME_SOURCE}`);
console.log(`缺失项         = ${config.env.DSH_DOCX_RUNTIME_MISSING ?? '(无)'}`);
for (const key of ['DOCX_PYTHON', 'DOCX_SOFFICE', 'DOCX_PDFTOPPM']) {
  const value = config.env[key];
  console.log(`  ${key.padEnd(14)} = ${value && existsSync(value) ? '存在' : '(未注入或不存在)'}`);
}
const ok = config.env.DSH_DOCX_RUNTIME_SOURCE === 'runtime-package' && !config.env.DSH_DOCX_RUNTIME_MISSING;
console.log(`\n结论：${ok ? '核心包在真实 pnpm 安装里自己找到了运行时 ✓' : '核心包没能找到运行时 ✗'}`);
process.exit(ok ? 0 : 1);
