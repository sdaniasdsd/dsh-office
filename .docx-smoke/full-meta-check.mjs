// 元包（一步装）到底装不装得上：在一个干净目录里只依赖 -full.tgz，
// 看 pnpm 会不会按它的直链把核心包与运行时包从 release 拉下来。
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

const DIST = 'D:\\开源团队作品\\dsh-office-clone\\dsh-office\\dist';
const ROOT = 'D:\\认真版agent\\.docx-smoke\\full-meta-check';
const FULL_TGZ = join(DIST, 'deepseek-ai-dsh-docx-full-0.9.1.tgz');

rmSync(ROOT, { recursive: true, force: true });
mkdirSync(ROOT, { recursive: true });
writeFileSync(join(ROOT, 'package.json'), `${JSON.stringify({
  name: 'full-meta-consumer', private: true, version: '0.0.0',
  dependencies: { '@deepseek-ai/dsh-docx-full': `file:${FULL_TGZ.replaceAll('\\', '/')}` },
}, null, 2)}\n`);

console.log('$ pnpm install   （元包的依赖是 release 直链，需要能访问 GitHub）');
const install = spawnSync('pnpm', ['install'], { cwd: ROOT, stdio: 'inherit', windowsHide: true, shell: process.platform === 'win32' });
console.log(`pnpm exit: ${install.status}`);
if (install.status !== 0) { console.log('\n结论：元包没能一步装好（看上面的错误）'); process.exit(1); }

const core = join(ROOT, 'node_modules', '@deepseek-ai', 'dsh-docx');
const runtime = join(ROOT, 'node_modules', '@deepseek-ai', 'dsh-docx-runtime');
console.log(`\n核心包装上: ${existsSync(join(core, 'lib', 'server.mjs'))}`);
console.log(`运行时包装上: ${existsSync(join(runtime, 'runtime.json'))}`);
if (!existsSync(join(core, 'dsh', 'runtime-config.mjs'))) { console.log('结论：核心包缺失，一步装不成立'); process.exit(1); }

const { resolveDshConfig } = await import(pathToFileURL(join(core, 'dsh', 'runtime-config.mjs')).href);
const config = await resolveDshConfig({ workspaceRoot: ROOT });
console.log(`运行时解析: source=${config.env.DSH_DOCX_RUNTIME_SOURCE}  缺失=${config.env.DSH_DOCX_RUNTIME_MISSING ?? '(无)'}`);
const ok = config.env.DSH_DOCX_RUNTIME_SOURCE === 'runtime-package' && !config.env.DSH_DOCX_RUNTIME_MISSING;
console.log(`\n结论：${ok ? '元包一步装成立，且核心包找到了运行时 ✓' : '装上了但运行时解析不对 ✗'}`);
process.exit(ok ? 0 : 1);
