// 复刻 build-dsh.mjs 的 bundled 包统计：root 的求法、循环、判定都照抄，
// 但把 catch 里被吞掉的错误打出来。目的是回答「为什么 recorded: 0」。
import { createRequire } from 'node:module';
import { readFile } from 'node:fs/promises';
import { resolve, join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

// 模拟「脚本位于 <repo>/scripts/diag.mjs」时的 import.meta.url
const base = 'file:///D:/%E5%BC%80%E6%BA%90%E5%9B%A2%E9%98%9F%E4%BD%9C%E5%93%81/dsh-office-clone/dsh-office/scripts/diag.mjs';
const root = fileURLToPath(new URL('..', base));
console.log(`root = ${JSON.stringify(root)}`);

const require = createRequire(join(root, 'package.json'));
const { build } = require('esbuild');

// 真脚本不传 absWorkingDir（默认取 process.cwd()）。这里用开关对照。
const useAwd = process.argv.includes('--awd');
const entry = process.argv.includes('--index') ? 'index' : 'server';
const failures = [];
{
  const result = await build({
    entryPoints: [join(root, 'src', `${entry}.ts`)], bundle: true, platform: 'node', format: 'esm',
    target: 'node22', external: ['pdfjs-dist'], write: false, metafile: true,
    ...(useAwd ? { absWorkingDir: root } : {}),
  });
  const keys = Object.keys(result.metafile.inputs);
  console.log(`cwd=${process.cwd()}  absWorkingDir=${useAwd ? 'root' : '(默认 cwd)'}`);
  console.log(`  前 3 个输入键: ${JSON.stringify(keys.slice(0, 3))}`);
  const bundledPackages = new Map();
  for (const input of keys) {
    let directory = dirname(resolve(root, input));
    while (directory !== root && directory.startsWith(root)) {
      if (directory.split(/[\\/]/).includes('node_modules')) {
        try {
          const packageJson = JSON.parse(await readFile(join(directory, 'package.json'), 'utf8'));
          if (packageJson.name) bundledPackages.set(packageJson.name, directory);
          break;
        } catch (error) { if (failures.length < 3) failures.push(`${directory} -> ${error.code ?? error.message}`); }
      }
      const parent = dirname(directory);
      if (parent === directory) break;
      directory = parent;
    }
  }
  console.log(`  → 记录到 ${bundledPackages.size} 个包`);
  for (const line of failures) console.log(`     读失败: ${line}`);
}
