// 只回答一个问题：esbuild 的 metafile.inputs 键形状随 cwd 变化吗？
// build-dsh.mjs 用「键 → resolve(root,键) → 向上找 node_modules」来记 bundled 包，
// 键形状一变，那段匹配就可能一条都命中不到（实测从 85 掉到 0）。
import { createRequire } from 'node:module';
import { resolve, join } from 'node:path';

const root = 'D:\\开源团队作品\\dsh-office-clone\\dsh-office';
// 脚本本身在仓库外，esbuild 只能按绝对路径从仓库的 node_modules 里取。
const require = createRequire(join(root, 'package.json'));
const { build } = require('esbuild');
const entry = join(root, 'src', 'server.ts');
const probe = async (label) => {
  const result = await build({
    entryPoints: [entry], bundle: true, platform: 'node', format: 'esm', target: 'node22',
    external: ['pdfjs-dist'], write: false, metafile: true,
  });
  const keys = Object.keys(result.metafile.inputs);
  const withNodeModules = keys.filter((key) => key.replaceAll('\\', '/').includes('node_modules'));
  // 复刻 build-dsh.mjs 的判定：resolve(root,key) 是否以 root 开头、路径里是否有 node_modules 段
  let matched = 0;
  for (const key of keys) {
    let directory = resolve(root, key).replace(/\\/g, '/');
    const rootSlash = root.replace(/\\/g, '/');
    while (directory !== rootSlash && directory.startsWith(rootSlash)) {
      if (directory.split('/').includes('node_modules')) { matched += 1; break; }
      const parent = directory.slice(0, directory.lastIndexOf('/'));
      if (parent === directory) break;
      directory = parent;
    }
  }
  console.log(`[${label}] cwd=${process.cwd()}`);
  console.log(`  输入键数=${keys.length}  含 node_modules 的键=${withNodeModules.length}  复刻判定命中=${matched}`);
  console.log(`  前 3 个键: ${JSON.stringify(keys.slice(0, 3))}`);
  console.log(`  含 node_modules 的键示例: ${JSON.stringify(withNodeModules.slice(0, 2))}`);
};

await probe('cwd=仓库目录');
process.chdir('D:\\认真版agent');
await probe('cwd=别处');
