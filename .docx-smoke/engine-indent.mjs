// Run the parse engine directly and print the first-line indent it observed for
// every paragraph: the fastest way to see style resolution working (or not).
import { spawnSync } from 'node:child_process';

const SCRIPT = process.argv[2];
const PYTHON = process.argv[3];
const CONFIG = JSON.stringify({
  engine: { driver: 'python', pythonPath: PYTHON, timeoutMs: 120000 },
  limits: {},
  featureFlags: { parseFormatting: true },
});

for (const path of process.argv.slice(4)) {
  const run = spawnSync(PYTHON, [SCRIPT, '--path', path, '--config', CONFIG], {
    cwd: SCRIPT.replace(/[\\/][^\\/]+$/u, ''),
    encoding: 'utf8',
    env: { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1', PYTHONDONTWRITEBYTECODE: '1' },
    maxBuffer: 64 * 1024 * 1024,
  });
  if (run.status !== 0) {
    console.log(`\n${path}: engine exited ${run.status}\n${(run.stderr ?? '').slice(0, 800)}`);
    continue;
  }
  const result = JSON.parse(run.stdout);
  const blocks = result.blocks ?? result.result?.blocks ?? [];
  const withFormatting = blocks.filter((block) => block.formatting);
  const byIndent = new Map();
  const samples = new Map();
  for (const block of withFormatting) {
    const indent = block.formatting.indent;
    const key = indent === undefined ? '(none)' : JSON.stringify(indent);
    byIndent.set(key, (byIndent.get(key) ?? 0) + 1);
    if (!samples.has(key)) samples.set(key, []);
    const list = samples.get(key);
    if (list.length < 3) list.push(String(block.text ?? '').trim().slice(0, 18));
  }
  console.log(`\n===== ${path} =====`);
  console.log(`blocks ${blocks.length}, observed ${withFormatting.length}`);
  for (const [key, count] of [...byIndent].sort((a, b) => b[1] - a[1])) {
    console.log(`  ${String(count).padStart(4)} × indent=${key}   e.g. ${samples.get(key).join(' | ')}`);
  }
}
