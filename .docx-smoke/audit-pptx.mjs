// 跑 run 级审计，解释计数差。
import { spawnSync } from 'node:child_process';
import { join } from 'node:path';

const PYTHON = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web', 'node_modules', '@deepseek-ai', 'dsh-docx', 'runtime', 'win32-x64', 'python', 'python.exe');
const SCRIPT = 'D:\\认真版agent\\.docx-smoke\\audit-pptx-runs.py';
for (const deck of process.argv.slice(2)) {
  const result = spawnSync(PYTHON, [SCRIPT, deck], { encoding: 'utf8', windowsHide: true, env: { ...process.env, PYTHONIOENCODING: 'utf-8' } });
  if (result.status !== 0) { console.error(`${deck}: python exited ${result.status}\n${result.stderr}`); continue; }
  const audit = JSON.parse(result.stdout);
  console.log(`\n${deck.split('\\').pop()}`);
  console.log(`  total runs ${audit.totalRuns}  empty runs ${audit.emptyRuns}  sizes ${JSON.stringify(audit.sizeHistogram)}`);
  console.log(`  runs not at 30/18 pt: ${audit.notAtTarget.length}`);
  for (const row of audit.notAtTarget.slice(0, 8)) console.log(`    slide ${row.slide} shape ${row.shapeId} p${row.p} r${row.r} placeholder=${row.placeholder} size=${row.size}`);
  for (const row of audit.emptyDetail.slice(0, 8)) console.log(`    empty: slide ${row.slide} shape ${row.shapeId} p${row.p} r${row.r} placeholder=${row.placeholder} size=${row.size}`);
}
