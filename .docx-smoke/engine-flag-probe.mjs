// Does the Python engine emit the formatting observation when the flag is on?
// Node's spawn passes arguments verbatim, unlike a PowerShell command line.
import { spawnSync } from 'node:child_process';

const [py, script, docx] = process.argv.slice(2);
const limits = {
  maxArchiveEntries: 4096, maxEntryUncompressedBytes: 33554432, maxTotalUncompressedBytes: 268435456,
  maxRelationships: 4096, maxBlocks: 50000, maxTableCells: 200000, maxAnnotations: 10000,
};

for (const parseFormatting of [true, false]) {
  const config = JSON.stringify({ limits, featureFlags: { parseFormatting } });
  const run = spawnSync(py, [script, '--path', docx, '--config', config], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
  console.log(`\n--- parseFormatting=${parseFormatting}  exit=${run.status}`);
  if (run.stderr?.trim()) console.log(`  stderr: ${run.stderr.trim().slice(0, 300)}`);
  try {
    const parsed = JSON.parse(run.stdout);
    const blocks = parsed.blocks ?? [];
    const carrying = blocks.filter((block) => block.formatting);
    console.log(`  blocks=${blocks.length}  carrying formatting=${carrying.length}`);
    if (carrying[0]) console.log(`  sample: ${JSON.stringify(carrying[0].formatting).slice(0, 260)}`);
  } catch (error) {
    console.log(`  stdout is not JSON (${run.stdout.length} chars): ${String(error).slice(0, 120)}`);
    console.log(`  first 200: ${run.stdout.slice(0, 200)}`);
  }
}
