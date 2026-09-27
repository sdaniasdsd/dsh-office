// Read the deck back with python-pptx (independent of the plugin's own report)
// and print a compact, parseable summary.
import { spawnSync } from 'node:child_process';
import { join } from 'node:path';

const PYTHON = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web', 'node_modules', '@deepseek-ai', 'dsh-docx', 'runtime', 'win32-x64', 'python', 'python.exe');
const SCRIPT = 'D:\\认真版agent\\.docx-smoke\\inspect-pptx.py';
const DECK = process.argv[2];

const result = spawnSync(PYTHON, [SCRIPT, DECK], { encoding: 'utf8', windowsHide: true, maxBuffer: 64 * 1024 * 1024, env: { ...process.env, PYTHONIOENCODING: 'utf-8' } });
if (result.status !== 0) {
  console.error(`python exited ${result.status}\n${result.stderr}`);
  process.exit(1);
}
const report = JSON.parse(result.stdout);
console.log(`slides ${report.slides}  ${report.widthInches}×${report.heightInches} in`);
console.log(`title sizes ${report.titleSizes.join(', ')} pt   body sizes ${report.bodySizes.join(', ')} pt`);
console.log(`unsized runs ${report.unsizedRuns.length}   speaker notes ${report.totalNotesChars} chars`);
for (const slide of report.perSlide) {
  const shapes = slide.shapes.map((shape) => `[${shape.placeholder} ${shape.sizes.join('/')}pt] ${shape.text}`).join('  ||  ');
  console.log(`\nslide ${String(slide.slide).padStart(2)}  notes=${slide.notesChars}字\n  ${shapes}`);
}
