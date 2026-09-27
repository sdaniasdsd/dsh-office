// What does the installed 0.8.0 bundle actually register, and what version does
// it claim? The handshake disagrees with package.json, so both are read here.
import { readFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';

const installed = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web', 'node_modules', '@deepseek-ai', 'dsh-docx');
for (const file of ['lib/server.mjs', 'dsh/index.mjs']) {
  const path = join(installed, file);
  if (!existsSync(path)) { console.log(`${file}: missing`); continue; }
  const text = readFileSync(path, 'utf8');
  console.log(`\n=== ${file} (${text.length} chars) ===`);
  console.log(`registerTool occurrences : ${(text.match(/registerTool/gu) ?? []).length}`);
  const names = [...text.matchAll(/registerTool\(\s*["'`]([^"'`]+)["'`]/gu)].map((match) => match[1]);
  console.log(`tool names               : ${names.join(', ') || '(none matched)'}`);
  // Every quoted string that looks like a tool name, as a fallback view.
  const candidates = [...new Set([...text.matchAll(/["'`](docx_[a-z_]+|pptx_[a-z_]+|xlsx_[a-z_]+)["'`]/gu)].map((match) => match[1]))];
  console.log(`tool-like strings        : ${candidates.join(', ') || '(none)'}`);
  for (const version of ['0.5.0', '0.7.0', '0.8.0', '0.0.0-dev']) {
    console.log(`  literal ${version.padEnd(10)}: ${(text.split(version).length - 1)}`);
  }
  const server = /new McpServer\(\{[^}]*\}/u.exec(text)?.[0];
  console.log(`McpServer construction   : ${server ?? '(not found as a literal)'}`);
}
