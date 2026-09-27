// Run the MCP server straight from the repository's TypeScript sources, from
// outside the repository: absolute paths only, so nothing has to be written into
// the checkout to exercise the real tool contract.
//
// A passthrough on purpose: the parent hands us pipes, and `stdio: 'inherit'`
// gives the child exactly those file descriptors, so the JSON-RPC stream and the
// stderr diagnostics keep their own channels.
import { spawn } from 'node:child_process';
import { join } from 'node:path';

const ROOT = 'D:\\开源团队作品\\the-last-docx';
const child = spawn(process.execPath, [join(ROOT, 'node_modules', 'tsx', 'dist', 'cli.mjs'), join(ROOT, 'src', 'server.ts')], {
  stdio: 'inherit',
  cwd: ROOT,
  env: process.env,
});
child.on('exit', (code, signal) => process.exit(signal ? 1 : code ?? 0));
child.on('error', (error) => {
  process.stderr.write(`source-server proxy failed: ${error.message}\n`);
  process.exit(1);
});
