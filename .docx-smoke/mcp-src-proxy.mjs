// Run the MCP server straight from this repository's TypeScript sources, so a
// case can be exercised through the real tool contract (docx_call, the safety
// guard, the artifact store) without packaging the plugin first.
//
// A passthrough on purpose: the parent hands us pipes, and `stdio: 'inherit'`
// gives the child exactly those file descriptors, so the JSON-RPC stream and the
// stderr diagnostics keep their own channels.
import { spawn } from 'node:child_process';
import { join } from 'node:path';

const root = import.meta.dirname;
const child = spawn(process.execPath, [join(root, 'node_modules', 'tsx', 'dist', 'cli.mjs'), 'src/server.ts'], {
  stdio: 'inherit',
  cwd: root,
  env: process.env,
});
child.on('exit', (code, signal) => process.exit(signal ? 1 : code ?? 0));
child.on('error', (error) => {
  process.stderr.write(`source-server proxy failed: ${error.message}\n`);
  process.exit(1);
});
