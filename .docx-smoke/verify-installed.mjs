// Post-install verification: exercise the INSTALLED copy's activation-time
// validation (this is the code that runs at DSH boot and would fail the profile).
import { readFileSync } from 'node:fs';

const root = 'file:///C:/Users/AA/AppData/Roaming/com.yeagoo.dsh-desktop/harness/profiles/web/node_modules/@deepseek-ai/dsh-docx/';

const { resolveDshConfig } = await import(root + 'dsh/runtime-config.mjs');

const cfg = await resolveDshConfig({ workspaceRoot: 'D:\\认真版agent' });

console.log('resolveDshConfig OK');
console.log(JSON.stringify({
  serverName: cfg.serverName,
  transport: cfg.transport,
  command: cfg.command,
  args: cfg.args,
  cwd: cfg.cwd,
  toolCallTimeoutMs: cfg.toolCallTimeoutMs,
  failOnStartupError: cfg.failOnStartupError,
  envKeys: Object.keys(cfg.env),
}, null, 2));

const mcpPkg = new URL(root + 'node_modules-does-not-exist').href; // placeholder, unused
const mcpPath = 'C:\\Users\\AA\\AppData\\Roaming\\com.yeagoo.dsh-desktop\\harness\\profiles\\web\\node_modules\\@deepseek-ai\\dsh-mcp-client\\package.json';
console.log('mcp-client on disk:', JSON.parse(readFileSync(mcpPath, 'utf8')).version);

// The plugin's own entry does `import * as mcpClient from '@deepseek-ai/dsh-mcp-client'`.
// Resolve it the way Node will, from the installed entry's real directory.
const resolved = import.meta.resolve('@deepseek-ai/dsh-mcp-client', root + 'dsh/index.mjs');
console.log('node resolves mcp-client to:', resolved);
const mod = await import(resolved);
console.log('mcp-client exports:', Object.keys(mod).join(', '));
