import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';

const root = fileURLToPath(new URL('..', import.meta.url));
const manifestPath = join(root, 'dist', 'dsh-docx', 'package.json');
const manifest = JSON.parse(await readFile(manifestPath, 'utf8'));

// dsh-mcp-client imports these peers during module evaluation.  They must be
// normal package dependencies of the delivered plugin, not accidental matches
// from a developer's npm cache or the host application's install tree.
const startupDependencies = [
  '@deepseek-ai/dsh-attachment',
  '@deepseek-ai/dsh-subprocess',
  '@deepseek-ai/dsh-timeout',
  '@deepseek-ai/dsh-tools',
];

const missing = startupDependencies.filter((name) => !manifest.dependencies?.[name]);
if (missing.length) {
  throw new Error(
    `DSH plugin package has an incomplete cold-start dependency closure: ${missing.join(', ')}`,
  );
}

if (!manifest.peerDependencies?.['@deepseek-ai/cordis']) {
  throw new Error('DSH plugin package must declare the host-owned @deepseek-ai/cordis peer dependency.');
}

console.log(`DSH package closure verified: ${startupDependencies.length} startup dependencies + host Cordis peer.`);
