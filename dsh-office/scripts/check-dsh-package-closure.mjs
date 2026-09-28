import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';

const root = fileURLToPath(new URL('..', import.meta.url));
const manifestPath = join(root, 'dist', 'dsh-docx', 'package.json');
const manifest = JSON.parse(await readFile(manifestPath, 'utf8'));

// The DSH bridge is a host runtime.  Shipping it inside a Profile creates a
// second module graph beside the Desktop host and can make ESM imports cross
// incompatible versions.  The package must carry only office-owned runtime
// dependencies and require this exact host family as peers.
const hostRuntimePeers = [
  '@deepseek-ai/dsh-mcp-client',
  '@deepseek-ai/dsh-attachment',
  '@deepseek-ai/dsh-llm',
  '@deepseek-ai/dsh-scope',
  '@deepseek-ai/dsh-subprocess',
  '@deepseek-ai/dsh-timeout',
  '@deepseek-ai/dsh-tools',
];
const hostRuntimeRange = '>=0.2.0-rc.1 <0.3.0';

const bundledHostPackages = Object.keys(manifest.dependencies ?? {}).filter((name) => name.startsWith('@deepseek-ai/dsh-'));
if (bundledHostPackages.length) {
  throw new Error(`DSH plugin must not bundle host runtime packages: ${bundledHostPackages.join(', ')}`);
}

const missing = hostRuntimePeers.filter((name) => !manifest.peerDependencies?.[name]);
if (missing.length) {
  throw new Error(
    `DSH plugin package has an incomplete host-runtime peer contract: ${missing.join(', ')}`,
  );
}

const incompatible = hostRuntimePeers.filter((name) => manifest.peerDependencies?.[name] !== hostRuntimeRange);
if (incompatible.length) {
  throw new Error(
    `DSH plugin host-runtime peers must declare ${hostRuntimeRange}: ${incompatible.join(', ')}`,
  );
}

if (!manifest.peerDependencies?.['@deepseek-ai/cordis']) {
  throw new Error('DSH plugin package must declare the host-owned @deepseek-ai/cordis peer dependency.');
}

console.log(`DSH host-runtime boundary verified: ${hostRuntimePeers.length} DSH peers + host Cordis peer; no duplicated DSH packages.`);
