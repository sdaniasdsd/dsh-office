// Stage a src/mcp.ts that carries the version fix, WITHOUT touching the clone's
// working copy of that file (another line of work is uncommitted there).
//
// The committed blob is written straight into the index: `git update-index
// --cacheinfo` takes a blob hash, so the file on disk never changes.
import { spawnSync } from 'node:child_process';
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';

const CLONE_ROOT = 'D:\\开源团队作品\\dsh-office-clone';
const PATH_IN_REPO = 'the-last-docx/src/mcp.ts';
const TEMP = join('D:\\认真版agent\\.docx-smoke', 'mcp-staged.ts');

const git = (args, options = {}) => {
  const result = spawnSync('git', ['-C', CLONE_ROOT, ...args], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024, ...options });
  if (result.status !== 0) throw new Error(`git ${args.join(' ')} failed: ${result.stderr}`);
  return result.stdout;
};

const committed = git(['show', `HEAD:${PATH_IN_REPO}`]);
console.log(`carryFormatting already published: ${committed.includes('Set options.carryFormatting')}`);

const before = `export function createMcpServer(profile:DocxProfile){\n  const server=new McpServer({name:'the-last-docx',version:'0.5.0'});`;
const after = [
  '/**',
  ' * The version this server reports in the MCP handshake.',
  ' *',
  ' * The packaging script (`scripts/build-dsh.mjs`) defines it from the version it',
  ' * stamps into the published manifest, so the handshake can no longer drift from',
  ' * the package it came out of — which it did: an installed 0.6.0 package still',
  ' * announced itself as 0.5.0. Running from TypeScript sources, where nothing',
  ' * defines it, says so instead of guessing.',
  ' */',
  'declare const __DSH_DOCX_VERSION__: string | undefined;',
  "const VERSION = typeof __DSH_DOCX_VERSION__ === 'string' ? __DSH_DOCX_VERSION__ : '0.0.0-dev';",
  'export function createMcpServer(profile:DocxProfile){',
  '  const server=new McpServer({name:\'the-last-docx\',version:VERSION});',
].join('\n');

if (!committed.includes(before)) throw new Error('the committed header is not the shape this patch expects');
let patched = committed.replace(before, after);

// The reference tool's description gained the `carryFormatting` sentence — taken
// verbatim from the working file, because re-typing it dropped an escape and the
// committed blob stopped parsing (caught by typechecking the commit, not the
// working tree).
const anchor = 'To restyle a document that carries annotations, edit that document with docx-edit instead.';
const mine = readFileSync('D:\\开源团队作品\\the-last-docx\\src\\mcp.ts', 'utf8');
const start = mine.indexOf('Set options.carryFormatting');
const end = mine.indexOf(anchor);
if (start < 0 || end < 0 || end < start) throw new Error('could not lift the sentence out of the working file');
const sentence = mine.slice(start, end);
console.log(`lifting ${sentence.length} chars verbatim: ${JSON.stringify(sentence.slice(0, 60))}…`);
if (!patched.includes('Set options.carryFormatting')) {
  if (!patched.includes(anchor)) throw new Error('the reference tool description anchor was not found');
  patched = patched.replace(anchor, `${sentence}${anchor}`);
}
writeFileSync(TEMP, patched);

const hash = git(['hash-object', '-w', TEMP]).trim();
git(['update-index', '--cacheinfo', `100644,${hash},${PATH_IN_REPO}`]);
console.log(`staged blob ${hash.slice(0, 12)}… for ${PATH_IN_REPO}`);
console.log(git(['status', '--short', '--', PATH_IN_REPO]).trim() || '(no status line)');
