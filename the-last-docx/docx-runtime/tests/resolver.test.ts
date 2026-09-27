import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { DSH_RUNTIME_SCHEMA, resolveRuntime } from '../src/index';

const roots: string[] = [];
function fixture(): { root: string; runtimeRoot: string; python: string; soffice: string } {
  const root = mkdtempSync(join(tmpdir(), 'dsh-runtime-'));
  roots.push(root);
  const runtimeRoot = join(root, 'runtime', 'win32-x64');
  const python = join(runtimeRoot, 'python', 'python.exe');
  const soffice = join(runtimeRoot, 'libreoffice', 'program', 'soffice.com');
  mkdirSync(join(runtimeRoot, 'python'), { recursive: true });
  mkdirSync(join(runtimeRoot, 'libreoffice', 'program'), { recursive: true });
  writeFileSync(python, '');
  writeFileSync(soffice, '');
  writeFileSync(join(root, 'runtime.json'), JSON.stringify({
    schema: DSH_RUNTIME_SCHEMA,
    components: {
      python: { present: true, entry: 'python\\python.exe' },
      libreoffice: { present: true, entry: 'libreoffice\\program\\soffice.com' },
    },
  }));
  return { root, runtimeRoot, python, soffice };
}
afterEach(() => { while (roots.length > 0) rmSync(roots.pop()!, { recursive: true, force: true }); });

const specs = [
  { id: 'python', envVar: 'DOCX_PYTHON', defaultEntry: 'python/python.exe' },
  { id: 'libreoffice', envVar: 'DOCX_SOFFICE', defaultEntry: 'libreoffice/program/soffice.com' },
];

describe('runtime resolution', () => {
  it('reads manifest entries from a configured package root', () => {
    const f = fixture();
    const result = resolveRuntime({ components: specs, runtimeRoot: f.root, environment: {} });
    expect(result.components.python).toMatchObject({ path: f.python, source: 'config-root', available: true });
    expect(result.components.libreoffice).toMatchObject({ path: f.soffice, source: 'config-root', available: true });
  });

  it('uses direct configuration before a runtime root and environment values', () => {
    const f = fixture();
    const result = resolveRuntime({
      components: specs,
      runtimeRoot: f.root,
      componentPaths: { python: f.soffice },
      environment: { DOCX_PYTHON: f.python },
    });
    expect(result.components.python).toMatchObject({ path: f.soffice, source: 'config-path' });
  });

  it('allows a future plugin to supply arbitrary component ids and package names', () => {
    const f = fixture();
    const caller = join(f.root, 'plugins', 'next-plugin', 'src', 'index.ts');
    const packageRoot = join(f.root, 'plugins', 'node_modules', '@example', 'future-runtime');
    const entry = join(packageRoot, 'runtime', 'win32-x64', 'engine', 'tool.exe');
    mkdirSync(join(packageRoot, 'runtime', 'win32-x64', 'engine'), { recursive: true });
    writeFileSync(entry, '');
    writeFileSync(join(packageRoot, 'runtime.json'), JSON.stringify({
      schema: DSH_RUNTIME_SCHEMA,
      components: { converter: { present: true, entry: 'engine/tool.exe' } },
    }));
    const result = resolveRuntime({
      components: [{ id: 'converter', defaultEntry: 'engine/tool.exe' }],
      runtimePackageNames: ['@example/future-runtime'],
      callerPath: caller,
      environment: {},
    });
    expect(result.components.converter).toMatchObject({ path: entry, source: 'runtime-package' });
  });

  it('reports absence without throwing or inventing a command path', () => {
    const result = resolveRuntime({ components: specs, environment: {} });
    expect(result.missing).toEqual(['python', 'libreoffice']);
    expect(result.components.python).toMatchObject({ path: undefined, source: 'unresolved', available: false });
  });
});

