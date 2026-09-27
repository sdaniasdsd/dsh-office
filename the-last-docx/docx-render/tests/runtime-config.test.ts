import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { resolveConfig } from '../src/config';

const roots: string[] = [];

function runtimeFixture(): { root: string; soffice: string; pdftoppm: string } {
  const root = mkdtempSync(join(tmpdir(), 'dsh-render-runtime-'));
  roots.push(root);
  const runtime = join(root, 'runtime', 'win32-x64');
  const soffice = join(runtime, 'libreoffice', 'program', 'soffice.com');
  const pdftoppm = join(runtime, 'poppler', 'current', 'Library', 'bin', 'pdftoppm.exe');
  mkdirSync(join(runtime, 'libreoffice', 'program'), { recursive: true });
  mkdirSync(join(runtime, 'poppler', 'current', 'Library', 'bin'), { recursive: true });
  writeFileSync(soffice, '');
  writeFileSync(pdftoppm, '');
  writeFileSync(join(root, 'runtime.json'), JSON.stringify({
    schema: 'dsh-office-runtime/v1',
    components: {
      libreoffice: { present: true, entry: 'libreoffice\\program\\soffice.com' },
      poppler: { present: true, entry: 'poppler\\current\\Library\\bin\\pdftoppm.exe' },
    },
  }));
  return { root, soffice, pdftoppm };
}

afterEach(() => { while (roots.length > 0) rmSync(roots.pop()!, { recursive: true, force: true }); });

describe('DSH toolchain runtime configuration', () => {
  it('derives both renderer executables from a runtime manifest', () => {
    const fixture = runtimeFixture();
    const config = resolveConfig({ runtimeRoot: fixture.root });
    expect(config.engine.sofficePath).toBe(fixture.soffice);
    expect(config.engine.pdftoppmPath).toBe(fixture.pdftoppm);
  });

  it('keeps an explicit executable path ahead of the shared runtime', () => {
    const fixture = runtimeFixture();
    const config = resolveConfig({
      runtimeRoot: fixture.root,
      engine: { sofficePath: 'host-managed-soffice' },
    });
    expect(config.engine.sofficePath).toBe('host-managed-soffice');
    expect(config.engine.pdftoppmPath).toBe(fixture.pdftoppm);
  });
});

