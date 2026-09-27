import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { resolveRuntime } from '../src/index';

const roots: string[] = [];

function runtimeFixture(): { root: string; python: string; soffice: string; pdftoppm: string } {
  const root = mkdtempSync(join(tmpdir(), 'dsh-runtime-stress-'));
  roots.push(root);
  const runtime = join(root, 'runtime', 'win32-x64');
  const python = join(runtime, 'python', 'python.exe');
  const soffice = join(runtime, 'libreoffice', 'program', 'soffice.com');
  const pdftoppm = join(runtime, 'poppler', 'current', 'Library', 'bin', 'pdftoppm.exe');
  for (const path of [python, soffice, pdftoppm]) mkdirSync(join(path, '..'), { recursive: true });
  writeFileSync(python, '');
  writeFileSync(soffice, '');
  writeFileSync(pdftoppm, '');
  writeFileSync(join(root, 'runtime.json'), JSON.stringify({
    schema: 'dsh-office-runtime/v1',
    components: {
      python: { present: true, entry: 'python\\python.exe' },
      libreoffice: { present: true, entry: 'libreoffice\\program\\soffice.com' },
      poppler: { present: true, entry: 'poppler\\current\\Library\\bin\\pdftoppm.exe' },
    },
  }));
  return { root, python, soffice, pdftoppm };
}

afterEach(() => { while (roots.length > 0) rmSync(roots.pop()!, { recursive: true, force: true }); });

const components = [
  { id: 'python', envVar: 'DOCX_PYTHON', defaultEntry: 'python/python.exe' },
  { id: 'libreoffice', envVar: 'DOCX_SOFFICE', defaultEntry: 'libreoffice/program/soffice.com' },
  { id: 'poppler', envVar: 'DOCX_PDFTOPPM', defaultEntry: 'poppler/current/Library/bin/pdftoppm.exe' },
];

describe('runtime resolution: 200-cycle engineering stress gate', () => {
  it('preserves precedence, manifest interpretation, and missing-runtime degradation for 200 cycles', () => {
    const fixture = runtimeFixture();
    const expected = [fixture.python, fixture.soffice, fixture.pdftoppm];
    for (let cycle = 0; cycle < 200; cycle += 1) {
      const mode = cycle % 5;
      const result = mode === 0
        ? resolveRuntime({ components, runtimeRoot: fixture.root, environment: {} })
        : mode === 1
          ? resolveRuntime({ components, environment: { DSH_OFFICE_RUNTIME_ROOT: fixture.root } })
          : mode === 2
            ? resolveRuntime({
              components,
              componentPaths: { python: fixture.python },
              environment: { DOCX_SOFFICE: fixture.soffice, DSH_OFFICE_RUNTIME_ROOT: fixture.root },
            })
            : mode === 3
              ? resolveRuntime({ components, environment: {} })
              : resolveRuntime({
                components,
                componentPaths: {
                  python: fixture.python,
                  libreoffice: fixture.soffice,
                  poppler: fixture.pdftoppm,
                },
                environment: {},
              });

      if (mode === 3) {
        expect(result.missing, `cycle ${cycle}`).toEqual(['python', 'libreoffice', 'poppler']);
        expect(result.source, `cycle ${cycle}`).toBe('unresolved');
      } else {
        expect(components.map((component) => result.components[component.id]?.path), `cycle ${cycle}`).toEqual(expected);
        expect(result.missing, `cycle ${cycle}`).toEqual([]);
        const sources = components.map((component) => result.components[component.id]?.source);
        expect(sources, `cycle ${cycle}`).toEqual(
          mode === 0
            ? ['config-root', 'config-root', 'config-root']
            : mode === 1
              ? ['env-root', 'env-root', 'env-root']
              : mode === 2
                ? ['config-path', 'env-path', 'env-root']
                : ['config-path', 'config-path', 'config-path'],
        );
      }
    }
  });
});

