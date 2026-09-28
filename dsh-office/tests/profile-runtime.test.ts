import { describe, expect, it } from 'vitest';
import { injectRuntimeConfig } from '../src/profile';

describe('Profile runtime integration', () => {
  it('passes trusted discovery settings without manufacturing engine-path overrides', () => {
    const config = injectRuntimeConfig({}, { runtimeRoot: 'D:/runtime', runtimePackageNames: ['@example/office-runtime'] });
    expect(config).toMatchObject({ runtimeRoot: 'D:/runtime', runtimePackageNames: ['@example/office-runtime'] });
    expect(config).not.toHaveProperty('engine');
  });

  it('does not manufacture undefined extension keys for modules with strict schemas', () => {
    expect(injectRuntimeConfig({}, {})).toEqual({});
  });

  it('keeps a module-specific trusted runtime selection ahead of the host default', () => {
    const config = injectRuntimeConfig({ runtimeRoot: 'D:/module-runtime', runtimePackageNames: ['@example/module-runtime'] }, { runtimeRoot: 'D:/host-runtime', runtimePackageNames: ['@example/host-runtime'] });
    expect(config).toMatchObject({ runtimeRoot: 'D:/module-runtime', runtimePackageNames: ['@example/module-runtime'] });
  });
});
