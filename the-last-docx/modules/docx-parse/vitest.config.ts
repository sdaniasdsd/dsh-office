import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    // 全局准备：重建 fixtures（生成器见 fixtures/generate_fixtures.py）。
    globalSetup: ['./tests/global-setup.ts'],
    environment: 'node',
    testTimeout: 60_000,
    hookTimeout: 60_000,
    sequence: { concurrent: false },
  },
});
