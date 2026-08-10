import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['{src,skills}/**/__tests__/**/*.{test,spec}.ts'],
    testTimeout: 10_000,
    coverage: {
      provider: 'v8',
      reporter: ['text', 'html', 'lcov'],
      include: ['src/**/*.ts'],
      exclude: ['src/debug.ts', 'src/**/__tests__/**'],
      // A floor, not a target: set a little below what the suite currently
      // achieves, so that a drop is noticed but small refactors aren't blocked.
      // What is left uncovered is defensive code that cannot be reached from
      // the public API (`node parent is null` guards, `?? null` fallbacks that
      // `noUncheckedIndexedAccess` requires, ...).
      thresholds: {
        statements: 97,
        branches: 93,
        functions: 98,
        lines: 98,
      },
    },
  },
});
