import js from '@eslint/js';
import vitest from '@vitest/eslint-plugin';
import prettier from 'eslint-config-prettier';
import { createTypeScriptImportResolver } from 'eslint-import-resolver-typescript';
import importPlugin from 'eslint-plugin-import-x';
import globals from 'globals';
import tseslint from 'typescript-eslint';

export default tseslint.config(
  {
    ignores: [
      'dist',
      'coverage',
      // Committed build output of the skill's agent tools.
      'skills/*/agent/dist',
      // Legacy demo apps, pinned to ancient toolchains and not part of the build.
      'examples',
      'src/__tests__/__snapshots__',
    ],
  },

  js.configs.recommended,
  tseslint.configs.recommended,
  importPlugin.flatConfigs.recommended,
  importPlugin.flatConfigs.typescript,

  {
    languageOptions: {
      globals: globals.node,
      parserOptions: {
        // The `.mjs` build/lint configs are not part of the TS program, but we
        // still want them linted.
        projectService: { allowDefaultProject: ['*.mjs'] },
        tsconfigRootDir: import.meta.dirname,
      },
    },
    settings: {
      'import-x/resolver-next': [createTypeScriptImportResolver()],
    },
    rules: {
      // `any` is part of the contract here: user data, sandbox contents and
      // JS-snippet results are arbitrary by design.
      '@typescript-eslint/no-explicit-any': 'off',
      '@typescript-eslint/consistent-type-imports': [
        'error',
        { fixStyle: 'inline-type-imports' },
      ],
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
          caughtErrorsIgnorePattern: '^_',
        },
      ],

      eqeqeq: ['error', 'allow-null'],
      'dot-notation': 'error',
      'guard-for-in': 'warn',
      'linebreak-style': ['error', 'unix'],
      'no-extra-bind': 'error',
      'no-loop-func': 'error',
      'no-throw-literal': 'error',
      'prefer-promise-reject-errors': 'error',
      radix: 'error',

      'import-x/no-extraneous-dependencies': [
        'error',
        {
          devDependencies: true,
          peerDependencies: true,
          optionalDependencies: false,
        },
      ],
      // `jszip`, `qrcode` and `sax` are CommonJS namespaces whose members are
      // meant to be reached through the default import (`QR.toDataURL()`).
      'import-x/no-named-as-default-member': 'off',
      'import-x/prefer-default-export': 'off',
      // NOTE: `import-x/no-unused-modules` (dead-export detection) is a no-op on
      // ESLint 10, which removed the FileEnumerator API the rule depends on.
      // Left out rather than enabled-but-silent so the config doesn't lie.
    },
  },

  {
    files: ['src/**/__tests__/**/*.ts'],
    ...vitest.configs.recommended,
  },

  // Must stay last: turns off every rule that would fight Prettier.
  prettier
);
