import { createRequire } from 'node:module';

import commonjs from '@rollup/plugin-commonjs';
import node from '@rollup/plugin-node-resolve';
import replace from '@rollup/plugin-replace';
import { defineConfig } from 'rollup';
import dtsPlugin from 'rollup-plugin-dts';
import esbuild from 'rollup-plugin-esbuild';

const require = createRequire(import.meta.url);

// Node built-ins the browser bundle has to stand in for. Resolved to absolute
// paths up front: calling `this.resolve()` from inside `resolveId` re-enters
// the plugin chain and deadlocks, because the polyfills themselves import the
// very built-ins being mapped.
const BROWSER_POLYFILLS = {
  vm: require.resolve('vm-browserify'),
  'node:vm': require.resolve('vm-browserify'),
};

// Optional Node-only requires that must stay unresolved in the browser bundle;
// see the `commonjs({ ignore })` call below.
const BROWSER_IGNORED_REQUIRES = new Set([
  'readable-stream',
  'stream',
  'node:stream',
]);

// Runtime dependencies are kept external in the Node builds so that consumers
// get a single copy of them; the browser bundle inlines everything instead.
const nodeExternal = ['jszip', 'sax', 'node:vm'];

// `tsc` adds `/// <reference types="node" />` to the declarations because the
// build itself is type-checked against Node's types, but the public API only
// speaks in web types (`Uint8Array`, `ArrayBuffer`). Dropping the directive
// keeps consumers from needing `@types/node` — with a guard so that this stays
// true if a Node-only type ever creeps back into the public surface.
const NODE_ONLY_TYPE = /\bNodeJS\.|(?::|<|\bextends )\s*Buffer\b/;

const selfContainedDts = {
  name: 'self-contained-dts',
  renderChunk(code) {
    const stripped = code
      .split('\n')
      .filter(line => !line.startsWith('/// <reference types="node"'))
      .join('\n');
    const offender = NODE_ONLY_TYPE.exec(stripped);
    if (offender) {
      this.error(
        `The generated declarations reference a Node-only type (${offender[0].trim()}). ` +
          `Either express it with a web type, or stop stripping the "node" type reference.`
      );
    }
    return stripped;
  },
};

export default defineConfig([
  // Node: ESM + CJS bundles from the same entry point.
  {
    input: './src/index.ts',
    external: nodeExternal,
    output: [
      {
        file: './dist/index.mjs',
        format: 'es',
        sourcemap: true,
      },
      {
        file: './dist/index.cjs',
        format: 'cjs',
        exports: 'named',
        sourcemap: true,
      },
    ],
    plugins: [esbuild({ target: 'node20' })],
  },

  // Type declarations. Both entry points expose exactly the same API, so one
  // bundle is emitted under every name the `exports` map can resolve to: a
  // different extension per module system, plus a plain `.d.ts` for legacy
  // (node10) resolution.
  {
    input: './src/index.ts',
    // `sax` is deliberately *not* external here: its `QualifiedAttribute` type
    // leaks into the public `Node` type, and `@types/sax` is only a
    // devDependency. Inlining it keeps the published declarations self-contained.
    // `stream` only shows up in the parts of `@types/sax` we don't pull in;
    // marking it external just silences the unresolved-import warning, and
    // `selfContainedDts` fails the build if it ever reaches the output.
    external: ['jszip', 'node:vm', 'stream'],
    output: [
      { file: './dist/index.d.ts', format: 'es' },
      { file: './dist/index.d.cts', format: 'es' },
      { file: './dist/index.d.mts', format: 'es' },
      { file: './dist/browser.d.ts', format: 'es' },
      { file: './dist/browser.d.mts', format: 'es' },
    ],
    plugins: [dtsPlugin({ respectExternal: true }), selfContainedDts],
  },

  // Agent tools for the `docx-template` skill. These are CLIs, not a package:
  // they bundle `jszip`, `sax` and the library itself so that the skill runs
  // straight after `/plugin install`, with nothing to install first. The output
  // is committed to git, so it carries no sourcemaps and no hashed filenames —
  // see CONTRIBUTING.md.
  {
    input: Object.fromEntries(
      ['analyze', 'generate', 'refine', 'verify'].map(tool => [
        tool,
        `./skills/docx-template/agent/${tool}.ts`,
      ])
    ),
    output: {
      dir: './skills/docx-template/agent/dist',
      format: 'es',
      entryFileNames: '[name].mjs',
      chunkFileNames: 'shared-[name].mjs',
      sourcemap: false,
    },
    plugins: [
      node({ preferBuiltins: true }),
      commonjs(),
      esbuild({ target: 'node20' }),
    ],
  },

  // Browser: self-contained, polyfilled bundle exposed as the `./browser`
  // subpath (and as the default entry point on unpkg/jsDelivr).
  {
    input: './src/browser.ts',
    output: {
      file: './dist/browser.mjs',
      format: 'es',
      exports: 'named',
      sourcemap: true,
    },
    plugins: [
      node({
        preferBuiltins: false,
      }),
      commonjs({
        // `sax` and `jszip` probe for Node's stream support behind try/catch and
        // fall back to a no-op when the require throws — which is exactly what
        // happens in a browser. Left to its own devices the plugin resolves those
        // probes and pulls ~90 kB of unreachable Node stream machinery into the
        // bundle (and, for `stream`, emits a bare `import ... from "stream"` that
        // no browser can load). Keep the requires literal so the fallbacks win.
        ignore: id =>
          BROWSER_IGNORED_REQUIRES.has(id) ||
          id.endsWith('nodejs/NodejsStreamOutputAdapter'),
      }),
      // Set some node specific globals
      replace({
        preventAssignment: true,
        // The default `\b` delimiters would also rewrite `process` inside
        // module specifiers such as 'process-nextick-args'. Exclude quotes and
        // hyphens so only real identifier references are substituted.
        delimiters: ['(?<![\\w\'"@/-])', '(?![\\w-])'],
        values: {
          'process.env.NODE_DEBUG': false,
          'process.pid': 42,
          'process.nextTick': 'setTimeout',
          'process.stdout': 'null',
          'process.stderr': 'null',
          'process.env.READABLE_STREAM': 'false',
          'process.browser': 'true',
          'process.env.NODE_ENV': '"production"',
          process: 'undefined',
        },
      }),
      esbuild({
        target: 'es2017',
        minify: true,
      }),
      // Map modules to polyfill
      {
        name: 'module-map',
        resolveId(id) {
          return BROWSER_POLYFILLS[id] ?? null;
        },
      },
    ],
  },
]);
