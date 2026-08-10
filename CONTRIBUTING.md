# Contributing

Thanks for taking the time to contribute!

## Getting set up

The project uses **npm** and requires **Node.js 20 or later** (see `.nvmrc`).

```
$ npm ci
$ npm test
```

## Everyday commands

| Command                 | What it does                                                 |
| ----------------------- | ------------------------------------------------------------ |
| `npm test`              | Runs the test suite once                                     |
| `npm run test:watch`    | Runs the tests in watch mode                                 |
| `npm run test:coverage` | Runs the tests and writes a coverage report to `coverage/`   |
| `npm run typecheck`     | Type-checks the whole project (`tsc --noEmit`)               |
| `npm run lint`          | Lints with ESLint (`npm run lint:fix` to autofix)            |
| `npm run format`        | Formats with Prettier (`npm run format:check` to only check) |
| `npm run build`         | Builds `dist/` and the skill's agent tools with Rollup       |
| `npm run check:package` | Validates the published package (`publint` + `attw`)         |
| `npm run verify`        | Everything above, in the same order CI runs it               |

Run `npm run verify` before opening a pull request.

## Tests

Tests live in `src/__tests__` and run on [Vitest](https://vitest.dev). Most of them
render a fixture template and compare the result against a snapshot.

- Use the `fixturePath` / `readFixture` helpers from `src/__tests__/helpers.ts`
  instead of building paths by hand.
- If a change intentionally alters rendered output, update the snapshots with
  `npx vitest run -u` and **review the resulting diff** — a snapshot diff is the
  main signal that the templating engine changed behaviour.

## The `docx-template` skill

`skills/docx-template/` is a Claude Code plugin that generates templates from
filled-out documents. Its four agent tools are written in TypeScript under
`skills/docx-template/agent/` and bundled by the same Rollup config as the
library, into `agent/dist/*.mjs`.

**Those bundles are committed to git**, so that the plugin runs straight after
installation with nothing to `npm install`. After changing anything under
`agent/`, run `npm run build` and commit the regenerated `dist/` files — CI
fails if they are out of date.

Its tests live in `skills/docx-template/agent/__tests__/` and run as part of
`npm test`.

## Changesets

Releases are automated with [changesets](https://github.com/changesets/changesets).
Any pull request that changes runtime behaviour, the public API or the published
package needs a changeset:

```
$ npm run changeset
```

Pick the appropriate bump (`patch` / `minor` / `major`), describe the change from
a user's point of view, and commit the generated file in `.changeset/`.
Documentation-only or internal-tooling changes don't need one.

Merging to `master` opens (or updates) a "Version Packages" pull request; merging
_that_ publishes to npm with provenance. See
[`.github/workflows/README.md`](./.github/workflows/README.md) for how CI and the
release pipeline are wired together.
