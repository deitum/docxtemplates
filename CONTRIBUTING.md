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

## How the source is laid out

Rendering a report is a pipeline: unzip the package, parse each XML part, walk
it while executing the commands it contains, and zip the result back up.

| Module                                                                      | Responsibility                                                            |
| --------------------------------------------------------------------------- | ------------------------------------------------------------------------- |
| `index.ts`                                                                  | The public API, and the only file that decides what is public             |
| `main.ts`                                                                   | Orchestration: `createReport` and `listCommands`                          |
| `docx/parts.ts`                                                             | Reading the package: which parts are templates, and `[Content_Types].xml` |
| `docx/relationships.ts`                                                     | Writing embedded images/HTML and the `.rels` entries that point at them   |
| `docx/metadata.ts`                                                          | `getMetadata`                                                             |
| `preprocessTemplate.ts`                                                     | Joins commands that Word split across several text nodes                  |
| `walk.ts`                                                                   | The walk itself: moves over the input tree and builds the output tree     |
| `commands/parse.ts`                                                         | Turns the raw text between delimiters into a command name and its payload |
| `commands/execute.ts`                                                       | Executes one command; control flow drives the loop stack `walk.ts` reads  |
| `commands/media.ts`                                                         | The OOXML that IMAGE, LINK and HTML insert                                |
| `jsSandbox.ts`                                                              | Evaluates the JS in a command                                             |
| `context.ts`, `options.ts`                                                  | Per-part mutable state, and the defaults behind every option              |
| `ooxml.ts`                                                                  | The OOXML vocabulary: tags, attributes, namespaces, paths, media types    |
| `types.ts`                                                                  | Shared types, the `Command` enum and the built-in command list            |
| `xml.ts`, `zip.ts`, `reportUtils.ts`, `aliases.ts`, `errors.ts`, `debug.ts` | Supporting layers                                                         |

Two conventions worth knowing:

- **No bare OOXML strings.** Tags, attributes and package paths come from
  `ooxml.ts`; a typo there fails the build instead of producing a document Word
  refuses to open.
- **Debug logging is off unless a sink is installed** (`setDebugLogSink`).
  Hot paths check `logger.enabled` before building a message, because
  serializing a node once per node is expensive on a large document.

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
