/**
 * The golden corpus.
 *
 * Renders every fixture end to end and snapshots the *whole package*: the zip
 * entry order, a `path -> sha256` map of every entry, and the full text of the
 * parts that carry the report. The `_probe` snapshots elsewhere only ever look
 * at `word/document.xml`; this is what watches the headers, footers, `.rels`
 * parts, `[Content_Types].xml` and `word/media/`.
 *
 * If a change here is intentional, `npm test -- -u` and *read the diff*.
 */
import crypto from 'node:crypto';
import fs from 'node:fs';
import JSZip from 'jszip';
import MockDate from 'mockdate';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { createReport, listCommands } from '../index';
import { isError } from '../errors';
import { caseName, CORPUS, uncoveredFixtures } from './corpus';
import { fixturePath } from './helpers';

// Zip entries carry a DOS timestamp; without a fixed clock the bytes differ
// between runs.
beforeEach(() => MockDate.set('1/1/2000'));
afterEach(() => MockDate.reset());

const sha256 = (data: Uint8Array) =>
  crypto.createHash('sha256').update(data).digest('hex');

/**
 * The parts whose text is spelled out in full, because they are the ones a
 * regression shows up in and a hash would say nothing useful about. Everything
 * else in the package is still covered, by its hash.
 */
const isContentPart = (path: string) =>
  path === '[Content_Types].xml' ||
  /^word\/_rels\/[^/]+\.rels$/.test(path) ||
  /^word\/(document|header|footer)\d*\.xml$/.test(path);

/** A readable, diffable description of a generated .docx. */
async function describePackage(report: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(report);
  const paths = Object.keys(zip.files);

  const lines: string[] = [];
  lines.push('# zip entries, in the order they are stored');
  for (const path of paths) lines.push(`  ${path}`);

  lines.push('', '# path -> sha256');
  for (const path of [...paths].sort()) {
    const file = zip.files[path];
    if (file == null || file.dir) continue;
    const bytes = await file.async('uint8array');
    lines.push(`  ${path}  ${sha256(bytes)}  (${bytes.length} bytes)`);
  }

  for (const path of [...paths].sort()) {
    if (!isContentPart(path)) continue;
    const file = zip.files[path];
    if (file == null || file.dir) continue;
    lines.push('', `# ${path}`, await file.async('text'));
  }
  return lines.join('\n');
}

/** The same, for a case that is expected to fail. */
const describeErrors = (err: unknown): string => {
  const list = Array.isArray(err) ? err : [err];
  return [
    '# throws',
    ...list.map(e =>
      isError(e) ? `  ${e.name}: ${e.message}` : `  (not an Error) ${String(e)}`
    ),
  ].join('\n');
};

describe('golden corpus', () => {
  it('covers every fixture', () => {
    expect(uncoveredFixtures()).toEqual([]);
  });

  // Both sandbox modes: the report must not depend on how the JS was evaluated.
  for (const noSandbox of [false, true]) {
    describe(noSandbox ? 'noSandbox' : 'sandbox', () => {
      for (const corpusCase of CORPUS) {
        const { file, options } = corpusCase;
        it(caseName(corpusCase), async () => {
          const template = fs.readFileSync(fixturePath(file));
          let described: string;
          try {
            const report = await createReport({
              ...options,
              noSandbox,
              template,
            });
            described = await describePackage(report);
          } catch (err) {
            described = describeErrors(err);
          }
          // One snapshot file per sandbox mode would double the corpus for no
          // gain, so the two modes share it: they must agree byte for byte.
          await expect(described).toMatchFileSnapshot(
            `./__corpus__/${caseName(corpusCase)}.txt`
          );
        });
      }
    });
  }
});

describe('golden listCommands', () => {
  for (const corpusCase of CORPUS) {
    const { file, options } = corpusCase;
    it(caseName(corpusCase), async () => {
      const template = fs.readFileSync(fixturePath(file));
      let described: string;
      try {
        // Spread rather than list the two keys: under
        // `exactOptionalPropertyTypes`, passing them as explicitly `undefined`
        // is not the same as leaving them out, and `listCommands` wants the
        // latter.
        const commands = await listCommands(template, options.cmdDelimiter, {
          ...(options.commandAliases
            ? { commandAliases: options.commandAliases }
            : {}),
          ...(options.operatorAliases
            ? { operatorAliases: options.operatorAliases }
            : {}),
        });
        described = commands
          .map(
            c =>
              `${c.type}\t${JSON.stringify(c.code)}\t${JSON.stringify(c.raw)}`
          )
          .join('\n');
      } catch (err) {
        described = describeErrors(err);
      }
      await expect(described).toMatchFileSnapshot(
        `./__corpus__/${caseName(corpusCase)}.commands.txt`
      );
    });
  }
});
