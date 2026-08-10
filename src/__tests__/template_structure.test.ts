/**
 * The FOR/IF nesting worked out from a template's commands, without rendering.
 *
 * The last test is the important one: it runs the analysis over every fixture
 * and lines the problems it finds up against the errors the engine actually
 * reports. That comparison is the evidence for why the structural diagnosis is
 * an observer rather than a replacement — see `analyzeStructure`.
 */
import fs from 'node:fs';
import { describe, expect, it } from 'vitest';
import { createReport } from '../index';
import { isError } from '../errors';
import { resolveOptions } from '../options';
import preprocessTemplate from '../preprocessTemplate';
import { compileTemplate } from '../template/compile';
import { analyzeStructure } from '../template/structure';
import { Command, type CreateReportOptions, type Node } from '../types';
import { parseXml } from '../xml';
import { caseName, CORPUS } from './corpus';
import { fixturePath } from './helpers';

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const parseBody = async (
  lines: string[],
  options: CreateReportOptions
): Promise<Node> => {
  const xml =
    `<w:document ${W_NS}><w:body>` +
    lines
      .map(
        line =>
          `<w:p><w:r><w:t>${line
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')}</w:t></w:r></w:p>`
      )
      .join('') +
    `</w:body></w:document>`;
  return preprocessTemplate(
    await parseXml(xml),
    options.cmdDelimiter,
    options.preserveSpace
  );
};

const analyze = async (lines: string[]) => {
  const options = resolveOptions({});
  const template = await parseBody(lines, options);
  return analyzeStructure(
    compileTemplate(template, options.cmdDelimiter),
    options
  );
};

describe('analyzeStructure', () => {
  it('pairs a FOR with its END-FOR', async () => {
    const { constructs, problems } = await analyze([
      '+++FOR item IN items+++',
      '+++$item+++',
      '+++END-FOR item+++',
    ]);
    expect(problems).toEqual([]);
    expect(constructs).toEqual([
      {
        kind: Command.FOR,
        open: 0,
        close: 2,
        branches: [],
        depth: 0,
        varName: 'item',
      },
    ]);
  });

  it('records the branches of an IF construct', async () => {
    const { constructs } = await analyze([
      '+++IF a+++',
      '+++ELSE-IF b+++',
      '+++ELSE+++',
      '+++END-IF+++',
    ]);
    expect(constructs[0]).toMatchObject({
      kind: Command.IF,
      open: 0,
      branches: [1, 2],
      close: 3,
    });
  });

  it('records the nesting depth', async () => {
    const { constructs } = await analyze([
      '+++FOR a IN as+++',
      '+++FOR b IN bs+++',
      '+++END-FOR b+++',
      '+++END-FOR a+++',
    ]);
    expect(constructs.map(c => c.depth)).toEqual([0, 1]);
  });

  it('reports a FOR that is never closed', async () => {
    const { problems } = await analyze(['+++FOR a IN as+++', '+++$a+++']);
    expect(problems).toEqual([
      { id: 'unclosed_for', at: 0, command: 'FOR a IN as' },
    ]);
  });

  it('reports an END-IF with no IF', async () => {
    const { problems } = await analyze(['+++END-IF+++']);
    expect(problems).toEqual([
      { id: 'unexpected_end', at: 0, command: 'END-IF' },
    ]);
  });

  it('reports an ELSE outside an IF', async () => {
    const { problems } = await analyze(['+++ELSE+++']);
    expect(problems).toEqual([
      { id: 'else_outside_if', at: 0, command: 'ELSE' },
    ]);
  });

  it('reports a second ELSE', async () => {
    const { problems } = await analyze([
      '+++IF a+++',
      '+++ELSE+++',
      '+++ELSE+++',
      '+++END-IF+++',
    ]);
    expect(problems.map(p => p.id)).toEqual(['else_after_else']);
  });

  it('ignores an END-FOR that names no open loop', async () => {
    // Not an error, and deliberately so: a construct spanning several table
    // cells legitimately meets the END-FOR of an earlier part of the same row.
    const { problems } = await analyze([
      '+++IF a+++',
      '+++END-FOR x+++',
      '+++END-IF+++',
    ]);
    expect(problems).toEqual([]);
  });

  it('reports an END-FOR that names an enclosing loop, not the innermost', async () => {
    const { problems } = await analyze([
      '+++FOR company IN companies+++',
      '+++FOR person IN persons+++',
      '+++END-FOR company+++',
      '+++END-FOR company+++',
    ]);
    expect(problems.map(p => p.id)).toEqual([
      'unexpected_end',
      'unexpected_end',
      'unclosed_for',
      'unclosed_for',
    ]);
  });

  it('survives a shorthand it cannot resolve', async () => {
    // `*name` needs the ALIAS that defines it to have run, which only happens
    // during a render. Analysis must not throw over it.
    const { problems, constructs } = await analyze([
      '+++ALIAS name INS $x+++',
      '+++*name+++',
    ]);
    expect(problems).toEqual([]);
    expect(constructs).toEqual([]);
  });
});

describe('structural diagnosis versus what the engine reports', () => {
  /** The errors `createReport` produces for a corpus case, in order. */
  const engineErrors = async (file: string, options: object) => {
    try {
      await createReport({
        ...options,
        template: fs.readFileSync(fixturePath(file)),
      } as never);
      return [];
    } catch (err) {
      const list = Array.isArray(err) ? err : [err];
      return list.map(e => (isError(e) ? e.message : String(e)));
    }
  };

  it('finds structural problems only where the engine also fails', async () => {
    // A template the analysis is unhappy about had better not be one that
    // renders cleanly today — that would mean the analysis is wrong about what
    // this template language allows.
    const falsePositives: string[] = [];
    const flagged: string[] = [];
    for (const corpusCase of CORPUS) {
      const options = resolveOptions(corpusCase.options);
      const template = preprocessTemplate(
        await parseXml(await readMainDocument(fixturePath(corpusCase.file))),
        options.cmdDelimiter,
        options.preserveSpace
      );
      const { problems } = analyzeStructure(
        compileTemplate(template, options.cmdDelimiter),
        options
      );
      if (problems.length === 0) continue;
      flagged.push(caseName(corpusCase));
      const errors = await engineErrors(corpusCase.file, corpusCase.options);
      if (errors.length === 0) falsePositives.push(caseName(corpusCase));
    }
    expect(falsePositives).toEqual([]);
    // Listed, so that the check above cannot pass by finding nothing at all.
    expect(flagged.sort()).toMatchInlineSnapshot(`
      [
        "elseOutsideIf.docx",
        "ifDoubleElse.docx",
        "invalidFor.docx",
        "invalidForCmd.docx",
        "invalidIf.docx",
        "invalidMultipleErrors.docx",
        "missingEndFor.docx",
        "missingEndIf.docx",
        "unmatchedEndFor.docx",
        "unmatchedEndIf.docx",
      ]
    `);
  }, 60_000);

  it('cannot replace the engine: the errors arrive in a different order', async () => {
    // This is why the structural diagnosis does not throw. With
    // `failFast: false` the engine reports errors in the order the walk meets
    // them, and the array it throws is part of the contract. Here the two
    // expression failures come *before* the structural ones, because that is
    // where they sit in the document — a diagnosis made up front would put the
    // structural problems first.
    const errors = await engineErrors('invalidMultipleErrors.docx', {
      data: { companies: [{ name: 'FIRST' }] },
      failFast: false,
    });
    expect(errors).toEqual([
      "Error executing command 'notavailable': ReferenceError: notavailable is not defined",
      "Error executing command 'something': ReferenceError: something is not defined",
      'Invalid command: END-FOR company',
      "Unterminated FOR-loop ('FOR company'). Make sure each FOR loop has a corresponding END-FOR command.",
    ]);

    const options = resolveOptions({ failFast: false });
    const template = preprocessTemplate(
      await parseXml(
        await readMainDocument(fixturePath('invalidMultipleErrors.docx'))
      ),
      options.cmdDelimiter,
      options.preserveSpace
    );
    const { problems } = analyzeStructure(
      compileTemplate(template, options.cmdDelimiter),
      options
    );
    // The structure alone knows nothing about the two runtime failures, so it
    // could only ever produce a prefix of that array, never the array itself.
    expect(problems.map(p => p.id)).toEqual(['unexpected_end', 'unclosed_for']);
  });
});

/** Reads `word/document.xml` out of a .docx on disk. */
async function readMainDocument(path: string): Promise<string> {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(fs.readFileSync(path));
  const main = zip.file('word/document.xml') ?? zip.file('word/document2.xml');
  if (main == null) throw new Error(`no main document in ${path}`);
  return main.async('text');
}
