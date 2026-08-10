import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import { createReport, listCommands } from '../../../../src/index';
import { fixturePath, readFixture } from '../../../../src/__tests__/helpers';
import { type NonTextNode } from '../../../../src/types';
import { analyze } from '../analyze';
import { aliasesFromCommandNames, applyMapping } from '../generate';
import { applyModifications } from '../refine';
import { checkCommands } from '../verify';
import {
  directChildren,
  findAll,
  firstChild,
  newElement,
  readDocParts,
  replaceText,
  replaceTextAll,
  textOf,
  topLevelTables,
  writeDocParts,
} from '../docxXml';

const DELIMITER: [string, string] = ['+++', '+++'];

let tmpDir: string;
const tmpFile = (name: string) => path.join(tmpDir, name);

beforeAll(() => {
  tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'docx-template-skill-'));
});

afterAll(() => {
  fs.rmSync(tmpDir, { recursive: true, force: true });
});

/**
 * Renders a template fixture and saves the result, giving us a realistic
 * "filled-out document" to feed back into `generate` — which is exactly the
 * situation the skill exists for.
 */
async function renderToFile(
  fixture: string,
  data: unknown,
  name: string
): Promise<string> {
  const report = await createReport({
    template: await readFixture(fixture),
    data,
    cmdDelimiter: DELIMITER,
  });
  const filePath = tmpFile(name);
  fs.writeFileSync(filePath, report);
  return filePath;
}

/** Applies a mapping or modification list to a docx and writes the result. */
async function edit(
  inputPath: string,
  outputName: string,
  apply: (parts: Awaited<ReturnType<typeof readDocParts>>['parts']) => void
): Promise<string> {
  const { zip, parts } = await readDocParts(inputPath);
  apply(parts);
  const outputPath = tmpFile(outputName);
  await writeDocParts(zip, parts, outputPath);
  return outputPath;
}

// ==========================================

describe('analyze', () => {
  it('reports the commands and structure of a template', async () => {
    const analysis = await analyze(fixturePath('for1.docx'), DELIMITER, {});

    expect(analysis.commands.map(c => c.type)).toEqual([
      'FOR',
      'INS',
      'END-FOR',
    ]);
    expect(analysis.plainText).toContain('+++FOR company IN companies+++');
    expect(analysis.main.paragraphs.map(p => p.text)).toContain(
      '+++INS $company.name+++'
    );
  });

  it('reassembles commands that Word split across runs', async () => {
    const analysis = await analyze(fixturePath('for1.docx'), DELIMITER, {});

    // The fixture stores this command in three separate runs; the paragraph
    // text has to read as one command anyway, or nothing downstream can match.
    const paragraph = analysis.main.paragraphs.find(p =>
      p.text.includes('$company.name')
    );
    expect(paragraph?.runs.length).toBeGreaterThan(1);
    expect(paragraph?.text).toBe('+++INS $company.name+++');
  });

  it('extracts table rows and cells', async () => {
    const analysis = await analyze(fixturePath('for-row1.docx'), DELIMITER, {});

    const table = analysis.main.tables[0];
    expect(table).toBeDefined();
    expect(table?.rows.map(r => r.cells[0]?.text)).toEqual([
      'Name',
      '+++FOR company IN companies+++',
      '+++INS $company.name+++',
      '+++END-FOR company+++',
    ]);
  });

  it('still reads a template whose commands do not balance', async () => {
    // A half-written template has to stay readable — that is exactly when the
    // agent most needs to see what is in the document.
    const analysis = await analyze(
      fixturePath('missingEndFor.docx'),
      DELIMITER,
      {}
    );

    expect(analysis.commandsError).toBeUndefined();
    expect(analysis.plainText).not.toBe('');
    expect(checkCommands(analysis.commands)).toEqual([
      'FOR a is never closed by an END-FOR',
    ]);
  });
});

// ==========================================

describe('generate', () => {
  it('turns a filled-out table back into a working loop', async () => {
    const companies = [
      { name: 'Acme Corp' },
      { name: 'Globex' },
      { name: 'Initech' },
    ];
    const filled = await renderToFile(
      'for-row1.docx',
      { companies },
      'filled.docx'
    );

    const template = await edit(filled, 'table_template.docx', parts => {
      applyMapping(parts, {
        tableLoops: [
          {
            var: 'company',
            over: 'companies',
            tableIndex: 0,
            startRow: 1,
            fields: { 0: '$company.name' },
          },
        ],
      });
    });

    expect(
      (await listCommands(fs.readFileSync(template), DELIMITER)).map(c => ({
        type: c.type,
        code: c.code,
      }))
    ).toEqual([
      { type: 'FOR', code: 'company IN companies' },
      { type: 'INS', code: '$company.name' },
      { type: 'END-FOR', code: 'company' },
    ]);

    // The real test: rendering the regenerated template reproduces the
    // document it was derived from.
    const report = await createReport({
      template: fs.readFileSync(template),
      data: { companies },
      cmdDelimiter: DELIMITER,
    });
    const roundTrip = tmpFile('roundtrip.docx');
    fs.writeFileSync(roundTrip, report);
    const analysis = await analyze(roundTrip, DELIMITER, {});
    expect(analysis.plainText).toBe('Name\nAcme Corp\nGlobex\nInitech\n');
  });

  it('puts FOR and END-FOR in rows of their own', async () => {
    const filled = await renderToFile(
      'for-row1.docx',
      { companies: [{ name: 'Acme Corp' }, { name: 'Globex' }] },
      'filled-rows.docx'
    );

    const template = await edit(filled, 'rows_template.docx', parts => {
      applyMapping(parts, {
        tableLoops: [
          {
            var: 'company',
            over: 'companies',
            tableIndex: 0,
            fields: { 0: '$company.name' },
          },
        ],
      });
    });

    const { parts } = await readDocParts(template);
    const table = topLevelTables(parts[0]!.root)[0];
    expect(table).toBeDefined();
    const rowTexts = directChildren(table as NonTextNode, 'w:tr').map(textOf);

    // Three rows collapse to one, wrapped by a command row on either side —
    // this library deletes rows that hold nothing but commands, which is what
    // makes the markers disappear from the report.
    expect(rowTexts).toEqual([
      'Name',
      '+++FOR company IN companies+++',
      '+++INS $company.name+++',
      '+++END-FOR company+++',
    ]);
  });

  it('wraps paragraph blocks in IF and substitutes variables', async () => {
    const filled = await renderToFile(
      'for1.docx',
      { companies: [{ name: 'Acme Corp' }, { name: 'Globex' }] },
      'filled-paras.docx'
    );

    const template = await edit(filled, 'paras_template.docx', parts => {
      applyMapping(parts, {
        conditionals: [{ expr: 'showPartner', paragraphText: 'Globex' }],
        variables: { 'Acme Corp': 'client.name', Globex: 'partner.name' },
      });
    });

    expect(
      (await listCommands(fs.readFileSync(template), DELIMITER)).map(
        c => c.type
      )
    ).toEqual(['INS', 'IF', 'INS', 'END-IF']);

    const report = await createReport({
      template: fs.readFileSync(template),
      data: {
        client: { name: 'Umbrella Ltd' },
        partner: { name: 'Soylent' },
        showPartner: false,
      },
      cmdDelimiter: DELIMITER,
    });
    const rendered = tmpFile('paras_report.docx');
    fs.writeFileSync(rendered, report);
    const analysis = await analyze(rendered, DELIMITER, {});
    expect(analysis.plainText).toContain('Umbrella Ltd');
    expect(analysis.plainText).not.toContain('Soylent');
  });

  it('refuses a variable it cannot find, rather than silently skipping it', async () => {
    const { parts } = await readDocParts(fixturePath('for1.docx'));
    expect(() =>
      applyMapping(parts, { variables: { 'not in the document': 'x' } })
    ).toThrow(/not found in the document/);
  });

  it('writes localized commands, and the aliases that decode them', async () => {
    const filled = await renderToFile(
      'for-row1.docx',
      { companies: [{ name: 'Acme Corp' }, { name: 'Globex' }] },
      'filled-ru.docx'
    );
    const mapping = {
      commandNames: {
        INS: '=',
        FOR: 'ДЛЯ',
        IN: 'ИЗ',
        'END-FOR': 'КОНЕЦ ДЛЯ',
      },
      tableLoops: [
        {
          var: 'компания',
          over: 'компании',
          tableIndex: 0,
          fields: { 0: '$компания.название' },
        },
      ],
    };

    const template = await edit(filled, 'ru_template.docx', parts => {
      applyMapping(parts, mapping);
    });

    // `=` is a built-in shorthand and `ДЛЯ`'s `IN` keyword is an operator, not
    // a command — the derived aliases have to reflect both.
    const aliases = aliasesFromCommandNames(mapping);
    expect(aliases).toEqual({
      commandAliases: { ДЛЯ: 'FOR', 'КОНЕЦ ДЛЯ': 'END-FOR' },
      operatorAliases: { ИЗ: 'IN' },
    });

    // Without the aliases the engine cannot tell the commands apart...
    const template2 = fs.readFileSync(template);
    expect((await listCommands(template2, DELIMITER)).map(c => c.type)).toEqual(
      ['INS', 'INS', 'INS']
    );
    // ...and with them, it renders.
    expect(
      (await listCommands(template2, DELIMITER, aliases)).map(c => c.type)
    ).toEqual(['FOR', 'INS', 'END-FOR']);

    const report = await createReport({
      template: template2,
      data: { компании: [{ название: 'Ромашка' }, { название: 'Восход' }] },
      cmdDelimiter: DELIMITER,
      ...aliases,
    });
    const rendered = tmpFile('ru_report.docx');
    fs.writeFileSync(rendered, report);
    expect((await analyze(rendered, DELIMITER, {})).plainText).toBe(
      'Name\nРомашка\nВосход\n'
    );
  });

  it('reports a table index that does not exist', async () => {
    const { parts } = await readDocParts(fixturePath('for1.docx'));
    expect(() =>
      applyMapping(parts, {
        tableLoops: [
          { var: 'x', over: 'xs', tableIndex: 3, fields: { 0: '$x' } },
        ],
      })
    ).toThrow(/no table with index 3/);
  });
});

// ==========================================

describe('style-preserving replacement', () => {
  /** `<w:r><w:rPr>…</w:rPr><w:t>text</w:t></w:r>`, optionally bold. */
  const run = (text: string, bold = false): NonTextNode =>
    newElement('w:r', {}, [
      ...(bold ? [newElement('w:rPr', {}, [newElement('w:b')])] : []),
      newElement('w:t', {}, [{ _fTextNode: true, _text: text, _children: [] }]),
    ]);

  const paragraph = (...runs: NonTextNode[]) => newElement('w:p', {}, runs);

  it('leaves the formatting of untouched runs alone', () => {
    const p = paragraph(run('Dear ', true), run('Acme Corp'), run(', hello'));

    expect(replaceText(p, 'Acme Corp', '+++INS client.name+++')).toBe(true);
    expect(textOf(p)).toBe('Dear +++INS client.name+++, hello');
    // The bold run is still bold, and still a separate run.
    const runs = directChildren(p, 'w:r');
    expect(findAll(runs[0]!, 'w:b')).toHaveLength(1);
    expect(textOf(runs[0]!)).toBe('Dear ');
  });

  it('handles a match split across several runs', () => {
    const p = paragraph(run('Acme'), run(' '), run('Corp'), run(' Ltd'));

    expect(replaceText(p, 'Acme Corp', '{X}')).toBe(true);
    expect(textOf(p)).toBe('{X} Ltd');
  });

  it('replaces every occurrence without disturbing the offsets', () => {
    const p = paragraph(run('a-b-'), run('a-b'));

    expect(replaceTextAll(p, 'a-b', 'Z')).toBe(2);
    expect(textOf(p)).toBe('Z-Z');
  });
});

// ==========================================

describe('refine', () => {
  it('renames an expression and swaps a whole command', async () => {
    const refined = await edit(
      fixturePath('for1.docx'),
      'refined.docx',
      parts => {
        applyModifications(parts, {
          modifications: [
            {
              type: 'renameExpression',
              from: '$company.name',
              to: '$company.legalName',
            },
            {
              type: 'replaceCommand',
              from: 'FOR company IN companies',
              to: 'FOR company IN companies.filter(c => c.active)',
            },
          ],
        });
      }
    );

    expect(
      (await listCommands(fs.readFileSync(refined), DELIMITER)).map(c => c.code)
    ).toEqual([
      'company IN companies.filter(c => c.active)',
      '$company.legalName',
      'company',
    ]);
  });

  it('rewrites commands that are split across runs', async () => {
    // The fixture stores `+++INS $company.name+++` in three runs; a naive
    // per-run edit would never see the whole expression.
    const refined = await edit(
      fixturePath('for1.docx'),
      'refined-split.docx',
      parts => {
        applyModifications(parts, {
          modifications: [
            { type: 'renameExpression', from: 'company.name', to: 'co.title' },
          ],
        });
      }
    );

    const commands = await listCommands(fs.readFileSync(refined), DELIMITER);
    expect(commands.map(c => c.code)).toContain('$co.title');
  });

  it('removes a command and leaves literal text behind', async () => {
    const refined = await edit(
      fixturePath('for1.docx'),
      'refined-removed.docx',
      parts => {
        applyModifications(parts, {
          modifications: [
            {
              type: 'removeCommand',
              code: 'INS $company.name',
              replaceWith: 'redacted',
            },
          ],
        });
      }
    );

    const analysis = await analyze(refined, DELIMITER, {});
    expect(analysis.plainText).toContain('redacted');
    expect(analysis.commands.map(c => c.type)).toEqual(['FOR', 'END-FOR']);
  });

  it('merges floating tables into one inline table', () => {
    // No fixture in this repo uses `w:tblpPr`, so the document is built by
    // hand: two absolutely-positioned single-cell tables with an empty
    // paragraph between them — Word's usual signature-block layout.
    const floatingTable = (text: string, width: string) =>
      newElement('w:tbl', {}, [
        newElement('w:tblPr', {}, [newElement('w:tblpPr', { 'w:tblpX': '1' })]),
        newElement('w:tblGrid', {}, [
          newElement('w:gridCol', { 'w:w': width }),
        ]),
        newElement('w:tr', {}, [
          newElement('w:tc', {}, [
            newElement('w:p', {}, [
              newElement('w:r', {}, [
                newElement('w:t', {}, [
                  { _fTextNode: true, _text: text, _children: [] },
                ]),
              ]),
            ]),
          ]),
        ]),
      ]);

    const body = newElement('w:body', {}, [
      floatingTable('Signed for the buyer', '4000'),
      newElement('w:p'),
      floatingTable('Signed for the seller', '5000'),
    ]);
    const root = newElement('w:document', {}, [body]);

    applyModifications(
      [{ filename: 'word/document.xml', root, isMain: true }],
      {
        modifications: [{ type: 'mergeFloatingTables' }],
      }
    );

    const tables = topLevelTables(root);
    expect(tables).toHaveLength(1);
    const merged = tables[0]!;

    // The two tables' cells now sit side by side in one row...
    const rows = directChildren(merged, 'w:tr');
    expect(rows).toHaveLength(1);
    expect(directChildren(rows[0]!, 'w:tc').map(textOf)).toEqual([
      'Signed for the buyer',
      'Signed for the seller',
    ]);

    // ...with both column widths carried over, no absolute positioning left,
    // and no borders to give away the seam.
    const grid = firstChild(merged, 'w:tblGrid')!;
    expect(directChildren(grid, 'w:gridCol')).toHaveLength(2);
    expect(findAll(merged, 'w:tblpPr')).toHaveLength(0);
    expect(findAll(merged, 'w:tblBorders')).toHaveLength(1);

    // The empty paragraph that sat between them is gone.
    expect(directChildren(body, 'w:p')).toHaveLength(0);
  });
});

// ==========================================

describe('verify', () => {
  const commandsOf = async (fixture: string) =>
    listCommands(await readFixture(fixture), DELIMITER);

  it('accepts a balanced template', async () => {
    expect(checkCommands(await commandsOf('for1.docx'))).toEqual([]);
  });

  it('catches an ELSE with no IF around it', async () => {
    expect(checkCommands(await commandsOf('elseOutsideIf.docx'))).toEqual([
      'ELSE appears outside an IF … END-IF block',
    ]);
  });

  it('catches a loop that is never closed', () => {
    expect(
      checkCommands([
        { type: 'FOR', code: 'p IN people', raw: 'FOR p IN people' },
        { type: 'INS', code: '$p.name', raw: 'INS $p.name' },
      ])
    ).toEqual(['FOR p is never closed by an END-FOR']);
  });

  it('catches an END-FOR with no loop to close', async () => {
    expect(checkCommands(await commandsOf('unmatchedEndFor.docx'))).toEqual([
      'END-FOR closes a loop that was never opened',
    ]);
  });

  it('catches an END-FOR that closes the wrong loop', () => {
    expect(
      checkCommands([
        { type: 'FOR', code: 'p IN people', raw: 'FOR p IN people' },
        { type: 'END-FOR', code: 'person', raw: 'END-FOR person' },
      ])
    ).toEqual([
      'END-FOR person closes FOR p — the variable names differ',
      // The mismatched END-FOR still popped the loop, so nothing is left open.
    ]);
  });

  it('flags the curly quotes Word introduces', async () => {
    const problems = checkCommands(await commandsOf('fixSmartQuotes.docx'));
    expect(problems.join('\n')).toMatch(/curly quote/);
  });

  it('collects every template error in one pass', async () => {
    // `failFast: false` is what makes verify able to report a whole broken
    // template at once instead of one error per run.
    const error = await createReport({
      template: await readFixture('commandExecutionError.docx'),
      data: {},
      cmdDelimiter: DELIMITER,
      failFast: false,
    }).catch((err: unknown) => err);

    expect(Array.isArray(error)).toBe(true);
  });
});
