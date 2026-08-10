import { describe, it, expect } from 'vitest';
import { fixturePath, makeDocx } from './helpers';
import fs from 'fs';
import { listCommands } from '../main';
import { type CommandSummary } from '../types';

describe('listCommands', () => {
  it('handles simple INS', async () => {
    const template = await fs.promises.readFile(
      fixturePath('noQuerySimpleInserts.docx')
    );
    expect(await listCommands(template)).toEqual([
      { raw: 'INS a', code: 'a', type: 'INS' },
      { raw: 'ins b', code: 'b', type: 'INS' },
    ]);
  });

  it('handles INS in header and footer', async () => {
    const template = await fs.promises.readFile(
      fixturePath('insertInHeaderAndFooter.docx')
    );
    expect(await listCommands(template)).toMatchInlineSnapshot(`
      [
        {
          "code": "body_command",
          "raw": "INS body_command",
          "type": "INS",
        },
        {
          "code": "footer_command",
          "raw": "INS footer_command",
          "type": "INS",
        },
        {
          "code": "header_command",
          "raw": "INS header_command",
          "type": "INS",
        },
      ]
    `);
  });

  it('handles IMAGE', async () => {
    const template = await fs.promises.readFile(fixturePath('imagesSVG.docx'));
    expect(await listCommands(template, '+++')).toEqual([
      { raw: 'IMAGE svgImgFile()', code: 'svgImgFile()', type: 'IMAGE' },
      { raw: 'IMAGE svgImgStr()', code: 'svgImgStr()', type: 'IMAGE' },
    ]);
  });

  it('handles IMAGE in header', async () => {
    const template = await fs.promises.readFile(
      fixturePath('imageHeader.docx')
    );
    expect(await listCommands(template, '+++')).toMatchInlineSnapshot(`
      [
        {
          "code": "image()",
          "raw": "IMAGE image()",
          "type": "IMAGE",
        },
        {
          "code": "image()",
          "raw": "IMAGE image()",
          "type": "IMAGE",
        },
      ]
    `);
  });

  it('handles inline FOR loops', async () => {
    const template = await fs.promises.readFile(fixturePath('for1inline.docx'));
    expect(await listCommands(template)).toMatchInlineSnapshot(`
      [
        {
          "code": "company IN companies",
          "raw": "FOR company IN companies",
          "type": "FOR",
        },
        {
          "code": "$company.name",
          "raw": "INS $company.name",
          "type": "INS",
        },
        {
          "code": "company",
          "raw": "END-FOR company",
          "type": "END-FOR",
        },
      ]
    `);
  });

  it('handles IF clausess', async () => {
    const template = await fs.promises.readFile(fixturePath('if2.docx'));
    expect(await listCommands(template)).toMatchInlineSnapshot(`
      [
        {
          "code": "4 > 3",
          "raw": "IF 4 > 3",
          "type": "IF",
        },
        {
          "code": "true",
          "raw": "IF true",
          "type": "IF",
        },
        {
          "code": "",
          "raw": "END-IF",
          "type": "END-IF",
        },
        {
          "code": "",
          "raw": "END-IF",
          "type": "END-IF",
        },
        {
          "code": "4 > 3",
          "raw": "IF 4 > 3",
          "type": "IF",
        },
        {
          "code": "false",
          "raw": "IF false",
          "type": "IF",
        },
        {
          "code": "",
          "raw": "END-IF",
          "type": "END-IF",
        },
        {
          "code": "",
          "raw": "END-IF",
          "type": "END-IF",
        },
        {
          "code": "4 < 3",
          "raw": "IF 4 < 3",
          "type": "IF",
        },
        {
          "code": "true",
          "raw": "IF true",
          "type": "IF",
        },
        {
          "code": "",
          "raw": "END-IF",
          "type": "END-IF",
        },
        {
          "code": "",
          "raw": "END-IF",
          "type": "END-IF",
        },
      ]
    `);
  });

  it('handles IF / ELSE-IF / ELSE', async () => {
    const template = await fs.promises.readFile(fixturePath('ifElseIf.docx'));
    expect(await listCommands(template)).toEqual([
      { raw: 'IF value > 10', code: 'value > 10', type: 'IF' },
      { raw: 'ELSE-IF value > 5', code: 'value > 5', type: 'ELSE-IF' },
      { raw: 'ELSE-IF value > 0', code: 'value > 0', type: 'ELSE-IF' },
      { raw: 'ELSE', code: '', type: 'ELSE' },
      { raw: 'END-IF', code: '', type: 'END-IF' },
    ]);
  });

  it('reports every kind of command', async () => {
    const template = await makeDocx({
      body: [
        '+++QUERY the_query+++',
        '+++ALIAS shorthand INS a+++',
        '+++EXEC 1 + 1+++',
        '+++INS b+++',
        '+++IMAGE img()+++',
        '+++LINK lnk()+++',
        `+++HTML '<p>hi</p>'+++`,
      ],
    });
    const expected: CommandSummary[] = [
      { raw: 'QUERY the_query', type: 'QUERY', code: 'the_query' },
      { raw: 'ALIAS shorthand INS a', type: 'ALIAS', code: 'shorthand INS a' },
      { raw: 'EXEC 1 + 1', type: 'EXEC', code: '1 + 1' },
      { raw: 'INS b', type: 'INS', code: 'b' },
      { raw: 'IMAGE img()', type: 'IMAGE', code: 'img()' },
      { raw: 'LINK lnk()', type: 'LINK', code: 'lnk()' },
      { raw: `HTML '<p>hi</p>'`, type: 'HTML', code: `'<p>hi</p>'` },
    ];
    expect(await listCommands(template)).toEqual(expected);
  });

  it('expands the `=` and `!` shorthands', async () => {
    const template = await makeDocx({ body: ['+++=b+++', '+++!c()+++'] });
    expect(await listCommands(template)).toEqual([
      { raw: 'INS b', type: 'INS', code: 'b' },
      { raw: 'EXEC c()', type: 'EXEC', code: 'c()' },
    ]);
  });

  it('fails on a template that uses an ALIAS shorthand', async () => {
    // NOTE: this is a known defect, pinned here so that a fix is a conscious
    // decision rather than an accident. `listCommands` never runs the ALIAS
    // command, so `ctx.shorthands` stays empty and any later `*name` reference
    // is reported as unknown — even though `createReport` renders the same
    // template just fine (see the `for1alias.docx` test in templating.test.ts).
    const template = await makeDocx({
      body: ['+++ALIAS x INS a+++', '+++*x+++'],
    });
    await expect(listCommands(template)).rejects.toThrow('Unknown alias: *x');
    await expect(
      listCommands(await fs.promises.readFile(fixturePath('for1alias.docx')))
    ).rejects.toThrow('Unknown alias');
  });

  it('handles a left/right delimiter pair', async () => {
    const template = await makeDocx({ body: ['{INS a}', 'text', '{INS b}'] });
    expect(await listCommands(template, ['{', '}'])).toEqual([
      { raw: 'INS a', type: 'INS', code: 'a' },
      { raw: 'INS b', type: 'INS', code: 'b' },
    ]);
  });

  it('returns an empty list for a template without commands', async () => {
    const template = await makeDocx({ body: ['just some text'] });
    expect(await listCommands(template)).toEqual([]);
  });

  it('handles custom delimiter', async () => {
    const template = await fs.promises.readFile(
      fixturePath('for1customDelimiter.docx')
    );
    expect(await listCommands(template, '***')).toMatchInlineSnapshot(`
      [
        {
          "code": "company IN companies",
          "raw": "FOR company IN companies",
          "type": "FOR",
        },
        {
          "code": "$company.name",
          "raw": "INS $company.name",
          "type": "INS",
        },
        {
          "code": "company",
          "raw": "END-FOR company",
          "type": "END-FOR",
        },
      ]
    `);
  });
});
