/**
 * The compile pass: finding a template's commands without rendering it.
 */
import { describe, expect, it } from 'vitest';
import { compileTemplate, extractQuery } from '../template/compile';
import { resolveOptions } from '../options';
import preprocessTemplate from '../preprocessTemplate';
import { createReport } from '../index';
import { type Node } from '../types';
import { parseXml } from '../xml';
import { makeDocx, reportText } from './helpers';

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** Parses a body of paragraphs the way `createReport` does. */
const parseBody = async (
  lines: string[],
  delimiter: [string, string] = ['+++', '+++']
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
  return preprocessTemplate(await parseXml(xml), delimiter, true);
};

const rawCommands = async (lines: string[], delimiter?: [string, string]) =>
  compileTemplate(
    await parseBody(lines, delimiter),
    delimiter ?? ['+++', '+++']
  ).commands.map(c => c.raw);

describe('compileTemplate', () => {
  it('finds the commands in document order', async () => {
    expect(
      await rawCommands([
        'before +++FOR x IN xs+++ after',
        '+++$x+++',
        '+++END-FOR x+++',
      ])
    ).toEqual(['FOR x IN xs', '$x', 'END-FOR x']);
  });

  it('finds no commands in a template that has none', async () => {
    expect(await rawCommands(['just text', 'more text'])).toEqual([]);
  });

  it('handles a left/right delimiter pair', async () => {
    expect(await rawCommands(['{INS a} text {INS b}'], ['{', '}'])).toEqual([
      'INS a',
      'INS b',
    ]);
  });

  it('reports the node each command was closed in', async () => {
    const template = await parseBody(['+++INS a+++']);
    const [site] = compileTemplate(template, ['+++', '+++']).commands;
    expect(site?.node._fTextNode).toBe(true);
  });

  it('leaves the template untouched', async () => {
    // The compile pass is a read: `listCommands` and `createReport` both run it
    // over the same tree.
    const template = await parseBody([
      '+++INS a+++',
      '+++IF b+++',
      '+++END-IF+++',
    ]);
    const before = JSON.stringify(template, (key, value) =>
      key === '_parent' ? undefined : value
    );
    compileTemplate(template, ['+++', '+++']);
    compileTemplate(template, ['+++', '+++']);
    const after = JSON.stringify(template, (key, value) =>
      key === '_parent' ? undefined : value
    );
    expect(after).toEqual(before);
  });
});

describe('extractQuery', () => {
  const options = resolveOptions({});

  it('returns the payload of the QUERY command', async () => {
    const template = await parseBody(['+++QUERY some query+++', '+++INS a+++']);
    expect(extractQuery(template, options)).toEqual('some query');
  });

  it('returns undefined when there is no QUERY', async () => {
    expect(
      extractQuery(await parseBody(['+++INS a+++']), options)
    ).toBeUndefined();
  });

  it('stops at the first QUERY', async () => {
    const template = await parseBody([
      '+++QUERY first+++',
      '+++QUERY second+++',
    ]);
    expect(extractQuery(template, options)).toEqual('first');
  });
});

describe('rendering does not consume the template', () => {
  it('renders the same parsed tree twice to the same result', async () => {
    // Render state used to be written into the input tree (`_ifName` on the IF
    // and END-IF nodes), so a template could only be rendered once.
    const template = await makeDocx({
      body: [
        '+++FOR item IN items+++',
        '+++IF $item.big+++big: +++$item.name+++',
        '+++ELSE+++small: +++$item.name+++',
        '+++END-IF+++',
        '+++END-FOR item+++',
      ],
    });
    const data = {
      items: [
        { name: 'one', big: true },
        { name: 'two', big: false },
      ],
    };

    const first = await createReport({ template, data }, 'JS');
    const second = await createReport({ template, data }, 'JS');
    expect(reportText(first)).toEqual('big: one\nsmall: two');
    expect(reportText(second)).toEqual(reportText(first));
  });
});
