import { describe, it, expect } from 'vitest';
import { readFixture } from './helpers';
import { createReport, listCommands } from '../index';
import { type Node } from '../types';
import { setDebugLogSink } from '../debug';

if (process.env.DEBUG) setDebugLogSink(console.log);

// Concatenates the text of all the text nodes below the given node
const nodeText = (node: Node): string =>
  node._fTextNode ? node._text : node._children.map(nodeText).join('');

// Returns the text of the report, one line per (non-empty) paragraph
const reportText = (node: Node): string => {
  const lines: string[] = [];
  const walk = (n: Node) => {
    if (!n._fTextNode && n._tag === 'w:p') {
      const text = nodeText(n).trim();
      if (text !== '') lines.push(text);
      return;
    }
    n._children.forEach(walk);
  };
  walk(node);
  return lines.join('\n');
};

const OPERATOR_ALIASES = {
  'больше или равно': '>=',
  'меньше или равно': '<=',
  'не равно': '!==',
  больше: '>',
  меньше: '<',
  равно: '===',
  и: '&&',
  или: '||',
  ИЗ: 'IN',
};

const COMMAND_ALIASES = {
  ЕСЛИ: 'IF',
  'ИНАЧЕ ЕСЛИ': 'ELSE-IF',
  ИНАЧЕ: 'ELSE',
  'КОНЕЦ ЕСЛИ': 'END-IF',
  ДЛЯ: 'FOR',
  'КОНЕЦ ДЛЯ': 'END-FOR',
};

['noSandbox', 'sandbox'].forEach(sbStatus => {
  const noSandbox = sbStatus === 'sandbox' ? false : true;

  const render = async (
    fixture: string,
    data: any,
    options: any = {}
  ): Promise<string> => {
    const template = await readFixture(fixture);
    const report = await createReport(
      { noSandbox, template, data, ...options },
      'JS'
    );
    return reportText(report);
  };

  describe(`${sbStatus}`, () => {
    describe('operatorAliases', () => {
      it('replaces operator aliases in command expressions', async () => {
        expect(
          await render(
            'aliasOperators.docx',
            { a: 3, b: 1 },
            { operatorAliases: OPERATOR_ALIASES }
          )
        ).toEqual(
          [
            'a больше b',
            'ge: true',
            'eq: false',
            'literal: a больше b',
            'shorthand: true',
          ].join('\n')
        );
      });

      it('picks the longest matching alias', async () => {
        expect(
          await render(
            'aliasOperators.docx',
            { a: 1, b: 1 },
            { operatorAliases: OPERATOR_ALIASES }
          )
        ).toEqual(
          [
            'a не больше b',
            // `больше или равно` wins over `больше`
            'ge: true',
            'eq: true',
            'literal: a больше b',
            'shorthand: false',
          ].join('\n')
        );
      });

      it('is case-insensitive', async () => {
        // Aliases declared in upper case, used in lower case in the template
        const upperCased: { [alias: string]: string } = {};
        Object.entries(OPERATOR_ALIASES).forEach(([alias, replacement]) => {
          upperCased[alias.toUpperCase()] = replacement;
        });
        expect(
          await render(
            'aliasOperators.docx',
            { a: 3, b: 1 },
            { operatorAliases: upperCased }
          )
        ).toEqual(
          [
            'a больше b',
            'ge: true',
            'eq: false',
            'literal: a больше b',
            'shorthand: true',
          ].join('\n')
        );
      });

      it('leaves the template alone when no aliases are configured', async () => {
        const template = await readFixture('aliasOperators.docx');
        await expect(
          createReport({ noSandbox, template, data: { a: 3, b: 1 } }, 'JS')
        ).rejects.toThrow(/больше/);
      });
    });

    describe('commandAliases', () => {
      const data = {
        значение: 42,
        товары: [{ название: 'один' }, { название: 'два' }],
      };
      const options = {
        commandAliases: COMMAND_ALIASES,
        operatorAliases: OPERATOR_ALIASES,
      };

      it('supports aliased IF / ELSE-IF / ELSE / END-IF', async () => {
        expect(await render('aliasCommands.docx', data, options)).toEqual(
          ['большое', 'товар: один', 'товар: два'].join('\n')
        );
        expect(
          await render('aliasCommands.docx', { ...data, значение: 7 }, options)
        ).toEqual(['среднее', 'товар: один', 'товар: два'].join('\n'));
        expect(
          await render('aliasCommands.docx', { ...data, значение: 1 }, options)
        ).toEqual(['маленькое', 'товар: один', 'товар: два'].join('\n'));
      });

      it('supports aliased FOR / END-FOR', async () => {
        expect(
          await render(
            'aliasCommands.docx',
            { ...data, товары: [{ название: 'три' }] },
            options
          )
        ).toEqual(['большое', 'товар: три'].join('\n'));
      });

      it('rejects aliases that do not point to a built-in command', async () => {
        const template = await readFixture('aliasCommands.docx');
        await expect(
          createReport({
            noSandbox,
            template,
            data,
            commandAliases: { ЕСЛИ: 'NOPE' },
          })
        ).rejects.toThrow('Command aliases must point to one of');
      });
    });
  });
});

describe('listCommands with aliases', () => {
  it('reports the built-in command names and the de-aliased code', async () => {
    const template = await readFixture('aliasCommands.docx');
    const commands = await listCommands(template, undefined, {
      commandAliases: COMMAND_ALIASES,
      operatorAliases: OPERATOR_ALIASES,
    });
    // `raw` keeps the command as written in the template (except for the
    // command name); `code` is the JS that actually gets evaluated.
    expect(commands).toEqual([
      { raw: 'IF значение больше 10', type: 'IF', code: 'значение > 10' },
      {
        raw: 'ELSE-IF значение больше 5',
        type: 'ELSE-IF',
        code: 'значение > 5',
      },
      { raw: 'ELSE', type: 'ELSE', code: '' },
      { raw: 'END-IF', type: 'END-IF', code: '' },
      { raw: 'FOR товар ИЗ товары', type: 'FOR', code: 'товар IN товары' },
      {
        raw: 'INS $товар.название',
        type: 'INS',
        code: '$товар.название',
      },
      { raw: 'END-FOR товар', type: 'END-FOR', code: 'товар' },
    ]);
  });
});
