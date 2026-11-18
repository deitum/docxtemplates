/**
 * `transformCode`, the hook that rewrites a command's JavaScript before it runs.
 *
 * What is pinned down here: which code it sees (every snippet, after the
 * aliases), that the rewrite reaches every sandbox, that it is not redone on
 * each iteration of a loop, and that errors still point at the template.
 */
import { describe, expect, it, vi } from 'vitest';
import { createReport } from '../index';
import { CommandExecutionError } from '../errors';
import { type SandBox } from '../types';
import { getError, makeDocx, reportText } from './helpers';

const render = async (
  body: string[],
  options: Record<string, unknown> = {}
): Promise<string> => {
  const template = await makeDocx({ body });
  return reportText(await createReport({ template, ...options }, 'JS'));
};

const unchanged = (code: string) => code;

for (const noSandbox of [false, true]) {
  describe(noSandbox ? 'noSandbox' : 'sandbox', () => {
    it('evaluates the code the transform returns', async () => {
      expect(
        await render(['+++greeting+++'], {
          noSandbox,
          data: { greeting: 'hello', farewell: 'goodbye' },
          transformCode: (code: string) => code.replace('greeting', 'farewell'),
        })
      ).toEqual('goodbye');
    });

    it('rewrites the code of every command that runs JavaScript', async () => {
      const transformCode = vi.fn(unchanged);
      await render(
        [
          '+++EXEC total = 0+++',
          '+++FOR item IN items+++',
          '+++IF $item > 1+++',
          '+++$item+++',
          '+++END-IF+++',
          '+++END-FOR item+++',
        ],
        { noSandbox, data: { items: [1, 2] }, transformCode }
      );
      expect(new Set(transformCode.mock.calls.map(([code]) => code))).toEqual(
        new Set(['total = 0', 'items', '$item > 1', '$item'])
      );
    });

    it('sees the code once operatorAliases have been applied', async () => {
      const transformCode = vi.fn(unchanged);
      expect(
        await render(['+++IF value больше 1+++big+++END-IF+++'], {
          noSandbox,
          data: { value: 2 },
          operatorAliases: { больше: '>' },
          transformCode,
        })
      ).toEqual('big');
      expect(transformCode).toHaveBeenCalledWith('value > 1');
    });

    it('rewrites a snippet once, however many times a loop runs it', async () => {
      const transformCode = vi.fn(unchanged);
      expect(
        await render(
          ['+++FOR item IN items+++', '+++$item+++', '+++END-FOR item+++'],
          { noSandbox, data: { items: [1, 2, 3] }, transformCode }
        )
      ).toEqual('1\n2\n3');
      expect(
        transformCode.mock.calls.filter(([code]) => code === '$item')
      ).toHaveLength(1);
    });

    it('lets the rewritten code call helpers from additionalJsContext', async () => {
      // The motivating case: comparing `dd.mm.yyyy` strings as dates.
      const toKey = (date: string) => date.split('.').reverse().join('');
      expect(
        await render(["+++start > end ? 'later' : 'earlier'+++"], {
          noSandbox,
          data: { start: '01.02.2024', end: '15.01.2024' },
          additionalJsContext: {
            isAfter: (a: string, b: string) => toKey(a) > toKey(b),
          },
          transformCode: (code: string) =>
            code.replace('start > end', 'isAfter(start, end)'),
        })
      ).toEqual('later');
    });

    it('quotes the code as the template has it when the snippet throws', async () => {
      const template = await makeDocx({ body: ['+++value+++'] });
      const err = await getError<CommandExecutionError>(() =>
        createReport({
          template,
          noSandbox,
          data: { value: 1 },
          transformCode: code => `missing.${code}`,
        })
      );
      expect(err).toBeInstanceOf(CommandExecutionError);
      expect(err.command).toEqual('value');
      expect(err.err.name).toEqual('ReferenceError');
    });

    it('hands errorHandler the code as the template has it', async () => {
      const errorHandler = vi.fn(() => 'fallback');
      expect(
        await render(['+++value+++'], {
          noSandbox,
          data: { value: 1 },
          transformCode: (code: string) => `missing.${code}`,
          errorHandler,
        })
      ).toEqual('fallback');
      // By name: the `vm` sandbox throws errors of its own realm.
      expect(errorHandler).toHaveBeenCalledWith(
        expect.objectContaining({ name: 'ReferenceError' }),
        'value'
      );
    });

    it('fails the command when the transform throws', async () => {
      const template = await makeDocx({ body: ['+++value+++'] });
      const err = await getError<CommandExecutionError>(() =>
        createReport({
          template,
          noSandbox,
          data: { value: 1 },
          transformCode: () => {
            throw new Error('cannot rewrite');
          },
        })
      );
      expect(err).toBeInstanceOf(CommandExecutionError);
      expect(err.command).toEqual('value');
      expect(err.err.message).toEqual('cannot rewrite');
    });

    it('fails the command when the transform does not return a string', async () => {
      const template = await makeDocx({ body: ['+++value+++'] });
      const err = await getError<CommandExecutionError>(() =>
        createReport({
          template,
          noSandbox,
          data: { value: 1 },
          transformCode: (() => undefined) as never,
        })
      );
      expect(err).toBeInstanceOf(CommandExecutionError);
      expect(err.err).toBeInstanceOf(TypeError);
      expect(err.err.message).toEqual(
        'transformCode must return a string, but returned undefined'
      );
    });
  });
}

describe('runJs', () => {
  it('receives the rewritten code', async () => {
    const runJs = ({ sandbox }: { sandbox: SandBox }) => ({
      modifiedSandbox: sandbox,
      result: sandbox.__code__,
    });
    expect(
      await render(['+++value+++'], {
        data: {},
        runJs,
        transformCode: (code: string) => `${code}!`,
      })
    ).toEqual('value!');
  });
});
