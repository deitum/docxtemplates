import { describe, expect, it } from 'vitest';
import { getCommand, splitCommand } from '../commands/parse';
import { COMMANDS, EXPRESSION_COMMANDS, specOf } from '../commands/registry';
import { BUILT_IN_COMMANDS, Command } from '../types';

const OPTIONS = { fixSmartQuotes: false, commandAliases: [] };
const resolve = (cmd: string) => getCommand(cmd, {}, OPTIONS);

describe('getCommand', () => {
  it('leaves a built-in command alone', () => {
    expect(resolve('FOR company IN companies')).toEqual(
      'FOR company IN companies'
    );
    expect(resolve('END-FOR company')).toEqual('END-FOR company');
  });

  it('adds the implicit INS to a bare expression', () => {
    expect(resolve('company.name')).toEqual('INS company.name');
    expect(resolve('姓名')).toEqual('INS 姓名');
  });

  it('expands the = and ! prefixes', () => {
    expect(resolve('= a + b')).toEqual('INS a + b');
    expect(resolve('! doThing()')).toEqual('EXEC doThing()');
  });

  it('expands a *shorthand defined by an ALIAS', () => {
    expect(getCommand('*name', { name: 'INS $company.name' }, OPTIONS)).toEqual(
      'INS $company.name'
    );
  });

  it('rejects an unknown *shorthand', () => {
    expect(() => resolve('*nope')).toThrow('Unknown alias');
  });

  // Recognising a built-in used to mean running one anchored regex per command
  // name against every command. This pins the replacement to exactly what those
  // regexes did — including the corners, e.g. `IF-x` starting with `IF` and a
  // word boundary, while `INSertable` does not start with `INS`.
  describe('matches the regex-based recognition it replaced', () => {
    const regexes = BUILT_IN_COMMANDS.map(name => new RegExp(`^${name}\\b`));
    const wasBuiltIn = (cmd: string) =>
      regexes.some(r => r.test(cmd.toUpperCase()));

    const CANDIDATES = [
      ...BUILT_IN_COMMANDS,
      ...BUILT_IN_COMMANDS.map(c => `${c} rest`),
      ...BUILT_IN_COMMANDS.map(c => `${c.toLowerCase()} rest`),
      ...BUILT_IN_COMMANDS.map(c => `${c}x`),
      ...BUILT_IN_COMMANDS.map(c => `${c}-x`),
      ...BUILT_IN_COMMANDS.map(c => `${c}(1)`),
      ...BUILT_IN_COMMANDS.map(c => `x${c}`),
      'INSertable',
      'FOREACH',
      'IF-x',
      'ELSE-IFX',
      'END-FORX',
      'END',
      'END-',
      'CMD_NODE',
      'CMD',
      'formatNumber(123)',
      'a > b',
      '姓名',
      '',
      '-',
      '_private',
      'html_content',
    ];

    for (const cmd of CANDIDATES) {
      it(`agrees on ${JSON.stringify(cmd)}`, () => {
        const expected = wasBuiltIn(cmd) ? cmd : `${Command.INS} ${cmd}`;
        expect(resolve(cmd)).toEqual(expected.trim());
      });
    }
  });
});

describe('splitCommand', () => {
  it('splits a command into its upper-cased name and the rest', () => {
    expect(splitCommand('for company IN companies')).toEqual({
      cmdName: 'FOR',
      cmdRest: 'company IN companies',
    });
  });

  it('reports no name for an empty command', () => {
    expect(splitCommand('')).toEqual({ cmdName: undefined, cmdRest: '' });
  });

  it('substitutes operator aliases only in expression commands', () => {
    const aliases = [{ tokens: ['больше'], replacement: '>' }];
    expect(splitCommand('IF a больше b', aliases).cmdRest).toEqual('a > b');
    // `ALIAS` defines a name; its payload is not an expression.
    expect(splitCommand('ALIAS x больше', aliases).cmdRest).toEqual('x больше');
  });
});

describe('command registry', () => {
  it('has an entry for every command', () => {
    // The point of the registry: a command that exists but was never wired up
    // would otherwise fail at run time, on whichever template happens to use it.
    expect(Object.keys(COMMANDS).sort()).toEqual([...BUILT_IN_COMMANDS].sort());
  });

  it('agrees with the set of commands whose payload is an expression', () => {
    expect([...EXPRESSION_COMMANDS].sort()).toEqual(
      [
        Command.ELSE_IF,
        Command.EXEC,
        Command.FOR,
        Command.HTML,
        Command.IF,
        Command.IMAGE,
        Command.INS,
        Command.LINK,
      ].sort()
    );
  });

  it('skips exactly the output-producing commands while exploring a loop', () => {
    const skipped = Object.entries(COMMANDS)
      .filter(([, spec]) => spec.skipWhileExploring)
      .map(([name]) => name)
      .sort();
    expect(skipped).toEqual(
      [
        Command.EXEC,
        Command.HTML,
        Command.IMAGE,
        Command.INS,
        Command.LINK,
      ].sort()
    );
  });

  it('resolves a command name to its spec, and nothing else', () => {
    expect(specOf(Command.FOR)).toBe(COMMANDS[Command.FOR]);
    expect(specOf('NOPE')).toBeUndefined();
    // Not a command, but a property of every object: the lookup must not find it.
    expect(specOf('constructor')).toBeUndefined();
    expect(specOf('toString')).toBeUndefined();
  });
});
