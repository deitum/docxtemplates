import { describe, it, expect } from 'vitest';
import {
  compileAliases,
  compileCommandAliases,
  resolveCommandAlias,
  substituteAliases,
} from '../aliases';
import { InvalidAliasError } from '../errors';
import { type AliasList } from '../types';

const OPERATORS = compileAliases({
  'больше или равно': '>=',
  больше: '>',
  равно: '===',
  'не равно': '!==',
  и: '&&',
});

describe('compileAliases', () => {
  it('lowercases the alias words and keeps the replacement as-is', () => {
    expect(compileAliases({ 'БОЛЬШЕ  ИЛИ  РАВНО': '>=' })).toEqual([
      { tokens: ['больше', 'или', 'равно'], replacement: '>=' },
    ]);
  });

  it('sorts the longest alias first, so that it wins', () => {
    const compiled = compileAliases({ больше: '>', 'больше или равно': '>=' });
    expect(compiled.map(a => a.tokens.join(' '))).toEqual([
      'больше или равно',
      'больше',
    ]);
  });

  it('returns an empty list when no aliases are given', () => {
    expect(compileAliases()).toEqual([]);
    expect(compileAliases({})).toEqual([]);
  });

  it('rejects an alias made up of whitespace only', () => {
    expect(() => compileAliases({ '   ': '>' })).toThrow(InvalidAliasError);
    expect(() => compileAliases({ '   ': '>' })).toThrow('Empty alias');
  });

  it('rejects a non-string replacement', () => {
    const aliases = { больше: 42 } as unknown as { [alias: string]: string };
    expect(() => compileAliases(aliases)).toThrow(
      'Alias replacement must be a string'
    );
  });
});

describe('compileCommandAliases', () => {
  it('normalizes the replacement to an upper-cased built-in command', () => {
    expect(compileCommandAliases({ ЕСЛИ: ' else-if ' })).toEqual([
      { tokens: ['если'], replacement: 'ELSE-IF' },
    ]);
  });

  it('rejects a replacement that is not a built-in command', () => {
    expect(() => compileCommandAliases({ ЕСЛИ: 'NOPE' })).toThrow(
      InvalidAliasError
    );
    expect(() => compileCommandAliases({ ЕСЛИ: 'NOPE' })).toThrow(
      'Command aliases must point to one of'
    );
  });
});

describe('substituteAliases', () => {
  it('replaces whole-word aliases', () => {
    expect(substituteAliases('a больше b', OPERATORS)).toEqual('a > b');
    expect(substituteAliases('a больше b и c равно d', OPERATORS)).toEqual(
      'a > b && c === d'
    );
  });

  it('is case-insensitive and tolerates extra whitespace between words', () => {
    expect(substituteAliases('a БОЛЬШЕ   ИЛИ\tРАВНО b', OPERATORS)).toEqual(
      'a >= b'
    );
  });

  it('prefers the longest matching alias', () => {
    expect(substituteAliases('a больше или равно b', OPERATORS)).toEqual(
      'a >= b'
    );
    expect(substituteAliases('a не равно b', OPERATORS)).toEqual('a !== b');
  });

  it('leaves the text alone when there are no aliases, or no text', () => {
    expect(substituteAliases('a больше b', [])).toEqual('a больше b');
    expect(substituteAliases('', OPERATORS)).toEqual('');
  });

  it('does not match an alias that is part of a longer word', () => {
    // `равно` is a prefix of `равного` and a suffix of `неравно`
    expect(substituteAliases('x равного y', OPERATORS)).toEqual('x равного y');
    expect(substituteAliases('x неравно y', OPERATORS)).toEqual('x неравно y');
  });

  it('does not match a multi-word alias whose words are run together', () => {
    expect(substituteAliases('x неравно y', OPERATORS)).toEqual('x неравно y');
    expect(substituteAliases('x не  равно y', OPERATORS)).toEqual('x !== y');
  });

  it('does not match a multi-word alias cut short by the end of the text', () => {
    expect(substituteAliases('x не', OPERATORS)).toEqual('x не');
    expect(substituteAliases('x не ', OPERATORS)).toEqual('x не ');
  });

  it('matches a symbolic alias even when it touches word characters', () => {
    const symbolic = compileAliases({ '=>': '>=' });
    expect(substituteAliases('a=>b', symbolic)).toEqual('a>=b');
  });

  it('never matches an alias with no words', () => {
    const broken: AliasList = [{ tokens: [], replacement: 'X' }];
    expect(substituteAliases('anything', broken)).toEqual('anything');
  });

  it('leaves the contents of string literals untouched', () => {
    expect(substituteAliases(`'a больше b' и c`, OPERATORS)).toEqual(
      `'a больше b' && c`
    );
    expect(substituteAliases(`"a больше b" и c`, OPERATORS)).toEqual(
      `"a больше b" && c`
    );
    expect(substituteAliases('`a больше b` и c', OPERATORS)).toEqual(
      '`a больше b` && c'
    );
  });

  it('handles escaped quotes inside string literals', () => {
    expect(substituteAliases(`'it\\'s больше' и c`, OPERATORS)).toEqual(
      `'it\\'s больше' && c`
    );
    // A trailing backslash must not run past the end of the string
    expect(substituteAliases(`'unterminated \\`, OPERATORS)).toEqual(
      `'unterminated \\`
    );
  });
});

describe('resolveCommandAlias', () => {
  const COMMANDS = compileCommandAliases({
    ЕСЛИ: 'IF',
    'ИНАЧЕ ЕСЛИ': 'ELSE-IF',
    ИНАЧЕ: 'ELSE',
  });

  it('replaces the leading alias with the built-in command', () => {
    expect(resolveCommandAlias('ЕСЛИ значение > 10', COMMANDS)).toEqual(
      'IF значение > 10'
    );
  });

  it('prefers the longest matching alias', () => {
    expect(resolveCommandAlias('ИНАЧЕ ЕСЛИ значение > 5', COMMANDS)).toEqual(
      'ELSE-IF значение > 5'
    );
  });

  it('handles an alias with no arguments', () => {
    expect(resolveCommandAlias('ИНАЧЕ', COMMANDS)).toEqual('ELSE');
    expect(resolveCommandAlias('ИНАЧЕ   ', COMMANDS)).toEqual('ELSE');
  });

  it('returns undefined when the command does not start with an alias', () => {
    expect(resolveCommandAlias('IF значение > 10', COMMANDS)).toBeUndefined();
    expect(resolveCommandAlias('значение ЕСЛИ', COMMANDS)).toBeUndefined();
    expect(resolveCommandAlias('ЕСЛИБО значение', COMMANDS)).toBeUndefined();
    expect(resolveCommandAlias('ЕСЛИ значение', [])).toBeUndefined();
  });
});
