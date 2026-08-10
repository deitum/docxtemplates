import {
  type AliasList,
  BUILT_IN_COMMANDS,
  type BuiltInCommand,
} from './types';
import { InvalidAliasError } from './errors';

// Characters that can never be part of an "alias word". Everything else
// (including non-ASCII letters, e.g. Cyrillic) is considered a word character,
// so that aliases are matched on whole words only.
const NON_WORD_CHAR = /[\s!"#%&'()*+,\-./:;<=>?[\\\]^`{|}~]/;

const isWordChar = (char: string | undefined): boolean =>
  char != null && char !== '' && !NON_WORD_CHAR.test(char);

const isWhitespace = (char: string | undefined): boolean =>
  char != null && /\s/.test(char);

/**
 * Turns a user-provided alias map into a list that is sorted so that the
 * longest (i.e. most specific) alias is always matched first. This is what
 * makes e.g. `больше или равно` win over `больше`.
 */
export function compileAliases(aliases?: {
  [alias: string]: string;
}): AliasList {
  const out: AliasList = [];
  if (aliases == null) return out;
  for (const alias of Object.keys(aliases)) {
    const tokens = alias
      .trim()
      .split(/\s+/)
      .filter(token => token !== '');
    if (!tokens.length) throw new InvalidAliasError('Empty alias', alias);
    const replacement = aliases[alias];
    if (typeof replacement !== 'string')
      throw new InvalidAliasError('Alias replacement must be a string', alias);
    out.push({ tokens: tokens.map(token => token.toLowerCase()), replacement });
  }
  out.sort((a, b) => b.tokens.join(' ').length - a.tokens.join(' ').length);
  return out;
}

/**
 * Same as `compileAliases`, but validates that every alias maps to one of the
 * built-in commands.
 */
export function compileCommandAliases(aliases?: {
  [alias: string]: string;
}): AliasList {
  const out = compileAliases(aliases);
  for (const alias of out) {
    alias.replacement = alias.replacement.trim().toUpperCase();
    if (!BUILT_IN_COMMANDS.includes(alias.replacement as BuiltInCommand))
      throw new InvalidAliasError(
        `Command aliases must point to one of ${BUILT_IN_COMMANDS.join(', ')}`,
        `${alias.tokens.join(' ')} -> ${alias.replacement}`
      );
  }
  return out;
}

// Tries to match `tokens` in `text`, starting at `idx`. Matching is
// case-insensitive, and any amount of whitespace is allowed between tokens.
// Returns the index just after the match, or -1 if there is no match.
const matchTokensAt = (text: string, idx: number, tokens: string[]): number => {
  let i = idx;
  for (let t = 0; t < tokens.length; t++) {
    // Tokens must be separated by at least one whitespace character
    if (t > 0) {
      const startOfGap = i;
      while (i < text.length && isWhitespace(text[i])) i++;
      if (i === startOfGap) return -1;
    }
    const token = tokens[t];
    if (token == null) return -1;
    const candidate = text.slice(i, i + token.length);
    if (candidate.length < token.length) return -1;
    if (candidate.toLowerCase() !== token) return -1;
    i += token.length;
  }
  return i;
};

// As `matchTokensAt`, but also checks that the match is not part of a longer
// word (only relevant on sides where the alias itself starts/ends with a word
// character, so that e.g. `>=` can be aliased too).
const matchAliasAt = (text: string, idx: number, tokens: string[]): number => {
  const firstToken = tokens[0];
  const lastToken = tokens[tokens.length - 1];
  // An empty alias can never match anything.
  if (firstToken == null || lastToken == null) return -1;
  if (isWordChar(firstToken[0]) && isWordChar(text[idx - 1])) return -1;
  const end = matchTokensAt(text, idx, tokens);
  if (end < 0) return -1;
  if (isWordChar(lastToken[lastToken.length - 1]) && isWordChar(text[end]))
    return -1;
  return end;
};

/**
 * Replaces all the aliases found in a command expression with their JS
 * equivalent. The contents of string literals are left untouched.
 */
export function substituteAliases(text: string, aliases: AliasList): string {
  if (!aliases.length || !text) return text;
  let out = '';
  let idx = 0;
  let openQuote: string | null = null;
  while (idx < text.length) {
    const char = text[idx];

    // Inside a string literal: copy verbatim until it is closed
    if (openQuote != null) {
      out += char;
      if (char === '\\') {
        if (idx + 1 < text.length) out += text[idx + 1];
        idx += 2;
        continue;
      }
      if (char === openQuote) openQuote = null;
      idx += 1;
      continue;
    }
    if (char === "'" || char === '"' || char === '`') {
      openQuote = char;
      out += char;
      idx += 1;
      continue;
    }

    let matched = false;
    for (const { tokens, replacement } of aliases) {
      const end = matchAliasAt(text, idx, tokens);
      if (end < 0) continue;
      out += replacement;
      idx = end;
      matched = true;
      break;
    }
    if (!matched) {
      out += char;
      idx += 1;
    }
  }
  return out;
}

/**
 * Replaces the leading command name of a command with the built-in command it
 * is an alias for. Returns `undefined` if the command doesn't start with a
 * known alias.
 */
export function resolveCommandAlias(
  cmd: string,
  aliases: AliasList
): string | undefined {
  for (const { tokens, replacement } of aliases) {
    const end = matchAliasAt(cmd, 0, tokens);
    if (end < 0) continue;
    return `${replacement} ${cmd.slice(end).trim()}`.trim();
  }
  return undefined;
}
