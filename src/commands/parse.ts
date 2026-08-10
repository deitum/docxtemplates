import { resolveCommandAlias, substituteAliases } from '../aliases';
import { logger } from '../debug';
import { InvalidCommandError } from '../errors';
import {
  type AliasList,
  BUILT_IN_COMMANDS,
  Command,
  CommandPrefix,
  type Context,
  type CreateReportOptions,
} from '../types';
import { EXPRESSION_COMMANDS } from './registry';

const BUILT_IN_NAMES: ReadonlySet<string> = new Set(BUILT_IN_COMMANDS);

/** The characters a word boundary (`\b`) separates on. */
const isWordChar = (char: string | undefined) =>
  char != null && /[A-Za-z0-9_]/.test(char);

/** How far the run of word characters starting at `from` reaches. */
const endOfWord = (cmd: string, from: number): number => {
  let idx = from;
  while (idx < cmd.length && isWordChar(cmd[idx])) idx += 1;
  return idx;
};

/**
 * Whether the command starts with the name of a built-in.
 *
 * Equivalent to testing `^NAME\b` for every built-in, which is what this used to
 * do — fourteen regular expressions against every command that has no prefix.
 * A single word is enough to decide, except for the hyphenated names
 * (`END-FOR`, `ELSE-IF`, `END-IF`), which need the word after the hyphen too.
 */
const isBuiltIn = (cmd: string): boolean => {
  const firstEnd = endOfWord(cmd, 0);
  if (BUILT_IN_NAMES.has(cmd.slice(0, firstEnd).toUpperCase())) return true;
  if (cmd[firstEnd] !== '-') return false;
  const secondEnd = endOfWord(cmd, firstEnd + 1);
  return BUILT_IN_NAMES.has(cmd.slice(0, secondEnd).toUpperCase());
};

/**
 * The quotes MS Word autocorrects straight ones into. They are not valid
 * JavaScript, so `fixSmartQuotes` puts them back.
 */
const SMART_DOUBLE_QUOTES = /[“”„]/g; // “ ” „
const SMART_SINGLE_QUOTES = /[‘’‚]/g; // ‘ ’ ‚

/** What a single-character command prefix expands into. */
type PrefixExpansion = (args: {
  /** Whatever followed the prefix, trimmed. */
  rest: string;
  /** The whole command, prefix included, for error messages. */
  cmd: string;
  shorthands: Context['shorthands'];
  options: Pick<CreateReportOptions, 'commandAliases'>;
}) => string;

const PREFIX_EXPANSIONS: { [prefix: string]: PrefixExpansion } = {
  [CommandPrefix.shorthand]: ({ rest, cmd, shorthands, options }) => {
    const shorthand = shorthands[rest];
    if (!shorthand) throw new InvalidCommandError('Unknown alias', cmd);
    logger.debug(`Alias for: ${shorthand}`);
    return resolveCommandAlias(shorthand, options.commandAliases) ?? shorthand;
  },
  [CommandPrefix.ins]: ({ rest }) => `${Command.INS} ${rest}`,
  [CommandPrefix.exec]: ({ rest }) => `${Command.EXEC} ${rest}`,
};

/**
 * Normalises the raw text found between the command delimiters into a full
 * command: expands `*shorthand` / `=` / `!` prefixes, resolves user-defined
 * names for the built-ins, and adds the implicit `INS` that lets a template
 * write `+++name+++` instead of `+++INS name+++`.
 */
export function getCommand(
  command: string,
  shorthands: Context['shorthands'],
  options: Pick<CreateReportOptions, 'fixSmartQuotes' | 'commandAliases'>
): string {
  let cmd = command.trim();
  const expandPrefix = PREFIX_EXPANSIONS[cmd[0] ?? ''];
  if (expandPrefix != null) {
    cmd = expandPrefix({ rest: cmd.slice(1).trim(), cmd, shorthands, options });
  } else {
    // A user-defined name for a built-in command (e.g. `ЕСЛИ` for `IF`)?
    const aliased = resolveCommandAlias(cmd, options.commandAliases);
    if (aliased != null) cmd = aliased;
    else if (!isBuiltIn(cmd)) cmd = `${Command.INS} ${cmd}`;
  }

  if (options.fixSmartQuotes) {
    cmd = cmd
      .replace(SMART_DOUBLE_QUOTES, '"')
      .replace(SMART_SINGLE_QUOTES, "'");
  }

  return cmd.trim();
}

/** Splits a command into its (upper-cased) name and the rest. */
export function splitCommand(cmd: string, operatorAliases?: AliasList) {
  const cmdNameMatch = /^(\S+)\s*/.exec(cmd);
  let cmdName;
  let cmdRest = '';
  if (cmdNameMatch?.[1] != null) {
    cmdName = cmdNameMatch[1].toUpperCase();
    cmdRest = cmd.slice(cmdName.length).trim();
    // Operator aliases (`больше` for `>`, say) are only substituted where the
    // payload is a JS expression; `EXPRESSION_COMMANDS` comes from the command
    // registry, so it cannot fall out of step with it.
    if (operatorAliases?.length && EXPRESSION_COMMANDS.has(cmdName)) {
      cmdRest = substituteAliases(cmdRest, operatorAliases);
    }
  }

  return { cmdName, cmdRest };
}
