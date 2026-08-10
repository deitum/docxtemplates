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

const builtInRegexes = BUILT_IN_COMMANDS.map(word => new RegExp(`^${word}\\b`));

const isBuiltIn = (cmd: string) =>
  builtInRegexes.some(r => r.test(cmd.toUpperCase()));

/**
 * The quotes MS Word autocorrects straight ones into. They are not valid
 * JavaScript, so `fixSmartQuotes` puts them back.
 */
const SMART_DOUBLE_QUOTES = /[“”„]/g; // “ ” „
const SMART_SINGLE_QUOTES = /[‘’‚]/g; // ‘ ’ ‚

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
  if (cmd[0] === CommandPrefix.shorthand) {
    const aliasName = cmd.slice(1).trim();
    const shorthand = shorthands[aliasName];
    if (!shorthand) throw new InvalidCommandError('Unknown alias', cmd);
    logger.debug(`Alias for: ${shorthand}`);
    cmd = resolveCommandAlias(shorthand, options.commandAliases) ?? shorthand;
  } else if (cmd[0] === CommandPrefix.ins) {
    cmd = `${Command.INS} ${cmd.slice(1).trim()}`;
  } else if (cmd[0] === CommandPrefix.exec) {
    cmd = `${Command.EXEC} ${cmd.slice(1).trim()}`;
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

// Commands whose payload is a JS expression, and hence the only ones in which
// operator aliases (e.g. `больше` for `>`) get substituted.
const EXPRESSION_COMMANDS: string[] = [
  Command.FOR,
  Command.IF,
  Command.ELSE_IF,
  Command.INS,
  Command.EXEC,
  Command.IMAGE,
  Command.LINK,
  Command.HTML,
];

/** Splits a command into its (upper-cased) name and the rest. */
export function splitCommand(cmd: string, operatorAliases?: AliasList) {
  const cmdNameMatch = /^(\S+)\s*/.exec(cmd);
  let cmdName;
  let cmdRest = '';
  if (cmdNameMatch?.[1] != null) {
    cmdName = cmdNameMatch[1].toUpperCase();
    cmdRest = cmd.slice(cmdName.length).trim();
    if (operatorAliases?.length && EXPRESSION_COMMANDS.includes(cmdName)) {
      cmdRest = substituteAliases(cmdRest, operatorAliases);
    }
  }

  return { cmdName, cmdRest };
}
