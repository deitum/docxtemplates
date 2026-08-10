/**
 * Dispatch of a single command, once the walk has collected it between two
 * delimiters.
 *
 * Everything specific to a command lives in `registry.ts`; what is left here is
 * what every command shares — resolving the text into a name and a payload,
 * skipping the commands that must not run while a loop is being explored, and
 * routing whatever they throw to the `errorHandler`.
 */
import { logger } from '../debug';
import { CommandSyntaxError, isError } from '../errors';
import { isLoopExploring } from '../reportUtils';
import { Command, type Context, type Node, type ReportData } from '../types';
import { getCommand, splitCommand } from './parse';
import { specOf } from './registry';

export type CommandProcessor = (
  data: ReportData | undefined,
  node: Node,
  ctx: Context
) => Promise<undefined | string | Error>;

export const processCmd: CommandProcessor = async (data, node, ctx) => {
  // Deliberately outside the `try`: a template naming an alias that nothing has
  // defined is broken in a way the `errorHandler` is not meant to paper over.
  const cmd = getCommand(ctx.walk.command, ctx.scope.shorthands, ctx.options);
  ctx.walk.command = ''; // flush the context
  const { cmdName, cmdRest } = splitCommand(cmd, ctx.options.operatorAliases);

  try {
    if (cmdName !== Command.CMD_NODE) logger.debug(`Processing cmd: ${cmd}`);

    // While seeking the QUERY, no other command may run.
    if (ctx.walk.seekingQuery) {
      if (cmdName === Command.QUERY) ctx.walk.query = cmdRest;
      return;
    }

    const spec = cmdName != null ? specOf(cmdName) : undefined;
    if (spec == null || cmdName == null) throw new CommandSyntaxError(cmd);
    if (spec.skipWhileExploring && isLoopExploring(ctx)) return;

    const result = await spec.run({
      data,
      node,
      ctx,
      cmd,
      name: cmdName,
      rest: cmdRest,
    });
    return result ?? undefined;
  } catch (err) {
    if (!isError(err)) throw err;
    if (ctx.options.errorHandler != null) {
      return ctx.options.errorHandler(err, cmdRest);
    }
    return err;
  }
};
