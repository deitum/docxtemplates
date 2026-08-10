/**
 * Execution of a single command, once the walk has collected it between two
 * delimiters. Control-flow commands (FOR/IF/ELSE-IF/ELSE/END-*) manipulate the
 * loop stack that `walkTemplate` reads to decide where to go next; the rest
 * either return text to insert or park a node on the context.
 */
import { logger } from '../debug';
import { WTag, XML_SPACE_PRESERVE, XmlAttr } from '../ooxml';
import {
  CommandSyntaxError,
  InvalidCommandError,
  ImageError,
  isError,
  ObjectCommandResultError,
} from '../errors';
import { runUserJsAndGetRaw } from '../jsSandbox';
import {
  findParentPorTrNode,
  getCurLoop,
  isLoopExploring,
  isLoopSkippingOutput,
  logLoop,
  markCellIfLoopSpansCells,
  tagOf,
} from '../reportUtils';
import {
  Command,
  type Context,
  EXPLORATION_PASS,
  type ImagePars,
  type LinkPars,
  type LoopStatus,
  type Node,
  type ReportData,
} from '../types';
import { processHtml, processImage, processLink } from './media';
import { getCommand, splitCommand } from './parse';

/** Prefix of the internal variable name an IF construct is tracked under. */
const IF_VAR_PREFIX = '__if_' as const;

/** `ifActiveBranch` while no branch condition has matched. */
const NO_BRANCH = -1;

const NEWLINE = /\n/g;

/**
 * `loopOver` for an IF construct: a single (arbitrary) item means "render the
 * selected branch on a second pass", an empty array means "render nothing".
 */
const IF_RENDER_ONCE = [1];
const IF_RENDER_NEVER: unknown[] = [];

export type CommandProcessor = (
  data: ReportData | undefined,
  node: Node,
  ctx: Context
) => Promise<undefined | string | Error>;

export const processCmd: CommandProcessor = async (data, node, ctx) => {
  const cmd = getCommand(ctx.cmd, ctx.shorthands, ctx.options);
  ctx.cmd = ''; // flush the context
  const { cmdName, cmdRest } = splitCommand(cmd, ctx.options.operatorAliases);
  try {
    if (cmdName !== Command.CMD_NODE) logger.debug(`Processing cmd: ${cmd}`);

    // While seeking the QUERY, no other command may run.
    if (ctx.fSeekQuery) {
      if (cmdName === Command.QUERY) ctx.query = cmdRest;
      return;
    }

    switch (cmdName) {
      // `CMD_NODE` marks a text node emptied out by `preprocessTemplate`;
      // `QUERY` was already consumed above.
      case Command.QUERY:
      case Command.CMD_NODE:
        return;

      case Command.ALIAS:
        processAlias(ctx, cmd, cmdRest);
        return;

      case Command.FOR:
      case Command.IF:
        await processForIf(data, node, ctx, cmd, cmdName, cmdRest);
        return;

      case Command.ELSE_IF:
      case Command.ELSE:
        await processElse(data, node, ctx, cmd, cmdName, cmdRest);
        return;

      case Command.END_FOR:
      case Command.END_IF:
        processEndForIf(node, ctx, cmd, cmdName, cmdRest);
        return;

      case Command.INS:
        if (isLoopExploring(ctx)) return;
        return await processIns(data, ctx, cmdRest);

      case Command.EXEC:
        if (isLoopExploring(ctx)) return;
        await runUserJsAndGetRaw(data, cmdRest, ctx);
        return;

      case Command.IMAGE:
        if (isLoopExploring(ctx)) return;
        await processImageCmd(data, ctx, cmd, cmdRest);
        return;

      case Command.LINK: {
        if (isLoopExploring(ctx)) return;
        const pars: LinkPars | undefined = await runUserJsAndGetRaw(
          data,
          cmdRest,
          ctx
        );
        if (pars != null) processLink(ctx, pars);
        return;
      }

      case Command.HTML: {
        if (isLoopExploring(ctx)) return;
        const html: string | undefined = await runUserJsAndGetRaw(
          data,
          cmdRest,
          ctx
        );
        if (html != null) processHtml(ctx, html);
        return;
      }

      default:
        throw new CommandSyntaxError(cmd);
    }
  } catch (err) {
    if (!isError(err)) throw err;
    if (ctx.options.errorHandler != null) {
      return ctx.options.errorHandler(err, cmdRest);
    }
    return err;
  }
};

// ==========================================
// ALIAS, INS, IMAGE
// ==========================================

// ALIAS <name> <the command it stands for>
const processAlias = (ctx: Context, cmd: string, cmdRest: string) => {
  const aliasMatch = /^(\S+)\s+(.+)/.exec(cmdRest);
  const [, aliasName, fullCmd] = aliasMatch ?? [];
  if (aliasName == null || fullCmd == null)
    throw new InvalidCommandError('Invalid ALIAS command', cmd);
  ctx.shorthands[aliasName] = fullCmd;
  logger.debug(`Defined alias '${aliasName}' for: ${fullCmd}`);
};

// INS <expression>
const processIns = async (
  data: ReportData | undefined,
  ctx: Context,
  cmdRest: string
): Promise<string> => {
  let result = await runUserJsAndGetRaw(data, cmdRest, ctx);
  if (result == null) return '';

  // An object would stringify to '[object Object]', which is never what the
  // template author meant.
  if (typeof result === 'object' && !Array.isArray(result)) {
    const err = new ObjectCommandResultError(cmdRest, result);
    if (ctx.options.errorHandler == null) throw err;
    result = await ctx.options.errorHandler(err, cmdRest);
  }

  const str = String(result);
  return ctx.options.processLineBreaks ? insertLineBreaks(str, ctx) : str;
};

/**
 * Replaces newlines with a `w:br` tag, protected by the `literalXmlDelimiter`
 * separators so that `buildXml` passes the markup through unescaped.
 */
const insertLineBreaks = (str: string, ctx: Context): string => {
  const { literalXmlDelimiter: d } = ctx.options;
  // Wrapped in the literal-XML delimiters, so that `buildXml` emits the markup
  // as-is instead of escaping it.
  const lineBreak = `${d}<${WTag.br}/>${d}`;
  if (!ctx.options.processLineBreaksAsNewText) {
    return str.replace(NEWLINE, lineBreak);
  }
  // Closing and reopening the `w:t` around the break renders better in
  // LibreOffice and Google Drive.
  const endOfText = `${d}</${WTag.t}>${d}`;
  const startOfText = `${d}<${WTag.t} ${XmlAttr.space}="${XML_SPACE_PRESERVE}">${d}`;
  return str.split('\n').join(`${endOfText}${lineBreak}${startOfText}`);
};

// IMAGE <expression>
const processImageCmd = async (
  data: ReportData | undefined,
  ctx: Context,
  cmd: string,
  cmdRest: string
) => {
  const img: ImagePars | undefined = await runUserJsAndGetRaw(
    data,
    cmdRest,
    ctx
  );
  if (img == null) return;
  try {
    processImage(ctx, img);
  } catch (e) {
    if (!isError(e)) throw e;
    throw new ImageError(e, cmd);
  }
};

// ==========================================
// FOR / IF / ELSE-IF / ELSE / END-FOR / END-IF
// ==========================================

// FOR <varName> IN <expression>
// IF <expression>
const processForIf = async (
  data: ReportData | undefined,
  node: Node,
  ctx: Context,
  cmd: string,
  cmdName: string,
  cmdRest: string
): Promise<void> => {
  const isIf = cmdName === Command.IF;

  // Identify the FOR/IF loop. IF constructs have no variable of their own, so
  // they get an arbitrary name that the matching END-IF picks up.
  let forMatch: RegExpExecArray | null = null;
  let varName: string;
  if (isIf) {
    if (!node._ifName) {
      node._ifName = `${IF_VAR_PREFIX}${ctx.gCntIf}`;
      ctx.gCntIf += 1;
    }
    varName = node._ifName;
  } else {
    forMatch = /^(\S+)\s+IN\s+(.+)/i.exec(cmdRest);
    if (forMatch?.[1] == null || forMatch[2] == null)
      throw new InvalidCommandError('Invalid FOR command', cmd);
    varName = forMatch[1];
  }

  // Have we already seen this node, or is this a new construct?
  const curLoop = getCurLoop(ctx);
  if (curLoop && curLoop.varName === varName) {
    // We're revisiting the IF command on the second pass (the one that renders
    // the selected branch); start again from the first branch.
    if (isIf) restartIfBranches(curLoop);
    logLoop(ctx.loops);
    return;
  }

  if (isIf) checkNoNestedIfInSameScope(ctx, node, cmd);

  const parentLoop = getCurLoop(ctx);
  const fParentIsExploring =
    parentLoop != null && isLoopSkippingOutput(parentLoop);

  // For IF constructs, the branch that gets rendered (if any) is only known
  // once every ELSE-IF/ELSE has been explored, i.e. when END-IF is reached;
  // that is where `loopOver` gets filled in.
  let loopOver: unknown[];
  let ifBranchTaken = false;
  let ifActiveBranch: number = NO_BRANCH;
  if (fParentIsExploring) {
    loopOver = [];
    // Nothing inside this IF must be rendered (nor any of its branches
    // selected), since the parent is not rendering its contents either.
    if (isIf) ifBranchTaken = true;
  } else if (isIf) {
    loopOver = [];
    ifBranchTaken = !!(await runUserJsAndGetRaw(data, cmdRest, ctx));
    if (ifBranchTaken) ifActiveBranch = 0;
  } else {
    const loopExpression = forMatch?.[2];
    if (loopExpression == null)
      throw new InvalidCommandError('Invalid FOR command', cmd);
    loopOver = await runUserJsAndGetRaw(data, loopExpression, ctx);
    if (!Array.isArray(loopOver))
      throw new InvalidCommandError(
        'Invalid FOR command (can only iterate over Array)',
        cmd
      );
  }

  ctx.loops.push({
    refNode: node,
    refNodeLevel: ctx.level,
    varName,
    loopOver,
    isIf,
    // Run through the loop once first without outputting anything; otherwise
    // empty loops could not be detected.
    idx: EXPLORATION_PASS,
    ...(isIf ? { ifCurrentBranch: 0, ifActiveBranch, ifBranchTaken } : {}),
  });
  logLoop(ctx.loops);
};

/**
 * Two IF constructs on the same paragraph or table row would fight over the
 * same node when the walk jumps back, and loop forever. Upstream issue #340.
 */
const checkNoNestedIfInSameScope = (ctx: Context, node: Node, cmd: string) => {
  const scopeNode = findParentPorTrNode(node);
  if (scopeNode == null) return;
  const tag = tagOf(scopeNode);
  const seen =
    tag === WTag.p
      ? ctx.pIfCheckMap
      : tag === WTag.tr
        ? ctx.trIfCheckMap
        : null;
  if (seen == null) return;
  if (seen.has(scopeNode) && seen.get(scopeNode) !== cmd) {
    throw new InvalidCommandError(
      `Invalid IF command nested into another IF command on the same ${
        tag === WTag.p ? 'line' : 'table row'
      }`,
      cmd
    );
  }
  seen.set(scopeNode, cmd);
};

/**
 * Prepares an IF construct for a new pass through its branches. The branch
 * selected during the exploration pass is kept, of course.
 */
const restartIfBranches = (loop: LoopStatus) => {
  loop.ifCurrentBranch = 0;
  loop.ifElseBranch = undefined;
};

// ELSE-IF <expression>
// ELSE
const processElse = async (
  data: ReportData | undefined,
  node: Node,
  ctx: Context,
  cmd: string,
  cmdName: string,
  cmdRest: string
): Promise<void> => {
  const isElseIf = cmdName === Command.ELSE_IF;
  const curLoop = getCurLoop(ctx);
  if (!curLoop || !curLoop.isIf)
    throw new InvalidCommandError(
      `Unexpected ${cmdName} outside of IF statement context`,
      cmd
    );
  if (isElseIf && !cmdRest)
    throw new InvalidCommandError(
      'Invalid ELSE-IF command (missing condition)',
      cmd
    );
  if (curLoop.ifElseBranch != null)
    throw new InvalidCommandError(
      `Unexpected ${cmdName} after an ELSE command`,
      cmd
    );
  markCellIfLoopSpansCells(ctx, node, curLoop);

  // Move on to the next branch of the IF construct
  const branch = (curLoop.ifCurrentBranch ?? 0) + 1;
  curLoop.ifCurrentBranch = branch;
  if (!isElseIf) curLoop.ifElseBranch = branch;

  // Conditions are only evaluated during the exploration pass; the second pass
  // simply walks the tree again, rendering the branch that was selected.
  if (curLoop.idx <= EXPLORATION_PASS && !curLoop.ifBranchTaken) {
    let shouldRun = true;
    if (isElseIf) {
      // Evaluate the expression as if we were outside of the IF loop, so that
      // e.g. `$idx` refers to the enclosing FOR loop, just like in an IF command
      const ifLoop = ctx.loops.pop();
      try {
        shouldRun = !!(await runUserJsAndGetRaw(data, cmdRest, ctx));
      } finally {
        if (ifLoop) ctx.loops.push(ifLoop);
      }
    }
    if (shouldRun) {
      curLoop.ifBranchTaken = true;
      curLoop.ifActiveBranch = branch;
    }
  }
  logLoop(ctx.loops);
};

// END-FOR [varName]
// END-IF
const processEndForIf = (
  node: Node,
  ctx: Context,
  cmd: string,
  cmdName: string,
  cmdRest: string
): void => {
  const isIf = cmdName === Command.END_IF;
  const curLoop = getCurLoop(ctx);
  if (!curLoop)
    throw new InvalidCommandError(
      `Unexpected ${cmdName} outside of ${
        isIf ? 'IF statement' : 'FOR loop'
      } context`,
      cmd
    );

  // The scope is free for another IF again
  const scopeNode = findParentPorTrNode(node);
  if (scopeNode != null) {
    const tag = tagOf(scopeNode);
    if (tag === WTag.p) ctx.pIfCheckMap.delete(scopeNode);
    else if (tag === WTag.tr) ctx.trIfCheckMap.delete(scopeNode);
  }

  // First time we visit an END-IF node, we assign it the arbitrary name
  // generated when the IF was processed
  if (isIf && !node._ifName) {
    node._ifName = curLoop.varName;
    ctx.gCntEndIf += 1;
  }

  // Is this the END-IF/END-FOR we're expecting? If not:
  // - If it names one of the enclosing loops, the template is malformed.
  // - Otherwise ignore it; an END-IF/END-FOR belonging to an earlier part of
  //   the current loop's paragraph shows up here legitimately.
  const varName = isIf ? node._ifName : cmdRest;
  if (curLoop.varName !== varName) {
    if (ctx.loops.find(o => o.varName === varName) == null) {
      logger.debug(
        `Ignoring ${cmd} (${varName}, but we're expecting ${curLoop.varName})`
      );
      return;
    }
    throw new InvalidCommandError('Invalid command', cmd);
  }
  markCellIfLoopSpansCells(ctx, node, curLoop);

  // Every branch of the IF construct has now been explored: if one of them was
  // selected, run a second pass through the construct to render it.
  if (isIf && curLoop.idx <= EXPLORATION_PASS) {
    curLoop.loopOver =
      (curLoop.ifActiveBranch ?? NO_BRANCH) >= 0
        ? IF_RENDER_ONCE
        : IF_RENDER_NEVER;
  }

  const nextIdx = curLoop.idx + 1;
  const nextItem = curLoop.loopOver[nextIdx];
  if (nextItem != null) {
    // Next iteration: jump back to the node the construct opened on
    ctx.vars[varName] = nextItem;
    ctx.fJump = true;
    curLoop.idx = nextIdx;
    if (isIf) restartIfBranches(curLoop);
  } else {
    ctx.loops.pop();
  }
};
