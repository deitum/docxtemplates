/**
 * What each command *is*, in one table.
 *
 * Knowledge about a command used to be spread over six places — the enum, the
 * list of built-in names, the list of commands whose payload is an expression,
 * the prefix chain, the dispatch `switch`, and a `isLoopExploring` guard
 * repeated inside five of its branches. Adding a command meant finding all six.
 *
 * Here a command is one entry: whether operator aliases are substituted in its
 * payload, whether it runs during a loop's exploration pass, and what it does.
 * `execute.ts` is then a dispatcher that knows about none of them in
 * particular, and `parse.ts` derives what it needs instead of restating it.
 */
import { logger } from '../debug';
import {
  ImageError,
  InvalidCommandError,
  isError,
  ObjectCommandResultError,
} from '../errors';
import { runUserJsAndGetRaw } from '../jsSandbox';
import { WTag, XML_SPACE_PRESERVE, XmlAttr } from '../ooxml';
import {
  findParentPorTrNode,
  getCurLoop,
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

/** Everything a command handler is given. */
export type CommandArgs = {
  data: ReportData | undefined;
  /** The text node the command was found in. */
  node: Node;
  ctx: Context;
  /** The whole command, as it appears in error messages. */
  cmd: string;
  /** The command name, upper-cased. */
  name: string;
  /** Whatever followed the name. */
  rest: string;
};

export type CommandSpec = {
  /**
   * Whether the payload is a JS expression, and hence the only thing operator
   * aliases (`больше` for `>`, say) are substituted in.
   */
  isExpression: boolean;
  /**
   * Whether to skip the command while a loop is being explored — walked once
   * without rendering, to find out whether it is empty and which branch of an
   * IF construct applies. Control-flow commands must still run then; commands
   * that produce output must not.
   */
  skipWhileExploring: boolean;
  /** The text to insert in the command's place, if any. */
  run: (args: CommandArgs) => Promise<string | void> | string | void;
};

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

// ==========================================
// ALIAS, INS, EXEC
// ==========================================

// ALIAS <name> <the command it stands for>
const processAlias = ({ ctx, cmd, rest }: CommandArgs): void => {
  const aliasMatch = /^(\S+)\s+(.+)/.exec(rest);
  const [, aliasName, fullCmd] = aliasMatch ?? [];
  if (aliasName == null || fullCmd == null)
    throw new InvalidCommandError('Invalid ALIAS command', cmd);
  ctx.scope.shorthands[aliasName] = fullCmd;
  logger.debug(`Defined alias '${aliasName}' for: ${fullCmd}`);
};

// INS <expression>
const processIns = async ({
  data,
  ctx,
  rest,
}: CommandArgs): Promise<string> => {
  let result = await runUserJsAndGetRaw(data, rest, ctx);
  if (result == null) return '';

  // An object would stringify to '[object Object]', which is never what the
  // template author meant.
  if (typeof result === 'object' && !Array.isArray(result)) {
    const err = new ObjectCommandResultError(rest, result);
    if (ctx.options.errorHandler == null) throw err;
    result = await ctx.options.errorHandler(err, rest);
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

// EXEC <code>
const processExec = async ({ data, ctx, rest }: CommandArgs): Promise<void> => {
  await runUserJsAndGetRaw(data, rest, ctx);
};

// ==========================================
// IMAGE, LINK, HTML
// ==========================================

// IMAGE <expression>
const processImageCmd = async ({ data, ctx, cmd, rest }: CommandArgs) => {
  const img: ImagePars | undefined = await runUserJsAndGetRaw(data, rest, ctx);
  if (img == null) return;
  try {
    processImage(ctx, img);
  } catch (e) {
    if (!isError(e)) throw e;
    throw new ImageError(e, cmd);
  }
};

// LINK <expression>
const processLinkCmd = async ({ data, ctx, rest }: CommandArgs) => {
  const pars: LinkPars | undefined = await runUserJsAndGetRaw(data, rest, ctx);
  if (pars != null) processLink(ctx, pars);
};

// HTML <expression>
const processHtmlCmd = async ({ data, ctx, rest }: CommandArgs) => {
  const html: string | undefined = await runUserJsAndGetRaw(data, rest, ctx);
  if (html != null) processHtml(ctx, html);
};

// ==========================================
// FOR / IF / ELSE-IF / ELSE / END-FOR / END-IF
// ==========================================

// FOR <varName> IN <expression>
// IF <expression>
const processForIf = async ({
  data,
  node,
  ctx,
  cmd,
  name,
  rest,
}: CommandArgs): Promise<void> => {
  const isIf = name === Command.IF;

  // Identify the FOR/IF loop. IF constructs have no variable of their own, so
  // they get an arbitrary name that the matching END-IF picks up.
  let forMatch: RegExpExecArray | null = null;
  let varName: string;
  if (isIf) {
    if (!node._ifName) {
      node._ifName = `${IF_VAR_PREFIX}${ctx.walk.openIfCount}`;
      ctx.walk.openIfCount += 1;
    }
    varName = node._ifName;
  } else {
    forMatch = /^(\S+)\s+IN\s+(.+)/i.exec(rest);
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
    logLoop(ctx.scope.loops);
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
    ifBranchTaken = !!(await runUserJsAndGetRaw(data, rest, ctx));
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

  ctx.scope.loops.push({
    refNode: node,
    refNodeLevel: ctx.walk.level,
    varName,
    loopOver,
    isIf,
    // Run through the loop once first without outputting anything; otherwise
    // empty loops could not be detected.
    idx: EXPLORATION_PASS,
    ...(isIf ? { ifCurrentBranch: 0, ifActiveBranch, ifBranchTaken } : {}),
  });
  logLoop(ctx.scope.loops);
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
      ? ctx.walk.ifByParagraph
      : tag === WTag.tr
        ? ctx.walk.ifByTableRow
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
const processElse = async ({
  data,
  node,
  ctx,
  cmd,
  name,
  rest,
}: CommandArgs): Promise<void> => {
  const isElseIf = name === Command.ELSE_IF;
  const curLoop = getCurLoop(ctx);
  if (!curLoop || !curLoop.isIf)
    throw new InvalidCommandError(
      `Unexpected ${name} outside of IF statement context`,
      cmd
    );
  if (isElseIf && !rest)
    throw new InvalidCommandError(
      'Invalid ELSE-IF command (missing condition)',
      cmd
    );
  if (curLoop.ifElseBranch != null)
    throw new InvalidCommandError(
      `Unexpected ${name} after an ELSE command`,
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
      const ifLoop = ctx.scope.loops.pop();
      try {
        shouldRun = !!(await runUserJsAndGetRaw(data, rest, ctx));
      } finally {
        if (ifLoop) ctx.scope.loops.push(ifLoop);
      }
    }
    if (shouldRun) {
      curLoop.ifBranchTaken = true;
      curLoop.ifActiveBranch = branch;
    }
  }
  logLoop(ctx.scope.loops);
};

// END-FOR [varName]
// END-IF
const processEndForIf = ({ node, ctx, cmd, name, rest }: CommandArgs): void => {
  const isIf = name === Command.END_IF;
  const curLoop = getCurLoop(ctx);
  if (!curLoop)
    throw new InvalidCommandError(
      `Unexpected ${name} outside of ${
        isIf ? 'IF statement' : 'FOR loop'
      } context`,
      cmd
    );

  // The scope is free for another IF again
  const scopeNode = findParentPorTrNode(node);
  if (scopeNode != null) {
    const tag = tagOf(scopeNode);
    if (tag === WTag.p) ctx.walk.ifByParagraph.delete(scopeNode);
    else if (tag === WTag.tr) ctx.walk.ifByTableRow.delete(scopeNode);
  }

  // First time we visit an END-IF node, we assign it the arbitrary name
  // generated when the IF was processed
  if (isIf && !node._ifName) {
    node._ifName = curLoop.varName;
    ctx.walk.closedIfCount += 1;
  }

  // Is this the END-IF/END-FOR we're expecting? If not:
  // - If it names one of the enclosing loops, the template is malformed.
  // - Otherwise ignore it; an END-IF/END-FOR belonging to an earlier part of
  //   the current loop's paragraph shows up here legitimately.
  const varName = isIf ? node._ifName : rest;
  if (curLoop.varName !== varName) {
    if (ctx.scope.loops.find(o => o.varName === varName) == null) {
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
    ctx.scope.vars[varName] = nextItem;
    ctx.walk.jumpRequested = true;
    curLoop.idx = nextIdx;
    if (isIf) restartIfBranches(curLoop);
  } else {
    ctx.scope.loops.pop();
  }
};

// ==========================================
// The table
// ==========================================

/** Commands that are consumed elsewhere and do nothing when dispatched. */
const noop = () => {};

export const COMMANDS: Record<Command, CommandSpec> = {
  // Read before rendering starts, by the compile-time query scan.
  [Command.QUERY]: {
    isExpression: false,
    skipWhileExploring: false,
    run: noop,
  },
  // Scaffolding left in emptied-out text nodes by `preprocessTemplate`, not
  // something a template author writes.
  [Command.CMD_NODE]: {
    isExpression: false,
    skipWhileExploring: false,
    run: noop,
  },
  [Command.ALIAS]: {
    isExpression: false,
    skipWhileExploring: false,
    run: processAlias,
  },
  [Command.FOR]: {
    isExpression: true,
    skipWhileExploring: false,
    run: processForIf,
  },
  [Command.END_FOR]: {
    isExpression: false,
    skipWhileExploring: false,
    run: processEndForIf,
  },
  [Command.IF]: {
    isExpression: true,
    skipWhileExploring: false,
    run: processForIf,
  },
  [Command.ELSE_IF]: {
    isExpression: true,
    skipWhileExploring: false,
    run: processElse,
  },
  [Command.ELSE]: {
    isExpression: false,
    skipWhileExploring: false,
    run: processElse,
  },
  [Command.END_IF]: {
    isExpression: false,
    skipWhileExploring: false,
    run: processEndForIf,
  },
  [Command.INS]: {
    isExpression: true,
    skipWhileExploring: true,
    run: processIns,
  },
  [Command.EXEC]: {
    isExpression: true,
    skipWhileExploring: true,
    run: processExec,
  },
  [Command.IMAGE]: {
    isExpression: true,
    skipWhileExploring: true,
    run: processImageCmd,
  },
  [Command.LINK]: {
    isExpression: true,
    skipWhileExploring: true,
    run: processLinkCmd,
  },
  [Command.HTML]: {
    isExpression: true,
    skipWhileExploring: true,
    run: processHtmlCmd,
  },
};

/**
 * Lookup by name. A `Map` rather than the record itself, so that the inherited
 * members of `Object.prototype` — `constructor`, `toString`, ... — cannot be
 * mistaken for commands; the `switch` this replaced could not be fooled that
 * way, and neither should this be.
 */
const SPECS = new Map<string, CommandSpec>(Object.entries(COMMANDS));

/** The spec for a command name, or `undefined` if it names no command. */
export const specOf = (name: string): CommandSpec | undefined =>
  SPECS.get(name);

/**
 * The commands whose payload is a JS expression — derived, so that a new
 * command cannot be forgotten here.
 */
export const EXPRESSION_COMMANDS: ReadonlySet<string> = new Set(
  Object.entries(COMMANDS)
    .filter(([, spec]) => spec.isExpression)
    .map(([name]) => name)
);
