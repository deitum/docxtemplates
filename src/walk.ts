/**
 * The template walk.
 *
 * `walkTemplate` moves a cursor over the input tree and builds a second tree as
 * it goes. Commands can make it jump backwards (that is how FOR loops repeat),
 * and can make nodes it already emitted disappear again (that is how a
 * paragraph holding nothing but `+++END-FOR+++` vanishes from the report).
 */
import { type CommandProcessor, processCmd } from './commands/execute';
import {
  DROP_RULES,
  fillRequiredChildren,
  PENDING_SLOTS,
} from './docx/structure';
import { BUFFER_TAGS, DrawAttr, isBufferTag, VTag, WpTag, WTag } from './ooxml';
import { DEFAULT_MAXIMUM_WALKING_DEPTH } from './options';
import { logger } from './debug';
import {
  IncompleteConditionalStatementError,
  InternalError,
  UnterminatedForLoopError,
} from './errors';
import {
  cloneNodeWithoutChildren,
  debugPrintNode,
  getCurLoop,
  getFirstChild,
  getNextSibling,
  isLoopExploring,
  tagOf,
} from './reportUtils';
import {
  type Context,
  type Htmls,
  type Images,
  type Links,
  type LoopStatus,
  type Node,
  type NonTextNode,
  type ReportData,
  type TextNode,
} from './types';

/** How the cursor got to the node it is on. */
enum Move {
  jump = 'JUMP',
  down = 'DOWN',
  side = 'SIDE',
  up = 'UP',
}

type ReportOutput =
  | {
      status: 'success';
      report: Node;
      images: Images;
      links: Links;
      htmls: Htmls;
    }
  | {
      status: 'errors';
      errors: Error[];
    };

export async function produceJsReport(
  data: ReportData | undefined,
  template: Node,
  ctx: Context
): Promise<ReportOutput> {
  return walkTemplate(data, template, ctx, processCmd);
}

const isTextNodeInsideWt = (node: Node): node is TextNode =>
  node._fTextNode && tagOf(node._parent) === WTag.t;

export async function walkTemplate(
  data: ReportData | undefined,
  template: Node,
  ctx: Context,
  processor: CommandProcessor
): Promise<ReportOutput> {
  const out: Node = cloneNodeWithoutChildren(template);
  let nodeIn: Node = template;
  let nodeOut: Node = out;
  let move: Move | undefined;
  const errors: Error[] = [];

  let loopCount = 0;
  const maximumWalkingDepth =
    ctx.options.maximumWalkingDepth || DEFAULT_MAXIMUM_WALKING_DEPTH;

  for (;;) {
    const curLoop = getCurLoop(ctx);

    // 1. Move the input cursor. `null` means the walk is over.
    const step = advance({
      nodeIn,
      ctx,
      previousMove: move,
      curLoop,
      loopCount,
      maximumWalkingDepth,
    });
    if (step == null) break;
    ({ node: nodeIn, move } = step);

    if (logger.enabled)
      logger.debug(
        `Next node [${move}, level ${ctx.walk.level}]`,
        debugPrintNode(nodeIn)
      );

    // 2. The node we just finished may have to be dropped from the output.
    if (move !== Move.down) dropDeadOutputNode(nodeIn, nodeOut, ctx);

    // 3. Apply the move to the output tree, and run the phase it triggers.
    if (move === Move.up) {
      nodeOut = moveOutputUp(nodeIn, nodeOut, ctx, curLoop);
    } else if (move === Move.down || move === Move.side) {
      nodeOut = await appendOutputNode(
        data,
        nodeIn,
        nodeOut,
        ctx,
        move,
        processor,
        errors
      );
    } else {
      // JUMP: climb the output tree back to the level the loop opened at.
      for (let climb = step.deltaJump; climb > 0; climb -= 1) {
        if (nodeOut._parent == null)
          throw new InternalError('node parent is null');
        nodeOut = nodeOut._parent;
      }
    }

    loopCount++;
  }

  collectUnterminatedConstructErrors(ctx, errors);

  if (errors.length > 0) return { status: 'errors', errors };
  return {
    status: 'success',
    report: out,
    images: ctx.resources.images,
    links: ctx.resources.links,
    htmls: ctx.resources.htmls,
  };
}

// ==========================================
// The phases of one step
// ==========================================

type Step = { node: Node; move: Move; deltaJump: number };

/**
 * Picks the next input node: back to a loop's reference node if a command asked
 * for a jump, otherwise down, sideways or up, in that order. Returns `null`
 * when the root has been walked out of.
 */
function advance({
  nodeIn,
  ctx,
  previousMove,
  curLoop,
  loopCount,
  maximumWalkingDepth,
}: {
  nodeIn: Node;
  ctx: Context;
  previousMove: Move | undefined;
  curLoop: LoopStatus | null;
  loopCount: number;
  maximumWalkingDepth: number;
}): Step | null {
  if (ctx.walk.jumpRequested) {
    if (!curLoop) throw new InternalError('jumping while curLoop is null');
    if (logger.enabled)
      logger.debug(
        `Jumping to level ${curLoop.refNodeLevel}...`,
        debugPrintNode(curLoop.refNode)
      );
    const deltaJump = ctx.walk.level - curLoop.refNodeLevel;
    ctx.walk.level = curLoop.refNodeLevel;
    ctx.walk.jumpRequested = false;
    return { node: curLoop.refNode, move: Move.jump, deltaJump };
  }

  // Down — but not if we have just come up, or we would walk in circles.
  if (previousMove !== Move.up) {
    const firstChild = getFirstChild(nodeIn);
    if (firstChild) {
      ctx.walk.level += 1;
      return { node: firstChild, move: Move.down, deltaJump: 0 };
    }
  }

  const nextSibling = getNextSibling(nodeIn);
  if (nextSibling) return { node: nextSibling, move: Move.side, deltaJump: 0 };

  const parent = nodeIn._parent;
  if (parent == null) {
    logger.debug(`=== parent is null, breaking after ${loopCount} loops...`);
    return null;
  }
  if (loopCount > maximumWalkingDepth) {
    // Emergency exit, in case a template manages to make the walk loop forever
    if (logger.enabled)
      logger.debug(
        `=== parent is still not null after ${loopCount} loops, something must be wrong ...`,
        debugPrintNode(parent)
      );
    throw new InternalError(
      'infinite loop or massive dataset detected. Please review and try again'
    );
  }
  ctx.walk.level -= 1;
  return { node: parent, move: Move.up, deltaJump: 0 };
}

/**
 * Removes the output node we have just finished building, in the cases where it
 * should never have been emitted: nodes produced while exploring a loop, and
 * paragraphs, rows or cells that held nothing but commands.
 *
 * Which of the latter survive anyway is stated in `DROP_RULES`, not here.
 */
function dropDeadOutputNode(nodeIn: Node, nodeOut: Node, ctx: Context): void {
  const tag = tagOf(nodeOut);
  if (!isBufferTag(tag) && tag !== WTag.tbl) return;

  let fRemoveNode = false;
  if (isLoopExploring(ctx)) {
    // Nothing generated during an exploration pass belongs in the output.
    fRemoveNode = true;
  } else if (isBufferTag(tag)) {
    const buffers = ctx.walk.buffers[tag];
    const heldOnlyCommands =
      buffers.text === '' && buffers.cmds !== '' && !buffers.hasInsertedText;
    const rule = DROP_RULES[tag];
    fRemoveNode = heldOnlyCommands && !rule.keep({ nodeIn, nodeOut, ctx });
    if (heldOnlyCommands && !fRemoveNode && logger.enabled)
      logger.debug(`Keeping empty ${tag}: ${rule.name}`);
  }

  // The node leaves the output, but keeps its parent link, so that the walk can
  // still move up the tree through it.
  if (fRemoveNode && nodeOut._parent != null) nodeOut._parent._children.pop();
}

/**
 * Moves the output cursor up one level, and finishes off the node being left
 * behind: splicing in whatever an IMAGE/LINK/HTML command produced, and keeping
 * table cells structurally valid.
 */
function moveOutputUp(
  nodeIn: Node,
  nodeOut: Node,
  ctx: Context,
  curLoop: LoopStatus | null
): Node {
  // Loop exploring? Update the reference node for the current loop
  if (isLoopExploring(ctx) && curLoop && nodeIn === curLoop.refNode._parent) {
    curLoop.refNode = nodeIn;
    curLoop.refNodeLevel -= 1;
    if (logger.enabled)
      logger.debug(
        `Updated loop '${curLoop.varName}' refNode: ` + debugPrintNode(nodeIn)
      );
  }

  const nodeOutParent = nodeOut._parent;
  if (nodeOutParent == null) throw new InternalError('node parent is null');
  nodeOut = nodeOutParent;

  const tag = tagOf(nodeOut);

  // Splice in whatever an IMAGE/LINK/HTML command parked for this level.
  for (const slot of PENDING_SLOTS) {
    if (tag !== slot.tag) continue;
    const pending = ctx.resources.takePending(slot.kind);
    if (pending != null)
      replaceOutputNode(nodeOut, ctx, pending.node, pending.extra);
  }

  fillRequiredChildren(nodeOut);

  // Remember the last `w:rPr` seen, so that a LINK can inherit its formatting
  if (tag === WTag.rPr) ctx.resources.textRunProps = nodeOut as NonTextNode;
  if (tagOf(nodeIn) === WTag.r) ctx.resources.textRunProps = undefined;

  return nodeOut;
}

/**
 * Swaps the node the cursor is leaving for the one a command produced, and
 * marks the enclosing paragraph/row/cell as non-empty so that it survives.
 */
function replaceOutputNode(
  nodeOut: Node,
  ctx: Context,
  replacement: Node,
  extra?: Node[]
): void {
  const parent = nodeOut._parent;
  if (!parent) return;
  replacement._parent = parent;
  parent._children.pop();
  parent._children.push(replacement);
  if (extra) parent._children.push(...extra);
  for (const key of BUFFER_TAGS) ctx.walk.buffers[key].hasInsertedText = true;
}

/**
 * Copies the input node into the output tree and, when it is a text node,
 * processes the commands it contains. Note that being copied is no guarantee of
 * survival: `dropDeadOutputNode` may remove it again later.
 */
async function appendOutputNode(
  data: ReportData | undefined,
  nodeIn: Node,
  nodeOut: Node,
  ctx: Context,
  move: Move.down | Move.side,
  processor: CommandProcessor,
  errors: Error[]
): Promise<Node> {
  // Point at the new node's parent
  if (move === Move.side) {
    if (nodeOut._parent == null) throw new InternalError('node parent is null');
    nodeOut = nodeOut._parent;
  }

  // Reset the buffers when a new `w:p`, `w:tr` or `w:tc` starts
  const tag = tagOf(nodeIn);
  if (isBufferTag(tag)) {
    ctx.walk.buffers[tag] = { text: '', cmds: '', hasInsertedText: false };
    if (tag === WTag.tc) ctx.walk.cell = { node: nodeIn, spansCells: false };
  }

  const newNode: Node = cloneNodeWithoutChildren(nodeIn);
  newNode._parent = nodeOut;
  nodeOut._children.push(newNode);

  // Renumber shapes, so that copies made by a FOR loop don't share an id
  if (!isLoopExploring(ctx) && (tag === WpTag.docPr || tag === VTag.shape)) {
    if (logger.enabled) logger.debug('detected a - ', debugPrintNode(newNode));
    assignNewShapeId(newNode as NonTextNode, ctx);
  }

  if (isTextNodeInsideWt(nodeIn)) {
    const result = await processText(data, nodeIn, ctx, processor);
    if (typeof result === 'string') {
      (newNode as TextNode)._text = result;
      if (logger.enabled)
        logger.debug(
          `Inserted command result string into node. Updated node: ` +
            debugPrintNode(newNode)
        );
    } else {
      errors.push(...result);
    }
  }

  return newNode;
}

function collectUnterminatedConstructErrors(ctx: Context, errors: Error[]) {
  const report = (err: Error) => {
    if (ctx.options.failFast) throw err;
    errors.push(err);
  };

  if (ctx.walk.openIfCount !== ctx.walk.closedIfCount) {
    report(new IncompleteConditionalStatementError());
  }
  const { loops } = ctx.scope;
  const innermostLoop = loops[loops.length - 1];
  if (innermostLoop != null && loops.some(l => !l.isIf)) {
    report(new UnterminatedForLoopError(innermostLoop));
  }
}

// ==========================================
// Text and command extraction
// ==========================================

/**
 * Splits a text node on the command delimiters and alternates between copying
 * text to the output and feeding commands to `onCommand`, whose result (if any)
 * is inserted in the command's place.
 */
const processText = async (
  data: ReportData | undefined,
  node: TextNode,
  ctx: Context,
  onCommand: CommandProcessor
): Promise<string | Error[]> => {
  const { cmdDelimiter, failFast } = ctx.options;
  const text = node._text;
  if (text == null || text === '') return '';
  const segments = text
    .split(cmdDelimiter[0])
    .flatMap(s => s.split(cmdDelimiter[1]));

  let outText = '';
  const errors: Error[] = [];
  for (let idx = 0; idx < segments.length; idx++) {
    // Include the separators in the buffers, so that a paragraph holding only a
    // command is recognised as such
    if (idx > 0)
      appendTextToTagBuffers(cmdDelimiter[0], ctx, { isCommand: true });

    // Append the segment either to the command being collected or to the output
    const segment = segments[idx] ?? '';
    if (ctx.walk.isCollectingCommand) ctx.walk.command += segment;
    else if (!isLoopExploring(ctx)) outText += segment;
    appendTextToTagBuffers(segment, ctx, {
      isCommand: ctx.walk.isCollectingCommand,
    });

    // A delimiter follows: run the command if one was being collected, then
    // toggle between "command" and "text" mode
    if (idx < segments.length - 1) {
      if (ctx.walk.isCollectingCommand) {
        const cmdResultText = await onCommand(data, node, ctx);
        if (cmdResultText != null) {
          if (typeof cmdResultText === 'string') {
            outText += cmdResultText;
            appendTextToTagBuffers(cmdResultText, ctx, {
              isCommand: false,
              hasInsertedText: true,
            });
          } else {
            if (failFast) throw cmdResultText;
            errors.push(cmdResultText);
          }
        }
      }
      ctx.walk.isCollectingCommand = !ctx.walk.isCollectingCommand;
    }
  }
  if (errors.length > 0) return errors;
  return outText;
};

const appendTextToTagBuffers = (
  text: string,
  ctx: Context,
  options: { isCommand?: boolean; hasInsertedText?: boolean }
) => {
  const { isCommand, hasInsertedText } = options;
  const type = isCommand ? 'cmds' : 'text';
  for (const key of BUFFER_TAGS) {
    const buf = ctx.walk.buffers[key];
    buf[type] += text;
    if (hasInsertedText) buf.hasInsertedText = true;
  }
};

function assignNewShapeId(newNode: NonTextNode, ctx: Context) {
  newNode._attrs = {
    ...newNode._attrs,
    [DrawAttr.id]: String(ctx.resources.nextShapeId()),
  };
}
