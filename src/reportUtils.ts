import { logger } from './debug';
import { TemplateParseError } from './errors';
import { WTag } from './ooxml';
import {
  type Context,
  EXPLORATION_PASS,
  type LoopStatus,
  type Node,
  type NonTextNode,
  type TextNode,
} from './types';

// ==========================================
// Nodes and trees
// ==========================================

/** The tag of a node, or `null` if it is a text node (or there is no node). */
export const tagOf = (node: Node | null | undefined): string | null =>
  node == null || node._fTextNode ? null : node._tag;

export const cloneNodeWithoutChildren = (node: Node): Node => {
  if (node._fTextNode) {
    return {
      _children: [],
      _fTextNode: true,
      _text: node._text,
    };
  }
  return {
    _children: [],
    _fTextNode: false,
    _tag: node._tag,
    _attrs: node._attrs,
  };
};

export const getFirstChild = (node: Node): Node | null =>
  node._children[0] ?? null;

/**
 * Where the last `getNextSibling` call on a given parent left off.
 *
 * A node does not know its own index, so finding the next one means searching
 * the parent's children — and visiting all of them in turn is quadratic in
 * their number. `w:body` routinely has tens of thousands of children, which
 * made that the most expensive thing about a long, otherwise trivial document.
 *
 * Children are visited in order, so the answer is nearly always the one right
 * after the previous answer for the same parent. It is one entry per parent
 * rather than one overall because the walk descends into each child before
 * asking for the next one, so a single slot would be evicted every time.
 *
 * The guess is checked by identity before it is used. A miss — an interleaved
 * report, or a tree spliced behind our back, which is exactly what
 * `preprocessTemplate` does — costs a search and nothing else.
 */
const nextIdxByParent = new WeakMap<Node, number>();

export const getNextSibling = (node: Node): Node | null => {
  const parent = node._parent;
  if (parent == null) return null;
  const siblings = parent._children;

  const guess = nextIdxByParent.get(parent);
  const idx =
    guess != null && siblings[guess] === node ? guess : siblings.indexOf(node);
  nextIdxByParent.set(parent, idx + 1);

  if (idx < 0 || idx >= siblings.length - 1) return null;
  return siblings[idx + 1] ?? null;
};

/**
 * Depth-first traversal: the node that comes after `node`, or `null` once the
 * tree has been walked out of. Nodes inserted while traversing are picked up,
 * which is what `preprocessTemplate` relies on.
 */
export const nextNodeInTree = (node: Node): Node | null => {
  const firstChild = getFirstChild(node);
  if (firstChild) return firstChild;
  let curNode: Node | null = node;
  while (curNode != null) {
    const nextSibling = getNextSibling(curNode);
    if (nextSibling) return nextSibling;
    curNode = curNode._parent ?? null;
  }
  return null;
};

export const insertTextSiblingAfter = (textNode: TextNode): TextNode => {
  const tNode = textNode._parent;
  if (tNode == null || tagOf(tNode) !== WTag.t) {
    throw new TemplateParseError(
      'Template syntax error: text node not within w:t'
    );
  }
  const tNodeParent = tNode._parent;
  if (tNodeParent == null)
    throw new TemplateParseError(
      'Template syntax error: w:t node has no parent'
    );
  const idx = tNodeParent._children.indexOf(tNode);
  if (idx < 0) throw new TemplateParseError('Template syntax error');
  const newTNode = cloneNodeWithoutChildren(tNode);
  newTNode._parent = tNodeParent;
  const newTextNode: Node = {
    _parent: newTNode,
    _children: [],
    _fTextNode: true,
    _text: '',
  };
  newTNode._children = [newTextNode];
  tNodeParent._children.splice(idx + 1, 0, newTNode);
  return newTextNode;
};

export const newNonTextNode = (
  tag: string,
  attrs = {},
  children: Array<Node> = []
): NonTextNode => {
  const node: NonTextNode = {
    _fTextNode: false,
    _tag: tag,
    _attrs: attrs,
    _children: children,
  };
  node._children.forEach(child => {
    child._parent = node;
  });
  return node;
};

export const newTextNode = (text: string): TextNode => ({
  _children: [],
  _fTextNode: true,
  _text: text,
});

export const addChild = (parent: Node, child: Node): Node => {
  parent._children.push(child);
  child._parent = parent;
  return child;
};

/** One node, as a single line, for the debug log. */
export const debugPrintNode = (node: Node) =>
  JSON.stringify(
    node._fTextNode
      ? {
          _ifName: node._ifName,
          _fTextNode: node._fTextNode,
          _text: node._text,
        }
      : {
          _ifName: node._ifName,
          _fTextNode: node._fTextNode,
          _tag: node._tag,
          _attrs: node._attrs,
        }
  );

// ==========================================
// Loops
// ==========================================

export const getCurLoop = (ctx: Context): LoopStatus | null =>
  ctx.loops[ctx.loops.length - 1] ?? null;

// Whether we're walking through a branch of an IF construct (IF / ELSE-IF / ELSE)
// that has not been selected. Its contents must not be rendered, in exactly the
// same way as during an exploration pass.
const isIfBranchSuppressed = (loop: LoopStatus) =>
  loop.isIf === true &&
  loop.ifCurrentBranch != null &&
  loop.ifActiveBranch != null &&
  loop.ifCurrentBranch !== loop.ifActiveBranch;

// Whether the contents inside the given loop must be skipped, either because
// the loop is being explored (first pass) or because it's an IF construct and
// the branch being walked is not the selected one.
export const isLoopSkippingOutput = (loop: LoopStatus) =>
  loop.idx <= EXPLORATION_PASS || isIfBranchSuppressed(loop);

export const isLoopExploring = (ctx: Context) => {
  const curLoop = getCurLoop(ctx);
  return curLoop != null && isLoopSkippingOutput(curLoop);
};

export const logLoop = (loops: Array<LoopStatus>) => {
  if (!logger.enabled) return;
  const level = loops.length - 1;
  const curLoop = loops[level];
  if (curLoop == null) return;
  const { varName, idx, loopOver, isIf, ifCurrentBranch, ifActiveBranch } =
    curLoop;
  const idxStr = idx > EXPLORATION_PASS ? idx + 1 : 'EXPLORATION';
  const branchStr = isIf
    ? ` [branch ${ifCurrentBranch}, selected: ${ifActiveBranch}]`
    : '';
  logger.debug(
    `${isIf ? 'IF' : 'FOR'} loop ` +
      `on ${level}:${varName} ` +
      `${idxStr}/${loopOver.length}${branchStr}`
  );
};

// ==========================================
// Paragraphs, rows and cells
// ==========================================

/**
 * The `w:p` a node belongs to — or the `w:tr` around it, when that paragraph
 * sits in a table row. This is the scope an IF construct may not be nested in
 * twice; see `checkNoNestedIfInSameScope`.
 */
export const findParentPorTrNode = (node: Node): Node | null => {
  let parentNode = node._parent;
  while (parentNode != null) {
    if (tagOf(parentNode) === WTag.p) {
      const grandParentNode = parentNode._parent?._parent;
      if (grandParentNode != null && tagOf(grandParentNode) === WTag.tr)
        return grandParentNode;
      return parentNode;
    }
    parentNode = parentNode._parent;
  }
  return null;
};

/**
 * The table cell a node belongs to (the node itself, if it is a cell): loop
 * reference nodes get hoisted up the tree while the loop is being explored, and
 * can end up being the `w:tc` node.
 */
const findCellNode = (node: Node): Node | null => {
  let curNode: Node | null = node;
  while (curNode != null) {
    if (tagOf(curNode) === WTag.tc) return curNode;
    curNode = curNode._parent ?? null;
  }
  return null;
};

/**
 * Flags the cell we're walking when the FOR/IF construct the command belongs to
 * was opened in another cell: such a cell is part of the scaffolding of a
 * multi-cell construct (see "dynamic columns" in the README), and is deleted if
 * it renders to nothing.
 */
export const markCellIfLoopSpansCells = (
  ctx: Context,
  node: Node,
  loop: LoopStatus
) => {
  const cell = ctx.cell;
  if (cell == null) return;
  const cmdCell = findCellNode(node);
  if (cmdCell !== cell.node) return;
  if (findCellNode(loop.refNode) !== cmdCell) cell.fSpansCells = true;
};

/**
 * Whether the commands of the cell we're walking are part of a FOR/IF construct
 * that spans several cells, either because the construct was opened in the cell
 * and is still open now that the cell ends, or because one of its commands
 * referred to a construct opened in another cell.
 */
export const doesCellSpanCells = (ctx: Context): boolean => {
  const cell = ctx.cell;
  if (cell == null) return false;
  return (
    cell.fSpansCells ||
    ctx.loops.some(loop => findCellNode(loop.refNode) === cell.node)
  );
};
