/**
 * The rules that keep the generated document something Word will open, stated
 * as data rather than as control flow.
 *
 * Three separate things used to be tangled into the walk: which tags may not be
 * left empty and what to fill them with, where a node built by IMAGE / LINK /
 * HTML gets spliced in, and when a paragraph, row or cell that rendered to
 * nothing should disappear. They were spread through `moveOutputUp` and a
 * four-branch `dropDeadOutputNode`, where none of them could be read on its own,
 * let alone tested on its own.
 *
 * Each rule carries the reason it exists: these are the corners where "why does
 * my table lose a column" gets decided, and the answer should be legible.
 */
import { WTag, WpTag } from '../ooxml';
import { doesCellSpanCells, tagOf } from '../reportUtils';
import { type BufferTag } from '../ooxml';
import { type PendingKind } from '../resources';
import { type Context, type Node } from '../types';

// ==========================================
// Tags that may not be left empty
// ==========================================

export type ContainmentRule = {
  /** The tag being closed. */
  tag: string;
  /** It must have at least one child with one of these tags... */
  oneOf: readonly string[];
  /** ...or one of these is added. */
  fill: string;
};

/**
 * Word refuses to open a document with a table cell that holds no block-level
 * content, which is what is left behind when the cell's only paragraph was
 * removed for holding nothing but commands.
 */
export const TAG_SHOULD_CONTAIN: readonly ContainmentRule[] = [
  { tag: WTag.tc, oneOf: [WTag.p, WTag.altChunk], fill: WTag.p },
];

/** Applies {@link TAG_SHOULD_CONTAIN} to a node that has just been finished. */
export function fillRequiredChildren(node: Node): void {
  const tag = tagOf(node);
  for (const rule of TAG_SHOULD_CONTAIN) {
    if (tag !== rule.tag) continue;
    if (node._children.some(child => rule.oneOf.includes(tagOf(child) ?? '')))
      continue;
    node._children.push({
      _parent: node,
      _children: [],
      _fTextNode: false,
      _tag: rule.fill,
      _attrs: {},
    });
  }
}

// ==========================================
// Where a command's node is spliced in
// ==========================================

export type PendingSlot = {
  /** The output node this replaces, identified by the tag being closed. */
  tag: string;
  /** Which command's parked node belongs at this level. */
  kind: PendingKind;
};

/**
 * IMAGE, LINK and HTML do not produce text: they build a node and park it,
 * because it can only be spliced in once the element it replaces is finished.
 * Each replaces a different level — an image stands in for the `w:t`, a link for
 * the whole `w:r`, an HTML chunk for the `w:p`.
 */
export const PENDING_SLOTS: readonly PendingSlot[] = [
  { tag: WTag.t, kind: 'image' },
  { tag: WTag.r, kind: 'link' },
  { tag: WTag.p, kind: 'html' },
];

// ==========================================
// Nodes that rendered to nothing
// ==========================================

export type DropRule = {
  /** For the debug log, and for reading this table. */
  name: string;
  /** Why a node this applies to survives despite looking empty. */
  reason: string;
  /** Whether the node must be kept after all. */
  keep: (args: {
    /** The input node the output was built from. */
    nodeIn: Node;
    /** The output node about to be dropped. */
    nodeOut: Node;
    ctx: Context;
  }) => boolean;
};

const hasDrawing = (node: Node): boolean => {
  const tag = tagOf(node);
  if (tag === WpTag.anchor || tag === WpTag.inline || tag === WTag.drawing)
    return true;
  return node._children.some(hasDrawing);
};

const hasChildTagged = (node: Node, tag: string): boolean =>
  node._children.some(child => tagOf(child) === tag);

/**
 * A paragraph, row or cell holding nothing but commands is removed — that is
 * what makes the line carrying `+++END-FOR+++` vanish from the report. These
 * are the cases where it must stay anyway, one per tag.
 */
export const DROP_RULES: { readonly [tag in BufferTag]: DropRule } = {
  [WTag.p]: {
    name: 'paragraph anchoring a drawing',
    reason:
      'A paragraph with no text still carries content when a drawing is ' +
      'anchored to it; removing it takes the image with it.',
    keep: ({ nodeOut }) => hasDrawing(nodeOut),
  },
  [WTag.tr]: {
    name: 'row wrapping a nested table',
    reason:
      'A row whose only content is a single nested row is the wrapper of a ' +
      'nested table, not an empty row.',
    keep: ({ nodeIn }) =>
      nodeIn._children.filter(child => tagOf(child) === WTag.tr).length === 1,
  },
  [WTag.tc]: {
    name: 'cell that is not part of a multi-cell construct',
    reason:
      'A cell is only removed when the commands in it belong to a FOR/IF ' +
      'construct spanning several cells (the "dynamic columns" pattern). A ' +
      'construct opening and closing inside one cell leaves the cell empty ' +
      'but present, since removing it would shift the rest of the row into ' +
      'the wrong columns. A cell holding a nested table always stays.',
    keep: ({ nodeOut, ctx }) =>
      !doesCellSpanCells(ctx) || hasChildTagged(nodeOut, WTag.tbl),
  },
};
