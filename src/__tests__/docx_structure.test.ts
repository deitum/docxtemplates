/**
 * The OOXML structure rules, tested without running the walk.
 *
 * These are the corners where "why did my table lose a column" is decided.
 * Reaching them through `createReport` means building a template that happens
 * to hit the case; here each rule is poked at directly.
 */
import { describe, expect, it } from 'vitest';
import { newContext } from '../context';
import {
  DROP_RULES,
  fillRequiredChildren,
  PENDING_SLOTS,
  TAG_SHOULD_CONTAIN,
} from '../docx/structure';
import { WTag, WpTag } from '../ooxml';
import { resolveOptions } from '../options';
import { newNonTextNode, newTextNode, tagOf } from '../reportUtils';
import { type Context, type Node, type NonTextNode } from '../types';

const ctxWith = (over: Partial<Context> = {}): Context =>
  Object.assign(newContext(resolveOptions({})), over);

const childTags = (node: Node) => node._children.map(child => tagOf(child));

describe('TAG_SHOULD_CONTAIN', () => {
  it('gives an emptied-out table cell a paragraph back', () => {
    const cell = newNonTextNode(WTag.tc);
    fillRequiredChildren(cell);
    expect(childTags(cell)).toEqual([WTag.p]);
    expect(cell._children[0]?._parent).toBe(cell);
  });

  it('leaves a cell that still has a paragraph alone', () => {
    const cell = newNonTextNode(WTag.tc, {}, [newNonTextNode(WTag.p)]);
    fillRequiredChildren(cell);
    expect(childTags(cell)).toEqual([WTag.p]);
  });

  it('accepts an altChunk in place of a paragraph', () => {
    // This is what an HTML command leaves behind.
    const cell = newNonTextNode(WTag.tc, {}, [newNonTextNode(WTag.altChunk)]);
    fillRequiredChildren(cell);
    expect(childTags(cell)).toEqual([WTag.altChunk]);
  });

  it('does not touch tags it says nothing about', () => {
    const row = newNonTextNode(WTag.tr);
    fillRequiredChildren(row);
    expect(row._children).toEqual([]);
  });

  it('only fills with a tag it also accepts', () => {
    // Otherwise the rule would add a child on every pass, forever.
    for (const rule of TAG_SHOULD_CONTAIN) {
      expect(rule.oneOf).toContain(rule.fill);
    }
  });
});

describe('PENDING_SLOTS', () => {
  const slotFor = (tag: string) => {
    const slot = PENDING_SLOTS.find(s => s.tag === tag);
    if (!slot) throw new Error(`no slot for ${tag}`);
    return slot;
  };

  it('has one slot per tag', () => {
    const tags = PENDING_SLOTS.map(s => s.tag);
    expect(new Set(tags).size).toEqual(tags.length);
  });

  it('takes an image, with its caption, and clears it', () => {
    const image = newNonTextNode('image') as NonTextNode;
    const caption = [newNonTextNode(WTag.p) as NonTextNode];
    const ctx = ctxWith({ pendingImageNode: { image, caption } });

    expect(slotFor(WTag.t).take(ctx)).toEqual({ node: image, extra: caption });
    expect(ctx.pendingImageNode).toBeUndefined();
    expect(slotFor(WTag.t).take(ctx)).toBeUndefined();
  });

  it('takes an image without a caption', () => {
    const image = newNonTextNode('image') as NonTextNode;
    const ctx = ctxWith({ pendingImageNode: { image } });
    expect(slotFor(WTag.t).take(ctx)).toEqual({ node: image });
  });

  it('takes a link, replacing the whole run', () => {
    const link = newNonTextNode(WTag.hyperlink) as NonTextNode;
    const ctx = ctxWith({ pendingLinkNode: link });
    expect(slotFor(WTag.r).take(ctx)).toEqual({ node: link });
    expect(ctx.pendingLinkNode).toBeUndefined();
  });

  it('takes an HTML chunk, replacing the whole paragraph', () => {
    const html = newTextNode('<p>hi</p>');
    const ctx = ctxWith({ pendingHtmlNode: html });
    expect(slotFor(WTag.p).take(ctx)).toEqual({ node: html });
    expect(ctx.pendingHtmlNode).toBeUndefined();
  });

  it('takes nothing when no command parked anything', () => {
    const ctx = ctxWith();
    for (const slot of PENDING_SLOTS) expect(slot.take(ctx)).toBeUndefined();
  });
});

describe('DROP_RULES', () => {
  const keeps = (tag: keyof typeof DROP_RULES, nodeIn: Node, nodeOut: Node) =>
    DROP_RULES[tag].keep({ nodeIn, nodeOut, ctx: ctxWith() });

  describe('paragraph', () => {
    it('keeps one that anchors an inline drawing', () => {
      const p = newNonTextNode(WTag.p, {}, [
        newNonTextNode(WTag.r, {}, [
          newNonTextNode(WTag.drawing, {}, [newNonTextNode(WpTag.inline)]),
        ]),
      ]);
      expect(keeps(WTag.p, p, p)).toBe(true);
    });

    it('keeps one that anchors a floating drawing', () => {
      const p = newNonTextNode(WTag.p, {}, [newNonTextNode(WpTag.anchor)]);
      expect(keeps(WTag.p, p, p)).toBe(true);
    });

    it('drops one that holds nothing', () => {
      const p = newNonTextNode(WTag.p, {}, [newNonTextNode(WTag.r)]);
      expect(keeps(WTag.p, p, p)).toBe(false);
    });
  });

  describe('table row', () => {
    it('keeps a row wrapping exactly one nested row', () => {
      const row = newNonTextNode(WTag.tr, {}, [newNonTextNode(WTag.tr)]);
      expect(keeps(WTag.tr, row, row)).toBe(true);
    });

    it('drops a row with no nested row', () => {
      const row = newNonTextNode(WTag.tr, {}, [newNonTextNode(WTag.tc)]);
      expect(keeps(WTag.tr, row, row)).toBe(false);
    });

    it('drops a row with more than one nested row', () => {
      const row = newNonTextNode(WTag.tr, {}, [
        newNonTextNode(WTag.tr),
        newNonTextNode(WTag.tr),
      ]);
      expect(keeps(WTag.tr, row, row)).toBe(false);
    });
  });

  describe('table cell', () => {
    it('keeps a cell whose construct opened and closed inside it', () => {
      // `doesCellSpanCells` is false for a fresh context: nothing marked the
      // cell, and no loop is open.
      const cell = newNonTextNode(WTag.tc);
      expect(keeps(WTag.tc, cell, cell)).toBe(true);
    });

    it('keeps a cell holding a nested table even when the construct spans cells', () => {
      const cell = newNonTextNode(WTag.tc, {}, [newNonTextNode(WTag.tbl)]);
      const ctx = ctxWith({ cell: { node: cell, fSpansCells: true } });
      expect(
        DROP_RULES[WTag.tc].keep({ nodeIn: cell, nodeOut: cell, ctx })
      ).toBe(true);
    });

    it('drops a cell that is scaffolding of a multi-cell construct', () => {
      const cell = newNonTextNode(WTag.tc);
      const ctx = ctxWith({ cell: { node: cell, fSpansCells: true } });
      expect(
        DROP_RULES[WTag.tc].keep({ nodeIn: cell, nodeOut: cell, ctx })
      ).toBe(false);
    });
  });

  it('explains itself', () => {
    // The reasons are the documentation of this corner of the engine; an empty
    // one means a rule was added without saying why.
    for (const rule of Object.values(DROP_RULES)) {
      expect(rule.name).not.toEqual('');
      expect(rule.reason.length).toBeGreaterThan(40);
    }
  });
});
