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
import { newNonTextNode, tagOf } from '../reportUtils';
import { type Context, type Node } from '../types';

const ctxWith = (): Context => newContext(resolveOptions({}));

/** A context walking `cell`, whose commands span more than one cell. */
const ctxWithCell = (cell: Node): Context => {
  const ctx = ctxWith();
  ctx.walk.cell = { node: cell, spansCells: true };
  return ctx;
};

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
  const kindFor = (tag: string) => {
    const slot = PENDING_SLOTS.find(s => s.tag === tag);
    if (!slot) throw new Error(`no slot for ${tag}`);
    return slot.kind;
  };

  it('has one slot per tag', () => {
    const tags = PENDING_SLOTS.map(s => s.tag);
    expect(new Set(tags).size).toEqual(tags.length);
  });

  it('takes an image, with its caption, and clears it', () => {
    const ctx = ctxWith();
    const image = newNonTextNode('image');
    const caption = [newNonTextNode(WTag.p)];
    ctx.resources.park('image', { node: image, extra: caption });

    expect(ctx.resources.takePending(kindFor(WTag.t))).toEqual({
      node: image,
      extra: caption,
    });
    expect(ctx.resources.takePending(kindFor(WTag.t))).toBeUndefined();
  });

  it('takes an image without a caption', () => {
    const ctx = ctxWith();
    const image = newNonTextNode('image');
    ctx.resources.park('image', { node: image });
    expect(ctx.resources.takePending(kindFor(WTag.t))).toEqual({ node: image });
  });

  it('puts a link at the run and an HTML chunk at the paragraph', () => {
    expect(kindFor(WTag.r)).toEqual('link');
    expect(kindFor(WTag.p)).toEqual('html');
  });

  it('takes nothing when no command parked anything', () => {
    const ctx = ctxWith();
    for (const slot of PENDING_SLOTS) {
      expect(ctx.resources.takePending(slot.kind)).toBeUndefined();
    }
  });
});

describe('Resources', () => {
  it('hands out one relationship id per resource, by kind', () => {
    const { resources } = ctxWith();
    expect(resources.addLink('https://a.test')).toEqual('link1');
    expect(resources.addLink('https://b.test')).toEqual('link2');
    expect(resources.addHtml('<p>x</p>')).toEqual('html1');
    expect(resources.links).toEqual({
      link1: { url: 'https://a.test' },
      link2: { url: 'https://b.test' },
    });
    expect(resources.htmls).toEqual({ html1: '<p>x</p>' });
  });

  it('continues shape numbering where the template left off', () => {
    // Images and shapes share one counter: both become `wp:docPr` ids, and a
    // generated one colliding with a template's own would corrupt the file.
    const ctx = newContext(resolveOptions({}), 7);
    expect(ctx.resources.lastShapeId).toEqual(7);
    expect(ctx.resources.addImage({ data: 'x', extension: '.png' })).toEqual(
      'img8'
    );
    expect(ctx.resources.nextShapeId()).toEqual(9);
    expect(ctx.resources.lastShapeId).toEqual(9);
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
      const ctx = ctxWithCell(cell);
      expect(
        DROP_RULES[WTag.tc].keep({ nodeIn: cell, nodeOut: cell, ctx })
      ).toBe(true);
    });

    it('drops a cell that is scaffolding of a multi-cell construct', () => {
      const cell = newNonTextNode(WTag.tc);
      const ctx = ctxWithCell(cell);
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
