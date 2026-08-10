import { describe, it, expect } from 'vitest';
import {
  addChild,
  cloneNodeWithoutChildren,
  getFirstChild,
  getNextSibling,
  insertTextSiblingAfter,
  logLoop,
  newNonTextNode,
  newTextNode,
} from '../reportUtils';
import { TemplateParseError } from '../errors';
import { type Node, type TextNode } from '../types';

/** `<w:r><w:t>text</w:t></w:r>`, with all the parent links wired up. */
const makeRun = (text: string) => {
  const textNode = newTextNode(text);
  const tNode = newNonTextNode('w:t', {}, [textNode]);
  const run = newNonTextNode('w:r', {}, [tNode]);
  return { run, tNode, textNode };
};

describe('newNonTextNode', () => {
  it('links the given children back to their new parent', () => {
    const child = newNonTextNode('w:t');
    const parent = newNonTextNode('w:r', { a: 'b' }, [child]);
    expect(parent._attrs).toEqual({ a: 'b' });
    expect(child._parent).toBe(parent);
  });
});

describe('cloneNodeWithoutChildren', () => {
  it('drops the children and the parent link of a non-text node', () => {
    const { run } = makeRun('hi');
    const clone = cloneNodeWithoutChildren(run);
    expect(clone).toEqual({
      _children: [],
      _fTextNode: false,
      _tag: 'w:r',
      _attrs: {},
    });
    expect(clone._parent).toBeUndefined();
  });

  it('keeps the text of a text node', () => {
    const { textNode } = makeRun('hi');
    expect(cloneNodeWithoutChildren(textNode)).toEqual({
      _children: [],
      _fTextNode: true,
      _text: 'hi',
    });
  });
});

describe('getFirstChild / getNextSibling', () => {
  it('returns null when there is nothing to return', () => {
    const lonely = newNonTextNode('w:p');
    expect(getFirstChild(lonely)).toBeNull();
    // No parent at all
    expect(getNextSibling(lonely)).toBeNull();
  });

  it('walks the children of a node', () => {
    const a = newNonTextNode('a');
    const b = newNonTextNode('b');
    const parent = newNonTextNode('p', {}, [a, b]);
    expect(getFirstChild(parent)).toBe(a);
    expect(getNextSibling(a)).toBe(b);
    // `b` is the last child
    expect(getNextSibling(b)).toBeNull();
  });

  it('returns null for a node that its parent does not know about', () => {
    const parent = newNonTextNode('p', {}, [newNonTextNode('a')]);
    const orphan: Node = { ...newNonTextNode('b'), _parent: parent };
    expect(getNextSibling(orphan)).toBeNull();
  });
});

describe('addChild', () => {
  it('appends the child and sets its parent', () => {
    const parent = newNonTextNode('p');
    const child = newNonTextNode('c');
    expect(addChild(parent, child)).toBe(child);
    expect(parent._children).toEqual([child]);
    expect(child._parent).toBe(parent);
  });
});

describe('insertTextSiblingAfter', () => {
  it('inserts an empty w:t sibling right after the current one', () => {
    const { run, tNode, textNode } = makeRun('hello');
    const inserted = insertTextSiblingAfter(textNode);

    expect(inserted).toMatchObject({
      _children: [],
      _fTextNode: true,
      _text: '',
    });
    expect(run._children).toHaveLength(2);
    expect(run._children[0]).toBe(tNode);
    const newTNode = run._children[1];
    expect(newTNode).toMatchObject({ _fTextNode: false, _tag: 'w:t' });
    expect(newTNode?._children).toEqual([inserted]);
    expect(inserted._parent).toBe(newTNode);
  });

  it('copies the attributes of the w:t node it clones', () => {
    const textNode = newTextNode('hello');
    const tNode = newNonTextNode('w:t', { 'xml:space': 'preserve' }, [
      textNode,
    ]);
    newNonTextNode('w:r', {}, [tNode]);
    const inserted = insertTextSiblingAfter(textNode);
    expect(inserted._parent).toMatchObject({
      _attrs: { 'xml:space': 'preserve' },
    });
  });

  it('throws when the text node is not inside a w:t', () => {
    const textNode = newTextNode('hello');
    newNonTextNode('w:r', {}, [textNode]);
    expect(() => insertTextSiblingAfter(textNode)).toThrow(TemplateParseError);
    expect(() => insertTextSiblingAfter(textNode)).toThrow(
      'text node not within w:t'
    );
  });

  it('throws when the text node has no parent at all', () => {
    expect(() => insertTextSiblingAfter(newTextNode('hello'))).toThrow(
      'text node not within w:t'
    );
  });

  it('throws when the w:t node has no parent', () => {
    const textNode = newTextNode('hello');
    newNonTextNode('w:t', {}, [textNode]);
    expect(() => insertTextSiblingAfter(textNode)).toThrow(
      'w:t node has no parent'
    );
  });

  it('throws when the w:t node is not among its parent children', () => {
    const textNode = newTextNode('hello');
    const tNode = newNonTextNode('w:t', {}, [textNode]);
    // A parent that does not list `tNode` as one of its children
    tNode._parent = newNonTextNode('w:r');
    expect(() => insertTextSiblingAfter(textNode)).toThrow(TemplateParseError);
  });
});

describe('logLoop', () => {
  it('does nothing when there is no loop to log', () => {
    expect(() => logLoop([])).not.toThrow();
  });

  it('does not choke on either kind of loop', () => {
    const refNode: TextNode = newTextNode('');
    expect(() =>
      logLoop([
        { refNode, refNodeLevel: 1, varName: 'c', loopOver: [1, 2], idx: 0 },
        {
          refNode,
          refNodeLevel: 2,
          varName: '__if_0',
          loopOver: [],
          idx: -1,
          isIf: true,
          ifCurrentBranch: 1,
          ifActiveBranch: -1,
        },
      ])
    ).not.toThrow();
  });
});
