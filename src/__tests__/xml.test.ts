import { describe, it, expect } from 'vitest';
import { parseXml, buildXml } from '../xml';
import { newNonTextNode, newTextNode } from '../reportUtils';
import { type Node, type NonTextNode } from '../types';

const OPTIONS = { literalXmlDelimiter: '||', indentXml: true };

const build = (node: Node, options = OPTIONS) =>
  buildXml(node, options).toString('utf-8');

describe('parseXml', () => {
  it('builds a tree of nodes with attributes, text and parents', async () => {
    const root = (await parseXml(
      `<w:p><w:r><w:t xml:space="preserve">hello</w:t></w:r></w:p>`
    )) as NonTextNode;

    expect(root._fTextNode).toBe(false);
    expect(root._tag).toEqual('w:p');
    const run = root._children[0] as NonTextNode;
    expect(run._tag).toEqual('w:r');
    expect(run._parent).toBe(root);
    const t = run._children[0] as NonTextNode;
    expect(t._attrs['xml:space']).toEqual('preserve');
    expect(t._children[0]).toMatchObject({ _fTextNode: true, _text: 'hello' });
  });

  it('keeps whitespace as-is', async () => {
    const root = await parseXml(`<a>  spaced  </a>`);
    expect(root._children[0]).toMatchObject({ _text: '  spaced  ' });
  });

  it('rejects malformed XML', async () => {
    await expect(parseXml('<a><b></a>')).rejects.toThrow();
    await expect(parseXml('<a>')).rejects.toThrow(/Unclosed root tag/);
  });
});

describe('buildXml', () => {
  it('prepends the XML declaration to the root node only', () => {
    const child = newNonTextNode('w:r');
    const root = newNonTextNode('w:p', {}, [child]);
    const xml = build(root);
    expect(xml.startsWith('<?xml version="1.0"')).toBe(true);
    expect(xml.match(/<\?xml/g)).toHaveLength(1);
  });

  it('escapes XML-hostile characters in text', () => {
    const root = newNonTextNode('w:t', {}, [newTextNode('3 < 4 & 5 > 2')]);
    expect(build(root)).toContain('3 &lt; 4 &amp; 5 &gt; 2');
  });

  it('escapes XML-hostile characters in attributes', () => {
    const root = newNonTextNode('a', { v: `<&>'"` });
    expect(build(root)).toContain('v="&lt;&amp;&gt;&apos;&quot;"');
  });

  it('skips attributes with a nullish value', () => {
    const root = newNonTextNode('a', {
      kept: 'yes',
      dropped: undefined,
    } as unknown as NonTextNode['_attrs']);
    const xml = build(root);
    expect(xml).toContain('kept="yes"');
    expect(xml).not.toContain('dropped');
  });

  it('supports qualified attribute objects as well as plain strings', () => {
    // `parseXml` yields plain strings, but the `QualifiedAttribute` shape is
    // part of the `NonTextNode` contract.
    const root = newNonTextNode('a', {
      v: { name: 'v', value: 'x&y', prefix: '', local: 'v', uri: '' },
    });
    expect(build(root)).toContain('v="x&amp;y"');
  });

  it('passes text between literal delimiters through unescaped', () => {
    const root = newNonTextNode('w:t', {}, [
      newTextNode('foo||<w:br/>||bar & baz'),
    ]);
    expect(build(root)).toContain('foo<w:br/>bar &amp; baz');
  });

  it('treats a trailing literal delimiter as opening a literal section', () => {
    const root = newNonTextNode('w:t', {}, [newTextNode('foo||<w:br/>')]);
    expect(build(root)).toContain('foo<w:br/>');
  });

  it('honours a custom literal delimiter', () => {
    const root = newNonTextNode('w:t', {}, [newTextNode('foo__<w:br/>__bar')]);
    const xml = build(root, { literalXmlDelimiter: '__', indentXml: true });
    expect(xml).toContain('foo<w:br/>bar');
  });

  it('self-closes childless nodes', () => {
    const root = newNonTextNode('w:p', {}, [newNonTextNode('w:br')]);
    expect(build(root)).toContain('<w:br/>');
    expect(build(root)).not.toContain('</w:br>');
  });

  it('indents nested nodes, but not when indentXml is false', () => {
    const root = newNonTextNode('w:p', {}, [
      newNonTextNode('w:r', {}, [
        newNonTextNode('w:t', {}, [newTextNode('hi')]),
      ]),
    ]);
    expect(build(root)).toContain('\n  <w:r>\n    <w:t');

    const flat = build(root, { literalXmlDelimiter: '||', indentXml: false });
    expect(flat).toEqual(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:p><w:r><w:t>hi</w:t></w:r></w:p>'
    );
  });

  it('round-trips a parsed document', async () => {
    const source =
      '<w:p><w:r><w:t xml:space="preserve">a &amp; b</w:t></w:r></w:p>';
    const parsed = await parseXml(source);
    const xml = build(parsed, { literalXmlDelimiter: '||', indentXml: false });
    expect(xml).toEqual(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + source
    );
  });
});
