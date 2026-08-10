import { describe, it, expect } from 'vitest';
import { readFixture } from './helpers';
import { createReport } from '../index';
import { type Node } from '../types';
import { setDebugLogSink } from '../debug';

if (process.env.DEBUG) setDebugLogSink(console.log);

// Concatenates the text of all the text nodes below the given node
const nodeText = (node: Node): string =>
  node._fTextNode ? node._text : node._children.map(nodeText).join('');

// Returns the text of the report, one line per (non-empty) paragraph
const reportText = (node: Node): string => {
  const lines: string[] = [];
  const walk = (n: Node) => {
    if (!n._fTextNode && n._tag === 'w:p') {
      const text = nodeText(n).trim();
      if (text !== '') lines.push(text);
      return;
    }
    n._children.forEach(walk);
  };
  walk(node);
  return lines.join('\n');
};

const render = async (
  fixture: string,
  data: any,
  noSandbox: boolean
): Promise<string> => {
  const template = await readFixture(fixture);
  const report = await createReport({ noSandbox, template, data }, 'JS');
  return reportText(report);
};

['noSandbox', 'sandbox'].forEach(sbStatus => {
  const noSandbox = sbStatus === 'sandbox' ? false : true;

  describe(`${sbStatus}`, () => {
    describe('IF / ELSE-IF / ELSE', () => {
      it('renders the IF branch when the condition is truthy', async () => {
        expect(await render('ifElse.docx', { value: 42 }, noSandbox)).toEqual(
          'big\ndone'
        );
      });

      it('renders the ELSE branch when the condition is falsy', async () => {
        expect(await render('ifElse.docx', { value: 1 }, noSandbox)).toEqual(
          'small\ndone'
        );
      });

      it('renders the first matching branch of an ELSE-IF chain', async () => {
        expect(await render('ifElseIf.docx', { value: 42 }, noSandbox)).toEqual(
          'big\ndone'
        );
        expect(await render('ifElseIf.docx', { value: 7 }, noSandbox)).toEqual(
          'medium\ndone'
        );
        expect(await render('ifElseIf.docx', { value: 3 }, noSandbox)).toEqual(
          'small\ndone'
        );
      });

      it('renders the final ELSE branch when no condition matches', async () => {
        expect(await render('ifElseIf.docx', { value: -1 }, noSandbox)).toEqual(
          'none\ndone'
        );
      });

      it('supports inline IF / ELSE-IF / ELSE within a single paragraph', async () => {
        expect(
          await render('ifElseInline.docx', { value: 42 }, noSandbox)
        ).toEqual('Result: big!');
        expect(
          await render('ifElseInline.docx', { value: 7 }, noSandbox)
        ).toEqual('Result: medium!');
        expect(
          await render('ifElseInline.docx', { value: 0 }, noSandbox)
        ).toEqual('Result: small!');
      });

      it('supports nested IF / ELSE constructs', async () => {
        expect(
          await render('ifElseNested.docx', { a: true, b: true }, noSandbox)
        ).toEqual('a-yes\nb-yes');
        expect(
          await render('ifElseNested.docx', { a: true, b: false }, noSandbox)
        ).toEqual('a-yes\nb-no');
        expect(
          await render('ifElseNested.docx', { a: false, b: true }, noSandbox)
        ).toEqual('a-no\nc-yes');
        expect(
          await render('ifElseNested.docx', { a: false, b: false }, noSandbox)
        ).toEqual('a-no\nc-no');
      });

      it('supports IF / ELSE-IF / ELSE inside a FOR loop', async () => {
        const data = {
          items: [
            { name: 'one', big: false },
            { name: 'two', big: false },
            { name: 'three', big: true },
          ],
        };
        expect(await render('ifElseFor.docx', data, noSandbox)).toEqual(
          ['small: one', 'SECOND: two', 'BIG: three'].join('\n')
        );
      });

      it('supports IF / ELSE-IF / ELSE over table rows', async () => {
        expect(
          await render('ifElseRow.docx', { value: 42 }, noSandbox)
        ).toEqual('big');
        expect(await render('ifElseRow.docx', { value: 7 }, noSandbox)).toEqual(
          'medium'
        );
        expect(await render('ifElseRow.docx', { value: 0 }, noSandbox)).toEqual(
          'small'
        );
      });

      it('supports a FOR loop inside an ELSE branch', async () => {
        const data = { a: false, list: ['one', 'two'] };
        expect(await render('ifElseForInside.docx', data, noSandbox)).toEqual(
          'item: one\nitem: two'
        );
        expect(
          await render('ifElseForInside.docx', { ...data, a: true }, noSandbox)
        ).toEqual('X');
      });

      it('does not evaluate commands of branches that are not selected', async () => {
        // The ELSE branch would throw if it were evaluated
        expect(
          await render('ifElseSkippedBranch.docx', { a: true }, noSandbox)
        ).toEqual('safe');
      });

      it('throws on an ELSE outside of an IF statement', async () => {
        const template = await readFixture('elseOutsideIf.docx');
        await expect(
          createReport({ noSandbox, template, data: {} }, 'JS')
        ).rejects.toThrow(
          'Unexpected ELSE outside of IF statement context: ELSE'
        );
      });

      it('throws on a second ELSE within the same IF statement', async () => {
        const template = await readFixture('ifDoubleElse.docx');
        await expect(
          createReport({ noSandbox, template, data: { value: true } }, 'JS')
        ).rejects.toThrow('Unexpected ELSE after an ELSE command: ELSE');
      });
    });
  });
});
