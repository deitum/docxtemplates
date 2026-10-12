import { describe, it, expect } from 'vitest';
import {
  makeDocx,
  readFixture,
  reportText,
  tableCells,
  tableXml,
} from './helpers';
import { createReport } from '../index';
import { setDebugLogSink } from '../debug';

if (process.env.DEBUG) setDebugLogSink(console.log);

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

      it('throws on an ELSE-IF after an ELSE', async () => {
        const template = await makeDocx({
          body: [
            '+++IF false+++',
            'a',
            '+++ELSE+++',
            'b',
            '+++ELSE-IF true+++',
            'c',
            '+++END-IF+++',
          ],
        });
        await expect(
          createReport({ noSandbox, template, data: {} }, 'JS')
        ).rejects.toThrow('Unexpected ELSE-IF after an ELSE command: ELSE-IF');
      });

      it('evaluates ELSE-IF conditions in the scope of the enclosing FOR loop', async () => {
        // `$idx` inside an ELSE-IF condition must refer to the FOR loop, not to
        // the IF construct the branch belongs to.
        const template = await makeDocx({
          body: [
            '+++FOR item IN items+++',
            '+++IF $idx === 0+++',
            'first: +++$item+++',
            '+++ELSE-IF $idx === 1+++',
            'second: +++$item+++',
            '+++ELSE+++',
            'rest: +++$item+++',
            '+++END-IF+++',
            '+++END-FOR item+++',
          ],
        });
        const report = await createReport(
          { noSandbox, template, data: { items: ['a', 'b', 'c', 'd'] } },
          'JS'
        );
        expect(reportText(report)).toEqual(
          ['first: a', 'second: b', 'rest: c', 'rest: d'].join('\n')
        );
      });
    });

    // Rejected outright until 1.1.0, for fear of upstream issue #340: an IF
    // jumping back over the IF it is nested into, and opening it again. Only
    // the ones that would actually do that are rejected now.
    describe('IF nested into another IF on the same line or table row', () => {
      const renderTable = async (rows: string[][][], data: object) => {
        const template = await makeDocx({ bodyXml: tableXml(rows) });
        return tableCells(
          await createReport({ noSandbox, template, data }, 'JS')
        );
      };

      it('renders an IF closed in its own paragraph of the cell', async () => {
        const rows = [
          [
            ['name'],
            ['+++IF a+++', 'A', '+++IF b+++B+++END-IF+++', '+++END-IF+++'],
          ],
        ];
        expect(await renderTable(rows, { a: true, b: true })).toEqual([
          ['name', 'AB'],
        ]);
        expect(await renderTable(rows, { a: true, b: false })).toEqual([
          ['name', 'A'],
        ]);
        expect(await renderTable(rows, { a: false, b: true })).toEqual([
          ['name', ''],
        ]);
      });

      it('renders an IF spanning several paragraphs of the cell', async () => {
        const rows = [
          [
            ['name'],
            [
              '+++IF a+++',
              'A',
              '+++IF b+++',
              'B',
              '+++END-IF+++',
              '+++END-IF+++',
            ],
          ],
        ];
        expect(await renderTable(rows, { a: true, b: true })).toEqual([
          ['name', 'AB'],
        ]);
        expect(await renderTable(rows, { a: true, b: false })).toEqual([
          ['name', 'A'],
        ]);
      });

      it('renders an IF opened in another cell of the row', async () => {
        const rows = [
          [
            ['+++IF a+++', 'A'],
            ['+++IF b+++', 'B', '+++END-IF+++', '+++END-IF+++'],
          ],
        ];
        expect(await renderTable(rows, { a: true, b: true })).toEqual([
          ['A', 'B'],
        ]);
        expect(await renderTable(rows, { a: true, b: false })).toEqual([
          ['A', ''],
        ]);
      });

      it('renders an ELSE branch of the nested IF', async () => {
        const rows = [
          [
            [
              '+++IF a+++',
              'A',
              '+++IF b+++B+++ELSE+++not B+++END-IF+++',
              '+++END-IF+++',
            ],
          ],
        ];
        expect(await renderTable(rows, { a: true, b: false })).toEqual([
          ['Anot B'],
        ]);
      });

      it('renders an IF in a FOR loop in an IF of the same cell', async () => {
        const rows = [
          [
            [
              '+++IF a+++',
              '+++FOR x IN xs+++',
              '+++IF $x > 1++++++$x++++++END-IF+++',
              '+++END-FOR x+++',
              '+++END-IF+++',
            ],
          ],
        ];
        expect(await renderTable(rows, { a: true, xs: [1, 2, 3] })).toEqual([
          ['23'],
        ]);
      });

      it('renders nested IFs within a single paragraph', async () => {
        const template = await makeDocx({
          body: ['+++IF a+++A+++IF b+++B+++END-IF++++++END-IF+++'],
        });
        const render = async (data: object) =>
          reportText(await createReport({ noSandbox, template, data }, 'JS'));
        expect(await render({ a: true, b: true })).toEqual('AB');
        expect(await render({ a: true, b: false })).toEqual('A');
        expect(await render({ a: false, b: true })).toEqual('');
      });

      it('rejects an IF that closes in another cell', async () => {
        // Jumping back to its cell would run into `IF a` again
        const template = await makeDocx({
          bodyXml: tableXml([
            [
              ['+++IF a+++', '+++IF b+++'],
              ['+++END-IF+++', '+++END-IF+++'],
            ],
          ]),
        });
        await expect(
          createReport(
            { noSandbox, template, data: { a: true, b: true } },
            'JS'
          )
        ).rejects.toThrow(
          'Invalid IF command nested into another IF command on the same table row: IF b'
        );
      });

      it('rejects it when it cannot be told where the IF closes', async () => {
        // A `*shorthand` might stand for an END-IF, which makes pairing up the
        // IFs and END-IFs of the part guesswork
        const template = await makeDocx({
          body: ['+++ALIAS one INS 1+++', '+++*one+++'],
          bodyXml: tableXml([
            [['+++IF a+++', 'A', '+++IF b+++B+++END-IF+++', '+++END-IF+++']],
          ]),
        });
        await expect(
          createReport(
            { noSandbox, template, data: { a: true, b: true } },
            'JS'
          )
        ).rejects.toThrow(
          'Invalid IF command nested into another IF command on the same table row: IF b'
        );
      });
    });
  });
});
