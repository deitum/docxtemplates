/**
 * What the JS sandbox guarantees, pinned down.
 *
 * These exist because the engine reuses one evaluation context per document
 * part rather than building a fresh one per command — the single biggest cost
 * in a template with loops. They are the definition of what "reuse" is allowed
 * to mean: what carries over between snippets, what overrides what, and where
 * the isolation boundaries are.
 *
 * Where the two evaluation paths (`vm` and `noSandbox`) genuinely differ, they
 * are asserted separately rather than papered over. `noSandbox` runs the
 * snippet with `with(this) { eval(...) }`, so an assignment to a name the
 * sandbox doesn't already carry creates a *host* global; that is a property of
 * the insecure path, and is recorded here rather than fixed.
 */
import JSZip from 'jszip';
import { afterEach, describe, expect, it } from 'vitest';
import { createReport } from '../index';
import { makeDocx, reportText } from './helpers';

const render = async (
  body: string[],
  options: Record<string, unknown> = {}
): Promise<string> => {
  const template = await makeDocx({ body });
  return reportText(await createReport({ template, ...options }, 'JS'));
};

// The `noSandbox` cases below deliberately leak into the host global scope.
const HOST_GLOBALS = ['leakedAcrossReports', 'leakedIntoHeader'];
afterEach(() => {
  for (const name of HOST_GLOBALS) {
    delete (globalThis as Record<string, unknown>)[name];
  }
});

for (const noSandbox of [false, true]) {
  describe(noSandbox ? 'noSandbox' : 'sandbox', () => {
    it('carries a value from one snippet to the next', async () => {
      expect(
        await render(
          [
            '+++EXEC counter = 1+++',
            'first: +++counter+++',
            '+++EXEC counter = counter + 1+++',
            'second: +++counter+++',
          ],
          { noSandbox, data: {} }
        )
      ).toEqual('first: 1\nsecond: 2');
    });

    it('lets a snippet see objects an earlier snippet built', async () => {
      expect(
        await render(
          [
            '+++EXEC cfg = { nested: { label: "deep" } }+++',
            '+++cfg.nested.label+++',
          ],
          { noSandbox, data: {} }
        )
      ).toEqual('deep');
    });

    it('lets a snippet delete a variable an earlier snippet defined', async () => {
      expect(
        await render(
          [
            '+++EXEC gone = "here"+++',
            'before: +++typeof gone+++',
            '+++EXEC delete gone+++',
            'after: +++typeof gone+++',
          ],
          { noSandbox, data: {} }
        )
      ).toEqual('before: string\nafter: undefined');
    });

    it('lets data override a value left behind by an earlier snippet', async () => {
      expect(
        await render(['+++EXEC label = "from snippet"+++', '+++label+++'], {
          noSandbox,
          data: { label: 'from data' },
        })
      ).toEqual('from data');
    });

    it('lets additionalJsContext override a value left behind by an earlier snippet', async () => {
      expect(
        await render(
          ['+++EXEC helper = () => "from snippet"+++', '+++helper()+++'],
          {
            noSandbox,
            data: {},
            additionalJsContext: { helper: () => 'from context' },
          }
        )
      ).toEqual('from context');
    });

    it('exposes loop variables and the loop index', async () => {
      expect(
        await render(
          [
            '+++FOR item IN items+++',
            '+++$idx+++: +++$item.name+++',
            '+++END-FOR item+++',
          ],
          { noSandbox, data: { items: [{ name: 'a' }, { name: 'b' }] } }
        )
      ).toEqual('0: a\n1: b');
    });
  });
}

describe('sandbox (vm) isolation', () => {
  it('keeps a snippet from reaching the host global scope', async () => {
    expect(await render(['+++typeof process+++'], { data: {} })).toEqual(
      'undefined'
    );
  });

  it('does not carry state between two reports', async () => {
    expect(
      await render(['+++EXEC leakedAcrossReports = "yes"+++', 'x'], {
        data: {},
      })
    ).toEqual('x');
    expect(
      await render(['+++typeof leakedAcrossReports+++'], { data: {} })
    ).toEqual('undefined');
  });

  it('does not carry state from the main document into a header', async () => {
    const template = await makeDocx({
      body: ['+++EXEC leakedIntoHeader = "from body"+++', 'body ok'],
      header: ['header: +++typeof leakedIntoHeader+++'],
    });
    const report = await createReport({ template, data: {} });
    const zip = await JSZip.loadAsync(report);
    const header = await zip.file('word/header1.xml')!.async('text');
    expect(header).toContain('undefined');
  });

  it('shares built-ins between the snippets of one part', async () => {
    // A consequence of evaluating every snippet of a part in one context: a
    // value built by one snippet is recognised by the next one's `instanceof`.
    // Before the contexts were shared, each snippet had its own `Array` and
    // this answered `false`.
    expect(
      await render(
        ['+++EXEC items = [1, 2]+++', '+++items instanceof Array+++'],
        { data: {} }
      )
    ).toEqual('true');
  });

  it('lets a function defined in one snippet mutate state read by the next', async () => {
    // Likewise: the closure `bump` captures the context it was defined in. When
    // every snippet had a context of its own, calling `bump()` from a later
    // snippet incremented a `counter` nobody could see again, and this answered
    // `1`.
    expect(
      await render(
        [
          '+++EXEC counter = 1; bump = () => (counter += 1)+++',
          '+++EXEC bump()+++',
          '+++counter+++',
        ],
        { data: {} }
      )
    ).toEqual('2');
  });
});

describe('noSandbox escapes the sandbox, as documented', () => {
  it('leaks assignments into the host global scope, across reports', async () => {
    expect(
      await render(['+++EXEC leakedAcrossReports = "yes"+++', 'x'], {
        noSandbox: true,
        data: {},
      })
    ).toEqual('x');
    expect(
      await render(['+++typeof leakedAcrossReports+++'], {
        noSandbox: true,
        data: {},
      })
    ).toEqual('string');
  });

  it('leaks assignments from the main document into a header', async () => {
    const template = await makeDocx({
      body: ['+++EXEC leakedIntoHeader = "from body"+++', 'body ok'],
      header: ['header: +++typeof leakedIntoHeader+++'],
    });
    const report = await createReport({ template, noSandbox: true, data: {} });
    const zip = await JSZip.loadAsync(report);
    const header = await zip.file('word/header1.xml')!.async('text');
    expect(header).toContain('string');
  });
});
