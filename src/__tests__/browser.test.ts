import { describe, it, expect, afterAll } from 'vitest';
import { makeDocx, reportText } from './helpers';

// The browser entry point installs a `Buffer` polyfill on `globalThis` as a
// side effect, so it lives in its own test file and puts the global back
// afterwards. Everything else about it should behave like the main entry.
const nodeBuffer = globalThis.Buffer;
afterAll(() => {
  globalThis.Buffer = nodeBuffer;
});

describe('browser entry point', () => {
  it('polyfills Buffer and re-exports the whole API', async () => {
    const browser = await import('../browser');

    expect(globalThis.Buffer).toBeDefined();
    expect(typeof globalThis.Buffer.from).toBe('function');

    expect(typeof browser.default).toBe('function');
    expect(browser.createReport).toBe(browser.default);
    expect(typeof browser.listCommands).toBe('function');
    expect(typeof browser.getMetadata).toBe('function');
    expect(typeof browser.CommandExecutionError).toBe('function');
  });

  it('renders a report through the browser bundle', async () => {
    const { createReport } = await import('../browser');
    const template = await makeDocx({ body: ['hello +++name+++'] });
    const report = await createReport(
      { noSandbox: true, template, data: { name: 'John' } },
      'JS'
    );
    expect(reportText(report)).toEqual('hello John');
  });
});
