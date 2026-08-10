import { afterEach, describe, expect, it } from 'vitest';
import { makeDocx, readFixture, tableXml } from './helpers';
import { createReport } from '../index';
import { setDebugLogSink } from '../debug';

// Building a debug message costs real time on a large document, so the hot
// paths of the walk only do it when a sink is installed. That makes those
// paths unreachable in a normal test run — this file is what exercises them,
// and what would catch a debug statement that throws on some node shape.

const captureLogs = async (run: () => Promise<unknown>): Promise<string[]> => {
  const lines: string[] = [];
  setDebugLogSink((message, ...rest) =>
    lines.push([message, ...rest.map(String)].join(' '))
  );
  try {
    await run();
  } finally {
    setDebugLogSink(null);
  }
  return lines;
};

describe('debug logging', () => {
  afterEach(() => setDebugLogSink(null));

  it('logs the walk, the loops and the commands of a render', async () => {
    const template = await makeDocx({
      body: [
        '+++FOR item IN items+++',
        '+++IF $item.big+++',
        'big: +++$item.name+++',
        '+++ELSE+++',
        'small: +++$item.name+++',
        '+++END-IF+++',
        '+++END-FOR item+++',
      ],
      bodyXml: tableXml([[['+++INS 1+++'], ['plain']]]),
    });

    const logs = await captureLogs(() =>
      createReport({
        template,
        data: {
          items: [
            { name: 'a', big: true },
            { name: 'b', big: false },
          ],
        },
      })
    );

    const joined = logs.join('\n');
    expect(joined).toContain('Unzipping...');
    // The per-node walk message, with the serialized node next to it.
    expect(joined).toContain('Next node [DOWN, level 2]');
    expect(joined).toContain('"_tag":"w:body"');
    // Loop bookkeeping: the exploration pass, then the two iterations.
    expect(joined).toContain('FOR loop on 0:item');
    expect(joined).toContain('EXPLORATION');
    expect(joined).toContain('IF loop on 1:__if_0');
    expect(joined).toContain('Jumping to level');
    expect(joined).toContain('Processing cmd: INS $item.name');
  });

  it('logs an image command without choking on the nodes it builds', async () => {
    const data = await readFixture('sample.png');
    const template = await makeDocx({ body: ['+++IMAGE img()+++'] });

    const logs = await captureLogs(() =>
      createReport({
        template,
        data: {},
        additionalJsContext: {
          img: () => ({ width: 6, height: 6, data, extension: '.png' }),
        },
      })
    );

    expect(logs.join('\n')).toContain('Processing cmd: IMAGE img()');
  });

  it('stops logging once the sink is removed', async () => {
    const template = await makeDocx({ body: ['+++name+++'] });
    const lines: string[] = [];
    setDebugLogSink(message => lines.push(String(message)));
    setDebugLogSink(null);

    await createReport({ template, data: { name: 'John' } });
    expect(lines).toEqual([]);
  });
});
