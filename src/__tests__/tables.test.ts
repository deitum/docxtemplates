import { describe, it, expect } from 'vitest';
import { makeDocx, tableCells, tableXml } from './helpers';
import { createReport } from '../index';
import { type ReportData } from '../types';
import { setDebugLogSink } from '../debug';

if (process.env.DEBUG) setDebugLogSink(console.log);

/** Renders a table (one array of paragraph lines per cell) and reads it back. */
const renderTable = async (
  rows: string[][][],
  data: ReportData
): Promise<string[][]> => {
  const template = await makeDocx({ bodyXml: tableXml(rows) });
  const report = await createReport({ template, data, noSandbox: true }, 'JS');
  return tableCells(report);
};

// A cell containing nothing but commands is only deleted when those commands
// are part of a construct spanning several cells; otherwise the cell stays
// (empty), so that the rest of the row keeps its columns.
describe('table cells that only contain commands', () => {
  it('keeps the cell of a FOR loop that iterates over nothing', async () => {
    const rows = await renderTable(
      [
        [
          ['name1'],
          ['+++FOR p IN people+++', '+++INS $p+++', '+++END-FOR p+++'],
        ],
      ],
      { people: [] }
    );
    expect(rows).toEqual([['name1', '']]);
  });

  it('keeps the cell of a FOR loop whose body renders nothing', async () => {
    const rows = await renderTable(
      [[['name1'], ['+++FOR p IN people+++', '+++END-FOR p+++']]],
      { people: ['a', 'b'] }
    );
    expect(rows).toEqual([['name1', '']]);
  });

  it('renders the cell of a FOR loop that produces output', async () => {
    const rows = await renderTable(
      [
        [
          ['name1'],
          ['+++FOR p IN people+++', '+++INS $p.name+++', '+++END-FOR p+++'],
        ],
      ],
      { people: [{ name: 'John' }] }
    );
    expect(rows).toEqual([['name1', 'John']]);
  });

  it('keeps the cell of an IF construct whose condition is falsy', async () => {
    const cell = ['+++IF ok+++', '+++INS value+++', '+++END-IF+++'];
    expect(
      await renderTable([[['name1'], cell]], { ok: false, value: 'V' })
    ).toEqual([['name1', '']]);
    expect(
      await renderTable([[['name1'], cell]], { ok: true, value: 'V' })
    ).toEqual([['name1', 'V']]);
  });

  it('keeps a cell that only contains an EXEC command', async () => {
    const rows = await renderTable([[['name1'], ['+++EXEC 1 + 1+++']]], {});
    expect(rows).toEqual([['name1', '']]);
  });

  it('keeps the rest of the table intact', async () => {
    const rows = await renderTable(
      [
        [['Name'], ['List']],
        [
          ['name1'],
          ['+++FOR p IN people1+++', '+++INS $p.name+++', '+++END-FOR p+++'],
        ],
        [
          ['name2'],
          ['+++FOR p IN people2+++', '+++INS $p.name+++', '+++END-FOR p+++'],
        ],
      ],
      { people1: [], people2: [{ name: 'John' }] }
    );
    expect(rows).toEqual([
      ['Name', 'List'],
      ['name1', ''],
      ['name2', 'John'],
    ]);
  });
});

describe('constructs spanning several table cells', () => {
  it('deletes the FOR/END-FOR cells of a dynamic-columns table', async () => {
    const rows = await renderTable(
      [
        [
          ['+++FOR col IN columns+++'],
          ['+++INS $col+++'],
          ['+++END-FOR col+++'],
        ],
      ],
      { columns: ['Column 1', 'Column 2', 'Column 3'] }
    );
    expect(rows).toEqual([['Column 1', 'Column 2', 'Column 3']]);
  });

  it('deletes every cell of a dynamic-columns table without columns', async () => {
    const rows = await renderTable(
      [
        [
          ['+++FOR col IN columns+++'],
          ['+++INS $col+++'],
          ['+++END-FOR col+++'],
        ],
      ],
      { columns: [] }
    );
    expect(rows).toEqual([[]]);
  });

  it('deletes the IF/ELSE/END-IF cells of a construct spread over cells', async () => {
    const rows = [
      [['+++IF ok+++'], ['yes'], ['+++ELSE+++'], ['no'], ['+++END-IF+++']],
    ];
    expect(await renderTable(rows, { ok: true })).toEqual([['yes']]);
    expect(await renderTable(rows, { ok: false })).toEqual([['no']]);
  });

  it('keeps the cells that follow the END-FOR of a dynamic-columns loop', async () => {
    const rows = await renderTable(
      [
        [
          ['+++FOR col IN columns+++'],
          ['+++INS $col+++'],
          ['+++END-FOR col+++'],
          ['Total'],
        ],
      ],
      { columns: ['Column 1', 'Column 2'] }
    );
    expect(rows).toEqual([['Column 1', 'Column 2', 'Total']]);
  });
});
