import { describe, it, expect, vi } from 'vitest';
import JSZip from 'jszip';
import { getError, makeDocx, reportText } from './helpers';
import { createReport } from '../index';
import { ErrorId, InternalError, InvalidOptionError } from '../errors';
import { type RunJsContext, type SandBox } from '../types';
import { setDebugLogSink } from '../debug';

if (process.env.DEBUG) setDebugLogSink(console.log);

describe('runJs (custom sandbox)', () => {
  it('runs every snippet through the user-provided sandbox', async () => {
    const seen: { code: string | undefined; hasCtx: boolean }[] = [];
    const runJs = ({
      sandbox,
      ctx,
    }: {
      sandbox: SandBox;
      ctx: RunJsContext;
    }) => {
      seen.push({ code: sandbox.__code__, hasCtx: ctx.options != null });
      return { modifiedSandbox: sandbox, result: `<${sandbox.__code__}>` };
    };

    const template = await makeDocx({ body: ['+++a+++', '+++b+++'] });
    const report = await createReport({ template, data: {}, runJs }, 'JS');

    expect(reportText(report)).toEqual('<a>\n<b>');
    expect(seen).toEqual([
      { code: 'a', hasCtx: true },
      { code: 'b', hasCtx: true },
    ]);
  });

  it('exposes the data and the additional context to the sandbox', async () => {
    const runJs = ({ sandbox }: { sandbox: SandBox }) => ({
      modifiedSandbox: sandbox,
      result: `${sandbox.fromData}/${sandbox.fromContext}`,
    });

    const template = await makeDocx({ body: ['+++whatever+++'] });
    const report = await createReport(
      {
        template,
        data: { fromData: 'DATA' },
        additionalJsContext: { fromContext: 'CONTEXT' },
        runJs,
      },
      'JS'
    );

    expect(reportText(report)).toEqual('DATA/CONTEXT');
  });

  it('carries the modified sandbox over to the next snippet', async () => {
    const runJs = ({ sandbox }: { sandbox: SandBox }) => {
      const count = ((sandbox.count as number | undefined) ?? 0) + 1;
      return { modifiedSandbox: { ...sandbox, count }, result: count };
    };

    const template = await makeDocx({
      body: ['+++a+++', '+++b+++', '+++c+++'],
    });
    const report = await createReport({ template, data: {}, runJs }, 'JS');

    expect(reportText(report)).toEqual('1\n2\n3');
  });
});

describe('indentXml', () => {
  it('indents the generated XML by default', async () => {
    const template = await makeDocx({ body: ['+++name+++'] });
    const xml = await createReport({ template, data: { name: 'John' } }, 'XML');
    expect(xml).toContain('\n  <w:body>');
  });

  it('emits everything on a single line when disabled', async () => {
    const template = await makeDocx({ body: ['+++name+++'] });
    const xml = await createReport(
      { template, data: { name: 'John' }, indentXml: false },
      'XML'
    );
    expect(xml).toContain('John');
    expect(xml).not.toContain('\n');
  });
});

describe('preserveSpace', () => {
  it('adds xml:space="preserve" to w:t nodes by default', async () => {
    const template = await makeDocx({ body: ['hello +++name+++'] });
    const xml = await createReport({ template, data: { name: 'John' } }, 'XML');
    expect(xml).toContain('xml:space="preserve"');
  });

  it('leaves w:t nodes alone when disabled', async () => {
    const template = await makeDocx({ body: ['hello +++name+++'] });
    const xml = await createReport(
      { template, data: { name: 'John' }, preserveSpace: false },
      'XML'
    );
    expect(xml).toContain('John');
    expect(xml).not.toContain('xml:space');
  });
});

describe('compressionLevel', () => {
  const repetitive = Array.from(
    { length: 500 },
    (_, i) => `lorem ipsum dolor sit amet ${i}`
  );

  it('produces a smaller file at higher compression levels', async () => {
    const template = await makeDocx({ body: repetitive });
    const fast = await createReport({
      template,
      data: {},
      compressionLevel: 1,
    });
    const small = await createReport({
      template,
      data: {},
      compressionLevel: 9,
    });
    expect(small.length).toBeLessThan(fast.length);
  });

  it('produces a readable docx at every level', async () => {
    const template = await makeDocx({ body: ['+++name+++'] });
    for (const compressionLevel of [0, 1, 6, 9]) {
      const report = await createReport({
        template,
        data: { name: 'John' },
        compressionLevel,
      });
      const zip = await JSZip.loadAsync(report);
      const xml = await zip.file('word/document.xml')?.async('text');
      expect(xml).toContain('John');
    }
  });
});

describe('maximumWalkingDepth', () => {
  const template = () =>
    makeDocx({
      body: ['+++FOR i IN [1, 2, 3, 4, 5]+++', '+++$i+++', '+++END-FOR i+++'],
    });

  it('aborts the walk once the limit is exceeded', async () => {
    await expect(
      createReport({
        template: await template(),
        data: {},
        maximumWalkingDepth: 5,
      })
    ).rejects.toThrow(InternalError);
  });

  it('renders normally when the limit is high enough', async () => {
    const report = await createReport(
      {
        template: await template(),
        data: {},
        maximumWalkingDepth: Infinity,
      },
      'JS'
    );
    expect(reportText(report)).toEqual('1\n2\n3\n4\n5');
  });
});

describe('data as a resolver function', () => {
  it('calls the resolver with `undefined` when the template has no QUERY', async () => {
    // Walks `extractQuery` all the way through the document without ever
    // finding a QUERY command.
    const template = await makeDocx({
      body: ['some text', '+++name+++', 'more text'],
    });
    const resolver = vi.fn(() => ({ name: 'John' }));

    const report = await createReport({ template, data: resolver }, 'JS');

    expect(resolver).toHaveBeenCalledTimes(1);
    expect(resolver.mock.calls[0]).toEqual([undefined, undefined]);
    expect(reportText(report)).toEqual('some text\nJohn\nmore text');
  });
});

describe('runJs receives a narrowed context', () => {
  it('exposes the options, the loop variables and the carried sandbox', async () => {
    // `ctx` holds live references to the engine's state, exactly as it did when
    // the whole context was handed over, so anything read out of it has to be
    // read while the call is happening.
    const seen: { keys: string[]; loopVar?: string; delimiter: string }[] = [];
    const runJs = ({
      sandbox,
      ctx,
    }: {
      sandbox: SandBox;
      ctx: RunJsContext;
    }) => {
      seen.push({
        keys: Object.keys(ctx).sort(),
        ...(ctx.loops[0] ? { loopVar: ctx.loops[0].varName } : {}),
        delimiter: ctx.options.cmdDelimiter[0],
      });
      // A real sandbox has to evaluate: `FOR` needs an actual array back.
      const result = new Function(
        's',
        `with (s) { return (${String(sandbox.__code__)}); }`
      )(sandbox);
      return { modifiedSandbox: sandbox, result };
    };

    const template = await makeDocx({
      body: ['+++FOR item IN items+++', '+++$item+++', '+++END-FOR item+++'],
    });
    const report = await createReport(
      { template, data: { items: ['a', 'b'] }, runJs },
      'JS'
    );

    expect(reportText(report)).toEqual('a\nb');
    // The keys are the ones `runJs` was given back when it received the whole
    // context, so code reading them did not have to change.
    expect(seen[0]?.keys).toEqual(['jsSandbox', 'loops', 'options', 'vars']);
    expect(seen[0]?.delimiter).toEqual('+++');
    expect(seen.map(s => s.loopVar)).toContain('item');
  });
});

describe('option validation', () => {
  const template = () => makeDocx({ body: ['text'] });

  it('rejects an option of the wrong type instead of ignoring it', async () => {
    // The motivating case: a non-function `errorHandler` used to be replaced
    // with `null`, so the handler silently never ran.
    const docx = await template();
    const err = await getError<InvalidOptionError>(() =>
      createReport({ template: docx, errorHandler: 'nope' } as never)
    );
    expect(err).toBeInstanceOf(InvalidOptionError);
    expect(err.option).toEqual('errorHandler');
    expect(err.message).toEqual(
      "Option 'errorHandler' must be a function, but received a string"
    );
    expect(err.properties.id).toEqual(ErrorId.invalidOption);
  });

  it.each([
    ['cmdDelimiter', 42],
    ['cmdDelimiter', ['only-one']],
    ['literalXmlDelimiter', 7],
    ['processLineBreaks', 'yes'],
    ['noSandbox', 1],
    ['runJs', {}],
    ['additionalJsContext', 'nope'],
    ['failFast', 'true'],
    ['rejectNullish', null],
    ['fixSmartQuotes', 'x'],
    ['maximumWalkingDepth', '100'],
    ['indentXml', 0],
    ['preserveSpace', 'yes'],
    ['compressionLevel', '9'],
    ['commandAliases', { ЕСЛИ: 5 }],
    ['operatorAliases', []],
  ])('rejects %s of the wrong type', async (option, value) => {
    await expect(
      createReport({ template: await template(), [option]: value } as never)
    ).rejects.toBeInstanceOf(InvalidOptionError);
  });

  it('accepts an option explicitly set to undefined', async () => {
    // Absent and `undefined` mean the same thing, and callers spread objects.
    await expect(
      createReport({
        template: await template(),
        errorHandler: undefined,
        failFast: undefined,
      } as never)
    ).resolves.toBeInstanceOf(Uint8Array);
  });

  it('reports a template of the wrong type instead of failing inside JSZip', async () => {
    const err = await getError<InvalidOptionError>(() =>
      createReport({ template: { not: 'a docx' } } as never)
    );
    expect(err).toBeInstanceOf(InvalidOptionError);
    expect(err.message).toEqual(
      "Option 'template' must be a Uint8Array (e.g. a Buffer), an ArrayBuffer " +
        'or a string, but received an object'
    );
  });

  it('reports a missing template', async () => {
    const err = await getError<InvalidOptionError>(() =>
      createReport({ data: {} } as never)
    );
    expect(err).toBeInstanceOf(InvalidOptionError);
    expect(err.message).toContain('but received nothing');
  });

  it('ignores an unknown option, but says so in the debug log', async () => {
    const lines: string[] = [];
    setDebugLogSink(msg => lines.push(String(msg)));
    try {
      await createReport({
        template: await template(),
        procesLineBreaks: true, // typo
        data: {},
      } as never);
    } finally {
      setDebugLogSink(process.env.DEBUG ? console.log : null);
    }
    expect(
      lines.some(l => l.includes('Ignoring unknown option: procesLineBreaks'))
    ).toBe(true);
  });

  it('still accepts every documented option', async () => {
    await expect(
      createReport({
        template: await template(),
        data: {},
        queryVars: { a: 1 },
        cmdDelimiter: ['{', '}'],
        literalXmlDelimiter: '||',
        processLineBreaks: true,
        noSandbox: false,
        additionalJsContext: {},
        failFast: true,
        rejectNullish: false,
        errorHandler: () => '',
        fixSmartQuotes: false,
        processLineBreaksAsNewText: false,
        maximumWalkingDepth: 1000,
        indentXml: true,
        preserveSpace: true,
        compressionLevel: 1,
        commandAliases: { ЕСЛИ: 'IF' },
        operatorAliases: { больше: '>' },
      })
    ).resolves.toBeInstanceOf(Uint8Array);
  });
});
