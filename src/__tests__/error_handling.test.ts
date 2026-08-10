import { describe, it, expect } from 'vitest';
import { fixturePath, getError, makeDocx, readFixture } from './helpers';
import fs from 'fs';
import QR from 'qrcode';
import { createReport, getMetadata } from '../index';
import { setDebugLogSink } from '../debug';
import {
  isError,
  NullishCommandResultError,
  CommandExecutionError,
  CommandSyntaxError,
  ImageError,
  IncompleteConditionalStatementError,
  InvalidCommandError,
  ObjectCommandResultError,
  TemplateParseError,
} from '../errors';

if (process.env.DEBUG) setDebugLogSink(console.log);

['noSandbox', 'sandbox'].forEach(sbStatus => {
  const noSandbox = sbStatus === 'sandbox' ? false : true;

  describe(`${sbStatus}`, () => {
    describe('rejectNullish setting', () => {
      it('INS', async () => {
        const template = await fs.promises.readFile(
          fixturePath('rejectNullishINS.docx')
        );

        // When not explicitly set, rejectNullish should be considered 'false' so this case should resolve.
        await expect(
          createReport({
            noSandbox,
            template,
            data: {
              testobj: {}, // accessing a non-existing property will result in `undefined`
              test2: 'second value!',
            },
          })
        ).resolves.toBeInstanceOf(Uint8Array);

        // The same case should throw when we decide NOT to accept nullish values.
        await expect(
          createReport({
            noSandbox,
            template,
            data: {
              testobj: {}, // accessing a non-existing property will result in `undefined`
              test2: 'second value!',
            },
            rejectNullish: true,
          })
        ).rejects.toBeInstanceOf(Error);

        // Should be ok when we actually set the value.
        await expect(
          createReport({
            noSandbox,
            template,
            data: {
              testobj: { value: 'the value is now set' },
              test2: 'second value!',
            },
            rejectNullish: true,
          })
        ).resolves.toBeInstanceOf(Uint8Array);
      });

      it('IMAGE', async () => {
        const template = await fs.promises.readFile(
          fixturePath('rejectNullishIMAGE.docx')
        );
        await expect(
          createReport({
            noSandbox,
            template,
            data: {},
            additionalJsContext: {
              qr: () => undefined,
            },
          })
        ).resolves.toBeInstanceOf(Uint8Array);

        await expect(
          createReport({
            noSandbox,
            template,
            data: {},
            rejectNullish: true,
            additionalJsContext: {
              qr: () => undefined,
            },
          })
        ).rejects.toThrowErrorMatchingSnapshot();

        await expect(
          createReport({
            noSandbox,
            template,
            data: {},
            rejectNullish: true,
            additionalJsContext: {
              qr: async (contents: string) => {
                const dataUrl = await QR.toDataURL(contents, { width: 500 });
                const data = dataUrl.slice('data:image/gif;base64,'.length);
                return { width: 6, height: 6, data, extension: '.gif' };
              },
            },
          })
        ).resolves.toBeInstanceOf(Uint8Array);
      });
    });

    describe('custom ErrorHandler', () => {
      it('allows graceful handling of NullishCommandResultError', async () => {
        expect.assertions(3);

        const template = await fs.promises.readFile(
          fixturePath('rejectNullishINS.docx')
        );

        const result = await createReport(
          {
            noSandbox,
            template,
            data: {
              testobj: {}, // accessing a non-existing property will result in `undefined`
              test2: 'second value!',
            },
            rejectNullish: true,
            errorHandler: (err, code) => {
              expect(err).toBeInstanceOf(NullishCommandResultError);
              expect(code).toStrictEqual('testobj.value');
              return `${err}`;
            },
          },
          'XML'
        );
        expect(result).toMatchSnapshot();
      });

      it('handles arbitrary errors occurring in command execution', async () => {
        const template = await fs.promises.readFile(
          fixturePath('commandExecutionError.docx')
        );

        // First check whether the CommandExecutionError is triggered correctly
        await expect(
          createReport({ noSandbox, template, data: {} })
        ).rejects.toThrow(CommandExecutionError);

        // Now try with an errorHandler
        expect(
          await createReport(
            {
              noSandbox,
              template,
              data: {},
              errorHandler: () => 'no problem dude',
            },
            'XML'
          )
        ).toMatchSnapshot();
      });

      it('properly handles InvalidCommandError', async () => {
        const template = await fs.promises.readFile(
          fixturePath('invalidCommand.docx')
        );

        const errs: Error[] = [];
        expect(
          await createReport(
            {
              noSandbox,
              template,
              data: {},
              errorHandler: err => {
                errs.push(err);
                return `${err}`;
              },
            },
            'XML'
          )
        ).toMatchSnapshot();

        expect(errs).toMatchSnapshot();
      });

      it('handler can decide to re-throw the error, crashing the render', async () => {
        const template = await fs.promises.readFile(
          fixturePath('invalidCommand.docx')
        );

        await expect(
          createReport({
            noSandbox,
            template,
            data: {},
            errorHandler: () => {
              throw new Error('yeah, no!');
            },
          })
        ).rejects.toThrow('yeah, no!');
      });

      it('properly handles nested InvalidCommandError from invalid FOR', async () => {
        const template = await fs.promises.readFile(
          fixturePath('invalidForCmd.docx')
        );

        const errs: Error[] = [];
        const cmds: (string | undefined)[] = [];
        expect(
          await createReport(
            {
              noSandbox,
              template,
              data: {},
              errorHandler: (err, code) => {
                errs.push(err);
                cmds.push(code);
                return `${err} (${code})`;
              },
            },
            'XML'
          )
        ).toMatchSnapshot();

        expect(errs).toMatchSnapshot();
        expect(cmds).toMatchSnapshot();
      });
    });
  });

  it('throw when user tries to iterate over non-array', async () => {
    const template = await fs.promises.readFile(
      fixturePath('forOverObject.docx')
    );
    await expect(
      createReport({
        noSandbox,
        template,
        data: {
          companies: {
            one: 'FIRST',
            two: 'SECOND',
            three: 'THIRD',
          },
        },
      })
    ).rejects.toThrowErrorMatchingSnapshot();
  });

  it('throw when result of INS command is an object', async () => {
    const template = await fs.promises.readFile(
      fixturePath('objectCommandResultError.docx')
    );
    await expect(
      createReport({
        noSandbox,
        template,
        data: {
          companies: {
            one: 'FIRST',
            two: 'SECOND',
            three: 'THIRD',
          },
        },
      })
    ).rejects.toThrowErrorMatchingSnapshot();
  });

  it('attaches the result to ObjectCommandResultError', async () => {
    const template = await fs.promises.readFile(
      fixturePath('objectCommandResultError.docx')
    );

    await expect(
      createReport({
        noSandbox,
        template,
        data: {
          companies: {
            one: 'FIRST',
            two: 'SECOND',
            three: 'THIRD',
          },
        },
      })
    ).rejects.toHaveProperty('result', {
      one: 'FIRST',
      two: 'SECOND',
      three: 'THIRD',
    });
  });

  it('Incomplete conditional statement: missing END-IF', async () => {
    const template = await fs.promises.readFile(
      fixturePath('missingEndIf.docx')
    );
    await expect(
      createReport({
        noSandbox,
        template,
        data: {},
        rejectNullish: true,

        // We use failFast:false to ensure the error handling of a missing END-IF and missing END-FOR are not accidentally mixed, see Github issue #322.
        failFast: false,
      })
    ).rejects.toMatchSnapshot();
  });

  it('Incomplete conditional statement: missing IF statement', async () => {
    const template = await fs.promises.readFile(
      fixturePath('unmatchedEndIf.docx')
    );
    await expect(
      createReport({
        noSandbox,
        template,
        data: {},
        rejectNullish: false, // needs to be false for the expected error to trigger instead of the NullishCommandResultError.
      })
    ).rejects.toThrow(
      `Unexpected END-IF outside of IF statement context: END-IF`
    );
  });

  it('Incomplete loop statement: unmatched END-FOR', async () => {
    const template = await fs.promises.readFile(
      fixturePath('unmatchedEndFor.docx')
    );
    await expect(
      createReport({
        noSandbox,
        template,
        data: {},
      })
    ).rejects.toThrow(
      `Unexpected END-FOR outside of FOR loop context: END-FOR`
    );
  });

  it('Incomplete loop statement: missing END-FOR', async () => {
    const template = await fs.promises.readFile(
      fixturePath('missingEndFor.docx')
    );
    await expect(
      createReport({
        noSandbox,
        template,
        data: {},
      })
    ).rejects.toThrow(
      `Unterminated FOR-loop ('FOR c'). Make sure each FOR loop has a corresponding END-FOR command.`
    );
  });

  it('Incomplete loop statement: invalid FOR', async () => {
    const template = await fs.promises.readFile(
      fixturePath('invalidForCmd.docx')
    );
    await expect(
      createReport({
        noSandbox,
        template,
        data: {},
      })
    ).rejects.toThrow('Invalid FOR command: FOR person');
  });
});

describe('errors from different realms', () => {
  it('sandbox', async () => {
    const template = await fs.promises.readFile(
      fixturePath('referenceError.docx')
    );

    const error = await getError(() =>
      createReport({ noSandbox: false, template, data: {} })
    );
    expect(error).toBeInstanceOf(CommandExecutionError);

    // We cannot check with instanceof as this Error is from another realm despite still being an error
    const commandExecutionError = error as CommandExecutionError;
    expect(commandExecutionError.err).not.toBeInstanceOf(ReferenceError);
    expect(commandExecutionError.err).not.toBeInstanceOf(Error);
    expect(commandExecutionError.err.name).toBe('ReferenceError');
    expect(commandExecutionError.err.message).toBe(
      'nonExistentVariable is not defined'
    );
  });

  it('noSandbox', async () => {
    const template = await fs.promises.readFile(
      fixturePath('referenceError.docx')
    );

    const error = await getError(() =>
      createReport({ noSandbox: true, template, data: {} })
    );
    expect(error).toBeInstanceOf(CommandExecutionError);

    // Without sandboxing, the error is from the same realm
    const commandExecutionError = error as CommandExecutionError;
    expect(commandExecutionError.err).toBeInstanceOf(ReferenceError);
    expect(commandExecutionError.err).toBeInstanceOf(Error);
    expect(commandExecutionError.err.name).toBe('ReferenceError');
    expect(commandExecutionError.err.message).toBe(
      'nonExistentVariable is not defined'
    );
  });
});

describe('isError', () => {
  it('Error is an error', () => {
    expect(isError(new Error())).toBeTruthy();
  });

  it('error-like object is an error', () => {
    expect(
      isError({
        name: 'ReferenceError',
        message: 'nonExistentVariable is not defined',
      })
    ).toBeTruthy();
  });

  it('primitive is not an error', () => {
    expect(isError(1)).toBeFalsy();
  });
});

describe('malformed commands', () => {
  const render = async (body: string[], options = {}) =>
    createReport({ template: await makeDocx({ body }), data: {}, ...options });

  it('CommandSyntaxError on a command that looks built-in but is not', async () => {
    // `IF-NOT` starts with the built-in `IF`, so it is not treated as an
    // implicit INS, but there is no command to run either.
    const error = await getError<CommandSyntaxError>(() =>
      render(['+++IF-NOT foo+++'])
    );
    expect(error).toBeInstanceOf(CommandSyntaxError);
    expect(error.message).toEqual('Invalid command syntax: IF-NOT foo');
    expect(error.command).toEqual('IF-NOT foo');
  });

  it('InvalidCommandError on an ALIAS without a definition', async () => {
    await expect(render(['+++ALIAS foo+++'])).rejects.toThrow(
      'Invalid ALIAS command: ALIAS foo'
    );
  });

  it('InvalidCommandError on an unknown shorthand', async () => {
    const error = await getError<InvalidCommandError>(() =>
      render(['+++ALIAS known INS 1+++', '+++*unknown+++'])
    );
    expect(error).toBeInstanceOf(InvalidCommandError);
    expect(error.message).toEqual('Unknown alias: *unknown');
  });

  it('InvalidCommandError on an ELSE-IF without a condition', async () => {
    await expect(
      render(['+++IF false+++', 'x', '+++ELSE-IF+++', 'y', '+++END-IF+++'])
    ).rejects.toThrow('Invalid ELSE-IF command (missing condition): ELSE-IF');
  });

  it('IncompleteConditionalStatementError on a missing END-IF (failFast)', async () => {
    await expect(
      render(['+++IF true+++', 'x'], { failFast: true })
    ).rejects.toBeInstanceOf(IncompleteConditionalStatementError);
  });

  it('collects an unterminated FOR loop when failFast is false', async () => {
    await expect(
      render(['+++FOR c IN [1, 2]+++', '+++$c+++'], { failFast: false })
    ).rejects.toEqual([
      expect.objectContaining({
        message:
          "Unterminated FOR-loop ('FOR c'). Make sure each FOR loop has a corresponding END-FOR command.",
      }),
    ]);
  });
});

['noSandbox', 'sandbox'].forEach(sbStatus => {
  const noSandbox = sbStatus === 'sandbox' ? false : true;

  describe(`${sbStatus}`, () => {
    describe('IMAGE parameter validation', () => {
      const renderImage = async (img: unknown) =>
        createReport({
          noSandbox,
          template: await makeDocx({ body: ['+++IMAGE img()+++'] }),
          data: {},
          additionalJsContext: { img: () => img },
        });

      const validPng = async () => ({
        width: 6,
        height: 6,
        data: await readFixture('sample.png'),
        extension: '.png',
      });

      it('rejects a non-numeric width or height', async () => {
        const error = await getError<ImageError>(async () =>
          renderImage({ ...(await validPng()), width: 'wide' })
        );
        expect(error).toBeInstanceOf(ImageError);
        expect(error.message).toContain('invalid image width: wide (in cm)');

        await expect(
          renderImage({ ...(await validPng()), height: NaN })
        ).rejects.toThrow('invalid image height: NaN (in cm)');
      });

      it('rejects image data that is not binary or base64', async () => {
        await expect(
          renderImage({ ...(await validPng()), data: 42 })
        ).rejects.toThrow(/image .data property needs to be provided as/);
      });

      it('rejects an unsupported extension', async () => {
        await expect(
          renderImage({ ...(await validPng()), extension: '.tiff' })
        ).rejects.toThrow(/An extension \(one of .*\) needs to be provided/);
      });

      it('rejects a thumbnail without an extension', async () => {
        await expect(
          renderImage({
            ...(await validPng()),
            extension: '.svg',
            thumbnail: { data: await readFixture('sample.png') },
          })
        ).rejects.toThrow(/An extension \(one of .*\) needs to be provided/);
      });

      it('accepts a base64-encoded string', async () => {
        const report = await renderImage({
          width: 6,
          height: 6,
          data: (await readFixture('sample.png')).toString('base64'),
          extension: '.png',
        });
        expect(report).toBeInstanceOf(Uint8Array);
      });
    });

    describe('snippets that throw a non-Error', () => {
      it('are still reported as a CommandExecutionError', async () => {
        const template = await makeDocx({
          body: [`+++!(() => { throw 'just a string' })()+++`],
        });
        const error = await getError<CommandExecutionError>(() =>
          createReport({ noSandbox, template, data: {} })
        );
        expect(error).toBeInstanceOf(CommandExecutionError);
        expect(error.err.message).toEqual('just a string');
        expect(error.message).toContain('just a string');
      });
    });

    describe('ObjectCommandResultError', () => {
      it('can be handled by a custom errorHandler', async () => {
        const seen: Error[] = [];
        const xml = await createReport(
          {
            noSandbox,
            template: await makeDocx({ body: ['+++obj+++'] }),
            data: { obj: { a: 1 } },
            errorHandler: err => {
              seen.push(err);
              return 'REPLACED';
            },
          },
          'XML'
        );
        expect(xml).toContain('REPLACED');
        expect(seen).toHaveLength(1);
        expect(seen[0]).toBeInstanceOf(ObjectCommandResultError);
        expect((seen[0] as ObjectCommandResultError).result).toEqual({ a: 1 });
      });
    });
  });
});

describe('broken templates', () => {
  it('rejects a zip without a [Content_Types].xml', async () => {
    const template = await makeDocx({
      body: ['hi'],
      files: { '[Content_Types].xml': null },
    });
    await expect(createReport({ template, data: {} })).rejects.toThrow(
      TemplateParseError
    );
  });

  it('rejects a [Content_Types].xml that lists no main document', async () => {
    const template = await makeDocx({
      body: ['hi'],
      files: {
        '[Content_Types].xml': `<?xml version="1.0"?>
          <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="xml" ContentType="application/xml"/>
          </Types>`,
      },
    });
    await expect(createReport({ template, data: {} })).rejects.toThrow(
      /Could not find main document/
    );
  });

  it('rejects when the main document is missing from the zip', async () => {
    const template = await makeDocx({
      body: ['hi'],
      files: { 'word/document.xml': null },
    });
    await expect(createReport({ template, data: {} })).rejects.toThrow(
      'document.xml could not be found'
    );
  });

  it('rejects a malformed main document', async () => {
    const template = await makeDocx({
      files: { 'word/document.xml': '<w:document><w:body>' },
    });
    await expect(createReport({ template, data: {} })).rejects.toThrow(
      /Unclosed root tag/
    );
  });
});

describe('getMetadata on unusual documents', () => {
  const APP_XML = (body: string) =>
    `<?xml version="1.0"?><Properties>${body}</Properties>`;
  const CORE_XML = `<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>`;

  it('rejects a document without metadata parts', async () => {
    await expect(getMetadata(await makeDocx({ body: ['hi'] }))).rejects.toThrow(
      'docProps/app.xml could not be read'
    );
  });

  it('rejects a document with an app.xml but no core.xml', async () => {
    const template = await makeDocx({
      body: ['hi'],
      files: { 'docProps/app.xml': APP_XML('<Pages>3</Pages>') },
    });
    await expect(getMetadata(template)).rejects.toThrow(
      'docProps/core.xml could not be read'
    );
  });

  it('returns undefined for missing, empty and non-numeric fields', async () => {
    const template = await makeDocx({
      body: ['hi'],
      files: {
        'docProps/app.xml': APP_XML(
          '<Pages>lots</Pages><Words>4</Words><Company/>'
        ),
        'docProps/core.xml': CORE_XML,
      },
    });
    const metadata = await getMetadata(template);
    expect(metadata.pages).toBeUndefined(); // not a number
    expect(metadata.words).toEqual(4);
    expect(metadata.company).toBeUndefined(); // empty element
    expect(metadata.lines).toBeUndefined(); // absent element
    expect(metadata.title).toBeUndefined(); // absent from core.xml
  });

  it('rejects a metadata field that holds markup instead of text', async () => {
    const template = await makeDocx({
      body: ['hi'],
      files: {
        'docProps/app.xml': APP_XML('<Pages><nested/></Pages>'),
        'docProps/core.xml': CORE_XML,
      },
    });
    await expect(getMetadata(template)).rejects.toThrow('Not a text node');
  });
});
