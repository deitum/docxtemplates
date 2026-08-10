/**
 * The `properties` every error carries: a stable code, the part of the .docx it
 * came from, the command that caused it, and an explanation.
 *
 * The classes, messages and the shape of what is thrown are unchanged; this
 * covers what was added.
 */
import { describe, expect, it } from 'vitest';
import {
  CommandExecutionError,
  createReport,
  ErrorId,
  ImageError,
  InvalidCommandError,
  TemplateError,
  TemplateParseError,
} from '../index';
import { getError, makeDocx } from './helpers';

const errorFrom = async (options: Record<string, unknown>) =>
  getError<TemplateError>(() => createReport(options as never, 'JS'));

describe('error properties', () => {
  it('says which part of the document the error came from', async () => {
    // The `'JS'` probe returns as soon as the main document is done, so an
    // error in a header only surfaces when the whole package is rendered.
    const template = await makeDocx({
      body: ['all good'],
      header: ['+++missingInHeader+++'],
    });
    const err = await getError<TemplateError>(() =>
      createReport({ template, data: {} })
    );
    expect(err.properties.part).toEqual('header1.xml');
    expect(err.properties.command).toEqual('missingInHeader');
  });

  it('names the main document when the error is in the body', async () => {
    const template = await makeDocx({ body: ['+++nope.missing+++'] });
    const err = await errorFrom({ template, data: {} });
    expect(err.properties.part).toEqual('document.xml');
  });

  it('addresses every error when failFast is off', async () => {
    // `failFast: false` still throws an array, as it always has. Each entry now
    // says which file to open and which command to look at — the array itself
    // is a flat list of messages that otherwise names neither.
    const template = await makeDocx({
      body: ['+++missingOne+++', 'text', '+++missingTwo+++'],
    });
    const errors = await getError<TemplateError[]>(() =>
      createReport({ template, data: {}, failFast: false } as never, 'JS')
    );

    expect(Array.isArray(errors)).toBe(true);
    expect(errors.map(e => [e.properties.part, e.properties.command])).toEqual([
      ['document.xml', 'missingOne'],
      ['document.xml', 'missingTwo'],
    ]);
  });

  it('reports a footer error against the footer', async () => {
    const template = await makeDocx({
      body: ['fine'],
      footer: ['+++missingInFooter+++'],
    });
    const err = await getError<TemplateError>(() =>
      createReport({ template, data: {} })
    );
    expect(err.properties.part).toEqual('footer1.xml');
  });

  it('carries a stable id to branch on', async () => {
    const template = await makeDocx({ body: ['+++nope.missing+++'] });
    const err = await errorFrom({ template, data: {} });
    expect(err).toBeInstanceOf(CommandExecutionError);
    expect(err.properties.id).toEqual(ErrorId.commandExecution);
    // The message is prose and may be reworded; the id is the contract.
    expect(err.message).toContain("Error executing command 'nope.missing'");
  });

  it('carries the command that caused it', async () => {
    const template = await makeDocx({ body: ['+++FOR x IN 5+++'] });
    const err = await errorFrom({ template, data: {} });
    expect(err).toBeInstanceOf(InvalidCommandError);
    expect(err.properties.id).toEqual(ErrorId.invalidCommand);
    expect(err.properties.command).toEqual('FOR x IN 5');
  });

  it('explains itself in terms of the template', async () => {
    const template = await makeDocx({ body: ['+++INS ({ a: 1 })+++'] });
    const err = await errorFrom({ template, data: {} });
    expect(err.properties.id).toEqual(ErrorId.objectCommandResult);
    expect(err.properties.explanation).toContain('[object Object]');
  });

  it('gives an IMAGE failure its own id, not the generic one', async () => {
    const template = await makeDocx({ body: ['+++IMAGE badImage()+++'] });
    const err = await errorFrom({
      template,
      data: {},
      additionalJsContext: { badImage: () => ({ width: 1, height: 1 }) },
    });
    expect(err).toBeInstanceOf(ImageError);
    // `ImageError` extends `CommandExecutionError`; the id is what tells them
    // apart without relying on the class hierarchy.
    expect(err).toBeInstanceOf(CommandExecutionError);
    expect(err.properties.id).toEqual(ErrorId.image);
  });

  it('leaves the part unset when the failure precedes rendering', async () => {
    const template = await makeDocx({
      body: ['text'],
      files: { '[Content_Types].xml': null },
    });
    const err = await getError<TemplateError>(() =>
      createReport({ template, data: {} })
    );
    expect(err).toBeInstanceOf(TemplateParseError);
    expect(err.properties.id).toEqual(ErrorId.templateParse);
    // Nothing had been rendered yet, so there is no part to name.
    expect(err.properties.part).toBeUndefined();
  });

  it('makes every error catchable as one type', async () => {
    const template = await makeDocx({ body: ['+++nope.missing+++'] });
    const err = await errorFrom({ template, data: {} });
    expect(err).toBeInstanceOf(TemplateError);
    expect(err).toBeInstanceOf(Error);
  });
});
