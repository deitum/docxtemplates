/**
 * A journal of the public surface, not a prohibition on changing it.
 *
 * Names and shapes are free to change; what must not happen is changing them by
 * accident. A rename shows up as a diff here, gets looked at, and is accepted
 * deliberately (and, if it is user-visible, with a changeset).
 */
import { describe, expect, it } from 'vitest';
import * as browserEntry from '../browser';
import * as mainEntry from '../index';

const surfaceOf = (mod: object) =>
  Object.entries(mod)
    .map(([name, value]) => `${name}: ${typeof value}`)
    .sort()
    .join('\n');

describe('public surface', () => {
  it('main entry point', () => {
    expect(surfaceOf(mainEntry)).toMatchInlineSnapshot(`
      "CommandExecutionError: function
      CommandSyntaxError: function
      ErrorId: object
      ImageError: function
      IncompleteConditionalStatementError: function
      InternalError: function
      InvalidAliasError: function
      InvalidCommandError: function
      InvalidOptionError: function
      NullishCommandResultError: function
      ObjectCommandResultError: function
      TemplateError: function
      TemplateParseError: function
      UnterminatedForLoopError: function
      createReport: function
      default: function
      getMetadata: function
      isError: function
      listCommands: function"
    `);
  });

  it('browser entry point', () => {
    expect(surfaceOf(browserEntry)).toMatchInlineSnapshot(`
      "CommandExecutionError: function
      CommandSyntaxError: function
      ErrorId: object
      ImageError: function
      IncompleteConditionalStatementError: function
      InternalError: function
      InvalidAliasError: function
      InvalidCommandError: function
      InvalidOptionError: function
      NullishCommandResultError: function
      ObjectCommandResultError: function
      TemplateError: function
      TemplateParseError: function
      UnterminatedForLoopError: function
      createReport: function
      default: function
      getMetadata: function
      isError: function
      listCommands: function"
    `);
  });
});
