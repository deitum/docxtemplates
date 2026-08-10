/**
 * The errors this library throws.
 *
 * Every one of them carries `properties`: a stable `id` to branch on instead of
 * matching against message text, the part of the package it came from, the
 * command that caused it, and an explanation aimed at whoever wrote the
 * template. Reports are rendered from several parts — `document.xml`,
 * `header1.xml`, `footer1.xml` — and with `failFast: false` the errors of all of
 * them arrive in one array, so "which file" is the first thing you need and the
 * message alone cannot tell you.
 *
 * Messages, class names and the shape of what is thrown are unchanged; the
 * properties are additive.
 */
import { type LoopStatus } from './types';

export function isError(err: unknown): err is Error {
  return (
    err instanceof Error ||
    (typeof err === 'object' && !!err && 'name' in err && 'message' in err)
  );
}

/**
 * Machine-readable error codes. Stable across releases: branch on these rather
 * than on `message`, which is prose and may be reworded.
 */
export const ErrorId = {
  nullishCommandResult: 'nullish_command_result',
  objectCommandResult: 'object_command_result',
  commandSyntax: 'command_syntax',
  invalidCommand: 'invalid_command',
  invalidAlias: 'invalid_alias',
  commandExecution: 'command_execution',
  image: 'image',
  internal: 'internal',
  templateParse: 'template_parse',
  incompleteConditional: 'incomplete_conditional',
  unterminatedForLoop: 'unterminated_for_loop',
} as const;

export type ErrorId = (typeof ErrorId)[keyof typeof ErrorId];

export type ErrorProperties = {
  /** A stable code; see {@link ErrorId}. */
  id: ErrorId;
  /**
   * The part of the .docx the error came from — `document.xml`, `header1.xml`,
   * and so on. Filled in by `createReport` once it knows which part it was
   * rendering, so it is absent on errors raised before rendering starts.
   */
  part?: string;
  /** The command that caused it, as it appears in the template. */
  command?: string;
  /** What went wrong, in terms of the template rather than of the engine. */
  explanation: string;
};

/**
 * The base of every error raised here. Existing code keeps working unchanged —
 * the classes, their names, their messages and their own fields are all as they
 * were — but anything can now be caught as one type and inspected uniformly.
 */
export class TemplateError extends Error {
  properties: ErrorProperties;

  constructor(message: string, properties: ErrorProperties) {
    super(message);
    // `new.target`, not `TemplateError`, so that subclasses keep their own
    // prototype and `instanceof` works for them after bundling.
    Object.setPrototypeOf(this, new.target.prototype);
    this.properties = properties;
  }
}

/**
 * Records which part of the package an error came from. Applied in one place,
 * once the part being rendered is known.
 */
export function withPart<T>(err: T, part: string): T {
  for (const one of Array.isArray(err) ? err : [err]) {
    if (one instanceof TemplateError && one.properties.part == null) {
      one.properties.part = part;
    }
  }
  return err;
}

/**
 * Thrown when `rejectNullish` is set to `true` and a command returns `null` or `undefined`.
 */
export class NullishCommandResultError extends TemplateError {
  command: string;
  constructor(command: string) {
    super(
      `Result of command ${command} is null or undefined and rejectNullish is set`,
      {
        id: ErrorId.nullishCommandResult,
        command,
        explanation:
          'The command returned nothing. With `rejectNullish` set that is ' +
          'treated as a mistake, on the assumption that a value was meant to ' +
          'be there.',
      }
    );
    Object.setPrototypeOf(this, NullishCommandResultError.prototype);
    this.command = command;
  }
}

/**
 * Thrown when the result of an `INS` command is an `Object`. This ensures you don't accidentally put `'[object Object]'` in your report.
 */
export class ObjectCommandResultError extends TemplateError {
  command: string;
  result: unknown;
  constructor(command: string, result: unknown) {
    super(`Result of command '${command}' is an object`, {
      id: ErrorId.objectCommandResult,
      command,
      explanation:
        'Inserting an object would put the text `[object Object]` in the ' +
        'report. Pick a property of it, or format it into a string first.',
    });
    Object.setPrototypeOf(this, ObjectCommandResultError.prototype);
    this.command = command;
    this.result = result;
  }
}

export class CommandSyntaxError extends TemplateError {
  command: string;
  constructor(command: string) {
    super(`Invalid command syntax: ${command}`, {
      id: ErrorId.commandSyntax,
      command,
      explanation:
        'The text between the delimiters does not name a command, and could ' +
        'not be read as an expression to insert either.',
    });
    Object.setPrototypeOf(this, CommandSyntaxError.prototype);
    this.command = command;
  }
}

export class InvalidCommandError extends TemplateError {
  command: string;
  constructor(msg: string, command: string) {
    super(`${msg}: ${command}`, {
      id: ErrorId.invalidCommand,
      command,
      explanation: msg,
    });
    Object.setPrototypeOf(this, InvalidCommandError.prototype);
    this.command = command;
  }
}

/**
 * Thrown when the `commandAliases` or `operatorAliases` options are invalid.
 */
export class InvalidAliasError extends TemplateError {
  alias: string;
  constructor(msg: string, alias: string) {
    super(`${msg}: ${alias}`, {
      id: ErrorId.invalidAlias,
      explanation: msg,
    });
    Object.setPrototypeOf(this, InvalidAliasError.prototype);
    this.alias = alias;
  }
}

export class CommandExecutionError extends TemplateError {
  command: string;
  err: Error;
  constructor(err: Error, command: string) {
    super(`Error executing command '${command}': ${err.name}: ${err.message}`, {
      id: ErrorId.commandExecution,
      command,
      explanation:
        'The JavaScript in the command threw. Usually the data the template ' +
        'expects is shaped differently, or is not there at all.',
    });
    // `new.target`, not `CommandExecutionError`, so that subclasses (e.g.
    // `ImageError`) keep their own prototype and `instanceof` works for them.
    Object.setPrototypeOf(this, new.target.prototype);
    this.command = command;
    this.err = err;
  }
}

export class ImageError extends CommandExecutionError {
  constructor(err: Error, command: string) {
    super(err, command);
    this.properties = {
      ...this.properties,
      id: ErrorId.image,
      explanation:
        'The IMAGE command returned something that could not be embedded. It ' +
        'must be an object with `width`, `height`, `data` and a supported ' +
        '`extension`.',
    };
  }
}

export class InternalError extends TemplateError {
  constructor(msg: string) {
    super(`INTERNAL ERROR: ${msg}`, {
      id: ErrorId.internal,
      explanation:
        'The engine reached a state it believes to be impossible. This is a ' +
        'bug in the library rather than in the template.',
    });
    Object.setPrototypeOf(this, InternalError.prototype);
  }
}

export class TemplateParseError extends TemplateError {
  constructor(msg: string) {
    super(msg, {
      id: ErrorId.templateParse,
      explanation:
        'The .docx could not be read. Either it is not a Word file, or the ' +
        'part the template was expected in is missing or malformed.',
    });
    Object.setPrototypeOf(this, TemplateParseError.prototype);
  }
}

export class IncompleteConditionalStatementError extends TemplateError {
  constructor() {
    super(
      'Incomplete IF/END-IF statement. Make sure each IF-statement has a corresponding END-IF command.',
      {
        id: ErrorId.incompleteConditional,
        explanation:
          'The template has a different number of IF and END-IF commands.',
      }
    );
    Object.setPrototypeOf(this, IncompleteConditionalStatementError.prototype);
  }
}

export class UnterminatedForLoopError extends TemplateError {
  constructor(loop: LoopStatus) {
    super(
      `Unterminated FOR-loop ('FOR ${loop.varName}'). Make sure each FOR loop has a corresponding END-FOR command.`,
      {
        id: ErrorId.unterminatedForLoop,
        command: `FOR ${loop.varName}`,
        explanation:
          'The template reached its end with a FOR loop still open. Every ' +
          'FOR needs a matching END-FOR.',
      }
    );
    Object.setPrototypeOf(this, UnterminatedForLoopError.prototype);
  }
}
