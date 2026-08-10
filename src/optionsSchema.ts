/**
 * Checking the options a caller passed.
 *
 * `resolveOptions` used to swallow mistakes. The worst of them:
 *
 * ```ts
 * errorHandler: typeof options.errorHandler === 'function' ? options.errorHandler : null,
 * ```
 *
 * Pass anything but a function and the handler simply never ran — no error, no
 * warning, and a report full of the failures it was meant to catch. `template`
 * was not checked at all, so the wrong type surfaced as an opaque complaint
 * from JSZip about not finding the end of a central directory.
 *
 * Types of known options are now checked and rejected. Unknown keys are only
 * logged: TypeScript already catches a typo for anyone who has types, and in
 * plain JavaScript an extra key is more often deliberate than not.
 */
import { logger } from './debug';
import { InvalidOptionError } from './errors';
import { type UserOptions } from './types';

type Check = {
  /** What the option must be, phrased for an error message. */
  expected: string;
  ok: (value: unknown) => boolean;
};

const isString = (v: unknown) => typeof v === 'string';
const isBoolean = (v: unknown) => typeof v === 'boolean';
const isNumber = (v: unknown) => typeof v === 'number' && !Number.isNaN(v);
const isFunction = (v: unknown) => typeof v === 'function';
const isObject = (v: unknown) =>
  typeof v === 'object' && v !== null && !Array.isArray(v);

const check = (expected: string, ok: (value: unknown) => boolean): Check => ({
  expected,
  ok,
});

/** Anything goes: user data and query variables are arbitrary by design. */
const ANYTHING = check('anything', () => true);

const BINARY = check(
  'a Uint8Array (e.g. a Buffer), an ArrayBuffer or a string',
  v => v instanceof Uint8Array || v instanceof ArrayBuffer || isString(v)
);

const DELIMITER = check(
  'a string, or an array of two strings',
  v => isString(v) || (Array.isArray(v) && v.length === 2 && v.every(isString))
);

const ALIAS_MAP = check(
  'an object mapping each alias to a string',
  v => isObject(v) && Object.values(v as object).every(isString)
);

/**
 * One entry per option. Exhaustive over `UserOptions` by construction: leaving
 * one out is a compile error, which is how a new option cannot go unchecked.
 */
const OPTION_CHECKS: { [K in keyof Required<UserOptions>]: Check } = {
  template: BINARY,
  data: ANYTHING,
  queryVars: ANYTHING,
  cmdDelimiter: DELIMITER,
  literalXmlDelimiter: check('a string', isString),
  processLineBreaks: check('a boolean', isBoolean),
  noSandbox: check('a boolean', isBoolean),
  runJs: check('a function', isFunction),
  additionalJsContext: check('an object', isObject),
  failFast: check('a boolean', isBoolean),
  rejectNullish: check('a boolean', isBoolean),
  errorHandler: check('a function', isFunction),
  fixSmartQuotes: check('a boolean', isBoolean),
  processLineBreaksAsNewText: check('a boolean', isBoolean),
  maximumWalkingDepth: check('a number', isNumber),
  indentXml: check('a boolean', isBoolean),
  preserveSpace: check('a boolean', isBoolean),
  compressionLevel: check('a number', isNumber),
  commandAliases: ALIAS_MAP,
  operatorAliases: ALIAS_MAP,
};

const KNOWN_OPTIONS = new Set(Object.keys(OPTION_CHECKS));

/** How a rejected value is shown in the error message. */
const describe = (value: unknown): string => {
  if (value === null) return 'null';
  if (Array.isArray(value)) return 'an array';
  const type = typeof value;
  return type === 'object' ? 'an object' : `a ${type}`;
};

/**
 * Rejects options of the wrong type, and mentions unknown keys in the debug
 * log. `undefined` always passes: leaving an option out and setting it to
 * `undefined` mean the same thing.
 */
export function validateOptions(options: object): void {
  if (!isObject(options)) {
    throw new InvalidOptionError(
      'options',
      'an object',
      describe(options as unknown)
    );
  }

  const unknown: string[] = [];
  for (const [name, value] of Object.entries(options)) {
    if (!KNOWN_OPTIONS.has(name)) {
      unknown.push(name);
      continue;
    }
    if (value === undefined) continue;
    const rule = OPTION_CHECKS[name as keyof UserOptions];
    if (!rule.ok(value)) {
      throw new InvalidOptionError(name, rule.expected, describe(value));
    }
  }

  if (unknown.length > 0) {
    logger.debug(
      `Ignoring unknown option${unknown.length > 1 ? 's' : ''}: ` +
        unknown.join(', ')
    );
  }
}

/** Checks the one option that has no useful default. */
export function validateTemplate(template: unknown): void {
  if (template === undefined || !BINARY.ok(template)) {
    throw new InvalidOptionError(
      'template',
      BINARY.expected,
      template === undefined ? 'nothing' : describe(template)
    );
  }
}
