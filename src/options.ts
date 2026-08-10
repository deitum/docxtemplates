/**
 * The defaults behind every `UserOptions` field, and the one place that turns a
 * user's (mostly optional) options into the fully-resolved `CreateReportOptions`
 * the engine works with. `createReport` and `listCommands` both go through it,
 * so the two cannot drift apart.
 */
import { compileAliases, compileCommandAliases } from './aliases';
import { validateOptions } from './optionsSchema';
import { type CreateReportOptions, type UserOptions } from './types';

export const DEFAULT_CMD_DELIMITER = '+++' as const;
export const DEFAULT_LITERAL_XML_DELIMITER = '||' as const;

/**
 * Deflate level for the generated .docx: 0 is store-only, 9 is smallest. The
 * default trades a slightly bigger file for a much faster save, which is what
 * report generation wants.
 */
export const DEFAULT_COMPRESSION_LEVEL = 1;

/** Emergency brake on the template walk; see `UserOptions.maximumWalkingDepth`. */
export const DEFAULT_MAXIMUM_WALKING_DEPTH = 1_000_000;

export const OPTION_DEFAULTS = {
  processLineBreaks: true,
  processLineBreaksAsNewText: false,
  noSandbox: false,
  failFast: true,
  rejectNullish: false,
  fixSmartQuotes: false,
  indentXml: true,
  preserveSpace: true,
  compressionLevel: DEFAULT_COMPRESSION_LEVEL,
} as const;

/** Normalises the delimiter option into its `[open, close]` form. */
export const getCmdDelimiter = (
  delimiter?: string | [string, string]
): [string, string] => {
  if (!delimiter) return [DEFAULT_CMD_DELIMITER, DEFAULT_CMD_DELIMITER];
  if (typeof delimiter === 'string') return [delimiter, delimiter];
  return delimiter;
};

/**
 * Everything a caller may set, with an explicit `| undefined` on each field:
 * under `exactOptionalPropertyTypes`, "absent" and "present but undefined" are
 * different types, and callers here pass the latter.
 */
type ResolvableOptions = {
  [K in keyof Omit<UserOptions, 'template'>]?:
    Omit<UserOptions, 'template'>[K] | undefined;
};

/** Fills in every default and compiles the alias maps. */
export function resolveOptions(
  options: ResolvableOptions
): CreateReportOptions {
  validateOptions(options);
  return {
    cmdDelimiter: getCmdDelimiter(options.cmdDelimiter),
    literalXmlDelimiter:
      options.literalXmlDelimiter || DEFAULT_LITERAL_XML_DELIMITER,
    processLineBreaks:
      options.processLineBreaks ?? OPTION_DEFAULTS.processLineBreaks,
    processLineBreaksAsNewText:
      options.processLineBreaksAsNewText ??
      OPTION_DEFAULTS.processLineBreaksAsNewText,
    noSandbox: options.noSandbox ?? OPTION_DEFAULTS.noSandbox,
    runJs: options.runJs,
    additionalJsContext: options.additionalJsContext ?? {},
    failFast: options.failFast ?? OPTION_DEFAULTS.failFast,
    rejectNullish: options.rejectNullish ?? OPTION_DEFAULTS.rejectNullish,
    errorHandler: options.errorHandler ?? null,
    fixSmartQuotes: options.fixSmartQuotes ?? OPTION_DEFAULTS.fixSmartQuotes,
    maximumWalkingDepth: options.maximumWalkingDepth,
    indentXml: options.indentXml ?? OPTION_DEFAULTS.indentXml,
    preserveSpace: options.preserveSpace ?? OPTION_DEFAULTS.preserveSpace,
    compressionLevel:
      options.compressionLevel ?? OPTION_DEFAULTS.compressionLevel,
    commandAliases: compileCommandAliases(options.commandAliases),
    operatorAliases: compileAliases(options.operatorAliases),
  };
}
