// ==========================================
// Docx nodes

import { type QualifiedAttribute } from 'sax';

// ==========================================
export type Node = TextNode | NonTextNode;
type BaseNode = {
  _parent?: Node;
  _children: Array<Node>;
  _ifName?: string;
};
export type TextNode = BaseNode & {
  _fTextNode: true;
  _text: string;
};
export type NonTextNode = BaseNode & {
  _fTextNode: false;
  _tag: string;

  // Simplified; only need this property
  _attrs: { [key: string]: QualifiedAttribute | string } & {
    Extension?: string;
    ContentType?: string;
    PartName?: string;
  };
};

// ==========================================
// Report creator
// ==========================================

/**
 * Binary contents accepted wherever a docx file (or one of the XML files
 * inside it) is read or written: a `Uint8Array` — which includes a NodeJS
 * `Buffer` — a raw `ArrayBuffer`, or a string.
 */
export type ZipInput = Uint8Array | ArrayBuffer | string;

export type ReportData = any;
export type QueryResolver = (
  query: string | undefined,
  queryVars: any
) => ReportData | Promise<ReportData>;

type ErrorHandler = (e: Error, raw_code?: string) => any;

type RunJSFunc = (o: { sandbox: SandBox; ctx: Context }) => {
  modifiedSandbox: SandBox;
  result: unknown;
};

export type UserOptions = {
  /**
   * Docx file template as a Uint8Array (or e.g. ArrayBuffer or NodeJS Buffer).
   */
  template: Uint8Array;
  /**
   * Object of data to be injected or a (async) function that resolves to the data. The function gets as an argument the contents of the QUERY command as a string.
   */
  data?: ReportData | QueryResolver;
  /**
   * Gets injected into data function as second argument.
   */
  queryVars?: any;
  /**
   * Defines a custom command delimiter. This can be a String e.g. '+++' or an Array of Strings with length 2: ['{', '}'] in which the first element serves as the start delimiter and the second as the end delimiter.
   */
  cmdDelimiter?: string | [string, string];
  /**
   * The delimiter that's used to indicate literal XML that should be inserted into the docx XML tree as-is. Defaults to `||`.
   */
  literalXmlDelimiter?: string;
  /**
   * Handle linebreaks in result of commands as actual linebreaks (Default: true)
   */
  processLineBreaks?: boolean; // true by default
  /**
   * INSECURE: Set this option to true to disable running all commands in a new JS-VM. USE ONLY WITH TRUSTED TEMPLATES. Beware of arbitrary code injection risks. Can slightly improve performance on complex templates.
   */
  noSandbox?: boolean;
  /**
   * Custom sandbox. See documentation for details.
   */
  runJs?: RunJSFunc;
  /**
   * Add functions or other static data to this option to have access to it in your commands.
   *
   * ```js
   * additionalJsContext: {
   *   qrCode: url => {
   *     const dataUrl = createQrImage(url, { size: 500 });
   *     const data = dataUrl.slice('data:image/gif;base64,'.length);
   *     return { width: 6, height: 6, data, extension: '.gif' };
   *   },
   * }
   * ```
   */
  additionalJsContext?: object;
  /**
   * Whether to fail on the first error encountered in the template. Defaults to true. Can be used to collect all errors in a template (e.g. misspelled commands) before failing.
   */
  failFast?: boolean;

  /**
   * When set to `true`, this setting ensures `createReport` throws a `NullishCommandResultError` when the result of an INS, HTML, IMAGE, or LINK command is `null` or `undefined`. This is useful as nullish return values usually indicate a mistake in the template or the invoking code. Defaults to `false`.
   */
  rejectNullish?: boolean;

  /**
   * Custom error handler to catch any errors that may occur evaluating commands in the template. The value returned from this handler will be inserted into the template instead.
   */
  errorHandler?: ErrorHandler;

  /**
   * MS Word usually autocorrects JS string literal quotes with unicode 'smart' quotes ('curly' quotes). E.g. 'aubergine' -> ‘aubergine’.
   * This causes an error when evaluating commands containing these smart quotes, as they are not valid JavaScript.
   * If you set fixSmartQuotes to 'true', these smart quotes will automatically get replaced with straight quotes (') before command evaluation.
   * Defaults to false.
   */
  fixSmartQuotes?: boolean;

  /**
   * Use the new way of injecting line breaks from command results (only applies when `processLineBreaks` is `true`)
   * which has better results in LibreOffice and Google Drive.
   * (Default: false)
   */
  processLineBreaksAsNewText?: boolean;

  /**
   * Maximum loop iterations allowed when walking through the template.
   * You can increase this to generate reports with large amount of FOR loop elements.
   * Tip: You can disable infinite loop protection by using the `Infinity` constant.
   * This may be useful if you implement a process timeout instead.
   * (Default: 1,000,000)
   */
  maximumWalkingDepth?: number;
  /**
   * Whether to indent the generated XML to make it more human-readable.
   * Tip: Set this to false if you want to minimize the size of the generated docx file.
   * (Default: true)
   */
  indentXml?: boolean;
  /**
   * Whether to preserve whitespace in the generated XML.*
   * Tip: Set this to false if you want to minimize the size of the generated docx file.
   * (Default: true)
   */
  preserveSpace?: boolean;
  /**
   * Compression level for the generated docx file.
   * Integer between 0 (no compression, fastest) and 9 (maximum compression, slowest).
   * (Default: 1)
   */
  compressionLevel?: number;

  /**
   * Alternative names for the built-in commands, so that templates can be written
   * in the template author's own language. Keys are the aliases (they may contain
   * several words and are matched case-insensitively), values must be one of the
   * built-in commands (`IF`, `FOR`, `END-FOR`, `INS`, ...).
   *
   * ```js
   * commandAliases: {
   *   'ЕСЛИ': 'IF',
   *   'ИНАЧЕ ЕСЛИ': 'ELSE-IF',
   *   'ИНАЧЕ': 'ELSE',
   *   'КОНЕЦ ЕСЛИ': 'END-IF',
   * }
   * ```
   */
  commandAliases?: { [alias: string]: string };

  /**
   * Alternative names for the operators and keywords used inside command
   * expressions, so that e.g. `+++значение1 больше значение2+++` can be written
   * instead of `+++значение1 > значение2+++`. Keys are the aliases (they may
   * contain several words and are matched case-insensitively), values are the
   * JS snippet they get replaced with. Aliases are only substituted when they
   * appear as whole words outside of string literals.
   *
   * ```js
   * operatorAliases: {
   *   'больше или равно': '>=',
   *   'больше': '>',
   *   'равно': '===',
   *   'и': '&&',
   * }
   * ```
   */
  operatorAliases?: { [alias: string]: string };
};

/**
 * A user-provided alias map, compiled for matching: the alias is split into
 * lowercased words, and the list is sorted longest-alias-first.
 */
export type AliasList = Array<{ tokens: string[]; replacement: string }>;

export type CreateReportOptions = {
  cmdDelimiter: [string, string];
  literalXmlDelimiter: string;
  processLineBreaks: boolean;
  noSandbox: boolean;
  // Explicit `| undefined`: these are passed straight through from `UserOptions`,
  // where leaving them out and setting them to `undefined` mean the same thing.
  runJs?: RunJSFunc | undefined;
  additionalJsContext: object;
  failFast: boolean;
  rejectNullish: boolean;
  errorHandler: ErrorHandler | null;
  fixSmartQuotes: boolean;
  processLineBreaksAsNewText: boolean;
  maximumWalkingDepth?: number | undefined;
  indentXml: boolean;
  preserveSpace: boolean;
  compressionLevel: number;
  commandAliases: AliasList;
  operatorAliases: AliasList;
};

export type SandBox = {
  __code__: string | undefined;
  __result__: unknown | undefined;
  [k: string]: unknown;
};

export type Context = {
  gCntIf: number;
  gCntEndIf: number;
  level: number;
  fCmd: boolean;
  cmd: string;
  fSeekQuery: boolean;
  query?: string;
  buffers: {
    'w:p': BufferStatus;
    'w:tr': BufferStatus;
    'w:tc': BufferStatus;
  };
  cell?: CellStatus;
  pendingImageNode?: { image: NonTextNode; caption?: NonTextNode[] };
  imageAndShapeIdIncrement: number;
  images: Images;
  pendingLinkNode?: NonTextNode;
  linkId: number;
  links: Links;
  pendingHtmlNode?: TextNode | NonTextNode;
  htmlId: number;
  htmls: Htmls;
  vars: { [name: string]: VarValue };
  loops: Array<LoopStatus>;
  fJump: boolean;
  shorthands: { [shorthand: string]: string };
  options: CreateReportOptions;
  jsSandbox?: SandBox;
  textRunPropsNode?: NonTextNode;

  // To verfiy we don't have a nested if within the same p or tr tag
  pIfCheckMap: Map<Node, string>;
  trIfCheckMap: Map<Node, string>;
};

export type Images = { [id: string]: Image };
export const ImageExtensions = [
  '.png',
  '.gif',
  '.jpg',
  '.jpeg',
  '.svg',
] as const;
type ImageExtension = (typeof ImageExtensions)[number];
export type Image = {
  extension: ImageExtension;
  data: ZipInput;
};
export type Links = { [id: string]: Link };
type Link = { url: string };
export type Htmls = { [id: string]: string };

type BufferStatus = {
  text: string;
  cmds: string;
  fInsertedText: boolean;
};

/** The table cell (`w:tc`) that is currently being walked. */
export type CellStatus = {
  /** The cell node, in the input tree. */
  node: Node;
  /**
   * Whether the commands in the cell are part of a FOR/IF construct that starts
   * or ends in another cell (as in the dynamic-columns pattern). Only such a
   * cell is deleted when it renders to nothing.
   */
  fSpansCells: boolean;
};

type VarValue = unknown;
export type LoopStatus = {
  refNode: Node;
  refNodeLevel: number;
  varName: string;
  loopOver: Array<VarValue>;
  idx: number;
  isIf?: boolean;

  // The following fields are only used by IF loops, to support
  // IF / ELSE-IF / ELSE / END-IF constructs. They are reset back to `undefined`
  // between passes, hence the explicit `| undefined`.

  /** Index of the branch that is currently being walked (0 is the IF branch). */
  ifCurrentBranch?: number | undefined;
  /** Index of the branch whose condition evaluated truthy (-1 if none). */
  ifActiveBranch?: number | undefined;
  /** Whether one of the branches has already been selected. */
  ifBranchTaken?: boolean | undefined;
  /** Index of the final (unconditional) ELSE branch, if it has been seen already. */
  ifElseBranch?: number | undefined;
};

export type ImagePars = {
  /**
   * Desired width of the image in centimeters.
   */
  width: number;

  /**
   * Desired height of the image in centimeters.
   */
  height: number;

  /**
   * The image data, as a Uint8Array (e.g. a NodeJS Buffer), an ArrayBuffer, or
   * a base64-encoded string.
   */
  data: ZipInput;

  /**
   * Optional. When injecting an SVG image, a fallback non-SVG (png/jpg/gif, etc.) image can be provided. This thumbnail is used when SVG images are not supported (e.g. older versions of Word) or when the document is previewed by e.g. Windows Explorer. See usage example below.
   */
  thumbnail?: Image;

  /**
   * One of '.png', '.gif', '.jpg', '.jpeg', '.svg'.
   */
  extension: ImageExtension;

  /**
   * Optional alt text.
   */
  alt?: string;

  /**
   * Optional rotation in degrees, with positive angles moving clockwise.
   */
  rotation?: number;

  /**
   * Optional caption
   */
  caption?: string;
};

export type LinkPars = {
  url: string;
  label?: string;
};

export type CommandSummary = {
  raw: string;
  type: BuiltInCommand;
  code: string;
};

export type BuiltInCommand = (typeof BUILT_IN_COMMANDS)[number];

export const BUILT_IN_COMMANDS = [
  'QUERY',
  'CMD_NODE',
  'ALIAS',
  'FOR',
  'END-FOR',
  'IF',
  'ELSE-IF',
  'ELSE',
  'END-IF',
  'INS',
  'EXEC',
  'IMAGE',
  'LINK',
  'HTML',
] as const;
