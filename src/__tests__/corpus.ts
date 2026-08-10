/**
 * The golden corpus: every fixture, with the data and options it is meant to be
 * rendered with.
 *
 * This is the safety net the architecture work leans on. Unlike the `_probe`
 * snapshots — which only ever see `word/document.xml` — the corpus test renders
 * the *whole package* and records every entry of the zip, so that regressions in
 * headers, footers, `.rels` parts, `[Content_Types].xml`, embedded media and
 * even the order of the zip entries are caught.
 *
 * A case that is expected to fail is just as valuable as one that renders: the
 * error it produces is recorded too.
 */
import fs from 'node:fs';
import { type ImagePars, type UserOptions } from '../types';
import { fixturePath, fixturesDir } from './helpers';

export type CorpusCase = {
  /** File name inside `fixtures/`. */
  file: string;
  /**
   * Name of the snapshot, when one fixture is rendered under more than one
   * configuration. Defaults to `file`.
   */
  name?: string;
  /** Everything but `template`; the corpus test reads the file itself. */
  options: Omit<UserOptions, 'template'>;
};

/** The snapshot a case is recorded under. */
export const caseName = (c: CorpusCase): string => c.name ?? c.file;

// ==========================================
// Fixed inputs
// ==========================================

const readSample = (name: string) => fs.readFileSync(fixturePath(name));

/**
 * A pre-encoded 1x1 PNG. Fixtures that call `qr()` get this instead of a real
 * QR code: the corpus records bytes, and must not depend on what version of the
 * `qrcode` package happens to be installed.
 */
const TINY_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const pngImage = (): ImagePars => ({
  width: 6,
  height: 6,
  data: readSample('sample.png'),
  extension: '.png',
});

const jpgImage = (): ImagePars => ({
  width: 6,
  height: 6,
  data: readSample('sample.jpg'),
  extension: '.jpg',
});

const svgImage = (): ImagePars => ({
  width: 6,
  height: 6,
  data: readSample('sample.svg'),
  extension: '.svg',
  thumbnail: { data: readSample('sample.png'), extension: '.png' },
});

const COMPANIES = [{ name: 'FIRST' }, { name: 'SECOND' }, { name: 'THIRD' }];

const COMPANIES_WITH_PEOPLE = [
  {
    name: 'FIRST',
    people: [
      { firstName: 'Pep', projects: [{ name: 'Project 1' }] },
      { firstName: 'Fina', projects: [] },
    ],
  },
  { name: 'SECOND', people: [{ firstName: 'Manel', projects: [] }] },
];

const OPERATOR_ALIASES = {
  'больше или равно': '>=',
  'меньше или равно': '<=',
  'не равно': '!==',
  больше: '>',
  меньше: '<',
  равно: '===',
  и: '&&',
  или: '||',
  ИЗ: 'IN',
};

const COMMAND_ALIASES = {
  ЕСЛИ: 'IF',
  'ИНАЧЕ ЕСЛИ': 'ELSE-IF',
  ИНАЧЕ: 'ELSE',
  'КОНЕЦ ЕСЛИ': 'END-IF',
  ДЛЯ: 'FOR',
  'КОНЕЦ ДЛЯ': 'END-FOR',
};

/**
 * The functions fixtures reach for through `additionalJsContext`. Every one of
 * them is pure and returns fixed bytes, so the report is reproducible.
 */
const JS_CONTEXT = {
  injectImg: (caption?: boolean) =>
    caption ? { ...pngImage(), caption: 'Sample caption' } : pngImage(),
  injectSvg: () => svgImage(),
  svgImgFile: () => svgImage(),
  svgImgStr: () => ({
    ...svgImage(),
    data: Buffer.from(
      '<svg xmlns="http://www.w3.org/2000/svg"><rect x="10" y="10" ' +
        'height="100" width="100" style="stroke:#ff0000; fill: #0000ff"/></svg>',
      'utf-8'
    ),
  }),
  image: () => pngImage(),
  getImage: () => pngImage(),
  getImage45: () => ({ ...pngImage(), rotation: 45 }),
  getImage180: () => ({ ...pngImage(), rotation: 180 }),
  qr: () => ({
    width: 6,
    height: 6,
    data: TINY_PNG_BASE64,
    extension: '.png' as const,
  }),
  toLowerCase: (str: string) => String(str).toLowerCase(),
  formatNumber: (n: number) => `#${n}`,
};

// ==========================================
// The cases
// ==========================================

const LONG_TEXT =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit.\n' +
  'Sed commodo sagittis erat, sed vehicula lorem molestie et.\n' +
  'Fusce ut scelerisque neque. Donec porta eleifend dolor.';

/** Data shared by the many fixtures that only differ in their structure. */
const DEFAULT_DATA = {
  companies: COMPANIES,
  a: true,
  b: false,
  value: 7,
  list: ['x', 'y'],
  items: [
    { name: 'big one', big: true },
    { name: 'second', big: false },
    { name: 'third', big: false },
  ],
};

export const CORPUS: CorpusCase[] = [
  {
    file: 'aliasCommands.docx',
    options: {
      data: {
        значение: 7,
        товары: [{ название: 'Молоток' }, { название: 'Отвёртка' }],
      },
      commandAliases: COMMAND_ALIASES,
      operatorAliases: OPERATOR_ALIASES,
    },
  },
  {
    file: 'aliasOperators.docx',
    options: { data: { a: 3, b: 1 }, operatorAliases: OPERATOR_ALIASES },
  },
  { file: 'anchor-empty.docx', options: { data: {} } },
  { file: 'commandExecutionError.docx', options: { data: {} } },
  {
    file: 'confusingCommandNames.docx',
    options: {
      data: { something: 'a thing', INSertable: 'insertable', ...DEFAULT_DATA },
      additionalJsContext: JS_CONTEXT,
    },
  },
  {
    file: 'dynamic-columns-with-dynamic-rows.docx',
    options: { data: { rows: ['r1', 'r2'], columns: ['c1', 'c2', 'c3'] } },
  },
  {
    file: 'dynamic-columns.docx',
    options: { data: { columns: ['c1', 'c2', 'c3'] } },
  },
  { file: 'elseOutsideIf.docx', options: { data: {} } },
  { file: 'exec.docx', options: { data: {} } },
  { file: 'execAndIns.docx', options: { data: {} } },
  { file: 'execPromise.docx', options: { data: {} } },
  { file: 'execShorthand.docx', options: { data: {} } },
  {
    file: 'execWithContext.docx',
    options: {
      data: { companies: COMPANIES },
      additionalJsContext: JS_CONTEXT,
    },
  },
  {
    file: 'existingUppercaseJPEGExtension.docx',
    options: {
      data: {},
      additionalJsContext: { ...JS_CONTEXT, injectImg: () => jpgImage() },
    },
  },
  { file: 'falsy-block.docx', options: { data: {} } },
  { file: 'fixSmartQuotes.docx', options: { data: {}, fixSmartQuotes: true } },
  { file: 'for-row1.docx', options: { data: { companies: COMPANIES } } },
  { file: 'for1.docx', options: { data: { companies: COMPANIES } } },
  { file: 'for1alias.docx', options: { data: { companies: COMPANIES } } },
  {
    file: 'for1customDelimiter.docx',
    options: { data: { companies: COMPANIES }, cmdDelimiter: '***' },
  },
  { file: 'for1inline.docx', options: { data: { companies: COMPANIES } } },
  {
    file: 'for1inlineWithSpaces.docx',
    options: { data: { companies: COMPANIES } },
  },
  { file: 'for1js.docx', options: { data: { companies: COMPANIES } } },
  {
    file: 'for1scalars.docx',
    options: { data: { companies: ['one', 'two', 'three'] } },
  },
  {
    file: 'for2.docx',
    options: { data: { companies: COMPANIES_WITH_PEOPLE } },
  },
  {
    file: 'for3.docx',
    options: { data: { companies: COMPANIES_WITH_PEOPLE } },
  },
  {
    file: 'forLoopWithTextBox.docx',
    options: { data: { companies: COMPANIES } },
  },
  {
    file: 'forOverObject.docx',
    options: { data: { companies: ['alpha', 'beta'] } },
  },
  {
    file: 'forOverObjectKeys.docx',
    options: { data: { companies: { first: 'FIRST', second: 'SECOND' } } },
  },
  {
    file: 'forWithIdx.docx',
    options: {
      data: {
        companies: [
          { name: 'FIRST', executives: ['Ann', 'Bob'] },
          { name: 'SECOND', executives: ['Cid'] },
        ],
      },
    },
  },
  { file: 'htmls.docx', options: { data: {} } },
  { file: 'if-row1.docx', options: { data: {} } },
  { file: 'if.docx', options: { data: {} } },
  { file: 'if2.docx', options: { data: {} } },
  { file: 'ifDoubleElse.docx', options: { data: { value: true } } },
  { file: 'ifElse.docx', options: { data: { value: 7 } } },
  { file: 'ifElseFor.docx', options: { data: DEFAULT_DATA } },
  {
    file: 'ifElseForInside.docx',
    options: { data: { a: false, list: ['x', 'y', 'z'] } },
  },
  { file: 'ifElseIf.docx', options: { data: { value: 7 } } },
  { file: 'ifElseInline.docx', options: { data: { value: 7 } } },
  { file: 'ifElseNested.docx', options: { data: { a: true, b: false } } },
  { file: 'ifElseRow.docx', options: { data: { value: 3 } } },
  { file: 'ifElseSkippedBranch.docx', options: { data: { a: true } } },
  { file: 'ifInline.docx', options: { data: {} } },
  // Two IF constructs on one line / row: these are expected to be rejected, and
  // the message they are rejected with is the thing worth pinning down.
  {
    file: 'ifStatementsOnSameLine.docx',
    options: {
      data: { a: true, b: true, counts: ['a', 'b', 'c'] },
      cmdDelimiter: ['{{', '}}'],
    },
  },
  {
    file: 'ifStatementsOnSameRow1.docx',
    options: {
      data: { a: true, b: true, counts: ['a', 'b', 'c'] },
      cmdDelimiter: ['{{', '}}'],
    },
  },
  {
    file: 'ifStatementsOnSameRow2.docx',
    options: {
      data: { a: true, b: true, counts: ['a', 'b', 'c'] },
      cmdDelimiter: ['{{', '}}'],
    },
  },
  {
    file: 'imageBase64.docx',
    options: { data: {}, additionalJsContext: JS_CONTEXT },
  },
  {
    file: 'imageCaption.docx',
    options: { data: {}, additionalJsContext: JS_CONTEXT },
  },
  {
    file: 'imageExisting.docx',
    options: {
      data: { cv: { ProfilePicture: { url: 'http://example.com/p.png' } } },
      additionalJsContext: JS_CONTEXT,
    },
  },
  { file: 'imageExistingMultiple.docx', options: { data: {} } },
  {
    file: 'imageHeader.docx',
    options: { data: {}, additionalJsContext: JS_CONTEXT },
  },
  {
    file: 'imageInShapeInFooter.docx',
    options: { data: {}, additionalJsContext: JS_CONTEXT },
  },
  // The same fixture holds the same IMAGE command under two different
  // delimiters, one of which is inert on any given run.
  {
    file: 'imageMultiDelimiter.docx',
    name: 'imageMultiDelimiter.docx (+++)',
    options: { data: {}, additionalJsContext: JS_CONTEXT, cmdDelimiter: '+++' },
  },
  {
    file: 'imageMultiDelimiter.docx',
    name: 'imageMultiDelimiter.docx (---)',
    options: { data: {}, additionalJsContext: JS_CONTEXT, cmdDelimiter: '---' },
  },
  {
    file: 'imageRotation.docx',
    options: { data: {}, additionalJsContext: JS_CONTEXT },
  },
  {
    file: 'imageSimple.docx',
    options: { data: {}, additionalJsContext: JS_CONTEXT },
  },
  {
    file: 'imagesSVG.docx',
    options: { data: {}, additionalJsContext: JS_CONTEXT },
  },
  {
    file: 'insJsComplex.docx',
    options: { data: { companies: ['one', 'two'] } },
  },
  { file: 'insJsSimple.docx', options: { data: {} } },
  { file: 'insJsWithLoops.docx', options: { data: { companies: COMPANIES } } },
  { file: 'insOmitted.docx', options: { data: { companies: COMPANIES } } },
  { file: 'insShorthand.docx', options: { data: { companies: COMPANIES } } },
  {
    file: 'insertArray.docx',
    options: { data: { companies: ['one', 'two'] } },
  },
  {
    file: 'insertInHeaderAndFooter.docx',
    options: {
      data: {
        body_command: 'in the body',
        header_command: 'in the header',
        footer_command: 'in the footer',
      },
    },
  },
  { file: 'invalidCommand.docx', options: { data: {} } },
  { file: 'invalidFor.docx', options: { data: { companies: COMPANIES } } },
  { file: 'invalidForCmd.docx', options: { data: {} } },
  { file: 'invalidIf.docx', options: { data: { companies: COMPANIES } } },
  {
    file: 'invalidMultipleErrors.docx',
    options: { data: { companies: COMPANIES }, failFast: false },
  },
  {
    file: 'link-regression-issue-133.docx',
    options: {
      data: {
        links: [
          { url: 'https://www.google.com/', name: 'Google' },
          { url: 'https://www.youtube.com/', name: 'Youtube' },
        ],
      },
    },
  },
  {
    file: 'link-regression-issue-83.docx',
    options: { data: { companies: COMPANIES } },
  },
  { file: 'links.docx', options: { data: {} } },
  {
    file: 'literalXml.docx',
    options: { data: { text: 'foo||<w:br/>||bar' } },
  },
  { file: 'longText.docx', options: { data: { longText: LONG_TEXT } } },
  { file: 'macroEnabledTemplate.docm', options: { data: {} } },
  { file: 'missingEndFor.docx', options: { data: {} } },
  { file: 'missingEndIf.docx', options: { data: {} } },
  {
    file: 'nestedInlineForLoopWithSurroundingText.docx',
    options: {
      data: {
        companies: COMPANIES,
        products: [{ name: 'Hammer' }, { name: 'Nail' }],
      },
    },
  },
  {
    file: 'newlineInVariableIssue143.docx',
    options: { data: { headline: 'first line\nsecond line' } },
  },
  { file: 'noQuery.docx', options: { data: {} } },
  {
    file: 'noQueryBrackets.docx',
    options: { data: { a: 'foo', b: 'bar' }, cmdDelimiter: ['{', '}'] },
  },
  {
    file: 'noQuerySimpleInserts.docx',
    options: { data: { a: 'foo', b: 'bar' } },
  },
  {
    file: 'nonAlphaCommandNames1.docx',
    options: { data: { 姓名: 'hong', 标题: 'junyao' } },
  },
  {
    file: 'nonAlphaCommandNames2.docx',
    options: {
      data: { 姓名: 'hong', 标题: 'junyao' },
      cmdDelimiter: ['{', '}'],
    },
  },
  {
    file: 'objectCommandResultError.docx',
    options: { data: { companies: { a: 1 } } },
  },
  {
    file: 'office365.docx',
    options: {
      data: { test: 'first', test2: 'second' },
      cmdDelimiter: ['{', '}'],
    },
  },
  { file: 'referenceError.docx', options: { data: {} } },
  {
    file: 'rejectNullishIMAGE.docx',
    options: {
      data: {},
      additionalJsContext: JS_CONTEXT,
      rejectNullish: true,
    },
  },
  {
    file: 'rejectNullishINS.docx',
    options: {
      data: { testobj: { value: 'the value is now set' }, test2: 'second' },
      rejectNullish: true,
    },
  },
  { file: 'simpleQuery.docx', options: { data: {} } },
  {
    file: 'simpleQuerySimpleInserts.docx',
    options: { data: { a: 'foo', b: 'bar' } },
  },
  { file: 'splitDelimiters.docx', options: { data: { foo: 'bar' } } },
  { file: 'tableWithHTML.docx', options: { data: {} } },
  { file: 'unmatchedEndFor.docx', options: { data: {} } },
  { file: 'unmatchedEndIf.docx', options: { data: {} } },
  {
    file: 'wbs.docx',
    options: {
      data: {
        project: {
          name: '@deitum/docxtemplates',
          workPackages: [
            {
              acronym: 'WP1',
              title: 'Work Package 1',
              startMilestone: { acronym: 'M1', plannedDelta: '0 m' },
              endMilestone: { acronym: 'M2', plannedDelta: '2 m' },
              leaderCompany: { acronym: 'me' },
            },
            {
              acronym: 'WP2',
              title: 'Work Package 2',
              startMilestone: { acronym: 'M2', plannedDelta: '2 m' },
              endMilestone: { acronym: 'M3', plannedDelta: '4 m' },
              leaderCompany: {},
            },
          ],
        },
      },
    },
  },
];

/**
 * Guards against a fixture being added without a corpus case — which would let
 * a whole template shape go unwatched.
 */
export const uncoveredFixtures = (): string[] => {
  const covered = new Set(CORPUS.map(c => c.file));
  return fs
    .readdirSync(fixturesDir)
    .filter(f => /\.doc[xm]$/.test(f) && !covered.has(f))
    .sort();
};
